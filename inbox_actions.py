"""
inbox_actions.py  —  Outlook Inbox Action Extractor  (consolidated)
--------------------------------------------------------------------
Scans the Outlook mailbox (last N days), extracts actionable items using
Claude, groups them by topic, and pushes approved tasks to Microsoft To Do
via a local review dashboard.

Changes from the previous standalone version:
  • Shared infrastructure moved to core/ (auth, Graph, retry, JSON parsing).
  • Access tokens are NO LONGER baked into the dashboard HTML (S1). Push and
    Resolve both go through the local review server, which also restores
    pushed-to-To-Do state tracking (B1).
  • One combined post-run Claude call updates inbox state, company knowledge,
    and briefing suggestions (previously three separate calls).
  • Report headers show the lookback you actually chose (B6).
  • Decision-log matching normalises RE:/FW: prefixes (B11).
  • Non-interactive friendly: pass --days to skip the prompt (for scheduling).

Usage:
  py inbox_actions.py setup                — one-time browser auth (shared by all tools)
  py inbox_actions.py test                 — verify connection + list task lists
  py inbox_actions.py run [--days N]       — scan → extract → review dashboard
  py inbox_actions.py analyse [--days N]   — scan + extract only (no dashboard)
  py inbox_actions.py review               — re-open the last run's dashboard
  py inbox_actions.py build-context [--days N]  — deep scan → draft context_briefing.txt
  py inbox_actions.py show-state           — print current inbox state
  py inbox_actions.py clear-state          — reset inbox_state.json
  py inbox_actions.py show-knowledge       — print company_knowledge.json

.env (same folder or parent):
  OUTLOOK_CLIENT_ID=...   OUTLOOK_TENANT_ID=consumers   ANTHROPIC_API_KEY=sk-ant-...
"""

import re
import sys
import json
import html
import argparse
import datetime
import time
import threading
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler

from core.env import load_env
load_env()

from core.config import BASE_DIR, TASK_LIST_NAME, SCHEDULING_RULES_FILE
from core import graph
from core.llm import get_client, call_json, api_call_with_retry
from core.knowledge import (CompanyKnowledge, InboxState,
                            load_context_briefing, append_briefing)

# Stage 2 of the review dashboard (diary blocks) is provided by the scheduler.
# Imported lazily-ish and guarded so a missing/broken scheduler degrades to the
# old task-only dashboard rather than breaking the inbox scan.
try:
    import outlook_scheduler as scheduler
    _SCHEDULER_AVAILABLE   = True
    _SCHEDULER_IMPORT_ERROR = ""
except Exception as _sched_err:            # pragma: no cover - env dependent
    scheduler               = None
    _SCHEDULER_AVAILABLE    = False
    _SCHEDULER_IMPORT_ERROR = str(_sched_err)

# ─────────────────────────────────────────────
# CONFIG (tool-specific)
# ─────────────────────────────────────────────

LOOKBACK_DAYS         = 14
MAX_EMAILS_PER_FOLDER = 75
MAX_EMAILS_TOTAL      = 400
MAX_BODY_CHARS        = 1000    # action signals are almost always early in the body
MAX_EMAILS_PER_BATCH  = 150
DECISION_BATCH        = 100

# The To Do list approved actions are pushed into. Defined once in
# core/config.py — outlook_scheduler.py reads its work queue from the same
# constant, so tasks created here are always visible to the scheduler.
DEFAULT_TASK_LIST_NAME = TASK_LIST_NAME

OUTPUT_REPORT           = BASE_DIR / "inbox_action_report.txt"
OUTPUT_SUMMARY_MD       = BASE_DIR / "inbox_action_summary.md"
OUTPUT_DECISION_LOG_TXT = BASE_DIR / "inbox_decision_log.txt"
OUTPUT_ACTIONS_JSON     = BASE_DIR / ".inbox_actions.json"
OUTPUT_DASHBOARD_HTML   = BASE_DIR / "inbox_dashboard.html"
OUTPUT_CONTEXT_BRIEFING = BASE_DIR / "context_briefing.txt"


# ─────────────────────────────────────────────
# EXTRACTION  (Call 1 — batched)
# ─────────────────────────────────────────────

def _extract_batch(client, emails: list, batch_num: int, total_batches: int,
                   briefing_text: str, state_context: str,
                   knowledge_block: str) -> list:
    today = datetime.date.today().strftime("%d %B %Y")

    briefing_section = (f"CONTEXT BRIEFING (read before reviewing emails):\n"
                        f"{briefing_text}\n\n" if briefing_text else "")
    state_section    = f"{state_context}\n" if state_context else ""
    sk_section       = f"{knowledge_block}\n" if knowledge_block else ""

    prompt = f"""You are an executive assistant reviewing an email inbox.
Today's date is {today}. You are reviewing batch {batch_num} of {total_batches}.
Your job is to read the emails below and identify every clear action that the
inbox owner needs to take.

{sk_section}{state_section}{briefing_section}INSTRUCTIONS:
- Extract only genuine, concrete actions — not vague observations.
- Include actions from received emails (things others are asking for) AND
  from sent emails (commitments the owner has made that need follow-up).
- Ignore newsletters, automated notifications, receipts, and purely
  informational emails with no action required.
- Group related actions under a meaningful topic label (e.g. a project name,
  client name, or functional area like "Finance", "HR", "Legal", "IT").
- For each action, note: what needs to be done, who/what triggered it,
  and whether there is a stated or implied deadline.
- If a deadline is explicitly mentioned, set suggested_due to that date in
  YYYY-MM-DD format. If the email implies urgency but no specific date, set
  it to tomorrow. Otherwise leave it as today's date ({datetime.date.today().isoformat()}).
- IMPORTANT: If the cumulative memory above lists open actions from prior runs,
  and new emails in this batch appear to progress, update, or resolve those
  prior actions — include them in the relevant topic group so the state updater
  can reconcile them. Note in the detail field if an action is "continuation of
  prior open action" or "appears resolved based on reply".

RESPOND ONLY WITH VALID JSON — a single array of action_group objects
(no wrapper object, no markdown, no preamble):

[
  {{
    "topic": "Topic Label",
    "theme_summary": "One or two sentences describing what is happening in this topic area.",
    "actions": [
      {{
        "title": "Short action title (max 80 chars, starts with a verb)",
        "detail": "One or two sentences of context — who asked, what for, any deadline",
        "from_name": "Name of person who triggered this",
        "email_subject": "Subject of the source email",
        "email_date": "YYYY-MM-DD",
        "priority": "high | normal | low",
        "suggested_due": "YYYY-MM-DD"
      }}
    ]
  }}
]

Aim for quality over quantity. If there are no genuine actions, return an empty array: []

{'-'*60}
EMAILS ({len(emails)} in this batch):
{'-'*60}

{graph.build_email_context(emails)}
"""

    groups = call_json(client, prompt, max_tokens=16000,
                       label=f"Batch {batch_num}/{total_batches} extraction")
    if groups is None:
        return []
    if isinstance(groups, dict):
        groups = groups.get("action_groups", [])
    if not isinstance(groups, list):
        return []
    total = sum(len(g.get("actions", [])) for g in groups)
    print(f"       ✅  Batch {batch_num}/{total_batches}: "
          f"{total} actions across {len(groups)} topics")
    return groups


def _merge_action_groups(all_groups: list) -> list:
    """Merge groups across batches; dedupe by (title, email_subject)."""
    merged: dict = {}
    for g in all_groups:
        topic = g.get("topic", "Uncategorised")
        if topic not in merged:
            merged[topic] = {"topic": topic,
                             "theme_summary": g.get("theme_summary", ""),
                             "actions": [], "_seen": set()}
        for a in g.get("actions", []):
            key = (a.get("title", "").lower(), a.get("email_subject", "").lower())
            if key not in merged[topic]["_seen"]:
                merged[topic]["_seen"].add(key)
                merged[topic]["actions"].append(a)
    result = []
    for v in merged.values():
        v.pop("_seen", None)
        result.append(v)
    return result


def _norm_subject(subject: str) -> str:
    """Normalise a subject for matching — strip RE:/FW:/FWD: prefixes (B11)."""
    s = subject.strip().lower()
    while True:
        new = re.sub(r'^(re|fw|fwd)\s*:\s*', '', s)
        if new == s:
            return re.sub(r'\s+', ' ', s)
        s = new


def extract_actions(emails: list, briefing_text: str, state_context: str,
                    knowledge_block: str, lookback_days: int) -> dict:
    """Call 1 (batched extraction) + Call 2 (decision log)."""
    if not emails:
        return {}
    client = get_client()

    batches = [emails[i:i + MAX_EMAILS_PER_BATCH]
               for i in range(0, len(emails), MAX_EMAILS_PER_BATCH)]
    total_batches = len(batches)

    print(f"  🤖  Call 1/2 — extracting action groups…")
    print(f"       {len(emails)} emails split into {total_batches} batch(es)\n")

    all_groups: list = []
    for i, batch in enumerate(batches, 1):
        print(f"       Batch {i}/{total_batches}: {len(batch)} emails",
              end=" … ", flush=True)
        all_groups.extend(_extract_batch(client, batch, i, total_batches,
                                         briefing_text, state_context,
                                         knowledge_block))

    merged = _merge_action_groups(all_groups)
    total_actions = sum(len(g.get("actions", [])) for g in merged)
    print(f"\n       ✅  {total_actions} actions across {len(merged)} topics (after merge)\n")
    if not merged:
        return {}

    data = {
        "summary": f"Inbox review: {total_actions} actions across {len(merged)} topics.",
        "total_emails_reviewed": len(emails),
        "lookback_days": lookback_days,
        "action_groups": merged,
    }

    # ── Call 2: decision log (metadata only) ──────────────────────────────
    included_subjects = []
    for g in merged:
        for a in g.get("actions", []):
            included_subjects.append(f"  • [{g['topic']}] {a.get('email_subject', '')}")
    included_text = "\n".join(included_subjects) if included_subjects else "  (none)"

    print("  🤖  Call 2/2 — generating decision log…")
    decision_log_full = []
    email_batches = [emails[i:i+DECISION_BATCH]
                     for i in range(0, len(emails), DECISION_BATCH)]

    for db_idx, email_batch in enumerate(email_batches, 1):
        offset = (db_idx - 1) * DECISION_BATCH
        meta_lines = [
            f"{j}. [{e['datetime'][:10]}] {e['from_name']} <{e['from_email']}> — {e['subject']}"
            for j, e in enumerate(email_batch, 1)
        ]
        print(f"       Batch {db_idx}/{len(email_batches)}: {len(email_batch)} emails",
              end=" … ", flush=True)

        prompt_log = f"""You are an executive assistant. Below is a numbered list of {len(email_batch)} emails
from an inbox review, followed by the list of action items that were extracted.

Your task: produce a decision_log JSON array with one entry per email, in order.

Cross-reference each email subject against the included subjects list (treat
RE:/FW: variants of the same subject as matching). Use decision value "action"
if it generated an action item, or "no action" if it did not.

For the "reason" field:
- If decision is "action": briefly state what the action is and why it was raised.
- If decision is "no action": write a specific sentence explaining what the email
  actually was and exactly why no action is needed. Do NOT use vague labels like
  "FYI only" — describe the content so the inbox owner can judge the decision
  without re-opening the email.

RESPOND ONLY WITH VALID JSON — a single array, no wrapper object, no markdown:

[
  {{
    "email_number": 1,
    "subject": "Exact subject from the list",
    "from_name": "Sender name",
    "email_date": "YYYY-MM-DD",
    "decision": "action | no action",
    "topic": "Topic Label if action, otherwise null",
    "reason": "Specific sentence explaining the decision"
  }}
]

Every email in this batch must have an entry. Do not skip any numbers.

{'-'*60}
EMAILS ({len(email_batch)} in this batch):
{'-'*60}
{chr(10).join(meta_lines)}

{'-'*60}
ACTIONS EXTRACTED (cross-reference subjects against these):
{'-'*60}
{included_text}
"""
        batch_log = call_json(client, prompt_log, max_tokens=16000,
                              label=f"Decision log batch {db_idx}")
        if isinstance(batch_log, list):
            for entry in batch_log:
                entry["email_number"] = entry.get("email_number", 0) + offset
            decision_log_full.extend(batch_log)
            print(f"✅  {len(batch_log)} entries")
        else:
            print("⚠️  skipped (fallback fills gaps)")

    data["decision_log"] = decision_log_full
    print(f"       ✅  Decision log total: {len(decision_log_full)} entries\n")
    return data


# ─────────────────────────────────────────────
# COMBINED POST-RUN UPDATE  (one Claude call)
# ─────────────────────────────────────────────

def post_run_update(state: InboxState, ck: CompanyKnowledge, data: dict,
                    emails: list, lookback_days: int, briefing_text: str):
    """
    Single Claude call replacing the previous three (state patch, shared
    knowledge patch, briefing suggestions). Cuts cost and removes the
    hard-coded rate-limit sleeps between calls.
    """
    client = get_client()
    today  = datetime.date.today().isoformat()
    groups = data.get("action_groups", [])

    run_lines = [
        f"RUN DATE: {today}", f"LOOKBACK: {lookback_days} days",
        f"EMAILS REVIEWED: {data.get('total_emails_reviewed', len(emails))}",
        f"ACTIONS FOUND: {sum(len(g.get('actions', [])) for g in groups)}",
        "", "THIS RUN'S ACTION GROUPS:",
    ]
    for g in groups:
        run_lines.append(f"  Topic: {g['topic']}")
        run_lines.append(f"  Theme: {g.get('theme_summary', '')[:200]}")
        for a in g.get("actions", []):
            run_lines.append(
                f"    • [{a.get('priority','normal')}] {a.get('title','')[:120]}"
                f"  |  From: {a.get('from_name','')}"
                f"  |  Due: {a.get('suggested_due','')}")
            if a.get("detail"):
                run_lines.append(f"      {a['detail'][:300]}")
        run_lines.append("")
    run_lines.append(graph.top_senders_block(emails, 20))
    run_summary = "\n".join(run_lines)

    def _trunc(a):
        out = dict(a)
        if out.get("detail") and len(out["detail"]) > 300:
            out["detail"] = out["detail"][:300]
        return out

    current_state_json = json.dumps({
        "open_actions": [_trunc(a) for a in state.data.get("open_actions", [])],
        "topics":       state.data.get("topics", {}),
        "key_people":   state.data.get("key_people", {}),
    }, indent=2, ensure_ascii=False)

    current_knowledge_json = json.dumps({
        "company_profile": ck.data.get("company_profile", ""),
        "matters":         ck.data.get("matters", {}),
        "key_people":      list(ck.data.get("key_people", {}).keys()),
        "facts":           ck.data.get("facts", []),
    }, indent=2, ensure_ascii=False)
    if len(current_knowledge_json) > 25000:
        current_knowledge_json = current_knowledge_json[:25000] + "\n… (truncated)"

    briefing_section = (f"CURRENT CONTEXT BRIEFING:\n{briefing_text}"
                        if briefing_text else
                        "CURRENT CONTEXT BRIEFING: (none yet)")

    prompt = f"""You are maintaining the persistent memory for an executive's inbox tooling.
Today's date is {today}. Produce ONE JSON object with three top-level keys:
"state_patch", "knowledge_patch", and "briefing_additions".

═══════════════════════════════════════════════════════
1. CURRENT PERSONAL TASK STATE (inbox_state.json)
═══════════════════════════════════════════════════════
{current_state_json}

═══════════════════════════════════════════════════════
2. CURRENT COMPANY KNOWLEDGE (company_knowledge.json)
═══════════════════════════════════════════════════════
{current_knowledge_json}

═══════════════════════════════════════════════════════
3. {briefing_section}
═══════════════════════════════════════════════════════

═══════════════════════════════════════════════════════
THIS RUN'S RESULTS
═══════════════════════════════════════════════════════
{run_summary}

═══════════════════════════════════════════════════════
INSTRUCTIONS
═══════════════════════════════════════════════════════

STATE_PATCH — personal task tracking:
1. new_open_actions: every action in this run's groups NOT already in
   open_actions (match by title similarity). Assign ids "act-NNN"
   incrementing from the highest existing id.
2. resolve_action_ids: previously open actions that appear addressed
   (a reply confirms it, or absent from this window AND >14 days old).
   Be conservative — when in doubt, leave open.
3. topic_updates: update/create topics seen this run (last_seen, rolling
   2-3 sentence summary, open_actions list, key_contacts, status "active").
4. people_updates: frequent or important senders only — not mailing lists.

KNOWLEDGE_PATCH — cross-cutting company facts (shared with the board tools):
Only include information that is (a) genuinely new, (b) relevant to BOTH task
management and board reporting, and (c) stable for weeks. Empty arrays if
nothing qualifies. Do NOT duplicate personal task detail here.

BRIEFING_ADDITIONS — plain-text additions to the standing context briefing:
Only genuinely new recurring topics, key people, or standing rules that would
improve future scans, formatted as short lines:
  NEW MATTER: name — one sentence
  NEW PERSON: name — role/relevance
  NEW INSTRUCTION: standing rule
If nothing is new, use an empty string "".

RESPOND ONLY WITH VALID JSON (no markdown, no preamble):

{{
  "state_patch": {{
    "new_open_actions": [
      {{"id": "act-NNN", "title": "...", "topic": "...", "detail": "...",
        "from_name": "...", "raised_date": "{today}",
        "suggested_due": "YYYY-MM-DD", "priority": "high | normal | low",
        "status": "open"}}
    ],
    "resolve_action_ids": ["act-001"],
    "topic_updates": [
      {{"topic": "...", "first_seen": "YYYY-MM-DD", "last_seen": "{today}",
        "status": "active", "summary": "...", "open_actions": ["..."],
        "resolved_actions": [], "key_contacts": ["Name <email>"], "notes": ""}}
    ],
    "people_updates": [
      {{"name": "...", "email": "...", "role": "...",
        "first_seen": "YYYY-MM-DD", "last_seen": "{today}", "notes": "..."}}
    ]
  }},
  "knowledge_patch": {{
    "matter_updates": [
      {{"name": "...", "category": "Commercial | Regulatory | JV/Partner | Operational | Financial | Stakeholder | Legal | Other",
        "status": "active | dormant | closed", "description": "...",
        "key_parties": ["..."], "watch_for": "..."}}
    ],
    "close_matter_names": [],
    "people_updates": [
      {{"name": "...", "email": "...", "organisation": "...", "role": "...", "notes": "..."}}
    ],
    "new_facts": [],
    "company_profile": ""
  }},
  "briefing_additions": ""
}}
"""

    print("  🧠  Post-run update (state + knowledge + briefing, one call)…")
    patch = call_json(client, prompt, max_tokens=16000, label="Post-run patch")

    if patch:
        sp = patch.get("state_patch") or {}
        state.apply_patch(sp)
        print(f"       ✅  State: +{len(sp.get('new_open_actions', []))} actions, "
              f"{len(sp.get('resolve_action_ids', []))} resolved, "
              f"{len(sp.get('topic_updates', []))} topics, "
              f"{len(sp.get('people_updates', []))} people")

        kp = patch.get("knowledge_patch") or {}
        ck.apply_patch(kp)
        print(f"       ✅  Knowledge: +{len(kp.get('matter_updates', []))} matters, "
              f"+{len(kp.get('people_updates', []))} people, "
              f"+{len(kp.get('new_facts', []))} facts")

        additions = (patch.get("briefing_additions") or "").strip()
        if additions and not additions.upper().startswith("NOTHING NEW"):
            append_briefing(additions)
            print(f"       📝  context_briefing.txt updated with new patterns")
    else:
        print("       ⚠️  Patch unavailable — run count and timestamp still saved.")

    state.data["run_count"] = state.data.get("run_count", 0) + 1
    state.prune()
    state.save()
    ck.prune()
    ck.save()
    print(f"       📁  State: {state.path.name}   Knowledge: {ck.path.name}\n")


# ─────────────────────────────────────────────
# REPORTS  (lookback now passed through — B6)
# ─────────────────────────────────────────────

def save_markdown_summary(data: dict, emails: list, lookback_days: int):
    now    = datetime.datetime.now()
    groups = data.get("action_groups", [])
    total  = sum(len(g.get("actions", [])) for g in groups)
    since  = (datetime.date.today()
              - datetime.timedelta(days=lookback_days)).strftime("%d %b %Y")
    today  = datetime.date.today().strftime("%d %b %Y")

    with open(OUTPUT_SUMMARY_MD, "w", encoding="utf-8") as f:
        f.write("# Inbox Action Summary\n\n")
        f.write(f"**Generated:** {now.strftime('%d %B %Y, %H:%M')}  \n")
        f.write(f"**Period:** {since} → {today}  \n")
        f.write(f"**Emails reviewed:** {len(emails)}  \n")
        f.write(f"**Actions identified:** {total} across {len(groups)} topic(s)\n\n")
        if data.get("summary"):
            f.write(f"> {data['summary']}\n\n")
        f.write("---\n\n")
        for g in groups:
            actions = g.get("actions", [])
            high    = sum(1 for a in actions if a.get("priority") == "high")
            badge   = f" 🔴 {high} high-priority" if high else ""
            f.write(f"## {g.get('topic', 'General')}  ·  {len(actions)} action(s){badge}\n\n")
            if g.get("theme_summary"):
                f.write(f"{g['theme_summary']}\n\n")
            for a in actions:
                icon = "🔴 " if a.get("priority") == "high" else "• "
                due  = a.get("suggested_due", "")
                f.write(f"- {icon}**{a.get('title','')}**"
                        f"{f' _(due {due})_' if due else ''}  \n")
                f.write(f"  {a.get('detail', '')}\n\n")
        f.write("---\n*Generated by inbox_actions.py — review groupings before pushing to To Do.*\n")
    print(f"       MD summary:   {OUTPUT_SUMMARY_MD}")


def save_decision_log(data: dict, emails: list):
    now    = datetime.datetime.now()
    log    = data.get("decision_log", [])
    groups = data.get("action_groups", [])

    # normalised subject → (topic, action title)   (B11)
    subject_to_action = {}
    for g in groups:
        for a in g.get("actions", []):
            key = _norm_subject(a.get("email_subject", ""))
            if key:
                subject_to_action[key] = (g.get("topic", ""), a.get("title", ""))

    logged_numbers = {e.get("email_number") for e in log}
    for i, email in enumerate(emails, 1):
        if i not in logged_numbers:
            match = subject_to_action.get(_norm_subject(email.get("subject", "")))
            if match:
                topic, title = match
                entry = ("action", topic,
                         f'Linked to task "{title}" under topic "{topic}" '
                         f"— Claude omitted this entry from its log.")
            else:
                entry = ("no action", None,
                         "No matching task found — Claude omitted this entry from its log.")
            log.append({
                "email_number": i,
                "subject":      email.get("subject", ""),
                "from_name":    email.get("from_name", ""),
                "email_date":   email.get("datetime", "")[:10],
                "decision":     entry[0],
                "topic":        entry[1],
                "reason":       entry[2],
            })

    log_sorted = sorted(log, key=lambda e: e.get("email_number", 0))
    actioned   = sum(1 for e in log_sorted if e.get("decision") == "action")

    with open(OUTPUT_DECISION_LOG_TXT, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\nEMAIL DECISION LOG\n" + "=" * 70 + "\n")
        f.write(f"Generated:       {now.strftime('%d %B %Y, %H:%M')}\n")
        f.write(f"Emails reviewed: {len(log_sorted)}\n")
        f.write(f"Action raised:   {actioned}\n")
        f.write(f"No action:       {len(log_sorted) - actioned}\n")
        f.write("=" * 70 + "\n\n")
        for entry in log_sorted:
            decision = entry.get("decision", "no action")
            if decision == "action":
                match = subject_to_action.get(_norm_subject(entry.get("subject", "")))
                if match:
                    detail = f'Topic: {match[0]}  |  Task: "{match[1]}"'
                elif entry.get("topic"):
                    detail = f"Topic: {entry['topic']}  |  {entry.get('reason', '')}"
                else:
                    detail = entry.get("reason", "")
                label = "ACTION"
            else:
                label, detail = "NO ACTION", entry.get("reason", "")
            f.write(f"[{entry.get('email_number', '?'):>3}]  "
                    f"{entry.get('email_date', '')}  {label:<9}  {detail}\n")
            f.write(f"       From:    {entry.get('from_name', '')}\n")
            f.write(f"       Subject: {entry.get('subject', '')}\n\n")


def save_report(data: dict, emails: list, lookback_days: int):
    now    = datetime.datetime.now()
    groups = data.get("action_groups", [])
    total  = sum(len(g.get("actions", [])) for g in groups)
    since  = (datetime.date.today()
              - datetime.timedelta(days=lookback_days)).strftime("%d %b %Y")

    with open(OUTPUT_REPORT, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\nINBOX ACTION REPORT\n" + "=" * 70 + "\n")
        f.write(f"Generated:      {now.strftime('%d %B %Y  %H:%M')}\n")
        f.write(f"Emails scanned: {len(emails)}  ({since} → today)\n")
        f.write(f"Actions found:  {total}\n" + "=" * 70 + "\n\n")
        if data.get("summary"):
            f.write(f"SUMMARY\n{'-'*40}\n{data['summary']}\n\n")
        for g in groups:
            f.write(f"\n{'─'*70}\n  TOPIC: {g['topic'].upper()}\n")
            if g.get("theme_summary"):
                f.write(f"  {g['theme_summary']}\n")
            f.write(f"{'─'*70}\n")
            for i, a in enumerate(g.get("actions", []), 1):
                flag = "🔴 " if a.get("priority") == "high" else ""
                f.write(f"\n  {i}. {flag}{a.get('title','')}\n")
                f.write(f"     Context:  {a.get('detail', '')}\n")
                f.write(f"     From:     {a.get('from_name', '')}  "
                        f"[{a.get('email_subject', '')}]\n")
                f.write(f"     Due:      {a.get('suggested_due', '')}\n")
        f.write("\n\n" + "=" * 70 + "\nEND OF REPORT\n" + "=" * 70 + "\n")

    save_markdown_summary(data, emails, lookback_days)
    save_decision_log(data, emails)
    with open(OUTPUT_ACTIONS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"\n  💾  Output files:")
    print(f"       Report:       {OUTPUT_REPORT}")
    print(f"       Summary:      {OUTPUT_SUMMARY_MD}")
    print(f"       Decision log: {OUTPUT_DECISION_LOG_TXT}")
    print(f"       JSON data:    {OUTPUT_ACTIONS_JSON}\n")


# ─────────────────────────────────────────────
# MICROSOFT TO DO
# ─────────────────────────────────────────────

def get_target_list_name() -> str:
    """
    The To Do list approved actions are pushed into — the same list
    outlook_scheduler.py reads its work queue from. scheduling_rules.json may
    override it via global.task_list, but note the rules dashboard rewrites
    global.* wholesale on save, so core/config.py is the reliable source.
    """
    try:
        if SCHEDULING_RULES_FILE.exists():
            rules = json.loads(SCHEDULING_RULES_FILE.read_text(encoding="utf-8"))
            name  = (rules.get("global", {}) or {}).get("task_list", "")
            if isinstance(name, str) and name.strip():
                return name.strip()
    except Exception as e:
        print(f"  ⚠️  Could not read task_list from scheduling_rules.json ({e}) "
              f"— using '{DEFAULT_TASK_LIST_NAME}'.")
    return DEFAULT_TASK_LIST_NAME


def get_target_task_list(token: str) -> str:
    """
    Resolve the id of the scheduler's source list, creating it if it does not
    exist yet. Only falls back to the mailbox default list if creation fails,
    and says so loudly — a silent fallback is what broke the handoff before.
    """
    wanted = get_target_list_name()
    data   = graph.graph_get(token, "/me/todo/lists")
    lists  = data.get("value", [])

    for lst in lists:
        if lst.get("displayName", "").strip().lower() == wanted.lower():
            return lst["id"]

    # Not there — create it, so pushed tasks land where the scheduler looks.
    try:
        created = graph.graph_post(token, "/me/todo/lists",
                                   {"displayName": wanted})
        if created.get("id"):
            print(f"  📋  Created To Do list '{wanted}' (scheduler source list).")
            return created["id"]
    except Exception as e:
        print(f"  ⚠️  Could not create To Do list '{wanted}': {e}")

    for lst in lists:
        if lst.get("wellknownListName") == "defaultList":
            print(f"  ⚠️  Falling back to the default list — tasks pushed there "
                  f"will NOT be picked up by outlook_scheduler.py until they "
                  f"go overdue.")
            return lst["id"]
    if lists:
        print(f"  ⚠️  Falling back to '{lists[0].get('displayName','')}' — "
              f"tasks pushed there will NOT be scheduled automatically.")
        return lists[0]["id"]
    raise RuntimeError("No To Do task lists found.")


# Back-compat alias — anything still calling the old name keeps working, but
# now gets the scheduler-aware behaviour.
get_default_task_list = get_target_task_list


def list_task_lists(token: str):
    try:
        lists  = graph.graph_get(token, "/me/todo/lists").get("value", [])
        wanted = get_target_list_name()
        print(f"\n  📋  Your Microsoft To Do task lists ({len(lists)} found):\n")
        found = False
        for lst in lists:
            name  = lst.get("displayName", "Unnamed")
            flags = []
            if lst.get("wellknownListName") == "defaultList":
                flags.append("DEFAULT")
            if name.strip().lower() == wanted.lower():
                flags.append("← SCHEDULER SOURCE")
                found = True
            suffix = ("  " + "  ".join(flags)) if flags else ""
            print(f"     • {name}{suffix}")
        if not found:
            print(f"\n     ⚠️  No list named '{wanted}' yet — it will be created "
                  f"on the first push.")
        print()
    except Exception as e:
        print(f"  ⚠️  Could not retrieve task lists: {e}")


def create_todo_task(token: str, list_id: str, title: str, detail: str,
                     due_date: str, priority: str) -> str:
    """
    Create a To Do task. Returns the new task id (truthy) or "" on failure.
    The id matters now: stage 2 of the dashboard needs it to link the diary
    block back to the task and to mark it complete later.
    """
    body = {
        "title":      title,
        "importance": priority if priority in ("high", "low") else "normal",
        "body":       {"contentType": "text", "content": detail},
        "dueDateTime": {"dateTime": f"{due_date}T00:00:00", "timeZone": "UTC"},
    }
    try:
        return graph.graph_post(
            token, f"/me/todo/lists/{list_id}/tasks", body).get("id", "") or ""
    except Exception as e:
        msg = str(e)
        print(f"  ⚠️  Failed to create task '{title[:40]}': {e}")
        if "401" in msg or "Unauthorized" in msg:
            print("      → Access token expired. This is refreshed per push now; "
                  "if it persists, run: py inbox_actions.py setup")
        elif "403" in msg or "Forbidden" in msg:
            print("      → Permission denied — the token lacks Tasks.ReadWrite. "
                  "Delete .outlook_token_cache.bin and run: "
                  "py inbox_actions.py setup")
        return ""


# ─────────────────────────────────────────────
# HTML DASHBOARD  (no token in the file — S1)
# ─────────────────────────────────────────────

def generate_dashboard_html(data: dict, state: "InboxState | None" = None) -> str:
    """
    Self-contained dashboard for reviewing actions. Push and Resolve both go
    through the local review server — no access token is embedded (S1), and
    successful pushes are recorded in inbox_state.json again (B1).
    """
    groups       = data.get("action_groups", [])
    summary      = data.get("summary", "")
    total        = sum(len(g.get("actions", [])) for g in groups)
    total_emails = data.get("total_emails_reviewed", 0)
    now_str      = datetime.datetime.now().strftime("%d %B %Y, %H:%M")

    task_data = []
    for gi, g in enumerate(groups):
        for ai, a in enumerate(g.get("actions", [])):
            pri = a.get("priority", "normal")
            task_data.append({
                "id":       f"{gi}_{ai}",
                "title":    a.get("title", ""),
                "detail":   a.get("detail", ""),
                "due":      a.get("suggested_due", datetime.date.today().isoformat()),
                "priority": pri,
                # Low-priority items are usually two-minute replies — blocking
                # diary time for every one of them clogs the calendar, so they
                # default off. Overridable per task in the dashboard.
                "block":    pri != "low",
            })
    tasks_json = json.dumps(task_data).replace("</script>", "<\\/script>")
    scheduler_ready = "true" if _SCHEDULER_AVAILABLE else "false"
    scheduler_error = json.dumps(
        f"Diary scheduling unavailable ({_SCHEDULER_IMPORT_ERROR}). "
        f"Tasks were still added to To Do."
        if _SCHEDULER_IMPORT_ERROR else "")

    sections = ""
    for gi, g in enumerate(groups):
        topic   = html.escape(g.get("topic", ""))
        theme   = html.escape(g.get("theme_summary", ""))
        actions = g.get("actions", [])
        cards   = ""
        for ai, a in enumerate(actions):
            tid     = f"{gi}_{ai}"
            pri     = a.get("priority", "normal")
            pri_cls = {"high": "pri-high", "normal": "pri-normal",
                       "low": "pri-low"}.get(pri, "pri-normal")
            cards += f"""
      <div class="card" id="card-{tid}">
        <label class="card-check"><input type="checkbox" class="action-cb" data-id="{tid}" onchange="updateCount()"></label>
        <div class="card-body">
          <div class="card-title">{html.escape(a.get("title", ""))}</div>
          <div class="card-detail">{html.escape(a.get("detail", ""))}</div>
          <div class="card-meta">
            <span class="meta-pill">&#128100; {html.escape(a.get("from_name", ""))}</span>
            <span class="meta-pill">&#9993; {html.escape(a.get("email_subject", ""))}</span>
            <span class="meta-pill">&#128197; {html.escape(a.get("suggested_due", ""))}</span>
            <span class="pri-badge {pri_cls}">{pri.upper()}</span>
            <label class="block-lbl" title="Also block time in the diary for this task">
              <input type="checkbox" class="block-cb" data-id="{tid}" onchange="toggleFields('{tid}')" {"checked" if pri != "low" else ""}> &#128197; block time
            </label>
          </div>
          <div class="task-fields{"" if pri != "low" else " hidden"}" id="fields-{tid}">
            <div class="field-group">
              <label>How long will this take?</label>
              <select id="dur-{tid}">
                <option value="15">15 minutes</option>
                <option value="30"{" selected" if pri != "high" else ""}>30 minutes</option>
                <option value="45">45 minutes</option>
                <option value="60"{" selected" if pri == "high" else ""}>1 hour</option>
                <option value="90">1.5 hours</option>
                <option value="120">2 hours</option>
                <option value="180">3 hours</option>
                <option value="240">Half day (4h)</option>
              </select>
            </div>
            <div class="field-group">
              <label>Priority today</label>
              <select id="pri-{tid}">
                <option value="urgent">&#9889; Urgent — must do today</option>
                <option value="high"{" selected" if pri == "high" else ""}>&#9679; High — do this week</option>
                <option value="normal"{" selected" if pri != "high" else ""}>&#9702; Normal — when time allows</option>
              </select>
            </div>
            <div class="field-group">
              <label>Slot preference</label>
              <select id="slot-{tid}">
                <option value="deep_work">Morning (deep work 9–12)</option>
                <option value="afternoon">Afternoon (12–4pm)</option>
                <option value="any" selected>Any available slot</option>
              </select>
            </div>
            <div class="field-group field-notes">
              <label>Notes for scheduler (optional)</label>
              <input type="text" id="notes-{tid}" placeholder="e.g. needs to happen before the Santos call">
            </div>
          </div>
        </div>
        <div class="card-result" id="result-{tid}"></div>
      </div>"""
        theme_html = f'<p class="topic-theme">{theme}</p>' if theme else ""
        n = len(actions)
        sections += f"""
  <div class="topic-block">
    <div class="topic-head">
      <div class="topic-row">
        <span class="topic-name">{topic}</span>
        <span class="topic-badge">{n} action{"s" if n != 1 else ""}</span>
      </div>
      {theme_html}
      <label class="sel-all-lbl">
        <input type="checkbox" class="sel-all-cb" onchange="toggleGroup({gi},this.checked)"> Select all in this topic
      </label>
    </div>
    <div class="cards" id="group-{gi}">{cards}
    </div>
  </div>"""

    # Prior open actions panel
    prior_panel = ""
    if state:
        open_prior = [a for a in state.data.get("open_actions", [])
                      if a.get("status") in ("open", "pushed_to_todo")]
        if open_prior:
            open_prior.sort(key=lambda a: (
                0 if a.get("priority") == "high" else 1,
                a.get("raised_date", "9999")))
            today = datetime.date.today()
            rows = ""
            for a in open_prior:
                try:
                    age = (today - datetime.date.fromisoformat(
                        a.get("raised_date", today.isoformat()))).days
                except ValueError:
                    age = 0
                pri     = a.get("priority", "normal")
                pri_cls = {"high": "pri-high", "normal": "pri-normal",
                           "low": "pri-low"}.get(pri, "pri-normal")
                status_badge = ('<span class="status-pushed">In To Do</span>'
                                if a.get("status") == "pushed_to_todo"
                                else '<span class="status-open">Open</span>')
                aid    = html.escape(a.get("id", ""))
                detail = html.escape(a.get("detail", ""))
                rows += f"""
        <div class="prior-row" id="prior-card-{aid}">
          <label class="prior-check"><input type="checkbox" class="resolve-cb" data-aid="{aid}" onchange="updateResolveCount()"></label>
          <div class="prior-body">
            <div class="prior-left">
              <span class="prior-title">{html.escape(a.get("title", ""))}</span>
              <span class="pri-badge {pri_cls}" style="font-size:10px">{pri.upper()}</span>
              {status_badge}
            </div>
            <div class="prior-meta">
              <span class="meta-pill">&#128193; {html.escape(a.get("topic", ""))}</span>
              <span class="meta-pill">&#128100; {html.escape(a.get("from_name", ""))}</span>
              <span class="meta-pill">&#128197; due {html.escape(a.get("suggested_due", ""))}</span>
              <span class="meta-pill">&#9203; raised {age}d ago</span>
              <span class="meta-pill prior-id">{aid}</span>
            </div>
            {f'<div class="prior-detail">{detail}</div>' if detail else ''}
          </div>
          <div class="prior-result" id="prior-result-{aid}"></div>
        </div>"""

            prior_panel = f"""
  <div class="topic-block prior-block">
    <div class="topic-head prior-head">
      <div class="topic-row">
        <span class="topic-name prior-title-hdr">&#129300; Prior Open Actions</span>
        <span class="topic-badge prior-badge">{len(open_prior)} carried forward</span>
      </div>
      <p class="topic-theme">Actions from previous runs not yet confirmed resolved.
      Tick any that are done and click <strong>Resolve Selected</strong>.
      Memory built across {state.data.get("run_count", 0)} run(s).</p>
      <div style="display:flex;align-items:center;gap:16px;margin-top:6px">
        <label class="sel-all-lbl"><input type="checkbox" onchange="toggleAllPrior(this.checked)"> Select all prior</label>
        <button class="resolve-btn" id="resolve-btn" onclick="resolvePrior()" disabled>&#10003; Resolve Selected</button>
        <span class="resolve-info" id="resolve-info" style="font-size:13px;color:#64748b"></span>
      </div>
    </div>
    <div class="prior-list">{rows}
    </div>
  </div>"""

    summary_html = (f'<div class="header-summary">{html.escape(summary)}</div>'
                    if summary else "")
    state_pill = ""
    if state and state.data.get("run_count"):
        na = len([a for a in state.data.get("open_actions", [])
                  if a.get("status") == "open"])
        state_pill = (f'<span class="stat-pill">&#129504; Run '
                      f'#{state.data["run_count"]+1} &middot; {na} prior open</span>')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Inbox Action Review</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f1f5f9;color:#1e293b;padding-bottom:90px}}
.page-header{{background:#1d4ed8;color:#fff;padding:20px 28px 16px;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2)}}
.hdr-top{{display:flex;justify-content:space-between;align-items:baseline}}
.hdr-title{{font-size:20px;font-weight:700}}
.hdr-date{{font-size:13px;opacity:.75}}
.hdr-stats{{display:flex;gap:12px;margin-top:10px;flex-wrap:wrap}}
.stat-pill{{background:rgba(255,255,255,.18);border-radius:20px;padding:3px 12px;font-size:13px}}
.header-summary{{margin-top:10px;font-size:13px;opacity:.88;border-left:3px solid rgba(255,255,255,.4);padding-left:10px;font-style:italic}}
.main{{max-width:880px;margin:0 auto;padding:20px 16px}}
.topic-block{{background:#fff;border-radius:10px;border:1px solid #e2e8f0;margin-bottom:18px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.topic-head{{padding:16px 18px 12px;border-bottom:1px solid #e2e8f0;background:#f8fafc}}
.topic-row{{display:flex;align-items:center;gap:10px;margin-bottom:6px}}
.topic-name{{font-size:16px;font-weight:700;color:#1d4ed8}}
.topic-badge{{background:#dbeafe;color:#1d4ed8;border-radius:12px;padding:2px 10px;font-size:12px;font-weight:600}}
.topic-theme{{font-size:13px;color:#64748b;margin-bottom:8px;line-height:1.5}}
.sel-all-lbl{{font-size:13px;color:#64748b;cursor:pointer;display:flex;align-items:center;gap:6px}}
.cards{{padding:10px 14px;display:flex;flex-direction:column;gap:8px}}
.card{{display:flex;align-items:flex-start;gap:12px;padding:13px 14px;border-radius:8px;border:1px solid #e2e8f0;background:#fdfdfd;position:relative;transition:border-color .15s,background .15s}}
.card:has(.action-cb:checked){{border-color:#1d4ed8;background:#eff6ff}}
.card-check{{padding-top:2px;flex-shrink:0}}
.action-cb{{width:18px;height:18px;cursor:pointer;accent-color:#1d4ed8}}
.card-body{{flex:1;min-width:0}}
.card-title{{font-size:14px;font-weight:600;margin-bottom:4px;line-height:1.4}}
.card-detail{{font-size:13px;color:#64748b;margin-bottom:9px;line-height:1.5}}
.card-meta{{display:flex;flex-wrap:wrap;gap:6px;align-items:center}}
.meta-pill{{font-size:12px;color:#64748b;background:#f1f5f9;padding:2px 8px;border-radius:4px;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.pri-badge{{font-size:11px;font-weight:700;padding:2px 8px;border-radius:4px;text-transform:uppercase;letter-spacing:.4px}}
.pri-high{{background:#fee2e2;color:#dc2626}}
.pri-normal{{background:#fef3c7;color:#d97706}}
.pri-low{{background:#d1fae5;color:#059669}}
.card-result{{position:absolute;top:10px;right:12px;font-size:13px;font-weight:600}}
.result-ok{{color:#059669}}
.result-fail{{color:#dc2626}}
.prior-block{{border-color:#e0e7ff}}
.prior-head{{background:#eef2ff}}
.prior-title-hdr{{color:#4338ca}}
.prior-badge{{background:#e0e7ff;color:#4338ca}}
.prior-list{{padding:10px 14px;display:flex;flex-direction:column;gap:8px}}
.prior-row{{display:flex;align-items:flex-start;gap:10px;padding:11px 13px;border-radius:8px;border:1px solid #e0e7ff;background:#f5f7ff;transition:border-color .15s,background .15s}}
.prior-row:has(.resolve-cb:checked){{border-color:#4338ca;background:#eef2ff}}
.prior-check{{padding-top:2px;flex-shrink:0}}
.resolve-cb{{width:17px;height:17px;cursor:pointer;accent-color:#4338ca}}
.prior-body{{flex:1;min-width:0}}
.prior-left{{display:flex;align-items:center;gap:8px;margin-bottom:6px;flex-wrap:wrap}}
.prior-title{{font-size:13px;font-weight:600;color:#1e293b}}
.prior-meta{{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:4px}}
.prior-detail{{font-size:12px;color:#64748b;margin-top:4px;line-height:1.4}}
.prior-id{{font-family:monospace;font-size:11px;color:#94a3b8}}
.prior-result{{font-size:12px;font-weight:600;flex-shrink:0;padding-top:2px;min-width:60px;text-align:right}}
.status-open{{font-size:11px;font-weight:600;padding:2px 7px;border-radius:4px;background:#fef3c7;color:#d97706}}
.status-pushed{{font-size:11px;font-weight:600;padding:2px 7px;border-radius:4px;background:#d1fae5;color:#059669}}
.resolve-btn{{background:#4338ca;color:#fff;border:none;border-radius:6px;padding:6px 16px;font-size:13px;font-weight:600;cursor:pointer;transition:background .15s,opacity .15s}}
.resolve-btn:hover{{background:#3730a3}}
.resolve-btn:disabled{{opacity:.4;cursor:not-allowed}}
.resolve-btn.done{{background:#059669}}
.footer{{position:fixed;bottom:0;left:0;right:0;background:#fff;border-top:1px solid #e2e8f0;padding:13px 28px;display:flex;justify-content:space-between;align-items:center;box-shadow:0 -2px 8px rgba(0,0,0,.08);z-index:100;flex-wrap:wrap}}
.sel-info{{font-size:14px;color:#64748b}}
.sel-info strong{{color:#1e293b}}
.push-btn{{background:#1d4ed8;color:#fff;border:none;border-radius:8px;padding:11px 24px;font-size:14px;font-weight:600;cursor:pointer;transition:background .15s,opacity .15s}}
.push-btn:hover{{background:#1e40af}}
.push-btn:disabled{{opacity:.45;cursor:not-allowed}}
.push-btn.done{{background:#059669}}
.close-lnk{{font-size:13px;color:#94a3b8;margin-left:16px;text-decoration:none}}
.close-lnk:hover{{color:#64748b}}
.srv-note{{font-size:12px;color:#94a3b8;width:100%;text-align:center;margin-top:4px}}
.block-lbl{{font-size:12px;color:#0369a1;background:#e0f2fe;padding:2px 8px;border-radius:4px;cursor:pointer;display:inline-flex;align-items:center;gap:5px;user-select:none}}
.block-cb{{width:14px;height:14px;cursor:pointer;accent-color:#0369a1;margin:0}}
.task-fields{{display:grid;grid-template-columns:repeat(3,1fr);gap:9px;margin-top:11px;padding-top:11px;border-top:1px dashed #e2e8f0}}
.task-fields.hidden{{display:none}}
.field-group{{display:flex;flex-direction:column;gap:3px}}
.field-group label{{font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.3px}}
.field-group select,.field-group input{{font-family:inherit;font-size:12px;padding:5px 7px;border:1px solid #cbd5e1;border-radius:5px;background:#fff;color:#1e293b;width:100%}}
.field-notes{{grid-column:1/-1}}
@media(max-width:640px){{.task-fields{{grid-template-columns:1fr}}}}
.sched-block{{border-color:#bae6fd}}
.sched-head{{background:#f0f9ff}}
.sched-title-hdr{{color:#0369a1}}
.sched-badge{{background:#e0f2fe;color:#0369a1}}
.sched-list{{padding:10px 14px;display:flex;flex-direction:column;gap:8px}}
.sched-row{{display:flex;align-items:flex-start;gap:10px;padding:11px 13px;border-radius:8px;border:1px solid #bae6fd;background:#f8fdff}}
.sched-body{{flex:1;min-width:0}}
.sched-name{{font-size:13px;font-weight:600;margin-bottom:6px}}
.sched-reason{{font-size:12px;color:#64748b;margin-top:5px;line-height:1.4}}
.sched-inputs{{display:flex;flex-wrap:wrap;gap:8px;align-items:center}}
.sched-inputs input,.sched-inputs select{{font-family:inherit;font-size:12px;padding:4px 7px;border:1px solid #cbd5e1;border-radius:5px;background:#fff;color:#1e293b}}
.sched-inputs label{{font-size:12px;color:#64748b;display:flex;align-items:center;gap:5px}}
.sched-result{{font-size:12px;font-weight:600;flex-shrink:0;padding-top:2px;min-width:70px;text-align:right}}
.sched-btn{{background:#0369a1;color:#fff;border:none;border-radius:6px;padding:6px 16px;font-size:13px;font-weight:600;cursor:pointer}}
.sched-btn:hover{{background:#075985}}
.sched-btn:disabled{{opacity:.4;cursor:not-allowed}}
.sched-btn.done{{background:#059669}}
.sched-warn{{font-size:12px;color:#b45309;background:#fef3c7;padding:6px 10px;border-radius:6px;margin-top:8px}}
</style>
</head>
<body>
<div class="page-header">
  <div class="hdr-top"><span class="hdr-title">Inbox Action Review</span><span class="hdr-date">{now_str}</span></div>
  <div class="hdr-stats">
    <span class="stat-pill">&#128236; {total_emails} emails reviewed</span>
    <span class="stat-pill">&#9989; {total} actions identified</span>
    <span class="stat-pill">&#128193; {len(groups)} topics</span>
    {state_pill}
    <span class="stat-pill" title="Build marker — if this is missing you are looking at a cached page">&#9881; v2 &middot; tasks + diary</span>
  </div>
  {summary_html}
</div>
<div class="main">
  {prior_panel}
  {sections}
  <div id="sched-panel"></div>
</div>
<div class="footer">
  <div style="display:flex;flex-direction:column;gap:4px">
    <span class="sel-info"><strong id="sel-count">0</strong> of <strong>{total}</strong> new actions selected</span>
    <span class="sel-info" id="resolve-footer-info" style="display:none"><strong id="resolve-count">0</strong> prior action(s) selected to resolve</span>
  </div>
  <div style="display:flex;align-items:center;gap:12px">
    <button class="push-btn" id="push-btn" onclick="pushTasks()" disabled>Push to Microsoft To Do</button>
    <a class="close-lnk" href="/shutdown">Close</a>
  </div>
  <div class="srv-note" id="srv-note">Push &amp; Resolve work while the review server is running.
  If this page was opened as a file, run: py inbox_actions.py review</div>
</div>
<script>
const TASKS={tasks_json};
const SCHEDULER_READY={scheduler_ready};
const SCHEDULER_ERROR={scheduler_error};

function serverAvailable(){{return location.protocol.startsWith('http');}}
(function(){{
  if(!serverAvailable()){{
    const n=document.getElementById('srv-note');
    n.style.color='#dc2626';
    n.textContent='Opened as a file — Push/Resolve disabled. Run: py inbox_actions.py review';
  }}
}})();

function toggleFields(id){{
  const cb=document.querySelector('.block-cb[data-id="'+id+'"]');
  const f=document.getElementById('fields-'+id);
  if(f)f.classList.toggle('hidden',!cb.checked);
}}
function updateCount(){{
  const n=document.querySelectorAll('.action-cb:checked').length;
  document.getElementById('sel-count').textContent=n;
  document.getElementById('push-btn').disabled=n===0||!serverAvailable();
}}
function updateResolveCount(){{
  const n=document.querySelectorAll('.resolve-cb:checked').length;
  document.getElementById('resolve-count').textContent=n;
  const infoEl=document.getElementById('resolve-footer-info');
  if(infoEl) infoEl.style.display=n>0?'':'none';
  const btn=document.getElementById('resolve-btn');
  if(btn) btn.disabled=n===0||!serverAvailable();
  const hdrInfo=document.getElementById('resolve-info');
  if(hdrInfo) hdrInfo.textContent=n>0?n+' selected':'';
}}
function toggleGroup(gi,checked){{
  document.querySelectorAll('#group-'+gi+' .action-cb').forEach(cb=>{{cb.checked=checked;}});
  updateCount();
}}
function toggleAllPrior(checked){{
  document.querySelectorAll('.resolve-cb').forEach(cb=>{{cb.checked=checked;}});
  updateResolveCount();
}}
async function resolvePrior(){{
  const checked=[...document.querySelectorAll('.resolve-cb:checked')];
  if(!checked.length)return;
  const btn=document.getElementById('resolve-btn');
  btn.disabled=true;btn.textContent='Resolving…';
  const ids=checked.map(cb=>cb.dataset.aid);
  try{{
    const resp=await fetch('/resolve',{{method:'POST',
      headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{action_ids:ids}})}});
    const data=await resp.json();
    if(data.ok){{
      ids.forEach(aid=>{{
        const card=document.getElementById('prior-card-'+aid);
        const el=document.getElementById('prior-result-'+aid);
        if(card) card.style.opacity='0.4';
        if(el) el.innerHTML='<span style="color:#059669">✓ Resolved</span>';
      }});
      btn.textContent='✓ '+ids.length+' resolved';
      btn.classList.add('done');
      document.querySelectorAll('.resolve-cb:checked').forEach(cb=>{{cb.disabled=true;}});
      updateResolveCount();
    }}else{{btn.textContent='Error — see terminal';btn.disabled=false;}}
  }}catch(e){{btn.textContent='Server error — check terminal';btn.disabled=false;}}
}}
async function pushTasks(){{
  const checked=[...document.querySelectorAll('.action-cb:checked')];
  if(!checked.length)return;
  const btn=document.getElementById('push-btn');
  btn.disabled=true;btn.textContent='Pushing…';
  const ids=checked.map(cb=>cb.dataset.id);
  // Carry each task's "block time" toggle through with it.
  const val=(p,id,d)=>{{const e=document.getElementById(p+'-'+id);return e?e.value:d;}};
  const tasks=TASKS.filter(t=>ids.includes(t.id)).map(t=>{{
    const cb=document.querySelector('.block-cb[data-id="'+t.id+'"]');
    return Object.assign({{}},t,{{
      block:         cb?cb.checked:t.block,
      duration_mins: parseInt(val('dur',t.id,30),10),
      day_priority:  val('pri',t.id,'normal'),
      slot_pref:     val('slot',t.id,'any'),
      user_notes:    val('notes',t.id,'').trim()
    }});
  }});
  let pushed=[];
  try{{
    const resp=await fetch('/push',{{method:'POST',
      headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{tasks:tasks}})}});
    const data=await resp.json();
    let ok=0,fail=0;
    (data.results||[]).forEach(r=>{{
      const el=document.getElementById('result-'+r.id);
      if(r.success){{
        ok++;
        if(el)el.innerHTML='<span class="result-ok">✓ Added</span>';
        const src=tasks.find(t=>t.id===r.id);
        if(src&&src.block)pushed.push(Object.assign({{}},src,{{task_id:r.task_id}}));
      }}else{{fail++;if(el)el.innerHTML='<span class="result-fail">✗ Failed</span>';}}
    }});
    btn.textContent='✓ '+ok+' added'+(fail?', '+fail+' failed':'');
    btn.classList.add('done');
    document.querySelectorAll('.action-cb, .block-cb').forEach(cb=>{{cb.disabled=true;}});
    const note=document.getElementById('srv-note');
    if(note&&ok>0)note.textContent=ok+' task'+(ok>1?'s':'')+' pushed to Microsoft To Do ✓ (recorded in inbox state)';
  }}catch(e){{
    btn.textContent='Server error — check terminal';
    btn.disabled=false;
    return;
  }}
  if(SCHEDULER_READY&&pushed.length)scheduleBlocks(pushed);
  else if(!SCHEDULER_READY&&pushed.length)renderSchedError(SCHEDULER_ERROR||'Scheduler unavailable — tasks were still added to To Do.');
}}

// ── STAGE 2: propose diary blocks for the tasks just created ───────────────
let BLOCKS=[];

function renderSchedError(msg){{
  document.getElementById('sched-panel').innerHTML=
    '<div class="topic-block sched-block"><div class="topic-head sched-head">'
    +'<div class="topic-row"><span class="topic-name sched-title-hdr">&#128197; Diary Blocks</span></div>'
    +'<div class="sched-warn">'+msg+'</div></div></div>';
}}

async function scheduleBlocks(pushed){{
  document.getElementById('sched-panel').innerHTML=
    '<div class="topic-block sched-block"><div class="topic-head sched-head">'
    +'<div class="topic-row"><span class="topic-name sched-title-hdr">&#128197; Diary Blocks</span>'
    +'<span class="topic-badge sched-badge">finding time…</span></div>'
    +'<p class="topic-theme">Looking for free slots for '+pushed.length+' task'
    +(pushed.length>1?'s':'')+' and estimating how long each will take.</p></div></div>';
  try{{
    const resp=await fetch('/schedule',{{method:'POST',
      headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{tasks:pushed}})}});
    const data=await resp.json();
    if(data.error){{renderSchedError(data.error);return;}}
    BLOCKS=data.blocks||[];
    renderBlocks(data);
  }}catch(e){{renderSchedError('Scheduling failed — check the terminal.');}}
}}

function renderBlocks(data){{
  const unplaced=data.unscheduled||[];
  if(!BLOCKS.length){{
    renderSchedError('No free slots found. Tasks are in To Do but not in the diary — '
      +'check scheduling_rules.json is not too restrictive.');
    return;
  }}
  let rows='';
  BLOCKS.forEach((b,i)=>{{
    rows+='<div class="sched-row" id="sched-row-'+i+'"><div class="sched-body">'
      +'<div class="sched-name">'+escapeHtml(b.title)+'</div>'
      +'<div class="sched-inputs">'
      +'<label>Date <input type="date" id="d-'+i+'" value="'+b.date+'"></label>'
      +'<label>Start <input type="time" id="t-'+i+'" value="'+b.start_time+'" step="900"></label>'
      +'<label>Duration '+durSelect(i,b.duration_mins)+'</label>'
      +'<label><input type="checkbox" class="keep-cb" id="k-'+i+'" checked> include</label>'
      +'</div>'
      +(b.reason?'<div class="sched-reason">'+escapeHtml(b.reason)+'</div>':'')
      +'</div><div class="sched-result" id="sched-result-'+i+'"></div></div>';
  }});
  let warn='';
  if(unplaced.length){{
    warn='<div class="sched-warn">'+unplaced.length+' task'+(unplaced.length>1?'s':'')
      +' could not be placed and remain in To Do without a diary block: '
      +unplaced.map(u=>escapeHtml(u.title)).join('; ')+'</div>';
  }}
  document.getElementById('sched-panel').innerHTML=
    '<div class="topic-block sched-block"><div class="topic-head sched-head">'
    +'<div class="topic-row"><span class="topic-name sched-title-hdr">&#128197; Diary Blocks</span>'
    +'<span class="topic-badge sched-badge">'+BLOCKS.length+' proposed</span></div>'
    +'<p class="topic-theme">Adjust any time or duration, untick anything you do not want, '
    +'then create the blocks in Outlook.</p>'
    +'<div style="display:flex;align-items:center;gap:16px;margin-top:6px">'
    +'<button class="sched-btn" id="sched-btn" onclick="pushEvents()">&#128197; Create Diary Blocks</button>'
    +'<span id="sched-info" style="font-size:13px;color:#64748b"></span></div>'
    +warn+'</div><div class="sched-list">'+rows+'</div></div>';
  document.getElementById('sched-panel').scrollIntoView({{behavior:'smooth',block:'start'}});
}}

function durSelect(i,mins){{
  const opts=[15,30,45,60,90,120,150,180,240];
  return '<select id="m-'+i+'">'+opts.map(m=>
    '<option value="'+m+'"'+(m==mins?' selected':'')+'>'
    +(m>=60?Math.floor(m/60)+'h'+(m%60?m%60+'m':''):m+'m')+'</option>').join('')+'</select>';
}}
function escapeHtml(t){{
  return String(t==null?'':t).replace(/[&<>"']/g,c=>(
    {{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]));
}}

async function pushEvents(){{
  const btn=document.getElementById('sched-btn');
  const items=[];
  BLOCKS.forEach((b,i)=>{{
    if(!document.getElementById('k-'+i).checked)return;
    items.push(Object.assign({{}},b,{{
      index:i,
      date:document.getElementById('d-'+i).value,
      start_time:document.getElementById('t-'+i).value,
      duration_mins:parseInt(document.getElementById('m-'+i).value,10)||b.duration_mins
    }}));
  }});
  if(!items.length){{document.getElementById('sched-info').textContent='Nothing selected.';return;}}
  btn.disabled=true;btn.textContent='Creating…';
  try{{
    const resp=await fetch('/push_events',{{method:'POST',
      headers:{{'Content-Type':'application/json'}},
      body:JSON.stringify({{blocks:items}})}});
    const data=await resp.json();
    let ok=0,fail=0;
    (data.results||[]).forEach(r=>{{
      const el=document.getElementById('sched-result-'+r.index);
      if(r.success){{ok++;if(el)el.innerHTML='<span class="result-ok">✓ In diary</span>';}}
      else{{fail++;if(el)el.innerHTML='<span class="result-fail">✗ Failed</span>';}}
    }});
    btn.textContent='✓ '+ok+' block'+(ok===1?'':'s')+' created'+(fail?', '+fail+' failed':'');
    btn.classList.add('done');
    document.querySelectorAll('.keep-cb').forEach(cb=>{{cb.disabled=true;}});
  }}catch(e){{btn.textContent='Server error — check terminal';btn.disabled=false;}}
}}
</script>
</body>
</html>"""


def save_dashboard_html(data: dict, state: "InboxState | None" = None):
    html_content = generate_dashboard_html(data, state)
    OUTPUT_DASHBOARD_HTML.write_text(html_content, encoding="utf-8")
    print(f"  🗂️   Dashboard saved: {OUTPUT_DASHBOARD_HTML}")


# ─────────────────────────────────────────────
# STAGE 2 — DIARY BLOCKS  (bridge to outlook_scheduler)
# ─────────────────────────────────────────────

def _as_scheduler_task(t: dict, list_id: str) -> dict:
    """
    Translate a dashboard task into the shape outlook_scheduler expects,
    carrying the user's own duration, today-priority, slot preference and
    notes so the AI places the block rather than re-guessing how long it takes.
    """
    today = datetime.date.today().isoformat()
    due   = t.get("due") or today
    return {
        "id":         t.get("task_id") or t.get("id", ""),
        "title":      t.get("title", ""),
        "body":       t.get("detail", ""),
        "due_date":   due,
        "priority":   t.get("day_priority") or t.get("priority", "normal"),
        "list_name":  get_target_list_name(),
        "is_overdue": due < today or t.get("day_priority") == "urgent",
        "list_id":    list_id,
        "user_duration_mins": int(t.get("duration_mins") or 30),
        "slot_pref":  t.get("slot_pref", "any"),
        "user_notes": t.get("user_notes", ""),
    }


def propose_blocks(tasks: list, list_id: str) -> dict:
    """
    Ask the scheduler for proposed diary blocks for the tasks just pushed.
    Nothing is written to the calendar here — the dashboard shows the proposal
    for review first. Returns {"blocks": [...], "unscheduled": [...]} or
    {"error": "..."}.
    """
    if not _SCHEDULER_AVAILABLE:
        return {"error": f"Scheduler unavailable ({_SCHEDULER_IMPORT_ERROR})."}
    if not tasks:
        return {"blocks": [], "unscheduled": []}

    rules = scheduler.load_rules()
    if not rules:
        return {"error": "scheduling_rules.json not found — open the "
                         "Scheduling Rules dashboard and save your rules first."}

    try:
        token = graph.get_access_token()
    except Exception as e:
        return {"error": f"Auth error: {e}"}

    sched_tasks = [_as_scheduler_task(t, list_id) for t in tasks]
    global_rules   = rules.get("global", {})
    days_ahead     = global_rules.get("days_ahead", 5)
    max_days_ahead = max(global_rules.get("max_days_ahead", 60), days_ahead)

    print(f"\n  📅  Finding diary time for {len(sched_tasks)} task(s)…")
    try:
        # Existing events are fetched over the full horizon so widening the
        # search later can't collide with something already in the diary.
        events = scheduler.fetch_upcoming_events(token, days=max_days_ahead + 1)
        scheduled, unscheduled = scheduler.schedule_with_retry(
            sched_tasks, events, rules, None,
            # Honours the exact duration set on each card instead of
            # estimating one — same path as the scheduling dashboard.
            scheduler_fn=scheduler.ai_schedule_tasks_with_durations,
            initial_days=days_ahead,
            max_days=max_days_ahead,
            verbose=True,
        )
    except Exception as e:
        print(f"  ⚠️  Scheduling failed: {e}")
        return {"error": f"Scheduling failed: {e}"}

    blocks = []
    for item in scheduled:
        start = item["start_dt"]
        # The scheduler clamps a block to the free slot it found. Say so rather
        # than silently handing back less time than was asked for.
        asked = item["task"].get("user_duration_mins", 0)
        got   = item.get("estimated_mins", 0)
        reason = item.get("reason", "")
        if asked and got and got < asked:
            reason = (f"⚠️ Shortened to {got} min — the free slot was only that "
                      f"long (you asked for {asked}). "
                      + reason).strip()
        blocks.append({
            "title":         item["task"]["title"],
            "task_id":       item["task"]["id"],
            "list_id":       item["task"].get("list_id", list_id),
            "task_body":     item["task"].get("body", ""),
            "task_due":      item["task"].get("due_date", ""),
            "task_overdue":  item["task"].get("is_overdue", False),
            "date":          start.strftime("%Y-%m-%d"),
            "start_time":    start.strftime("%H:%M"),
            "duration_mins": item.get("estimated_mins",
                                      item["task"].get("user_duration_mins", 30)),
            "reason":        reason,
        })
    print(f"       ✅  {len(blocks)} block(s) proposed, "
          f"{len(unscheduled)} unplaced")
    return {
        "blocks": blocks,
        "unscheduled": [{"title": u["task"]["title"], "reason": u.get("reason", "")}
                        for u in unscheduled],
    }


def create_blocks(blocks: list) -> dict:
    """
    Create the reviewed (and possibly edited) blocks as Outlook calendar
    events. Times come back from the dashboard as local Brisbane wall-clock
    values and are converted here.
    """
    if not _SCHEDULER_AVAILABLE:
        return {"error": f"Scheduler unavailable ({_SCHEDULER_IMPORT_ERROR})."}
    if not blocks:
        return {"results": []}

    rules = scheduler.load_rules()
    try:
        token = graph.get_access_token()
    except Exception as e:
        return {"error": f"Auth error: {e}"}

    cal_blocks = rules.get("calendar_blocks", {})
    results    = []
    for b in blocks:
        idx = b.get("index")
        try:
            start = datetime.datetime.fromisoformat(
                f"{b['date']}T{b['start_time']}:00").replace(
                    tzinfo=scheduler.AEST_OFFSET)
            mins  = max(15, int(b.get("duration_mins", 30)))
            task  = {
                "id":         b.get("task_id", ""),
                "title":      b.get("title", ""),
                "body":       b.get("task_body", ""),
                "due_date":   b.get("task_due", ""),
                "is_overdue": b.get("task_overdue", False),
                "list_name":  get_target_list_name(),
            }
            item = {
                "task":           task,
                "start_dt":       start,
                "end_dt":         start + datetime.timedelta(minutes=mins),
                "estimated_mins": mins,
                "title":          f"{cal_blocks.get('block_prefix', '🎯 ')}{task['title']}",
                "description":    scheduler._build_event_body(task, cal_blocks),
            }
            ok, fail = scheduler.create_calendar_events(token, [item], rules)
            results.append({"index": idx, "success": ok == 1})
        except Exception as e:
            print(f"  ⚠️  Block failed ({b.get('title','')[:40]}): {e}")
            results.append({"index": idx, "success": False})
    return {"results": results}


def serve_dashboard(state: InboxState, token: str):
    """
    Local server for the dashboard: GET / (dashboard), POST /push (create
    To Do tasks server-side + record pushed state — B1/S1), POST /resolve,
    GET /shutdown.
    """
    if not OUTPUT_DASHBOARD_HTML.exists():
        print("  ⚠️  No dashboard HTML found — run 'run' first.")
        return

    def _fresh_token() -> str:
        """
        The token captured at the start of the run expires while the dashboard
        sits open — a scan plus extraction plus review easily exceeds its
        lifetime, and the push then 401s. MSAL silent-refreshes from the cache,
        so ask for a current token at the moment of use and fall back to the
        original only if that fails. (outlook_scheduler.py does the same.)
        """
        try:
            return graph.get_access_token()
        except Exception:
            return token

    try:
        list_id = get_target_task_list(_fresh_token())
    except Exception as e:
        print(f"  ⚠️  Could not resolve To Do list ({e}) — Push will fail.")
        list_id = ""

    dashboard_html = OUTPUT_DASHBOARD_HTML.read_text(encoding="utf-8")
    shutdown_event = threading.Event()

    class Handler(BaseHTTPRequestHandler):
        def _respond(self, body: bytes, ctype: str = "application/json"):
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            # Without this the browser happily re-serves the previous run's
            # dashboard from cache — same URL, same port, stale HTML.
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            path = self.path.split("?")[0]
            if path == "/":
                self._respond(dashboard_html.encode("utf-8"),
                              "text/html; charset=utf-8")
            elif path == "/shutdown":
                self._respond(b"Closing dashboard...", "text/plain")
                shutdown_event.set()
            else:
                self.send_response(404)
                self.end_headers()

        def do_POST(self):
            length = int(self.headers.get("Content-Length", 0))
            try:
                body = json.loads(self.rfile.read(length))
            except Exception:
                body = {}

            if self.path == "/push":
                results, pushed_titles = [], []
                push_token = _fresh_token()
                # The list id was resolved when the server started; re-resolve
                # if that failed, now that we have a valid token again.
                nonlocal_list_id = list_id or get_target_task_list(push_token)
                for t in body.get("tasks", []):
                    task_id = create_todo_task(
                        push_token, nonlocal_list_id, t.get("title", ""),
                        t.get("detail", ""),
                        t.get("due", datetime.date.today().isoformat()),
                        t.get("priority", "normal"))
                    results.append({"id": t.get("id"),
                                    "success": bool(task_id),
                                    "task_id": task_id,
                                    "list_id": nonlocal_list_id})
                    print(f"     {'✓' if task_id else '✗'}  {t.get('title','')[:60]}")
                    if task_id:
                        pushed_titles.append(t.get("title", ""))
                if pushed_titles:
                    state.mark_pushed(pushed_titles)          # B1 restored
                self._respond(json.dumps({"results": results}).encode())

            elif self.path == "/schedule":
                # Stage 2 — propose diary blocks for the tasks just created.
                self._respond(json.dumps(
                    propose_blocks(body.get("tasks", []),
                                   list_id or "")).encode())

            elif self.path == "/push_events":
                # Stage 2 — create the (possibly edited) blocks in Outlook.
                self._respond(json.dumps(
                    create_blocks(body.get("blocks", []))).encode())

            elif self.path == "/resolve":
                action_ids = body.get("action_ids", [])
                try:
                    if action_ids:
                        titles = [a.get("title", "") for a in state.data.get("open_actions", [])
                                  if a.get("id") in set(action_ids)]
                        state.mark_resolved(action_ids)
                        print(f"     ✓  Resolved {len(action_ids)} prior action(s)")
                        try:
                            from core.todo import complete_matching
                            n = complete_matching(_fresh_token(), titles)
                            if n:
                                print(f"     ✓  Ticked off {n} matching task(s) in To Do")
                        except Exception as e:
                            print(f"     ⚠️  Could not update To Do ({e})")
                    resp = {"ok": True, "resolved": len(action_ids)}
                except Exception as e:
                    resp = {"ok": False, "error": str(e)}
                self._respond(json.dumps(resp).encode())
            else:
                self.send_response(404)
                self.end_headers()

        def log_message(self, fmt, *args):
            pass

    for port in (8765, 8766, 8767):
        try:
            server = HTTPServer(("localhost", port), Handler)
            break
        except OSError:
            continue
    else:
        print("  ❌  No free port (8765-8767) for the dashboard.")
        return

    url = f"http://localhost:{port}"
    print(f"\n  🌐  Opening dashboard at {url}")
    print(f"       Push and Resolve are handled by this server — no token in the HTML.")
    if _SCHEDULER_AVAILABLE:
        print(f"       Tasks ticked with 'block time' get proposed diary blocks "
              f"after they're pushed.")
    else:
        print(f"       ⚠️  Diary scheduling unavailable ({_SCHEDULER_IMPORT_ERROR}) "
              f"— tasks only.")
    print(f"       Click 'Close' in the dashboard when done, or Ctrl+C here.\n")
    bust = f"{url}/?v={int(time.time())}"
    threading.Timer(0.8, lambda: webbrowser.open(bust)).start()
    try:
        while not shutdown_event.is_set():
            server.handle_request()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        print("\n  Dashboard closed.\n")


# ─────────────────────────────────────────────
# BUILD-CONTEXT  (one-off deep scan)
# ─────────────────────────────────────────────

def build_context_briefing(token: str, lookback_days: int,
                           refresh_folders: bool = False):
    print(f"\n{'='*60}")
    print(f"  BUILDING CONTEXT BRIEFING — scanning {lookback_days} days")
    print(f"{'='*60}\n")

    emails = graph.run_scan(token, lookback_days,
                            max_per_folder=MAX_EMAILS_PER_FOLDER,
                            max_total=MAX_EMAILS_TOTAL,
                            body_chars=MAX_BODY_CHARS,
                            label="INBOX DEEP SCAN",
                            refresh_folders=refresh_folders)
    if not emails:
        print("  ⚠️  No emails found.")
        return

    client = get_client()
    today  = datetime.date.today().strftime("%d %B %Y")

    BRIEFING_BATCH  = 200
    batches         = [emails[i:i+BRIEFING_BATCH]
                       for i in range(0, len(emails), BRIEFING_BATCH)]
    batch_summaries = []

    print(f"  🤖  Step 1/2 — scanning {len(batches)} batch(es) for patterns…\n")
    for bi, batch in enumerate(batches, 1):
        print(f"       Batch {bi}/{len(batches)}: {len(batch)} emails",
              end=" … ", flush=True)
        batch_prompt = f"""You are an expert executive assistant analysing a batch of emails
({bi} of {len(batches)}) from the past {lookback_days} days.

Extract the following in plain text (no JSON, no markdown):

ONGOING MATTERS: recurring topics that generate actions — name, description, key contacts
KEY PEOPLE: important contacts — name, role, why they matter
PATTERNS: any standing rules or always-monitor keywords you observe

Be concise — bullet points only. Today is {today}.

{'-'*60}
{graph.build_email_context(batch)}"""
        try:
            msg = api_call_with_retry(client, max_tokens=3000,
                                      messages=[{"role": "user",
                                                 "content": batch_prompt}])
            batch_summaries.append(msg.content[0].text.strip())
            print("✅")
        except Exception as e:
            print(f"⚠️  {e}")
            batch_summaries.append(f"[Batch {bi} failed: {e}]")

    print(f"\n  🤖  Step 2/2 — consolidating into briefing document…\n")
    combined = "\n\n".join(f"=== BATCH {i+1} ===\n{s}"
                           for i, s in enumerate(batch_summaries))
    consolidate_prompt = f"""You are an expert executive assistant. Below are pattern summaries
extracted from {len(emails)} emails over {lookback_days} days, split across {len(batches)} batches.

Consolidate these into a single CONTEXT BRIEFING DOCUMENT. Deduplicate across batches.

FORMAT RULES:
- Plain text, not markdown or JSON
- Use the section headers exactly as shown below
- Be specific and factual; 2-4 lines per entry
- Today's date is {today}

OUTPUT FORMAT:

# INBOX CONTEXT BRIEFING
# Generated: {today}  |  Based on: {lookback_days} days / {len(emails)} emails
# Edit this file freely. Claude reads it before every scan.

ONGOING MATTERS
===============
[Matter Name]
  Description: ...
  Status: ...
  Key contacts: ...
  Watch for: ...

KEY PEOPLE
==========
[Full Name] — [Role/Organisation] — [Why they matter to this inbox]

STANDING INSTRUCTIONS
=====================
- [Rule]

TOPICS TO ALWAYS MONITOR
=========================
- [Topic or keyword]: [Why]

{'-'*60}
BATCH SUMMARIES TO CONSOLIDATE:
{'-'*60}

{combined}
"""
    try:
        msg = api_call_with_retry(client, max_tokens=8000,
                                  messages=[{"role": "user",
                                             "content": consolidate_prompt}])
        OUTPUT_CONTEXT_BRIEFING.write_text(msg.content[0].text.strip(),
                                           encoding="utf-8")
        print(f"  ✅  Context briefing saved to: {OUTPUT_CONTEXT_BRIEFING}")
        print(f"\n  NEXT: open context_briefing.txt, review and edit, save.\n")
    except Exception as e:
        print(f"  ❌  Failed to consolidate briefing: {e}\n")


# ─────────────────────────────────────────────
# COMMANDS
# ─────────────────────────────────────────────

def _resolve_days(args, default: int) -> int:
    if args.days is not None:
        return args.days
    try:
        raw = input(f"\n  How many days to look back? [default: {default}]: ").strip()
        return int(raw) if raw else default
    except (ValueError, EOFError, KeyboardInterrupt):
        return default


def _run_pipeline(args, open_dashboard: bool):
    token    = graph.get_access_token()
    lookback = _resolve_days(args, LOOKBACK_DAYS)

    state = InboxState()
    ck    = CompanyKnowledge()
    _sync_from_todo(state, token)
    state_context   = state.build_context_block()
    knowledge_block = ck.build_prompt_block()               # loaded once (P4)
    briefing_text   = load_context_briefing()

    if state_context:
        print(f"  🧠  Loaded cumulative state (run #{state.data.get('run_count',0)+1}, "
              f"{len(state.data.get('open_actions',[]))} open actions, "
              f"{len(state.data.get('topics',{}))} topics)\n")

    emails = graph.run_scan(token, lookback,
                            max_per_folder=MAX_EMAILS_PER_FOLDER,
                            max_total=MAX_EMAILS_TOTAL,
                            body_chars=MAX_BODY_CHARS,
                            label="INBOX ACTION SCAN",
                            refresh_folders=args.refresh_folders)
    if not emails:
        print(f"  ⚠️  No emails found in the last {lookback} days.")
        return

    data = extract_actions(emails, briefing_text, state_context,
                           knowledge_block, lookback)
    if not data:
        print("  ⚠️  No actions extracted.")
        return
    save_report(data, emails, lookback)

    for g in data.get("action_groups", []):
        print(f"       • {g['topic']}  ({len(g.get('actions', []))} actions)")

    post_run_update(state, ck, data, emails, lookback, briefing_text)

    save_dashboard_html(data, state)
    print(f"  ✅  Run complete.")
    if open_dashboard:
        serve_dashboard(state, token)


def _sync_from_todo(state, token):
    """Actions you have already ticked off in To Do (any list) are resolved here too."""
    try:
        from core.todo import sync_completed
        n = sync_completed(state, token)
        if n:
            print(f"  ✓  {n} prior action(s) resolved - completed in To Do")
    except Exception as e:
        print(f"  ⚠️  Could not check To Do for completed actions ({e})")


def cmd_run(args):
    print("\n" + "=" * 60)
    print("  INBOX ACTION EXTRACTOR")
    print("=" * 60)
    try:
        _run_pipeline(args, open_dashboard=True)
    except RuntimeError as e:
        print(f"\n❌  {e}\n")


def cmd_analyse(args):
    print("\n" + "=" * 60)
    print("  INBOX ACTION EXTRACTOR  (analyse only)")
    print("=" * 60)
    try:
        _run_pipeline(args, open_dashboard=False)
        print("  Run 'py inbox_actions.py review' to review and push to To Do.\n")
    except RuntimeError as e:
        print(f"\n❌  {e}\n")


def cmd_review(args):
    """Re-open the last run's dashboard with a live server (fresh token)."""
    if not OUTPUT_ACTIONS_JSON.exists():
        print(f"\n❌  No saved actions found. Run 'py inbox_actions.py run' first.\n")
        return
    try:
        with open(OUTPUT_ACTIONS_JSON, encoding="utf-8") as f:
            data = json.load(f)
        token = graph.get_access_token()
        state = InboxState()
        _sync_from_todo(state, token)
        save_dashboard_html(data, state)     # regenerate with current state
        serve_dashboard(state, token)
    except RuntimeError as e:
        print(f"\n❌  {e}\n")


def _utcnow():
    return datetime.datetime.now(datetime.timezone.utc)


def cmd_doctor(args):
    """
    Diagnose the task → diary pipeline. Prints exactly which file is loaded,
    whether the new dashboard is present, and where the chain breaks.
    """
    import inspect
    print("\n" + "=" * 60)
    print("  PIPELINE DOCTOR")
    print("=" * 60)

    print(f"\n  1. Script actually loaded:")
    print(f"     {inspect.getfile(inspect.currentframe())}")
    print(f"     BASE_DIR: {BASE_DIR}")

    print(f"\n  2. Scheduler import:")
    if _SCHEDULER_AVAILABLE:
        print(f"     ✓ outlook_scheduler loaded from {scheduler.__file__}")
    else:
        print(f"     ✗ FAILED — {_SCHEDULER_IMPORT_ERROR}")
        print(f"       Stage 2 (diary blocks) will not appear.")

    print(f"\n  3. core/ package:")
    print(f"     graph.py: {graph.__file__}")
    print(f"     Calendars.ReadWrite in SCOPES: "
          f"{'✓ yes' if 'Calendars.ReadWrite' in graph.SCOPES else '✗ NO — core/graph.py is the old copy'}")
    print(f"     graph_patch present: "
          f"{'✓ yes' if hasattr(graph, 'graph_patch') else '✗ NO — core/graph.py is the old copy'}")

    print(f"\n  4. Saved dashboard HTML:")
    if not OUTPUT_DASHBOARD_HTML.exists():
        print(f"     ✗ {OUTPUT_DASHBOARD_HTML} does not exist — run 'run' first.")
    else:
        h = OUTPUT_DASHBOARD_HTML.read_text(encoding="utf-8")
        age = datetime.datetime.fromtimestamp(
            OUTPUT_DASHBOARD_HTML.stat().st_mtime).strftime("%d %b %H:%M")
        print(f"     {OUTPUT_DASHBOARD_HTML}  (written {age})")
        for marker, label in (("v2 &middot; tasks + diary", "build marker"),
                              ("block-cb", "block-time toggles"),
                              ("sched-panel", "diary panel"),
                              ("/push_events", "event push endpoint")):
            print(f"     {'✓' if marker in h else '✗'} {label}")
        if "block-cb" not in h:
            print(f"       → This file was generated by the OLD script. "
                  f"Re-run 'run', and confirm you replaced inbox_actions.py.")

    print(f"\n  5. Microsoft Graph:")
    try:
        token = graph.get_access_token()
        print(f"     ✓ token acquired")
    except Exception as e:
        print(f"     ✗ {e}")
        print("\n" + "=" * 60 + "\n")
        return

    try:
        lists  = graph.graph_get(token, "/me/todo/lists").get("value", [])
        wanted = get_target_list_name()
        names  = [l.get("displayName", "") for l in lists]
        print(f"     target list: '{wanted}'")
        print(f"     {'✓ exists' if wanted in names else '✗ missing (will be created on push)'}")
        print(f"     all lists: {', '.join(names)}")
    except Exception as e:
        print(f"     ✗ could not list To Do lists: {e}")

    try:
        graph.graph_get(token, "/me/calendarView", params={
            "startDateTime": _utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
            "endDateTime": (_utcnow()
                            + datetime.timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "$top": 1})
        print(f"     ✓ calendar readable")
    except Exception as e:
        print(f"     ✗ calendar read failed: {e}")
        print(f"       Delete .outlook_token_cache.bin and run: "
              f"py inbox_actions.py setup")

    print("\n" + "=" * 60 + "\n")


def cmd_test(args):
    print("\n🔍  Testing Outlook connection…\n")
    try:
        token = graph.get_access_token()
        me    = graph.graph_get(token, "/me")
        print(f"✅  Connected as: {me.get('displayName')} "
              f"({me.get('mail') or me.get('userPrincipalName')})\n")
        folders = graph.get_all_folders(token, refresh=args.refresh_folders)
        print(f"\n  📁  Folders found: {len(folders)}")
        list_task_lists(token)
    except RuntimeError as e:
        print(f"❌  {e}\n")


def cmd_show_state(args):
    state = InboxState()
    d = state.data
    if not d.get("run_count"):
        print("\n  ℹ️  No state recorded yet. Run 'py inbox_actions.py run' first.\n")
        return
    print("\n" + "=" * 60)
    print("  INBOX STATE SUMMARY")
    print("=" * 60)
    print(f"  Last updated:  {d.get('last_updated', 'n/a')}")
    print(f"  Total runs:    {d.get('run_count', 0)}")
    open_act = d.get("open_actions", [])
    print(f"\n  OPEN ACTIONS ({len(open_act)}):")
    today = datetime.date.today()
    for a in sorted(open_act, key=lambda x: x.get("priority", "") == "high",
                    reverse=True):
        try:
            age = (today - datetime.date.fromisoformat(
                a.get("raised_date", today.isoformat()))).days
        except ValueError:
            age = 0
        pri = "🔴 " if a.get("priority") == "high" else "   "
        print(f"  {pri}[{a.get('id','')}]  {a.get('title','')}  ({age}d old)")
        print(f"         Topic: {a.get('topic','')}  Due: {a.get('suggested_due','')}")
    topics = d.get("topics", {})
    print(f"\n  TOPICS ({len(topics)}):")
    for name, t in topics.items():
        icon = "▸" if t.get("status") == "active" else "·"
        print(f"  {icon} {name}  [{t.get('status','')}]  last: {t.get('last_seen','?')}")
    print(f"\n  KEY PEOPLE: {len(d.get('key_people', {}))}    "
          f"CLOSED (recent): {len(d.get('closed_actions', []))}\n")


def cmd_clear_state(args):
    from core.config import INBOX_STATE_FILE
    if INBOX_STATE_FILE.exists():
        try:
            confirm = input("\n  ⚠️  This deletes all accumulated inbox memory."
                            "\n  Type 'yes' to confirm: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\n  Cancelled.\n")
            return
        if confirm != "yes":
            print("  Cancelled.\n")
            return
        INBOX_STATE_FILE.unlink()
        print(f"  ✅  State file deleted.\n")
    else:
        print("  ℹ️  No state file found — nothing to clear.\n")


def main():
    parser = argparse.ArgumentParser(description="Outlook inbox action extractor")
    parser.add_argument("command", nargs="?", default="help",
                        choices=["setup", "test", "run", "analyse", "review",
                                 "doctor", "build-context", "show-state",
                                 "clear-state", "show-knowledge", "help"])
    parser.add_argument("--days", type=int, default=None,
                        help="Lookback window in days (skips the interactive prompt)")
    parser.add_argument("--refresh-folders", action="store_true",
                        help="Force folder rediscovery (ignore the 7-day cache)")
    args = parser.parse_args()

    if args.command == "setup":
        graph.setup_auth()
    elif args.command == "test":
        cmd_test(args)
    elif args.command == "run":
        cmd_run(args)
    elif args.command == "analyse":
        cmd_analyse(args)
    elif args.command == "review":
        cmd_review(args)
    elif args.command == "doctor":
        cmd_doctor(args)
    elif args.command == "build-context":
        days = args.days if args.days is not None else 365
        try:
            token = graph.get_access_token()
            build_context_briefing(token, days, args.refresh_folders)
        except RuntimeError as e:
            print(f"\n❌  {e}\n")
    elif args.command == "show-state":
        cmd_show_state(args)
    elif args.command == "clear-state":
        cmd_clear_state(args)
    elif args.command == "show-knowledge":
        CompanyKnowledge().print_summary()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
