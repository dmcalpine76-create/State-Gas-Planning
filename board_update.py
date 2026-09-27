"""
board_update.py  —  Weekly Board Update Generator  (consolidated)
------------------------------------------------------------------
Scans the Outlook mailbox for the chosen window and produces a concise,
board-appropriate narrative email — markdown file + Outlook draft.

Changes from the previous standalone version:
  • Shared infrastructure moved to core/ (auth, Graph, retry, JSON parsing).
  • board_context.json is retired — company knowledge now lives in ONE file
    (company_knowledge.json) shared with the other tools; run
    migrate_knowledge.py once to merge your existing files.
  • One post-run knowledge patch call instead of two, and no more fixed
    30-second rate-limit sleeps.
  • Dead code removed (unused dormant section, no-op update-context command).

Usage:
  py board_update.py run [--days N]           — scan → markdown + Outlook draft
  py board_update.py preview [--days N]       — print to terminal only
  py board_update.py build-context [--days N] — deep scan to seed company knowledge
  py board_update.py show-knowledge           — print company_knowledge.json
  py board_update.py setup / test             — shared auth / connection check
"""

import json
import argparse
import datetime

import requests

from core.env import load_env
load_env()

from core.config import (BASE_DIR, COMPANY_NAME, COMPANY_TICKER,
                         MD_NAME, MD_TITLE, VOICE_RULES)
from core import graph
from core.llm import get_client, call_json
from core.knowledge import (CompanyKnowledge, InboxState,
                            load_context_briefing)

# ─────────────────────────────────────────────
# CONFIG (tool-specific)
# ─────────────────────────────────────────────

LOOKBACK_DAYS         = 7
MAX_BODY_CHARS        = 1500     # board narrative needs more context than tasks
MAX_EMAILS_PER_BATCH  = 250
MAX_EMAILS_PER_FOLDER = 75
MAX_EMAILS_TOTAL      = 400

OUTPUT_DIR = BASE_DIR


# ─────────────────────────────────────────────
# CONTEXT ASSEMBLY
# ─────────────────────────────────────────────

def load_supporting_context(ck: CompanyKnowledge) -> str:
    """Company knowledge (primary) + active inbox topics (supplementary)."""
    blocks = []
    block = ck.build_prompt_block(label="COMPANY KNOWLEDGE")
    if block:
        blocks.append(block)

    state = InboxState()
    topics = state.data.get("topics", {})
    active = {k: v for k, v in topics.items() if v.get("status") == "active"}
    if active:
        lines = [
            "INBOX TOOL — ACTIVE TOPICS (from inbox_actions.py):",
            "  (May overlap with company knowledge above — supporting context only.)",
            "",
        ]
        for name, t in active.items():
            lines.append(f"  ▸ {name}  (last: {t.get('last_seen','?')})")
            if t.get("summary"):
                lines.append(f"    {t['summary']}")
        lines += ["", "─" * 60, ""]
        blocks.append("\n".join(lines))

    return "\n".join(blocks)


# ─────────────────────────────────────────────
# EXTRACTION
# ─────────────────────────────────────────────

def extract_board_update_batch(client, emails: list, batch_num: int,
                               total_batches: int, briefing_text: str,
                               supporting_context: str,
                               lookback_days: int) -> list:
    today = datetime.date.today().strftime("%d %B %Y")
    since = (datetime.date.today()
             - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")

    context_section  = f"{supporting_context}\n" if supporting_context else ""
    briefing_section = (f"INBOX CONTEXT BRIEFING:\n{briefing_text}\n\n"
                        if briefing_text else "")

    prompt = f"""You are preparing a weekly board update email on behalf of {MD_NAME},
{MD_TITLE} of {COMPANY_NAME} ({COMPANY_TICKER}).

{context_section}{briefing_section}TODAY: {today}
PERIOD COVERED: {since} to {today}
EMAIL BATCH: {batch_num} of {total_batches}

This is a direct email from the MD to the directors, not a formal board paper.

{VOICE_RULES}

APPLY THE RIGHT SIGNIFICANCE FILTER. The test is not "is this legally or
regulatorily significant" — it is "does the board need to know this right now?"

INCLUDE matters about:
  • Commercial negotiations, deals, or strategic approaches with material developments
  • Regulatory matters requiring board awareness or where the outcome is uncertain
  • JV and partner activity with strategic implications
  • Project milestones or decisions made
  • Financial matters of strategic (not operational) significance
  • Legal matters with material developments or upcoming deadlines

DO NOT INCLUDE:
  • Routine operational matters being managed without board input needed
  • Administrative matters (insurance renewals, standard compliance, routine payments)
  • Matters where the update would be "nothing has changed"
  • Internal logistics or meeting scheduling with no substantive outcome
  • Approaches the MD has assessed as not credible or not worth pursuing

RESPOND ONLY WITH VALID JSON — a single array, no wrapper, no markdown:

[
  {{
    "matter": "Short descriptive title matching Doug's style (e.g. 'Santos JV — ATP 2068/2069 Work Program')",
    "category": "Commercial | Regulatory | JV/Partner | Operational | Financial | Stakeholder | Legal | Other",
    "significance": "high | medium | low",
    "update": "Written in first person as Doug, 2-4 sentences, includes judgment and next step.",
    "date_range": "e.g. '7-10 April 2026'",
    "key_parties": ["Name or organisation mentioned"]
  }}
]

If no emails in this batch contain board-relevant content, return an empty array: []

{'-'*60}
EMAILS ({len(emails)} in this batch):
{'-'*60}

{graph.build_email_context(emails)}
"""

    matters = call_json(client, prompt, max_tokens=8000,
                        label=f"Board batch {batch_num}/{total_batches}")
    if not isinstance(matters, list):
        return []
    print(f"✅  {len(matters)} matter{'s' if len(matters) != 1 else ''}")
    return matters


def merge_matters(all_matters: list) -> list:
    """
    Merge matters across batches (single implementation — previously
    duplicated in board_update.py and board_briefing.py). Keeps the higher
    significance, merges parties, appends genuinely new update text.
    """
    sig_rank = {"high": 0, "medium": 1, "low": 2}
    merged   = {}
    for m in all_matters:
        key = m.get("matter", "").strip().lower()
        if not key:
            continue
        if key not in merged:
            merged[key] = dict(m)
        else:
            ex = merged[key]
            if sig_rank.get(m.get("significance", "low"), 2) < \
               sig_rank.get(ex.get("significance", "low"), 2):
                ex["significance"] = m["significance"]
            parties = set(ex.get("key_parties", []))
            parties.update(m.get("key_parties", []))
            ex["key_parties"] = list(parties)
            new_update = m.get("update", "")
            if new_update and new_update not in ex.get("update", ""):
                ex["update"] = ex.get("update", "").rstrip(".") + " " + new_update
    return sorted(
        merged.values(),
        key=lambda m: (sig_rank.get(m.get("significance", "low"), 2),
                       m.get("category", ""), m.get("matter", "")))


def extract_board_update(emails: list, lookback_days: int,
                         ck: CompanyKnowledge) -> list:
    if not emails:
        return []
    client             = get_client()
    briefing_text      = load_context_briefing()
    supporting_context = load_supporting_context(ck)

    if ck.is_empty():
        print("  ℹ️  company_knowledge.json is empty — consider running "
              "build-context or migrate_knowledge.py first.\n")

    batches = [emails[i:i + MAX_EMAILS_PER_BATCH]
               for i in range(0, len(emails), MAX_EMAILS_PER_BATCH)]
    print(f"  🤖  Extracting board-relevant matters…")
    print(f"       {len(emails)} emails in {len(batches)} batch(es)\n")

    all_matters = []
    for i, batch in enumerate(batches, 1):
        print(f"       Batch {i}/{len(batches)}: {len(batch)} emails … ",
              end="", flush=True)
        all_matters.extend(extract_board_update_batch(
            client, batch, i, len(batches),
            briefing_text, supporting_context, lookback_days))

    merged = merge_matters(all_matters)
    high   = sum(1 for m in merged if m.get("significance") == "high")
    print(f"\n       ✅  {len(merged)} matters identified ({high} high significance)\n")
    return merged


# ─────────────────────────────────────────────
# POST-RUN KNOWLEDGE UPDATE  (one call)
# ─────────────────────────────────────────────

def update_knowledge(ck: CompanyKnowledge, matters: list, emails: list,
                     lookback_days: int):
    """Single patch call to company_knowledge.json (previously two calls
    maintaining two overlapping files)."""
    if not matters:
        ck.save()
        return

    client = get_client()
    today  = datetime.date.today().isoformat()

    lines = [f"SOURCE: board_update.py", f"DATE: {today}",
             f"LOOKBACK: {lookback_days} days",
             f"MATTERS IDENTIFIED: {len(matters)}", ""]
    for m in matters:
        lines.append(f"  [{m.get('significance','').upper()}] "
                     f"[{m.get('category','')}] {m.get('matter','')}")
        lines.append(f"    {m.get('update','')[:200]}")
        if m.get("key_parties"):
            lines.append(f"    Parties: {', '.join(m['key_parties'])}")
        lines.append("")
    lines.append(graph.top_senders_block(emails, 15))

    prompt = ck.build_patch_prompt("\n".join(lines),
                                   source_tool="weekly board update generator")

    print("  🧠  Updating company knowledge…")
    patch = call_json(client, prompt, max_tokens=8000, label="Knowledge patch")
    if patch:
        ck.apply_patch(patch)
        print(f"       ✅  +{len(patch.get('matter_updates', []))} matters, "
              f"{len(patch.get('close_matter_names', []))} closed, "
              f"+{len(patch.get('people_updates', []))} people, "
              f"+{len(patch.get('new_facts', []))} facts")
    else:
        print("       ℹ️  No knowledge updates this run.")
    ck.prune()
    ck.save()
    print(f"       📁  {ck.path.name}\n")


# ─────────────────────────────────────────────
# BUILD-CONTEXT  (deep scan → seed company knowledge)
# ─────────────────────────────────────────────

def build_company_context(token: str, lookback_days: int,
                          refresh_folders: bool = False):
    print(f"\n{'='*60}")
    print(f"  BUILDING COMPANY KNOWLEDGE — scanning {lookback_days} days")
    print(f"{'='*60}\n")

    emails = graph.run_scan(token, lookback_days,
                            max_per_folder=MAX_EMAILS_PER_FOLDER,
                            max_total=MAX_EMAILS_TOTAL,
                            body_chars=MAX_BODY_CHARS,
                            label="BOARD DEEP SCAN",
                            refresh_folders=refresh_folders)
    if not emails:
        print("  ⚠️  No emails found.")
        return

    client = get_client()
    today  = datetime.date.today().strftime("%d %B %Y")

    BATCH   = 200
    batches = [emails[i:i+BATCH] for i in range(0, len(emails), BATCH)]
    batch_summaries = []

    print(f"  🤖  Step 1/2 — scanning {len(batches)} batch(es)…\n")
    for bi, batch in enumerate(batches, 1):
        print(f"       Batch {bi}/{len(batches)}: {len(batch)} emails … ",
              end="", flush=True)
        batch_prompt = f"""You are building a company knowledge file for {COMPANY_NAME} ({COMPANY_TICKER}).
Analyse this batch of emails ({bi} of {len(batches)}) from the past {lookback_days} days.

Extract in plain text (no JSON, no markdown):

ONGOING MATTERS: board-significant recurring topics — name, category, current status, key parties
KEY PEOPLE: external contacts important at board level — name, org, role, why they matter
STANDING FACTS: stable facts about the company or its relationships

Today is {today}. Be specific and concise.

{'-'*60}
{graph.build_email_context(batch)}"""
        try:
            from core.llm import api_call_with_retry
            msg = api_call_with_retry(client, max_tokens=3000,
                                      messages=[{"role": "user",
                                                 "content": batch_prompt}])
            batch_summaries.append(msg.content[0].text.strip())
            print("✅")
        except Exception as e:
            print(f"⚠️  {e}")
            batch_summaries.append(f"[Batch {bi} failed: {e}]")

    print(f"\n  🤖  Step 2/2 — consolidating into company_knowledge.json…\n")
    combined = "\n\n".join(f"=== BATCH {i+1} ===\n{s}"
                           for i, s in enumerate(batch_summaries))
    today_iso = datetime.date.today().isoformat()

    consolidate_prompt = f"""You are building company_knowledge.json for {COMPANY_NAME} ({COMPANY_TICKER}).

Consolidate the batch summaries below into a single structured JSON file.
Deduplicate — merge matters/people that appear in multiple batches.

RESPOND ONLY WITH VALID JSON (no markdown, no preamble):

{{
  "company_profile": "2-3 sentence description of the company and its current strategic position",
  "matter_updates": [
    {{
      "name":        "<matter name>",
      "category":    "Commercial | Regulatory | JV/Partner | Operational | Financial | Stakeholder | Legal | Other",
      "status":      "active | dormant",
      "description": "1-2 sentence current status",
      "key_parties": ["name or org"],
      "watch_for":   "What signals in future emails indicate developments",
      "first_seen":  "{today_iso}",
      "last_active": "{today_iso}"
    }}
  ],
  "people_updates": [
    {{"name": "Full name", "organisation": "...", "role": "...",
      "notes": "Why this person matters"}}
  ],
  "new_facts": ["Stable fact or standing instruction"]
}}

{'-'*60}
BATCH SUMMARIES:
{'-'*60}
{combined}
"""
    patch = call_json(client, consolidate_prompt, max_tokens=8000,
                      label="Knowledge consolidation")
    if patch:
        ck = CompanyKnowledge()
        ck.apply_patch(patch)
        ck.save()
        print(f"  ✅  company_knowledge.json seeded: "
              f"{len(patch.get('matter_updates', []))} matters, "
              f"{len(patch.get('people_updates', []))} people, "
              f"{len(patch.get('new_facts', []))} facts")
        print(f"\n  NEXT: py board_update.py show-knowledge   to review\n"
              f"        py board_update.py run              for your first update\n")
    else:
        print("  ❌  Consolidation failed.\n")


# ─────────────────────────────────────────────
# RENDERERS
# ─────────────────────────────────────────────

def render_markdown(matters: list, lookback_days: int) -> str:
    today      = datetime.date.today()
    since      = today - datetime.timedelta(days=lookback_days)
    period_str = f"{since.strftime('%d %B')} – {today.strftime('%d %B %Y')}"
    sig_rank   = {"high": 0, "medium": 1, "low": 2}
    sorted_matters = sorted(matters,
                            key=lambda m: sig_rank.get(m.get("significance", "low"), 2))

    lines = [
        f"# {COMPANY_NAME} — Board Update",
        f"**Period:** {period_str}  ",
        f"**Prepared:** {today.strftime('%d %B %Y')}  ",
        f"**Distribution:** Directors  ",
        "", "---", "",
        "Dear Directors", "",
        "A quick update on key matters for the period:", "",
    ]
    if not sorted_matters:
        lines += ["*No matters of board significance were identified in this period.*", ""]
    else:
        for i, m in enumerate(sorted_matters, 1):
            lines += [f"### <u>**{i}. {m.get('matter', 'Untitled')}**</u>", ""]
            lines += [m.get("update", "").strip() or "*No detail available.*", ""]
    lines += [
        "---", "",
        "Happy to take any questions on these points or any other matters.",
        "", "Kind regards", "", MD_NAME, MD_TITLE, "",
        f"*Prepared from a {lookback_days}-day review of management email "
        f"correspondence ending {today.strftime('%d %B %Y')}.*", "",
    ]
    return "\n".join(lines)


def render_html(matters: list, lookback_days: int) -> str:
    import html as _html
    today      = datetime.date.today()
    since      = today - datetime.timedelta(days=lookback_days)
    period_str = f"{since.strftime('%d %B')} – {today.strftime('%d %B %Y')}"
    sig_rank   = {"high": 0, "medium": 1, "low": 2}
    base_style = ("font-family: Aptos, 'Aptos (Body)', Calibri, Arial, sans-serif; "
                  "font-size: 11pt; color: #1a1a1a; line-height: 1.6;")

    def e(s):
        return _html.escape(str(s))

    sorted_matters = sorted(matters,
                            key=lambda m: sig_rank.get(m.get("significance", "low"), 2))
    parts = [f'<div style="{base_style} max-width: 700px;">']
    parts.append(f'''
<table width="100%" cellpadding="0" cellspacing="0"
       style="border-bottom: 2px solid #1d4ed8; margin-bottom: 16px; padding-bottom: 10px;">
  <tr>
    <td style="{base_style}">
      <div style="font-size: 15pt; font-weight: bold; color: #1d4ed8;">
        {e(COMPANY_NAME)} — Board Update
      </div>
      <div style="font-size: 10pt; color: #64748b; margin-top: 4px;">
        {e(period_str)} &nbsp;|&nbsp; {e(today.strftime("%d %B %Y"))}
      </div>
    </td>
  </tr>
</table>''')
    parts.append(
        f'<p style="{base_style} margin-bottom: 16px;">Dear Directors</p>'
        f'<p style="{base_style} margin-bottom: 20px;">A quick update on key '
        f'matters for the period:</p>')
    if not sorted_matters:
        parts.append(f'<p style="{base_style} color:#64748b;"><em>No matters of '
                     f'board significance were identified in this period.</em></p>')
    else:
        for i, m in enumerate(sorted_matters, 1):
            parts.append(
                f'<p style="{base_style} font-weight: bold; text-decoration: underline; '
                f'margin: 18px 0 6px 0;">{i}. {e(m.get("matter", "Untitled"))}</p>')
            update_text = m.get("update", "").strip()
            if update_text:
                parts.append(f'<p style="{base_style} margin: 0 0 4px 0;">{e(update_text)}</p>')
            else:
                parts.append(f'<p style="{base_style} color:#94a3b8; margin:0 0 4px 0;">'
                             f'<em>No detail available.</em></p>')
    parts.append(
        f'<p style="{base_style} margin-top: 24px;">Happy to take any questions '
        f'on these points or any other matters.</p>'
        f'<p style="{base_style} margin-top: 16px;">Kind regards</p>'
        f'<p style="{base_style} margin-top: 8px; margin-bottom: 0;">'
        f'<strong>{e(MD_NAME)}</strong><br>{e(MD_TITLE)}</p>')
    parts.append(
        f'<hr style="border:none; border-top:1px solid #e2e8f0; margin: 24px 0 10px 0;">'
        f'<p style="{base_style} font-size:9pt; color:#94a3b8;">'
        f'Prepared from a {lookback_days}-day review of management email '
        f'correspondence ending {e(today.strftime("%d %B %Y"))}.</p>')
    parts.append('</div>')
    return "\n".join(parts)


def create_outlook_draft(token: str, html_body: str) -> bool:
    subject = (f"{COMPANY_NAME} — Board Update "
               f"{datetime.date.today().strftime('%d %B %Y')}")
    try:
        requests.post(
            f"{graph.GRAPH_BASE}/me/messages",
            headers={"Authorization": f"Bearer {token}",
                     "Content-Type": "application/json"},
            json={"subject": subject,
                  "body": {"contentType": "HTML", "content": html_body},
                  "isDraft": True},
            timeout=graph.REQUEST_TIMEOUT,
        ).raise_for_status()
        print(f"  📧  Outlook draft created: \"{subject}\"")
        print(f"       Open Outlook → Drafts to review, address, and send.")
        return True
    except Exception as e:
        print(f"  ⚠️  Could not create Outlook draft: {e}")
        print(f"       The markdown file was saved — you can copy from there.")
        return False


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


def cmd_run(args, preview: bool = False):
    print("\n" + "=" * 60)
    print(f"  {COMPANY_NAME.upper()} — BOARD UPDATE GENERATOR")
    print("=" * 60)
    try:
        token    = graph.get_access_token()
        lookback = _resolve_days(args, LOOKBACK_DAYS)

        ck = CompanyKnowledge()
        if not ck.is_empty():
            n_active = len([m for m in ck.data.get("matters", {}).values()
                            if m.get("status") == "active"])
            print(f"  🧠  Loaded company knowledge: {n_active} active matters\n")

        emails = graph.run_scan(token, lookback,
                                max_per_folder=MAX_EMAILS_PER_FOLDER,
                                max_total=MAX_EMAILS_TOTAL,
                                body_chars=MAX_BODY_CHARS,
                                label="BOARD UPDATE SCAN",
                                refresh_folders=args.refresh_folders)
        if not emails:
            print(f"  ⚠️  No emails found in the last {lookback} days.")
            return

        matters = extract_board_update(emails, lookback, ck)
        if not matters:
            print("  ⚠️  No board-relevant matters identified.")
            print("       Try a longer lookback window.\n")

        md_content = render_markdown(matters, lookback)

        if preview:
            print("\n" + "=" * 60 + "\n  PREVIEW\n" + "=" * 60)
            print(md_content)
        else:
            path = OUTPUT_DIR / f"board_update_{datetime.date.today().isoformat()}.md"
            path.write_text(md_content, encoding="utf-8")
            print(f"  💾  Saved: {path}")
            print(f"\n  ✅  Board update complete — {len(matters)} matters.\n")
            create_outlook_draft(token, render_html(matters, lookback))

        # Post-run knowledge update (single call — no fixed sleeps needed)
        update_knowledge(ck, matters, emails, lookback)

    except RuntimeError as e:
        print(f"\n❌  {e}\n")


def main():
    parser = argparse.ArgumentParser(description="Weekly board update generator")
    parser.add_argument("command", nargs="?", default="help",
                        choices=["setup", "test", "run", "preview",
                                 "build-context", "show-knowledge", "help"])
    parser.add_argument("--days", type=int, default=None,
                        help="Lookback window in days (skips the prompt)")
    parser.add_argument("--refresh-folders", action="store_true",
                        help="Force folder rediscovery")
    args = parser.parse_args()

    if args.command == "setup":
        graph.setup_auth()
    elif args.command == "test":
        try:
            token = graph.get_access_token()
            me    = graph.graph_get(token, "/me")
            print(f"\n✅  Connected as: {me.get('displayName')} "
                  f"({me.get('mail') or me.get('userPrincipalName')})")
            folders = graph.get_all_folders(token, refresh=args.refresh_folders)
            print(f"  📁  {len(folders)} folders\n")
        except RuntimeError as e:
            print(f"\n❌  {e}\n")
    elif args.command == "run":
        cmd_run(args, preview=False)
    elif args.command == "preview":
        cmd_run(args, preview=True)
    elif args.command == "build-context":
        days = args.days if args.days is not None else 90
        try:
            token = graph.get_access_token()
            build_company_context(token, days, args.refresh_folders)
        except RuntimeError as e:
            print(f"\n❌  {e}\n")
    elif args.command == "show-knowledge":
        CompanyKnowledge().print_summary()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
