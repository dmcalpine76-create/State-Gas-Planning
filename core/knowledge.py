"""
core/knowledge.py — the unified memory layer.

Replaces four near-identical classes (StateManager, SharedKnowledgeManager,
BoardContextManager, BoardBriefingContext) with three purpose-distinct stores:

  CompanyKnowledge  company_knowledge.json — matters / key people / standing
                    facts / company profile. Merges the old shared_knowledge.json
                    AND board_context.json (run migrate_knowledge.py once).
                    Read + written by every tool.

  InboxState        inbox_state.json — personal open/closed actions and inbox
                    topics. inbox_actions.py only. (Unchanged schema.)

  PillarHistory     pillar_history.json — last month's RAG status + summary per
                    strategic pillar. board_briefing.py only. (Replaces
                    board_briefing_context.json; matters/people/facts now live
                    in CompanyKnowledge instead of being duplicated here.)

Fixes over the originals:
  • Prompt blocks sort newest-first and budget per section, so the newest
    matters/people/facts are injected instead of truncated away.          (B4)
  • CompanyKnowledge gets prune() — active→dormant→removed over time,
    like inbox_state always had.                                          (B5)
  • context_briefing.txt is read head+tail, so auto-appended additions
    at the bottom of a large file are no longer invisible.                (B8)
"""

import json
import copy
import datetime

from .config import (STORE_DIR, STORE_ACTIVE)
from .config import (COMPANY_KNOWLEDGE_FILE, INBOX_STATE_FILE,
                     PILLAR_HISTORY_FILE, CONTEXT_BRIEFING_FILE,
                     COMPANY_NAME, COMPANY_TICKER, STRATEGIC_PILLARS)


def _today() -> str:
    return datetime.date.today().isoformat()


def _days_since(iso: str) -> int:
    try:
        return (datetime.date.today() - datetime.date.fromisoformat(iso)).days
    except (ValueError, TypeError):
        return 99999


def _load_json(path, empty: dict) -> dict:
    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                raw = json.load(f)
            merged = copy.deepcopy(empty)
            merged.update(raw)
            return merged
        except Exception as e:
            print(f"  ⚠️  Could not read {path.name} ({e}) — starting fresh.")
    return copy.deepcopy(empty)


def _save_json(path, data: dict):
    data["last_updated"] = _today()
    # Write to a temp file and swap it in, so a crash (or OneDrive syncing
    # mid-write) can never leave a half-written knowledge file behind.
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    tmp.replace(path)


# ═════════════════════════════════════════════════════════════════════════════
# COMPANY KNOWLEDGE  (shared by all tools)
# ═════════════════════════════════════════════════════════════════════════════

class CompanyKnowledge:
    """
    Schema
    ──────
    {
      "last_updated":    "YYYY-MM-DD",
      "company_profile": "2-3 sentence description",
      "matters": {
        "<matter name>": {
          "first_seen": "YYYY-MM-DD", "last_active": "YYYY-MM-DD",
          "status": "active | dormant | closed",
          "category": "Commercial | Regulatory | JV/Partner | Operational |
                       Financial | Stakeholder | Legal | Other",
          "description": "...", "key_parties": [...], "watch_for": "..."
        }
      },
      "key_people": {
        "<full name>": { "email", "organisation", "role", "notes",
                         "last_seen": "YYYY-MM-DD" }
      },
      "facts": ["stable factual statement", ...]
    }
    """

    EMPTY = {
        "last_updated":    "",
        "company_profile": "",
        "matters":         {},
        "key_people":      {},
        "facts":           [],
    }

    # Pruning thresholds (B5)
    DORMANT_AFTER_DAYS   = 60    # active matter with no activity → dormant
    CLOSED_RETAIN_DAYS   = 90    # closed matters removed after this
    DORMANT_RETAIN_DAYS  = 365   # dormant matters removed after this

    # In the shared store the same data is kept as three files so each can be
    # read and edited on its own; in memory it is exactly the schema above.
    _SPLIT = (("matters.json", ("matters",)),
              ("people.json",  ("key_people",)),
              ("facts.json",   ("company_profile", "facts")))

    def __init__(self):
        if STORE_ACTIVE:
            self.path = STORE_DIR
            self.data = copy.deepcopy(self.EMPTY)
            for fname, keys in self._SPLIT:
                part = _load_json(STORE_DIR / fname, {})
                for key in keys:
                    if key in part:
                        self.data[key] = part[key]
            self.data["last_updated"] = _load_json(
                STORE_DIR / "facts.json", {}).get("last_updated", "")
        else:
            self.path = COMPANY_KNOWLEDGE_FILE
            self.data = _load_json(self.path, self.EMPTY)

    def save(self):
        if STORE_ACTIVE:
            for fname, keys in self._SPLIT:
                _save_json(STORE_DIR / fname,
                           {key: self.data.get(key, self.EMPTY[key]) for key in keys})
            self.data["last_updated"] = _today()
        else:
            _save_json(self.path, self.data)

    def is_empty(self) -> bool:
        return not (self.data.get("company_profile") or self.data.get("matters")
                    or self.data.get("key_people") or self.data.get("facts"))

    # ── PRUNING (B5) ───────────────────────────────────────────────────────

    def prune(self):
        matters = self.data.get("matters", {})
        remove  = []
        for name, m in matters.items():
            age = _days_since(m.get("last_active", ""))
            status = m.get("status", "active")
            if status == "active" and age > self.DORMANT_AFTER_DAYS:
                m["status"] = "dormant"
            elif status == "closed" and age > self.CLOSED_RETAIN_DAYS:
                remove.append(name)
            elif status == "dormant" and age > self.DORMANT_RETAIN_DAYS:
                remove.append(name)
        for name in remove:
            del matters[name]
        # People: drop entries not seen for 18 months (only if dated)
        people = self.data.get("key_people", {})
        for name in [n for n, p in people.items()
                     if p.get("last_seen") and _days_since(p["last_seen"]) > 540]:
            del people[name]

    # ── PROMPT BLOCK (B4 fixed) ────────────────────────────────────────────

    def build_prompt_block(self, label: str = "COMPANY KNOWLEDGE",
                           max_chars: int = 9000, max_matters: int = 25,
                           max_people: int = 20, max_facts: int = 40) -> str:
        """
        Newest-first with enforced per-section character budgets (matters 55%,
        people 18%, facts 18%, dormant 5%). Previously active matters were
        sorted OLDEST-first and the whole block hard-truncated at 6,000 chars,
        so the newest matters and all people/facts were routinely cut.
        """
        if self.is_empty():
            return ""

        d       = self.data
        matters = d.get("matters", {})
        people  = d.get("key_people", {})
        facts   = d.get("facts", [])

        def fill(entries, budget, limit):
            """Add rendered entries (lists of lines) until the budget is spent."""
            out, used, shown = [], 0, 0
            for entry_lines in entries[:limit]:
                cost = sum(len(l) + 1 for l in entry_lines)
                if used + cost > budget and shown > 0:
                    break
                out.extend(entry_lines)
                used  += cost
                shown += 1
            return out, shown

        lines = [
            "╔══════════════════════════════════════════════════════════╗",
            f"  {label}",
            f"  (company_knowledge.json — last updated: {d.get('last_updated', 'never')})",
            "╚══════════════════════════════════════════════════════════╝",
            "",
        ]

        if d.get("company_profile"):
            lines += ["COMPANY PROFILE:", f"  {d['company_profile']}", ""]

        active = sorted(
            [(k, v) for k, v in matters.items() if v.get("status") == "active"],
            key=lambda x: x[1].get("last_active", ""), reverse=True)   # newest first
        dormant = sorted(
            [(k, v) for k, v in matters.items() if v.get("status") == "dormant"],
            key=lambda x: x[1].get("last_active", ""), reverse=True)

        if active:
            rendered = []
            for name, m in active:
                e = [f"  ▸ [{m.get('category', '')}] {name}"
                     f"  (last active: {m.get('last_active', '?')})"]
                if m.get("description"):
                    e.append(f"    {m['description'][:300]}")
                if m.get("key_parties"):
                    e.append(f"    Parties: {', '.join(m['key_parties'][:4])}")
                if m.get("watch_for"):
                    e.append(f"    Watch for: {m['watch_for'][:150]}")
                rendered.append(e)
            body, shown = fill(rendered, int(max_chars * 0.55), max_matters)
            hdr = f"ONGOING MATTERS (active — {shown} most recent"
            hdr += f" of {len(active)}):" if len(active) > shown else "):"
            lines += [hdr] + body + [""]

        if dormant:
            rendered = []
            for name, m in dormant:
                desc = m.get("description", "")[:80]
                rendered.append([f"  · {name}  (last: {m.get('last_active', '?')})"
                                 + (f" — {desc}" if desc else "")])
            body, shown = fill(rendered, int(max_chars * 0.05), 10)
            lines += ["DORMANT MATTERS (include if emails resurface them):"] + body + [""]

        if people:
            # Most recently added/updated entries sit last in insertion order.
            rendered = []
            for name, p in reversed(list(people.items())):
                parts = [x for x in [p.get("organisation", ""), p.get("role", "")] if x]
                e = [f"  {name}  —  {' / '.join(parts)}"]
                if p.get("notes"):
                    e.append(f"    {p['notes'][:150]}")
                rendered.append(e)
            body, shown = fill(rendered, int(max_chars * 0.18), max_people)
            hdr = f"KEY PEOPLE ({shown}"
            hdr += f" most recent of {len(people)}):" if len(people) > shown else "):"
            lines += [hdr] + body + [""]

        if facts:
            rendered = [[f"  • {f}"] for f in reversed(facts)]   # newest first
            body, shown = fill(rendered, int(max_chars * 0.18), max_facts)
            hdr = f"STANDING FACTS ({shown}"
            hdr += f" most recent of {len(facts)}):" if len(facts) > shown else "):"
            lines += [hdr] + body + [""]

        lines += ["─" * 60, ""]
        block = "\n".join(lines)
        if len(block) > max_chars:
            block = block[:max_chars] + "\n  … (company knowledge truncated)\n"
        return block

    # ── PATCH APPLICATION ──────────────────────────────────────────────────

    def apply_patch(self, patch: dict):
        """
        Patch structure (returned by Claude):
        {
          "matter_updates":     [{name, category, status, description,
                                  key_parties, watch_for, first_seen?, last_active?}],
          "close_matter_names": ["..."],
          "people_updates":     [{name, email, organisation, role, notes}],
          "new_facts":          ["..."],
          "company_profile":    ""   (only if materially changed)
        }
        """
        if not patch:
            return
        today = _today()

        for m in patch.get("matter_updates", []):
            name = (m.get("name") or "").strip()
            if not name:
                continue
            existing = self.data["matters"].get(name, {})
            if existing.get("first_seen") and not m.get("first_seen"):
                m["first_seen"] = existing["first_seen"]
            elif not m.get("first_seen"):
                m["first_seen"] = today
            if not m.get("last_active"):
                m["last_active"] = today
            self.data["matters"][name] = m

        for name in patch.get("close_matter_names", []):
            if name in self.data["matters"]:
                self.data["matters"][name]["status"]      = "closed"
                self.data["matters"][name]["last_active"] = today

        for p in patch.get("people_updates", []):
            name = (p.get("name") or "").strip()
            if not name:
                continue
            existing = self.data["key_people"].get(name, {})
            merged   = {**existing, **{k: v for k, v in p.items() if v}}
            merged["last_seen"] = today
            # Re-insert so newest entries sit at the end (prompt block relies on it)
            self.data["key_people"].pop(name, None)
            self.data["key_people"][name] = merged

        existing_facts = set(self.data.get("facts", []))
        for fact in patch.get("new_facts", []):
            fact = (fact or "").strip()
            if fact and fact not in existing_facts:
                self.data["facts"].append(fact)
                existing_facts.add(fact)

        new_profile = (patch.get("company_profile") or "").strip()
        if new_profile and len(new_profile) > 20:
            self.data["company_profile"] = new_profile

    # ── PATCH PROMPT ───────────────────────────────────────────────────────

    def build_patch_prompt(self, run_summary: str, source_tool: str) -> str:
        today        = _today()
        current_json = json.dumps({
            "company_profile": self.data.get("company_profile", ""),
            "matters":         self.data.get("matters", {}),
            "key_people":      {k: {kk: vv for kk, vv in v.items() if kk != "last_seen"}
                                for k, v in self.data.get("key_people", {}).items()},
            "facts":           self.data.get("facts", []),
        }, indent=2, ensure_ascii=False)
        if len(current_json) > 30000:
            current_json = current_json[:30000] + "\n… (truncated)"

        return f"""You are maintaining the shared company knowledge file for {COMPANY_NAME} ({COMPANY_TICKER}).
This file is used by an inbox action reviewer, a weekly board update generator,
and a monthly board briefing generator. It contains facts that are genuinely
cross-cutting and stable.

SOURCE TOOL: {source_tool}
DATE: {today}

CURRENT COMPANY KNOWLEDGE:
{current_json}

THIS RUN'S FINDINGS:
{run_summary}

YOUR TASK:
Compare the findings against the current knowledge. Identify only information that is:
  (a) genuinely new — not already captured, AND
  (b) cross-cutting — relevant to task management AND board reporting, AND
  (c) stable — will remain true for weeks, not just today

DO NOT add personal task details, one-off events, or anything already present.
Be conservative — it is better to add nothing than to add noise.
If nothing is genuinely new, return empty arrays for each key.

RESPOND ONLY WITH VALID JSON (no markdown, no preamble):

{{
  "matter_updates": [
    {{
      "name":        "Exact or best matter name",
      "category":    "Commercial | Regulatory | JV/Partner | Operational | Financial | Stakeholder | Legal | Other",
      "status":      "active | dormant | closed",
      "description": "Current 1-2 sentence factual status",
      "key_parties": ["name or org"],
      "watch_for":   "What email signals indicate developments"
    }}
  ],
  "close_matter_names": ["matter name if clearly resolved"],
  "people_updates": [
    {{"name": "Full name", "email": "if seen", "organisation": "...", "role": "...",
      "notes": "Why they matter / typical interaction type"}}
  ],
  "new_facts": ["Stable structural fact"],
  "company_profile": ""
}}

Leave company_profile empty unless the current profile is materially wrong or missing.
"""

    # ── DISPLAY ────────────────────────────────────────────────────────────

    def print_summary(self):
        d = self.data
        if self.is_empty():
            print(f"\n  ℹ️  {self.path.name} is empty or not yet created.")
            print(f"      Run migrate_knowledge.py or a build-context command.\n")
            return
        matters = d.get("matters", {})
        people  = d.get("key_people", {})
        facts   = d.get("facts", [])
        by_status = {}
        for m in matters.values():
            by_status[m.get("status", "?")] = by_status.get(m.get("status", "?"), 0) + 1

        print("\n" + "=" * 60)
        print("  COMPANY KNOWLEDGE")
        print("=" * 60)
        print(f"  File:         {self.path}")
        print(f"  Last updated: {d.get('last_updated', 'never')}")
        print(f"\n  PROFILE: {d.get('company_profile', '(none)')}")
        print(f"\n  MATTERS ({len(matters)}): "
              + ", ".join(f"{v} {k}" for k, v in sorted(by_status.items())))
        for name, m in sorted(matters.items(),
                              key=lambda x: x[1].get("last_active", ""), reverse=True):
            if m.get("status") != "active":
                continue
            print(f"\n  ▸ [{m.get('category', '')}] {name}  "
                  f"(last: {m.get('last_active', '?')})")
            print(f"    {m.get('description', '')}")
            if m.get("key_parties"):
                print(f"    Parties: {', '.join(m['key_parties'])}")
        print(f"\n  KEY PEOPLE: {len(people)}    STANDING FACTS: {len(facts)}")
        print()


# ═════════════════════════════════════════════════════════════════════════════
# INBOX STATE  (inbox_actions.py only — schema unchanged)
# ═════════════════════════════════════════════════════════════════════════════

class InboxState:
    """
    Personal task memory. Same schema as before:
    { last_updated, run_count, topics{}, key_people{},
      open_actions[], closed_actions[] }
    """

    EMPTY = {
        "last_updated":   "",
        "run_count":      0,
        "topics":         {},
        "key_people":     {},
        "open_actions":   [],
        "closed_actions": [],
    }

    CLOSED_RETENTION_DAYS = 14
    TOPIC_DORMANT_DAYS    = 45
    TOPIC_PRUNE_DAYS      = 90
    MAX_STATE_CHARS       = 8000

    def __init__(self):
        self.path = INBOX_STATE_FILE
        self.data = _load_json(self.path, self.EMPTY)

    def save(self):
        _save_json(self.path, self.data)

    # ── PRUNING ────────────────────────────────────────────────────────────

    def prune(self):
        for name, t in list(self.data["topics"].items()):
            age = _days_since(t.get("last_seen", ""))
            if age > self.TOPIC_PRUNE_DAYS:
                del self.data["topics"][name]
            elif age > self.TOPIC_DORMANT_DAYS and t.get("status") == "active":
                t["status"] = "dormant"
        self.data["closed_actions"] = [
            a for a in self.data.get("closed_actions", [])
            if _days_since(a.get("closed_date", "")) <= self.CLOSED_RETENTION_DAYS
        ]

    # ── CONTEXT BLOCK ──────────────────────────────────────────────────────

    def build_context_block(self) -> str:
        if not self.data.get("run_count"):
            return ""

        lines    = []
        topics   = self.data.get("topics", {})
        people   = self.data.get("key_people", {})
        open_act = self.data.get("open_actions", [])

        lines += [
            "╔══════════════════════════════════════════════════════════╗",
            "  CUMULATIVE INBOX MEMORY  (built across previous runs)",
            f"  Last updated: {self.data.get('last_updated', 'n/a')}"
            f"  |  Total runs: {self.data.get('run_count', 0)}",
            "╚══════════════════════════════════════════════════════════╝",
            "",
        ]

        open_sorted = sorted(
            [a for a in open_act if a.get("status") in ("open", "pushed_to_todo")],
            key=lambda a: (0 if a.get("priority") == "high" else 1,
                           a.get("raised_date", "9999")))
        if open_sorted:
            lines += [
                "OPEN ACTIONS FROM PREVIOUS RUNS (not yet confirmed resolved):",
                "  Flag them if new emails relate to them; note if they appear resolved.",
                "",
            ]
            for a in open_sorted:
                age = _days_since(a.get("raised_date", _today()))
                pri = "🔴 HIGH  " if a.get("priority") == "high" else "        "
                lines.append(f"  {pri}[{a.get('id','')}] {a.get('title','')}"
                             f"  (raised {age}d ago, due {a.get('suggested_due','?')})")
                lines.append(f"          Topic: {a.get('topic','')}  |  "
                             f"From: {a.get('from_name','')}")
                if a.get("detail"):
                    lines.append(f"          {a['detail'][:200]}")
            lines.append("")

        active  = {k: v for k, v in topics.items() if v.get("status") == "active"}
        dormant = {k: v for k, v in topics.items() if v.get("status") == "dormant"}

        if active:
            lines.append("ACTIVE TOPICS (seen in recent runs):")
            for name, t in active.items():
                lines.append(f"  ▸ {name}  (first: {t.get('first_seen','?')}"
                             f", last: {t.get('last_seen','?')})")
                if t.get("summary"):
                    lines.append(f"    {t['summary']}")
                if t.get("open_actions"):
                    lines.append(f"    Open: {'; '.join(t['open_actions'][:4])}")
                if t.get("key_contacts"):
                    lines.append(f"    Contacts: {', '.join(t['key_contacts'][:3])}")
            lines.append("")

        if dormant:
            lines.append("DORMANT TOPICS (include if emails resurface them):")
            for name, t in dormant.items():
                lines.append(f"  · {name}  (last: {t.get('last_seen','?')})"
                             f"  —  {t.get('summary','')[:100]}")
            lines.append("")

        if people:
            lines.append("KEY PEOPLE (identified across previous runs):")
            for name, p in people.items():
                lines.append(f"  {name} <{p.get('email','')}>  —  "
                             f"{p.get('role','')}  —  {p.get('notes','')}")
            lines.append("")

        lines += ["END OF CUMULATIVE MEMORY", "─" * 60, ""]
        block = "\n".join(lines)
        if len(block) > self.MAX_STATE_CHARS:
            block = block[:self.MAX_STATE_CHARS] + "\n  … (state truncated)\n"
        return block

    # ── ACTION STATE TRANSITIONS ───────────────────────────────────────────

    def mark_pushed(self, action_titles: list):
        """Mark open actions as pushed_to_todo after a successful push (B1)."""
        title_set = {t.lower() for t in action_titles}
        for a in self.data.get("open_actions", []):
            if a.get("title", "").lower() in title_set:
                a["status"] = "pushed_to_todo"
        self.save()

    def mark_resolved(self, action_ids: list):
        """Move actions from open to closed by id."""
        today      = _today()
        id_set     = set(action_ids)
        still_open = []
        for a in self.data.get("open_actions", []):
            if a.get("id") in id_set:
                a["closed_date"] = today
                a["status"]      = "resolved"
                self.data.setdefault("closed_actions", []).append(a)
            else:
                still_open.append(a)
        self.data["open_actions"] = still_open
        self.save()

    def apply_patch(self, patch: dict):
        """Apply the state portion of the post-run patch."""
        if not patch:
            return
        today = _today()

        resolve_ids  = set(patch.get("resolve_action_ids", []))
        newly_closed, still_open = [], []
        for a in self.data["open_actions"]:
            if a.get("id") in resolve_ids:
                a["closed_date"] = today
                a["status"]      = "resolved"
                newly_closed.append(a)
            else:
                still_open.append(a)
        self.data["open_actions"]   = still_open
        self.data["closed_actions"] = self.data.get("closed_actions", []) + newly_closed

        existing_ids = {a.get("id") for a in self.data["open_actions"]}
        for a in patch.get("new_open_actions", []):
            if a.get("id") not in existing_ids:
                self.data["open_actions"].append(a)
                existing_ids.add(a.get("id"))

        for t in patch.get("topic_updates", []):
            name = t.get("topic", "")
            if not name:
                continue
            existing = self.data["topics"].get(name, {})
            if existing.get("first_seen") and not t.get("first_seen"):
                t["first_seen"] = existing["first_seen"]
            self.data["topics"][name] = t

        for p in patch.get("people_updates", []):
            name = p.get("name", "")
            if not name:
                continue
            existing = self.data["key_people"].get(name, {})
            if existing.get("first_seen") and not p.get("first_seen"):
                p["first_seen"] = existing["first_seen"]
            self.data["key_people"][name] = p


# ═════════════════════════════════════════════════════════════════════════════
# PILLAR HISTORY  (board_briefing.py only)
# ═════════════════════════════════════════════════════════════════════════════

def _default_pillars() -> dict:
    return {
        p["id"]: {
            "title":         p["title"],
            "last_rag":      None,
            "last_summary":  None,
            "last_run_date": None,
            "key_watch":     "Emails mentioning: " + ", ".join(p["keywords"][:12]),
        }
        for p in STRATEGIC_PILLARS
    }


class PillarHistory:
    """
    Remembers last month's RAG status and summary per strategic pillar so the
    briefing reports CHANGES, not cold facts. Matters/people/facts that used
    to live in board_briefing_context.json now go to CompanyKnowledge.
    """

    def __init__(self):
        self.path = PILLAR_HISTORY_FILE
        self.data = _load_json(self.path, {
            "last_updated":      "",
            "run_count":         0,
            "strategic_pillars": {},
        })
        # Deep-merge defaults so new pillars are never lost
        merged = _default_pillars()
        merged.update(self.data.get("strategic_pillars", {}))
        self.data["strategic_pillars"] = merged

    def save(self):
        _save_json(self.path, self.data)

    def increment_run_count(self):
        self.data["run_count"] = self.data.get("run_count", 0) + 1

    def is_empty(self) -> bool:
        return not any(p.get("last_summary")
                       for p in self.data["strategic_pillars"].values())

    def build_prompt_block(self, max_chars: int = 5000) -> str:
        pillars = self.data["strategic_pillars"]
        if self.is_empty():
            return ""
        lines = [
            "╔══════════════════════════════════════════════════════════╗",
            f"  STRATEGIC PILLAR STATUS FROM LAST RUN  (run #{self.data.get('run_count', 0)},"
            f" updated {self.data.get('last_updated', '?')})",
            "  Report what has CHANGED since this snapshot, not just current state.",
            "╚══════════════════════════════════════════════════════════╝",
            "",
        ]
        for pid, p in pillars.items():
            if not p.get("last_summary"):
                continue
            lines.append(f"  ▸ {p['title']}  "
                         f"[{p.get('last_rag', '?')} as of {p.get('last_run_date', '?')}]")
            lines.append(f"    {p['last_summary']}")
            if p.get("key_watch"):
                lines.append(f"    Watch for: {p['key_watch']}")
            lines.append("")
        lines += ["─" * 60, ""]
        block = "\n".join(lines)
        if len(block) > max_chars:
            block = block[:max_chars] + "\n  … (pillar history truncated)\n"
        return block

    def apply_patch(self, pillar_updates: list):
        """pillar_updates: [{id, last_rag, last_summary, key_watch}]"""
        today = _today()
        for pu in pillar_updates or []:
            pid = pu.get("id", "")
            if pid in self.data["strategic_pillars"]:
                p = self.data["strategic_pillars"][pid]
                if pu.get("last_rag"):
                    p["last_rag"] = pu["last_rag"]
                if pu.get("last_summary"):
                    p["last_summary"] = pu["last_summary"]
                if pu.get("key_watch"):
                    p["key_watch"] = pu["key_watch"]
                p["last_run_date"] = today

    def print_summary(self):
        print("\n" + "=" * 60)
        print("  PILLAR HISTORY")
        print("=" * 60)
        print(f"  File:         {self.path}")
        print(f"  Last updated: {self.data.get('last_updated', 'never')}")
        print(f"  Runs:         {self.data.get('run_count', 0)}")
        for pid, p in self.data["strategic_pillars"].items():
            rag  = p.get("last_rag") or "no data"
            date = p.get("last_run_date") or "never"
            print(f"\n  ▸ {p['title']}  [{rag} — {date}]")
            print(f"    {p.get('last_summary') or '(no history yet)'}")
        print()


# ═════════════════════════════════════════════════════════════════════════════
# CONTEXT BRIEFING  (human-editable standing instructions)
# ═════════════════════════════════════════════════════════════════════════════

BRIEFING_MAX_CHARS = 12000

def load_context_briefing(max_chars: int = BRIEFING_MAX_CHARS) -> str:
    """
    Read context_briefing.txt, skipping '#' comment lines.
    If oversized, keep the head AND the tail (auto-appends go at the bottom,
    so tail content must survive — B8). Previously the file was cut at the
    top 6,000 chars and appended items became permanently invisible.
    """
    if not CONTEXT_BRIEFING_FILE.exists():
        return ""
    lines = [ln.strip() for ln in
             CONTEXT_BRIEFING_FILE.read_text(encoding="utf-8").splitlines()
             if ln.strip() and not ln.strip().startswith("#")]
    text = "\n".join(lines)
    if len(text) > max_chars:
        head = text[: int(max_chars * 0.6)]
        tail = text[-int(max_chars * 0.35):]
        text = head + "\n  … (middle of briefing truncated) …\n" + tail
        print(f"  ⚠️  context_briefing.txt is large — consider consolidating it.")
    return text


def append_briefing(additions: str):
    """Append auto-discovered patterns with a dated header."""
    with open(CONTEXT_BRIEFING_FILE, "a", encoding="utf-8") as f:
        f.write(f"\n\n# ── Auto-appended {_today()} ──\n")
        f.write(additions.strip() + "\n")


def ensure_briefing_template():
    """Create a starter context_briefing.txt if none exists."""
    if CONTEXT_BRIEFING_FILE.exists():
        return
    CONTEXT_BRIEFING_FILE.write_text(
        "# context_briefing.txt — standing instructions for all tools\n"
        "# Edit freely. Lines starting with # are comments.\n"
        "# These instructions are injected into every Claude analysis prompt.\n"
        "#\n"
        "# Examples:\n"
        "#   - Ignore automated ASX notification emails and market announcements\n"
        "#   - Our litigation case numbers are ... (so emails quoting them are linked)\n"
        "#   - <Partner> is our preferred pipeline partner for <project>\n",
        encoding="utf-8")
    print(f"  📝  Created {CONTEXT_BRIEFING_FILE.name} — edit it to add standing instructions")
