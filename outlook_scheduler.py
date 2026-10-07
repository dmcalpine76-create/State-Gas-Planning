"""
outlook_scheduler.py  —  AI-powered diary scheduling engine
------------------------------------------------------------
Reads tasks from Microsoft To Do, finds free slots in Outlook calendar,
and uses Claude to schedule tasks according to the rules in scheduling_rules.json.

Now runs on the shared core/ package — same token cache, same Graph wrapper,
same Anthropic client and retry logic as inbox_actions.py. outlook_email.py is
no longer required. Its scheduling functions are imported directly by the
inbox_actions.py review dashboard, which runs task creation and diary
scheduling as one flow.

Callable standalone:
  py outlook_scheduler.py dashboard — plan, then open the scheduling dashboard
  py outlook_scheduler.py preview   — show proposed schedule, create nothing
  py outlook_scheduler.py flagged   — show yesterday's unfinished blocks

NOTE: calendar writes need the Calendars.ReadWrite scope. If your token cache
predates that being added to core/graph.py, delete .outlook_token_cache.bin
and run:  py inbox_actions.py setup
"""

import os
import json
import datetime
# requests is no longer used directly — all HTTP goes through core.graph
import re
from pathlib import Path

from core.env import load_env
load_env()

from core.config import (BASE_DIR, TASK_LIST_NAME, SCHEDULING_RULES_FILE,
                         SCHEDULING_PLAN_FILE, DURATION_HISTORY_FILE,
                         AEST_OFFSET_HOURS)
from core import graph
from core.llm import get_client, call_json, api_call_with_retry

# Kept only so existing guard clauses below stay truthy — the Anthropic client
# now comes from core.llm, which handles retries and JSON repair.
_ANTHROPIC_AVAILABLE = True
_OUTLOOK_AVAILABLE   = True

RULES_FILE   = SCHEDULING_RULES_FILE
PLAN_FILE    = SCHEDULING_PLAN_FILE
GRAPH_BASE   = graph.GRAPH_BASE
AEST_OFFSET  = datetime.timezone(datetime.timedelta(hours=AEST_OFFSET_HOURS))
BLOCK_TAG    = "🎯"          # used to identify scheduler-created events

BLOCK_CATEGORY = "Scheduled Task"  # every scheduler-created event carries this

# Icon in front of a diary block, chosen from the task's title and notes.
# First match wins, so the more specific groups come first. Override or extend
# with "calendar_blocks": {"icons": [{"icon": "⚖️", "keywords": ["court"]}]}
# in scheduling_rules.json — those are checked before the defaults.
DEFAULT_BLOCK_ICONS = [
    ("⚖️", ["litigation", "court", "lawyer", "legal", "counsel", "hopgood", " h&g", "h & g",
            "mcdonald", "wildhorse", "affidavit", "mediation", "expert report", "insurer",
            "claim", "subpoena"]),
    ("🏛️", ["board", "director", "minutes", "chair", "agm", "annual report", "governance",
            "audit", "remuneration", "constitution"]),
    ("📋", ["asx", "quarterly", "appendix 5b", "5b", "disclosure", "accc", "compliance",
            "lodge", "lodgement", "relinquish", "tenure", "permit", "atp ", "pl application",
            "renewal", "department", "regulator", "ato", "tax", "r&d", "r & d", "royalt"]),
    ("💰", ["capital", "raise", "raising", "funding", "investor", "broker", "alpine",
            "saffron", "fensom", "ignite", "convertible", "cash flow", "cashflow",
            "budget", "invoice", "payment", "loan", "equity", "placement", "bank"]),
    ("🔥", ["hdng", "leip", "treasury", "jellinbah", "foundation customer", "x-cell",
            "cradle", "rapid refuel", "truck", "diesel", "mes ", "conversion kit",
            "cylinder", "hot tap", "qpm", "blue energy"]),
    ("⛽", ["rolleston", "concept study", "pre-feed", "feed", "jemena", "qgp", "pipeline",
            "reserve", "nsai", "farm-in", "farm in", "farmout", "farm-out", "beach",
            "senex", "elementa", "santos", "glng", "reid's dome", "pl231", "wells", "drilling"]),
    ("📞", ["call ", "call with", "phone", "meeting", "meet ", "catch up", "catch-up",
            "coffee", "visit", "workshop", "session"]),
    ("✉️", ["reply", "respond", "email", "send ", "follow up", "follow-up", "chase"]),
]
import re
_KNOWN_ICONS = {BLOCK_TAG} | {i for i, _ in DEFAULT_BLOCK_ICONS}


def block_icon(title: str, body: str = "", cal_blocks: dict | None = None) -> str:
    """Pick the icon for a diary block; falls back to the configured prefix (🎯)."""
    cal_blocks = cal_blocks or {}
    text = f"{title} {body}".lower()
    groups = [(g.get("icon", ""), g.get("keywords", []))
              for g in cal_blocks.get("icons", []) if g.get("icon")]
    for icon, words in groups + DEFAULT_BLOCK_ICONS:
        for w in words:
            w = w.strip().lower()
            if w and re.search(r"(?<![a-z0-9])" + re.escape(w) + r"(?![a-z0-9])", text):
                return icon
    return (cal_blocks.get("block_prefix", BLOCK_TAG + " ").strip() or BLOCK_TAG)


def block_title(title: str, body: str = "", cal_blocks: dict | None = None) -> str:
    """Event subject for a diary block: '<icon> <task title>'."""
    title = strip_block_icon(title)
    return f"{block_icon(title, body, cal_blocks)} {title}"


def strip_block_icon(subject: str) -> str:
    """Remove a leading block icon (any of the known ones, or a custom one)."""
    s = (subject or "").strip()
    for icon in sorted(_KNOWN_ICONS, key=len, reverse=True):
        if s.startswith(icon):
            return s[len(icon):].lstrip("\ufe0f ").strip()
    return s


def is_block_event(subject: str, categories=None) -> bool:
    """True for events this scheduler created (category, or a known icon prefix)."""
    if BLOCK_CATEGORY in (categories or []):
        return True
    s = (subject or "").strip()
    return any(s.startswith(i) for i in _KNOWN_ICONS)
REQUEST_TIMEOUT = graph.REQUEST_TIMEOUT
DAYFMT       = "%#d" if os.name == "nt" else "%-d"   # no-pad day: Windows vs Linux

# Scheduling is a cheap, high-volume call — keep it on Haiku rather than the
# core default model used for inbox extraction.
SCHEDULER_MODEL = "claude-haiku-4-5-20251001"


# ── AUTH ─────────────────────────────────────────────────────────────────────
# Thin wrappers over core.graph so the ~40 existing call sites below are
# unchanged, but there is now exactly one token cache and one HTTP layer.

def _get_token() -> str:
    return graph.get_access_token()


def _graph_get(token: str, path: str, params: dict = None) -> dict:
    return graph.graph_get(token, path, params=params)


def _graph_post(token: str, path: str, body: dict) -> dict:
    return graph.graph_post(token, path, body)


def _graph_patch(token: str, path: str, body: dict) -> dict:
    return graph.graph_patch(token, path, body)


# ── RULES ─────────────────────────────────────────────────────────────────────

def load_rules() -> dict:
    """Load scheduling_rules.json. Returns empty dict if not found."""
    if not RULES_FILE.exists():
        return {}
    try:
        return json.loads(RULES_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"  ⚠️  Could not load scheduling_rules.json: {e}")
        return {}


def get_day_rules(rules: dict, dt: datetime.datetime) -> dict:
    """
    Return the rules for a specific date, accounting for the fortnightly cycle.
    Week A / Week B is determined by ISO week number (even = A, odd = B).
    Returns a day dict with: enabled, start_time, end_time, notes.
    """
    day_key = dt.strftime("%a").lower()  # mon, tue, etc.
    iso_week = dt.isocalendar()[1]
    week_key = "a" if iso_week % 2 == 0 else "b"

    weeks = rules.get("weeks", {})
    week  = weeks.get(week_key, {})
    day   = week.get(day_key, {})

    global_rules = rules.get("global", {})

    return {
        "enabled":    day.get("enabled", True),
        "start_time": day.get("start_time", global_rules.get("global_start", "09:00")),
        "end_time":   day.get("end_time",   global_rules.get("global_end",   "16:00")),
        "notes":      day.get("notes", ""),
        "week":       week_key.upper(),
        "day_key":    day_key,
    }


# ── FETCH TASKS FROM TO DO ────────────────────────────────────────────────────

def get_target_list_name(rules: dict = None) -> str:
    """
    The To Do list this scheduler treats as its work queue. Defined once in
    core/config.py so inbox_actions.py pushes approved actions into exactly
    the list this reads from. scheduling_rules.json may override it, but the
    rules dashboard doesn't write that key, so config is the practical source.
    """
    if rules:
        name = (rules.get("global", {}) or {}).get("task_list", "")
        if isinstance(name, str) and name.strip():
            return name.strip()
    return TASK_LIST_NAME


def fetch_todo_tasks(token: str, rules: dict, lookback_days: int = 2) -> list[dict]:
    """
    Fetch tasks from:
      1. The scheduler's source list (global.task_list, default 'Daily Priorities')
      2. Any other list with tasks due today or overdue within lookback_days

    Tasks due TODAY are now included, not just ones already late — previously a
    task created this morning by inbox_actions.py was neither in the source list
    nor overdue, so it was invisible to the scheduler until it slipped.

    lookback_days: how far back to look for overdue tasks in other lists.
    Returns list of task dicts with: id, title, body, due_date, priority,
    list_name, is_overdue.
    """
    today     = datetime.date.today()
    lookback_cutoff = today - datetime.timedelta(days=lookback_days)
    tasks = []
    seen_ids = set()

    try:
        lists_data = _graph_get(token, "/me/todo/lists")
        all_lists  = lists_data.get("value", [])
    except Exception as e:
        print(f"  ⚠️  Could not fetch To Do lists: {e}")
        return []

    target_list_name = get_target_list_name(rules)

    for lst in all_lists:
        list_id   = lst["id"]
        list_name = lst.get("displayName", "")
        is_priority = list_name.lower() == target_list_name.lower()

        # Pull open tasks from every list; filtering happens per task below.
        try:
            tasks_data = _graph_get(
                token,
                f"/me/todo/lists/{list_id}/tasks",
                {"$filter": "status ne 'completed'", "$top": 100}
            )
        except Exception:
            continue

        for t in tasks_data.get("value", []):
            tid = t.get("id", "")
            if tid in seen_ids:
                continue

            # Parse due date
            due_raw = t.get("dueDateTime", {})
            due_date = None
            if due_raw and due_raw.get("dateTime"):
                try:
                    due_date = datetime.date.fromisoformat(
                        due_raw["dateTime"][:10]
                    )
                except Exception:
                    pass

            is_overdue = bool(due_date and lookback_cutoff <= due_date < today)
            is_due_today = bool(due_date and due_date == today)

            # Include if: in the scheduler's source list, OR due today in any
            # list, OR overdue in any list.
            if not is_priority and not is_due_today and not is_overdue:
                continue

            seen_ids.add(tid)

            # Get body text
            body_content = t.get("body", {}).get("content", "").strip()
            # Strip HTML tags if present
            body_clean = re.sub(r"<[^>]+>", " ", body_content).strip()

            importance = t.get("importance", "normal")
            priority   = "high" if importance == "high" else (
                         "urgent" if is_overdue else "normal")

            tasks.append({
                "id":         tid,
                "title":      t.get("title", ""),
                "body":       body_clean[:400] if body_clean else "",
                "due_date":   due_date.isoformat() if due_date else None,
                "priority":   priority,
                "list_name":  list_name,
                "is_overdue": is_overdue,
                "list_id":    list_id,
                "raw":        t,
            })

    # Sort: overdue first, then by due date, then by importance
    tasks.sort(key=lambda t: (
        0 if t["is_overdue"] else 1,
        t["due_date"] or "9999-99-99",
        0 if t["priority"] == "urgent" else (1 if t["priority"] == "high" else 2)
    ))

    return tasks


# ── FETCH CALENDAR EVENTS ─────────────────────────────────────────────────────

def fetch_upcoming_events(token: str, days: int = 7) -> list[dict]:
    """
    Fetch calendar events for the next `days` days (and yesterday for flagging).
    Returns list of event dicts with: subject, start_dt, end_dt, is_scheduler_block
    """
    now_aest    = datetime.datetime.now(AEST_OFFSET)
    yest_start  = (now_aest - datetime.timedelta(days=1)).replace(
        hour=0, minute=0, second=0, microsecond=0)
    window_end  = now_aest.replace(
        hour=0, minute=0, second=0, microsecond=0) + datetime.timedelta(days=days)

    start_utc = yest_start.astimezone(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    end_utc   = window_end.astimezone(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    try:
        data = graph.graph_get(
            token, "/me/calendarView",
            params={
                "startDateTime": start_utc,
                "endDateTime":   end_utc,
                "$select": "subject,start,end,isAllDay,isCancelled,showAs,body,categories",
                "$orderby": "start/dateTime",
                "$top": 200,
            },
            headers={"Prefer": 'outlook.timezone="AUS Eastern Standard Time"'},
        )
    except Exception as e:
        print(f"  ⚠️  Could not fetch calendar events: {e}")
        if "403" in str(e) or "Forbidden" in str(e):
            print("      This usually means the cached token predates the "
                  "Calendars.ReadWrite scope.\n"
                  "      Delete .outlook_token_cache.bin and run: "
                  "py inbox_actions.py setup")
        return []

    events = []
    for ev in data.get("value", []):
        if ev.get("isCancelled"):
            continue
        if ev.get("isAllDay"):
            continue

        raw_start = ev["start"].get("dateTime", "")
        raw_end   = ev["end"].get("dateTime", "")

        def _parse(s):
            s = s.split(".")[0]
            if s.endswith("Z"):
                return datetime.datetime.fromisoformat(
                    s[:-1] + "+00:00").astimezone(AEST_OFFSET)
            elif "+" in s[10:] or s.count("-") > 2:
                return datetime.datetime.fromisoformat(s).astimezone(AEST_OFFSET)
            else:
                return datetime.datetime.fromisoformat(s).replace(tzinfo=AEST_OFFSET)

        try:
            start_dt = _parse(raw_start)
            end_dt   = _parse(raw_end)
        except Exception:
            continue

        subject = ev.get("subject", "")
        is_scheduler = is_block_event(subject, ev.get("categories"))

        events.append({
            "subject":          subject,
            "start_dt":         start_dt,
            "end_dt":           end_dt,
            "is_scheduler_block": is_scheduler,
            "event_id":         ev.get("id", ""),
        })

    return events


# ── FIND FREE SLOTS ───────────────────────────────────────────────────────────

def find_free_slots(events: list[dict], rules: dict, days_ahead: int = 5) -> list[dict]:
    """
    For each schedulable day in the next days_ahead days, find free time slots
    respecting the rules (working hours, deep work window, buffers, protected times).

    Returns list of slot dicts:
      { date, start_dt, end_dt, duration_mins, slot_type, week }
      slot_type: 'deep_work' | 'afternoon'
    """
    global_rules = rules.get("global", {})
    buffer_mins  = global_rules.get("buffer_after", 15)
    snap_mins    = global_rules.get("snap_to", 30)
    min_block    = global_rules.get("min_block", 30)
    deep_start   = global_rules.get("deep_start", "09:00")
    deep_end     = global_rules.get("deep_end", "12:00")
    protected    = rules.get("protected_times", [])
    work_only    = global_rules.get("work_days_only", True)

    now_aest = datetime.datetime.now(AEST_OFFSET)
    slots    = []

    for day_offset in range(0, days_ahead):
        date = (now_aest + datetime.timedelta(days=day_offset)).date()
        dt   = datetime.datetime.combine(date, datetime.time(0, 0), tzinfo=AEST_OFFSET)

        day_rules = get_day_rules(rules, dt)
        if not day_rules["enabled"]:
            continue

        # Skip weekends for work tasks if configured
        if work_only and dt.weekday() >= 5:
            continue

        # Parse day boundaries
        day_start = _parse_time_on_date(date, day_rules["start_time"])
        day_end   = _parse_time_on_date(date, day_rules["end_time"])
        dw_start  = _parse_time_on_date(date, deep_start)
        dw_end    = _parse_time_on_date(date, deep_end)

        # Don't schedule in the past
        if day_end < now_aest:
            continue
        if day_start < now_aest:
            day_start = _snap_forward(now_aest, snap_mins)

        # Build list of blocked periods for this day
        blocked = []

        # From actual calendar events — including previously-created 🎯 scheduler
        # blocks. Once a block has been pushed to the real calendar it's a genuine
        # commitment, so it must block new bookings the same as any other event
        # (otherwise a later run/dashboard session can schedule right on top of it).
        for ev in events:
            if ev["start_dt"].date() != date:
                continue
            buf_end = ev["end_dt"] + datetime.timedelta(minutes=buffer_mins)
            blocked.append((ev["start_dt"], buf_end))

        # From protected times
        day_name = dt.strftime("%a").lower()
        for pt in protected:
            if day_name not in (pt.get("days") or []):
                continue
            if not pt.get("start") or not pt.get("end"):
                continue
            try:
                pt_start = _parse_time_on_date(date, pt["start"])
                pt_end   = _parse_time_on_date(date, pt["end"])
                blocked.append((pt_start, pt_end))
            except Exception:
                continue

        # Sort and merge blocked periods
        blocked.sort(key=lambda x: x[0])
        merged = []
        for b in blocked:
            if merged and b[0] <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], b[1]))
            else:
                merged.append(list(b))

        # Find free slots in deep work window and afternoon window
        for window_start, window_end_t, slot_type in [
            (dw_start, dw_end, "deep_work"),
            (dw_end,   day_end, "afternoon"),
        ]:
            cursor = max(window_start, day_start)
            if cursor < now_aest:
                cursor = _snap_forward(now_aest, snap_mins)

            for block_start, block_end in merged + [(window_end_t, window_end_t)]:
                # Free time between cursor and block_start
                free_end = min(block_start, window_end_t)
                if free_end > cursor:
                    # Snap cursor to hour/half-hour
                    snapped = _snap_forward(cursor, snap_mins)
                    if snapped < free_end:
                        duration = int((free_end - snapped).total_seconds() / 60)
                        if duration >= min_block:
                            slots.append({
                                "date":         date.isoformat(),
                                "start_dt":     snapped,
                                "end_dt":       free_end,
                                "duration_mins": duration,
                                "slot_type":    slot_type,
                                "week":         day_rules["week"],
                            })
                cursor = max(cursor, block_end)

    return slots


def _parse_time_on_date(date: datetime.date, time_str: str) -> datetime.datetime:
    h, m = map(int, time_str.split(":"))
    return datetime.datetime.combine(
        date, datetime.time(h, m), tzinfo=AEST_OFFSET)


def _snap_forward(dt: datetime.datetime, snap_mins: int) -> datetime.datetime:
    """
    Snap datetime forward to the next snap_mins boundary.
    Uses timedelta arithmetic so a snap past midnight rolls into the next
    day instead of crashing on hour=24 (e.g. running the scheduler at 23:45).
    """
    base = dt.replace(hour=0, minute=0, second=0, microsecond=0)
    mins = dt.hour * 60 + dt.minute
    if dt.second or dt.microsecond:
        mins += 1   # never snap backwards past a partial minute
    snapped = ((mins + snap_mins - 1) // snap_mins) * snap_mins
    return base + datetime.timedelta(minutes=snapped)


# ── DURATION LEARNING (D8) ────────────────────────────────────────────────────
# Logs every block the scheduler creates, marks blocks that were flagged as
# unfinished the next morning, and feeds a short calibration summary back
# into the AI scheduling prompt so duration estimates improve over time.

DURATION_HISTORY_FILE = Path(__file__).parent / "duration_history.json"


def _load_duration_history() -> list[dict]:
    if not DURATION_HISTORY_FILE.exists():
        return []
    try:
        return json.loads(DURATION_HISTORY_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []


def _save_duration_history(entries: list[dict]) -> None:
    try:
        DURATION_HISTORY_FILE.write_text(
            json.dumps(entries[-100:], indent=1, ensure_ascii=False),
            encoding="utf-8")
    except Exception as e:
        print(f"  ⚠️  Could not save duration history: {e}")


def _log_scheduled_duration(title: str, mins: int) -> None:
    """Record a block the moment it's pushed to the calendar."""
    entries = _load_duration_history()
    entries.append({
        "date":    datetime.date.today().isoformat(),
        "title":   strip_block_icon(title)[:80],
        "mins":    int(mins),
        "outcome": "scheduled",
    })
    _save_duration_history(entries)


def _record_flagged_outcomes(flagged: list[dict]) -> None:
    """Mark yesterday's unfinished blocks so the AI learns they ran over."""
    if not flagged:
        return
    entries = _load_duration_history()
    if not entries:
        return
    changed = False
    for ev in flagged:
        title = strip_block_icon(ev["subject"])[:80]
        for entry in reversed(entries):
            if entry["title"] == title and entry.get("outcome") == "scheduled":
                entry["outcome"] = "unfinished"
                changed = True
                break
    if changed:
        _save_duration_history(entries)


def _duration_hints(max_lines: int = 12) -> str:
    """Compact calibration summary of recent history for the AI prompt."""
    entries = _load_duration_history()
    if not entries:
        return ""
    lines = []
    for e in entries[-max_lines:]:
        outcome = "was NOT finished in the allocated time" \
            if e.get("outcome") == "unfinished" else "was completed"
        lines.append(f"  - \"{e['title']}\" allocated {e['mins']}min — {outcome}")
    return (
        "\nRECENT DURATION HISTORY (use this to calibrate your estimates — "
        "if similar tasks ran over, allocate MORE time):\n" + "\n".join(lines) + "\n"
    )


# ── AI SCHEDULING ─────────────────────────────────────────────────────────────

def _resolve_overlaps(candidates: list[dict]) -> list[dict]:
    """
    Guard against the AI assigning two different tasks into overlapping time
    ranges (e.g. picking the same slot twice, or two slots that touch).
    Keeps items in the order the AI returned them — earlier items (its own
    priority order) win the time; anything that would overlap an already-
    accepted item is dropped here rather than silently double-booked. Dropped
    tasks aren't lost — they come back as unscheduled and get picked up by
    the next (wider) scheduling pass in schedule_with_retry().
    """
    accepted = []
    for item in candidates:
        overlap = any(
            item["start_dt"] < a["end_dt"] and item["end_dt"] > a["start_dt"]
            for a in accepted
        )
        if overlap:
            continue
        accepted.append(item)
    return accepted


def _subtract_booked(slots: list[dict], booked_ranges: list[tuple],
                      min_block: int = 30) -> list[dict]:
    """
    Trim or drop the portion of each free slot that overlaps a range already
    booked earlier in this scheduling run. Needed because widening the search
    window re-derives slots from the real calendar only — it has no idea
    about tasks this same run already placed a moment ago — so without this,
    a second (wider) pass could offer the same minutes to a different task.
    """
    if not booked_ranges:
        return slots
    out = []
    for slot in slots:
        pieces = [(slot["start_dt"], slot["end_dt"])]
        for b_start, b_end in booked_ranges:
            next_pieces = []
            for p_start, p_end in pieces:
                if b_end <= p_start or b_start >= p_end:
                    next_pieces.append((p_start, p_end))   # no overlap
                    continue
                if b_start > p_start:
                    next_pieces.append((p_start, b_start))  # slice before
                if b_end < p_end:
                    next_pieces.append((b_end, p_end))      # slice after
            pieces = next_pieces
        for p_start, p_end in pieces:
            dur = int((p_end - p_start).total_seconds() / 60)
            if dur >= min_block:
                trimmed = dict(slot)
                trimmed["start_dt"]      = p_start
                trimmed["end_dt"]        = p_end
                trimmed["duration_mins"] = dur
                out.append(trimmed)
    return out


def schedule_with_retry(
    tasks: list[dict],
    events: list[dict],
    rules: dict,
    api_key: str = None,
    scheduler_fn=None,
    initial_days: int = 5,
    max_days: int = 60,
    verbose: bool = False,
) -> tuple[list[dict], list[dict]]:
    """
    Place tasks in three passes, never booking beyond the coming week:

      1. Due date first — each task is offered only the free slots on its
         To Do due date (overdue tasks: today), when that date falls inside
         the scheduling window.
      2. Best fit in the window — anything not placed by pass 1 (including
         tasks with no due date, or a due date beyond the window) is placed
         wherever it fits best within the next `global.schedule_window_days`
         days (default 7).
      3. Weekend holding slots — whatever is still left is parked on the
         coming Saturday and Sunday as "to reallocate" blocks, for Doug to
         drag into next week by hand.

    `initial_days` / `max_days` are accepted for compatibility but no longer
    widen the search. Returns (scheduled, unscheduled); unscheduled is only
    non-empty if the weekend itself is full.
    """
    scheduler_fn = scheduler_fn or ai_schedule_tasks
    g          = rules.get("global", {})
    min_block  = g.get("min_block", 30)
    window     = int(g.get("schedule_window_days", 7))
    today      = datetime.datetime.now(AEST_OFFSET).date()
    last_day   = today + datetime.timedelta(days=window - 1)
    remaining  = list(tasks)
    scheduled  = []
    booked     = []

    def _place(batch, slots):
        nonlocal remaining
        slots = _subtract_booked(slots, booked, min_block=min_block)
        if not batch or not slots:
            return
        newly = scheduler_fn(batch, slots, rules, api_key) or []
        newly = [n for n in newly
                 if not any(n["start_dt"] < e and n["end_dt"] > s for s, e in booked)]
        scheduled.extend(newly)
        booked.extend((n["start_dt"], n["end_dt"]) for n in newly)
        placed = {n["task"]["id"] for n in newly}
        remaining = [t for t in remaining if t["id"] not in placed]

    all_slots = find_free_slots(events, rules, days_ahead=window)

    # Pass 1 — on the due date
    by_date = {}
    for t in remaining:
        due = (t.get("due_date") or "")[:10]
        if not due:
            continue
        try:
            d = datetime.date.fromisoformat(due)
        except ValueError:
            continue
        d = max(d, today)                       # overdue → today
        if d <= last_day:
            by_date.setdefault(d.isoformat(), []).append(t)
    for day in sorted(by_date):
        _place(by_date[day], [s for s in all_slots if s["date"] == day])
    if verbose and by_date:
        print(f"   ↳ due-date pass: {len(scheduled)} placed, {len(remaining)} left")

    # Pass 2 — best fit anywhere in the window
    if remaining:
        _place(list(remaining), all_slots)
        if verbose:
            print(f"   ↳ week pass: {len(scheduled)} placed, {len(remaining)} left")

    # Pass 3 — park the rest on the weekend
    unscheduled = []
    if remaining:
        parked, unscheduled = _park_on_weekend(remaining, events, booked, rules)
        scheduled.extend(parked)
        if verbose:
            print(f"   ↳ {len(parked)} parked on the weekend to reallocate")

    return scheduled, unscheduled


def _park_on_weekend(tasks: list[dict], events: list[dict], booked: list[tuple],
                     rules: dict) -> tuple[list[dict], list[dict]]:
    """
    Put tasks the week couldn't fit onto the coming Saturday and Sunday,
    back to back from 09:00, avoiding anything already in the diary. They're
    holding slots, not commitments: marked free and titled to be reallocated.
    """
    g          = rules.get("global", {})
    cal_blocks = rules.get("calendar_blocks", {})
    start_t    = g.get("weekend_hold_start", "09:00")
    end_t      = g.get("weekend_hold_end", "17:00")
    now        = datetime.datetime.now(AEST_OFFSET)
    today      = now.date()
    sat        = today + datetime.timedelta(days=(5 - today.weekday()) % 7)
    days       = [d for d in (sat, sat + datetime.timedelta(days=1)) if d >= today]
    if today.weekday() == 6:                      # run on a Sunday → today only
        days = [today]

    busy = [(e["start_dt"], e["end_dt"]) for e in events] + list(booked)
    parked, left = [], []
    queue = list(tasks)
    for day in days:
        cursor = _parse_time_on_date(day, start_t)
        if cursor < now:
            cursor = _snap_forward(now, 30)
        day_end = _parse_time_on_date(day, end_t)
        while queue:
            t    = queue[0]
            mins = int(t.get("user_duration_mins") or g.get("min_block", 30) or 30)
            mins = max(15, round(mins / 15) * 15)
            end  = cursor + datetime.timedelta(minutes=mins)
            clash = [b for b in busy if cursor < b[1] and end > b[0]]
            if clash:
                cursor = max(b[1] for b in clash)
                continue
            if end > day_end:
                break
            body = (_build_event_body(t, cal_blocks)
                    + "\n\nCould not be fitted into the coming week — move this to a "
                      "time next week.")
            parked.append({
                "task":           t,
                "slot":           None,
                "estimated_mins": mins,
                "start_dt":       cursor,
                "end_dt":         end,
                "title":          block_title(f"TO REALLOCATE – {t['title']}",
                                              t.get("body", ""), cal_blocks),
                "description":    body,
                "reason":         "No room this week — parked on the weekend to reallocate",
                "holding":        True,
            })
            busy.append((cursor, end))
            cursor = end
            queue.pop(0)
    for t in queue:
        left.append({"task": t, "reason": "No room this week and the weekend holding area is full"})
    return parked, left


def ai_schedule_tasks(
    tasks: list[dict],
    slots: list[dict],
    rules: dict,
    api_key: str = None,
) -> list[dict]:
    """
    Use Claude to match tasks to slots, estimating durations and respecting rules.

    Returns list of scheduled items:
      { task, slot, estimated_mins, title, description }
    """
    if not _ANTHROPIC_AVAILABLE or not tasks or not slots:
        return []

    global_rules = rules.get("global", {})
    task_types   = rules.get("task_types", [])
    scheduling   = rules.get("scheduling", {})
    cal_blocks   = rules.get("calendar_blocks", {})

    max_hours_day = global_rules.get("max_hours", 3)
    max_blocks    = global_rules.get("max_blocks", 4)
    extra_instr   = scheduling.get("extra_instructions", "")

    # Build task type reference for prompt
    tt_lines = "\n".join(
        f"  - {tt['category']}: keywords={tt['keywords']}, default={tt['duration']}min, "
        f"preferred={tt['time_pref']}, type={tt['type']}"
        for tt in task_types
    )

    # Build task list for prompt
    task_lines = "\n".join(
        f"[{i+1}] {t['title']}"
        + (f" (OVERDUE — due {t['due_date']})" if t['is_overdue'] else
           f" (due {t['due_date']})" if t['due_date'] else "")
        + (f"\n    Priority: {t['priority']}" if t['priority'] != 'normal' else "")
        + (f"\n    Detail: {t['body']}" if t['body'] else "")
        + f"\n    List: {t['list_name']}"
        for i, t in enumerate(tasks[:20])  # cap at 20 tasks
    )

    # Build slot list for prompt
    slot_lines = "\n".join(
        f"[S{i+1}] {s['date']} {s['start_dt'].strftime('%H:%M')}–{s['end_dt'].strftime('%H:%M')} "
        f"({s['duration_mins']}min available, {s['slot_type'].replace('_',' ')}, Week {s['week']})"
        for i, s in enumerate(slots[:40])
    )

    prompt = f"""You are a professional diary scheduler for Doug McAlpine, State Gas, Brisbane (AEST).

SCHEDULING RULES:
- Hard start: 09:00, Hard stop: 16:00 every day
- Deep work window (contracts, documents, financial, regulatory): 09:00–12:00
- Admin/calls/emails window: 12:00–16:00
- Max {max_hours_day} hours of task blocks per day, max {max_blocks} blocks per day
- Snap all tasks to start on the hour or half-hour
- Leave 15 min buffer after each task block
- Overdue tasks must be scheduled TODAY or TOMORROW at latest
- High priority tasks before normal priority
{f'- {extra_instr}' if extra_instr else ''}

TASK TYPE DURATION GUIDE:
{tt_lines}
{_duration_hints()}
AVAILABLE SLOTS (use slot index S1, S2 etc):
{slot_lines}

TASKS TO SCHEDULE (use task index 1, 2 etc):
{task_lines}

For each task:
1. Estimate duration in minutes (use task type guide, use your judgement based on description)
2. Pick the best slot (match deep_work tasks to deep work slots, admin to afternoon slots)
3. If a slot is too small for a task, split the task across two slots only if it makes sense
4. Don't exceed the slot's available duration
5. Don't schedule more than {max_hours_day}h total per day across all tasks

Respond ONLY as JSON — no markdown, no preamble:
{{
  "schedule": [
    {{
      "task_index": 1,
      "slot_index": "S1",
      "estimated_mins": 60,
      "reason": "one sentence why this slot fits this task"
    }}
  ],
  "unscheduled": [
    {{
      "task_index": 2,
      "reason": "why it couldn't be scheduled"
    }}
  ]
}}"""

    data = call_json(get_client(), prompt, max_tokens=2000,
                     model=SCHEDULER_MODEL, label="AI scheduling")
    if not isinstance(data, dict):
        return []

    # Map indices back to actual tasks and slots
    prefix  = cal_blocks.get("block_prefix", "🎯 ")
    results = []

    for item in data.get("schedule", []):
        try:
            t_idx = int(item["task_index"]) - 1
            s_key = item["slot_index"]          # e.g. "S3"
            s_idx = int(s_key[1:]) - 1
            est   = int(item["estimated_mins"])
            task  = tasks[t_idx]
            slot  = slots[s_idx]

            # Clamp estimated duration to slot availability
            est = min(est, slot["duration_mins"])
            # Snap to nearest 15 mins
            est = max(15, round(est / 15) * 15)

            end_dt = slot["start_dt"] + datetime.timedelta(minutes=est)

            results.append({
                "task":           task,
                "slot":           slot,
                "estimated_mins": est,
                "start_dt":       slot["start_dt"],
                "end_dt":         end_dt,
                "title":          block_title(task['title'], task.get("body", ""), cal_blocks),
                "description":    _build_event_body(task, cal_blocks),
                "reason":         item.get("reason", ""),
            })
        except (IndexError, KeyError, ValueError):
            continue

    # Guard against the AI assigning two tasks into overlapping time ranges
    return _resolve_overlaps(results)


def _build_event_body(task: dict, cal_blocks: dict) -> str:
    parts = []
    if task.get("body"):
        parts.append(task["body"])
    if task.get("is_overdue"):
        parts.append(f"⚠️ OVERDUE (was due {task['due_date']})")
    elif task.get("due_date"):
        parts.append(f"Due: {task['due_date']}")
    parts.append(f"From To Do list: {task['list_name']}")
    if cal_blocks.get("add_todo_link", True):
        parts.append(f"\nView in To Do: https://to-do.office.com/tasks/id/{task['id']}/details")
    return "\n".join(parts)


# ── CREATE CALENDAR EVENTS ────────────────────────────────────────────────────

def create_calendar_events(token: str, scheduled: list[dict], rules: dict) -> tuple[int, int]:
    """
    Create calendar events in Outlook for each scheduled item.
    Returns (ok_count, fail_count).
    """
    global_rules = rules.get("global", {})
    cal_blocks   = rules.get("calendar_blocks", {})

    show_as     = cal_blocks.get("block_show_as", "free")
    reminder    = cal_blocks.get("block_reminder", 10)
    mark_free   = global_rules.get("mark_as_free", True)
    if mark_free:
        show_as = "free"

    ok, fail = 0, 0

    for item in scheduled:
        try:
            start_str = item["start_dt"].astimezone(datetime.timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%S")
            end_str   = item["end_dt"].astimezone(datetime.timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%S")

            body = {
                "subject": item["title"],
                "body": {
                    "contentType": "text",
                    "content": item["description"],
                },
                "start": {"dateTime": start_str, "timeZone": "UTC"},
                "end":   {"dateTime": end_str,   "timeZone": "UTC"},
                "showAs": show_as,
                "isReminderOn": reminder > 0,
                "reminderMinutesBeforeStart": reminder,
                "categories": [BLOCK_CATEGORY],
            }

            _graph_post(token, "/me/events", body)
            _log_scheduled_duration(item["title"], item["estimated_mins"])
            ok += 1
            print(f"   ✓ Created: {item['title'][:60]} @ "
                  f"{item['start_dt'].strftime('%a %d %b %H:%M')}")
        except Exception as e:
            fail += 1
            print(f"   ✗ Failed: {item['task']['title'][:50]}: {e}")

    return ok, fail


# ── FLAG UNFINISHED BLOCKS ────────────────────────────────────────────────────

def get_flagged_blocks(events: list[dict]) -> list[dict]:
    """
    Return yesterday's scheduler-created blocks (potential unfinished tasks).
    """
    yesterday = (datetime.datetime.now(AEST_OFFSET) - datetime.timedelta(days=1)).date()
    return [
        ev for ev in events
        if ev["is_scheduler_block"] and ev["start_dt"].date() == yesterday
    ]


# ── HTML SUMMARY FOR BRIEFING ─────────────────────────────────────────────────

def build_schedule_summary_html(
    scheduled: list[dict],
    unscheduled: list[dict],
    flagged: list[dict],
) -> str:
    """
    Build a compact HTML summary to embed in the morning briefing.
    Shows: today's scheduled blocks, flagged yesterday items, unscheduled tasks.
    """
    if not scheduled and not flagged:
        return ""

    today = datetime.date.today()

    # Today's blocks
    today_items = [s for s in scheduled if s["start_dt"].date() == today]
    future_items = [s for s in scheduled if s["start_dt"].date() > today]

    html_parts = []

    if today_items:
        cards = ""
        for s in today_items:
            t = s["task"]
            time_str = f"{s['start_dt'].strftime('%I:%M %p').lstrip('0')} – {s['end_dt'].strftime('%I:%M %p').lstrip('0')}"
            overdue_badge = '<span style="font-size:0.58rem;font-weight:700;background:#fde8e8;color:#c0392b;padding:0.1rem 0.35rem;border-radius:2px;margin-left:0.35rem">⚠️ OVERDUE</span>' if t.get("is_overdue") else ""
            cards += f"""<div style="padding:0.7rem 1rem;border-bottom:1px solid var(--rule)">
  <div style="font-size:0.65rem;font-weight:700;color:var(--ink-light);text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.2rem">{time_str} · {s['estimated_mins']}min</div>
  <div style="font-size:0.85rem;font-weight:700;font-family:var(--font-display)">{t['title']}{overdue_badge}</div>
  {f'<div style="font-size:0.73rem;color:var(--ink-light);margin-top:0.15rem;font-style:italic">{s["reason"]}</div>' if s.get("reason") else ""}
</div>"""
        html_parts.append(f"""<div style="margin-bottom:0.75rem">
<div style="font-size:0.65rem;font-weight:800;letter-spacing:0.1em;text-transform:uppercase;color:var(--ink-light);padding:0.5rem 1rem;background:var(--paper-2);border-bottom:1px solid var(--rule)">🎯 Scheduled for today</div>
{cards}</div>""")

    if future_items:
        # Group by date
        by_date = {}
        for s in future_items:
            d = s["start_dt"].strftime("%a %d %b")
            by_date.setdefault(d, []).append(s)
        rows = ""
        for date_str, items in list(by_date.items())[:3]:
            titles = ", ".join(s["task"]["title"][:35] for s in items[:3])
            if len(items) > 3:
                titles += f" +{len(items)-3} more"
            rows += f'<div style="padding:0.45rem 1rem;border-bottom:1px solid var(--rule);font-size:0.8rem"><span style="font-weight:700;min-width:80px;display:inline-block">{date_str}</span><span style="color:var(--ink-light)">{titles}</span></div>'
        html_parts.append(f"""<div style="margin-bottom:0.75rem">
<div style="font-size:0.65rem;font-weight:800;letter-spacing:0.1em;text-transform:uppercase;color:var(--ink-light);padding:0.5rem 1rem;background:var(--paper-2);border-bottom:1px solid var(--rule)">📅 Upcoming scheduled tasks</div>
{rows}</div>""")

    if flagged:
        rows = ""
        for ev in flagged:
            rows += f'<div style="padding:0.45rem 1rem;border-bottom:1px solid var(--rule);font-size:0.8rem;display:flex;gap:0.5rem"><span style="color:#d35400">⚑</span><span>{strip_block_icon(ev["subject"])}</span><span style="margin-left:auto;font-size:0.7rem;color:var(--ink-light)">{ev["start_dt"].strftime("%H:%M")}</span></div>'
        html_parts.append(f"""<div style="margin-bottom:0.75rem">
<div style="font-size:0.65rem;font-weight:800;letter-spacing:0.1em;text-transform:uppercase;padding:0.5rem 1rem;background:#fef3e2;border-bottom:1px solid #f5cba7;color:#d35400">⚑ Yesterday's blocks — review and reschedule if needed</div>
{rows}</div>""")

    if unscheduled:
        rows = "".join(
            f'<div style="padding:0.35rem 1rem;font-size:0.78rem;color:var(--ink-light)">• {u["task"]["title"][:60]} — {u.get("reason","no slot available")}</div>'
            for u in unscheduled[:5]
        )
        html_parts.append(f"""<div>
<div style="font-size:0.65rem;font-weight:800;letter-spacing:0.1em;text-transform:uppercase;color:var(--ink-light);padding:0.5rem 1rem;background:var(--paper-2);border-bottom:1px solid var(--rule)">⚠️ Could not schedule</div>
{rows}</div>""")

    return f'<div style="border:1px solid var(--rule);border-radius:3px;overflow:hidden">{"".join(html_parts)}</div>'


# ── MAIN ENTRY POINT ──────────────────────────────────────────────────────────

def save_plan(result: dict) -> None:
    """Serialise the scheduler result to scheduling_plan.json for the dashboard."""
    import copy

    def _serialise(obj):
        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        if isinstance(obj, datetime.date):
            return obj.isoformat()
        return str(obj)

    def _clean_scheduled(items):
        out = []
        for s in items:
            task = {k: v for k, v in s["task"].items() if k != "raw"}
            out.append({
                "task":           task,
                "date":           s["start_dt"].date().isoformat(),
                "start_time":     s["start_dt"].strftime("%H:%M"),
                "end_time":       s["end_dt"].strftime("%H:%M"),
                "estimated_mins": s["estimated_mins"],
                "title":          s["title"],
                "description":    s["description"],
                "reason":         s.get("reason", ""),
            })
        return out

    def _clean_flagged(items):
        return [
            {
                "subject":    ev["subject"],
                "start_time": ev["start_dt"].strftime("%H:%M"),
                "date":       ev["start_dt"].date().isoformat(),
            }
            for ev in items
        ]

    def _clean_unscheduled(items):
        return [
            {
                "task":   {k: v for k, v in u["task"].items() if k != "raw"},
                "reason": u.get("reason", "no suitable slot"),
            }
            for u in items
        ]

    plan = {
        "generated_at": datetime.datetime.now(AEST_OFFSET).strftime(f"%A {DAYFMT} %B %Y, %H:%M AEST"),
        "scheduled":    _clean_scheduled(result.get("scheduled", [])),
        "unscheduled":  _clean_unscheduled(result.get("unscheduled", [])),
        "flagged":      _clean_flagged(result.get("flagged", [])),
        "error":        result.get("error"),
    }

    PLAN_FILE.write_text(json.dumps(plan, indent=2, default=str), encoding="utf-8")



def run_scheduler(
    api_key: str = None,
    dry_run: bool = False,
    verbose: bool = True,
) -> dict:
    """
    Full scheduling run. Called from briefing.py or standalone.

    Returns dict with:
      scheduled, unscheduled, flagged, html_summary, ok_count, fail_count
    """
    result = {
        "scheduled":   [],
        "unscheduled": [],
        "flagged":     [],
        "html_summary": "",
        "ok_count":    0,
        "fail_count":  0,
        "error":       None,
    }

    if not _OUTLOOK_AVAILABLE:
        result["error"] = "outlook_email.py not available"
        return result

    rules = load_rules()
    if not rules:
        result["error"] = "scheduling_rules.json not found — run the Scheduling Rules dashboard first"
        return result

    global_rules   = rules.get("global", {})
    days_ahead     = global_rules.get("days_ahead", 5)
    # Safety-valve horizon for the "keep pushing into the future" retry below —
    # override with global.max_days_ahead in scheduling_rules.json if needed.
    max_days_ahead = max(global_rules.get("max_days_ahead", 60), days_ahead)

    try:
        token = _get_token()
    except Exception as e:
        result["error"] = f"Auth error: {e}"
        return result

    if verbose:
        print("\n📋  Fetching To Do tasks…")
    tasks = fetch_todo_tasks(token, rules)
    if verbose:
        print(f"   {len(tasks)} tasks to schedule "
              f"({sum(1 for t in tasks if t['is_overdue'])} overdue)")

    if verbose:
        print("📅  Fetching calendar events…")
    # Fetch the full possible search window up front, not just the initial
    # days_ahead window, so widening the search later doesn't miss real events.
    events = fetch_upcoming_events(token, days=max_days_ahead + 1)
    if verbose:
        print(f"   {len(events)} events found")

    # Flag yesterday's unfinished blocks
    flagged = get_flagged_blocks(events)
    _record_flagged_outcomes(flagged)   # duration learning (D8)
    if verbose and flagged:
        print(f"   ⚑  {len(flagged)} unfinished block(s) from yesterday flagged")

    if not tasks:
        if verbose:
            print(f"   No tasks to schedule — '{get_target_list_name(rules)}' is empty "
                  f"and nothing else is due today or overdue")
        result["flagged"]     = flagged
        result["html_summary"] = build_schedule_summary_html([], [], flagged)
        return result

    if verbose:
        print("🧩  Finding free slots…")
        print("🤖  AI scheduling tasks into slots…")

    scheduled, unscheduled = schedule_with_retry(
        tasks, events, rules, api_key,
        scheduler_fn=ai_schedule_tasks,
        initial_days=days_ahead,
        max_days=max_days_ahead,
        verbose=verbose,
    )

    if verbose:
        print(f"   {len(scheduled)} tasks scheduled, {len(unscheduled)} unscheduled")

    result["scheduled"]   = scheduled
    result["unscheduled"] = unscheduled
    result["flagged"]     = flagged

    if not dry_run and scheduled:
        if verbose:
            print("📆  Creating calendar events…")
        ok, fail = create_calendar_events(token, scheduled, rules)
        result["ok_count"]   = ok
        result["fail_count"] = fail
        if verbose:
            print(f"   ✓ {ok} events created, {fail} failed")

    result["html_summary"] = build_schedule_summary_html(
        scheduled, unscheduled, flagged)

    # Always save plan for dashboard — regardless of dry_run
    try:
        save_plan(result)
    except Exception as e:
        print(f"  ⚠️  Could not save plan file: {e}")

    return result


# ── CLI ───────────────────────────────────────────────────────────────────────

# ── DASHBOARD SERVER ─────────────────────────────────────────────────────────


def ai_schedule_tasks_with_durations(
    tasks: list[dict],
    slots: list[dict],
    rules: dict,
    api_key: str = None,
) -> list[dict]:
    """
    Schedule tasks using USER-SPECIFIED durations instead of AI estimates.
    The AI still picks the best slot and time, but honours the exact duration
    the user set for each task.
    """
    if not _ANTHROPIC_AVAILABLE or not tasks or not slots:
        return []

    global_rules = rules.get("global", {})
    scheduling   = rules.get("scheduling", {})
    cal_blocks   = rules.get("calendar_blocks", {})
    max_hours    = global_rules.get("max_hours", 3)
    max_blocks   = global_rules.get("max_blocks", 4)
    extra_instr  = scheduling.get("extra_instructions", "")

    # Build task lines — include user-set duration explicitly
    task_lines = "\n".join(
        f"[{i+1}] {t['title']}"
        + (f" (OVERDUE — due {t['due_date']})" if t["is_overdue"] else
           f" (due {t['due_date']})" if t["due_date"] else "")
        + f"\n    Duration: EXACTLY {t['user_duration_mins']} minutes (set by user — do not change)"
        + f"\n    Priority: {t['priority']}"
        + (f"\n    Notes: {t['body']}" if t.get("body") else "")
        # slot_pref and user_notes were collected by the dashboard but never
        # reached the prompt — the AI could not honour either.
        + (f"\n    Slot preference: {t['slot_pref']} (user set — respect unless impossible)"
           if t.get("slot_pref") and t["slot_pref"] != "any" else "")
        + (f"\n    User instruction: {t['user_notes']}" if t.get("user_notes") else "")
        for i, t in enumerate(tasks[:20])
    )

    slot_lines = "\n".join(
        f"[S{i+1}] {s['date']} {s['start_dt'].strftime('%H:%M')}–{s['end_dt'].strftime('%H:%M')} "
        f"({s['duration_mins']}min available, {s['slot_type'].replace('_', ' ')}, Week {s['week']})"
        for i, s in enumerate(slots[:40])
    )

    prompt = f"""You are a professional diary scheduler for Doug McAlpine, State Gas, Brisbane (AEST).

SCHEDULING RULES:
- Hard start: 09:00, Hard stop: 16:00 every day
- Deep work (contracts, documents, financial, regulatory): 09:00–12:00
- Admin/calls/emails: 12:00–16:00
- Max {max_hours} hours of task blocks per day, max {max_blocks} blocks per day
- Snap all tasks to start on the hour or half-hour
- Overdue tasks must be scheduled TODAY or TOMORROW at latest
- High priority tasks scheduled before normal priority tasks
- IMPORTANT: Use EXACTLY the duration specified for each task — the user has set these
{f'- {extra_instr}' if extra_instr else ''}

AVAILABLE SLOTS:
{slot_lines}

TASKS TO SCHEDULE:
{task_lines}

Rules for slot selection:
1. The task duration MUST fit within the slot's available time
2. Match deep work tasks to morning slots (deep work), admin to afternoon slots
3. Do not schedule more than {max_hours}h total per day

Respond ONLY as JSON:
{{
  "schedule": [
    {{
      "task_index": 1,
      "slot_index": "S1",
      "estimated_mins": 60,
      "reason": "one sentence why this slot works"
    }}
  ],
  "unscheduled": [
    {{
      "task_index": 2,
      "reason": "why it could not be scheduled"
    }}
  ]
}}"""

    data = call_json(get_client(), prompt, max_tokens=2000,
                     model=SCHEDULER_MODEL, label="AI scheduling")
    if not isinstance(data, dict):
        return []

    prefix  = cal_blocks.get("block_prefix", "🎯 ")
    results = []

    for item in data.get("schedule", []):
        try:
            t_idx = int(item["task_index"]) - 1
            s_idx = int(item["slot_index"][1:]) - 1
            task  = tasks[t_idx]
            slot  = slots[s_idx]

            # Use user-specified duration — not AI estimate
            est = task["user_duration_mins"]
            est = min(est, slot["duration_mins"])  # can't exceed slot
            est = max(15, round(est / 15) * 15)    # snap to 15min

            end_dt = slot["start_dt"] + datetime.timedelta(minutes=est)

            results.append({
                "task":           task,
                "slot":           slot,
                "estimated_mins": est,
                "start_dt":       slot["start_dt"],
                "end_dt":         end_dt,
                "title":          block_title(task['title'], task.get("body", ""), cal_blocks),
                "description":    _build_event_body(task, cal_blocks),
                "reason":         item.get("reason", ""),
            })
        except (IndexError, KeyError, ValueError):
            continue

    # Guard against the AI assigning two tasks into overlapping time ranges
    return _resolve_overlaps(results)


def _auto_describe_task(title: str, body: str, due: str, overdue: bool,
                         list_name: str, api_key: str = None, web_link: str = "") -> str:
    """
    Generate a 1-2 sentence calendar event description using Claude.
    Falls back to a plain text description if the API call fails.
    """
    # If body is already substantial, just format it cleanly
    parts = []
    if overdue:
        parts.append(f"⚠️ OVERDUE (was due {due})")
    elif due:
        parts.append(f"Due: {due}")
    if list_name:
        parts.append(f"From: {list_name}")
    if web_link:
        parts.append(f"Source email: {web_link}")

    if body and len(body.strip()) > 40:
        # Body is already descriptive — use it directly
        return body.strip() + ("\n" + "\n".join(parts) if parts else "")

    # Body is sparse or empty — ask Claude to write a description
    try:
        _client = get_client()
        context = body.strip() if body else ""
        prompt = f"""Write a 1-2 sentence calendar event description for this task.
Be specific and actionable — what needs to happen during this time block.

Task title: {title}
{f"Context notes: {context}" if context else "No additional context provided."}
{f"Due: {due}" if due else ""}
{f"List: {list_name}" if list_name else ""}

Write just the description text, no headings or labels. Max 2 sentences."""

        resp = api_call_with_retry(
            _client,
            model="claude-haiku-4-5-20251001",
            max_tokens=120,
            messages=[{"role": "user", "content": prompt}],
        )
        generated = resp.content[0].text.strip()
        if parts:
            generated += "\n" + "\n".join(parts)
        return generated
    except Exception:
        # Fallback to plain text
        return (body + "\n" + "\n".join(parts)).strip() if body else "\n".join(parts)



def serve_dashboard(api_key: str = None) -> None:
    """
    Serve the scheduling dashboard on localhost and handle:
      GET  /plan       — return scheduling_plan.json
      POST /refresh    — regenerate plan, return new plan JSON
      POST /push_event — create one calendar event, return {success, error}
    Opens browser automatically.
    """
    import threading
    import webbrowser
    from http.server import HTTPServer, BaseHTTPRequestHandler

    dashboard_file = Path(__file__).parent / "scheduling_dashboard.html"
    if not dashboard_file.exists():
        raise SystemExit("❌  scheduling_dashboard.html not found in project folder")

    dashboard_html = dashboard_file.read_text(encoding="utf-8")

    try:
        token = _get_token()
    except Exception as e:
        raise SystemExit(f"❌  Outlook auth error: {e}\n    Run: py outlook_email.py setup")

    rules = load_rules()
    if not rules:
        raise SystemExit("❌  scheduling_rules.json not found\n    Open scheduling_rules.html and save your rules first")

    cal_blocks = rules.get("calendar_blocks", {})

    def _create_one_event(item: dict) -> dict:
        """Create a single calendar event from a dashboard push_event request."""
        import datetime as _dt
        try:
            date_str  = item["date"]
            time_str  = item["start_time"]
            dur_mins  = int(item["duration_mins"])
            title     = item["task_title"]
            body_text = item.get("task_body", "")
            due       = item.get("task_due", "")
            overdue   = item.get("task_overdue", False)
            list_name = item.get("task_list", "")
            web_link  = item.get("web_link", "")

            prefix    = cal_blocks.get("block_prefix", "🎯 ")
            show_as   = "free" if rules.get("global", {}).get("mark_as_free", True) else cal_blocks.get("block_show_as", "free")
            reminder  = cal_blocks.get("block_reminder", 10)

            # Auto-generate a description using Claude if body is sparse
            description = _auto_describe_task(
                title=title,
                body=body_text,
                due=due,
                overdue=overdue,
                list_name=list_name,
                web_link=web_link,
                api_key=api_key,
            )

            # Parse datetime in AEST then convert to UTC for Graph
            dt_local  = _dt.datetime.fromisoformat(f"{date_str}T{time_str}:00").replace(tzinfo=AEST_OFFSET)
            dt_end    = dt_local + _dt.timedelta(minutes=dur_mins)
            start_utc = dt_local.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
            end_utc   = dt_end.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")

            body = {
                "subject": block_title(title, body_text, cal_blocks),
                "body": {"contentType": "text", "content": description},
                "start": {"dateTime": start_utc, "timeZone": "UTC"},
                "end":   {"dateTime": end_utc,   "timeZone": "UTC"},
                "showAs": show_as,
                "isReminderOn": reminder > 0,
                "reminderMinutesBeforeStart": reminder,
                "categories": [BLOCK_CATEGORY],
            }
            # Fresh token per push (A6) — MSAL silent-refreshes, so a
            # dashboard left open for hours keeps working.
            try:
                _tok = _get_token()
            except Exception:
                _tok = token
            _graph_post(_tok, "/me/events", body)
            _log_scheduled_duration(f"{prefix}{title}", dur_mins)
            print(f"   ✓  Created: {prefix}{title[:50]} @ {date_str} {time_str}")
            return {"success": True}
        except Exception as e:
            print(f"   ✗  Failed: {item.get('task_title','?')[:50]}: {e}")
            return {"success": False, "error": str(e)}

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path == "/":
                self._serve_html(dashboard_html)
            elif self.path == "/plan":
                if PLAN_FILE.exists():
                    self._serve_json(PLAN_FILE.read_text(encoding="utf-8"))
                else:
                    self._serve_json(json.dumps({"error": "No plan file found."}))
            elif self.path.startswith("/tasks"):
                # Return raw tasks for the Task Review step (no scheduling yet)
                # Optional ?lookback=72 query param to override default 48h lookback
                from urllib.parse import urlparse, parse_qs
                _qs = parse_qs(urlparse(self.path).query)
                _lb = int(_qs.get("lookback", ["48"])[0])
                try:
                    try:
                        _tok = _get_token()
                    except Exception:
                        _tok = token
                    raw_tasks = fetch_todo_tasks(_tok, rules, lookback_days=max(1, -(-_lb // 24)))  # page sends hours
                    tasks_out = [
                        {
                            "id":        t["id"],
                            "title":     t["title"],
                            "body":      t["body"],
                            "due_date":  t["due_date"],
                            "priority":  t["priority"],
                            "list_name": t["list_name"],
                            "is_overdue": t["is_overdue"],
                        }
                        for t in raw_tasks
                    ]
                    self._serve_json(json.dumps({"tasks": tasks_out}))
                except Exception as e:
                    self._serve_json(json.dumps({"error": str(e)}))
            else:
                self.send_response(404); self.end_headers()

        def do_POST(self):
            length  = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length)) if length else {}

            if self.path == "/refresh":
                print("\n   ↺  Regenerating plan…")
                result = run_scheduler(api_key, dry_run=True, verbose=True)
                plan_json = PLAN_FILE.read_text(encoding="utf-8") if PLAN_FILE.exists() else json.dumps({})
                self._serve_json(plan_json)

            elif self.path == "/schedule_with_inputs":
                # Receive user-reviewed tasks with durations/priorities,
                # run AI scheduling using those exact values, return plan
                user_tasks    = payload.get("tasks", [])
                lookback_hours = int(payload.get("lookback_hours", 48))
                print(f"\n   🧠  Scheduling {len(user_tasks)} user-reviewed tasks (lookback {lookback_hours}h)…")
                try:
                    days_ahead     = rules.get("global", {}).get("days_ahead", 5)
                    max_days_ahead = max(rules.get("global", {}).get("max_days_ahead", 60), days_ahead)

                    # Fetch the full possible search window up front so widening
                    # the search later doesn't miss real events further out.
                    try:
                        _tok = _get_token()
                    except Exception:
                        _tok = token
                    events = fetch_upcoming_events(_tok, days=max_days_ahead + 1)

                    # Build task dicts from user inputs, overriding AI estimates
                    reviewed_tasks = []
                    for ut in user_tasks:
                        reviewed_tasks.append({
                            "id":         ut["id"],
                            "title":      ut["title"],
                            "body":       ut.get("body", ""),
                            "due_date":   ut.get("due_date"),
                            "priority":   ut["priority"],
                            "list_name":  ut.get("list_name", ""),
                            "is_overdue": ut.get("is_overdue", False),
                            "list_id":    ut.get("list_id", ""),
                            # User-specified duration stored for the AI prompt
                            "user_duration_mins": ut.get("duration_mins", 0),
                            "slot_pref":  ut.get("slot_pref", "any"),
                            "user_notes": ut.get("user_notes", ""),
                        })

                    # Run AI scheduling — widening the window until every task
                    # is placed (or max_days_ahead is hit), pushing tasks that
                    # don't fit the near-term calendar further into the future.
                    scheduled, unscheduled = schedule_with_retry(
                        reviewed_tasks, events, rules, api_key,
                        scheduler_fn=ai_schedule_tasks_with_durations,
                        initial_days=days_ahead,
                        max_days=max_days_ahead,
                        verbose=True,
                    )
                    flagged = get_flagged_blocks(events)

                    result = {
                        "scheduled":   scheduled,
                        "unscheduled": unscheduled,
                        "flagged":     flagged,
                    }
                    save_plan(result)
                    plan_json = PLAN_FILE.read_text(encoding="utf-8") if PLAN_FILE.exists() else json.dumps({})
                    self._serve_json(plan_json)
                    print(f"   ✓  Scheduled {len(scheduled)} tasks"
                          + (f", {len(unscheduled)} still unplaced" if unscheduled else ""))
                except Exception as e:
                    import traceback
                    traceback.print_exc()
                    self._serve_json(json.dumps({"error": str(e)}))

            elif self.path == "/push_event":
                result = _create_one_event(payload)
                self._serve_json(json.dumps(result))

            elif self.path == "/carry_forward":
                # D7 — convert yesterday's flagged 🎯 blocks into tasks and
                # schedule them into upcoming free slots, keeping each block's
                # original duration. Returns the updated plan JSON.
                print("\n   ⚑  Carrying forward yesterday's unfinished blocks…")
                try:
                    try:
                        _tok = _get_token()
                    except Exception:
                        _tok = token
                    days_ahead     = rules.get("global", {}).get("days_ahead", 5)
                    max_days_ahead = max(rules.get("global", {}).get("max_days_ahead", 60), days_ahead)
                    events  = fetch_upcoming_events(_tok, days=max_days_ahead + 1)
                    flagged = get_flagged_blocks(events)
                    if not flagged:
                        self._serve_json(json.dumps({"error": "No unfinished blocks from yesterday to carry forward."}))
                        return
                    today_iso = datetime.date.today().isoformat()
                    carry_tasks = []
                    for i, ev in enumerate(flagged):
                        dur = max(30, int((ev["end_dt"] - ev["start_dt"]).total_seconds() / 60))
                        carry_tasks.append({
                            "id":         f"carry_{i}",
                            "title":      strip_block_icon(ev["subject"]),
                            "body":       "Carried forward — yesterday's block was not completed.",
                            "due_date":   today_iso,
                            "priority":   "high",
                            "list_name":  "Carried forward",
                            "is_overdue": True,
                            "list_id":    "",
                            "user_duration_mins": dur,
                        })
                    scheduled, unscheduled = schedule_with_retry(
                        carry_tasks, events, rules, api_key,
                        scheduler_fn=ai_schedule_tasks_with_durations,
                        initial_days=days_ahead,
                        max_days=max_days_ahead,
                        verbose=True,
                    )
                    # Merge into the existing saved plan so the dashboard
                    # shows carried blocks alongside the day's normal plan.
                    existing = {}
                    if PLAN_FILE.exists():
                        try:
                            existing = json.loads(PLAN_FILE.read_text(encoding="utf-8"))
                        except Exception:
                            existing = {}
                    merged = {
                        "scheduled":   scheduled,
                        "unscheduled": unscheduled,
                        "flagged":     [],
                    }
                    save_plan(merged)
                    new_plan = json.loads(PLAN_FILE.read_text(encoding="utf-8"))
                    new_plan["scheduled"]   = (existing.get("scheduled", []) or []) + new_plan["scheduled"]
                    new_plan["unscheduled"] = (existing.get("unscheduled", []) or []) + new_plan["unscheduled"]
                    PLAN_FILE.write_text(json.dumps(new_plan, indent=2, default=str), encoding="utf-8")
                    print(f"   ✓  {len(scheduled)} block(s) carried forward"
                          + (f", {len(unscheduled)} could not be placed" if unscheduled else ""))
                    self._serve_json(json.dumps(new_plan, default=str))
                except Exception as e:
                    import traceback
                    traceback.print_exc()
                    self._serve_json(json.dumps({"error": str(e)}))

            elif self.path == "/complete_task":
                # D9 — mark a Microsoft To Do task as completed straight
                # from the dashboard. Payload: {list_id, task_id}
                list_id = payload.get("list_id", "")
                task_id = payload.get("task_id", "")
                if not list_id or not task_id:
                    self._serve_json(json.dumps({"success": False, "error": "list_id and task_id required"}))
                    return
                try:
                    try:
                        _tok = _get_token()
                    except Exception:
                        _tok = token
                    _graph_patch(_tok, f"/me/todo/lists/{list_id}/tasks/{task_id}",
                                 {"status": "completed"})
                    print(f"   ✓  Task marked complete: {task_id[:20]}…")
                    self._serve_json(json.dumps({"success": True}))
                except Exception as e:
                    print(f"   ✗  Complete failed: {e}")
                    self._serve_json(json.dumps({"success": False, "error": str(e)}))

            else:
                self.send_response(404); self.end_headers()

        def _serve_html(self, html: str):
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _serve_json(self, json_str: str):
            body = json_str.encode("utf-8") if isinstance(json_str, str) else json_str
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, fmt, *args):
            pass  # suppress HTTP noise

    for port in (8766, 8767, 8768):
        try:
            server = HTTPServer(("localhost", port), Handler)
            break
        except OSError:
            continue
    else:
        raise SystemExit("❌  No free port found (tried 8766–8768)")

    url = f"http://localhost:{port}"
    print(f"\n   🌐  Scheduling Dashboard: {url}")
    print(f"       Review proposed blocks, adjust times, then click Push to Outlook.")
    print(f"       Press Ctrl+C to close.\n")
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    try:
        while True:
            server.handle_request()
    except KeyboardInterrupt:
        print("\n   Server closed.")
    finally:
        server.server_close()


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise SystemExit("❌  Set ANTHROPIC_API_KEY environment variable")

    mode = sys.argv[1] if len(sys.argv) > 1 else "dashboard"

    if mode == "dashboard":
        # Generate plan then open interactive dashboard
        print("🗓️  Generating scheduling plan…")
        result = run_scheduler(api_key, dry_run=True, verbose=True)
        if result.get("error"):
            print(f"\n⚠️  {result['error']}")
        print("\n🖥️   Opening Scheduling Dashboard…")
        serve_dashboard(api_key)

    elif mode == "preview":
        print("🗓️  Generating plan (preview only — no events created)")
        result = run_scheduler(api_key, dry_run=True)
        if result.get("error"):
            print(f"\n❌  {result['error']}"); sys.exit(1)
        print(f"\n{'='*55}\nPROPOSED SCHEDULE\n{'='*55}")
        for s in result["scheduled"]:
            print(f"  {s['start_dt'].strftime('%a %d %b %H:%M')} – {s['end_dt'].strftime('%H:%M')}  {s['task']['title'][:50]}  ({s['estimated_mins']}min)")
        if result["unscheduled"]:
            print(f"\nUnscheduled: {len(result['unscheduled'])} task(s)")
        print("\nPlan saved to scheduling_plan.json")
        print("Run 'py outlook_scheduler.py dashboard' to open the interactive dashboard")

    elif mode == "flagged":
        result = run_scheduler(api_key, dry_run=True, verbose=False)
        if result["flagged"]:
            print(f"⚑  {len(result['flagged'])} unfinished block(s) from yesterday:")
            for ev in result["flagged"]:
                print(f"   • {ev['subject']} @ {ev['start_dt'].strftime('%H:%M')}")
        else:
            print("✓ No unfinished blocks from yesterday")
    else:
        print("Usage: py outlook_scheduler.py [dashboard|preview|flagged]")
        sys.exit(1)
