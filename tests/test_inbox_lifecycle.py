"""
test_inbox_lifecycle.py
-----------------------
The inbox_actions carried-forward list, as redesigned on 15 October:

  - every action is tied to its source email (msg_id), by the EMAIL n number
  - this run's actions do not also appear as "carried forward"
  - pushed items leave the list (they live in To Do) and are never re-offered
  - dismissed and resolved emails are never re-offered
  - anything already in To Do (by [src:] marker) is closed and hidden
  - normal/low items expire after CARRY_EXPIRY_DAYS; high never do
  - new actions are recorded deterministically, deduped by source email

Uses a temporary state file - never the real inbox_state.json.

    py tests/test_inbox_lifecycle.py
"""
import sys, os, re, json, datetime, tempfile, subprocess, contextlib, io
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
with contextlib.redirect_stdout(io.StringIO()):
    import inbox_actions as ia
from core.knowledge import InboxState

TODAY = datetime.date.today()
def ago(n): return (TODAY - datetime.timedelta(days=n)).isoformat()

tmp = Path(tempfile.mkdtemp()) / "inbox_state.json"
st = InboxState(); st.path = tmp
st.data = json.loads(json.dumps(InboxState.EMPTY))
st.data["run_count"] = 6
st.data["open_actions"] = [
    {"id": "act-001", "title": "Old carried item", "status": "open", "priority": "normal",
     "raised_date": ago(5), "msg_id": "M-OLD", "topic": "T"},
    {"id": "act-002", "title": "Stale normal item", "status": "open", "priority": "normal",
     "raised_date": ago(30), "msg_id": "M-STALE", "topic": "T"},
    {"id": "act-003", "title": "Stale but high", "status": "open", "priority": "high",
     "raised_date": ago(30), "msg_id": "M-HIGH", "topic": "T"},
    {"id": "act-004", "title": "Legacy pushed", "status": "pushed_to_todo", "priority": "normal",
     "raised_date": ago(10), "topic": "T"},
    {"id": "act-005", "title": "Already in To Do via briefing", "status": "open",
     "priority": "normal", "raised_date": ago(3), "msg_id": "M-INTODO", "topic": "T"},
]

# _attach_sources: the model returns "EMAIL n" numbers, we look up the ids
emails = [{"id": f"M-{i}", "subject": f"Subject {i}", "datetime": f"{ago(1)} 09:00:00"}
          for i in range(1, 6)]
groups = [{"topic": "Rolleston", "theme_summary": "x", "actions": [
    {"title": "Reply to Caroline", "email_number": "EMAIL 2", "priority": "high"},
    {"title": "Send farm-in pack", "email_number": "3", "priority": "normal"},
    {"title": "Answer the ACCC notice", "email_number": "[4]", "priority": "high"},
    {"title": "Unresolvable number", "email_number": "99"},
]}]
ia._attach_sources(groups, emails)
acts = groups[0]["actions"]

todo_src = {"M-INTODO", "M-4"}            # M-4 = the ACCC email, already in To Do
n_todo = st.sync_with_todo(todo_src)
n_exp  = st.expire_stale()

data = {"action_groups": groups, "total_emails_reviewed": 5, "summary": ""}
st.mark_dismissed([{"msg_id": "M-3", "title": "Send farm-in pack"}])   # dismissed earlier
hidden = ia._filter_handled(data, st.handled_sources() | todo_src)
shown  = [a for g in data["action_groups"] for a in g["actions"]]

flat  = [dict(a, topic=g["topic"]) for g in data["action_groups"] for a in g["actions"]]
added = st.add_new_actions(flat)
again = st.add_new_actions(flat)          # second run, same emails -> no duplicates

with contextlib.redirect_stdout(io.StringIO()):
    page = ia.generate_dashboard_html(data, st)
prior_ids = set(re.findall(r'id="prior-card-(act-\d+)"', page))

# push the remaining card, then make sure it can't come back
st.mark_pushed(["Reply to Caroline"], ["M-2"])
data2 = {"action_groups": [{"topic": "Rolleston", "actions": [
    {"title": "Reply to Caroline (reworded)", "msg_id": "M-2"}]}]}
hidden2 = ia._filter_handled(data2, st.handled_sources())

st.mark_resolved(["act-001"])
data3 = {"action_groups": [{"topic": "T", "actions": [{"title": "x", "msg_id": "M-OLD"}]}]}
hidden3 = ia._filter_handled(data3, st.handled_sources())

# the page's script must parse
scripts = "\n".join(re.findall(r"<script>(.*?)</script>", page, re.S))
js = Path(tempfile.mkdtemp()) / "inbox.js"; js.write_text(scripts, encoding="utf-8")
try:
    node_ok = subprocess.run(["node", "--check", str(js)], capture_output=True).returncode == 0
except FileNotFoundError:
    node_ok = None

CHECKS = [
    ("EMAIL n resolves to the source message id",
     [a.get("msg_id") for a in acts[:3]] == ["M-2", "M-3", "M-4"]),
    ("an out-of-range number leaves the action unlinked, not mislinked",
     not acts[3].get("msg_id")),
    ("each linked action gets an Outlook link",
     acts[0].get("web_link", "").startswith("https://outlook.office365.com/owa/?ItemID=M-2")),
    ("legacy 'pushed_to_todo' items leave the carried list", "act-004" not in prior_ids),
    ("an item already in To Do (via the briefing) is closed", n_todo == 2),
    ("a stale normal item expires", n_exp == 1 and "act-002" not in prior_ids),
    ("a stale HIGH item does not expire", "act-003" in prior_ids),
    ("an email already in To Do is not offered as new",
     all(a.get("msg_id") != "M-4" for a in shown)),
    ("a dismissed email is not offered as new",
     all(a.get("msg_id") != "M-3" for a in shown)),
    ("hidden count is reported", hidden == 2 and data.get("hidden_handled") == 2),
    ("new actions are recorded once, with their source", len(added) == 2 and not again
     and {a["msg_id"] for a in added} == {"M-2", ""}),
    ("this run's actions are NOT also shown as carried forward",
     not ({a["id"] for a in added} & prior_ids)),
    ("a genuinely carried item is shown", "act-001" in prior_ids),
    ("a pushed email can't come back, even reworded", hidden2 == 1),
    ("a resolved email can't come back", hidden3 == 1),
    ("handled sources survive pruning inside the retention window",
     (st.prune() or True) and "M-2" in st.handled_sources()),
    ("the page has a Dismiss control and the /dismiss call",
     'class="dismiss-btn"' in page and "fetch('/dismiss'" in page),
    ("the page uses the Control Room's look (crimson, light header, no black strip)",
     "--crimson:#A51C30" in page and "Aptos" in page and "Playfair" not in page
     and 'class="page-header"' in page and "background:var(--ink);color:var(--white)" not in page),
    ("the dashboard's JavaScript parses", node_ok is not False),
]

if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    bad = 0
    for label, ok in CHECKS:
        print(("ok   " if ok else "FAIL ") + label)
        bad += not ok
    if node_ok is None:
        print("     (node not installed - JS parse check skipped)")
    print(f"\n{len(CHECKS) - bad}/{len(CHECKS)} passed")
    if os.environ.get("DUMP_PAGE"):
        Path(os.environ["DUMP_PAGE"]).write_text(page, encoding="utf-8")
    sys.exit(1 if bad else 0)
