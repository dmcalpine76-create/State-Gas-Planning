"""
weekly_plan.py - the Sunday weekly planning run.

Every Sunday at 5pm (GitHub Actions) it:
  1. reads your week: email, the Daily Priorities list in To Do, your calendar
  2. brings the knowledge store up to date: matters, people, deadlines
  3. works out what slipped, next week's shape, deadlines in the next 6 weeks,
     matters that have gone quiet, and what needs preparing
  4. writes a one-page weekly plan: kept in the knowledge folder, published as
     an encrypted page, and emailed to you with the link that opens it

  py weekly_plan.py setup          one-time sign-in (also refreshes the cloud copy)
  py weekly_plan.py run [--days 7] run on the laptop; opens the plan when done
  py weekly_plan.py run --email    ...and email it as the Sunday run does

Nothing company-specific is printed while it runs: cloud run logs are public.
"""
import os
import io
import sys
import json
import base64
import secrets
import argparse
import datetime
import tempfile
import contextlib
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from dotenv import load_dotenv
load_dotenv(HERE / ".env")

AEST       = datetime.timezone(datetime.timedelta(hours=10))
CLOUD      = os.environ.get("CI") == "true"
# One Microsoft sign-in for every process (morning briefing, Friday wrap,
# Sunday plan, board tools, inbox_actions): the union of what each needs.
SCOPES     = ["Mail.ReadWrite", "Mail.Read", "Mail.Send", "Calendars.ReadWrite",
              "Calendars.Read", "Tasks.ReadWrite", "Files.ReadWrite", "User.Read"]
MB_REPO    = "dmcalpine76-create/morning-briefing"
CACHE_FILE = HERE / ".weekly_token_cache.bin"
TASK_LIST  = "Daily Priorities"
SITE       = "https://dmcalpine76-create.github.io/State-Gas-Planning"
REPO       = "dmcalpine76-create/State-Gas-Planning"
PUBLISH_DIR = HERE / "docs" / "weekly"   # Pages publishes docs/


def say(msg: str):
    """Progress lines. Keep these free of names, subjects and content."""
    print(f"  {msg}", flush=True)


@contextlib.contextmanager
def quiet():
    """In the cloud, swallow library output (it can include folder names)."""
    if not CLOUD:
        yield
        return
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        yield


# ═══════════════════════════════════════════════════════════════════════════
# SIGN-IN
# ═══════════════════════════════════════════════════════════════════════════

def _msal_app(cache):
    import msal
    tenant = os.environ.get("OUTLOOK_TENANT_ID", "organizations")
    return msal.PublicClientApplication(
        os.environ["OUTLOOK_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{tenant}",
        token_cache=cache)


def _load_cache():
    import msal
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text(encoding="utf-8"))
    return cache


def get_token() -> str:
    cache = _load_cache()
    app = _msal_app(cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not result or "access_token" not in result:
        raise SystemExit("Sign-in needed: on the laptop, press 'Microsoft sign-in (all processes)' "
                         "in the Control Room (or run: py weekly_plan.py setup).")
    if cache.has_state_changed:
        CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def _github_token() -> str:
    tok = os.environ.get("GH_SECRETS_TOKEN") or os.environ.get("GITHUB_UPLOAD_TOKEN")
    if not tok:
        mb = HERE.parent / "morning briefing system" / ".env"
        if mb.exists():
            from dotenv import dotenv_values
            tok = dotenv_values(mb).get("GITHUB_UPLOAD_TOKEN")
    return (tok or "").strip()


def put_secret(name: str, value: str, repo: str = REPO) -> bool:
    """Store a GitHub Actions secret on a repository (the planning one by default)."""
    import requests
    try:
        from nacl import encoding, public
    except ImportError:                       # first run on the laptop
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pynacl"], check=False)
        from nacl import encoding, public
    tok = _github_token()
    if not tok:
        return False
    h = {"Authorization": f"Bearer {tok}", "Accept": "application/vnd.github+json"}
    k = requests.get(f"https://api.github.com/repos/{repo}/actions/secrets/public-key",
                     headers=h, timeout=30)
    k.raise_for_status()
    k = k.json()
    box = public.SealedBox(public.PublicKey(k["key"].encode(), encoding.Base64Encoder()))
    enc = base64.b64encode(box.encrypt(value.encode())).decode()
    r = requests.put(f"https://api.github.com/repos/{repo}/actions/secrets/{name}",
                     headers=h, json={"encrypted_value": enc, "key_id": k["key_id"]},
                     timeout=30)
    return r.status_code in (201, 204)


def share_signin(text: str) -> list:
    """Give every process the same sign-in. Returns what could not be updated."""
    failed = []
    if not put_secret("WEEKLY_TOKEN_CACHE", text):
        failed.append("Sunday run (cloud)")
    if not (put_secret("OUTLOOK_TOKEN_CACHE", text, MB_REPO)
            and put_secret("OUTLOOK_TOKEN_SETUP_DATE", datetime.date.today().isoformat(), MB_REPO)):
        failed.append("morning briefing and Friday wrap (cloud)")
    return failed


def _share_locally(text: str):
    """Laptop copies: inbox_actions / diary / board tools, and the local briefing."""
    for f in (HERE / ".outlook_token_cache.bin",
              HERE.parent / "morning briefing system" / ".outlook_token_cache.bin"):
        if not f.parent.exists():
            continue
        if f.exists():
            bak = f.with_name(f.name + ".before-shared-signin")
            if not bak.exists():
                bak.write_bytes(f.read_bytes())
        f.write_text(text, encoding="utf-8")
    setup = HERE.parent / "morning briefing system" / ".outlook_token_setup"
    if setup.parent.exists():
        setup.write_text(datetime.date.today().isoformat(), encoding="utf-8")


def cmd_setup(_args):
    cache = _load_cache()
    app = _msal_app(cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise SystemExit(f"Could not start sign-in: {flow.get('error_description')}")
        print("\n" + "=" * 64)
        print("  MICROSOFT SIGN-IN FOR ALL STATE GAS PROCESSES")
        print("=" * 64)
        print(f"\n  1. Go to:    {flow['verification_uri']}")
        print(f"  2. Enter:    {flow['user_code']}")
        print("  3. Sign in with your State Gas account and accept the permissions")
        print("     (mail, calendar, To Do and OneDrive files).\n")
        try:
            webbrowser.open(flow["verification_uri"])
        except Exception:
            pass
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise SystemExit(f"Sign-in failed: {result.get('error_description')}")
    text = cache.serialize()
    CACHE_FILE.write_text(text, encoding="utf-8")
    print("  Signed in.")
    _share_locally(text)
    print("  Laptop tools updated (inbox actions, diary, board tools, local briefing).")
    failed = share_signin(text)
    if not failed:
        print("  Cloud copies updated - morning briefing, Friday wrap and Sunday run "
              "all use this sign-in.")
    else:
        print("  !! Could not update: " + ", ".join(failed) + ". Tell Claude.")


# ═══════════════════════════════════════════════════════════════════════════
# READING THE WEEK
# ═══════════════════════════════════════════════════════════════════════════

def _pages(token, url, params=None, headers=None, cap=1000):
    from core.graph import graph_get
    out, first = [], True
    while url and len(out) < cap:
        data = graph_get(token, url, params=params if first else None, headers=headers)
        out.extend(data.get("value", []))
        url, first = data.get("@odata.nextLink"), False
    return out


def fetch_tasks(token) -> list:
    lists = _pages(token, "/me/todo/lists")
    target = next((l for l in lists if l.get("displayName", "").strip().lower()
                   == TASK_LIST.lower()), None) \
        or next((l for l in lists if l.get("wellknownListName") == "defaultList"), None)
    if not target:
        return []
    raw = _pages(token, f"/me/todo/lists/{target['id']}/tasks", params={"$top": "100"})
    tasks = []
    for t in raw:
        tasks.append({
            "id":        t.get("id"),
            "title":     (t.get("title") or "").strip(),
            "status":    t.get("status"),
            "importance": t.get("importance"),
            "created":   (t.get("createdDateTime") or "")[:10],
            "modified":  (t.get("lastModifiedDateTime") or "")[:10],
            "completed": ((t.get("completedDateTime") or {}).get("dateTime") or "")[:10],
            "due":       ((t.get("dueDateTime") or {}).get("dateTime") or "")[:10],
            "note":      ((t.get("body") or {}).get("content") or "")[:300],
        })
    return tasks


def fetch_calendar(token, start: datetime.datetime, end: datetime.datetime) -> list:
    raw = _pages(token, "/me/calendarView",
                 params={"startDateTime": start.isoformat(), "endDateTime": end.isoformat(),
                         "$select": "subject,start,end,isAllDay,showAs,attendees,isCancelled,location",
                         "$top": "200", "$orderby": "start/dateTime"},
                 headers={"Prefer": 'outlook.timezone="E. Australia Standard Time"'})
    events = []
    for e in raw:
        if e.get("isCancelled"):
            continue
        att = [a.get("emailAddress", {}).get("address", "").lower() for a in e.get("attendees", [])]
        events.append({
            "subject":  (e.get("subject") or "(no subject)").strip(),
            "start":    e["start"]["dateTime"][:16],
            "end":      e["end"]["dateTime"][:16],
            "all_day":  e.get("isAllDay", False),
            "show_as":  e.get("showAs", "busy"),
            "external": any(a and not a.endswith("@stategas.com") for a in att),
            "people":   len(att),
            "location": ((e.get("location") or {}).get("displayName") or "")[:60],
        })
    return events


# ═══════════════════════════════════════════════════════════════════════════
# WORKING IT OUT (no AI: plain rules, so the numbers can be trusted)
# ═══════════════════════════════════════════════════════════════════════════

def _d(s):
    try:
        return datetime.date.fromisoformat(s[:10])
    except (TypeError, ValueError):
        return None


def analyse(tasks, events, prev_snapshot, deadlines, matters, today):
    open_tasks = [t for t in tasks if t["status"] != "completed"]
    done = [t for t in tasks if t["status"] == "completed" and _d(t["completed"])
            and (today - _d(t["completed"])).days <= 7]

    prev_ids = {x["id"] for x in (prev_snapshot or {}).get("open", [])}
    slipped = []
    for t in open_tasks:
        carried = t["id"] in prev_ids
        untouched = _d(t["modified"]) and (today - _d(t["modified"])).days
        if carried or (untouched and untouched >= 14):
            slipped.append({**t, "carried": carried,
                            "age": (today - _d(t["created"])).days if _d(t["created"]) else None,
                            "untouched": untouched})
    slipped.sort(key=lambda t: -(t["age"] or 0))

    # next working week: the Monday after today (the run is on Sunday)
    monday = today + datetime.timedelta(days=(7 - today.weekday()) % 7 or 7)
    days = []
    for i in range(5):
        day = monday + datetime.timedelta(days=i)
        evs = [e for e in events if e["start"][:10] == day.isoformat() and not e["all_day"]
               and e["show_as"] in ("busy", "tentative", "oof", "workingElsewhere")]
        mins = 0
        for e in evs:
            s = datetime.datetime.fromisoformat(e["start"]); f = datetime.datetime.fromisoformat(e["end"])
            mins += max(0, int((f - s).total_seconds() // 60))
        clashes = []
        for a in range(len(evs)):
            for b in range(a + 1, len(evs)):
                if evs[a]["start"] < evs[b]["end"] and evs[b]["start"] < evs[a]["end"]:
                    clashes.append((evs[a]["subject"], evs[b]["subject"]))
        days.append({"date": day.isoformat(), "label": day.strftime("%a %d %b").replace(" 0", " "),
                     "meetings": len(evs), "hours": round(mins / 60, 1),
                     "external": [e for e in evs if e["external"]],
                     "all_day": [e["subject"] for e in events if e["start"][:10] == day.isoformat() and e["all_day"]],
                     "clashes": clashes, "events": evs})

    horizon = today + datetime.timedelta(days=42)
    due_soon = sorted([d for d in deadlines if _d(d.get("date")) and today <= _d(d["date"]) <= horizon
                       and not d.get("done")], key=lambda d: d["date"])

    active = {n: m for n, m in matters.items() if m.get("status", "active") == "active"}
    quiet_m = sorted([(n, m) for n, m in active.items()
                      if _d(m.get("last_active")) and (today - _d(m["last_active"])).days >= 14],
                     key=lambda x: x[1].get("last_active", ""))
    no_step = [n for n, m in active.items() if not (m.get("next_step") or "").strip()]

    return {"open": open_tasks, "done": done, "slipped": slipped, "week": days,
            "monday": monday, "deadlines": due_soon, "quiet": quiet_m, "no_step": no_step,
            "active": active}


# ═══════════════════════════════════════════════════════════════════════════
# AI STEPS
# ═══════════════════════════════════════════════════════════════════════════

def _compact_matters(active: dict, limit=70) -> str:
    rows = sorted(active.items(), key=lambda x: x[1].get("last_active", ""), reverse=True)[:limit]
    return json.dumps({n: {k: m.get(k) for k in ("category", "description", "next_step",
                                                  "next_step_due", "last_active", "key_parties")
                           if m.get(k)} for n, m in rows}, ensure_ascii=False)


def consolidate(client, emails, a, inbox_state, instructions, today) -> dict:
    """Step 1: fold the week into matters, people and deadlines."""
    from core.graph import build_email_context
    from core.llm import call_json
    week_ago = (today - datetime.timedelta(days=7)).isoformat()
    topics = {n: {"summary": t.get("summary", "")[:400], "open_actions": t.get("open_actions", [])[:6]}
              for n, t in inbox_state.get("topics", {}).items() if (t.get("last_seen") or "") >= week_ago}
    email_ctx = build_email_context(emails, 7)[:90000]
    prompt = f"""You maintain the company knowledge base for the Managing Director of an ASX-listed
gas explorer. Once a week you fold the week's evidence into it. Today is {today:%A %d %B %Y}.

STANDING INSTRUCTIONS FROM THE MD:
{instructions[:4000] or "(none)"}

ACTIVE MATTERS (current knowledge):
{_compact_matters(a["active"])}

INBOX TOPICS UPDATED THIS WEEK (from the MD's inbox review tool):
{json.dumps(topics, ensure_ascii=False)[:15000]}

OPEN TASKS IN THE MD'S TO DO LIST:
{json.dumps([t["title"] for t in a["open"]][:120], ensure_ascii=False)}

THIS WEEK'S EMAIL:
{email_ctx}

YOUR TASK
1. For every matter that moved this week, return its updated record. Use the existing
   name exactly where it matches. Add genuinely new board-level matters. Mark matters
   clearly finished as closed. Leave untouched matters out.
2. Give each updated matter a concrete next_step (who does what) and next_step_due
   (YYYY-MM-DD) when the evidence gives one.
3. List dated obligations you can see: statutory, tenure, lodgement, contract, board,
   hearing or payment dates. Only real dates stated in the evidence.
4. List anything you are unsure about as a short question for the MD to confirm.
Be conservative: no speculation, no duplicates of existing matters under new names.

Respond with JSON only:
{{"matter_updates":[{{"name":"","category":"Commercial|Regulatory|JV/Partner|Operational|Financial|Stakeholder|Legal|Other",
   "status":"active|dormant|closed","description":"1-2 sentence current status","key_parties":[],
   "watch_for":"","next_step":"","next_step_due":""}}],
 "close_matter_names":[],
 "people_updates":[{{"name":"","email":"","organisation":"","role":"","notes":""}}],
 "new_facts":[],
 "deadlines":[{{"title":"","date":"YYYY-MM-DD","matter":"","kind":"statutory|tenure|lodgement|contract|board|hearing|payment|other","source":"email subject or task"}}],
 "confirm":[{{"question":"","matter":""}}]}}"""
    with quiet():
        out = call_json(client, prompt, max_tokens=12000, label="weekly consolidation")
    return out or {}


def plan(client, a, instructions, confirm, today) -> dict:
    """Step 2: turn the facts into next week's plan."""
    from core.llm import call_json
    week = [{"day": d["label"], "meetings": d["meetings"], "hours": d["hours"],
             "external": [e["subject"] for e in d["external"]], "all_day": d["all_day"],
             "clashes": d["clashes"]} for d in a["week"]]
    prompt = f"""You are the chief of staff to the Managing Director of an ASX-listed gas explorer.
Write his plan for the week starting {a["monday"]:%A %d %B %Y}. Address him as "you".
Be specific, brief and practical. No filler, no generic advice.

STANDING INSTRUCTIONS: {instructions[:3000] or "(none)"}

NEXT WEEK'S CALENDAR: {json.dumps(week, ensure_ascii=False)}
DEADLINES IN THE NEXT 6 WEEKS: {json.dumps(a["deadlines"], ensure_ascii=False)[:6000]}
TASKS THAT SLIPPED: {json.dumps([{"task": t["title"], "age_days": t["age"], "carried_from_last_week": t["carried"]} for t in a["slipped"][:25]], ensure_ascii=False)}
OPEN TASKS: {json.dumps([t["title"] for t in a["open"]][:80], ensure_ascii=False)}
ACTIVE MATTERS WITH NEXT STEPS: {_compact_matters(a["active"], 40)[:20000]}
MATTERS QUIET FOR 14+ DAYS: {json.dumps([n for n, _ in a["quiet"]][:20], ensure_ascii=False)}
QUESTIONS OPEN FOR THE MD: {json.dumps(confirm, ensure_ascii=False)[:3000]}

Respond with JSON only:
{{"headline":"2-3 sentences: what matters most this week and why",
 "priorities":[{{"title":"","why":"one line","matter":"","when":"day or by-date"}}],
 "prep_needed":[{{"meeting":"","day":"","what":"what to prepare"}}],
 "diary_blocks":[{{"task":"","day":"","minutes":60,"why":""}}],
 "risks":[{{"risk":"","action":""}}]}}
At most 6 priorities, 6 prep items, 8 diary blocks and 4 risks. Put diary blocks on
days with room, never in a clash."""
    with quiet():
        out = call_json(client, prompt, max_tokens=6000, label="weekly plan")
    return out or {}


def merge_deadlines(existing: list, found: list, today) -> tuple:
    have = {((d.get("title") or "").lower().strip(), d.get("date")) for d in existing}
    added = 0
    for d in found or []:
        key = ((d.get("title") or "").lower().strip(), d.get("date"))
        if not key[0] or not _d(d.get("date")) or key in have:
            continue
        existing.append({**d, "added": today.isoformat(), "confirmed": False})
        have.add(key); added += 1
    cutoff = today - datetime.timedelta(days=60)
    kept = [d for d in existing if not _d(d.get("date")) or _d(d["date"]) >= cutoff]
    return sorted(kept, key=lambda d: d.get("date", "")), added


# ═══════════════════════════════════════════════════════════════════════════
# THE PAGE
# ═══════════════════════════════════════════════════════════════════════════

CSS = """
:root{--paper:#F6F4F3;--surface:#fff;--ink:#1A1A1A;--muted:#6B6264;--rule:#E3DCDC;--crimson:#A51C30;
--good:#1E7A3C;--warn:#9A6A12;--bad:#B3261E;color-scheme:light}
@media (prefers-color-scheme:dark){:root{--paper:#151314;--surface:#1F1C1D;--ink:#EEE9E9;--muted:#A79C9E;
--rule:#352F31;--crimson:#E0566B;--good:#6FCF8E;--warn:#E3B55C;--bad:#F2877F;color-scheme:dark}}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);
font:15px/1.55 Aptos,"Segoe UI",system-ui,sans-serif;padding-inline:16px;padding-block:24px 48px}
.wrap{max-width:860px;margin:0 auto;display:grid;gap:26px}
header{border-bottom:3px solid var(--crimson);padding-bottom:12px}
header small{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:var(--crimson);font-weight:700}
h1{margin:4px 0 0;font-size:26px;line-height:1.2;text-wrap:balance}
.lead{font-size:17px;max-width:68ch;margin:0}
h2{margin:0 0 8px;font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted)}
.card{background:var(--surface);border:1px solid var(--rule);border-radius:10px;padding:6px 18px}
.item{padding:12px 0;border-top:1px solid var(--rule)}.item:first-child{border-top:0}
.item b{display:block}.item span{color:var(--muted);font-size:13.5px}
table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}
th,td{text-align:left;padding:9px 8px;border-top:1px solid var(--rule);vertical-align:top;font-size:14px}
th{color:var(--muted);font-weight:600;font-size:12.5px;border-top:0}
.scroll{overflow-x:auto}.pill{font-size:11.5px;padding:1px 7px;border-radius:999px;border:1px solid currentColor;white-space:nowrap}
.bad{color:var(--bad)}.warn{color:var(--warn)}.good{color:var(--good)}.muted{color:var(--muted)}
details summary{cursor:pointer;color:var(--muted);padding:10px 0}
footer{color:var(--muted);font-size:12.5px}
"""


def render(a, p, patch, stats, today) -> str:
    from html import escape as e
    wk = a["monday"]
    items = lambda rows: "".join(rows) or '<div class="item muted">Nothing this week.</div>'

    pri = [f'<div class="item"><b>{e(x.get("title",""))}</b><span>{e(x.get("why",""))}'
           f'{" · " + e(x["when"]) if x.get("when") else ""}</span></div>' for x in p.get("priorities", [])]

    dl_rows = []
    for d in a["deadlines"]:
        days = (_d(d["date"]) - today).days
        cls = "bad" if days <= 7 else "warn" if days <= 21 else "muted"
        flag = "" if d.get("confirmed", True) else ' <span class="pill muted">to confirm</span>'
        dl_rows.append(f'<tr><td class="{cls}">{_d(d["date"]):%a %d %b}</td><td>{days} days</td>'
                       f'<td>{e(d.get("title",""))}{flag}</td><td class="muted">{e(d.get("matter",""))}</td></tr>')
    dl = (f'<div class="scroll"><table><tr><th>Date</th><th>In</th><th>What</th><th>Matter</th></tr>'
          f'{"".join(dl_rows)}</table></div>') if dl_rows else '<div class="item muted">No dated obligations found in the next six weeks.</div>'

    wk_rows = []
    for d in a["week"]:
        load = "bad" if d["hours"] >= 6 else "warn" if d["hours"] >= 4 else "good"
        ext = ", ".join(x["subject"] for x in d["external"][:4])
        cl = f'<br><span class="bad">Clash: {e(d["clashes"][0][0])} / {e(d["clashes"][0][1])}</span>' if d["clashes"] else ""
        ad = f'<br><span class="muted">All day: {e(", ".join(d["all_day"]))}</span>' if d["all_day"] else ""
        wk_rows.append(f'<tr><td>{e(d["label"])}</td><td class="{load}">{d["hours"]} h · {d["meetings"]}</td>'
                       f'<td>{e(ext) or "<span class=muted>—</span>"}{cl}{ad}</td></tr>')
    week = (f'<div class="scroll"><table><tr><th>Day</th><th>Booked</th><th>External meetings</th></tr>'
            f'{"".join(wk_rows)}</table></div>')

    prep = [f'<div class="item"><b>{e(x.get("meeting",""))} <span class="muted">· {e(x.get("day",""))}</span></b>'
            f'<span>{e(x.get("what",""))}</span></div>' for x in p.get("prep_needed", [])]
    blocks = [f'<tr><td>{e(x.get("day",""))}</td><td>{int(x.get("minutes") or 0)} min</td>'
              f'<td>{e(x.get("task",""))}<br><span class="muted">{e(x.get("why",""))}</span></td></tr>'
              for x in p.get("diary_blocks", [])]
    blocks_html = (f'<div class="scroll"><table><tr><th>Day</th><th>Time</th><th>For</th></tr>{"".join(blocks)}</table></div>'
                   f'<p class="muted" style="font-size:13px">Book them with <b>Plan diary blocks</b> in the Control Room.</p>') \
        if blocks else '<div class="item muted">No blocks suggested.</div>'

    slipped = [f'<div class="item"><b>{e(t["title"])}</b><span>'
               f'{"Carried from last week · " if t["carried"] else ""}'
               f'{t["age"]} days old' + (f' · untouched {t["untouched"]} days' if t.get("untouched") else "")
               + '</span></div>' for t in a["slipped"][:15]]
    risks = [f'<div class="item"><b>{e(x.get("risk",""))}</b><span>{e(x.get("action",""))}</span></div>'
             for x in p.get("risks", [])]
    quiet_m = [f'<div class="item"><b>{e(n)}</b><span>Last activity {e(m.get("last_active",""))}'
               f'{" · next step: " + e(m["next_step"]) if m.get("next_step") else ""}</span></div>'
               for n, m in a["quiet"][:12]]
    confirm = [f'<div class="item"><b>{e(c.get("question",""))}</b><span>{e(c.get("matter",""))}</span></div>'
               for c in patch.get("confirm", [])]
    done = "".join(f"<li>{e(t['title'])}</li>" for t in a["done"])

    return f"""<!doctype html><html lang="en-AU"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex">
<title>Weekly Plan {wk:%d %b}</title><style>{CSS}</style></head><body><div class="wrap">
<header><small>State Gas · Weekly plan</small><h1>Week of {wk:%A %d %B %Y}</h1></header>
<p class="lead">{e(p.get("headline","The planning step did not return a summary this week."))}</p>
<section><h2>Priorities</h2><div class="card">{items(pri)}</div></section>
<section><h2>Deadlines · next six weeks</h2><div class="card">{dl}</div></section>
<section><h2>Next week's shape</h2><div class="card">{week}</div></section>
<section><h2>Prepare for</h2><div class="card">{items(prep)}</div></section>
<section><h2>Proposed diary blocks</h2><div class="card">{blocks_html}</div></section>
<section><h2>What slipped</h2><div class="card">{items(slipped)}</div></section>
<section><h2>Risks</h2><div class="card">{items(risks)}</div></section>
<section><h2>Matters gone quiet · 14+ days</h2><div class="card">{items(quiet_m)}</div></section>
<section><h2>Please confirm</h2><div class="card">{items(confirm)}</div></section>
<section><h2>Knowledge base this week</h2><div class="card"><div class="item">
<span>{stats["updated"]} matters updated · {stats["new"]} new · {stats["closed"]} closed · {stats["deadlines"]} deadlines added ·
{stats["emails"]} emails, {len(a["open"])} open tasks and {sum(d["meetings"] for d in a["week"])} meetings read</span></div>
{'<div class="item"><b>Board update drafted</b><span>' + str(stats["board"]) + ' matters, waiting in Outlook Drafts. Review notes are in the knowledge folder under board/drafts.</span></div>' if stats.get("board") is not None else ''}
<details><summary>Done this week ({len(a["done"])})</summary><ul>{done}</ul></details></div></section>
<footer>Generated {datetime.datetime.now(AEST):%a %d %b %Y %H:%M} AEST. A copy is kept in your
state gas knowledge folder under weekly.</footer></div></body></html>"""


def encrypt_page(html_text: str, title: str) -> tuple:
    """AES-256-GCM. The key travels only in the link's #fragment, never to a server."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    key, nonce = secrets.token_bytes(32), secrets.token_bytes(12)
    ct = AESGCM(key).encrypt(nonce, html_text.encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode()
    k = base64.urlsafe_b64encode(key).decode().rstrip("=")
    page = f"""<!doctype html><html lang="en-AU"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex">
<title>{title}</title><style>body{{font:15px Aptos,"Segoe UI",system-ui,sans-serif;padding:40px 16px;
color:#6B6264;background:#F6F4F3;text-align:center}}</style></head><body><p id="m">Opening your plan…</p><script>
(async()=>{{const m=document.getElementById("m");try{{
const k=new URLSearchParams(location.hash.slice(1)).get("k");if(!k)throw 0;
const b=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
const key=await crypto.subtle.importKey("raw",b(k.replace(/-/g,"+").replace(/_/g,"/")+"=".repeat((4-k.length%4)%4)),"AES-GCM",false,["decrypt"]);
const pt=await crypto.subtle.decrypt({{name:"AES-GCM",iv:b("{b64(nonce)}")}},key,b("{b64(ct)}"));
document.open();document.write(new TextDecoder().decode(pt));document.close();
}}catch(e){{m.textContent="This plan opens from the link in your Sunday email. A copy is also in your state gas knowledge folder, under weekly.";}}}})();
</script></body></html>"""
    return page, k


def send_email(token, to_addr, a, p, url) -> None:
    import requests
    from html import escape as e
    pri = "".join(f"<li><b>{e(x.get('title',''))}</b> — {e(x.get('why',''))}</li>" for x in p.get("priorities", [])[:6])
    due = "".join(f"<li>{_d(d['date']):%a %d %b}: {e(d.get('title',''))}</li>" for d in a["deadlines"][:5])
    body = f"""<div style="font-family:Aptos,Segoe UI,sans-serif;font-size:11pt;color:#1A1A1A">
<p>{e(p.get('headline',''))}</p>
<p><a href="{url}" style="color:#A51C30;font-weight:bold">Open the full weekly plan</a></p>
<p><b>Priorities</b></p><ol>{pri}</ol>
{'<p><b>Next deadlines</b></p><ul>' + due + '</ul>' if due else ''}
<p style="color:#6B6264;font-size:9pt">The link carries the key that unlocks the page; keep this email to yourself.
A copy is saved in your state gas knowledge folder under weekly.</p></div>"""
    msg = {"message": {"subject": f"Weekly plan — week of {a['monday']:%a %d %b}",
                       "body": {"contentType": "HTML", "content": body},
                       "toRecipients": [{"emailAddress": {"address": to_addr}}]},
           "saveToSentItems": False}
    r = requests.post("https://graph.microsoft.com/v1.0/me/sendMail",
                      headers={"Authorization": f"Bearer {token}"}, json=msg, timeout=60)
    r.raise_for_status()


# ═══════════════════════════════════════════════════════════════════════════
# THE RUN
# ═══════════════════════════════════════════════════════════════════════════

_REMOTE = None

STORE_FILES = ["matters.json", "people.json", "facts.json", "company_config.json",
               "deadlines.json", "pillars.json", "inbox/inbox_state.json", "instructions.md"]


def cmd_run(args):
    today = datetime.datetime.now(AEST).date()
    say(f"Weekly plan run - {today:%a %d %b %Y}")
    token = get_token()
    say("Signed in")

    global _REMOTE
    remote = None
    if CLOUD:
        from core.onedrive import OneDriveStore, DEFAULT_PATH
        tmp = Path(tempfile.mkdtemp(prefix="store_"))
        remote = OneDriveStore(token, os.environ.get("STORE_DRIVE_PATH", DEFAULT_PATH), tmp)
        for f in STORE_FILES:
            remote.pull(f, required=f == "matters.json")
        for f in ("board/lessons.json", "board/directors.json"):
            remote.pull(f)
        cutoff = (today - datetime.timedelta(days=30)).isoformat()
        for n in remote.list("board/drafts"):
            if n.endswith(".json") and n[:10] >= cutoff:
                remote.pull(f"board/drafts/{n}")
        if args.only in ("pack", "brief"):
            for n in remote.list("board/briefs"):
                remote.pull(f"board/briefs/{n}")
            remote.pull("board/style/style_guide.md")
            if args.meeting:
                remote.pull(f"board/style/style_guide_before_{args.meeting}.md")
            remote.pull("board/pack_lessons.json")
            remote.pull("board/pack_guidance.md")
            for n in remote.list("board/minutes"):
                remote.pull(f"board/minutes/{n}")
            remote.pull("board/notes_for_next_meeting.md")
            for n in remote.list("board/packs/drafts"):
                if n.endswith("spec.json"):
                    remote.pull(f"board/packs/drafts/{n}")
            for n in remote.list("board/sent"):
                remote.pull(f"board/sent/{n}")
            decks = sorted(n for n in remote.list("board/packs")
                           if n.endswith(".pptx") or n.endswith(".deck.json"))
            for n in decks[-3:]:
                remote.pull(f"board/packs/{n}")
        snaps = sorted(n for n in remote.list("snapshots") if n.endswith(".json"))
        prev_name = next((n for n in reversed(snaps) if n[:10] < today.isoformat()), None)
        if prev_name:
            remote.pull(f"snapshots/{prev_name}")
        os.environ["STATEGAS_STORE_DIR"] = str(tmp)
        _REMOTE = remote
        say(f"Knowledge store copied down ({len(remote.hashes)} files)")

    with quiet():
        from core.config import STORE_DIR, VOICE_RULES, COMPANY_NAME, MD_NAME, MD_TITLE
        from core.knowledge import CompanyKnowledge
        from core import graph
        from core.llm import get_client
    store = Path(STORE_DIR)
    load = lambda rel, empty: json.loads((store / rel).read_text(encoding="utf-8")) \
        if (store / rel).exists() else empty

    snaps_local = sorted((store / "snapshots").glob("*.json")) if (store / "snapshots").exists() else []
    prev = next((json.loads(p.read_text(encoding="utf-8")) for p in reversed(snaps_local)
                 if p.stem < today.isoformat()), None)

    if args.only in ("pack", "brief"):
        return run_pack(args, token, store, remote, today)

    # 1. read the week
    with quiet():
        emails = graph.run_scan(token, args.days, max_per_folder=75, max_total=400,
                                body_chars=1200, label="WEEKLY PLAN")
    tasks = fetch_tasks(token)
    start = datetime.datetime.combine(today - datetime.timedelta(days=7), datetime.time(), AEST)
    events = fetch_calendar(token, start, start + datetime.timedelta(days=22))
    say(f"Read {len(emails)} emails, {len(tasks)} tasks, {len(events)} calendar entries")

    ck = CompanyKnowledge()
    deadlines = load("deadlines.json", {"deadlines": []}).get("deadlines", [])
    inbox_state = load("inbox/inbox_state.json", {})
    instr_path = store / "instructions.md"
    instructions = instr_path.read_text(encoding="utf-8") if instr_path.exists() else ""
    client = get_client()

    def do_board():
        import board_weekly as bw
        result = None
        try:
            with quiet():
                result = bw.make_update(token, client, store, emails, args.days, VOICE_RULES,
                                        COMPANY_NAME, MD_NAME, MD_TITLE, instructions, say=lambda m: None)
            say(f"Board update draft in Outlook: {len(result['items'])} matters, "
                f"{result['sent_found']} past updates read, {result['lessons_added']} new lessons")
        except Exception as ex:
            say(f"!! Board update step failed: {type(ex).__name__}")
            if not CLOUD:
                raise
        if remote:
            for f in ("board/lessons.json", "board/directors.json"):
                remote.push(f)
            for sub in ("board/drafts", "board/sent"):
                d = store / sub
                if d.exists():
                    for fp in d.iterdir():
                        rel = f"{sub}/{fp.name}"
                        remote.push(rel, force_new=rel not in remote.hashes)
        return result

    if args.only == "board":
        do_board()
        say("Done (board update only)")
        return

    # 2. bring the knowledge up to date
    a0 = analyse(tasks, events, prev, deadlines, ck.data["matters"], today)
    patch = consolidate(client, emails, a0, inbox_state, instructions, today)
    before = set(ck.data["matters"])
    ck.apply_patch({k: patch.get(k, []) for k in ("matter_updates", "close_matter_names",
                                                  "people_updates", "new_facts")})
    ck.prune(); ck.save()
    deadlines, added = merge_deadlines(deadlines, patch.get("deadlines"), today)
    (store / "deadlines.json").write_text(json.dumps({"deadlines": deadlines}, indent=2,
                                                     ensure_ascii=False), encoding="utf-8")
    stats = {"updated": len(patch.get("matter_updates", [])),
             "new": len(set(ck.data["matters"]) - before),
             "closed": len(patch.get("close_matter_names", [])),
             "deadlines": added, "emails": len(emails)}
    say(f"Knowledge updated: {stats['updated']} matters, {stats['new']} new, "
        f"{stats['closed']} closed, {added} deadlines added")
    board = do_board() if args.only == "all" else None

    # 3. plan the week
    a = analyse(tasks, events, prev, deadlines, ck.data["matters"], today)
    p = plan(client, a, instructions, patch.get("confirm", []), today)
    say(f"Plan written: {len(p.get('priorities', []))} priorities, "
        f"{len(p.get('diary_blocks', []))} diary blocks")
    stats["board"] = len(board["items"]) if board else None
    html_text = render(a, p, patch, stats, today)

    # 4. keep it: plan + this week's task snapshot in the store
    (store / "weekly").mkdir(exist_ok=True)
    (store / "snapshots").mkdir(exist_ok=True)
    stamp = today.isoformat()
    (store / "weekly" / f"{stamp}.html").write_text(html_text, encoding="utf-8")
    (store / "snapshots" / f"{stamp}.json").write_text(json.dumps(
        {"date": stamp, "open": [{"id": t["id"], "title": t["title"][:120]} for t in a["open"]]},
        indent=2, ensure_ascii=False), encoding="utf-8")

    if remote:
        pushed = [f for f in ["matters.json", "people.json", "facts.json", "deadlines.json"] if remote.push(f)]
        remote.push(f"weekly/{stamp}.html", force_new=True)
        remote.push(f"snapshots/{stamp}.json", force_new=True)
        say(f"Knowledge store saved ({len(pushed)} files changed, plan and snapshot added)")

    # 5. deliver
    if CLOUD or args.email:
        page, key = encrypt_page(html_text, f"Weekly Plan {a['monday']:%d %b}")
        PUBLISH_DIR.mkdir(exist_ok=True)
        (PUBLISH_DIR / f"{stamp}.html").write_text(page, encoding="utf-8")
        url = f"{SITE}/weekly/{stamp}.html#k={key}"
        to_addr = os.environ.get("BRIEFING_EMAIL_TO", "").strip()
        if to_addr:
            send_email(token, to_addr, a, p, url)
            say("Email sent")
        else:
            say("!! No email address set (BRIEFING_EMAIL_TO)")
    if not CLOUD:
        webbrowser.open((store / "weekly" / f"{stamp}.html").as_uri())
        say(f"Opened the plan (saved in the knowledge folder under weekly/{stamp}.html)")

    if CLOUD and CACHE_FILE.exists() and os.environ.get("GH_SECRETS_TOKEN"):
        if CACHE_FILE.read_text(encoding="utf-8") != os.environ.get("WEEKLY_TOKEN_CACHE_ORIG", ""):
            # The refreshed sign-in goes to every cloud process, which keeps the
            # morning briefing's sign-in alive too (no more 90-day re-sign-in).
            failed = share_signin(CACHE_FILE.read_text(encoding="utf-8"))
            say("Sign-in refreshed for all processes" if not failed
                else f"!! Could not save the refreshed sign-in ({len(failed)} place(s))")
    if remote and remote.conflicts:
        say(f"!! {len(remote.conflicts)} file(s) left as they were - changed in OneDrive during the run")
    say("Done")


def run_pack(args, token, store, remote, today):
    """Monthly board pack: deck + Word bridge, into board/packs/drafts."""
    import board_pack
    from core import graph
    from core.llm import get_client
    days = args.days if args.days != 7 else 30
    with quiet():
        emails = graph.run_scan(token, days, max_per_folder=75, max_total=400,
                                body_chars=1200, label="BOARD PACK")
    start = datetime.datetime.combine(today, datetime.time(), AEST)
    events = fetch_calendar(token, start, start + datetime.timedelta(days=45))
    say(f"Read {len(emails)} emails and {len(events)} calendar entries")
    try:                                   # refresh the record of what you told the board
        import board_weekly as bw
        with quiet():
            sent = bw.fetch_sent_updates(token, store, days=max(days + 60, 180))
        say(f"Read {len(sent)} past board updates and notes")
        if remote:
            d = store / "board" / "sent"
            for fp in (d.iterdir() if d.exists() else []):
                rel = f"board/sent/{fp.name}"
                remote.push(rel, force_new=rel not in remote.hashes)
    except Exception as ex:
        say(f"!! Could not refresh past board updates ({type(ex).__name__})")
    if args.only == "brief":
        with quiet():
            out = board_pack.make_brief(token, get_client(), store, emails, events, say=lambda m: None,
                                        prior_name=args.prior or None, meeting=args.meeting or None,
                                        model=os.environ.get("BOARD_PACK_MODEL") or None)
        say("Board pack brief prepared")
        if remote:
            for key in ("json", "md"):
                remote.push(out[key].relative_to(store).as_posix(), force_new=True)
            for rel in ("board/pack_lessons.json", "board/notes_for_next_meeting.md"):
                remote.push(rel, force_new=rel not in remote.hashes)
            d = store / "board" / "notes_archive"
            for fp in (d.glob("*") if d.exists() else []):
                rel = f"board/notes_archive/{fp.name}"
                if rel not in remote.hashes:
                    remote.push(rel, force_new=True)
            say("Saved to the knowledge folder: board/briefs")
        say("Done (brief)")
        return
    with quiet():
        out = board_pack.make_pack(token, get_client(), store, emails, events, say=lambda m: None,
                                   prior_name=args.prior or None, meeting=args.meeting or None,
                                   model=os.environ.get("BOARD_PACK_MODEL") or None)
    say("Board pack drafted")
    if remote:
        for key in ("pptx", "docx", "spec"):
            rel = out[key].relative_to(store).as_posix()
            remote.push(rel, force_new=True)
        for rel in ("board/pack_lessons.json", "board/notes_for_next_meeting.md"):
            remote.push(rel, force_new=rel not in remote.hashes)
        for sub in ("board/packs", "board/notes_archive"):
            d = store / sub
            for fp in (d.glob("*") if d.exists() else []):
                rel = f"{sub}/{fp.name}"
                if fp.is_file() and rel not in remote.hashes:
                    remote.push(rel, force_new=True)
        say("Saved to the knowledge folder: board/packs/drafts")
    say("Done (board pack)")


def main():
    ap = argparse.ArgumentParser(description="Sunday weekly planning run")
    ap.add_argument("command", choices=["setup", "run"])
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--email", action="store_true")
    ap.add_argument("--only", choices=["all", "board", "plan", "pack", "brief"], default="all")
    ap.add_argument("--prior", default="", help="board pack: roll forward from this deck (file name prefix)")
    ap.add_argument("--meeting", default="", help="board pack: meeting date YYYY-MM-DD")
    args = ap.parse_args()
    try:
        (cmd_setup if args.command == "setup" else cmd_run)(args)
    except SystemExit:
        raise
    except Exception as ex:
        # Type only: the message can carry company content and cloud logs are public.
        if CLOUD:
            print(f"  !! Run failed: {type(ex).__name__}", flush=True)
            # The full detail goes to the private knowledge folder, not the public log.
            try:
                import traceback
                if _REMOTE:
                    log = _REMOTE.local / "logs" / "weekly_plan_error.txt"
                    log.parent.mkdir(exist_ok=True)
                    log.write_text(traceback.format_exc(), encoding="utf-8")
                    _REMOTE.push("logs/weekly_plan_error.txt", force_new=True)
                    print("  Details saved to the knowledge folder: logs/weekly_plan_error.txt", flush=True)
            except Exception:
                pass
            sys.exit(1)
        raise


if __name__ == "__main__":
    main()
