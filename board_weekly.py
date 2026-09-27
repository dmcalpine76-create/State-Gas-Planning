"""
board_weekly.py - the weekly board update, rebuilt on the knowledge store.

What it does differently from the old version:
  * starts from the matters that moved this week (knowledge store), with the
    week's email as evidence, instead of summarising raw email from scratch
  * reads your sent board updates, so it reports the change since you last
    wrote and follows up the next steps you promised
  * learns from your edits: it compares each draft with what you actually sent
    and keeps the lessons (items cut, items added, wording) for next time

Used by board_update.py (the Control Room button) and weekly_plan.py (Sunday).
Everything it keeps lives in the store's board/ folder:
  directors.json   who the directors are (learned from sent updates)
  sent/DATE.txt    your sent updates, as text
  drafts/DATE.json each draft it made, and drafts/DATE.html review notes
  lessons.json     what it has learned from your edits
"""
import json
import html
import datetime
import requests
from pathlib import Path

GRAPH = "https://graph.microsoft.com/v1.0"
BOARD_FORMAT = """FORMAT AND STANDARDS (from the MD's own updates):
- A simple numbered list of matters in priority order; no category headers.
- Each matter: a short heading, then 2-4 sentences in the MD's first person.
- Include the MD's judgment and opinion, not just facts.
- Every matter ends with the specific next step and its timing.
- Strict board filter: leave out matters managed operationally that the board
  does not need to know about, and anything with no real change.
- Collegial but professional; the directors know the MD personally.
- Do not include approaches or information the MD has judged not credible."""


def _today():
    return datetime.date.today()


def _read(p: Path, empty):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return empty


def _write(p: Path, data):
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)


# ── sent updates ────────────────────────────────────────────────────────────

def fetch_sent_updates(token: str, store: Path, days: int = 180) -> list:
    """Your sent board updates, newest first. Learns the directors as it goes."""
    h = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.body-content-type="text"'}
    me = requests.get(f"{GRAPH}/me", headers=h, params={"$select": "mail,userPrincipalName"},
                      timeout=30).json()
    mine = (me.get("mail") or me.get("userPrincipalName") or "").lower()
    since = (datetime.datetime.utcnow() - datetime.timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url, params, msgs = f"{GRAPH}/me/mailFolders/sentitems/messages", {
        "$filter": f"sentDateTime ge {since}", "$orderby": "sentDateTime desc", "$top": "100",
        "$select": "subject,sentDateTime,toRecipients,ccRecipients,body"}, []
    while url and len(msgs) < 600:
        r = requests.get(url, headers=h, params=params, timeout=60)
        r.raise_for_status()
        d = r.json()
        msgs += d.get("value", [])
        url, params = d.get("@odata.nextLink"), None

    dir_file = store / "board" / "directors.json"
    directors = set(_read(dir_file, {"emails": []}).get("emails", []))
    recips = lambda m: {x["emailAddress"]["address"].lower()
                        for x in m.get("toRecipients", []) + m.get("ccRecipients", [])
                        if x.get("emailAddress", {}).get("address")} - {mine}
    for m in msgs:                                  # learn directors from titled updates
        if "board update" in (m.get("subject") or "").lower():
            directors |= recips(m)
    if directors:
        _write(dir_file, {"emails": sorted(directors), "updated": _today().isoformat()})

    out = []
    for m in msgs:
        subj = (m.get("subject") or "")
        if subj.lower().startswith(("re:", "fw:", "fwd:")):
            continue
        rc = recips(m)
        if "board update" in subj.lower() or (directors and len(rc & directors) >= 2
                                              and len(rc & directors) >= 0.6 * len(rc)):
            text = (m.get("body") or {}).get("content", "")
            out.append({"date": m["sentDateTime"][:10], "subject": subj, "text": text[:12000]})
    for u in out[:12]:                              # keep a readable history
        f = store / "board" / "sent" / f"{u['date']}.txt"
        if not f.exists():
            f.parent.mkdir(parents=True, exist_ok=True)
            f.write_text(f"{u['subject']}\n\n{u['text']}", encoding="utf-8")
    return out


# ── learning from your edits ────────────────────────────────────────────────

def learn_from_edits(client, store: Path, sent: list) -> int:
    from core.llm import call_json
    lf = store / "board" / "lessons.json"
    lessons = _read(lf, {"lessons": [], "learned_from": []})
    done = set(lessons.get("learned_from", []))
    new = 0
    for dfile in sorted((store / "board" / "drafts").glob("*.json")):
        d = dfile.stem
        if d in done:
            continue
        draft = _read(dfile, {})
        after = [u for u in sent if d <= u["date"] <= (datetime.date.fromisoformat(d)
                                                        + datetime.timedelta(days=10)).isoformat()]
        if not after:
            if (_today() - datetime.date.fromisoformat(d)).days > 10:
                done.add(d)                          # never sent: nothing to learn
            continue
        actual = sorted(after, key=lambda u: u["date"])[0]
        prompt = f"""Below is a board update DRAFT written for a Managing Director, and the
version he actually SENT to his directors. Work out what he changed and why,
as short, reusable rules for writing his next update. Consider: matters he cut
(and what that says about his relevance filter), matters he added, facts he
corrected, tone and wording, length, structure, how he states next steps.

Only rules that will generalise. Ignore changes that are purely about this
week's facts. At most 6 rules. If he sent it essentially unchanged, return [].

DRAFT:
{draft.get("text", "")[:10000]}

SENT:
{actual["text"][:10000]}

Respond with JSON only: {{"lessons": ["rule", ...]}}"""
        out = call_json(client, prompt, max_tokens=1500, label="learn from edits") or {}
        have = {x["lesson"].lower() for x in lessons["lessons"]}
        for rule in out.get("lessons", [])[:6]:
            if rule and rule.lower() not in have:
                lessons["lessons"].append({"lesson": rule, "added": _today().isoformat(), "from": d})
                new += 1
        done.add(d)
    lessons["lessons"] = lessons["lessons"][-30:]
    lessons["learned_from"] = sorted(done)
    _write(lf, lessons)
    return new


# ── the draft ───────────────────────────────────────────────────────────────

def generate(client, store: Path, emails: list, days: int, sent: list, voice_rules: str,
             instructions: str = "") -> dict:
    from core.llm import call_json
    from core.graph import build_email_context
    today = _today()
    since = (today - datetime.timedelta(days=days)).isoformat()
    matters = _read(store / "matters.json", {"matters": {}}).get("matters", {})
    moved = {n: m for n, m in matters.items() if (m.get("last_active") or "") >= since}
    live = {n: {k: m.get(k) for k in ("description", "next_step", "next_step_due", "status")}
            for n, m in matters.items() if m.get("status", "active") == "active"}
    horizon = (today + datetime.timedelta(days=30)).isoformat()
    deadlines = [d for d in _read(store / "deadlines.json", {"deadlines": []}).get("deadlines", [])
                 if today.isoformat() <= (d.get("date") or "") <= horizon]
    lessons = [x["lesson"] for x in _read(store / "board" / "lessons.json", {"lessons": []})["lessons"]]
    last = sent[:2]

    prompt = f"""You are drafting the weekly board update email that the Managing Director sends
to his directors. Today is {today:%A %d %B %Y}; the update covers {since} to today.

{voice_rules}

{BOARD_FORMAT}

LESSONS FROM HOW HE EDITED EARLIER DRAFTS (follow these closely):
{json.dumps(lessons, ensure_ascii=False) if lessons else "(none yet)"}

STANDING INSTRUCTIONS: {instructions[:3000] or "(none)"}

WHAT HE TOLD THE BOARD MOST RECENTLY (do not repeat unchanged news; report the change
since then, and follow up any next step he promised):
{json.dumps(last, ensure_ascii=False)[:14000] if last else "(no previous update found)"}

MATTERS THAT MOVED THIS PERIOD (company knowledge):
{json.dumps(moved, ensure_ascii=False)[:20000]}

ALL ACTIVE MATTERS (for context):
{json.dumps(live, ensure_ascii=False)[:12000]}

DATED OBLIGATIONS IN THE NEXT 30 DAYS: {json.dumps(deadlines, ensure_ascii=False)[:4000]}

THIS PERIOD'S EMAIL (evidence):
{build_email_context(emails, days)[:80000]}

Write the update. Respond with JSON only:
{{"items":[{{"matter":"short heading","update":"2-4 sentences, first person, judgment, ends with next step and timing",
   "follows_up":"what he said last time that this updates, or empty","sources":["email subject or matter name"]}}],
 "omitted":[{{"matter":"","why":"why it is not board-level this week"}}]}}
Items in priority order, most important first."""
    out = call_json(client, prompt, max_tokens=8000, label="board update") or {}
    out.setdefault("items", [])
    return out


def render_email(items: list, md_name: str, md_title: str) -> tuple:
    """Plain, personal email: Aptos 11pt, numbered bold underlined headings."""
    s = "font-family:Aptos,'Aptos (Body)',Calibri,Arial,sans-serif;font-size:11pt;color:#1a1a1a;line-height:1.5;"
    e = html.escape
    parts = [f'<div style="{s}"><p style="{s}">Dear Directors</p>']
    text = ["Dear Directors", ""]
    for i, it in enumerate(items, 1):
        parts.append(f'<p style="{s}margin:16px 0 4px 0;"><b><u>{i}. {e(it.get("matter",""))}</u></b></p>'
                     f'<p style="{s}margin:0;">{e(it.get("update",""))}</p>')
        text += [f"{i}. {it.get('matter','')}", it.get("update", ""), ""]
    parts.append(f'<p style="{s}margin-top:20px;">Happy to take any questions</p>'
                 f'<p style="{s}">Kind regards</p><p style="{s}">{e(md_name)}<br>{e(md_title)}</p></div>')
    text += ["Happy to take any questions", "", "Kind regards", "", md_name, md_title]
    return "\n".join(parts), "\n".join(text)


def render_review(out: dict, days: int) -> str:
    """Private review notes: what each item rests on, and what was left out."""
    e = html.escape
    rows = "".join(f"<li><b>{e(i.get('matter',''))}</b>"
                   f"{'<br><i>Follows up: ' + e(i['follows_up']) + '</i>' if i.get('follows_up') else ''}"
                   f"<br><small>Sources: {e('; '.join(i.get('sources', [])))}</small></li>"
                   for i in out.get("items", []))
    om = "".join(f"<li><b>{e(o.get('matter',''))}</b> – {e(o.get('why',''))}</li>" for o in out.get("omitted", []))
    return (f"<!doctype html><meta charset='utf-8'><title>Board update review notes</title>"
            f"<body style='font:15px Aptos,Segoe UI,sans-serif;max-width:760px;margin:24px auto;padding:0 16px'>"
            f"<h2>Board update draft – review notes</h2><p>{days}-day window ending {_today():%d %B %Y}.</p>"
            f"<h3>In the draft</h3><ol>{rows}</ol><h3>Left out</h3><ul>{om or '<li>Nothing</li>'}</ul></body>")


def create_draft(token: str, subject: str, html_body: str, to: list) -> None:
    body = {"subject": subject, "body": {"contentType": "HTML", "content": html_body}}
    if to:
        body["toRecipients"] = [{"emailAddress": {"address": a}} for a in to]
    r = requests.post(f"{GRAPH}/me/messages", headers={"Authorization": f"Bearer {token}"},
                      json=body, timeout=60)
    r.raise_for_status()


def make_update(token, client, store: Path, emails: list, days: int, voice_rules: str,
                company: str, md_name: str, md_title: str, instructions: str = "",
                say=print) -> dict:
    """The whole thing: learn, draft, save, and leave it in Outlook Drafts."""
    store = Path(store)
    sent = fetch_sent_updates(token, store)
    learned = learn_from_edits(client, store, sent)
    out = generate(client, store, emails, days, sent, voice_rules, instructions)
    html_body, text = render_email(out["items"], md_name, md_title)
    stamp = _today().isoformat()
    _write(store / "board" / "drafts" / f"{stamp}.json",
           {"date": stamp, "days": days, "items": out["items"], "omitted": out.get("omitted", []),
            "text": text})
    (store / "board" / "drafts" / f"{stamp}.html").write_text(render_review(out, days), encoding="utf-8")
    directors = _read(store / "board" / "directors.json", {"emails": []}).get("emails", [])
    subject = f"{company} — Board Update {_today():%d %B %Y}"
    create_draft(token, subject, html_body, [])   # you address and send it yourself
    say(f"Board update: {len(out['items'])} matters drafted, {len(out.get('omitted', []))} left out, "
        f"{len(sent)} past updates read, {learned} new lessons from your edits")
    return {"items": out["items"], "text": text, "html": html_body, "subject": subject,
            "sent_found": len(sent), "lessons_added": learned, "stamp": stamp}
