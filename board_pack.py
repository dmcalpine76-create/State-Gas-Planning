"""
board_pack.py - the monthly board pack: a complete PowerPoint in the MD's own
template, plus a Word bridge document with the reasoning and sources.

How it works
  * Reads the previous final deck(s) saved in the knowledge store (board/packs)
    so each month rolls forward from what the board last saw.
  * Evidence, in priority order: last deck, the weekly updates and notes sent to
    directors since, the company knowledge (matters, deadlines), then email.
  * Writes page by page, each with examples of that page type from his own
    decks and the style guide in board/style/style_guide.md.
  * Draws the deck on his template (pack_render.py) and saves both files to
    board/packs/drafts. When he saves the final deck into board/packs, the next
    run learns from it.
"""
import re
import json
import datetime
from pathlib import Path

from deck_reader import read_deck, deck_as_text
import pack_render


def _read(p: Path, empty):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return empty


def _deck_files(store: Path):
    d = store / "board" / "packs"
    files = [f for f in d.glob("*") if f.suffix == ".pptx" or f.name.endswith(".deck.json")]
    return sorted(files, key=lambda f: f.name)


def _load_deck(f: Path):
    return _read(f, []) if f.name.endswith(".deck.json") else read_deck(str(f))


def _date_of(f: Path) -> str:
    m = re.match(r"(\d{4}-\d{2}-\d{2})", f.name)
    return m.group(1) if m else "1900-01-01"


def _section(slides, *keys):
    """Text of the slides whose titles contain any of the keys."""
    keep = [s for s in slides if any(k.lower() in (s["title"] or "").lower() for k in keys)]
    return deck_as_text(keep, 14000)


def _strip_num(title):
    return re.sub(r"^\s*([A-Z]|\d+(\.\d+)*)\.?\s+", "", title or "").strip()


def _strategy_slides(slides):
    skip = ("executive summary", "operating plan", "strategic objectives", "cashflow",
            "other matters", "appendix")
    return [s for s in slides[1:] if s["title"] and not any(k in s["title"].lower() for k in skip)]


def gather(store: Path, prior_name: str = None, cutoff: str = None):
    """cutoff (ISO date) = back-test: only what was known before that date."""
    files = _deck_files(store)
    if prior_name:
        files = [f for f in files if f.name <= prior_name or f.name.startswith(prior_name)]
    if not files:
        raise FileNotFoundError("No previous board decks in board/packs")
    prior_f = files[-1]
    prior = _load_deck(prior_f)
    history = [_load_deck(f) for f in files[-3:]]
    since = _date_of(prior_f)
    sent = []
    for f in sorted((store / "board" / "sent").glob("*.txt")):
        if f.stem >= since:
            t = f.read_text(encoding="utf-8", errors="ignore")
            if "teams.microsoft.com" not in t and len(t) > 300:
                sent.append({"date": f.stem, "text": t[:6000]})
    sent.sort(key=lambda x: x["date"], reverse=True)     # newest first survives the cap
    matters = _read(store / "matters.json", {"matters": {}}).get("matters", {})
    keys = ("category", "description", "next_step", "next_step_due", "last_active")
    if cutoff:
        # A matter last touched after the cutoff has had its notes rewritten
        # with later knowledge, so it is left out rather than leak the future.
        lo = (datetime.date.fromisoformat(cutoff) - datetime.timedelta(days=90)).isoformat()
        chosen = [(n, m) for n, m in matters.items()
                  if lo <= (m.get("last_active") or "") < cutoff and m.get("status") != "closed"]
    else:
        chosen = [(n, m) for n, m in matters.items() if m.get("status", "active") == "active"]
    chosen.sort(key=lambda x: x[1].get("last_active") or "", reverse=True)
    active = {n: {k: m.get(k) for k in keys} for n, m in chosen}
    ref = cutoff or datetime.date.today().isoformat()
    deadlines = [d for d in _read(store / "deadlines.json", {"deadlines": []}).get("deadlines", [])
                 if (d.get("date") or "") >= ref and (not cutoff or (d.get("added") or "9999") < cutoff)][:30]
    style = store / "board" / "style"
    guide_p = style / f"style_guide_before_{cutoff}.md" if cutoff else style / "style_guide.md"
    if cutoff and not guide_p.exists():
        guide_p = style / "style_guide.md"
    guidance_p = store / "board" / "pack_guidance.md"
    lessons = [x["lesson"] for x in _read(store / "board" / "pack_lessons.json", {"lessons": []})["lessons"]]
    instr_p = store / "instructions.md"
    return {
        "prior_file": prior_f.name, "prior": prior, "prior_text": deck_as_text(prior, 30000),
        "history": history, "since": since, "sent": sent, "matters": active, "deadlines": deadlines,
        "guide": guide_p.read_text(encoding="utf-8") if guide_p.exists() else "",
        "lessons": [] if cutoff else lessons,
        "guidance": "" if cutoff or not guidance_p.exists() else guidance_p.read_text(encoding="utf-8")[:3000],
        "guide_file": guide_p.name,
        "instructions": instr_p.read_text(encoding="utf-8")[:3000] if instr_p.exists() else "",
    }


def _answers_block(ctx) -> str:
    a = ctx.get("answers") or {}
    if not a:
        return ""
    return ("THE MD'S DECISIONS FOR THIS REPORT - BINDING. He reviewed a proposed outline and decided the\n"
            "following. Follow them exactly: these headlines and emphasis, these summary boxes in this order,\n"
            "these strategy papers in this order, these Other matters calls, and treat his answers as fact\n"
            "(they override anything in the evidence below):\n"
            + json.dumps(a, ensure_ascii=False, indent=1)[:12000] + "\n")


def _common(ctx, meeting_date):
    return f"""You are drafting part of the monthly Executive Board Report that the Managing
Director of an ASX-listed Queensland gas company presents to his board on {meeting_date}.
You write exactly as he does. His style guide:

{ctx["guide"]}

LESSONS FROM HOW HE CHANGED EARLIER DRAFTS: {json.dumps(ctx["lessons"], ensure_ascii=False) if ctx["lessons"] else "(none yet)"}
STANDING INSTRUCTIONS: {ctx["instructions"] or "(none)"}
HIS STANDING GUIDANCE FOR THE PACK: {ctx.get("guidance") or "(none)"}

{_answers_block(ctx)}
EVIDENCE SINCE THE LAST BOARD REPORT ({ctx["since"]}), in priority order:
A. HIS OWN NOTES FOR THIS MEETING (calls, meetings, his current thinking - treat as authoritative
   and give it prominence; these are often the strategic points email does not show):
{ctx.get("notes") or "(none)"}

B. MINUTES OF THE LAST BOARD MEETING ({ctx.get("minutes_ref") or "not found"}) - what the board decided,
   asked for and actioned. Every decision and board request must be reflected (e.g. a forecast
   case the board adopted, advice the board asked for) and every action item reported back on:
{ctx.get("minutes") or "(not available)"}

C. MARKET AND POLICY CONTEXT FOR THE PERIOD (use for the headlines' external framing and the
   strategic papers where it changes the picture; do not overstate):
{ctx.get("market") or "(not available)"}

1. What he has told the directors since (weekly updates and notes):
{json.dumps(ctx["sent"], ensure_ascii=False)[:40000] or "(none found)"}

2. Company knowledge - active matters (status, next steps):
{json.dumps(ctx["matters"], ensure_ascii=False)[:22000]}

3. Dated obligations ahead: {json.dumps(ctx["deadlines"], ensure_ascii=False)[:4000]}

4. Email this period (supporting detail only - never narrate email events):
{ctx["email"][:45000]}

ALTITUDE - this is a board paper, not an email digest:
- Write at director level: what it means, management's view, what happens next, what the
  board must note or decide. Judgement and implication, not chronology.
- No staff or junior names, no contractor or counterparty contact names; name organisations
  and roles instead. No reference numbers, email mechanics ("X emailed Y"), meeting logistics.
- Never comment on individuals' behaviour or attitude.
- Short, dense bullets. If a fact is uncertain or a number must come from the model, write
  [confirm] rather than guess.
"""


def _call(client, prompt, max_tokens, label, model=None):
    from core.llm import call_json
    return call_json(client, prompt, max_tokens=max_tokens, label=label, model=model) or {}


def plan_and_summary(client, ctx, meeting_date, model=None):
    ex = "\n\n".join(_section(d, "executive summary") for d in ctx["history"])
    strat_titles = [[s["title"] for s in _strategy_slides(d)] for d in ctx["history"]]
    cf = _section(ctx["prior"], "cashflow")
    prior_pages = [_strip_num(s["title"]) for s in _strategy_slides(ctx["prior"])
                   if "continued" not in (s["title"] or "").lower()
                   and not re.match(r"^[A-Z]\.\s", s["title"])]
    prior_boxes = [t["paras"][0]["text"] for s in ctx["prior"] if "executive summary" in (s["title"] or "").lower()
                   for t in s["texts"] if t.get("heading") and t.get("paras")
                   and not re.match(r"(\d+\.|executive summary|key )", t["paras"][0]["text"].lower())]
    prompt = _common(ctx, meeting_date) + f"""
LAST BOARD REPORT (what the board saw last time - roll forward from this):
{ctx["prior_text"]}

HIS LAST EXECUTIVE SUMMARIES (match length, tone and level):
{ex[:12000]}

STRATEGY PAGES HE HAS WRITTEN IN RECENT REPORTS: {json.dumps(strat_titles, ensure_ascii=False)}

TASK: plan this month's report and write the Executive Summary and cashflow commentary.
- headlines: 4-5 bullets, each one sentence of at most 30 words - the month's most important
  strategic points, as in his summaries. Open with the external environment (policy, market,
  value markers) when something material moved, then the month's key progress.
- boxes: START from the workstream boxes in last month's Executive Summary ({json.dumps(prior_boxes, ensure_ascii=False)})
  and keep their titles unless a workstream is clearly finished or superseded. Add a box only
  when a new matter has become one of the few things the board must hold in mind. Usually 4
  (2 to 6). Each 3-5 concise bullets (under 25 words each) at board altitude.
- strategy_pages: the matters that need their own page this month because the board must
  understand, discuss or decide something. Last month's strategy pages were
  {json.dumps(prior_pages, ensure_ascii=False)} - carry each forward unless it is clearly resolved.
  Add new pages only where the evidence demands. Titles without numbers, under 55 characters.
  Put first the paper the board most needs to discuss or decide at this meeting.
- Any matter important enough for its own strategy page and still live should normally also
  have an Executive Summary box; a box whose matter is resolved should be dropped or folded
  into a headline. For each give a title in his
  style, its purpose (what the board needs from it), the format (narrative, or narrative plus
  table, or table), and image_note if a map, chart or model extract belongs on it.
- cashflow: title and 5-9 commentary bullets rolled forward from last month's, changing only
  what the evidence supports. Reflect any forecast decisions or changes the board asked for at
  the last meeting (see the minutes) - e.g. which case is now the base case. Where a number must come from the model, write it as [confirm].
- flags: private notes to the MD - risks, contradictions, things the board may ask.

Respond with JSON only:
{{"headlines":[""],"boxes":[{{"title":"","bullets":[""]}}],
 "strategy_pages":[{{"title":"","purpose":"","format":"narrative|narrative+table|table","image_note":""}}],
 "cashflow":{{"title":"","bullets":[""]}},
 "flags":[""]}}"""
    return _call(client, prompt, 9000, "board pack plan", model)


def op_plan(client, ctx, meeting_date, model=None):
    tables = [t for s in ctx["prior"] if "operating plan" in (s["title"] or "").lower() for t in s["tables"]]
    if not tables:
        return None
    grid = tables[0]
    prompt = _common(ctx, meeting_date) + f"""
LAST MONTH'S OPERATING PLAN TABLE (rows of cells, first two rows are headers):
{json.dumps(grid, ensure_ascii=False)}

TASK: roll this table forward to {meeting_date}. Keep the same six objectives and row labels.
Update Status and Actions from the evidence. Change Timing, Resources, Risks and Cost
estimates only where the evidence says so. Mark objectives "Done" when complete. Cells are
short phrases, several per cell separated by newlines, as he writes them. Leave cells
unchanged when nothing has moved.

Respond with JSON only:
{{"title":"Progress Against FY 27 Operating Plan","columns":["objective 1", "..."],
 "rows":[{{"label":"Status","cells":["", "..."]}}],"changes":["what you changed and why"]}}"""
    return _call(client, prompt, 7000, "operating plan", model)


def other_matters(client, ctx, meeting_date, page_titles, model=None):
    prior_rows = [r for s in ctx["prior"] if "other matters" in (s["title"] or "").lower()
                  for t in s["tables"] for r in t[1:]]
    prompt = _common(ctx, meeting_date) + f"""
LAST MONTH'S OTHER MATTERS TABLE (Area | Matter | Update | Next Steps | Risk | Action):
{json.dumps(prior_rows, ensure_ascii=False)}

MATTERS ALREADY COVERED ON THEIR OWN PAGES THIS MONTH (leave these out): {json.dumps(page_titles, ensure_ascii=False)}

TASK: write this month's Other matters table. Carry forward EVERY row from last month's table
(same Area and Matter wording), updating it - if nothing has moved say so briefly. Drop a row
only if it is clearly closed or now has its own page, and list any dropped row in "dropped"; add board-relevant matters from
the evidence that have no page of their own - including a row reporting back on each action item
from the last meeting's minutes that has no page of its own, with Action "Discuss" or "Decision"
where the board asked to consider it. Usually 5 to 10 rows. Update is 2-3 sentences (under
60 words) of substance and view in his voice; Next Steps under 30 words; Next Steps concrete; Risk Nil/Low/Med/High; Action Nil,
Note, Discuss or Decision.

Respond with JSON only:
{{"rows":[{{"area":"","matter":"","update":"","next_steps":"","risk":"","action":"","basis":"evidence used"}}],
 "dropped":[{{"matter":"","reason":""}}]}}"""
    return _call(client, prompt, 9000, "other matters", model)


def strategy_page(client, ctx, meeting_date, page, model=None):
    prior_match = _section(ctx["prior"], page["title"].split("–")[0].split("-")[0].strip()[:25])
    examples = "\n\n".join(deck_as_text(_strategy_slides(d)[:4], 9000) for d in ctx["history"][-2:])
    words = 200 if page.get("image_note") else 320
    prompt = _common(ctx, meeting_date) + f"""
HIS STRATEGY PAGES FROM RECENT REPORTS (match structure, length, tone - argument, not narrative):
{examples[:16000]}

THIS TOPIC IN LAST MONTH'S REPORT (if it appeared):
{prior_match or "(new this month)"}

TASK: write the page "{page["title"]}". Purpose: {page.get("purpose","")}. Format: {page.get("format","narrative")}.
Structure it as he does: the situation and what has changed, management's analysis and view,
options or implications, and a clear recommendation or ask of the board. Use sub-bullets
(level 1) for lists of factors, reasons or options. It must fit ONE slide: {words} words of
bullets at most (his pages run 220-360 words). Cut chronology before cutting judgement. Where a
figure is needed but not in the evidence, write [confirm]. If the format includes a table,
give it (header row plus rows; 3-6 columns; short cell text), as he does for partner options,
project plans by quarter, or observation/implication. Only include a table when it carries
the substance better than bullets; when it does, keep the bullets under 120 words.

Respond with JSON only:
{{"paras":[{{"level":0,"text":""}}],"table":null or {{"intro":"","header":[""],"rows":[[""]]}},
 "basis":["evidence used"],"to_confirm":["facts or numbers he must check"]}}"""
    return _call(client, prompt, 6000, f"page: {page['title'][:30]}", model)


def _words(paras):
    return sum(len((p.get("text") or "").split()) for p in paras)


def tighten(client, ctx, what, content, limit_note, model=None):
    """An editor's pass: same structure, board altitude, within the length limit."""
    prompt = f"""You are the Managing Director's editor, preparing his board paper. His style guide:

{ctx["guide"][:6000]}

Edit this {what} so it reads as he would write it to his board. Keep the JSON structure and
keys exactly. {limit_note}
- Remove staff, junior, contractor and counterparty contact names (keep organisations and
  roles), reference numbers, email and meeting mechanics, and any remark about an individual.
- Keep every judgement, number, risk and ask. Cut chronology and process detail first.
- Keep [confirm] markers. Do not add facts.

CONTENT:
{json.dumps(content, ensure_ascii=False)}

Respond with the edited JSON only."""
    out = _call(client, prompt, 6000, f"edit: {what[:30]}", model)
    return out if isinstance(out, dict) and out else content


def _meeting_iso(text: str):
    try:
        return datetime.datetime.strptime(text.strip(), "%d %B %Y").date().isoformat()
    except Exception:
        return None


def learn_from_finals(token, client, store: Path, model=None) -> int:
    """
    For each earlier draft whose meeting has now happened, find the final deck
    the MD saved into the knowledge store (board/packs, named with the meeting
    date) and record what he changed as lessons for future drafts
    (board/pack_lessons.json).
    """
    import board_sources as bs
    lf = store / "board" / "pack_lessons.json"
    book = _read(lf, {"lessons": []})
    book.setdefault("learned", [])
    drafts = sorted((store / "board" / "packs" / "drafts").glob("*spec.json"))
    done = 0
    for sp in drafts:
        if sp.name in book["learned"]:
            continue
        spec = _read(sp, {})
        miso = _meeting_iso(spec.get("meeting_date", ""))
        if not miso or miso >= datetime.date.today().isoformat():
            continue
        final = None
        for f in _deck_files(store):
            if _date_of(f) == miso:
                final = _load_deck(f)
        if final is None:
            continue
        slim = {k: spec.get(k) for k in ("exec", "strategy_pages", "cashflow", "other_matters")}
        prompt = f"""Below is a machine draft of a monthly board report and the final version the
Managing Director actually presented. Work out what he changed and why, as reusable lessons for
future drafts: what he chose to cover or leave out, how he framed the month, which papers he
wrote, what level of detail and tone he used, what evidence he drew on that the draft lacked.
Lessons must be general rules, not this month's facts. Then merge them with the existing lessons
into one list of at most 25, most important first, removing duplicates.

EXISTING LESSONS: {json.dumps([x["lesson"] for x in book["lessons"]], ensure_ascii=False)}

DRAFT (JSON): {json.dumps(slim, ensure_ascii=False)[:30000]}

FINAL AS PRESENTED:
{deck_as_text(final, 30000)}

Respond with JSON only: {{"lessons": [""], "new_this_time": [""]}}"""
        res = _call(client, prompt, 4000, "learn from final deck", model)
        if res.get("lessons"):
            book["lessons"] = [{"lesson": l, "updated": datetime.date.today().isoformat()}
                               for l in res["lessons"][:25]]
        book["learned"].append(sp.name)
        book.setdefault("log", []).append({"draft": sp.name, "meeting": miso,
                                           "new": res.get("new_this_time", [])})
        if not spec.get("_evidence", {}).get("backtest"):
            notes = store / "board" / "notes_for_next_meeting.md"
            if notes.exists() and bs.running_notes(store):
                arch = store / "board" / "notes_archive" / f"{miso}.md"
                arch.parent.mkdir(parents=True, exist_ok=True)
                arch.write_text(notes.read_text(encoding="utf-8"), encoding="utf-8")
                notes.write_text(NOTES_TEMPLATE, encoding="utf-8")
        done += 1
    if done:
        lf.parent.mkdir(parents=True, exist_ok=True)
        lf.write_text(json.dumps(book, indent=1, ensure_ascii=False), encoding="utf-8")
    return done


NOTES_TEMPLATE = """# Notes for the next board meeting

> Jot anything the board pack should know that isn't in email: calls, meetings, what you're
> thinking, what you want the board to discuss. One line each is fine. The board pack treats
> these as top priority. After the meeting they are archived and this page resets.

"""


def next_meeting(events, today):
    for e in events:
        if "board" in e["subject"].lower() and "meeting" in e["subject"].lower() and e["start"][:10] >= today.isoformat():
            return datetime.date.fromisoformat(e["start"][:10])
    return None


def write_bridge(path, spec, parts, ctx):
    """The Word bridge: same content, plus the basis and what to confirm."""
    from docx import Document
    from docx.shared import Pt
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Aptos"; st.font.size = Pt(10)
    doc.add_heading(f"Board Report draft – {spec['meeting_date']}", 0)
    doc.add_paragraph(f"Rolled forward from: {ctx['prior_file']}. Evidence since {ctx['since']}: "
                      f"{len(ctx['sent'])} notes to directors, {len(ctx['matters'])} active matters. "
                      f"Style guide: {ctx.get('guide_file','')}. "
                      f"Minutes used: {ctx.get('minutes_ref') or 'none found'}. "
                      f"Market context: {'yes' if ctx.get('market') else 'none'}. "
                      f"Your notes: {'yes' if ctx.get('notes') else 'none'}. "
                      f"Your decisions from the brief: {'yes' if ctx.get('answers') else 'none'}. "
                      "Draft for the MD's review - not for circulation.")
    if parts["plan"].get("flags"):
        doc.add_heading("Flags for you", 1)
        for f in parts["plan"]["flags"]:
            doc.add_paragraph(f, style="List Bullet")
    doc.add_heading("1. Executive Summary", 1)
    for h in spec["exec"]["headlines"]:
        doc.add_paragraph(h, style="List Bullet")
    for b in spec["exec"]["boxes"]:
        doc.add_heading(b["title"], 2)
        for x in b["bullets"]:
            doc.add_paragraph(x, style="List Bullet")
    if parts.get("op") and parts["op"].get("changes"):
        doc.add_heading("2. Operating plan – what changed", 1)
        for c in parts["op"]["changes"]:
            doc.add_paragraph(c, style="List Bullet")
    for pg in spec["strategy_pages"]:
        res = pg
        doc.add_heading(pg["title"], 1)
        for p in pg["paras"]:
            doc.add_paragraph(p["text"], style="List Bullet 2" if p.get("level") else "List Bullet")
        if res.get("to_confirm"):
            doc.add_paragraph("To confirm: " + "; ".join(res["to_confirm"])).italic = True
        if res.get("basis"):
            doc.add_paragraph("Basis: " + "; ".join(res["basis"]))
    if spec.get("other_matters"):
        doc.add_heading("Other matters – basis", 1)
        for r in spec["other_matters"]:
            doc.add_paragraph(f"{r.get('matter','')}: {r.get('basis','')}", style="List Bullet")
    if parts.get("dropped"):
        doc.add_heading("Other matters – rows dropped from last month", 1)
        for d in parts["dropped"]:
            doc.add_paragraph(f"{d.get('matter','')}: {d.get('reason','')}", style="List Bullet")
    doc.save(path)


def build_context(token, client, store: Path, emails: list, events: list, say=print,
                  prior_name: str = None, meeting: str = None, model: str = None):
    """Everything the brief and the pack draw on. Returns (ctx, md, backtest)."""
    from core.graph import build_email_context
    store = Path(store)
    today = datetime.date.today()
    md = datetime.date.fromisoformat(meeting) if meeting else (next_meeting(events, today) or today)
    backtest = md < today
    try:
        learned = learn_from_finals(token, client, store, model)
        if learned:
            say(f"Learned from {learned} final deck(s)")
    except Exception as ex:
        say(f"!! Could not learn from final decks ({type(ex).__name__})")
    ctx = gather(store, prior_name, md.isoformat() if backtest else None)
    import board_sources as bs
    until = md.isoformat() if backtest else today.isoformat()
    try:
        mins = bs.last_minutes(store, until)
        ctx["minutes"], ctx["minutes_ref"] = mins.get("text", ""), (
            f"{mins['folder']}/{mins['name']}, saved {mins['saved']}" if mins else "")
    except Exception as ex:
        ctx["minutes"], ctx["minutes_ref"] = "", ""
        say(f"!! Could not read the last minutes ({type(ex).__name__})")
    ctx["market"] = bs.market_context(client, ctx["since"], until, model)
    ctx["notes"] = "" if backtest else bs.running_notes(store)
    if backtest:   # re-drafting a past meeting: only what was known before it
        cut = md.isoformat()
        emails = [m for m in emails
                  if (m.get("datetime") or m.get("receivedDateTime") or m.get("sentDateTime") or "9999")[:10] < cut]
        ctx["sent"] = [x for x in ctx["sent"] if x["date"] < cut]
    ctx["email"] = build_email_context(emails)
    return ctx, md, backtest


def _brief_paths(store: Path, md):
    d = store / "board" / "briefs"
    k = md.isoformat()
    return d / f"{k} brief.json", d / f"{k} answers.json", d / f"{k} brief.md"


def make_brief(token, client, store: Path, emails: list, events: list, say=print,
               prior_name: str = None, meeting: str = None, model: str = None) -> dict:
    """
    Step 1 of the board pack: a proposed outline and the questions only the MD
    can answer. Saved to board/briefs/<meeting> brief.json (and .md to read).
    """
    store = Path(store)
    ctx, md, backtest = build_context(token, client, store, emails, events, say,
                                      prior_name, meeting, model)
    meeting_date = f"{md.day} {md:%B %Y}"
    prior_boxes = [t["paras"][0]["text"] for s in ctx["prior"] if "executive summary" in (s["title"] or "").lower()
                   for t in s["texts"] if t.get("heading") and t.get("paras")
                   and not re.match(r"(\d+\.|executive summary|key )", t["paras"][0]["text"].lower())]
    prior_pages = [_strip_num(s["title"]) for s in _strategy_slides(ctx["prior"])
                   if "continued" not in (s["title"] or "").lower() and not re.match(r"^[A-Z]\.\s", s["title"])]
    prior_rows = [r[:2] for s in ctx["prior"] if "other matters" in (s["title"] or "").lower()
                  for t in s["tables"] for r in t[1:]]
    prompt = _common(ctx, meeting_date) + f"""
LAST BOARD REPORT:
{ctx["prior_text"][:25000]}

LAST MONTH'S SUMMARY BOXES: {json.dumps(prior_boxes, ensure_ascii=False)}
LAST MONTH'S STRATEGY PAPERS: {json.dumps(prior_pages, ensure_ascii=False)}
LAST MONTH'S OTHER MATTERS ROWS (area, matter): {json.dumps(prior_rows, ensure_ascii=False)}

TASK: do NOT write the report. Prepare a short brief the MD will review with an assistant before
the report is drafted, so that he makes the judgement calls. Be concrete and brief.
- headlines: 5-7 candidate headlines (one sentence each) with the evidence for each.
- boxes: last month's boxes (carry / drop, with why) plus any new candidates, each with 2-4
  key points you would make.
- papers: candidate strategy papers - last month's still live plus new ones the evidence
  suggests - each with why, what the board would be asked to do, and your recommendation
  (include / optional / not needed).
- other_matters: every row from last month plus new candidates: proposal (update / drop /
  new) and a one-line proposed update.
- action_items: every decision, request and action item in the last minutes, what the
  evidence shows about it, and status (done / in progress / no evidence).
- cashflow: points the commentary needs that the evidence cannot settle.
- questions: 5-10 pointed questions only the MD can answer - gaps (matters with no evidence
  this period), judgement calls (lead paper, emphasis, what to ask the board), facts to
  confirm. Each with why it matters and, where sensible, 2-4 suggested answers.

Respond with JSON only:
{{"headlines":[{{"text":"","evidence":""}}],
 "boxes":[{{"title":"","status":"carry|drop|new","why":"","points":[""]}}],
 "papers":[{{"title":"","why":"","ask":"","recommend":"include|optional|not needed"}}],
 "other_matters":[{{"area":"","matter":"","proposal":"update|drop|new","draft":""}}],
 "action_items":[{{"item":"","found":"","status":""}}],
 "cashflow":[""],
 "questions":[{{"q":"","why":"","options":[""]}}]}}"""
    brief = _call(client, prompt, 9000, "board pack brief", model)
    if not brief:
        raise RuntimeError("brief step returned nothing")
    brief["meeting"], brief["meeting_date"] = md.isoformat(), meeting_date
    brief["evidence"] = {"prior": ctx["prior_file"], "since": ctx["since"],
                         "minutes": ctx.get("minutes_ref", ""), "market": bool(ctx.get("market")),
                         "notes": bool(ctx.get("notes")), "notes_to_directors": len(ctx["sent"]),
                         "backtest": backtest}
    bj, _, bm = _brief_paths(store, md)
    bj.parent.mkdir(parents=True, exist_ok=True)
    bj.write_text(json.dumps(brief, indent=1, ensure_ascii=False), encoding="utf-8")
    bm.write_text(_brief_md(brief), encoding="utf-8")
    return {"json": bj, "md": bm}


def _brief_md(b: dict) -> str:
    L = [f"# Board pack brief - {b.get('meeting_date','')}", "",
         f"Evidence: {json.dumps(b.get('evidence', {}))}", "", "## Questions for you"]
    for i, q in enumerate(b.get("questions", []), 1):
        L.append(f"{i}. {q.get('q','')}  _({q.get('why','')})_")
        if q.get("options"):
            L.append("   Options: " + " / ".join(q["options"]))
    L += ["", "## Candidate headlines"] + [f"- {h.get('text','')}  _[{h.get('evidence','')}]_" for h in b.get("headlines", [])]
    L += ["", "## Summary boxes"]
    for x in b.get("boxes", []):
        L.append(f"- **{x.get('title','')}** ({x.get('status','')}) - {x.get('why','')}")
        L += [f"  - {p}" for p in x.get("points", [])]
    L += ["", "## Strategy papers"] + [f"- **{p.get('title','')}** [{p.get('recommend','')}] - {p.get('why','')} Ask: {p.get('ask','')}"
                                        for p in b.get("papers", [])]
    L += ["", "## Other matters"] + [f"- {r.get('area','')} | {r.get('matter','')} [{r.get('proposal','')}] {r.get('draft','')}"
                                      for r in b.get("other_matters", [])]
    L += ["", "## Last meeting's actions"] + [f"- {a.get('item','')} - {a.get('status','')}: {a.get('found','')}"
                                              for a in b.get("action_items", [])]
    L += ["", "## Cash flow points"] + [f"- {c}" for c in b.get("cashflow", [])]
    return "\n".join(L) + "\n"


def make_pack(token, client, store: Path, emails: list, events: list, say=print,
              prior_name: str = None, meeting: str = None, model: str = None) -> dict:
    store = Path(store)
    today = datetime.date.today()
    ctx, md, backtest = build_context(token, client, store, emails, events, say,
                                      prior_name, meeting, model)
    meeting_date = f"{md.day} {md:%B %Y}"
    _, aj, _ = _brief_paths(store, md)
    ctx["answers"] = _read(aj, {})

    plan = plan_and_summary(client, ctx, meeting_date, model)
    agreed = [p for p in ctx["answers"].get("papers", []) if isinstance(p, dict) and p.get("title")]
    if agreed:        # the MD's chosen papers, in his order, replace the model's choice
        mine = {_strip_num(p.get("title", "")).lower(): p for p in plan.get("strategy_pages", [])}
        plan["strategy_pages"] = [{**mine.get(p["title"].lower(), {}), "title": p["title"],
                                   "purpose": p.get("purpose") or mine.get(p["title"].lower(), {}).get("purpose", ""),
                                   "format": p.get("format") or mine.get(p["title"].lower(), {}).get("format", "narrative"),
                                   "image_note": p.get("image_note") or mine.get(p["title"].lower(), {}).get("image_note")}
                                  for p in agreed]
    if not plan.get("boxes"):
        raise RuntimeError("planning step returned nothing")
    for pg in plan.get("strategy_pages", []):
        pg["title"] = _strip_num(pg.get("title", ""))
    ed = tighten(client, ctx, "executive summary",
                 {"headlines": plan.get("headlines", []), "boxes": plan["boxes"]},
                 "Headlines: at most 5, each one sentence under 30 words. Boxes: 3-5 bullets each, "
                 "each under 25 words.", model)
    if ed.get("boxes"):
        plan["headlines"], plan["boxes"] = ed.get("headlines", plan.get("headlines", [])), ed["boxes"]
    op = op_plan(client, ctx, meeting_date, model)
    pages_out = []
    for pg in plan.get("strategy_pages", [])[:7]:
        res = strategy_page(client, ctx, meeting_date, pg, model)
        limit = 200 if pg.get("image_note") else (120 if res.get("table") else 320)
        if res.get("paras") and _words(res["paras"]) > limit * 1.1:
            e = tighten(client, ctx, f"board page '{pg['title']}'", {"paras": res["paras"]},
                        f"The bullets must total no more than {limit} words so the page fits one slide.", model)
            if e.get("paras"):
                res["paras"] = e["paras"]
        pages_out.append(res)
    om = other_matters(client, ctx, meeting_date, [p["title"] for p in plan.get("strategy_pages", [])], model)

    spec = {"meeting_date": meeting_date,
            "exec": {"headlines": plan.get("headlines", []), "boxes": plan["boxes"][:6]},
            "op_plan": op if op and op.get("rows") else None,
            "strategy_pages": [],
            "cashflow": {**plan.get("cashflow", {}), "image_note": "cashflow forecast chart from the model"},
            "other_matters": om.get("rows", [])}
    for pg, res in zip(plan.get("strategy_pages", []), pages_out):
        if not res.get("paras"):
            continue
        t = res.get("table")
        if t and t.get("header") and t.get("rows"):
            n = len(t["header"]); first = 1.25
            t["widths"] = [first] + [(9.23 - first) / (n - 1)] * (n - 1) if n > 1 else [9.23]
        else:
            t = None
        spec["strategy_pages"].append({"title": pg["title"], "paras": res["paras"], "table": t,
                                       "image_note": pg.get("image_note") or None,
                                       "basis": res.get("basis", []), "to_confirm": res.get("to_confirm", [])})

    out_dir = store / "board" / "packs" / "drafts"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = f"{md.isoformat()} BACKTEST" if backtest else today.isoformat()
    template = next((f for f in reversed(_deck_files(store)) if f.suffix == ".pptx"), None)
    pptx_path = out_dir / f"{stamp} Board Report DRAFT.pptx"
    pack_render.build(spec, str(template), str(pptx_path))
    docx_path = out_dir / f"{stamp} Board Report DRAFT - notes.docx"
    spec["_evidence"] = {"minutes": ctx.get("minutes_ref", ""), "market": bool(ctx.get("market")),
                         "notes": bool(ctx.get("notes")), "backtest": backtest}
    write_bridge(str(docx_path), spec, {"plan": plan, "op": op, "dropped": om.get("dropped", [])}, ctx)
    (out_dir / f"{stamp} spec.json").write_text(json.dumps(spec, indent=1, ensure_ascii=False), encoding="utf-8")
    say(f"Board pack drafted: {len(spec['exec']['boxes'])} summary boxes, "
        f"{len(spec['strategy_pages'])} strategy pages, {len(spec['other_matters'])} other matters "
        f"(rolled forward from {ctx['prior_file'][:10]})")
    return {"pptx": pptx_path, "docx": docx_path, "spec": out_dir / f"{stamp} spec.json"}
