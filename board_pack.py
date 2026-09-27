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


def _strategy_slides(slides):
    skip = ("executive summary", "operating plan", "strategic objectives", "cashflow",
            "other matters", "appendix")
    return [s for s in slides[1:] if s["title"] and not any(k in s["title"].lower() for k in skip)]


def gather(store: Path, prior_name: str = None):
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
                sent.append({"date": f.stem, "text": t[:9000]})
    matters = _read(store / "matters.json", {"matters": {}}).get("matters", {})
    active = {n: {k: m.get(k) for k in ("category", "description", "next_step", "next_step_due", "last_active")}
              for n, m in matters.items() if m.get("status", "active") == "active"}
    deadlines = [d for d in _read(store / "deadlines.json", {"deadlines": []}).get("deadlines", [])
                 if (d.get("date") or "") >= datetime.date.today().isoformat()][:30]
    guide_p = store / "board" / "style" / "style_guide.md"
    lessons = [x["lesson"] for x in _read(store / "board" / "pack_lessons.json", {"lessons": []})["lessons"]]
    instr_p = store / "instructions.md"
    return {
        "prior_file": prior_f.name, "prior": prior, "prior_text": deck_as_text(prior, 30000),
        "history": history, "since": since, "sent": sent, "matters": active, "deadlines": deadlines,
        "guide": guide_p.read_text(encoding="utf-8") if guide_p.exists() else "",
        "lessons": lessons,
        "instructions": instr_p.read_text(encoding="utf-8")[:3000] if instr_p.exists() else "",
    }


def _common(ctx, meeting_date):
    return f"""You are drafting part of the monthly Executive Board Report that the Managing
Director of an ASX-listed Queensland gas company presents to his board on {meeting_date}.
You write exactly as he does. His style guide:

{ctx["guide"]}

LESSONS FROM HOW HE CHANGED EARLIER DRAFTS: {json.dumps(ctx["lessons"], ensure_ascii=False) if ctx["lessons"] else "(none yet)"}
STANDING INSTRUCTIONS: {ctx["instructions"] or "(none)"}

EVIDENCE SINCE THE LAST BOARD REPORT ({ctx["since"]}), in priority order:
1. What he has told the directors since (weekly updates and notes):
{json.dumps(ctx["sent"], ensure_ascii=False)[:30000] or "(none found)"}

2. Company knowledge - active matters (status, next steps):
{json.dumps(ctx["matters"], ensure_ascii=False)[:22000]}

3. Dated obligations ahead: {json.dumps(ctx["deadlines"], ensure_ascii=False)[:4000]}

4. Email this period (supporting detail only - never narrate email events):
{ctx["email"][:45000]}
"""


def _call(client, prompt, max_tokens, label, model=None):
    from core.llm import call_json
    return call_json(client, prompt, max_tokens=max_tokens, label=label, model=model) or {}


def plan_and_summary(client, ctx, meeting_date, model=None):
    ex = "\n\n".join(_section(d, "executive summary") for d in ctx["history"])
    strat_titles = [[s["title"] for s in _strategy_slides(d)] for d in ctx["history"]]
    cf = _section(ctx["prior"], "cashflow")
    prompt = _common(ctx, meeting_date) + f"""
LAST BOARD REPORT (what the board saw last time - roll forward from this):
{ctx["prior_text"]}

HIS LAST EXECUTIVE SUMMARIES (match length, tone and level):
{ex[:12000]}

STRATEGY PAGES HE HAS WRITTEN IN RECENT REPORTS: {json.dumps(strat_titles, ensure_ascii=False)}

TASK: plan this month's report and write the Executive Summary and cashflow commentary.
- headlines: 4-6 bullets, the month's most important strategic points.
- boxes: one per key matter this month (usually 4, anywhere from 2 to 6), each 3-6 bullets,
  titled like his (e.g. "HDNG Foundation Customer Project", "Capital Management").
- strategy_pages: the matters that need their own page this month because the board must
  understand, discuss or decide something. Carry forward pages that are still live; add new
  ones where the evidence demands; drop ones that are resolved. For each give a title in his
  style, its purpose (what the board needs from it), the format (narrative, or narrative plus
  table, or table), and image_note if a map, chart or model extract belongs on it.
- cashflow: title and 5-9 commentary bullets rolled forward from last month's, changing only
  what the evidence supports. Where a number must come from the model, write it as [confirm].
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

TASK: write this month's Other matters table. Roll forward rows that are still live, updating
them; drop rows that are closed or no longer board-relevant; add board-relevant matters from
the evidence that have no page of their own. Usually 5 to 10 rows. Update is 2-5 sentences of
substance and view in his voice; Next Steps concrete; Risk Nil/Low/Med/High; Action Nil,
Note, Discuss or Decision.

Respond with JSON only:
{{"rows":[{{"area":"","matter":"","update":"","next_steps":"","risk":"","action":"","basis":"evidence used"}}]}}"""
    return _call(client, prompt, 9000, "other matters", model)


def strategy_page(client, ctx, meeting_date, page, model=None):
    prior_match = _section(ctx["prior"], page["title"].split("–")[0].split("-")[0].strip()[:25])
    examples = "\n\n".join(deck_as_text(_strategy_slides(d)[:4], 9000) for d in ctx["history"][-2:])
    prompt = _common(ctx, meeting_date) + f"""
HIS STRATEGY PAGES FROM RECENT REPORTS (match structure, length, tone - argument, not narrative):
{examples[:16000]}

THIS TOPIC IN LAST MONTH'S REPORT (if it appeared):
{prior_match or "(new this month)"}

TASK: write the page "{page["title"]}". Purpose: {page.get("purpose","")}. Format: {page.get("format","narrative")}.
Structure it as he does: the situation and what has changed, management's analysis and view,
options or implications, and a clear recommendation or ask of the board. Use sub-bullets
(level 1) for lists of factors, reasons or options. Roughly 180-380 words of bullets. Where a
figure is needed but not in the evidence, write [confirm]. If the format includes a table,
give it (header row plus rows; 3-6 columns; short cell text), as he does for partner options,
project plans by quarter, or observation/implication.

Respond with JSON only:
{{"paras":[{{"level":0,"text":""}}],"table":null or {{"intro":"","header":[""],"rows":[[""]]}},
 "basis":["evidence used"],"to_confirm":["facts or numbers he must check"]}}"""
    return _call(client, prompt, 6000, f"page: {page['title'][:30]}", model)


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
    for pg, res in zip(spec["strategy_pages"], parts["pages"]):
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
    doc.save(path)


def make_pack(token, client, store: Path, emails: list, events: list, say=print,
              prior_name: str = None, meeting: str = None, model: str = None) -> dict:
    from core.graph import build_email_context
    store = Path(store)
    today = datetime.date.today()
    ctx = gather(store, prior_name)
    ctx["email"] = build_email_context(emails)
    md = datetime.date.fromisoformat(meeting) if meeting else (next_meeting(events, today) or today)
    meeting_date = f"{md.day} {md:%B %Y}"

    plan = plan_and_summary(client, ctx, meeting_date, model)
    if not plan.get("boxes"):
        raise RuntimeError("planning step returned nothing")
    op = op_plan(client, ctx, meeting_date, model)
    pages_out = [strategy_page(client, ctx, meeting_date, pg, model) for pg in plan.get("strategy_pages", [])[:7]]
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
                                       "image_note": pg.get("image_note") or None})

    out_dir = store / "board" / "packs" / "drafts"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = today.isoformat()
    template = next((f for f in reversed(_deck_files(store)) if f.suffix == ".pptx"), None)
    pptx_path = out_dir / f"{stamp} Board Report DRAFT.pptx"
    pack_render.build(spec, str(template), str(pptx_path))
    docx_path = out_dir / f"{stamp} Board Report DRAFT - notes.docx"
    write_bridge(str(docx_path), spec, {"plan": plan, "op": op, "pages": pages_out}, ctx)
    (out_dir / f"{stamp} spec.json").write_text(json.dumps(spec, indent=1, ensure_ascii=False), encoding="utf-8")
    say(f"Board pack drafted: {len(spec['exec']['boxes'])} summary boxes, "
        f"{len(spec['strategy_pages'])} strategy pages, {len(spec['other_matters'])} other matters "
        f"(rolled forward from {ctx['prior_file'][:10]})")
    return {"pptx": pptx_path, "docx": docx_path, "spec": out_dir / f"{stamp} spec.json"}
