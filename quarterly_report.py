"""
quarterly_report.py - the quarterly activities report (Appendix 5B narrative),
drafted from the knowledge store in the company's own Word format, plus a
verification document showing where every statement came from.

Two steps, like the board pack:
  brief  - reviews the quarter and proposes highlights, sections and the
           judgement calls (lead story, disclosures, funding message, changes to
           guidance). Saved to reports/briefs/<period> brief.json / .md.
           The MD answers the questions in a Claude session, which saves
           reports/briefs/<period> answers.json - binding on the build.
  build  - writes the draft (reports/drafts/<period> Quarterly DRAFT <date>.docx)
           and the verification document beside it.

Knowledge store layout (OneDrive "state gas knowledge"):
  reports/published/   released PDFs, named "<YYYY-MM> Quarterly Activities Report.pdf",
                       "<YYYY-MM> Half-year Report.pdf", "<YYYY-MM> Annual Report.pdf"
  reports/templates/   the latest quarterly as Word, "<YYYY-MM> Quarterly Activities Report.docx"
  reports/strategy/    the latest strategy / company deck (.pptx) - the messaging source
  reports/briefs/      brief and answers files
  reports/drafts/      drafts and verification documents
  reports/quarterly_guidance.md   standing rules, edited by the MD
Also read: board/minutes, board/packs (decks), board/sent (updates to directors),
matters.json, deadlines.json, board/notes_for_next_meeting.md, and email.

Figures from the Appendix 5B stay as [x] unless the answers file supplies them
under "figures".
"""
import re
import json
import datetime
from pathlib import Path

R = "reports"


# ── periods and files ──────────────────────────────────────────────────────

def default_period(today: datetime.date) -> str:
    """The quarter being reported: the last quarter end before today."""
    y, m = today.year, today.month
    qm = ((m - 1) // 3) * 3          # month of the previous quarter end (0 = Dec last year)
    if qm == 0:
        return f"{y - 1}-12"
    return f"{y}-{qm:02d}"


def bounds(period: str):
    y, m = (int(x) for x in period.split("-"))
    end = (datetime.date(y + (m == 12), m % 12 + 1, 1) - datetime.timedelta(days=1))
    sm = m - 2
    start = datetime.date(y, sm, 1)
    return start, end


def _quarter_name(period: str) -> str:
    _, end = bounds(period)
    return f"{end.day} {end:%B %Y}"


def _files(store: Path, sub: str, exts=(".pdf", ".docx", ".pptx", ".md", ".txt", ".json")):
    d = store / R / sub
    if not d.exists():
        return []
    return sorted(f for f in d.iterdir() if f.is_file() and f.suffix.lower() in exts
                  and not f.name.startswith(("~", ".")) and "readme" not in f.name.lower())


def _text(f: Path, limit=None) -> str:
    if f.suffix.lower() == ".pptx":
        from deck_reader import read_deck, deck_as_text
        t = deck_as_text(read_deck(str(f)), limit or 40000)
    else:
        import board_sources as bs
        t = bs.doc_text(f)
    t = re.sub(r"For personal use only\s*", "", t)
    return t[:limit] if limit else t


def _narrative(t: str, limit: int) -> str:
    """The narrative part of a report: up to the reserves tables / statutory text."""
    m = re.search(r"\n\s*(Resource and Reserve Estimates|Reserves and Resources|Competent Person)", t)
    return (t[:m.start()] if m and m.start() > 2000 else t)[:limit]


def _section(t: str, start_pat: str, end_pat: str, limit: int) -> str:
    """The longest stretch from a start heading to the next end heading (skips the contents page)."""
    best = ""
    for a in re.finditer(start_pat, t, re.I):
        b = re.search(end_pat, t[a.end():], re.I)
        part = t[a.start(): a.end() + (b.start() if b else limit)]
        if len(part) > len(best):
            best = part
    return (best or t)[:limit]


def _read(p: Path, empty):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return empty


def _paths(store: Path, period: str):
    b = store / R / "briefs"
    return b / f"{period} brief.json", b / f"{period} answers.json", b / f"{period} brief.md"


# ── evidence ───────────────────────────────────────────────────────────────

def gather(store: Path, period: str, emails: list) -> dict:
    from core.graph import build_email_context
    store = Path(store)
    start, end = bounds(period)
    today = datetime.date.today()
    pub = _files(store, "published")
    quarterlies = [f for f in pub if "quarterly" in f.name.lower() and f.name[:7] < period]
    annuals = [f for f in pub if re.search(r"annual|half", f.name, re.I) and f.name[:7] <= period]
    templates = [f for f in _files(store, "templates", (".docx",)) if f.name[:7] < period]
    decks = _files(store, "strategy", (".pptx",))
    if not templates:
        raise FileNotFoundError("No Word template in reports/templates")

    import quarterly_render as qr
    ctx = {"period": period, "quarter_end": _quarter_name(period),
           "start": start.isoformat(), "end": end.isoformat(),
           "template": templates[-1],
           "template_text": _narrative(qr.template_text(str(templates[-1])), 18000)}
    older = [f for f in quarterlies if f.stem[:7] != templates[-1].stem[:7]]
    ctx["previous"] = [{"name": f.name, "text": _narrative(_text(f), 9000)} for f in older[-1:]]
    ctx["annual"] = ({"name": annuals[-1].name,
                      "text": _section(_text(annuals[-1]), r"(Chairman.s letter|Review of operations)",
                                       r"Directors.? report|Remuneration report|Corporate governance", 12000)}
                     if annuals else None)
    ctx["deck"] = ({"name": decks[-1].name, "text": _text(decks[-1], 14000)} if decks else None)

    # board material from the start of the quarter to today
    import board_sources as bs
    mins = []
    md = store / "board" / "minutes"
    for f in (sorted(md.iterdir()) if md.exists() else []):
        if f.is_file() and f.suffix.lower() in (".docx", ".md", ".txt", ".pdf") \
                and "readme" not in f.name.lower() and bs._file_date(f) >= start.isoformat():
            mins.append({"name": f.name, "text": bs.doc_text(f)[:7000]})
    ctx["minutes"] = mins[-2:]
    packs = []
    pd = store / "board" / "packs"
    for f in (sorted(pd.iterdir()) if pd.exists() else []):
        if f.is_file() and (f.suffix == ".pptx" or f.name.endswith(".deck.json")) and f.name[:10] >= start.isoformat():
            try:
                if f.suffix == ".pptx":
                    t = _text(f, 9000)
                else:
                    from deck_reader import deck_as_text
                    t = deck_as_text(_read(f, []), 9000)
                packs.append({"name": f.name, "text": t})
            except Exception:
                pass
    ctx["packs"] = packs[-1:]
    sent = []
    for f in sorted((store / "board" / "sent").glob("*.txt")):
        if f.stem >= start.isoformat():
            t = f.read_text(encoding="utf-8", errors="ignore")
            if "teams.microsoft.com" not in t and len(t) > 300:
                sent.append({"date": f.stem, "text": t[:2200]})
    ctx["sent"] = sorted(sent, key=lambda x: x["date"], reverse=True)[:8]
    matters = _read(store / "matters.json", {"matters": {}}).get("matters", {})
    keys = ("category", "description", "next_step", "next_step_due", "last_active")
    ctx["matters"] = {n: {k: m.get(k) for k in keys} for n, m in matters.items()
                      if m.get("status", "active") == "active"}
    notes = store / "board" / "notes_for_next_meeting.md"
    ctx["notes"] = notes.read_text(encoding="utf-8")[:6000] if notes.exists() else ""
    g = store / R / "quarterly_guidance.md"
    ctx["guidance"] = g.read_text(encoding="utf-8")[:6000] if g.exists() else ""
    ctx["earlier_drafts"] = [{"name": f.name, "text": _narrative(_text(f), 14000)} for f in _files(store, "drafts", (".docx",))
                             if f.name.startswith(period) and "verification" not in f.name.lower()][-1:]
    ctx["email"] = build_email_context(emails)[:22000]
    ctx["today"] = today.isoformat()
    return ctx


def _evidence(ctx) -> str:
    def block(title, items, key="text"):
        if not items:
            return ""
        return f"\n{title}\n" + "\n\n".join(f"--- {i.get('name') or i.get('date')}\n{i[key]}" for i in items) + "\n"
    s = f"""COMPANY: State Gas Limited (ASX: GAS). REPORTING PERIOD: quarter ended {ctx['quarter_end']}
({ctx['start']} to {ctx['end']}). Today is {ctx['today']}; events after quarter end and before release
can be reported as subsequent events.

STANDING GUIDANCE FOR THE QUARTERLY (from the MD - follow it):
{ctx['guidance']}

LAST QUARTERLY, AS THE WORD TEMPLATE ({ctx['template'].name}) - structure, tone and statutory text:
{ctx['template_text']}
"""
    if ctx.get("deck"):
        s += f"\nLATEST STRATEGY DECK ({ctx['deck']['name']}) - THE MOST RECENT STATEMENT OF THE COMPANY'S MESSAGING:\n{ctx['deck']['text']}\n"
    s += block("EARLIER RELEASED QUARTERLIES:", ctx["previous"])
    if ctx.get("annual"):
        s += f"\nLATEST ANNUAL / HALF-YEAR REPORT ({ctx['annual']['name']}):\n{ctx['annual']['text']}\n"
    s += block("BOARD MINUTES IN THE PERIOD:", ctx["minutes"])
    s += block("BOARD PACKS IN THE PERIOD:", ctx["packs"])
    s += block("UPDATES SENT TO DIRECTORS (newest first):", ctx["sent"])
    s += "\nCOMPANY KNOWLEDGE - ACTIVE MATTERS:\n" + json.dumps(ctx["matters"], ensure_ascii=False)[:7000] + "\n"
    if ctx.get("notes"):
        s += f"\nMD'S RUNNING NOTES:\n{ctx['notes']}\n"
    s += block("AN EARLIER DRAFT OF THIS QUARTER'S REPORT (worked through with the MD):", ctx["earlier_drafts"])
    s += f"\nEMAIL FROM THE PERIOD:\n{ctx['email']}\n"
    return s


def _answers(ctx) -> str:
    a = ctx.get("answers") or {}
    if not a:
        return ""
    return ("\nTHE MD'S DECISIONS FOR THIS REPORT - BINDING. Follow them exactly; they override the "
            "evidence and the earlier draft:\n" + json.dumps(a, ensure_ascii=False, indent=1)[:12000] + "\n")


def _call(client, prompt, max_tokens, label, model=None):
    """Model call that raises with the reason (the error goes to the private log, not the public one)."""
    from core.llm import api_call_with_retry, robust_json_parse
    msg = api_call_with_retry(client, model=model or None, max_tokens=max_tokens,
                              messages=[{"role": "user", "content": prompt}]) if model else \
        api_call_with_retry(client, max_tokens=max_tokens, messages=[{"role": "user", "content": prompt}])
    raw = msg.content[0].text.strip()
    try:
        return robust_json_parse(raw) or {}
    except Exception:
        raise RuntimeError(f"{label}: could not read the model's answer "
                           f"(stop_reason={getattr(msg, 'stop_reason', '?')}, {len(raw)} chars)")


# ── step 1: brief ──────────────────────────────────────────────────────────

def make_brief(client, store: Path, emails: list, period: str = None, model: str = None) -> dict:
    store = Path(store)
    period = period or default_period(datetime.date.today())
    ctx = gather(store, period, emails)
    prompt = _evidence(ctx) + """
TASK: do NOT write the report. Prepare a short brief the MD will review with an assistant before the
quarterly is drafted, so that he makes the judgement calls. Be concrete.
- highlights: 4-6 candidate highlights (one sentence each) with the evidence for each.
- sections: the sections the report should have, in order - last quarter's sections (carry / drop,
  why) plus new ones - each with 2-4 points you would make and the source.
- messaging: where the strategy deck's framing differs from the last quarterly, what you propose.
- sensitive: items that may or may not be suitable for public disclosure (commercial negotiations,
  litigation, counterparties, unreleased numbers), each with a recommendation.
- guidance_changes: anything that would change previously released guidance (timelines, volumes,
  costs, commitments) and how you propose to say it.
- figures: the figures the report needs from the Appendix 5B or the accounts.
- questions: 5-10 pointed questions only the MD can answer, each with why it matters and 2-4
  suggested answers.

Respond with JSON only:
{"highlights":[{"text":"","evidence":""}],
 "sections":[{"heading":"","status":"carry|drop|new","why":"","points":[""]}],
 "messaging":[""],
 "sensitive":[{"item":"","recommend":"disclose|hold|confirm","why":""}],
 "guidance_changes":[{"item":"","previous":"","proposed":""}],
 "figures":[""],
 "questions":[{"q":"","why":"","options":[""]}]}"""
    brief = _call(client, prompt, 9000, "quarterly brief", model)
    if not brief:
        raise RuntimeError("brief step returned nothing")
    brief["period"], brief["quarter_end"] = period, ctx["quarter_end"]
    brief["evidence"] = {"template": ctx["template"].name, "deck": (ctx["deck"] or {}).get("name"),
                         "previous": [p["name"] for p in ctx["previous"]],
                         "annual": (ctx["annual"] or {}).get("name"),
                         "minutes": [m["name"] for m in ctx["minutes"]],
                         "packs": [p["name"] for p in ctx["packs"]],
                         "updates_to_directors": len(ctx["sent"]),
                         "earlier_draft": [d["name"] for d in ctx["earlier_drafts"]]}
    bj, _, bm = _paths(store, period)
    bj.parent.mkdir(parents=True, exist_ok=True)
    bj.write_text(json.dumps(brief, indent=1, ensure_ascii=False), encoding="utf-8")
    bm.write_text(_brief_md(brief), encoding="utf-8")
    return {"json": bj, "md": bm}


def _brief_md(b: dict) -> str:
    L = [f"# Quarterly brief - quarter ended {b.get('quarter_end','')}", "",
         f"Evidence: {json.dumps(b.get('evidence', {}), ensure_ascii=False)}", "", "## Questions for you"]
    for i, q in enumerate(b.get("questions", []), 1):
        L.append(f"{i}. {q.get('q','')}  _({q.get('why','')})_")
        if q.get("options"):
            L.append("   Options: " + " / ".join(q["options"]))
    L += ["", "## Candidate highlights"] + [f"- {h.get('text','')}  _[{h.get('evidence','')}]_" for h in b.get("highlights", [])]
    L += ["", "## Sections"]
    for s in b.get("sections", []):
        L.append(f"- **{s.get('heading','')}** ({s.get('status','')}) - {s.get('why','')}")
        L += [f"  - {p}" for p in s.get("points", [])]
    L += ["", "## Messaging"] + [f"- {m}" for m in b.get("messaging", [])]
    L += ["", "## Sensitive items"] + [f"- {x.get('item','')} [{x.get('recommend','')}] {x.get('why','')}"
                                       for x in b.get("sensitive", [])]
    L += ["", "## Changes to previous guidance"] + [f"- {x.get('item','')}: was {x.get('previous','')}; proposed {x.get('proposed','')}"
                                                     for x in b.get("guidance_changes", [])]
    L += ["", "## Figures needed"] + [f"- {f}" for f in b.get("figures", [])]
    return "\n".join(L) + "\n"


# ── step 2: build ──────────────────────────────────────────────────────────

def make_report(client, store: Path, emails: list, period: str = None, model: str = None,
                say=print) -> dict:
    store = Path(store)
    today = datetime.date.today()
    period = period or default_period(today)
    ctx = gather(store, period, emails)
    _, aj, _ = _paths(store, period)
    ctx["answers"] = _read(aj, {})
    figures = (ctx["answers"] or {}).get("figures") or {}
    end = bounds(period)[1]

    prompt = _evidence(ctx) + _answers(ctx) + f"""
TASK: write the quarterly activities report for the quarter ended {ctx['quarter_end']} in the company's
format, following the Word template's structure and voice. Australian English, plain, factual, ASX
continuous-disclosure appropriate.
- Highlights: exactly 3 bullets.
- lead_paras: the market/strategy paragraphs between the intro line and the short/medium/long-term
  table (2-3 paragraphs; the last one leads into the table).
- horizons: the table rows if they need updating (same labels as the template), else null.
- after_table: one paragraph straight after the table.
- sections: the narrative sections in order, ending with Outlook (objectives for next quarter as
  bullets). Blocks are paragraphs (t "p") or bullets (t "b").
- Use [x] for any figure not supplied here: {json.dumps(figures, ensure_ascii=False)}
- Put text the MD must decide on in [square brackets] beginning OPTIONAL or CONFIRM.
- lr_5_31_6: replacement text for the LR 5.31.6 statutory row ONLY if the development timeline or
  pre-conditions have changed; else "".
- related_parties / e_and_e: the two payment paragraphs, with [x] for amounts not supplied.
- date_line: "[xx] {(end + datetime.timedelta(days=31)):%B %Y}" unless supplied.

Respond with JSON only:
{{"date_line":"","title":"QUARTERLY ACTIVITIES REPORT – {ctx['quarter_end']}","highlights":["","",""],
 "intro":"","lead_paras":[""],"horizons":null,"after_table":"",
 "sections":[{{"heading":"","blocks":[{{"t":"p","text":""}}]}}],
 "lr_5_31_6":"","related_parties":"","e_and_e":""}}"""
    spec = _call(client, prompt, 16000, "quarterly draft", model)
    if not spec.get("sections"):
        raise RuntimeError("drafting step returned nothing")
    if figures.get("date_line"):
        spec["date_line"] = figures["date_line"]

    ver_prompt = _evidence(ctx) + _answers(ctx) + f"""
THE DRAFT REPORT (JSON):
{json.dumps(spec, ensure_ascii=False)[:30000]}

TASK: prepare the verification document for this draft, for the MD only.
- statements: every factual statement in the draft (highlights, each section), with the source it
  rests on (document name / email date and sender / board minutes) and a status:
  "verified" (source says it), "previously released" (in an ASX release), "confirm" (needs the MD
  or a third party), "placeholder" (figure or bracketed text to fill).
- decisions: how each of the MD's binding decisions was applied.
- left_out: things in the evidence deliberately not included, and why.
- changes: what changed against the last quarterly (structure, messaging, guidance).
- deck_alignment: where the draft follows or departs from the strategy deck, and what in the deck
  conflicts with released information.
- app5b_8_8: draft answers to Appendix 5B items 8.8.1-8.8.3 (needed if funding is under 2 quarters).

Respond with JSON only:
{{"statements":[{{"statement":"","source":"","status":""}}],"decisions":[{{"decision":"","applied":""}}],
 "left_out":[""],"changes":[""],"deck_alignment":[{{"deck":"","draft":"","note":""}}],
 "app5b_8_8":[{{"q":"","a":""}}]}}"""
    ver = _call(client, ver_prompt, 12000, "quarterly verification", model)

    import quarterly_render as qr
    out_dir = store / R / "drafts"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = f"{period} Quarterly DRAFT {today.isoformat()}"
    docx_path = out_dir / f"{stamp}.docx"
    res = qr.build(spec, str(ctx["template"]), str(docx_path))
    ver_path = out_dir / f"{stamp} - verification.docx"
    write_verification(str(ver_path), spec, ver, ctx, res.get("notes", []))
    (out_dir / f"{stamp} spec.json").write_text(json.dumps({"spec": spec, "verification": ver},
                                                           indent=1, ensure_ascii=False), encoding="utf-8")
    say(f"Quarterly drafted: {len(spec['sections'])} sections, "
        f"{len(ver.get('statements', []))} statements traced")
    return {"docx": docx_path, "verification": ver_path, "spec": out_dir / f"{stamp} spec.json"}


def write_verification(path, spec, ver, ctx, notes):
    import docx
    from docx.shared import Pt, RGBColor
    d = docx.Document()
    st = d.styles["Normal"]; st.font.name = "Aptos"; st.font.size = Pt(10)

    def h(t):
        p = d.add_heading(t, 1)
        for r in p.runs:
            r.font.color.rgb = RGBColor(0xA5, 0x1C, 0x30)

    def table(head, rows):
        if not rows:
            d.add_paragraph("None.")
            return
        t = d.add_table(rows=1, cols=len(head)); t.style = "Light Grid Accent 1"
        for c, x in zip(t.rows[0].cells, head):
            c.text = ""; c.paragraphs[0].add_run(x).bold = True
        for r in rows:
            for c, x in zip(t.add_row().cells, r):
                c.text = str(x or "")
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
        d.add_paragraph()

    d.add_heading(f"Quarterly verification – quarter ended {ctx['quarter_end']}", 0)
    d.add_paragraph(f"Prepared {datetime.date.today():%d %B %Y} · for MD review only, not for release")
    ev = [f"Template: {ctx['template'].name}",
          f"Strategy deck: {(ctx.get('deck') or {}).get('name', 'none')}",
          f"Earlier quarterlies: {', '.join(p['name'] for p in ctx['previous']) or 'none'}",
          f"Annual / half-year: {(ctx.get('annual') or {}).get('name', 'none')}",
          f"Board minutes: {', '.join(m['name'] for m in ctx['minutes']) or 'none'}",
          f"Board packs: {', '.join(p['name'] for p in ctx['packs']) or 'none'}",
          f"Updates to directors: {len(ctx['sent'])}",
          f"Answers file: {'yes' if ctx.get('answers') else 'NO - draft made without your decisions'}"]
    h("1. Evidence used")
    for e in ev + [f"Note: {n}" for n in notes]:
        d.add_paragraph(e, style="List Bullet")
    h("2. Your decisions and how they were applied")
    table(["Decision", "Applied"], [(x.get("decision"), x.get("applied")) for x in ver.get("decisions", [])])
    h("3. Statement-by-statement sources")
    table(["Statement", "Source", "Status"],
          [(x.get("statement"), x.get("source"), x.get("status")) for x in ver.get("statements", [])])
    h("4. Alignment with the strategy deck")
    table(["Deck", "Draft", "Note"], [(x.get("deck"), x.get("draft"), x.get("note")) for x in ver.get("deck_alignment", [])])
    h("5. Changes against the last quarterly")
    for t in ver.get("changes", []) or ["None noted."]:
        d.add_paragraph(t, style="List Bullet")
    h("6. Left out deliberately")
    for t in ver.get("left_out", []) or ["None noted."]:
        d.add_paragraph(t, style="List Bullet")
    h("7. Appendix 5B section 8 – draft answers")
    table(["Question", "Draft answer"], [(x.get("q"), x.get("a")) for x in ver.get("app5b_8_8", [])])
    d.save(path)
