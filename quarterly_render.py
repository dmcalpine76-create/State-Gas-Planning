"""
quarterly_render.py - writes a quarterly activities report into the company's
own Word format, using the previous quarter's report as the template.

The template is found by markers rather than fixed positions, so a template
saved from Word with small layout changes still works:
  * a date line near the top and the "QUARTERLY ACTIVITIES REPORT" title
  * the highlights box (first table, bulleted paragraphs)
  * the "provides/releases this update" intro line
  * the short/medium/long-term table (a table whose first row says "Short term")
  * the "Resource and Reserve Estimates" heading, where the narrative ends
  * the related-party and E&E payment paragraphs
Everything from the reserves heading down (tables, competent persons
statements, About sections) is kept as it is, except for the LR 5.31.6 cell
and the payment paragraphs when new text is supplied.
"""
import copy
import re

import docx
from docx.text.paragraph import Paragraph

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _set(p, text):
    runs = p.runs
    if not runs:
        p.add_run(text)
        return p
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    return p


def _is_bullet(el):
    return el.find(f".//{W}numPr") is not None


def _bold(p):
    return bool(p.runs) and bool(p.runs[0].bold)


def build(spec: dict, template: str, out: str) -> dict:
    d = docx.Document(template)
    body = d.element.body
    els = list(body.iterchildren())
    paras = {id(e): Paragraph(e, d) for e in els if e.tag == f"{W}p"}
    P = lambda e: paras[id(e)]
    notes = []

    def find(pred, start=0):
        for i in range(start, len(els)):
            e = els[i]
            if e.tag == f"{W}p" and pred(P(e)):
                return i
        return None

    # header
    i_title = find(lambda p: p.text.strip().upper().startswith("QUARTERLY ACTIVITIES REPORT"))
    if i_title is None:
        raise ValueError("template: no QUARTERLY ACTIVITIES REPORT title")
    i_date = find(lambda p: bool(re.match(r"^\s*(\[?\w+\]?\s+)?\w+ \d{4}\s*$", p.text)) and len(p.text) < 30)
    if i_date is not None and i_date < i_title and spec.get("date_line"):
        _set(P(els[i_date]), spec["date_line"])
    if spec.get("title"):
        _set(P(els[i_title]), spec["title"])

    # highlights box
    hl = spec.get("highlights") or []
    tbl0 = d.tables[0]
    hps = [Paragraph(p, d) for p in tbl0._tbl.iter(f"{W}p") if _is_bullet(p)]
    if hps and hl:
        while len(hps) < len(hl):                       # add bullets if needed
            n = copy.deepcopy(hps[-1]._p); hps[-1]._p.addnext(n); hps.append(Paragraph(n, d))
        for p, t in zip(hps, hl):
            _set(p, t)
        for p in hps[len(hl):]:
            p._p.getparent().remove(p._p)

    # intro line
    i_intro = find(lambda p: re.search(r"(provides|releases) this update", p.text, re.I) is not None)
    if i_intro is None:
        raise ValueError("template: no intro line")
    if spec.get("intro"):
        _set(P(els[i_intro]), spec["intro"])

    # anchors
    i_res = find(lambda p: p.text.strip().lower().startswith("resource and reserve estimates"), i_intro)
    if i_res is None:
        raise ValueError("template: no 'Resource and Reserve Estimates' heading")
    i_tab = next((i for i in range(i_intro, i_res) if els[i].tag == f"{W}tbl"
                  and "short term" in "".join(t.text or "" for t in els[i].iter(f"{W}t")).lower()), None)

    # paragraph templates, taken from the old narrative before it is removed
    narr = [P(els[i]) for i in range(i_intro + 1, i_res) if els[i].tag == f"{W}p"]
    t_head = next((p._p for p in narr if _bold(p) and 0 < len(p.text) < 90), None)
    plain = [p._p for p in narr if not _bold(p) and len(p.text) > 120 and not _is_bullet(p._p)]
    t_body = next((p._p for p in narr if p._p in plain and p.style.name == "Normal"), None)
    if t_body is None and plain:
        t_body = plain[0]
    t_lead = plain[0] if plain else t_body
    t_bull = next((p._p for p in narr if _is_bullet(p._p)), None)
    t_gap = next((p._p for p in narr if not p.text.strip()), None)
    if not (t_head is not None and t_body is not None and t_bull is not None):
        raise ValueError("template: could not find heading/body/bullet styles in the narrative")
    t_head, t_body, t_lead, t_bull = (copy.deepcopy(x) for x in (t_head, t_body, t_lead, t_bull))
    t_gap = copy.deepcopy(t_gap) if t_gap is not None else None

    def add(after, tmpl, text):
        el = copy.deepcopy(tmpl)
        after.addnext(el)
        _set(Paragraph(el, d), text)
        return el

    # lead-in paragraphs (between the intro line and the horizons table)
    lead = spec.get("lead_paras") or []
    if i_tab is not None:
        old = [els[i] for i in range(i_intro + 1, i_tab)]
        after_tab = [els[i] for i in range(i_tab + 1, i_res)]
    else:
        old, after_tab = [els[i] for i in range(i_intro + 1, i_res)], []
    if lead:
        for e in old:
            body.remove(e)
        cur = els[i_intro]
        for t in lead:
            if t_gap is not None:
                g = copy.deepcopy(t_gap); cur.addnext(g); cur = g
            cur = add(cur, t_lead, t)
        if t_gap is not None:
            g = copy.deepcopy(t_gap); cur.addnext(g)

    # horizons table
    if i_tab is not None and spec.get("horizons"):
        tb = docx.table.Table(els[i_tab], d)
        for row in tb.rows[1:]:
            label = row.cells[0].text.strip().lower()
            for k, vals in spec["horizons"].items():
                if k.strip().lower() == label and isinstance(vals, list):
                    for c, v in zip(row.cells[1:], vals):
                        if v and c.paragraphs:
                            _set(c.paragraphs[0], v)

    # narrative sections (after the table, down to the reserves heading)
    region = after_tab if i_tab is not None else ([] if lead else old)
    sections = spec.get("sections") or []
    if sections:
        anchor = els[i_tab] if i_tab is not None else (els[i_intro])
        for e in region:
            if e.getparent() is not None:
                body.remove(e)
        cur = anchor
        if spec.get("after_table"):
            if t_gap is not None:
                g = copy.deepcopy(t_gap); cur.addnext(g); cur = g
            cur = add(cur, t_lead, spec["after_table"])
        for s in sections:
            cur = add(cur, t_head, s.get("heading", ""))
            for b in s.get("blocks", []):
                txt = (b.get("text") or "").strip()
                if txt:
                    cur = add(cur, t_bull if b.get("t") == "b" else t_body, txt)
        if t_gap is not None:
            g = copy.deepcopy(t_gap); cur.addnext(g)

    # LR 5.31.6 statement
    if spec.get("lr_5_31_6"):
        done = False
        for t in d.tables:
            for r in t.rows:
                if r.cells and "5.31.6" in r.cells[0].text and len(r.cells) > 1:
                    ps = [p for p in r.cells[1].paragraphs if p.text.strip()]
                    if ps:
                        _set(ps[0], spec["lr_5_31_6"])
                        for p in ps[1:]:
                            p._p.getparent().remove(p._p)
                        done = True
        if not done:
            notes.append("LR 5.31.6 row not found - statutory timeline text not updated")

    # payments
    for p in d.paragraphs:
        tx = p.text.strip()
        if tx.startswith("A total of $") and spec.get("related_parties"):
            _set(p, spec["related_parties"])
        elif tx.startswith("In accordance with ASX Listing Rule 5.3.1") and spec.get("e_and_e"):
            _set(p, spec["e_and_e"])

    d.save(out)
    return {"notes": notes}


def template_text(path: str) -> str:
    """Plain text of a Word report, tables flattened, for prompts."""
    d = docx.Document(path)
    out = []
    for el in d.element.body.iterchildren():
        if el.tag == f"{W}p":
            t = Paragraph(el, d).text.strip()
            if t:
                out.append(("- " if _is_bullet(el) else "") + t)
        elif el.tag == f"{W}tbl":
            for r in docx.table.Table(el, d).rows:
                cells = []
                for c in r.cells:
                    if c.text.strip() not in cells:
                        cells.append(c.text.strip())
                out.append(" | ".join(cells))
    return "\n".join(out)
