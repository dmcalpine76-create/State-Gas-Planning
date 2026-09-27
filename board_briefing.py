"""
board_briefing.py  —  Monthly Board Report Briefing Tool  (v3, consolidated)
-----------------------------------------------------------------------------
Scans the Outlook mailbox for the preceding N days (default 30) and produces:

  board_report_data.json           — structured data for generate_board_deck.js
                                     (format UNCHANGED — the PPTX pipeline still works)
  StateGas_BoardBriefing_DATE.md   — human-readable draft for review
  StateGas_BoardReport_Draft_DATE.docx  — optional Word draft (--docx, needs node)

Changes from v2:
  • Shared infrastructure moved to core/.
  • FIXED: the review markdown now reads the fields the synthesis actually
    produces (executive_summary_intro/boxes, slide_briefing, sections) —
    previously it read the old v1 schema and rendered empty sections.     (B2)
  • FIXED: pillar history now persists — post-run updates write to
    pillar_history.json via canonical pillar ids, so next month's run
    really does see last month's RAG status.                              (B3)
  • Matters/people/facts discovered here now go to the shared
    company_knowledge.json instead of a third private context file.
  • board_report_analyser.py is retired — its Word-document output is
    available here via --docx.
  • Replaces board_briefing_context.py (see core/knowledge.py).

Usage:
  py board_briefing.py run [--days 30] [--docx]
  py board_briefing.py build-context [--days 90]
  py board_briefing.py show-context
  py board_briefing.py setup / test
"""

import json
import argparse
import datetime

from core.env import load_env
load_env()

from core.config import (BASE_DIR, COMPANY_NAME, COMPANY_TICKER, MD_NAME,
                         VOICE_RULES, STRATEGIC_PILLARS, PILLAR_ID_MAP)
from core import graph
from core.llm import get_client, call_json
from core.knowledge import (CompanyKnowledge, PillarHistory,
                            load_context_briefing, ensure_briefing_template)
from board_update import merge_matters   # single shared implementation

# ─────────────────────────────────────────────
# CONFIG (tool-specific)
# ─────────────────────────────────────────────

DEFAULT_LOOKBACK_DAYS = 30
MAX_EMAILS_PER_FOLDER = 100
MAX_EMAILS_TOTAL      = 600
MAX_BODY_CHARS        = 1500
MAX_EMAILS_PER_BATCH  = 150

OUTPUT_DIR  = BASE_DIR
OUTPUT_JSON = OUTPUT_DIR / "board_report_data.json"   # feeds generate_board_deck.js

# Token-budget guard for oversized batches
_CHARS_PER_TOKEN  = 3.5
_MAX_BATCH_TOKENS = 160000


def _estimate_tokens(text: str) -> int:
    return int(len(text) / _CHARS_PER_TOKEN)


def _split_oversized_batch(emails: list, prompt_overhead: int) -> list:
    budget = _MAX_BATCH_TOKENS - prompt_overhead
    sub_batches, current, current_tok = [], [], 0
    for email in emails:
        email_tokens = _estimate_tokens(
            f"Subject: {email.get('subject','')}\n"
            f"From: {email.get('from_name','')}\n"
            f"Body: {email.get('body','')[:MAX_BODY_CHARS]}\n")
        if current and current_tok + email_tokens > budget:
            sub_batches.append(current)
            current, current_tok = [], 0
        current.append(email)
        current_tok += email_tokens
    if current:
        sub_batches.append(current)
    return sub_batches


def _pillar_definitions_block() -> str:
    lines = []
    for i, p in enumerate(STRATEGIC_PILLARS, 1):
        lines += [f"  {i}. {p['title']}",
                  f"     Scope: {p['description']}",
                  f"     Key terms: {', '.join(p['keywords'][:12])}", ""]
    return "\n".join(lines)


# ─────────────────────────────────────────────
# EXTRACTION
# ─────────────────────────────────────────────

def extract_matters_from_batch(client, emails: list, batch_num: int,
                               total_batches: int, context_block: str,
                               briefing_text: str, lookback_days: int) -> list:
    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today()
                 - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")

    context_section  = f"{context_block}\n" if context_block else ""
    briefing_section = (f"STANDING INSTRUCTIONS:\n{briefing_text}\n\n"
                        if briefing_text else "")

    prompt = f"""You are a senior executive assistant preparing the monthly board paper for
{COMPANY_NAME} ({COMPANY_TICKER}).

{context_section}{briefing_section}TODAY: {date_to}
PERIOD COVERED: {date_from} to {date_to}
EMAIL BATCH: {batch_num} of {total_batches}

STATE GAS — 6 STRATEGIC PILLARS:
{_pillar_definitions_block()}

YOUR TASK:
Read the emails below. For each email that contains board-relevant content,
identify the matter, which strategic pillar it relates to, and write a
factual update suitable for inclusion in a board paper.

IMPORTANT — if you have accumulated context from previous runs above, use it:
  • Report CHANGES from last month's status, not just current facts
  • Flag when something has improved or deteriorated from last month
  • Note when a matter that was previously active has gone quiet

Focus on matters where the Board would reasonably want to be informed.
Ignore: routine personal admin, travel bookings, automated notifications,
newsletters, receipts, internal scheduling with no substantive outcome.

RESPOND ONLY WITH VALID JSON — a single array, no markdown:

[
  {{
    "pillar_id":    "reserves_production | litigation | access_rights | rolleston_west | alt_fuels | capital | other",
    "matter":       "Short matter title",
    "significance": "high | medium | low",
    "update":       "2-4 sentence factual update, board-appropriate tone. Reference any change from prior status if context is available.",
    "date_range":   "e.g. '1-11 March 2026'",
    "key_parties":  ["Name or organisation"],
    "source_subject": "Subject line of the primary email"
  }}
]

Return [] if this batch contains no board-relevant content.

{'─'*60}
EMAILS ({len(emails)} in this batch):
{'─'*60}

{graph.build_email_context(emails)}
"""
    matters = call_json(client, prompt, max_tokens=8000,
                        label=f"Briefing batch {batch_num}")
    if not isinstance(matters, list):
        return []
    print(f"✅  {len(matters)} matter{'s' if len(matters) != 1 else ''}")
    return matters


# ─────────────────────────────────────────────
# SYNTHESIS  (schema unchanged — feeds the PPTX generator)
# ─────────────────────────────────────────────

BOARD_SYNTHESIS_SCHEMA = """{
  "report_title":  "Executive Board Report",
  "period_label":  "e.g. Quarter 1 FY26",
  "meeting_date":  "dd Month YYYY",
  "author":        "Doug McAlpine",
  "author_title":  "Managing Director",
  "emails_reviewed": <number>,

  "executive_summary_intro": [
    "First top-level bullet — 1-2 sentences on the most important strategic theme this period.",
    "Second top-level bullet — 1-2 sentences on capital or risk position.",
    "Third top-level bullet — 1-2 sentences on a key decision or upcoming milestone."
  ],

  "executive_summary_boxes": [
    {"title": "Theme title", "bullets": ["First-person bullet.", "Second bullet."]},
    {"title": "Second theme", "bullets": ["bullet", "bullet"]},
    {"title": "Third theme", "bullets": ["bullet", "bullet"]},
    {"title": "Fourth theme", "bullets": ["bullet", "bullet"]}
  ],

  "strategic_pillars": [
    {"id": "reserves",   "title": "2P Reserve",              "rag_status": "ON_TRACK | IN_PROGRESS | WATCH | URGENT", "slide_briefing": "2-3 sentence summary.", "extended_bullets": null},
    {"id": "litigation", "title": "Litigation",             "rag_status": "IN_PROGRESS", "slide_briefing": "2-3 sentences.", "extended_bullets": null},
    {"id": "access",     "title": "Access Rights Extension", "rag_status": "IN_PROGRESS", "slide_briefing": "2-3 sentences.", "extended_bullets": null},
    {"id": "rw",         "title": "Rolleston West Concept Study", "rag_status": "IN_PROGRESS", "slide_briefing": "2-3 sentences.", "extended_bullets": null},
    {"id": "hdng",       "title": "HDNG Strategy",           "rag_status": "IN_PROGRESS", "slide_briefing": "2-3 sentences.", "extended_bullets": null},
    {"id": "capital",    "title": "Capital Management",      "rag_status": "WATCH", "slide_briefing": "Brief summary.",
     "extended_bullets": ["Detailed first-person bullet for the expanded Capital card.", "Second bullet.", "Third bullet."]}
  ],

  "sections": [
    {"number": 3, "title": "Section title",
     "bullets": ["First-person bullet summarising this topic.", "Second bullet."],
     "right_title": "Optional right-column heading",
     "right_bullets": ["Right column bullet."],
     "right_sub_title": "Optional second heading in right column",
     "right_sub_bullets": ["sub-bullet"]}
  ],

  "analyst_flags": [
    "Candid observation for Doug's attention — not for the deck itself."
  ]
}"""


def synthesise_board_briefing(client, matters: list, emails: list,
                              lookback_days: int, context_block: str) -> dict:
    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today()
                 - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")

    if matters:
        m_lines = []
        for m in matters:
            m_lines.append(f"[{m.get('significance','?').upper()}] "
                           f"[{m.get('pillar_id','?')}] {m.get('matter','')}")
            m_lines.append(f"  {m.get('update','')}")
            if m.get("key_parties"):
                m_lines.append(f"  Parties: {', '.join(m['key_parties'])}")
            if m.get("source_subject"):
                m_lines.append(f"  Source: {m['source_subject']}")
            m_lines.append("")
        matters_text = "\n".join(m_lines)
    else:
        matters_text = "No board-relevant matters extracted from emails this period."
    if len(matters_text) > 20000:
        matters_text = matters_text[:20000] + "\n  … (matters text truncated)"

    ctx = context_block or "(No prior context — first run)"
    if len(ctx) > 6000:
        ctx = ctx[:6000] + "\n  … (context truncated)\n"

    prompt = f"""You are a senior corporate advisor preparing the monthly board paper for
{COMPANY_NAME} ({COMPANY_TICKER}).

You have been given:
1. Extracted matters from this month's emails (below)
2. Accumulated context from previous runs (if available)

YOUR TASK:
Synthesise the extracted matters into structured JSON that feeds the State Gas
Executive Board Report PowerPoint presentation. Your output must match the
schema precisely.

BOARD REPORT STRUCTURE:
  Slide 1  — Title (auto-generated from meta fields)
  Slide 2  — Executive Summary: 3 intro bullets + 4 themed sub-boxes
  Slide 3  — Strategic Objectives: 6 RAG pillars (Capital expanded with detail)
  Slides 4+ — Deep-dive sections (one per major topic)

{VOICE_RULES}

If accumulated context is available, describe CHANGES from last period.

RESPOND ONLY WITH VALID JSON — no markdown, no preamble.

{BOARD_SYNTHESIS_SCHEMA}

PERIOD: {date_from} to {date_to}
MEETING DATE: {date_to}
EMAILS REVIEWED: {len(emails)}

ACCUMULATED CONTEXT:
{ctx}

EXTRACTED MATTERS FROM THIS MONTH'S EMAILS:
{matters_text}
"""
    print("  🤖  Step 2/2 — synthesising board briefing…", end=" ", flush=True)
    data = call_json(client, prompt, max_tokens=8000, label="Synthesis")
    if data:
        print("✅")
    return data or {}


# ─────────────────────────────────────────────
# POST-RUN UPDATE  (pillar history + company knowledge, one call — B3)
# ─────────────────────────────────────────────

def update_context_after_run(ph: PillarHistory, ck: CompanyKnowledge,
                             briefing_data: dict, matters: list, emails: list):
    client = get_client()
    today  = datetime.date.today().isoformat()

    # Summary built from the fields the synthesis ACTUALLY produces (B2 fix)
    lines = [f"RUN DATE: {today}", ""]
    for pt in briefing_data.get("executive_summary_intro", []):
        lines.append(f"EXEC SUMMARY: {pt}")
    lines.append("")
    for p in briefing_data.get("strategic_pillars", []):
        lines.append(f"PILLAR: {p.get('title','')}  [{p.get('rag_status','?')}]")
        lines.append(f"  {p.get('slide_briefing','')}")
        for b in (p.get("extended_bullets") or []):
            lines.append(f"  • {b}")
        lines.append("")
    for s in briefing_data.get("sections", []):
        lines.append(f"SECTION: {s.get('title','')}")
        for b in s.get("bullets", []):
            lines.append(f"  • {b}")
        lines.append("")
    lines.append("EXTRACTED MATTERS THIS RUN:")
    for m in matters[:40]:
        lines.append(f"  [{m.get('pillar_id','?')}] {m.get('matter','')}: "
                     f"{m.get('update','')[:150]}")
    lines.append("")
    lines.append(graph.top_senders_block(emails, 15))
    synthesis_summary = "\n".join(lines)

    current_pillars = ph.build_prompt_block() or "(no pillar history yet)"

    prompt = f"""You are maintaining persistent memory for {COMPANY_NAME} ({COMPANY_TICKER})'s
monthly board briefing tool. Produce ONE JSON object with two top-level keys:
"pillar_updates" and "knowledge_patch".

Be conservative — better to add nothing than noise. Only include information
that is genuinely new or changed, stable, and factual.

CURRENT PILLAR HISTORY:
{current_pillars}

THIS MONTH'S SYNTHESIS:
{synthesis_summary}

RESPOND ONLY WITH VALID JSON — no markdown, no preamble:

{{
  "pillar_updates": [
    {{
      "id":           "one of: reserves_production, litigation, access_rights, rolleston_west, alt_fuels, capital",
      "last_rag":     "ON_TRACK | IN_PROGRESS | WATCH | URGENT",
      "last_summary": "3-4 sentence plain-English summary of this pillar's status this month",
      "key_watch":    "Updated watch signals for next month"
    }}
  ],
  "knowledge_patch": {{
    "matter_updates": [
      {{"name": "...", "category": "Commercial | Regulatory | JV/Partner | Operational | Financial | Stakeholder | Legal | Other",
        "status": "active | dormant | closed", "description": "...",
        "key_parties": ["..."], "watch_for": "..."}}
    ],
    "close_matter_names": [],
    "people_updates": [
      {{"name": "...", "email": "...", "organisation": "...", "role": "...", "notes": "..."}}
    ],
    "new_facts": [],
    "company_profile": ""
  }}
}}

Provide a pillar_update for EVERY pillar that had activity this month.
Leave knowledge_patch arrays empty if nothing is genuinely new.
"""
    print("  🧠  Updating pillar history + company knowledge…", end=" ", flush=True)
    patch = call_json(client, prompt, max_tokens=8000, label="Post-run patch")
    if patch:
        pillar_updates = patch.get("pillar_updates", [])
        # Accept either canonical or short deck ids (mapped via PILLAR_ID_MAP)
        for pu in pillar_updates:
            pid = pu.get("id", "")
            if pid in PILLAR_ID_MAP:
                pu["id"] = PILLAR_ID_MAP[pid]
        ph.apply_patch(pillar_updates)
        kp = patch.get("knowledge_patch") or {}
        ck.apply_patch(kp)
        ck.prune()
        ck.save()
        print(f"✅  ({len(pillar_updates)} pillars, "
              f"{len(kp.get('matter_updates', []))} matters, "
              f"{len(kp.get('new_facts', []))} facts)")
    else:
        print("⚠️  patch unavailable — run count still saved.")
    ph.increment_run_count()
    ph.save()


# ─────────────────────────────────────────────
# MARKDOWN OUTPUT  (now matches the synthesis schema — B2)
# ─────────────────────────────────────────────

RAG_EMOJI = {"ON_TRACK": "🟢", "IN_PROGRESS": "🟡", "WATCH": "🟠", "URGENT": "🔴"}


def generate_markdown(data: dict, emails: list, lookback_days: int) -> str:
    now       = datetime.datetime.now()
    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today()
                 - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")

    lines = [
        f"# {COMPANY_NAME} — Board Report Briefing",
        "",
        f"> **Period:** {date_from} → {date_to}  ",
        f"> **Emails reviewed:** {len(emails)}  ",
        f"> **Generated:** {now.strftime('%d %B %Y, %H:%M')}  ",
        f"> **Status:** DRAFT — AI-assisted. Review all content before use.",
        "", "---", "",
    ]

    # Section 1 — Executive summary (intro bullets + themed boxes)
    lines += ["## Section 1 — Executive Summary", "",
              "_Copy into the Executive Summary PowerPoint slide._", ""]
    for pt in data.get("executive_summary_intro", []):
        lines.append(f"- {pt}")
    lines.append("")
    for box in data.get("executive_summary_boxes", []):
        lines += [f"**{box.get('title', '')}**", ""]
        for b in box.get("bullets", []):
            lines.append(f"- {b}")
        lines.append("")
    lines += ["---", ""]

    # Section 2 — Strategic pillars
    lines += ["## Section 2 — Strategic Pillars", "",
              "_Each **Slide Briefing** pastes into the Strategic Milestone "
              "Tracker table for that pillar._", ""]
    for p in data.get("strategic_pillars", []):
        rag   = p.get("rag_status", "IN_PROGRESS")
        emoji = RAG_EMOJI.get(rag, "⚪")
        lines += [f"### {emoji} {p.get('title', '')}", "",
                  f"**Status:** `{rag}`", ""]
        if p.get("slide_briefing"):
            lines += ["**Slide Briefing** _(paste into PowerPoint table)_", "",
                      f"> {p['slide_briefing']}", ""]
        for b in (p.get("extended_bullets") or []):
            lines.append(f"- {b}")
        if p.get("extended_bullets"):
            lines.append("")
        lines += ["---", ""]

    # Section 3 — Deep-dive sections
    sections = data.get("sections", [])
    if sections:
        lines += ["## Section 3 — Deep-Dive Sections", ""]
        for s in sections:
            lines += [f"### {s.get('number', '')}. {s.get('title', '')}", ""]
            for b in s.get("bullets", []):
                lines.append(f"- {b}")
            if s.get("right_title"):
                lines += ["", f"**{s['right_title']}**", ""]
                for b in s.get("right_bullets", []):
                    lines.append(f"- {b}")
            if s.get("right_sub_title"):
                lines += ["", f"**{s['right_sub_title']}**", ""]
                for b in s.get("right_sub_bullets", []):
                    lines.append(f"- {b}")
            lines += ["", "---", ""]

    # Analyst flags
    flags = data.get("analyst_flags", [])
    if flags:
        lines += ["## ⚠️ Analyst Flags", "",
                  "_For Doug's attention before the board meeting. "
                  "Not for distribution._", ""]
        for flag in flags:
            lines.append(f"- {flag}")
        lines += ["", "---", ""]

    lines += ["", "_Generated by board_briefing.py v3 — review all content before use._"]
    return "\n".join(lines)


def _generate_briefing_docx(data: dict, docx_path):
    """Generate a formatted Word document from briefing data using python-docx."""
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.section import WD_ORIENT
        from docx.oxml.ns import nsdecls
        from docx.oxml import parse_xml
    except ImportError:
        print("  ❌  python-docx not installed — skipping Word output")
        print("       Install: pip install python-docx")
        return

    print(f"  📝  Generating Word document…", end=" ", flush=True)

    CRIMSON = RGBColor(0xA5, 0x1C, 0x30)
    DARK    = RGBColor(0x1A, 0x1A, 0x1A)
    SLATE   = RGBColor(0x5A, 0x5A, 0x5A)
    GREEN   = RGBColor(0x1E, 0x7B, 0x34)
    AMBER   = RGBColor(0xB3, 0x5A, 0x00)
    RAG_COLORS = {"ON_TRACK": GREEN, "IN_PROGRESS": AMBER, "WATCH": AMBER, "URGENT": CRIMSON}
    RAG_LABELS = {"ON_TRACK": "Done", "IN_PROGRESS": "In Progress", "WATCH": "Watch", "URGENT": "Urgent"}

    meta = data.get("meta", {})
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # Header
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hr = hp.add_run("State Gas Limited — Executive Board Report")
    hr.bold = True; hr.font.size = Pt(9); hr.font.color.rgb = DARK; hr.font.name = "Calibri"
    pPr = hp._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="A51C30"/></w:pBdr>')
    pPr.append(pBdr)

    # Footer
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fr = fp.add_run("CONFIDENTIAL — FOR BOARD USE ONLY")
    fr.italic = True; fr.font.size = Pt(7); fr.font.color.rgb = SLATE; fr.font.name = "Calibri"

    # Cover page
    p = doc.add_paragraph()
    p.space_before = Pt(72)
    r = p.add_run("STATE GAS LIMITED")
    r.bold = True; r.font.size = Pt(14); r.font.color.rgb = CRIMSON; r.font.name = "Calibri"

    p = doc.add_paragraph()
    r = p.add_run(meta.get("report_title", "Executive Board Report"))
    r.bold = True; r.font.size = Pt(36); r.font.color.rgb = DARK; r.font.name = "Calibri"

    meeting = meta.get("meeting_date", "")
    period  = meta.get("period_label", "")
    if meeting or period:
        p = doc.add_paragraph()
        r = p.add_run(f"{period}  |  {meeting}" if period and meeting else period or meeting)
        r.font.size = Pt(14); r.font.color.rgb = SLATE; r.font.name = "Calibri"

    author = meta.get("author", "Doug McAlpine")
    if author:
        p = doc.add_paragraph()
        p.space_before = Pt(24)
        r = p.add_run(f"Author: {author}")
        r.font.size = Pt(11); r.font.color.rgb = SLATE; r.font.name = "Calibri"

    doc.add_page_break()

    # Executive Summary
    h = doc.add_heading("1.    Executive Summary", level=1)
    for run in h.runs:
        run.font.color.rgb = CRIMSON; run.font.name = "Calibri"

    for bullet in data.get("executive_summary_intro", []):
        p = doc.add_paragraph(str(bullet), style="List Bullet")
        p.paragraph_format.space_after = Pt(6)
        for run in p.runs:
            run.font.size = Pt(11); run.font.name = "Calibri"

    for box in data.get("executive_summary_boxes", []):
        p = doc.add_paragraph()
        p.space_before = Pt(12)
        r = p.add_run(box.get("title", ""))
        r.bold = True; r.font.size = Pt(12); r.font.color.rgb = CRIMSON; r.font.name = "Calibri"
        for bullet in box.get("bullets", []):
            bp = doc.add_paragraph(str(bullet), style="List Bullet")
            bp.paragraph_format.space_after = Pt(3)
            for run in bp.runs:
                run.font.size = Pt(10); run.font.name = "Calibri"

    # Strategic Objectives
    h = doc.add_heading("2.    Progress Against Strategic Objectives", level=1)
    for run in h.runs:
        run.font.color.rgb = CRIMSON; run.font.name = "Calibri"

    for pillar in data.get("strategic_pillars", []):
        title = pillar.get("title", "")
        rag   = pillar.get("rag_status", "IN_PROGRESS")
        label = RAG_LABELS.get(rag, rag)
        color = RAG_COLORS.get(rag, AMBER)

        p = doc.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run(f"{title}  ")
        r.bold = True; r.font.size = Pt(12); r.font.name = "Calibri"
        r2 = p.add_run(f"[{label}]")
        r2.bold = True; r2.font.size = Pt(10); r2.font.color.rgb = color; r2.font.name = "Calibri"

        briefing = pillar.get("slide_briefing", "")
        if briefing:
            bp = doc.add_paragraph(str(briefing))
            bp.paragraph_format.space_after = Pt(4)
            for run in bp.runs:
                run.font.size = Pt(10); run.font.name = "Calibri"

        for bullet in pillar.get("extended_bullets", []) or []:
            bp = doc.add_paragraph(str(bullet), style="List Bullet")
            bp.paragraph_format.space_after = Pt(3)
            for run in bp.runs:
                run.font.size = Pt(10); run.font.name = "Calibri"

    # Deep-dive sections
    for sec in data.get("sections", []):
        num   = sec.get("number", "")
        title = sec.get("title", "")
        h = doc.add_heading(f"{num}.    {title}" if num else title, level=1)
        for run in h.runs:
            run.font.color.rgb = CRIMSON; run.font.name = "Calibri"

        for bullet in sec.get("bullets", []):
            p = doc.add_paragraph(str(bullet), style="List Bullet")
            p.paragraph_format.space_after = Pt(6)
            for run in p.runs:
                run.font.size = Pt(11); run.font.name = "Calibri"

        if sec.get("right_title"):
            p = doc.add_paragraph()
            p.space_before = Pt(12)
            r = p.add_run(sec["right_title"])
            r.bold = True; r.underline = True; r.font.size = Pt(11); r.font.name = "Calibri"

        for bullet in sec.get("right_bullets", []):
            p = doc.add_paragraph(str(bullet), style="List Bullet")
            p.paragraph_format.space_after = Pt(4)
            for run in p.runs:
                run.font.size = Pt(10); run.font.name = "Calibri"

    # Analyst flags
    flags = data.get("analyst_flags", [])
    if flags:
        h = doc.add_heading("Notes for MD Review", level=1)
        for run in h.runs:
            run.font.color.rgb = CRIMSON; run.font.name = "Calibri"
        for flag in flags:
            p = doc.add_paragraph(str(flag), style="List Bullet")
            for run in p.runs:
                run.font.size = Pt(10); run.font.color.rgb = SLATE; run.font.name = "Calibri"

    from pathlib import Path
    doc.save(str(docx_path))
    size_kb = Path(docx_path).stat().st_size // 1024
    print(f"✅  ({size_kb} KB)")



def save_outputs(data: dict, emails: list, lookback_days: int):
    today   = datetime.date.today().strftime("%Y%m%d")
    md_path = OUTPUT_DIR / f"StateGas_BoardBriefing_{today}.md"

    md_path.write_text(generate_markdown(data, emails, lookback_days),
                       encoding="utf-8")
    OUTPUT_JSON.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                           encoding="utf-8")

    print(f"\n{'='*60}\n  OUTPUT FILES\n{'='*60}")
    print(f"  📄  Markdown briefing:   {md_path.name}")
    print(f"  📊  JSON (→ PowerPoint): {OUTPUT_JSON.name}")
    print(f"{'='*60}")
    print(f"\n  NEXT: review the markdown, then run:")
    print(f"  node generate_board_deck.js --data {OUTPUT_JSON.name}\n")
    return md_path


# ─────────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────────

def _extract_all(client, emails, context_block, briefing_text, lookback_days):
    batches = [emails[i:i+MAX_EMAILS_PER_BATCH]
               for i in range(0, len(emails), MAX_EMAILS_PER_BATCH)]
    print(f"  🤖  Step 1/2 — extracting matters across {len(batches)} batch(es)…\n")
    all_matters = []
    for i, batch in enumerate(batches, 1):
        overhead    = _estimate_tokens(context_block + briefing_text) + 2000
        sub_batches = _split_oversized_batch(batch, overhead)
        if len(sub_batches) > 1:
            print(f"       Batch {i}/{len(batches)}: split into "
                  f"{len(sub_batches)} sub-batches due to size")
        for si, sub in enumerate(sub_batches):
            label = f"{i}.{si+1}" if len(sub_batches) > 1 else str(i)
            print(f"       Batch {label}/{len(batches)}: {len(sub)} emails … ",
                  end="", flush=True)
            all_matters.extend(extract_matters_from_batch(
                client, sub, i, len(batches),
                context_block, briefing_text, lookback_days))
    return merge_matters(all_matters)


def cmd_run(args):
    lookback = args.days
    print(f"\n{'='*60}")
    print(f"  STATE GAS — BOARD BRIEFING GENERATOR  v3")
    print(f"{'='*60}")
    print(f"  Lookback: {lookback} days")
    print(f"{'='*60}")

    try:
        token = graph.get_access_token()

        ph = PillarHistory()
        ck = CompanyKnowledge()
        ensure_briefing_template()
        briefing_text = load_context_briefing()

        if ph.is_empty():
            print("  ℹ️  No pillar history yet — first run.")
            print("       Tip: 'build-context' gives a richer first output.\n")
        else:
            print(f"  🧠  Pillar history loaded: run #{ph.data.get('run_count', 0)}, "
                  f"updated {ph.data.get('last_updated', '?')}\n")

        # Accumulated context = pillar history + company knowledge
        context_block = (ph.build_prompt_block()
                         + ck.build_prompt_block(label="COMPANY KNOWLEDGE"))

        emails = graph.run_scan(token, lookback,
                                max_per_folder=MAX_EMAILS_PER_FOLDER,
                                max_total=MAX_EMAILS_TOTAL,
                                body_chars=MAX_BODY_CHARS,
                                label="BOARD BRIEFING SCAN",
                                refresh_folders=args.refresh_folders)
        if not emails:
            print(f"  ⚠️  No emails found in the last {lookback} days.")
            return

        client = get_client()
        merged = _extract_all(client, emails, context_block,
                              briefing_text, lookback)
        high = sum(1 for m in merged if m.get("significance") == "high")
        print(f"\n       {len(merged)} matters identified ({high} high significance)\n")

        briefing_data = synthesise_board_briefing(client, merged, emails,
                                                  lookback, context_block)
        if not briefing_data:
            print("  ⚠️  Synthesis failed — check API connection.")
            return

        # Meta fields the deck generator expects
        briefing_data.setdefault("meta", {})
        briefing_data["meta"].setdefault("report_title", "Board Report")
        briefing_data["meta"].setdefault(
            "period_label",
            f"Period to {datetime.date.today().strftime('%d %B %Y')}")
        briefing_data["meta"].setdefault(
            "meeting_date", datetime.date.today().strftime("%d %B %Y"))
        briefing_data["meta"].setdefault("author", MD_NAME)
        briefing_data["meta"].setdefault("author_title", "CEO & Managing Director")

        save_outputs(briefing_data, emails, lookback)

        today = datetime.date.today().strftime("%Y%m%d")
        _generate_briefing_docx(
            briefing_data,
            OUTPUT_DIR / f"StateGas_BoardReport_Draft_{today}.docx")

        update_context_after_run(ph, ck, briefing_data, merged, emails)

    except RuntimeError as e:
        print(f"\n❌  {e}\n")


def cmd_build_context(args):
    """Deep scan to seed pillar history + company knowledge."""
    lookback = args.days if args.days != DEFAULT_LOOKBACK_DAYS else 90
    print(f"\n{'='*60}")
    print(f"  BUILD CONTEXT — scanning {lookback} days of history")
    print(f"{'='*60}\n")
    try:
        token  = graph.get_access_token()
        emails = graph.run_scan(token, lookback,
                                max_per_folder=MAX_EMAILS_PER_FOLDER,
                                max_total=MAX_EMAILS_TOTAL,
                                body_chars=MAX_BODY_CHARS,
                                label="BOARD BRIEFING DEEP SCAN",
                                refresh_folders=args.refresh_folders)
        if not emails:
            print("  ⚠️  No emails found.")
            return

        client        = get_client()
        ph            = PillarHistory()
        ck            = CompanyKnowledge()
        ensure_briefing_template()
        briefing_text = load_context_briefing()

        merged        = _extract_all(client, emails, "", briefing_text, lookback)
        briefing_data = synthesise_board_briefing(client, merged, emails,
                                                  lookback, "")
        if briefing_data:
            update_context_after_run(ph, ck, briefing_data, merged, emails)
            print(f"\n  ✅  Context built from {lookback} days of history.")
            print(f"       py board_briefing.py show-context   to review")
            print(f"       py board_briefing.py run             for your first briefing\n")
    except RuntimeError as e:
        print(f"\n❌  {e}\n")


def cmd_test(args):
    print("\n🔍  Testing Outlook connection…\n")
    try:
        token = graph.get_access_token()
        me    = graph.graph_get(token, "/me")
        print(f"✅  Connected as: {me.get('displayName')} "
              f"({me.get('mail') or me.get('userPrincipalName')})\n")
        folders = graph.get_all_folders(token, refresh=args.refresh_folders)
        print(f"  📁  {len(folders)} folders discovered")
        ph = PillarHistory()
        print(f"\n  🧠  Pillar history: run #{ph.data.get('run_count', 0)}, "
              f"last updated {ph.data.get('last_updated') or 'never'}")
        briefing = load_context_briefing()
        print(f"      context_briefing.txt: "
              f"{len(briefing.splitlines()) if briefing else 0} active instructions")
    except RuntimeError as e:
        print(f"❌  {e}\n")


def main():
    parser = argparse.ArgumentParser(
        description="State Gas board briefing generator v3",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    parser.add_argument("command", nargs="?", default="help",
                        choices=["setup", "test", "run", "build-context",
                                 "show-context", "help"])
    parser.add_argument("--days", type=int, default=DEFAULT_LOOKBACK_DAYS,
                        help=f"Lookback window (default: {DEFAULT_LOOKBACK_DAYS})")
    parser.add_argument("--docx", action="store_true",
                        help="Also generate a Word draft (requires python-docx)")
    parser.add_argument("--refresh-folders", action="store_true",
                        help="Force folder rediscovery")
    args = parser.parse_args()

    if args.command == "setup":
        graph.setup_auth()
    elif args.command == "test":
        cmd_test(args)
    elif args.command == "run":
        cmd_run(args)
    elif args.command == "build-context":
        cmd_build_context(args)
    elif args.command == "show-context":
        PillarHistory().print_summary()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
