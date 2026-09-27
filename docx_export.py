"""
docx_export.py — Word-document draft of the board briefing.

Ported from the retired board_report_analyser.py and adapted to the v3
synthesis schema (executive_summary_intro/boxes, slide_briefing,
extended_bullets, sections, analyst_flags).

Requires Node.js with the 'docx' package (same dependency as the
PowerPoint generator):  npm install docx
"""

import json
import datetime
import subprocess
from pathlib import Path

from core.config import BASE_DIR, COMPANY_NAME

DOCX_TEMPLATE = r"""
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
        VerticalAlign, LevelFormat, PageBreak } = require('docx');
const fs = require('fs');

const data = __DATA_JSON__;
const meta = { emails: __N_EMAILS__, days: __N_DAYS__ };

const C = { navy: "0D1E3D", teal: "0D7377", amber: "D97706",
            green: "22863A", red: "DC2626", grey: "6B7280" };

const ragColor = (rag) => ({ ON_TRACK: C.green, IN_PROGRESS: C.teal,
                             WATCH: C.amber, URGENT: C.red }[rag] || C.grey);
const ragLabel = (rag) => (rag || "").replace("_", " ");

const border  = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };

const heading1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, bold: true, font: "Arial", size: 32, color: C.navy })],
  spacing: { before: 360, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.teal, space: 1 } },
});

const heading2 = (text, color) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, bold: true, font: "Arial", size: 26, color: color || C.navy })],
  spacing: { before: 240, after: 120 },
});

const para = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text: String(text || ""), font: "Arial",
    size: opts.size || 20, color: opts.color || "1A202C",
    bold: opts.bold || false, italics: opts.italic || false })],
  spacing: { before: opts.before || 80, after: opts.after || 80 },
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text: String(text || ""), font: "Arial", size: 20 })],
  spacing: { before: 40, after: 40 },
});

// ── Cover ────────────────────────────────────────────────────────────────────
const coverSection = [
  new Paragraph({ spacing: { before: 1440, after: 0 } }),
  new Paragraph({
    children: [new TextRun({ text: "__COMPANY_UPPER__", bold: true, font: "Arial",
      size: 28, color: C.teal, characterSpacing: 60 })],
    spacing: { before: 0, after: 240 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Board Report — Draft for Review", bold: true,
      font: "Arial", size: 56, color: C.navy })],
    spacing: { before: 0, after: 360 },
  }),
  para(`Period: ${data.period_label || (data.meta && data.meta.period_label) || ""}`, { size: 24, color: C.grey }),
  para(`Emails reviewed: ${meta.emails} over ${meta.days} days`, { size: 24, color: C.grey }),
  para(`Generated: ${new Date().toLocaleDateString('en-AU', { day: '2-digit', month: 'long', year: 'numeric' })}`, { size: 24, color: C.grey }),
  new Paragraph({ spacing: { before: 240, after: 0 } }),
  new Paragraph({
    children: [new TextRun({ text: "DRAFT — AI-ASSISTED ANALYSIS. REVIEW BEFORE USE.",
      bold: true, font: "Arial", size: 20, color: C.red })],
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.red },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: C.red } },
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── Analyst flags ────────────────────────────────────────────────────────────
const flags = data.analyst_flags || [];
const flagsSection = flags.length ? [
  heading1("Analyst Flags"),
  para("For Doug's attention before the meeting — not for distribution.",
       { italic: true, color: C.grey }),
  ...flags.map(f => bullet(f)),
  new Paragraph({ spacing: { before: 200, after: 0 } }),
] : [];

// ── Executive summary ────────────────────────────────────────────────────────
const execSection = [
  heading1("Executive Summary"),
  ...(data.executive_summary_intro || []).map(b => bullet(b)),
  ...(data.executive_summary_boxes || []).map(box => [
    heading2(box.title || ""),
    ...(box.bullets || []).map(b => bullet(b)),
  ]).flat(),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── Strategic pillars ────────────────────────────────────────────────────────
const pillarSections = (data.strategic_pillars || []).map(p => [
  heading1(p.title || ""),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1500, 7860],
    rows: [new TableRow({ children: [
      new TableCell({
        borders, width: { size: 1500, type: WidthType.DXA },
        shading: { fill: ragColor(p.rag_status), type: ShadingType.CLEAR },
        verticalAlign: VerticalAlign.MIDDLE,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: ragLabel(p.rag_status), bold: true,
            font: "Arial", size: 18, color: "FFFFFF" })],
        })],
      }),
      new TableCell({
        borders, width: { size: 7860, type: WidthType.DXA },
        shading: { fill: "F8FAFB", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: p.slide_briefing || "", font: "Arial",
            size: 20, color: "4A5568" })],
        })],
      }),
    ]})],
  }),
  new Paragraph({ spacing: { before: 160, after: 0 } }),
  ...((p.extended_bullets || []).map(b => bullet(b))),
]).flat();

// ── Deep-dive sections ───────────────────────────────────────────────────────
const diveSections = (data.sections || []).map(s => [
  heading1(`${s.number != null ? s.number + ". " : ""}${s.title || ""}`),
  ...(s.bullets || []).map(b => bullet(b)),
  s.right_title ? heading2(s.right_title, C.teal) : null,
  ...((s.right_bullets || []).map(b => bullet(b))),
  s.right_sub_title ? heading2(s.right_sub_title, C.teal) : null,
  ...((s.right_sub_bullets || []).map(b => bullet(b))),
]).flat().filter(Boolean);

const doc = new Document({
  numbering: { config: [{
    reference: "bullets",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
  }]},
  styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
  sections: [{
    properties: { page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
    }},
    children: [
      ...coverSection, ...flagsSection, ...execSection,
      ...pillarSections,
      new Paragraph({ children: [new PageBreak()] }),
      ...diveSections,
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("__OUTPUT_PATH__", buffer);
  console.log("Word document saved: __OUTPUT_PATH__");
}).catch(e => { console.error("Error:", e.message); process.exit(1); });
"""


def generate_word_doc(briefing_data: dict, n_emails: int, lookback_days: int,
                      output_path: Path):
    """Render the briefing data to a .docx via a temporary Node script."""
    js_code = (DOCX_TEMPLATE
               .replace("__DATA_JSON__", json.dumps(briefing_data, indent=2,
                                                    ensure_ascii=False))
               .replace("__N_EMAILS__", str(n_emails))
               .replace("__N_DAYS__", str(lookback_days))
               .replace("__COMPANY_UPPER__", COMPANY_NAME.upper())
               .replace("__OUTPUT_PATH__", str(output_path).replace("\\", "/")))
    js_path = BASE_DIR / ".board_docx_gen.js"
    js_path.write_text(js_code, encoding="utf-8")

    print(f"  📝  Generating Word document…")
    try:
        result = subprocess.run(["node", str(js_path)],
                                capture_output=True, text=True, timeout=120)
        if result.returncode == 0:
            print(f"  ✅  {result.stdout.strip()}")
        else:
            print(f"  ❌  {result.stderr.strip()}")
            print(f"      (Is the 'docx' npm package installed?  npm install docx)")
    except FileNotFoundError:
        print("  ❌  node not found — install Node.js to use --docx.")
    finally:
        js_path.unlink(missing_ok=True)
