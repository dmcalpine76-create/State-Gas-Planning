/**
 * generate_matter_log_docx.js  —  State Gas Matter Log Word Document Generator
 * ──────────────────────────────────────────────────────────────────────────────
 * Reads StateGas_MatterLog_YYYYMMDD.md and produces a landscape A4 Word
 * document with a branded table — same content as the PowerPoint version,
 * formatted for editing and annotation in Word.
 *
 * Usage:
 *   node generate_matter_log_docx.js                              (auto-finds latest)
 *   node generate_matter_log_docx.js --data StateGas_MatterLog_20260505.md
 *   node generate_matter_log_docx.js --out MyMatterLog.docx
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");
const fs   = require("fs");
const path = require("path");

// ─── CLI args ────────────────────────────────────────────────────────────────
const args    = process.argv.slice(2);
const dataArg = args.indexOf("--data");
const outArg  = args.indexOf("--out");

let dataFile = dataArg !== -1 ? args[dataArg + 1] : null;
if (!dataFile) {
  const files = fs.readdirSync(__dirname)
    .filter(f => f.startsWith("StateGas_MatterLog_") && f.endsWith(".md"))
    .sort().reverse();
  if (!files.length) {
    console.error("❌  No StateGas_MatterLog_*.md found. Run: py matter_log.py run");
    process.exit(1);
  }
  dataFile = path.join(__dirname, files[0]);
  console.log(`  Auto-selected: ${files[0]}`);
}
if (!fs.existsSync(dataFile)) {
  console.error(`❌  File not found: ${dataFile}`);
  process.exit(1);
}

const baseName  = path.basename(dataFile, ".md");
const defaultOut = path.join(__dirname, baseName + ".docx");
const OUTPUT    = outArg !== -1 ? args[outArg + 1] : defaultOut;

// ─── Brand colours (hex without #) ──────────────────────────────────────────
const C = {
  dark:      "1A1A1A",
  crimson:   "A51C30",
  crimsonLt: "C8293F",
  white:     "FFFFFF",
  lightGrey: "E8E8E8",
  rowAlt:    "FAF5F5",
  green:     "1A6B38",
  amber:     "B35C00",
  red:       "A51C30",
  textDark:  "1A1A1A",
  textMid:   "5A5A5A",
  slate:     "5A5A5A",
};

const SIG_COLORS = {
  high:   { bg: "A51C30", fg: "FFFFFF" },
  medium: { bg: "B35C00", fg: "FFFFFF" },
  low:    { bg: "1A6B38", fg: "FFFFFF" },
};

const AREA_LABELS = {
  commercial:  "Commercial & BD",
  regulatory:  "Regulatory",
  legal:       "Legal",
  financial:   "Financial",
  jv_partner:  "JV / Partner",
  operational: "Operational",
  stakeholder: "Stakeholder & IR",
  other:       "Other",
};
const AREA_ORDER = ["commercial","regulatory","legal","financial","jv_partner","operational","stakeholder","other"];

// ─── Landscape A4 dimensions (DXA, 1440 = 1 inch) ───────────────────────────
// Per skill: pass portrait dimensions + LANDSCAPE orientation
// A4 portrait: 11906 wide × 16838 tall
// Landscape content width = 16838 - 720 (left 0.5") - 720 (right 0.5") = 15398 DXA
const PAGE_W   = 11906;
const PAGE_H   = 16838;
const MARGIN   = 720;   // 0.5" all sides
const CONTENT_W = PAGE_H - MARGIN * 2;  // 15398 DXA (landscape content)

// Column widths (DXA, must sum to CONTENT_W = 15398)
// Sig: 0.45", Area: 1.1", Matter: 1.8", Update: 6.7", Next Steps: 2.15", Dec: 0.55"
const COL = {
  sig:       648,    // 0.45"
  area:      1584,   // 1.1"
  matter:    2592,   // 1.8"
  update:    7776,   // ~5.4" — gets the most space
  next:      2232,   // 1.55"
  dec:       566,    // 0.39"
};
// Sum: 648+1584+2592+7776+2232+566 = 15398 ✓
const COL_WIDTHS = [COL.sig, COL.area, COL.matter, COL.update, COL.next, COL.dec];

// ─── Markdown parser (shared with PowerPoint generator) ──────────────────────
function parseMatterLog(mdText) {
  const lines   = mdText.split(/\r?\n/);
  const matters = [];
  let current   = null;
  let section   = "other";
  let meta      = { period: "", generated: "", count: "" };

  for (const line of lines) {
    if (line.includes("**Period:**"))
      meta.period = line.replace(/.*\*\*Period:\*\*/, "").replace(/\s+/g, " ").trim();
    if (line.includes("**Generated:**"))
      meta.generated = line.replace(/.*\*\*Generated:\*\*/, "").replace(/\s+/g, " ").trim();
    if (line.includes("**Matters identified:**"))
      meta.count = line.replace(/.*\*\*Matters identified:\*\*/, "").replace(/\s+/g, " ").trim();
  }

  for (const line of lines) {
    if (/^## [^⚡]/.test(line)) {
      section = "other";
      if      (line.includes("Commercial"))  section = "commercial";
      else if (line.includes("Regul"))       section = "regulatory";
      else if (line.includes("Legal"))       section = "legal";
      else if (line.includes("Financ"))      section = "financial";
      else if (line.includes("JV") || line.includes("Partner")) section = "jv_partner";
      else if (line.includes("Operat"))      section = "operational";
      else if (line.includes("Stake") || line.includes("Investor")) section = "stakeholder";
      continue;
    }

    if (/^### /.test(line)) {
      if (current) matters.push(current);
      const raw        = line.replace(/^### /, "").trim();
      const sig        = raw.startsWith("🔴") ? "high" : raw.startsWith("🟡") ? "medium" : raw.startsWith("🟢") ? "low" : "medium";
      const requiresDec = raw.includes("⚡");
      const title      = raw.replace(/^[🔴🟡🟢]\s*/, "").replace(/⚡.*$/, "").replace(/_Decision required_/i, "").trim();
      current = { area: section, significance: sig, matter: title, requires_decision: requiresDec,
                  date_range: "", key_parties: "", key_figures: "", narrative: "",
                  next_steps: "", doug_note: "", decision_detail: "" };
      continue;
    }
    if (!current) continue;

    if (/^_[^_]/.test(line) && line.includes("Parties:")) {
      const parts = line.split("  ");
      for (const p of parts) {
        if (p.startsWith("_") && !p.includes("Parties")) current.date_range = p.replace(/_/g, "").trim();
        if (p.includes("**Parties:**")) current.key_parties = p.replace("**Parties:**", "").trim();
      }
      continue;
    }
    if (line.startsWith("**Key figures:**")) { current.key_figures = line.replace("**Key figures:**", "").trim(); continue; }
    if (line.startsWith("> **Decision required:**")) { current.decision_detail = line.replace(/^> \*\*Decision required:\*\*/, "").trim(); continue; }
    if (line.startsWith("**Next steps:**")) { current.next_steps = line.replace("**Next steps:**", "").trim(); continue; }

    if (line.trim() && !line.startsWith("#") && !line.startsWith(">") &&
        !line.startsWith("**") && !line.startsWith("_") && !line.startsWith("<") &&
        !line.startsWith("-") && !line.startsWith("|") && !line.startsWith("!") && line !== "---") {
      if (!current.next_steps)
        current.narrative = current.narrative ? current.narrative + " " + line.trim() : line.trim();
    }
  }
  if (current) matters.push(current);
  return { matters, meta };
}

// ─── Truncate helper ─────────────────────────────────────────────────────────
function trunc(str, max) {
  if (!str) return "";
  str = String(str).replace(/\s+/g, " ").trim();
  return str.length <= max ? str : str.slice(0, max - 1) + "…";
}

// ─── Cell builders ───────────────────────────────────────────────────────────
const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellPad = { top: 72, bottom: 72, left: 108, right: 108 };  // 0.05" top/bottom, 0.075" sides

function makeCell(width, paragraphs, opts = {}) {
  return new TableCell({
    width:   { size: width, type: WidthType.DXA },
    borders: opts.noBorder ? noBorders : borders,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    verticalAlign: opts.vAlign || VerticalAlign.TOP,
    margins: cellPad,
    children: Array.isArray(paragraphs) ? paragraphs : [paragraphs],
  });
}

function run(text, opts = {}) {
  return new TextRun({
    text:      String(text || ""),
    bold:      opts.bold   || false,
    italics:   opts.italic || false,
    color:     opts.color  || C.textDark,
    size:      opts.size   || 18,   // 9pt (size in half-points)
    font:      "Aptos",
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing:   { before: opts.before || 0, after: opts.after || 0 },
    children:  Array.isArray(children) ? children : [children],
  });
}

// ─── Header row ──────────────────────────────────────────────────────────────
function makeHeaderRow() {
  const hdr = (text, width) => makeCell(width,
    para(run(text, { bold: true, color: C.white, size: 16 }),
         { align: AlignmentType.LEFT }),
    { fill: C.dark, vAlign: VerticalAlign.CENTER }
  );
  return new TableRow({
    tableHeader: true,
    children: [
      hdr("Sig",         COL.sig),
      hdr("Area",        COL.area),
      hdr("Matter",      COL.matter),
      hdr("Update",      COL.update),
      hdr("Next Steps",  COL.next),
      hdr("Dec?",        COL.dec),
    ]
  });
}

// ─── Data row ────────────────────────────────────────────────────────────────
function makeDataRow(m, rowIndex) {
  const sigConf = SIG_COLORS[m.significance] || SIG_COLORS.medium;
  const sigLabel = { high: "HIGH", medium: "MED", low: "LOW" }[m.significance] || "MED";
  const rowBg   = rowIndex % 2 === 0 ? C.white : C.rowAlt;
  const decReq  = m.requires_decision;

  // Narrative: combine narrative + key figures
  const rawNarr = [
    m.narrative,
    m.key_figures ? `[${m.key_figures}]` : "",
  ].filter(Boolean).join(" ").replace(/\s+/g, " ").trim();
  const narrativeText = trunc(rawNarr, 1200);

  // Next steps: combine with decision detail
  const rawNext = [
    m.next_steps,
    m.decision_detail ? `Decision required: ${m.decision_detail}` : "",
  ].filter(Boolean).join(" ").trim();
  const nextText = trunc(rawNext, 400);

  // Sig badge cell
  const sigCell = makeCell(COL.sig,
    para(run(sigLabel, { bold: true, color: sigConf.fg, size: 16 }),
         { align: AlignmentType.CENTER }),
    { fill: sigConf.bg, vAlign: VerticalAlign.CENTER }
  );

  // Area cell (italic, muted)
  const areaCell = makeCell(COL.area,
    para(run(AREA_LABELS[m.area] || m.area, { italic: true, color: C.textMid, size: 16 })),
    { fill: rowBg }
  );

  // Matter title — bold, crimson for high
  const matterColor = m.significance === "high" ? C.crimson : C.textDark;
  const matterCell  = makeCell(COL.matter, [
    para(run(m.matter, { bold: true, color: matterColor, size: 18 })),
    m.date_range ? para(run(m.date_range, { italic: true, color: C.slate, size: 14 }),
                        { before: 40 }) : null,
    m.key_parties ? para(run(`Parties: ${trunc(m.key_parties, 120)}`,
                             { color: C.slate, size: 14 }), { before: 20 }) : null,
  ].filter(Boolean), { fill: rowBg });

  // Narrative
  const narrativeCell = makeCell(COL.update, [
    para(run(narrativeText, { size: 18 })),
  ], { fill: rowBg });

  // Next steps — bold crimson if decision required
  const nextCell = makeCell(COL.next, [
    para(run(nextText, { size: 18, bold: decReq, color: decReq ? C.crimson : C.textDark })),
  ], { fill: rowBg });

  // Decision flag
  const decCell = makeCell(COL.dec,
    para(run(decReq ? "⚡ YES" : "", { bold: true, color: decReq ? C.white : C.textMid, size: 16 }),
         { align: AlignmentType.CENTER }),
    { fill: decReq ? C.crimson : rowBg, vAlign: VerticalAlign.CENTER }
  );

  return new TableRow({ children: [sigCell, areaCell, matterCell, narrativeCell, nextCell, decCell] });
}

// ─── Section header row (area break) ─────────────────────────────────────────
function makeAreaRow(label) {
  return new TableRow({
    children: [
      new TableCell({
        columnSpan: 6,
        width:      { size: CONTENT_W, type: WidthType.DXA },
        shading:    { fill: C.crimson, type: ShadingType.CLEAR },
        borders,
        margins:    cellPad,
        children:   [para(run(label.toUpperCase(), { bold: true, color: C.white, size: 18 }))],
      })
    ]
  });
}

// ─── Cover page content ───────────────────────────────────────────────────────
function makeCoverPage(meta, matterCount, highCount, decCount) {
  const spacer = new Paragraph({ spacing: { before: 0, after: 200 }, children: [] });
  return [
    new Paragraph({
      children: [run("STATE GAS LIMITED", { bold: true, color: C.crimson, size: 24 })],
      spacing: { before: 720, after: 80 },
    }),
    new Paragraph({
      children: [run("Matter Log", { bold: true, color: C.dark, size: 56 })],
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      children: [run(meta.period || "", { color: C.textMid, size: 24 })],
      spacing: { before: 0, after: 400 },
    }),
    // Stats table
    new Table({
      width:        { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W / 3, CONTENT_W / 3, CONTENT_W / 3].map(Math.round),
      rows: [
        new TableRow({ children: [
          makeCell(Math.round(CONTENT_W/3), para(run(String(matterCount), { bold: true, color: C.dark,    size: 72 }), { align: AlignmentType.CENTER }), { fill: C.lightGrey }),
          makeCell(Math.round(CONTENT_W/3), para(run(String(highCount),   { bold: true, color: C.crimson, size: 72 }), { align: AlignmentType.CENTER }), { fill: C.lightGrey }),
          makeCell(Math.round(CONTENT_W/3), para(run(String(decCount),    { bold: true, color: C.crimson, size: 72 }), { align: AlignmentType.CENTER }), { fill: C.lightGrey }),
        ]}),
        new TableRow({ children: [
          makeCell(Math.round(CONTENT_W/3), para(run("Matters",           { italic: true, color: C.textMid, size: 20 }), { align: AlignmentType.CENTER }), { fill: C.lightGrey }),
          makeCell(Math.round(CONTENT_W/3), para(run("High significance", { italic: true, color: C.textMid, size: 20 }), { align: AlignmentType.CENTER }), { fill: C.lightGrey }),
          makeCell(Math.round(CONTENT_W/3), para(run("Decisions required",{ italic: true, color: C.textMid, size: 20 }), { align: AlignmentType.CENTER }), { fill: C.lightGrey }),
        ]}),
      ]
    }),
    new Paragraph({ spacing: { before: 400, after: 200 }, children: [] }),
    new Paragraph({
      children: [run("WORKING DOCUMENT — NOT FOR DISTRIBUTION", { bold: true, color: C.crimson, size: 18 })],
      spacing: { before: 0, after: 0 },
      border:  { top:    { style: BorderStyle.SINGLE, size: 4, color: C.crimson, space: 4 },
                 bottom: { style: BorderStyle.SINGLE, size: 4, color: C.crimson, space: 4 } },
    }),
    new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),
  ];
}

// ─── Main ────────────────────────────────────────────────────────────────────
async function build() {
  console.log(`\n${"═".repeat(56)}`);
  console.log(`  STATE GAS — MATTER LOG WORD DOCUMENT GENERATOR`);
  console.log(`${"═".repeat(56)}`);
  console.log(`  Input:  ${path.basename(dataFile)}`);
  console.log(`  Output: ${path.basename(OUTPUT)}`);
  console.log(`${"═".repeat(56)}\n`);

  const mdText  = fs.readFileSync(dataFile, "utf8");
  const { matters, meta } = parseMatterLog(mdText);

  if (!matters.length) {
    console.error("❌  No matters parsed.");
    process.exit(1);
  }

  // Sort: decisions first, then significance, then area order, then alpha
  const sigOrder  = { high: 0, medium: 1, low: 2 };
  const areaOrder = Object.fromEntries(AREA_ORDER.map((a, i) => [a, i]));
  matters.sort((a, b) => {
    if (a.requires_decision !== b.requires_decision) return a.requires_decision ? -1 : 1;
    const sd = sigOrder[a.significance] - sigOrder[b.significance]; if (sd) return sd;
    const ad = (areaOrder[a.area]??99) - (areaOrder[b.area]??99); if (ad) return ad;
    return a.matter.localeCompare(b.matter);
  });

  const highCount = matters.filter(m => m.significance === "high").length;
  const decCount  = matters.filter(m => m.requires_decision).length;

  console.log(`  Matters:     ${matters.length}`);
  console.log(`  High:        ${highCount}`);
  console.log(`  Decisions:   ${decCount}\n`);

  // Build table rows with area break rows
  const tableRows = [makeHeaderRow()];
  let   lastArea  = null;
  let   rowIndex  = 0;

  for (const m of matters) {
    if (m.area !== lastArea) {
      tableRows.push(makeAreaRow(AREA_LABELS[m.area] || m.area));
      lastArea = m.area;
    }
    tableRows.push(makeDataRow(m, rowIndex++));
  }

  // Header and footer
  const hdr = new Header({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          run("State Gas Limited — Matter Log", { bold: true, color: C.dark, size: 18 }),
          new TextRun({ text: "\t", size: 18 }),
          run(meta.period || "", { color: C.textMid, size: 16 }),
        ],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.crimson, space: 2 } },
        spacing: { after: 120 },
      })
    ]
  });

  const ftr = new Footer({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          run("WORKING DOCUMENT — NOT FOR DISTRIBUTION", { italic: true, color: C.textMid, size: 14 }),
          new TextRun({ text: "\t", size: 14 }),
          run("Page ", { color: C.textMid, size: 14 }),
          new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.textMid }),
          run(" of ", { color: C.textMid, size: 14 }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 14, color: C.textMid }),
        ],
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.lightGrey, space: 2 } },
        spacing: { before: 120 },
      })
    ]
  });

  // Build document
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Aptos", size: 18 } }
      }
    },
    sections: [
      // ── Cover page section (portrait for cover is fine in landscape doc)
      {
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.LANDSCAPE },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          }
        },
        headers: { default: hdr },
        footers: { default: ftr },
        children: makeCoverPage(meta, matters.length, highCount, decCount),
      },
      // ── Main table section
      {
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.LANDSCAPE },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN + 200, left: MARGIN },
          }
        },
        headers: { default: hdr },
        footers: { default: ftr },
        children: [
          new Table({
            width:        { size: CONTENT_W, type: WidthType.DXA },
            columnWidths: COL_WIDTHS,
            rows:         tableRows,
          }),
        ],
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT, buffer);
  console.log(`✅  Saved: ${OUTPUT}`);
  console.log(`    ${matters.length} matters across ${new Set(matters.map(m=>m.area)).size} areas\n`);
}

build().catch(e => { console.error("❌  Error:", e.message); process.exit(1); });
