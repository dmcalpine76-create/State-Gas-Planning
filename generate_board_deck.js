/**
 * generate_board_deck.js  —  State Gas Executive Board Report Generator
 * -----------------------------------------------------------------------
 * Generates a board deck matching the May 2026 format established by Doug:
 *   Slide 1:  Title (dark, logo, "Executive Board Report")
 *   Slide 2:  1. Executive Summary (3 bullets + 4 themed sub-boxes)
 *   Slide 3:  2. Progress Against Strategic Objectives (RAG pillars)
 *   Slide 4+: Deep-dive sections from board_decisions and other_matters
 *
 * Usage:
 *   node generate_board_deck.js
 *   node generate_board_deck.js --data briefing.json
 *   node generate_board_deck.js --out MyDeck.pptx
 *   node generate_board_deck.js --sample
 */

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");
const { FaCheckCircle, FaExclamationTriangle, FaTimesCircle } = require("react-icons/fa");

// ── CLI ──────────────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
const getArg = (f) => { const i = args.indexOf(f); return i !== -1 && args[i+1] ? args[i+1] : null; };
const useSample  = args.includes("--sample");
const dataFile   = getArg("--data") || "board_report_draft.json";
const outputFile = getArg("--out")  || null;

// ── State Gas logo (embedded — no external file dependency) ──────────────────
// Extracted from StateGas_Board_Report_May_2026.pptx
const LOGO_B64 = fs.existsSync(path.join(__dirname, "stategas_logo.png"))
  ? "image/png;base64," + fs.readFileSync(path.join(__dirname, "stategas_logo.png")).toString("base64")
  : null;

// ── Icons ────────────────────────────────────────────────────────────────────
function renderSvg(Icon, color, size=256) {
  return ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
}
async function iconPng(Icon, color) {
  return "image/png;base64," + (await sharp(Buffer.from(renderSvg(Icon, color))).png().toBuffer()).toString("base64");
}

// ── Brand palette ─────────────────────────────────────────────────────────────
const C = {
  dark:     "1A1A1A",   // near-black background (title slide)
  crimson:  "A51C30",   // State Gas crimson — header bars, badges
  crimsonLt:"C8293F",   // lighter crimson for sub-labels
  bgLight:  "F2F2F2",   // content slide background (light grey)
  white:    "FFFFFF",
  lgrey:    "E0E0E0",   // borders, dividers
  green:    "1E7B34",   // RAG green (Done)
  amber:    "B35A00",   // RAG amber (In Progress / Watch)
  red:      "A51C30",   // RAG red (Urgent)
  textDark: "1A1A1A",
  textMid:  "555555",
  textLight:"888888",
};

const RAG_COLOR = { ON_TRACK:"1E7B34", IN_PROGRESS:"B35A00", WATCH:"B35A00", URGENT:"A51C30" };
const RAG_LABEL = { ON_TRACK:"Done", IN_PROGRESS:"IN PROGRESS", WATCH:"WATCH", URGENT:"URGENT" };
const sh = () => ({ type:"outer", blur:6, offset:2, angle:135, color:"000000", opacity:0.10 });

// ── Sample / fallback data ────────────────────────────────────────────────────
// Default slide content lives in the knowledge store (board/deck_defaults.json),
// not in this file, so the code can be public. Fields missing from a run fall
// back to it, exactly as before.
const SAMPLE = (() => {
  const dir = process.env.STATEGAS_STORE_DIR || path.join(__dirname, "..", "state gas knowledge");
  try { return JSON.parse(fs.readFileSync(path.join(dir, "board", "deck_defaults.json"), "utf8")); }
  catch { console.log("  deck_defaults.json not found in the knowledge store - no default content"); return {}; }
})();

// ── Load data ─────────────────────────────────────────────────────────────────
function loadData() {
  if (useSample) { console.log("  Using built-in sample data (--sample)"); return SAMPLE; }
  const candidates = [
    dataFile, path.join(__dirname, dataFile),
    ...(() => { try { return fs.readdirSync(__dirname).filter(f => f.startsWith("StateGas_BoardBriefing_") && f.endsWith(".json")).sort().reverse().map(f => path.join(__dirname, f)); } catch(e) { return []; } })(),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      try { const d = JSON.parse(fs.readFileSync(p, "utf8")); console.log("  Data loaded: " + path.basename(p)); return { ...SAMPLE, ...d }; }
      catch (e) { console.error("  Could not parse " + p + ": " + e.message); }
    }
  }
  console.log("  No JSON data file found — using built-in sample data.");
  return SAMPLE;
}

function outPath(data) {
  if (outputFile) return outputFile;
  const months = {January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12};
  let ds = new Date().getFullYear() + String(new Date().getMonth()+1).padStart(2,"0");
  const pt = data.meeting_date || data.period_to || "";
  if (pt) { const m = pt.match(/(\w+)\s+(\d{4})/); if (m && months[m[1]]) ds = m[2] + String(months[m[1]]).padStart(2,"0"); }
  return path.join(__dirname, "StateGas_BoardReport_" + ds + ".pptx");
}

// ── Layout constants ──────────────────────────────────────────────────────────
const FONT  = "Calibri";   // matches actual board report font
const W     = 10;          // slide width inches
const H     = 5.625;       // slide height inches
const HDR_H = 0.55;        // header bar height
const LOGO_W = 1.35;
const LOGO_H = 0.25;
const LOGO_X = W - LOGO_W - 0.15;
const LOGO_Y = 0.12;
const FOOTER_Y = H - 0.28;
const FOOTER_H = 0.28;
const CONTENT_Y = HDR_H + 0.18;  // top of content area
const CONTENT_H = FOOTER_Y - CONTENT_Y - 0.1;

// ── Shared helpers ────────────────────────────────────────────────────────────

/** Crimson header bar: "N.   Title" */
function addHeader(s, pres, num, title) {
  s.background = { color: C.bgLight };
  s.addShape(pres.shapes.RECTANGLE, {
    x:0, y:0, w:W, h:HDR_H,
    fill:{color:C.crimson}, line:{color:C.crimson}
  });
  const label = num ? num + ".    " + title : title;
  s.addText(label, {
    x:0.3, y:0, w: LOGO_W ? (W - LOGO_W - 0.6) : W - 0.6, h:HDR_H,
    fontSize:16, bold:true, color:C.white, fontFace:FONT,
    valign:"middle", margin:0,
  });
  if (LOGO_B64) {
    s.addImage({ data: LOGO_B64, x: LOGO_X, y: LOGO_Y, w: LOGO_W, h: LOGO_H });
  }
  s.addShape(pres.shapes.RECTANGLE, {
    x:0, y:FOOTER_Y, w:W, h:FOOTER_H,
    fill:{color:C.crimson}, line:{color:C.crimson}
  });
  s.addText("CONFIDENTIAL \u2013 FOR BOARD USE ONLY", {
    x:0.3, y:FOOTER_Y, w:W-0.6, h:FOOTER_H,
    fontSize:8, bold:true, color:C.white, fontFace:FONT,
    valign:"middle", margin:0,
  });
}

/** Add a bullet list to a text box (returns pptxgenjs options) */
function bulletText(lines, opts={}) {
  return {
    text: lines.map(l => ({ text: l, options: { bullet: { type: "bullet" } } })),
    options: {
      x: opts.x || 0.3, y: opts.y || CONTENT_Y,
      w: opts.w || W - 0.6, h: opts.h || CONTENT_H,
      fontSize: opts.fontSize || 11,
      color: opts.color || C.textDark,
      fontFace: FONT,
      valign: "top",
      margin: [4, 0, 4, 8],
      ...opts,
    }
  };
}

// ── Build deck ────────────────────────────────────────────────────────────────
async function build() {
  console.log("\n" + "=".repeat(56));
  console.log("  STATE GAS BOARD DECK GENERATOR");
  console.log("=".repeat(56));

  const D = loadData();
  const meetingDate = D.meeting_date || D.period_to || "";
  const periodLabel = D.period_label || "";

  // Pre-render RAG icons
  const iGreen = await iconPng(FaCheckCircle,         "#1E7B34");
  const iAmber = await iconPng(FaExclamationTriangle, "#B35A00");
  const iRed   = await iconPng(FaTimesCircle,         "#A51C30");
  const ragIco = r => r === "ON_TRACK" ? iGreen : r === "URGENT" ? iRed : iAmber;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = "State Gas Executive Board Report " + meetingDate;

  // ════════════════════════════════════════════════════════
  // SLIDE 1 — TITLE
  // Matches: dark bg, logo top-right, large bold title,
  //          period + author line, crimson bottom bar only
  // ════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Logo — top right
    if (LOGO_B64) {
      s.addImage({ data: LOGO_B64, x: W - 2.1 - 0.2, y: 0.2, w: 2.1, h: 0.39 });
    }

    // Large title — two lines matching the actual deck
    s.addText(D.report_title || "Executive\nBoard Report", {
      x:0.5, y:1.5, w:7.5, h:2.5,
      fontSize:64, bold:true, color:C.white, fontFace:FONT,
      valign:"middle", margin:0,
    });

    // Period | date line
    s.addText(periodLabel + (periodLabel && meetingDate ? "  |  " : "") + meetingDate, {
      x:0.5, y:4.1, w:6, h:0.4,
      fontSize:14, color:"D0D0D0", fontFace:FONT, margin:0,
    });

    // Crimson rule under the period line
    s.addShape(pres.shapes.RECTANGLE, {
      x:0.5, y:4.55, w:3.2, h:0.05,
      fill:{color:C.crimson}, line:{color:C.crimson}
    });

    // Author line
    s.addText([{text:"Author: ", options:{bold:true}}, {text:D.author||"Doug McAlpine"}], {
      x:0.5, y:4.65, w:5, h:0.32,
      fontSize:12, color:"D0D0D0", fontFace:FONT, margin:0,
    });

    // Crimson bottom bar with confidential footer
    s.addShape(pres.shapes.RECTANGLE, {
      x:0, y:H-0.35, w:W, h:0.35,
      fill:{color:C.crimson}, line:{color:C.crimson}
    });
    s.addText("CONFIDENTIAL \u2013 FOR BOARD USE ONLY", {
      x:0.3, y:H-0.35, w:W-0.6, h:0.35,
      fontSize:9, bold:true, color:C.white, fontFace:FONT,
      valign:"middle", margin:0,
    });
  }

  // ════════════════════════════════════════════════════════
  // SLIDE 2 — 1. EXECUTIVE SUMMARY
  // Three top-level bullets + 2×2 grid of themed sub-boxes
  // ════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    addHeader(s, pres, "1", "Executive Summary");

    const intro   = D.executive_summary_intro || [];
    const boxes   = D.executive_summary_boxes || [];

    // Three intro bullets — left-aligned, above the boxes
    const introBullets = intro.slice(0, 3).map(l => ({ text: l, options: { bullet: { type:"bullet" } } }));
    s.addText(introBullets.length ? introBullets : [{ text:"Key period summary to be inserted.", options:{} }], {
      x:0.3, y:CONTENT_Y, w:W-0.6, h:1.1,
      fontSize:11, color:C.textDark, fontFace:FONT,
      valign:"top", margin:[3,0,3,8],
    });

    // 2×2 sub-boxes
    const boxY = CONTENT_Y + 1.15;
    const boxH = FOOTER_Y - boxY - 0.08;
    const boxW = (W - 0.7) / 2;
    boxes.slice(0,4).forEach((box, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const bx  = 0.3 + col * (boxW + 0.1);
      const by  = boxY + row * (boxH / 2 + 0.05);
      const bh  = boxH / 2 - 0.05;

      s.addShape(pres.shapes.RECTANGLE, {
        x:bx, y:by, w:boxW, h:bh,
        fill:{color:C.white}, line:{color:C.lgrey, pt:0.75}
      });
      // Box header bar
      s.addShape(pres.shapes.RECTANGLE, {
        x:bx, y:by, w:boxW, h:0.28,
        fill:{color:C.crimson}, line:{color:C.crimson}
      });
      s.addText(box.title||"", {
        x:bx+0.08, y:by, w:boxW-0.1, h:0.28,
        fontSize:10, bold:true, color:C.white, fontFace:FONT,
        valign:"middle", margin:0,
      });
      // Bullets inside box
      const buls = (box.bullets||[]).map(l => ({ text:l, options:{ bullet:{type:"bullet"} } }));
      s.addText(buls.length ? buls : [{text:"(no content)", options:{}}], {
        x:bx+0.06, y:by+0.31, w:boxW-0.1, h:bh-0.35,
        fontSize:9.5, color:C.textDark, fontFace:FONT,
        valign:"top", margin:[2,0,2,6],
      });
    });
  }

  // ════════════════════════════════════════════════════════
  // SLIDE 3 — 2. PROGRESS AGAINST STRATEGIC OBJECTIVES
  // Left: intro paragraph + expanded Capital Management detail
  // Right: 5 RAG pillar cards (Done / In Progress / Watch)
  // ════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    addHeader(s, pres, "2", "Progress Against FY26 Strategic Objectives");

    const pillars = D.strategic_pillars || [];
    const capPillar = pillars.find(p => p.id === "capital") || pillars[pillars.length-1] || {};
    const otherPillars = pillars.filter(p => p.id !== "capital").slice(0, 5);

    // Left column — intro + Capital Management detail
    const leftX = 0.3, leftW = 4.3, leftY = CONTENT_Y;

    // Two intro bullets (last period reporting, next meeting plan)
    const introBullets = [
      { text:"This is the last period where management will report against the FY26 strategic objectives agreed at the meeting in September 2025. Failure to stabilise the balance sheet during the period has been a significant miss by management against the plan.", options:{ bullet:{type:"bullet"} } },
      { text:"At the next meeting, management will propose a new operational and financial plan for board consideration and approval. Some of the objectives from this year will naturally continue into next year, but we anticipate the short-term focus will broadly be on the asset sale strategy and capital conservation.", options:{ bullet:{type:"bullet"} } },
    ];
    s.addText(introBullets, {
      x:leftX, y:leftY, w:leftW, h:1.35,
      fontSize:10, color:C.textDark, fontFace:FONT,
      valign:"top", margin:[3,0,3,8],
    });

    // Capital Management card (lower left, expanded)
    const capY = leftY + 1.4;
    const capH = FOOTER_Y - capY - 0.05;
    s.addShape(pres.shapes.RECTANGLE, {
      x:leftX, y:capY, w:leftW, h:capH,
      fill:{color:C.white}, line:{color:C.lgrey, pt:0.75}
    });
    const capRag   = capPillar.rag_status || "WATCH";
    const capColor = RAG_COLOR[capRag] || C.amber;
    const capLabel = RAG_LABEL[capRag] || capRag;

    // Card header
    s.addShape(pres.shapes.RECTANGLE, {
      x:leftX, y:capY, w:leftW, h:0.3,
      fill:{color:capColor}, line:{color:capColor}
    });
    // RAG icon
    s.addImage({ data: ragIco(capRag), x: leftX+0.06, y: capY+0.03, w:0.24, h:0.24 });
    s.addText(capPillar.title || "Capital Management", {
      x:leftX+0.36, y:capY, w:leftW-1.0, h:0.3,
      fontSize:10.5, bold:true, color:C.white, fontFace:FONT,
      valign:"middle", margin:0,
    });
    // WATCH badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x:leftX+leftW-0.75, y:capY+0.04, w:0.68, h:0.22,
      fill:{color:C.white, transparency:20}, line:{color:C.white, pt:0.75}, rectRadius:0.04,
    });
    s.addText(capLabel, {
      x:leftX+leftW-0.75, y:capY+0.04, w:0.68, h:0.22,
      fontSize:7, bold:true, color:C.white, fontFace:FONT,
      align:"center", valign:"middle", margin:0,
    });
    // Bullets inside Capital Management card
    const capBullets = (capPillar.extended_bullets || (capPillar.slide_briefing ? [capPillar.slide_briefing] : ["No detail available."])).map(l => ({text:l, options:{bullet:{type:"bullet"}}}));
    s.addText(capBullets, {
      x:leftX+0.06, y:capY+0.33, w:leftW-0.1, h:capH-0.37,
      fontSize:9.5, color:C.textDark, fontFace:FONT,
      valign:"top", margin:[2,0,2,6],
    });

    // Right column — 5 RAG pillar cards
    const rightX = leftX + leftW + 0.15;
    const rightW = W - rightX - 0.25;
    const cardH  = (FOOTER_Y - CONTENT_Y - 0.06) / Math.max(otherPillars.length, 1);

    otherPillars.forEach((p, i) => {
      const by  = CONTENT_Y + i * cardH;
      const rag = p.rag_status || "IN_PROGRESS";
      const rc  = RAG_COLOR[rag] || C.amber;
      const rl  = RAG_LABEL[rag] || rag;

      s.addShape(pres.shapes.RECTANGLE, {
        x:rightX, y:by, w:rightW, h:cardH-0.04,
        fill:{color:C.white}, line:{color:C.lgrey, pt:0.75}
      });

      // RAG icon + title
      s.addImage({ data: ragIco(rag), x: rightX+0.06, y: by+0.06, w:0.28, h:0.28 });
      s.addText(p.title||"", {
        x:rightX+0.4, y:by+0.06, w:rightW-1.3, h:0.28,
        fontSize:11, bold:true, color:C.textDark, fontFace:FONT, margin:0,
      });
      // RAG badge
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x:rightX+rightW-0.92, y:by+0.07, w:0.85, h:0.22,
        fill:{color:rc}, line:{color:rc}, rectRadius:0.04,
      });
      s.addText(rl, {
        x:rightX+rightW-0.92, y:by+0.07, w:0.85, h:0.22,
        fontSize:7, bold:true, color:C.white, fontFace:FONT,
        align:"center", valign:"middle", margin:0,
      });
      // Brief text from slide_briefing
      const brief = p.slide_briefing || "";
      s.addText(brief, {
        x:rightX+0.06, y:by+0.38, w:rightW-0.12, h:cardH-0.46,
        fontSize:8.5, color:C.textMid, fontFace:FONT, valign:"top", margin:0,
      });
      // Light divider between cards
      if (i < otherPillars.length-1) {
        s.addShape(pres.shapes.RECTANGLE, {
          x:rightX, y:by+cardH-0.06, w:rightW, h:0.02,
          fill:{color:C.lgrey}, line:{color:C.lgrey}
        });
      }
    });
  }

  // ════════════════════════════════════════════════════════
  // SLIDES 4+ — SECTION DEEP DIVES
  // Each section: numbered header + bullets left +
  //               optional right column with own header + bullets
  // ════════════════════════════════════════════════════════
  const sections = D.sections || [];
  for (const sec of sections) {
    const s = pres.addSlide();
    addHeader(s, pres, String(sec.number||""), sec.title||"");

    const hasRight = !!(sec.right_bullets && sec.right_bullets.length);
    const leftW    = hasRight ? 4.5 : W - 0.6;

    // Left bullets
    const leftBullets = (sec.bullets||[]).map(l => ({text:l, options:{bullet:{type:"bullet"}}}));
    s.addText(leftBullets.length ? leftBullets : [{text:"(no content)", options:{}}], {
      x:0.3, y:CONTENT_Y, w:leftW, h:CONTENT_H,
      fontSize:11, color:C.textDark, fontFace:FONT,
      valign:"top", margin:[3,0,3,8],
    });

    if (hasRight) {
      const rightX = 0.3 + leftW + 0.15;
      const rightW = W - rightX - 0.15;
      let curY = CONTENT_Y;

      // Optional right column title (underlined bold)
      if (sec.right_title) {
        s.addText(sec.right_title, {
          x:rightX, y:curY, w:rightW, h:0.3,
          fontSize:11, bold:true, underline:true,
          color:C.textDark, fontFace:FONT, margin:0,
        });
        curY += 0.32;
      }

      const rightBullets = (sec.right_bullets||[]).map(l => ({text:l, options:{bullet:{type:"bullet"}}}));
      const rightH = sec.right_sub_bullets ? CONTENT_H * 0.52 : CONTENT_H - (curY - CONTENT_Y);
      s.addText(rightBullets, {
        x:rightX, y:curY, w:rightW, h:rightH,
        fontSize:11, color:C.textDark, fontFace:FONT,
        valign:"top", margin:[3,0,3,8],
      });

      // Optional right sub-section
      if (sec.right_sub_title && sec.right_sub_bullets) {
        const subY = curY + rightH + 0.1;
        s.addText(sec.right_sub_title, {
          x:rightX, y:subY, w:rightW, h:0.3,
          fontSize:11, bold:true, underline:true,
          color:C.textDark, fontFace:FONT, margin:0,
        });
        const subBullets = sec.right_sub_bullets.map(l => ({text:l, options:{bullet:{type:"bullet"}}}));
        s.addText(subBullets, {
          x:rightX, y:subY+0.32, w:rightW, h:FOOTER_Y - subY - 0.38,
          fontSize:11, color:C.textDark, fontFace:FONT,
          valign:"top", margin:[3,0,3,8],
        });
      }
    }
  }

  // ── Write ────────────────────────────────────────────────────────────────────
  const out = outPath(D);
  await pres.writeFile({ fileName: out });
  const nSlides = 3 + sections.length;
  console.log("\n" + "=".repeat(56));
  console.log("  Board deck saved: " + path.basename(out));
  console.log("  Slides: " + nSlides + " (title + exec summary + strategic objectives + " + sections.length + " sections)");
  console.log("  Meeting: " + meetingDate);
  console.log("=".repeat(56) + "\n");
}

build().catch(e => { console.error("Error:", e.message, e.stack); process.exit(1); });
