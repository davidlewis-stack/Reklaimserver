// =============================================================================
// REKLAIM DECK TEMPLATE — reusable engine
// =============================================================================
// Usage: node reklaim_deck_template.js path/to/deck_config.json
//
// All content (brand, vertical, audiences, stats, contact) lives in the JSON
// config. This file should not need editing for new decks.
// =============================================================================

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

// -----------------------------------------------------------------------------
// Load config
// -----------------------------------------------------------------------------
const configPath = process.argv[2];
if (!configPath) {
  console.error("Usage: node reklaim_deck_template.js <config.json>");
  process.exit(1);
}
const cfg = JSON.parse(fs.readFileSync(configPath, "utf8"));

// Resolve icon names (strings in JSON) to react-icons components
function iconFor(name, fallback = "FaCheckCircle") {
  return FA[name] || FA[fallback];
}

// -----------------------------------------------------------------------------
// Constants
// -----------------------------------------------------------------------------
const W = 13.333, H = 7.5;
const C = {
  black:    "0A0A0A",
  charcoal: "1F1F1F",
  midGray:  "4A4A4A",
  softGray: "8A8A8A",
  lineGray: "E5E5E5",
  bgLight:  "FAFAFA",
  white:    "FFFFFF",
  yellow:   "F5C518",
  yellowD:  "C9A013",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Reklaim";
pres.title = cfg.deckTitle || `Reklaim — ${cfg.brand}`;

// -----------------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------------
async function iconPng(IconComponent, color = C.black, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const softShadow = () => ({ type: "outer", blur: 12, offset: 2, angle: 90, color: "000000", opacity: 0.08 });

function addChrome(slide, pageLabel, sectionLabel) {
  slide.addText(sectionLabel.toUpperCase(), {
    x: 0.6, y: 0.4, w: 8, h: 0.3,
    fontSize: 9, fontFace: "Arial", bold: true,
    color: C.softGray, charSpacing: 4, margin: 0,
  });
  slide.addText("REKLAIM", {
    x: W - 2.2, y: 0.4, w: 1.6, h: 0.3,
    fontSize: 11, fontFace: "Arial Black", bold: true,
    color: C.black, charSpacing: 4, align: "right", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: W - 0.55, y: 0.46, w: 0.18, h: 0.18,
    fill: { color: C.yellow }, line: { type: "none" },
  });
  slide.addText(`Reklaim  ·  Confidential  ·  Prepared for ${cfg.brand}`, {
    x: 0.6, y: H - 0.45, w: 8, h: 0.3,
    fontSize: 9, fontFace: "Arial",
    color: C.softGray, margin: 0,
  });
  slide.addText(pageLabel, {
    x: W - 1.6, y: H - 0.45, w: 1.0, h: 0.3,
    fontSize: 9, fontFace: "Arial",
    color: C.softGray, align: "right", margin: 0,
  });
}

function titleBlock(slide, eyebrow, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0.95, w: 0.08, h: 0.95,
    fill: { color: C.yellow }, line: { type: "none" },
  });
  slide.addText(eyebrow.toUpperCase(), {
    x: 0.85, y: 0.95, w: 8, h: 0.3,
    fontSize: 10, fontFace: "Arial", bold: true,
    color: C.yellowD, charSpacing: 4, margin: 0,
  });
  slide.addText(title, {
    x: 0.85, y: 1.25, w: 11.5, h: 0.75,
    fontSize: 32, fontFace: "Calibri", bold: true,
    color: C.black, margin: 0,
  });
}

function drawGrille(slide) {
  const slatX0 = 8.6, slatGap = 0.55, slatW = 0.32;
  for (let i = 0; i < 8; i++) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: slatX0 + i * slatGap, y: 0, w: slatW, h: H,
      fill: { color: C.charcoal }, line: { type: "none" },
    });
  }
  slide.addShape(pres.shapes.RECTANGLE, {
    x: slatX0 + 3 * slatGap, y: 0, w: slatW, h: H,
    fill: { color: C.yellow }, line: { type: "none" },
  });
}

// -----------------------------------------------------------------------------
// SLIDE 1 — Cover
// -----------------------------------------------------------------------------
async function slide01() {
  const s = pres.addSlide();
  s.background = { color: C.black };
  drawGrille(s);

  s.addText("REKLAIM", {
    x: 0.7, y: 0.55, w: 4, h: 0.5,
    fontSize: 22, fontFace: "Arial Black", bold: true,
    color: C.yellow, charSpacing: 6, margin: 0,
  });
  s.addText(`CLIENT PRESENTATION  ·  ${cfg.year}`, {
    x: 0.7, y: 2.4, w: 8, h: 0.3,
    fontSize: 11, fontFace: "Arial", bold: true,
    color: C.yellow, charSpacing: 6, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.85, w: 0.7, h: 0.06,
    fill: { color: C.yellow }, line: { type: "none" },
  });
  s.addText([
    { text: cfg.coverHeadline.line1, options: { breakLine: true, color: C.white } },
    { text: cfg.coverHeadline.line2, options: { color: C.yellow } },
  ], {
    x: 0.7, y: 3.05, w: 8.0, h: 1.9,
    fontSize: 36, fontFace: "Calibri", bold: true, margin: 0, paraSpaceAfter: 0,
  });
  s.addText(cfg.coverSubtitle, {
    x: 0.7, y: 5.05, w: 7.6, h: 1.0,
    fontSize: 17, fontFace: "Calibri", color: "CCCCCC", margin: 0,
  });

  // Footer meta
  const meta = [
    ["PREPARED FOR", cfg.brand, 0.7],
    ["PRESENTED BY", cfg.presenter || "Reklaim", 3.5],
    ["DATE", String(cfg.year), 6.3],
  ];
  for (const [label, value, x] of meta) {
    s.addText(label, {
      x, y: 6.35, w: 2.5, h: 0.25,
      fontSize: 8, fontFace: "Arial", bold: true,
      color: C.softGray, charSpacing: 3, margin: 0,
    });
    s.addText(value, {
      x, y: 6.6, w: 2.5, h: 0.4,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.white, margin: 0,
    });
  }
  s.addText("CONFIDENTIAL  /  01", {
    x: 0.7, y: H - 0.4, w: 3.0, h: 0.28,
    fontSize: 9, fontFace: "Arial", bold: true,
    color: C.softGray, charSpacing: 4, align: "left", margin: 0,
  });
}

// -----------------------------------------------------------------------------
// SLIDE 2 — User Journey
// -----------------------------------------------------------------------------
async function slide02() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "02 / 10", "02 / User Journey");
  titleBlock(s, "How it works", "How Reklaim captures explicit consent.");

  s.addText("Five lightweight steps. Every audience attached to a verified, opted-in identity.", {
    x: 0.85, y: 2.05, w: 11.5, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.midGray, margin: 0,
  });

  // Phone mock
  const px = 0.85, py = 2.7, pw = 3.1, ph = 4.1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px, y: py, w: pw, h: ph,
    fill: { color: C.black }, line: { color: C.charcoal, width: 1 },
    rectRadius: 0.25, shadow: softShadow(),
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.12, y: py + 0.18, w: pw - 0.24, h: ph - 0.36,
    fill: { color: "111111" }, line: { type: "none" }, rectRadius: 0.18,
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px + pw / 2 - 0.45, y: py + 0.22, w: 0.9, h: 0.18,
    fill: { color: C.black }, line: { type: "none" }, rectRadius: 0.09,
  });
  s.addText("IT'S YOUR DATA.", {
    x: px + 0.2, y: py + 1.6, w: pw - 0.4, h: 0.45,
    fontSize: 20, fontFace: "Arial Black", bold: true,
    color: C.white, align: "center", margin: 0,
  });
  s.addText("Take back control of it.", {
    x: px + 0.2, y: py + 2.05, w: pw - 0.4, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: "BBBBBB", align: "center", margin: 0,
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.35, y: py + 2.6, w: pw - 0.7, h: 0.42,
    fill: { color: C.yellow }, line: { type: "none" }, rectRadius: 0.06,
  });
  s.addText("Sign up now", {
    x: px + 0.35, y: py + 2.6, w: pw - 0.7, h: 0.42,
    fontSize: 12, fontFace: "Arial", bold: true,
    color: C.black, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.35, y: py + 3.1, w: pw - 0.7, h: 0.42,
    fill: { color: "111111" }, line: { color: C.yellow, width: 1 }, rectRadius: 0.06,
  });
  s.addText("Login to my account", {
    x: px + 0.35, y: py + 3.1, w: pw - 0.7, h: 0.42,
    fontSize: 11, fontFace: "Arial", bold: true,
    color: C.yellow, align: "center", valign: "middle", margin: 0,
  });

  // 5 numbered steps (config-driven)
  const colX = 4.4, colW = 8.3, gap = 0.05;
  const journeyN = cfg.userJourney.length;
  const journeyAvailH = 4.1;
  const rowH = journeyAvailH / journeyN;
  let y = 2.7;
  cfg.userJourney.forEach((step, i) => {
    s.addShape(pres.shapes.OVAL, {
      x: colX, y: y + 0.08, w: 0.6, h: 0.6,
      fill: { color: C.yellow }, line: { type: "none" },
    });
    s.addText(String(i + 1), {
      x: colX, y: y + 0.08, w: 0.6, h: 0.6,
      fontSize: 18, fontFace: "Arial Black", bold: true,
      color: C.black, align: "center", valign: "middle", margin: 0,
    });
    s.addText(step.title, {
      x: colX + 0.85, y: y + 0.05, w: 4, h: 0.35,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.black, margin: 0,
    });
    s.addText(step.body, {
      x: colX + 0.85, y: y + 0.4, w: 6.0, h: 0.35,
      fontSize: 12, fontFace: "Calibri",
      color: C.midGray, margin: 0,
    });
    if (i < cfg.userJourney.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: colX, y: y + rowH - gap, w: colW - 0.4, h: 0,
        line: { color: C.lineGray, width: 0.75 },
      });
    }
    y += rowH;
  });
}

// -----------------------------------------------------------------------------
// SLIDE 3 — Data Catalogue
// -----------------------------------------------------------------------------
async function slide03() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "03 / 10", "03 / Data Catalogue");
  titleBlock(s, "Breadth", cfg.catalogue.title);

  s.addText(cfg.catalogue.intro, {
    x: 0.85, y: 2.05, w: 11.5, h: 0.6,
    fontSize: 13, fontFace: "Calibri", color: C.midGray, margin: 0,
  });
  s.addText("These categories represent a selection from Reklaim's 500+ opt-in SPI audiences.", {
    x: 0.85, y: 2.65, w: 11.5, h: 0.25,
    fontSize: 10, fontFace: "Calibri", italic: true,
    color: C.softGray, margin: 0,
  });

  const colY = 3.05;
  s.addText(cfg.catalogue.leftHeader, {
    x: 0.85, y: colY, w: 5.6, h: 0.3,
    fontSize: 11, fontFace: "Arial", bold: true,
    color: C.yellowD, charSpacing: 3, margin: 0,
  });
  s.addText(cfg.catalogue.rightHeader, {
    x: 6.85, y: colY, w: 5.6, h: 0.3,
    fontSize: 11, fontFace: "Arial", bold: true,
    color: C.yellowD, charSpacing: 3, margin: 0,
  });

  const cardH = 0.58, cardGap = 0.06, cardY0 = 3.45;
  const drawCol = (items, x0, w) => {
    items.forEach((it, i) => {
      const y = cardY0 + i * (cardH + cardGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x: x0, y, w, h: cardH,
        fill: { color: C.white }, line: { color: C.lineGray, width: 0.75 },
        shadow: softShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: x0, y, w: 0.05, h: cardH,
        fill: { color: C.yellow }, line: { type: "none" },
      });
      s.addText(it.name, {
        x: x0 + 0.2, y: y + 0.05, w: w - 0.3, h: 0.26,
        fontSize: 12, fontFace: "Calibri", bold: true,
        color: C.black, margin: 0,
      });
      s.addText(it.desc, {
        x: x0 + 0.2, y: y + 0.3, w: w - 0.3, h: 0.25,
        fontSize: 10.5, fontFace: "Calibri",
        color: C.midGray, margin: 0,
      });
    });
  };
  drawCol(cfg.catalogue.left,  0.85, 5.6);
  drawCol(cfg.catalogue.right, 6.85, 5.6);
}

// -----------------------------------------------------------------------------
// SLIDE 4 — Verticals (3 cards, first one accented)
// -----------------------------------------------------------------------------
async function slide04() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "04 / 10", "04 / Verticals");
  titleBlock(s, "Focus areas", cfg.verticals.title);

  s.addText(cfg.verticals.intro, {
    x: 0.85, y: 2.05, w: 11.5, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.midGray, margin: 0,
  });

  const cards = cfg.verticals.cards;
  const n = cards.length;
  const cgap = 0.2, cy = 2.9, ch = 4.0;
  const cw = (W - 1.7 - (n - 1) * cgap) / n;
  const startX = (W - (cw * n + cgap * (n - 1))) / 2;

  for (let i = 0; i < cards.length; i++) {
    const c = cards[i];
    const x = startX + i * (cw + cgap);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cy, w: cw, h: ch,
      fill: { color: c.accent ? C.black : C.white },
      line: { color: c.accent ? C.black : C.lineGray, width: 0.75 },
      shadow: softShadow(),
    });
    if (c.accent) {
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: cy, w: cw, h: 0.08,
        fill: { color: C.yellow }, line: { type: "none" },
      });
    }
    const iconColor = c.accent ? C.yellow : C.yellowD;
    const iconData = await iconPng(iconFor(c.icon), iconColor);
    s.addImage({ data: iconData, x: x + 0.35, y: cy + 0.4, w: 0.55, h: 0.55 });
    s.addText(c.tag, {
      x: x + 0.35, y: cy + 1.1, w: cw - 0.7, h: 0.25,
      fontSize: 9, fontFace: "Arial", bold: true,
      color: c.accent ? C.yellow : C.softGray, charSpacing: 3, margin: 0,
    });
    s.addText(c.title, {
      x: x + 0.35, y: cy + 1.35, w: cw - 0.7, h: 0.5,
      fontSize: 22, fontFace: "Calibri", bold: true,
      color: c.accent ? C.white : C.black, margin: 0,
    });
    s.addText(c.stat, {
      x: x + 0.35, y: cy + 1.95, w: cw - 0.7, h: 0.7,
      fontSize: 44, fontFace: "Calibri", bold: true,
      color: c.accent ? C.yellow : C.black, margin: 0,
    });
    s.addText(c.statLabel, {
      x: x + 0.35, y: cy + 2.65, w: cw - 0.7, h: 0.6,
      fontSize: 10, fontFace: "Calibri",
      color: c.accent ? "AAAAAA" : C.softGray, margin: 0,
    });
    s.addText(c.body, {
      x: x + 0.35, y: cy + 3.25, w: cw - 0.7, h: 0.7,
      fontSize: 11, fontFace: "Calibri",
      color: c.accent ? "DDDDDD" : C.midGray, margin: 0,
    });
  }
}

// -----------------------------------------------------------------------------
// SLIDE 5 — Tier 1 Audiences (2x2 grid)
// -----------------------------------------------------------------------------
async function slide05() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "05 / 10", "05 / Tier 1 Opportunity");
  titleBlock(s, "Tier 1 opportunity", cfg.tier1.title);

  s.addText(cfg.tier1.intro, {
    x: 0.85, y: 2.05, w: 11.5, h: 0.6,
    fontSize: 13, fontFace: "Calibri", color: C.midGray, margin: 0,
  });

  const aud = cfg.tier1.audiences;
  const n5 = aud.length;
  const cols5 = n5 <= 2 ? n5 : 2;
  const rows5 = Math.ceil(n5 / cols5);
  const gx0 = 0.85, gy0 = 2.85, gxg = 0.2, gyg = 0.2;
  const availW5 = W - 1.7;
  const availH5 = H - gy0 - 0.6;
  const gw = (availW5 - (cols5 - 1) * gxg) / cols5;
  const gh = (availH5 - (rows5 - 1) * gyg) / rows5;
  for (let i = 0; i < n5; i++) {
    const a = aud[i];
    const col = i % cols5, row = Math.floor(i / cols5);
    const x = gx0 + col * (gw + gxg);
    const y = gy0 + row * (gh + gyg);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: gw, h: gh,
      fill: { color: C.white }, line: { color: C.lineGray, width: 0.75 },
      shadow: softShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: gh,
      fill: { color: C.yellow }, line: { type: "none" },
    });
    s.addText(`AUDIENCE 0${i + 1}`, {
      x: x + 0.3, y: y + 0.18, w: 2, h: 0.25,
      fontSize: 9, fontFace: "Arial", bold: true,
      color: C.softGray, charSpacing: 3, margin: 0,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + gw - 1.4, y: y + 0.15, w: 1.2, h: 0.5,
      fill: { color: C.black }, line: { type: "none" }, rectRadius: 0.06,
    });
    s.addText([
      { text: "REACH  ", options: { fontSize: 8, color: C.softGray, charSpacing: 2 } },
      { text: a.reach, options: { fontSize: 14, color: C.yellow, bold: true } },
    ], {
      x: x + gw - 1.4, y: y + 0.15, w: 1.2, h: 0.5,
      fontFace: "Arial", align: "center", valign: "middle", margin: 0,
    });
    s.addText(a.name, {
      x: x + 0.3, y: y + 0.5, w: gw - 1.6, h: 0.45,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.black, margin: 0,
    });
    s.addText(a.desc, {
      x: x + 0.3, y: y + 1.0, w: gw - 0.55, h: 0.6,
      fontSize: 11, fontFace: "Calibri",
      color: C.midGray, margin: 0,
    });
    s.addText(a.signals, {
      x: x + 0.3, y: y + gh - 0.4, w: gw - 0.55, h: 0.3,
      fontSize: 9, fontFace: "Arial", bold: true,
      color: C.yellowD, charSpacing: 2, margin: 0,
    });
  }
}

// -----------------------------------------------------------------------------
// SLIDE 6 — Tier 2 Opportunity (3 stats + 3 WHO/WHY/HOW blocks)
// -----------------------------------------------------------------------------
async function slide06() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "06 / 10", "06 / Tier 2 Opportunity");
  titleBlock(s, "Tier 2 opportunity", cfg.tier2.title);

  // Stats
  const stats = cfg.tier2.stats;
  const ns = stats.length;
  const stx0 = 0.85, sty = 2.15, sth = 1.5, stg = 0.25;
  const stw = (W - 1.7 - (ns - 1) * stg) / ns;
  for (let i = 0; i < ns; i++) {
    const x = stx0 + i * (stw + stg);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: sty, w: stw, h: sth,
      fill: { color: C.black }, line: { type: "none" }, shadow: softShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: sty, w: stw, h: 0.06,
      fill: { color: C.yellow }, line: { type: "none" },
    });
    s.addText(stats[i].num, {
      x: x + 0.3, y: sty + 0.2, w: stw - 0.6, h: 0.7,
      fontSize: 44, fontFace: "Calibri", bold: true,
      color: C.yellow, margin: 0,
    });
    s.addText(stats[i].label, {
      x: x + 0.3, y: sty + 0.95, w: stw - 0.6, h: 0.5,
      fontSize: 11, fontFace: "Calibri", color: "DDDDDD", margin: 0,
    });
  }

  // WHO/WHY/HOW
  const blocks = cfg.tier2.blocks;
  const nb = blocks.length;
  const bx0 = 0.85, by0 = 3.95, bh = 2.7, bg = 0.25;
  const bw = (W - 1.7 - (nb - 1) * bg) / nb;
  for (let i = 0; i < nb; i++) {
    const b = blocks[i];
    const x = bx0 + i * (bw + bg);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: by0, w: bw, h: bh,
      fill: { color: C.white }, line: { color: C.lineGray, width: 0.75 },
      shadow: softShadow(),
    });
    const iconData = await iconPng(iconFor(b.icon), C.yellowD);
    s.addImage({ data: iconData, x: x + 0.3, y: by0 + 0.3, w: 0.45, h: 0.45 });
    s.addText(b.tag, {
      x: x + 0.3, y: by0 + 0.85, w: bw - 0.6, h: 0.25,
      fontSize: 10, fontFace: "Arial", bold: true,
      color: C.yellowD, charSpacing: 4, margin: 0,
    });
    s.addText(b.title, {
      x: x + 0.3, y: by0 + 1.1, w: bw - 0.6, h: 0.5,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.black, margin: 0,
    });
    s.addText(b.body, {
      x: x + 0.3, y: by0 + 1.6, w: bw - 0.6, h: 1.0,
      fontSize: 11, fontFace: "Calibri", color: C.midGray, margin: 0,
    });
  }
}

// -----------------------------------------------------------------------------
// SLIDE 7 — Advantages (2x2)
// -----------------------------------------------------------------------------
async function slide07() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "07 / 10", "07 / Advantage");
  titleBlock(s, "Why Reklaim", cfg.advantages.title);

  s.addText(cfg.advantages.intro, {
    x: 0.85, y: 2.05, w: 11.5, h: 0.45,
    fontSize: 14, fontFace: "Calibri", italic: true,
    color: C.midGray, margin: 0,
  });

  const advs = cfg.advantages.items;
  const na = advs.length;
  const colsA = na <= 2 ? na : 2;
  const rowsA = Math.ceil(na / colsA);
  const ax0 = 0.85, ay0 = 2.7, agx = 0.25, agy = 0.25;
  const availWA = W - 1.7;
  const availHA = H - ay0 - 0.6;
  const aw = (availWA - (colsA - 1) * agx) / colsA;
  const ah = (availHA - (rowsA - 1) * agy) / rowsA;
  for (let i = 0; i < na; i++) {
    const a = advs[i];
    const col = i % colsA, row = Math.floor(i / colsA);
    const x = ax0 + col * (aw + agx);
    const y = ay0 + row * (ah + agy);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: aw, h: ah,
      fill: { color: C.white }, line: { color: C.lineGray, width: 0.75 },
      shadow: softShadow(),
    });
    s.addText(`0${i + 1}`, {
      x: x + 0.25, y: y + 0.25, w: 1.3, h: 1.4,
      fontSize: 56, fontFace: "Calibri", bold: true,
      color: C.yellow, margin: 0,
    });
    const iconData = await iconPng(iconFor(a.icon), C.black);
    s.addImage({ data: iconData, x: x + aw - 0.65, y: y + 0.25, w: 0.42, h: 0.42 });
    s.addText(a.title, {
      x: x + 1.55, y: y + 0.3, w: aw - 2.3, h: 0.5,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.black, margin: 0,
    });
    s.addText(a.body, {
      x: x + 1.55, y: y + 0.85, w: aw - 1.85, h: 1.0,
      fontSize: 11, fontFace: "Calibri", color: C.midGray, margin: 0,
    });
  }
}

// -----------------------------------------------------------------------------
// SLIDE 8 — CPL Solution (4 horizontal cards with chevrons)
// -----------------------------------------------------------------------------
async function slide08() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "08 / 10", "08 / Solution");
  titleBlock(s, cfg.solution.eyebrow, cfg.solution.title);

  s.addText(cfg.solution.intro, {
    x: 0.85, y: 2.05, w: 11.5, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.midGray, margin: 0,
  });

  const steps = cfg.solution.steps;
  const nSt = steps.length;
  const sx0 = 0.85, sy = 2.85, arrowW = 0.35;
  const totalW = W - 1.7;
  const cardWidth = (totalW - (nSt - 1) * arrowW) / nSt;
  const cardH = 3.5;

  for (let i = 0; i < nSt; i++) {
    const st = steps[i];
    const x = sx0 + i * (cardWidth + arrowW);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: sy, w: cardWidth, h: cardH,
      fill: { color: C.white }, line: { color: C.lineGray, width: 0.75 },
      shadow: softShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: sy, w: cardWidth, h: 0.08,
      fill: { color: C.yellow }, line: { type: "none" },
    });
    s.addText(`0${i + 1}`, {
      x: x + 0.3, y: sy + 0.3, w: cardWidth - 0.6, h: 0.5,
      fontSize: 12, fontFace: "Arial", bold: true,
      color: C.yellowD, charSpacing: 4, margin: 0,
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + cardWidth/2 - 0.5, y: sy + 0.85, w: 1.0, h: 1.0,
      fill: { color: C.black }, line: { type: "none" },
    });
    const iconData = await iconPng(iconFor(st.icon), C.yellow);
    s.addImage({ data: iconData, x: x + cardWidth/2 - 0.3, y: sy + 1.05, w: 0.6, h: 0.6 });
    s.addText(st.title, {
      x: x + 0.2, y: sy + 2.0, w: cardWidth - 0.4, h: 0.4,
      fontSize: 18, fontFace: "Calibri", bold: true,
      color: C.black, align: "center", margin: 0,
    });
    s.addText(st.body, {
      x: x + 0.25, y: sy + 2.45, w: cardWidth - 0.5, h: 1.0,
      fontSize: 11, fontFace: "Calibri",
      color: C.midGray, align: "center", margin: 0,
    });
    if (i < steps.length - 1) {
      const ax = x + cardWidth + 0.05;
      const ay = sy + cardH / 2;
      s.addText(">", {
        x: ax, y: ay - 0.3, w: arrowW - 0.05, h: 0.6,
        fontSize: 36, fontFace: "Arial", bold: true,
        color: C.yellow, align: "center", valign: "middle", margin: 0,
      });
    }
  }

  if (cfg.solution.footnote) {
    s.addText(cfg.solution.footnote, {
      x: 0.85, y: H - 0.85, w: 11.5, h: 0.3,
      fontSize: 11, fontFace: "Calibri", italic: true,
      color: C.softGray, margin: 0,
    });
  }
}

// -----------------------------------------------------------------------------
// SLIDE 9 — Activation (3x2 grid)
// -----------------------------------------------------------------------------
async function slide09() {
  const s = pres.addSlide();
  s.background = { color: C.bgLight };
  addChrome(s, "09 / 10", "09 / Activation");
  titleBlock(s, "Flexible activation", cfg.activation.title);

  s.addText(cfg.activation.intro, {
    x: 0.85, y: 2.05, w: 11.5, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.midGray, margin: 0,
  });

  const channels = cfg.activation.channels;
  const nCh = channels.length;
  const cgcols = nCh <= 3 ? nCh : (nCh <= 6 ? 3 : 4);
  const cgrows = Math.ceil(nCh / cgcols);
  const cgx0 = 0.85, cgy0 = 2.85, cggx = 0.2, cggy = 0.2;
  const availWC = W - 1.7;
  const availHC = H - cgy0 - 0.6;
  const cgw = (availWC - (cgcols - 1) * cggx) / cgcols;
  const cgh = (availHC - (cgrows - 1) * cggy) / cgrows;

  for (let i = 0; i < nCh; i++) {
    const c = channels[i];
    const col = i % cgcols, row = Math.floor(i / cgcols);
    const x = cgx0 + col * (cgw + cggx);
    const y = cgy0 + row * (cgh + cggy);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cgw, h: cgh,
      fill: { color: C.white }, line: { color: C.lineGray, width: 0.75 },
      shadow: softShadow(),
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7,
      fill: { color: C.yellow }, line: { type: "none" },
    });
    const iconData = await iconPng(iconFor(c.icon), C.black);
    s.addImage({ data: iconData, x: x + 0.43, y: y + 0.43, w: 0.44, h: 0.44 });
    s.addText(`0${i + 1}`, {
      x: x + cgw - 0.7, y: y + 0.3, w: 0.5, h: 0.3,
      fontSize: 11, fontFace: "Arial", bold: true,
      color: C.softGray, charSpacing: 3, align: "right", margin: 0,
    });
    s.addText(c.title, {
      x: x + 0.3, y: y + 1.05, w: cgw - 0.6, h: 0.35,
      fontSize: 14, fontFace: "Calibri", bold: true,
      color: C.black, margin: 0,
    });
    s.addText(c.body, {
      x: x + 0.3, y: y + 1.4, w: cgw - 0.6, h: 0.5,
      fontSize: 10, fontFace: "Calibri", color: C.midGray, margin: 0,
    });
  }
}

// -----------------------------------------------------------------------------
// SLIDE 10 — Thank You
// -----------------------------------------------------------------------------
async function slide10() {
  const s = pres.addSlide();
  s.background = { color: C.black };
  drawGrille(s);

  s.addText("REKLAIM", {
    x: 0.7, y: 0.55, w: 4, h: 0.5,
    fontSize: 22, fontFace: "Arial Black", bold: true,
    color: C.yellow, charSpacing: 6, margin: 0,
  });
  s.addText(cfg.closing.eyebrow, {
    x: 0.7, y: 2.5, w: 8, h: 0.3,
    fontSize: 11, fontFace: "Arial", bold: true,
    color: C.yellow, charSpacing: 6, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.95, w: 0.7, h: 0.06,
    fill: { color: C.yellow }, line: { type: "none" },
  });
  s.addText(cfg.closing.headline, {
    x: 0.7, y: 3.15, w: 8, h: 1.3,
    fontSize: 48, fontFace: "Calibri", bold: true,
    color: C.white, margin: 0,
  });
  s.addText(cfg.closing.subline, {
    x: 0.7, y: 4.55, w: 7.6, h: 1.0,
    fontSize: 16, fontFace: "Calibri", color: "CCCCCC", margin: 0,
  });

  const contact = [
    ["CONTACT", cfg.contact.name, 0.7, C.white],
    ["EMAIL", cfg.contact.email, 4.0, C.yellow],
    ["WEB", cfg.contact.web, 6.5, C.white],
  ];
  for (const [label, value, x, color] of contact) {
    s.addText(label, {
      x, y: 5.85, w: 2.5, h: 0.25,
      fontSize: 8, fontFace: "Arial", bold: true,
      color: C.softGray, charSpacing: 3, margin: 0,
    });
    s.addText(value, {
      x, y: 6.1, w: 3.5, h: 0.4,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color, margin: 0,
    });
  }
  s.addText("CONFIDENTIAL  /  10", {
    x: 0.7, y: H - 0.4, w: 3.0, h: 0.28,
    fontSize: 9, fontFace: "Arial", bold: true,
    color: C.softGray, charSpacing: 4, align: "left", margin: 0,
  });
}

// -----------------------------------------------------------------------------
// BUILD
// -----------------------------------------------------------------------------
(async () => {
  await slide01();
  await slide02();
  await slide03();
  await slide04();
  await slide05();
  await slide06();
  await slide07();
  await slide08();
  await slide09();
  await slide10();

  const outFile = cfg.outputPath || `Reklaim_${cfg.brand.replace(/\s+/g, "_")}_${cfg.shortVertical || "Deck"}.pptx`;
  await pres.writeFile({ fileName: outFile });
  console.log("Wrote:", outFile);
})();
