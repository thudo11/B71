/**
 * generate.js
 * Generates Startup_Profile_Template_1920x1080_CanvaReady.pptx
 * using PptxGenJS (Node.js).
 *
 * Design: Image-1 inspired – white background, left mesh placeholder (~42% width),
 * orange badge top-right, navy/orange accent lines.
 *
 * Usage:
 *   node generate.js
 */

"use strict";

const PptxGenJS = require("pptxgenjs");
const path = require("path");

// ─── Theme Tokens ────────────────────────────────────────────────────────────
const COLOR = {
  orange: "F57C00",
  navy: "0B2E5E",
  dark: "1E1E1E",
  lightGray: "EEF1F5",
  white: "FFFFFF",
  accent: "CCCCCC",
};

// Slide dimensions in inches (PptxGenJS default unit)
// 1920×1080 px ≡ 20 in × 11.25 in at 96 dpi
const W = 20;   // inches wide
const H = 11.25; // inches tall

const MESH_W = W * 0.42;  // ~8.4 in  (42% left zone)
const BADGE_SIZE = 1.5;   // orange badge (square)
const FONT_TITLE = "Montserrat";
const FONT_BODY  = "Arial";

// ─── Helper: add shared chrome (mesh placeholder + badge + footer) ───────────
function addChrome(slide, { hideFooter = false } = {}) {
  // Left mesh placeholder
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: MESH_W, h: H,
    fill: { color: COLOR.lightGray },
    line: { color: COLOR.accent, width: 0.5 },
  });
  slide.addText("MESH IMAGE", {
    x: 0, y: H / 2 - 0.4, w: MESH_W, h: 0.8,
    align: "center",
    color: COLOR.accent,
    fontSize: 18,
    fontFace: FONT_BODY,
    italic: true,
  });

  // Thin vertical accent line between mesh and content area
  slide.addShape(pptx.ShapeType.line, {
    x: MESH_W, y: 0, w: 0, h: H,
    line: { color: COLOR.navy, width: 2 },
  });

  // Orange badge top-right
  const bx = W - BADGE_SIZE - 0.25;
  const by = 0.25;
  slide.addShape(pptx.ShapeType.rect, {
    x: bx, y: by, w: BADGE_SIZE, h: BADGE_SIZE,
    fill: { color: COLOR.orange },
    line: { color: COLOR.orange, width: 0 },
  });
  slide.addText("DEMO\n71", {
    x: bx, y: by, w: BADGE_SIZE, h: BADGE_SIZE,
    align: "center",
    valign: "middle",
    color: COLOR.white,
    fontSize: 22,
    fontFace: FONT_TITLE,
    bold: true,
  });

  // Bottom-right logo placeholder
  if (!hideFooter) {
    slide.addShape(pptx.ShapeType.rect, {
      x: W - 2.2, y: H - 0.65, w: 1.9, h: 0.45,
      fill: { color: COLOR.lightGray },
      line: { color: COLOR.accent, width: 0.5 },
    });
    slide.addText("LOGO", {
      x: W - 2.2, y: H - 0.65, w: 1.9, h: 0.45,
      align: "center",
      valign: "middle",
      color: COLOR.accent,
      fontSize: 10,
      fontFace: FONT_BODY,
    });

    // Thin navy footer rule
    slide.addShape(pptx.ShapeType.line, {
      x: MESH_W + 0.2, y: H - 0.7, w: W - MESH_W - 0.4, h: 0,
      line: { color: COLOR.navy, width: 1 },
    });
  }
}

// ─── Content-area helpers ─────────────────────────────────────────────────────
const CX = MESH_W + 0.4;  // content area left edge
const CW = W - MESH_W - 0.6;  // content area width

function sectionTitle(slide, text, y, x = CX, w = CW - 0.5) {
  // Orange underline accent
  slide.addShape(pptx.ShapeType.rect, {
    x: x, y: y + 0.35, w: 0.35, h: 0.05,
    fill: { color: COLOR.orange },
    line: { color: COLOR.orange, width: 0 },
  });
  slide.addText(text, {
    x: x + 0.45, y: y, w: w, h: 0.45,
    color: COLOR.navy,
    fontSize: 13,
    fontFace: FONT_TITLE,
    bold: true,
  });
}

// ─── PptxGenJS instance ───────────────────────────────────────────────────────
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE"; // 10×7.5 in by default; we'll override
pptx.defineLayout({ name: "HD1920x1080", width: W, height: H });
pptx.layout = "HD1920x1080";

pptx.author = "Block71";
pptx.title  = "Startup Profile Template – Block71";

// ─── Slide 1: Cover ───────────────────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addChrome(slide);

  const rx = CX;
  const rw = CW - BADGE_SIZE - 0.5;

  // Main title
  slide.addText("STARTUP PROFILE\nPRESENTATION", {
    x: rx, y: 1.8, w: rw, h: 2.2,
    color: COLOR.dark,
    fontSize: 44,
    fontFace: FONT_TITLE,
    bold: true,
    breakLine: true,
    autoFit: true,
  });

  // Orange underline
  slide.addShape(pptx.ShapeType.rect, {
    x: rx, y: 4.1, w: 2.5, h: 0.1,
    fill: { color: COLOR.orange },
    line: { color: COLOR.orange, width: 0 },
  });

  // Subtitle
  slide.addText("BLOCK71 SINGAPORE", {
    x: rx, y: 4.3, w: rw, h: 0.6,
    color: COLOR.orange,
    fontSize: 22,
    fontFace: FONT_TITLE,
    bold: true,
  });

  // Meta row
  const metaY = 5.2;
  const metaFontSize = 13;
  const metaItems = [
    { icon: "📍", label: "[Location]" },
    { icon: "📅", label: "[Date]" },
    { icon: "🕐", label: "[Time]" },
  ];
  metaItems.forEach((item, i) => {
    slide.addText(`${item.icon}  ${item.label}`, {
      x: rx + i * 3.2, y: metaY, w: 3.0, h: 0.45,
      color: COLOR.dark,
      fontSize: metaFontSize,
      fontFace: FONT_BODY,
    });
  });
}

// ─── Slide 2: Startup Index ────────────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addChrome(slide);

  slide.addText("STARTUPS", {
    x: CX, y: 0.5, w: CW, h: 0.9,
    color: COLOR.navy,
    fontSize: 36,
    fontFace: FONT_TITLE,
    bold: true,
  });
  // Orange underline
  slide.addShape(pptx.ShapeType.rect, {
    x: CX, y: 1.45, w: 2.0, h: 0.08,
    fill: { color: COLOR.orange },
    line: { color: COLOR.orange, width: 0 },
  });

  const col1 = [];
  const col2 = [];
  for (let i = 1; i <= 10; i++) {
    (i <= 5 ? col1 : col2).push(`${i < 10 ? "0" + i : i}.  Startup ${i}`);
  }

  const listOpts = {
    color: COLOR.dark,
    fontSize: 16,
    fontFace: FONT_BODY,
    lineSpacingMultiple: 1.6,
  };
  slide.addText(col1.join("\n"), { x: CX, y: 1.7, w: CW / 2 - 0.2, h: 5.5, ...listOpts });
  slide.addText(col2.join("\n"), { x: CX + CW / 2 + 0.1, y: 1.7, w: CW / 2 - 0.2, h: 5.5, ...listOpts });
}

// ─── Slides 3–12: Startup Profile (one per startup) ──────────────────────────
for (let n = 1; n <= 10; n++) {
  const slide = pptx.addSlide();
  addChrome(slide);

  const label = `Startup ${n}`;

  // Slide header bar
  slide.addShape(pptx.ShapeType.rect, {
    x: CX, y: 0, w: CW, h: 0.85,
    fill: { color: COLOR.navy },
    line: { color: COLOR.navy, width: 0 },
  });
  slide.addText(`STARTUP ${n}`, {
    x: CX + 0.2, y: 0, w: CW - 0.4, h: 0.85,
    color: COLOR.white,
    fontSize: 24,
    fontFace: FONT_TITLE,
    bold: true,
    valign: "middle",
  });

  // Website + tagline
  slide.addText("🌐  www.startup-website.com", {
    x: CX, y: 0.95, w: CW, h: 0.35,
    color: COLOR.navy,
    fontSize: 11,
    fontFace: FONT_BODY,
    hyperlink: { url: "http://www.startup-website.com" },
  });
  slide.addText('"One-liner tagline goes here – what does the startup do?"', {
    x: CX, y: 1.35, w: CW, h: 0.4,
    color: COLOR.dark,
    fontSize: 13,
    fontFace: FONT_TITLE,
    italic: true,
    bold: false,
  });

  // Orange divider
  slide.addShape(pptx.ShapeType.rect, {
    x: CX, y: 1.82, w: CW, h: 0.04,
    fill: { color: COLOR.orange },
    line: { color: COLOR.orange, width: 0 },
  });

  // ── Problem & Solution (two columns) ──
  const colW = (CW - 0.3) / 2;
  const probX = CX;
  const solX  = CX + colW + 0.3;
  const secY  = 1.95;

  // Problem
  sectionTitle(slide, "PROBLEM", secY, probX, colW - 0.5);
  const bullets = ["• Bullet point 1", "• Bullet point 2", "• Bullet point 3"];
  slide.addText(bullets.join("\n"), {
    x: probX, y: secY + 0.55, w: colW, h: 2.0,
    color: COLOR.dark,
    fontSize: 12,
    fontFace: FONT_BODY,
    lineSpacingMultiple: 1.4,
  });

  // Solution
  sectionTitle(slide, "SOLUTION", secY, solX, colW - 0.5);
  slide.addText(bullets.join("\n"), {
    x: solX, y: secY + 0.55, w: colW, h: 2.0,
    color: COLOR.dark,
    fontSize: 12,
    fontFace: FONT_BODY,
    lineSpacingMultiple: 1.4,
  });

  // ── Divider ──
  slide.addShape(pptx.ShapeType.line, {
    x: CX, y: 4.2, w: CW, h: 0,
    line: { color: COLOR.accent, width: 0.5 },
  });

  // ── Quick Facts row ──
  const chips = [
    { label: "INDUSTRY",  value: "Industry" },
    { label: "STAGE",     value: "Stage" },
    { label: "HQ",        value: "HQ" },
    { label: "TEAM SIZE", value: "Size" },
    { label: "ASK",       value: "Amount" },
  ];
  const chipW = CW / chips.length - 0.12;
  chips.forEach((chip, i) => {
    const cx2 = CX + i * (chipW + 0.12);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: cx2, y: 4.35, w: chipW, h: 0.85,
      fill: { color: COLOR.lightGray },
      line: { color: COLOR.navy, width: 0.5 },
      rectRadius: 0.08,
    });
    slide.addText(chip.label, {
      x: cx2, y: 4.38, w: chipW, h: 0.3,
      align: "center",
      color: COLOR.orange,
      fontSize: 8,
      fontFace: FONT_TITLE,
      bold: true,
    });
    slide.addText(`[${chip.value}]`, {
      x: cx2, y: 4.68, w: chipW, h: 0.5,
      align: "center",
      color: COLOR.dark,
      fontSize: 12,
      fontFace: FONT_BODY,
    });
  });
}

// ─── Slide 13: Closing ────────────────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addChrome(slide);

  slide.addText("THANK YOU", {
    x: CX, y: 2.5, w: CW, h: 1.5,
    color: COLOR.navy,
    fontSize: 54,
    fontFace: FONT_TITLE,
    bold: true,
    align: "center",
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: CX + CW / 4, y: 4.1, w: CW / 2, h: 0.1,
    fill: { color: COLOR.orange },
    line: { color: COLOR.orange, width: 0 },
  });

  slide.addText("Contact: [name@email.com]  |  [Phone]  |  [Website]", {
    x: CX, y: 4.35, w: CW, h: 0.55,
    align: "center",
    color: COLOR.dark,
    fontSize: 14,
    fontFace: FONT_BODY,
  });

  slide.addText("block71.co  |  @block71sg", {
    x: CX, y: 5.0, w: CW, h: 0.45,
    align: "center",
    color: COLOR.navy,
    fontSize: 13,
    fontFace: FONT_BODY,
    italic: true,
  });
}

// ─── Write output ─────────────────────────────────────────────────────────────
const outFile = path.join(__dirname, "Startup_Profile_Template_1920x1080_CanvaReady.pptx");
pptx.writeFile({ fileName: outFile }).then(() => {
  console.log(`✅  PPTX written to: ${outFile}`);
}).catch((err) => {
  console.error("❌  Error writing PPTX:", err);
  process.exit(1);
});
