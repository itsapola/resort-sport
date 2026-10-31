// Deck 02: Brand Identity & Guidelines, in the same system as Deck 01.
// Copy comes from the guidelines deck (with the edits from review). Run: NODE_PATH=decks/node_modules node decks/build-guidelines.js
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/b4cf305d-94fc-40e7-bc0e-635802eda6fb_520b0384-9679-424c-8e35-7ac2d1cfffc1/pptx/scripts/apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "Resort_Sport_Deck_02_Brand_Guidelines_Designed.pptx");

const THEME = {
  name: "Resort Sport Ltd.",
  headFontFace: "Instrument Serif",
  bodyFontFace: "Instrument Serif",
  colors: {
    dk1: "11293A", lt1: "F1E9DD", dk2: "811E23", lt2: "E7DBCB",
    accent1: "811E23", accent2: "98C6C5", accent3: "CFB9A2", accent4: "123F34",
    accent5: "96562F", accent6: "5B6670", hlink: "811E23", folHlink: "5B6670",
  },
};
const HEX = { navy: "11293A", ivory: "F1E9DD", red: "811E23", panel: "E7DBCB", aqua: "98C6C5", sand: "CFB9A2", muted: "5B6670", line: "D9CCB8" };
const SANS = "Space Grotesk";
const W = 13.333, H = 7.5, M = 0.7;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Resort Sport Ltd. Brand Identity and Guidelines";
pres.author = "Resort Sport Ltd.";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- layouts ----------
const footer = (color) => [
  { text: { text: "RESORT SPORT LTD.  /  IDENTITY & GUIDELINES", options: { x: M, y: H - 0.5, w: 6, h: 0.3, fontFace: SANS, fontSize: 9, charSpacing: 3, color, margin: 0 } } },
];
const slideNum = (color) => ({ x: W - M - 0.6, y: H - 0.5, w: 0.6, h: 0.3, fontFace: SANS, fontSize: 9, color, align: "right" });
const eyebrowPh = (color, y = 0.55) => ({ placeholder: { options: { name: "eyebrow", type: "body", x: M, y, w: 10, h: 0.3, fontFace: SANS, fontSize: 10, charSpacing: 4, color, margin: 0, valign: "top" }, text: "" } });
const titlePh = (color, y = 0.85) => ({ placeholder: { options: { name: "title", type: "title", x: M, y, w: W - 2 * M, h: 0.95, fontSize: 38, color, margin: 0, valign: "top", align: "left", fit: "none" }, text: "" } });

pres.defineSlideMaster({ title: "CONTENT", background: { color: C.background1 },
  objects: [eyebrowPh(C.accent1), titlePh(C.text1), ...footer(C.accent6)], slideNumber: slideNum(C.accent6) });
pres.defineSlideMaster({ title: "PANEL", background: { color: C.background2 },
  objects: [eyebrowPh(C.accent1), titlePh(C.text1), ...footer(C.accent6)], slideNumber: slideNum(C.accent6) });
pres.defineSlideMaster({ title: "NAVY", background: { color: C.text1 },
  objects: [eyebrowPh(C.accent2), titlePh(C.background1), ...footer(C.accent2)], slideNumber: slideNum(C.accent2) });
pres.defineSlideMaster({ title: "RED", background: { color: C.accent1 },
  objects: [eyebrowPh(C.background1), ...footer(C.background1)], slideNumber: slideNum(C.background1) });
pres.defineSlideMaster({ title: "BLANK", background: { color: C.text1 }, objects: [] });

// ---------- helpers ----------
const meta = {};
async function dims(f) { if (!meta[f]) meta[f] = await sharp(A(f)).metadata(); return meta[f]; }
// crop an image to fill a box exactly (focus 0..1)
async function cover(slide, f, x, y, w, h, fx = 0.5, fy = 0.5, extra = {}) {
  const m = await dims(f); const r = w / h;
  let cw = m.width, ch = Math.round(m.width / r);
  if (ch > m.height) { ch = m.height; cw = Math.round(m.height * r); }
  const left = Math.round((m.width - cw) * fx), top = Math.round((m.height - ch) * fy);
  const buf = await sharp(A(f)).extract({ left, top, width: cw, height: ch }).jpeg({ quality: 88 }).toBuffer();
  slide.addImage({ data: "image/jpeg;base64," + buf.toString("base64"), x, y, w, h, ...extra });
}
// fit an image inside a box, centered
async function contain(slide, f, x, y, w, h, extra = {}) {
  const m = await dims(f); const s = Math.min(w / m.width, h / m.height);
  const iw = m.width * s, ih = m.height * s;
  slide.addImage({ path: A(f), x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih, ...extra });
}
const label = (s, text, x, y, w, color = C.accent1, extra = {}) =>
  s.addText(text.toUpperCase(), { x, y, w, h: 0.28, fontFace: SANS, fontSize: 10, charSpacing: 3, color, margin: 0, isTextBox: true, valign: "top", ...extra });
const body = (s, text, x, y, w, h, extra = {}) =>
  s.addText(text, { x, y, w, h, fontSize: 16, color: C.text1, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 6, ...extra });
const serif = (s, text, x, y, w, h, size, extra = {}) =>
  s.addText(text, { x, y, w, h, fontSize: size, color: C.text1, margin: 0, isTextBox: true, valign: "top", ...extra });
const box = (s, x, y, w, h, color, extra = {}) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color }, line: { type: "none" }, ...extra });
const rule = (s, x, y, w, color = C.accent3) => s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width: 0.75 } });
function slide(master, section, eyebrow, title) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (eyebrow) s.addText(eyebrow.toUpperCase(), { placeholder: "eyebrow" });
  if (title) s.addText(title, { placeholder: "title" });
  return s;
}
const it = (t) => ({ text: t, options: { italic: true } });
const rn = (t, o = {}) => ({ text: t, options: o });

const dash = (s, x, y, w, h) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: "none" }, line: { color: HEX.red, width: 1, dashType: "dash" } });
async function logo(s, f, x, y, w, extra = {}) { const m = await dims(f); s.addImage({ path: A(f), x, y, w, h: w * m.height / m.width, altText: "Resort Sport Ltd.", ...extra }); return w * m.height / m.width; }
async function logoH(s, f, x, y, h, extra = {}) { const m = await dims(f); s.addImage({ path: A(f), x, y, w: h * m.width / m.height, h, altText: "Resort Sport Ltd.", ...extra }); }
const swatchText = (s, x, y, w, name, hexv, role, dark) => {
  serif(s, name, x + 0.25, y + 0.25, w - 0.5, 0.5, 26, { color: dark ? C.background1 : C.text1 });
  label(s, "#" + hexv, x + 0.25, y + 0.8, w - 0.5, dark ? C.background1 : C.text1);
  if (role) body(s, role, x + 0.25, y + 1.15, w - 0.5, 0.9, { fontSize: 13.5, color: dark ? C.background1 : C.text1 });
};

(async () => {
  // ================= 01 COVER =================
  pres.addSection({ title: "Foundation" });
  let s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Foundation" });
  box(s, 0, 0, 7.4, H, C.background1);
  box(s, 7.4, 0, W - 7.4, H, C.accent2);
  await contain(s, "portofino-cognac.png", 7.9, 1.2, 4.9, 5.1);
  label(s, "Creative system  /  Issue 01", M, 0.75, 6, C.accent6);
  await logo(s, "logo-wordmark-stacked-cabana-red.png", M, 1.55, 5.2);
  label(s, "Deck 02  /  Identity & guidelines", M, 4.55, 6);
  serif(s, [rn("The art of being "), it("off duty.")], M, 4.9, 6.3, 1.0, 40, { color: C.text1 });
  label(s, "Hypothetical case study  ·  October 2026", M, 6.55, 6, C.accent6);

  // ================= 02 FOUNDATION =================
  s = slide("CONTENT", "Foundation", "The foundation", "One brand. Many expressions.");
  serif(s, "Resort Sport Ltd. brings the design discipline of luxury leather goods to the optimism of leisure.", M, 1.95, 10.5, 1.3, 30, { color: C.accent1 });
  [["Considered", "Precision in product, material, and visual execution.", "Pillar 02  ·  Considered for real life"],
    ["Expressive", "Confident character without novelty or costume.", "Pillar 01  ·  Design worth recognizing"],
    ["Off duty", "Ease, movement, and pleasure without pretension.", "Pillar 03  ·  A life off the clock"]].forEach(([t, d, p], i) => {
    const x = M + i * 4.05;
    box(s, x, 3.65, 3.75, 2.75, C.background2);
    serif(s, [it(t)], x + 0.3, 3.9, 3.2, 0.7, 40, { color: C.accent1 });
    body(s, d, x + 0.3, 4.75, 3.2, 0.9, { fontSize: 17 });
    label(s, p, x + 0.3, 5.85, 3.3, C.accent6, { fontSize: 9 });
  });

  // ================= 03 ARCHITECTURE =================
  s = slide("CONTENT", "Foundation", "Brand architecture", "The identity architecture");
  const arch = [["Master brand", "Resort Sport Ltd. Customer-facing wordmarks and corporate identity.", "logo-wordmark-cabana-red.png"],
    ["Product signature", "RS only: embossing, leather tags, zipper hardware, product marks.", "logo-monogram-cabana-red.png"],
    ["Secondary shorthand", "RSL: approved editorial, stationery and packaging contexts only.", "RSL"],
    ["Cultural properties", "The Time Off Club; The Journal editorial; Limited Issues seasonal products.", "PROPS"]];
  for (const [i, [t, d, v]] of arch.entries()) {
    const x = M + (i % 2) * 6.05, y = 1.95 + Math.floor(i / 2) * 2.05;
    box(s, x, y, 2.3, 1.8, C.background2);
    if (v === "RSL") serif(s, "RSL", x, y + 0.45, 2.3, 0.9, 48, { color: C.accent1, align: "center" });
    else if (v === "PROPS") { ["The Time Off Club", "The Journal", "Limited Issues"].forEach((p, k) => serif(s, [it(p)], x + 0.2, y + 0.3 + k * 0.42, 2.0, 0.4, 18, { color: C.accent1 })); }
    else await contain(s, v, x + 0.3, y + 0.35, 1.7, 1.1);
    label(s, t, x + 2.55, y + 0.15, 3.2); body(s, d, x + 2.55, y + 0.5, 3.2, 1.25, { fontSize: 15 });
  }
  box(s, M, 6.05, W - 2 * M, 0.55, C.accent1);
  label(s, "Rule", M + 0.25, 6.2, 1, C.background1);
  body(s, "Never use RSL, the full wordmark or the Time Off Club on physical bags.", 1.9, 6.15, 10.4, 0.35, { fontSize: 15, color: C.background1 });

  // ================= 04 PRIMARY WORDMARK =================
  pres.addSection({ title: "Logo" });
  s = slide("CONTENT", "Logo", "Logo system", "Primary wordmark");
  box(s, M, 1.95, 7.6, 4.65, C.background2);
  await contain(s, "logo-wordmark-stacked-cabana-red.png", M + 0.8, 2.45, 6.0, 3.65);
  box(s, 8.55, 1.95, 4.08, 2.2, C.text1); await contain(s, "logo-wordmark-stacked-ivory.png", 9.05, 2.25, 3.08, 1.6);
  box(s, 8.55, 4.4, 4.08, 2.2, C.accent1); await contain(s, "logo-wordmark-stacked-ivory.png", 9.05, 4.7, 3.08, 1.6);
  label(s, "Approved master artwork", M, 6.75, 6, C.accent6, { fontSize: 9 });
  s.addNotes("The stacked italic serif wordmark is the leading brand signature. Preserve the approved drawing, baseline, proportions and underline. Use original artwork; never recreate by typing the name.");

  // ================= 05 SECONDARY =================
  s = slide("CONTENT", "Logo", "Logo system", "Compact wordmark and monogram");
  box(s, M, 1.95, 7.2, 3.0, C.background2); await contain(s, "logo-wordmark-cabana-red.png", M + 0.6, 2.6, 6.0, 1.7);
  box(s, 8.15, 1.95, 4.48, 3.0, C.background2); await contain(s, "logo-monogram-cabana-red.png", 8.75, 2.3, 3.28, 2.3);
  label(s, "Compact wordmark", M, 5.2, 6); body(s, "Editorial, navigation and narrow formats: the site header, email headers and sign-offs.", M, 5.55, 6.8, 0.8, { fontSize: 16 });
  label(s, "RS monogram", 8.15, 5.2, 4.4); body(s, "The only mark permitted on physical products. May also appear in editorial and packaging.", 8.15, 5.55, 4.4, 0.9, { fontSize: 16 });

  // ================= 06 CLEAR SPACE =================
  s = slide("CONTENT", "Logo", "Logo system", "Clear space and scale");
  box(s, M, 1.95, 7.6, 4.65, C.background2);
  const ls = await dims("logo-wordmark-stacked-cabana-red.png"); const lw = 4.6, lh = lw * ls.height / ls.width, lx = M + (7.6 - lw) / 2, ly = 1.95 + (4.65 - lh) / 2;
  s.addImage({ path: A("logo-wordmark-stacked-cabana-red.png"), x: lx, y: ly, w: lw, h: lh, altText: "Resort Sport Ltd." });
  const cs = 0.42; dash(s, lx - cs, ly - cs, lw + 2 * cs, lh + 2 * cs);
  label(s, "R", lx - cs + 0.05, ly - cs - 0.3, 0.4, C.accent1);
  label(s, "Safe area", 8.75, 1.95, 3.8); body(s, "Keep a clear margin at least the height of the capital R stem on every side of the wordmark. Treat this as a working rule until the vector masters are approved.", 8.75, 2.3, 3.85, 1.9, { fontSize: 16 });
  label(s, "Minimum sizes", 8.75, 4.45, 3.8); body(s, "No hard minimums until vector masters exist. Test small-format legibility before production.", 8.75, 4.8, 3.85, 1.2, { fontSize: 16 });

  // ================= 07 DO / DON'T =================
  s = slide("CONTENT", "Logo", "Logo system", "Correct uses. Incorrect uses.");
  label(s, "Do: approved color combinations", M, 1.95, 6);
  const combos = [["Cabana Red on Ivory", "BG1", "logo-wordmark-cabana-red.png"], ["Ivory on Deep Navy", "NAVY", "logo-wordmark-ivory.png"], ["Ivory on Cabana Red", "RED", "logo-wordmark-ivory.png"]];
  for (const [i, [t, bg, f]] of combos.entries()) {
    const x = M + i * 4.05; box(s, x, 2.3, 3.75, 1.45, bg === "BG1" ? C.background2 : bg === "NAVY" ? C.text1 : C.accent1);
    await contain(s, f, x + 0.5, 2.6, 2.75, 0.85); label(s, t, x, 3.85, 3.7, C.text1, { fontSize: 9 });
  }
  label(s, "Don't", M, 4.45, 6);
  const bad = [["Retype the name", "TYPED"], ["Stretch", "bad-stretch.png"], ["Rotate", "bad-rotate.png"], ["Recolor", "bad-recolor.png"], ["Add effects", "bad-shadow.png"]];
  for (const [i, [t, f]] of bad.entries()) {
    const x = M + i * 2.42, w = 2.22; box(s, x, 4.8, w, 1.35, C.background2);
    if (f === "TYPED") serif(s, "Resort Sport Ltd.", x, 5.2, w, 0.55, 24, { color: C.accent1, align: "center", italic: true });
    else await contain(s, f, x + 0.25, 4.95, w - 0.5, 1.05);
    s.addText("×", { x: x + w - 0.42, y: 4.82, w: 0.38, h: 0.38, fontFace: SANS, fontSize: 18, color: C.accent1, align: "center", margin: 0, isTextBox: true });
    label(s, t, x, 6.25, w, C.text1, { fontSize: 9 });
  }

  // ================= 08 PALETTE =================
  pres.addSection({ title: "Color" });
  s = slide("CONTENT", "Color", "Color system", "A palette made for escape");
  const pal = [["Cabana Red", "811E23", C.accent1, true], ["Deep Navy", "11293A", C.text1, true], ["Pool Aqua", "98C6C5", C.accent2, false],
    ["Ivory", "F1E9DD", C.background1, false], ["Sand", "CFB9A2", C.accent3, false], ["Palm Green", "123F34", C.accent4, true]];
  pal.forEach(([n, h, c, dark], i) => {
    const x = M + i * 2.02, w = 1.9;
    box(s, x, 1.95, w, 4.6, c, i === 3 ? { line: { color: HEX.line, width: 0.75 } } : {});
    serif(s, n, x + 0.2, 5.25, w - 0.3, 0.5, 22, { color: dark ? C.background1 : C.text1 });
    label(s, "#" + h, x + 0.2, 5.85, w - 0.3, dark ? C.background1 : C.text1);
  });

  // ================= 09 COLOR ROLES =================
  s = slide("CONTENT", "Color", "Color system", "Color roles, not just colors");
  const roles = [["Cabana Red", C.accent1, "Brand authority, signature headlines, logo, primary buttons."], ["Deep Navy", C.text1, "Text, navigation, editorial structure and quiet contrast."],
    ["Ivory", C.background1, "The primary canvas for presentation and commerce."], ["Pool Aqua", C.accent2, "Resort accent and campaign cover; inspiration for the Curaçao Limited Issue."],
    ["Palm Green", C.accent4, "Supporting accent; Bijou is its racing-green Limited Issue expression."], ["Sand", C.accent3, "The lining of every bag; the packaging neutral."]];
  roles.forEach(([n, c, d], i) => {
    const x = M + (i % 2) * 6.05, y = 1.95 + Math.floor(i / 2) * 1.5;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.1, w: 0.75, h: 0.75, fill: { color: c }, line: { color: HEX.line, width: 0.75 } });
    serif(s, n, x + 1.0, y + 0.05, 4.6, 0.5, 26, { color: C.accent1 }); body(s, d, x + 1.0, y + 0.55, 4.7, 0.8, { fontSize: 15 });
  });
  label(s, "Product colors (Cognac, Avorio, Nero) appear only as colorway swatches, never as decoration.", M, 6.55, 11, C.accent6, { fontSize: 9 });

  // ================= 10 TYPE SYSTEM =================
  pres.addSection({ title: "Typography" });
  s = slide("CONTENT", "Typography", "Typography", "Typography is a system");
  serif(s, [rn("The art of being "), it("off duty.")], M, 1.95, 11.8, 1.3, 72, { color: C.accent1 });
  label(s, "Instrument Serif  ·  display and body text, regular and selective italic", M, 3.35, 11, C.accent6);
  s.addText("OFFICIAL EQUIPMENT FOR TIME OFF.", { x: M, y: 4.05, w: 11.8, h: 0.6, fontFace: SANS, fontSize: 26, charSpacing: 6, color: C.text1, margin: 0, isTextBox: true });
  label(s, "Space Grotesk  ·  labels, navigation, buttons, prices and specs", M, 4.75, 11, C.accent6);
  box(s, M, 5.4, W - 2 * M, 0.85, C.background2);
  body(s, "Body text stays direct and readable. Product specifications are never sacrificed for tone. The wordmark is logo artwork, not type.", M + 0.3, 5.6, 11.4, 0.5, { fontSize: 16 });

  // ================= 11 HIERARCHY =================
  s = slide("CONTENT", "Typography", "Typography", "Type hierarchy and usage");
  const hier = [["Display", "40–54 pt", [rn("A very good "), it("entrance.")], 44, null, "Instrument Serif Regular and Italic. Editorial statements in sentence case; italic emphasis sparingly."],
    ["Section", "26–34 pt", "The Details", 30, null, "Instrument Serif, or Space Grotesk in functional contexts."],
    ["Body", "14–18 pt", "Supple leather. A polished aluminum arch.", 18, null, "Instrument Serif Regular. Practical language, comfortable leading."],
    ["Utility", "9–12 pt", "RS-04  ·  $695  ·  30-DAY RETURNS", 11, SANS, "Space Grotesk Medium. Navigation, captions, SKUs and UI labels."]];
  hier.forEach(([n, sz, sample, fs, ff, d], i) => {
    const y = 1.95 + i * 1.15; rule(s, M, y, W - 2 * M, C.accent3);
    label(s, n, M, y + 0.2, 1.6); label(s, sz, M, y + 0.5, 1.6, C.accent6, { fontSize: 9 });
    s.addText(sample, { x: 2.5, y: y + 0.12, w: 5.6, h: 0.9, fontSize: fs, color: C.text1, margin: 0, isTextBox: true, valign: "middle", ...(ff ? { fontFace: ff, charSpacing: 3 } : {}) });
    body(s, d, 8.4, y + 0.2, 4.2, 0.85, { fontSize: 14 });
  });

  // ================= 12 FAMILY =================
  pres.addSection({ title: "Product" });
  s = slide("CONTENT", "Product", "Product design", "One family. Five signatures.");
  const fam = [["Monaco", "Heritage", "monaco-heritage.png"], ["Rally", "Cognac", "rally-cognac.png"], ["Riviera", "Heritage", "riviera-heritage.png"], ["Portofino", "Nero", "portofino-nero.png"], ["Regatta", "Avorio", "regatta-avorio.png"]];
  for (const [i, [n, c, f]] of fam.entries()) {
    const x = M + i * 2.42, w = 2.22; box(s, x, 1.95, w, 3.3, C.background2); await contain(s, f, x + 0.2, 2.25, w - 0.4, 2.7);
    serif(s, n, x, 5.45, w, 0.5, 26, { color: C.accent1 }); label(s, c, x, 6.0, w, C.text1);
  }

  // ================= 13 / 14 SIGNATURES =================
  const sig = [["Monaco", "RS-01  ·  Travel duffel", "Aerodynamic east–west duffel; vertical leather panels; two handles; removable strap.", "monaco-cognac.png"],
    ["Rally", "RS-02  ·  Belt bag", "Low curved belt bag; diagonal panel; single-zip access; adjustable waist or shoulder strap.", "rally-heritage.png"],
    ["Riviera", "RS-03  ·  Oversized carryall", "Extra-wide winged carryall; dipped center; flowing leather corner panels and handles.", "riviera-cognac.png"],
    ["Portofino", "RS-04  ·  Handbag", "Compact trapezoid; polished aluminum tubular arch handle (prototype direction); leather collar mounts; removable strap.", "portofino-cognac.png"],
    ["Regatta", "RS-05  ·  Backpack", "Sculptural zip-top daypack; asymmetric front panel; diagonal zip front pocket; separate 16-inch laptop access.", "regatta-cognac.png"]];
  for (const [part, items] of [["01–03", sig.slice(0, 3)], ["04–05", sig.slice(3)]]) {
    s = slide("CONTENT", "Product", "Product design", "Product signatures: " + part);
    const n = items.length, cw = (W - 2 * M - (n - 1) * 0.35) / n;
    for (const [i, [nm, code, d, f]] of items.entries()) {
      const x = M + i * (cw + 0.35); box(s, x, 1.95, cw, 2.75, C.background2); await contain(s, f, x + 0.3, 2.2, cw - 0.6, 2.25);
      label(s, code, x, 4.9, cw, C.accent6); serif(s, nm, x, 5.2, cw, 0.5, 28, { color: C.accent1 }); body(s, d, x, 5.75, cw - 0.1, 0.9, { fontSize: 14.5 });
    }
  }

  // ================= 15 STANDARDS =================
  s = slide("PANEL", "Product", "Product standards", "Non-negotiable construction rules");
  const std = [["Product mark", "RS only. Tonal blind debossing or discreet RS metal engraving.", "emboss.jpg"], ["Metal", "Brushed silver-tone hardware. Portofino: polished 6061-T6 aluminum arch, under prototype testing.", "handle.jpg"],
    ["Surfaces", "Supple, protected fine-grain leather. Heritage: natural canvas and cognac leather.", "edges.jpg"], ["Seams", "Tonal stitching that preserves each model's geometry and curvature.", "zips.jpg"],
    ["Interior", "Sand lining on every silhouette and every exterior colorway.", "lining.jpg"], ["No exceptions", "No gold metal, prominent exterior wordmark, alternate linings or material substitutions.", null]];
  for (const [i, [t, d, f]] of std.entries()) {
    const x = M + (i % 3) * 4.05, y = 1.95 + Math.floor(i / 3) * 2.3;
    if (f) await cover(s, f, x, y, 1.6, 2.0); else { box(s, x, y, 1.6, 2.0, C.accent1); s.addText("×", { x, y, w: 1.6, h: 2.0, fontFace: SANS, fontSize: 54, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true }); }
    label(s, t, x + 1.85, y + 0.05, 1.9); body(s, d, x + 1.85, y + 0.4, 1.9, 1.6, { fontSize: 13.5 });
  }

  // ================= 16 COLORWAYS =================
  pres.addSection({ title: "Colorways" });
  s = slide("CONTENT", "Colorways", "Colorways", "Four permanent expressions");
  const cws = [["Heritage", "Natural canvas & cognac leather", "portofino-heritage.png"], ["Cognac", "Cognac leather", "portofino-cognac.png"], ["Avorio", "Ivory leather", "portofino-avorio.png"], ["Nero", "Black leather", "portofino-nero.png"]];
  for (const [i, [n, d, f]] of cws.entries()) {
    const x = M + i * 3.03, w = 2.8; box(s, x, 1.95, w, 3.1, C.background2); await contain(s, f, x + 0.25, 2.25, w - 0.5, 2.5);
    serif(s, n, x, 5.25, w, 0.5, 30, { color: C.accent1 }); label(s, d, x, 5.85, w, C.text1, { fontSize: 9 });
  }
  label(s, "Every colorway uses a Sand lining  ·  Shown on the Portofino", M, 6.55, 9, C.accent6, { fontSize: 9 });

  // ================= 17 AT SCALE =================
  s = slide("CONTENT", "Colorways", "Colorways", "The colorway system at scale");
  for (const [r, prod] of [["Riviera", "riviera"], ["Rally", "rally"]].entries()) {
    for (const [i, c] of ["heritage", "cognac", "avorio", "nero"].entries()) {
      const x = M + i * 3.03, y = 1.95 + r * 2.35, w = 2.8; box(s, x, y, w, 1.75, C.background2); await contain(s, `${prod[1]}-${c}.png`, x + 0.3, y + 0.15, w - 0.6, 1.45);
      label(s, `${prod[0]}  /  ${c}`, x, y + 1.85, w, C.text1, { fontSize: 9 });
    }
  }

  // ================= 18 MATERIAL LOGIC =================
  s = slide("CONTENT", "Colorways", "Craft and finish", "Material logic by use");
  const mat = [["City / fashion", "Portofino and Rally: soft hand, refined drape, tactile appeal, protected leather finishes.", "carried.jpg", 0.7], ["Travel / daily", "Monaco and Regatta: supple but abrasion-resilient leather, functional hardware, reinforced stress points.", "jet.jpg", 0.35], ["Resort / carryall", "Riviera: soft sculptural body, durable anchoring, practical interior organization.", "lining.jpg", 0.5]];
  for (const [i, [t, d, f, fx]] of mat.entries()) {
    const x = M + i * 4.05; await cover(s, f, x, 1.95, 3.75, 2.6, fx, 0.5);
    label(s, t, x, 4.75, 3.7); body(s, d, x, 5.1, 3.7, 1.1, { fontSize: 15 });
  }
  label(s, "Material performance specifications are intended design standards, not certified manufacturing claims.", M, 6.55, 11, C.accent6, { fontSize: 9 });

  // ================= 19 TWO WORLDS =================
  pres.addSection({ title: "Photography" });
  s = slide("CONTENT", "Photography", "Photography", "Two visual worlds. One identity.");
  box(s, M, 1.95, 5.85, 3.8, C.background2); await contain(s, "monaco-heritage.png", M + 0.5, 2.3, 4.85, 3.1);
  await cover(s, "cover.jpg", 6.8, 1.95, 5.83, 3.8, 0.45, 0.5);
  label(s, "Product  /  precise", M, 5.95, 5.6); body(s, "Centered, true transparent PNG, front and 45° angles, faithful materials.", M, 6.25, 5.6, 0.5, { fontSize: 15 });
  label(s, "Editorial  /  cinematic", 6.8, 5.95, 5.6); body(s, "Mediterranean light, real leisure, the bag clearly visible in the scene.", 6.8, 6.25, 5.8, 0.5, { fontSize: 15 });

  // ================= 20 ART DIRECTION =================
  s = slide("CONTENT", "Photography", "Photography", "Editorial art direction");
  const ad = [["Light", "Mediterranean daytime clarity; late-afternoon warmth; considered nocturnal scenes."], ["Framing", "Architectural lines, directional movement, negative space for copy."],
    ["Styling", "Lived-in luxury; effortless clothes; no costume nostalgia or artificial excess."], ["Humanity", "Natural gesture, real leisure, people living in the scene, not posing for wealth."],
    ["Retouching", "Keep leather grain and true color. No harsh halos, inconsistent shadows or reshaped products."], ["Exclusions", "No stock wealth clichés, cartoon retro filters, neon overload or irrelevant props."]];
  ad.forEach(([t, d], i) => { const x = M + (i % 3) * 2.65, y = 1.95 + Math.floor(i / 3) * 2.3; rule(s, x, y, 2.4, C.accent3); label(s, t, x, y + 0.18, 2.4); body(s, d, x, y + 0.5, 2.4, 1.6, { fontSize: 14 }); });
  await cover(s, "cafe.jpg", 8.75, 1.95, 3.88, 2.15, 0.5, 0.4); await cover(s, "rally-worn.jpg", 8.75, 4.25, 3.88, 2.15, 0.5, 0.35);

  // ================= 21 VOICE =================
  pres.addSection({ title: "Voice" });
  s = slide("RED", "Voice", "Verbal identity");
  serif(s, [rn("Clever, "), it("never cute.")], M, 1.0, 11, 1.2, 66, { color: C.background1 });
  [["Confident", "Assertive. Understated. Never boastful."], ["Observant", "Culturally fluent, attentive to the details of leisure."], ["Dryly witty", "A subtle wink. No puns, exclamation points or manufactured urgency."], ["Clear when it counts", "Dimensions, materials, warranty, delivery and care are plain and exact."]].forEach(([t, d], i) => {
    const x = M + i * 3.03; rule(s, x, 3.9, 2.75, C.background1); label(s, t, x, 4.1, 2.75, C.background1); serif(s, d, x, 4.5, 2.75, 1.6, 22, { color: C.background1 });
  });

  // ================= 22 JOBS =================
  s = slide("CONTENT", "Voice", "Verbal identity", "One voice, different jobs");
  [["Campaign", [rn("Out of office. "), it("In our element.")]], ["Product", [it("A very good entrance.")]], ["Specification", "Supple leather. Detachable shoulder strap. Sand lining."], ["Service", "We're happy to help you choose the right bag."], ["Editorial", "A strong case for the long weekend."], ["Microcopy", [it("A fine choice.  ·  Until next time.  ·  Plans optional.")]]].forEach(([l, t], i) => {
    const y = 1.95 + i * 0.75; rule(s, M, y, W - 2 * M, C.accent3); label(s, l, M, y + 0.25, 2.5, C.accent6);
    serif(s, t, 3.4, y + 0.1, 9.2, 0.6, 28, { color: C.accent1 });
  });

  // ================= 23 WRITE THIS =================
  s = slide("CONTENT", "Voice", "Verbal identity", "Writing guidelines in practice");
  box(s, M, 1.95, 5.85, 3.5, C.text1); box(s, M + 6.1, 1.95, 5.85, 3.5, C.background2);
  label(s, "Write this", M + 0.35, 2.25, 5, C.accent2); label(s, "Not this", M + 6.45, 2.25, 5, C.accent1);
  serif(s, "“The Portofino. Sculptural metal handle. Soft leather body. Made for a very good entrance.”", M + 0.35, 2.7, 5.2, 2.5, 28, { color: C.background1 });
  serif(s, "“Introducing the ultimate must-have iconic luxury handbag! Elevate your lifestyle with unmatched quality!”", M + 6.45, 2.7, 5.2, 2.5, 28, { color: C.accent6, strike: "sngStrike" });
  label(s, "Principle", M, 5.8, 2); serif(s, [it("Personality creates interest; details earn trust. Never substitute wit for useful information.")], 2.3, 5.7, 10.3, 0.6, 24, { color: C.accent1 });

  // ================= 24–29 APPLICATIONS =================
  pres.addSection({ title: "Applications" });
  const app = async (title, eyebrow, f, fx, fy, note) => {
    const sl = slide("PANEL", "Applications", eyebrow, title);
    await cover(sl, f, M, 1.95, 8.3, 4.65, fx, fy, { line: { color: HEX.line, width: 0.75 } });
    body(sl, note, 9.3, 1.95, 3.3, 4.5, { fontSize: 16 });
    return sl;
  };
  await app("Identity in commerce", "Applications  /  homepage", "app-home.jpg", 0.5, 0, "The homepage opens on the campaign line, the Pool Aqua cover and the Portofino in Cognac. The compact wordmark leads the navigation; Cabana Red marks the announcement bar and every primary action.");
  await app("A product page that earns trust", "Applications  /  product page", "app-product.jpg", 0.5, 0, "Show, explain, reassure, close. Colorways first, then the line, the price and the button, then size, what fits, materials, care, delivery and returns, stated plainly.");

  s = slide("PANEL", "Applications", "Applications  /  packaging", "Unboxing is a brand moment");
  box(s, M, 1.95, 4.8, 3.2, C.background1, { line: { color: HEX.line, width: 0.75 } }); await contain(s, "logo-wordmark-cabana-red.png", M + 0.8, 3.05, 3.2, 0.9);
  label(s, "The box", M, 5.3, 4.8); body(s, "Ivory board, master wordmark in Cabana Red.", M, 5.6, 4.8, 0.5, { fontSize: 15 });
  box(s, 5.85, 1.95, 3.2, 3.2, C.accent3); await contain(s, "logo-monogram-ivory.png", 6.85, 2.95, 1.2, 1.2);
  label(s, "The dust bag", 5.85, 5.3, 3.2); body(s, "Sand cotton, RS monogram only.", 5.85, 5.6, 3.2, 0.5, { fontSize: 15 });
  box(s, 9.4, 1.95, 1.4, 2.2, C.background1, { line: { color: HEX.line, width: 0.75 } }); await contain(s, "logo-monogram-cabana-red.png", 9.6, 2.6, 1.0, 0.8);
  box(s, 11.0, 1.95, 1.63, 2.2, C.accent1); await contain(s, "logo-wordmark-ivory.png", 11.1, 2.85, 1.43, 0.4);
  label(s, "Hang tag and card", 9.4, 5.3, 3.2); body(s, "Monogram on the tag; wordmark on the card.", 9.4, 5.6, 3.3, 0.5, { fontSize: 15 });
  label(s, "Sand interior is a bag-lining rule, not a packaging requirement", M, 6.55, 10, C.accent6, { fontSize: 9 });

  await app("Editorial culture, not decorative nostalgia", "Applications  /  The Journal", "app-journal.jpg", 0.5, 0.15, "Field notes, leisure studies and destinations. The Journal gives people a reason to return when they aren't shopping, and shows the bags living in the scene.");

  s = slide("PANEL", "Applications", "Applications  /  correspondence", "Email that keeps its promises");
  await cover(s, "app-email.jpg", M, 1.95, 3.3, 4.65, 0.5, 0, { line: { color: HEX.line, width: 0.75 } });
  body(s, "The Time Off Club writes when something good comes up. Five lifecycle emails share one template: the compact wordmark, Instrument Serif, square Cabana Red buttons. No countdowns, no discount codes, full price always.", 4.4, 1.95, 4.0, 3.2, { fontSize: 16 });
  [["Welcome", "We'll keep this brief."], ["Launch", "Out of office. In our element."], ["Waitlist", "For leaving well."], ["Order", "A fine choice."], ["Limited Issue", "Some things are worth making an exception for."]].forEach(([l, t], i) => {
    const y = 1.95 + i * 0.9; rule(s, 8.8, y, 3.83, C.accent3); label(s, l, 8.8, y + 0.15, 3.8, C.accent6); serif(s, [it(t)], 8.8, y + 0.42, 3.83, 0.45, 18, { color: C.accent1 });
  });

  s = slide("PANEL", "Applications", "Applications  /  campaign", "One idea, every placement");
  await cover(s, "app-billboard.png", M, 1.95, 7.6, 3.26, 0.5, 0.5);
  await cover(s, "app-story.png", 8.55, 1.95, 1.95, 3.47, 0.5, 0.5); await cover(s, "app-ad.png", 10.7, 1.95, 1.93, 2.41, 0.5, 0.5);
  body(s, "Text-free photography with live type and the master wordmark, so every placement matches the guidelines exactly: billboard, story, feed and paid social.", M, 5.45, 7.6, 1.0, { fontSize: 16 });

  // ================= 30 LIMITED ISSUES =================
  pres.addSection({ title: "Limited Issues" });
  s = slide("CONTENT", "Limited Issues", "Color system  /  Limited Issues", "The cocktail menu");
  body(s, "Seasonal color names evoke leisure culture. They never replace the four permanent colorways.", M, 1.85, 8.2, 0.5, { fontSize: 17 });
  const li = [["Curaçao", "Pool Aqua", "98C6C5", false], ["Negroni", "Bitter Orange", "C2502E", true], ["Limoncello", "Butter Yellow", "F4E19A", false], ["Bijou", "Racing Green", "183D31", true]];
  li.forEach(([n, c, h, dark], i) => {
    const x = M + i * 2.15, w = 1.95; box(s, x, 2.55, w, 2.75, h);
    serif(s, [it(n)], x + 0.2, 4.05, w - 0.3, 0.5, 26, { color: dark ? C.background1 : C.text1 }); label(s, c + "  ·  #" + h, x + 0.2, 4.65, w - 0.3, dark ? C.background1 : C.text1, { fontSize: 8, h: 0.5 });
  });
  box(s, 9.45, 2.55, 3.18, 2.75, C.text1); await contain(s, "rally-curacao.png", 9.65, 2.85, 2.78, 1.6);
  label(s, "Issue 001  /  The Rally in Curaçao", 9.65, 4.7, 2.9, C.accent2, { fontSize: 9 });
  body(s, "Rules: cocktail names only for Limited Issues. RS-only branding, silver-tone hardware and Sand lining stay constant. Release numbers and products beyond Issue 001 are not yet assigned.", M, 5.6, 11.9, 0.8, { fontSize: 15 });

  // ================= 31 CLOSING =================
  pres.addSection({ title: "Closing" });
  s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Closing" });
  label(s, "The system in the world", M, 0.8, 6, C.accent2);
  serif(s, [rn("One mark. Five forms. "), it("Endless time off.")], M, 1.15, 11.5, 1.2, 50, { color: C.background1 });
  const fin = ["monaco-heritage.png", "rally-cognac.png", "riviera-heritage.png", "portofino-nero.png", "regatta-avorio.png"];
  for (const [i, f] of fin.entries()) await contain(s, f, M + i * 2.42, 2.9, 2.2, 2.3);
  await logo(s, "logo-wordmark-ivory.png", M, 6.25, 2.6);
  label(s, "Official equipment for time off.", 8.4, 6.45, 4.25, C.accent2, { align: "right" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
