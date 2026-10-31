// Deck 03: Product & Go-to-Market, in the same system as Decks 01 and 02.
// Copy comes from the GTM deck (with the edits from review). Run: NODE_PATH=decks/node_modules node decks/build-gtm.js
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/b4cf305d-94fc-40e7-bc0e-635802eda6fb_520b0384-9679-424c-8e35-7ac2d1cfffc1/pptx/scripts/apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "Resort_Sport_Deck_03_Product_GTM_Designed.pptx");

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
pres.title = "Resort Sport Ltd. Product and Go-to-Market";
pres.author = "Resort Sport Ltd.";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- layouts ----------
const footer = (color) => [
  { text: { text: "RESORT SPORT LTD.  /  PRODUCT & GO-TO-MARKET", options: { x: M, y: H - 0.5, w: 6, h: 0.3, fontFace: SANS, fontSize: 9, charSpacing: 3, color, margin: 0 } } },
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

const tblText = (t, o = {}) => ({ text: t, options: { fontFace: THEME.bodyFontFace, fontSize: 13, color: HEX.navy, valign: "middle", ...o } });
const tblHead = (t, o = {}) => ({ text: t.toUpperCase(), options: { fontFace: SANS, fontSize: 9, charSpacing: 2, color: HEX.muted, valign: "bottom", ...o } });
const tblBorder = { type: "solid", pt: 0.75, color: HEX.line };
const chartText = { catAxisLabelFontFace: SANS, valAxisLabelFontFace: SANS, dataLabelFontFace: SANS, catAxisLabelColor: HEX.navy, valAxisLabelColor: HEX.muted, dataLabelColor: HEX.navy, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 12 };
const tag = (s, t, x, y, fill, color) => s.addText(t.toUpperCase(), { x, y, w: 1.2, h: 0.32, fontFace: SANS, fontSize: 9, charSpacing: 3, color, fill: { color: fill }, align: "center", valign: "middle", margin: 0, isTextBox: true });

(async () => {
  // ================= 01 COVER =================
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Opening" });
  await cover(s, "gtm-cover.jpg", 7.3, 0, W - 7.3, H, 0.5, 0.6);
  label(s, "Deck 03  /  Commercial strategy", M, 0.8, 6, C.accent2);
  serif(s, [rn("Create desire.\nEstablish value.\n"), it("Capture demand.")], M, 1.3, 6.3, 3.3, 58, { color: C.background1 });
  body(s, "Product positioning and go-to-market strategy: a proposed commercial system for the art of being off duty.", M, 4.85, 5.8, 0.9, { fontSize: 18, color: C.background1 });
  const wm = await dims("wordmark-ivory.png");
  s.addImage({ path: A("wordmark-ivory.png"), x: M, y: 6.3, w: 2.4, h: 2.4 * wm.height / wm.width, altText: "Resort Sport Ltd." });
  label(s, "Hypothetical case study  ·  October 2026", 3.6, 6.42, 3.6, C.accent2, { fontSize: 9 });

  // ================= 02 MANDATE =================
  s = slide("CONTENT", "Opening", "Executive brief", "The executive mandate");
  serif(s, [rn("Build desire for an unknown brand. Earn trust at a premium price. "), it("Scale without diluting the point of view.")], M, 1.95, 11.6, 1.4, 32, { color: C.accent1 });
  [["The business question", "How does an emerging accessories label become a credible alternative to established leather-goods brands?"],
    ["The product question", "Which distinctive silhouette earns the first purchase, and which builds long-term brand recognition?"],
    ["The marketing question", "How does a world of leisure translate into qualified discovery, confidence and conversion?"]].forEach(([l, q], i) => {
    const x = M + i * 4.05; box(s, x, 3.75, 3.75, 2.65, C.background2); label(s, l, x + 0.3, 4.0, 3.2); serif(s, q, x + 0.3, 4.35, 3.2, 1.9, 21);
  });

  // ================= 03 THESIS =================
  s = slide("NAVY", "Opening", "Commercial thesis", "One proposition. Three engines.");
  [["01", "Desirability", "Recognizable silhouettes and an editorial world worth entering.", "Pillars 01 + 03"], ["02", "Credibility", "Material clarity, design details and low-friction product evidence.", "Pillar 02"], ["03", "Accessibility", "Accessible-premium price positioning and intentional entry products.", "Deck 01 price role"]].forEach(([n, t, d, p], i) => {
    const x = M + i * 4.05; rule(s, x, 2.1, 3.7, C.accent2);
    label(s, n, x, 2.3, 1, C.accent2); serif(s, t, x, 2.6, 3.7, 0.7, 40, { color: C.background1 });
    body(s, d, x, 3.45, 3.5, 1.0, { fontSize: 17, color: C.background1 }); label(s, p, x, 4.55, 3.5, C.accent2, { fontSize: 9 });
  });
  label(s, "Operating principle", M, 5.5, 3, C.accent2);
  serif(s, [it("Emotional differentiation earns attention; product substance earns the transaction.")], M, 5.85, 11.6, 0.6, 26, { color: C.background1 });

  // ================= 04 SEGMENTATION =================
  pres.addSection({ title: "Customer & competition" });
  s = slide("CONTENT", "Customer & competition", "Segmentation", "Prioritize buyers, not personas");
  const seg = [["Primary", "The design seeker", [["Occasion", "A distinctive everyday handbag"], ["Trigger", "Tired of interchangeable premium bags"], ["Values", "Silhouette, leather, signature hardware"], ["Barrier", "Unfamiliar brand, unproven quality"]], "Hero: Portofino", "portofino-cognac.png", C.text1, C.background1, C.accent2],
    ["Secondary", "The off-duty mover", [["Occasion", "Travel, weekend, hands-free use"], ["Trigger", "Wants a beautiful bag that earns its keep"], ["Values", "Ease, durability, modular versatility"], ["Barrier", "Practicality versus price"]], "Entry: Rally  ·  Expansion: Monaco, Regatta", "rally-heritage.png", C.background2, C.text1, C.accent1]];
  for (const [i, [l, t, rows, hero, img, bg, fg, ac]] of seg.entries()) {
    const x = M + i * 6.05; box(s, x, 1.95, 5.85, 4.3, bg);
    label(s, l, x + 0.35, 2.2, 3, ac); serif(s, t, x + 0.35, 2.5, 3.6, 0.6, 32, { color: fg });
    await contain(s, img, x + 3.95, 2.25, 1.6, 1.0);
    rows.forEach(([k, v], j) => { label(s, k, x + 0.35, 3.4 + j * 0.55, 1.4, ac, { fontSize: 9 }); body(s, v, x + 1.75, 3.34 + j * 0.55, 3.9, 0.45, { fontSize: 15, color: fg }); });
    label(s, hero, x + 0.35, 5.75, 5.2, ac);
  }
  body(s, [rn("Decision: ", { bold: true }), rn("lead with design-led handbag demand; use utility-led buyers to widen occasions. A lower price does not imply a lower acquisition cost.")], M, 6.4, 11.9, 0.45, { fontSize: 14, color: C.accent6 });

  // ================= 05 SWITCHING =================
  s = slide("CONTENT", "Customer & competition", "Competitive choice", "The switching argument");
  const swRows = [["Polène", "Sculptural shapes", "Leisure-led versatility beyond handbags"], ["Mansur Gavriel", "Soft leather minimalism", "A more differentiated cultural world"], ["Cuyana", "Utility and value", "More ownable shape and personality"], ["DeMellier", "Refined hardware", "Sport and resort movement language"], ["STAUD / BY FAR", "Playful fashion", "Timeless materials within a broader system"]];
  s.addTable([[tblHead("Current alternative"), tblHead("What it already owns"), tblHead("Our intended reason to switch")],
    ...swRows.map(([a, b, c]) => [tblText(a, { fontSize: 20, color: HEX.red }), tblText(b, { fontSize: 15 }), tblText(c, { fontSize: 15 })])],
    { x: M, y: 1.95, w: W - 2 * M, colW: [3.4, 3.6, 4.93], rowH: [0.4, 0.78, 0.78, 0.78, 0.78, 0.78], border: [{ type: "none" }, { type: "none" }, tblBorder, { type: "none" }], margin: [0.05, 0.05, 0.05, 0], fill: { color: HEX.ivory } });

  // ================= 06 PORTOFINO COMPARISON =================
  s = slide("CONTENT", "Customer & competition", "Product-level competitive intelligence", "The trade-offs a buyer actually makes");
  const cmp = [["Retail price", "$695 leather launch colorways; $495–695 band", "$540 US", "£425 UK"], ["Body size (W×H×D)", "24 × 17 × 9.5 cm target", "22 × 14 × 12.5 cm", "24 × 16.5 × 10 cm base"],
    ["Weight", "≤700 g design target", "430 g published", "435 g published"], ["Exterior and finish", "Protected fine-grain leather; polished aluminum arch (target)", "Full-grain Italian calfskin; leather drape", "Italian grained cowhide; gold hardware"],
    ["Carry and closure", "Rigid arch handle; detachable strap; top zip", "Hand, shoulder, crossbody; magnetic", "Top handle, crossbody; magnetic flap"], ["Distinctive strength", "Polished aluminum signature; secure top zip", "Proven sculptural execution; lighter bag", "Established craft and service; expandable sides"]];
  s.addTable([[tblHead("Decision criterion"), tblHead("Portofino  /  concept", { color: HEX.red }), tblHead("Polène Numéro Neuf Mini"), tblHead("DeMellier Small Hudson")],
    ...cmp.map(([a, b, c, d]) => [tblText(a, { fontFace: SANS, fontSize: 10, color: HEX.muted }), tblText(b, { color: HEX.red }), tblText(c), tblText(d)])],
    { x: M, y: 1.95, w: W - 2 * M, colW: [2.3, 3.4, 3.1, 3.13], rowH: [0.38, 0.55, 0.55, 0.55, 0.62, 0.62, 0.62], border: [{ type: "none" }, { type: "none" }, tblBorder, { type: "none" }], margin: [0.04, 0.08, 0.04, 0], fill: { color: HEX.ivory } });
  box(s, M, 5.95, W - 2 * M, 0.75, C.background2);
  label(s, "PMM implication", M + 0.3, 6.18, 2.2);
  body(s, "Position the aluminum arch as a lightweight design target, not a proven advantage. Validate finish wear, structural load and comfort before making claims.", 3.0, 6.08, 9.4, 0.55, { fontSize: 14 });
  s.addNotes("Source: Polène US Numéro Neuf Mini; DeMellier Small Hudson Mocha Small Grain (official product pages). Portofino specifications are illustrative design targets. Prices in native market currencies.");

  // ================= 07 REQUIREMENTS =================
  pres.addSection({ title: "Product" });
  s = slide("CONTENT", "Product", "RS-04  /  Proposed product requirements", "Compact, not impractically small");
  [["24 cm", "Base width"], ["17 cm", "Body height"], ["9.5 cm", "Max depth"], ["2.5–3.0 L", "Usable capacity"], ["≤700 g", "Empty weight target"]].forEach(([n, l], i) => {
    const x = M + i * 2.42; serif(s, n, x, 1.95, 2.35, 0.9, 46, { color: C.accent1 }); label(s, l, x, 2.9, 2.2, C.accent6);
  });
  rule(s, M, 3.45, W - 2 * M);
  label(s, "What fits", M, 3.7, 5.6); body(s, "Phone up to about 6.7 in, a compact wallet, sunglasses in a slim case, keys, lipstick. Not made for a tablet, a large water bottle or full-size headphones.", M, 4.05, 5.6, 1.3, { fontSize: 16 });
  label(s, "Engineering trade-offs", 6.75, 3.7, 5.8);
  body(s, "Prototype hollow 6061-T6 polished aluminum and benchmark steel. Compare clear protective finishes for wear. Verify reinforced mounts, strength, comfort and corrosion isolation. Preserve the top zip and check real usable volume.", 6.75, 4.05, 5.85, 1.6, { fontSize: 15 });
  box(s, M, 5.75, W - 2 * M, 0.8, C.accent1);
  label(s, "Decision gate", M + 0.3, 6.0, 2.2, C.background1);
  body(s, "Approve the aluminum arch only if strength, comfort, finish durability and the ≤700 g target pass prototype review.", 3.0, 5.92, 9.4, 0.55, { fontSize: 15, color: C.background1 });

  // ================= 08 PORTFOLIO =================
  s = slide("CONTENT", "Product", "Portfolio architecture", "The complete line. The selective launch.");
  const port = [["RS-01", "Monaco", "Travel authority", "$695–895", "monaco-heritage.png", false], ["RS-02", "Rally", "Price entry / trial", "$275–350", "rally-heritage.png", true], ["RS-03", "Riviera", "Leisure halo", "$595–795", "riviera-heritage.png", false], ["RS-04", "Portofino", "Design hero", "$495–695", "portofino-cognac.png", true], ["RS-05", "Regatta", "Utility bridge", "$525–725", "regatta-heritage.png", false]];
  for (const [i, [c, n, r, p, f, launch]] of port.entries()) {
    const x = M + i * 2.42, w = 2.22; box(s, x, 1.95, w, 2.55, launch ? C.accent2 : C.background2); await contain(s, f, x + 0.2, 2.2, w - 0.4, 2.0);
    if (launch) tag(s, "Launch", x + 0.12, 2.07, HEX.navy, HEX.ivory);
    label(s, c, x, 4.65, w, C.accent6); serif(s, n, x, 4.95, w, 0.5, 26, { color: C.accent1 }); label(s, r, x, 5.5, w, C.text1); serif(s, p, x, 5.85, w, 0.45, 20);
  }
  label(s, "Price bands are proposed US retail hypotheses, not quotes or validated willingness to pay.", M, 6.55, 11.9, C.accent6, { fontSize: 9 });

  // ================= 09 / 10 HERO + ENTRY =================
  const heroSlide = async (eyebrow, title, img, fx, rows, msg) => {
    const sl = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Product" });
    await cover(sl, img, 0, 0, 5.6, H, fx, 0.5);
    label(sl, eyebrow, 6.3, 0.55, 6); serif(sl, title, 6.3, 0.85, 6.3, 1.4, 38, { color: C.text1 });
    rows.forEach(([k, v], i) => { const y = 2.55 + i * 0.85; rule(sl, 6.3, y, 6.3, C.accent3); label(sl, k, 6.3, y + 0.17, 1.9, C.accent6); body(sl, v, 8.3, y + 0.12, 4.3, 0.72, { fontSize: 15 }); });
    serif(sl, [it(msg)], 6.3, 6.05, 6.3, 0.6, 32, { color: C.accent1 });
  };
  await heroSlide("Hero product  /  RS-04 The Portofino", "Portofino creates the image of the brand", "carried.jpg", 0.7,
    [["Customer job", "Own a recognizable, versatile handbag without conspicuous logos."], ["Hero asset", "The polished aluminum arch handle (proposed; finish validation pending)."], ["Value proposition", "Architectural elegance, everyday ease."], ["Launch", "Cognac, Avorio, Nero  ·  $695"]], "“A very good entrance.”");
  await heroSlide("Entry product  /  RS-02 The Rally", "Rally expands the entry occasion", "rally-worn.jpg", 0.5,
    [["Customer job", "Hands-free freedom without the sporty-nylon aesthetic."], ["Signature", "A low aerodynamic profile, sweeping seam, fine-grain leather."], ["Value proposition", "Lighter commitment, daily usefulness."], ["Launch", "Heritage $275  ·  Cognac, Nero $350"]], "“Hands free. Plans optional.”");

  // ================= 11 FEEDBACK LOOP =================
  s = slide("CONTENT", "Product", "Product-market feedback loop", "Marketing shapes what gets built");
  const fb = [["Brand trust", "Leather sourcing proof; finish close-ups", "Set the buyer's evidence requirements"], ["Carry comfort", "Prototype strap width, weight and grip", "Turn objections into usability criteria"], ["Value for price", "Prioritize visible quality over hidden cost", "Separate must-have from nice-to-have features"], ["Versatility", "Detachable straps, Sand interiors, capacity", "Specify usage scenarios and acceptance tests"], ["Distinctiveness", "Protect the recognizable arch and geometry", "Defend the features that create preference"]];
  s.addTable([[tblHead("Customer signal"), tblHead("Product response"), tblHead("PMM role")], ...fb.map(([a, b, c]) => [tblText(a, { fontSize: 20, color: HEX.red }), tblText(b, { fontSize: 15 }), tblText(c, { fontSize: 15 })])],
    { x: M, y: 1.95, w: W - 2 * M, colW: [3.2, 4.4, 4.33], rowH: [0.4, 0.7, 0.7, 0.7, 0.7, 0.7], border: [{ type: "none" }, { type: "none" }, tblBorder, { type: "none" }], margin: [0.05, 0.05, 0.05, 0], fill: { color: HEX.ivory } });
  body(s, "PMM leads customer insight and market requirements; Product, Design and Operations own technical specification and feasibility.", M, 6.1, 11.9, 0.5, { fontSize: 14, color: C.accent6 });

  // ================= 12 PORTFOLIO STRATEGY =================
  s = slide("PANEL", "Product", "Portfolio strategy", "Use breadth to expand, never to blur");
  for (const [i, [n, r, f]] of [["Riviera", "Resort carry  /  leisure credibility", "riviera-heritage.png"], ["Monaco", "Weekend travel  /  brand authority", "monaco-heritage.png"], ["Regatta", "Work-to-weekend  /  functional bridge", "regatta-heritage.png"]].entries()) {
    const x = M + i * 4.05; box(s, x, 1.95, 3.75, 3.1, C.background1); await contain(s, f, x + 0.35, 2.2, 3.05, 2.6);
    serif(s, n, x, 5.25, 3.75, 0.55, 32, { color: C.accent1 }); label(s, r, x, 5.85, 3.75, C.text1);
  }

  // ================= 13 OPTIONS =================
  pres.addSection({ title: "Assortment & price" });
  s = slide("CONTENT", "Assortment & price", "Investment options", "The launch choice is not inevitable");
  [["Portofino only", "Sharpest visual and message focus", "Higher entry price; narrower daily use", "Reject", C.background2, C.text1, HEX.muted, HEX.ivory],
    ["Portofino + Rally", "Recognition plus a second use occasion", "Needs distinct messaging and buyer proof", "Select", C.accent1, C.background1, HEX.ivory, HEX.red],
    ["All five", "Broad ecosystem and more categories", "20-SKU complexity; inventory and focus risk", "Defer", C.background2, C.text1, HEX.muted, HEX.ivory]].forEach(([t, u, d, v, bg, fg, tf, tc], i) => {
    const x = M + i * 4.05; box(s, x, 1.95, 3.75, 4.6, bg); tag(s, v, x + 0.3, 2.25, tf, tc);
    serif(s, t, x + 0.3, 2.8, 3.2, 0.7, 34, { color: fg });
    label(s, "Upside", x + 0.3, 3.85, 3, fg); body(s, u, x + 0.3, 4.15, 3.2, 0.7, { fontSize: 16, color: fg });
    label(s, "Trade-off", x + 0.3, 5.05, 3, fg); body(s, d, x + 0.3, 5.35, 3.2, 0.9, { fontSize: 16, color: fg });
  });

  // ================= 14 SIX SKUS =================
  s = slide("CONTENT", "Assortment & price", "Assortment discipline", "20 permanent options are not 20 launch bets");
  const six = [["Portofino", "Cognac", "portofino-cognac.png"], ["Portofino", "Avorio", "portofino-avorio.png"], ["Portofino", "Nero", "portofino-nero.png"], ["Rally", "Heritage", "rally-heritage.png"], ["Rally", "Cognac", "rally-cognac.png"], ["Rally", "Nero", "rally-nero.png"]];
  for (const [i, [n, c, f]] of six.entries()) {
    const x = M + i * 2.02, w = 1.85; box(s, x, 1.95, w, 1.85, C.background2); await contain(s, f, x + 0.15, 2.1, w - 0.3, 1.55);
    serif(s, n, x, 3.95, w, 0.45, 20, { color: C.accent1 }); label(s, c, x, 4.4, w, C.text1);
  }
  label(s, "Chosen  /  six SKUs", M, 5.05, 5.6); body(s, "A recognizable hero plus an accessible entry; controlled inventory and a coherent visual assortment.", M, 5.4, 5.6, 0.9, { fontSize: 15 });
  label(s, "Expansion  /  staged", 6.75, 5.05, 5.8); body(s, "Add Riviera and Monaco after hero evidence; test Regatta separately for fit, weight and laptop use. Extend color depth only after demand signals.", 6.75, 5.4, 5.85, 1.0, { fontSize: 15 });

  // ================= 15 CHOICE LOGIC =================
  s = slide("PANEL", "Assortment & price", "Assortment choice logic", "Six SKUs: concentrate where the story is strongest");
  [["The hero  /  three colors", "Portofino: Cognac, Avorio, Nero", "Cognac establishes rich leather character; Avorio supports editorial fashion imagery; Nero answers classic demand. Heritage is held back for later learning."],
    ["The entry  /  three colors", "Rally: Heritage, Cognac, Nero", "Heritage communicates resort identity; Cognac carries the leather signature; Nero supports everyday use. Avorio can follow observed demand."]].forEach(([l, t, d], i) => {
    const x = M + i * 6.05; box(s, x, 1.95, 5.85, 3.55, C.background1); label(s, l, x + 0.35, 2.25, 5); serif(s, t, x + 0.35, 2.6, 5.2, 0.6, 28, { color: C.accent1 }); body(s, d, x + 0.35, 3.4, 5.1, 1.9, { fontSize: 16 });
  });
  label(s, "Why not four each?", M, 5.85, 3); body(s, "Six options keep the core contrast while limiting early color-by-product inventory fragmentation.", 3.4, 5.78, 9.2, 0.5, { fontSize: 15 });

  // ================= 16 MERCHANDISING =================
  s = slide("CONTENT", "Assortment & price", "Merchandising", "Permanent discipline. Seasonal permission.");
  [["Heritage", "Natural canvas & cognac leather", "portofino-heritage.png"], ["Avorio", "Soft ivory leather", "portofino-avorio.png"], ["Cognac", "Warm cognac leather", "portofino-cognac.png"], ["Nero", "Black leather", "portofino-nero.png"]].entries().forEach(() => {});
  for (const [i, [n, d, f]] of [["Heritage", "Natural canvas & cognac leather", "portofino-heritage.png"], ["Avorio", "Soft ivory leather", "portofino-avorio.png"], ["Cognac", "Warm cognac leather", "portofino-cognac.png"], ["Nero", "Black leather", "portofino-nero.png"]].entries()) {
    const x = M + i * 2.15, w = 1.95; box(s, x, 1.95, w, 1.85, C.background2); await contain(s, f, x + 0.15, 2.1, w - 0.3, 1.55);
    serif(s, n, x, 3.95, w, 0.45, 22, { color: C.accent1 }); label(s, d, x, 4.45, w, C.text1, { fontSize: 8, h: 0.5 });
  }
  box(s, 9.45, 1.95, 3.18, 3.1, C.text1);
  label(s, "Limited Issues  /  cocktail colorways", 9.7, 2.2, 2.8, C.accent2, { fontSize: 9, h: 0.45 });
  ["Curaçao", "Negroni", "Limoncello", "Bijou"].forEach((n, i) => serif(s, [it(n)], 9.7, 2.75 + i * 0.52, 2.8, 0.45, 24, { color: C.background1 }));
  body(s, "The four permanent colorways never change. Limited Issues add a seasonal reason to talk, without replacing them.", M, 5.35, 11.9, 0.8, { fontSize: 16 });

  // ================= 17 PRICE POSITIONING =================
  s = slide("CONTENT", "Assortment & price", "Price positioning", "Price follows value and portfolio role");
  const lo = 200, hi = 950, x0 = 3.6, x1 = W - M, px = (v) => x0 + (v - lo) / (hi - lo) * (x1 - x0);
  [200, 400, 600, 800].forEach((v) => { s.addShape(pres.shapes.LINE, { x: px(v), y: 2.0, w: 0, h: 3.45, line: { color: HEX.line, width: 0.5 } }); s.addText("$" + v, { x: px(v) - 0.4, y: 5.5, w: 0.8, h: 0.25, fontFace: SANS, fontSize: 10, color: C.accent6, align: "center", margin: 0, isTextBox: true }); });
  [["Rally", "Entry to brand", 275, 350], ["Portofino", "Iconic handbag", 495, 695], ["Regatta", "Utility investment", 525, 725], ["Riviera", "Large-scale leisure", 595, 795], ["Monaco", "Travel flagship", 695, 895]].forEach(([n, r, a, b], i) => {
    const y = 2.1 + i * 0.68; serif(s, n, M, y - 0.05, 1.6, 0.4, 20, { color: C.accent1 }); label(s, r, M + 1.5, y + 0.06, 1.5, C.accent6, { fontSize: 8, h: 0.4 });
    box(s, px(a), y + 0.02, px(b) - px(a), 0.26, i < 2 ? C.accent1 : C.accent3);
    s.addText(`$${a}–${b}`, { x: px(b) + 0.08, y: y - 0.02, w: 1.3, h: 0.34, fontFace: SANS, fontSize: 11, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  });
  box(s, M, 5.95, W - 2 * M, 0.65, C.background2);
  body(s, [rn("Full-price DTC positioning. ", { bold: true }), rn("No introductory discount; perceived value comes from product evidence and service. Launch products shown in Cabana Red.")], M + 0.3, 6.08, 11.3, 0.45, { fontSize: 14 });

  // ================= 18 GUARDRAILS =================
  pres.addSection({ title: "Economics" });
  s = slide("CONTENT", "Economics", "Commercial guardrails", "Price and unit economics must agree");
  [["Retail price", "Fix the market position and the customer's value perception."], ["Landed cost", "Target a product cost compatible with contribution margin."], ["Servicing cost", "Include packaging, complimentary delivery, free returns and support."], ["Channel cost", "Budget creative, partnerships and acquisition against net margin."]].forEach(([t, d], i) => {
    const x = M + i * 3.03; label(s, String(i + 1).padStart(2, "0"), x, 2.0, 1, C.accent6); serif(s, t, x, 2.3, 2.75, 0.55, 28, { color: C.accent1 }); body(s, d, x, 2.95, 2.7, 1.3, { fontSize: 15 });
    if (i < 3) s.addText("→", { x: x + 2.68, y: 2.3, w: 0.35, h: 0.5, fontSize: 22, color: C.accent1, margin: 0, align: "center", isTextBox: true });
  });
  box(s, M, 4.75, W - 2 * M, 1.0, C.accent1);
  label(s, "Gate", M + 0.3, 5.1, 1.2, C.background1);
  serif(s, "Do not authorize a price and product combination before sourcing and modeled contribution clear the hurdle.", 2.0, 4.95, 10.4, 0.7, 22, { color: C.background1 });

  // ================= 19 UNIT ECONOMICS =================
  s = slide("CONTENT", "Economics", "Commercial model  /  $695 DTC case", "Illustrative unit economics: Portofino");
  const steps = [["Retail price, net of tax", 695, 0, 695, "total"], ["Landed unit cost", -160, 535, 695, "neg"], ["Fulfillment, payment, packaging, returns", -65, 470, 535, "neg"],
    ["Contribution before acquisition", 470, 0, 470, "total"], ["New-customer CAC", -110, 360, 470, "neg"], ["Contribution after acquisition", 360, 0, 360, "final"]];
  const bx0 = 3.75, bx1 = 7.9, bv = (v) => bx0 + v / 695 * (bx1 - bx0);
  steps.forEach(([l, v, from, to, kind], i) => {
    const y = 2.0 + i * 0.72;
    body(s, l, M, y + 0.06, 2.95, 0.5, { fontSize: 14, align: "right", color: kind === "final" ? C.accent1 : C.text1, bold: kind === "final" });
    box(s, bv(from), y, Math.max(0.04, bv(to) - bv(from)), 0.48, kind === "neg" ? C.accent3 : C.accent1);
    s.addText((v < 0 ? "−$" : "$") + Math.abs(v), { x: bv(to) + 0.1, y, w: 1.0, h: 0.48, fontFace: SANS, fontSize: 13, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  });
  serif(s, "51.8%", 8.85, 2.0, 3.8, 1.2, 80, { color: C.accent1 });
  label(s, "Contribution after acquisition, as a share of net revenue", 8.85, 3.3, 3.7, C.accent6, { h: 0.5 });
  rule(s, 8.85, 4.05, 3.75);
  body(s, "The $65 servicing line covers complimentary US delivery, free return labels and packaging. All cost and acquisition values are invented scenario inputs; the model excludes fixed operating expenses and income taxes.", 8.85, 4.25, 3.75, 2.2, { fontSize: 14 });

  // ================= 20 SENSITIVITY =================
  s = slide("CONTENT", "Economics", "Economics  /  sensitivity", "A plan is only useful if it survives stress");
  s.addChart(pres.charts.BAR, [{ name: "Contribution after acquisition", labels: ["Base", "Cost +20%", "CAC ×2", "Both"], values: [360, 328, 250, 218] }],
    { x: M, y: 1.85, w: W - 2 * M, h: 2.9, barDir: "col", chartColors: [HEX.red], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"#,##0', valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 420, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 110, ...chartText, catAxisLabelFontSize: 12, dataLabelFontSize: 14 });
  [["CAC $110  ·  Cost $160", "Continue controlled learning"], ["CAC $110  ·  Cost $192", "Pressure-test construction and sourcing"], ["CAC $220  ·  Cost $160", "Stop scaling; fix efficiency and conversion"], ["CAC $220  ·  Cost $192", "Revisit offer, channel, materials and launch scale"]].forEach(([a, d], i) => {
    const cw = (W - 2 * M) / 4, x = M + i * cw + 0.15; label(s, a, x, 4.9, cw - 0.3, C.accent6, { fontSize: 9 }); body(s, d, x, 5.2, cw - 0.3, 0.8, { fontSize: 15 });
  });
  s.addText("Scenario math: contribution = $695 retail − landed cost − $65 servicing − CAC. Before fixed operating expenses.", { x: M, y: 6.4, w: 11, h: 0.3, fontFace: SANS, fontSize: 9, color: C.accent6, margin: 0, isTextBox: true });

  // ================= 21 MESSAGE ARCHITECTURE =================
  pres.addSection({ title: "Go-to-market" });
  s = slide("CONTENT", "Go-to-market", "Message architecture", "Messaging that moves from feeling to proof");
  const ma = [["Awareness", "Out of office. In our element.", "The leisure world; strong product geometry", "ad-aware.jpg"], ["Consideration", "Design worth recognizing.", "Leather, silver hardware, Sand lining, silhouette", "ad-consider.jpg"],
    ["Decision", "A very good entrance.", "Dimensions, capacity, scale, delivery, returns", "ad-decide.jpg"], ["Post-purchase", "See you out there.", "Care, use, support, collection cross-sell", "app-email.jpg"]];
  for (const [i, [st, line, proof, f]] of ma.entries()) {
    const x = M + i * 3.03, w = 2.8;
    await cover(s, f, x, 1.95, w, 2.45, 0.5, f === "app-email.jpg" ? 0.02 : 0.25);
    label(s, st, x, 4.55, w); serif(s, [it(line)], x, 4.85, w, 0.75, 22, { color: C.accent1 }); body(s, proof, x, 5.65, w, 0.8, { fontSize: 14 });
  }

  // ================= 22 CONVERSION =================
  s = slide("PANEL", "Go-to-market", "Conversion design", "Make the product page do the selling");
  await cover(s, "app-product.jpg", M, 1.95, 6.4, 4.6, 0.5, 0.0, { line: { color: HEX.line, width: 0.75 } });
  [["Show", "Front, 45°, on-body scale, interior"], ["Explain", "Leather, shape, silver hardware, Sand lining"], ["Reassure", "Dimensions, weight, capacity, care; complimentary US delivery; 30-day returns"], ["Close", "Clear colorway selection, price and one action"]].forEach(([t, d], i) => {
    const y = 1.95 + i * 1.15; label(s, "0" + (i + 1), 7.45, y + 0.05, 0.5, C.accent6); serif(s, t, 7.95, y - 0.05, 4.6, 0.5, 28, { color: C.accent1 }); body(s, d, 7.95, y + 0.5, 4.6, 0.6, { fontSize: 15 });
  });

  // ================= 23 CAMPAIGN =================
  s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Go-to-market" });
  s.background = { color: HEX.navy };
  await contain(s, "billboard.jpg", 0, 0.45, W, 5.75);
  label(s, "Campaign platform", M, 6.35, 3, C.accent2);
  body(s, "One idea, adapted by channel, never rewritten into generic luxury advertising. Collection 01  ·  “Your time off looks good from here.”", 3.2, 6.28, 9.4, 0.6, { fontSize: 16, color: C.background1 });

  // ================= 24 CHANNELS =================
  s = slide("CONTENT", "Go-to-market", "Channel priorities", "Earn attention before buying reach");
  const ch = [["Owned  /  site + email", "High", "Editorial destination, product education, waitlist and lifecycle email."], ["PR + cultural seeding", "High", "Design, travel, fashion and resort culture; earned credibility."], ["Creator partnerships", "High", "Selected aesthetic alignment; credible on-body demonstrations."], ["Paid social + search", "Test", "Retargeting and high-intent capture after message and creative signals."], ["Physical experiences", "Selective", "Pop-ups and partnerships only when audience fit is compelling."]];
  ch.forEach(([n, p, d], i) => {
    const y = 1.95 + i * 0.85; rule(s, M, y, W - 2 * M, C.accent3);
    serif(s, n, M, y + 0.15, 4.0, 0.55, 22, { color: C.text1 });
    tag(s, p, 4.85, y + 0.24, p === "High" ? HEX.red : p === "Test" ? HEX.aqua : HEX.panel, p === "High" ? HEX.ivory : HEX.navy);
    body(s, d, 6.4, y + 0.2, 6.2, 0.6, { fontSize: 15 });
  });

  // ================= 25 INVESTMENT MIX =================
  s = slide("CONTENT", "Go-to-market", "Launch  /  illustrative investment mix", "Allocate capital in learning stages");
  s.addChart(pres.charts.BAR, [{ name: "Share of launch marketing", labels: ["Creative & proof content", "Earned media & creators", "Paid testing", "Owned commerce & CRM", "Experiential reserve"], values: [30, 25, 20, 15, 10] }],
    { x: M, y: 1.9, w: 7.4, h: 4.4, barDir: "bar", catAxisOrientation: "maxMin", chartColors: [HEX.red], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 36, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 55, ...chartText, catAxisLabelFontSize: 12, dataLabelFontSize: 13 });
  [["Creative & proof", "Must-have product, scale and detail assets"], ["Earned & creators", "Selective brand-fit relationships"], ["Paid testing", "Controlled tests after content and readiness"], ["Owned & CRM", "Conversion information and retention"], ["Experiential", "Released only if fit and economics work"]].forEach(([t, d], i) => {
    const y = 2.0 + i * 0.85; label(s, t, 8.6, y, 4, C.accent1); body(s, d, 8.6, y + 0.3, 4.0, 0.5, { fontSize: 14 });
  });
  label(s, "Illustrative allocation of discretionary launch marketing, not total company budget.", M, 6.5, 11, C.accent6, { fontSize: 9 });

  // ================= 26 12 WEEKS =================
  s = slide("PANEL", "Go-to-market", "12-week illustrative plan", "Sequence the launch to build conviction");
  [["01", "Prep", "Weeks 1–4", "Prototype proof, product content, press kit, creator shortlist, waitlist."], ["02", "Reveal", "Weeks 5–8", "Campaign launch, Portofino hero, Rally education, earned placements."], ["03", "Convert", "Weeks 9–12", "Release six SKUs, lifecycle email, retargeting, product feedback and restock signals."]].forEach(([n, t, w, d], i) => {
    const x = M + i * 4.05; box(s, x, 2.6, 3.75 + (i < 2 ? 0.3 : 0), 0.12, i === 2 ? C.accent1 : C.accent3);
    s.addShape(pres.shapes.OVAL, { x, y: 2.43, w: 0.45, h: 0.45, fill: { color: i === 2 ? C.accent1 : C.text1 }, line: { type: "none" } });
    label(s, w, x, 3.15, 3.5, C.accent6); serif(s, t, x, 3.45, 3.6, 0.7, 44, { color: C.accent1 }); body(s, d, x, 4.35, 3.5, 1.5, { fontSize: 16 });
  });
  label(s, "Dates, availability and launch targets are case-study scenarios", M, 6.4, 8, C.accent6, { fontSize: 9 });

  // ================= 27 SEASONAL DROP =================
  s = pres.addSlide({ masterName: "NAVY", sectionTitle: "Go-to-market" });
  s.addText("SEASONAL DROP", { placeholder: "eyebrow" });
  await contain(s, "rally-curacao.png", 0.6, 1.4, 6.6, 4.6);
  label(s, "Limited Issue 001", 7.6, 1.6, 5, C.accent2);
  serif(s, [rn("The Rally in "), it("Curaçao.")], 7.6, 1.95, 5.1, 1.0, 46, { color: C.background1 });
  label(s, "The role", 7.6, 3.4, 5, C.accent2);
  body(s, "A collectible color story after core recognition. Not a discount event, and not a replacement for the permanent collection.", 7.6, 3.75, 5.0, 1.4, { fontSize: 18, color: C.background1 });
  serif(s, [it("Issue 001 creates the next conversation.")], 7.6, 5.4, 5.0, 0.6, 24, { color: C.accent2 });

  // ================= 28 SCORECARD =================
  pres.addSection({ title: "Governance" });
  s = slide("CONTENT", "Governance", "Executive scorecard", "Measure the decisions, not just the metrics");
  const sc = [["Desire", "Qualified waitlist; hero preference", "Revise hero storytelling or assortment"], ["Trust", "Product-page engagement; objections", "Add proof, change copy, solve product friction"], ["Conversion", "Add-to-cart, checkout, return reasons", "Fix fit, product information, offer or checkout"], ["Economics", "Contribution margin; acquisition payback", "Limit paid scale or redesign cost structure"], ["Loyalty", "Repeat engagement; second-category interest", "Advance the next silhouette or retention program"]];
  s.addTable([[tblHead("Outcome"), tblHead("Signal"), tblHead("If it underperforms")], ...sc.map(([a, b, c]) => [tblText(a, { fontSize: 22, color: HEX.red }), tblText(b, { fontSize: 15 }), tblText(c, { fontSize: 15 })])],
    { x: M, y: 1.95, w: W - 2 * M, colW: [2.6, 4.6, 4.73], rowH: [0.4, 0.75, 0.75, 0.75, 0.75, 0.75], border: [{ type: "none" }, { type: "none" }, tblBorder, { type: "none" }], margin: [0.05, 0.05, 0.05, 0], fill: { color: HEX.ivory } });
  label(s, "Thresholds come from prototype, financial and pilot evidence, not invented history.", M, 6.4, 10, C.accent6, { fontSize: 9 });

  // ================= 29 CONTROL PANEL =================
  s = slide("PANEL", "Governance", "Executive control panel", "Define the decision, then choose the metric");
  [["Hero not preferred", "Compare product-level preference and reasons", "Refine product claims or hero selection"], ["Conversion weak", "Examine product comprehension and checkout", "Improve proof, sizing or checkout"], ["Return rate elevated", "Segment by defect, fit and expectation", "Fix product or merchandising before scaling"], ["CAC exceeds headroom", "Evaluate contribution after acquisition", "Reallocate spend; stop unprofitable scale"], ["Second-category interest", "Track qualified demand for follow-on models", "Authorize a Riviera or Monaco test"]].forEach(([a, b, c], i) => {
    const y = 1.95 + i * 0.88; box(s, M, y, W - 2 * M, 0.78, C.background1);
    serif(s, a, M + 0.25, y + 0.17, 3.6, 0.5, 21, { color: C.accent1 }); body(s, b, 4.6, y + 0.22, 4.0, 0.5, { fontSize: 14 });
    s.addText("→", { x: 8.55, y: y + 0.17, w: 0.4, h: 0.45, fontSize: 20, color: C.accent1, align: "center", margin: 0, isTextBox: true });
    body(s, c, 9.05, y + 0.22, 3.5, 0.5, { fontSize: 14, bold: true });
  });

  // ================= 30 OPERATING MODEL =================
  s = slide("CONTENT", "Governance", "Leadership  /  operating model", "Decision rights and accountability");
  const om = [["Customer and competitive intelligence", "PMM leads synthesis", "Research + Product"], ["Positioning and messaging", "PMM accountable", "Brand + Creative"], ["Product readiness and claims", "Product and Ops approve; PMM sets the buyer bar", "Quality + Legal"], ["Price and financial guardrails", "Finance and GM approve; PMM recommends", "Finance + Product"], ["Integrated launch", "PMM accountable for the plan", "Growth + PR + Ecommerce"], ["Scaling capital", "GM and Finance authorize", "PMM + Growth propose"]];
  s.addTable([[tblHead("Decision area"), tblHead("Who decides"), tblHead("Partners")], ...om.map(([a, b, c]) => [tblText(a, { fontSize: 17, color: HEX.red }), tblText(b, { fontSize: 14 }), tblText(c, { fontFace: SANS, fontSize: 10, color: HEX.muted })])],
    { x: M, y: 1.95, w: W - 2 * M, colW: [4.4, 4.6, 2.93], rowH: [0.4, 0.62, 0.62, 0.62, 0.62, 0.62, 0.62], border: [{ type: "none" }, { type: "none" }, tblBorder, { type: "none" }], margin: [0.05, 0.05, 0.05, 0], fill: { color: HEX.ivory } });
  body(s, "Cadence: weekly launch-readiness reviews; monthly demand and margin decisions; explicit escalation on cost, quality or claims.", M, 6.2, 11.9, 0.45, { fontSize: 14, color: C.accent6 });

  // ================= 31 GATES =================
  s = slide("NAVY", "Governance", "Decision system", "Three executive authorization gates");
  [["Gate 1", "Product", "Approve only after prototype feel, weight, comfort, hardware and target cost meet requirements."], ["Gate 2", "Market", "Release only after competitive product tests and price acceptance support the launch proposition."], ["Gate 3", "Scale", "Expand paid investment only with positive contribution after acquisition and manageable returns."]].forEach(([g, t, d], i) => {
    const x = M + i * 4.05; rule(s, x, 2.15, 3.7, C.accent2);
    label(s, g, x, 2.35, 2, C.accent2); serif(s, t, x, 2.65, 3.7, 0.8, 48, { color: C.background1 }); body(s, d, x, 3.7, 3.55, 1.8, { fontSize: 17, color: C.background1 });
  });

  // ================= 32 OUTCOME =================
  pres.addSection({ title: "Closing" });
  s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Closing" });
  await cover(s, "lineup.jpg", 6.2, 0, W - 6.2, H, 0.5, 0.5);
  label(s, "Strategic outcome", M, 0.9, 5, C.accent2);
  serif(s, [rn("An ambitious brand.\n"), it("A disciplined launch.")], M, 1.3, 5.3, 2.0, 44, { color: C.background1 });
  ["Build recognition through the Portofino.", "Earn the first purchase through the Rally.", "Expand the customer's world through the collection."].forEach((t, i) => serif(s, t, M, 3.6 + i * 0.6, 5.3, 0.55, 21, { color: C.background1 }));
  s.addImage({ path: A("wordmark-ivory.png"), x: M, y: 6.25, w: 2.4, h: 2.4 * wm.height / wm.width, altText: "Resort Sport Ltd." });
  label(s, "Official equipment for time off.", M, 6.85, 5, C.accent2, { fontSize: 9 });

  // ================= 33 APPENDIX =================
  s = slide("CONTENT", "Closing", "Appendix  /  model boundaries", "Working assumptions and handoff");
  box(s, M, 1.95, 5.85, 4.1, C.background2); box(s, M + 6.1, 1.95, 5.85, 4.1, C.background1, { line: { color: HEX.line, width: 0.75 } });
  label(s, "Locked brand standards", M + 0.35, 2.2, 5); label(s, "Open commercial variables", M + 6.45, 2.2, 5, C.accent6);
  ["Product marks: RS only.", "Hardware: brushed silver-tone; aluminum arch pending tests.", "Interior: Sand lining in every bag.", "Permanent: Heritage, Avorio, Cognac, Nero.", "Limited: Curaçao, Negroni, Limoncello, Bijou.", "Service: complimentary US delivery; 30-day returns."].forEach((t, i) => body(s, t, M + 0.35, 2.6 + i * 0.55, 5.2, 0.5, { fontSize: 15 }));
  ["Actual materials cost, fabrication and quality.", "Consumer preference and willingness to pay.", "Channel CAC and conversion performance.", "Packaging, logistics and returns cost.", "Launch timing and feasible inventory."].forEach((t, i) => body(s, t, M + 6.45, 2.6 + i * 0.55, 5.2, 0.5, { fontSize: 15 }));
  body(s, "This is a hypothetical, self-initiated commercial strategy. Assortment, pricing, segments, channels and sequencing are proposed decisions; no consumer testing, sales or manufactured-product performance is claimed.", M, 6.2, 11.9, 0.5, { fontSize: 13, color: C.accent6 });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
