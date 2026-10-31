// Deck 01: Brand Strategy, rebuilt in the Resort Sport Ltd. brand system.
// Copy comes from the strategy deck (with the edits from review). Run: node decks/build-strategy.js
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/b4cf305d-94fc-40e7-bc0e-635802eda6fb_520b0384-9679-424c-8e35-7ac2d1cfffc1/pptx/scripts/apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "Resort_Sport_Deck_01_Brand_Strategy_Designed.pptx");

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
pres.title = "Resort Sport Ltd. Brand Strategy";
pres.author = "Resort Sport Ltd.";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- layouts ----------
const footer = (color) => [
  { text: { text: "RESORT SPORT LTD.  /  BRAND STRATEGY", options: { x: M, y: H - 0.5, w: 6, h: 0.3, fontFace: SANS, fontSize: 9, charSpacing: 3, color, margin: 0 } } },
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

(async () => {
  // ================= 01 COVER =================
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Opening" });
  await cover(s, "cover.jpg", 0, 0, W, H, 0.45, 0.5);
  box(s, 7.35, 0, W - 7.35, H, C.background1, { fill: { color: C.background1, transparency: 12 } });
  const wr = await dims("wordmark-red.png");
  s.addImage({ path: A("wordmark-red.png"), x: 7.95, y: 1.4, w: 4.4, h: 4.4 * wr.height / wr.width, altText: "Resort Sport Ltd." });
  label(s, "Deck 01  /  The brand strategy", 7.95, 2.75, 4.6);
  serif(s, [rn("The art of being "), it("off duty.")], 7.95, 3.15, 4.8, 2.0, 54, { color: C.accent1 });
  label(s, "Hypothetical case study  ·  October 2026", 7.95, 6.55, 5.0, C.accent6);

  // ================= 02 THE TEST =================
  s = slide("CONTENT", "Opening", "The strategic test", "Three things the brand has to earn");
  const tests = [
    ["01", "Differentiation", "A point of view no established leather-goods label already owns."],
    ["02", "Desirability", "Product and a world people want before they need either."],
    ["03", "Credibility", "Proof that an unknown brand's materials and build are worth the price."],
  ];
  tests.forEach(([n, t, d], i) => {
    const x = M + i * 4.05;
    label(s, n, x, 2.55, 3.6, C.accent6);
    serif(s, t, x, 2.9, 3.7, 1.0, 50, { color: C.accent1 });
    rule(s, x, 4.15, 3.6);
    body(s, d, x, 4.35, 3.5, 1.2, { fontSize: 18 });
  });

  // ================= 03 MARKET =================
  pres.addSection({ title: "The market" });
  s = slide("CONTENT", "The market", "Market context", "A large market, growing slowly");
  [["€358B", "2025 personal luxury goods market"], ["2–4%", "Bain 2026 base-case growth forecast"]].forEach(([n, l], i) => {
    const x = M + i * 3.3;
    serif(s, n, x, 2.3, 3.2, 1.3, 80, { color: C.accent1 });
    label(s, l, x, 3.7, 2.9, C.accent6, { h: 0.5 });
  });
  box(s, 7.65, 2.25, 4.98, 3.6, C.background2);
  label(s, "Signal", 8.0, 2.55, 4.3);
  serif(s, "Experience-led luxury interest is relevant context, not category demand evidence.", 8.0, 2.9, 4.3, 1.3, 24);
  label(s, "Interpretation", 8.0, 4.3, 4.3);
  body(s, "An accessory may express an experience-oriented way of living, but this does not establish bag demand, switching or price acceptance.", 8.0, 4.65, 4.3, 1.1, { fontSize: 15 });
  s.addText("Source: Bain & Company / Altagamma luxury market update, June 25, 2026.", { x: M, y: 6.45, w: 8, h: 0.3, fontFace: SANS, fontSize: 9, color: C.accent6, margin: 0, isTextBox: true });

  // ================= 04 CATEGORY TENSION =================
  s = slide("RED", "The market", "The category tension");
  serif(s, "Luxury goods can communicate excellence through heritage and restraint. Not every shopper wants those codes to define their personal style.", M, 1.3, 10.6, 2.9, 40, { color: C.background1 });
  rule(s, M, 4.55, 1.0, C.background1);
  body(s, [rn("Expressive leisure brands demonstrate the power of a cultural worldview. "), rn("The question: ", { italic: true }), rn("can that worldview create preference for a distinctive, premium leather accessory?", { italic: true })],
    M, 4.8, 8.4, 1.4, { fontSize: 20, color: C.background1 });

  // ================= 05 COMPETITION =================
  s = slide("CONTENT", "The market", "Competitive landscape", "Different benchmarks. Distinct competitive jobs.");
  box(s, M, 2.0, 3.9, 4.6, C.background2);
  const bench = [
    ["Design & craft", "Loewe: sculptural, recognizable leather silhouettes. The Row: restrained material expression. Peserico: refined Italian material sensibility."],
    ["Price anchor", "Longchamp full-leather collections: a direct price and product reference, not its nylon lines."],
    ["Cultural benchmarks", "Sporty & Rich: lifestyle identity. Vacation Inc.: an ownable editorial world. Neither replaces a leather handbag."],
  ];
  bench.forEach(([t, d], i) => { label(s, t, M + 0.3, 2.25 + i * 1.47, 3.4); body(s, d, M + 0.3, 2.55 + i * 1.47, 3.35, 1.15, { fontSize: 13.5 }); });
  label(s, "Direct leather-goods competitors  /  the buyer's real alternatives", 5.0, 2.0, 7.6, C.accent6);
  const comp = [["Polène", "Sculptural, signature leather forms", "Product distinctiveness"], ["Mansur Gavriel", "Material-led, understated luxury", "Tactile value and restraint"],
    ["Cuyana", "Everyday versatility and practical utility", "Functional value"], ["DeMellier", "Recognizable metalwork and polished build", "Hardware signature and finish"],
    ["STAUD", "Playful silhouettes and expressive appeal", "Recognizable personality"], ["BY FAR", "Retro-inflected styling and relevance", "Nostalgia and desirability"]];
  comp.forEach(([n, d, t], i) => {
    const x = 5.0 + (i % 3) * 2.58, y = 2.4 + Math.floor(i / 3) * 2.15;
    rule(s, x, y, 2.38, C.text1);
    serif(s, n, x, y + 0.12, 2.38, 0.5, 24, { color: C.accent1 });
    body(s, d, x, y + 0.62, 2.35, 0.75, { fontSize: 14 });
    label(s, "Threat: " + t, x, y + 1.42, 2.38, C.accent6, { h: 0.5, fontSize: 9 });
  });

  // ================= 06 WHITESPACE =================
  s = slide("CONTENT", "The market", "Existing poles", "The whitespace is in the combination, not the components");
  const poles = [["Luxury houses", "Heritage, prestige and recognizable design."], ["Leisure labels", "Cultural expression, often led by apparel or beauty."], ["Premium leather", "Competing alternatives on design, function and price."]];
  poles.forEach(([t, d], i) => { const x = M + i * 4.05; serif(s, t, x, 2.05, 3.7, 0.5, 26, { color: C.accent1 }); body(s, d, x, 2.55, 3.6, 0.7, { fontSize: 15 }); });
  const rows = [["Product design", "Polène, DeMellier", "Recognizable geometry + architectural hardware"], ["Material value", "Mansur Gavriel, Cuyana", "Supple leather and finish standards at an accessible-premium target"],
    ["Personality", "STAUD, BY FAR", "Optimistic resort culture, not conventional fashion nostalgia"], ["Product breadth", "Across the competitors", "A cohesive bag family spanning daily, beach and travel use"]];
  label(s, "Strength", M, 3.55, 3, C.accent6); label(s, "Who delivers it today", 3.75, 3.55, 3, C.accent6); label(s, "Resort Sport unites it as", 7.0, 3.55, 5, C.accent6);
  rows.forEach(([a, b, c], i) => {
    const y = 3.95 + i * 0.6; rule(s, M, y, W - 2 * M, C.accent3);
    serif(s, a, M, y + 0.1, 3, 0.45, 20); body(s, b, 3.75, y + 0.14, 3.1, 0.4, { fontSize: 15, color: C.accent6 }); body(s, c, 7.0, y + 0.14, 5.6, 0.45, { fontSize: 15 });
  });

  // ================= 07 SIGNAL TO PRODUCT =================
  s = slide("PANEL", "The market", "Strategic interpretation", "From signal to product");
  const flow = [["Signal", "Experiences carry meaning", "Luxury demand increasingly reflects a search for personal value beyond visible prestige."],
    ["Hypothesis", "Objects can carry that meaning", "A distinctive accessory may express a way of living, not merely function as a status symbol."],
    ["Product response", "Design for the life, not the logo", "Sculptural bags with intended everyday utility and a recognizable off-duty point of view."]];
  flow.forEach(([l, t, d], i) => {
    const x = M + i * 4.12;
    box(s, x, 2.1, 3.6, 2.9, C.background1);
    label(s, l, x + 0.3, 2.35, 3.0);
    serif(s, t, x + 0.3, 2.7, 3.0, 1.0, 28, { color: C.accent1 });
    body(s, d, x + 0.3, 3.75, 3.0, 1.1, { fontSize: 15 });
    if (i < 2) s.addText("→", { x: x + 3.62, y: 3.2, w: 0.48, h: 0.6, fontSize: 28, color: C.accent1, align: "center", margin: 0, isTextBox: true });
  });
  label(s, "Key test", M, 5.4, 3);
  serif(s, [it("Does this promise create preference over a known leather-goods label at a comparable price?")], M, 5.7, 11.5, 0.6, 24, { color: C.accent1 });

  // ================= 08 CUSTOMER TENSION =================
  pres.addSection({ title: "The customer" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The customer" });
  await cover(s, "cafe.jpg", 0, 0, 5.6, H, 0.55, 0.5);
  label(s, "The customer tension", 6.3, 0.55, 6);
  serif(s, [it("“I want the substance of luxury without having to perform the seriousness of luxury.”")], 6.3, 0.95, 6.3, 2.6, 38, { color: C.accent1 });
  [["Desires", "Design authorship and tactile quality."], ["Resists", "Generic silhouettes or conspicuous status codes."], ["Needs to believe", "Unknown-brand material claims, comfort and durability must be proven before purchase."]].forEach(([l, d], i) => {
    const y = 3.95 + i * 0.85; rule(s, 6.3, y, 6.3, C.accent3);
    label(s, l, 6.3, y + 0.17, 1.9); body(s, d, 8.3, y + 0.12, 4.3, 0.7, { fontSize: 16 });
  });

  // ================= 09 WHO WE BUILD FOR =================
  s = slide("CONTENT", "The customer", "Target buyer", "Who we build for");
  box(s, M, 2.0, 5.2, 4.6, C.text1);
  label(s, "Primary buyer", M + 0.35, 2.3, 4.5, C.accent2);
  serif(s, "Leather-accessory shoppers comparing established premium and luxury options, drawn to signature design over logos.", M + 0.35, 2.65, 4.5, 1.9, 24, { color: C.background1 });
  label(s, "Triggers", M + 0.35, 4.75, 4.5, C.accent2); body(s, "Style refresh, travel, everyday replacement.", M + 0.35, 5.03, 4.5, 0.4, { fontSize: 15, color: C.background1 });
  label(s, "Barriers", M + 0.35, 5.55, 4.5, C.accent2); body(s, "Unknown materials, durability, fit and price-value.", M + 0.35, 5.83, 4.5, 0.4, { fontSize: 15, color: C.background1 });
  label(s, "Jobs to be done", 6.3, 2.0, 6, C.accent6);
  [["Functional", "Carry essentials across daily and leisure use."], ["Emotional", "Feel prepared and spontaneous."], ["Social", "Express individual taste without logos."], ["Switching", "Better design must outweigh familiar-brand trust."]].forEach(([t, d], i) => {
    const x = 6.3 + (i % 2) * 3.2, y = 2.4 + Math.floor(i / 2) * 2.15;
    box(s, x, y, 3.0, 1.95, C.background2);
    serif(s, t, x + 0.25, y + 0.2, 2.6, 0.5, 26, { color: C.accent1 }); body(s, d, x + 0.25, y + 0.8, 2.55, 1.0, { fontSize: 15 });
  });

  // ================= 10 SWITCHING =================
  s = slide("CONTENT", "The customer", "The switching argument", "What earns a switch from a trusted label");
  const sw = [["Customer wants", "Distinctive leather design, tactile quality and versatility across workaday and off-duty settings."],
    ["Current trade-off", "Established options may deliver trusted quality, prestige or utility, but not combined with this design and leisure identity."],
    ["Why consider Resort Sport", "Strong silhouette signatures, intentionally useful forms and a consistent, expressive brand world."],
    ["Why hesitate", "Unfamiliar brand; uncertain leather quality, practical comfort, durability, servicing and value."]];
  sw.forEach(([l, d], i) => { const y = 2.0 + i * 0.82; rule(s, M, y, W - 2 * M, C.accent3); label(s, l, M, y + 0.2, 3.3, C.accent6); body(s, d, 4.2, y + 0.14, 8.4, 0.65, { fontSize: 16 }); });
  box(s, M, 5.45, W - 2 * M, 1.1, C.accent1);
  label(s, "What earns the switch", M + 0.3, 5.68, 3.2, C.background1);
  serif(s, "Physical product proof, transparent specs, fit demonstrations, credible policies and direct value comparison.", 4.2, 5.62, 8.2, 0.8, 22, { color: C.background1 });

  // ================= 11 POSITIONING =================
  pres.addSection({ title: "The brand" });
  s = slide("NAVY", "The brand", "Positioning statement");
  [["For", "design-conscious consumers seeking luxury design with more freedom and personality,"],
    ["Resort Sport Ltd. is", "an accessible-premium leather accessories brand inspired by the art of being off duty."],
    ["Unlike", "brands built on prestige codes or utility-first propositions, it combines identifiable form, intended utility and a leisure-led cultural identity."]].forEach(([l, d], i) => {
    const y = 1.2 + i * 1.75; label(s, l, M, y + 0.12, 2.8, C.accent2);
    serif(s, d, 3.6, y, 9.0, 1.6, 30, { color: C.background1 });
  });

  // ================= 12 PILLARS =================
  s = slide("CONTENT", "The brand", "Brand pillars", "One idea, three strategic pillars");
  const pil = [["01", "Design worth recognizing", "Signature geometry and thoughtful detailing that earn attention without overt logos.", "Portofino handle, Riviera scoop, Regatta sweep.", "handle.jpg", 0.5, 0.45],
    ["02", "Considered for real life", "Function, construction and category-appropriate materials that justify use beyond the photograph.", "Fit, weight, finish and wear to validate.", "lining.jpg", 0.5, 0.5],
    ["03", "A life off the clock", "Optimistic leisure storytelling and a culturally rich editorial identity: a heritage of leisure, not of aristocracy.", "", "carried.jpg", 0.7, 0.5]];
  for (const [i, [n, t, d, p, img, fx, fy]] of pil.entries()) {
    const x = M + i * 4.05;
    await cover(s, img, x, 1.95, 3.75, 1.85, fx, fy);
    label(s, n, x, 4.0, 1, C.accent6);
    serif(s, t, x, 4.28, 3.75, 0.55, 26, { color: C.accent1 });
    body(s, d, x, 4.85, 3.7, 1.05, { fontSize: 14.5 });
    if (p) label(s, p, x, 6.0, 3.7, C.accent6, { fontSize: 9, h: 0.4 });
  }

  // ================= 13 PLATFORM =================
  s = slide("RED", "The brand", "The brand platform in a sentence");
  serif(s, [rn("The art of being "), it("off duty.")], M, 1.25, 11.5, 1.6, 84, { color: C.background1 });
  [["Customer promise", "Distinctive, considered accessories that bring the freedom and optimism of leisure into everyday life."],
    ["Tagline", "Official equipment for time off."], ["Creative provocation", "Luxury doesn't have to be serious to be considered."]].forEach(([l, d], i) => {
    const x = M + i * 4.05; rule(s, x, 4.2, 3.6, C.background1);
    label(s, l, x, 4.4, 3.6, C.background1); serif(s, d, x, 4.8, 3.6, 1.6, i ? 26 : 20, { color: C.background1, italic: i > 0 });
  });

  // ================= 14 STRATEGY BECOMES PRODUCT =================
  pres.addSection({ title: "The product" });
  s = slide("PANEL", "The product", "Product principles", "Strategy becomes product design");
  const pp = [["No oversized logos", "Let silhouettes, hardware and materials build recognition.", "emboss.jpg", "cover"],
    ["Signature handle on Portofino", "A focal design code, not a repeating ornament.", "portofino-cognac.png", "contain"],
    ["Heritage + tonal leathers", "Continuity, utility and seasonless merchandising.", "rally-heritage.png", "contain"],
    ["Riviera wide scoop", "An ownable carryall shape for resort leisure.", "riviera-heritage.png", "contain"],
    ["Regatta zip + laptop fit", "Sculptural design that works in daily use.", "regatta-cognac.png", "contain"]];
  for (const [i, [t, d, img, mode]] of pp.entries()) {
    const x = M + i * 2.42, w = 2.22;
    box(s, x, 2.0, w, 2.2, C.background1);
    if (mode === "cover") await cover(s, img, x, 2.0, w, 2.2); else await contain(s, img, x + 0.15, 2.15, w - 0.3, 1.9);
    serif(s, t, x, 4.4, w, 0.85, 22, { color: C.accent1 }); body(s, d, x, 5.3, w, 1.0, { fontSize: 14 });
  }

  // ================= 15 FIVE SILHOUETTES =================
  s = slide("CONTENT", "The product", "Collection 01", "Five silhouettes, one philosophy");
  const five = [["RS-01", "The Monaco", "Travel duffel", "monaco-heritage.png"], ["RS-02", "The Rally", "Belt bag", "rally-heritage.png"], ["RS-03", "The Riviera", "Oversized carryall", "riviera-heritage.png"],
    ["RS-04", "The Portofino", "Handbag", "portofino-heritage.png"], ["RS-05", "The Regatta", "Backpack", "regatta-heritage.png"]];
  for (const [i, [c, n, t, img]] of five.entries()) {
    const x = M + i * 2.42, w = 2.22;
    box(s, x, 2.0, w, 2.9, C.background2);
    await contain(s, img, x + 0.2, 2.3, w - 0.4, 2.3);
    label(s, c, x, 5.1, w, C.accent6); serif(s, n, x, 5.4, w, 0.5, 26, { color: C.accent1 }); label(s, t, x, 5.95, w, C.text1);
  }
  label(s, "Shown in Heritage: natural canvas and cognac leather", M, 6.55, 8, C.accent6, { fontSize: 9 });

  // ================= 16 PRICE CONTEXT =================
  s = slide("CONTENT", "The product", "Competitive price context", "Priced against full leather, not discount luxury");
  label(s, "Longchamp  /  Le Pliage Xtra leather, US examples", M, 2.0, 8, C.accent6);
  ["$425", "$570", "$700", "$780"].forEach((p, i) => serif(s, p, M + i * 2.2, 2.35, 2.1, 1.0, 54, { color: C.text1 }));
  rule(s, M, 3.55, W - 2 * M, C.accent3);
  label(s, "Resort Sport Ltd.  /  Collection 01, proposed bands", M, 3.8, 8);
  [["Rally", "$275–350"], ["Portofino", "$495–695"], ["Regatta", "$525–725"], ["Riviera", "$595–795"], ["Monaco", "$695–895"]].forEach(([n, p], i) => {
    const x = M + i * 2.42; serif(s, p, x, 4.15, 2.3, 0.6, 30, { color: C.accent1 }); label(s, n, x, 4.8, 2.2, C.text1);
  });
  box(s, M, 5.4, W - 2 * M, 0.85, C.background2);
  label(s, "Strategic use", M + 0.3, 5.68, 2.2);
  body(s, "A directional retail price reference: not a ceiling, a quality-equivalence claim or a validated margin model.", 3.0, 5.6, 9.4, 0.6, { fontSize: 15 });
  s.addText("Source: Longchamp USA, Le Pliage Xtra leather collection, accessed October 30, 2026. Prices change.", { x: M, y: 6.45, w: 9, h: 0.3, fontFace: SANS, fontSize: 9, color: C.accent6, margin: 0, isTextBox: true });

  // ================= 17 WILL / WON'T =================
  pres.addSection({ title: "The choices" });
  s = slide("CONTENT", "The choices", "Positioning matters only when it constrains decisions", "What we will, and will not, compete on");
  box(s, M, 2.0, 5.85, 4.55, C.text1); box(s, M + 6.1, 2.0, 5.85, 4.55, C.background2);
  label(s, "We will compete on", M + 0.35, 2.3, 5, C.accent2); label(s, "We will not chase", M + 6.45, 2.3, 5, C.accent1);
  const will = ["Recognizable sculptural design", "Material and finish credibility", "Thoughtful daily and leisure versatility", "A distinctive editorial worldview", "A disciplined, coherent assortment"];
  const wont = ["Oversized status logos", "The lowest ticket price", "Feature volume for its own sake", "Short-lived seasonal trend dependence", "Unlimited SKUs at launch"];
  will.forEach((t, i) => serif(s, t, M + 0.35, 2.75 + i * 0.7, 5.2, 0.6, 24, { color: C.background1 }));
  wont.forEach((t, i) => serif(s, t, M + 6.45, 2.75 + i * 0.7, 5.2, 0.6, 24, { color: C.text1 }));

  // ================= 18 PRODUCT / BRAND =================
  s = slide("CONTENT", "The choices", "Sources of advantage", "Two advantages, both still to be proven");
  box(s, M, 2.0, 2.6, 4.55, C.background2); await contain(s, "portofino-cognac.png", M + 0.2, 2.3, 2.2, 3.9);
  label(s, "Product", 3.6, 2.0, 2.8); body(s, "Distinctive geometry offers a visible reason to notice and evaluate the products. Memorability, preference and price justification must be tested.", 3.6, 2.35, 2.75, 3.2, { fontSize: 16 });
  await cover(s, "rally-worn.jpg", 6.75, 2.0, 2.6, 4.55, 0.5, 0.5);
  label(s, "Brand", 9.65, 2.0, 2.9); body(s, "A repeatable leisure-led worldview can connect product, copy, campaign and commerce without depending on seasonal trend cues.", 9.65, 2.35, 2.95, 3.2, { fontSize: 16 });

  // ================= 19 REJECTED =================
  s = slide("CONTENT", "The choices", "Design decisions", "What we rejected, and why");
  const rej = [["Generic barrel duffel", "old-duffel.jpg", "monaco-heritage.png", "A lower, aerodynamic Monaco creates stronger distinction while keeping travel utility."],
    ["Conventional beach tote", "old-riviera.jpg", "riviera-heritage.png", "The extra-wide, scooped Riviera is more expressive and ownable."],
    ["Drawstring heritage backpack", "old-regatta.jpg", "regatta-heritage.png", "Regatta trades flap and drawstring cues for zip access and a laptop compartment."],
    ["Logo-heavy resort graphics", null, "emboss.jpg", "A small blind-debossed RS keeps the focus on the silhouettes."]];
  for (const [i, [r, old, nw, d]] of rej.entries()) {
    const x = M + i * 3.03, w = 2.8;
    label(s, "Rejected", x, 2.0, w, C.accent6);
    if (old) await cover(s, old, x, 2.3, 1.25, 1.25, 0.5, 0.5, { transparency: 35 });
    else { box(s, x, 2.3, 1.25, 1.25, C.background2); s.addText("RESORT\nSPORT\nRESORT", { x, y: 2.3, w: 1.25, h: 1.25, fontFace: SANS, fontSize: 12, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, transparency: 40 }); }
    body(s, r, x + 1.4, 2.32, 1.4, 1.2, { fontSize: 14, color: C.accent6, strike: "sngStrike" });
    label(s, "Chosen", x, 3.85, w);
    box(s, x, 4.15, w, 1.5, C.background2);
    if (nw === "emboss.jpg") await cover(s, nw, x, 4.15, w, 1.5); else await contain(s, nw, x + 0.15, 4.25, w - 0.3, 1.3);
    body(s, d, x, 5.8, w, 0.85, { fontSize: 14 });
  }

  // ================= 20 THREE DECISIONS =================
  s = slide("PANEL", "The choices", "What would change the decision", "Three decisions that determine success");
  [["01", "Desirability", "Do the designs create preference?", "Compare Portofino and Riviera against established sculptural bags. If preference is weak, refine product geometry before expanding the range."],
    ["02", "Value", "Will the target buyer trust the proposition?", "Test material perception, price expectations and credibility signals. If price resistance dominates, reconsider specification or assortment, not just messaging."],
    ["03", "Delivery", "Can the intended product promise be built?", "Validate supplier costs, hardware, abrasion resistance and carry comfort. If quality at target prices is infeasible, prioritize fewer hero silhouettes."]].forEach(([n, t, q, d], i) => {
    const x = M + i * 4.05;
    box(s, x, 2.0, 3.75, 4.5, C.background1);
    serif(s, n, x + 0.3, 2.2, 1.2, 0.9, 54, { color: C.accent3 });
    label(s, t, x + 0.3, 3.15, 3.2);
    serif(s, q, x + 0.3, 3.45, 3.2, 1.0, 24, { color: C.accent1 });
    body(s, d, x + 0.3, 4.5, 3.2, 1.9, { fontSize: 14 });
  });

  // ================= 21 CHOICES =================
  s = slide("NAVY", "The choices", "An intentional brand platform", "The strategy makes choices");
  [["Positioning", "Design-led leisure accessories, not logo-led status goods."], ["Customer", "Buyers seeking distinctive design, tactile quality and versatile carry."],
    ["Portfolio", "Five complementary silhouettes; hero signatures over assortment breadth."], ["Price role", "Accessible-premium, anchored to full-leather competitors, not discount luxury."],
    ["Expression", "An optimistic resort editorial world with quiet product branding."], ["Growth", "Permanent, recognizable products first; limited color issues later."]].forEach(([t, d], i) => {
    const x = M + (i % 3) * 4.05, y = 2.1 + Math.floor(i / 3) * 2.25;
    rule(s, x, y, 3.7, C.accent2);
    label(s, t, x, y + 0.2, 3.6, C.accent2);
    serif(s, d, x, y + 0.55, 3.65, 1.5, 22, { color: C.background1 });
  });

  // ================= 22 SOURCES =================
  s = slide("CONTENT", "The choices", "Sources and methodology", "Sources and scope");
  const src = [["Research", "Bain & Company / Altagamma, June 25, 2026 luxury market outlook (market slide)."],
    ["Price benchmark", "Longchamp US, Le Pliage Xtra full-leather collection (US retail)."],
    ["Direct brands", "polene-paris.com  ·  mansurgavriel.com  ·  cuyana.com  ·  demellierlondon.com  ·  staud.clothing  ·  byfar.com"],
    ["Quality and design reference", "loewe.com  ·  therow.com  ·  peserico.com"],
    ["Method", "Comparisons reflect publicly presented product categories and brand positioning; assessment is qualitative."],
    ["Scope", "Customer, market and product statements are hypotheses for validation, not research findings. Product images are conceptual renders. No equivalence of materials, build quality or consumer preference is asserted."]];
  src.forEach(([l, d], i) => { const y = 2.0 + i * 0.72; rule(s, M, y, W - 2 * M, C.accent3); label(s, l, M, y + 0.18, 3.2, C.accent6); body(s, d, 4.2, y + 0.12, 8.4, 0.6, { fontSize: 14 }); });

  // ================= 23 OUTCOME =================
  pres.addSection({ title: "Closing" });
  s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Closing" });
  await cover(s, "closing.jpg", 6.2, 0, W - 6.2, H, 0.5, 0.55);
  label(s, "The strategic outcome", M, 0.9, 5, C.accent2);
  serif(s, [rn("A different expression of "), it("luxury.")], M, 1.3, 5.2, 2.0, 48, { color: C.background1 });
  body(s, "Resort Sport Ltd. brings the discipline of luxury design to the freedom of leisure. A deliberate position. Five complementary silhouettes. A cohesive creative and product language.", M, 3.75, 4.9, 1.5, { fontSize: 17, color: C.background1 });
  serif(s, [it("Luxury doesn't have to be serious to be considered.")], M, 5.3, 5.1, 0.9, 24, { color: C.accent2 });
  const wm = await dims("wordmark-ivory.png");
  s.addImage({ path: A("wordmark-ivory.png"), x: M, y: 6.45, w: 2.4, h: 2.4 * wm.height / wm.width, altText: "Resort Sport Ltd." });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
