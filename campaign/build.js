// Renders the "Out of office" key visual at exact placement sizes from the text-free crops.
// Type and logo are live layers: Instrument Serif, Space Grotesk, real wordmark in Cabana Red.
const { chromium } = require('playwright');
const path = require('path');
const dir = __dirname;
const FONTS = '<link href="https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Space+Grotesk:wght@500&display=swap" rel="stylesheet">';
const css = (W, H) => `
  *{margin:0;box-sizing:border-box}
  body{width:${W}px;height:${H}px;overflow:hidden}
  .f{position:relative;width:${W}px;height:${H}px;background-size:cover;background-position:var(--pos,center)}
  .t{position:absolute;color:#11293A}
  h1{font-family:'Instrument Serif',Georgia,serif;font-weight:400;line-height:.95;letter-spacing:-.01em}
  h1 em{font-style:italic}
  .wm{display:block}
  .rule{width:44px;height:1px;background:#811E23;opacity:.7}
  .lbl{font-family:'Space Grotesk',Arial,sans-serif;font-weight:500;text-transform:uppercase;color:#11293A}
  .btn{display:inline-block;font-family:'Space Grotesk',Arial,sans-serif;font-weight:500;text-transform:uppercase;background:#811E23;color:#F1E9DD}`;
const block = (o) => `
  <div class="t" style="${o.pos};text-align:${o.align};display:flex;flex-direction:column;align-items:${o.align==='right'?'flex-end':o.align==='center'?'center':'flex-start'};gap:${o.gap}px">
    <h1 style="font-size:${o.h}px">Out of office.<br><em>In our element.</em></h1>
    <img class="wm" src="src/wordmark.png" style="width:${o.wm}px;margin-top:${o.gap*.4}px">
    <div class="rule"></div>
    <div class="lbl" style="font-size:${o.l}px;letter-spacing:${o.l*.3}px">Collection 01</div>
    ${o.btn ? `<div class="btn" style="margin-top:${o.gap}px;font-size:${o.btn}px;letter-spacing:${o.btn*.28}px;padding:${o.btn*1.6}px ${o.btn*3}px">Shop Collection 01</div>` : ''}
  </div>`;
const pieces = [
  // name, W, H, source, background position, wash gradient, text block
  ['kv-square-1080x1080', 1080, 1080, 'kv-square.jpg', 'center', 'linear-gradient(200deg,rgba(241,233,221,.55) 0%,rgba(241,233,221,0) 42%)',
    {pos:'top:72px;right:70px', align:'right', h:92, wm:300, l:14, gap:22}],
  ['kv-story-1080x1920', 1080, 1920, 'kv-vertical.jpg', 'center top', 'linear-gradient(180deg,rgba(241,233,221,.6) 0%,rgba(241,233,221,0) 36%)',
    {pos:'top:150px;left:0;right:0', align:'center', h:118, wm:380, l:17, gap:28}],
  ['kv-feed-1080x1350', 1080, 1350, 'kv-vertical.jpg', 'center 62%', 'linear-gradient(180deg,rgba(241,233,221,.55) 0%,rgba(241,233,221,0) 40%)',
    {pos:'top:60px;left:0;right:0', align:'center', h:100, wm:320, l:15, gap:24}],
  ['kv-wide-1200x628', 1200, 628, 'kv-wide.jpg', 'center', 'linear-gradient(250deg,rgba(241,233,221,.6) 0%,rgba(241,233,221,0) 45%)',
    {pos:'top:56px;right:56px', align:'right', h:64, wm:220, l:11, gap:16}],
  ['ad-feed-1080x1350', 1080, 1350, 'kv-vertical.jpg', 'center 62%', 'linear-gradient(180deg,rgba(241,233,221,.55) 0%,rgba(241,233,221,0) 42%)',
    {pos:'top:54px;left:0;right:0', align:'center', h:96, wm:300, l:14, gap:22, btn:15}],
  ['billboard-1920x823', 1920, 823, 'kv-wide.jpg', 'center', 'linear-gradient(250deg,rgba(241,233,221,.6) 0%,rgba(241,233,221,0) 45%)',
    {pos:'top:84px;right:96px', align:'right', h:108, wm:360, l:16, gap:26}],
];
(async () => {
  const b = await chromium.launch(); const errs = [];
  for (const [name, W, H, src, bp, wash, o] of pieces) {
    const html = `<!doctype html><html><head><meta charset="utf-8">${FONTS}<style>${css(W,H)}</style></head><body>
      <div class="f" style="background-image:${wash},url('src/${src}');--pos:${bp}">${block(o)}</div></body></html>`;
    const f = path.join(dir, `.${name}.html`); require('fs').writeFileSync(f, html);
    const p = await b.newPage({ viewport: { width: W, height: H } }); p.on('pageerror', e => errs.push(e.message));
    await p.goto('file://' + f); await p.waitForLoadState('networkidle'); await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(300);
    await p.screenshot({ path: path.join(dir, name + '.png') }); await p.close(); require('fs').unlinkSync(f);
    console.log(name);
  }
  console.log('errors', errs); await b.close();
})();
