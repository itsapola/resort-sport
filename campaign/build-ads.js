// Funnel ads (Deck 03 message architecture) set on text-free photography.
// Live type: Instrument Serif headlines and body, Space Grotesk labels/prices/buttons, master wordmark.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const dir = __dirname;
const FONTS = '<link href="https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Space+Grotesk:wght@500&display=swap" rel="stylesheet">';
const IV = 'rgba(241,233,221,';
const css = (W, H) => `*{margin:0;box-sizing:border-box}body{width:${W}px;height:${H}px;overflow:hidden}
  .f{position:relative;width:${W}px;height:${H}px;background-size:cover}
  .t{position:absolute;color:#11293A;display:flex;flex-direction:column}
  h1{font-family:'Instrument Serif',Georgia,serif;font-weight:400;line-height:.98;letter-spacing:-.01em}
  h1 em{font-style:italic}
  p{font-family:'Instrument Serif',Georgia,serif;line-height:1.3}
  .lbl,.meta,.btn{font-family:'Space Grotesk',Arial,sans-serif;font-weight:500;text-transform:uppercase}
  .btn{background:#811E23;color:#F1E9DD;display:inline-block}
  .rule{width:40px;height:1px;background:#811E23;opacity:.7}`;
// s = scale unit (px) so one layout works at every size
function block(o) {
  const s = o.s, al = o.align, ai = al === 'right' ? 'flex-end' : al === 'center' ? 'center' : 'flex-start';
  return `<div class="t" style="${o.pos};text-align:${al};align-items:${ai};gap:${s*1.1}px;${o.w?`width:${o.w}px`:''}">
    <img src="src/wordmark.png" style="width:${s*11}px;margin-bottom:${s*.4}px">
    <h1 style="font-size:${s*o.h}px">${o.head}</h1>
    ${o.sub ? `<p style="font-size:${s*1.55}px;max-width:${s*22}px">${o.sub}</p>` : ''}
    ${o.meta ? `<div class="rule"></div><div class="meta" style="font-size:${s*.7}px;letter-spacing:${s*.18}px;line-height:1.9">${o.meta}</div>` : ''}
    <div class="btn" style="margin-top:${s*.5}px;font-size:${s*.66}px;letter-spacing:${s*.19}px;padding:${s*1.05}px ${s*1.9}px">${o.btn}</div>
  </div>`;
}
function band(o) { // ivory strip along the bottom, for full-bleed group shots
  const s = o.s;
  return `<div style="position:absolute;left:0;right:0;bottom:0;height:${o.bh}px;background:#F1E9DD;display:flex;align-items:center;justify-content:space-between;padding:0 ${s*2.4}px;gap:${s*2}px">
    <div style="display:flex;flex-direction:column;gap:${s*.5}px">
      <img src="src/wordmark.png" style="width:${s*9}px">
      <h1 style="font-size:${s*2.6}px;color:#11293A">${o.head}</h1></div>
    <div style="display:flex;flex-direction:column;align-items:flex-end;gap:${s*.8}px">
      <div class="meta" style="font-size:${s*.66}px;letter-spacing:${s*.16}px;line-height:1.9;text-align:right;color:#11293A">${o.meta}</div>
      <div class="btn" style="font-size:${s*.62}px;letter-spacing:${s*.18}px;padding:${s*.95}px ${s*1.7}px">${o.btn}</div></div>
  </div>`;
}
const P = '$695 &nbsp;·&nbsp; 30-day returns';
const ads = [
  // ---- Consideration: "Design worth recognizing."
  ['consider-portofino-1080x1350',1080,1350,'cognac-left.jpg','40% center',`linear-gradient(180deg,${IV}.88) 0%,${IV}.6) 30%,${IV}0) 52%)`,
    {s:40,pos:'top:64px;left:0;right:0',align:'center',h:2.7,head:'Design worth <em>recognizing.</em>',
     meta:'Supple leather · Brushed silver-tone hardware',btn:'Discover the Portofino'}],
  ['consider-zip-1200x628',1200,628,'cognac-tr.jpg','center',`linear-gradient(270deg,${IV}.55) 0%,${IV}0) 55%)`,
    {s:22,pos:'top:0;bottom:0;right:56px;justify-content:center',align:'right',h:2.8,head:'Design worth <em>recognizing.</em>',
     sub:'RS-engraved pulls. Tonal stitching. Nothing extra.',btn:'Discover Collection 01'}],
  ['consider-arch-1200x628',1200,628,'cognac-br.jpg','center',`linear-gradient(270deg,${IV}.55) 0%,${IV}0) 55%)`,
    {s:22,pos:'top:0;bottom:0;right:56px;justify-content:center',align:'right',h:2.8,head:'Design worth <em>recognizing.</em>',
     sub:'A polished aluminum arch, held by hand-finished leather collars.',btn:'Discover the Portofino'}],
  // ---- Decision: price and reassurance
  ['decide-portofino-1080x1350',1080,1350,'decision-left.jpg','center',`linear-gradient(180deg,${IV}.85) 0%,${IV}.55) 28%,${IV}0) 46%)`,
    {s:40,pos:'top:64px;left:0;right:0',align:'center',h:2.9,head:'A very good <em>entrance.</em>',
     meta:'The Portofino in Cognac<br>'+P,btn:'Shop the Portofino'}],
  ['decide-rally-1200x628',1200,628,'decision-tr.jpg','center',`linear-gradient(270deg,${IV}.6) 0%,${IV}0) 55%)`,
    {s:22,pos:'top:0;bottom:0;right:56px;justify-content:center',align:'right',h:2.8,head:'Hands free. <em>Plans optional.</em>',
     meta:'The Rally in Heritage<br>$275 &nbsp;·&nbsp; 30-day returns',btn:'Shop the Rally'}],
  ['decide-collection-1080x1080',1080,1080,'decision-br.jpg','center 30%','none',null,
    {band:true,s:30,bh:230,head:'Collection 01.<br><em>Now available.</em>',meta:'The Portofino $695<br>The Rally from $275',btn:'Shop Collection 01'}],
  // ---- Retargeting
  ['retarget-portofino-1080x1350',1080,1350,'nero-left.jpg','center',`linear-gradient(180deg,${IV}.88) 0%,${IV}.6) 32%,${IV}0) 52%)`,
    {s:40,pos:'top:64px;left:0;right:0',align:'center',h:2.9,head:'Still thinking<br><em>about this one?</em>',
     meta:'The Portofino in Nero<br>'+P,btn:'Take another look'}],
  ['retarget-waiting-1200x628',1200,628,'nero-tr.jpg','center 40%',`linear-gradient(270deg,${IV}.7) 0%,${IV}0) 52%)`,
    {s:22,pos:'top:0;bottom:0;right:56px;justify-content:center',align:'right',h:2.8,head:'Your Portofino<br><em>is waiting.</em>',
     meta:P,btn:'Take another look'}],
  ['retarget-still-here-1200x628',1200,628,'nero-br.jpg','center',`linear-gradient(270deg,${IV}.55) 0%,${IV}0) 55%)`,
    {s:22,pos:'top:0;bottom:0;right:56px;justify-content:center',align:'right',h:2.8,head:'The one you<br>looked at<br><em>is still here.</em>',
     meta:P,btn:'Take another look'}],
];
(async () => {
  const b = await chromium.launch(); const errs = [];
  for (const a of ads) {
    const [name, W, H, src, bp, wash] = a; const o = a[a.length - 1];
    const bg = wash === 'none' ? `url('src/${src}')` : `${wash},url('src/${src}')`;
    const inner = o.band ? `<div style="position:absolute;left:0;right:0;top:0;height:${H-o.bh}px;background:url('src/${src}') ${bp}/cover"></div>` + band(o) : block(o);
    const html = `<!doctype html><html><head><meta charset="utf-8">${FONTS}<style>${css(W,H)}</style></head><body>
      <div class="f" style="${o.band?'background:#F1E9DD':`background-image:${bg};background-position:${bp}`}">${inner}</div></body></html>`;
    const f = path.join(dir, `.${name}.html`); fs.writeFileSync(f, html);
    const p = await b.newPage({ viewport: { width: W, height: H } }); p.on('pageerror', e => errs.push(e.message));
    await p.goto('file://' + f); await p.waitForLoadState('networkidle'); await p.evaluate(() => document.fonts.ready); await p.waitForTimeout(250);
    await p.screenshot({ path: path.join(dir, 'ads', name + '.png') }); await p.close(); fs.unlinkSync(f);
  }
  console.log('errors', errs); await b.close();
})();
