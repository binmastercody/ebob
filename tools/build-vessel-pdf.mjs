/**
 * Binventory 1.0 Vessel Configuration PDF.
 * Same layout as the Binventory 2.0 guide PDF (BinMaster Quick Setup Guide style): cover, PART bands,
 * numbered STEP bands, header, footer, page numbers. Every step of the two lessons is listed and pictured,
 * because the reader may never have used Binventory before.
 *   Part 1: Size and shape   (lesson vessel-shape)
 *   Part 2: Contents         (lesson vessel-contents)
 * Screens are taken from the practice copy (emulator.html), driven by the guide itself.
 *   node tools/build-vessel-pdf.mjs   ->  pdf/Binventory1.0-Vessel-Configuration.pdf
 * Needs Chrome and playwright-core (uses the copy in node_modules here, or the Binventory 2.0 guide's).
 */
import fs from 'fs';
import path from 'path';
import vm from 'vm';
import { createRequire } from 'module';
import { fileURLToPath, pathToFileURL } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
function loadPlaywright() {
  const tries = [path.join(ROOT, 'node_modules'), path.resolve(ROOT, '../BInCloud Emulator/Binventory2.0-Tutorial/node_modules')];
  for (const t of tries) { try { return require(require.resolve('playwright-core', { paths: [t] })); } catch (e) {} }
  throw new Error('playwright-core not found. Run: npm i -D playwright-core');
}
const { chromium } = loadPlaywright();

const OUTDIR = path.join(ROOT, 'pdf'); const SH = path.join(ROOT, 'tools', 'out', 'vessel-shots');
for (const d of [OUTDIR, SH]) fs.mkdirSync(d, { recursive: true });
const OUTFILE = path.join(OUTDIR, 'Binventory1.0-Vessel-Configuration.pdf');

const sb = { window: {}, document: { querySelector() { return null; }, querySelectorAll() { return []; } } };
vm.runInNewContext(fs.readFileSync(path.join(ROOT, 'app/lessons.js'), 'utf8'), sb);
const LS = sb.window.EbobLessons;
const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// step text: keep <b>, drop every other tag
const txt = (html) => esc(String(html || '').replace(/<(?!\/?b>)[^>]+>/g, '')).replace(/&lt;(\/?)b&gt;/g, '<$1b>');

const PARTS = [
  { part: 1, id: 'vessel-shape', title: 'Size and shape', blurb: 'Type/Shape tab: shape, straight wall, width, cone.' },
  { part: 2, id: 'vessel-contents', title: 'Contents', blurb: 'Contents tab: product name and density.' },
];

// ---------- 1. one screenshot per step, cropped around the circled target ----------
const browser = await chromium.launch({ channel: 'chrome', headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.5 });
const shots = {};
for (const P of PARTS) {
  const L = LS.byId(P.id); shots[P.id] = [];
  await page.goto(pathToFileURL(path.join(ROOT, 'emulator.html')).href + '?tutorial=free-mode&lesson=' + P.id);
  await page.addStyleTag({ content: '.eg-coach,.eg-topbar,.eg-fab,.qm-panel,.qm-welcome,#guideBackToMenu{visibility:hidden!important} .eg-ring,.eg-hand{transition:none!important;animation:none!important}' });
  await page.waitForFunction(() => { const s = window.EbobGuide && window.EbobGuide.state(); return s && s.total && (s.phase === 'ready' || s.phase === 'done'); }, null, { timeout: 60000 });
  for (let k = 0; k < L.stepCount; k++) {
    await page.waitForFunction((k) => { const s = window.EbobGuide.state(); return s.idx === k && (s.phase === 'ready' || s.phase === 'done'); }, k, { timeout: 20000 });
    await page.waitForTimeout(700);
    // typing and list steps are pictured with the new value already in, so the reader sees what it should look like
    const st = L.steps[k];
    if ((st.do === 'type' || st.do === 'select') && typeof st.target === 'string') {
      await page.evaluate(({ sel, how, value }) => {
        const e = document.querySelector(sel); if (!e) return;
        if (how === 'type') e.value = value;
        else { for (const o of e.options) if (o.text.trim() === value) e.value = o.value; }
        e.dispatchEvent(new Event('input', { bubbles: true })); e.dispatchEvent(new Event('change', { bubbles: true }));
        try { e.blur(); } catch (x) {}
      }, { sel: st.target, how: st.do, value: st.value });
      await page.waitForTimeout(500);
    }
    const box = await page.evaluate(() => {
      const R = (e) => { if (!e) return null; const r = e.getBoundingClientRect(); return r.width > 2 ? { l: r.left, t: r.top, r: r.right, b: r.bottom } : null; };
      const U = (a, b) => (!a ? b : !b ? a : { l: Math.min(a.l, b.l), t: Math.min(a.t, b.t), r: Math.max(a.r, b.r), b: Math.max(a.b, b.b) });
      const shell = R(document.getElementById('appShell'));
      const ring = R(document.querySelector('.eg-ring')); const hand = R(document.querySelector('.eg-hand.on'));
      // topmost open Binventory window, and an open drop-down menu
      let dlg = null, z = -1;
      document.querySelectorAll('.backdrop.show').forEach((b) => { const zz = parseInt(getComputedStyle(b).zIndex, 10) || 0; const m = b.querySelector('.modal, .modal-dialog, [class*="modal"]'); const r = R(m); if (r && zz >= z) { z = zz; dlg = r; } });
      const menu = R(document.querySelector('#menuStrip .menu-root.open .menu-drop, #menuStrip .menu-root.open ul, #menuStrip .menu-root.open .dropdown'));
      let u = U(U(dlg, ring), U(hand, menu));
      if (!u) u = shell;
      const pad = 18; u = { l: u.l - pad, t: u.t - pad, r: u.r + pad, b: u.b + pad };
      // at least 640 x 400, so a small target still shows where it is in the window
      const minW = 640, minH = 400; const cx = (u.l + u.r) / 2, cy = (u.t + u.b) / 2;
      if (u.r - u.l < minW) { u.l = cx - minW / 2; u.r = cx + minW / 2; }
      if (u.b - u.t < minH) { u.t = cy - minH / 2; u.b = cy + minH / 2; }
      const S = shell || { l: 0, t: 0, r: innerWidth, b: innerHeight };
      const sh = (lo, hi, a, b) => { if (a < lo) { b += lo - a; a = lo; } if (b > hi) { a -= b - hi; b = hi; } return [Math.max(lo, a), Math.min(hi, b)]; };
      [u.l, u.r] = sh(S.l, S.r, u.l, u.r); [u.t, u.b] = sh(S.t, S.b, u.t, u.b);
      return { x: u.l, y: u.t, width: u.r - u.l, height: u.b - u.t };
    });
    const file = path.join(SH, P.id + '-' + String(k + 1).padStart(2, '0') + '.png');
    await page.screenshot({ path: file, clip: { x: Math.round(box.x), y: Math.round(box.y), width: Math.round(box.width), height: Math.round(box.height) } });
    shots[P.id].push(file);
    process.stdout.write('.');
    if (k < L.stepCount - 1) {
      await page.evaluate(() => window.EbobGuide.act());
      await page.waitForFunction((k) => { const s = window.EbobGuide.state(); return s.idx > k || s.finished; }, k, { timeout: 20000 });
    }
  }
}
console.log(' shots done');

// ---------- 2. HTML, laid out like the Binventory 2.0 guide PDF ----------
const img = (f) => pathToFileURL(f).href;
const logo = 'data:image/png;base64,' + fs.readFileSync(path.join(ROOT, 'app/binmaster-logo-white.png')).toString('base64');
const CONE = img(path.join(ROOT, 'assets/vessel-types/02_VerticalCylinderWithCone_.PNG'));
const CSS = `
@page { size: Letter; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif; color: #1a1a1a; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.page { width: 8.5in; height: 11in; position: relative; overflow: hidden; page-break-after: always; break-after: page; }
.page:last-child { page-break-after: auto; break-after: auto; }
.hdr { position: absolute; top: 0; left: 0; right: 0; height: 0.95in; display: flex; align-items: center; padding: 0 0.6in; gap: 14px; }
.logo { display: inline-block; width: 1.75in; height: 0.41in; background: #d6202f; -webkit-mask: url("${logo}") center / contain no-repeat; mask: url("${logo}") center / contain no-repeat; }
.hdr .rule { flex: 1; height: 1px; background: #1a1a1a; margin-top: 10px; }
.hdr .ht { font-size: 21px; margin-top: 2px; color: #1a1a1a; }
.ftr { position: absolute; bottom: 0; left: 0; right: 0; height: 0.55in; display: flex; align-items: center; justify-content: space-between; padding: 0 0.6in; font-size: 9.5px; color: #333; border-top: 1px solid #ddd; }
.body { position: absolute; top: 0.95in; bottom: 0.55in; left: 0; right: 0; }
.blk { padding: 12px 0.6in 14px; }
.band.alt { background: #ececec; }
.band.note { background: #fff7f7; border-top: 2px solid #d6202f; border-bottom: 2px solid #d6202f; }
.part { background: #2b2b2b; color: #fff; display: flex; align-items: baseline; gap: 16px; padding-top: 14px; padding-bottom: 14px; }
.part-no { font-weight: 700; color: #ff7a82; letter-spacing: .08em; font-size: 13px; white-space: nowrap; }
.part-t { font-size: 24px; font-weight: 700; color: #fff; }
.part-b { font-size: 12px; color: #e2e2e2; margin-left: auto; text-align: right; }
.les { display: flex; align-items: baseline; gap: 12px; margin: 2px 0 10px; padding-bottom: 6px; border-bottom: 2px solid #d6202f; }
.les-no { font-size: 22px; font-weight: 700; color: #c41424; }
.les-t { font-size: 22px; font-weight: 700; color: #1a1a1a; }
.les-d { margin-left: auto; font-size: 11.5px; color: #444; }
.step { font-size: 13.5px; margin-bottom: 4px; color: #1a1a1a; }
.step-k { font-weight: 700; }
.step-h { font-weight: 700; } .step-h::before { content: "- "; }
ul.say { margin: 0 0 8px; padding: 0; list-style: none; font-size: 12.5px; line-height: 1.45; color: #1a1a1a; }
ul.say li { padding-left: 14px; text-indent: -14px; margin-bottom: 2px; }
ul.say li::before { content: "> "; color: #555; }
.lt { color: #c41424; font-weight: 700; }
.pics { display: grid; grid-template-columns: 1fr 1fr; gap: 0.25in; align-items: start; }
.pics.n1 { grid-template-columns: 1fr; width: 55%; margin: 0 auto; }
figure { margin: 0; position: relative; }
figure img { width: 100%; display: block; border: 1px solid #bdbdbd; border-radius: 3px; box-shadow: 0 1px 3px rgba(0,0,0,.12); }
figcaption { position: absolute; left: -16px; top: -2px; font-size: 13px; color: #1a1a1a; }
.say1 { font-size: 12.5px; margin: 0 0 8px; line-height: 1.45; }
.measure { display: flex; gap: 0.3in; align-items: center; }
.measure img { width: 1.9in; flex: none; border: 1px solid #d0d0d0; background: #fff; }
table.dims { border-collapse: collapse; font-size: 12px; width: 100%; }
table.dims th { background: #c41424; color: #fff; text-align: left; padding: 5px 8px; }
table.dims td { padding: 5px 8px; border-bottom: 1px solid #e3c9cb; vertical-align: top; }
table.dims tr:nth-child(even) td { background: #fdeeee; }
table.dims td:first-child { font-weight: 700; white-space: nowrap; }
.rank td:first-child { white-space: nowrap; }
.done { text-align: center; padding-top: 16px; padding-bottom: 16px; }
.done-t { font-size: 20px; font-weight: 700; color: #15803d; } .done-b { font-size: 12.5px; margin-top: 4px; color: #1a1a1a; }
.cov-top { background: #e6e6e6; text-align: center; padding: 0.35in 0.6in 0.25in; }
.cov-top .logo { width: 3.3in; height: 0.78in; }
.cov-k { font-size: 22px; letter-spacing: .02em; margin-top: 4px; color: #1a1a1a; }
.cov-t { font-size: 40px; font-weight: 700; margin-top: 12px; color: #1a1a1a; }
.cov-s { font-size: 16px; color: #333; margin-top: 6px; }
.glance { display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; padding: 0.25in 0.6in 0; }
.gl { border-top: 4px solid #d6202f; background: #f4f4f4; padding: 10px 8px; text-align: center; font-size: 12px; font-weight: 700; color: #1a1a1a; }
.gl b { display: block; font-size: 22px; color: #c41424; margin-bottom: 2px; }
.gl.go { border-color: #15803d; } .gl.go b { color: #15803d; }
.need { padding: 0.25in 0.6in 0.1in; font-size: 12.5px; line-height: 1.5; color: #1a1a1a; }
.need h3 { font-size: 14px; margin: 0 0 4px; }
.need ul { margin: 0; padding: 0; list-style: none; columns: 2; } .need li::before { content: "> "; color: #555; }
`;

const blocks = []; let band = 0;
const note = (html) => blocks.push(`<section class="blk band note">${html}</section>`);
const MEASURE = `<div class="step"><span class="step-k">BEFORE YOU START</span> <span class="step-h">What to measure</span></div>
<div class="measure"><img src="${CONE}"><table class="dims"><thead><tr><th>Box in Binventory</th><th>What it is</th></tr></thead><tbody>
<tr><td>Height (H)</td><td>The <b>straight wall</b>: from the top of the wall down to where the cone starts.</td></tr>
<tr><td>Width (W)</td><td>The distance across the bin (the diameter).</td></tr>
<tr><td>Cone Height (CH)</td><td>From where the cone starts down to the opening at the bottom.</td></tr>
<tr><td>Outlet Width (OW)</td><td>The width of the opening at the very bottom of the cone.</td></tr>
</tbody></table></div><p class="say1" style="margin-top:8px">Every number is in <b>feet</b>. Take them from the bin drawings, or measure with a tape. This example is a round bin with a cone bottom (Vertical Cylinder with Cone).</p>`;
const DENSITY = `<div class="step"><span class="step-k">BEFORE YOU START</span> <span class="step-h">What is density, and where do I get it?</span></div>
<p class="say1"><b>Density</b> is how much one cubic foot of your material weighs. Binventory measures how much is in the bin (the volume) and multiplies it by the density to get the <b>weight</b>. A wrong density means a wrong weight.</p>
<table class="dims rank"><thead><tr><th>Where the number comes from</th><th>How good it is</th></tr></thead><tbody>
<tr><td>1. Ask your supplier</td><td><b>Most accurate.</b> Use this whenever you can.</td></tr>
<tr><td>2. Weigh a box of it</td><td><b>Second best.</b> Fill a box 1 ft long, 1 ft wide and 1 ft tall with the material and weigh it. The pounds it weighs is the pounds per cubic foot.</td></tr>
<tr><td>3. Look it up online</td><td><b>Least accurate.</b> Only a starting point; real material varies.</td></tr>
</tbody></table>`;

for (const P of PARTS) {
  const L = LS.byId(P.id);
  blocks.push(`<section class="blk part" data-keep><div class="part-no">PART ${P.part}</div><div class="part-t">${esc(P.title)}</div><div class="part-b">${esc(P.blurb)}</div></section>`);
  note(P.id === 'vessel-shape' ? MEASURE : DENSITY);
  // every step gets a line and its own picture, two steps per band
  for (let i = 0, g = 0; i < L.stepCount; i += 2, g++) {
    const idx = [i, i + 1].filter((x) => x < L.stepCount);
    const lines = idx.map((x, j) => `<li><b>${x + 1}. ${esc(L.steps[x].title)}${/[?:.]$/.test(L.steps[x].title) ? '' : ':'}</b> ${txt(L.steps[x].msg)}${idx.length > 1 ? ` <span class="lt">(${'ab'[j]})</span>` : ''}</li>`).join('');
    const pics = idx.map((x, j) => `<figure>${idx.length > 1 ? `<figcaption>${'ab'[j]}.</figcaption>` : ''}<img src="${img(shots[P.id][x])}"></figure>`).join('');
    const head = g === 0 ? `<div class="les"><span class="les-no">${P.part}.1</span><span class="les-t">${esc(L.title)}</span><span class="les-d">${esc(L.desc)}</span></div>` : '';
    const range = idx.length > 1 ? `Steps ${idx[0] + 1} and ${idx[1] + 1}` : `Step ${idx[0] + 1}`;
    blocks.push(`<section class="blk band ${band++ % 2 ? 'alt' : ''}">${head}<div class="step"><span class="step-k">${range.toUpperCase()}</span></div><ul class="say">${lines}</ul><div class="pics n${idx.length}">${pics}</div></section>`);
  }
  blocks.push(`<section class="blk done"><div class="done-t">${P.part === 1 ? 'Size and shape done' : 'Contents done'}</div><div class="done-b">${esc(L.outro)}</div></section>`);
}

const HDR_T = 'Binventory 1.0 Vessel Configuration';
const FTR_T = 'BV1-VC &nbsp; BINVENTORY 1.0 VESSEL CONFIGURATION';
const cover = `<div class="cov-top"><span class="logo"></span><div class="cov-k">VESSEL CONFIGURATION</div><div class="cov-t">Binventory 1.0</div><div class="cov-s">Change the size, shape and contents of a vessel you already have</div></div>
<div class="glance"><div class="gl"><b>1</b>Open the vessel</div><div class="gl"><b>2</b>Pick the shape</div><div class="gl"><b>3</b>Type the sizes</div><div class="gl"><b>4</b>Set the density</div><div class="gl go"><b>5</b>Save</div></div>
<div class="need"><h3>What you need</h3><ul><li>Straight wall height (feet)</li><li>Width across the bin (feet)</li><li>Cone height (feet)</li><li>Outlet width (feet)</li><li>What is in the bin</li><li>Density from your supplier</li></ul></div>`;
const html = `<!doctype html><html><head><meta charset="utf-8"><title>${HDR_T}</title><style>${CSS}</style></head><body>
<div id="src" style="display:none">${blocks.join('\n')}</div>
<div id="cover" style="display:none">${cover}</div>
<div id="pages"></div>
<script>
window.addEventListener('load', function () {
var HDR = '<div class="hdr"><span class="logo"></span><span class="rule"></span><span class="ht">${HDR_T}</span></div>';
var FTR = function (n) { return '<div class="ftr"><span>${FTR_T}</span><span class="pn">Page ' + n + ' of <span class="pt"></span></span></div>'; };
var pages = document.getElementById('pages');
function mk(inner) { var p = document.createElement('div'); p.className = 'page'; p.innerHTML = inner; pages.appendChild(p); return p; }
var n = 0, cur = null, body = null;
function newPage() { n++; cur = mk(HDR + '<div class="body"></div>' + FTR(n)); body = cur.querySelector('.body'); }
var src = Array.prototype.slice.call(document.getElementById('src').children);
n = 1; cur = mk(document.getElementById('cover').innerHTML + '<div class="body"></div>' + FTR(1)); body = cur.querySelector('.body');
var cv = cur.querySelector('.need'); body.style.top = (cv.offsetTop + cv.offsetHeight) + 'px';
src.forEach(function (b, i) {
  body.appendChild(b);
  var keepNext = b.hasAttribute('data-keep');
  if (body.scrollHeight > body.clientHeight + 1 && body.children.length > 1) { body.removeChild(b); newPage(); body.appendChild(b); }
  if (keepNext && src[i + 1]) { body.appendChild(src[i + 1]); var over = body.scrollHeight > body.clientHeight + 1; body.removeChild(src[i + 1]); if (over && body.children.length > 1) { body.removeChild(b); newPage(); body.appendChild(b); } }
});
var total = pages.children.length;
document.querySelectorAll('.pt').forEach(function (e) { e.textContent = total; });
var overflow = []; Array.prototype.forEach.call(document.querySelectorAll('.page .body'), function (b, i) { if (b.scrollHeight > b.clientHeight + 1) overflow.push(i + 1); });
window.__layout = { total: total, overflow: overflow };
});
</script></body></html>`;

// ---------- 3. PDF (+ page PNGs for review in tools/out/vessel-pages) ----------
const htmlFile = path.join(ROOT, 'tools', 'out', 'vessel-config.html'); fs.writeFileSync(htmlFile, html);
const pp = await browser.newPage({ viewport: { width: 816, height: 1056 }, deviceScaleFactor: 1.25 });
await pp.goto(pathToFileURL(htmlFile).href); await pp.waitForFunction(() => window.__layout); await pp.waitForTimeout(600);
const lay = await pp.evaluate(() => window.__layout);
const PG = path.join(ROOT, 'tools', 'out', 'vessel-pages'); fs.mkdirSync(PG, { recursive: true });
for (const f of fs.readdirSync(PG)) fs.unlinkSync(path.join(PG, f));
const els = await pp.$$('.page');
for (let i = 0; i < els.length; i++) await els[i].screenshot({ path: path.join(PG, 'page-' + String(i + 1).padStart(2, '0') + '.png') });
await pp.emulateMedia({ media: 'print' });
await pp.pdf({ path: OUTFILE, format: 'Letter', printBackground: true, preferCSSPageSize: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
await browser.close();
console.log(path.relative(ROOT, OUTFILE), lay.total + ' pages', Math.round(fs.statSync(OUTFILE).size / 1024) + ' KB', lay.overflow.length ? 'OVERFLOW ' + JSON.stringify(lay.overflow) : '');
