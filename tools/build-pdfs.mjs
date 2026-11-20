/**
 * PDF guides for every lesson, plus the Binventory 1.0 Vessel Configuration book.
 * All in the same layout as the Binventory 2.0 guide PDF (BinMaster Quick Setup Guide style):
 * cover, PART bands, numbered STEP bands with a picture of every step, header, footer, page numbers.
 *
 *   node tools/build-pdfs.mjs                 every lesson -> pdf/lessons/<id>.pdf, and the vessel book
 *   node tools/build-pdfs.mjs add-vessel tour only those lessons (and the vessel book if a vessel lesson is named)
 *   node tools/build-pdfs.mjs --book          only the vessel book
 *
 * Screens come from the practice copy itself: each lesson page is opened with ?pdf=1 and
 * app/pdf-capture.js steps through it, asking this script for a real screenshot of each step.
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

const SITE = 'https://binmastercody.github.io/ebob/';
const OUT = path.join(ROOT, 'tools', 'out');
const SH = path.join(OUT, 'shots'), PAGES = path.join(OUT, 'pages');
const LESSON_DIR = path.join(ROOT, 'pdf', 'lessons');
for (const d of [SH, PAGES, LESSON_DIR]) fs.mkdirSync(d, { recursive: true });

const sb = { window: {}, document: { querySelector() { return null; }, querySelectorAll() { return []; } } };
vm.runInNewContext(fs.readFileSync(path.join(ROOT, 'app/lessons.js'), 'utf8'), sb);
const LS = sb.window.EbobLessons;
const SECTION = Object.fromEntries(LS.sections.map((s) => [s.id, s.title]));

const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// step text: keep <b>, drop every other tag
// entities in the lesson text (&gt; for "Next >") are decoded first, so they are not escaped twice
const unent = (s) => s.replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const txt = (html) => esc(unent(String(html || '').replace(/<(?!\/?b>)[^>]+>/g, ''))).replace(/&lt;(\/?)b&gt;/g, '<$1b>');
// doc lessons: keep simple text markup, no attributes
const DOC_TAGS = new Set(['p', 'ul', 'ol', 'li', 'b', 'strong', 'em', 'i', 'br', 'dl', 'dt', 'dd', 'h2', 'h3', 'h4', 'table', 'thead', 'tbody', 'tr', 'th', 'td', 'code']);
const docHtml = (html) => String(html || '').replace(/<\/?([a-z0-9]+)[^>]*>/gi, (m, t) => {
  t = t.toLowerCase(); if (!DOC_TAGS.has(t)) return '';
  return m.startsWith('</') ? `</${t}>` : `<${t}>`;
});
const hrefOf = (l) => l.page === 'installer' ? 'installer.html?lesson=' + l.id
  : l.kind === 'doc' ? 'doc.html?lesson=' + l.id
  : l.scenario ? 'emulator.html?tutorial=' + l.scenario + '&fresh=1'
  : 'emulator.html?tutorial=free-mode&lesson=' + l.id;

// ---------- what to build ----------
const args = process.argv.slice(2);
const bookOnly = args.includes('--book');
const named = args.filter((a) => !a.startsWith('--'));
for (const id of named) if (!LS.byId(id)) throw new Error('no lesson ' + id);
const lessons = bookOnly ? [] : named.length ? named.map((id) => LS.byId(id)) : LS.lessons.slice();
const BOOK_IDS = ['vessel-shape', 'vessel-contents'];
const buildBook = bookOnly || !named.length || named.some((id) => BOOK_IDS.includes(id));
const needShots = new Set(lessons.filter((l) => l.kind !== 'doc').map((l) => l.id));
if (buildBook) BOOK_IDS.forEach((id) => needShots.add(id));

// ---------- 1. screenshots: every step of every lesson ----------
const browser = await chromium.launch({ channel: 'chrome', headless: true });
const shots = {};
const HIDE = '.eg-coach,.eg-topbar,.eg-fab,.qm-panel,.qm-welcome,#guideBackToMenu,#ebobTutorialTooltip,.eg-share-btn{visibility:hidden!important}' +
  '.eg-ring,.eg-hand,#ebobTutorialCircle{transition:none!important;animation:none!important}';
async function capture(id) {
  const l = LS.byId(id);
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.5 });
  let n = 0;
  await page.exposeFunction('__ebobPdfShot', async (box) => {
    const file = path.join(SH, id + '-' + String(++n).padStart(2, '0') + '.jpg');
    await page.screenshot({ path: file, type: 'jpeg', quality: 85, clip: box });
    return file;
  });
  page.on('pageerror', (e) => console.log('  [' + id + '] page error: ' + e.message));
  const u = hrefOf(l);
  await page.goto(pathToFileURL(path.join(ROOT, u.split('?')[0])).href + '?' + u.split('?')[1] + '&pdf=1');
  await page.addStyleTag({ content: HIDE });
  await page.waitForFunction(() => window.EbobPdfCapture, null, { timeout: 30000 });
  const steps = await page.evaluate(() => window.EbobPdfCapture.driveSelf());
  await page.close();
  return steps;
}
const queue = [...needShots]; let failed = [];
async function worker() {
  while (queue.length) {
    const id = queue.shift();
    const t0 = Date.now();
    for (let tryNo = 1; tryNo <= 2; tryNo++) {
      try { shots[id] = await capture(id); console.log('shots ' + id + ': ' + shots[id].length + ' steps, ' + Math.round((Date.now() - t0) / 1000) + ' s'); break; }
      catch (e) { console.log('  [' + id + '] try ' + tryNo + ' failed: ' + (e.message || e).split('\n')[0]); if (tryNo === 2) failed.push(id); }
    }
  }
}
await Promise.all([worker(), worker(), worker()]);

// ---------- 2. layout, the same as the Binventory 2.0 guide PDF ----------
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
.logo { display: inline-block; width: 1.75in; height: 0.41in; background: #d6202f; -webkit-mask: url("${logo}") center / contain no-repeat; mask: url("${logo}") center / contain no-repeat; flex: none; }
.hdr .rule { flex: 1; height: 1px; background: #1a1a1a; margin-top: 10px; min-width: 0.4in; }
.hdr .ht { font-size: 18px; margin-top: 2px; color: #1a1a1a; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 4.6in; }
.ftr { position: absolute; bottom: 0; left: 0; right: 0; height: 0.55in; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 0 0.6in; font-size: 9.5px; color: #333; border-top: 1px solid #ddd; }
.ftr .pn { white-space: nowrap; }
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
.les-d { margin-left: auto; font-size: 11.5px; color: #444; text-align: right; }
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
.doc { font-size: 12.5px; line-height: 1.5; }
.doc p { margin: 0 0 6px; } .doc ul, .doc ol { margin: 0 0 6px; padding-left: 20px; } .doc li { margin-bottom: 3px; }
.doc h2, .doc h3, .doc h4 { font-size: 13.5px; margin: 6px 0 4px; }
.doc dl { display: grid; grid-template-columns: auto 1fr; gap: 3px 14px; margin: 0 0 6px; } .doc dt { font-weight: 700; } .doc dd { margin: 0; }
.measure { display: flex; gap: 0.3in; align-items: center; }
.measure img { width: 1.9in; flex: none; border: 1px solid #d0d0d0; background: #fff; }
table.dims { border-collapse: collapse; font-size: 12px; width: 100%; }
table.dims th { background: #c41424; color: #fff; text-align: left; padding: 5px 8px; }
table.dims td { padding: 5px 8px; border-bottom: 1px solid #e3c9cb; vertical-align: top; }
table.dims tr:nth-child(even) td { background: #fdeeee; }
table.dims td:first-child { font-weight: 700; white-space: nowrap; }
.done { text-align: center; padding-top: 16px; padding-bottom: 16px; }
.done-t { font-size: 20px; font-weight: 700; color: #15803d; } .done-b { font-size: 12.5px; margin-top: 4px; color: #1a1a1a; }
.cov-top { background: #e6e6e6; text-align: center; padding: 0.35in 0.6in 0.25in; }
.cov-top .logo { width: 3.3in; height: 0.78in; }
.cov-k { font-size: 20px; letter-spacing: .04em; margin-top: 4px; color: #1a1a1a; text-transform: uppercase; }
.cov-t { font-size: 36px; font-weight: 700; margin-top: 12px; color: #1a1a1a; line-height: 1.15; }
.cov-s { font-size: 16px; color: #333; margin-top: 6px; }
.glance { display: grid; grid-template-columns: repeat(var(--n), 1fr); gap: 10px; padding: 0.25in 0.6in 0; }
.gl { border-top: 4px solid #d6202f; background: #f4f4f4; padding: 10px 8px; text-align: center; font-size: 12px; font-weight: 700; color: #1a1a1a; }
.gl b { display: block; font-size: 22px; color: #c41424; margin-bottom: 2px; }
.gl.go { border-color: #15803d; } .gl.go b { color: #15803d; }
.need { padding: 0.25in 0.6in 0.1in; font-size: 12.5px; line-height: 1.5; color: #1a1a1a; }
.need h3 { font-size: 14px; margin: 0 0 4px; }
.need ul { margin: 0; padding: 0; list-style: none; } .need ul.two { columns: 2; } .need li::before { content: "> "; color: #555; }
.need a { color: #c41424; }
`;

// notes shown before the steps of a lesson (also in the vessel book)
const NOTES = {
  'vessel-shape': `<div class="step"><span class="step-k">BEFORE YOU START</span> <span class="step-h">What to measure</span></div>
<div class="measure"><img src="${CONE}"><table class="dims"><thead><tr><th>Box in Binventory</th><th>What it is</th></tr></thead><tbody>
<tr><td>Height (H)</td><td>The <b>straight wall</b>: from the top of the wall down to where the cone starts.</td></tr>
<tr><td>Width (W)</td><td>The distance across the bin (the diameter).</td></tr>
<tr><td>Cone Height (CH)</td><td>From where the cone starts down to the opening at the bottom.</td></tr>
<tr><td>Outlet Width (OW)</td><td>The width of the opening at the very bottom of the cone.</td></tr>
</tbody></table></div><p class="say1" style="margin-top:8px">Every number is in <b>feet</b>. Take them from the bin drawings, or measure with a tape. This example is a round bin with a cone bottom (Vertical Cylinder with Cone).</p>`,
  'vessel-contents': `<div class="step"><span class="step-k">BEFORE YOU START</span> <span class="step-h">What is density, and where do I get it?</span></div>
<p class="say1"><b>Density</b> is how much one cubic foot of your material weighs. Binventory measures how much is in the bin (the volume) and multiplies it by the density to get the <b>weight</b>. A wrong density means a wrong weight.</p>
<table class="dims"><thead><tr><th>Where the number comes from</th><th>How good it is</th></tr></thead><tbody>
<tr><td>1. Ask your supplier</td><td><b>Most accurate.</b> Use this whenever you can.</td></tr>
<tr><td>2. Weigh a box of it</td><td><b>Second best.</b> Fill a box 1 ft long, 1 ft wide and 1 ft tall with the material and weigh it. The pounds it weighs is the pounds per cubic foot.</td></tr>
<tr><td>3. Look it up online</td><td><b>Least accurate.</b> Only a starting point; real material varies.</td></tr>
</tbody></table>`,
};

// the blocks for one lesson: optional note, then every step with its picture, two steps per band
function lessonBlocks(l, num, state) {
  const blocks = [];
  if (NOTES[l.id]) blocks.push(`<section class="blk band note">${NOTES[l.id]}</section>`);
  const head = `<div class="les">${num ? `<span class="les-no">${num}</span>` : ''}<span class="les-t">${esc(l.title)}</span><span class="les-d">${esc(l.desc)}</span></div>`;
  if (l.kind === 'doc') {
    l.doc.forEach((d, i) => blocks.push(`<section class="blk band ${state.band++ % 2 ? 'alt' : ''}">${i === 0 ? head : ''}<div class="step"><span class="step-k">PART ${i + 1} OF ${l.doc.length}</span> <span class="step-h">${esc(d.title)}</span></div><div class="doc">${docHtml(d.html)}</div></section>`));
  } else {
    const S = shots[l.id];
    for (let i = 0, g = 0; i < S.length; i += 2, g++) {
      const idx = [i, i + 1].filter((x) => x < S.length);
      const lines = idx.map((x, j) => { const t = /^step \d+$/i.test(S[x].title || '') ? '' : (S[x].title || ''); return `<li><b>${x + 1}.${t ? ' ' + esc(t) + (/[?:.!]$/.test(t) ? '' : ':') : ''}</b> ${txt(S[x].html)}${idx.length > 1 ? ` <span class="lt">(${'ab'[j]})</span>` : ''}</li>`; }).join('');
      const pics = idx.map((x, j) => `<figure>${idx.length > 1 ? `<figcaption>${'ab'[j]}.</figcaption>` : ''}<img src="${img(S[x].file)}"></figure>`).join('');
      const range = idx.length > 1 ? `Steps ${idx[0] + 1} and ${idx[1] + 1}` : `Step ${idx[0] + 1}`;
      blocks.push(`<section class="blk band ${state.band++ % 2 ? 'alt' : ''}">${g === 0 ? head : ''}<div class="step"><span class="step-k">${range.toUpperCase()}</span></div><ul class="say">${lines}</ul><div class="pics n${idx.length}">${pics}</div></section>`);
    }
  }
  return blocks;
}
const doneBlock = (t, b) => `<section class="blk done"><div class="done-t">${esc(t)}</div><div class="done-b">${esc(b)}</div></section>`;
const partBlock = (n, t, b) => `<section class="blk part" data-keep><div class="part-no">PART ${n}</div><div class="part-t">${esc(t)}</div><div class="part-b">${esc(b)}</div></section>`;
function coverHtml(c) {
  return `<div class="cov-top"><span class="logo"></span><div class="cov-k">${esc(c.kicker)}</div><div class="cov-t">${esc(c.title)}</div><div class="cov-s">${esc(c.sub)}</div></div>
<div class="glance" style="--n:${c.glance.length}">${c.glance.map((g, i) => `<div class="gl${i === c.glance.length - 1 && c.goLast ? ' go' : ''}"><b>${esc(g[0])}</b>${esc(g[1])}</div>`).join('')}</div>
<div class="need">${c.need}</div>`;
}
const HOWTO = (l) => `<h3>How to use this guide</h3><ul>
<li>Do the steps in order, from the top down.</li>
${l.kind === 'doc' ? '<li>Read each part before you start the install.</li>' : '<li>In each picture, the <b>red circle</b> shows where to click or type.</li>'}
<li>Words in <b>bold</b> are the buttons, tabs and boxes on your screen.</li>
<li>Practice it online: <a href="${SITE + hrefOf(l)}">${esc(SITE + hrefOf(l))}</a></li></ul>`;

async function renderPdf(name, doc, outFile) {
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${esc(doc.title)}</title><style>${CSS}</style></head><body>
<div id="src" style="display:none">${doc.blocks.join('\n')}</div>
<div id="cover" style="display:none">${doc.cover}</div>
<div id="pages"></div>
<script>
window.addEventListener('load', function () {
var HDR = ${JSON.stringify(`<div class="hdr"><span class="logo"></span><span class="rule"></span><span class="ht">${esc(doc.header)}</span></div>`)};
var FT = ${JSON.stringify(doc.footer)};
var FTR = function (n) { return '<div class="ftr"><span>' + FT + '</span><span class="pn">Page ' + n + ' of <span class="pt"></span></span></div>'; };
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
  else if (body.scrollHeight > body.clientHeight + 1 && cur === pages.firstChild) { body.removeChild(b); newPage(); body.appendChild(b); }
  if (keepNext && src[i + 1]) { body.appendChild(src[i + 1]); var over = body.scrollHeight > body.clientHeight + 1; body.removeChild(src[i + 1]); if (over && body.children.length > 1) { body.removeChild(b); newPage(); body.appendChild(b); } }
});
var total = pages.children.length;
document.querySelectorAll('.pt').forEach(function (e) { e.textContent = total; });
var overflow = []; Array.prototype.forEach.call(document.querySelectorAll('.page .body'), function (b, i) { if (b.scrollHeight > b.clientHeight + 1) overflow.push(i + 1); });
window.__layout = { total: total, overflow: overflow };
});
</script></body></html>`;
  const htmlFile = path.join(OUT, name + '.html'); fs.writeFileSync(htmlFile, html);
  const pp = await browser.newPage({ viewport: { width: 816, height: 1056 }, deviceScaleFactor: 1.25 });
  await pp.goto(pathToFileURL(htmlFile).href); await pp.waitForFunction(() => window.__layout); await pp.waitForTimeout(400);
  const lay = await pp.evaluate(() => window.__layout);
  const PG = path.join(PAGES, name); fs.mkdirSync(PG, { recursive: true });
  for (const f of fs.readdirSync(PG)) fs.unlinkSync(path.join(PG, f));
  const els = await pp.$$('.page');
  for (let i = 0; i < els.length; i++) await els[i].screenshot({ path: path.join(PG, 'page-' + String(i + 1).padStart(2, '0') + '.png') });
  await pp.emulateMedia({ media: 'print' });
  await pp.pdf({ path: outFile, format: 'Letter', printBackground: true, preferCSSPageSize: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
  await pp.close();
  console.log(path.relative(ROOT, outFile), lay.total + ' pages', Math.round(fs.statSync(outFile).size / 1024) + ' KB', lay.overflow.length ? 'OVERFLOW on page ' + lay.overflow.join(', ') : '');
  return lay;
}

// ---------- 3. one PDF per lesson ----------
const problems = [];
for (const l of lessons) {
  if (l.kind !== 'doc' && !shots[l.id]) { problems.push(l.id + ': no screenshots'); continue; }
  const state = { band: 0 };
  const blocks = lessonBlocks(l, '', state);
  blocks.push(doneBlock('Done', l.outro || ''));
  const count = l.kind === 'doc' ? l.doc.length : shots[l.id].length;
  const glance = [[String(count), l.kind === 'doc' ? 'Parts to read' : 'Steps'], [String(l.minutes || 1), 'Minutes'], [String(l.manual || '-'), 'Manual section']];
  const vesselNeed = l.id === 'vessel-shape' ? '<h3>What you need</h3><ul class="two"><li>Straight wall height (feet)</li><li>Width across the bin (feet)</li><li>Cone height (feet)</li><li>Outlet width (feet)</li></ul>'
    : l.id === 'vessel-contents' ? '<h3>What you need</h3><ul class="two"><li>What is in the bin</li><li>Density from your supplier</li></ul>' : '';
  const cover = coverHtml({ kicker: 'Binventory 1.0 Guide · ' + (SECTION[l.section] || ''), title: l.title, sub: l.desc || '', glance, need: vesselNeed + HOWTO(l) });
  const lay = await renderPdf(l.id, {
    title: 'Binventory 1.0 - ' + l.title, header: l.title, blocks, cover,
    footer: esc('BV1-' + l.id.toUpperCase()) + ' &nbsp; ' + esc('BINVENTORY 1.0 GUIDE: ' + l.title.toUpperCase()),
  }, path.join(LESSON_DIR, l.id + '.pdf'));
  if (lay.overflow.length) problems.push(l.id + ': overflow on page ' + lay.overflow.join(', '));
}

// ---------- 4. the vessel book (both vessel lessons in one PDF) ----------
if (buildBook && BOOK_IDS.every((id) => shots[id])) {
  const PARTS = [
    { id: 'vessel-shape', title: 'Size and shape', blurb: 'Type/Shape tab: shape, straight wall, width, cone.', done: 'Size and shape done' },
    { id: 'vessel-contents', title: 'Contents', blurb: 'Contents tab: product name and density.', done: 'Contents done' },
  ];
  const state = { band: 0 }; const blocks = [];
  PARTS.forEach((P, i) => {
    const l = LS.byId(P.id);
    blocks.push(partBlock(i + 1, P.title, P.blurb));
    blocks.push(...lessonBlocks(l, (i + 1) + '.1', state));
    blocks.push(doneBlock(P.done, l.outro));
  });
  const cover = coverHtml({ kicker: 'Vessel configuration', title: 'Binventory 1.0', sub: 'Change the size, shape and contents of a vessel you already have', goLast: true,
    glance: [['1', 'Open the vessel'], ['2', 'Pick the shape'], ['3', 'Type the sizes'], ['4', 'Set the density'], ['5', 'Save']],
    need: '<h3>What you need</h3><ul class="two"><li>Straight wall height (feet)</li><li>Width across the bin (feet)</li><li>Cone height (feet)</li><li>Outlet width (feet)</li><li>What is in the bin</li><li>Density from your supplier</li></ul>' });
  const lay = await renderPdf('vessel-config', { title: 'Binventory 1.0 Vessel Configuration', header: 'Binventory 1.0 Vessel Configuration', blocks, cover,
    footer: 'BV1-VC &nbsp; BINVENTORY 1.0 VESSEL CONFIGURATION' }, path.join(ROOT, 'pdf', 'Binventory1.0-Vessel-Configuration.pdf'));
  if (lay.overflow.length) problems.push('vessel book: overflow on page ' + lay.overflow.join(', '));
} else if (buildBook) problems.push('vessel book: missing screenshots');

await browser.close();
if (failed.length) problems.push('capture failed: ' + failed.join(', '));
console.log(problems.length ? 'PROBLEMS:\n  ' + problems.join('\n  ') : 'All PDFs built.');
process.exitCode = problems.length ? 1 : 0;
