/* Build the self-contained Global Purchase Report (3D dashboard) HTML.
 *
 *   node build.mjs                 → ../Global_Purchase_3D_WarRoom_TEMP.html (working file)
 *   node build.mjs 202609          → also writes ../Global_Purchase_3D_WarRoom_202609.html (snapshot)
 *
 * Inputs live in this folder: src/ (hand-written parts), data/ (report-seed.json +
 * land-110m.json), vendor/ (echarts, three, fonts, lucide icons). Re-run after
 * refreshing data/report-seed.json from the dated Excel deliverable. */
import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..');
const P = (...parts) => join(HERE, ...parts);
const read = (p) => readFileSync(p, 'utf8');

/* ── 1. styles: embedded fonts + kit tokens + big-screen page CSS ────────── */
const b64 = (p) => readFileSync(p).toString('base64');
const fonts = `@font-face{font-family:"Geist Variable";font-style:normal;font-weight:100 900;font-display:swap;
  src:url(data:font/woff2;base64,${b64(P('vendor/fonts/geist-latin-wght.woff2'))}) format("woff2");}
@font-face{font-family:"Geist Mono Variable";font-style:normal;font-weight:100 900;font-display:swap;
  src:url(data:font/woff2;base64,${b64(P('vendor/fonts/geist-mono-latin-wght.woff2'))}) format("woff2");}`;
/* original kit.css loaded once with the local fonts — same cascade, no @font-face there. */
const kitSrc = read(P('styles-original/kit.css')).replace(/@font-face\{[\s\S]*?format\("woff2"\);\}\n/g, '');
if (kitSrc.includes('@font-face')) throw new Error('kit.css @font-face strip failed');
const appSrc = read(P('styles-original/app.css'));
const assistantCss = read(P('src/assistant.css'));
/* single-file additions: hide the globe loading layer once the canvas is live;
   中文/ENG switch chrome (same placement habit as the tabular report's toggle) */
const extra = `
.globe-fallback.ready{display:none}
.lang-btn{gap:2px;padding:5px 9px;font:600 12px/1 var(--mono);letter-spacing:.04em}
.lang-btn .lang-opt{color:var(--muted);padding:3px 4px;border-radius:3px;transition:color var(--fast),background var(--fast)}
.lang-btn .lang-opt.on{color:var(--accent);background:var(--accent-soft)}
.lang-btn .lang-sep{color:var(--line);margin:0 1px}
`;
const styles = fonts + '\n' + kitSrc + '\n' + appSrc + '\n' + assistantCss + '\n' + extra;

/* ── 2. icons: lucide svg bodies keyed for icon() helper ─────────────────── */
const KEY_MAP = { 'triangle-alert': 'alert-triangle' };  // renamed upstream
const icons = {};
for (const file of readdirSync(P('vendor/icons'))) {
  if (!file.endsWith('.svg')) continue;
  const name = file.slice(0, -4);
  const svg = read(P('vendor/icons', file));
  const inner = svg.replace(/^[\s\S]*?<svg[^>]*>/, '').replace(/<\/svg>[\s\S]*$/, '');
  icons[KEY_MAP[name] ?? name] = inner.trim();
}

/* ── 3. payloads: data / land / three, each escape-checked ───────────────── */
const seed = JSON.parse(read(P('data/report-seed.json')));
const dataJson = JSON.stringify(seed).replaceAll('<', '\\u003c');
const landJson = JSON.stringify(JSON.parse(read(P('data/land-110m.json')))).replaceAll('<', '\\u003c');
const threeSrc = read(P('vendor/lib/three.module.min.js'));
const threeJson = JSON.stringify(threeSrc).replaceAll('<', '\\u003c');
const echartsSrc = read(P('vendor/lib/echarts.min.js'));

for (const [label, text] of [['echarts', echartsSrc], ['three (JSON)', threeJson], ['data', dataJson], ['land', landJson]]) {
  if (text.includes('</script') || text.includes('<!--')) throw new Error(`${label} contains a script-terminating sequence — refusing to inline`);
}

/* ── 4. app scripts (concatenated into one ES module) ────────────────────── */
const core = read(P('src/core.js'));
const globe = read(P('src/globe.js'));
let main = read(P('src/main.js'));
main = main.replace('/*__ICONS__*/ {}', '/*__ICONS__*/ ' + JSON.stringify(icons));
if (main.includes('/*__ICONS__*/ {}')) throw new Error('icon placeholder not replaced');
const assistant = read(P('src/assistant.js'));   /* must come last: binds the AI button after boot() */

/* ── 5. assemble ─────────────────────────────────────────────────────────── */
const skeleton = read(P('src/skeleton.html'));
const buildDate = new Date().toISOString().slice(0, 10);
const html = `<!doctype html>
<html lang="zh-Hant" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#050810">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%23050810'/%3E%3Ccircle cx='32' cy='32' r='20' fill='none' stroke='%2322d3ee' stroke-width='3'/%3E%3Cpath d='M12 32h40M32 12c8 8 8 32 0 40M32 12c-8 8-8 32 0 40' fill='none' stroke='%2322d3ee' stroke-width='3'/%3E%3C/svg%3E">
<title>Global Purchase Report 全球採購報告</title>
<meta name="description" content="Global Purchase Report 全球採購報告 — self-contained 3D dashboard. Data: FIN Global Purchase Report Jan–Sep 2026. Built ${buildDate}.">
<style>
${styles}
</style>
</head>
<body>
${skeleton}
<script id="gpr-data" type="application/json">${dataJson}</script>
<script id="gpr-land" type="application/json">${landJson}</script>
<script id="gpr-three" type="application/json">${threeJson}</script>
<script>${echartsSrc}</script>
<script type="module">
${core}
${globe}
${main}
${assistant}
</script>
</body>
</html>
`;

const tempPath = join(ROOT, 'Global_Purchase_3D_WarRoom_TEMP.html');
writeFileSync(tempPath, html);
console.log(`WROTE ${tempPath}  ${(Buffer.byteLength(html) / 1048576).toFixed(1)} MB  rows=${seed.rows.length}`);

const publishArg = process.argv[2];
if (publishArg) {
  if (!/^\d{6}$/.test(publishArg)) throw new Error('publish argument must be YYYYMM');
  const snapPath = join(ROOT, `Global_Purchase_3D_WarRoom_${publishArg}.html`);
  writeFileSync(snapPath, html);
  console.log(`PUBLISHED ${snapPath}`);
}
