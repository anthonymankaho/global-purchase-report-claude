/* Build the self-contained Global Purchase Report (3D dashboard) HTML.
 *
 *   node build.mjs                 → ../Global_Purchase_3D_WarRoom_TEMP.html (working file)
 *   node build.mjs 202609          → also writes ../Global_Purchase_3D_WarRoom_202609.html (snapshot)
 *
 * Inputs live in this folder: src/ (hand-written parts), data/ (report-seed.json +
 * land-110m.json), vendor/ (echarts, three, fonts, lucide icons). Re-run after
 * refreshing data/report-seed.json from the dated Excel deliverable. */
import { readFileSync, writeFileSync, readdirSync, statSync, mkdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..');
const P = (...parts) => join(HERE, ...parts);
const read = (p) => readFileSync(p, 'utf8');
const args = process.argv.slice(2);
const hostedIndex = args.indexOf('--hosted');
const hosted = hostedIndex >= 0;
const hostedOutputArg = hosted ? args[hostedIndex + 1] : null;
const publishArg = args.find((arg) => /^\d{6}$/.test(arg));
if (hosted && (!hostedOutputArg || hostedOutputArg.startsWith('--'))) {
  throw new Error('Usage: node build.mjs --hosted <output-html-path>');
}
if (hosted && publishArg) throw new Error('hosted mode does not create a local dated snapshot');
if (args.some((arg, index) => arg.startsWith('--') && !(arg === '--hosted' && index === hostedIndex))) {
  throw new Error('unknown build option');
}
if (!hosted && args.some((arg) => !/^\d{6}$/.test(arg))) {
  throw new Error('snapshot argument must be YYYYMM');
}

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
const seed = hosted ? null : JSON.parse(read(P('data/report-seed.json')));
const dataJson = seed ? JSON.stringify(seed).replaceAll('<', '\\u003c') : '';
const landJson = JSON.stringify(JSON.parse(read(P('data/land-110m.json')))).replaceAll('<', '\\u003c');
const threeSrc = read(P('vendor/lib/three.module.min.js'));
const threeJson = JSON.stringify(threeSrc).replaceAll('<', '\\u003c');
const echartsSrc = read(P('vendor/lib/echarts.min.js'));

for (const [label, text] of [['echarts', echartsSrc], ['three (JSON)', threeJson], ['data', dataJson], ['land', landJson]]) {
  if (text.includes('</script') || text.includes('<!--')) throw new Error(`${label} contains a script-terminating sequence — refusing to inline`);
}

/* ── 4. app scripts (concatenated into one ES module) ────────────────────── */
const coreSource = read(P('src/core.js'));
const globe = read(P('src/globe.js'));
const mainSource = read(P('src/main.js'));
const assistantSource = read(P('src/assistant.js'));   /* must come last: binds the AI button after boot() */

/* ── 5. assemble ─────────────────────────────────────────────────────────── */
let skeleton = read(P('src/skeleton.html'));
const buildDate = new Date().toISOString().slice(0, 10);
if (hosted) {
  const beforeLanguageButton = skeleton;
  skeleton = skeleton.replace(
    /<button\b(?=[^>]*\bid="lang-btn")[^>]*>[\s\S]*?<\/button>/,
    '<button class="btn lang-btn" id="lang-btn" type="button" hidden aria-hidden="true"></button>',
  );
  if (skeleton === beforeLanguageButton) throw new Error('hosted language control marker changed');
  skeleton = skeleton
    .replace('aria-label="採購篩選"', 'aria-label="Procurement filters"')
    .replace('aria-label="採購關鍵指標"', 'aria-label="Purchase KPIs"')
    .replace('aria-label="可旋轉的全球採購實體地球"', 'aria-label="Rotatable global procurement globe"')
    .replace('aria-label="搜尋採購明細"', 'aria-label="Search purchase details"')
    .replace('aria-label="清除搜尋"', 'aria-label="Clear search"');
}
function buildHtml(isHosted) {
let core = coreSource;
if (isHosted) {
  const localLoad = "const RAW = readJson('gpr-data');";
  const hostedLoad = `const RAW = await (async () => {
  try {
    const response = await fetch('../api/data', { credentials: 'same-origin', cache: 'no-store' });
    if (!response.ok) throw new Error('report_data_unavailable');
    return await response.json();
  } catch {
    document.body.innerHTML = '<main style="font:16px Arial,sans-serif;padding:32px;color:#222"><h1>Report data unavailable</h1><p>Please reload later.</p></main>';
    throw new Error('hosted_report_data_unavailable');
  }
})();`;
  if (core.split(localLoad).length !== 2) throw new Error('local data-load marker changed; hosted build refused');
  core = core.replace(localLoad, hostedLoad);
}

let main = mainSource.replace('/*__ICONS__*/ {}', '/*__ICONS__*/ ' + JSON.stringify(icons));
if (main.includes('/*__ICONS__*/ {}')) throw new Error('icon placeholder not replaced');
if (isHosted) {
  const languageInitializer = /const initialLang = \(\(\) => \{[\s\S]*?\}\)\(\);/g;
  if ([...main.matchAll(languageInitializer)].length !== 1) throw new Error('language initializer changed; hosted English-only build refused');
  main = main.replace(languageInitializer, "const initialLang = 'en';");
  main = main.replace(/^\s*\$\('#lang-btn'\)\.addEventListener\('click'.*\);\r?\n/m, '\n');
  if (main.includes("addEventListener('click', () => setLang(state.lang")) throw new Error('hosted language toggle remains enabled');
  const hostedLabels = [
    ['資料只送到本機 AI 中繼,不含任何憑證。', '僅傳送目前畫面的彙總資料；不傳送任何憑證。'],
    ['無法連線 AI 中繼服務。', '無法連線報告 AI 服務。'],
    ['AI 中繼已連線', 'AI 服務已連線'],
    ['AI 助理未啟動:在本機執行 .mpl-qwen\\\\start-proxy.bat 後重試;報表其餘功能不受影響。', 'AI 服務暫時無法使用；報表其餘功能不受影響。'],
    ['Data only goes to the local relay - never any credentials.', 'Only aggregate report context is sent to the report AI service; credentials are never sent.'],
    ['Cannot reach the AI relay service.', 'Cannot reach the report AI service.'],
    ['AI relay connected', 'AI service connected'],
    ['Assistant not running: start .mpl-qwen\\\\start-proxy.bat on this PC and retry; the rest of the report is unaffected.', 'AI service is unavailable. The rest of the report is unaffected.'],
  ];
  for (const [from, to] of hostedLabels) main = main.replace(from, to);
}

let assistant = assistantSource;
if (isHosted) {
  const localRelay = "const RELAY = 'http://127.0.0.1:8765';";
  if (assistant.split(localRelay).length !== 2) throw new Error('local relay marker changed; hosted build refused');
  assistant = assistant.replace(localRelay, "const RELAY = '../api';").replace('http://127.0.0.1:8765', 'the local relay')
    .replace('Global Purchase Report 3D dashboard (self-contained HTML, NOT MPL_System)', 'Global Purchase Report 3D hosted dashboard');
}

const dataIsland = isHosted ? '' : `<script id="gpr-data" type="application/json">${dataJson}</script>`;
const hostedNav = isHosted ? `<a class="gpr-main-report-link" href="../" aria-label="Back to the main report">← Main report</a>` : '';
const hostedStyle = isHosted ? `.gpr-main-report-link{position:fixed;z-index:10000;left:12px;bottom:12px;padding:9px 13px;border-radius:8px;background:#123b70;color:#fff;text-decoration:none;font:600 13px Arial,sans-serif;box-shadow:0 2px 10px #0005}.gpr-main-report-link:hover{background:#005daa}` : '';
return `<!doctype html>
<html lang="${isHosted ? 'en' : 'zh-Hant'}" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#050810">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%23050810'/%3E%3Ccircle cx='32' cy='32' r='20' fill='none' stroke='%2322d3ee' stroke-width='3'/%3E%3Cpath d='M12 32h40M32 12c8 8 8 32 0 40M32 12c-8 8-8 32 0 40' fill='none' stroke='%2322d3ee' stroke-width='3'/%3E%3C/svg%3E">
<title>${isHosted ? 'Global Purchase Report' : 'Global Purchase Report 全球採購報告'}</title>
<meta name="description" content="Global Purchase Report dashboard. ${isHosted ? 'Hosted 3D view with report data loaded from the authorized Worker API.' : 'Self-contained 3D dashboard.'} Data: FIN Global Purchase Report Jan–Sep 2026. Built ${buildDate}.">
<style>
${styles}
${hostedStyle}
</style>
</head>
<body>
${hostedNav}
${skeleton}
${dataIsland}
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
}

if (hosted) {
  const hostedPath = resolve(process.cwd(), hostedOutputArg);
  const html = buildHtml(true);
  if (html.includes('<script id="gpr-data"') || (seed && html.includes(dataJson))) {
    throw new Error('hosted build contains embedded purchase data; refusing to write');
  }
  mkdirSync(dirname(hostedPath), { recursive: true });
  writeFileSync(hostedPath, html);
  console.log(`WROTE HOSTED ${hostedPath}  ${(Buffer.byteLength(html) / 1048576).toFixed(1)} MB  data=Worker API`);
} else {
  const html = buildHtml(false);
  const tempPath = join(ROOT, 'Global_Purchase_3D_WarRoom_TEMP.html');
  writeFileSync(tempPath, html);
  console.log(`WROTE ${tempPath}  ${(Buffer.byteLength(html) / 1048576).toFixed(1)} MB  rows=${seed.rows.length}`);
  if (publishArg) {
  const snapPath = join(ROOT, `Global_Purchase_3D_WarRoom_${publishArg}.html`);
  writeFileSync(snapPath, html);
  console.log(`PUBLISHED ${snapPath}`);
  }
}
