/* expected.mjs — recompute the headline numbers straight from the seed file,
 * as an independent check on what the War Room HTML renders. */
import { readFileSync } from 'node:fs';
const d = JSON.parse(readFileSync(new URL('data/report-seed.json', import.meta.url), 'utf8'));
const F = d.fields;
const iTot = F.indexOf('TOT_PO_PRICE_USD'), iDb = F.indexOf('DB'), iSt = F.indexOf('ignore / consider'), iM = F.indexOf('yyyy-mm'), iV = F.indexOf('Std Vendor Name');
const tot = d.rows.reduce((s, r) => s + r[iTot], 0);
const byDb = {}; for (const r of d.rows) byDb[r[iDb]] = (byDb[r[iDb]] || 0) + r[iTot];
const ignore = d.rows.filter((r) => r[iSt] === 'Ignore').reduce((s, r) => s + r[iTot], 0);
const usd = (v) => v >= 1e9 ? '$' + (v / 1e9).toFixed(2) + 'B' : v >= 1e6 ? '$' + (v / 1e6).toFixed(2) + 'M' : v >= 1e3 ? '$' + (v / 1e3).toFixed(2) + 'K' : '$' + v.toFixed(0);
console.log('totalUsd', tot.toFixed(2), '→', usd(tot));
console.log('byEntity', Object.entries(byDb).map(([k, v]) => [k, Math.round(v)]).sort((a, b) => b[1] - a[1]));
console.log('ignoreShare', (ignore / tot * 100).toFixed(1) + '%');
const byV = {}; for (const r of d.rows) byV[r[iV]] = (byV[r[iV]] || 0) + r[iTot];
const topV = Object.entries(byV).sort((a, b) => b[1] - a[1])[0];
console.log('topVendor', topV[0], usd(topV[1]), (topV[1] / tot * 100).toFixed(1) + '%');
const m = {}; for (const r of d.rows) m[r[iM]] = (m[r[iM]] || 0) + r[iTot];
const lastTwo = Object.entries(m).sort();
console.log('months', lastTwo.map(([k, v]) => k + ':' + (v / 1e6).toFixed(2) + 'M').join(' '));
const prev = lastTwo.at(-2)[1], last = lastTwo.at(-1)[1];
console.log('monthChange', ((last - prev) / prev * 100).toFixed(1) + '%');
