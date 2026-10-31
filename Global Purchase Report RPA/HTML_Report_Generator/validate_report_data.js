'use strict';
/**
 * Read-only validation of the extracted purchase-report JSON before HTML publication.
 * Usage: node validate_report_data.js <workDir> <YYYY-MM> [--strict]
 * Never changes the source workbook or any report file.
 */
const fs = require('fs');
const path = require('path');
function validate(sheet, period) {
  const errors = [], warnings = [];
  if (!sheet || !Array.isArray(sheet.header) || !Array.isArray(sheet.rows)) {
    return { ok: false, errors: ['Expected JSON object with header[] and rows[]'], warnings, summary: {} };
  }
  const header = sheet.header;
  const index = Object.fromEntries(header.map((h, i) => [String(h), i]));
  if (header.length !== new Set(header).size) errors.push('Duplicate column headers');
  for (const required of ['DB', 'PO_REC_DATE']) {
    if (!(required in index)) errors.push('Missing required column: ' + required);
  }
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(period)) errors.push('Invalid expected period: ' + period);
  const widths = [];
  let invalidDates = 0, inPeriod = 0, outOfPeriod = 0, emptyDb = 0;
  const entities = {};
  for (let i = 0; i < sheet.rows.length; i++) {
    const row = sheet.rows[i];
    if (!Array.isArray(row) || row.length !== header.length) {
      if (widths.length < 5) widths.push(i + 2);
      continue;
    }
    if ('DB' in index) {
      const db = String(row[index.DB] ?? '').trim();
      if (!db) emptyDb++;
      else entities[db] = (entities[db] || 0) + 1;
    }
    if ('PO_REC_DATE' in index) {
      const value = row[index.PO_REC_DATE];
      let month = null;
      if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
        const d = new Date(value + 'T00:00:00Z');
        if (!Number.isNaN(d.valueOf()) && d.toISOString().slice(0, 10) === value) month = value.slice(0, 7);
      } else if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
        const d = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
        month = d.toISOString().slice(0, 7);
      }
      if (!month) invalidDates++;
      else if (month === period) inPeriod++;
      else outOfPeriod++;
    }
  }
  if (widths.length) errors.push('Row width mismatch (Excel row examples): ' + widths.join(', '));
  if (!sheet.rows.length) errors.push('No transaction rows');
  if (emptyDb) warnings.push(emptyDb + ' rows have empty DB/entity');
  if (invalidDates) warnings.push(invalidDates + ' rows have missing or invalid PO_REC_DATE');
  // The report is cumulative: older-period rows are expected, not an error.
  if (!inPeriod) errors.push('No rows found for expected reporting period ' + period);
  return { ok: errors.length === 0, errors, warnings, summary: {
    expectedPeriod: period, totalRows: sheet.rows.length, columns: header.length,
    currentPeriodRows: inPeriod, otherPeriodRows: outOfPeriod, invalidDates, entities
  }};
}
if (require.main === module) {
  const [dir, period, flag] = process.argv.slice(2);
  if (!dir || !period) { console.error('Usage: node validate_report_data.js <workDir> <YYYY-MM> [--strict]'); process.exit(2); }
  try {
    const result = validate(JSON.parse(fs.readFileSync(path.join(dir, 'data_sheet.json'), 'utf8').replace(/^\uFEFF/, '')), period);
    console.log(JSON.stringify(result, null, 2));
    if (!result.ok || (flag === '--strict' && result.warnings.length)) process.exitCode = 1;
  } catch (err) { console.error('Validation failed:', err.message); process.exitCode = 2; }
}
module.exports = { validate };
