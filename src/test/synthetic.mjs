// synthetic.mjs — exercises buildPlan/exportXlsx against hand-crafted data that
// actually has matching IDs, so we can verify: overnight rollover, REST DAY /
// UNPAID LEAVE exclusion, unparseable-shift handling, and unmatched-ID handling
// all in one pass (the real sample files can't test the "matched" path since
// their ID schemes don't overlap).
import { writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

import { buildPlan } from '../core/generate.js';
import { runValidation } from '../core/validate.js';
import { buildFlatFile, buildBiologsTemplate } from '../core/exportXlsx.js';

const require = createRequire(import.meta.url);
const ExcelJS = require('../../vendor/exceljs.min.js');
const here = dirname(fileURLToPath(import.meta.url));

const headers = ['Date', 'Day', 'Shift Type', 'Shift', 'Biologs'];
const detailed = {
  headers,
  rows: [
    // normal day shift, matched -> should produce IN/OUT same date
    { name: 'Alice Tan', idNumber: 'RB-9058', cells: ['8/3/2026', 'Mon', 'Follows Schedule Per Week', '08:00 AM To 05:00 PM', 'NO LOGS'] },
    // overnight shift, matched -> OUT should roll to 8/4/2026
    { name: 'Bob Reyes', idNumber: 'RB-9062', cells: ['8/3/2026', 'Mon', 'Follows Schedule Per Week', '10:00 PM To 06:00 AM', 'NO LOGS'] },
    // rest day -> excluded entirely
    { name: 'Chloe Ng', idNumber: 'RB-9065', cells: ['8/3/2026', 'Mon', 'Follows Schedule Per Week', 'REST DAY', 'NO LOGS'] },
    // unpaid leave with a time range still present -> excluded (keyword wins regardless of time)
    { name: 'Dave Cruz', idNumber: 'RB-9066', cells: ['8/3/2026', 'Mon', 'Follows Schedule Per Week', '01:00 AM To 12:00 AM (UNPAID LEAVE)', 'NO LOGS'] },
    // unmatched ID -> should show up as a validation warning, not in biologsRows
    { name: 'Ghost Employee', idNumber: 'NOPE-0000', cells: ['8/3/2026', 'Mon', 'Follows Schedule Per Week', '09:00 AM To 06:00 PM', 'NO LOGS'] },
    // unparseable shift (only one time token) -> validation warning
    { name: 'Eve Santos', idNumber: 'RB-9067', cells: ['8/3/2026', 'Mon', 'Follows Schedule Per Week', 'Something weird 08:00 AM only', 'NO LOGS'] },
  ],
};

const mapping = new Map([
  ['RB-9058', 9058],
  ['RB-9062', 9062],
  ['RB-9065', 9065],
  ['RB-9066', 9066],
  ['RB-9067', 9067],
]);

const plan = buildPlan(detailed, mapping);
console.log('Stats:', plan.stats);
console.log('\nbiologsRows:');
plan.biologsRows.forEach((r) => console.log(' ', r));

console.log('\nValidation:');
for (const rule of runValidation(plan)) {
  console.log(`[${rule.sev.toUpperCase()}] ${rule.title}`);
  (rule.items || []).forEach((it) => console.log(`    - ${it}`));
}

// Assertions
const assert = (cond, msg) => { if (!cond) { console.error('ASSERT FAILED:', msg); process.exitCode = 1; } else console.log('OK:', msg); };

const alice = plan.biologsRows.filter((r) => r.biometricId === 9058);
assert(alice.length === 2, 'Alice produces exactly 2 rows (IN+OUT)');
assert(alice[0].mode === 'IN' && alice[0].time === '08:00 AM' && alice[0].date === '08/03/2026', 'Alice IN correct');
assert(alice[1].mode === 'OUT' && alice[1].time === '05:00 PM' && alice[1].date === '08/03/2026', 'Alice OUT same-day (non-overnight)');

const bob = plan.biologsRows.filter((r) => r.biometricId === 9062);
assert(bob.length === 2, 'Bob produces exactly 2 rows (IN+OUT)');
assert(bob[0].date === '08/03/2026', 'Bob IN on 08/03/2026');
assert(bob[1].date === '08/04/2026', 'Bob OUT rolls to 08/04/2026 (overnight)');
assert(bob[1].time === '06:00 AM', 'Bob OUT time correct');

assert(!plan.biologsRows.some((r) => r.biometricId === 9065), 'Chloe (REST DAY) excluded');
assert(!plan.biologsRows.some((r) => r.biometricId === 9066), 'Dave (UNPAID LEAVE, has time) excluded');
assert(plan.issues.unmatched.has('NOPE-0000'), 'Ghost Employee flagged as unmatched');
assert(plan.issues.shiftErrors.some((e) => e.idNumber === 'RB-9067'), 'Eve flagged as shift parse error');
assert(plan.stats.excludedDays === 2, 'excludedDays === 2 (Chloe + Dave)');
assert(plan.stats.biologsRowsGenerated === 4, 'biologsRowsGenerated === 4 (Alice + Bob)');

(async () => {
  const flatWb = buildFlatFile(ExcelJS, plan);
  writeFileSync(join(here, 'out_flat_synth.xlsx'), Buffer.from(await flatWb.xlsx.writeBuffer()));
  const bioWb = buildBiologsTemplate(ExcelJS, plan);
  const bioPath = join(here, 'out_biologs_synth.xlsx');
  writeFileSync(bioPath, Buffer.from(await bioWb.xlsx.writeBuffer()));

  const check = new ExcelJS.Workbook();
  await check.xlsx.load(require('node:fs').readFileSync(bioPath));
  const tpl = check.getWorksheet('Template');
  console.log('\nRound-trip Template rows:');
  for (let r = 1; r <= 5; r++) console.log(' ', tpl.getRow(r).values);
  console.log('Sheet order:', check.worksheets.map((w) => w.name));
  console.log('definedNames:', check.definedNames.model);
  console.log('col widths:', [1, 2, 3, 4].map((c) => tpl.getColumn(c).width));
  console.log('col C numFmt:', tpl.getColumn(3).numFmt);

  console.log(process.exitCode === 1 ? '\nFAIL' : '\nALL ASSERTIONS PASSED');
})();
