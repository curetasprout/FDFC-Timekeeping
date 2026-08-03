// pipeline.mjs — runs the real core logic against the actual sample files and
// dumps the results for manual inspection. Node ESM (import works directly on
// the core/*.js files since they use `export` syntax); ExcelJS is a UMD bundle
// so it's pulled in via createRequire.
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

import { parseDetailed } from '../core/parse.js';
import { parseEmployeeMapping } from '../core/mapping.js';
import { buildPlan } from '../core/generate.js';
import { runValidation } from '../core/validate.js';
import { buildFlatFile, buildBiologsTemplate } from '../core/exportXlsx.js';

const require = createRequire(import.meta.url);
const ExcelJS = require('../../vendor/exceljs.min.js');

const here = dirname(fileURLToPath(import.meta.url));

const ATTENDANCE = 'C:\\Users\\CUreta\\Desktop\\AttendanceReport(08032026-08032026).xlsx';
const EMPLIST = 'C:\\Users\\CUreta\\Downloads\\Employee list report - 2026-08-03-11_01.xls';

(async () => {
  const attWb = new ExcelJS.Workbook();
  await attWb.xlsx.load(readFileSync(ATTENDANCE));
  const detailedWs = attWb.getWorksheet('Detailed');
  if (!detailedWs) throw new Error('No Detailed sheet found in attendance file');
  const detailed = parseDetailed(detailedWs.getSheetValues());
  console.log(`Parsed Detailed: ${detailed.headers.length} headers, ${detailed.rows.length} day-rows`);
  console.log('Headers:', detailed.headers.join(' | '));

  const summaryWs = attWb.getWorksheet('Summary');
  const summaryValues = summaryWs ? summaryWs.getSheetValues() : null;
  console.log(`Summary sheet present: ${!!summaryWs}${summaryWs ? `, ${summaryValues.length - 1} rows` : ''}`);

  const empWb = new ExcelJS.Workbook();
  await empWb.xlsx.load(readFileSync(EMPLIST));
  const empWs = empWb.worksheets[0];
  const mapping = parseEmployeeMapping(empWs.getSheetValues());
  console.log(`\nParsed Employee mapping: ${mapping.size} entries`);

  const plan = buildPlan(detailed, mapping);
  console.log('\nStats:', plan.stats);

  console.log('\n--- Validation ---');
  for (const rule of runValidation(plan)) {
    console.log(`[${rule.sev.toUpperCase()}] ${rule.title}`);
    (rule.items || []).forEach((it) => console.log(`    - ${it}`));
    if (rule.hint) console.log(`    hint: ${rule.hint}`);
  }

  console.log('\n--- Sample biologsRows (first 10) ---');
  plan.biologsRows.slice(0, 10).forEach((r) => console.log(r));

  console.log('\n--- Sample flatRows (first 3) ---');
  plan.flatRows.slice(0, 3).forEach((r) => console.log({ name: r.name, idNumber: r.idNumber, cells: r.cells }));

  const flatWb = buildFlatFile(ExcelJS, plan, summaryValues);
  const flatOut = join(here, 'out_flat.xlsx');
  writeFileSync(flatOut, Buffer.from(await flatWb.xlsx.writeBuffer()));
  console.log(`\nWrote ${flatOut}`);
  console.log('Flat file sheets:', flatWb.worksheets.map((w) => w.name));

  const bioWb = buildBiologsTemplate(ExcelJS, plan);
  const bioOut = join(here, 'out_biologs.xlsx');
  writeFileSync(bioOut, Buffer.from(await bioWb.xlsx.writeBuffer()));
  console.log(`Wrote ${bioOut}`);

  // Round-trip sanity check: re-read the biologs output and confirm the named
  // range / dropdown / text-format survived the write.
  const check = new ExcelJS.Workbook();
  await check.xlsx.load(readFileSync(bioOut));
  const tpl = check.getWorksheet('Template');
  console.log('\nRound-trip check — Template col C numFmt:', tpl.getColumn(3).numFmt);
  console.log('Round-trip check — defined names:', check.definedNames.model);
  console.log('Round-trip check — row2 values:', tpl.getRow(2).values);
})().catch((e) => { console.error('FAILED:', e); process.exit(1); });
