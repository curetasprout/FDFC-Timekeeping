// core/exportXlsx.js — builds the two output workbooks from a buildPlan() result.

// Deliverable 1: the Detailed sheet flattened, with Name/ID Number prepended to
// every row (a clean rebuild — values only, no attempt to mirror the source's
// per-block layout/styling since the whole point is to de-nest it). The source
// Summary sheet, if provided, is carried over as-is (raw values) so the flat
// file stays self-contained/reconcilable with the original report.
export function buildFlatFile(ExcelJS, plan, summaryValues) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'HRIS Biologs Generator';

  if (summaryValues) {
    const sOut = wb.addWorksheet('Summary');
    for (let r = 1; r < summaryValues.length; r++) {
      const row = summaryValues[r];
      if (!row) continue;
      for (let c = 1; c < row.length; c++) {
        const v = row[c];
        if (v !== undefined && v !== null) sOut.getRow(r).getCell(c).value = v;
      }
    }
    sOut.getRow(1).eachCell((cell) => { cell.font = { name: 'Calibri', size: 11, bold: true }; });
    sOut.views = [{ state: 'frozen', ySplit: 1 }];
  }

  const ws = wb.addWorksheet('Detailed (Flat)');

  const headers = ['Name', 'ID Number', ...plan.flatHeaders];
  const headerRow = ws.getRow(1);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Calibri', size: 11, bold: true };
  });
  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 16;
  for (let i = 3; i <= headers.length; i++) ws.getColumn(i).width = 14;
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  plan.flatRows.forEach((r, i) => {
    const row = ws.getRow(i + 2);
    row.getCell(1).value = r.name;
    row.getCell(2).value = r.idNumber;
    r.cells.forEach((v, j) => { row.getCell(j + 3).value = v === undefined ? null : v; });
  });

  return wb;
}

// Deliverable 2: exact-fidelity rebuild of HRISBiologsTemplate.xlsx (Template +
// Sheet3 sheets, the LogMode named range, the Log Mode dropdown on column B,
// and the same column widths / text-formatted Log Time column), populated with
// the IN/OUT rows from plan.biologsRows.
export function buildBiologsTemplate(ExcelJS, plan) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'HRIS Biologs Generator';

  const ws = wb.addWorksheet('Template');
  const headerLabels = ['Biometric ID', 'Log Mode (IN/OUT)', 'Log Time (hh:mm am/pm)', 'Log Date (mm/dd/yyyy)'];
  const headerRow = ws.getRow(1);
  headerLabels.forEach((label, i) => { headerRow.getCell(i + 1).value = label; });

  ws.getColumn(1).width = 11.85546875;
  ws.getColumn(2).width = 18.28515625;
  ws.getColumn(3).width = 24.28515625;
  ws.getColumn(4).width = 22.28515625;
  ws.getColumn(3).numFmt = '@'; // Log Time stored as literal text, matching the source template

  plan.biologsRows.forEach((r, i) => {
    const row = ws.getRow(i + 2);
    row.getCell(1).value = r.biometricId;
    row.getCell(2).value = r.mode;
    row.getCell(3).value = r.time;
    row.getCell(4).value = r.date;
  });

  const sheet3 = wb.addWorksheet('Sheet3');
  sheet3.getCell('A1').value = 'Log Mode';
  sheet3.getCell('A2').value = 'IN';
  sheet3.getCell('A3').value = 'OUT';

  wb.definedNames.add('Sheet3!$A$2:$A$5', 'LogMode');
  ws.dataValidations.add('B2:B1048576', { type: 'list', allowBlank: true, formulae: ['LogMode'] });

  return wb;
}
