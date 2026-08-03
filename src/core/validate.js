// core/validate.js — validation rules over a buildPlan() result.
// Each rule returns { sev: 'error'|'warn'|'info'|'ok', title, items?, hint? }.
// Warnings/info never block export — they just inform (same convention as the
// GL Report Builder). Nothing here is a hard 'error' because every gap has an
// agreed-upon fallback (skip the row).

import { cellText } from './parse.js';

export function runValidation(plan) {
  const out = [];

  if (plan.issues.unmatched.size) {
    const items = [...plan.issues.unmatched].map(([idNumber, info]) => `${idNumber} — ${info.name} (${info.count} day${info.count === 1 ? '' : 's'} skipped)`);
    out.push({
      sev: 'warn',
      title: `${plan.issues.unmatched.size} employee(s) have no Biometric ID match`,
      items,
      hint: 'These employees are left out of the Biologs template entirely. Fix the ID Number in the Attendance export or add them to the Employee List Report mapping, then re-run.',
    });
  }

  if (plan.issues.shiftErrors.length) {
    const items = plan.issues.shiftErrors.map((e) => `${e.name} (${e.idNumber}) — ${cellText(e.date)}: ${e.message}`);
    out.push({
      sev: 'warn',
      title: `${plan.issues.shiftErrors.length} day(s) had no parseable shift time`,
      items,
      hint: 'These day-rows are skipped from the Biologs template — the Shift cell had no recognizable "h:mm AM/PM To h:mm AM/PM" pattern.',
    });
  }

  out.push({
    sev: 'info',
    title: `${plan.stats.employees} employee(s), ${plan.stats.dayRows} day-row(s) parsed`,
    items: [
      `${plan.stats.excludedDays} day(s) excluded (Rest Day / Unpaid Leave)`,
      `${plan.stats.biologsRowsGenerated} log row(s) will be written to the Biologs template (${plan.stats.biologsRowsGenerated / 2} working day(s) × IN+OUT)`,
      `${plan.flatRows.length} row(s) will be written to the flat file`,
    ],
  });

  if (!plan.issues.unmatched.size && !plan.issues.shiftErrors.length) {
    out.push({ sev: 'ok', title: 'No mapping or parsing issues found' });
  }

  return out;
}
