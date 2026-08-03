// core/generate.js — ties parse + shift + dateutil + mapping together into the
// two outputs the tool produces: the flat file rows and the Biologs template rows.

import { cellText, colIndex } from './parse.js';
import { parseShift } from './shift.js';
import { toYMD, addDays, formatMDY } from './dateutil.js';
import { lookupBiometricId } from './mapping.js';

// detailed = parseDetailed() result. mappingMap = parseEmployeeMapping() result.
export function buildPlan(detailed, mappingMap) {
  const dateIdx = colIndex(detailed.headers, 'Date');
  const shiftIdx = colIndex(detailed.headers, 'Shift');
  if (dateIdx === -1 || shiftIdx === -1) {
    throw new Error('Detailed sheet is missing the expected "Date"/"Shift" columns.');
  }

  const biologsRows = [];
  const unmatched = new Map(); // idNumber -> { name, count }
  const shiftErrors = [];
  let excludedDays = 0;

  for (const row of detailed.rows) {
    const dateRaw = row.cells[dateIdx];
    const shiftRaw = cellText(row.cells[shiftIdx]).trim();
    const parsed = parseShift(shiftRaw);

    if (parsed.excluded) { excludedDays++; continue; }

    if (parsed.error) {
      shiftErrors.push({ name: row.name, idNumber: row.idNumber, date: dateRaw, shift: shiftRaw, message: parsed.error });
      continue;
    }

    const ymd = toYMD(dateRaw);
    if (!ymd) {
      shiftErrors.push({ name: row.name, idNumber: row.idNumber, date: dateRaw, shift: shiftRaw, message: `Could not parse Date cell: "${cellText(dateRaw)}"` });
      continue;
    }

    const biometricId = lookupBiometricId(mappingMap, row.idNumber);
    if (biometricId === undefined) {
      const key = row.idNumber || '(blank ID Number)';
      if (!unmatched.has(key)) unmatched.set(key, { name: row.name, count: 0 });
      unmatched.get(key).count++;
      continue;
    }

    biologsRows.push({ biometricId, mode: 'IN', time: parsed.inTime, date: formatMDY(ymd) });
    biologsRows.push({ biometricId, mode: 'OUT', time: parsed.outTime, date: formatMDY(addDays(ymd, parsed.outDayOffset)) });
  }

  return {
    flatHeaders: detailed.headers,
    flatRows: detailed.rows,
    biologsRows,
    issues: { unmatched, shiftErrors },
    stats: {
      employees: new Set(detailed.rows.map((r) => r.idNumber)).size,
      dayRows: detailed.rows.length,
      excludedDays,
      biologsRowsGenerated: biologsRows.length,
    },
  };
}
