// core/mapping.js — Employee List Report -> Biometric ID lookup.
//
// Expected shape: a single sheet with a header row containing "Employee ID" and
// "Biometric ID" columns (order-agnostic, matched by header text). Employee ID is
// normalized (trimmed, uppercased) so lookups tolerate whitespace/casing drift
// against the Detailed sheet's "ID Number" values.

import { cellText, normId } from './parse.js';

export function parseEmployeeMapping(values) {
  const maxR = values.length - 1;
  if (maxR < 1) throw new Error('Employee List Report looks empty.');

  const header = values[1] || [];
  let idCol = null, biometricCol = null;
  for (let c = 1; c < header.length; c++) {
    const h = cellText(header[c]).trim().toLowerCase();
    if (h === 'employee id') idCol = c;
    else if (h === 'biometric id') biometricCol = c;
  }
  if (!idCol || !biometricCol) {
    throw new Error('Could not find "Employee ID" and "Biometric ID" columns in the Employee List Report.');
  }

  const map = new Map(); // normalized Employee ID -> raw Biometric ID value
  for (let r = 2; r <= maxR; r++) {
    const row = values[r];
    if (!row) continue;
    const empId = normId(row[idCol]);
    if (!empId) continue;
    map.set(empId.toUpperCase(), row[biometricCol]);
  }
  return map;
}

export function lookupBiometricId(map, idNumber) {
  return map.get(String(idNumber || '').toUpperCase());
}
