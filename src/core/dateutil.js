// core/dateutil.js — normalizes the Detailed sheet's Date cell (seen as a plain
// "8/3/2026"-style string in samples, but handled as a real Date too in case a
// future export stores it as an actual Excel date) into a plain {y,m,d} triple,
// then formats/shifts it for the Biologs template's "mm/dd/yyyy" column.

export function toYMD(value) {
  if (value instanceof Date) {
    return { y: value.getFullYear(), m: value.getMonth() + 1, d: value.getDate() };
  }
  const s = String(value ?? '').trim();
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (!m) return null;
  let [, mo, da, yr] = m;
  yr = yr.length === 2 ? 2000 + parseInt(yr, 10) : parseInt(yr, 10);
  return { y: yr, m: parseInt(mo, 10), d: parseInt(da, 10) };
}

export function addDays(ymd, n) {
  const dt = new Date(ymd.y, ymd.m - 1, ymd.d);
  dt.setDate(dt.getDate() + n);
  return { y: dt.getFullYear(), m: dt.getMonth() + 1, d: dt.getDate() };
}

export function formatMDY(ymd) {
  return `${String(ymd.m).padStart(2, '0')}/${String(ymd.d).padStart(2, '0')}/${ymd.y}`;
}
