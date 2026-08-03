// core/parse.js — Detailed sheet parser (marker-based, tolerant of any date range).
//
// The Attendance Report's "Detailed" tab is not a grid — it's stacked per-employee
// blocks (Name: / ID Number: / Days Present / Days Absent / header row / N day rows /
// Totals: row / blank rows), and the number of day rows varies with the report's date
// range. We walk the sheet by "Name:" markers rather than assuming fixed row offsets.

export function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.text !== undefined) return String(v.text);
    if (v.result !== undefined) return String(v.result);
    if (v instanceof Date) return v.toISOString();
    return String(v);
  }
  return String(v);
}

export function normId(x) {
  if (x === null || x === undefined) return '';
  if (typeof x === 'object' && x.text !== undefined) x = x.text;
  if (typeof x === 'number' && Number.isInteger(x)) return String(x);
  return String(x).trim();
}

// values = worksheet.getSheetValues() (1-indexed rows, 1-indexed cols).
// Returns { headers, rows } where headers is the Detailed tab's per-day column
// labels (Date, Day, Shift Type, Shift, Biologs, ...) starting at column A, and
// rows is a flat list of { name, idNumber, cells } — one entry per employee-day,
// with cells aligned 1:1 to headers. Totals rows and blank separator rows are
// dropped (identified by an empty Date cell, same rule the report itself uses).
export function parseDetailed(values) {
  const maxR = values.length - 1;
  const colA = (r) => cellText(values[r] && values[r][1]).trim();

  const starts = [];
  for (let r = 1; r <= maxR; r++) if (colA(r) === 'Name:') starts.push(r);
  if (!starts.length) {
    throw new Error('No employee blocks found — expected "Name:" markers in column A of the Detailed sheet.');
  }

  let headers = null;
  const rows = [];

  for (let bi = 0; bi < starts.length; bi++) {
    const start = starts[bi];
    const end = bi + 1 < starts.length ? starts[bi + 1] : maxR + 1;

    let name = null, idNumber = null, hdrRow = null;
    for (let r = start; r < end; r++) {
      const a = colA(r);
      if (a === 'Name:') name = values[r][2];
      else if (a === 'ID Number:') idNumber = values[r][2];
      else if (a === 'Date') hdrRow = r;
    }
    if (hdrRow === null) continue; // malformed block — no per-day table, skip

    const hrow = values[hdrRow] || [];
    if (!headers) {
      headers = [];
      for (let c = 1; c < hrow.length; c++) headers.push(cellText(hrow[c]).trim());
    }

    for (let r = hdrRow + 1; r < end; r++) {
      const row = values[r];
      if (!row) continue;
      if (cellText(row[1]).trim() === '') continue; // blank row or the block's "Totals:" row
      const cells = [];
      for (let c = 1; c < hrow.length; c++) cells.push(row[c] !== undefined ? row[c] : null);
      rows.push({ name: cellText(name).trim(), idNumber: normId(idNumber), cells });
    }
  }

  return { headers: headers || [], rows };
}

export function colIndex(headers, label) {
  return headers.findIndex((h) => h === label);
}
