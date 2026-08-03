// core/shift.js — turns a Detailed-sheet "Shift" cell (e.g. "08:00 AM To 05:00 PM")
// into the IN/OUT log the Biologs template wants.
//
// Decision (confirmed with the client): the Biologs template is seeded from the
// SCHEDULED shift, not the actual Biologs punches — we are dictating what the
// logs should say, not reporting what was actually scanned.

const TIME_RE = /(\d{1,2}):(\d{2})\s*(AM|PM)/gi;

// Extracts every "h:mm AM/PM" token in order, each as { raw, mins } where mins
// is minutes-from-midnight and raw is reformatted to a consistent "hh:mm AM/PM".
export function extractTimeTokens(text) {
  const out = [];
  const re = new RegExp(TIME_RE.source, 'gi');
  let m;
  while ((m = re.exec(text || ''))) {
    let h = parseInt(m[1], 10) % 12;
    const ampm = m[3].toUpperCase();
    if (ampm === 'PM') h += 12;
    const mins = h * 60 + parseInt(m[2], 10);
    out.push({ raw: formatMinutes(mins), mins });
  }
  return out;
}

export function formatMinutes(mins) {
  const ampm = mins >= 720 ? 'PM' : 'AM';
  let h = Math.floor(mins / 60) % 12;
  if (h === 0) h = 12;
  const m = mins % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')} ${ampm}`;
}

export function isExcludedShift(shiftText) {
  const s = (shiftText || '').toUpperCase();
  return s.includes('REST DAY') || s.includes('UNPAID LEAVE');
}

// Returns one of:
//   { excluded: true }                                          — Rest Day / Unpaid Leave, no logs generated
//   { excluded: false, error: '...' }                           — couldn't find a start+end time to seed from
//   { excluded: false, inTime, outTime, outDayOffset }          — outDayOffset is 0 or 1 (overnight rollover)
export function parseShift(shiftText) {
  const s = (shiftText || '').trim();
  if (isExcludedShift(s)) return { excluded: true };

  const tokens = extractTimeTokens(s);
  if (tokens.length < 2) {
    return { excluded: false, error: `Could not find a start and end time in Shift cell: "${s || '(blank)'}"` };
  }
  const [start, end] = tokens;
  const outDayOffset = end.mins <= start.mins ? 1 : 0; // overnight shift: OUT lands on the next calendar day
  return { excluded: false, inTime: start.raw, outTime: end.raw, outDayOffset };
}
