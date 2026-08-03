# HRIS Biologs Generator

Standalone browser tool (no backend, no install) that turns an Attendance
Report export into two deliverables:

1. **Flat file** — the `Detailed` sheet de-nested into one row per employee-day,
   with `Name` and `ID Number` added as columns A/B.
2. **HRIS Biologs import file** — an exact-fidelity rebuild of
   `HRISBiologsTemplate.xlsx` (same `Template` + `Sheet3` sheets, same `LogMode`
   named range and dropdown, same column widths/text formatting), populated with
   one `IN` row and one `OUT` row per scheduled working day.

Runs entirely client-side via [ExcelJS](https://github.com/exceljs/exceljs) —
uploaded data never leaves the browser.

## Design decisions (confirmed with the client)

- **The Biologs template is seeded from the *scheduled* shift, not the actual
  punches.** Column C/D times come from the Detailed sheet's `Shift` column
  (e.g. `08:00 AM To 05:00 PM`), not the `Biologs` column. We're dictating what
  the logs should say, not reporting what was actually scanned.
- **Exclusions:** any day whose `Shift` cell contains `REST DAY` or
  `UNPAID LEAVE` (case-insensitive) produces no rows — even if a time range is
  also present in that cell.
- **Overnight shifts:** if the shift's end time is earlier than (or equal to)
  its start time, the `OUT` row's date rolls to the next calendar day; `IN`
  keeps the row's date as-is.
- **Unmatched Biometric IDs:** surfaced as a validation warning (not blocking);
  those employee-days are simply skipped from the Biologs template export.
- **Two independent export buttons** — the flat file only needs the Attendance
  Report; the Biologs template additionally needs the Employee List Report
  (Employee ID → Biometric ID mapping).

## Project layout

```
src/core/       pure logic (parser, shift-time rules, mapping, validation, xlsx builders)
src/app/        UI wiring (app.js) + page shell (index.template.html)
src/test/       Node-based tests — run against real sample files + synthetic edge cases
vendor/         ExcelJS (vendored, no CDN dependency)
build.js        assembles vendor + core + app into the single self-contained index.html
```

## Build

```
node build.js
```

Rebuilds `index.html` from `src/`. Open the resulting file directly in a
browser — no server needed.

## Test

```
node src/test/synthetic.mjs   # hand-crafted edge cases: overnight rollover,
                               # REST DAY/UNPAID LEAVE exclusion, unmatched IDs,
                               # unparseable shift times — with assertions
node src/test/pipeline.mjs    # runs against the real sample files, dumps output
                               # for manual inspection
```

## Known gap

Both sample files provided so far use unrelated dummy ID schemes (attendance
export IDs like `ABC-00001` vs. employee list IDs like `RB-9058`), so the
"matched" path (an employee whose ID Number *does* resolve to a Biometric ID)
has only been exercised with synthetic data, not real production files. Run it
against a real matching pair before fully trusting the Biologs template output.
