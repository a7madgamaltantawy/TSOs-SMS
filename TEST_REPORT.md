# v6.3.2 Validation Report

## Static validation
- `node --check` passed for Code.js, Model.js, Groups.js, KPIEngine.js, Dashboard.js and ParticipationGap.js.
- Embedded JavaScript extracted from App.html and passed `node --check`.
- Required UI controls verified: Team membership, Task assignment, member/task selectors, assignment handler and executive gap loader.

## Workbook smoke test
Using the bundled stable workbook:
- Active TSOs detected: 18.
- Active Technical Committees detected: 5.
- Active Tasks detected: 18.
- Task assignment records are present in `Task Force Members`.
- Sample calculated coverage from the source data is non-zero, confirming that the repaired engine can populate the dashboard without `Groups` or `Group Members`.

## Runtime boundary
Apps Script server execution and browser clicks must still be confirmed after `clasp push`, because this environment cannot execute Google Apps Script against the live Google Sheet.
