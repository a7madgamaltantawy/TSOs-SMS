# MED-TSO SMS v6.3 — Structure and Gap Analysis

## Added
- Renamed **Operational Team Membership** to **Team Membership**.
- Added Task Assignment to the Structure page.
- Members can be assigned to active tasks as Owner, Contributor, Reviewer, Observer or Assignee.
- Added KPI Gap Analysis for Technical Committees, Teams, Task Forces and active Tasks.
- Added an optional generated `Gap Analysis` sheet.

## Deployment
1. Copy all files into the clasp project.
2. Run `clasp push`.
3. Refresh the spreadsheet and reopen the Management App.
4. In Apps Script, run `setupSmsV63()` once if the `Task Force Members` sheet does not already exist.

## Test checklist
- Structure page displays Team Membership.
- Task dropdown and active-member dropdown load.
- Task assignment saves and duplicate active assignment is rejected.
- Member search reflects assigned task.
- KPI Gap Analysis loads all four structure types.
- Gap Analysis sheet is created/refreshed.
