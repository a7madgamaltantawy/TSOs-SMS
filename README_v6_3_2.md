# MED-TSO SMS v6.3.2 — Executive Gap Engine Fix

## Why the previous view was empty
The v6.3.1 dashboard expected `Groups` and `Group Members`, but the stable workbook stores the authoritative data in `TCs`, `TC Members`, `Tasks`, and `Task Force Members`. This release reads those source sheets directly, so the executive dashboard no longer depends on the optional unified-groups migration.

## Included
- Working committee → related tasks → missing TSOs analysis.
- Committee and task coverage based on all active TSOs.
- Executive metric cards, ranked committee chart, risk bands, and task drill-down.
- Visible in-panel error message if the backend fails.
- `Operational Team membership` renamed to `Team membership`.
- Task assignment panel restored on the Structure page.
- Task dropdown includes the related Technical Committee.
- Duplicate active task assignment protection retained.
- Gap Analysis sheet refresh retained.

## Deployment
```bash
clasp push --force
```
Refresh Google Sheets with `Command + Shift + R`, reopen the Management App, and open **KPIs**.

No unified-groups setup is required for the executive gap dashboard.

## Test sequence
1. Open **KPIs** and confirm the metric cards populate.
2. Confirm all five active Technical Committees appear.
3. Confirm committee tasks are listed with coverage and missing TSO badges.
4. Use the committee filter.
5. Open **Structure** and confirm **Team membership** and **Task assignment** are visible.
6. Assign a member to a task and refresh the KPI view.
7. Confirm that member's TSO disappears from the task's missing list.
8. Click **Create / refresh Gap Analysis sheet**.
