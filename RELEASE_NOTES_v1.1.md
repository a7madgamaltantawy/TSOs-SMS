# MED-TSO SMS v1.1 — KPI Refinement

## Included
- Centralised KPI defaults in `Constants.js`.
- Added `KPI Guide` sheet with transparent definitions and formulas.
- Unified TSO ranking methodology: 40% attendance, 30% TC coverage, 30% task participation.
- Added member activity ranking: 40% attendance, 40% TC membership, 20% leadership.
- Added configurable leadership scores for Chairman, Vice-Chairman, Technical Secretariat and SPOC.
- Replaced Tasks with Gaps with an actionable Task Issues table.
- Committee Health now uses attendance, active-TSO coverage and task activity.
- Removed Critical Tasks and Critical Risks from the Steering Dashboard.
- Improved dashboard tables and horizontal ranking charts.
- TC coverage continues to include active TSOs only.

## Deployment
1. Create/check out branch `feature/v1.1-kpi-refinement`.
2. Replace/add the supplied Apps Script files.
3. Run `clasp push`.
4. Open the spreadsheet and run **MED-TSO SMS → KPIs → Refresh KPIs & Dashboard** once.
5. Review `KPI_Settings`, then refresh again after changing any setting.
