# MED-TSO SMS v7.0 — Feature Branch 1

## Git branch
`feature/steering-dashboard-foundation`

## Scope
This branch adds the first safe v7 feature set only:

- Renames the Home navigation item to **Steering Dashboard**.
- Adds the **Steering Committee Dashboard** landing page.
- Adds six live KPI cards.
- Adds Technical Committee participation-health bars.
- Adds critical/overdue active-task exceptions.
- Adds meetings scheduled in the next 60 days.
- Adds a read-only backend service in `SteeringDashboard.js`.
- Removes no sheets and changes no spreadsheet records.
- Does not include an AI summary.

## Deploy
```bash
cd ~/TSOs-SMS
git checkout develop
git pull
git checkout -b feature/steering-dashboard-foundation

# Copy the package files into this repository, then:
git add App.html Model.js SteeringDashboard.js README_v7_0_dashboard_foundation.md
git commit -m "Add Steering Committee Dashboard foundation"
clasp push --force
```

Reload the Google Sheet and reopen the Management App. The header should show `Version 7.0.0-alpha.1`.

## Smoke test
1. Management App opens without an error.
2. Steering Dashboard is the first page.
3. Six KPI cards display values.
4. Committee-health rows display for active TCs.
5. Critical-task and upcoming-meeting sections display either data or a clear empty state.
6. Members, Structure, Meetings, Attendance and KPIs still open.

## Rollback
```bash
git checkout develop
clasp push --force
```
