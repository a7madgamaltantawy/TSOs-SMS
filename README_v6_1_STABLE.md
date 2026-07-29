# MED-TSO SMS v6.1 — Stabilized Unified Groups Build

## Included fixes

- One clean `onOpen()` menu only.
- Removed menu items for model preparation, KPI settings, and late-joiner refresh.
- Unified Groups initialization and validation retained under Administration.
- Removed duplicate helper functions that were overriding `nextId_()` and causing:
  `TypeError: existingIds.forEach is not a function`.
- Meeting creation uses the unified `Groups` table.
- Meeting records retain compatibility with both legacy and v6.1 column names.
- Attendance roster loads active eligible members from `Group Members`.
- Existing KPI and Dashboard code is preserved.

## Deployment

From Terminal:

```bash
cd ~/MED-TSO-SMS

# Make a safety commit before replacing files
git add .
git commit -m "Backup before v6.1 stabilized build"

# Copy the files from this package into the project folder, then:
clasp push --force
```

Refresh Google Sheets with `Command + Shift + R`.

## Test sequence

1. MED-TSO SMS → Administration → Initialize Unified Groups.
2. Confirm the result reports groups and memberships.
3. MED-TSO SMS → Administration → Validate Unified Groups.
4. Open Management App.
5. Create a new meeting using a group with active members.
6. Open Attendance and select that new meeting.
7. Confirm eligible members appear, set statuses, and save attendance.
8. Refresh KPIs & Dashboard.

## Important

Old meetings created before Group IDs were introduced may not load an attendance roster. Create a new meeting through the v6.1 form, or populate its `Group ID`, `Group Name`, and `Group Type` fields manually.
