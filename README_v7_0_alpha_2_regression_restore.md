# MED-TSO SMS v7.0.0-alpha.2 — Regression Restore

Branch: `fix/restore-task-assignment-member-intelligence`

Restored without changing the Steering Committee Dashboard data engine:

- Structure → **Team membership** naming
- Structure → **Task assignment** panel
- Member search cards → Technical Committee memberships
- Member search cards → Assigned tasks
- Member search cards → Primary phone when available

Files changed:

- `App.html`
- `Model.js` (version only)

Backend task assignment and member participation fields already existed in `Code.js`; no data migration is required.
