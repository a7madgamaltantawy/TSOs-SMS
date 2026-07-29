#!/usr/bin/env bash
set -euo pipefail
BRANCH="feature/steering-dashboard-foundation"
git checkout develop
git pull
git checkout -b "$BRANCH"
git add App.html Model.js SteeringDashboard.js README_v7_0_dashboard_foundation.md
git commit -m "Add Steering Committee Dashboard foundation"
echo "Branch committed. Run: clasp push --force"
