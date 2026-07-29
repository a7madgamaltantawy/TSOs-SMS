/**
 * KPIEngine.gs
 * Header-driven KPI engine for the MED-TSO Secretariat workbook.
 *
 * The script deliberately looks up columns by header names/synonyms rather
 * than assuming a rigid column order. This makes it safer when sheets evolve.
 */

const KPI_SHEETS_ = {
  SETTINGS: 'KPI_Settings',
  RANKING: 'KPI_TSO_Ranking',
  ATTENDANCE: 'KPI_Attendance',
  COVERAGE: 'KPI_TC_Coverage',
  TASKS: 'KPI_Task_Statistics',
  GAPS: 'KPI_Gap_Analysis',
  DASHBOARD: 'Dashboard'
};

function refreshAllKPIs() {
  const ui = SpreadsheetApp.getUi();
  try {
    initialiseKPISettings();
    const db = buildKPIDatabase_();

    const attendance = calculateAttendanceKPI_(db);
    const coverage = calculateTCCoverageKPI_(db);
    const tasks = calculateTaskKPI_(db);
    const ranking = calculateTSORanking_(db, attendance, coverage, tasks);
    const gaps = calculateGapAnalysis_(db);

    writeAttendanceKPI_(attendance);
    writeCoverageKPI_(coverage);
    writeTaskKPI_(tasks);
    writeRankingKPI_(ranking);
    writeGapKPI_(gaps);
    buildDashboard_(db, attendance, coverage, tasks, ranking, gaps);

    SpreadsheetApp.flush();
    ui.alert(
      'KPI refresh completed',
      'The KPI sheets and Dashboard were refreshed successfully.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('KPI refresh failed', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

function initialiseKPISettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(KPI_SHEETS_.SETTINGS);
  if (!sh) sh = ss.insertSheet(KPI_SHEETS_.SETTINGS);

  const defaults = [
    ['Setting', 'Value', 'Description'],
    ['Attendance Weight', 40, 'Weight in final TSO ranking (%)'],
    ['TC Coverage Weight', 20, 'Weight in final TSO ranking (%)'],
    ['Task Participation Weight', 25, 'Weight in final TSO ranking (%)'],
    ['Task Ownership Weight', 15, 'Weight in final TSO ranking (%)'],
    ['Target Technical Committees', 5, 'Expected number of Technical Committees'],
    ['Minimum Task Participation', 1, 'Assignments needed for full participation score'],
    ['Include Completed Meetings Only', 'Yes', 'Use only completed meetings in attendance KPIs'],
    ['Dashboard Top N', 10, 'Number of TSOs shown in the ranking section'],
    ['Last Refreshed', '', 'Updated automatically']
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, defaults.length, defaults[0].length).setValues(defaults);
    formatOutputSheet_(sh, 3);
  } else {
    const existing = sh.getRange(1, 1, sh.getLastRow(), Math.max(3, sh.getLastColumn()))
      .getDisplayValues()
      .map(r => normalizeKPI_(r[0]));
    defaults.slice(1).forEach(row => {
      if (existing.indexOf(normalizeKPI_(row[0])) < 0) sh.appendRow(row);
    });
  }
}

function openDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KPI_SHEETS_.DASHBOARD);
  if (!sh) {
    refreshAllKPIs();
    return;
  }
  ss.setActiveSheet(sh);
}

function buildKPIDatabase_() {
  const names = [
    'TSOs', 'Members', 'TCs', 'TCs Members', 'TC Members',
    'Meetings', 'Attendance', 'Tasks', 'Task Force Members'
  ];

  const db = {};
  names.forEach(name => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (sh) db[name] = readTable_(sh);
  });

  requireTable_(db, 'TSOs');
  requireTable_(db, 'Members');
  requireTable_(db, 'TCs');
  requireTable_(db, 'Meetings');
  requireTable_(db, 'Tasks');
  requireTable_(db, 'Task Force Members');

  db.tcMembers = db['TCs Members'] || db['TC Members'] || {headers: [], rows: []};
  db.settings = readKPISettings_();
  db.activeTSOs = getActiveTSOs_(db);
  db.activeMembers = getActiveMembers_(db);

  return db;
}

function readTable_(sh) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    return {name: sh.getName(), headers: [], rows: []};
  }

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(v => String(v || '').trim());
  const rows = values.slice(1).filter(r => r.some(v => String(v || '').trim() !== ''));

  return {name: sh.getName(), headers: headers, rows: rows};
}

function requireTable_(db, name) {
  if (!db[name]) throw new Error('Missing required sheet: ' + name);
}

function col_(table, synonyms, required) {
  const wanted = synonyms.map(normalizeHeader_);

  // First try an exact normalized match.
  let idx = table.headers.findIndex(h =>
    wanted.indexOf(normalizeHeader_(h)) >= 0
  );

  // Then allow safe partial matches such as:
  // "Attendance Status", "Present (X)", "Member Attendance", etc.
  if (idx < 0) {
    idx = table.headers.findIndex(header => {
      const normalized = normalizeHeader_(header);
      return wanted.some(candidate =>
        normalized === candidate ||
        normalized.indexOf(candidate + ' ') === 0 ||
        normalized.endsWith(' ' + candidate) ||
        normalized.indexOf(' ' + candidate + ' ') >= 0
      );
    });
  }

  if (idx < 0 && required) {
    throw new Error(
      'Sheet "' + table.name + '" is missing a required column. Expected one of: ' +
      synonyms.join(', ') +
      '. Existing headers: ' + table.headers.join(' | ')
    );
  }
  return idx;
}

function normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[_\-\/]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function normalizeKPI_(value) {
  return String(value || '').trim().toLowerCase();
}

function truthy_(value) {
  const normalized = normalizeKPI_(value);
  return [
    'yes', 'y', 'true', '1', 'active', 'present', 'x',
    'attended', 'participated', 'p', 'checked', '✓', '✔'
  ].indexOf(normalized) >= 0;
}

function statusActive_(value) {
  const n = normalizeKPI_(value);
  return n === '' || n === 'active' || n === 'yes';
}

function readKPISettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KPI_SHEETS_.SETTINGS);
  const values = sh.getDataRange().getDisplayValues();
  const settings = {};
  values.slice(1).forEach(r => {
    if (r[0]) settings[normalizeKPI_(r[0])] = r[1];
  });
  return settings;
}

function settingNumber_(db, name, fallback) {
  const n = Number(db.settings[normalizeKPI_(name)]);
  return isFinite(n) ? n : fallback;
}

function settingYes_(db, name, fallback) {
  const value = db.settings[normalizeKPI_(name)];
  if (value === undefined || value === '') return fallback;
  return truthy_(value);
}

function getActiveTSOs_(db) {
  const t = db['TSOs'];
  const code = col_(t, ['TSO Code', 'TSO', 'Organisation Code', 'Organization Code', 'Name'], true);
  const status = col_(t, ['Status', 'Active Status', 'Is Active'], false);
  const country = col_(t, ['Country Code', 'Country'], false);

  return t.rows
    .filter(r => status < 0 || statusActive_(r[status]))
    .map(r => ({
      code: String(r[code] || '').trim(),
      country: country >= 0 ? String(r[country] || '').trim() : ''
    }))
    .filter(x => x.code);
}

function getActiveMembers_(db) {
  const t = db['Members'];
  const id = col_(t, ['Member ID', 'ID'], true);
  const name = col_(t, ['Display Name', 'Member Name', 'Name', 'Full Name'], true);
  const tso = col_(t, ['TSO Code', 'TSO', 'Organisation', 'Organization'], true);
  const status = col_(t, ['Status', 'Member Status'], false);

  return t.rows
    .filter(r => status < 0 || statusActive_(r[status]))
    .map(r => ({
      id: String(r[id] || '').trim(),
      name: String(r[name] || '').trim(),
      tso: String(r[tso] || '').trim()
    }))
    .filter(x => x.id && x.tso);
}

function calculateAttendanceKPI_(db) {
  const meetings = db['Meetings'];
  const attendance = db['Attendance'];
  const result = {};

  db.activeTSOs.forEach(t => {
    result[t.code] = {tso: t.code, invited: 0, attended: 0, rate: 0};
  });

  if (!attendance || attendance.rows.length === 0) {
    return Object.keys(result).map(k => result[k]);
  }

  const mtgIdM = col_(meetings, ['Meeting ID', 'ID'], true);
  const statusM = col_(meetings, ['Status', 'Meeting Status'], false);
  const countCompletedOnly = settingYes_(db, 'Include Completed Meetings Only', true);

  const eligibleMeetingIds = {};
  meetings.rows.forEach(r => {
    if (!countCompletedOnly || statusM < 0 || normalizeKPI_(r[statusM]) === 'completed') {
      eligibleMeetingIds[String(r[mtgIdM] || '').trim()] = true;
    }
  });

  const mtgIdA = col_(attendance, ['Meeting ID', 'Meeting', 'Meeting Code'], true);
  const memberIdA = col_(attendance, ['Member ID', 'Expert ID', 'Participant ID'], false);
  const tsoA = col_(attendance, ['TSO Code', 'TSO', 'Organisation', 'Organization'], false);
  const presentA = col_(attendance, [
    'Present',
    'Attendance',
    'Attended',
    'Attendance Status',
    'Presence',
    'Present X',
    'Present?',
    'Participation',
    'Status'
  ], true);
  const invitedA = col_(attendance, ['Invited', 'Eligible', 'Expected'], false);

  const memberToTSO = {};
  db.activeMembers.forEach(m => memberToTSO[m.id] = m.tso);

  const meetingTso = {};
  attendance.rows.forEach(r => {
    const meetingId = String(r[mtgIdA] || '').trim();
    if (!eligibleMeetingIds[meetingId]) return;

    const tso = tsoA >= 0
      ? String(r[tsoA] || '').trim()
      : memberToTSO[String(r[memberIdA] || '').trim()];

    if (!tso || !result[tso]) return;

    const key = meetingId + '||' + tso;
    if (!meetingTso[key]) meetingTso[key] = {invited: false, attended: false};

    const explicitlyInvited = invitedA < 0 ? true : truthy_(r[invitedA]);
    if (explicitlyInvited) meetingTso[key].invited = true;
    if (truthy_(r[presentA])) {
      meetingTso[key].invited = true;
      meetingTso[key].attended = true;
    }
  });

  Object.keys(meetingTso).forEach(key => {
    const tso = key.split('||')[1];
    const rec = meetingTso[key];
    if (rec.invited) result[tso].invited++;
    if (rec.attended) result[tso].attended++;
  });

  Object.keys(result).forEach(tso => {
    const rec = result[tso];
    rec.rate = rec.invited ? rec.attended / rec.invited : 0;
  });

  return Object.keys(result).map(k => result[k]);
}

function calculateTCCoverageKPI_(db) {
  const memberships = db.tcMembers;
  const result = {};
  const target = settingNumber_(db, 'Target Technical Committees', 5);

  db.activeTSOs.forEach(t => {
    result[t.code] = {tso: t.code, committees: {}, covered: 0, target: target, rate: 0};
  });

  if (!memberships || memberships.rows.length === 0) {
    return Object.keys(result).map(k => result[k]);
  }

  const memberId = col_(memberships, ['Member ID', 'Expert ID'], false);
  const tsoCol = col_(memberships, ['TSO Code', 'TSO', 'Organisation', 'Organization'], false);
  const tcCol = col_(memberships, ['TC Name', 'Technical Committee', 'Committee', 'TC'], true);
  const activeCol = col_(memberships, ['Active', 'Status', 'Membership Status'], false);

  const memberToTSO = {};
  db.activeMembers.forEach(m => memberToTSO[m.id] = m.tso);

  memberships.rows.forEach(r => {
    if (activeCol >= 0 && !statusActive_(r[activeCol])) return;
    const tso = tsoCol >= 0
      ? String(r[tsoCol] || '').trim()
      : memberToTSO[String(r[memberId] || '').trim()];
    const tc = String(r[tcCol] || '').trim();
    if (!tso || !tc || !result[tso]) return;
    result[tso].committees[tc] = true;
  });

  Object.keys(result).forEach(tso => {
    const rec = result[tso];
    rec.covered = Object.keys(rec.committees).length;
    rec.rate = Math.min(1, rec.covered / Math.max(1, rec.target));
  });

  return Object.keys(result).map(k => result[k]);
}

function calculateTaskKPI_(db) {
  const assignments = db['Task Force Members'];
  const tasks = db['Tasks'];
  const result = {};

  db.activeTSOs.forEach(t => {
    result[t.code] = {
      tso: t.code,
      activeAssignments: 0,
      ownedTasks: 0,
      participationRate: 0,
      ownershipRate: 0
    };
  });

  const minimum = settingNumber_(db, 'Minimum Task Participation', 1);
  const memberToTSO = {};
  db.activeMembers.forEach(m => memberToTSO[m.id] = m.tso);

  const taskId = col_(assignments, ['Task ID', 'Task'], true);
  const memberId = col_(assignments, ['Member ID', 'Expert ID', 'Assignee ID'], false);
  const tsoCol = col_(assignments, ['TSO Code', 'TSO', 'Organisation', 'Organization'], false);
  const role = col_(assignments, ['Role', 'Assignment Role'], false);
  const status = col_(assignments, ['Status', 'Assignment Status', 'Active'], false);

  const uniqueAssignments = {};
  const uniqueOwners = {};

  assignments.rows.forEach(r => {
    if (status >= 0 && !statusActive_(r[status])) return;

    const tso = tsoCol >= 0
      ? String(r[tsoCol] || '').trim()
      : memberToTSO[String(r[memberId] || '').trim()];
    const tid = String(r[taskId] || '').trim();
    if (!tso || !tid || !result[tso]) return;

    uniqueAssignments[tso + '||' + tid] = true;
    if (role >= 0 && /owner|lead|coordinator/i.test(String(r[role] || ''))) {
      uniqueOwners[tso + '||' + tid] = true;
    }
  });

  Object.keys(uniqueAssignments).forEach(k => {
    const tso = k.split('||')[0];
    result[tso].activeAssignments++;
  });
  Object.keys(uniqueOwners).forEach(k => {
    const tso = k.split('||')[0];
    result[tso].ownedTasks++;
  });

  Object.keys(result).forEach(tso => {
    const rec = result[tso];
    rec.participationRate = Math.min(1, rec.activeAssignments / Math.max(1, minimum));
    rec.ownershipRate = rec.activeAssignments
      ? Math.min(1, rec.ownedTasks / rec.activeAssignments)
      : 0;
  });

  return Object.keys(result).map(k => result[k]);
}

function calculateTSORanking_(db, attendance, coverage, tasks) {
  const weights = {
    attendance: settingNumber_(db, 'Attendance Weight', 40),
    coverage: settingNumber_(db, 'TC Coverage Weight', 20),
    tasks: settingNumber_(db, 'Task Participation Weight', 25),
    ownership: settingNumber_(db, 'Task Ownership Weight', 15)
  };
  const totalWeight = weights.attendance + weights.coverage + weights.tasks + weights.ownership || 100;

  const a = indexBy_(attendance, 'tso');
  const c = indexBy_(coverage, 'tso');
  const t = indexBy_(tasks, 'tso');

  const rows = db.activeTSOs.map(item => {
    const tso = item.code;
    const attendanceRate = a[tso] ? a[tso].rate : 0;
    const coverageRate = c[tso] ? c[tso].rate : 0;
    const participationRate = t[tso] ? t[tso].participationRate : 0;
    const ownershipRate = t[tso] ? t[tso].ownershipRate : 0;

    const score = 100 * (
      attendanceRate * weights.attendance +
      coverageRate * weights.coverage +
      participationRate * weights.tasks +
      ownershipRate * weights.ownership
    ) / totalWeight;

    return {
      tso: tso,
      attendanceRate: attendanceRate,
      coverageRate: coverageRate,
      participationRate: participationRate,
      ownershipRate: ownershipRate,
      score: score,
      trafficLight: score >= 70 ? 'Green' : score >= 45 ? 'Amber' : 'Red'
    };
  });

  rows.sort((x, y) => y.score - x.score || x.tso.localeCompare(y.tso));
  rows.forEach((r, i) => r.rank = i + 1);
  return rows;
}

function calculateGapAnalysis_(db) {
  const tasks = db['Tasks'];
  const assignments = db['Task Force Members'];
  const gaps = [];

  const idT = col_(tasks, ['Task ID', 'ID'], true);
  const tcT = col_(tasks, ['TC Name', 'Technical Committee', 'Committee', 'TC'], false);
  const nameT = col_(tasks, ['Task Name', 'Name', 'Task'], true);
  const statusT = col_(tasks, ['Status', 'Task Status'], false);
  const dueT = col_(tasks, ['Due Date', 'Deadline'], false);

  const taskA = col_(assignments, ['Task ID', 'Task'], true);
  const roleA = col_(assignments, ['Role', 'Assignment Role'], false);
  const statusA = col_(assignments, ['Status', 'Assignment Status', 'Active'], false);

  const byTask = {};
  assignments.rows.forEach(r => {
    if (statusA >= 0 && !statusActive_(r[statusA])) return;
    const tid = String(r[taskA] || '').trim();
    if (!tid) return;
    if (!byTask[tid]) byTask[tid] = {assignees: 0, owners: 0};
    byTask[tid].assignees++;
    if (roleA >= 0 && /owner|lead|coordinator/i.test(String(r[roleA] || ''))) {
      byTask[tid].owners++;
    }
  });

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  tasks.rows.forEach(r => {
    const taskId = String(r[idT] || '').trim();
    const status = statusT >= 0 ? String(r[statusT] || '').trim() : '';
    if (normalizeKPI_(status) === 'completed') return;

    const rec = byTask[taskId] || {assignees: 0, owners: 0};
    const due = dueT >= 0 && r[dueT] ? new Date(r[dueT]) : null;
    const overdue = due && !isNaN(due.getTime()) && due < today;

    const issues = [];
    if (rec.assignees === 0) issues.push('No active assignee');
    if (rec.owners === 0) issues.push('No owner/lead');
    if (overdue) issues.push('Overdue');

    if (issues.length) {
      gaps.push({
        tc: tcT >= 0 ? String(r[tcT] || '').trim() : '',
        taskId: taskId,
        taskName: String(r[nameT] || '').trim(),
        status: status,
        dueDate: due && !isNaN(due.getTime()) ? due : '',
        assignees: rec.assignees,
        owners: rec.owners,
        issue: issues.join('; ')
      });
    }
  });

  return gaps;
}

function indexBy_(rows, key) {
  const out = {};
  rows.forEach(r => out[r[key]] = r);
  return out;
}

function writeAttendanceKPI_(rows) {
  const data = [['TSO', 'Invited Meetings', 'Attended Meetings', 'Attendance Rate']];
  rows.sort((a, b) => b.rate - a.rate || a.tso.localeCompare(b.tso))
    .forEach(r => data.push([r.tso, r.invited, r.attended, r.rate]));
  writeTableToSheet_(KPI_SHEETS_.ATTENDANCE, data, [4]);
}

function writeCoverageKPI_(rows) {
  const data = [['TSO', 'TCs Covered', 'Target TCs', 'Coverage Rate', 'Covered Committees']];
  rows.sort((a, b) => b.rate - a.rate || a.tso.localeCompare(b.tso))
    .forEach(r => data.push([
      r.tso, r.covered, r.target, r.rate, Object.keys(r.committees).sort().join(', ')
    ]));
  writeTableToSheet_(KPI_SHEETS_.COVERAGE, data, [4]);
}

function writeTaskKPI_(rows) {
  const data = [[
    'TSO', 'Active Task Assignments', 'Owned Tasks',
    'Task Participation Rate', 'Task Ownership Rate'
  ]];
  rows.sort((a, b) => b.activeAssignments - a.activeAssignments || a.tso.localeCompare(b.tso))
    .forEach(r => data.push([
      r.tso, r.activeAssignments, r.ownedTasks,
      r.participationRate, r.ownershipRate
    ]));
  writeTableToSheet_(KPI_SHEETS_.TASKS, data, [4, 5]);
}

function writeRankingKPI_(rows) {
  const data = [[
    'Rank', 'TSO', 'Final Score', 'Attendance', 'TC Coverage',
    'Task Participation', 'Task Ownership', 'Traffic Light'
  ]];
  rows.forEach(r => data.push([
    r.rank, r.tso, r.score / 100, r.attendanceRate, r.coverageRate,
    r.participationRate, r.ownershipRate, r.trafficLight
  ]));
  const sh = writeTableToSheet_(KPI_SHEETS_.RANKING, data, [3, 4, 5, 6, 7]);
  applyTrafficLights_(sh, 8, 2, rows.length);
}

function writeGapKPI_(rows) {
  const data = [[
    'Technical Committee', 'Task ID', 'Task Name', 'Status',
    'Due Date', 'Active Assignees', 'Owners', 'Issue'
  ]];
  rows.forEach(r => data.push([
    r.tc, r.taskId, r.taskName, r.status, r.dueDate,
    r.assignees, r.owners, r.issue
  ]));
  writeTableToSheet_(KPI_SHEETS_.GAPS, data, []);
}

function writeTableToSheet_(name, data, percentColumns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear({contentsOnly: false});
  sh.getCharts().forEach(c => sh.removeChart(c));

  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  formatOutputSheet_(sh, data[0].length);

  percentColumns.forEach(col => {
    if (data.length > 1) sh.getRange(2, col, data.length - 1, 1).setNumberFormat('0.0%');
  });

  sh.autoResizeColumns(1, data[0].length);
  sh.setFrozenRows(1);
  return sh;
}

function formatOutputSheet_(sh, width) {
  if (sh.getLastRow() < 1) return;
  sh.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setBackground('#0f3d56')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function applyTrafficLights_(sh, column, startRow, count) {
  if (!count) return;
  const range = sh.getRange(startRow, column, count, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Green').setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Amber').setBackground('#fff2cc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Red').setBackground('#f4cccc').setRanges([range]).build()
  ];
  sh.setConditionalFormatRules(rules);
}
