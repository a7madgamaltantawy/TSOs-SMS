
/**
 * MED-TSO Secretariat Management System
 * Core application + KPI menu integration
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('MED-TSO SMS')
    .addItem('Open Management App', 'showApp')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Administration')
        .addItem('Initialize Unified Groups', 'initializeUnifiedGroups')
        .addItem('Validate Unified Groups', 'validateUnifiedGroups')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('KPIs')
        .addItem('Refresh KPIs & Dashboard', 'refreshAllKPIs')
        .addItem('Open Dashboard', 'openDashboard')
    )
    .addToUi();
}

function showApp() {
  const html = HtmlService.createHtmlOutputFromFile('App')
    .setTitle('MED-TSO Secretariat Management');
  SpreadsheetApp.getUi().showSidebar(html);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function rows_(sheetName) {
  const sh = sheet_(sheetName);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
}

function normalize_(value) {
  return String(value || '').trim().toLowerCase();
}

function getFormOptions() {
  return {
    tsos: rows_('TSOs')
      .filter(r => normalize_(r[6]) === 'active')
      .map(r => r[1])
      .filter(String),

    tcs: rows_('TCs')
      .map(r => r[1])
      .filter(String),

    workstreams: rows_('Workstreams')
      .filter(r => normalize_(r[4]) === 'active')
      .map(r => r[1])
      .filter(String),

    tasks: rows_('Tasks')
      .filter(r => normalize_(r[10]) !== 'completed')
      .map(r => ({ id: r[0], name: r[5] }))
      .filter(x => x.id && x.name),

    members: rows_('Members')
      .filter(r => normalize_(r[10]) === 'active')
      .map(r => ({ id: r[0], name: r[3], tso: r[7], email: r[4] }))
      .filter(x => x.id && x.name)
      .sort((a, b) => a.name.localeCompare(b.name))
  };
}

function nextId_(sheetName, prefix, idColumn, digits) {
  const sh = sheet_(sheetName);
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, idColumn, last - 1, 1).getDisplayValues().flat().forEach(id => {
      const m = String(id).match(/(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return prefix + '-' + String(max + 1).padStart(digits || 4, '0');
}

function findRow_(sheetName, column, value) {
  const sh = sheet_(sheetName);
  if (sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, column, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const idx = vals.findIndex(v => String(v).trim() === String(value).trim());
  return idx < 0 ? null : idx + 2;
}

function lateJoiner_(joinedDate) {
  if (!joinedDate) return {phase: 'Unknown', late: 'Unknown'};
  const cfg = rows_('Configuration');
  const thresholdRow = cfg.find(r => r[0] === 'Late Joiner Threshold');
  const threshold = thresholdRow && thresholdRow[1]
    ? new Date(thresholdRow[1])
    : new Date('2025-07-01');
  const joined = new Date(joinedDate);

  return joined > threshold
    ? {phase: 'Late joiner', late: 'Yes'}
    : {phase: 'Project start / early phase', late: 'No'};
}

function validateMember_(data, existingId) {
  if (!String(data.firstName || '').trim()) throw new Error('First name is required.');
  if (!String(data.lastName || '').trim()) throw new Error('Last name is required.');
  if (!String(data.primaryEmail || '').trim()) throw new Error('Primary email is required.');

  if (!String(data.tsoCode || '').trim() &&
      String(data.memberType || '').indexOf('Secretariat') < 0) {
    throw new Error('TSO is required for TSO members.');
  }

  const email = normalize_(data.primaryEmail);
  const duplicate = rows_('Members').find(r =>
    normalize_(r[4]) === email &&
    String(r[0]) !== String(existingId || '')
  );

  if (duplicate) {
    throw new Error('This primary email already belongs to ' + duplicate[3] + '.');
  }
}

function saveMember(data) {
  validateMember_(data, null);

  const sh = sheet_('Members');
  const id = nextId_('Members', 'MEM', 1, 4);
  const display = [data.firstName, data.lastName].filter(Boolean).join(' ').trim();
  const join = lateJoiner_(data.joinedDate);

  sh.appendRow([
    id,
    data.firstName || '',
    data.lastName || '',
    display,
    data.primaryEmail || '',
    data.secondaryEmail || '',
    data.memberType || 'TSO Member',
    data.tsoCode || '',
    data.countryCode || '',
    data.role || '',
    data.status || 'Active',
    data.joinedDate || '',
    join.phase,
    join.late,
    'Manual entry via App.html',
    data.notes || ''
  ]);

  sh.getRange(sh.getLastRow(), memberPhoneColumn_()).setValue(data.primaryPhone || '');

  syncSecretariatMember_(id, data);
  return {id: id, message: 'Member successfully added.'};
}


function memberPhoneColumn_() {
  const sh = sheet_('Members');
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const accepted = ['primary phone', 'phone', 'mobile', 'mobile phone', 'phone number'];
  const index = headers.findIndex(h => accepted.indexOf(normalize_(h)) >= 0);
  if (index >= 0) return index + 1;

  const column = lastColumn + 1;
  sh.getRange(1, column).setValue('Primary Phone');
  return column;
}

function memberParticipationIndex_() {
  const committees = {};
  rows_('TC Members').forEach(r => {
    if (normalize_(r[7]) === 'inactive') return;
    const memberId = String(r[1] || '').trim();
    const tcName = String(r[5] || '').trim();
    if (!memberId || !tcName) return;
    if (!committees[memberId]) committees[memberId] = [];
    if (committees[memberId].indexOf(tcName) < 0) committees[memberId].push(tcName);
  });

  const tasks = {};
  rows_('Task Force Members').forEach(r => {
    if (normalize_(r[7]) === 'inactive') return;
    const memberId = String(r[3] || '').trim();
    const taskName = String(r[2] || '').trim();
    if (!memberId || !taskName) return;
    if (!tasks[memberId]) tasks[memberId] = [];
    if (tasks[memberId].indexOf(taskName) < 0) tasks[memberId].push(taskName);
  });

  return {committees: committees, tasks: tasks};
}

function getMembers(filters) {
  filters = filters || {};
  const q = normalize_(filters.query);
  const tso = String(filters.tso || '');
  const status = String(filters.status || '');

  const phoneColumn = memberPhoneColumn_();
  const participation = memberParticipationIndex_();

  return rows_('Members')
    .filter(r => {
      const haystack = normalize_([r[0], r[3], r[4], r[7], r[9]].join(' '));
      return (!q || haystack.indexOf(q) >= 0) &&
             (!tso || r[7] === tso) &&
             (!status || r[10] === status);
    })
    .map(r => ({
      id: r[0],
      displayName: r[3],
      email: r[4],
      tsoCode: r[7],
      role: r[9],
      status: r[10],
      memberType: r[6],
      primaryPhone: r[phoneColumn - 1] || '',
      technicalCommittees: participation.committees[r[0]] || [],
      assignedTasks: participation.tasks[r[0]] || []
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
}

function getMember(memberId) {
  const row = findRow_('Members', 1, memberId);
  if (!row) throw new Error('Member not found.');

  const sh = sheet_('Members');
  const phoneColumn = memberPhoneColumn_();
  const r = sh.getRange(row, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const participation = memberParticipationIndex_();

  return {
    memberId: r[0],
    firstName: r[1],
    lastName: r[2],
    displayName: r[3],
    primaryEmail: r[4],
    secondaryEmail: r[5],
    memberType: r[6],
    tsoCode: r[7],
    countryCode: r[8],
    role: r[9],
    status: r[10],
    joinedDate: r[11],
    joinedPhase: r[12],
    lateJoiner: r[13],
    notes: r[15],
    primaryPhone: r[phoneColumn - 1] || '',
    technicalCommittees: participation.committees[r[0]] || [],
    assignedTasks: participation.tasks[r[0]] || []
  };
}

function updateMember(data) {
  if (!data.memberId) throw new Error('Member ID is required.');

  validateMember_(data, data.memberId);

  const row = findRow_('Members', 1, data.memberId);
  if (!row) throw new Error('Member not found.');

  const display = [data.firstName, data.lastName].filter(Boolean).join(' ').trim();
  const join = lateJoiner_(data.joinedDate);
  const sh = sheet_('Members');

  sh.getRange(row, 2, 1, 15).setValues([[
    data.firstName || '',
    data.lastName || '',
    display,
    data.primaryEmail || '',
    data.secondaryEmail || '',
    data.memberType || 'TSO Member',
    data.tsoCode || '',
    data.countryCode || '',
    data.role || '',
    data.status || 'Active',
    data.joinedDate || '',
    join.phase,
    join.late,
    'Updated via App.html on ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    data.notes || ''
  ]]);

  sh.getRange(row, memberPhoneColumn_()).setValue(data.primaryPhone || '');
  syncSecretariatMember_(data.memberId, data);
  return {id: data.memberId, message: 'Member successfully updated.'};
}

function softDeleteMember(memberId) {
  const row = findRow_('Members', 1, memberId);
  if (!row) throw new Error('Member not found.');

  sheet_('Members').getRange(row, 11).setValue('Inactive');

  const secRow = findRow_('Secretariat Members', 2, memberId);
  if (secRow) {
    sheet_('Secretariat Members').getRange(secRow, 6).setValue('Inactive');
  }

  return {
    id: memberId,
    message: 'Member set to Inactive. Historical records were preserved.'
  };
}

function syncSecretariatMember_(memberId, data) {
  const isSecretariat =
    String(data.memberType || '').indexOf('Secretariat') >= 0;

  const sh = sheet_('Secretariat Members');
  const existingRow = findRow_('Secretariat Members', 2, memberId);
  const display = [data.firstName, data.lastName].filter(Boolean).join(' ').trim();

  if (isSecretariat && existingRow) {
    sh.getRange(existingRow, 3, 1, 6).setValues([[
      display,
      data.primaryEmail || '',
      data.role || '',
      data.status || 'Active',
      data.joinedDate || '',
      'Updated via App.html'
    ]]);
  } else if (isSecretariat && !existingRow) {
    sh.appendRow([
      nextId_('Secretariat Members', 'SEC', 1, 4),
      memberId,
      display,
      data.primaryEmail || '',
      data.role || '',
      data.status || 'Active',
      data.joinedDate || '',
      'Manual entry via App.html'
    ]);
  } else if (!isSecretariat && existingRow) {
    sh.getRange(existingRow, 6).setValue('Inactive');
    sh.getRange(existingRow, 8)
      .setValue('Removed from Secretariat role via App.html');
  }
}

function saveMeeting(data) {
  if (!data.tcName || !data.meetingName || !data.meetingDate) {
    throw new Error('TC, meeting name and date are required.');
  }

  const sh = sheet_('Meetings');
  const tc = rows_('TCs').find(r => r[1] === data.tcName) || ['', data.tcName];
  const id = nextId_('Meetings', 'MTG', 1, 4);

  sh.appendRow([
    id,
    tc[0],
    tc[1],
    data.meetingName,
    data.meetingDate,
    data.meetingType || 'Regular',
    data.description || '',
    data.status || 'Planned',
    'Manual entry via App.html',
    data.status === 'Cancelled' ? 'No' : 'Yes',
    'Inactive-TSO participants must not be invited or counted.'
  ]);

  return id;
}

function saveTask(data) {
  if (!data.taskName) throw new Error('Task name is required.');

  const sh = sheet_('Tasks');
  const tc = rows_('TCs').find(r => r[1] === data.tcName) ||
    ['', data.tcName || ''];
  const ws = rows_('Workstreams').find(r => r[1] === data.workstream) ||
    ['', data.workstream || ''];
  const id = nextId_('Tasks', 'TSK', 1, 4);

  sh.appendRow([
    id,
    tc[0],
    tc[1],
    ws[0],
    ws[1],
    data.taskName,
    data.description || '',
    data.originType || 'Manual',
    data.originMeetingId || '',
    data.originReference || '',
    data.status || 'Not Started',
    data.priority || 'Normal',
    data.dueDate || '',
    new Date(),
    Session.getActiveUser().getEmail() || 'App user',
    '',
    new Date(),
    'Manual entry via App.html',
    ''
  ]);

  return id;
}

function saveTaskAssignment(data) {
  if (!data.taskId || !data.memberId) {
    throw new Error('Task and member are required.');
  }

  const task = rows_('Tasks').find(r =>
    String(r[0]).trim() === String(data.taskId).trim()
  );

  const member = rows_('Members').find(r =>
    String(r[0]).trim() === String(data.memberId).trim() &&
    normalize_(r[10]) === 'active'
  );

  if (!task) throw new Error('Task not found.');
  if (!member) throw new Error('The selected member is missing or inactive.');

  const inactiveTso = rows_('TSOs').some(r =>
    String(r[1]).trim() === String(member[7]).trim() &&
    normalize_(r[6]) !== 'active'
  );

  if (inactiveTso) {
    throw new Error(
      'Members of inactive TSOs cannot receive operational assignments.'
    );
  }

  const existing = rows_('Task Force Members').find(r =>
    String(r[1]).trim() === String(task[0]).trim() &&
    String(r[3]).trim() === String(member[0]).trim() &&
    normalize_(r[7]) === 'active'
  );

  if (existing) {
    throw new Error(member[3] + ' is already assigned to this task.');
  }

  const sh = sheet_('Task Force Members');
  const id = nextId_('Task Force Members', 'TFM', 1, 4);

  sh.appendRow([
    id,
    task[0],
    task[5],
    member[0],
    member[3],
    member[7],
    data.role || 'Assignee',
    data.status || 'Active',
    'Manual entry via App.html',
    '',
    data.notes || ''
  ]);

  return id;
}

function refreshLateJoinerFlags() {
  const sh = sheet_('Members');
  if (sh.getLastRow() < 2) return;

  const dates = sh.getRange(2, 12, sh.getLastRow() - 1, 1).getValues();
  const phase = [];
  const late = [];

  dates.forEach(r => {
    const result = lateJoiner_(r[0]);
    phase.push([result.phase]);
    late.push([result.late]);
  });

  sh.getRange(2, 13, phase.length, 1).setValues(phase);
  sh.getRange(2, 14, late.length, 1).setValues(late);
}
