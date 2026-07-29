/**
 * MED-TSO SMS v6 data-model and UI service layer.
 * Additive only: existing sheets and historical records are preserved.
 */

const SMS_V6_ = {
  version: '7.1.0',
  sheets: {
    teams: 'Teams',
    teamMembers: 'Team Members',
    spocs: 'TSO SPOCs',
    executive: 'Executive Board Members',
    steering: 'Steering Committee Members',
    assembly: 'General Assembly Members'
  }
};

function setupSmsV6() {
  const ss = SpreadsheetApp.getActive();
  const created = [];
  const extended = [];

  const definitions = [
    [SMS_V6_.sheets.teams, ['Team ID','Team Name','Team Type','Lead TC ID','Lead TC','Status','Description','Created Date','Notes']],
    [SMS_V6_.sheets.teamMembers, ['Team Membership ID','Team ID','Team Name','Member ID','Member Name','TSO Code','Team Role','Status','Start Date','End Date','Source','Notes']],
    [SMS_V6_.sheets.spocs, ['SPOC ID','TSO Code','Member ID','Member Name','SPOC Type','Primary','Status','Start Date','End Date','Notes']],
    [SMS_V6_.sheets.executive, ['Governance Membership ID','Member ID','Member Name','TSO Code','Governance Role','Status','Start Date','End Date','Notes']],
    [SMS_V6_.sheets.steering, ['Governance Membership ID','Member ID','Member Name','TSO Code','Governance Role','Status','Start Date','End Date','Notes']],
    [SMS_V6_.sheets.assembly, ['Governance Membership ID','Member ID','Member Name','TSO Code','Governance Role','Status','Start Date','End Date','Notes']]
  ];

  definitions.forEach(def => {
    let sh = ss.getSheetByName(def[0]);
    if (!sh) {
      sh = ss.insertSheet(def[0]);
      sh.getRange(1,1,1,def[1].length).setValues([def[1]]);
      formatHeaderV6_(sh, def[1].length);
      created.push(def[0]);
    }
  });

  const tcAdded = ensureHeadersV6_('TC Members', ['Planning Data SPOC','Start Date','End Date','Notes']);
  if (tcAdded.length) extended.push('TC Members: ' + tcAdded.join(', '));

  const meetingAdded = ensureHeadersV6_('Meetings', ['Invite Group Type','Invite Group ID','Invite Group Name','Start Time','End Time','Location','Online Link']);
  if (meetingAdded.length) extended.push('Meetings: ' + meetingAdded.join(', '));

  const attendanceAdded = ensureHeadersV6_('Attendance', ['Invite Group Type','Invite Group ID','Recorded By','Recorded At']);
  if (attendanceAdded.length) extended.push('Attendance: ' + attendanceAdded.join(', '));

  seedTeamsV6_();

  return {
    ok: true,
    version: SMS_V6_.version,
    created: created,
    extended: extended,
    message: created.length || extended.length
      ? 'SMS v6 data model prepared successfully.'
      : 'SMS v6 data model is already prepared.'
  };
}

function ensureHeadersV6_(sheetName, headersToAdd) {
  const sh = sheet_(sheetName);
  const headers = getHeadersV6_(sh);
  const missing = headersToAdd.filter(h => headers.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    formatHeaderV6_(sh, headers.length + missing.length);
  }
  return missing;
}

function seedTeamsV6_() {
  const sh = sheet_(SMS_V6_.sheets.teams);
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['TEAM-001','Modelling Team','Operational Team','','','Active','Power-system and market modelling support',new Date(),''],
    ['TEAM-002','Network Study Team','Operational Team','','TC Planning','Active','Regional network studies and modelling',new Date(),''],
    ['TEAM-003','Adequacy Study Team','Operational Team','','TC Economic Studies & Scenarios','Active','Adequacy study preparation and review',new Date(),''],
    ['TEAM-004','Market Study Team','Operational Team','','TC Economic Studies & Scenarios','Active','Market study preparation and review',new Date(),''],
    ['TEAM-005','MSMT & Climate Data Team','Operational Team','','TC Economic Studies & Scenarios','Active','MSMT and climate-data coordination',new Date(),'']
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function getSmsBootstrap() {
  const options = getFormOptions();
  return {
    version: SMS_V6_.version,
    options: options,
    teams: safeRecordsV6_(SMS_V6_.sheets.teams).filter(r => normalize_(r.Status) !== 'inactive'),
    groups: getActiveGroupsV61(),
    counts: getSmsCountsV6_(),
    setupRequired: !SpreadsheetApp.getActive().getSheetByName(SMS_V61_.groupsSheet)
  };
}

function getSmsCountsV6_() {
  const count = name => {
    const sh = SpreadsheetApp.getActive().getSheetByName(name);
    return !sh || sh.getLastRow() < 2 ? 0 : sh.getLastRow() - 1;
  };
  return {
    members: count('Members'),
    activeMembers: safeRecordsV6_('Members').filter(r => normalize_(r.Status) === 'active').length,
    tsos: safeRecordsV6_('TSOs').filter(r => normalize_(r.Status) === 'active').length,
    tcs: count('TCs'),
    teams: count(SMS_V6_.sheets.teams),
    taskForces: count('Tasks'),
    meetings: count('Meetings'),
    attendance: count('Attendance')
  };
}

function listTcMembershipsV6(filters) {
  filters = filters || {};
  return safeRecordsV6_('TC Members').filter(r => {
    return (!filters.tcName || r['TC Name'] === filters.tcName) &&
      (!filters.memberId || r['Member ID'] === filters.memberId) &&
      (!filters.status || r.Status === filters.status);
  });
}

function saveTcMembershipV6(data) {
  return withLockV6_(() => {
    requireV6_(data, ['memberId','tcName']);
    const member = activeMemberV6_(data.memberId);
    const tc = safeRecordsV6_('TCs').find(r => r['TC Name'] === data.tcName);
    if (!tc) throw new Error('Technical Committee not found.');

    if (String(data.planningDataSpoc || 'No') === 'Yes' && normalize_(data.tcName).indexOf('planning') < 0) {
      throw new Error('Planning Data SPOC can only be assigned within TC Planning.');
    }

    const existing = safeRecordsV6_('TC Members').find(r =>
      r['Member ID'] === member['Member ID'] && r['TC ID'] === tc['TC ID'] && normalize_(r.Status) === 'active');
    if (existing) throw new Error(member['Display Name'] + ' already has an active membership in this committee.');

    appendByHeadersV6_('TC Members', {
      'TC Membership ID': nextId_('TC Members','TCM',1,4),
      'Member ID': member['Member ID'], 'Member Name': member['Display Name'], 'TSO Code': member['TSO Code'],
      'TC ID': tc['TC ID'], 'TC Name': tc['TC Name'], 'Committee Role': data.role || 'Member',
      'Status': data.status || 'Active', 'Source Sheet': 'App v6', 'Source Row': '',
      'Planning Data SPOC': data.planningDataSpoc || 'No', 'Start Date': data.startDate || '',
      'End Date': data.endDate || '', 'Notes': data.notes || ''
    });
    return {ok:true, message:'Technical Committee membership saved.'};
  });
}

function listTeamsV6() {
  return safeRecordsV6_(SMS_V6_.sheets.teams);
}

function saveTeamMembershipV6(data) {
  return withLockV6_(() => {
    requireV6_(data, ['memberId','teamId']);
    const member = activeMemberV6_(data.memberId);
    const team = safeRecordsV6_(SMS_V6_.sheets.teams).find(r => r['Team ID'] === data.teamId);
    if (!team) throw new Error('Operational team not found.');

    const existing = safeRecordsV6_(SMS_V6_.sheets.teamMembers).find(r =>
      r['Member ID'] === member['Member ID'] && r['Team ID'] === team['Team ID'] && normalize_(r.Status) === 'active');
    if (existing) throw new Error(member['Display Name'] + ' is already an active member of this team.');

    appendByHeadersV6_(SMS_V6_.sheets.teamMembers, {
      'Team Membership ID': nextId_(SMS_V6_.sheets.teamMembers,'TM',1,4),
      'Team ID': team['Team ID'], 'Team Name': team['Team Name'],
      'Member ID': member['Member ID'], 'Member Name': member['Display Name'], 'TSO Code': member['TSO Code'],
      'Team Role': data.role || 'Member', 'Status': data.status || 'Active',
      'Start Date': data.startDate || '', 'End Date': data.endDate || '', 'Source': 'App v6', 'Notes': data.notes || ''
    });
    return {ok:true, message:'Operational team membership saved.'};
  });
}

function saveTsoSpocV6(data) {
  return withLockV6_(() => {
    requireV6_(data, ['memberId','spocType']);
    const member = activeMemberV6_(data.memberId);
    if (!member['TSO Code']) throw new Error('The selected member has no TSO code.');

    if (String(data.primary || 'No') === 'Yes') {
      const duplicatePrimary = safeRecordsV6_(SMS_V6_.sheets.spocs).find(r =>
        r['TSO Code'] === member['TSO Code'] && r['SPOC Type'] === data.spocType &&
        String(r.Primary) === 'Yes' && normalize_(r.Status) === 'active');
      if (duplicatePrimary) throw new Error('This TSO already has an active primary SPOC of this type.');
    }

    appendByHeadersV6_(SMS_V6_.sheets.spocs, {
      'SPOC ID': nextId_(SMS_V6_.sheets.spocs,'SPOC',1,4), 'TSO Code': member['TSO Code'],
      'Member ID': member['Member ID'], 'Member Name': member['Display Name'],
      'SPOC Type': data.spocType, 'Primary': data.primary || 'No', 'Status': data.status || 'Active',
      'Start Date': data.startDate || '', 'End Date': data.endDate || '', 'Notes': data.notes || ''
    });
    return {ok:true, message:'TSO SPOC assignment saved.'};
  });
}

function getMeetingOptionsV6() {
  return safeRecordsV6_('Meetings')
    .filter(r => normalize_(r.Status) !== 'cancelled')
    .map(r => ({id:r['Meeting ID'], name:r['Meeting Name'], date:dateTextV6_(r['Meeting Date']), tc:r['TC Name']}))
    .sort((a,b) => String(b.date).localeCompare(String(a.date)));
}

function getAttendanceRosterV61(meetingId) {
  const meetings = safeRecordsV6_('Meetings');
  const memberships = safeRecordsV6_('Group Members');
  const members = safeRecordsV6_('Members');

  const meeting = meetings.find(
    row => String(row['Meeting ID']).trim() === String(meetingId).trim()
  );

  if (!meeting) {
    throw new Error('Meeting not found: ' + meetingId);
  }

  const groupId = String(
    meeting['Group ID'] ||
    meeting['Invite Group ID'] ||
    meeting['TC ID'] ||
    ''
  ).trim();

  if (!groupId) {
    throw new Error('The selected meeting has no Group ID.');
  }

  const memberIndex = {};
  members.forEach(member => {
    memberIndex[String(member['Member ID']).trim()] = member;
  });

  const roster = memberships
    .filter(row =>
      String(row['Group ID']).trim() === groupId &&
      ['yes', 'active', 'true', '1'].includes(
        String(row['Active']).trim().toLowerCase()
      )
    )
    .map(row => {
      const memberId = String(row['Member ID']).trim();
      const member = memberIndex[memberId] || {};

      return {
        memberId: memberId,
        memberName:
          member['Display Name'] ||
          row['Member Name'] ||
          memberId,
        tsoCode:
          member['TSO Code'] ||
          row['TSO Code'] ||
          '',
        role: row['Role'] || member['Role / Function'] || 'Member',
        status: 'Absent'
      };
    });

  return {
    meeting: meeting,
    groupId: groupId,
    groupName:
      meeting['Group Name'] ||
      meeting['Invite Group Name'] ||
      meeting['TC Name'] ||
      '',
    members: roster
  };
}

function saveAttendanceRosterV6(payload) {
  return withLockV6_(() => {
    requireV6_(payload, ['meetingId']);
    const meeting = safeRecordsV6_('Meetings').find(r => r['Meeting ID'] === payload.meetingId);
    if (!meeting) throw new Error('Meeting not found.');
    const sh = sheet_('Attendance');
    const headers = getHeadersV6_(sh);
    const existingRows = safeRecordsV6_('Attendance').filter(r => r['Meeting ID'] === payload.meetingId);
    const byMember = {};
    existingRows.forEach(r => byMember[r['Member ID']] = r);

    (payload.members || []).forEach(item => {
      const record = {
        'Meeting ID': meeting['Meeting ID'], 'Meeting Name': meeting['Meeting Name'],
        'Meeting Date': meeting['Meeting Date'], 'TC ID': meeting['TC ID'], 'TC Name': meeting['TC Name'],
        'Member ID': item.memberId, 'Member Name': item.memberName, 'TSO Code': item.tsoCode,
        'Attendance Status': item.status || 'Absent', 'Operational KPI Eligible': item.status === 'Present' ? 'Yes' : 'No',
        'Notes': item.notes || '', 'Invite Group Type': meeting['Invite Group Type'] || 'Technical Committee',
        'Invite Group ID': meeting['Invite Group ID'] || meeting['TC ID'],
        'Recorded By': Session.getActiveUser().getEmail() || 'App user', 'Recorded At': new Date()
      };
      const old = byMember[item.memberId];
      if (old) {
        const rowValues = headers.map(h => Object.prototype.hasOwnProperty.call(record,h) ? record[h] : old[h]);
        sh.getRange(old.__row,1,1,headers.length).setValues([rowValues]);
      } else {
        record['Attendance ID'] = nextId_('Attendance','ATT',1,5);
        appendByHeadersV6_('Attendance', record);
      }
    });
    return {ok:true, message:'Attendance saved for ' + (payload.members || []).length + ' members.'};
  });
}

function safeRecordsV6_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues();
  const headers = values.shift().map(String);
  return values.filter(row => row.some(v => v !== '' && v !== null)).map((row,i) => {
    const obj = {__row:i+2};
    headers.forEach((h,j) => obj[h] = row[j]);
    return obj;
  });
}

function getHeadersV6_(sh) {
  return sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
}

function appendByHeadersV6_(sheetName, record) {
  const sh = sheet_(sheetName);
  const headers = getHeadersV6_(sh);
  sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(record,h) ? record[h] : ''));
}

function activeMemberV6_(memberId) {
  const member = safeRecordsV6_('Members').find(r => r['Member ID'] === memberId);
  if (!member) throw new Error('Member not found.');
  if (normalize_(member.Status) !== 'active') throw new Error('The selected member is inactive.');
  return member;
}

function requireV6_(data, fields) {
  fields.forEach(f => { if (!data || !String(data[f] || '').trim()) throw new Error(f + ' is required.'); });
}

function withLockV6_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function formatHeaderV6_(sh, columns) {
  sh.getRange(1,1,1,columns).setFontWeight('bold').setBackground('#123b5d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function dateTextV6_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * MED-TSO SMS v6.1 unified Groups model.
 * Additive migration: existing sheets and records are preserved.
 */
const SMS_V61_ = {
  version: '7.1.0',
  groupsSheet: 'Groups',
  groupMembersSheet: 'Group Members'
};

function setupUnifiedGroupsV61() {
  return withLockV6_(() => {
    const ss = SpreadsheetApp.getActive();
    const created = [];
    const ensure = (name, headers) => {
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        sh.getRange(1,1,1,headers.length).setValues([headers]);
        formatHeaderV6_(sh, headers.length);
        created.push(name);
      } else {
        ensureHeadersV6_(name, headers);
      }
      return sh;
    };

    ensure(SMS_V61_.groupsSheet, ['Group ID','Group Name','Group Type','Parent Group ID','Parent Group Name','Active','Description','Source Sheet','Source ID','Created At']);
    ensure(SMS_V61_.groupMembersSheet, ['Membership ID','Group ID','Group Name','Group Type','Member ID','Member Name','TSO Code','Role','Active','Start Date','End Date','Source Sheet','Source Membership ID','Notes']);
    ensureHeadersV6_('Meetings', ['Group ID','Group Name','Group Type']);
    ensureHeadersV6_('Attendance', ['Group ID','Group Name','Group Type']);

    const migration = migrateUnifiedGroupsV61_();
    const links = backfillUnifiedGroupLinksV61_();
    return {ok:true, version:SMS_V61_.version, created:created, migration:migration, links:links,
      message:'Unified Groups model prepared. ' + migration.groupsAdded + ' groups and ' + migration.membershipsAdded +
        ' memberships added. ' + links.meetingsUpdated + ' meetings and ' + links.attendanceUpdated + ' attendance rows linked.'};
  });
}

function migrateUnifiedGroupsV61_() {
  const groupRows = safeRecordsV6_(SMS_V61_.groupsSheet);
  const groupKeys = {};
  groupRows.forEach(r => groupKeys[String(r['Group ID'])] = true);
  let groupsAdded = 0;

  const addGroup = (id,name,type,parentId,parentName,description,sourceSheet,sourceId) => {
    id = String(id || '').trim(); name = String(name || '').trim();
    if (!id || !name || groupKeys[id]) return;
    appendByHeadersV6_(SMS_V61_.groupsSheet, {
      'Group ID':id,'Group Name':name,'Group Type':type,'Parent Group ID':parentId||'',
      'Parent Group Name':parentName||'','Active':'Yes','Description':description||'',
      'Source Sheet':sourceSheet||'','Source ID':sourceId||id,'Created At':new Date()
    });
    groupKeys[id] = true; groupsAdded++;
  };

  safeRecordsV6_('TCs').forEach(r => addGroup(r['TC ID'],r['TC Name'],'Technical Committee','','',r.Description,'TCs',r['TC ID']));
  safeRecordsV6_('Teams').forEach(r => addGroup(r['Team ID'],r['Team Name'],'Team',r['Lead TC ID'],r['Lead TC'],r.Description,'Teams',r['Team ID']));
  safeRecordsV6_('Tasks').forEach(r => addGroup(r['Task ID'],r['Task Name'],'Task Force',r['TC ID'],r['TC Name'],r.Description,'Tasks',r['Task ID']));

  [
    ['GOV-EXEC','Executive Board','Governance','Executive Board Members'],
    ['GOV-STEER','Steering Committee','Governance','Steering Committee Members'],
    ['GOV-GA','General Assembly','Governance','General Assembly Members'],
    ['SEC-001','Secretariat','Secretariat','Secretariat Members']
  ].forEach(x => { if (SpreadsheetApp.getActive().getSheetByName(x[3])) addGroup(x[0],x[1],x[2],'','',x[1],x[3],x[0]); });

  const existing = {};
  safeRecordsV6_(SMS_V61_.groupMembersSheet).forEach(r => existing[String(r['Group ID'])+'|'+String(r['Member ID'])] = true);
  let membershipsAdded = 0, skipped = 0;
  const addMembership = (groupId,groupName,groupType,memberId,memberName,tso,role,active,start,end,sourceSheet,sourceId,notes) => {
    groupId=String(groupId||'').trim(); memberId=String(memberId||'').trim();
    if (!groupId || !memberId) { skipped++; return; }
    const key=groupId+'|'+memberId;
    if (existing[key]) return;
    appendByHeadersV6_(SMS_V61_.groupMembersSheet, {
      'Membership ID':nextId_(SMS_V61_.groupMembersSheet,'GM',1,5),
      'Group ID':groupId,'Group Name':groupName||'','Group Type':groupType||'',
      'Member ID':memberId,'Member Name':memberName||'','TSO Code':tso||'',
      'Role':role||'Member','Active':normalize_(active)==='inactive'?'No':'Yes',
      'Start Date':start||'','End Date':end||'','Source Sheet':sourceSheet||'',
      'Source Membership ID':sourceId||'','Notes':notes||''
    });
    existing[key]=true; membershipsAdded++;
  };

  safeRecordsV6_('TC Members').forEach(r => addMembership(r['TC ID'],r['TC Name'],'Technical Committee',r['Member ID'],r['Member Name'],r['TSO Code'],r['Committee Role'],r.Status,r['Start Date'],r['End Date'],'TC Members',r['TC Membership ID'],r.Notes));
  safeRecordsV6_('Team Members').forEach(r => addMembership(r['Team ID'],r['Team Name'],'Team',r['Member ID'],r['Member Name'],r['TSO Code'],r['Team Role'],r.Status,r['Start Date'],r['End Date'],'Team Members',r['Team Membership ID'],r.Notes));
  safeRecordsV6_('Task Force Members').forEach(r => addMembership(r['Task ID'],r['Task Name'],'Task Force',r['Member ID'],r['Member Name'],r['TSO Code'],r['Assignment Role']||r.Role,r.Status,'','Task Force Members',r['Task Force Membership ID']||r['Assignment ID'],r.Notes));

  const govSources = [
    ['Executive Board Members','GOV-EXEC','Executive Board'],
    ['Steering Committee Members','GOV-STEER','Steering Committee'],
    ['General Assembly Members','GOV-GA','General Assembly'],
    ['Secretariat Members','SEC-001','Secretariat']
  ];
  govSources.forEach(x => safeRecordsV6_(x[0]).forEach(r => addMembership(x[1],x[2],x[1]==='SEC-001'?'Secretariat':'Governance',r['Member ID'],r['Member Name'],r['TSO Code'],r['Governance Role']||r.Role,r.Status,r['Start Date'],r['End Date'],x[0],r['Governance Membership ID']||r['Membership ID'],r.Notes)));

  return {groupsAdded:groupsAdded,membershipsAdded:membershipsAdded,skipped:skipped,totalGroups:safeRecordsV6_(SMS_V61_.groupsSheet).length,totalMemberships:safeRecordsV6_(SMS_V61_.groupMembersSheet).length};
}


/**
 * Backfills the unified Group fields on historical Meetings and Attendance rows.
 * Existing explicit Group values are preserved. Historical TC ID / TC Name values
 * are used as the primary mapping source.
 */
function backfillUnifiedGroupLinksV61_() {
  const groups = safeRecordsV6_(SMS_V61_.groupsSheet);
  const byId = {};
  const byName = {};
  groups.forEach(g => {
    const id = String(g['Group ID'] || '').trim();
    const name = normalize_(g['Group Name']);
    if (id) byId[id] = g;
    if (name && !byName[name]) byName[name] = g;
  });

  const meetingSheet = sheet_('Meetings');
  const meetingHeaders = getHeadersV6_(meetingSheet);
  const meetings = safeRecordsV6_('Meetings');
  const meetingGroups = {};
  let meetingsUpdated = 0;
  let meetingsUnmatched = 0;

  meetings.forEach(m => {
    let group = null;
    const existingGroupId = String(m['Group ID'] || m['Invite Group ID'] || '').trim();
    const tcId = String(m['TC ID'] || '').trim();
    const existingName = normalize_(m['Group Name'] || m['Invite Group Name'] || m['TC Name']);

    if (existingGroupId && byId[existingGroupId]) group = byId[existingGroupId];
    else if (tcId && byId[tcId]) group = byId[tcId];
    else if (existingName && byName[existingName]) group = byName[existingName];

    if (!group) {
      meetingsUnmatched++;
      return;
    }

    meetingGroups[String(m['Meeting ID'] || '').trim()] = group;
    const changes = {
      'Invite Group Type': group['Group Type'] || '',
      'Invite Group ID': group['Group ID'] || '',
      'Invite Group Name': group['Group Name'] || '',
      'Group ID': group['Group ID'] || '',
      'Group Name': group['Group Name'] || '',
      'Group Type': group['Group Type'] || ''
    };

    const needsUpdate = Object.keys(changes).some(h => !String(m[h] || '').trim() && String(changes[h] || '').trim());
    if (!needsUpdate) return;

    const row = meetingHeaders.map(h => {
      if (Object.prototype.hasOwnProperty.call(changes, h) && !String(m[h] || '').trim()) return changes[h];
      return m[h];
    });
    meetingSheet.getRange(m.__row, 1, 1, meetingHeaders.length).setValues([row]);
    meetingsUpdated++;
  });

  const attendanceSheet = sheet_('Attendance');
  const attendanceHeaders = getHeadersV6_(attendanceSheet);
  const attendance = safeRecordsV6_('Attendance');
  let attendanceUpdated = 0;
  let attendanceUnmatched = 0;

  attendance.forEach(a => {
    const meetingId = String(a['Meeting ID'] || '').trim();
    let group = meetingGroups[meetingId] || null;
    if (!group) {
      const explicitId = String(a['Group ID'] || a['Invite Group ID'] || a['TC ID'] || '').trim();
      const name = normalize_(a['Group Name'] || a['TC Name']);
      group = (explicitId && byId[explicitId]) || (name && byName[name]) || null;
    }
    if (!group) {
      attendanceUnmatched++;
      return;
    }

    const changes = {
      'Invite Group Type': group['Group Type'] || '',
      'Invite Group ID': group['Group ID'] || '',
      'Group ID': group['Group ID'] || '',
      'Group Name': group['Group Name'] || '',
      'Group Type': group['Group Type'] || ''
    };
    const needsUpdate = Object.keys(changes).some(h => !String(a[h] || '').trim() && String(changes[h] || '').trim());
    if (!needsUpdate) return;

    const row = attendanceHeaders.map(h => {
      if (Object.prototype.hasOwnProperty.call(changes, h) && !String(a[h] || '').trim()) return changes[h];
      return a[h];
    });
    attendanceSheet.getRange(a.__row, 1, 1, attendanceHeaders.length).setValues([row]);
    attendanceUpdated++;
  });

  return {
    meetingsUpdated: meetingsUpdated,
    meetingsUnmatched: meetingsUnmatched,
    attendanceUpdated: attendanceUpdated,
    attendanceUnmatched: attendanceUnmatched
  };
}

function repairUnifiedGroupLinks() {
  const result = withLockV6_(() => backfillUnifiedGroupLinksV61_());
  SpreadsheetApp.getUi().alert(
    'Group links repaired',
    'Meetings updated: ' + result.meetingsUpdated + '\n' +
    'Meetings unmatched: ' + result.meetingsUnmatched + '\n' +
    'Attendance rows updated: ' + result.attendanceUpdated + '\n' +
    'Attendance rows unmatched: ' + result.attendanceUnmatched,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function getActiveGroupsV61() {
  return safeRecordsV6_(SMS_V61_.groupsSheet)
    .filter(r => normalize_(r.Active) !== 'no' && normalize_(r.Active) !== 'inactive')
    .map(r => ({id:r['Group ID'],name:r['Group Name'],type:r['Group Type'],parent:r['Parent Group Name']||''}))
    .sort((a,b) => (a.type+' '+a.name).localeCompare(b.type+' '+b.name));
}

function saveMeetingV61(data) {
  return withLockV6_(() => {
    requireV6_(data,['groupId','meetingName','meetingDate']);
    const group = safeRecordsV6_(SMS_V61_.groupsSheet).find(r => String(r['Group ID'] || '').trim() === String(data.groupId || '').trim());
    if (!group) throw new Error('Meeting group not found. Run Unified Groups setup first.');
    const id = nextId_('Meetings','MTG',1,4);
    appendByHeadersV6_('Meetings', {
      'Meeting ID':id,'TC ID':group['Group Type']==='Technical Committee'?group['Group ID']:(group['Parent Group ID']||''),
      'TC Name':group['Group Type']==='Technical Committee'?group['Group Name']:(group['Parent Group Name']||''),
      'Meeting Name':data.meetingName,'Meeting Date':data.meetingDate,'Meeting Type':data.meetingType||'Regular',
      'Description':data.description||'','Status':data.status||'Planned',
      'Source':'Manual entry via App v6.1','Source Sheet':'Manual entry via App v6.1',
      'Operational KPI Eligible':data.status==='Cancelled'?'No':'Yes','KPI Eligible':data.status==='Cancelled'?'No':'Yes','Notes':'',
      'Invite Group Type':group['Group Type'],'Invite Group ID':group['Group ID'],'Invite Group Name':group['Group Name'],
      'Group ID':group['Group ID'],'Group Name':group['Group Name'],'Group Type':group['Group Type'],
      'Start Time':data.startTime||'','End Time':data.endTime||'','Location':data.location||'','Online Link':data.onlineLink||''
    });
    return id;
  });
}

function getMeetingOptionsV61() {
  return safeRecordsV6_('Meetings').filter(r=>normalize_(r.Status)!=='cancelled').map(r=>({
    id:r['Meeting ID'],name:r['Meeting Name'],date:dateTextV6_(r['Meeting Date']),
    group:r['Group Name']||r['Invite Group Name']||r['TC Name']||'',type:r['Group Type']||r['Invite Group Type']||'Technical Committee'
  })).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
}

function getAttendanceRosterV61(meetingId) {
  const clean = v =>
    String(v == null ? '' : v)
      .replace(/\u00A0/g, ' ')
      .trim()
      .toUpperCase();

  const targetMeetingId = clean(meetingId);

  const meetings = safeRecordsV6_('Meetings');
  const meeting = meetings.find(
    r => clean(r['Meeting ID']) === targetMeetingId
  );

  if (!meeting) {
    throw new Error('Meeting not found: ' + meetingId);
  }

  const groupId = String(
    meeting['Group ID'] ||
    meeting['Invite Group ID'] ||
    meeting['TC ID'] ||
    ''
  ).trim();

  if (!groupId) {
    throw new Error('This meeting has no group assigned.');
  }

  const targetGroupId = clean(groupId);

  const groupMemberships = safeRecordsV6_('Group Members');

  const matching = groupMemberships.filter(
    r => clean(r['Group ID']) === targetGroupId
  );

  const active = matching.filter(r => {
    const value = clean(r['Active']);
    return !['NO', 'INACTIVE', 'FALSE', '0'].includes(value);
  });

  const masterMembers = safeRecordsV6_('Members');
  const memberMaster = {};

  masterMembers.forEach(r => {
    memberMaster[clean(r['Member ID'])] = r;
  });

  const attendanceRows = safeRecordsV6_('Attendance');
  const existing = {};

  attendanceRows
    .filter(r => clean(r['Meeting ID']) === targetMeetingId)
    .forEach(r => {
      existing[clean(r['Member ID'])] = r;
    });

  const seen = {};

  const roster = active
    .map(r => {
      const memberId = String(r['Member ID'] || '').trim();
      const key = clean(memberId);

      if (!key || seen[key]) {
        return null;
      }

      seen[key] = true;

      const master = memberMaster[key] || {};
      const oldAttendance = existing[key] || {};

      return {
        memberId: memberId,
        memberName: String(
          master['Display Name'] ||
          r['Member Name'] ||
          [master['First Name'], master['Last Name']]
            .filter(Boolean)
            .join(' ') ||
          memberId
        ),
        tsoCode: String(
          master['TSO Code'] ||
          r['TSO Code'] ||
          ''
        ),
        role: String(
          r['Role'] ||
          master['Role / Function'] ||
          'Member'
        ),
        status: String(
          oldAttendance['Attendance Status'] ||
          'Absent'
        ),
        notes: String(
          oldAttendance['Notes'] ||
          ''
        )
      };
    })
    .filter(Boolean)
    .sort((a, b) =>
      a.memberName.localeCompare(b.memberName)
    );

  return {
    meeting: {
      'Meeting ID': String(meeting['Meeting ID'] || ''),
      'Meeting Name': String(meeting['Meeting Name'] || ''),
      'Meeting Date': dateTextV6_(meeting['Meeting Date'])
    },
    groupId: groupId,
    groupName: String(
      meeting['Group Name'] ||
      meeting['Invite Group Name'] ||
      meeting['TC Name'] ||
      ''
    ),
    members: roster,
    diagnostics: {
      selectedMeetingId: String(meetingId || ''),
      resolvedGroupId: groupId,
      totalGroupMemberRows: groupMemberships.length,
      matchingGroupRows: matching.length,
      activeMatchingRows: active.length,
      returnedMembers: roster.length
    }
  };
}

function saveAttendanceRosterV61(payload) {
  return withLockV6_(() => {
    requireV6_(payload,['meetingId']);
    const meeting=safeRecordsV6_('Meetings').find(r=>String(r['Meeting ID']||'').trim()===String(payload.meetingId||'').trim());
    if(!meeting) throw new Error('Meeting not found.');
    const groupId=meeting['Group ID']||meeting['Invite Group ID']||meeting['TC ID'];
    const groupName=meeting['Group Name']||meeting['Invite Group Name']||meeting['TC Name']||'';
    const groupType=meeting['Group Type']||meeting['Invite Group Type']||'Technical Committee';
    const sh=sheet_('Attendance'), headers=getHeadersV6_(sh), byMember={};
    safeRecordsV6_('Attendance').filter(r=>r['Meeting ID']===payload.meetingId).forEach(r=>byMember[r['Member ID']]=r);
    (payload.members||[]).forEach(item=>{
      const record={'Meeting ID':meeting['Meeting ID'],'Meeting Name':meeting['Meeting Name'],'Meeting Date':meeting['Meeting Date'],
        'TC ID':meeting['TC ID'],'TC Name':meeting['TC Name'],'Member ID':item.memberId,'Member Name':item.memberName,'TSO Code':item.tsoCode,
        'Attendance Status':item.status||'Absent','Operational KPI Eligible':['Present','Online','Remote'].indexOf(item.status)>=0?'Yes':'No','Notes':item.notes||'',
        'Invite Group Type':groupType,'Invite Group ID':groupId,'Group ID':groupId,'Group Name':groupName,'Group Type':groupType,
        'Recorded By':Session.getActiveUser().getEmail()||'App user','Recorded At':new Date()};
      const old=byMember[item.memberId];
      if(old) sh.getRange(old.__row,1,1,headers.length).setValues([headers.map(h=>Object.prototype.hasOwnProperty.call(record,h)?record[h]:old[h])]);
      else {record['Attendance ID']=nextId_('Attendance','ATT',1,5);appendByHeadersV6_('Attendance',record);}
    });
    return {ok:true,message:'Attendance saved for '+(payload.members||[]).length+' members.'};
  });
}
