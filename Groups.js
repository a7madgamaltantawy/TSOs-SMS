/**
 * MED-TSO SMS — Unified Groups administration.
 *
 * The canonical Groups and Group Members implementation lives in Model.gs.
 * This file exposes only safe menu actions and validation, avoiding duplicate
 * helper names that can override the core application functions.
 */

function initializeUnifiedGroups() {
  const result = setupUnifiedGroupsV61();
  SpreadsheetApp.getUi().alert(
    'Unified Groups initialized',
    [
      result.message || 'Initialization completed.',
      '',
      'Groups added: ' + ((result.migration && result.migration.groupsAdded) || 0),
      'Memberships added: ' + ((result.migration && result.migration.membershipsAdded) || 0),
      'Skipped rows: ' + ((result.migration && result.migration.skipped) || 0),
      'Total groups: ' + ((result.migration && result.migration.totalGroups) || 0),
      'Total memberships: ' + ((result.migration && result.migration.totalMemberships) || 0),
      'Meetings linked: ' + ((result.links && result.links.meetingsUpdated) || 0),
      'Attendance rows linked: ' + ((result.links && result.links.attendanceUpdated) || 0)
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function validateUnifiedGroups() {
  const ss = SpreadsheetApp.getActive();
  const issues = [];
  const groups = safeRecordsV6_(SMS_V61_.groupsSheet);
  const memberships = safeRecordsV6_(SMS_V61_.groupMembersSheet);
  const members = safeRecordsV6_('Members');

  if (!ss.getSheetByName(SMS_V61_.groupsSheet)) issues.push('Missing sheet: Groups');
  if (!ss.getSheetByName(SMS_V61_.groupMembersSheet)) issues.push('Missing sheet: Group Members');

  const groupIds = new Set(groups.map(r => String(r['Group ID'] || '').trim()).filter(Boolean));
  const memberIds = new Set(members.map(r => String(r['Member ID'] || '').trim()).filter(Boolean));
  const seen = new Set();

  groups.forEach(r => {
    if (!String(r['Group ID'] || '').trim()) issues.push('A Groups row has no Group ID (row ' + r.__row + ').');
    if (!String(r['Group Name'] || '').trim()) issues.push('A Groups row has no Group Name (row ' + r.__row + ').');
  });

  memberships.forEach(r => {
    const groupId = String(r['Group ID'] || '').trim();
    const memberId = String(r['Member ID'] || '').trim();
    const key = groupId + '|' + memberId;
    if (!groupId || !groupIds.has(groupId)) issues.push('Unknown Group ID at Group Members row ' + r.__row + ': ' + groupId);
    if (!memberId || !memberIds.has(memberId)) issues.push('Unknown Member ID at Group Members row ' + r.__row + ': ' + memberId);
    if (groupId && memberId && seen.has(key)) issues.push('Duplicate membership at row ' + r.__row + ': ' + key);
    seen.add(key);
  });

  let log = ss.getSheetByName('Unified Groups Validation');
  if (!log) log = ss.insertSheet('Unified Groups Validation');
  log.clearContents();
  log.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Validation result']]);
  const rows = (issues.length ? issues : ['Validation passed.']).map(x => [new Date(), x]);
  log.getRange(2, 1, rows.length, 2).setValues(rows);
  log.setFrozenRows(1);
  log.autoResizeColumns(1, 2);

  const message = [
    'Groups: ' + groups.length,
    'Group memberships: ' + memberships.length,
    'Validation issues: ' + issues.length,
    '',
    issues.length ? 'Open “Unified Groups Validation” for details.' : 'Validation passed.'
  ].join('\n');

  SpreadsheetApp.getUi().alert('Unified Groups validation', message, SpreadsheetApp.getUi().ButtonSet.OK);
  return {ok: issues.length === 0, groups: groups.length, memberships: memberships.length, issues: issues};
}
