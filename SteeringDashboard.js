/**
 * MED-TSO SMS v7.0 — Steering Committee Dashboard foundation.
 * Read-only service. No spreadsheet structure or records are changed.
 */
function getSteeringCommitteeDashboardV70() {
  const activeTsos = steeringRowsV70_('TSOs').filter(function(r) {
    return steeringNormV70_(steeringPickV70_(r, ['Status'])) !== 'inactive';
  });
  const tsoCodes = activeTsos.map(function(r) {
    return String(steeringPickV70_(r, ['TSO Code','Code','TSO']) || '').trim();
  }).filter(Boolean);
  const denominator = tsoCodes.length;

  const members = steeringRowsV70_('Members');
  const activeMembers = members.filter(function(r) {
    return steeringNormV70_(steeringPickV70_(r, ['Status'])) === 'active';
  });
  const tasks = steeringRowsV70_('Tasks').filter(function(r) {
    const status = steeringNormV70_(steeringPickV70_(r, ['Status','Task Status']));
    return status !== 'completed' && status !== 'cancelled' && status !== 'archived';
  });
  const meetings = steeringRowsV70_('Meetings');
  const tcRows = steeringRowsV70_('TCs').filter(function(r) {
    return steeringNormV70_(steeringPickV70_(r, ['Status'])) !== 'inactive';
  });
  const tcMembers = steeringRowsV70_('TC Members').filter(function(r) {
    return steeringNormV70_(steeringPickV70_(r, ['Status'])) !== 'inactive';
  });

  const committees = tcRows.map(function(tc) {
    const name = String(steeringPickV70_(tc, ['TC Name','Committee Name','Name']) || '').trim();
    const id = String(steeringPickV70_(tc, ['TC ID','Committee ID','ID']) || '').trim();
    const rows = tcMembers.filter(function(m) {
      const memberTcId = String(steeringPickV70_(m, ['TC ID','Committee ID']) || '').trim();
      const memberTcName = String(steeringPickV70_(m, ['TC Name','Committee Name']) || '').trim();
      return (id && memberTcId === id) || (name && memberTcName === name);
    });
    const represented = steeringUniqueV70_(rows.map(function(m) {
      return String(steeringPickV70_(m, ['TSO Code','TSO']) || '').trim();
    }).filter(Boolean));
    const coverage = denominator ? Math.round(represented.length / denominator * 1000) / 10 : 0;
    const committeeTasks = tasks.filter(function(t) {
      return String(steeringPickV70_(t, ['TC Name','Committee Name','Lead TC']) || '').trim() === name;
    });
    return {
      id: id,
      name: name || 'Unnamed committee',
      coverage: coverage,
      participatingTsos: represented.length,
      missingTsos: tsoCodes.filter(function(code) { return represented.indexOf(code) < 0; }),
      activeMembers: rows.length,
      activeTasks: committeeTasks.length,
      risk: coverage < 60 ? 'Critical' : coverage < 80 ? 'Watch' : 'Healthy'
    };
  }).sort(function(a,b) { return b.coverage - a.coverage; });

  const criticalTasks = tasks.filter(function(t) {
    const priority = steeringNormV70_(steeringPickV70_(t, ['Priority','Task Priority']));
    const due = steeringDateV70_(steeringPickV70_(t, ['Due Date','Deadline']));
    const overdue = due && due.getTime() < new Date().setHours(0,0,0,0);
    return priority === 'critical' || overdue;
  }).slice(0, 8).map(function(t) {
    return {
      id: String(steeringPickV70_(t, ['Task ID','ID']) || ''),
      name: String(steeringPickV70_(t, ['Task Name','Name']) || 'Unnamed task'),
      committee: String(steeringPickV70_(t, ['TC Name','Committee Name','Lead TC']) || 'Unassigned committee'),
      priority: String(steeringPickV70_(t, ['Priority','Task Priority']) || 'Normal'),
      dueDate: steeringDateTextV70_(steeringPickV70_(t, ['Due Date','Deadline'])),
      status: String(steeringPickV70_(t, ['Status','Task Status']) || '')
    };
  });

  const upcomingMeetings = meetings.map(function(m) {
    const d = steeringDateV70_(steeringPickV70_(m, ['Meeting Date','Date']));
    return {
      id: String(steeringPickV70_(m, ['Meeting ID','ID']) || ''),
      name: String(steeringPickV70_(m, ['Meeting Name','Name']) || 'Meeting'),
      group: String(steeringPickV70_(m, ['Group Name','TC Name','Invite Group Name']) || ''),
      dateObj: d,
      date: steeringDateTextV70_(d)
    };
  }).filter(function(m) {
    if (!m.dateObj) return false;
    const start = new Date(); start.setHours(0,0,0,0);
    const end = new Date(start); end.setDate(end.getDate() + 60);
    return m.dateObj >= start && m.dateObj <= end;
  }).sort(function(a,b) { return a.dateObj - b.dateObj; }).slice(0,6).map(function(m) {
    delete m.dateObj; return m;
  });

  const avgCoverage = committees.length
    ? Math.round(committees.reduce(function(sum,c){ return sum + c.coverage; },0) / committees.length * 10) / 10
    : 0;

  return {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Cairo', 'dd MMM yyyy HH:mm'),
    summary: {
      overallParticipation: avgCoverage,
      activeTsos: denominator,
      activeMembers: activeMembers.length,
      activeTasks: tasks.length,
      criticalRisks: committees.filter(function(c){return c.risk === 'Critical';}).length + criticalTasks.length,
      upcomingMeetings: upcomingMeetings.length
    },
    committees: committees,
    criticalTasks: criticalTasks,
    upcomingMeetings: upcomingMeetings
  };
}

function steeringRowsV70_(sheetName) {
  try {
    if (typeof safeRecordsV6_ === 'function') return safeRecordsV6_(sheetName) || [];
  } catch (e) {}
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.map(function(row) {
    const out = {};
    headers.forEach(function(h,i){ out[h] = row[i]; });
    return out;
  });
}
function steeringPickV70_(row, keys) {
  for (var i=0;i<keys.length;i++) if (row && row[keys[i]] !== undefined && row[keys[i]] !== '') return row[keys[i]];
  return '';
}
function steeringNormV70_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function steeringUniqueV70_(rows) { return rows.filter(function(v,i,a){ return a.indexOf(v) === i; }); }
function steeringDateV70_(v) {
  if (!v) return null;
  const d = v instanceof Date ? new Date(v) : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function steeringDateTextV70_(v) {
  const d = steeringDateV70_(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone() || 'Africa/Cairo', 'dd MMM yyyy') : '';
}
