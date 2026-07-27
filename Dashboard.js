/**
 * Dashboard.gs
 */

function buildDashboard_(db, attendance, coverage, tasks, ranking, gaps) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(KPI_SHEETS_.DASHBOARD);
  if (!sh) sh = ss.insertSheet(KPI_SHEETS_.DASHBOARD);

  sh.clear({contentsOnly: false});
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.setHiddenGridlines(true);

  const completedMeetings = countByStatus_(db['Meetings'], 'Completed');
  const openTasks = countOpenTasks_(db['Tasks']);
  const completedTasks = countByStatus_(db['Tasks'], 'Completed');
  const avgAttendance = attendance.length
    ? attendance.reduce((s, r) => s + r.rate, 0) / attendance.length
    : 0;

  sh.getRange('A1:H1').merge()
    .setValue('MED-TSO Secretariat Dashboard')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground('#0f3d56')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const cards = [
    ['Active TSOs', db.activeTSOs.length],
    ['Active Members', db.activeMembers.length],
    ['Completed Meetings', completedMeetings],
    ['Average Attendance', avgAttendance],
    ['Open Tasks', openTasks],
    ['Completed Tasks', completedTasks],
    ['Tasks with Gaps', gaps.length],
    ['Last Refreshed', new Date()]
  ];

  cards.forEach((card, i) => {
    const row = 3 + Math.floor(i / 4) * 3;
    const col = 1 + (i % 4) * 2;
    sh.getRange(row, col, 1, 2).merge()
      .setValue(card[0])
      .setFontWeight('bold')
      .setBackground('#d9eaf7')
      .setHorizontalAlignment('center');
    sh.getRange(row + 1, col, 1, 2).merge()
      .setValue(card[1])
      .setFontSize(16)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false);
  });

  sh.getRange('G4:H4').setNumberFormat('0.0%');
  sh.getRange('G7:H7').setNumberFormat('yyyy-mm-dd hh:mm');

  const topN = settingNumber_(db, 'Dashboard Top N', 10);
  const rankingStart = 10;

  sh.getRange(rankingStart, 1, 1, 4)
    .setValues([['Rank', 'TSO', 'Final Score', 'Traffic Light']])
    .setFontWeight('bold')
    .setBackground('#0f3d56')
    .setFontColor('#ffffff');

  const rankingRows = ranking.slice(0, topN).map(r => [
    r.rank, r.tso, r.score / 100, r.trafficLight
  ]);
  if (rankingRows.length) {
    sh.getRange(rankingStart + 1, 1, rankingRows.length, 4).setValues(rankingRows);
    sh.getRange(rankingStart + 1, 3, rankingRows.length, 1).setNumberFormat('0.0%');
    applyTrafficLights_(sh, 4, rankingStart + 1, rankingRows.length);
  }

  const gapStart = rankingStart;
  sh.getRange(gapStart, 6, 1, 3)
    .setValues([['Task ID', 'Task', 'Issue']])
    .setFontWeight('bold')
    .setBackground('#0f3d56')
    .setFontColor('#ffffff');

  const gapRows = gaps.slice(0, 12).map(r => [r.taskId, r.taskName, r.issue]);
  if (gapRows.length) {
    sh.getRange(gapStart + 1, 6, gapRows.length, 3).setValues(gapRows);
    sh.getRange(gapStart + 1, 8, gapRows.length, 1).setWrap(true);
  }

  sh.setColumnWidths(1, 8, 115);
  sh.setColumnWidth(7, 180);
  sh.setColumnWidth(8, 220);

  createDashboardCharts_(sh, ranking, attendance, coverage);
  updateLastRefreshed_();
}

function countByStatus_(table, wantedStatus) {
  const status = col_(table, ['Status', table.name === 'Meetings' ? 'Meeting Status' : 'Task Status'], false);
  if (status < 0) return 0;
  return table.rows.filter(r => normalizeKPI_(r[status]) === normalizeKPI_(wantedStatus)).length;
}

function countOpenTasks_(tasks) {
  const status = col_(tasks, ['Status', 'Task Status'], false);
  if (status < 0) return tasks.rows.length;
  return tasks.rows.filter(r => {
    const s = normalizeKPI_(r[status]);
    return s !== 'completed' && s !== 'cancelled' && s !== 'archived';
  }).length;
}

function createDashboardCharts_(dashboard, ranking, attendance, coverage) {
  if (ranking.length) {
    const rankSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KPI_SHEETS_.RANKING);
    const n = Math.min(10, ranking.length);
    const chart = dashboard.newChart()
      .asBarChart()
      .addRange(rankSheet.getRange(1, 2, n + 1, 2))
      .setPosition(27, 1, 0, 0)
      .setOption('title', 'Top TSO Ranking')
      .setOption('legend', {position: 'none'})
      .setOption('hAxis', {format: '0%'})
      .build();
    dashboard.insertChart(chart);
  }

  if (attendance.length) {
    const aSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KPI_SHEETS_.ATTENDANCE);
    const n = Math.min(10, attendance.length);
    const chart = dashboard.newChart()
      .asColumnChart()
      .addRange(aSheet.getRange(1, 1, n + 1, 1))
      .addRange(aSheet.getRange(1, 4, n + 1, 1))
      .setPosition(27, 5, 0, 0)
      .setOption('title', 'Attendance by TSO')
      .setOption('legend', {position: 'none'})
      .setOption('vAxis', {format: '0%'})
      .build();
    dashboard.insertChart(chart);
  }

  if (coverage.length) {
    const cSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KPI_SHEETS_.COVERAGE);
    const n = Math.min(10, coverage.length);
    const chart = dashboard.newChart()
      .asColumnChart()
      .addRange(cSheet.getRange(1, 1, n + 1, 1))
      .addRange(cSheet.getRange(1, 4, n + 1, 1))
      .setPosition(45, 1, 0, 0)
      .setOption('title', 'Technical Committee Coverage')
      .setOption('legend', {position: 'none'})
      .setOption('vAxis', {format: '0%'})
      .build();
    dashboard.insertChart(chart);
  }
}

function updateLastRefreshed_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KPI_SHEETS_.SETTINGS);
  if (!sh) return;
  const values = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues().flat();
  const row = values.findIndex(v => normalizeKPI_(v) === 'last refreshed') + 1;
  if (row > 0) sh.getRange(row, 2).setValue(new Date());
}
