/**
 * MED-TSO SMS v6.3.2 — resilient participation gap analysis.
 * Uses the workbook's canonical source sheets directly:
 * TCs, TC Members, Tasks, Task Force Members and TSOs.
 * Unified Groups remains optional and is not required for the executive view.
 */
const GAP_V632_ = {version:'6.3.2', sheet:'Gap Analysis'};

function getGapAnalysisV63() {
  const activeTsos = safeRecordsV6_('TSOs')
    .filter(function(r){ return gapIsActive_(r.Status); })
    .map(function(r){
      return {
        code: gapText_(r['TSO Code'] || r['Organisation Code']),
        name: gapText_(r['Official Name'] || r['TSO Name'] || r['Organisation Name'])
      };
    })
    .filter(function(x){ return x.code; })
    .sort(function(a,b){ return a.code.localeCompare(b.code); });

  const tsoCodes = activeTsos.map(function(x){ return x.code; });
  const tsoLookup = {};
  tsoCodes.forEach(function(code){ tsoLookup[gapNorm_(code)] = code; });

  const tcRows = safeRecordsV6_('TCs')
    .filter(function(r){ return gapIsActive_(r.Status); })
    .map(function(r){
      return {
        id: gapText_(r['TC ID']),
        name: gapText_(r['TC Name']),
        shortName: gapText_(r['Short Name'])
      };
    })
    .filter(function(tc){ return tc.id || tc.name; });

  const tcById = {};
  const tcByName = {};
  tcRows.forEach(function(tc){
    if (tc.id) tcById[gapNorm_(tc.id)] = tc;
    if (tc.name) tcByName[gapNorm_(tc.name)] = tc;
    if (tc.shortName) tcByName[gapNorm_(tc.shortName)] = tc;
  });

  const committeeRepresentation = {};
  safeRecordsV6_('TC Members').forEach(function(r){
    if (!gapCountsAsParticipation_(r.Status)) return;
    const tc = gapResolveTc_(r['TC ID'], r['TC Name'], tcById, tcByName);
    const tso = gapCanonicalTso_(r['TSO Code'], tsoLookup);
    if (tc && tso) committeeRepresentation[tc.id + '|' + tso] = true;
  });

  const taskRepresentation = {};
  safeRecordsV6_('Task Force Members').forEach(function(r){
    if (!gapCountsAsParticipation_(r.Status)) return;
    const taskId = gapText_(r['Task ID']);
    const tso = gapCanonicalTso_(r['TSO Code'], tsoLookup);
    if (taskId && tso) taskRepresentation[taskId + '|' + tso] = true;
  });

  const tasks = safeRecordsV6_('Tasks')
    .filter(function(r){ return !gapIsClosedTask_(r.Status); })
    .map(function(r){
      const tc = gapResolveTc_(r['TC ID'], r['TC Name'], tcById, tcByName);
      const taskId = gapText_(r['Task ID']);
      const participating = tsoCodes.filter(function(code){ return !!taskRepresentation[taskId + '|' + code]; });
      const missing = tsoCodes.filter(function(code){ return !taskRepresentation[taskId + '|' + code]; });
      return gapRowV632_(
        'Task',
        taskId,
        gapText_(r['Task Name']) || taskId,
        participating,
        missing,
        tsoCodes.length,
        tc ? tc.name : gapText_(r['TC Name']),
        tc ? tc.id : gapText_(r['TC ID'])
      );
    })
    .filter(function(t){ return t.id; });

  const committees = tcRows.map(function(tc){
    const participating = tsoCodes.filter(function(code){ return !!committeeRepresentation[tc.id + '|' + code]; });
    const missing = tsoCodes.filter(function(code){ return !committeeRepresentation[tc.id + '|' + code]; });
    const committeeCoverage = gapCoverage_(participating.length, tsoCodes.length);
    const relatedTasks = tasks
      .filter(function(t){
        return (t.parentId && gapNorm_(t.parentId) === gapNorm_(tc.id)) ||
          (!t.parentId && gapNorm_(t.parent) === gapNorm_(tc.name));
      })
      .sort(function(a,b){ return b.missingCount-a.missingCount || a.name.localeCompare(b.name); });
    const taskCoverage = relatedTasks.length
      ? gapRound1_(relatedTasks.reduce(function(sum,t){ return sum+t.coverage; },0) / relatedTasks.length)
      : 0;
    const taskMissingSet = {};
    relatedTasks.forEach(function(t){ t.missing.forEach(function(code){ taskMissingSet[code]=true; }); });
    return {
      id: tc.id,
      name: tc.name,
      shortName: tc.shortName,
      committeeCoverage: committeeCoverage,
      committeeParticipating: participating,
      committeeMissing: missing,
      taskCount: relatedTasks.length,
      taskCoverage: taskCoverage,
      taskMissingTsos: Object.keys(taskMissingSet).sort(),
      criticalTasks: relatedTasks.filter(function(t){ return t.coverage < 50; }).length,
      watchTasks: relatedTasks.filter(function(t){ return t.coverage >= 50 && t.coverage < 75; }).length,
      healthyTasks: relatedTasks.filter(function(t){ return t.coverage >= 75; }).length,
      tasks: relatedTasks
    };
  }).sort(function(a,b){
    const aNoTasks = a.taskCount ? 0 : 1;
    const bNoTasks = b.taskCount ? 0 : 1;
    return aNoTasks-bNoTasks || a.taskCoverage-b.taskCoverage || b.criticalTasks-a.criticalTasks || a.name.localeCompare(b.name);
  });

  const committeeGapRows = committees.map(function(c){
    return gapRowV632_('Technical Committee', c.id, c.name, c.committeeParticipating, c.committeeMissing, tsoCodes.length, '', '');
  });
  const rows = committeeGapRows.concat(tasks).sort(function(a,b){
    return gapTypeOrder_(a.type)-gapTypeOrder_(b.type) || b.missingCount-a.missingCount || a.name.localeCompare(b.name);
  });

  const avgCommitteeCoverage = committees.length
    ? gapRound1_(committees.reduce(function(sum,c){ return sum+c.committeeCoverage; },0) / committees.length)
    : 0;
  const avgTaskCoverage = tasks.length
    ? gapRound1_(tasks.reduce(function(sum,t){ return sum+t.coverage; },0) / tasks.length)
    : 0;

  return {
    ok: true,
    version: GAP_V632_.version,
    generatedAt: new Date(),
    activeTsos: tsoCodes.length,
    tsoCodes: tsoCodes,
    rows: rows,
    committees: committees,
    summary: {
      committeeCount: committees.length,
      taskCount: tasks.length,
      avgCommitteeCoverage: avgCommitteeCoverage,
      avgTaskCoverage: avgTaskCoverage,
      criticalTasks: tasks.filter(function(t){ return t.coverage < 50; }).length,
      watchTasks: tasks.filter(function(t){ return t.coverage >= 50 && t.coverage < 75; }).length,
      healthyTasks: tasks.filter(function(t){ return t.coverage >= 75; }).length,
      completeTasks: tasks.filter(function(t){ return t.missingCount === 0; }).length,
      committeesWithoutTasks: committees.filter(function(c){ return c.taskCount === 0; }).length
    },
    diagnostics: {
      source: 'TCs + TC Members + Tasks + Task Force Members',
      activeTsoCount: tsoCodes.length,
      committeeCount: committees.length,
      taskCount: tasks.length,
      assignmentRows: safeRecordsV6_('Task Force Members').length,
      tcMembershipRows: safeRecordsV6_('TC Members').length
    }
  };
}

function gapResolveTc_(id, name, byId, byName) {
  return byId[gapNorm_(id)] || byName[gapNorm_(name)] || null;
}
function gapCanonicalTso_(value, lookup) {
  return lookup[gapNorm_(value)] || '';
}
function gapText_(value) {
  return String(value == null ? '' : value).trim();
}
function gapNorm_(value) {
  return gapText_(value).toLowerCase().replace(/\s+/g,' ');
}
function gapIsActive_(status) {
  const s = gapNorm_(status);
  return !s || s === 'active' || s === 'yes' || s === 'true';
}
function gapCountsAsParticipation_(status) {
  const s = gapNorm_(status);
  return ['inactive','cancelled','canceled','archived','removed'].indexOf(s) < 0;
}
function gapIsClosedTask_(status) {
  const s = gapNorm_(status);
  return ['completed','cancelled','canceled','archived'].indexOf(s) >= 0;
}
function gapCoverage_(participating, total) {
  return total ? gapRound1_(participating * 100 / total) : 0;
}
function gapRound1_(value) {
  return Math.round(Number(value || 0) * 10) / 10;
}
function gapRowV632_(type,id,name,present,missing,total,parent,parentId) {
  return {
    type:type,
    id:id,
    name:name,
    parent:parent || '',
    parentId:parentId || '',
    participating:present,
    missing:missing,
    participatingCount:present.length,
    missingCount:missing.length,
    coverage:gapCoverage_(present.length,total)
  };
}
function gapTypeOrder_(type) {
  return {'Technical Committee':1,'Team':2,'Task Force':3,'Task':4}[type] || 99;
}

function refreshGapAnalysisV63() {
  const data = getGapAnalysisV63();
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(GAP_V632_.sheet);
  if (!sh) sh = ss.insertSheet(GAP_V632_.sheet);
  sh.clear();
  const headers = ['Type','Name','Parent / TC','Coverage %','Participating TSO Count','Missing TSO Count','Participating TSOs','Missing TSOs'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  formatHeaderV6_(sh,headers.length);
  const values = data.rows.map(function(r){
    return [r.type,r.name,r.parent,r.coverage,r.participatingCount,r.missingCount,r.participating.join(', '),r.missing.join(', ')];
  });
  if (values.length) sh.getRange(2,1,values.length,headers.length).setValues(values);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,headers.length);
  sh.setColumnWidth(2,260);
  sh.setColumnWidth(7,320);
  sh.setColumnWidth(8,420);
  if (sh.getFilter()) sh.getFilter().remove();
  if (sh.getLastRow()>1) sh.getDataRange().createFilter();
  sh.activate();
  return {ok:true,message:'Gap Analysis refreshed: '+data.summary.committeeCount+' committees and '+data.summary.taskCount+' active tasks.'};
}
