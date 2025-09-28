/***** 07_auto_sync_reloj.gs *******************************************/
function enableAutoSyncEvery10m(){ ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='autoSyncGuides_').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoSyncGuides_').timeBased().everyMinutes(10).create(); }
function disableAutoSync(){ ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='autoSyncGuides_').forEach(t=>ScriptApp.deleteTrigger(t)); }
function autoSyncGuides_(){ const ss=SpreadsheetApp.getActive(); ss.getSheets().forEach(sh=>{ if(/^\d{2}_\d{4}$/.test(sh.getName())){ ss.setActiveSheet(sh); syncActiveMonthFromGuides(); } }); }
function syncActiveMonthFromGuides(){ const sh=SpreadsheetApp.getActiveSheet(); if(!/^\d{2}_\d{4}$/.test(sh.getName())) return;
  const dates=sh.getRange(3,1,sh.getLastRow()-2,1).getValues().map(r=>r[0]?new Date(r[0]):null).filter(Boolean); const lastCol=sh.getLastColumn();
  for(let c=3;c<=lastCol;c+=2){ const code=String(sh.getRange(1,c).getValue()||'').split('—')[0].trim(); if(!code) continue;
    const guideInfo=JSON.parse(P.getProperty('guide:'+code)||'{}'); if(!guideInfo.id) continue;
    const guide=SpreadsheetApp.openById(guideInfo.id); const gsh=guide.getSheetByName(sh.getName()); if(!gsh) continue;
    const current=sh.getRange(3,c,dates.length,2).getValues(); const out=current.map(r=>[r[0],r[1]]);
    for(let i=0;i<dates.length;i++){
      if(String(current[i][0]||'').indexOf('ASIGNADO')!==0){ const pM=locateGuideCell_(gsh, dates[i], 'M'); const vM=pM?String(gsh.getRange(pM.row,pM.col).getValue()).trim():''; out[i][0]=(vM==='NO DISPONIBLE')?'NO DISPONIBLE':''; }
      if(String(current[i][1]||'').indexOf('ASIGNADO')!==0){ const pT=locateGuideCell_(gsh, dates[i], 'T'); const vT=pT?String(gsh.getRange(pT.row,pT.col).getValue()).trim():''; out[i][1]=(vT==='NO DISPONIBLE')?'NO DISPONIBLE':''; }
    } sh.getRange(3,c,dates.length,2).setValues(out);
  }
}
