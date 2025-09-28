/***** 05_sync_master_onEdit.gs ****************************************
 * Asignar/liberar desde MASTER y re-aplicar asignaciones.
 ***********************************************************************/

function createOnEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t=>{ if (t.getHandlerFunction()==='onEditMaster_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditMaster_').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

function onEditMaster_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
  const row = e.range.getRow(), col = e.range.getColumn();
  if (row < 3 || col < 3) return;

  const isMorning = (col % 2 === 1);
  const guideColStart = isMorning ? col : col-1;
  const code = String(sh.getRange(1, guideColStart).getValue()).split('—')[0].trim();
  const date = sh.getRange(row, 1).getValue();
  const action = String(sh.getRange(row, col).getValue()||'').trim();
  if (!code || !date || !action) return;

  const guideInfo = JSON.parse(P.getProperty('guide:'+code) || '{}');
  if (!guideInfo.id) { sh.getRange(row, col).setValue('').setNote('Guía no registrado.'); return; }

  const slot = isMorning ? 'M' : 'T';
  try {
    LOCK.tryLock(3000);
    if (action.indexOf('ASIGNAR')===0) {
      const status = readGuideStatus_(guideInfo.id, sh.getName(), date, slot);
      if (status === 'NO DISPONIBLE') { sh.getRange(row, col).setValue('NO DISPONIBLE').setNote('Guía marcó NO DISPONIBLE.'); return; }
      const assignLabel = slot==='M' ? 'ASIGNADO M' : ('ASIGNADO ' + action.split(' ')[1]);
      sh.getRange(row, col).setValue(assignLabel);
      writeGuideAssignment_(guideInfo.id, sh.getName(), date, slot, assignLabel, true);
      sendMailAssignment_(guideInfo.email, guideInfo.name, date, assignLabel);
    } else if (action === 'LIBERAR') {
      const st = readGuideStatus_(guideInfo.id, sh.getName(), date, slot);
      if (st === 'NO DISPONIBLE') { sh.getRange(row, col).setValue('NO DISPONIBLE').setNote('Solo el guía puede LIBERAR.'); return; }
      sh.getRange(row, col).setValue('');
      writeGuideAssignment_(guideInfo.id, sh.getName(), date, slot, '', false);
      sendMailLiberation_(guideInfo.email, guideInfo.name, date, slot);
    }
  } finally { try { LOCK.releaseLock(); } catch(err){} }
}

// Idempotente: reescribe en GUÍAS lo ASIGNADO del MASTER y re-protege
function reapplyAssignmentsActiveMonth() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) { SpreadsheetApp.getActive().toast('No es una pestaña mensual'); return; }
  const dates = sh.getRange(3,1,sh.getLastRow()-2,1).getValues().map(r=>r[0]).filter(Boolean);
  const lastCol = sh.getLastColumn();

  for (let c=3; c<=lastCol; c+=2) {
    const code = String(sh.getRange(1,c).getValue()||'').split('—')[0].trim();
    if (!code) continue;
    const guideInfo = JSON.parse(P.getProperty('guide:'+code)||'{}');
    if (!guideInfo.id) continue;

    const colM = c, colT = c+1;
    for (let i=0;i<dates.length;i++){
      const vM = String(sh.getRange(3+i,colM).getValue()||'');
      if (vM.indexOf('ASIGNADO')===0) writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], 'M', vM, true);
      const vT = String(sh.getRange(3+i,colT).getValue()||'');
      if (vT.indexOf('ASIGNADO')===0) writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], 'T', vT, true);
    }
  }
  SpreadsheetApp.getActive().toast('Asignaciones reaplicadas');
}
