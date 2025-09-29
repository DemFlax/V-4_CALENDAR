/***** 05_sync_master_onEdit.gs ****************************************/
/** Crea el trigger instalable de onEdit para el MASTER */
function createOnEditTrigger(){
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEditMaster_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onEditMaster_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

function onEditMaster_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) return; // solo pestañas mensuales
  const row = e.range.getRow(), col = e.range.getColumn();
  if (row < 3 || col < 3) return;

  const action = String((e.value!=null?e.value:sh.getRange(row,col).getValue())||'').trim();
  if (!action) return;
  const isAssign = action.indexOf('ASIGNAR') === 0;
  const isRelease = action === 'LIBERAR';
  if (!isAssign && !isRelease) return;

  const isMorning = (col % 2 === 1);
  const guideColStart = isMorning ? col : col-1;
  const header = String(sh.getRange(1, guideColStart).getValue()||'');
  const code = header.split('—')[0].trim(); // "G01 — DAN"
  const date = sh.getRange(row, 1).getValue();
  if (!code || !date) return;

  const guideInfo = JSON.parse(P.getProperty('guide:'+code) || '{}');
  if (!guideInfo.id) { e.range.setValue('').setNote('Guía no registrado.'); return; }

  const uiToast = msg => SpreadsheetApp.getActive().toast(msg);
  const slotMT = isMorning ? 'M' : 'T';
  const slotKey = isMorning ? 'M' : String(action).replace('ASIGNAR ','').trim(); // T1/T2/T3

  const LOCK = LockService.getDocumentLock();
  LOCK.waitLock(5000);
  try {
    if (isAssign) {
      // 0) Respeto estricto al NO DISPONIBLE del guía
      const status = readGuideStatus_(guideInfo.id, sh.getName(), date, slotMT);
      if (status === 'NO DISPONIBLE') {
        e.range.setValue('NO DISPONIBLE').setNote('Guía marcó NO DISPONIBLE. Solo el guía puede liberar.');
        uiToast('Bloqueado: guía NO DISPONIBLE.');
        return;
      }

      // 1) verificar evento Bookeo
      const ev = findBookeoEventBySlot_(date, slotKey);
      if (!ev) {
        e.range.setValue('').setNote('EVENTO NO EXISTE');
        uiToast('EVENTO NO EXISTE en Bookeo; no se invita.');
        return;
      }

      // 2) escribir en GUÍA
      const assignLabel = isMorning ? 'ASIGNADO M' : ('ASIGNADO ' + slotKey);
      const ok = writeGuideAssignment_(guideInfo.id, sh.getName(), date, slotMT, assignLabel, true);
      if (ok !== true) {
        e.range.setValue('').setNote('Turno inexistente en calendario del guía.');
        uiToast('No existe la celda del turno en la guía (mes/día).');
        return;
      }

      // 3) actualizar MASTER, invitar y email
      e.range.setValue(assignLabel).setNote('');
      inviteGuideToEvent_(ev, guideInfo.email); try { ev.addPopupReminder(240); } catch(_) {}
      const link = getEventHtmlLink_(ev);
      sendMailAssignment_(guideInfo.email, guideInfo.name, date, assignLabel, link);
    }

    if (isRelease) {
      const st = readGuideStatus_(guideInfo.id, sh.getName(), date, slotMT);

      // 0) Si el guía fijó NO DISPONIBLE, el MASTER no puede liberar
      if (st === 'NO DISPONIBLE') {
        e.range.setValue('NO DISPONIBLE').setNote('Solo el guía puede liberar este NO DISPONIBLE.');
        uiToast('Bloqueado: NO DISPONIBLE fijado por el guía.');
        return;
      }

      const wasAssigned = (st && st.indexOf('ASIGNADO') === 0);

      // 1) quitar invitación si hay evento
      const key = (slotMT==='M')?'M':(wasAssigned ? st.replace('ASIGNADO ','').trim() : '');
      if (key){
        const ev2 = findBookeoEventBySlot_(date, key);
        if (ev2 && guideInfo.email) { try { ev2.removeGuest(guideInfo.email); } catch(_) {} }
      }

      // 2) escribir en GUÍA y MASTER
      const ok2 = writeGuideAssignment_(guideInfo.id, sh.getName(), date, slotMT, '', true); // '' => vuelve a MAÑANA/TARDE
      if (ok2 !== true) {
        e.range.setValue('').setNote('Turno inexistente en guía.');
        uiToast('No existe la celda del turno en la guía (mes/día).');
        return;
      }
      e.range.setValue('').setNote('');

      // 3) email
      if (wasAssigned) sendMailLiberation_(guideInfo.email, guideInfo.name, date, slotMT);
    }
  } finally {
    try { LOCK.releaseLock(); } catch(_){}
  }
}

/** Reaplica en GUÍA lo que en MASTER figure como ASIGNADO en el mes activo */
function reapplyAssignmentsActiveMonth(){
  const sh = SpreadsheetApp.getActiveSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) { SpreadsheetApp.getActive().toast('No es una pestaña mensual'); return; }
  const dates = sh.getRange(3,1,Math.max(0,sh.getLastRow()-2),1).getValues().map(r=>r[0]).filter(Boolean);
  const lastCol = sh.getLastColumn();
  for (let c=3; c<=lastCol; c+=2){
    const code = String(sh.getRange(1,c).getValue()||'').split('—')[0].trim();
    if (!code) continue;
    const guideInfo = JSON.parse(P.getProperty('guide:'+code)||'{}');
    if (!guideInfo.id) continue;
    const colM=c, colT=c+1;
    for (let i=0;i<dates.length;i++){
      const vM = String(sh.getRange(3+i,colM).getValue()||'');
      if (vM.indexOf('ASIGNADO')===0) writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], 'M', vM, true);
      const vT = String(sh.getRange(3+i,colT).getValue()||'');
      if (vT.indexOf('ASIGNADO')===0) writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], 'T', vT, true);
    }
  }
  SpreadsheetApp.getActive().toast('Asignaciones reaplicadas');
}
