/***** 05_sync_master_onEdit.gs ****************************************/
function createOnEditTrigger(){ ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='onEditMaster_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditMaster_').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create(); }
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
  if (!guideInfo.id) { sh.getRange(row,col).setValue('').setNote('Guía no registrado.'); return; }

  const slotMT = isMorning ? 'M' : 'T';                 // para hoja del guía
  let   slotKey = isMorning ? 'M' : null;               // para Bookeo: M/T1/T2/T3
  if (!isMorning && action.indexOf('ASIGNAR') === 0) {  // ej. "ASIGNAR T2"
    slotKey = action.split(' ')[1];                     // T1|T2|T3
  }

  try {
    LOCK.tryLock(3000);

    if (action.indexOf('ASIGNAR')===0) {
      const status = readGuideStatus_(guideInfo.id, sh.getName(), date, slotMT);
      if (status === 'NO DISPONIBLE') { sh.getRange(row,col).setValue('NO DISPONIBLE').setNote('Guía marcó NO DISPONIBLE.'); return; }

      const ev = findBookeoEventBySlot_(date, slotKey); // usa M/T1/T2/T3
      if (!ev) { sh.getRange(row,col).setValue('').setNote('EVENTO NO EXISTE'); SpreadsheetApp.getActive().toast('EVENTO NO EXISTE en Bookeo; no se invita.'); return; }

      const assignLabel = isMorning ? 'ASIGNADO M' : ('ASIGNADO ' + slotKey);
      sh.getRange(row,col).setValue(assignLabel);
      writeGuideAssignment_(guideInfo.id, sh.getName(), date, slotMT, assignLabel, true);

      inviteGuideToEvent_(ev, guideInfo.email); try { ev.addPopupReminder(240); } catch(_) {}
      const link = getEventHtmlLink_(ev);
      sendMailAssignment_(guideInfo.email, guideInfo.name, date, assignLabel, link);

    } else if (action === 'LIBERAR') {
      const st = readGuideStatus_(guideInfo.id, sh.getName(), date, slotMT);
      if (st === 'NO DISPONIBLE') { sh.getRange(row,col).setValue('NO DISPONIBLE').setNote('Solo el guía puede LIBERAR.'); return; }

      // Determinar T1/T2/T3 desde el valor anterior o desde el estado del guía
      if (!isMorning) {
        const prev = String(e.oldValue || '');                          // disponible en onEdit(e) para una sola celda
        const m = /ASIGNADO\s+(T\d)/.exec(prev) || /ASIGNADO\s+(T\d)/.exec(st);
        slotKey = m ? m[1] : null;
      }

      sh.getRange(row,col).setValue('');
      writeGuideAssignment_(guideInfo.id, sh.getName(), date, slotMT, '', false);

      if (slotKey) { const ev2 = findBookeoEventBySlot_(date, slotKey); if (ev2) { try { ev2.removeGuest(guideInfo.email); } catch(_) {} } }

      sendMailLiberation_(guideInfo.email, guideInfo.name, date, slotMT);
    }
  } finally { try { LOCK.releaseLock(); } catch(_){} }
}

function reapplyAssignmentsActiveMonth(){ const sh=SpreadsheetApp.getActiveSheet(); if(!/^\d{2}_\d{4}$/.test(sh.getName())){ SpreadsheetApp.getActive().toast('No es una pestaña mensual'); return; }
  const dates=sh.getRange(3,1,sh.getLastRow()-2,1).getValues().map(r=>r[0]).filter(Boolean); const lastCol=sh.getLastColumn();
  for(let c=3;c<=lastCol;c+=2){ const code=String(sh.getRange(1,c).getValue()||'').split('—')[0].trim(); if(!code) continue;
    const guideInfo=JSON.parse(P.getProperty('guide:'+code)||'{}'); if(!guideInfo.id) continue;
    const colM=c, colT=c+1; for(let i=0;i<dates.length;i++){ const vM=String(sh.getRange(3+i,colM).getValue()||''); if(vM.indexOf('ASIGNADO')===0) writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], 'M', vM, true);
      const vT=String(sh.getRange(3+i,colT).getValue()||''); if(vT.indexOf('ASIGNADO')===0) writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], 'T', vT, true); } }
  SpreadsheetApp.getActive().toast('Asignaciones reaplicadas'); }
