/***** 06_listener_edicion_guia.gs *************************************
 * Trigger instalable en cada GUÍA: revertir ASIGNADO y espejar NO DISP/LIB.
 ***********************************************************************/

function ensureGuideEditTriggersForAllGuides() {
  const reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  const rows = reg.getDataRange().getValues().slice(1).filter(r=>r[4]);
  rows.forEach(r => ensureGuideEditTriggerForGuide_(String(r[4])));
  SpreadsheetApp.getActive().toast('Guardián activo en todas las guías');
}

function ensureGuideEditTriggerForGuide_(guideId) {
  ScriptApp.getProjectTriggers()
    .filter(t => (t.getHandlerFunction()==='guideEditHandler_') && t.getTriggerSourceId && t.getTriggerSourceId()===guideId)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('guideEditHandler_').forSpreadsheet(guideId).onEdit().create();
}

function guideEditHandler_(e) {
  if (!e || !e.range) return;
  const gSh = e.range.getSheet();
  if (!/^\d{2}_\d{4}$/.test(gSh.getName())) return;
  const row = e.range.getRow(), col = e.range.getColumn();
  if ((row - 3) % 3 === 0) return; // fila de números
  const oldV = String(e.oldValue || '');
  const newV = String(e.value || '').trim();

  // 1) Intento de cambiar ASIGNADO => revertir y re-proteger
  if (oldV && oldV.indexOf('ASIGNADO')===0 && newV !== oldV) {
    const r = e.range.setValue(oldV);
    // limpiar protecciones previas del rango y re-proteger
    const protections = gSh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(p=>{ const rn=p.getRange(); if (rn.getRow()===r.getRow() && rn.getColumn()===r.getColumn()) p.remove(); });
    const p = r.protect().setDescription('Asignado por MASTER');
    p.setWarningOnly(false);
    const me = Session.getEffectiveUser(); p.addEditor(me);
    const toRemove = p.getEditors().filter(u=>u.getEmail()!==me.getEmail()); if (toRemove.length) p.removeEditors(toRemove);
    if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
    return;
  }

  // 2) Sanitizar entrada
  const allowed = {'':1, 'NO DISPONIBLE':1, 'LIBERAR':1};
  if (!allowed[newV]) { e.range.setValue(''); return; }

  // 3) Espejo al MASTER si no está ASIGNADO
  const tab = gSh.getName();
  const [year, month] = [Number(tab.split('_')[1]), Number(tab.split('_')[0])];
  const w = Math.floor((row-3)/3), dow = col-1;
  const numberRow = 3 + w*3;
  const day = Number(gSh.getRange(numberRow, col).getValue());
  if (!day) return;
  const date = new Date(year, month-1, day);
  const slot = ((row - 3) % 3 === 1) ? 'M' : 'T';

  const master = SpreadsheetApp.openById(P.getProperty('MASTER_ID') || SpreadsheetApp.getActive().getId());
  const mSh = master.getSheetByName(tab); if (!mSh) return;

  const title = e.source.getName();
  const m = title.match(/GUÍA\s+([^\s—]+)/); const code = m && m[1] ? m[1] : '';
  if (!code) return;

  let startCol = null;
  for (let c=3; c<=mSh.getLastColumn(); c+=2) {
    const head = String(mSh.getRange(1,c).getValue()||'');
    if (head.indexOf(code)===0) { startCol = c; break; }
  }
  if (!startCol) return;
  const targetCol = slot==='M' ? startCol : startCol+1;

  const dates = mSh.getRange(3,1,mSh.getLastRow()-2,1).getValues().map(r=>r[0] && new Date(r[0]).setHours(0,0,0,0));
  const idx = dates.findIndex(x => x === new Date(date).setHours(0,0,0,0));
  if (idx < 0) return;
  const mCell = mSh.getRange(3+idx, targetCol);
  const mVal = String(mCell.getValue()||'');

  if (mVal.indexOf('ASIGNADO')===0) return;
  if (newV === 'NO DISPONIBLE') mCell.setValue('NO DISPONIBLE'); else mCell.setValue('');
}
