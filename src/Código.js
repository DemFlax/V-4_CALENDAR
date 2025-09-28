/*******************************************************************************
 * Tours Madrid — MASTER v1.0.7
 * - Meses MM_YYYY, registro, DV/colores.
 * - Pull: “Actualizar disponibilidad (mes activo)” trae NO DISPONIBLE desde GUÍAS.
 * - Push: “Empujar asignaciones (mes activo)” MASTER→GUÍA.
 * - Refuerzo: “Reforzar asignaciones (mes activo)” reimpone y blinda ASIGNADO.
 * - Guardián: onEdit en hojas de GUÍA que revierte cambios sobre ASIGNADO.
 * - Auto-sync: trigger cada 10 min que ejecuta el pull en todas las pestañas.
 * - onEdit local: ASIGNAR/LIBERAR en tiempo real respetando NO DISPONIBLE.
 *******************************************************************************/

const CFG = {
  TZ: 'Europe/Madrid',
  REGISTRY_SHEET: 'REGISTRO',
  REGISTRY_HEADERS: ['TIMESTAMP','CODIGO','NOMBRE','EMAIL','FILE_ID','URL'],
  MASTER_M_LIST: ['', 'LIBERAR', 'ASIGNAR M'],
  MASTER_T_LIST: ['', 'LIBERAR', 'ASIGNAR T1', 'ASIGNAR T2', 'ASIGNAR T3'],
  GUIDE_DV_LIST: ['', 'NO DISPONIBLE', 'LIBERAR'],
  MONTHS_INITIAL: ['2025-10','2025-11','2025-12'],
  COLORS: { ASSIGNED: '#A5D6A7', NODISP: '#EF9A9A', BLANK: '#FFFFFF' },
  GUIDES_FOLDER_ID: '1ibz8PUeaFbUraTgRS9VgfjZ_hqs80J-p'
};

const P = PropertiesService.getScriptProperties();
const LOCK = LockService.getScriptLock();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tours Madrid')
    .addItem('Setup inicial', 'setupMaster')
    .addSeparator()
    .addItem('Añadir guía', 'menuAddGuide')
    .addItem('Eliminar guía…', 'removeGuideCompletely')
    .addItem('Sincronizar meses a guías', 'syncMonthsToGuides')
    .addItem('Reconstruir DV/colores (mes activo)', 'rebuildMonthFormatting')
    .addItem('Actualizar disponibilidad (mes activo)', 'syncActiveMonthFromGuides')
    .addItem('Empujar asignaciones (mes activo)', 'pushAssignmentsActiveMonth')
    .addItem('Reforzar asignaciones (mes activo)', 'enforceAssignmentsActiveMonth')
    .addItem('Activar guardián anti-ediciones', 'ensureGuardTriggersForAllGuides')
    .addItem('Activar auto-actualización (cada 10 min)', 'enableAutoSyncEvery10m')
    .addItem('Desactivar auto-actualización', 'disableAutoSync')
    .addSeparator()
    .addItem('Crear disparador onEdit', 'createOnEditTrigger')
    .addToUi();
}
// === Flujo de borrado completo de un guía ===
function removeGuideCompletely() {
  const ui = SpreadsheetApp.getUi();
  const ask = ui.prompt('Eliminar guía', 'Escribe el CÓDIGO exacto (ej.: G01)', ui.ButtonSet.OK_CANCEL);
  if (ask.getSelectedButton() !== ui.Button.OK) return;
  const code = String(ask.getResponseText() || '').trim();
  if (!code) return;

  const reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  const data = reg.getDataRange().getValues();
  const h = data[0];
  const idxCode = h.indexOf('CODIGO');
  const idxId   = h.indexOf('FILE_ID');
  const idxName = h.indexOf('NOMBRE');
  const idxMail = h.indexOf('EMAIL');

  let row = -1, fileId = '', name = '', email = '';
  for (let i=1;i<data.length;i++){
    if (String(data[i][idxCode]).trim() === code) {
      row   = i+1;
      fileId= String(data[i][idxId]||'').trim();
      name  = String(data[i][idxName]||'').trim();
      email = String(data[i][idxMail]||'').trim();
      break;
    }
  }
  if (row < 0) { ui.alert('Código no encontrado'); return; }

  const conf = ui.prompt('Confirmación', `Escribe BORRAR para eliminar a ${code} (${name})`, ui.ButtonSet.OK_CANCEL);
  if (conf.getSelectedButton() !== ui.Button.OK || String(conf.getResponseText()).trim() !== 'BORRAR') return;

  // 1) Quitar triggers "guardián" de esa hoja de guía
  deleteGuardTriggersForGuide_(fileId); // ScriptApp.deleteTrigger :contentReference[oaicite:1]{index=1}

  // 2) Borrar columnas del guía en todas las pestañas MM_YYYY
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
    const lastCol = sh.getLastColumn();
    for (let c=3; c<=lastCol; c+=2) {
      const header = String(sh.getRange(1, c).getValue() || '');
      if (header.indexOf(code) === 0) {
        // Deshacer merge y borrar las dos columnas (MAÑANA/TARDE)
        sh.getRange(1, c, 1, 2).breakApart();        // Range.breakApart :contentReference[oaicite:2]{index=2}
        sh.deleteColumns(c, 2);                       // Sheet.deleteColumns :contentReference[oaicite:3]{index=3}
        break;
      }
    }
    applyMasterDataValidations_(sh); // reconstruye DV/colores para el resto
  });

  // 3) Eliminar fila del REGISTRO
  reg.deleteRow(row);

  // 4) Borrar propiedad cacheada
  PropertiesService.getScriptProperties().deleteProperty('guide:'+code);

  // 5) Borrar el archivo del guía (Drive avanzado si está habilitado; si no, papelera)
  if (fileId) {
    try {
      if (typeof Drive !== 'undefined' && Drive.Files) {
        Drive.Files.remove(fileId);                   // Advanced Drive Service :contentReference[oaicite:4]{index=4}
      } else {
        DriveApp.getFileById(fileId).setTrashed(true); // File.setTrashed(true) :contentReference[oaicite:5]{index=5}
      }
    } catch (err) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e){}
    }
  }

  SpreadsheetApp.getActive().toast(`Guía ${code} eliminado`);
}

// Borra triggers "guardián" asociados a ese spreadsheet de guía
function deleteGuardTriggersForGuide_(guideId){
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'guardAssignments_' &&
        t.getTriggerSourceId && t.getTriggerSourceId() === guideId) {
      ScriptApp.deleteTrigger(t);                     // ScriptApp.deleteTrigger :contentReference[oaicite:6]{index=6}
    }
  });
}

function setupMaster() {
  SpreadsheetApp.getActive().setSpreadsheetTimeZone(CFG.TZ);
  ensureRegistry_();
  ensureInitialMonths_();
  rebuildMonthFormatting();
  createOnEditTrigger();
}

/* ===== Registro y meses ===== */

function ensureRegistry_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.REGISTRY_SHEET);
  sh.getRange(1,1,1,CFG.REGISTRY_HEADERS.length).setValues([CFG.REGISTRY_HEADERS]);
  sh.setFrozenRows(1);
}

function ensureInitialMonths_() {
  const ss = SpreadsheetApp.getActive();
  CFG.MONTHS_INITIAL.forEach(tag => {
    if (!ss.getSheetByName(toTabName_(tag))) createMonthSheet_(ss, tag);
  });
}

function createMonthSheet_(ss, tag) {
  const sh = ss.insertSheet(toTabName_(tag));
  sh.getRange('A1').setValue('FECHA');
  sh.getRange('B1').setValue('DÍA');
  sh.setFrozenRows(2);
  const parts = tag.split('-');
  const y = Number(parts[0]), m = Number(parts[1]);
  const dates = enumerateMonthDates_(y, m);
  sh.getRange(3,1,dates.length,1).setValues(dates.map(d=>[d])).setNumberFormat('dd/mm/yyyy');
  sh.getRange(3,2,dates.length,1).setValues(dates.map(d=>[getWeekdayShort_(d)]));
}

function menuAddGuide() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Nuevo guía','"Nombre; CODIGO" o "Nombre,CODIGO"', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const parts = String(r1.getResponseText()||'').split(/[;,]/).map(function(s){return s.trim();}).filter(function(s){return s;});
  const name = parts[0]||''; const code = parts[1]||'';
  if (!name || !code) throw new Error('Faltan datos del guía: usa "Nombre; CODIGO".');
  const r2 = ui.prompt('Email del guía','Email de Google del guía', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const email = String(r2.getResponseText()||'').trim();
  if (!email) throw new Error('Falta email del guía.');
  addGuide_({code, name, email});
}

function addGuide_({code, name, email}) {
  if (!CFG.GUIDES_FOLDER_ID) throw new Error('Configura CFG.GUIDES_FOLDER_ID');
  const ssMaster = SpreadsheetApp.getActive();
  const folder = DriveApp.getFolderById(CFG.GUIDES_FOLDER_ID);
  const guideFile = SpreadsheetApp.create('GUÍA ' + code + ' — ' + name);
  DriveApp.getFileById(guideFile.getId()).moveTo(folder);

  const guide = SpreadsheetApp.openById(guideFile.getId());
  buildGuideScaffold_(guide, name, code);

  ssMaster.getSheetByName(CFG.REGISTRY_SHEET)
    .appendRow([new Date(), code, name, email, guideFile.getId(), guideFile.getUrl()]);

  ssMaster.getSheets().forEach(function(sh){
    if (/^\d{2}_\d{4}$/.test(sh.getName())) {
      addGuideColumnsToMonth_(sh,{code,name});
      applyMasterDataValidations_(sh);
    }
  });

  syncMonthsToOneGuide_(guide);
  P.setProperty('guide:'+code, JSON.stringify({code, name, email, id: guideFile.getId(), url: guideFile.getUrl()}));
}

function buildGuideScaffold_(guideSS, name, code) {
  guideSS.setSpreadsheetTimeZone(CFG.TZ);
  const sh0 = guideSS.getActiveSheet();
  sh0.setName('PORTADA');
  sh0.getRange('A1').setValue('Guía: ' + name + ' (' + code + ')');
  CFG.MONTHS_INITIAL.forEach(function(tag){
    if (!guideSS.getSheetByName(toTabName_(tag))) createGuideMonthSheet_(guideSS, tag);
  });
}

function createGuideMonthSheet_(ss, tag) {
  const sh = ss.insertSheet(toTabName_(tag));
  const parts = tag.split('-');
  const y = Number(parts[0]), m = Number(parts[1]);
  const obj = buildMonthlyGrid_(y, m);
  const grid = obj.grid, labels = obj.labels;
  sh.getRange(1,1,1,7).setValues([labels]);
  sh.setFrozenRows(2);
  sh.getRange(3,1,grid.length,7).setValues(grid);
  applyGuideDataValidationsForMaster_(sh);
}

function addGuideColumnsToMonth_(sh, info) {
  const code = info.code, name = info.name;
  const lastCol = sh.getLastColumn();
  sh.insertColumnsAfter(lastCol || 2, 2);
  const startCol = (lastCol || 2) + 1;
  sh.getRange(1, startCol, 1, 2).merge().setValue(code + ' — ' + name);
  sh.getRange(2, startCol).setValue('MAÑANA');
  sh.getRange(2, startCol+1).setValue('TARDE');
}

/* ===== DV/colores ===== */

function applyGuideDataValidationsForMaster_(sh) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.GUIDE_DV_LIST, true).build();
  for (var w=0; w<6; w++) {
    var rowM = 3 + w*3 + 1, rowT = 3 + w*3 + 2;
    sh.getRange(rowM,1,1,7).setDataValidation(rule);
    sh.getRange(rowT,1,1,7).setDataValidation(rule);
  }
  var rules = sh.getConditionalFormatRules();
  var body = sh.getRange(3,1,(3+6*3-1)-2,7);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('ASIGNADO').setBackground(CFG.COLORS.ASSIGNED).setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('NO DISPONIBLE').setBackground(CFG.COLORS.NODISP).setRanges([body]).build()
  );
  sh.setConditionalFormatRules(rules);
}

function applyMasterDataValidations_(sh) {
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastCol <= 2) return;
  var mRule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.MASTER_M_LIST, true).build();
  var tRule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.MASTER_T_LIST, true).build();
  for (var c=3; c<=lastCol; c+=2) {
    sh.getRange(3,c,lastRow-2,1).setDataValidation(mRule);
    sh.getRange(3,c+1,lastRow-2,1).setDataValidation(tRule);
  }
  var rules = sh.getConditionalFormatRules();
  var body = sh.getRange(3,3,lastRow-2,lastCol-2);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('ASIGNADO').setBackground(CFG.COLORS.ASSIGNED).setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('NO DISPONIBLE').setBackground(CFG.COLORS.NODISP).setRanges([body]).build()
  );
  sh.setConditionalFormatRules(rules);
}

/* ===== Triggers locales ===== */

function createOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0;i<triggers.length;i++){
    if (triggers[i].getHandlerFunction()==='onEditMaster_') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('onEditMaster_').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

/* ===== onEdit MASTER: tiempo real ===== */

function onEditMaster_(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
  var row = e.range.getRow(), col = e.range.getColumn();
  if (row < 3 || col < 3) return;

  var isMorning = (col % 2 === 1);
  var guideColStart = isMorning ? col : col-1;
  var code = String(sh.getRange(1, guideColStart).getValue()).split('—')[0].trim();
  var date = sh.getRange(row, 1).getValue();
  var action = String(sh.getRange(row, col).getValue()||'').trim();
  if (!code || !date || !action) return;

  var guideInfo = JSON.parse(P.getProperty('guide:'+code) || '{}');
  if (!guideInfo.id) { sh.getRange(row, col).setValue(''); sh.getRange(row, col).setNote('Guía '+code+' no registrado.'); return; }

  var slot = isMorning ? 'M' : 'T';
  try {
    LOCK.tryLock(3000);
    if (action.indexOf('ASIGNAR')===0) {
      var status = readGuideStatus_(guideInfo.id, sh.getName(), date, slot);
      if (status === 'NO DISPONIBLE') { sh.getRange(row, col).setValue(''); sh.getRange(row, col).setNote('No se puede asignar. Guía marcó NO DISPONIBLE.'); return; }
      var assignLabel = slot === 'M' ? 'ASIGNADO M' : ('ASIGNADO ' + action.split(' ')[1]);
      sh.getRange(row, col).setValue(assignLabel);
      writeGuideAssignment_(guideInfo.id, sh.getName(), date, slot, assignLabel, true);
      sendMailAssignment_(guideInfo.email, guideInfo.name, date, assignLabel);
    } else if (action === 'LIBERAR') {
      // NUEVO: si el guía tiene NO DISPONIBLE, el manager no puede liberar
      var st = readGuideStatus_(guideInfo.id, sh.getName(), date, slot);
      if (st === 'NO DISPONIBLE') {
        sh.getRange(row, col).setValue('NO DISPONIBLE')
          .setNote('Turno bloqueado por el guía. Solo el guía puede LIBERAR.');
        return;
      }
      sh.getRange(row, col).setValue('');
      writeGuideAssignment_(guideInfo.id, sh.getName(), date, slot, '', false);
      sendMailLiberation_(guideInfo.email, guideInfo.name, date, slot);
    }
  } finally {
    try { LOCK.releaseLock(); } catch(err){}
  }
}

/* ===== Botón: PUSH MASTER→GUÍA ===== */

function pushAssignmentsActiveMonth() {
  var sh = SpreadsheetApp.getActiveSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) { SpreadsheetApp.getActive().toast('Abre una pestaña MM_YYYY'); return; }
  var dates = sh.getRange(3,1,sh.getLastRow()-2,1).getValues().map(function(r){return r[0] ? new Date(r[0]) : null;}).filter(function(d){return d;});
  var lastCol = sh.getLastColumn();

  for (var c=3; c<=lastCol; c+=2) {
    var header = String(sh.getRange(1,c).getValue()||'');
    var code = header.split('—')[0].trim();
    if (!code) continue;
    var guideInfo = JSON.parse(P.getProperty('guide:'+code)||'{}');
    if (!guideInfo.id) continue;

    var rng = sh.getRange(3,c,dates.length,2);
    var vals = rng.getValues();

    for (var i=0; i<dates.length; i++) {
      vals[i][0] = processPushCell_(sh, dates[i], 'M', vals[i][0], guideInfo);
      vals[i][1] = processPushCell_(sh, dates[i], 'T', vals[i][1], guideInfo);
    }
    rng.setValues(vals);
  }
}

function processPushCell_(masterSheet, dateObj, slot, cellValue, guideInfo) {
  var text = String(cellValue||'').trim();
  if (!text) return '';
  if (text.indexOf('ASIGNADO')===0) return text;

  if (text === 'LIBERAR') {
    // NUEVO: respeta NO DISPONIBLE del guía
    var status = readGuideStatus_(guideInfo.id, masterSheet.getName(), dateObj, slot);
    if (status === 'NO DISPONIBLE') return 'NO DISPONIBLE';
    writeGuideAssignment_(guideInfo.id, masterSheet.getName(), dateObj, slot, '', false);
    sendMailLiberation_(guideInfo.email, guideInfo.name, dateObj, slot);
    return '';
  }
  if (text.indexOf('ASIGNAR')===0) {
    var st = readGuideStatus_(guideInfo.id, masterSheet.getName(), dateObj, slot);
    if (st === 'NO DISPONIBLE') return '';
    var assignLabel = slot==='M' ? 'ASIGNADO M' : ('ASIGNADO ' + text.split(' ')[1]);
    writeGuideAssignment_(guideInfo.id, masterSheet.getName(), dateObj, slot, assignLabel, true);
    sendMailAssignment_(guideInfo.email, guideInfo.name, dateObj, assignLabel);
    return assignLabel;
  }
  return text;
}

/* ===== Botón: PULL GUÍA→MASTER ===== */

function syncActiveMonthFromGuides() {
  var sh = SpreadsheetApp.getActiveSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) { SpreadsheetApp.getActive().toast('Abre una pestaña MM_YYYY'); return; }

  var dates = sh.getRange(3,1,sh.getLastRow()-2,1).getValues()
    .map(function(r){return r[0] ? new Date(r[0]) : null;}).filter(function(d){return d;});

  var lastCol = sh.getLastColumn();
  for (var c=3; c<=lastCol; c+=2) {
    var head = String(sh.getRange(1,c).getValue()||'');
    var code = head.split('—')[0].trim();
    if (!code) continue;
    var guideInfo = JSON.parse(P.getProperty('guide:'+code)||'{}');
    if (!guideInfo.id) continue;

    var guide = SpreadsheetApp.openById(guideInfo.id);
    var gsh = guide.getSheetByName(sh.getName());
    if (!gsh) continue;

    var current = sh.getRange(3,c,dates.length,2).getValues();
    var out = current.map(function(r){ return [r[0], r[1]]; });
    for (var i=0; i<dates.length; i++) {
      if (String(current[i][0]||'').indexOf('ASIGNADO')!==0) {
        var pM = locateGuideCell_(gsh, dates[i], 'M');
        var vM = pM ? String(gsh.getRange(pM.row, pM.col).getValue()).trim() : '';
        out[i][0] = (vM === 'NO DISPONIBLE') ? 'NO DISPONIBLE' : '';
      }
      if (String(current[i][1]||'').indexOf('ASIGNADO')!==0) {
        var pT = locateGuideCell_(gsh, dates[i], 'T');
        var vT = pT ? String(gsh.getRange(pT.row, pT.col).getValue()).trim() : '';
        out[i][1] = (vT === 'NO DISPONIBLE') ? 'NO DISPONIBLE' : '';
      }
    }
    sh.getRange(3,c,dates.length,2).setValues(out);
  }
}

/* ===== Botón: refuerzo protecciones ===== */

function enforceAssignmentsActiveMonth() {
  var sh = SpreadsheetApp.getActiveSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) { SpreadsheetApp.getActive().toast('Abre una pestaña MM_YYYY'); return; }

  var dates = sh.getRange(3,1,sh.getLastRow()-2,1).getValues()
    .map(function(r){return r[0] ? new Date(r[0]) : null;}).filter(function(d){return d;});
  var lastCol = sh.getLastColumn();

  for (var c=3; c<=lastCol; c+=2) {
    var header = String(sh.getRange(1,c).getValue()||'');
    var code = header.split('—')[0].trim();
    var guideInfo = JSON.parse(P.getProperty('guide:'+code)||'{}');
    if (!guideInfo.id) continue;

    for (var i=0; i<dates.length; i++) {
      for (var off=0; off<=1; off++) {
        var cell = sh.getRange(3+i, c+off);
        var v = String(cell.getValue()||'');
        if (v.indexOf('ASIGNADO')!==0) continue;
        var slot = off===0 ? 'M' : 'T';
        writeGuideAssignment_(guideInfo.id, sh.getName(), dates[i], slot, v, true);
      }
    }
  }
}

/* ===== Guardián anti-ediciones en GUÍAS ===== */

function ensureGuardTriggersForAllGuides() {
  var reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  var rows = reg.getDataRange().getValues().slice(1).filter(function(r){return r[4];});
  for (var i=0;i<rows.length;i++) ensureGuardTriggerForGuide_(String(rows[i][4]));
  SpreadsheetApp.getActive().toast('Guardián activado en todas las guías');
}

function ensureGuardTriggerForGuide_(guideId) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0;i<triggers.length;i++) {
    if (triggers[i].getHandlerFunction()==='guardAssignments_' &&
        triggers[i].getTriggerSourceId && triggers[i].getTriggerSourceId()===guideId) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('guardAssignments_').forSpreadsheet(guideId).onEdit().create();
}

function guardAssignments_(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
  if ((e.range.getRow() - 3) % 3 === 0) return;
  var oldV = String(e.oldValue || '');
  var newV = String(e.value || '');
  if (!oldV || oldV.indexOf('ASIGNADO')!==0) return;
  if (newV === oldV) return;

  var r = e.range.setValue(oldV);
  var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i=0;i<protections.length;i++){
    var rn = protections[i].getRange();
    if (rn.getSheet().getName()===sh.getName() && rn.getRow()===r.getRow() && rn.getColumn()===r.getColumn()) {
      protections[i].remove();
    }
  }
  var p = r.protect().setDescription('Asignado por MASTER');
  p.setWarningOnly(false);
  var me = Session.getEffectiveUser();
  p.addEditor(me);
  var editors = p.getEditors();
  var toRemove = [];
  for (var j=0;j<editors.length;j++){ if (editors[j].getEmail() !== me.getEmail()) toRemove.push(editors[j]); }
  if (toRemove.length) p.removeEditors(toRemove);
  if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
}

/* ===== Auto-sync (reloj) ===== */

function enableAutoSyncEvery10m() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0;i<triggers.length;i++){
    if (triggers[i].getHandlerFunction()==='autoSyncGuides_') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('autoSyncGuides_').timeBased().everyMinutes(10).create();
}

function disableAutoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0;i<triggers.length;i++){
    if (triggers[i].getHandlerFunction()==='autoSyncGuides_') ScriptApp.deleteTrigger(triggers[i]);
  }
}

function autoSyncGuides_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (var i=0;i<sheets.length;i++){
    var sh = sheets[i];
    if (/^\d{2}_\d{4}$/.test(sh.getName())) {
      ss.setActiveSheet(sh);
      syncActiveMonthFromGuides();
    }
  }
}

/* ===== IO con hojas de GUÍA ===== */

function readGuideStatus_(guideId, tag, dateObj, slot) {
  var ss = SpreadsheetApp.openById(guideId);
  var sh = ss.getSheetByName(tag);
  if (!sh) return '';
  var pos = locateGuideCell_(sh, dateObj, slot);
  return pos ? String(sh.getRange(pos.row, pos.col).getValue()).trim() : '';
}

function writeGuideAssignment_(guideId, tag, dateObj, slot, value, lockCell) {
  var ss = SpreadsheetApp.openById(guideId);
  var sh = ss.getSheetByName(tag) || createGuideMonthSheet_(ss, fromTabTag_(tag));
  var pos = locateGuideCell_(sh, dateObj, slot);
  if (!pos) return;

  var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i=0;i<protections.length;i++){
    var rn = protections[i].getRange();
    if (rn.getSheet().getName()===sh.getName() && rn.getRow()===pos.row && rn.getColumn()===pos.col) protections[i].remove();
  }

  var r = sh.getRange(pos.row, pos.col).setValue(value);

  if (lockCell && value) {
    var p = r.protect().setDescription('Asignado por MASTER');
    p.setWarningOnly(false);
    var me = Session.getEffectiveUser();
    p.addEditor(me);
    var editors = p.getEditors();
    var toRemove = [];
    for (var j=0;j<editors.length;j++){ if (editors[j].getEmail() !== me.getEmail()) toRemove.push(editors[j]); }
    if (toRemove.length) p.removeEditors(toRemove);
    if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
  }
}

/* ===== Localizar celda en cuadrícula de GUÍA ===== */

function locateGuideCell_(sh, dateObj, slot) {
  var year = Number(sh.getName().split('_')[1]);
  var month = Number(sh.getName().split('_')[0]);
  var first = new Date(year, month-1, 1);
  var startDow = (first.getDay()+6)%7; // 0=Lun
  var d = new Date(dateObj);
  var day = d.getDate();
  var idx = startDow + (day-1);
  var week = Math.floor(idx/7), dow = idx%7;
  var baseRow = 3 + week*3;
  var row = baseRow + (slot==='M' ? 1 : 2);
  var col = 1 + dow;
  var numberCell = sh.getRange(baseRow, col).getValue();
  if (String(numberCell).trim() !== String(day)) return null;
  return {row: row, col: col};
}

/* ===== Utilidades ===== */

function rebuildMonthFormatting() {
  var sh = SpreadsheetApp.getActiveSheet();
  if (/^\d{2}_\d{4}$/.test(sh.getName())) applyMasterDataValidations_(sh);
}
function syncMonthsToGuides() {
  var reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  var rows = reg.getDataRange().getValues().slice(1).filter(function(r){return r[4];});
  for (var i=0;i<rows.length;i++){
    var guide = SpreadsheetApp.openById(rows[i][4]);
    syncMonthsToOneGuide_(guide);
  }
}
function syncMonthsToOneGuide_(guide) {
  var masterMonths = SpreadsheetApp.getActive().getSheets().map(function(s){return s.getName();}).filter(function(n){return /^\d{2}_\d{4}$/.test(n);});
  var existing = guide.getSheets().map(function(s){return s.getName();});
  for (var i=0;i<masterMonths.length;i++){
    var tag = masterMonths[i];
    if (existing.indexOf(tag)===-1) createGuideMonthSheet_(guide, fromTabTag_(tag));
  }
  for (var j=0;j<existing.length;j++){
    var name = existing[j];
    if (/^\d{2}_\d{4}$/.test(name) && masterMonths.indexOf(name)===-1) guide.deleteSheet(guide.getSheetByName(name));
  }
}
function toTabName_(tag){ var parts = tag.split('-'); var y = Number(parts[0]), m = Number(parts[1]); return (String(m).length===1?'0'+m:m) + '_' + y; }
function fromTabTag_(tab){ var a = tab.split('_'); return a[1] + '-' + a[0]; }
function enumerateMonthDates_(year, month){ var a=[]; var last = new Date(year, month, 0).getDate(); for(var d=1; d<=last; d++) a.push(new Date(year, month-1, d)); return a; }
function getWeekdayShort_(d){ var arr = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb']; return arr[new Date(d).getDay()]; }

/* ===== Emails ===== */

function sendMailAssignment_(email, name, dateObj, label) {
  if (!email) return;
  var fecha = Utilities.formatDate(new Date(dateObj), CFG.TZ, 'dd/MM/yyyy');
  MailApp.sendEmail({ to: email, subject: 'Asignado: ' + label + ' — ' + fecha, htmlBody: '<p>Hola '+name+',</p><p>Se te ha <b>asignado</b>: <b>'+label+'</b> el <b>'+fecha+'</b>.</p>' });
}
function sendMailLiberation_(email, name, dateObj, slot) {
  if (!email) return;
  var fecha = Utilities.formatDate(new Date(dateObj), CFG.TZ, 'dd/MM/yyyy');
  var s = slot==='M' ? 'MAÑANA' : 'TARDE';
  MailApp.sendEmail({ to: email, subject: 'Liberación de turno — ' + fecha, htmlBody: '<p>Hola '+name+',</p><p>Se ha <b>liberado</b> tu turno de <b>'+s+'</b> el <b>'+fecha+'</b>.</p>' });
}

/* ===== Grid mensual para GUÍA ===== */

function buildMonthlyGrid_(year, month) {
  var labels = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
  var first = new Date(year, month-1, 1);
  var startDow = (first.getDay()+6)%7;
  var last = new Date(year, month, 0).getDate();
  var rows = 6*3;
  var grid = [];
  for (var i=0;i<rows;i++){ grid.push(['','','','','','','']); }
  for (var d=1; d<=last; d++) {
    var idx = startDow + (d-1);
    var w = Math.floor(idx/7);
    var dow = idx%7;
    grid[w*3][dow] = d;
  }
  return {grid: grid, labels: labels};
}
