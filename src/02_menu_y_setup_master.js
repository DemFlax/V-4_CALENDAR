/***** 02_menu_y_setup_master.gs ***************************************
 * Menú, setup inicial, registro y meses en MASTER.
 ***********************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tours Madrid')
    .addItem('Setup inicial', 'setupMaster')
    .addSeparator()
      .addItem('Añadir guía', 'menuAddGuide')
      .addItem('Eliminar guía...', 'menuRemoveGuide')
      .addItem('Sincronizar meses a guías', 'syncMonthsToGuides')
      .addItem('Reaplicar asignaciones (mes activo)', 'reapplyAssignmentsActiveMonth')
    .addSeparator()
      .addItem('Activar guardián anti-ediciones', 'ensureGuideEditTriggersForAllGuides')
      .addItem('Activar auto-actualización (cada 10 min)', 'enableAutoSyncEvery10m')
      .addItem('Desactivar auto-actualización', 'disableAutoSync')
    .addSeparator()
      .addItem('Crear disparador onEdit', 'createOnEditTrigger')
    .addToUi();
}

function setupMaster() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  if (!P.getProperty('MASTER_ID')) P.setProperty('MASTER_ID', ss.getId());
  ensureRegistry_();
  ensureInitialMonths_();
  refreshGuideIndexFromRegistry_();
  applyAllMasterDV_();
}

function ensureRegistry_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.REGISTRY_SHEET);
  sh.getRange(1,1,1,CFG.REGISTRY_HEADERS.length).setValues([CFG.REGISTRY_HEADERS]);
  sh.setFrozenRows(1);
}

function ensureInitialMonths_() {
  const ss = SpreadsheetApp.getActive();
  CFG.MONTHS_INITIAL.forEach(tag => { if (!ss.getSheetByName(toTabName_(tag))) createMonthSheet_(ss, tag); });
}

function createMonthSheet_(ss, tag) {
  const sh = ss.insertSheet(toTabName_(tag));
  sh.getRange('A1').setValue('FECHA'); sh.getRange('B1').setValue('DÍA');
  sh.setFrozenRows(2);
  const [y, m] = tag.split('-').map(Number);
  const dates = enumerateMonthDates_(y, m);
  sh.getRange(3,1,dates.length,1).setValues(dates.map(d=>[d])).setNumberFormat('dd/MM/yyyy');
  sh.getRange(3,2,dates.length,1).setValues(dates.map(d=>[getWeekdayShort_(d)]));
}

function refreshGuideIndexFromRegistry_() {
  const reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  if (!reg) return;
  const rows = reg.getDataRange().getValues().slice(1).filter(r=>r[1] && r[4]);
  rows.forEach(r=>{
    const code = String(r[1]).trim(), name = String(r[2]).trim(), email = String(r[3]).trim(), id = String(r[4]).trim();
    P.setProperty('guide:'+code, JSON.stringify({id, email, name, code}));
  });
}
