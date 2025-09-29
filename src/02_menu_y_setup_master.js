/***** 02_menu_y_setup_master.gs ***************************************
 * Menú y setup del MASTER. Sin funciones duplicadas de CRUD ni triggers.
 ***********************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tours Madrid')
    .addItem('Setup inicial', 'setupMaster')
    .addSeparator()
      .addItem('Añadir guía', 'menuAddGuide')                     // en 03
      .addItem('Eliminar guía...', 'menuRemoveGuide')             // en 03
      .addItem('Sincronizar meses a guías', 'syncMonthsToGuides') // en 03
      .addItem('Reaplicar asignaciones (mes activo)', 'reapplyAssignmentsActiveMonth') // en 05
    .addSeparator()
      .addItem('Activar guardián anti-ediciones', 'ensureGuideEditTriggersForAllGuides') // en 06
      .addItem('Activar auto-actualización (cada 10 min)', 'enableAutoSyncEvery10m')     // en 07
      .addItem('Desactivar auto-actualización', 'disableAutoSync')                       // en 07
    .addSeparator()
      .addItem('Crear disparador onEdit', 'createOnEditTrigger') // en 05
      .addItem('Reaplicar formato (MASTER + guías)', 'applyAllFormatting_')
    .addToUi();
}

function setupMaster(){
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  if (!P.getProperty('MASTER_ID')) P.setProperty('MASTER_ID', ss.getId());
  ensureRegistry_();
  ensureInitialMonths_();
  refreshGuideIndexFromRegistry_();
  applyAllMasterDV_(); // en 04
  SpreadsheetApp.getActive().toast('Setup completado');
}

/** Hoja REGISTRO con cabecera fija */
function ensureRegistry_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.REGISTRY_SHEET);
  sh.getRange(1,1,1,CFG.REGISTRY_HEADERS.length).setValues([CFG.REGISTRY_HEADERS]);
  sh.setFrozenRows(1);
}

/** Crear pestañas MM_YYYY iniciales en MASTER */
function ensureInitialMonths_(){
  const ss = SpreadsheetApp.getActive();
  CFG.MONTHS_INITIAL.forEach(tag => {
    const name = toTabName_(tag); // helper en 01
    if (!ss.getSheetByName(name)) createMonthSheet_(ss, tag);
  });
}

/** Construye una pestaña mensual vertical para MASTER */
function createMonthSheet_(ss, tag){
  const sh = ss.insertSheet(toTabName_(tag));
  sh.getRange('A1').setValue('FECHA');
  sh.getRange('B1').setValue('DÍA');
  sh.setFrozenRows(2);
  const [y,m] = tag.split('-').map(Number);
  const dates = enumerateMonthDates_(y,m); // helper en 01
  sh.getRange(3,1,dates.length,1).setValues(dates.map(d=>[d])).setNumberFormat('dd/MM/yyyy');
  sh.getRange(3,2,dates.length,1).setValues(dates.map(d=>[getWeekdayShort_(d)])); // helper en 01
}

/** Indexa guías del REGISTRO en ScriptProperties */
function refreshGuideIndexFromRegistry_(){
  const ss = SpreadsheetApp.getActive();
  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!reg) return;
  const rows = reg.getDataRange().getValues().slice(1);
  // Limpia índice previo
  Object.keys(P.getProperties()).forEach(k=>{
    if (k.indexOf('guide:')===0 || k.indexOf('guideById:')===0) P.deleteProperty(k);
  });
  // Indexa
  rows.forEach(r=>{
    const code = String(r[1]||'').trim().toUpperCase();
    const name = String(r[2]||'').trim();
    const email= String(r[3]||'').trim().toLowerCase();
    const id   = String(r[4]||'').trim();
    const url  = String(r[5]||'').trim();
    if (!code || !id) return;
    const payload = JSON.stringify({code,name,email,id,url});
    P.setProperty('guide:'+code, payload);
    P.setProperty('guideById:'+id, code);
  });
}

/** Activa el guardián onEdit en todos los archivos de guía del REGISTRO */
function ensureGuideEditTriggersForAllGuides(){
  const ss = SpreadsheetApp.getActive();
  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!reg) return;
  const rows = reg.getDataRange().getValues().slice(1);
  rows.forEach(r=>{
    const id = String(r[4]||'').trim();
    if (id) ensureGuideEditTriggerForGuide_(id); // función en 06
  });
}
