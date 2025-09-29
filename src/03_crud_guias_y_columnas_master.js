/***** 03_crud_guias_y_columnas_master.gs ******************************
 * Alta/baja de guías, creación de archivos y columnas en MASTER.
 ***********************************************************************/

function menuAddGuide() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Nuevo guía','Formato: "Nombre; CODIGO"', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const parts = String(r1.getResponseText()||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
  const name = parts[0]||''; const code = parts[1]||'';
  if (!name || !code) throw new Error('Faltan datos: "Nombre; CODIGO".');

  const r2 = ui.prompt('Email del guía','Email de Google del guía', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const email = String(r2.getResponseText()||'').trim();
  if (!email) throw new Error('Falta email.');

  addGuide_({code, name, email});
}

function addGuide_({code,name,email}){
  const folderId = (CFG.GUIDES_FOLDER_ID || CFG.DEST_FOLDER_ID);
  const folder = DriveApp.getFolderById(folderId);
  const ssGuide = SpreadsheetApp.create(`CALENDARIO_${name}-${code}`);
  const file = DriveApp.getFileById(ssGuide.getId());
  folder.addFile(file); DriveApp.getRootFolder().removeFile(file);

  buildGuideScaffold_(ssGuide, name, code); // crea meses iniciales
  const url = ssGuide.getUrl();

  // Registrar en REGISTRO
  const master = SpreadsheetApp.getActive();
  const reg = master.getSheetByName(CFG.REGISTRY_SHEET) || master.insertSheet(CFG.REGISTRY_SHEET);
  if (reg.getLastRow() === 0) reg.getRange(1,1,1,CFG.REGISTRY_HEADERS.length).setValues([CFG.REGISTRY_HEADERS]);
  reg.appendRow([new Date(), code, name, email, ssGuide.getId(), url]);

  // Índices en ScriptProperties
  P.setProperty('guide:'+code, JSON.stringify({code,name,email,id:ssGuide.getId(),url}));
  P.setProperty('guideById:'+ssGuide.getId(), code);

  // Añadir columnas del guía en todos los meses del MASTER (idempotente)
  master.getSheets().forEach(sh=>{
    if (/^\d{2}_\d{4}$/.test(sh.getName())) addGuideColumnsToMonth_(sh,{code,name});
  });
  applyAllMasterDV_(); // en 04

  // Disparador onEdit para ese Spreadsheet de guía (guardían)
  ensureGuideEditTriggerForGuide_(ssGuide.getId()); // en 06

  // Email al guía
  sendMailCalendarCreated_(email, name, url); // en 08

  master.toast(`Guía ${code} añadido`);
}

function menuRemoveGuide(){
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Eliminar guía','Escribe el CÓDIGO del guía a eliminar (también se enviará el archivo a la papelera)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const code = String(r.getResponseText()||'').trim();
  if (!code) return;
  removeGuide_(code);
}

/** Elimina columnas en MASTER, REGISTRO, propiedades, triggers y envía a papelera el archivo del guía. */
function removeGuide_(code){
  const ss = SpreadsheetApp.getActive();
  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET);
  let guideId = '', rowToDelete = -1;

  if (reg){
    const data = reg.getDataRange().getValues();
    for (let i=1; i<data.length; i++){
      if (String(data[i][1]).trim() === code){
        guideId = String(data[i][4]||'').trim();
        rowToDelete = i+1;
        break;
      }
    }
  }

  // 1) Borrar columnas en cada mes
  ss.getSheets().forEach(sh => {
    if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
    const colStart = findGuideColumnsInMonth_(sh, code);
    if (colStart) sh.deleteColumns(colStart, 2);
  });

  // 2) Borrar fila de REGISTRO
  if (reg && rowToDelete > 1) reg.deleteRow(rowToDelete);

  // 3) Borrar propiedades índice
  Object.keys(P.getProperties()).forEach(k=>{
    if (k === 'guide:'+code) P.deleteProperty(k);
    if (k.indexOf('guideById:')===0 && P.getProperty(k)===code) P.deleteProperty(k);
  });

  // 4) Eliminar triggers del proyecto asociados a ese Spreadsheet de guía
  if (guideId){
    ScriptApp.getProjectTriggers()
      .filter(t => t.getTriggerSourceId && t.getTriggerSourceId() === guideId)
      .forEach(t => ScriptApp.deleteTrigger(t));
  }

  // 5) Enviar archivo del guía a papelera
  if (guideId){
    try { DriveApp.getFileById(guideId).setTrashed(true); } catch(e){}
  }

  ss.toast(`Guía ${code} eliminado por completo`);
}

/** Inserta columnas M/T para un guía en el MASTER si no existen */
function addGuideColumnsToMonth_(mSheet,{code,name}){
  const existsAt = findGuideColumnsInMonth_(mSheet, code);
  if (existsAt) return; // idempotente

  const lastCol = Math.max(2, mSheet.getLastColumn());
  mSheet.insertColumnsAfter(lastCol, 2);
  const colM = lastCol+1, colT = lastCol+2;
  mSheet.getRange(1,colM).setValue(`${code} — ${name}`);
  mSheet.getRange(2,colM).setValue('MAÑANA');
  mSheet.getRange(2,colT).setValue('TARDE');
}

/** Localiza el par de columnas del guía en un mes del MASTER */
function findGuideColumnsInMonth_(mSheet, code){
  const lastCol = mSheet.getLastColumn();
  for (let c=3; c<=lastCol; c+=2){
    const head = String(mSheet.getRange(1,c).getValue()||'').trim();
    if (head && head.split('—')[0].trim() === code) return c;
  }
  return 0;
}

/** Sincroniza meses/pestañas entre MASTER y guías y asegura DV/formato */
function syncMonthsToGuides(){
  const ss = SpreadsheetApp.getActive();
  ensureInitialMonths_(); // en 02

  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!reg) { ss.toast('No hay REGISTRO'); return; }
  const rows = reg.getDataRange().getValues().slice(1).filter(r=>r[1] && r[4]);
  const months = CFG.MONTHS_INITIAL.map(toTabName_);

  rows.forEach(r=>{
    const code = String(r[1]).trim(), name = String(r[2]).trim(), id = String(r[4]).trim();
    let gss; try { gss = SpreadsheetApp.openById(id); } catch(e){ return; }
    months.forEach(tab=>{
      if (!gss.getSheetByName(tab)) createGuideMonthSheet_(gss, fromTabTag_(tab));
      applyGuideDV_(gss.getSheetByName(tab)); // en 04
      const mSh = ss.getSheetByName(tab);
      if (mSh) addGuideColumnsToMonth_(mSh, {code,name}); // idempotente
    });
  });

  applyAllMasterDV_(); // en 04
  ss.toast('Meses y DV sincronizados');
}

/** Construye portada y meses iniciales en el archivo del guía */
function buildGuideScaffold_(guideSS, name, code) {
  guideSS.setSpreadsheetTimeZone(CFG.TZ);
  const sh0 = guideSS.getActiveSheet(); sh0.setName('PORTADA');
  sh0.getRange('A1').setValue(`Guía: ${name} (${code})`);
  CFG.MONTHS_INITIAL.forEach(tag => { if (!guideSS.getSheetByName(toTabName_(tag))) createGuideMonthSheet_(guideSS, tag); });
}

/** Crea la pestaña mensual del guía con cuadrícula Lun–Dom */
function createGuideMonthSheet_(ss, tag) {
  const sh = ss.insertSheet(toTabName_(tag));
  const [y, m] = tag.split('-').map(Number);
  const {grid, labels} = buildMonthlyGrid_(y, m); // en 01
  sh.getRange(1,1,1,7).setValues([labels]);
  sh.setFrozenRows(2);
  sh.getRange(3,1,grid.length,7).setValues(grid);
  applyGuideDV_(sh); // en 04
}
