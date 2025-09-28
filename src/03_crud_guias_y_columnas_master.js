/***** 03_crud_guias_y_columnas_master.gs ******************************
 * Alta/baja de guías, creación de archivos y columnas en MASTER.
 ***********************************************************************/

function menuAddGuide() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Nuevo guía','Formato: "Nombre; CODIGO"', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const [name, code] = String(r1.getResponseText()||'').split(/[;,]/).map(s=>s.trim());
  if (!name || !code) throw new Error('Faltan datos: "Nombre; CODIGO".');

  const r2 = ui.prompt('Email del guía','Email de Google del guía', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const email = String(r2.getResponseText()||'').trim();
  if (!email) throw new Error('Falta email.');

  const folder = DriveApp.getFolderById(CFG.GUIDES_FOLDER_ID);
  const guideFile = SpreadsheetApp.create(`GUÍA ${code} — ${name}`);
  DriveApp.getFileById(guideFile.getId()).moveTo(folder);

  const guide = SpreadsheetApp.openById(guideFile.getId());
  buildGuideScaffold_(guide, name, code);

  const reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  reg.appendRow([new Date(), code, name, email, guideFile.getId(), guideFile.getUrl()]);
  P.setProperty('guide:'+code, JSON.stringify({id: guideFile.getId(), email, name, code}));

  SpreadsheetApp.getActive().getSheets().forEach(sh=>{
    if (/^\d{2}_\d{4}$/.test(sh.getName())) { addGuideColumnsToMonth_(sh,{code,name}); }
  });
  applyAllMasterDV_();

  ensureGuideEditTriggerForGuide_(guideFile.getId());
  SpreadsheetApp.getActive().toast(`Guía ${code} añadido`);
}

function buildGuideScaffold_(guideSS, name, code) {
  guideSS.setSpreadsheetTimeZone(CFG.TZ);
  const sh0 = guideSS.getActiveSheet(); sh0.setName('PORTADA');
  sh0.getRange('A1').setValue(`Guía: ${name} (${code})`);
  CFG.MONTHS_INITIAL.forEach(tag => { if (!guideSS.getSheetByName(toTabName_(tag))) createGuideMonthSheet_(guideSS, tag); });
}

function createGuideMonthSheet_(ss, tag) {
  const sh = ss.insertSheet(toTabName_(tag));
  const [y, m] = tag.split('-').map(Number);
  const {grid, labels} = buildMonthlyGrid_(y, m);
  sh.getRange(1,1,1,7).setValues([labels]);
  sh.setFrozenRows(2);
  sh.getRange(3,1,grid.length,7).setValues(grid);
  applyGuideDV_(sh);
}

function addGuideColumnsToMonth_(sh, {code, name}) {
  // Evitar duplicados: buscar cabecera existente "CODE — NAME"
  const lastCol = Math.max(2, sh.getLastColumn());
  for (let c=3; c<=lastCol; c+=2) {
    const head = String(sh.getRange(1,c).getValue()||'').trim();
    if (head.startsWith(code)) return; // ya existe
  }
  sh.insertColumnsAfter(lastCol, 2);
  const startCol = lastCol + 1;
  sh.getRange(1, startCol, 1, 2).merge().setValue(`${code} — ${name}`);
  sh.getRange(2, startCol).setValue('MAÑANA'); 
  sh.getRange(2, startCol+1).setValue('TARDE');
}

// ---------- Baja completa ----------
function menuRemoveGuide() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.prompt('Eliminar guía','Introduce CODIGO o EMAIL', ui.ButtonSet.OK_CANCEL);
  if (ans.getSelectedButton() !== ui.Button.OK) return;
  const token = String(ans.getResponseText()||'').trim();
  if (!token) return;

  const reg = SpreadsheetApp.getActive().getSheetByName(CFG.REGISTRY_SHEET);
  const data = reg.getDataRange().getValues();
  let idx = -1, row = null;
  for (let i=1;i<data.length;i++){
    if (String(data[i][1]).trim()===token || String(data[i][3]).trim()===token){ idx=i; row=data[i]; break; }
  }
  if (idx<0) { ui.alert('No encontrado en REGISTRO'); return; }

  const code = String(row[1]).trim(), name = String(row[2]).trim(), guideId = String(row[4]).trim();
  // 1) Eliminar archivo (a papelera)
  try { DriveApp.getFileById(guideId).setTrashed(true); } catch(e){}

  // 2) Borrar columnas en todos los meses
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh=>{
    if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
    const lastCol = sh.getLastColumn();
    for (let c=3; c<=lastCol; c+=2){
      const head = String(sh.getRange(1,c).getValue()||'').trim();
      if (head.startsWith(code)) { sh.deleteColumns(c,2); break; }
    }
  });

  // 3) Quitar del REGISTRO
  reg.deleteRow(idx+1);

  // 4) Borrar propiedades y triggers asociados
  P.deleteProperty('guide:'+code);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction()==='guideEditHandler_' && t.getTriggerSourceId && t.getTriggerSourceId()===guideId)
    .forEach(t => ScriptApp.deleteTrigger(t));

  SpreadsheetApp.getActive().toast(`Guía ${code} eliminado`);
}

// ---------- Sincronizar meses a guías ----------
function syncMonthsToGuides() {
  const ss = SpreadsheetApp.getActive();
  ensureInitialMonths_();

  // Para cada guía del REGISTRO, asegurar meses y DV en GUÍA y columnas en MASTER
  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET);
  const rows = reg.getDataRange().getValues().slice(1).filter(r=>r[1] && r[4]);
  const months = CFG.MONTHS_INITIAL.map(toTabName_);

  rows.forEach(r=>{
    const code = String(r[1]).trim(), name = String(r[2]).trim(), id = String(r[4]).trim();
    let gss; try { gss = SpreadsheetApp.openById(id); } catch(e){ return; }
    months.forEach(tab=>{
      if (!gss.getSheetByName(tab)) createGuideMonthSheet_(gss, fromTabTag_(tab));
      else applyGuideDV_(gss.getSheetByName(tab)); // reasegurar DV/colores
      const mSh = ss.getSheetByName(tab); if (mSh) addGuideColumnsToMonth_(mSh, {code,name});
    });
  });

  applyAllMasterDV_();
  ss.toast('Meses y DV sincronizados');
}
