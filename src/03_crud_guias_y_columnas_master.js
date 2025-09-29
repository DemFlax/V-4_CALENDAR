/***** 03_crud_guias_y_columnas_master.gs ******************************
 * Alta/baja de guías, creación de archivos y columnas en MASTER.
 * Bloqueo de duplicados en el alta. Auto-compartir por archivo.
 ***********************************************************************/

function menuAddGuide() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Nuevo guía','Formato: "Nombre; CODIGO"', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const parts = String(r1.getResponseText()||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
  const name = parts[0]||''; const code = (parts[1]||'').toUpperCase();
  if (!name || !code) throw new Error('Faltan datos: "Nombre; CODIGO".');

  const r2 = ui.prompt('Email del guía','Email de Google del guía', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const email = String(r2.getResponseText()||'').trim().toLowerCase();
  if (!email) throw new Error('Falta email.');

  addGuide_({code, name, email});
}

function addGuide_({code,name,email}){
  // Normaliza
  code = String(code||'').trim().toUpperCase();
  name = String(name||'').trim();
  email= String(email||'').trim().toLowerCase();

  // Bloqueos de duplicados en REGISTRO
  const master = SpreadsheetApp.getActive();
  const reg = master.getSheetByName(CFG.REGISTRY_SHEET) || master.insertSheet(CFG.REGISTRY_SHEET);
  if (reg.getLastRow() === 0) reg.getRange(1,1,1,CFG.REGISTRY_HEADERS.length).setValues([CFG.REGISTRY_HEADERS]);
  const rows = reg.getDataRange().getValues().slice(1);

  const codeExists = rows.some(r => String(r[1]||'').trim().toUpperCase() === code);
  if (codeExists) { master.toast('Código ya existente. Operación cancelada.'); return; }

  const emailExists = rows.some(r => String(r[3]||'').trim().toLowerCase() === email);
  if (emailExists) { master.toast('Email ya usado por otro guía. Operación cancelada.'); return; }

  if (P.getProperty('guide:'+code)) { master.toast('Código ya indexado. Operación cancelada.'); return; }

  // Crear archivo del guía
  const folderId = (CFG.GUIDES_FOLDER_ID || CFG.DEST_FOLDER_ID);
  const folder = DriveApp.getFolderById(folderId);
  const ssGuide = SpreadsheetApp.create(`CALENDARIO_${name}-${code}`);
  const file = DriveApp.getFileById(ssGuide.getId());
  folder.addFile(file); DriveApp.getRootFolder().removeFile(file);

  // Seguridad y compartición automática: solo manager + guía
  lockDownGuideFile_(ssGuide.getId(), email); // <<<<<<

  // Scaffold + meses
  buildGuideScaffold_(ssGuide, name, code);
  const url = ssGuide.getUrl();

  // Registrar en REGISTRO
  reg.appendRow([new Date(), code, name, email, ssGuide.getId(), url]);

  // Índices
  P.setProperty('guide:'+code, JSON.stringify({code,name,email,id:ssGuide.getId(),url}));
  P.setProperty('guideById:'+ssGuide.getId(), code);

  // Añadir columnas idempotentes en todos los meses
  master.getSheets().forEach(sh=>{
    if (/^\d{2}_\d{4}$/.test(sh.getName())) addGuideColumnsToMonth_(sh,{code,name});
  });
  applyAllMasterDV_(); // en 04

  // Trigger guardián en el archivo del guía
  ensureGuideEditTriggerForGuide_(ssGuide.getId()); // en 06

  // Email al guía
  sendMailCalendarCreated_(email, name, url); // en 08

  master.toast(`Guía ${code} añadido`);
}

function menuRemoveGuide(){
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Eliminar guía','Escribe el CÓDIGO del guía a eliminar (se enviará el archivo a la papelera)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const code = String(r.getResponseText()||'').trim().toUpperCase();
  if (!code) return;
  removeGuide_(code, {deleteAllWithSameCode:true});
}

/** Elimina columnas en MASTER, REGISTRO, propiedades, triggers y envía a papelera el archivo del guía. */
function removeGuide_(code, opts){
  code = String(code||'').trim().toUpperCase();
  const deleteAll = !!(opts && opts.deleteAllWithSameCode);

  const ss = SpreadsheetApp.getActive();
  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET);
  if (!reg) { ss.toast('No hay REGISTRO'); return; }
  const data = reg.getDataRange().getValues();

  // Recoge filas con ese código
  const rowsIdx = [];
  for (let i=1;i<data.length;i++){
    if (String(data[i][1]||'').trim().toUpperCase() === code) rowsIdx.push(i+1);
  }
  if (!rowsIdx.length){ ss.toast('Código no encontrado'); return; }

  // Borra columnas en cada mes del MASTER una sola vez
  ss.getSheets().forEach(sh => {
    if (!/^\d{2}_\d{4}$/.test(sh.getName())) return;
    const colStart = findGuideColumnsInMonth_(sh, code);
    if (colStart) sh.deleteColumns(colStart, 2);
  });

  // Para cada fila encontrada: limpia propiedades, triggers y mueve archivo a papelera
  const toDeleteRows = deleteAll ? rowsIdx.slice() : [rowsIdx[0]];
  toDeleteRows.sort((a,b)=>b-a).forEach(rowNum=>{
    const row = reg.getRange(rowNum,1,1,reg.getLastColumn()).getValues()[0];
    const guideId = String(row[4]||'').trim();
    if (guideId){
      ScriptApp.getProjectTriggers()
        .filter(t => t.getTriggerSourceId && t.getTriggerSourceId() === guideId)
        .forEach(t => ScriptApp.deleteTrigger(t));
      try { DriveApp.getFileById(guideId).setTrashed(true); } catch(_){}
      P.deleteProperty('guideById:'+guideId);
    }
    reg.deleteRow(rowNum);
  });

  P.deleteProperty('guide:'+code);
  ss.toast(`Guía ${code} eliminado`);
}

/** Inserta columnas M/T para un guía en el MASTER si no existen */
function addGuideColumnsToMonth_(mSheet,{code,name}){
  code = String(code||'').trim().toUpperCase();
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
  code = String(code||'').trim().toUpperCase();
  const lastCol = mSheet.getLastColumn();
  for (let c=3; c<=lastCol; c+=2){
    const head = String(mSheet.getRange(1,c).getValue()||'').trim();
    if (!head) continue;
    const cCode = head.split('—')[0].trim().toUpperCase();
    if (cCode === code) return c;
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
    const code = String(r[1]).trim().toUpperCase(), name = String(r[2]).trim(), id = String(r[4]).trim();
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

/** ===== Compartición segura por archivo =====
 *  Restringe el archivo y da edición solo al guía. Evita que editores re-compartan.
 */
function lockDownGuideFile_(fileId, guideEmail){
  const file = DriveApp.getFileById(fileId);
  // Restringido, sin enlace público ni de dominio
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); // restricción total :contentReference[oaicite:2]{index=2}
  // Evitar que editores vuelvan a compartir
  file.setShareableByEditors(false); // :contentReference[oaicite:3]{index=3}
  // Añadir al guía como editor
  file.addEditor(guideEmail); // :contentReference[oaicite:4]{index=4}
  // Limpieza de espectadores/editores extra
  const me = Session.getEffectiveUser().getEmail();
  file.getEditors().forEach(u => {
    const e = u.getEmail && u.getEmail();
    if (e && e !== guideEmail && e !== me) file.removeEditor(u);
  });
  file.getViewers().forEach(u => file.removeViewer(u));
}
