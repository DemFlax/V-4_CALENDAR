/***** 04_validaciones_y_formato.gs ************************************
 * DV del MASTER y de los CALENDARIOS DE GUÍA + formato visual.
 * Seguro con valores fuera de lista (ASIGNADO…, MAÑANA/TARDE) – allowInvalid=true.
 **********************************************************************/

/** ========== HELPERS ========== */
function dvList_(values, showDropdown, allowInvalid){
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(values, showDropdown !== false)
    .setAllowInvalid(allowInvalid !== false) // true por defecto
    .build();
}
function isMonthTab_(name){ return /^\d{2}_\d{4}$/.test(name); }

/** ========== DATA VALIDATION ========== */
/** MASTER: aplica DV a todas las pestañas MM_YYYY */
function applyAllMasterDV_(){
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => { if (isMonthTab_(sh.getName())) applyMasterDV_(sh); });
}

/** MASTER: DV por columnas pares (Mañana/Tarde) */
function applyMasterDV_(sh){
  const lastRow = Math.max(3, sh.getLastRow());
  const lastCol = Math.max(3, sh.getLastColumn());
  if (lastCol < 3 || lastRow < 3) return;

  // Menús del manager
  const dvM = dvList_(['', 'ASIGNAR M', 'LIBERAR'], true, true);
  const dvT = dvList_(['', 'ASIGNAR T1', 'ASIGNAR T2', 'ASIGNAR T3', 'LIBERAR'], true, true);

  const rangesM = [];
  const rangesT = [];
  for (let c=3; c<=lastCol; c+=2){
    rangesM.push(sh.getRange(3, c, lastRow-2, 1));     // MAÑANA
    if (c+1 <= lastCol) rangesT.push(sh.getRange(3, c+1, lastRow-2, 1)); // TARDE
  }

  if (rangesM.length) sh.getRangeList(rangesM.map(r=>r.getA1Notation())).setDataValidation(dvM);
  if (rangesT.length) sh.getRangeList(rangesT.map(r=>r.getA1Notation())).setDataValidation(dvT);
}

/** GUÍA: DV para todas las celdas de MAÑANA/TARDE en cuadrícula */
function applyGuideDV_(sh){
  const lastRow = Math.max(3, sh.getLastRow());
  if (lastRow < 5) { formatGuideMonth_(sh); return; }

  const dvGuide = dvList_(['NO DISPONIBLE', 'LIBERAR'], true, true);

  // Filas de MAÑANA y TARDE: bloques de 3 filas (número, MAÑANA, TARDE)
  const ranges = [];
  for (let base=3; base<=lastRow; base+=3){
    if (base+1 <= lastRow) ranges.push(sh.getRange(base+1, 1, 1, 7)); // MAÑANA
    if (base+2 <= lastRow) ranges.push(sh.getRange(base+2, 1, 1, 7)); // TARDE
  }
  if (ranges.length) sh.getRangeList(ranges.map(r=>r.getA1Notation())).setDataValidation(dvGuide);

  // Formato visual
  formatGuideMonth_(sh);
}

/** ========== FORMATO VISUAL ========== */
/** Aplica formato a todas las pestañas MM_YYYY del MASTER y a todos los guías */
function applyAllFormatting_(){
  const ss = SpreadsheetApp.getActive();

  // MASTER
  ss.getSheets().forEach(sh=>{
    if (isMonthTab_(sh.getName())) formatMasterMonth_(sh);
  });

  // GUÍAS desde REGISTRO
  const reg = ss.getSheetByName(CFG.REGISTRY_SHEET || 'REGISTRO');
  if (!reg) return;
  const rows = reg.getDataRange().getValues().slice(1);
  rows.forEach(r=>{
    const gid = String(r[4]||'').trim();
    if (!gid) return;
    let gss; try { gss = SpreadsheetApp.openById(gid); } catch(e){ return; }
    gss.getSheets().forEach(gsh=>{
      if (isMonthTab_(gsh.getName())) formatGuideMonth_(gsh);
    });
  });
}

/** MASTER: oculta “DÍA”, banding, reglas, centrado */
function formatMasterMonth_(sh){
  const lastRow = Math.max(3, sh.getLastRow());
  const lastCol = Math.max(3, sh.getLastColumn());
  if (lastCol < 3) return;

  sh.setFrozenRows(2);
  try { sh.hideColumns(2); } catch(_) {}

  // Banding sobre datos
  try {
    (sh.getBandings()||[]).forEach(b=>b.remove());
    sh.getRange(3,1,lastRow-2,lastCol)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  } catch(_){}

  // Reglas de color
  const data = sh.getRange(3,3,lastRow-2,Math.max(0,lastCol-2));
  const keep = (sh.getConditionalFormatRules()||[]).filter(r=>{
    const rs = r.getRanges()[0];
    return !(rs && rs.getRow()==3 && rs.getColumn()==3);
  });
  const red = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('NO DISPONIBLE')
    .setBackground('#F8D7DA').setFontColor('#7A1F23')
    .setRanges([data]).build();
  const green = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('ASIGNADO')
    .setBackground('#C6E6C3').setFontColor('#185C37')
    .setRanges([data]).build();
  sh.setConditionalFormatRules(keep.concat([red, green]));

  data.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
}

/** GUÍA: cabecera, fines de semana, bordes, reglas */
function formatGuideMonth_(sh){
  const lastRow = Math.max(3, sh.getLastRow());

  try { sh.setHiddenGridlines(true); } catch(_) {}
  sh.getRange(1,1,1,7).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(2);

  const block = sh.getRange(3,1,lastRow-2,7);
  block.setBorder(true,true,true,true,true,true);

  // Sombrear Sáb (F) y Dom (G)
  sh.getRange(3,6,lastRow-2,1).setBackground('#F3F3F3');
  sh.getRange(3,7,lastRow-2,1).setBackground('#F3F3F3');

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('NO DISPONIBLE')
      .setBackground('#F8D7DA').setFontColor('#7A1F23')
      .setRanges([block]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ASIGNADO')
      .setBackground('#C6E6C3').setFontColor('#185C37')
      .setRanges([block]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('MAÑANA')
      .setBackground('#EFEFEF').setRanges([block]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('TARDE')
      .setBackground('#EFEFEF').setRanges([block]).build()
  ]);

  block.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
}
