/***** 04_validaciones_y_formato.gs ************************************
 * Data validation y formato condicional en MASTER y GUÍAS.
 ***********************************************************************/

function applyGuideDV_(sh) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.GUIDE_DV_LIST, true).build();
  for (let w=0; w<6; w++) {
    const rowM = 3 + w*3 + 1, rowT = 3 + w*3 + 2;
    sh.getRange(rowM,1,1,7).setDataValidation(rule);
    sh.getRange(rowT,1,1,7).setDataValidation(rule);
  }
  const rules = sh.getConditionalFormatRules();
  const body = sh.getRange(3,1,18,7); // 6 semanas * 3 filas = 18
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('ASIGNADO').setBackground(CFG.COLORS.ASSIGNED).setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('NO DISPONIBLE').setBackground(CFG.COLORS.NODISP).setRanges([body]).build()
  );
  sh.setConditionalFormatRules(rules);
}

function applyMasterDataValidations_(sh) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastCol <= 2 || lastRow <= 2) return;
  const numRows = lastRow - 2;
  const mRule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.MASTER_M_LIST, true).build();
  const tRule = SpreadsheetApp.newDataValidation().requireValueInList(CFG.MASTER_T_LIST, true).build();
  for (let c=3; c<=lastCol; c+=2) {
    sh.getRange(3,c,numRows,1).setDataValidation(mRule);
    sh.getRange(3,c+1,numRows,1).setDataValidation(tRule);
  }
  const rules = sh.getConditionalFormatRules();
  const body = sh.getRange(3,3,numRows,Math.max(0,lastCol-2));
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('ASIGNADO').setBackground(CFG.COLORS.ASSIGNED).setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('NO DISPONIBLE').setBackground(CFG.COLORS.NODISP).setRanges([body]).build()
  );
  sh.setConditionalFormatRules(rules);
}
function applyAllMasterDV_(){ SpreadsheetApp.getActive().getSheets().forEach(sh=>{ if (/^\d{2}_\d{4}$/.test(sh.getName())) applyMasterDataValidations_(sh); }); }
