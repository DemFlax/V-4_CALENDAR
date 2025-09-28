/***** 04_validaciones_y_formato.gs ************************************
 * DV y formato en MASTER y GUÍAS.
 * - GUÍA: por defecto muestra "MAÑANA"/"TARDE" en todas las celdas de turno.
 * - Al liberar, vuelve a "MAÑANA"/"TARDE".
 * - Respeta "NO DISPONIBLE" y "ASIGNADO".
 ***********************************************************************/

/** Data validation y formato en GUÍA */
function applyGuideDV_(sh) {
  // Desplegable guía: '', 'NO DISPONIBLE', 'LIBERAR' pero permitimos ver otros valores.
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CFG.GUIDE_DV_LIST, true)
    .setAllowInvalid(true)
    .build();

  // Aplicar DV a 6 semanas (filas: números, MAÑANA, TARDE)
  for (let w = 0; w < 6; w++) {
    const rowM = 3 + w * 3 + 1, rowT = 3 + w * 3 + 2;
    sh.getRange(rowM, 1, 1, 7).setDataValidation(rule);
    sh.getRange(rowT, 1, 1, 7).setDataValidation(rule);
  }

  // Colores por texto
  const rules = sh.getConditionalFormatRules();
  const body = sh.getRange(3, 1, 18, 7); // 6 semanas * 3 filas
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ASIGNADO')
      .setBackground(CFG.COLORS.ASSIGNED)
      .setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('NO DISPONIBLE')
      .setBackground(CFG.COLORS.NODISP)
      .setRanges([body]).build()
  );
  sh.setConditionalFormatRules(rules);

  // Asegurar etiquetas visibles "MAÑANA"/"TARDE" en celdas libres o liberadas.
  ensureGuideSlotLabels_(sh);
}

/** Fuerza etiquetas de turno en celdas con día válido que estén vacías o en "LIBERAR". */
function ensureGuideSlotLabels_(sh) {
  for (let w = 0; w < 6; w++) {
    const base = 3 + w * 3; // fila de números
    for (let c = 1; c <= 7; c++) {
      const hasDay = String(sh.getRange(base, c).getValue() || '').trim() !== '';
      if (!hasDay) continue;

      // MAÑANA
      const rM = sh.getRange(base + 1, c);
      const vM = String(rM.getValue() || '').trim();
      if (vM === '' || vM === 'LIBERAR') rM.setValue('MAÑANA');

      // TARDE
      const rT = sh.getRange(base + 2, c);
      const vT = String(rT.getValue() || '').trim();
      if (vT === '' || vT === 'LIBERAR') rT.setValue('TARDE');
    }
  }
}

/** Data validation y formato en MASTER */
function applyMasterDataValidations_(sh) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastCol <= 2 || lastRow <= 2) return;

  const numRows = lastRow - 2;
  const mRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CFG.MASTER_M_LIST, true).build();
  const tRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CFG.MASTER_T_LIST, true).build();

  for (let c = 3; c <= lastCol; c += 2) {
    sh.getRange(3, c,   numRows, 1).setDataValidation(mRule); // MAÑANA
    sh.getRange(3, c+1, numRows, 1).setDataValidation(tRule); // TARDE
  }

  const rules = sh.getConditionalFormatRules();
  const body = sh.getRange(3, 3, numRows, Math.max(0, lastCol - 2));
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ASIGNADO')
      .setBackground(CFG.COLORS.ASSIGNED)
      .setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('NO DISPONIBLE')
      .setBackground(CFG.COLORS.NODISP)
      .setRanges([body]).build()
  );
  sh.setConditionalFormatRules(rules);
}

/** Aplicar DV a todas las pestañas mensuales del MASTER */
function applyAllMasterDV_() {
  SpreadsheetApp.getActive().getSheets()
    .forEach(sh => { if (/^\d{2}_\d{4}$/.test(sh.getName())) applyMasterDataValidations_(sh); });
}
