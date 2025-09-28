/***** 08_master_guia_utils_y_emails.gs ********************************
 * Localización de celdas, lecturas/escrituras en GUÍA y emails.
 ***********************************************************************/

function locateGuideCell_(sh, dateObj, slot) {
  const year = Number(sh.getName().split('_')[1]);
  const month = Number(sh.getName().split('_')[0]);
  const first = new Date(year, month-1, 1);
  const startDow = (first.getDay()+6)%7; // 0=Lun
  const d = new Date(dateObj); const day = d.getDate();
  const idx = startDow + (day-1);
  const week = Math.floor(idx/7), dow = idx%7;
  const baseRow = 3 + week*3;
  const row = baseRow + (slot==='M' ? 1 : 2);
  const col = 1 + dow;
  const numberCell = sh.getRange(baseRow, col).getValue();
  if (String(numberCell).trim() !== String(day)) return null;
  return {row, col};
}

function readGuideStatus_(guideId, tag, dateObj, slot) {
  const ss = SpreadsheetApp.openById(guideId);
  const sh = ss.getSheetByName(tag); if (!sh) return '';
  const pos = locateGuideCell_(sh, dateObj, slot);
  return pos ? String(sh.getRange(pos.row, pos.col).getValue()).trim() : '';
}

function writeGuideAssignment_(guideId, tag, dateObj, slot, value, lockCell) {
  const ss = SpreadsheetApp.openById(guideId);
  const sh = ss.getSheetByName(tag) || createGuideMonthSheet_(ss, fromTabTag_(tag));
  const pos = locateGuideCell_(sh, dateObj, slot); if (!pos) return;

  // limpia protección previa
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p=>{ const rn=p.getRange(); if (rn.getRow()==pos.row && rn.getColumn()==pos.col) p.remove(); });
  const r = sh.getRange(pos.row, pos.col).setValue(value);

  if (lockCell && value) {
    const p = r.protect().setDescription('Asignado por MASTER');
    p.setWarningOnly(false);
    const me = Session.getEffectiveUser(); p.addEditor(me);
    const toRemove = p.getEditors().filter(u=>u.getEmail()!==me.getEmail()); if (toRemove.length) p.removeEditors(toRemove);
    if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false);
  }
}

/* ===== Emails ===== */
function sendMailAssignment_(email, name, dateObj, label) {
  if (!email) return;
  const fecha = Utilities.formatDate(new Date(dateObj), CFG.TZ, 'dd/MM/yyyy');
  MailApp.sendEmail({ to: email, subject: `Asignado: ${label} — ${fecha}`, htmlBody: `<p>Hola ${name},</p><p>Se te ha <b>asignado</b>: <b>${label}</b> el <b>${fecha}</b>.</p>` });
}
function sendMailLiberation_(email, name, dateObj, slot) {
  if (!email) return;
  const fecha = Utilities.formatDate(new Date(dateObj), CFG.TZ, 'dd/MM/yyyy');
  const s = slot==='M' ? 'MAÑANA' : 'TARDE';
  MailApp.sendEmail({ to: email, subject: `Liberación de turno — ${fecha}`, htmlBody: `<p>Hola ${name},</p><p>Se ha <b>liberado</b> tu turno de <b>${s}</b> el <b>${fecha}</b>.</p>` });
}
