/***** 08_master_guia_utils_y_emails.gs ********************************/
/* ====== Localización y escritura en calendarios de GUÍA ====== */
function locateGuideCell_(sh,dateObj,slot){
  const year = Number(sh.getName().split('_')[1]);
  const month = Number(sh.getName().split('_')[0]);
  const first = new Date(year,month-1,1);
  const startDow = (first.getDay()+6)%7; // 0=Lun
  const d = new Date(dateObj);
  const day = d.getDate();
  const idx = startDow + (day-1);
  const week = Math.floor(idx/7), dow = idx%7;
  const baseRow = 3 + week*3; // fila de números
  const row = baseRow + (slot==='M'?1:2);
  const col = 1 + dow;
  const numberCell = sh.getRange(baseRow,col).getValue();
  if (String(numberCell).trim() !== String(day)) return null;
  return {row,col};
}

function readGuideStatus_(guideId,tag,dateObj,slot){
  const ss = SpreadsheetApp.openById(guideId);
  const sh = ss.getSheetByName(tag);
  if (!sh) return '';
  const pos = locateGuideCell_(sh,dateObj,slot);
  return pos ? String(sh.getRange(pos.row,pos.col).getValue()).trim() : '';
}

function writeGuideAssignment_(guideId,tag,dateObj,slot,value,lockCell){
  const ss = SpreadsheetApp.openById(guideId);
  const sh = ss.getSheetByName(tag) || createGuideMonthSheet_(ss, fromTabTag_(tag)); // helper en 03/01
  const pos = locateGuideCell_(sh,dateObj,slot);
  if (!pos) return false;

  // Quitar protección previa de esa celda
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p=>{
    const rn=p.getRange();
    if (rn.getRow()==pos.row && rn.getColumn()==pos.col) p.remove();
  });

  // Escribir display: '' => etiqueta del turno
  const display = (value==='' ? (slot==='M'?'MAÑANA':'TARDE') : value);
  const r = sh.getRange(pos.row,pos.col).setValue(display).setNote('');

  // Proteger si asignado
  if (lockCell && value){
    const p = r.protect().setDescription('Asignado por MASTER');
    p.setWarningOnly(false);
    const me = Session.getEffectiveUser();
    p.addEditor(me);
    try {
      const editors = r.getSheet().getEditors().map(e=>e.getEmail && e.getEmail()).filter(Boolean);
      const toRemove = editors.filter(e => e !== me.getEmail());
      if (toRemove.length) p.removeEditors(toRemove);
    } catch(_){}
    try { if (p.canDomainEdit && p.canDomainEdit()) p.setDomainEdit(false); } catch(_){}
  }
  return true;
}

/* ====== Bookeo utils ====== */
function findBookeoEventBySlot_(dateObj,slot){
  const cal = CalendarApp.getCalendarById(CFG.BOOKEO_CAL_ID);
  if (!cal) return null;
  const start = slotStartDate_(dateObj,slot); // helper en 01
  if (!start) return null;
  const end = new Date(start.getTime()+60*1000);
  const evs = cal.getEvents(start,end);
  return evs && evs.length ? evs[0] : null;
}

function inviteGuideToEvent_(event,email){
  try{
    const guests = event.getGuestList().map(g=>g.getEmail());
    if (guests.indexOf(email) === -1) event.addGuest(email);
  }catch(_e){}
}

function getEventHtmlLink_(event){
  try{
    const id=event.getId();
    try{
      const ev=Calendar.Events.get(CFG.BOOKEO_CAL_ID,id);
      if(ev && ev.htmlLink) return ev.htmlLink;
    }catch(_){ // fallback
      const start=event.getStartTime(), end=new Date(start.getTime()+60*1000);
      const res=Calendar.Events.list(CFG.BOOKEO_CAL_ID,{timeMin:start.toISOString(), timeMax:end.toISOString(), singleEvents:true, maxResults:1});
      if(res && res.items && res.items.length && res.items[0].htmlLink) return res.items[0].htmlLink;
    }
  }catch(_e){}
  return '';
}

/* ====== Emails ====== */
function sendMailAssignment_(email,name,dateObj,label,eventLink){
  if (!email) return;
  const fecha = Utilities.formatDate(new Date(dateObj), CFG.TZ, 'dd/MM/yyyy');
  const link = eventLink ? `<p><a href="${eventLink}">Abrir evento en Google Calendar</a></p>` : '';
  MailApp.sendEmail({
    to: email,
    replyTo: 'madrid@spainfoodsherpas',
    name: 'Spain Food Sherpas — Operaciones',
    subject: `Asignado: ${label} — ${fecha}`,
    htmlBody: `<p>Hola ${name||''},</p><p>Se te ha <b>asignado</b>: <b>${label}</b> el <b>${fecha}</b>.</p>${link}`
  });
}

function sendMailLiberation_(email,name,dateObj,slot){
  if (!email) return;
  const fecha = Utilities.formatDate(new Date(dateObj), CFG.TZ, 'dd/MM/yyyy');
  const s = slot==='M'?'MAÑANA':'TARDE';
  MailApp.sendEmail({
    to: email,
    replyTo: 'madrid@spainfoodsherpas',
    name: 'Spain Food Sherpas — Operaciones',
    subject: `Liberación de turno — ${fecha}`,
    htmlBody: `<p>Hola ${name||''},</p><p>Se ha <b>liberado</b> tu turno de <b>${s}</b> el <b>${fecha}</b>.</p>`
  });
}

function sendMailCalendarCreated_(email, name, url){
  if (!email) return;
  const html = [
    `<p>Hola ${name||''},</p>`,
    `<p>Tu calendario ya está creado. Acceso: <a href="${url}">${url}</a></p>`,
    `<p><b>Recuerda:</b> mantén tu calendario actualizado. Marca <b>NO DISPONIBLE</b> cuando no puedas y luego usa <b>LIBERAR</b> si cambia.</p>`,
    `<p>Cuando el manager asigne verás <b>ASIGNADO</b> en verde y bloqueado.</p>`,
    `<p>Soporte: <a href="mailto:madrid@spainfoodsherpas">madrid@spainfoodsherpas</a></p>`
  ].join('');
  MailApp.sendEmail({
    to: email,
    replyTo: 'madrid@spainfoodsherpas',
    name: 'Spain Food Sherpas — Operaciones',
    subject: 'Tu calendario de guías está listo',
    htmlBody: html
  });
}
