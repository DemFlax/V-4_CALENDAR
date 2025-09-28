/***** 01_config_y_helpers.gs ******************************************/
const CFG = {
  TZ: 'Europe/Madrid',
  REGISTRY_SHEET: 'REGISTRO',
  REGISTRY_HEADERS: ['TIMESTAMP','CODIGO','NOMBRE','EMAIL','FILE_ID','URL'],
  MASTER_M_LIST: ['', 'LIBERAR', 'ASIGNAR M'],
  MASTER_T_LIST: ['', 'LIBERAR', 'ASIGNAR T1', 'ASIGNAR T2', 'ASIGNAR T3'],
  GUIDE_DV_LIST: ['', 'NO DISPONIBLE', 'LIBERAR'],
  MONTHS_INITIAL: ['2025-10','2025-11','2025-12'],
  COLORS: { ASSIGNED: '#A5D6A7', NODISP: '#EF9A9A' },
  GUIDES_FOLDER_ID: '1ibz8PUeaFbUraTgRS9VgfjZ_hqs80J-p',
  BOOKEO_CAL_ID: 'c_61981c641dc3c970e63f1713ccc2daa49d8fe8962b6ed9f2669c4554496c7bdd@group.calendar.google.com',
  SLOT_TIME: { M: '12:00', T1: '17:15', T2: '18:15', T3: '19:15' }
};
const P = PropertiesService.getScriptProperties();
const LOCK = LockService.getScriptLock();
function toTabName_(tag){ const [y,m]=tag.split('-').map(Number); return `${String(m).padStart(2,'0')}_${y}`; }
function fromTabTag_(tab){ const [mm,yyyy]=tab.split('_'); return `${yyyy}-${mm}`; }
function enumerateMonthDates_(year, month){ const a=[], last=new Date(year, month, 0).getDate(); for(let d=1; d<=last; d++) a.push(new Date(year, month-1, d)); return a; }
function getWeekdayShort_(d){ return ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'][new Date(d).getDay()]; }
function slotStartDate_(dateObj, slot){ const t = CFG.SLOT_TIME[slot]; if (!t) return null; const [hh,mm]=t.split(':').map(Number); const d=new Date(dateObj); d.setHours(hh,mm,0,0); return d; }
function buildMonthlyGrid_(year, month){
  const first = new Date(year, month-1, 1);
  const startDow = (first.getDay()+6)%7; // 0=Lun
  const lastDay = new Date(year, month, 0).getDate();
  const labels = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];

  const grid = []; let cursor = 1 - startDow;
  for (let w=0; w<6; w++){
    const nums = Array(7).fill('').map((_,i)=>{
      const d = cursor + i; return d>=1 && d<=lastDay ? String(d) : '';
    });
    const rowM = nums.map(n => n ? 'MAÑANA' : '');
    const rowT = nums.map(n => n ? 'TARDE'  : '');
    grid.push(nums, rowM, rowT);
    cursor += 7;
  }
  return {grid, labels};
}

