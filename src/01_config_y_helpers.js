/***** 01_config_y_helpers.gs ******************************************
 * Tours Madrid — CORE v1.0.9
 * - Constantes, propiedades, lock y utilidades comunes.
 ***********************************************************************/

const CFG = {
  TZ: 'Europe/Madrid',
  REGISTRY_SHEET: 'REGISTRO',
  REGISTRY_HEADERS: ['TIMESTAMP','CODIGO','NOMBRE','EMAIL','FILE_ID','URL'],
  MASTER_M_LIST: ['', 'LIBERAR', 'ASIGNAR M'],
  MASTER_T_LIST: ['', 'LIBERAR', 'ASIGNAR T1', 'ASIGNAR T2', 'ASIGNAR T3'],
  GUIDE_DV_LIST: ['', 'NO DISPONIBLE', 'LIBERAR'],
  MONTHS_INITIAL: ['2025-10','2025-11','2025-12'],
  COLORS: { ASSIGNED: '#A5D6A7', NODISP: '#EF9A9A' },
  GUIDES_FOLDER_ID: '1ibz8PUeaFbUraTgRS9VgfjZ_hqs80J-p'
};

const P = PropertiesService.getScriptProperties();
const LOCK = LockService.getScriptLock();

// ---------- Helpers de fechas/nombres ----------
function toTabName_(tag){ const [y,m]=tag.split('-').map(Number); return `${String(m).padStart(2,'0')}_${y}`; }
function fromTabTag_(tab){ const [mm,yyyy]=tab.split('_'); return `${yyyy}-${mm}`; }
function enumerateMonthDates_(year, month){ const a=[], last=new Date(year, month, 0).getDate(); for(let d=1; d<=last; d++) a.push(new Date(year, month-1, d)); return a; }
function getWeekdayShort_(d){ return ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'][new Date(d).getDay()]; }

// ---------- Grid mensual para GUÍA (6 semanas x 7 días; filas: Nº, M, T) ----------
function buildMonthlyGrid_(year, month){
  const first = new Date(year, month-1, 1);
  const startDow = (first.getDay()+6)%7; // 0=Lun
  const lastDay = new Date(year, month, 0).getDate();
  const labels = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];

  const grid = [];
  let cursor = 1 - startDow; // puede empezar negativo
  for (let w=0; w<6; w++){
    const nums = new Array(7).fill('').map((_,i)=>{
      const day = cursor + i;
      return day>=1 && day<=lastDay ? String(day) : '';
    });
    const rowM = new Array(7).fill('');
    const rowT = new Array(7).fill('');
    grid.push(nums, rowM, rowT);
    cursor += 7;
  }
  return {grid, labels};
}
