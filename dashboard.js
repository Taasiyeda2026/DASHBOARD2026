function excelSerialToJSDate(serial) {
  if (!serial || isNaN(serial)) return null;
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const msPerDay = 86400000;
  const date = new Date(excelEpoch.getTime() + serial * msPerDay);
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function excelDecimalToTime(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') {
    const totalMinutes = Math.round(v * 24 * 60);
    const h = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    const m = String(totalMinutes % 60).padStart(2, '0');
    return `${h}:${m}`;
  }
  return String(v);
}

function escapeHtml(str){
  if(typeof str !== 'string') return '';
  return str
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");
}

const API_URL = "https://script.google.com/macros/s/AKfycbynNpk-XEfMXVqKfSXK5TIgFxTcd1I9LcX8BFhJW1sbN8j3DMIHwkaKh89L7a2uscQzeQ/exec";
async function saveZoomScheduling(data){
  return saveZoomAssignment(data);
}

function normalizeData(data){
  return data.map(r=>({
    ...r,
    StartTime: excelDecimalToTime(r.StartTime),
    EndTime: excelDecimalToTime(r.EndTime),
    End: (typeof r.End === 'number') ? excelSerialToJSDate(r.End) : (r.End ? new Date(r.End) : null),
    Dates: Array.isArray(r.Dates)
      ? r.Dates.map(d => excelSerialToJSDate(d)).filter(Boolean)
      : []
  }));
}

function sortByDateAndTime(list) {
  return [...list].sort((a, b) => {
    const aDate = a.date instanceof Date ? a.date : new Date(a.date);
    const bDate = b.date instanceof Date ? b.date : new Date(b.date);

    if (aDate.getTime() !== bDate.getTime()) {
      return aDate - bDate;
  }

    const [aH, aM] = String(a.start || '99:99').split(':').map(Number);
    const [bH, bM] = String(b.start || '99:99').split(':').map(Number);

    return (aH * 60 + aM) - (bH * 60 + bM);
  });
}

function getEarliestDate(dates) {
  const validDates = (dates || []).filter(Boolean);
  if(validDates.length === 0) return null;
  return new Date(Math.min(...validDates.map(d => d.getTime())));
}

function toDateAndTimeSortable(item, date, start) {
  return {
    ...item,
    date,
    start: start || '99:99'
  };
}

function createHourSelect(value){
  const select = document.createElement('select');
  select.className = 'zoom-time-select';

  const blank = document.createElement('option');
  blank.value = '';
  blank.textContent = '—';
  select.appendChild(blank);

  for(let h = 8; h <= 18; h++){
    const hour = String(h).padStart(2, '0') + ':00';
    const opt = document.createElement('option');
    opt.value = hour;
    opt.textContent = hour;
    if(hour === value) opt.selected = true;
    select.appendChild(opt);
  }

  return select;
}

const ZOOM_PROGRAMS = [
  'בינה מלאכותית',
  'ביומימיקרי',
  'ביומימיקרי לחטיבה',
  'השמיים אינם הגבול',
  'התנסות בתעשייה',
  'טכנולוגיות החלל',
  'יישומי AI',
  'מייקרים',
  'מנהיגות ירוקה',
  'משחקי קופסה',
  'פורצות דרך',
  'רוקחים עולם',
  'תלמידים להייטק',
];

rawData = normalizeData(rawData);

let schedulingJson = null;
let employmentTypeByEmployeeId = new Map();
let notesByKey = new Map();

function eventTypeOf(record){
  return String(record?.EventType || '').trim().toUpperCase();
}

function isEvent(record){
  return eventTypeOf(record) === 'EVENT';
}

async function loadSchedulingJson(){
  try{
    const res = await fetch('data/Scheduling/scheduling.json', { cache: 'no-store' });
    if(!res.ok) throw new Error('Failed to fetch scheduling.json: ' + res.status);
    schedulingJson = await res.json();

    const instructors = Array.isArray(schedulingJson?.instructors) ? schedulingJson.instructors : [];

    employmentTypeByEmployeeId = new Map(
      instructors.map(i => [
        String(i.EmployeeID ?? '').trim(),
        String(i.EmploymentType ?? '').trim()
      ])
    );
  }catch(err){
    console.error('loadSchedulingJson error:', err);
    schedulingJson = null;
    employmentTypeByEmployeeId = new Map();
  }
}

async function loadNotesJson(){
  try{
    const res = await fetch('data/notes/notes.json', { cache: 'no-store' });
    if(res.status === 404){
      notesByKey = new Map();
      return;
  }
    if(!res.ok) throw new Error('Failed to fetch notes.json: ' + res.status);

    const payload = await res.json();
    notesByKey = new Map(Object.entries(payload?.notesByKey || {}));
  }catch(err){
    console.warn('loadNotesJson warning:', err);
    notesByKey = new Map();
  }
}

function getEmploymentTypeForEmployeeId(employeeId){
  if(userRole === 'instructor' || window._dualViewMode === 'instructor') return '—';

  const id = String(employeeId ?? '').trim();
  if(!id) return '—';

  return employmentTypeByEmployeeId.get(id) || '—';
}
function enforceInstructorMode(){
  if(userRole === 'instructor' || window._dualViewMode === 'instructor'){
    window.mode = 'month';
  }
}

const view=document.getElementById('view');
const titleEl=document.getElementById('title');
const filtersEl=document.getElementById('filters');
const side=document.getElementById('side');
const sideContent=document.getElementById('sideContent');
const daySheet=document.getElementById('daySheet');
const daySheetBackdrop=document.getElementById('daySheetBackdrop');
const daySheetTitle=document.getElementById('daySheetTitle');
const daySheetContent=document.getElementById('daySheetContent');
const daySheetClose=document.getElementById('daySheetClose');
const btnMonth=document.getElementById('btnMonth');
const btnWeek=document.getElementById('btnWeek');
const btnSummary=document.getElementById('btnSummary');
const btnInstructors=document.getElementById('btnInstructors');
const btnEndDates=document.getElementById('btnEndDates');
const btnZoom=document.getElementById('btnZoom');
const goCalendar = document.getElementById('goCalendar');
const managerFilter=document.getElementById('managerFilter');
const employeeFilter=document.getElementById('employeeFilter');
const summaryMonth=document.getElementById('summaryMonth');
let activeSidePanelType = '';

function updateSchedulingButtonVisibility(){
  const btn = document.getElementById('btnScheduling');
  if(!btn) return;

  if(!window.ENABLE_SCHEDULING){
    btn.style.display = 'none';
    return;
  }

  // Show the scheduling button for all admins (not instructors)
  const isAdmin = (userRole === 'admin' || userRole === 'both') && window._dualViewMode !== 'instructor';
  if(isAdmin){
    btn.style.display = '';
  }else{
    btn.style.display = 'none';
  }
}

function updateEndDatesButtonVisibility(){
  if(!btnEndDates) return;
  const id = String(window.EmployeeID || '').trim();
  if(id === '7000' || id === '8000'){
    btnEndDates.style.display = '';
  }else{
    btnEndDates.style.display = 'none';
  }
}

function updateZoomButtonVisibility(){
  if(!btnZoom) return;
  if(userRole === 'instructor'){
    btnZoom.style.display = 'none';
  }else{
    btnZoom.style.display = '';
  }
}

if(userRole === 'instructor'){
  btnSummary.style.display = 'none';
  btnInstructors.style.display = 'none';
  btnMonth.style.display = 'none';
  btnWeek.style.display = 'none';
  filtersEl.style.display = 'none';
  window.mode = 'month';
}

// Dual role: manager + instructor — only for employee 1500
window._dualViewMode = 'admin'; // 'admin' | 'instructor'
function getDualRoleToggleLabel(){
  return window._dualViewMode === 'admin' ? 'תצוגת מדריך' : 'תצוגת מנהל';
}

function syncDualRoleToggleButtons(){
  const label = getDualRoleToggleLabel();
  document.querySelectorAll('.dual-role-toggle-btn').forEach((btn)=>{
    btn.textContent = label;
  });
}

function createDualRoleToggle(targetEl, extraClass=''){
  if(!targetEl) return;

  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = `dual-role-toggle-btn ${extraClass}`.trim();
  btn.onclick = switchDualRoleView;
  targetEl.appendChild(btn);
}

if(userRole === 'both'){
  // Start in admin view — all buttons visible by default
  // Add toggle button for desktop and mobile
  (function addDualRoleToggle(){
    const navModes = document.querySelector('.nav-modes');
    createDualRoleToggle(navModes, 'desktop-only');

    const navContainer = document.querySelector('.nav.week-header');
    createDualRoleToggle(navContainer, 'mobile-only');

    syncDualRoleToggleButtons();
  })();
}

function switchDualRoleView(){
  if(window._dualViewMode === 'admin'){
    // Switch to instructor view
    window._dualViewMode = 'instructor';
    document.body.dataset.role = 'instructor';
    rawData = normalizeData(window.personalData || []);
    btnSummary.style.display = 'none';
    btnInstructors.style.display = 'none';
    btnMonth.style.display = 'none';
    btnWeek.style.display = 'none';
    if(btnZoom) btnZoom.style.display = 'none';
    if(filtersEl) filtersEl.style.display = 'none';
    window.mode = 'month';
  } else {
    // Switch to admin view
    window._dualViewMode = 'admin';
    document.body.dataset.role = 'both';
    rawData = normalizeData(window.allAdminData || []);
    btnSummary.style.display = '';
    btnInstructors.style.display = '';
    btnMonth.style.display = '';
    btnWeek.style.display = '';
    if(btnZoom) btnZoom.style.display = '';
    if(filtersEl) filtersEl.style.display = '';
    window.mode = 'summary';
  }
  syncDualRoleToggleButtons();
  render();
}

let dataRange=null;
let _mode='month';
Object.defineProperty(window, 'mode', {
  get(){ return _mode; },
  set(value){
    if(userRole === 'instructor' || window._dualViewMode === 'instructor'){
      _mode = 'month';
      return;
    }
    _mode = value;
  },
  configurable: false
});
const isMobile = () => window.innerWidth < 800;

function logViewRuntimeState(context = 'render'){
  const currentViewMode = window._dualViewMode || userRole;
  console.log(`[${context}] viewMode:`, currentViewMode);
  console.log(`[${context}] screenWidth:`, window.innerWidth);
  console.log(`[${context}] isMobile:`, isMobile());
}

let openWeekId = null;
let currentDate=new Date(); currentDate.setHours(0,0,0,0);
const dayNames=['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];

const employeeColors = {
  "הנאא אבו אמזה": "#ffe1fb",
  "יונתן יהונתן פתייה": "#F7F5ED",
  "אביב בלנדר": "#FCF2FD",
  "אסיל ג'בר": "#ffd3ef",
  "ברקת קטעי": "#baffce",
  "אלכס זפקה": "#FFF5EC",
  "עליזה מולה": "#dccfff",
  "ליאל בן חמו": "#FFFFE9",
  "אפרת אוחיון": "#EFEAFF",
  "אלדר מיכאל טייב": "#E4F6FF",
  "הילה רוזן": "#ffd2f8",
  "תמר שפיר": "#F9EEEE",
  "אילנה טיטייבסקי": "#E6FFEB",
  "אמיר מלמוד": "#E4F6FF",
  "אוריה פדידה": "#DDFFFA",
  "אושרי רם": "#DDFFFA",
  "כרמית סמנדרוב": "#ffecda",
  "מיכל שכטמן": "#ffffd1",
  "ראנה סאלח": "#FCF2FD",
  "סוהא סאלם": "#ffefc5",
  "קרן גורביץ": "#F2FCFD",
  "אביגדור שרון": "#DDFFFA"
};
window.employeeColors = employeeColors;

function getEmployeeColor(name) {
  if (!name || name.trim() === "") return "#ffffff";
  return employeeColors[name.trim()] || "#f1f5f9";
}

function hashStringToHue(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) >>> 0;
  return h % 360;
}

function ensureColorContrast(colorValue) {
  const tuple = toRgbTuple(colorValue).split(',').map(v => Number(v.trim()));
  if (tuple.length !== 3 || tuple.some(v => Number.isNaN(v))) return colorValue;

  let [r, g, b] = tuple;
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  if (luminance <= 0.82) return colorValue;

  const darkenFactor = luminance > 0.9 ? 0.72 : 0.84;
  r = Math.max(0, Math.round(r * darkenFactor));
  g = Math.max(0, Math.round(g * darkenFactor));
  b = Math.max(0, Math.round(b * darkenFactor));
  return `rgb(${r}, ${g}, ${b})`;
}

function hasStrongSaturation(colorValue) {
  const tuple = toRgbTuple(colorValue).split(',').map(v => Number(v.trim()));
  if (tuple.length !== 3 || tuple.some(v => Number.isNaN(v))) return false;
  const [r, g, b] = tuple;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  if (max === 0) return false;
  const saturation = (max - min) / max;
  return saturation >= 0.2;
}

const instructorColors = [
  '#b2ebf2',
  '#f8bbd0',
  '#fff9c4',
  '#ffe0b2',
  '#c8e6c9',
  '#bbdefb',
  '#e0f7fa',
  '#fce4ec',
  '#fff3e0',
  '#f3e5f5',
  '#e3f2fd',
  '#ffebee'
];

const instructorColorMap = {};

function getInstructorPaletteColor(name) {
  const instructor = String(name || '').trim();
  if (!instructor) return instructorColors[0];

  if (!instructorColorMap[instructor]) {
    const index = Object.keys(instructorColorMap).length % instructorColors.length;
    instructorColorMap[instructor] = instructorColors[index];
  }

  return instructorColorMap[instructor];
}

function instructorColor(name) {
  return getInstructorPaletteColor(name);
}

function toRgbTuple(colorValue) {
  if (!colorValue) return '148, 163, 184';
  const normalized = String(colorValue).trim();
  const hexMatch = normalized.match(/^#([\da-f]{3}|[\da-f]{6})$/i);
  if (hexMatch) {
    const hex = hexMatch[1];
    const fullHex = hex.length === 3 ? hex.split('').map(ch => ch + ch).join('') : hex;
    const r = parseInt(fullHex.slice(0, 2), 16);
    const g = parseInt(fullHex.slice(2, 4), 16);
    const b = parseInt(fullHex.slice(4, 6), 16);
    return `${r}, ${g}, ${b}`;
  }

  const rgbMatch = normalized.match(/^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)(?:\s*,\s*(\d*\.?\d+))?\s*\)$/i);
  if (rgbMatch) {
    const clamp = (v) => Math.max(0, Math.min(255, Math.round(Number(v))));
    return `${clamp(rgbMatch[1])}, ${clamp(rgbMatch[2])}, ${clamp(rgbMatch[3])}`;
  }

  const hslMatch = normalized.match(/^hsl\((\d+)\s+(\d+)%\s+(\d+)%\)$/i);
  if (hslMatch) {
    const h = Number(hslMatch[1]) / 360;
    const s = Number(hslMatch[2]) / 100;
    const l = Number(hslMatch[3]) / 100;
    const hue2rgb = (p, q, t) => {
      let v = t;
      if (v < 0) v += 1;
      if (v > 1) v -= 1;
      if (v < 1 / 6) return p + (q - p) * 6 * v;
      if (v < 1 / 2) return q;
      if (v < 2 / 3) return p + (q - p) * (2 / 3 - v) * 6;
      return p;
  };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    const r = Math.round(hue2rgb(p, q, h + 1 / 3) * 255);
    const g = Math.round(hue2rgb(p, q, h) * 255);
    const b = Math.round(hue2rgb(p, q, h - 1 / 3) * 255);
    return `${r}, ${g}, ${b}`;
  }

  return '148, 163, 184';
}

function applyInstructorColorVars(node, colorValue) {
  node.style.setProperty('--instructor-color', colorValue);
  node.style.setProperty('--instructor-color-rgb', toRgbTuple(colorValue));
}

function formatTime(v){
  if(v==null||v==='') return '';
  if(typeof v==='number'){
    const totalMinutes = Math.round(v * 24 * 60);
    const h = String(Math.floor(totalMinutes/60)).padStart(2,'0');
    const m = String(totalMinutes%60).padStart(2,'0');
    return `${h}:${m}`;
  }
  return String(v);
}

const sameDay=(a,b)=>a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate();

function parseDate(v){
  if(!v) return null;
  const d=new Date(v);
  return isNaN(d)?null:d;
}

function isCourse(r){
  return eventTypeOf(r) === 'COURSE';
}

function getSessionNumberForItem(item){
  if(Number.isFinite(item.meetingIdx)) return item.meetingIdx;

  const dates = Array.isArray(item.Dates) ? item.Dates : [];
  const selectedDate = item.selectedDate instanceof Date ? item.selectedDate : null;
  if(!selectedDate) return null;

  const idx = dates.findIndex(d => sameDay(d, selectedDate));
  return idx >= 0 ? idx + 1 : null;
}

function getNotesForCourseItem(item){
  if(userRole !== 'instructor' || !isCourse(item)) return null;

  const program = String(item.Program || '').trim();
  const sessionNumber = getSessionNumberForItem(item);
  if(!program || !sessionNumber) return null;

  const noteData = notesByKey.get(`${program}|${sessionNumber}`);
  if(!noteData || typeof noteData !== 'object') return null;

  const toArray = (v) => Array.isArray(v) ? v.filter(Boolean) : [];
  // The Rem of session X is the note (הודעה) for session X
  const message = toArray(noteData.reminder);
  const general = toArray(noteData.general);

  // "תזכורת לשיעור הבא" = the Rem of the next session (if exists)
  const nextNoteData = notesByKey.get(`${program}|${sessionNumber + 1}`);
  const reminder = nextNoteData ? toArray(nextNoteData.reminder) : [];

  if(!message.length && !reminder.length && !general.length) return null;

  return { message, reminder, general };
}

function applyNotesBoxColor(){
  const boxes = document.querySelectorAll('.notes-box');
  if(!boxes.length) return;

  const name =
    (window.currentUserName ||
     window.currentUser?.name ||
     window.currentUser?.Employee ||
     '').trim();

  const baseColor =
    (window.employeeColors && name && window.employeeColors[name])
      ? window.employeeColors[name]
      : '#f8fafc';

  boxes.forEach(box => {
    box.style.backgroundColor = baseColor;
  });
}

function renderNotesBlock(notes, employeeName){
  if(!notes) return '';

  const sections = [
    { title: 'הודעה', items: notes.message },
    { title: 'תזכורת לשיעור הבא', items: notes.reminder },
    { title: 'מידע כללי', items: notes.general }
  ].filter(section => section.items.length > 0)
    .map(section => `
      <div class="notes-section">
        <div class="note-section-title">${section.title}</div>
        <ul>
          ${section.items.map(line => `<li>${line}</li>`).join('')}
        </ul>
    </div>
    `).join('');

  if(!sections) return '';

  return `
    <div class="notes-box" id="notesBox">
      <div class="notes-title">פתקים</div>
      <div class="notes-content">
        ${sections}
    </div>
    </div>
  `;
}


function getCourseManager(r){
  return String(r.CourseManager ?? '').trim();
}

function getInstructorManager(r){
  return String(r.InstructorManager ?? '').trim();
}

function getManagerForCourseViews(r){
  return (userRole === 'instructor' || window._dualViewMode === 'instructor') ? getInstructorManager(r) : getCourseManager(r);
}


function isEventVisibleToCurrentUser(record){
  if(!isEvent(record)) return true;
  if(userRole !== 'instructor' && window._dualViewMode !== 'instructor') return true;

  const eventEmployeeId = String(record.EmployeeID || '').trim();
  const currentEmployeeId = String(window.EmployeeID || '').trim();
  return !!eventEmployeeId && eventEmployeeId === currentEmployeeId;
}

function isCourseActiveInMonth(r, year, month){
  return isCourse(r) &&
    r.Dates.some(d =>
      d &&
      d.getFullYear() === year &&
      d.getMonth() === month
    );
}

function isCourseEndingInMonth(r, year, month){
  if(!isCourse(r) || !r.End) return false;
  const d = parseDate(r.End);
  return d &&
         d.getFullYear() === year &&
         d.getMonth() === month;
}

function getBusiestWeekWorkDays(courses, year, month){
  const weeksMap = {};

  courses.forEach(r=>{
    if(String(r.EventType || '').trim().toUpperCase() !== 'COURSE') return;

    r.Dates.forEach(d=>{
      if(
        d &&
        d.getFullYear() === year &&
        d.getMonth() === month
      ){
        const weekStart = new Date(d);
        weekStart.setDate(d.getDate() - d.getDay());
        weekStart.setHours(0,0,0,0);

        const key = weekStart.toISOString();

        if(!weeksMap[key]){
          weeksMap[key] = new Set();
      }

        weeksMap[key].add(d.toDateString());
    }
  });
  });

  let maxDays = 0;

  Object.values(weeksMap).forEach(set=>{
    if(set.size > maxDays){
      maxDays = set.size;
  }
  });

  return maxDays;
}

function clampDateToDataRange(date){
  const d = new Date(date);
  d.setHours(0,0,0,0);
  if(!dataRange) return d;
  if(d < dataRange.min) return new Date(dataRange.min);
  if(d > dataRange.max) return new Date(dataRange.max);
  return d;
}

function weekOverlapsDataRange(date){
  if(!dataRange) return true;
  const weekStart = new Date(date);
  weekStart.setDate(date.getDate() - date.getDay());
  weekStart.setHours(0,0,0,0);

  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  weekEnd.setHours(0,0,0,0);

  return weekEnd >= dataRange.min && weekStart <= dataRange.max;
}

function getMinAllowedMonth(){
  const today = new Date();
  today.setHours(0,0,0,0);
  return new Date(today.getFullYear(), today.getMonth()-1, 1);
}

function canGoPrev(){
  if(window.mode === 'summary'){
    if(summaryMonth.selectedIndex <= 0) return false;

    const prevOption = summaryMonth.options[summaryMonth.selectedIndex-1].value;
    const [y,m] = prevOption.split('-').map(Number);
    const prevDate = new Date(y,m,1);

    return prevDate >= getMinAllowedMonth();
  }
  if(!dataRange) return false;

  if(window.mode === 'month'){
    if((userRole === 'instructor' || window._dualViewMode === 'instructor') && isMobile()){
      // מדריך במובייל – ניווט שבועי
      const temp = new Date(currentDate);
      temp.setDate(temp.getDate() - 7);
      return temp >= getMinAllowedMonth();
    }
    const temp = new Date(currentDate);
    temp.setMonth(temp.getMonth()-1);

    return temp >= getMinAllowedMonth();
  }

  if(window.mode === 'week'){
    const temp = new Date(currentDate);
    temp.setDate(temp.getDate()-7);

    const minAllowed = getMinAllowedMonth();

    return temp >= minAllowed;
  }

  return false;
}

function canGoNext(){
  if(window.mode === 'summary') return summaryMonth.selectedIndex < summaryMonth.options.length-1;
  if(!dataRange) return false;

  if(window.mode === 'month'){
    if((userRole === 'instructor' || window._dualViewMode === 'instructor') && isMobile()){
      // מדריך במובייל – ניווט שבועי
      const temp = new Date(currentDate);
      temp.setDate(temp.getDate() + 7);
      return weekOverlapsDataRange(temp);
    }
    const temp = new Date(currentDate);
    temp.setMonth(temp.getMonth()+1);
    return temp.getFullYear() < dataRange.max.getFullYear() ||
      (temp.getFullYear() === dataRange.max.getFullYear() && temp.getMonth() <= dataRange.max.getMonth());
  }

  if(window.mode === 'week'){
    const temp = new Date(currentDate);
    temp.setDate(temp.getDate()+7);
    return weekOverlapsDataRange(temp);
  }

  return false;
}

function updateNavButtons(){
  const prevBtn = document.getElementById('prev');
  const nextBtn = document.getElementById('next');
  prevBtn.disabled = !canGoPrev();
  nextBtn.disabled = !canGoNext();
}

function updateModeButtons(){
  btnMonth.classList.remove('active');
  btnWeek.classList.remove('active');
  btnSummary.classList.remove('active');
  btnInstructors.classList.remove('active');
  if(btnEndDates) btnEndDates.classList.remove('active');
  if(btnZoom) btnZoom.classList.remove('active');

  if(window.mode === 'month') btnMonth.classList.add('active');
  if(window.mode === 'week') btnWeek.classList.add('active');
  if(window.mode === 'summary') btnSummary.classList.add('active');
  if(window.mode === 'instructors') btnInstructors.classList.add('active');
  if(window.mode === 'enddates' && btnEndDates) btnEndDates.classList.add('active');
  if(window.mode === 'zoom' && btnZoom) btnZoom.classList.add('active');
}

function endDate(r){
  const d=r.Dates.filter(Boolean);
  return d.length?new Date(Math.max(...d.map(x=>x.getTime()))):null;
}

function getDataDateRange(){
  const allDates = rawData.flatMap(r => r.Dates.filter(Boolean));

  if(allDates.length === 0) return null;

  const min = new Date(Math.min(...allDates.map(d=>d.getTime())));
  const max = new Date(Math.max(...allDates.map(d=>d.getTime())));

  min.setHours(0,0,0,0);
  max.setHours(0,0,0,0);

  return { min, max };
}

function updatePageUserName(user){
  const header = document.getElementById('pageEmployeeHeader') || document.getElementById('greetingName');
  if(!header) return;

  if(!user){
    header.textContent = '';
    return;
  }

  if(user.Role === 'instructor'){
    header.textContent = user.Employee || '';
    return;
  }

  if(user.Role === 'manager' || user.Role === 'admin'){
    header.textContent = user.Employee || user.Name || user.Manager || '';
    return;
  }

  header.textContent = user.Employee || user.Manager || user.Name || '';
}

async function initFromRawData(){
  dataRange = getDataDateRange();
  window.dataRange = dataRange;
  initFilters();
  initSummaryMonths();
  currentDate = clampDateToDataRange(new Date());

  const sessionName = sessionStorage.getItem('dash_name') || '';
  const currentUser = rawData.find(r => String(r.EmployeeID || '').trim() === String(window.EmployeeID || '').trim());
  const currentUserForHeader = {
    ...(currentUser || {}),
    Role: userRole,
    Name: sessionName || currentUser?.Name || '',
    Manager: currentUser?.Manager || sessionName || ''
  };
  window.currentUser = currentUserForHeader;
  window.currentUserName = currentUserForHeader.Employee || currentUserForHeader.Name || '';
  window.currentUserEmployeeID = String(window.EmployeeID || '').trim();
  updatePageUserName(currentUserForHeader);

  updateSchedulingButtonVisibility();
  updateEndDatesButtonVisibility();
  updateZoomButtonVisibility();

  window.mode = (userRole === 'instructor') ? 'month' : 'summary';

  if(userRole === 'instructor'){
    await loadSchedulingJson();
    await loadNotesJson();
    if(schedulingJson && Array.isArray(schedulingJson.courses)){
      const holidays = schedulingJson.courses.filter(
        r => eventTypeOf(r) === 'HOLIDAY'
      );
      if(holidays.length){
        rawData = rawData.concat(normalizeData(holidays));
    }
  }

    if(schedulingJson){
      const currentEmployeeId = String(window.EmployeeID || '').trim();
      const eventsFromCourses = Array.isArray(schedulingJson.courses)
        ? schedulingJson.courses.filter(isEvent)
        : [];
      const eventsLegacy = Array.isArray(schedulingJson.events) ? schedulingJson.events : [];
      const visibleEvents = eventsFromCourses
        .concat(eventsLegacy)
        .filter(e => String(e.EmployeeID || '').trim() === currentEmployeeId);

      if(visibleEvents.length){
        const existingEventKeys = new Set(
          rawData
            .filter(isEvent)
            .map(r => `${String(r.EmployeeID || '').trim()}|${String(r.Program || '').trim()}|${String(r.Date1 || '').trim()}|${String(r.StartTime || '').trim()}|${String(r.EndTime || '').trim()}`)
        );

        const missingEvents = visibleEvents.filter(e => {
          const key = `${String(e.EmployeeID || '').trim()}|${String(e.Program || '').trim()}|${String(e.Date1 || '').trim()}|${String(e.StartTime || '').trim()}|${String(e.EndTime || '').trim()}`;
          return !existingEventKeys.has(key);
      });

        if(missingEvents.length){
          rawData = rawData.concat(normalizeData(missingEvents));
      }
    }
  }
  } else {
    await loadSchedulingJson();
    notesByKey = new Map();
    if(schedulingJson){
      const eventsFromCourses = Array.isArray(schedulingJson.courses)
        ? schedulingJson.courses.filter(isEvent)
        : [];
      const eventsLegacy = Array.isArray(schedulingJson.events) ? schedulingJson.events : [];
      const allEvents = eventsFromCourses.concat(eventsLegacy);
      if(allEvents.length){
        const existingEventKeys = new Set(
          rawData
            .filter(isEvent)
            .map(r => `${String(r.EmployeeID || '').trim()}|${String(r.Program || '').trim()}|${String(r.Date1 || '').trim()}|${String(r.StartTime || '').trim()}|${String(r.EndTime || '').trim()}`)
        );

        const missingEvents = allEvents.filter(e => {
          const key = `${String(e.EmployeeID || '').trim()}|${String(e.Program || '').trim()}|${String(e.Date1 || '').trim()}|${String(e.StartTime || '').trim()}|${String(e.EndTime || '').trim()}`;
          return !existingEventKeys.has(key);
      });

        if(missingEvents.length){
          rawData = rawData.concat(normalizeData(missingEvents));
      }
    }
  }
  }

  render();
  applyNotesBoxColor();
}

function initFilters(){
  const managers=[...new Set(rawData.map(r=>getCourseManager(r)).filter(Boolean))];
  const employees=[...new Set(rawData.map(r=>r.Employee).filter(Boolean))];
  managerFilter.innerHTML='<option value="">כל המנהלים</option>'+managers.map(v=>`<option>${v}</option>`).join('');
  employeeFilter.innerHTML='<option value="">כל המדריכים</option>'+employees.map(v=>`<option>${v}</option>`).join('');
}

function initSummaryMonths(){
  const months=[...new Set(rawData.flatMap(r=>r.Dates.filter(Boolean)).map(d=>`${d.getFullYear()}-${d.getMonth()}`))].sort();
  summaryMonth.innerHTML=months.map(k=>{
    const [y,m]=k.split('-').map(Number);
    return `<option value="${k}">${new Date(y,m).toLocaleString('he-IL',{month:'long',year:'numeric'})}</option>`;
  }).join('');
  const todayKey = `${currentDate.getFullYear()}-${currentDate.getMonth()}`;
  const idx = months.indexOf(todayKey);
  if(idx >= 0) summaryMonth.selectedIndex = idx;
}

function fitViewToScreen() {
  if (window.innerWidth <= 800) return;
  const view = document.getElementById('view');
  if (!view) return;
  Array.from(view.children).forEach(c => { c.style.zoom = ''; });
  if(window.mode === 'enddates') return;
  requestAnimationFrame(() => requestAnimationFrame(() => {
    const viewH = view.clientHeight;
    const contentH = view.scrollHeight;
    if (!viewH || !contentH) return;
    const z = Math.max(0.70, Math.min(0.88, viewH / contentH));
    Array.from(view.children).forEach(c => { c.style.zoom = z; });
  }));
}
window.addEventListener('resize', fitViewToScreen);

async function render(){

  if(window._lastRenderedMode === 'zoom' && window.mode !== 'zoom'){
    clearZoomCache();
  }

  logViewRuntimeState('render');

  enforceInstructorMode();
  view.innerHTML=''; view.style.display=''; view.style.flexDirection=''; view.style.alignItems=''; view.style.justifyContent=''; view.style.width=''; view.scrollTop=0; window.scrollTo(0,0); document.documentElement.scrollTop=0; document.body.scrollTop=0; view.classList.toggle('week-mode', window.mode === 'week');
  view.classList.toggle('view-week', window.mode === 'week');
  view.classList.toggle('view-month', window.mode === 'month');
  view.classList.toggle('view-instructors', window.mode === 'instructors');
  view.classList.toggle('view-summary', window.mode === 'summary');
  view.classList.toggle('view-managers', window.mode === 'instructors');
  view.classList.toggle('view-enddates', window.mode === 'enddates');
  view.classList.toggle('view-zoom', window.mode === 'zoom');
  closeSidePanel();

  if(userRole === 'instructor' || window._dualViewMode === 'instructor' || window.mode === 'summary' || window.mode === 'instructors' || window.mode === 'enddates' || window.mode === 'zoom' || isMobile()){
    filtersEl.style.display = 'none';
  }else{
    filtersEl.style.display = 'flex';
  }

  if(window.mode==='summary'){
    renderSummary();
  }
  else if(window.mode==='week'){
    renderWeekView();
  }
  else if(window.mode==='instructors'){
    renderInstructors();
  }
  else if(window.mode==='enddates'){
    renderEndDates();
  }
  else if(window.mode==='zoom'){
    await renderZoom();
  }
  else{
    renderMonthView();
  }

  updateNavButtons();
  updateModeButtons();

  if(window.mode === 'month' || window.mode === 'enddates' || window.mode === 'zoom'){
    goCalendar.style.display = 'none';
  } else {
    goCalendar.style.display = 'inline-flex';
  }

  applyNotesBoxColor();
  fitViewToScreen();
  window._lastRenderedMode = window.mode;
}

function renderMonthView(){
  if(userRole === 'instructor' || window._dualViewMode === 'instructor'){
    if(isMobile()){
      renderInstructorMobileWeek();
    } else {
      renderInstructorGridMonth();
    }
    return;
  }
  if(isMobile()){
    renderMobileMonth();
    return;
  }
  renderDesktopMonth();
}

function renderInstructorMobileWeek(){
  renderMobileMonthAccordion(applyFilters());
}

function renderInstructorGridMonth(){
  const y = currentDate.getFullYear();
  const m = currentDate.getMonth();
  const monthTitle = new Date(y,m,1).toLocaleString('he-IL',{month:'long',year:'numeric'});
  titleEl.textContent = monthTitle;

  const data = applyFilters();
  const today = new Date(); today.setHours(0,0,0,0);

  const wrap = document.createElement('div');
  wrap.className = 'instructor-cal-wrap';

  const grid = document.createElement('div');
  grid.className = 'instructor-cal-grid';

  const first = new Date(y,m,1);
  const last  = new Date(y,m+1,0);

  // שורת כותרת ימים
  ['א׳','ב׳','ג׳','ד׳','ה׳','ו׳','ש׳'].forEach((d, i) => {
    const h = document.createElement('div');
    h.className = 'instructor-cal-header-cell' + (i === 6 ? ' ic-shabbat-header' : '');
    h.textContent = d;
    grid.appendChild(h);
  });

  // תאים ריקים לפני תחילת החודש
  for(let i = 0; i < first.getDay(); i++){
    const empty = document.createElement('div');
    empty.className = 'instructor-cal-cell ic-empty';
    grid.appendChild(empty);
  }

  for(let d = 1; d <= last.getDate(); d++){
    const date = new Date(y,m,d);
    const isToday = sameDay(date, today);

    // איסוף כל האירועים של היום
    const dailyItems = [];
    data.forEach(r => r.Dates.forEach((dd, idx) => {
      if(sameDay(dd, date)) dailyItems.push({ ...r, meetingIdx: idx+1, selectedDate: dd });
  }));

    // קיבוץ לפי תוכנית (זהה ל-buildDay)
    const groupsMap = {};
    dailyItems.forEach(ev => {
      if(ev.EventType === 'HOLIDAY'){
        const key = `holiday-${ev.Program}`;
        if(!groupsMap[key]) groupsMap[key] = { type:'holiday', items:[ev] };
    } else if(isEvent(ev)){
        const key = `event-${ev.Employee}-${ev.Program}`;
        if(!groupsMap[key]) groupsMap[key] = { type:'event', time: ev.StartTime||'99:99', items:[] };
        groupsMap[key].items.push(ev);
    } else {
        const key = `${ev.Employee}-${ev.Program}`;
        if(!groupsMap[key]) groupsMap[key] = { type:'course', time: ev.StartTime||'99:99', items:[] };
        groupsMap[key].items.push(ev);
    }
  });
    const groups = Object.values(groupsMap)
      .sort((a,b) => (a.time||'').localeCompare(b.time||''));
    const activityGroups = groups.filter(g => g.type !== 'holiday');
    const hasActivity = activityGroups.length > 0;

    const isShabbat = date.getDay() === 6;
    const cell = document.createElement('div');
    cell.className = 'instructor-cal-cell calendar-day' +
      (isShabbat ? ' ic-shabbat' : '') +
      (date.getDay() === 5 ? ' ic-friday' : '') +
      (isToday ? ' ic-today is-today' : '');

    // תצוגת יום: יום בשבוע מקוצר + יום/חודש
    const numWrap = document.createElement('div');
    numWrap.className = 'instructor-cal-day-num day-number date-number' + (isToday ? ' ic-today-num' : '');
    const day = date.getDate();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const weekdays = ['א׳','ב׳','ג׳','ד׳','ה׳','ו׳','ש׳'];
    const weekday = weekdays[date.getDay()];
    numWrap.textContent = `${weekday} ${day}/${month}`;

    if(hasActivity && !isToday && !isShabbat){
      const firstActivity = activityGroups[0]?.items?.[0] || null;
      const instructorName = firstActivity?.Employee || firstActivity?.Instructor || firstActivity?.EmployeeName || '';
      const mappedEmployeeColor = getEmployeeColor(instructorName);
      const activityColor = (instructorName && mappedEmployeeColor !== '#f1f5f9' && mappedEmployeeColor !== '#ffffff')
        ? instructorColor(instructorName)
        : getProgramColor(firstActivity?.Program || '');

      numWrap.classList.add('has-day-activity');
      numWrap.style.setProperty('--activity-color', activityColor);
    }

    cell.appendChild(numWrap);

    // פילים של אירועים (מקסימום 3) – לא בשבת
    const maxPills = 3;
    if(!isShabbat) groups.slice(0, maxPills).forEach(g => {
      const firstItem = g.items[0];
      const pill = document.createElement('div');
      pill.className = 'instructor-cal-pill calendar-event';
      const instructorName = firstItem.Employee || firstItem.Instructor || firstItem.EmployeeName || '';
      const eventColor = g.type === 'holiday' ? '#cbd5e1' : instructorColor(instructorName);
      applyInstructorColorVars(pill, eventColor);
      pill.style.backgroundColor = eventColor;
      if(g.type === 'holiday') pill.classList.add('holiday');
      if(g.type === 'holiday') pill.addEventListener('click', e => e.stopPropagation());
      const txt = firstItem.Program || '';
      const text = document.createElement('span');
      text.className = 'instructor-cal-pill-text';
      text.textContent = txt.length > 13 ? txt.slice(0,12)+'…' : txt;
      pill.appendChild(text);
      cell.appendChild(pill);
  });

    if(!isShabbat && groups.length > maxPills){
      const more = document.createElement('div');
      more.className = 'instructor-cal-more';
      more.textContent = `+${groups.length - maxPills}`;
      cell.appendChild(more);
  }

    // לחיצה → פתח פאנל צד (לא לחגים)
    const nonHolidayItems = dailyItems.filter(item => String(item.EventType || '').trim().toUpperCase() !== 'HOLIDAY');
    if(nonHolidayItems.length > 0){
      cell.addEventListener('click', (e) => {
        e.stopPropagation();
        openSideGrouped(nonHolidayItems);
    });
  }

    grid.appendChild(cell);
  }

  wrap.appendChild(grid);
  view.appendChild(wrap);

  if(window.currentUserRole === 'instructor' || window._dualViewMode === 'instructor'){
    const currentYear = y;
    const currentMonth = m;
    const currentEmployeeID = String(window.currentUserEmployeeID || window.EmployeeID || '').trim();

    const activeCourses = rawData.filter(r =>
      isCourse(r) &&
      String(r.EmployeeID || '').trim() === currentEmployeeID &&
      isCourseActiveByRange(r, currentYear, currentMonth)
    ).length;

    const dailyActivitiesCount = rawData.filter(r => {
      const type = String(r.EventType || '').trim().toUpperCase();
      const isDaily = type === 'WORKSHOP' || type === 'TOUR';
      const isOwn = String(r.EmployeeID || '').trim() === currentEmployeeID;
      const date1 = parseDate(r.Date1);
      const inMonthByDate1 = date1 && date1.getFullYear() === currentYear && date1.getMonth() === currentMonth;

      return isDaily && isOwn && inMonthByDate1;
  }).length;

    const distinctDays = new Set(
      rawData
        .filter(r =>
          r.EmployeeID == currentEmployeeID &&
          r.Dates?.some(d =>
            d.getFullYear() === currentYear &&
            d.getMonth() === currentMonth
          )
        )
        .flatMap(r =>
          r.Dates
            .filter(d =>
              d.getFullYear() === currentYear &&
              d.getMonth() === currentMonth
            )
            .map(d => d.toDateString())
        )
    ).size;

    const personalSummary = document.createElement('div');
    personalSummary.className = 'personal-summary-row';

    personalSummary.innerHTML = `
      <div class="kpi-card">
        <div class="kpi-label summary-label">קורסים פעילים</div>
        <div class="kpi-value summary-number">${activeCourses}</div>
    </div>
      <div class="kpi-card">
        <div class="kpi-label summary-label">סדנאות/סיורים</div>
        <div class="kpi-value summary-number">${dailyActivitiesCount}</div>
    </div>
      <div class="kpi-card">
        <div class="kpi-label summary-label">ימי פעילות</div>
        <div class="kpi-value summary-number">${distinctDays}</div>
    </div>
  `;

    view.appendChild(personalSummary);
  }
}

function renderWeekView(){
  if (isMobile()) {
    renderMobileMonthAccordion();
  } else {
    renderDesktopWeekView();
  }
}

function renderDesktopMonth(){
  titleEl.textContent=currentDate.toLocaleString('he-IL',{month:'long',year:'numeric'});
  const data=applyFilters();
  const grid=document.createElement('div'); grid.className='grid';
  const y=currentDate.getFullYear(),m=currentDate.getMonth();
  const first=new Date(y,m,1),last=new Date(y,m+1,0);
  for(let i=0;i<first.getDay();i++) grid.appendChild(Object.assign(document.createElement('div'), {className:'day inactive'}));
  for(let d=1;d<=last.getDate();d++){ grid.appendChild(buildDay(new Date(y,m,d),data)); }
  view.appendChild(grid);
}

function renderDesktopWeekView(){
  const s=new Date(currentDate); s.setDate(s.getDate()-s.getDay());
  const e=new Date(s); e.setDate(e.getDate()+6);
  titleEl.textContent=`${s.toLocaleDateString('he-IL')} – ${e.toLocaleDateString('he-IL')}`;
  const data=applyFilters();
  const grid=document.createElement('div'); grid.className='grid';
  for(let i=0;i<7;i++){
    const cur=new Date(s); cur.setDate(s.getDate()+i);
    grid.appendChild(buildDay(cur,data));
  }
  view.appendChild(grid);
}

function renderMobileWeekView(){
  const weekStart = new Date(currentDate);
  weekStart.setDate(currentDate.getDate() - currentDate.getDay());
  weekStart.setHours(0,0,0,0);

  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  titleEl.textContent = `${weekStart.toLocaleDateString('he-IL')} – ${weekEnd.toLocaleDateString('he-IL')}`;

  const data = applyFilters();
  const wrapper = document.createElement('div');
  wrapper.className = 'mobile-narrow-wrap';

  const container = document.createElement('div');
  container.style.cssText = 'display:flex;flex-direction:column;gap:12px;padding:10px 10px 200px;';

  for(let i = 0; i < 7; i++){
    const date = new Date(weekStart);
    date.setDate(weekStart.getDate() + i);
    container.appendChild(buildDay(date, data));
  }

  wrapper.appendChild(container);
  view.appendChild(wrapper);

  requestAnimationFrame(() => {
    const todayEl = container.querySelector('.today');
    if(todayEl){
      const viewEl = document.getElementById('view');
      const rect = todayEl.getBoundingClientRect();
      const viewRect = viewEl.getBoundingClientRect();
      viewEl.scrollTo({
        top: viewEl.scrollTop + rect.top - viewRect.top - 10,
        behavior: 'auto'
    });
  }
  });
}

function initMobileAccordion(){

  if (window.innerWidth > 768) return;

  const weeks = document.querySelectorAll('.mobile-week');

  weeks.forEach(week => {

    const header = week.querySelector('.mobile-week-header');
    if (!header) return;

    header.addEventListener('click', () => {

      const isOpen = week.classList.contains('open');

      weeks.forEach(w => w.classList.remove('open'));

      if (!isOpen) {
        week.classList.add('open');
    }

  });

  });
}


function renderMobileMonthAccordion(data){
  const y = currentDate.getFullYear();
  const m = currentDate.getMonth();
  titleEl.textContent = new Date(y,m,1).toLocaleString('he-IL',{month:'long',year:'numeric'});

  if(!data) data = applyFilters();
  const first = new Date(y,m,1);
  const last  = new Date(y,m+1,0);
  const start = new Date(first);
  start.setHours(0,0,0,0);

  const today = new Date(); today.setHours(0,0,0,0);

  const wrapper = document.createElement('div');
  wrapper.className = 'mobile-narrow-wrap';

  const container = document.createElement('div');
  container.className = 'mobile-accordion';

  const sideNav = document.createElement('div');
  sideNav.className = 'week-side-nav';

  const setActiveWeek = (weekKey) => {
    sideNav.querySelectorAll('button').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.weekKey === weekKey);
    });
  };

  let todayWeekEl = null;
  let cursor = new Date(start);
  let weekNumber = 1;

  while(cursor <= last){
    const weekStart = new Date(cursor);
    const weekEnd   = new Date(cursor);
    weekEnd.setDate(weekStart.getDate() + 6);
    if(weekEnd > last){
      weekEnd.setTime(last.getTime());
    }

    const containsToday = today >= weekStart && today <= weekEnd;

    const weekEl = document.createElement('div');
    weekEl.className = 'accordion-week' + (containsToday ? ' accordion-today' : '');

    const header = document.createElement('div');
    header.className = 'accordion-header';
    header.innerHTML = `
      <div>
        <div class="accordion-header-text">${weekStart.toLocaleDateString('he-IL')} – ${weekEnd.toLocaleDateString('he-IL')}</div>
      </div>
      <span class="accordion-arrow">▼</span>
    `;

    const content = document.createElement('div');
    content.className = 'accordion-content';

    const weekKey = weekStart.toISOString();

    const openWeek = () => {
      const isOpen = weekEl.classList.contains('open');
      container.querySelectorAll('.accordion-week.open').forEach(w => w.classList.remove('open'));
      if(!isOpen){
        weekEl.classList.add('open');
        setActiveWeek(weekKey);
        if(!content.dataset.loaded){
          for(let i = 0; i < 7; i++){
            const date = new Date(weekStart);
            date.setDate(weekStart.getDate() + i);
            if(date > weekEnd) break;
            content.appendChild(buildDay(date, data));
          }
          content.dataset.loaded = 'true';
        }
        setTimeout(() => {
          const viewEl = document.getElementById('view');
          const rect = weekEl.getBoundingClientRect();
          const viewRect = viewEl.getBoundingClientRect();
          viewEl.scrollTo({ top: viewEl.scrollTop + rect.top - viewRect.top - 10, behavior: 'smooth' });
        }, 50);
      } else {
        setActiveWeek('');
      }
    };

    header.addEventListener('click', openWeek);

    const navBtn = document.createElement('button');
    navBtn.type = 'button';
    navBtn.dataset.weekKey = weekKey;
    navBtn.textContent = String(weekNumber);
    navBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      openWeek();
    });
    sideNav.appendChild(navBtn);

    weekEl.appendChild(header);
    weekEl.appendChild(content);

    container.appendChild(weekEl);
    cursor.setDate(cursor.getDate() + 7);
    weekNumber += 1;
  }

  wrapper.appendChild(sideNav);
  wrapper.appendChild(container);
  view.appendChild(wrapper);
}

function renderMobileMonth(){
  renderMobileMonthAccordion();
}

function openMobileWeekDetail(weekStart, data){
  renderMobileMonthAccordion(data);
}

function toggleWeek(weekId) {
  if (!isMobile()) return;

  const clicked = document.getElementById(`week-${weekId}`);
  if (!clicked) return;

  if (openWeekId === weekId) {
    clicked.classList.remove('open');
    openWeekId = null;
    return;
  }

  if (openWeekId !== null) {
    const previous = document.getElementById(`week-${openWeekId}`);
    if (previous) previous.classList.remove('open');
  }

  clicked.classList.add('open');
  openWeekId = weekId;
}

function buildDay(date,data){
  const cell=document.createElement('div');
  cell.className='day day-column';
  const isToday = sameDay(date, new Date());
  const dayNumberMarkup = isToday ? `<span class='day-number'>${date.getDate()}</span>` : `${date.getDate()}`;
  cell.innerHTML=`<div class='day-header' style='text-align:center;'>${dayNames[date.getDay()]} | <span dir='ltr'>${dayNumberMarkup}/${date.getMonth()+1}</span></div>`;
  if(isToday){
    cell.classList.add('today');
  }

  const dailyPool = [];
  data.forEach(r => r.Dates.forEach((dd, i) => {
    if(sameDay(dd, date)) dailyPool.push({ ...r, meetingIdx: i + 1, selectedDate: dd });
  }));

  const groupsMap = {};
  dailyPool.forEach(ev => {
    if(ev.EventType === 'HOLIDAY') {
      const key = `holiday-${ev.Program}`;
      if(!groupsMap[key]) groupsMap[key] = { type:'holiday', time: '00:00', items:[ev] };
  } else if(isEvent(ev)) {
      const key = `event-${ev.Employee}-${ev.Program}`;
      if(!groupsMap[key]) groupsMap[key] = { type:'event', time: ev.StartTime || '99:99', items:[] };
      groupsMap[key].items.push(ev);
  } else {
      const key = `${ev.Employee}-${ev.Program}`;
      if(!groupsMap[key]) groupsMap[key] = { type:'course', time: ev.StartTime || '99:99', items:[] };
      groupsMap[key].items.push(ev);
  }
  });

  const sortedGroups = Object.values(groupsMap).sort((a, b) => a.time.localeCompare(b.time));

  if(sortedGroups.length === 0) {
    cell.classList.add('empty');
    const emptyDay = document.createElement('div');
    emptyDay.className = 'empty-day';
    emptyDay.textContent = 'אין פעילות';
    cell.appendChild(emptyDay);
  }

  sortedGroups.forEach(g => {
    const evDiv = document.createElement('div');
    const first = g.items[0];
    const instructorName = first.Employee || first.Instructor || first.EmployeeName || '';
    const eventColor = g.type === 'holiday' ? '#cbd5e1' : instructorColor(instructorName);
    applyInstructorColorVars(evDiv, eventColor);
    evDiv.style.backgroundColor = eventColor;

    if(g.type === 'holiday') {
      evDiv.className = 'event schedule-card holiday';
      evDiv.innerHTML = `<div class="title">${first.Program || ""}</div>`;
  } else if(g.type === 'event') {
      evDiv.className = 'event schedule-card';
      const hourStr = (first.StartTime || first.EndTime) ? `<div class="event-hour">${first.StartTime || '—'}–${first.EndTime || '—'}</div>` : '';
      evDiv.innerHTML = `${hourStr}<strong class="title">${first.Program}</strong><div class="meta">פעילות יומית</div>`;
      evDiv.onclick = (e) => { e.stopPropagation(); openSideGrouped(g.items); };
  } else {
      const hasEmp = !!(first.Employee && first.Employee.trim());
      evDiv.className = 'event schedule-card' + (!hasEmp ? ' missing' : '');

      const count = g.items.length > 1 ? `<div class="group-count" role="button" tabindex="0" aria-label="צפייה בקבוצות">➕ ${g.items.length}</div>` : '';
      const hourStr = first.StartTime ? `<div class="event-hour">${first.StartTime}</div>` : '';
      const empName = hasEmp ? `<strong class="title">${first.Employee}</strong>` : `<strong class="title" style="color:var(--danger)">חסר מדריך</strong>`;
      const meta = `<div class="meta">${first.Program}</div>`;

      evDiv.innerHTML = `${count}${hourStr}${empName}${meta}`;
      evDiv.onclick = (e) => { e.stopPropagation(); openSideGrouped(g.items); };
      const groupCountEl = evDiv.querySelector('.group-count');
      if(groupCountEl){
        const openGroup = (e) => {
          e.stopPropagation();
          openSideGrouped(g.items);
      };
        groupCountEl.addEventListener('click', openGroup);
        groupCountEl.addEventListener('keydown', (e) => {
          if(e.key === 'Enter' || e.key === ' '){
            e.preventDefault();
            openGroup(e);
        }
      });
    }
  }
    cell.appendChild(evDiv);
  });
  return cell;
}

const PROGRAM_COLORS = {
  'ביומימיקרי': '#02b79c',
  'מנהיגות ירוקה': '#4caf50',
  'טכנולוגיות החלל': '#3b82f6',
  'ביומימיקרי לחטיבה': '#006717',
  'בינה מלאכותית': '#800080',
  'רוקחים עולם': '#a91515',
  'השמיים אינם הגבול': '#545454',
  'פורצות דרך': '#e61ca1',
  'יישומי AI': '#8106cd',
  'פרימיום': '#ff6700',

  'תלמידים להייטק': '#3b82f6',
  'מייקרים': '#0292b7',

  'תמיר - קווסט חדר בריחה': '#ff6700',
  'תמיר - איפה דדי?': '#ff6700'
};
function getProgramColor(name){ return PROGRAM_COLORS[name] || '#1e293b'; }

function shouldUseInstructorDaySheet(){
  return (userRole === 'instructor' || window._dualViewMode === 'instructor') && window.mode === 'month' && window.innerWidth <= 768;
}

function buildGroupedDetailsContent(items){
  const sortedItems = sortByDateAndTime(
    items.map(item => toDateAndTimeSortable(item, item.selectedDate || getEarliestDate(item.Dates), item.StartTime))
  );

  if(sortedItems.length === 0) return null;

  const first = sortedItems[0];
  let html = '';

  if(isEvent(first)){
    const timeRange = (first.StartTime || first.EndTime) ? `${first.StartTime} – ${first.EndTime}` : '—';
    html = `
      <h2>${first.Program}</h2>
      <div class='subtitle'>מנהל: ${getManagerForCourseViews(first) || '—'}</div>
      <div style="border-top:1px solid var(--border); margin-top:10px; padding-top:10px;"></div>
      <div class="group-item">
        <div class='row'><span class='label'>מדריך</span><span class='value'>${first.Employee || '—'}</span></div>
        <div class='row'><span class='label'>שעות</span><span class='value'>${timeRange}</span></div>
        ${first.Note ? `<div class='row'><span class='label'>הערה</span><span class='value'>${first.Note}</span></div>` : ''}
    </div>
  `;
    return { title: first.Program || 'פרטי פעילות', html };
  }

  const allSameProgram = sortedItems.every(i => i.Program === first.Program);

  if(allSameProgram){
    html = `
      <h2 style="color:${getProgramColor(first.Program)}">${first.Program}</h2>
      <div class='subtitle'>מנהל: ${getManagerForCourseViews(first) || '—'}</div>
      <div style="border-top:1px solid var(--border); margin-top:10px; padding-top:10px;"></div>
  `;
  } else {
    html = `
      <h2>פעילויות</h2>
      <div style="border-top:1px solid var(--border); margin-top:10px; padding-top:10px;"></div>
  `;
  }

  sortedItems.forEach(item => {
    const type = eventTypeOf(item);
    const isDaily = type === 'WORKSHOP' || type === 'TOUR';
    const end = endDate(item);
    const timeRange = (item.StartTime || item.EndTime) ? `${item.StartTime} – ${item.EndTime}` : '—';
    const empDisplay = (item.Employee && item.Employee.trim()) ? item.Employee : `<span style="color:var(--danger); font-weight:bold;">חסר מדריך</span>`;
    const notesHtml = renderNotesBlock(getNotesForCourseItem(item), item.Employee || '');
    const activityDate = item.selectedDate || getEarliestDate(item.Dates);
    const activityDateText = activityDate ? activityDate.toLocaleDateString('he-IL') : '—';
    const programHeader = !allSameProgram ? `<div style="font-weight:700;font-size:14px;color:${getProgramColor(item.Program)};margin-bottom:8px">${item.Program || '—'}</div>` : '';

    if(isDaily){
      html += `
        <div class="group-item">
          ${programHeader}
          <div class='row'><span class='label'>תאריך</span><span class='value'>${activityDateText}</span></div>
          <div class='row'><span class='label'>רשות</span><span class='value'>${item.Authority || '—'}</span></div>
          <div class='row'><span class='label'>בית ספר</span><span class='value'>${item.School || '—'}</span></div>
          <div class='row'><span class='label'>שעות</span><span class='value'>${timeRange}</span></div>
          <div class='row'><span class='label'>מדריך</span><span class='value'>${empDisplay}</span></div>
    </div>
    `;
      return;
  }

    html += `
      <div class="group-item">
        ${programHeader}
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
          <div class='badge'>מפגש ${item.meetingIdx}</div>
          <div style="background:#334155; color:#fff; padding:2px 8px; border-radius:6px; font-size:12px; font-weight:bold;">${timeRange}</div>
    </div>
        <div class='row'><span class='label'>מדריך</span><span class='value'>${empDisplay}</span></div>
        <div class='row'><span class='label'>בית ספר</span><span class='value'>${item.School || '—'}</span></div>
        <div class='row'><span class='label'>רשות</span><span class='value'>${item.Authority || '—'}</span></div>
        <div class='row'><span class='label'>סיום קורס</span><span class='value'>${end ? end.toLocaleDateString('he-IL') : '—'}</span></div>
        ${notesHtml}
    </div>
  `;
  });

  const sheetTitle = allSameProgram ? (first.Program || 'פרטי יום') : 'פרטי יום';
  return { title: sheetTitle, html };
}

function openSideGrouped(items) {
  const content = buildGroupedDetailsContent(items);
  if(!content) return;

  sideContent.innerHTML = content.html;
  applyNotesBoxColor();
  openSidePanel();
}

function applyFilters(){
  return rawData.filter(r =>
    isEventVisibleToCurrentUser(r) &&
    (!managerFilter.value || getManagerForCourseViews(r) === managerFilter.value) &&
    (!employeeFilter.value || r.Employee === employeeFilter.value)
  );
}

function getCourseStartDate(r){
  if(!Array.isArray(r.Dates) || r.Dates.length === 0) return null;
  return new Date(Math.min(...r.Dates.map(d => d.getTime())));
}

function getCourseEndDate(r){
  if(!Array.isArray(r.Dates) || r.Dates.length === 0) return null;
  return new Date(Math.max(...r.Dates.map(d => d.getTime())));
}

function isCourseActiveByRange(r, currentYear, currentMonth){
  if(!isCourse(r)) return false;

  const startDate = getCourseStartDate(r);
  const endDate = getCourseEndDate(r);

  if(!startDate || !endDate) return false;

  const monthStart = new Date(currentYear, currentMonth, 1);
  const monthEnd = new Date(currentYear, currentMonth + 1, 0);

  return startDate <= monthEnd && endDate >= monthStart;
}

function openMissingCourses(year, month){
  const missingCourses = rawData.filter(r => {

    if (String(r.EventType || '').trim().toUpperCase() !== 'COURSE')
      return false;

    if (r.Employee && r.Employee.trim())
      return false;

    const activeInMonth = r.Dates.some(d =>
      d &&
      d.getFullYear() === year &&
      d.getMonth() === month
    );

    const startDate = getCourseStartDate(r);
    const nextMonthStart = new Date(year, month + 1, 1);
    nextMonthStart.setHours(0,0,0,0);

    const isFuture =
      startDate &&
      startDate >= nextMonthStart;

    return activeInMonth || isFuture;
  });

  const sortedMissingCourses = sortByDateAndTime(
    missingCourses.map(r => toDateAndTimeSortable(r, getCourseStartDate(r), r.StartTime))
  );

  sideContent.innerHTML = `
    <h2>קורסים ללא מדריך</h2>
    <div class="subtitle">${sortedMissingCourses.length} קורסים</div>
    <div style="border-top:1px solid var(--border); margin:10px 0;"></div>
  `;

  sortedMissingCourses.forEach(r => {

    const startDate = getCourseStartDate(r);
    const end = endDate(r);

    sideContent.innerHTML += `
      <div class="course-card">
        <div style="font-weight:800;font-size:16px;margin-bottom:6px">
          ${r.Program || '—'}
    </div>
        <div>🏫 בית ספר: ${r.School || '—'}</div>
        <div>🌍 רשות: ${r.Authority || '—'}</div>
        <div>👨‍💼 מנהל: ${getCourseManager(r) || '—'}</div>
        <div>📅 התחלה: ${startDate ? startDate.toLocaleDateString('he-IL') : '—'}</div>
        <div>🏁 סיום: ${end ? end.toLocaleDateString('he-IL') : '—'}</div>
    </div>
  `;
  });

  openSidePanel();
  activeSidePanelType = 'missing-courses';
}

function openFutureOpenings(year, month){
  const nextMonthStart = new Date(year, month + 1, 1);
  nextMonthStart.setHours(0,0,0,0);

  const futureCourses = rawData.filter(r => {
    if(String(r.EventType || '').trim().toUpperCase() !== 'COURSE') return false;
    const startDate = getCourseStartDate(r);
    if(!startDate) return false;
    startDate.setHours(0,0,0,0);
    return startDate >= nextMonthStart;
  });

  const sortedFutureCourses = sortByDateAndTime(
    futureCourses.map(r => toDateAndTimeSortable(r, getCourseStartDate(r), r.StartTime))
  );

  sideContent.innerHTML = `
    <h2>קורסים נפתחים בעתיד</h2>
    <div class="subtitle">${sortedFutureCourses.length} קורסים</div>
    <div style="border-top:1px solid var(--border); margin:10px 0;"></div>
  `;

  sortedFutureCourses.forEach(r => {
    const startDate = getCourseStartDate(r);
    const instructor = (r.Employee && r.Employee.trim()) ? r.Employee : '—';

    sideContent.innerHTML += `
      <div class="course-card">
        <div style="font-weight:800;font-size:16px;margin-bottom:6px">${r.Program || '—'}</div>
        <div>📅 פתיחה: ${startDate ? startDate.toLocaleDateString('he-IL') : '—'}</div>
        <div>🌍 רשות: ${r.Authority || '—'}</div>
        <div>📘 קורס: ${r.Program || '—'}</div>
        <div>👤 מדריך: ${instructor}</div>
      </div>
    `;
  });

  openSidePanel();
  activeSidePanelType = 'future-openings';
}

function openJuneWarningCourses(juneWarningCourses){
  const sortedJuneWarningCourses = sortByDateAndTime(
    juneWarningCourses.map(r => toDateAndTimeSortable(r, getCourseStartDate(r), r.StartTime))
  );

  sideContent.innerHTML = `
    <h2>⚠ חודש יוני</h2>
    <div class="subtitle">${sortedJuneWarningCourses.length} קורסים</div>
    <div style="border-top:1px solid var(--border); margin:10px 0;"></div>
  `;

  sortedJuneWarningCourses.forEach(r => {
    const startDate = getCourseStartDate(r);
    const courseEndDate = getCourseEndDate(r);
    const instructor = (r.Employee && r.Employee.trim()) ? r.Employee : '—';

    sideContent.innerHTML += `
      <div class="course-card">
        <div>🌍 רשות: ${r.Authority || '—'}</div>
        <div>🏫 בית ספר: ${r.School || '—'}</div>
        <div>📘 תוכנית: ${r.Program || '—'}</div>
        <div>👤 מדריך: ${instructor}</div>
        <div>📅 תאריך התחלה: ${startDate ? startDate.toLocaleDateString('he-IL') : '—'}</div>
        <div>🏁 תאריך סיום: ${courseEndDate ? courseEndDate.toLocaleDateString('he-IL') : '—'}</div>
      </div>
    `;
  });

  openSidePanel();
  activeSidePanelType = 'june-warning';
}

function openManagerOverlay(mgr, year, month){
  const courses = rawData.filter(isCourse);
  const mgrCourses = courses.filter(r=>getCourseManager(r) === mgr);
  const mgrEndedThisMonth = mgrCourses.filter(r =>
    isCourseEndingInMonth(r, year, month)
  ).sort((a,b)=>parseDate(a.End)-parseDate(b.End));

  const mgrMissingActive = mgrCourses.filter(r =>
    isCourseActiveByRange(r, year, month) &&
    (!r.Employee || !r.Employee.trim())
  );

  sideContent.innerHTML = `
    <h2>${mgr}</h2>
    <div class="subtitle">קורסים מסתיימים החודש: ${mgrEndedThisMonth.length}</div>
    <div style="border-top:1px solid var(--border);margin:10px 0 14px;"></div>
    ${mgrMissingActive.length > 0 ? `
      <div style="background:#fef2f2;border:1.5px solid #dc2626;border-radius:12px;padding:14px;margin-bottom:16px">
        <div style="font-weight:800;color:#dc2626;margin-bottom:8px">⚠ קורסים פעילים ללא מדריך: ${mgrMissingActive.length}</div>
        ${mgrMissingActive.map(r=>`
          <div style="font-size:13px;padding:5px 0;border-bottom:1px solid #fecaca">
            ${r.Program || '—'}${r.School ? ` · ${r.School}` : ''}
      </div>`).join('')}
    </div>` : ''}
    ${mgrEndedThisMonth.length
      ? mgrEndedThisMonth.map(r=>{
          const empName = (r.Employee && r.Employee.trim())
            ? r.Employee
            : `<span style="color:#dc2626;font-weight:700">חסר מדריך</span>`;
          return `
            <div class="group-item">
              <div style="font-weight:800;font-size:15px;margin-bottom:6px">${r.Program}</div>
              <div style="font-size:13px;color:#475569;line-height:1.7">
                👤 מדריך: ${empName}<br>
                🏫 בית ספר: ${r.School || '—'}<br>
                🌍 רשות: ${r.Authority || '—'}<br>
                📅 סיום: ${parseDate(r.End).toLocaleDateString('he-IL')}
          </div>
        </div>`;
      }).join('')
      : '<div style="color:#94a3b8;text-align:center;padding:20px 0">אין קורסים המסתיימים החודש</div>'
  }
  `;
  openSidePanel();
}

function renderSummary(){
  const selectedValue = summaryMonth.value;

  let currentYear;
  let currentMonth;

  if(selectedValue){
    const [y,m] = selectedValue.split('-').map(Number);
    currentYear = y;
    currentMonth = m;
  } else {
    const today = new Date();
    currentYear = today.getFullYear();
    currentMonth = today.getMonth();
  }

  const currentMonthStart = new Date(currentYear, currentMonth, 1);

  let nextMonth = currentMonth + 1;
  let nextYear = currentYear;

  if(nextMonth > 11){
    nextMonth = 0;
    nextYear++;
  }

  const courses = rawData.filter(isCourse);
  const juneWarningCutoff = new Date('2026-06-18');
  const juneWarningCourses = courses.filter(c => {
    const courseEndDate = getCourseEndDate(c);
    return courseEndDate && courseEndDate > juneWarningCutoff;
  });

  const activeThisMonth = rawData.filter(r => {
    if(String(r.EventType || '').trim().toUpperCase() !== 'COURSE')
      return false;

    return isCourseActiveByRange(r, currentYear, currentMonth);
  }).length;

  const nextMonthStart = new Date(currentYear, currentMonth + 1, 1);
  nextMonthStart.setHours(0,0,0,0);

  const startingFuture = rawData.filter(r => {
    if(String(r.EventType || '').trim().toUpperCase() !== 'COURSE')
      return false;

    const startDate = getCourseStartDate(r);
    if(!startDate) return false;

    startDate.setHours(0,0,0,0);

    return startDate >= nextMonthStart;
  }).length;

  const totalCourses = activeThisMonth + startingFuture;

  const dailyActivitiesThisMonth = rawData.filter(r => {

    const type = String(r.EventType).trim().toUpperCase();

    const isDaily = type === 'WORKSHOP' || type === 'TOUR';

    const inMonth = r.Dates?.some(d =>
      d.getFullYear() === currentYear &&
      d.getMonth() === currentMonth
    );

    const isAllowed =
      (window.currentUserRole === 'admin' && window._dualViewMode !== 'instructor')
        ? true
        : r.EmployeeID == (window.currentUserEmployeeID || window.EmployeeID);

    return isDaily && inMonth && isAllowed;
  });

  const dailyCount = dailyActivitiesThisMonth.length;

  const missingInstructorCount = rawData.filter(r => {

    if(String(r.EventType || '').trim().toUpperCase() !== 'COURSE')
      return false;

    if(r.Employee && r.Employee.trim())
      return false;

    const activeInMonth = r.Dates.some(d =>
      d &&
      d.getFullYear() === currentYear &&
      d.getMonth() === currentMonth
    );

    const startDate = getCourseStartDate(r);
    const isFuture =
      startDate &&
      startDate >= nextMonthStart;

    return activeInMonth || isFuture;

  }).length;

  titleEl.textContent = currentMonthStart.toLocaleString('he-IL',{month:'long',year:'numeric'});

  const wrap = document.createElement('div');
  wrap.className = 'summary-wrapper';

  const instructorHoursButton = document.createElement('button');
  instructorHoursButton.type = 'button';
  instructorHoursButton.className = 'summary-action-btn instructor-hours-button';
  instructorHoursButton.innerHTML = 'סה"כ שעות <span class="ihb-arrow" aria-hidden="true">▼</span>';
  instructorHoursButton.setAttribute('aria-expanded', 'false');

  const instructorHoursPanel = document.createElement('section');
  instructorHoursPanel.className = 'summary-instructor-hours-panel instructor-hours-panel';

  const selectedMonth = currentMonth;
  const selectedYear = currentYear;

  const instructorHoursMap = {};
  let totalHoursAll = 0;

  rawData.forEach(record => {
    if(String(record.EventType || '').trim().toUpperCase() !== 'COURSE') return;

    const instructorName = String(record.Employee || '').trim();
    if(!instructorName) return;
    if(!Array.isArray(record.Dates)) return;

    record.Dates.forEach(date => {
      const sessionDate = new Date(date);
      if(Number.isNaN(sessionDate.getTime())) return;
      if(sessionDate.getMonth() !== selectedMonth) return;
      if(sessionDate.getFullYear() !== selectedYear) return;

      if(!instructorHoursMap[instructorName]){
        instructorHoursMap[instructorName] = { sessions: 0, totalHours: 0 };
      }

      instructorHoursMap[instructorName].sessions += 1;
      instructorHoursMap[instructorName].totalHours += 2;
      totalHoursAll += 2;
    });
  });

  const instructorRows = Object.entries(instructorHoursMap)
    .map(([name, totals]) => ({
      name,
      sessions: totals.sessions,
      totalHours: totals.totalHours
    }))
    .sort((a,b) => b.totalHours - a.totalHours || b.sessions - a.sessions || a.name.localeCompare(b.name, 'he'));

  if(instructorRows.length){
    const tableWrap = document.createElement('div');
    tableWrap.className = 'table-container';

    const table = document.createElement('table');
    table.className = 'summary-instructor-hours-table instructor-hours-table';
    table.innerHTML = `
      <thead>
        <tr>
          <th>מדריך</th>
          <th>שעות</th>
        </tr>
      </thead>
      <tbody>
        ${instructorRows.map(row => `
          <tr>
            <td>${escapeHtml(row.name)}</td>
            <td>${row.totalHours}</td>
          </tr>
        `).join('')}
        <tr>
          <td><strong>סה״כ</strong></td>
          <td><strong>${totalHoursAll}</strong></td>
        </tr>
      </tbody>
    `;

    tableWrap.appendChild(table);
    instructorHoursPanel.appendChild(tableWrap);
  } else {
    const emptyState = document.createElement('div');
    emptyState.className = 'summary-instructor-hours-empty';
    emptyState.textContent = 'אין מפגשי קורס בחודש הנבחר.';
    instructorHoursPanel.appendChild(emptyState);
  }

  instructorHoursButton.onclick = () => {
    const willOpen = !instructorHoursPanel.classList.contains('is-open');
    instructorHoursPanel.classList.toggle('is-open', willOpen);
    instructorHoursButton.setAttribute('aria-expanded', String(willOpen));
    if (willOpen) {
      const arrow = instructorHoursButton.querySelector('.ihb-arrow');
      if (arrow) {
        arrow.classList.remove('ihb-bounce');
        void arrow.offsetWidth;
        arrow.classList.add('ihb-bounce');
        setTimeout(() => arrow.classList.remove('ihb-bounce'), 900);
      }
    }
  };

  wrap.innerHTML = `
    <div class="kpi-total">
      <div class="kpi-title">סה"כ קורסים פעילים</div>
      <div class="kpi-number">${totalCourses}</div>
    </div>

    <div class="kpi-row">
      <div class="kpi-small blue">
        <div class="kpi-title">פעילים החודש</div>
        <div class="kpi-number">${activeThisMonth}</div>
    </div>

      <div class="kpi-small green" data-action="future-openings" role="button" aria-label="קורסים נפתחים בעתיד" style="cursor:pointer">
        <div class="kpi-title">נפתחים בעתיד</div>
        <div class="kpi-number">${startingFuture}</div>
    </div>

      <div class="kpi-small orange">
        <div class="kpi-title">סדנאות וסיורים</div>
        <div class="kpi-number">${dailyCount}</div>
    </div>

      <div class="kpi-small" data-action="june-warning" role="button" aria-label="חודש יוני" style="cursor:pointer;background:#fee2e2;border:1px solid #ef4444;color:#7f1d1d;">
        <div class="kpi-title">⚠ חודש יוני</div>
        <div class="kpi-number" style="color:#7f1d1d;font-weight:900;">${juneWarningCourses.length}</div>
    </div>
    </div>

    ${missingInstructorCount > 0 ? `
      <div class="alert-missing" data-action="missing" style="cursor:pointer">
        ⚠ חסרים ${missingInstructorCount} מדריכים לשיבוץ
    </div>
    ` : ''}
  `;

  const managers = [...new Set(rawData.filter(isCourse).map(r=>getCourseManager(r)).filter(Boolean))]
    .sort((a,b)=>a.localeCompare(b,'he'));
  
  const split = document.createElement('div');
  split.className = 'managers-row';
  split.style.marginTop = '20px';

  managers.forEach((mgr,index)=>{
    const mgrCourses = courses.filter(r=>getCourseManager(r) === mgr);
    const mgrActive = mgrCourses.filter(r =>
      isCourseActiveByRange(r, currentYear, currentMonth)
    ).length;
    const mgrEndedThisMonth = mgrCourses.filter(r =>
      isCourseEndingInMonth(r, currentYear, currentMonth)
    ).sort((a,b)=>parseDate(a.End)-parseDate(b.End));

    const mgrFuture = mgrCourses.filter(r => {
      const start = getCourseStartDate(r);
      return start && start >= nextMonthStart;
  }).length;

    const col = document.createElement('div');
    col.className = `manager-card${index === 1 ? ' secondary' : ''}`;
    col.dataset.manager = mgr;
    col.innerHTML = `
      <div class="manager-name">${mgr}</div>
      <div class="manager-metric">
        <span>קורסים פעילים</span>
        <strong>${mgrActive}</strong>
    </div>
      <div class="manager-metric">
        <span>מסתיימים החודש</span>
        <strong>${mgrEndedThisMonth.length}</strong>
    </div>
      <div class="manager-metric">
        <span>נפתחים בעתיד</span>
        <strong>${mgrFuture}</strong>
    </div>`;

    split.appendChild(col);
  });
  wrap.appendChild(split);
  wrap.appendChild(instructorHoursButton);
  wrap.appendChild(instructorHoursPanel);
  view.appendChild(wrap);

  wrap.addEventListener('click', function(e){

    // === חסר מדריך ===
    if(e.target.closest('[data-action="missing"]')){
      e.stopPropagation();
      openMissingCourses(currentYear, currentMonth);
      return;
    }

    if(e.target.closest('[data-action="future-openings"]')){
      e.stopPropagation();
      if(side.classList.contains('open') && activeSidePanelType === 'future-openings'){
        closeSidePanel();
        return;
      }
      openFutureOpenings(currentYear, currentMonth);
      return;
    }

    if(e.target.closest('[data-action="june-warning"]')){
      e.stopPropagation();
      if(side.classList.contains('open') && activeSidePanelType === 'june-warning'){
        closeSidePanel();
        return;
      }
      openJuneWarningCourses(juneWarningCourses);
      return;
    }

    // === מנהל ===
    const managerCol = e.target.closest('[data-manager]');
    if(managerCol){
      e.stopPropagation();
      const mgrName = managerCol.dataset.manager;
      openManagerOverlay(mgrName, currentYear, currentMonth);
  }

  });
}

function getUniqueInstructorMonths(){
  const monthsMap = new Map();
  const minAllowed = getMinAllowedMonth();
  rawData.forEach(r=>{
    r.Dates.forEach(d=>{
      if(!d) return;
      const year = d.getFullYear();
      const month = d.getMonth();
      const key = `${year}-${String(month+1).padStart(2,'0')}`;
      if(!monthsMap.has(key)){
        monthsMap.set(key,{year,month});
    }
  });
  });

  return [...monthsMap.entries()]
    .sort((a,b)=>
      (a[1].year - b[1].year) ||
      (a[1].month - b[1].month)
    )
    .map(([,obj])=>obj)
    .filter(obj =>
      new Date(obj.year,obj.month,1) >= minAllowed
    )
    .map(({year,month})=>({
      value:`${year}-${String(month+1).padStart(2,'0')}`,
      label:new Date(year,month,1)
        .toLocaleDateString('he-IL',{month:'long',year:'numeric'})
  }));
}

function renderInstructors(){

  titleEl.textContent = "מדריכים";

  const managers=[...new Set(rawData.map(r=>getInstructorManager(r)).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'he'));
  const selectedManager = renderInstructors.selectedManager || '';

  const controls = document.createElement('div');
  controls.className = 'controls';
  controls.style.margin = '10px auto 0';
  controls.style.maxWidth = '1200px';

  const managerSelect = document.createElement('select');
  managerSelect.innerHTML = '<option value="">כל המנהלים</option>' + managers.map(v=>`<option value="${v}">${v}</option>`).join('');
  managerSelect.value = selectedManager;
  managerSelect.onchange = () => {
    renderInstructors.selectedManager = managerSelect.value;
    render();
  };
  controls.appendChild(managerSelect);

  const monthSelect = document.createElement('select');
  const monthOptions = getUniqueInstructorMonths();
  monthSelect.innerHTML = monthOptions
    .map(opt=>`<option value="${opt.value}">${opt.label}</option>`)
    .join('');

  const todayValue =
    `${currentDate.getFullYear()}-${String(currentDate.getMonth()+1).padStart(2,'0')}`;

  if(!renderInstructors.selectedMonthValue){

    if(monthOptions.some(opt=>opt.value === todayValue)){
      renderInstructors.selectedMonthValue = todayValue;
  }
    else if(monthOptions.length > 0){
      renderInstructors.selectedMonthValue = monthOptions[0].value;
  }

  }

  monthSelect.value = renderInstructors.selectedMonthValue;
  monthSelect.onchange = () => {
    renderInstructors.selectedMonthValue = monthSelect.value;
    render();
  };
  controls.appendChild(monthSelect);

  view.appendChild(controls);

  let selectedMonth = currentDate.getMonth();
  let selectedYear = currentDate.getFullYear();

  const selectedMonthValue = renderInstructors.selectedMonthValue;

  const [y,m] = selectedMonthValue.split('-').map(Number);
  selectedYear = y;
  selectedMonth = m - 1;

  const instructorsMap = {};
  const instructorMetaByName = {};
  const missingInstructorCourses = [];
  const allActiveCourses = [];

  rawData.forEach(r => {
    if(!isCourse(r)) return;
    if(!isCourseActiveByRange(r, selectedYear, selectedMonth)) return;

    if(selectedManager && getInstructorManager(r) !== selectedManager) return;

    allActiveCourses.push(r);

    if(!r.Employee || !r.Employee.trim()){
      missingInstructorCourses.push(r);
      return;
  }

    if(!instructorsMap[r.Employee]){
      instructorsMap[r.Employee] = [];
      instructorMetaByName[r.Employee] = {
        EmployeeID: String(r.EmployeeID || '').trim(),
        Employee: r.Employee
    };
  }
    instructorsMap[r.Employee].push(r);
  });

  const instructorDailyCountByName = {};

  Object.keys(instructorsMap).forEach(name => {
    const instructor = instructorMetaByName[name] || {};

    const instructorDailyCount = rawData.filter(r => {

      const type = String(r.EventType).trim().toUpperCase();

      const isDaily = type === 'WORKSHOP' || type === 'TOUR';

      const inMonth = r.Dates?.some(d =>
        d.getFullYear() === selectedYear &&
        d.getMonth() === selectedMonth
      );

      return isDaily &&
             inMonth &&
             r.EmployeeID == instructor.EmployeeID;

  }).length;

    instructorDailyCountByName[name] = instructorDailyCount;
  });

  const visibleInstructorNames = Object.keys(instructorsMap).filter(name => {
    if(userRole === 'admin') return true;
    if(userRole === 'instructor'){
      const instructor = instructorMetaByName[name] || {};
      return String(instructor.EmployeeID || '').trim() === String(window.currentUserEmployeeID || window.EmployeeID || '').trim();
  }
    return true;
  });

  const names = visibleInstructorNames
    .sort((a,b)=> instructorsMap[b].length - instructorsMap[a].length || a.localeCompare(b,'he'));

  const totalCourses = allActiveCourses.length;

  const summaryHeader = document.createElement('div');
  summaryHeader.style.textAlign = 'center';
  summaryHeader.style.margin = '10px 0 20px 0';
  summaryHeader.style.fontSize = '16px';
  summaryHeader.style.fontWeight = '700';
  summaryHeader.style.lineHeight = '1.8';

  summaryHeader.innerHTML = `
    כמות מדריכים: ${names.length}<br>
    כמות קורסים: ${totalCourses}
    ${userRole === 'admin' && missingInstructorCourses.length > 0
      ? `<br><span style="color:#dc2626;font-weight:800">⚠ חסר מדריך: ${missingInstructorCourses.length} קורסים</span>`
      : ''}
  `;

  view.appendChild(summaryHeader);

  if(names.length===0 && (userRole !== 'admin' || missingInstructorCourses.length===0)){
    const empty = document.createElement('div');
    empty.style.textAlign = 'center';
    empty.style.padding = '40px';
    empty.style.color = '#94a3b8';
    empty.textContent = 'לא נמצאו מדריכים פעילים';
    view.appendChild(empty);
    return;
  }

  const grid = document.createElement('div');
  grid.className = 'instructors-grid';

  if(userRole === 'admin' && missingInstructorCourses.length > 0){
    const box = document.createElement('div');
    box.className = 'instructor-card instructor-card-missing';

    box.innerHTML = `
      <div class="instructor-card-name" style="color:var(--danger)">חסר מדריך</div>
      <div class="instructor-card-count" style="color:var(--danger)">${missingInstructorCourses.length}</div>
      <div class="instructor-card-label">תוכניות פעילות</div>
  `;

    box.onclick = (e)=>{ e.stopPropagation(); openInstructorModal("חסר מדריך", missingInstructorCourses, selectedMonth, selectedYear); };

    grid.appendChild(box);
  }

  names.forEach(name=>{
    const instructorDailyCount = instructorDailyCountByName[name] || 0;
    const dailyWorkshopsContent = instructorDailyCount > 0
      ? `סדנאות/סיורים: ${instructorDailyCount}`
      : '&nbsp;';

    const box = document.createElement('div');
    box.className = 'instructor-card';
    box.style.background = getEmployeeColor(name);

    box.addEventListener('mouseenter', ()=>{ if(!window.matchMedia('(hover:none)').matches) box.style.transform='scale(1.03)'; });
    box.addEventListener('mouseleave', ()=>{ box.style.transform='scale(1)'; });

    box.innerHTML = `
      <div class="instructor-card-name">${name}</div>
      <div class="instructor-card-count">${instructorsMap[name].length}</div>
      <div class="instructor-card-label">קורסים פעילים</div>
      <div class="instructor-card-daily">${dailyWorkshopsContent}</div>
  `;

    box.onclick = (e)=>{ e.stopPropagation(); openInstructorModal(name, instructorsMap[name], selectedMonth, selectedYear); };

    grid.appendChild(box);
  });

  view.appendChild(grid);
}

function openInstructorModal(name, courses, selectedMonth, selectedYear){
  console.log('פתיחת מודל מדריך:', name);
  const month = selectedMonth ?? currentDate.getMonth();
  const year  = selectedYear ?? currentDate.getFullYear();

  let totalWorkDays = 0;

  function normalize(d){
    const n = new Date(d);
    n.setHours(0,0,0,0);
    return n;
  }

  const weeks = {};

  const courseOnlyRecords = courses.filter(r => isCourse(r));

  const instructorEmpID = courseOnlyRecords?.[0]?.EmployeeID;
  const dailyRecords = instructorEmpID ? rawData.filter(r => {
    const type = String(r.EventType || '').trim().toUpperCase();
    const isDaily = type === 'WORKSHOP' || type === 'TOUR';
    const inMonth = r.Dates?.some(d =>
      d.getFullYear() === year &&
      d.getMonth() === month
    );
    return isDaily && inMonth && r.EmployeeID == instructorEmpID;
  }).sort((a, b) => {
    const da = getEarliestDate(a.Dates) || new Date(0);
    const db = getEarliestDate(b.Dates) || new Date(0);
    return da - db;
  }) : [];

  const sortedCourses = sortByDateAndTime(
    courseOnlyRecords.map(r => toDateAndTimeSortable(r, getEarliestDate(r.Dates), r.StartTime))
  );

  sortedCourses.forEach(r=>{

    if(!isCourse(r)) return;

    r.Dates.forEach(d=>{

      if(!d) return;

      const date = normalize(d);

      if(
        date.getFullYear() === selectedYear &&
        date.getMonth() === selectedMonth
      ){

        const weekStart = new Date(date);
        weekStart.setDate(date.getDate() - date.getDay());
        weekStart.setHours(0,0,0,0);

        const key = weekStart.toISOString();

        if(!weeks[key]){
          weeks[key] = new Set();
      }

        weeks[key].add(date.toDateString());
    }

  });

  });

  let maxDays = 0;

  Object.values(weeks).forEach(set=>{
    if(set.size > maxDays){
      maxDays = set.size;
  }
  });

  totalWorkDays = maxDays;
  const employmentType = getEmploymentTypeForEmployeeId(courseOnlyRecords?.[0]?.EmployeeID);
  const managerName = getInstructorManager(courseOnlyRecords[0]) || '—';

  // הוספת שורת רשות אם מדובר ב"חסר מדריך"
  let authorityRow = '';
  if (name === "חסר מדריך" && courseOnlyRecords[0]?.Authority) {
    authorityRow = `<span class="badge" style="background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;">רשות: ${courseOnlyRecords[0].Authority}</span>`;
  }

  sideContent.innerHTML = `
    <div class="instructor-header">
      <div class="instructor-name">${name}</div>
      <div class="instructor-badges">
        <span class="badge type">${employmentType}</span>
        <span class="badge days">${totalWorkDays} ימי עבודה בשבוע</span>
        <span class="badge courses">${sortedCourses.length} קורסים</span>
          ${dailyRecords.length > 0 ? `<span class="badge" style="background:#d1fae5;color:#065f46;">${dailyRecords.length} סדנאות/סיורים</span>` : ''}
        <span class="badge">מנהל: ${managerName}</span>
        ${authorityRow}
      </div>
      </div>
  `;

  sortedCourses.forEach(r=>{

    const end = endDate(r);

    const courseStartDate = getEarliestDate(r.Dates);

    const startDate = courseStartDate
      ? courseStartDate.toLocaleDateString('he-IL')
      : '—';

    const endDateFormatted = end
      ? end.toLocaleDateString('he-IL')
      : '—';

    sideContent.innerHTML += `
      <div class="course-card">
        <div class="course-title" style="background:${getEmployeeColor(name)};">${r.Program || '—'}</div>
        <div>
          <span style="font-weight:700;color:#0f172a;">בית ספר:</span> ${r.School || '—'}<br>
          <span style="font-weight:700;color:#0f172a;">רשות:</span> ${r.Authority || '—'}
        </div>
        <div>
          <div style="font-weight:700;color:#0f172a;">תאריכי פעילות</div>
          <div>(${startDate}) - (${endDateFormatted})</div>
        </div>
      </div>
  `;
  });

  if(dailyRecords.length > 0){
    sideContent.innerHTML += `
      <div style="font-weight:800;font-size:15px;margin:16px 0 8px;color:#1b7895;border-top:1px solid var(--border);padding-top:14px;">
        סדנאות וסיורים
      </div>
  `;
    dailyRecords.forEach(r => {
      const type = String(r.EventType || '').trim().toUpperCase();
      const typeLabel = type === 'TOUR' ? 'סיור' : 'סדנה';
      const activityDate = getEarliestDate(r.Dates);
      const activityDateText = activityDate ? activityDate.toLocaleDateString('he-IL') : '—';
      const timeRange = (r.StartTime || r.EndTime) ? `${r.StartTime || ''} – ${r.EndTime || ''}` : '—';
      sideContent.innerHTML += `
        <div class="course-card">
          <div class="course-title" style="background:#d1fae5;color:#065f46;">
            ${r.Program || typeLabel}
            <span style="font-size:12px;font-weight:600;margin-right:6px;">(${typeLabel})</span>
          </div>
          <div>
            <span style="font-weight:700;color:#0f172a;">בית ספר:</span> ${r.School || '—'}<br>
            <span style="font-weight:700;color:#0f172a;">רשות:</span> ${r.Authority || '—'}
          </div>
          <div>
            <div style="font-weight:700;color:#0f172a;">תאריך</div>
            <div>${activityDateText}${timeRange !== '—' ? ' | ' + timeRange : ''}</div>
          </div>
        </div>
  `;
  });
  }

  openSidePanel();
}

document.getElementById('prev').onclick = ()=>{
  enforceInstructorMode();
  if(!canGoPrev()) return;

  if(window.mode==='summary'){
    summaryMonth.selectedIndex = Math.max(0, summaryMonth.selectedIndex-1);
  }
  else if(window.mode==='week'){
    const temp = new Date(currentDate);
    temp.setDate(temp.getDate()-7);

    if(temp >= getMinAllowedMonth()){
      currentDate = temp;
  }
  }
  else if(window.mode==='month'){
    if((userRole === 'instructor' || window._dualViewMode === 'instructor') && isMobile()){
      // מדריך במובייל – ניווט שבועי
      const temp = new Date(currentDate);
      temp.setDate(temp.getDate() - 7);
      if(temp >= getMinAllowedMonth()) currentDate = temp;
  } else {
      const temp = new Date(currentDate);
      temp.setMonth(temp.getMonth()-1);
      if(temp >= getMinAllowedMonth()) currentDate = temp;
  }
  }

  render();
};
document.getElementById('next').onclick = ()=>{
  enforceInstructorMode();
  if(!canGoNext()) return;

  if(window.mode==='summary'){
    summaryMonth.selectedIndex = Math.min(summaryMonth.options.length-1, summaryMonth.selectedIndex+1);
  }
  else if(window.mode==='week'){
    const temp = new Date(currentDate);
    temp.setDate(temp.getDate()+7);

    const weekStart = new Date(temp);
    weekStart.setDate(temp.getDate() - temp.getDay());
    weekStart.setHours(0,0,0,0);

    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    weekEnd.setHours(0,0,0,0);

    if(dataRange && weekEnd >= dataRange.min && weekStart <= dataRange.max){
      currentDate = temp;
  }
  }
  else if(window.mode==='month'){
    if((userRole === 'instructor' || window._dualViewMode === 'instructor') && isMobile()){
      // מדריך במובייל – ניווט שבועי
      const temp = new Date(currentDate);
      temp.setDate(temp.getDate() + 7);
      if(weekOverlapsDataRange(temp)) currentDate = temp;
  } else {
      const temp = new Date(currentDate.getFullYear(), currentDate.getMonth()+1, 1);
      if(dataRange){
        const maxMonth = new Date(dataRange.max.getFullYear(), dataRange.max.getMonth(), 1);
        if(temp <= maxMonth) currentDate = temp;
    }
  }
  }

  render();
};
btnMonth.onclick = ()=>{
  window.mode='month';
  currentDate = clampDateToDataRange(new Date());
  render();
};
btnWeek.onclick = ()=>{
  if(userRole === 'instructor' || window._dualViewMode === 'instructor') return;
  window.mode='week';
  currentDate = clampDateToDataRange(new Date());
  render();
};
btnSummary.onclick = ()=>{
  if(userRole === 'instructor' || window._dualViewMode === 'instructor') return;
  window.mode='summary';
  render();
};
btnInstructors.onclick = ()=>{
  if(userRole === 'instructor' || window._dualViewMode === 'instructor') return;
  window.mode='instructors';
  render();
};
if(btnEndDates){
  btnEndDates.onclick = ()=>{
    window.mode='enddates';
    render();
  };
}

if(btnZoom){
  btnZoom.onclick = ()=>{
    if(userRole === 'instructor' || window._dualViewMode === 'instructor') return;
    window.mode='zoom';
    render();
  };
}

// ─── ZOOM management helpers ──────────────────────────────────────────────────
function zoomCourseKey(dayNum, course) {
  const dateKey = normalizeZoomDateKey(course?.Date || course?.date || zoomDateString(dayNum));
  return [
    dateKey,
    course?.EmployeeID || course?.employeeId || '',
    course?.Employee || '',
    course?.Program || '',
    normalizeZoomTime(course?.StartTime || course?.startTime || ''),
    normalizeZoomTime(course?.EndTime || course?.endTime || ''),
    course?.Authority || '',
    course?.School || ''
  ].map(v => String(v || '').trim()).join('|');
}
function zoomCourseId(dayNum, course){
  const directId = course?.Id ?? course?.CourseId ?? course?.id;
  if(directId != null && String(directId).trim() !== '') return String(directId).trim();
  return zoomCourseKey(dayNum, course);
}

function clearZoomCache(){
  window.zoomDataCache = null;
  window.zoomGoogleCourses = [];
  window.zoomAssignments = {};
  window.zoomReadOnlyMode = false;
}

function normalizeZoomDateKey(value){
  if(!value) return '';
  if(value instanceof Date && !Number.isNaN(value.getTime())){
    return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;
  }
  const raw = String(value).trim();
  if(!raw) return '';

  // Keep date-only / ISO date keys timezone-stable for ZOOM.
  // Never pass these through `new Date(...)` because timezone conversion can shift day.
  const isoLike = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(isoLike) return `${isoLike[1]}-${isoLike[2]}-${isoLike[3]}`;

  // Handle full ISO strings explicitly and preserve the encoded calendar date.
  const isoDateTime = raw.match(/^(\d{4})-(\d{2})-(\d{2})[T\s]\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$/);
  if(isoDateTime) return `${isoDateTime[1]}-${isoDateTime[2]}-${isoDateTime[3]}`;

  const slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if(slash){
    const yyyy = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
    return `${yyyy}-${String(Number(slash[2])).padStart(2,'0')}-${String(Number(slash[1])).padStart(2,'0')}`;
  }

  const dotted = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if(dotted){
    const yyyy = dotted[3].length === 2 ? `20${dotted[3]}` : dotted[3];
    return `${yyyy}-${String(Number(dotted[2])).padStart(2,'0')}-${String(Number(dotted[1])).padStart(2,'0')}`;
  }

  return '';
}

function normalizeZoomTime(value){
  if(value == null || value === '') return '';
  if(typeof value === 'number'){
    if(value >= 0 && value < 1){
      const totalMinutes = Math.round(value * 24 * 60);
      const hh = String(Math.floor(totalMinutes / 60) % 24).padStart(2, '0');
      const mm = String(totalMinutes % 60).padStart(2, '0');
      return `${hh}:${mm}`;
    }
    const totalMinutes = Math.round(value);
    if(totalMinutes >= 0 && totalMinutes < (24 * 60)){
      const hh = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
      const mm = String(totalMinutes % 60).padStart(2, '0');
      return `${hh}:${mm}`;
    }
  }
  const raw = String(value).trim();
  if(!raw) return '';
  const m = raw.match(/^(\d{1,2})[:.](\d{1,2})(?::\d{1,2})?$/);
  if(m){
    const hh = Math.max(0, Math.min(23, Number(m[1])));
    const mm = Math.max(0, Math.min(59, Number(m[2])));
    return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }
  const parsed = new Date(raw);
  if(!Number.isNaN(parsed.getTime())){
    return `${String(parsed.getHours()).padStart(2, '0')}:${String(parsed.getMinutes()).padStart(2, '0')}`;
  }
  return '';
}

function getZoomEmployeeMap(){
  if(window.zoomDataCache?.employeeMap) return window.zoomDataCache.employeeMap;
  const employeeMap = {};
  rawData.forEach(r => {
    if(r.Employee && r.EmployeeID) employeeMap[r.Employee] = String(r.EmployeeID);
  });
  if(!window.zoomDataCache) window.zoomDataCache = {};
  window.zoomDataCache.employeeMap = employeeMap;
  return employeeMap;
}

function updateZoomAssignmentState(courseKey, patch){
  if(!courseKey) return;
  const key = String(courseKey).trim();
  if(!key) return;
  if(!window.zoomAssignments) window.zoomAssignments = {};
  if(!window.zoomAssignments[key]) {
    window.zoomAssignments[key] = { account: null, notes: '', conflict: false };
  }
  // Mutate in-place so existing `asgn` references in event handlers stay valid
  const next = window.zoomAssignments[key];
  Object.assign(next, patch);
  if(patch.startTime !== undefined) next.startTime = normalizeZoomTime(patch.startTime);
  if(patch.endTime !== undefined) next.endTime = normalizeZoomTime(patch.endTime);
  if(patch.date !== undefined) next.date = normalizeZoomDateKey(patch.date);
  console.log('[ZOOM][State] updateZoomAssignmentState', { key, patch, next });

  if(!window.zoomDataCache) window.zoomDataCache = {};
  if(!window.zoomDataCache.assignments) window.zoomDataCache.assignments = {};
  window.zoomDataCache.assignments[key] = { ...window.zoomDataCache.assignments[key], ...next };
}

function refreshZoomCalendarView(){
  console.log('[ZOOM][Calendar] refreshZoomCalendarView called', {
    mode: window.mode,
    subView: window.zoomSubView,
    dirty: !!window.zoomCalendarDirty
  });
  if(window.mode !== 'zoom' || window.zoomSubView !== 'calendar') return;
  const content = view.querySelector('.zoom-content');
  if(!content) return;
  content.innerHTML = '';
  renderZoomCalendar(content, window.zoomCoursesForView || [], window.zoomDaysForView || [], window.zoomHdaysForView || []);
  window.zoomCalendarDirty = false;
}

function requestZoomCalendarRefresh(){
  console.log('[ZOOM][Calendar] requestZoomCalendarRefresh called', {
    mode: window.mode,
    subView: window.zoomSubView
  });
  window.zoomCalendarDirty = true;
  console.log('[ZOOM][Calendar] zoomCalendarDirty=true');
  if(window.mode === 'zoom' && window.zoomSubView === 'calendar'){
    console.log('[ZOOM][Calendar] in calendar subView; refreshing immediately');
    refreshZoomCalendarView();
  } else {
    console.log('[ZOOM][Calendar] refresh deferred until calendar tab is active');
  }
}

async function getZoomData(forceReload = false){
  if(!window.zoomDataCache) window.zoomDataCache = {};
  if(!forceReload && window.zoomDataCache.loaded){
    return {
      courses: window.zoomDataCache.courses || [],
      assignments: window.zoomDataCache.assignments || {}
    };
  }
  const [googleCourses, rawAssignments] = await Promise.all([
    loadCoursesFromGoogle(),
    loadZoomAssignmentsFromGoogle()
  ]);
  const assignments = mapZoomAssignmentsByCourseKey(rawAssignments);
  window.zoomDataCache.courses = googleCourses;
  window.zoomDataCache.assignments = assignments;
  window.zoomDataCache.loaded = true;
  console.log('[ZOOM] loaded courses:', googleCourses.length);
  console.log('[ZOOM] loaded assignments:', Array.isArray(rawAssignments) ? rawAssignments.length : Object.keys(assignments).length);
  return { courses: googleCourses, assignments };
}
async function loadZoomAssignments(){
  try {
    const res = await fetch(API_URL, { cache:"no-store" });
    const data = await res.json();
    window.zoomReadOnlyMode = false;
    return data;
  } catch(err){
    console.error("Failed loading zoom assignments", err);
    console.error('Zoom API connection failed');
    window.zoomReadOnlyMode = true;
    return [];
  }
}
async function loadCoursesFromGoogle() {
  try {
    const res = await fetch(API_URL + '?type=courses', { cache: 'no-store' });
    const data = await res.json();
    return Array.isArray(data) ? data : [];
  } catch(err) {
    console.error('Failed loading courses from Google', err);
    return [];
  }
}
async function loadZoomAssignmentsFromGoogle() {
  try {
    const res = await fetch(API_URL + '?type=assignments', { cache: 'no-store' });
    const data = await res.json();
    window.zoomReadOnlyMode = false;
    return Array.isArray(data) ? data : [];
  } catch(err) {
    console.error('Failed loading zoom assignments from Google', err);
    window.zoomReadOnlyMode = true;
    return [];
  }
}
async function saveZoomAssignment(data){
  try {
    const body = new URLSearchParams();
    body.append("CourseId",   data.CourseId   || "");
    body.append("Date",       data.Date       || "");
    body.append("Authority",  data.Authority  || "");
    body.append("School",     data.School     || "");
    body.append("Program",    data.Program    || "");
    body.append("Employee",   data.Employee   || "");
    body.append("EmployeeID", data.EmployeeID || "");
    body.append("StartTime",  data.StartTime  || "");
    body.append("EndTime",    data.EndTime    || "");
    body.append("ZoomAccount",data.ZoomAccount|| "");
    body.append("Notes",      data.Notes      || "");
    const res = await fetch(API_URL, { method: "POST", body });
    const result = await res.json();
    return result;
  } catch(err){
    console.error("Failed saving zoom assignment", err);
    throw err;
  }
}

function canAssignZoom(){
  const employeeId = String(window.EmployeeID || '').trim();
  return employeeId === '8000' || employeeId === '6000';
}

function notifyZoomNoPermission(){
  const message = 'אין הרשאה לבצע שיבוץ';
  console.warn(message);
  alert(message);
}

function mapZoomAssignmentsByCourseKey(rows){
  const mapped = {};
  if(Array.isArray(rows)){
    rows.forEach((row)=>{
      const courseKey = String(row?.CourseId || row?.courseId || row?.courseKey || row?.CourseKey || row?.Id || row?.id || '').trim();
      if(!courseKey) return;
      const account = row?.ZoomAccount || row?.zoom || row?.Zoom || row?.account || null;
      const notes = row?.Notes || row?.notes || '';
      mapped[courseKey] = {
        account,
        notes,
        startTime:  normalizeZoomTime(row?.StartTime  || row?.startTime  || row?.start || ''),
        endTime:    normalizeZoomTime(row?.EndTime    || row?.endTime    || row?.end || ''),
        date:       normalizeZoomDateKey(row?.Date || row?.date || ''),
        authority:  row?.Authority  || row?.authority  || '',
        school:     row?.School     || row?.school     || '',
        program:    row?.Program    || row?.program    || '',
        employee:   row?.Employee   || row?.employee   || '',
        employeeId: String(row?.EmployeeID || row?.employeeId || ''),
        conflict: false
      };
    });
    return mapped;
  }

  if(rows && typeof rows === 'object'){
    Object.entries(rows).forEach(([courseKey, value])=>{
      const account = value?.ZoomAccount || value?.zoom || value?.Zoom || value?.account || null;
      const notes = value?.Notes || value?.notes || '';
      mapped[String(courseKey)] = {
        account,
        notes,
        startTime:  normalizeZoomTime(value?.StartTime  || value?.startTime  || ''),
        endTime:    normalizeZoomTime(value?.EndTime    || value?.endTime    || ''),
        date:       normalizeZoomDateKey(value?.Date || value?.date || ''),
        authority:  value?.Authority  || value?.authority  || '',
        school:     value?.School     || value?.school     || '',
        program:    value?.Program    || value?.program    || '',
        employee:   value?.Employee   || value?.employee   || '',
        employeeId: value?.EmployeeID || value?.employeeId || '',
        conflict: !!value?.conflict
      };
    });
  }

  return mapped;
}

function zoomDateString(dayNum){
  const d = new Date(2026, 2, dayNum);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function logZoomCourseDiagnostics(courses, days){
  const totalLoaded = Array.isArray(courses) ? courses.length : 0;
  const normalized = (Array.isArray(courses) ? courses : []).map((course, idx) => ({
    idx,
    date: normalizeZoomDateKey(course.Date || course.date || ''),
    key: zoomCourseId(Number(String(normalizeZoomDateKey(course.Date || course.date || '')).slice(-2)) || -1, course),
    course
  }));
  const withDate = normalized.filter(x => !!x.date);
  console.log('[ZOOM][Diag] total courses loaded:', totalLoaded);
  console.log('[ZOOM][Diag] total courses after normalizeZoomDateKey:', withDate.length);

  const requestedDays = Array.isArray(days) ? days : [];
  requestedDays.forEach(dayNum => {
    const dateStr = zoomDateString(dayNum);
    const dayRows = withDate.filter(x => x.date === dateStr);
    const keyCounts = dayRows.reduce((acc, row) => {
      acc[row.key] = (acc[row.key] || 0) + 1;
      return acc;
    }, {});
    const duplicates = Object.entries(keyCounts).filter(([, count]) => count > 1);
    console.log(`[ZOOM][Diag] total courses for ${dateStr}:`, dayRows.length);
    console.log(`[ZOOM][Diag] keys generated for ${dateStr}:`, dayRows.map(r => r.key));
    console.log(`[ZOOM][Diag] duplicate keys for ${dateStr}:`, duplicates);
  });
}

async function persistZoomAssignment(dayNum, course){
  if(window.zoomReadOnlyMode) return;
  if(!canAssignZoom()){
    notifyZoomNoPermission();
    return;
  }

  const courseKey = zoomCourseId(dayNum, course);
  const assignment = window.zoomAssignments?.[courseKey] || {};
  const startTime = normalizeZoomTime(assignment.startTime || course.StartTime || '');
  const endTime = normalizeZoomTime(assignment.endTime || course.EndTime || '');
  const payload = {
    CourseId:   courseKey,
    Date:       assignment.date      || course.Date      || zoomDateString(dayNum),
    Authority:  assignment.authority || course.Authority || '',
    School:     assignment.school    || course.School    || '',
    Program:    assignment.program   || course.Program   || '',
    Employee:   assignment.employee  || course.Employee  || '',
    EmployeeID: assignment.employeeId|| String(course.EmployeeID || ''),
    StartTime:  startTime,
    EndTime:    endTime,
    ZoomAccount: assignment.account  || '',
    Notes:      assignment.notes     || ''
  };
  await saveZoomAssignment(payload);
  updateZoomAssignmentState(courseKey, {
    account: payload.ZoomAccount || null,
    notes: payload.Notes || '',
    startTime: payload.StartTime || '',
    endTime: payload.EndTime || '',
    date: payload.Date || '',
    authority: payload.Authority || '',
    school: payload.School || '',
    program: payload.Program || '',
    employee: payload.Employee || '',
    employeeId: payload.EmployeeID || ''
  });
  requestZoomCalendarRefresh();
}

async function persistZoomAssignmentsForDay(dayNum, dayCourses){
  await Promise.all(dayCourses.map(course => persistZoomAssignment(dayNum, course)));
}

async function ensureZoomAssignmentsLoaded(){
  if(window.zoomAssignmentsLoaded) return;
  const zoomAssignments = await loadZoomAssignments();
  window.zoomAssignments = mapZoomAssignmentsByCourseKey(zoomAssignments);
  window.zoomAssignmentsLoaded = true;
}
function zoomTimeToMinutes(t) {
  if (!t) return null;
  const parts = String(t).split(':');
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1] || '0', 10);
}
function zoomTimesOverlap(s1, e1, s2, e2) {
  if (s1 == null || e1 == null || s2 == null || e2 == null) return false;
  return s1 < e2 && s2 < e1;
}
function copyZoomLink(url, btn) {
  const orig = btn.textContent;
  const done = () => { btn.textContent = orig + ' ✓'; setTimeout(() => { btn.textContent = orig; }, 1600); };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(url).then(done).catch(() => { zoomFallbackCopy(url); done(); });
  } else { zoomFallbackCopy(url); done(); }
}
function zoomFallbackCopy(text) {
  const ta = Object.assign(document.createElement('textarea'), { value: text });
  document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove();
}
async function autoAssignZoomDay(dayNum, dayCourses) {
  const ACCOUNTS = ['Z1', 'Z2', 'Z3'];
  let roundRobinIndex = 0;
  const assignedSlots = [];
  const toAssign = dayCourses
    .slice()
    .sort((a, b) => {
      const aKey = zoomCourseId(dayNum, a);
      const bKey = zoomCourseId(dayNum, b);
      const aStart = window.zoomAssignments[aKey]?.startTime || a.StartTime || '';
      const bStart = window.zoomAssignments[bKey]?.startTime || b.StartTime || '';
      return aStart.localeCompare(bStart);
    });

  dayCourses.forEach(c => {
    const k = zoomCourseId(dayNum, c);
    if (window.zoomAssignments[k]) {
      window.zoomAssignments[k].account = null;
      window.zoomAssignments[k].conflict = false;
    }
  });

  function isZoomBusy(startMin, endMin, account) {
    return assignedSlots.some(slot =>
      slot.zoom === account && zoomTimesOverlap(startMin, endMin, slot.startMin, slot.endMin)
    );
  }

  function isInstructorBusy(startMin, endMin, employee) {
    return assignedSlots.some(slot =>
      slot.employee === employee && zoomTimesOverlap(startMin, endMin, slot.startMin, slot.endMin)
    );
  }

  function pickZoomForSlot(startMin, endMin, preferred, employee) {
    let ordered;
    if (preferred) {
      ordered = [preferred, ...ACCOUNTS.filter(a => a !== preferred)];
    } else {
      ordered = [...ACCOUNTS.slice(roundRobinIndex), ...ACCOUNTS.slice(0, roundRobinIndex)];
    }
    for (const acc of ordered) {
      if (!isZoomBusy(startMin, endMin, acc) && !isInstructorBusy(startMin, endMin, employee)) {
        if (!preferred) roundRobinIndex = (roundRobinIndex + 1) % ACCOUNTS.length;
        return acc;
      }
    }
    return null;
  }

  toAssign.forEach(course => {
    const key = zoomCourseId(dayNum, course);
    const assignment = window.zoomAssignments[key] || {};
    const s = zoomTimeToMinutes(assignment.startTime || course.StartTime);
    const e = zoomTimeToMinutes(assignment.endTime || course.EndTime);
    const emp = course.Employee || '';

    let preferred = null;
    assignedSlots.forEach(slot => {
      if (slot.employee === emp && slot.endMin === s) preferred = slot.zoom;
    });

    const assigned = pickZoomForSlot(s, e, preferred, emp);

    if (assigned) {
      assignedSlots.push({ startMin: s, endMin: e, employee: emp, zoom: assigned });
      window.zoomAssignments[key].account = assigned;
      window.zoomAssignments[key].conflict = false;
    } else {
      window.zoomAssignments[key].conflict = true;
    }
  });
  await persistZoomAssignmentsForDay(dayNum, dayCourses);
}
// ─── End ZOOM helpers ─────────────────────────────────────────────────────────

async function renderZoom() {
  titleEl.textContent = 'ניהול מפגשי ZOOM';
  if (window.zoomWeekPage === undefined) window.zoomWeekPage = 0;
  if (!window.zoomSubView) window.zoomSubView = 'calendar';

  const WEEK_PAGES = [
    { days: [8, 9, 10, 11, 12, 13] },
    { days: [15, 16, 17, 18, 19] },
    { days: [22, 23] },
  ];
  const ZOOM_LINKS_MAP = {
    Z1: 'https://zoom.us/j/6023602336?omn=96962875568',
    Z2: 'https://zoom.us/j/7601360450?omn=98989531483',
    Z3: 'https://zoom.us/j/97448258082',
  };
  const HDAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

  const { courses: googleCourses, assignments } = await getZoomData(false);
  window.zoomGoogleCourses = googleCourses;
  window.zoomAssignments = assignments;

  // Fall back to rawData if Google COURSES sheet is empty
  let courses = googleCourses;
  if (!courses.length) {
    courses = rawData
      .filter(r => r.Date1)
      .map(r => {
        const d = parseDate(r.Date1);
        if (!d) return null;
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return {
          Id: `raw-${r.EmployeeID || ''}-${yyyy}${mm}${dd}-${r.Program || ''}`,
          Date: `${yyyy}-${mm}-${dd}`,
          Authority: r.Authority || '',
          School: r.School || '',
          Program: r.Program || '',
          Employee: r.Employee || '',
          EmployeeID: String(r.EmployeeID || ''),
          StartTime: normalizeZoomTime(r.StartTime || ''),
          EndTime: normalizeZoomTime(r.EndTime || ''),
          Notes: r.Notes || ''
        };
      })
      .filter(Boolean);
  }
  courses = courses.map(c => ({
    ...c,
    Id: c.Id != null ? String(c.Id).trim() : c.Id,
    Date: normalizeZoomDateKey(c.Date || c.date || ''),
    StartTime: normalizeZoomTime(c.StartTime || c.startTime || ''),
    EndTime: normalizeZoomTime(c.EndTime || c.endTime || '')
  }));
  const currentPage = WEEK_PAGES[window.zoomWeekPage];
  logZoomCourseDiagnostics(courses, currentPage.days);
  window.zoomCoursesForView = courses;
  window.zoomDaysForView = currentPage.days;
  window.zoomHdaysForView = HDAYS;
  const weekRangeLabel = formatZoomWeekRange(currentPage.days);

  const wrap = document.createElement('div');
  wrap.className = 'zoom-page';

  const header = document.createElement('div');
  header.className = 'zoom-header';

  // ── Z1 / Z2 / Z3 copy buttons ──
  const linksBar = document.createElement('div');
  linksBar.className = 'zoom-links-bar';
[['Z1', 'Zoom1', 'zoom-btn-1'], ['Z2', 'Zoom2', 'zoom-btn-2'], ['Z3', 'Zoom3', 'zoom-btn-3']].forEach(([key, label, cls]) => {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'zoom-link-btn zoom-btn ' + cls;
  btn.textContent = label;
  btn.addEventListener('click', () => copyZoomLink(ZOOM_LINKS_MAP[key], btn));
  linksBar.appendChild(btn);
  });
  header.appendChild(linksBar);
  const caption = document.createElement('p');
  caption.className = 'zoom-links-caption';
  caption.textContent = 'לחיצה מעתיקה את קישור הפגישה';
  header.appendChild(caption);

  // ── Sub-nav: יומן שבועי / הכנת שיבוץ (הכנת שיבוץ רק למורשים) ──
  if (!canAssignZoom() && window.zoomSubView === 'prep') window.zoomSubView = 'calendar';
  const subNav = document.createElement('div');
  subNav.className = 'zoom-sub-nav';
  const subViews = [['calendar', 'יומן שבועי']];
  if (canAssignZoom()) subViews.push(['prep', 'הכנת שיבוץ']);
  subViews.forEach(([sv, label]) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'zoom-sub-btn zoom-tab' + (window.zoomSubView === sv ? ' active' : '');
    btn.textContent = label;
    btn.addEventListener('click', async () => {
      const dirtyBeforeSwitch = !!window.zoomCalendarDirty;
      window.zoomSubView = sv;
      await renderZoom();
      let didRefreshAfterSwitch = false;
      if(sv === 'calendar' && dirtyBeforeSwitch){
        console.log('[ZOOM][Calendar] calendar tab clicked', {
          dirtyBeforeSwitch,
          dirtyAfterRender: !!window.zoomCalendarDirty
        });
        refreshZoomCalendarView();
        didRefreshAfterSwitch = true;
      }
      if(sv === 'calendar'){
        console.log('[ZOOM][Calendar] calendar tab transition result', {
          dirtyBeforeSwitch,
          dirtyAfterRender: !!window.zoomCalendarDirty,
          didRefreshAfterSwitch
        });
      }
    });
    subNav.appendChild(btn);
  });
  const refreshBtn = document.createElement('button');
  refreshBtn.type = 'button';
  refreshBtn.className = 'zoom-sub-btn zoom-tab';
  refreshBtn.textContent = 'רענון';
  refreshBtn.addEventListener('click', async () => {
    clearZoomCache();
    await renderZoom();
  });
  subNav.appendChild(refreshBtn);
  header.appendChild(subNav);

  // ── Week navigation ──
  const weekNav = document.createElement('div');
  weekNav.className = 'zoom-week-nav';
  const prevBtn = document.createElement('button');
  prevBtn.type = 'button'; prevBtn.className = 'zoom-week-arrow'; prevBtn.textContent = '▶';
  prevBtn.disabled = window.zoomWeekPage === 0;
  prevBtn.addEventListener('click', () => { window.zoomWeekPage--; renderZoom(); });
  const weekLabel = document.createElement('span');
  weekLabel.className = 'zoom-week-label';
  weekLabel.setAttribute('dir', 'ltr');
  weekLabel.textContent = weekRangeLabel;
  const nextBtn = document.createElement('button');
  nextBtn.type = 'button'; nextBtn.className = 'zoom-week-arrow'; nextBtn.textContent = '◀';
  nextBtn.disabled = window.zoomWeekPage === WEEK_PAGES.length - 1;
  nextBtn.addEventListener('click', () => { window.zoomWeekPage++; renderZoom(); });
  weekNav.append(prevBtn, weekLabel, nextBtn);
  header.appendChild(weekNav);
  wrap.appendChild(header);

  // ── Main content ──
  const content = document.createElement('div');
  content.className = 'zoom-content';
  if (window.zoomSubView === 'calendar' || !canAssignZoom()) {
    renderZoomCalendar(content, courses, currentPage.days, HDAYS);
    window.zoomCalendarDirty = false;
  } else {
    renderZoomPrep(content, courses, currentPage.days, HDAYS);
  }
  wrap.appendChild(content);
  view.innerHTML = '';
  view.appendChild(wrap);
}

function formatZoomWeekRange(days) {
  if (!Array.isArray(days) || days.length === 0) return '';
  const sortedDays = days.slice().sort((a, b) => a - b);
  const start = new Date(2026, 2, sortedDays[0]);
  const end = new Date(2026, 2, sortedDays[sortedDays.length - 1]);
  const format = (dateObj) => {
    const dd = String(dateObj.getDate()).padStart(2, '0');
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const yy = String(dateObj.getFullYear()).slice(-2);
    return `${dd}.${mm}.${yy}`;
  };
  return `\u200E${format(start)} - ${format(end)}\u200E`;
}

function zoomPastelColor(instructorName) {
  const [r, g, b] = toRgbTuple(instructorColor(instructorName || ''))
    .split(',')
    .map(v => Number(v.trim()));
  const blend = (value) => Math.round((value * 0.35) + (255 * 0.65));
  return `rgb(${blend(r)}, ${blend(g)}, ${blend(b)})`;
}

function computeZoomOverlapLayout(items) {
  const sorted = items.slice().sort((a, b) => (a.startMin - b.startMin) || (a.endMin - b.endMin));
  const groups = [];
  let currentGroup = null;

  sorted.forEach(item => {
    if (!currentGroup || item.startMin >= currentGroup.endMax) {
      currentGroup = { endMax: item.endMin, items: [item] };
      groups.push(currentGroup);
      return;
    }
    currentGroup.items.push(item);
    currentGroup.endMax = Math.max(currentGroup.endMax, item.endMin);
  });

  groups.forEach(group => {
    const active = [];
    let colCount = 0;
    const groupItems = group.items.sort((a, b) => (a.startMin - b.startMin) || (a.endMin - b.endMin));
    groupItems.forEach(item => {
      for (let i = active.length - 1; i >= 0; i--) {
        if (active[i].endMin <= item.startMin) active.splice(i, 1);
      }

      const usedCols = new Set(active.map(ev => ev._overlapIndex));
      let colIndex = 0;
      while (usedCols.has(colIndex)) colIndex++;

      item._overlapIndex = colIndex;
      active.push(item);
      colCount = Math.max(colCount, active.length);
    });

    groupItems.forEach(item => {
      item._overlapCount = colCount || 1;
    });
  });

  return sorted;
}

function renderZoomCalendar(container, courses, days, hdays) {
  const DAY_START  = 8 * 60;
  const DAY_END    = 16 * 60;
  const displayDays = days.slice().reverse(); // ראשון מופיע בצד ימין
  const ZOOM_ACCOUNT_ORDER = { Z1: 1, Z2: 2, Z3: 3 };

  // Gather assigned items for this week
  const assignedItems = [];
  console.log('[ZOOM][Calendar] renderZoomCalendar input', {
    courses: Array.isArray(courses) ? courses.length : 0,
    assignments: window.zoomAssignments ? Object.keys(window.zoomAssignments).length : 0,
    days: Array.isArray(days) ? days.length : 0
  });
  displayDays.forEach(dayNum => {
    const dateStr = zoomDateString(dayNum);
    const dayCourseCount = courses.filter(course => normalizeZoomDateKey(course.Date || course.date || '') === dateStr).length;
    console.log(`[ZOOM][Calendar] day ${dateStr} courses before assignment filter:`, dayCourseCount);
    const renderedKeys = new Set();
    let dayRendered = 0;
    courses.forEach(course => {
      const key = zoomCourseId(dayNum, course);
      const asgn = window.zoomAssignments[key];
      const effectiveDate = normalizeZoomDateKey(asgn?.date || course.Date || course.date || '');
      if(effectiveDate !== dateStr) return;
      if (asgn && asgn.account) {
        const startTime = normalizeZoomTime(asgn.startTime || course.StartTime || '');
        const endTime = normalizeZoomTime(asgn.endTime || course.EndTime || '');
        const startMin = zoomTimeToMinutes(startTime);
        const endMin   = zoomTimeToMinutes(endTime);
        assignedItems.push({
          dayNum, course, account: asgn.account,
          courseId: key,
          startTime,
          endTime,
          startMin,
          endMin,
        });
        renderedKeys.add(key);
        dayRendered += 1;
      }
    });

    Object.entries(window.zoomAssignments || {}).forEach(([key, asgn]) => {
      if(!asgn || !asgn.account || renderedKeys.has(key)) return;
      const asgnDate = normalizeZoomDateKey(asgn.date || '');
      if(asgnDate !== dateStr) return;
      const startTime = normalizeZoomTime(asgn.startTime || '');
      const endTime = normalizeZoomTime(asgn.endTime || '');
      const startMin = zoomTimeToMinutes(startTime);
      const endMin = zoomTimeToMinutes(endTime);
      assignedItems.push({
        dayNum,
        course: {
          Program: asgn.program || '',
          Authority: asgn.authority || '',
          School: asgn.school || '',
          Employee: asgn.employee || ''
        },
        account: asgn.account,
        courseId: key,
        startTime,
        endTime,
        startMin,
        endMin,
      });
      dayRendered += 1;
    });
    console.log(`[ZOOM][Calendar] day ${dateStr} rows rendered (assigned events):`, dayRendered);
  });
  console.log('[ZOOM][Calendar] assignedItems before build:', assignedItems.length);

  const grid = document.createElement('div');
  grid.className = 'zoom-cal-grid';

  const calendarContainer = document.createElement('div');
  calendarContainer.className = 'zoom-calendar-container';

  // Header
  const header = document.createElement('div');
  header.className = 'zoom-cal-header';
  const timeHdr = document.createElement('div');
  timeHdr.className = 'zoom-cal-timecol-header';
  header.appendChild(timeHdr);
  displayDays.forEach(dayNum => {
    const date = new Date(2026, 2, dayNum);
    const cell = document.createElement('div');
    cell.className = 'zoom-cal-day-header';
    cell.innerHTML =
      `<strong>יום ${hdays[date.getDay()]}</strong>` +
      `<span>${String(dayNum).padStart(2, '0')}.03</span>`;
    header.appendChild(cell);
  });
  grid.appendChild(header);

  // Body
  const body = document.createElement('div');
  body.className = 'zoom-cal-body';

  // Time column
  const timeCol = document.createElement('div');
  timeCol.className = 'zoom-cal-timecol';
  for (let h = 8; h <= 16; h++) {
    const lbl = document.createElement('div');
    lbl.className = 'zoom-cal-timeslot-label';
    lbl.textContent = String(h).padStart(2, '0') + ':00';
    timeCol.appendChild(lbl);
  }
  body.appendChild(timeCol);

  // Day columns
  displayDays.forEach(dayNum => {
    const dayItems = assignedItems.filter(item => item.dayNum === dayNum);
    const col = document.createElement('div');
    col.className = 'zoom-cal-daycol';
    for (let h = 8; h < 16; h++) {
      const slot = document.createElement('div');
      slot.className = 'zoom-cal-slot';

      const slotStart = h * 60;
      const slotEnd = slotStart + 60;
      const slotItems = dayItems
        .filter(({ startMin, endMin }) => {
          const s = Math.max(startMin != null ? startMin : DAY_START, DAY_START);
          const e = Math.min(endMin != null ? endMin : s + 120, DAY_END);
          return s < slotEnd && e > slotStart;
        })
        .sort((a, b) => {
          const accCmp = (ZOOM_ACCOUNT_ORDER[a.account] || 99) - (ZOOM_ACCOUNT_ORDER[b.account] || 99);
          if (accCmp !== 0) return accCmp;
          return (a.course.Employee || '').localeCompare(b.course.Employee || '', 'he');
        });

      if (slotItems.length >= 3) slot.classList.add('zoom-cal-slot--crowded');
      if (slotItems.length) {
        const slotList = document.createElement('div');
        slotList.className = 'zoom-slot';
        slotItems.forEach(({ account, course, courseId, startTime, endTime }) => {
          const line = document.createElement('div');
          const accountKey = String(account || '').toLowerCase();
          line.className = `zoom-line zoom-${accountKey}`;
          line.title = [
            `קורס: ${course.Program || ''}`,
            `רשות: ${course.Authority || ''}`,
            `בית ספר: ${course.School || ''}`,
            `מדריך: ${course.Employee || ''}`,
            `שעה: ${startTime}${endTime ? '–' + endTime : ''}`,
            `CourseId: ${courseId}`
          ].join('\n');
          line.innerHTML = `<strong>${escapeHtml(account || '')}</strong><span>${escapeHtml(course.Employee || '')}</span>`;
          slotList.appendChild(line);
        });
        slot.appendChild(slotList);
      }

      col.appendChild(slot);
    }
    body.appendChild(col);
  });
  grid.appendChild(body);
  calendarContainer.appendChild(grid);
  container.appendChild(calendarContainer);

  if (assignedItems.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'zoom-empty';
    empty.textContent = 'אין שיבוצי ZOOM לשבוע זה. עברו למסך הכנת שיבוץ כדי לבצע שיבוץ יומי.';
    container.appendChild(empty);
  }
}

function addNewZoomRow(dayNum, tbody, defaultDate) {
  const key = 'new-' + dayNum + '-' + Date.now();
  const course = { Id: key, Date: defaultDate || zoomDateString(dayNum) };
  updateZoomAssignmentState(key, {
    account: null, notes: '', conflict: false,
    startTime: '', endTime: '',
    date: defaultDate || zoomDateString(dayNum),
    authority: '', school: '', program: '', employee: '', employeeId: ''
  });
  requestZoomCalendarRefresh();
  const asgn = window.zoomAssignments[key];

  const tr = document.createElement('tr');

  const tdCheck = document.createElement('td');
  tdCheck.style.textAlign = 'center';
  const chk = document.createElement('input'); chk.type = 'checkbox'; chk.className = 'zoom-row-select';
  tdCheck.appendChild(chk); tr.appendChild(tdCheck);

  const tdZoom = document.createElement('td');
  tdZoom.style.textAlign = 'center'; tdZoom.style.whiteSpace = 'nowrap';
  const badge = document.createElement('span'); badge.className = 'zoom-account-badge'; badge.textContent = '';
  tdZoom.appendChild(badge); tr.appendChild(tdZoom);

  const tdDate = document.createElement('td'); tdDate.setAttribute('data-label', 'תאריך');
  const dateInp = document.createElement('input');
  dateInp.type = 'date'; dateInp.className = 'zoom-field-input'; dateInp.dir = 'ltr';
  dateInp.value = asgn.date;
  dateInp.addEventListener('change', async () => { asgn.date = dateInp.value; await persistZoomAssignment(dayNum, course); });
  tdDate.appendChild(dateInp); tr.appendChild(tdDate);

  const tdAuth = document.createElement('td'); tdAuth.setAttribute('data-label', 'רשות');
  const authInp = document.createElement('input');
  authInp.type = 'text'; authInp.className = 'zoom-field-input'; authInp.dir = 'rtl';
  authInp.addEventListener('input', async () => { asgn.authority = authInp.value; await persistZoomAssignment(dayNum, course); });
  tdAuth.appendChild(authInp); tr.appendChild(tdAuth);

  const tdSchool = document.createElement('td'); tdSchool.setAttribute('data-label', 'בית ספר');
  const schoolInp = document.createElement('input');
  schoolInp.type = 'text'; schoolInp.className = 'zoom-field-input'; schoolInp.dir = 'rtl';
  schoolInp.addEventListener('input', async () => { asgn.school = schoolInp.value; await persistZoomAssignment(dayNum, course); });
  tdSchool.appendChild(schoolInp); tr.appendChild(tdSchool);

  const tdProg = document.createElement('td'); tdProg.setAttribute('data-label', 'קורס');
  const progSelect = document.createElement('select'); progSelect.className = 'zoom-emp-select'; progSelect.dir = 'rtl';
  const blankProgOpt2 = document.createElement('option'); blankProgOpt2.value = ''; blankProgOpt2.textContent = '— בחר —';
  progSelect.appendChild(blankProgOpt2);
  ZOOM_PROGRAMS.forEach(name => {
    const opt = document.createElement('option'); opt.value = name; opt.textContent = name;
    progSelect.appendChild(opt);
  });
  progSelect.addEventListener('change', async () => { asgn.program = progSelect.value; await persistZoomAssignment(dayNum, course); });
  tdProg.appendChild(progSelect); tr.appendChild(tdProg);

  const tdEmp = document.createElement('td'); tdEmp.setAttribute('data-label', 'מדריך');
  const empSelect = document.createElement('select'); empSelect.className = 'zoom-emp-select'; empSelect.dir = 'rtl';
  const employeeMap = {};
  rawData.forEach(r => { if (r.Employee && r.EmployeeID) employeeMap[r.Employee] = String(r.EmployeeID); });
  const blankOpt = document.createElement('option'); blankOpt.value = ''; blankOpt.textContent = '— בחר —';
  empSelect.appendChild(blankOpt);
  Object.entries(employeeMap).sort(([a], [b]) => a.localeCompare(b, 'he')).forEach(([name, id]) => {
    const opt = document.createElement('option'); opt.value = name; opt.dataset.id = id; opt.textContent = name;
    empSelect.appendChild(opt);
  });
  empSelect.addEventListener('change', async () => {
    const sel = empSelect.selectedOptions[0];
    asgn.employee = sel.value; asgn.employeeId = sel.dataset.id || '';
    await persistZoomAssignment(dayNum, course);
  });
  tdEmp.appendChild(empSelect); tr.appendChild(tdEmp);

  const tdStart = document.createElement('td'); tdStart.className = 'zoom-col-start'; tdStart.setAttribute('data-label', 'התחלה');
  const startInput = createHourSelect('');
  const tdEnd = document.createElement('td'); tdEnd.className = 'zoom-col-end'; tdEnd.setAttribute('data-label', 'סיום');
  const endInput = createHourSelect('');
  startInput.addEventListener('change', async () => {
    const h = parseInt(startInput.value.split(':')[0], 10);
    if (!Number.isNaN(h)) endInput.value = String(Math.min(h + 1, 18)).padStart(2, '0') + ':00';
    asgn.startTime = startInput.value; asgn.endTime = endInput.value;
    await persistZoomAssignment(dayNum, course);
  });
  endInput.addEventListener('change', async () => { asgn.endTime = endInput.value; await persistZoomAssignment(dayNum, course); });
  tdStart.appendChild(startInput); tr.appendChild(tdStart);
  tdEnd.appendChild(endInput); tr.appendChild(tdEnd);

  const tdNotes = document.createElement('td'); tdNotes.className = 'zoom-col-notes'; tdNotes.setAttribute('data-label', 'הערות');
  const notesInp = document.createElement('input');
  notesInp.type = 'text'; notesInp.className = 'zoom-notes-input'; notesInp.dir = 'rtl'; notesInp.placeholder = 'הערה...';
  notesInp.addEventListener('input', async () => { asgn.notes = notesInp.value; await persistZoomAssignment(dayNum, course); });
  tdNotes.appendChild(notesInp); tr.appendChild(tdNotes);

  tbody.appendChild(tr);
  dateInp.focus();
}

function renderZoomPrep(container, courses, days, hdays) {
  const area = document.createElement('div');
  area.className = 'zoom-days-area';
  const employeeMap = getZoomEmployeeMap();
  const sortedEmployees = Object.entries(employeeMap).sort(([a], [b]) => a.localeCompare(b, 'he'));
  console.log('[ZOOM][Prep] input courses:', Array.isArray(courses) ? courses.length : 0);
  console.log('[ZOOM][Prep] input assignments:', window.zoomAssignments ? Object.keys(window.zoomAssignments).length : 0);

  let prepRowsTotal = 0;
  let prepRowsRendered = 0;

  days.forEach(dayNum => {
    const date = new Date(2026, 2, dayNum);
    const displayDate = String(dayNum).padStart(2, '0') + '.03.26';
    const dateStr = zoomDateString(dayNum);
    const dayCourses = courses
      .filter(c => normalizeZoomDateKey(c.Date || c.date || '') === dateStr)
      .slice()
      .sort((a, b) => normalizeZoomTime(a.StartTime || a.startTime || '').localeCompare(normalizeZoomTime(b.StartTime || b.startTime || '')));
    prepRowsTotal += dayCourses.length;
    console.log(`[ZOOM][Prep] day ${dateStr} courses after day filter:`, dayCourses.length);

    const card = document.createElement('div');
    card.className = 'zoom-day-card';

    const titleDiv = document.createElement('div');
    titleDiv.className = 'zoom-day-title';
    titleDiv.textContent = 'יום ' + hdays[date.getDay()] + ' – ' + displayDate;
    card.appendChild(titleDiv);

    const tableWrap = document.createElement('div');
    tableWrap.className = 'zoom-table-wrap';
    const table = document.createElement('table');
    table.className = 'zoom-table';
    const thead = document.createElement('thead');
    thead.innerHTML =
      '<tr>' +
      '<th style="width:2rem"><input type="checkbox" class="zoom-select-all" title="סמן הכל"></th>' +
      '<th>ZOOM</th>' +
      '<th>תאריך</th>' +
      '<th>רשות</th>' +
      '<th>בית ספר</th>' +
      '<th>קורס</th>' +
      '<th>מדריך</th>' +
      '<th class="zoom-col-start">התחלה</th>' +
      '<th class="zoom-col-end">סיום</th>' +
      '<th>הערות</th>' +
      '</tr>';
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    const rowByCourseKey = {};
    const dayKeyCounts = {};
    dayCourses.forEach(course => {
      const key = zoomCourseId(dayNum, course);
      dayKeyCounts[key] = (dayKeyCounts[key] || 0) + 1;
      if (!window.zoomAssignments[key]) {
        updateZoomAssignmentState(key, { account: null, notes: '', conflict: false });
      }
      const asgn = window.zoomAssignments[key];
      asgn.startTime = normalizeZoomTime(asgn.startTime || course.StartTime || '');
      asgn.endTime = normalizeZoomTime(asgn.endTime || course.EndTime || '');

      const tr = document.createElement('tr');
      tr.dataset.zoomCourseKey = key;
      if(!rowByCourseKey[key]) rowByCourseKey[key] = [];
      rowByCourseKey[key].push(tr);
      if (asgn.account)   tr.classList.add('zoom-assigned-row');
      if (asgn.conflict)  tr.classList.add('zoom-conflict-row');

      // Checkbox cell
      const tdCheck = document.createElement('td');
      tdCheck.style.textAlign = 'center';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.className = 'zoom-row-select';
      tdCheck.appendChild(chk);
      tr.appendChild(tdCheck);

      // ZOOM cell: account badge
      const tdZoom = document.createElement('td');
      tdZoom.style.textAlign = 'center';
      tdZoom.style.whiteSpace = 'nowrap';
      const badge = document.createElement('span');
      badge.style.marginRight = '5px';
      badge.className = 'zoom-account-badge' + (asgn.account ? ' zoom-account-badge-' + asgn.account.toLowerCase() : '');
      badge.textContent = asgn.account || '';
      tdZoom.append(badge);
      tr.appendChild(tdZoom);

      // Date cell
      const tdDate = document.createElement('td');
      tdDate.setAttribute('data-label', 'תאריך');
      const dateInp = document.createElement('input');
      dateInp.type = 'date'; dateInp.className = 'zoom-field-input'; dateInp.dir = 'ltr';
      dateInp.value = asgn.date || course.Date || '';
      dateInp.addEventListener('change', async () => {
        asgn.date = dateInp.value;
        await persistZoomAssignment(dayNum, course);
      });
      tdDate.appendChild(dateInp);
      tr.appendChild(tdDate);

      // Authority cell
      const tdAuth = document.createElement('td');
      tdAuth.setAttribute('data-label', 'רשות');
      const authInp = document.createElement('input');
      authInp.type = 'text'; authInp.className = 'zoom-field-input'; authInp.dir = 'rtl';
      authInp.value = asgn.authority || course.Authority || '';
      authInp.addEventListener('input', async () => {
        asgn.authority = authInp.value;
        await persistZoomAssignment(dayNum, course);
      });
      tdAuth.appendChild(authInp);
      tr.appendChild(tdAuth);

      // School cell
      const tdSchool = document.createElement('td');
      tdSchool.setAttribute('data-label', 'בית ספר');
      const schoolInp = document.createElement('input');
      schoolInp.type = 'text'; schoolInp.className = 'zoom-field-input'; schoolInp.dir = 'rtl';
      schoolInp.value = asgn.school || course.School || '';
      schoolInp.addEventListener('input', async () => {
        asgn.school = schoolInp.value;
        await persistZoomAssignment(dayNum, course);
      });
      tdSchool.appendChild(schoolInp);
      tr.appendChild(tdSchool);

      // Program cell
      const tdProg = document.createElement('td');
      tdProg.setAttribute('data-label', 'קורס');
      const progSelect = document.createElement('select');
      progSelect.className = 'zoom-emp-select'; progSelect.dir = 'rtl';
      const currentProg = asgn.program || course.Program || '';
      const blankProgOpt = document.createElement('option'); blankProgOpt.value = ''; blankProgOpt.textContent = '— בחר —';
      progSelect.appendChild(blankProgOpt);
      if (!ZOOM_PROGRAMS.includes(currentProg) && currentProg) {
        const opt = document.createElement('option'); opt.value = currentProg; opt.textContent = currentProg; opt.selected = true;
        progSelect.appendChild(opt);
      }
      ZOOM_PROGRAMS.forEach(name => {
        const opt = document.createElement('option'); opt.value = name; opt.textContent = name;
        if (name === currentProg) opt.selected = true;
        progSelect.appendChild(opt);
      });
      progSelect.addEventListener('change', async () => {
        asgn.program = progSelect.value;
        await persistZoomAssignment(dayNum, course);
      });
      tdProg.appendChild(progSelect);
      tr.appendChild(tdProg);

      // Employee dropdown cell
      const tdEmp = document.createElement('td');
      tdEmp.setAttribute('data-label', 'מדריך');
      const empSelect = document.createElement('select');
      empSelect.className = 'zoom-emp-select'; empSelect.dir = 'rtl';
      const currentEmp = asgn.employee || course.Employee || '';
      if (!employeeMap[currentEmp] && currentEmp) {
        const blankOpt = document.createElement('option');
        blankOpt.value = currentEmp; blankOpt.dataset.id = asgn.employeeId || String(course.EmployeeID || '');
        blankOpt.textContent = currentEmp; blankOpt.selected = true;
        empSelect.appendChild(blankOpt);
      }
      sortedEmployees.forEach(([name, id]) => {
        const opt = document.createElement('option');
        opt.value = name; opt.dataset.id = id; opt.textContent = name;
        if (name === currentEmp) opt.selected = true;
        empSelect.appendChild(opt);
      });
      empSelect.addEventListener('change', async () => {
        const sel = empSelect.selectedOptions[0];
        asgn.employee = sel.value;
        asgn.employeeId = sel.dataset.id || '';
        await persistZoomAssignment(dayNum, course);
      });
      tdEmp.appendChild(empSelect);
      tr.appendChild(tdEmp);

      const tdStart = document.createElement('td');
      tdStart.className = 'zoom-col-start';
      tdStart.setAttribute('data-label', 'התחלה');
      const startInput = createHourSelect(normalizeZoomTime(asgn.startTime || course.StartTime || ''));

      const tdEnd = document.createElement('td');
      tdEnd.className = 'zoom-col-end';
      tdEnd.setAttribute('data-label', 'סיום');
      const endInput = createHourSelect(normalizeZoomTime(asgn.endTime || course.EndTime || ''));

      if (!asgn.endTime && startInput.value) {
        const startHour = parseInt(startInput.value.split(':')[0], 10);
        if (!Number.isNaN(startHour)) {
          const endHour = Math.min(startHour + 1, 18);
          endInput.value = String(endHour).padStart(2, '0') + ':00';
        }
      }

      asgn.startTime = startInput.value;
      asgn.endTime = endInput.value;

      startInput.addEventListener('change', async () => {
        const startHour = parseInt(startInput.value.split(':')[0], 10);
        if (!Number.isNaN(startHour)) {
          const endHour = Math.min(startHour + 1, 18);
          endInput.value = String(endHour).padStart(2, '0') + ':00';
        }
        asgn.startTime = startInput.value;
        asgn.endTime = endInput.value;
        await persistZoomAssignment(dayNum, course);
      });

      endInput.addEventListener('change', async () => {
        asgn.endTime = endInput.value;
        await persistZoomAssignment(dayNum, course);
      });

      tdStart.appendChild(startInput);
      tdEnd.appendChild(endInput);
      tr.appendChild(tdStart);
      tr.appendChild(tdEnd);

      // Notes
      const tdNotes = document.createElement('td');
      tdNotes.className = 'zoom-col-notes';
      tdNotes.setAttribute('data-label', 'הערות');
      const inp = document.createElement('input');
      inp.type = 'text'; inp.className = 'zoom-notes-input'; inp.dir = 'rtl';
      inp.placeholder = 'הערה...'; inp.value = asgn.notes || course.Notes || '';
      if (!asgn.notes && course.Notes) asgn.notes = course.Notes;
      inp.addEventListener('input', async () => {
        window.zoomAssignments[key].notes = inp.value;
        await persistZoomAssignment(dayNum, course);
      });
      tdNotes.appendChild(inp);
      tr.appendChild(tdNotes);
      tbody.appendChild(tr);
      prepRowsRendered += 1;
    });
    const duplicateKeys = Object.entries(dayKeyCounts).filter(([, count]) => count > 1);
    console.log(`[ZOOM][Prep] day ${dateStr} duplicate keys:`, duplicateKeys);
    console.log(`[ZOOM][Prep] day ${dateStr} rows rendered to DOM:`, tbody.querySelectorAll('tr').length);

    table.appendChild(tbody);
    tableWrap.appendChild(table);
    card.appendChild(tableWrap);

    // Wire select-all checkbox
    const selectAllChk = thead.querySelector('.zoom-select-all');
    if (selectAllChk) {
      selectAllChk.addEventListener('change', () => {
        tbody.querySelectorAll('.zoom-row-select').forEach(c => { c.checked = selectAllChk.checked; });
      });
      tbody.querySelectorAll('.zoom-row-select').forEach(c => {
        c.addEventListener('change', () => {
          const all  = Array.from(tbody.querySelectorAll('.zoom-row-select'));
          selectAllChk.checked      = all.every(x => x.checked);
          selectAllChk.indeterminate = !selectAllChk.checked && all.some(x => x.checked);
        });
      });
    }

    // Assign button
    const assignBtn = document.createElement('button');
    assignBtn.type = 'button';
    assignBtn.className = 'zoom-assign-btn';
    assignBtn.textContent = 'שיבוץ';
    assignBtn.addEventListener('click', async () => {
      if (!canAssignZoom()) { notifyZoomNoPermission(); return; }
      // Only assign checked rows
      const rowCheckboxes = Array.from(tbody.querySelectorAll('.zoom-row-select'));
      const selectedCourses = dayCourses.filter((_, i) => rowCheckboxes[i] && rowCheckboxes[i].checked);

      if (!selectedCourses.length) {
        alert('יש לסמן לפחות קורס אחד לשיבוץ');
        return;
      }

      assignBtn.disabled = true;
      assignBtn.textContent = 'מבצע שיבוץ...';

      // Ensure selected courses have an entry in window.zoomAssignments
      selectedCourses.forEach(c => {
        const k = zoomCourseId(dayNum, c);
        if (!window.zoomAssignments[k]) {
          updateZoomAssignmentState(k, { account: null, notes: '', startTime: c.StartTime || '', endTime: c.EndTime || '', conflict: false });
        } else {
          if (!window.zoomAssignments[k].startTime) updateZoomAssignmentState(k, { startTime: c.StartTime || '' });
          if (!window.zoomAssignments[k].endTime)   updateZoomAssignmentState(k, { endTime: c.EndTime || '' });
        }
      });

      // Perform assignment only for selected courses and save to Google Sheets
      await autoAssignZoomDay(dayNum, selectedCourses);
      requestZoomCalendarRefresh();

      selectedCourses.forEach(c => {
        const k = zoomCourseId(dayNum, c);
        const rows = rowByCourseKey[k] || [];
        if(!rows.length) return;
        const asgn = window.zoomAssignments[k] || {};
        rows.forEach(row => {
          row.classList.toggle('zoom-assigned-row', !!asgn.account);
          row.classList.toggle('zoom-conflict-row', !!asgn.conflict);
          const rowBadge = row.querySelector('.zoom-account-badge');
          if(rowBadge){
            rowBadge.className = 'zoom-account-badge' + (asgn.account ? ` zoom-account-badge-${String(asgn.account).toLowerCase()}` : '');
            rowBadge.textContent = asgn.account || '';
          }
        });
      });

      assignBtn.textContent = 'שיבוץ';
      assignBtn.disabled = false;
      alert('השיבוץ הושלם');
    });
    const newRowBtn = document.createElement('button');
    newRowBtn.type = 'button';
    newRowBtn.className = 'zoom-new-row-btn';
    newRowBtn.textContent = 'חדש';
    newRowBtn.addEventListener('click', () => addNewZoomRow(dayNum, tbody, dateStr));

    const btnRow = document.createElement('div');
    btnRow.className = 'zoom-btn-row';
    btnRow.appendChild(assignBtn);
    btnRow.appendChild(newRowBtn);
    card.appendChild(btnRow);
    area.appendChild(card);
  });

  console.log('[ZOOM][Prep] rows before filter:', prepRowsTotal);
  console.log('[ZOOM][Prep] rows after filter:', prepRowsRendered);

  container.appendChild(area);
}

function renderEndDates(){
  titleEl.textContent = 'תאריכי סיום קורסים';

  const courses = rawData
    .filter(r => isCourse(r) && r.End instanceof Date && !isNaN(r.End.getTime()))
    .filter(r => r.End.getMonth() >= 0 && r.End.getMonth() <= 5)
    .slice()
    .sort((a, b) => a.End - b.End);

  const groupedByMonth = new Map();
  courses.forEach(course => {
    const endDate = course.End;
    const key = `${endDate.getFullYear()}-${endDate.getMonth()}`;
    if(!groupedByMonth.has(key)){
      groupedByMonth.set(key, {
        monthLabel: endDate.toLocaleDateString('he-IL', { month: 'long', year: 'numeric' }),
        courses: []
      });
    }
    groupedByMonth.get(key).courses.push(course);
  });

  const monthSections = [...groupedByMonth.values()].map((group, groupIndex) => {
    const rows = group.courses
      .slice()
      .sort((a, b) => a.End - b.End)
      .map(course => {
        const endDateText = course.End.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const courseIndex = courses.indexOf(course);
        const searchText = [endDateText, course.School, course.Authority, course.Program]
          .map(s => (s || '').toLowerCase())
          .join(' ');
        return `
          <tr class="end-courses-row" data-course-index="${courseIndex}" data-search="${escapeHtml(searchText)}">
            <td class="col-end-date" data-label="תאריך סיום">${escapeHtml(endDateText)}</td>
            <td class="col-school" data-label="בית ספר">${escapeHtml(course.School || '—')}</td>
            <td class="col-authority" data-label="רשות">${escapeHtml(course.Authority || '—')}</td>
            <td class="col-course" data-label="קורס">${escapeHtml(course.Program || '—')}</td>
          </tr>
        `;
      })
      .join('');

    const monthTitleId = `endCoursesMonthTitle-${groupIndex}`;

    return `
      <section class="end-courses-month-group">
        <div id="${monthTitleId}" class="end-courses-month-title">${escapeHtml(group.monthLabel)}</div>
        <div class="end-courses-table-wrap table-container" role="region" aria-labelledby="${monthTitleId}" tabindex="0">
          <table class="end-courses-table">
            <colgroup>
              <col class="col-end-date">
              <col class="col-school">
              <col class="col-authority">
              <col class="col-course">
            </colgroup>
            <thead>
              <tr>
                <th>תאריך סיום</th>
                <th>בית ספר</th>
                <th>רשות</th>
                <th>קורס</th>
              </tr>
            </thead>
            <tbody>
              ${rows || '<tr><td colspan="4" class="end-courses-empty" data-label="מצב">אין קורסים לחודש זה</td></tr>'}
            </tbody>
          </table>
        </div>
      </section>
    `;
  }).join('');

  view.innerHTML = `
    <div class="end-courses-page">
      <h2 class="end-courses-page-title">תאריכי סיום קורסים</h2>
      <div class="end-courses-search-wrap">
        <input
          id="endCoursesSearch"
          class="end-courses-search-input"
          type="search"
          dir="rtl"
          placeholder="חיפוש בטבלה..."
          aria-label="חיפוש בטבלת תאריכי סיום"
        />
      </div>
      <div class="end-courses-content">
        ${monthSections || '<div class="end-courses-empty">אין קורסים עם תאריך סיום</div>'}
      </div>
    </div>
  `;

  const searchInput = view.querySelector('#endCoursesSearch');
  if(searchInput){
    searchInput.addEventListener('input', () => {
      const query = searchInput.value.trim().toLowerCase();
      view.querySelectorAll('.end-courses-month-group').forEach(monthGroup => {
        let hasVisibleRows = false;
        monthGroup.querySelectorAll('.end-courses-table tbody tr.end-courses-row').forEach(row => {
          const rowText = (row.dataset.search || row.textContent).toLowerCase();
          const isVisible = !query || rowText.includes(query);
          row.style.display = isVisible ? '' : 'none';
          if(isVisible) hasVisibleRows = true;
        });
        monthGroup.style.display = hasVisibleRows ? '' : 'none';
      });
    });
  }

  view.querySelectorAll('.end-courses-row').forEach(row => {
    row.addEventListener('click', () => {
      const index = Number(row.dataset.courseIndex);
      const course = courses[index];
      if(course) openEndDateDetail(course);
    });
  });
}

function openEndDateDetail(course){
  const dates = (course.Dates || [])
    .filter(d => d instanceof Date && !isNaN(d.getTime()))
    .sort((a, b) => a - b);

  const endDateText = course.End instanceof Date && !isNaN(course.End.getTime())
    ? course.End.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' })
    : '—';

  const dateItems = dates.map((d, i) => {
    const dateStr = d.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
    return `<li class="end-detail-meeting-item"><span>${dateStr}</span><span>מפגש ${i + 1}</span></li>`;
  }).join('');

  sideContent.innerHTML = `
    <div class="end-detail-panel">
      <div class="end-detail-title">${escapeHtml(course.Program || '—')}</div>
      <div class="end-detail-meta-grid">
        <div class="end-detail-meta-item"><span>בית ספר</span><strong>${escapeHtml(course.School || '—')}</strong></div>
        <div class="end-detail-meta-item"><span>רשות</span><strong>${escapeHtml(course.Authority || '—')}</strong></div>
        <div class="end-detail-meta-item"><span>תאריך סיום</span><strong>${escapeHtml(endDateText)}</strong></div>
      </div>
      <div class="end-detail-section-title">מפגשים (${dates.length})</div>
      <ul class="end-detail-meetings-list">${dateItems || '<li class="end-detail-empty">אין תאריכים</li>'}</ul>
    </div>
  `;
  openSidePanel();
}

goCalendar.onclick = ()=>{
  window.mode = 'month';
  currentDate = clampDateToDataRange(new Date());
  render();
};
document.getElementById('goToday').onclick = ()=>{
  const today = new Date();
  today.setHours(0,0,0,0);

  currentDate = clampDateToDataRange(today);

  if(window.mode === 'summary'){
    const key = `${today.getFullYear()}-${today.getMonth()}`;
    const options = [...summaryMonth.options].map(o=>o.value);
    const index = options.indexOf(key);
    if(index >= 0){
      summaryMonth.selectedIndex = index;
  }
  }

  if(window.mode === 'instructors'){
    renderInstructors.selectedMonthValue =
      `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}`;
  }

  render();
};
managerFilter.onchange=render;
employeeFilter.onchange=render;
document.getElementById('clearFilters').onclick=()=>{managerFilter.value='';employeeFilter.value='';render();};
summaryMonth.onchange=render;

const sideBackdrop = document.getElementById('side-backdrop');

function isManagerOverlayOpen(){
  return !!document.querySelector('.manager-overlay-bg, .manager-details-overlay');
}

function syncBodyScrollLock(){
  const shouldLock = (!!side && side.classList.contains('open')) ||
                     (!!daySheet && !daySheet.classList.contains('day-sheet-hidden')) ||
                     isManagerOverlayOpen();
  if(shouldLock){
    if(view.dataset.savedScroll === undefined){
      view.dataset.savedScroll = view.scrollTop;
    }
    view.style.overflow = 'hidden';
  } else {
    const saved = view.dataset.savedScroll;
    view.style.overflow = '';
    if(saved !== undefined){
      view.scrollTop = parseInt(saved, 10);
      delete view.dataset.savedScroll;
    }
  }
}

function closeManagerOverlay(){
  document.querySelectorAll('.manager-overlay-bg, .manager-details-overlay').forEach(el => el.remove());
  syncBodyScrollLock();
}

function closeSidePanel(){
  console.log('סגירת פאנל צדדי');
  const side = document.getElementById('side');
  if(!side) return;
  side.classList.remove('open');
  if(sideBackdrop){
    sideBackdrop.classList.remove('active');
  }
  activeSidePanelType = '';
  syncBodyScrollLock();
}

function closeDaySheet(){
  if(!daySheet || !daySheetBackdrop) return;
  daySheet.classList.add('day-sheet-hidden');
  daySheetBackdrop.classList.add('day-sheet-hidden');
  syncBodyScrollLock();
}

function closeAllOverlays(){
  closeSidePanel();
  closeDaySheet();
  closeManagerOverlay();
}

function openSidePanel(){
  const side = document.getElementById('side');
  if(!side) return;
  closeAllOverlays();
  sideContent.scrollTop = 0;
  side.classList.add('open');
  if(isMobile() && sideBackdrop){
    sideBackdrop.classList.add('active');
  }
  syncBodyScrollLock();
}

function openDaySheet(title, htmlContent){
  if(!daySheet || !daySheetBackdrop || !daySheetContent) return;
  closeAllOverlays();
  daySheetTitle.textContent = title || 'פרטי יום';
  daySheetContent.innerHTML = htmlContent || '';
  daySheetContent.scrollTop = 0;
  daySheet.classList.remove('day-sheet-hidden');
  daySheetBackdrop.classList.remove('day-sheet-hidden');

  syncBodyScrollLock();
  applyNotesBoxColor();
}

document.addEventListener('click', function (e) {
  if (e.target.closest('#closeSide')) {
    e.preventDefault();
    e.stopPropagation();
    closeSidePanel();
  }
});

document.addEventListener('touchend', function (e) {
  if (e.target.closest('#closeSide')) {
    e.preventDefault();
    e.stopPropagation();
    closeSidePanel();
  }
}, { passive: false });

sideBackdrop.addEventListener('click', closeSidePanel);
sideBackdrop.addEventListener('touchend', e=>{ e.preventDefault(); closeSidePanel(); }, { passive: false });

if(daySheetClose) daySheetClose.addEventListener('click', closeDaySheet);
if(daySheetBackdrop){
  daySheetBackdrop.addEventListener('click', closeDaySheet);
  daySheetBackdrop.addEventListener('touchend', e=>{ e.preventDefault(); closeDaySheet(); }, { passive: false });
}
document.addEventListener('keydown', e => {
  if(e.key === 'Escape'){
    closeAllOverlays();
  }
});

/* ===== סגירת bottom-sheet בהחלקה למטה (swipe down) ===== */
(function(){
  let _tx = 0, _ty = 0;
  side.addEventListener('touchstart', e=>{
    _tx = e.touches[0].clientX;
    _ty = e.touches[0].clientY;
  }, { passive: true });
  side.addEventListener('touchend', e=>{
    const dx = Math.abs(e.changedTouches[0].clientX - _tx);
    const dy = e.changedTouches[0].clientY - _ty;
    // החלקה למטה (סגירה) — לפחות 60px ואנכית יותר מאופקית
    if(dy > 60 && dx < dy * 0.8 && sideContent.scrollTop <= 0){
      closeSidePanel();
  }
  }, { passive: true });
})();

initFromRawData();

window.addEventListener('popstate', (e)=>{
  if(isMobile() && (window.mode === 'month' || window.mode === 'week')){
    closeAllOverlays();
    view.innerHTML = '';
    renderMobileMonth();
  }
});
