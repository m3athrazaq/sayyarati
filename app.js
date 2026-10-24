'use strict';
/* ======================================================================
   سيّارتي — سجل السيارة
   البيانات تُحفظ على الجهاز فقط (localStorage، وفي تطبيق أندرويد أيضاً ملف داخل التطبيق).
   قواعد الأمان في هذا الملف (راجع SECURITY.md):
   - كل نص يُدرج في HTML يمر عبر esc()، وكل معرّف/لون/نوع يأتي من ملف يُتحقق منه في sanitize*().
   - لا eval ولا سكربتات خارجية؛ سياسة CSP في index.html تمنع أي سكربت غير ملفات التطبيق نفسه.
   ====================================================================== */
const APP_VERSION = '1.3.0';
const KEY = 'sayyarati.state.v1';
const NATIVE_FILE = 'sayyarati-state.json';
/* a backup with invoice photos can be large; the native picker has the same limit */
const MAX_IMPORT_BYTES = 150 * 1024 * 1024;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
/* files loaded on demand and their integrity hashes — written by tools/stamp.py; the browser refuses
   any file whose content does not match. The address carries part of the hash, so caches never mix versions. */
const VENDOR = {"vendor/exceljs.min.js": "sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz", "vendor/pdf-lib.min.js": "sha384-weMABwrltA6jWR8DDe9Jp5blk+tZQh7ugpCsF3JwSA53WZM9/14PjS5LAJNHNjAI", "vendor/fontkit.umd.min.js": "sha384-2p6U+1mmqF10USehFeRiyG2ESG9FwIqN+jxULn5w9jjQIihSn9Pt13dVCn/Hawjn", "vendor/bidi.min.js": "sha384-wQtTdqzwTFhzi4xIX4wrLGvJHshKoK61RCcjZ4lwr/sZBggODJHSpl5bHBcUfeV2", "report.js": "sha384-PbMBZxGmhPVv1iqD2xwVEq2zxKiUkuCEEvmGinsNI9a8dOFcCTYEobi34FkNn73+", "fonts/readex-report-400.ttf": "sha384-ZM9Zf46J0OTwwwL0lozkZa1qmDdWydinkx/JWuHjNv6S+0tSnrVMfJETkg5ZJl4l", "fonts/readex-report-700.ttf": "sha384-3xKwULn0RBRanuHVN8nTl+EbdhgM5YYhZQH4mk9u0MUTYD0fHrZkskxRt1gBAjVy"};

/* refuse to run inside a frame (another site embedding the app) */
if (window.top !== window.self) {
  document.documentElement.innerHTML = '';
  throw new Error('framed');
}

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"'`]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;', '`': '&#96;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const clone = o => JSON.parse(JSON.stringify(o));

const isNative = !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform());
const plug = n => (window.Capacitor && window.Capacitor.Plugins && window.Capacitor.Plugins[n]) || null;
let xlsxPluginRef = null;
function xlsxPlugin() {
  if (!isNative) return null;
  if (!xlsxPluginRef) xlsxPluginRef = plug('XlsxImport') || (window.Capacitor.registerPlugin ? window.Capacitor.registerPlugin('XlsxImport') : null);
  return xlsxPluginRef;
}

const TYPES = {
  text: 'نص',
  number: 'رقم',
  money: 'مبلغ',
  date: 'تاريخ',
  odometer: 'العداد (كم)',
  distance: 'المسافة (تُحسب تلقائياً)'
};
const CAR_COLORS = ['#E9A23B', '#3F8FD2', '#2F7D5B', '#B9452F', '#7B5CC4', '#8A9499', '#D46A9A', '#1F1F1F'];

const ICONS = {
  oil: '<path d="M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z"/>',
  wrench: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.6-.6-.6-2.6z"/>',
  fuel: '<path d="M4 21V5a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v16M3 21h12M4 10h10"/><path d="M14 8l3 3v6a1.5 1.5 0 0 0 3 0V9l-3-3"/>',
  tire: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="3.5"/><path d="M12 3v5.5M12 15.5V21M3 12h5.5M15.5 12H21"/>',
  shield: '<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"/>',
  doc: '<path d="M7 3h7l5 5v13H7z"/><path d="M14 3v5h5M10 13h6M10 17h6"/>',
  drop: '<path d="M7 16a5 5 0 0 0 10 0c0-3-5-9-5-9s-5 6-5 9z"/><path d="M4 21h16"/>',
  star: '<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z"/>'
};
const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
const icon = (k, s = 20) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${has(ICONS, k) ? ICONS[k] : ICONS.doc}</svg>`;
/* small UI icons (not offered as section icons) */
const UI_ICONS = {
  camera: '<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
  image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-9 9"/>',
  bell: '<path d="M6 16V11a6 6 0 0 1 12 0v5l2 2H4z"/><path d="M10 21h4"/>',
  cal: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  share: '<path d="M12 15V3M7 8l5-5 5 5"/><path d="M5 13v8h14v-8"/>',
  trash: '<path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>',
  gauge: '<path d="M4 15a8 8 0 1 1 16 0"/><path d="M12 15l4-5"/>'
};
const uiIcon = (k, s = 18) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${UI_ICONS[k] || ''}</svg>`;

/* ---------- templates ---------- */
const COL = (id, label, type) => ({ id, label, type });
/* reminderKm / reminderMonths: the next one is due after that distance or that time, whichever comes first (0 = off) */
const SECTION_TEMPLATES = {
  oil: () => ({ id: 'oil', name: 'الزيوت', icon: 'oil', note: '', reminderKm: 7500, reminderMonths: 6,
    columns: [COL('desc','بيان','text'), COL('odo','العداد','odometer'), COL('dist','المسافة','distance'), COL('date','التاريخ','date'), COL('amount','المبلغ','money'), COL('oil','الزيت','text'), COL('notes','ملاحظات','text')], rows: [] }),
  maint: () => ({ id: 'maint', name: 'الصيانة', icon: 'wrench', note: '', reminderKm: 0, reminderMonths: 0,
    columns: [COL('desc','بيان','text'), COL('odo','العداد','odometer'), COL('date','التاريخ','date'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  fuel: () => ({ id: 'fuel', name: 'الوقود', icon: 'fuel', note: '', reminderKm: 0, reminderMonths: 0,
    columns: [COL('date','التاريخ','date'), COL('odo','العداد','odometer'), COL('dist','المسافة','distance'), COL('liters','اللترات','number'), COL('amount','المبلغ','money'), COL('station','المحطة','text')], rows: [] }),
  tires: () => ({ id: 'tires', name: 'الإطارات', icon: 'tire', note: '', reminderKm: 0, reminderMonths: 0,
    columns: [COL('desc','بيان','text'), COL('odo','العداد','odometer'), COL('date','التاريخ','date'), COL('brand','النوع والمقاس','text'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  ins: () => ({ id: 'ins', name: 'التأمين والترخيص', icon: 'shield', note: '', reminderKm: 0, reminderMonths: 0,
    columns: [COL('desc','بيان','text'), COL('date','التاريخ','date'), COL('expiry','تاريخ الانتهاء','date'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  wash: () => ({ id: 'wash', name: 'الغسيل والعناية', icon: 'drop', note: '', reminderKm: 0, reminderMonths: 0,
    columns: [COL('desc','بيان','text'), COL('date','التاريخ','date'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  blank: () => ({ id: uid(), name: 'قسم جديد', icon: 'star', note: '', reminderKm: 0, reminderMonths: 0,
    columns: [COL(uid(),'بيان','text'), COL(uid(),'التاريخ','date'), COL(uid(),'المبلغ','money'), COL(uid(),'ملاحظات','text')], rows: [] })
};
const TEMPLATE_NAMES = { oil: 'الزيوت', maint: 'الصيانة', fuel: 'الوقود', tires: 'الإطارات', ins: 'التأمين والترخيص', wash: 'الغسيل والعناية', blank: 'قسم فارغ' };
const defaultCardFields = () => [
  { id: 'price', label: 'سعر الشراء', type: 'money', value: '' },
  { id: 'pdate', label: 'تاريخ الشراء', type: 'date', value: '' },
  { id: 'model', label: 'الموديل', type: 'number', value: '' },
  { id: 'odo0', label: 'العداد عند الشراء', type: 'odometer', value: '' },
  { id: 'plate', label: 'رقم اللوحة', type: 'text', value: '' }
];
function newCar(name, fromCar) {
  const car = { id: uid(), name: name || 'سيارة جديدة', color: CAR_COLORS[0], odoNow: null, created: Date.now() };
  if (fromCar) {
    car.cardFields = fromCar.cardFields.map(f => ({ ...clone(f), value: '' }));
    car.sections = fromCar.sections.map(s => ({ ...clone(s), note: '', rows: [] }));
  } else {
    car.cardFields = defaultCardFields();
    car.sections = [SECTION_TEMPLATES.oil(), SECTION_TEMPLATES.maint()];
  }
  return car;
}

/* ======================================================================
   Validation of everything that comes from storage or from a file.
   A backup file is untrusted input: every id, colour, type and value is
   checked here, so nothing in it can inject HTML or break the app.
   ====================================================================== */
const LIMITS = { cars: 50, cardFields: 60, sections: 40, columns: 40, rows: 5000, totalRows: 15000, text: 2000, label: 80, backupChars: 40e6,
  photosPerRow: 6, photos: 3000, photoBytes: 4 * 1024 * 1024 };
const ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
const COLOR_RE = /^#[0-9A-Fa-f]{6}$/;
const ISO_RE = /^\d{4}-\d{2}-\d{2}$/;
const isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const arr = v => (Array.isArray(v) ? v : []);
const str = (v, max) => {
  let t = (v == null || typeof v === 'object' ? '' : String(v)).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, max);
  if (/[\uD800-\uDBFF]$/.test(t)) t = t.slice(0, -1); /* never cut an emoji in half */
  return t;
};
const finite = (v, d = 0) => (typeof v === 'number' && isFinite(v) ? v : d);
function cleanId(v, used) {
  let id = typeof v === 'string' && ID_RE.test(v) && !(v in Object.prototype) ? v : uid();
  while (used.has(id)) id = uid();
  used.add(id);
  return id;
}
function cleanValue(type, v) {
  if (v instanceof Date) v = parseDateAny(v);
  if (v === '' || v == null || typeof v === 'object') return '';
  if (type === 'number' || type === 'money' || type === 'odometer') {
    if (typeof v === 'number') return isFinite(v) ? v : '';
    const n = toNum(v); return n == null ? str(v, 100) : n;
  }
  if (type === 'date') return parseDateAny(v) || str(v, 40);
  if (type === 'distance') return '';
  return str(v, LIMITS.text);
}
function sanitizeColumns(list) {
  const used = new Set(); const out = [];
  for (const c of arr(list).slice(0, LIMITS.columns)) {
    if (!isObj(c)) continue;
    out.push({ id: cleanId(c.id, used), label: str(c.label, LIMITS.label).trim() || 'بند', type: has(TYPES, c.type) ? c.type : 'text', _src: c.id });
  }
  if (!out.some(c => c.type === 'odometer')) out.forEach(c => { if (c.type === 'distance') c.type = 'number'; });
  return out;
}
function sanitizeSection(s, usedSec) {
  const columns = sanitizeColumns(s.columns);
  if (!columns.length) columns.push({ id: uid(), label: 'بيان', type: 'text', _src: null });
  const usedRow = new Set(); const rows = [];
  for (const r of arr(s.rows).slice(0, LIMITS.rows)) {
    if (!isObj(r)) continue;
    const src = isObj(r.v) ? r.v : {};
    const v = {};
    for (const c of columns) if (c.type !== 'distance') v[c.id] = typeof c._src === 'string' && has(src, c._src) ? cleanValue(c.type, src[c._src]) : '';
    const row = { id: cleanId(r.id, usedRow), t: finite(r.t, Date.now()), v };
    const p = cleanPhotoIds(r.p);
    if (p.length) row.p = p;
    rows.push(row);
  }
  columns.forEach(c => delete c._src);
  const reminderKm = Math.min(1e6, Math.max(0, finite(typeof s.reminderKm === 'string' ? toNum(s.reminderKm) : s.reminderKm, 0)));
  const months = finite(typeof s.reminderMonths === 'string' ? toNum(s.reminderMonths) : s.reminderMonths, NaN);
  return {
    id: cleanId(s.id, usedSec),
    name: str(s.name, LIMITS.label).trim() || 'قسم',
    icon: typeof s.icon === 'string' && has(ICONS, s.icon) ? s.icon : 'doc',
    note: str(s.note, 200),
    reminderKm,
    /* data from before 1.3 has no time interval: the oil section gets "every 6 months" like a new one */
    reminderMonths: isFinite(months) ? Math.min(120, Math.max(0, Math.round(months))) : (s.id === 'oil' && reminderKm > 0 ? 6 : 0),
    columns, rows
  };
}
/* invoice photo references of an entry: ids only (the photos themselves are stored separately) */
const cleanPhotoIds = list => [...new Set(arr(list).filter(x => typeof x === 'string' && ID_RE.test(x) && !(x in Object.prototype)))].slice(0, LIMITS.photosPerRow);
function sanitizeCar(c, usedCar) {
  if (!isObj(c) || (!Array.isArray(c.sections) && !Array.isArray(c.cardFields))) return null;
  const usedF = new Set(); const cardFields = [];
  for (const f of arr(c.cardFields).slice(0, LIMITS.cardFields)) {
    if (!isObj(f)) continue;
    const type = has(TYPES, f.type) && f.type !== 'distance' ? f.type : 'text';
    cardFields.push({ id: cleanId(f.id, usedF), label: str(f.label, LIMITS.label).trim() || 'بند', type, value: cleanValue(type, f.value) });
  }
  const usedS = new Set();
  const sections = arr(c.sections).slice(0, LIMITS.sections).filter(isObj).map(s => sanitizeSection(s, usedS));
  let odoNow = null;
  if (isObj(c.odoNow)) { const n = typeof c.odoNow.value === 'number' ? c.odoNow.value : toNum(c.odoNow.value); if (n != null && isFinite(n)) odoNow = { value: n, date: parseDateAny(c.odoNow.date) || '' }; }
  return {
    id: cleanId(c.id, usedCar),
    name: str(c.name, LIMITS.label).trim() || 'سيارة',
    color: typeof c.color === 'string' && COLOR_RE.test(c.color) ? c.color : CAR_COLORS[0],
    odoNow, created: finite(c.created, Date.now()), cardFields, sections
  };
}
const LEAD_DAYS = [7, 14, 30, 60];
const NUDGE_DAYS = [0, 14, 30, 60];
function sanitizeSettings(x) {
  const s = isObj(x) ? x : {};
  const dec = Number(s.decimals);
  const pick = (v, list, d) => (typeof v === 'number' && list.includes(v) ? v : d);
  return {
    currency: str(s.currency, 12).trim() || 'د.ك',
    decimals: [0, 2, 3].includes(dec) ? dec : 3,
    theme: ['auto', 'light', 'dark'].includes(s.theme) ? s.theme : 'auto',
    /* reminders: phone notifications (Android app), the hour they arrive, how early before an expiry,
       and how often to ask for a fresh odometer reading (0 = never) */
    notify: s.notify !== false,
    notifyHour: typeof s.notifyHour === 'number' && Number.isInteger(s.notifyHour) && s.notifyHour >= 0 && s.notifyHour <= 23 ? s.notifyHour : 9,
    expiryLead: pick(s.expiryLead, LEAD_DAYS, 30),
    odoNudge: pick(s.odoNudge, NUDGE_DAYS, 30)
  };
}
function sanitizeCars(list) { const used = new Set(); return arr(list).slice(0, LIMITS.cars).map(c => sanitizeCar(c, used)).filter(Boolean); }
const countRows = cars => cars.reduce((a, c) => a + c.sections.reduce((b, s) => b + s.rows.length, 0), 0);
/* last line of defence for stored data: never more rows than the app can handle */
function capRows(cars) {
  let left = LIMITS.totalRows;
  for (const c of cars) for (const s of c.sections) { if (s.rows.length > left) s.rows = s.rows.slice(0, Math.max(0, left)); left -= s.rows.length; }
  return cars;
}
function sanitizeState(x) {
  const s = isObj(x) ? x : {};
  const cars = capRows(sanitizeCars(s.cars));
  return {
    v: 1, cars,
    activeCar: cars.some(c => c.id === s.activeCar) ? s.activeCar : (cars[0] ? cars[0].id : null),
    settings: sanitizeSettings(s.settings),
    lastBackup: finite(s.lastBackup, 0), rev: finite(s.rev, 0)
  };
}

/* ---------- state ---------- */
let S = null;
function blankState() { return sanitizeState({}); }
function loadLocal() {
  try { const t = localStorage.getItem(KEY); if (t) return JSON.parse(t); } catch (e) {}
  return null;
}
async function loadNative() {
  const fs = plug('Filesystem'); if (!fs) return null;
  try { const r = await fs.readFile({ path: NATIVE_FILE, directory: 'DATA', encoding: 'utf8' }); return JSON.parse(r.data); } catch (e) { return null; }
}
let saveTimer = null, lastWarn = 0, saveOk = true;
function save() {
  S.rev = (S.rev || 0) + 1;
  const json = JSON.stringify(S);
  let ok = true;
  try { localStorage.setItem(KEY, json); } catch (e) { ok = false; }
  if (isNative) {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const fs = plug('Filesystem');
      if (!fs) { setSaveOk(ok); return; }
      fs.writeFile({ path: NATIVE_FILE, data: JSON.stringify(S), directory: 'DATA', encoding: 'utf8', recursive: true })
        .then(() => setSaveOk(true)).catch(e => { console.warn(e); setSaveOk(false); });
    }, 150);
    scheduleNotifSync();
  } else setSaveOk(ok);
}
function setSaveOk(ok) {
  const was = saveOk; saveOk = ok;
  if (!ok && Date.now() - lastWarn > 60e3) { lastWarn = Date.now(); toast('تعذّر حفظ آخر تعديل على الجهاز — صدّر نسخة احتياطية الآن'); }
  if (was !== ok && view === 'home' && S.cars.length) render();
}
const car = () => S.cars.find(c => c.id === S.activeCar) || S.cars[0];

/* ---------- formatting ---------- */
const nf = (n, d = 0) => (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
const money = n => `<span class="num">${nf(n, S.settings.decimals)}</span> ${esc(S.settings.currency)}`;
const km = n => `<span class="num">${nf(n)}</span> كم`;
/* plain text — always pass through esc() before putting it in HTML */
function fmtDate(v) {
  const s = String(v ?? '');
  if (ISO_RE.test(s)) { const [y, m, d] = s.split('-'); return `${d}.${m}.${y}`; }
  return s;
}
const dateHTML = v => `<span class="num">${esc(fmtDate(v))}</span>`;
function todayISO() { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); }
function isoOk(y, m, d) { return y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31; }
function parseDateAny(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return isNaN(v) ? '' : v.toISOString().slice(0, 10);
  if (typeof v === 'number') { if (!isFinite(v) || v < 1 || v > 200000) return ''; const d = new Date(Math.round((v - 25569) * 864e5)); return isNaN(d) ? '' : d.toISOString().slice(0, 10); }
  const s = String(v).trim().slice(0, 40);
  let m = s.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{4})$/);
  if (m && isoOk(+m[3], +m[2], +m[1])) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  m = s.match(/^(\d{4})[.\/\-](\d{1,2})[.\/\-](\d{1,2})(?:$|[T\s])/);
  if (m && isoOk(+m[1], +m[2], +m[3])) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return '';
}
const toNum = v => {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const t = String(v).slice(0, 60).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[,\s]/g, '');
  if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return null;
  const n = Number(t); return isFinite(n) ? n : null;
};
const daysBetween = (a, b) => (new Date(b) - new Date(a)) / 864e5;
const isoParts = iso => iso.split('-').map(Number);
function addDaysISO(iso, n) { const [y, m, d] = isoParts(iso); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); }
/* same day n months later (31 Aug + 6 months = 28/29 Feb) */
function addMonthsISO(iso, n) {
  const [y, m, d] = isoParts(iso);
  const last = new Date(Date.UTC(y, m + n, 0)).getUTCDate();
  return new Date(Date.UTC(y, m - 1 + n, Math.min(d, last))).toISOString().slice(0, 10);
}
const daysUntil = iso => Math.round(daysBetween(todayISO(), iso));
/* Arabic counted nouns: 1 يوم واحد، 2 يومين، 3–10 أيام، 11–99 يوماً */
function counted(n, one, two, few, many, plain) {
  n = Math.abs(Math.round(n));
  if (n === 1) return one;
  if (n === 2) return two;
  const t = n % 100;
  return `${nf(n)} ${t >= 3 && t <= 10 ? few : t >= 11 ? many : plain}`;
}
const nDays = n => counted(n, 'يوم واحد', 'يومين', 'أيام', 'يوماً', 'يوم');
const nMonths = n => counted(n, 'شهر واحد', 'شهرين', 'أشهر', 'شهراً', 'شهر');
const nPhotos = n => counted(n, 'صورة واحدة', 'صورتان', 'صور', 'صورة', 'صورة');
/* "بعد 5 أيام" / "اليوم" / "غداً" / "منذ يومين" */
function inDays(n) {
  if (n === 0) return 'اليوم';
  if (n === 1) return 'غداً';
  if (n === -1) return 'أمس';
  return n > 0 ? `بعد ${nDays(n)}` : `منذ ${nDays(-n)}`;
}

/* ---------- computations ---------- */
const colOf = (sec, type) => sec.columns.find(c => c.type === type);
const textCol = sec => sec.columns.find(c => c.type === 'text');
/* rows in chronological order (odometer first, then date) */
function chrono(sec) {
  const o = colOf(sec, 'odometer'), d = colOf(sec, 'date');
  return sec.rows.map((r, i) => ({ r, i })).sort((a, b) => {
    const ao = o ? toNum(a.r.v[o.id]) : null, bo = o ? toNum(b.r.v[o.id]) : null;
    if (ao != null && bo != null && ao !== bo) return ao - bo;
    const ad = d ? String(a.r.v[d.id] || '') : '', bd = d ? String(b.r.v[d.id] || '') : '';
    if (ad && bd && ad !== bd) return ad < bd ? -1 : 1;
    return (a.r.t || a.i) - (b.r.t || b.i);
  }).map(x => x.r);
}
/* distance since previous entry in the same section */
function distances(sec) {
  const o = colOf(sec, 'odometer'); const out = new Map(); if (!o) return out;
  let prev = null;
  for (const r of chrono(sec)) { const v = toNum(r.v[o.id]); if (v == null) continue; if (prev != null) out.set(r.id, v - prev); prev = v; }
  return out;
}
function sectionTotal(sec) {
  let t = 0; for (const c of sec.columns) if (c.type === 'money') for (const r of sec.rows) t += toNum(r.v[c.id]) || 0; return t;
}
function odoPoints(c) {
  const pts = [];
  for (const s of c.sections) {
    const o = colOf(s, 'odometer'), d = colOf(s, 'date'); if (!o) continue;
    for (const r of s.rows) { const v = toNum(r.v[o.id]); if (v != null) pts.push({ odo: v, date: d && ISO_RE.test(r.v[d.id]) ? r.v[d.id] : '' }); }
  }
  const pd = c.cardFields.find(x => x.type === 'date');
  for (const f of c.cardFields) if (f.type === 'odometer' && toNum(f.value) != null) pts.push({ odo: toNum(f.value), date: pd && ISO_RE.test(pd.value) ? pd.value : '' });
  if (c.odoNow && toNum(c.odoNow.value) != null) pts.push({ odo: toNum(c.odoNow.value), date: c.odoNow.date || '' });
  return pts;
}
function currentOdo(c) { let m = 0; for (const p of odoPoints(c)) if (p.odo > m) m = p.odo; return m; }
function firstOdo(c) { let m = Infinity; for (const p of odoPoints(c)) if (p.odo < m) m = p.odo; return m === Infinity ? 0 : m; }
/* average km/day from dated readings (ignores readings whose date goes backwards) */
function kmPerDay(c) {
  const pts = odoPoints(c).filter(p => p.date).sort((a, b) => a.odo - b.odo);
  const kept = [];
  for (const p of pts) if (!kept.length || p.date > kept[kept.length - 1].date) kept.push(p);
  if (kept.length < 2) return null;
  const last = kept[kept.length - 1];
  let first = kept[0];
  for (const p of kept) if (daysBetween(p.date, last.date) <= 365) { first = p; break; }
  if (first === last) first = kept[kept.length - 2];
  const days = daysBetween(first.date, last.date);
  if (days < 7) return null;
  return (last.odo - first.odo) / days;
}
/* the latest odometer reading that has a date */
function lastDated(c) {
  let best = null;
  for (const p of odoPoints(c)) if (p.date && (!best || p.date > best.date || (p.date === best.date && p.odo > best.odo))) best = p;
  return best;
}
/* the day the odometer should reach `target`, from the latest dated reading and the driving rate */
function etaFor(c, target, rate) {
  const base = lastDated(c);
  if (!base || !rate || rate <= 0) return null;
  const days = Math.round((target - base.odo) / rate);
  return isFinite(days) && Math.abs(days) <= 3650 ? addDaysISO(base.date, days) : null;
}
/* next service of each section that has a reminder: by distance (reminderKm), by time (reminderMonths),
   or both — whichever comes first */
function reminders(c) {
  const out = [];
  const cur = currentOdo(c), rate = kmPerDay(c), today = todayISO();
  for (const s of c.sections) {
    const o = colOf(s, 'odometer'), d = colOf(s, 'date');
    const useKm = !!o && s.reminderKm > 0, useTime = !!d && s.reminderMonths > 0;
    if (!useKm && !useTime) continue;
    let kmP = null, timeP = null;
    if (useKm) {
      let lastOdo = null;
      for (const r of s.rows) { const v = toNum(r.v[o.id]); if (v != null && (lastOdo == null || v > lastOdo)) lastOdo = v; }
      if (lastOdo != null) {
        const next = lastOdo + s.reminderKm, used = Math.max(0, cur - lastOdo);
        kmP = { lastOdo, next, used, left: next - cur, ratio: used / s.reminderKm, eta: etaFor(c, next, rate) };
      }
    }
    if (useTime) {
      let last = '';
      for (const r of s.rows) { const v = r.v[d.id]; if (ISO_RE.test(v) && v > last) last = v; }
      if (last) {
        const due = addMonthsISO(last, s.reminderMonths);
        timeP = { last, due, days: Math.round(daysBetween(today, due)), ratio: Math.max(0, daysBetween(last, today)) / Math.max(1, daysBetween(last, due)) };
      }
    }
    if (!kmP && !timeP) { out.push({ sec: s, none: true }); continue; }
    let by = kmP ? 'km' : 'time';
    if (kmP && timeP) {
      if (kmP.left <= 0) by = 'km';
      else if (timeP.days < 0) by = 'time';
      else if (kmP.eta) by = kmP.eta <= timeP.due ? 'km' : 'time';
      else by = kmP.ratio >= timeP.ratio ? 'km' : 'time';
    }
    /* the date it falls due: the time limit, or the day the distance should be reached at the usual driving
       rate (an estimate). Once overdue, the date it became due; otherwise the next one still ahead. */
    const late = !!((kmP && kmP.left <= 0) || (timeP && timeP.days < 0));
    let due = null;
    if (late) due = [kmP && kmP.left <= 0 ? kmP.eta : null, timeP && timeP.days < 0 ? timeP.due : null].filter(Boolean).sort()[0] || null;
    else due = [timeP && timeP.due, kmP && kmP.eta].filter(x => x && x >= today).sort()[0] || null;
    out.push({ sec: s, km: kmP, time: timeP, by, due, late, est: !!due && !(timeP && timeP.due === due),
      /* the last reading is old: at the usual rate the distance has probably been reached already */
      stale: !!(kmP && kmP.eta && kmP.eta < today && kmP.left > 0),
      ratio: Math.max(kmP ? kmP.ratio : 0, timeP ? timeP.ratio : 0) });
  }
  return out;
}
/* "عند 15,000 كم أو في 01.04.2027" (plain text) */
function svcWhen(x) {
  const parts = [];
  if (x.km) parts.push(`عند ${nf(x.km.next)} كم`);
  if (x.time) parts.push(`في ${fmtDate(x.time.due)}`);
  return parts.join(' أو ') + (parts.length > 1 ? '، أيهما أسبق' : '');
}

/* expiry reminders (insurance, registration…): date items whose name says they expire, in the car card
   and in the sections. Entries are grouped by what they are, so a section holding both the insurance and
   the registration reminds of each; anything that expired more than 60 days ago is no longer shown. */
const EXPIRY_RE = /انتهاء|ينتهي|صلاحية|تجديد|expir|renew/i;
function expiryGroup(title) {
  const t = norm(title);
  if (/تأمين|تامين|insurance/.test(t)) return 'ins';
  if (/دفتر|ترخيص|استمار|registration/.test(t)) return 'reg';
  if (/رخص|licen[cs]e/.test(t)) return 'lic';
  if (/فحص|inspection/.test(t)) return 'insp';
  return t.replace(/[0-9٠-٩]+/g, ' ').replace(/\s+/g, ' ').trim() || '-';
}
function expiries(c) {
  const out = [], today = todayISO();
  const add = (key, title, date, sec) => { const days = Math.round(daysBetween(today, date)); if (days >= -60) out.push({ key, title, date, days, sec }); };
  for (const f of c.cardFields) if (f.type === 'date' && EXPIRY_RE.test(f.label) && ISO_RE.test(f.value)) add('f-' + f.id, f.label, f.value, null);
  for (const s of c.sections) for (const col of s.columns) {
    if (col.type !== 'date' || !EXPIRY_RE.test(col.label)) continue;
    const tc = textCol(s), groups = new Map();
    for (const r of s.rows) {
      const v = r.v[col.id]; if (!ISO_RE.test(v)) continue;
      const title = String((tc && r.v[tc.id]) || '').trim() || s.name;
      const g = expiryGroup(title), cur = groups.get(g);
      if (!cur || v > cur.date) groups.set(g, { date: v, title });
    }
    for (const [g, x] of groups) add(`s-${s.id}-${col.id}-${hashId(g).toString(36)}`, x.title, x.date, s);
  }
  return out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}
/* every upcoming date of a car, for the list, the calendar file and the phone notifications */
function dueItems(c) {
  const items = [];
  for (const x of reminders(c)) if (!x.none) items.push({ type: 'svc', key: 'svc-' + x.sec.id, c, x, date: x.due });
  for (const e of expiries(c)) items.push({ type: 'exp', key: 'exp-' + e.key, c, e, date: e.date });
  return items;
}
/* stable 31-bit number for a text (notification ids, calendar ids) */
function hashId(s) {
  let h = 0x811c9dc5;
  for (const ch of String(s)) { h ^= ch.codePointAt(0); h = Math.imul(h, 0x01000193); }
  return (h >>> 1) || 1;
}

/* ---------- UI helpers ---------- */
let toastT;
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('on'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('on'), 2800); }
function openSheet(title, html, onMount) {
  releaseUrls();
  $('#sheetTitle').textContent = title;
  $('#sheetBody').innerHTML = html;
  $('#sheetBody').scrollTop = 0;
  $('#scrim').classList.add('on'); $('#sheet').classList.add('on');
  if (onMount) onMount($('#sheetBody'));
}
function closeSheet() { $('#scrim').classList.remove('on'); $('#sheet').classList.remove('on'); setTimeout(() => { if (!sheetOpen()) releaseUrls(); }, 400); }
const sheetOpen = () => $('#sheet').classList.contains('on');
function confirmBox(msg, okLabel, onOk, danger = true, onCancel = null) {
  openSheet('تأكيد', `<p style="margin:4px 0 20px">${esc(msg)}</p>
    <div class="btn-row"><button class="btn ghost" id="cNo">إلغاء</button><button class="btn ${danger ? 'danger' : ''}" id="cYes">${esc(okLabel)}</button></div>`, b => {
    b.querySelector('#cNo').onclick = () => { closeSheet(); if (onCancel) onCancel(); };
    b.querySelector('#cYes').onclick = () => { closeSheet(); onOk(); };
  });
}

/* ---------- theme ---------- */
function applyTheme() {
  const t = S.settings.theme;
  if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t; else delete document.documentElement.dataset.theme;
}

/* ---------- navigation ---------- */
let view = 'home', activeSec = null, query = '';
function go(v) {
  view = v;
  document.querySelectorAll('#nav button[data-v]').forEach(b => { if (b.dataset.v === v) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
  render(); window.scrollTo(0, 0);
}

function render() {
  const hasCars = S.cars.length > 0;
  $('#top').classList.toggle('hidden', !hasCars);
  $('#nav').classList.toggle('hidden', !hasCars);
  document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
  if (!hasCars) { renderWelcome(); $('#v-welcome').classList.add('on'); return; }
  const c = car();
  document.documentElement.style.setProperty('--car', COLOR_RE.test(c.color) ? c.color : CAR_COLORS[0]);
  $('#carPillName').textContent = c.name;
  try { ({ home: renderHome, logs: renderLogs, stats: renderStats, settings: renderSettings })[view](); }
  catch (e) { console.error(e); renderRescue(e); }
  $('#v-' + view).classList.add('on');
}
/* shown instead of a screen that failed to draw: the way out is always available */
function renderRescue(e) {
  $('#v-' + view).innerHTML = `<div class="block lift" style="margin-top:16px">
    <h2>تعذّر عرض هذه الصفحة</h2>
    <p style="margin:0 0 14px;color:var(--muted);font-size:14px">يبدو أن بيانات «${esc(car().name)}» فيها شيء غير متوقع. احفظ نسخة احتياطية، أو انتقل إلى سيارة أخرى، أو احذف هذه السيارة.</p>
    <p style="margin:0 0 14px;color:var(--muted);font-size:12px;direction:ltr;text-align:left">${esc(String((e && e.message) || e).slice(0, 160))}</p>
    <div class="btn-row"><button class="btn amber" id="rsExp">نسخة احتياطية</button><button class="btn ghost" id="rsCars">السيارات</button></div>
    <button class="btn danger wide" id="rsDel" style="margin-top:10px">حذف هذه السيارة</button></div>`;
  $('#rsExp').onclick = openBackup;
  $('#rsCars').onclick = openCarSwitcher;
  $('#rsDel').onclick = () => { const c = car(); confirmBox(`حذف «${c.name}» وكل سجلاتها؟`, 'حذف السيارة', () => { S.cars = S.cars.filter(x => x.id !== c.id); S.activeCar = S.cars[0] ? S.cars[0].id : null; save(); go('home'); }); };
}

/* ---------- welcome ---------- */
function renderWelcome() {
  $('#v-welcome').innerHTML = `<div class="welcome">
    ${odoHTML(0, true)}
    <h1>سيّارتي</h1>
    <p>سجّل تبديل الزيت والصيانة وكل مصاريف سياراتك، وتابع موعد التبديل القادم، واحتفظ بنسخة إكسل منسّقة في الدرايف.</p>
    <button class="btn amber wide" id="wImport">استيراد من ملف إكسل</button>
    <button class="btn wide" id="wNew">إضافة سيارة جديدة</button>
  </div>`;
  $('#wImport').onclick = () => pickImport();
  $('#wNew').onclick = () => openCarEditor(null);
}

/* ---------- odometer ---------- */
function odoHTML(n, still) {
  const s = String(Math.max(0, Math.min(9999999, Math.round(Number(n) || 0)))).padStart(6, '0');
  return `<div class="odo" aria-label="${nf(n)} كم" role="img">${[...s].map((d, i) => `<span class="d ${i === s.length - 1 ? 'km' : ''}"><span class="strip" data-d="${d}" style="transform:translateY(${still ? -d * 10 : 0}%)">${'0123456789'.split('').map(x => `<span>${x}</span>`).join('')}</span></span>`).join('')}</div>`;
}
let rolledFor = null;
function rollOdo(root, key) {
  const strips = root.querySelectorAll('.odo .strip');
  const run = () => strips.forEach(st => st.style.transform = `translateY(${-st.dataset.d * 10}%)`);
  if (rolledFor === key) { strips.forEach(st => { st.style.transition = 'none'; }); run(); return; }
  rolledFor = key;
  requestAnimationFrame(() => requestAnimationFrame(run));
}

/* ---------- home ---------- */
function gaugeSVG(ratio) {
  const r = Math.max(0, Math.min(1, ratio));
  const col = ratio >= 1 ? 'var(--bad)' : ratio >= .85 ? 'var(--warn)' : 'var(--ok)';
  const L = Math.PI * 50;
  return `<svg class="gauge" viewBox="0 0 120 70" aria-hidden="true">
    <path d="M10 62a50 50 0 0 1 100 0" fill="none" stroke="var(--bg)" stroke-width="12" stroke-linecap="round"/>
    <path d="M10 62a50 50 0 0 1 100 0" fill="none" stroke="${col}" stroke-width="12" stroke-linecap="round" stroke-dasharray="${L * r} ${L}"/>
    <text x="60" y="60" text-anchor="middle" font-size="19" font-weight="700" fill="var(--text)" font-family="Readex">${Math.round(ratio * 100)}%</text></svg>`;
}
function renderHome() {
  const c = car(), cur = currentOdo(c), rate = kmPerDay(c);
  const sub = [];
  const model = c.cardFields.find(f => f.label.includes('موديل') && f.value !== '');
  if (model) sub.push('موديل ' + model.value);
  const pl = c.cardFields.find(f => f.label.includes('لوحة') && f.value !== '');
  if (pl) sub.push('لوحة ' + pl.value);
  const rem = reminders(c), exp = expiries(c);
  const total = c.sections.reduce((a, s) => a + sectionTotal(s), 0);
  const driven = cur - firstOdo(c);

  let dueHTML = '';
  for (const x of rem) {
    if (x.none) { dueHTML += `<div class="due"><div class="due-txt"><b>${esc(x.sec.name)}</b><small>سجّل أول عملية ليبدأ حساب الموعد القادم.</small></div></div>`; continue; }
    const k = x.km, t = x.time;
    let head, foot;
    if (x.by === 'km') {
      head = k.left > 0 ? `باقي ${km(k.left)}` : `متأخر ${km(-k.left)}`;
      const d = k.eta && k.left > 0 ? daysUntil(k.eta) : null;
      foot = x.stale ? '<span class="warn-t">آخر قراءة قديمة، وحسب معدل قيادتك ربما حان الموعد. حدّث العداد.</span>'
        : d != null && d >= 0 ? (d <= 1 ? 'خلال يوم تقريباً حسب معدل قيادتك' : `بعد نحو ${esc(nDays(d))} حسب معدل قيادتك`) : `قطعت ${km(k.used)} من ${km(x.sec.reminderKm)}`;
    } else {
      head = t.days > 0 ? `باقي ${esc(nDays(t.days))}` : t.days === 0 ? 'الموعد اليوم' : `متأخر ${esc(nDays(-t.days))}`;
      foot = `آخر مرة ${dateHTML(t.last)}`;
    }
    const when = [k ? `عند ${km(k.next)}` : '', t ? `في ${dateHTML(t.due)}` : ''].filter(Boolean).join(' أو ');
    dueHTML += `<div class="due">${gaugeSVG(x.ratio)}<div class="due-txt"><small>${esc(x.sec.name)} القادم ${when}${k && t ? ' (أيهما أسبق)' : ''}</small><b>${head}</b><small>${foot}</small></div></div>`;
  }
  for (const e of exp) {
    const head = e.days > 0 ? `ينتهي ${esc(inDays(e.days))}` : e.days === 0 ? 'ينتهي اليوم' : `انتهى ${esc(inDays(e.days))}`;
    dueHTML += `<div class="due"><div class="due-txt"><small>${esc(e.title)}</small><b style="color:${e.days < 0 ? 'var(--bad)' : e.days < 30 ? 'var(--warn)' : 'inherit'}">${head}</b><small>${dateHTML(e.date)}</small></div></div>`;
  }
  if (dueHTML) dueHTML += `<div class="due-foot"><button id="hUpcoming">${uiIcon('cal', 16)} كل المواعيد${isNative ? ' والإشعارات' : ' والتقويم'}</button></div>`;
  /* the odometer reading the distance reminders rely on is getting old */
  const ld = lastDated(c);
  const usesKm = c.sections.some(s => s.reminderKm > 0 && colOf(s, 'odometer'));
  const staleDays = ld && usesKm && S.settings.odoNudge > 0 ? -daysUntil(ld.date) : 0;
  const askNotify = isNative && S.settings.notify && /^prompt/.test(notifPerm || '') && (rem.some(x => !x.none) || exp.length);

  const recent = [];
  for (const s of c.sections) for (const r of s.rows) recent.push({ s, r });
  const dkey = x => { const d = colOf(x.s, 'date'); return String((d && x.r.v[d.id]) || ''); };
  recent.sort((a, b) => dkey(b).localeCompare(dkey(a)) || (b.r.t || 0) - (a.r.t || 0));

  const daysSinceBackup = S.lastBackup ? Math.floor((Date.now() - S.lastBackup) / 864e5) : null;
  const backupWarn = (daysSinceBackup == null || daysSinceBackup >= 14) && S.cars.some(x => x.sections.some(s => s.rows.length));

  $('#v-home').innerHTML = `
  <div class="cluster">
    <div class="car-name">${esc(c.name)}</div>
    <p class="car-sub">${esc(sub.join('، ')) || '&nbsp;'}</p>
    <div class="odo-label">العداد الحالي${staleDays >= S.settings.odoNudge && staleDays > 0 ? ` · <span class="stale">آخر قراءة ${esc(inDays(-staleDays))}</span>` : ''}</div>
    <div class="odo-wrap">
      <div style="display:flex;align-items:flex-end">${odoHTML(cur)}<span class="odo-unit">كم</span></div>
      <button class="odo-btn" id="hOdo">تحديث العداد</button>
    </div>
  </div>
  ${dueHTML ? `<div class="block lift due-list">${dueHTML}</div>` : ''}
  ${askNotify ? `<div class="backup-note notify-note">${uiIcon('bell', 18)}<span>فعّل الإشعارات ليذكّرك التطبيق بموعد الزيت وانتهاء التأمين حتى وهو مغلق.</span><button id="hNotifyOn">تفعيل</button><button id="hNotifyNo" class="mute">ليس الآن</button></div>` : ''}
  ${!saveOk ? `<div class="backup-note" style="background:color-mix(in srgb,var(--bad) 18%,transparent)">آخر تعديل لم يُحفظ على الجهاز.<button id="hSaveFail">صدّر نسخة احتياطية</button></div>` : ''}
  ${backupWarn ? `<div class="backup-note">${daysSinceBackup == null ? 'لم تحفظ نسخة احتياطية بعد.' : `آخر نسخة احتياطية قبل ${daysSinceBackup} يوماً.`}<button id="hBackup">احفظ الآن</button></div>` : ''}
  <div class="block ${dueHTML ? '' : 'lift'}">
    <div class="stats">
      <div class="stat"><small>إجمالي المصاريف</small><b>${money(total)}</b></div>
      <div class="stat"><small>تكلفة الكيلو</small><b>${driven > 0 ? `<span class="num">${nf(total / driven * 1000, 1)}</span><em>فلس</em>` : '—'}</b></div>
      <div class="stat"><small>${rate ? 'معدل القيادة' : 'المسافة المسجلة'}</small><b>${rate ? `<span class="num">${nf(rate * 30)}</span><em>كم/شهر</em>` : km(driven)}</b></div>
    </div>
  </div>
  <div class="block">
    <h2>بطاقة السيارة <button class="act" id="hCard">تعديل</button></h2>
    <dl class="kv">${c.cardFields.map(f => `<dt>${esc(f.label)}</dt><dd>${fmtVal(f.type, f.value) || '<span style="color:var(--muted)">—</span>'}</dd>`).join('')}</dl>
  </div>
  <div class="block">
    <h2>آخر التسجيلات <button class="act" id="hAll">الكل</button></h2>
    ${recent.length ? `<div class="list">${recent.slice(0, 4).map(x => recentItem(x.s, x.r)).join('')}</div>` : `<div class="empty"><b>لا توجد تسجيلات بعد</b>اضغط زر + لتسجيل تبديل زيت أو صيانة.</div>`}
  </div>`;
  rollOdo($('#v-home'), c.id + ':' + cur);
  $('#hOdo').onclick = openOdoUpdate;
  $('#hCard').onclick = () => openCardEditor();
  $('#hAll').onclick = () => go('logs');
  if ($('#hBackup')) $('#hBackup').onclick = openBackup;
  if ($('#hSaveFail')) $('#hSaveFail').onclick = openBackup;
  if ($('#hUpcoming')) $('#hUpcoming').onclick = openUpcoming;
  if ($('#hNotifyOn')) $('#hNotifyOn').onclick = () => enableNotifications(true);
  if ($('#hNotifyNo')) $('#hNotifyNo').onclick = () => { S.settings.notify = false; save(); render(); toast('يمكنك تفعيل الإشعارات لاحقاً من الإعدادات'); };
  $('#v-home').querySelectorAll('[data-sec][data-row]').forEach(b => b.onclick = () => openEntry(b.dataset.sec, b.dataset.row));
}
function recentItem(s, r) {
  const t = textCol(s), d = colOf(s, 'date'), m = s.columns.find(c => c.type === 'money');
  return `<button class="list-item" data-sec="${esc(s.id)}" data-row="${esc(r.id)}"><span class="ic">${icon(s.icon)}</span>
    <span class="grow"><b style="font-weight:500">${esc((t && r.v[t.id]) || s.name)}</b><small>${esc(s.name)}${d && r.v[d.id] ? ` · ${dateHTML(r.v[d.id])}` : ''}${r.p && r.p.length ? ` · ${uiIcon('camera', 13)}` : ''}</small></span>
    ${m && toNum(r.v[m.id]) ? `<b style="font-weight:500;white-space:nowrap">${money(r.v[m.id])}</b>` : ''}</button>`;
}
function fmtVal(type, v) {
  if (v === '' || v == null) return '';
  if (type === 'money') return money(v);
  if (type === 'odometer' || type === 'distance') return km(v);
  if (type === 'date') return dateHTML(v);
  if (type === 'number') return `<span class="num">${esc(v)}</span>`;
  return esc(v);
}

/* ---------- logs ---------- */
function renderLogs() {
  const c = car();
  if (!c.sections.find(s => s.id === activeSec)) activeSec = c.sections[0] && c.sections[0].id;
  const s = c.sections.find(x => x.id === activeSec);
  let body = '';
  if (!s) body = `<div class="empty"><b>لا توجد أقسام</b>أضف قسماً من الإعدادات.</div>`;
  else {
    const dist = distances(s);
    const q = query.trim();
    let rows = chrono(s).reverse();
    if (q) rows = rows.filter(r => Object.values(r.v).some(v => String(v).includes(q)) || s.columns.some(c2 => c2.type === 'date' && fmtDate(r.v[c2.id]).includes(q)));
    const t = textCol(s), m = s.columns.filter(x => x.type === 'money'), o = colOf(s, 'odometer'), d = colOf(s, 'date'), dc = colOf(s, 'distance');
    const others = s.columns.filter(x => x !== t && !m.includes(x) && x !== o && x !== d && x.type !== 'distance');
    body = `
    <div class="sec-head"><div><h1>${esc(s.name)}</h1><p>${s.rows.length} ${s.rows.length === 1 ? 'تسجيل' : 'تسجيلات'}${s.note ? ' · ' + esc(s.note) : ''}</p></div>
      ${m.length ? `<div class="sec-total"><small>المجموع</small><b>${money(sectionTotal(s))}</b></div>` : ''}</div>
    ${s.rows.length > 5 ? `<input class="search" id="lq" type="search" placeholder="ابحث في ${esc(s.name)}" value="${esc(query)}">` : ''}
    ${rows.length ? `<ul class="log">${rows.map(r => {
      const amt = m.reduce((a, x) => a + (toNum(r.v[x.id]) || 0), 0);
      const meta = [];
      if (o && r.v[o.id] !== '' && r.v[o.id] != null) meta.push(`<span>العداد <b>${km(r.v[o.id])}</b></span>`);
      if (d && r.v[d.id]) meta.push(`<span><b>${dateHTML(r.v[d.id])}</b></span>`);
      if (dc && dist.has(r.id)) meta.push(`<span class="gap-tag">+${km(dist.get(r.id))}</span>`);
      for (const x of others) if (r.v[x.id] !== '' && r.v[x.id] != null) meta.push(`<span>${esc(x.label)}: <b>${fmtVal(x.type, r.v[x.id])}</b></span>`);
      if (r.p && r.p.length) meta.push(`<span class="clip" aria-label="${r.p.length} صورة">${uiIcon('camera', 14)}<b class="num">${r.p.length}</b></span>`);
      return `<li><button class="entry" data-row="${esc(r.id)}"><div class="entry-top"><span class="entry-title">${esc((t && r.v[t.id]) || s.name)}</span>${m.length ? `<span class="entry-amt">${amt ? money(amt) : ''}</span>` : ''}</div>
        <div class="entry-meta">${meta.join('')}</div></button></li>`;
    }).join('')}</ul>` : `<div class="empty"><b>${q ? 'لا نتائج' : 'القسم فارغ'}</b>${q ? 'جرّب كلمة أخرى.' : 'اضغط زر + لإضافة أول تسجيل.'}</div>`}`;
  }
  $('#v-logs').innerHTML = `<div class="seg" role="toolbar">${c.sections.map(x => `<button class="chip" data-s="${esc(x.id)}" aria-pressed="${x.id === activeSec}">${icon(x.icon, 17)} ${esc(x.name)} <span class="cnt">${x.rows.length}</span></button>`).join('')}
    <button class="chip" id="lAddSec" aria-label="إضافة قسم">+ قسم</button></div>${body}`;
  $('#v-logs').querySelectorAll('[data-s]').forEach(b => b.onclick = () => { activeSec = b.dataset.s; query = ''; renderLogs(); });
  $('#v-logs').querySelectorAll('.entry').forEach(b => b.onclick = () => openEntry(activeSec, b.dataset.row));
  $('#lAddSec').onclick = openAddSection;
  const lq = $('#lq'); if (lq) lq.oninput = () => { query = lq.value; const pos = lq.selectionStart; renderLogs(); const n = $('#lq'); if (n) { n.focus(); n.setSelectionRange(pos, pos); } };
}

/* ---------- entry form ---------- */
let draft = null;
function openEntry(secId, rowId) {
  const c = car();
  draft = { carId: c.id, secId: secId || activeSec || (c.sections[0] && c.sections[0].id), rowId: rowId || null, v: {}, p: [] };
  if (rowId) {
    const s = c.sections.find(x => x.id === draft.secId); const r = s && s.rows.find(x => x.id === rowId);
    if (r) { draft.v = clone(r.v); draft.p = (r.p || []).slice(); } else draft.rowId = null;
  }
  renderEntry();
}
/* ---------- invoice photos in the entry form ---------- */
let photoBusy = 0;
function photoTilesHTML() {
  const tiles = draft.p.map((id, i) => `<button class="ph" data-pi="${i}" aria-label="صورة الفاتورة ${i + 1}"><img data-ph="${esc(id)}" alt=""></button>`).join('');
  return tiles + (photoBusy ? `<span class="ph busy" role="status">جارٍ التجهيز…</span>` : '');
}
function redrawPhotos() {
  const box = $('#phList'); if (!box || !draft) return;
  box.innerHTML = photoTilesHTML();
  box.classList.toggle('empty', !draft.p.length && !photoBusy);
  const full = draft.p.length + photoBusy >= LIMITS.photosPerRow;
  $('#sheetBody').querySelectorAll('#phCam,#phPick').forEach(b => { b.disabled = full; });
  box.querySelectorAll('[data-pi]').forEach(b => b.onclick = () => openViewer(draft.p.slice(), +b.dataset.pi, id => { const k = draft.p.indexOf(id); if (k >= 0) draft.p.splice(k, 1); redrawPhotos(); }));
  fillPhotos(box);
}
/* puts the stored images into <img data-ph="id"> (a missing one is marked, never an error) */
async function fillPhotos(root) {
  for (const img of root.querySelectorAll('img[data-ph]')) {
    const u = await Photos.url(img.dataset.ph).catch(() => null);
    if (!img.isConnected) continue;
    if (u) { img.onerror = () => img.closest('.ph').classList.add('missing'); img.src = u; }
    else img.closest('.ph').classList.add('missing');
  }
}
async function addPhotoFiles(files) {
  const d = draft;
  if (!d || !files || !files.length) return;
  const room = LIMITS.photosPerRow - d.p.length - photoBusy;
  if (files.length > room) toast(`يمكن إرفاق ${LIMITS.photosPerRow} صور بحد أقصى لكل تسجيل`);
  for (const f of [...files].slice(0, Math.max(0, room))) {
    photoBusy++; redrawPhotos();
    try {
      const blob = await compressPhoto(f);
      const id = await Photos.put(blob);
      d.p.push(id);
    } catch (e) { console.warn('photo', e); toast('تعذّر فتح الصورة. جرّب صورة أخرى (JPG أو PNG).'); }
    finally { photoBusy--; if (draft === d) redrawPhotos(); }
  }
}
function renderEntry() {
  const c = S.cars.find(x => x.id === draft.carId);
  if (!c) { closeSheet(); return; }
  if (!c.sections.find(x => x.id === draft.secId)) draft.secId = c.sections[0] && c.sections[0].id;
  const s = c.sections.find(x => x.id === draft.secId);
  const editing = !!draft.rowId;
  const cur = currentOdo(c);
  let fields = '';
  if (!s) fields = `<div class="empty"><b>لا توجد أقسام لهذه السيارة</b>أضف قسماً من الإعدادات.</div>`;
  else {
    s.columns.forEach((col, ci) => {
      if (!editing && col.type === 'date' && !has(draft.v, col.id) && !/انتهاء|ينتهي|صلاحية/.test(col.label)) draft.v[col.id] = todayISO();
      const v = has(draft.v, col.id) ? draft.v[col.id] : '';
      if (col.type === 'distance') { fields += `<div class="field"><span>${esc(col.label)}</span><div class="calc" id="calcDist">—</div></div>`; return; }
      const attrs = { text: 'type="text"', number: 'type="text" inputmode="decimal"', money: 'type="text" inputmode="decimal"', odometer: 'type="text" inputmode="numeric"', date: 'type="date"' }[col.type] || 'type="text"';
      const hint = col.type === 'odometer' && cur ? `<div class="hint" id="odoHint">آخر قراءة معروفة: ${nf(cur)} كم</div>` : col.type === 'money' ? `<div class="hint">${esc(S.settings.currency)}</div>` : '';
      const listId = col.type === 'text' ? `dl-${ci}` : '';
      const isNum = col.type === 'number' || col.type === 'money' || col.type === 'odometer';
      fields += `<label class="field"><span>${esc(col.label)}</span><input ${attrs} data-ci="${ci}" value="${esc(v)}" maxlength="${LIMITS.text}" ${listId ? `list="${listId}" autocomplete="off"` : ''} ${isNum ? 'class="num"' : ''}>${hint}</label>`;
      if (listId) { const opts = [...new Set(s.rows.map(r => r.v[col.id]).filter(Boolean))].slice(-30); fields += `<datalist id="${listId}">${opts.map(o => `<option value="${esc(o)}">`).join('')}</datalist>`; }
    });
  }
  const carsPick = S.cars.length > 1 && !editing ? `<div class="pick" aria-label="السيارة">${S.cars.map((x, i) => `<button class="chip" data-car-i="${i}" aria-pressed="${x.id === draft.carId}"><span class="dot" style="background:${esc(x.color)}"></span>${esc(x.name)}</button>`).join('')}</div>` : '';
  const secPick = !editing ? `<div class="pick" aria-label="القسم">${c.sections.map((x, i) => `<button class="chip" data-sec-i="${i}" aria-pressed="${x.id === draft.secId}">${icon(x.icon, 16)} ${esc(x.name)}</button>`).join('')}</div>` : '';
  const photos = s ? `<div class="field"><span>صور الفاتورة</span><div class="photos" id="phList"></div>
    <div class="ph-btns"><button class="btn ghost" id="phCam">${uiIcon('camera')} تصوير</button><button class="btn ghost" id="phPick">${uiIcon('image')} من الصور</button></div></div>` : '';
  openSheet(editing ? `تعديل — ${s ? s.name : ''}` : 'تسجيل جديد', `${carsPick}${secPick}${fields}${photos}
    ${s ? `<div class="btn-row" style="margin-top:6px">${editing ? `<button class="btn danger" id="eDel">حذف</button>` : ''}<button class="btn amber" id="eSave" style="flex:2">${editing ? 'حفظ التعديل' : 'حفظ'}</button></div>` : ''}`, b => {
    if (b.querySelector('#phList')) {
      redrawPhotos();
      b.querySelector('#phCam').onclick = () => { const i = $('#photoCam'); i.value = ''; i.click(); };
      b.querySelector('#phPick').onclick = () => { const i = $('#photoPick'); i.value = ''; i.click(); };
    }
    b.querySelectorAll('[data-car-i]').forEach(x => x.onclick = () => { const t = S.cars[+x.dataset.carI]; if (!t) return; collect(b, s); draft.carId = t.id; draft.v = {}; S.activeCar = t.id; save(); render(); renderEntry(); });
    b.querySelectorAll('[data-sec-i]').forEach(x => x.onclick = () => { const t = c.sections[+x.dataset.secI]; if (!t) return; collect(b, s); draft.secId = t.id; draft.v = {}; renderEntry(); });
    b.querySelectorAll('input[data-ci]').forEach(i => i.oninput = () => { collect(b, s); updateCalc(b, c, s); });
    updateCalc(b, c, s);
    if (b.querySelector('#eSave')) b.querySelector('#eSave').onclick = () => saveEntry(b, c, s);
    if (b.querySelector('#eDel')) b.querySelector('#eDel').onclick = () => confirmBox('حذف هذا التسجيل نهائياً؟', 'حذف', () => { s.rows = s.rows.filter(r => r.id !== draft.rowId); save(); render(); toast('حُذف التسجيل'); });
    const first = b.querySelector('input[data-ci]'); if (first && !editing && !('ontouchstart' in window)) first.focus();
  });
}
function collect(b, s) { if (!s) return; b.querySelectorAll('input[data-ci]').forEach(i => { const col = s.columns[+i.dataset.ci]; if (col) draft.v[col.id] = i.value.trim(); }); }
function updateCalc(b, c, s) {
  if (!s) return;
  const o = colOf(s, 'odometer'); const el = b.querySelector('#calcDist');
  if (!o) { if (el) el.textContent = 'يحتاج بنداً من نوع العداد'; return; }
  const v = toNum(draft.v[o.id]);
  const prevRows = chrono(s).filter(r => r.id !== draft.rowId && toNum(r.v[o.id]) != null && (v == null || toNum(r.v[o.id]) < v));
  const prev = prevRows.length ? toNum(prevRows[prevRows.length - 1].v[o.id]) : null;
  if (el) el.innerHTML = v != null && prev != null ? `${km(v - prev)} <small style="color:var(--muted);font-weight:400">منذ آخر ${esc(s.name)} (${nf(prev)})</small>` : '—';
  const h = b.querySelector('#odoHint');
  if (h) { const cur = currentOdo(c); const low = v != null && cur && v < cur && !draft.rowId; h.classList.toggle('warn', !!low); h.textContent = low ? `القراءة أقل من آخر قراءة معروفة (${nf(cur)} كم) — تأكد منها` : `آخر قراءة معروفة: ${nf(cur)} كم`; }
}
function saveEntry(b, c, s) {
  collect(b, s);
  const v = {};
  for (const col of s.columns) {
    if (col.type === 'distance') continue;
    let x = has(draft.v, col.id) ? String(draft.v[col.id]) : '';
    if (['number', 'money', 'odometer'].includes(col.type) && x !== '') { const n = toNum(x); if (n == null) { toast(`«${col.label}» يجب أن يكون رقماً`); return; } x = n; }
    else x = cleanValue(col.type, x);
    v[col.id] = x;
  }
  const p = cleanPhotoIds(draft.p);
  if (!Object.values(v).some(x => x !== '') && !p.length) { toast('أدخل بياناً واحداً على الأقل'); return; }
  if (photoBusy) { toast('انتظر حتى تُجهَّز الصورة'); return; }
  if (draft.rowId) { const r = s.rows.find(x => x.id === draft.rowId); if (r) { r.v = v; if (p.length) r.p = p; else delete r.p; } }
  else {
    if (s.rows.length >= LIMITS.rows) { toast('وصل القسم إلى الحد الأقصى للتسجيلات'); return; }
    const row = { id: uid(), t: Date.now(), v }; if (p.length) row.p = p;
    s.rows.push(row);
  }
  activeSec = s.id; save(); closeSheet(); render(); toast(draft.rowId ? 'حُفظ التعديل' : `سُجّل في ${s.name} — ${c.name}`);
}

/* ---------- odometer quick update ---------- */
function openOdoUpdate() {
  const c = car(), cur = currentOdo(c);
  openSheet('تحديث العداد', `<label class="field"><span>قراءة العداد الآن (كم)</span><input id="ou" class="num" type="text" inputmode="numeric" value="" placeholder="${cur ? nf(cur) : ''}"><div class="hint">آخر قراءة: ${nf(cur)} كم. تُستخدم لحساب موعد تبديل الزيت القادم.</div></label>
    <button class="btn amber wide" id="ouSave">حفظ القراءة</button>`, b => {
    const i = b.querySelector('#ou'); setTimeout(() => i.focus(), 250);
    b.querySelector('#ouSave').onclick = () => { const n = toNum(i.value); if (n == null || n < 0) { toast('أدخل رقم العداد'); return; } c.odoNow = { value: n, date: todayISO() }; save(); closeSheet(); render(); toast('حُدّث العداد'); };
  });
}

/* ---------- car switcher & editor ---------- */
function openCarSwitcher() {
  openSheet('سياراتي', `<div class="list">${S.cars.map((x, i) => `<button class="list-item" data-car-i="${i}"><span class="ic" style="background:${esc(x.color)};color:#fff">${x.id === S.activeCar ? '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round"><path d="M5 12l5 5 9-10"/></svg>' : ''}</span>
    <span class="grow"><b style="font-weight:500">${esc(x.name)}</b><small>${nf(currentOdo(x))} كم · ${x.sections.reduce((a, s) => a + s.rows.length, 0)} تسجيل</small></span></button>`).join('')}</div>
    <div class="btn-row" style="margin-top:14px"><button class="btn ghost" id="csEdit">تعديل السيارة الحالية</button><button class="btn" id="csAdd">إضافة سيارة</button></div>`, b => {
    b.querySelectorAll('[data-car-i]').forEach(x => x.onclick = () => { const t = S.cars[+x.dataset.carI]; if (!t) return; S.activeCar = t.id; save(); closeSheet(); rolledFor = null; render(); });
    b.querySelector('#csAdd').onclick = () => openCarEditor(null);
    b.querySelector('#csEdit').onclick = () => openCarEditor(car().id);
  });
}
function openCarEditor(id) {
  const c = id ? S.cars.find(x => x.id === id) : null;
  let color = c ? c.color : CAR_COLORS[S.cars.length % CAR_COLORS.length];
  const tmplOpts = S.cars.length ? `<label class="field"><span>البنود والأقسام</span><select id="ceTmpl"><option value="">الأقسام الافتراضية (الزيوت والصيانة)</option>${S.cars.map((x, i) => `<option value="${i}" ${x.id === S.activeCar ? 'selected' : ''}>نفس أقسام ${esc(x.name)}</option>`).join('')}</select></label>` : '';
  openSheet(c ? 'تعديل السيارة' : 'سيارة جديدة', `
    <label class="field"><span>اسم السيارة</span><input id="ceName" type="text" maxlength="${LIMITS.label}" value="${esc(c ? c.name : '')}" placeholder="مثلاً: F150 أو سيارة البيت"></label>
    <div class="field"><span>اللون</span><div class="swatches">${CAR_COLORS.map((k, i) => `<button class="swatch" data-ci="${i}" style="background:${k}" aria-label="لون" aria-pressed="${k === color}"></button>`).join('')}</div></div>
    ${c ? '' : tmplOpts}
    <button class="btn amber wide" id="ceSave">${c ? 'حفظ' : 'إضافة السيارة'}</button>
    ${c ? `<button class="btn ghost wide" id="cePdf" style="margin-top:10px">تقرير PDF لهذه السيارة</button>
    <button class="btn danger wide" id="ceDel" style="margin-top:10px">حذف السيارة وكل سجلاتها</button>` : ''}`, b => {
    if (b.querySelector('#cePdf')) b.querySelector('#cePdf').onclick = () => openBackup(c.id);
    b.querySelectorAll('.swatch').forEach(s => s.onclick = () => { color = CAR_COLORS[+s.dataset.ci] || CAR_COLORS[0]; b.querySelectorAll('.swatch').forEach(x => x.setAttribute('aria-pressed', x === s)); });
    b.querySelector('#ceSave').onclick = () => {
      const name = str(b.querySelector('#ceName').value, LIMITS.label).trim(); if (!name) { toast('اكتب اسم السيارة'); return; }
      if (c) { c.name = name; c.color = color; }
      else {
        if (S.cars.length >= LIMITS.cars) { toast('وصلت إلى الحد الأقصى لعدد السيارات'); return; }
        const t = b.querySelector('#ceTmpl'); const from = t && t.value !== '' ? S.cars[+t.value] : null;
        const n = newCar(name, from); n.color = color; S.cars.push(n); S.activeCar = n.id;
      }
      save(); closeSheet(); rolledFor = null; render(); toast(c ? 'حُفظت السيارة' : `أُضيفت ${name}`);
    };
    if (b.querySelector('#ceDel')) b.querySelector('#ceDel').onclick = () => confirmBox(`حذف «${c.name}» وكل سجلاتها؟ احفظ نسخة احتياطية أولاً إن لم تفعل.`, 'حذف السيارة', () => {
      S.cars = S.cars.filter(x => x.id !== c.id); S.activeCar = S.cars[0] ? S.cars[0].id : null; save(); go('home'); toast('حُذفت السيارة');
    });
  });
}

/* ---------- item (column / card field) list editor ---------- */
function colEditRow(col, i, n, allowDistance) {
  const types = Object.entries(TYPES).filter(([k]) => allowDistance || k !== 'distance');
  return `<div class="col-edit" data-i="${i}">
    <input data-k="label" maxlength="${LIMITS.label}" value="${esc(col.label)}" aria-label="اسم البند">
    <select data-k="type" aria-label="نوع البند">${types.map(([k, l]) => `<option value="${k}" ${k === col.type ? 'selected' : ''}>${l}</option>`).join('')}</select>
    <div class="tools">
      <button data-a="up" ${i === 0 ? 'disabled style="opacity:.35"' : ''} aria-label="أعلى"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M6 15l6-6 6 6"/></svg></button>
      <button data-a="down" ${i === n - 1 ? 'disabled style="opacity:.35"' : ''} aria-label="أسفل"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M6 9l6 6 6-6"/></svg></button>
      <button data-a="del" class="del" aria-label="حذف"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/></svg></button>
    </div></div>`;
}
/* reads the inputs back into `list` (same order as on screen) */
function readItems(b, list) {
  b.querySelectorAll('.col-edit').forEach(el => {
    const f = list[+el.dataset.i]; if (!f) return;
    f.label = str(el.querySelector('[data-k=label]').value, LIMITS.label).trim() || f.label;
    const t = el.querySelector('[data-k=type]').value; if (has(TYPES, t)) f.type = t;
  });
}
/* wires up/down/delete: sync() saves the screen into the data first, then the change, then redraw() */
function bindItems(b, list, sync, redraw) {
  b.querySelectorAll('.col-edit').forEach(el => {
    const i = +el.dataset.i;
    el.querySelectorAll('[data-a]').forEach(btn => btn.onclick = () => {
      sync();
      const a = btn.dataset.a;
      if (a === 'up' && i > 0) [list[i - 1], list[i]] = [list[i], list[i - 1]];
      else if (a === 'down' && i < list.length - 1) [list[i + 1], list[i]] = [list[i], list[i + 1]];
      else if (a === 'del') { if (list.length <= 1) { toast('يجب أن يبقى بند واحد على الأقل'); return; } list.splice(i, 1); }
      redraw();
    });
  });
}

/* ---------- card fields editor ---------- */
function openCardEditor(structure) {
  const c = car();
  const work = clone(c.cardFields);
  const draw = b => {
    if (structure) {
      b.innerHTML = `
      <p style="color:var(--muted);margin:0 0 14px;font-size:14px">غيّر أسماء البنود وأنواعها، أو أضف بنوداً جديدة مثل رقم الشاصي أو تاريخ الترخيص.</p>
      ${work.map((f, i) => colEditRow(f, i, work.length, false)).join('')}
      <button class="btn ghost wide" id="cfAdd">+ إضافة بند</button>
      <label class="check"><input type="checkbox" id="cfAll" ${S.cars.length > 1 ? '' : 'disabled'}> طبّق نفس البنود على كل السيارات</label>
      <button class="btn amber wide" id="cfSave">حفظ البنود</button>`;
      const sync = () => readItems(b, work);
      bindItems(b, work, sync, () => draw(b));
      b.querySelector('#cfAdd').onclick = () => { sync(); if (work.length >= LIMITS.cardFields) return; work.push({ id: uid(), label: 'بند جديد', type: 'text', value: '' }); draw(b); };
      b.querySelector('#cfSave').onclick = () => {
        sync();
        const plans = [planCard(c, work)];
        if (b.querySelector('#cfAll').checked) S.cars.forEach(x => { if (x !== c) plans.push(planCard(x, work)); });
        runPlans(plans, 'حُفظت البنود', () => openSheet('بنود بطاقة السيارة', '', draw));
      };
    } else {
      b.innerHTML = `${work.map((f, i) => `<label class="field"><span>${esc(f.label)}</span><input data-i="${i}" maxlength="${LIMITS.text}" ${f.type === 'date' ? 'type="date"' : f.type === 'text' ? 'type="text"' : 'type="text" inputmode="decimal" class="num"'} value="${esc(f.value)}"></label>`).join('')}
      <button class="btn amber wide" id="cfSave">حفظ</button>
      <button class="btn ghost wide" id="cfStruct" style="margin-top:10px">تعديل البنود (إضافة / حذف / إعادة تسمية)</button>`;
      b.querySelector('#cfSave').onclick = () => {
        b.querySelectorAll('input[data-i]').forEach(i => { const f = work[+i.dataset.i]; if (f) f.value = cleanValue(f.type, i.value.trim()); });
        c.cardFields = work; save(); closeSheet(); render(); toast('حُفظت بطاقة السيارة');
      };
      b.querySelector('#cfStruct').onclick = () => openCardEditor(true);
    }
  };
  openSheet(structure ? 'بنود بطاقة السيارة' : 'بطاقة السيارة', '', draw);
}

/* ---------- section editor ---------- */
function openAddSection() {
  const c = car();
  openSheet('إضافة قسم', `<p style="color:var(--muted);margin:0 0 6px;font-size:14px">اختر قالباً جاهزاً، ثم عدّل بنوده كما تريد.</p>
    <div class="list">${Object.keys(SECTION_TEMPLATES).map(k => { const t = SECTION_TEMPLATES[k](); const exists = c.sections.some(s => s.id === t.id); return `<button class="list-item" data-t="${k}" ${exists ? 'disabled style="opacity:.45"' : ''}><span class="ic">${icon(t.icon)}</span><span class="grow"><b style="font-weight:500">${TEMPLATE_NAMES[k]}</b><small>${exists ? 'موجود' : t.columns.map(x => esc(x.label)).join('، ')}</small></span></button>`; }).join('')}</div>`, b => {
    b.querySelectorAll('[data-t]').forEach(x => x.onclick = () => {
      if (!has(SECTION_TEMPLATES, x.dataset.t)) return;
      if (c.sections.length >= LIMITS.sections) { toast('وصلت إلى الحد الأقصى لعدد الأقسام'); return; }
      const s = SECTION_TEMPLATES[x.dataset.t](); c.sections.push(s); activeSec = s.id; save(); openSectionEditor(s.id, true);
    });
  });
}
/* ---------- applying edited items to a car (keeps existing values) ---------- */
const norm = v => String(v ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
/* pairs each new item with an old one: same id first, then same name */
function matchItems(next, old, key) {
  const used = new Set(), m = new Map();
  for (const n of next) { const o = old.find(x => !used.has(x) && x.id === n.id); if (o) { used.add(o); m.set(n.id, o); } }
  for (const n of next) { if (m.has(n.id)) continue; const o = old.find(x => !used.has(x) && norm(x[key]) === norm(n[key])); if (o) { used.add(o); m.set(n.id, o); } }
  return { m, dropped: old.filter(x => !used.has(x)) };
}
const filled = v => v !== '' && v != null;
function planCard(target, fields) {
  const { m, dropped } = matchItems(fields, target.cardFields, 'label');
  const next = fields.map(f => { const o = m.get(f.id); return { id: f.id, label: f.label, type: f.type, value: o ? cleanValue(f.type, o.value) : '' }; });
  return { lost: dropped.filter(f => filled(f.value)).length, apply: () => { target.cardFields = next; } };
}
/* new columns for a section: each row's values move to the matching column; values of removed columns go */
function planSection(sec, columns, meta) {
  const { m, dropped } = matchItems(columns, sec.columns, 'label');
  let lost = 0;
  for (const r of sec.rows) for (const d of dropped) if (filled(r.v[d.id])) lost++;
  return { lost, apply: () => {
    for (const r of sec.rows) {
      const v = {};
      for (const c of columns) { if (c.type === 'distance') continue; const o = m.get(c.id); v[c.id] = o && has(r.v, o.id) ? cleanValue(c.type, r.v[o.id]) : ''; }
      r.v = v;
    }
    Object.assign(sec, meta, { columns: clone(columns) });
  } };
}
function runPlans(plans, done, reopen) {
  const lost = plans.reduce((a, p) => a + p.lost, 0);
  const run = () => { plans.forEach(p => p.apply()); save(); closeSheet(); render(); toast(done); };
  if (lost) confirmBox(`سيُحذف ${lost} ${lost === 1 ? 'قيمة مسجّلة' : 'من القيم المسجّلة'} في بنود لم تعد موجودة. متابعة؟`, 'حفظ وحذف القيم', run, true, reopen);
  else run();
}
/* reminder fields of the section editor: by distance (needs an odometer item) and by time (needs a date item) */
function remindHTML(work) {
  const hasOdo = work.columns.some(x => x.type === 'odometer'), hasDate = work.columns.some(x => x.type === 'date');
  if (!hasOdo && !hasDate) return '';
  const both = hasOdo && hasDate;
  return `<div class="sub-h">التذكير بالموعد القادم</div><div class="two">
    ${hasOdo ? `<label class="field"><span>كل (كم)</span><input id="seRem" class="num" type="text" inputmode="numeric" value="${esc(String(Math.round(work.reminderKm || 0)))}"></label>` : ''}
    ${hasDate ? `<label class="field"><span>${hasOdo ? 'أو كل (شهر)' : 'كل (شهر)'}</span><input id="seRemM" class="num" type="text" inputmode="numeric" value="${esc(String(work.reminderMonths || 0))}"></label>` : ''}
    </div><p class="hint-p">0 لإيقافه.${both ? ' إن حدّدت الاثنين فالموعد بأيهما أسبق، كتبديل الزيت كل 7,500 كم أو 6 أشهر.' : ''} يظهر في الصفحة الرئيسية${isNative ? '، ويصلك إشعار قبله' : ''}.</p>`;
}
function openSectionEditor(id, isNew) {
  const c = car(); const s = c.sections.find(x => x.id === id); if (!s) return;
  const work = clone(s);
  const iconKeys = Object.keys(ICONS);
  const draw = b => {
    b.innerHTML = `
      <label class="field"><span>اسم القسم</span><input id="seName" maxlength="${LIMITS.label}" value="${esc(work.name)}"></label>
      <label class="field"><span>ملاحظة ثابتة (اختياري)</span><input id="seNote" maxlength="200" value="${esc(work.note || '')}" placeholder="مثلاً: Motorcraft 5W-30"></label>
      <div class="field"><span>الأيقونة</span><div class="swatches">${iconKeys.map(k => `<button class="swatch" data-ic="${k}" aria-pressed="${k === work.icon}" style="background:var(--bg);display:grid;place-items:center;color:var(--steel)">${icon(k, 18)}</button>`).join('')}</div></div>
      ${remindHTML(work)}
      <div class="sub-h">البنود (أعمدة الجدول)</div>
      ${work.columns.map((f, i) => colEditRow(f, i, work.columns.length, true)).join('')}
      <button class="btn ghost wide" id="seAdd">+ إضافة بند</button>
      <label class="check"><input type="checkbox" id="seAll" ${S.cars.length > 1 ? '' : 'disabled'}> طبّق هذا القسم وبنوده على كل السيارات</label>
      <button class="btn amber wide" id="seSave">حفظ القسم</button>
      <button class="btn danger wide" id="seDel" style="margin-top:10px">حذف القسم وسجلاته</button>`;
    const sync = () => {
      work.name = str(b.querySelector('#seName').value, LIMITS.label).trim() || work.name;
      work.note = str(b.querySelector('#seNote').value, 200).trim();
      const r = b.querySelector('#seRem'); if (r) work.reminderKm = Math.min(1e6, Math.max(0, toNum(r.value) || 0));
      const rm = b.querySelector('#seRemM'); if (rm) work.reminderMonths = Math.min(120, Math.max(0, Math.round(toNum(rm.value) || 0)));
      readItems(b, work.columns);
    };
    bindItems(b, work.columns, sync, () => draw(b));
    b.querySelectorAll('[data-ic]').forEach(x => x.onclick = () => { sync(); if (has(ICONS, x.dataset.ic)) work.icon = x.dataset.ic; draw(b); });
    b.querySelector('#seAdd').onclick = () => { sync(); if (work.columns.length >= LIMITS.columns) return; work.columns.push({ id: uid(), label: 'بند جديد', type: 'text' }); draw(b); };
    b.querySelector('#seSave').onclick = () => {
      sync();
      if (work.columns.some(x => x.type === 'distance') && !work.columns.some(x => x.type === 'odometer')) { toast('بند المسافة يحتاج بنداً من نوع العداد'); return; }
      const meta = { name: work.name, icon: work.icon, reminderKm: work.reminderKm, reminderMonths: work.reminderMonths || 0 };
      const plans = [planSection(s, work.columns, { ...meta, note: work.note })];
      if (b.querySelector('#seAll').checked) S.cars.forEach(x => {
        if (x === c) return;
        const t = x.sections.find(y => y.id === s.id) || x.sections.find(y => norm(y.name) === norm(work.name) || norm(y.name) === norm(s.name));
        if (t) plans.push(planSection(t, work.columns, meta));
        else if (x.sections.length < LIMITS.sections) plans.push({ lost: 0, apply: () => x.sections.push({ id: s.id, ...meta, note: '', columns: clone(work.columns), rows: [] }) });
      });
      runPlans(plans, 'حُفظ القسم', () => openSheet('تعديل القسم', '', draw));
    };
    b.querySelector('#seDel').onclick = () => confirmBox(`حذف قسم «${s.name}» و${s.rows.length} تسجيل فيه؟`, 'حذف القسم', () => { c.sections = c.sections.filter(x => x !== s); save(); render(); toast('حُذف القسم'); });
  };
  openSheet(isNew ? 'قسم جديد' : 'تعديل القسم', '', draw);
}

/* ---------- stats ---------- */
function renderStats() {
  const c = car();
  const bySec = c.sections.map(s => ({ s, t: sectionTotal(s) })).filter(x => x.s.columns.some(k => k.type === 'money'));
  const total = bySec.reduce((a, x) => a + x.t, 0);
  const years = {};
  for (const s of c.sections) { const d = colOf(s, 'date'); for (const col of s.columns) if (col.type === 'money') for (const r of s.rows) { const dv = d ? String(r.v[d.id] || '') : ''; const y = ISO_RE.test(dv) ? dv.slice(0, 4) : 'بدون تاريخ'; years[y] = (years[y] || 0) + (toNum(r.v[col.id]) || 0); } }
  const ys = Object.entries(years).sort((a, b) => b[0].localeCompare(a[0]));
  const ymax = Math.max(1, ...ys.map(x => x[1]));
  const smax = Math.max(1, ...bySec.map(x => x.t));
  const oil = c.sections.find(s => colOf(s, 'distance'));
  let avgGap = null;
  if (oil) { const d = [...distances(oil).values()].filter(x => x > 0); if (d.length) avgGap = d.reduce((a, b) => a + b, 0) / d.length; }
  const price = c.cardFields.find(f => f.type === 'money' && toNum(f.value));
  const pct = (a, b) => Math.max(0, Math.min(100, a / b * 100));
  $('#v-stats').innerHTML = `
    <div class="sec-head"><div><h1>المصاريف</h1><p>${esc(c.name)}</p></div><div class="sec-total"><small>الإجمالي</small><b>${money(total)}</b></div></div>
    <div class="block"><h2>حسب القسم</h2><div class="bars">${bySec.length ? bySec.map(x => `<div class="bar-row" style="grid-template-columns:96px 1fr auto"><span>${esc(x.s.name)}</span><span class="bar-track"><span class="bar-fill" style="display:block;width:${pct(x.t, smax)}%"></span></span><span>${money(x.t)}</span></div>`).join('') : '<div class="empty">لا مصاريف بعد</div>'}</div></div>
    <div class="block"><h2>حسب السنة</h2><div class="bars">${ys.length ? ys.map(([y, t]) => `<div class="bar-row"><span class="num">${esc(y)}</span><span class="bar-track"><span class="bar-fill" style="display:block;width:${pct(t, ymax)}%"></span></span><span>${money(t)}</span></div>`).join('') : '<div class="empty">لا مصاريف بعد</div>'}</div></div>
    <div class="block"><h2>أرقام</h2><dl class="kv">
      ${avgGap ? `<dt>متوسط المسافة بين ${esc(oil.name)}</dt><dd>${km(Math.round(avgGap))}</dd>` : ''}
      <dt>العداد الحالي</dt><dd>${km(currentOdo(c))}</dd>
      <dt>المسافة منذ أول قراءة</dt><dd>${km(currentOdo(c) - firstOdo(c))}</dd>
      ${price ? `<dt>${esc(price.label)}</dt><dd>${money(price.value)}</dd><dt>السعر + المصاريف</dt><dd>${money(toNum(price.value) + total)}</dd>` : ''}
    </dl></div>
    ${S.cars.length > 1 ? `<div class="block"><h2>كل السيارات</h2><dl class="kv">${S.cars.map(x => `<dt>${esc(x.name)}</dt><dd>${money(x.sections.reduce((a, s) => a + sectionTotal(s), 0))}</dd>`).join('')}</dl></div>` : ''}`;
}

/* ---------- settings ---------- */
function renderSettings() {
  const c = car();
  const last = S.lastBackup ? new Date(S.lastBackup) : null;
  $('#v-settings').innerHTML = `
    <div class="sec-head"><div><h1>الإعدادات</h1></div></div>
    <div class="block"><h2>التصدير والنسخة الاحتياطية</h2>
      <p style="margin:0 0 14px;color:var(--muted);font-size:14px">تقرير PDF لكل سيارة على حدة، وملف إكسل منسّق بكل السيارات والأقسام يُحفظ في الدرايف أو أي مكان من قائمة المشاركة ويمكن استرجاعه في أي وقت.${last && !isNaN(last) ? `<br>آخر نسخة احتياطية: ${dateHTML(last.toISOString().slice(0, 10))}` : ''}</p>
      <div class="btn-row"><button class="btn amber" id="sPdf">تقرير PDF</button><button class="btn" id="sExp">تصدير إكسل</button></div>
      <button class="btn ghost wide" id="sImp" style="margin-top:10px">استيراد / استرجاع</button>
    </div>
    <div class="block"><h2>التذكيرات</h2>
      ${isNative ? `<label class="check"><input type="checkbox" id="sNotify" ${S.settings.notify && notifPerm === 'granted' ? 'checked' : ''}> إشعارات على الهاتف قبل المواعيد</label>
      <p class="hint-p" id="sNotifyState">${esc(notifStateText())}</p>`
      : `<p class="hint-p" style="margin-top:0">في نسخة الويب لا يستطيع المتصفح إرسال تذكير والتطبيق مغلق، فأضف المواعيد إلى تقويم هاتفك لتصلك تنبيهاته في وقتها.</p>`}
      <label class="field"><span>وقت التذكير</span><select id="sHour">${Array.from({ length: 17 }, (_, i) => i + 6).map(h => `<option value="${h}" ${h === S.settings.notifyHour ? 'selected' : ''}>${hourLabel(h)}</option>`).join('')}</select></label>
      <label class="field"><span>التذكير بانتهاء التأمين والترخيص</span><select id="sLead">${LEAD_DAYS.map(d => `<option value="${d}" ${d === S.settings.expiryLead ? 'selected' : ''}>قبل الانتهاء بـ ${esc(nDays(d))}</option>`).join('')}</select></label>
      <label class="field"><span>تذكير بتحديث العداد (لدقة مواعيد الكيلومترات)</span><select id="sNudge">${NUDGE_DAYS.map(d => `<option value="${d}" ${d === S.settings.odoNudge ? 'selected' : ''}>${d ? 'إن لم تُحدَّث قراءته منذ ' + esc(nDays(d)) : 'بدون'}</option>`).join('')}</select></label>
      <div class="btn-row"><button class="btn ghost" id="sUpcoming">المواعيد القادمة</button><button class="btn ghost" id="sCal">${uiIcon('cal', 18)} إلى التقويم</button></div>
    </div>
    <div class="block"><h2>أقسام ${esc(c.name)} وبنودها</h2>
      <div class="list">
        <button class="list-item" id="sCard"><span class="ic">${icon('doc')}</span><span class="grow"><b style="font-weight:500">بطاقة السيارة</b><small>${c.cardFields.map(f => esc(f.label)).join('، ')}</small></span></button>
        ${c.sections.map((s, i) => `<button class="list-item" data-sec-i="${i}"><span class="ic">${icon(s.icon)}</span><span class="grow"><b style="font-weight:500">${esc(s.name)}</b><small>${s.columns.map(f => esc(f.label)).join('، ')}</small></span></button>`).join('')}
      </div>
      <button class="btn ghost wide" id="sAddSec" style="margin-top:10px">+ إضافة قسم</button>
    </div>
    <div class="block"><h2>السيارات</h2>
      <div class="list">${S.cars.map((x, i) => `<button class="list-item" data-car-i="${i}"><span class="ic" style="background:${esc(x.color)}"></span><span class="grow"><b style="font-weight:500">${esc(x.name)}</b><small>${nf(currentOdo(x))} كم</small></span></button>`).join('')}</div>
      <button class="btn ghost wide" id="sAddCar" style="margin-top:10px">+ إضافة سيارة</button>
    </div>
    <div class="block"><h2>عام</h2>
      <label class="field"><span>العملة</span><input id="sCur" maxlength="12" value="${esc(S.settings.currency)}"></label>
      <label class="field"><span>المنازل العشرية للمبالغ</span><select id="sDec">${[0, 2, 3].map(n => `<option value="${n}" ${n === S.settings.decimals ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
      <label class="field"><span>المظهر</span><select id="sTheme">${[['auto', 'حسب الجهاز'], ['light', 'فاتح'], ['dark', 'داكن']].map(([k, l]) => `<option value="${k}" ${k === S.settings.theme ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    </div>
    <p style="text-align:center;color:var(--muted);font-size:13px;margin:22px 0">سيّارتي ${APP_VERSION} · البيانات محفوظة على هذا الجهاز فقط</p>`;
  $('#sExp').onclick = () => openBackup();
  $('#sPdf').onclick = () => openBackup(c.id);
  $('#sImp').onclick = pickImport;
  $('#sCard').onclick = () => openCardEditor(true);
  $('#sAddSec').onclick = openAddSection;
  $('#sAddCar').onclick = () => openCarEditor(null);
  $('#v-settings').querySelectorAll('[data-sec-i]').forEach(b => b.onclick = () => { const s = c.sections[+b.dataset.secI]; if (s) openSectionEditor(s.id); });
  $('#v-settings').querySelectorAll('[data-car-i]').forEach(b => b.onclick = () => { const x = S.cars[+b.dataset.carI]; if (x) openCarEditor(x.id); });
  $('#sCur').onchange = e => { S.settings.currency = str(e.target.value, 12).trim() || 'د.ك'; save(); };
  $('#sDec').onchange = e => { S.settings = sanitizeSettings({ ...S.settings, decimals: +e.target.value }); save(); };
  $('#sTheme').onchange = e => { S.settings = sanitizeSettings({ ...S.settings, theme: e.target.value }); save(); applyTheme(); };
  const setNum = (k, v) => { S.settings = sanitizeSettings({ ...S.settings, [k]: Number(v) }); save(); };
  $('#sHour').onchange = e => setNum('notifyHour', e.target.value);
  $('#sLead').onchange = e => setNum('expiryLead', e.target.value);
  $('#sNudge').onchange = e => setNum('odoNudge', e.target.value);
  $('#sUpcoming').onclick = openUpcoming;
  $('#sCal').onclick = () => openCalendarSheet();
  if ($('#sNotify')) $('#sNotify').onchange = e => { if (e.target.checked) enableNotifications(false); else { S.settings.notify = false; save(); renderSettings(); } };
}
const hourLabel = h => (h === 12 ? '12 ظهراً' : h < 12 ? `${h} صباحاً` : `${h - 12} ${h < 18 ? 'عصراً' : 'مساءً'}`);

/* ======================================================================
   Reminders: the list of upcoming dates, phone notifications (Android app)
   and a calendar file (.ics) for the phone's own calendar (iPhone / web).
   ====================================================================== */
function allDue() {
  const items = [];
  for (const c of S.cars) items.push(...dueItems(c));
  return items.sort((a, b) => ((a.date || '9999') < (b.date || '9999') ? -1 : (a.date || '9999') > (b.date || '9999') ? 1 : 0));
}
function dueRowHTML(it) {
  const ci = S.cars.indexOf(it.c);
  let title, sub, late;
  if (it.type === 'svc') {
    const x = it.x;
    title = x.sec.name; late = x.late;
    sub = esc(svcWhen(x)) + (x.late ? ' · متأخر' : x.due ? ` · ${esc(inDays(daysUntil(x.due)))}${x.est ? ' تقريباً' : ''}` : x.stale ? ' · حدّث العداد لتقدير الموعد' : '');
  } else {
    const e = it.e;
    title = e.title; late = e.days < 0;
    sub = `${e.days >= 0 ? 'ينتهي' : 'انتهى'} ${dateHTML(e.date)} · ${esc(inDays(e.days))}`;
  }
  return `<button class="list-item" data-car-i="${ci}"><span class="ic" style="${late ? 'color:var(--bad)' : ''}">${it.type === 'svc' ? icon(it.x.sec.icon) : icon('shield')}</span>
    <span class="grow"><b style="font-weight:500">${esc(title)}${S.cars.length > 1 ? ` <span class="car-tag"><span class="dot" style="background:${esc(it.c.color)}"></span>${esc(it.c.name)}</span>` : ''}</b><small>${sub}</small></span></button>`;
}
function openUpcoming() {
  const items = allDue();
  openSheet('المواعيد القادمة', `
    ${items.length ? `<div class="list">${items.map(dueRowHTML).join('')}</div>` : `<div class="empty"><b>لا توجد مواعيد بعد</b>حدّد في قسم الزيوت مثلاً «التذكير كل كذا كم أو شهر» (الإعدادات ← القسم)، أو أضف تاريخ انتهاء التأمين.</div>`}
    ${isNative ? `<p class="hint-p">${esc(notifStateText())}</p>` : ''}
    <button class="btn ${isNative ? 'ghost' : 'amber'} wide" id="upCal" style="margin-top:12px">${uiIcon('cal', 18)} إضافة المواعيد إلى التقويم</button>`, b => {
    b.querySelectorAll('[data-car-i]').forEach(x => x.onclick = () => { const t = S.cars[+x.dataset.carI]; if (!t) return; S.activeCar = t.id; save(); closeSheet(); rolledFor = null; go('home'); });
    b.querySelector('#upCal').onclick = () => openCalendarSheet();
  });
}

/* ---------- phone notifications (Android app only) ---------- */
let notifPerm = null;            /* 'granted' | 'denied' | 'prompt' | … once known */
let notifTimer = null, notifSig = null;
const notifPlugin = () => (isNative ? plug('LocalNotifications') : null);
function notifStateText() {
  if (!S.settings.notify) return 'الإشعارات متوقفة. تظهر المواعيد هنا وفي الصفحة الرئيسية فقط.';
  if (notifPerm === 'granted') return `تصلك الإشعارات الساعة ${hourLabel(S.settings.notifyHour)}: قبل الموعد بأسبوع، وفي يومه، وبعده بأسبوع إن لم تسجّله.`;
  if (notifPerm === 'denied') return 'الإشعارات مرفوضة في إعدادات الهاتف. فعّلها من: الإعدادات ← التطبيقات ← سيّارتي ← الإشعارات.';
  return 'اسمح للتطبيق بالإشعارات لتصلك التذكيرات.';
}
/* time of an alert: that day at the chosen hour (local time) */
const alertAt = iso => { const [y, m, d] = isoParts(iso); return new Date(y, m - 1, d, S.settings.notifyHour, 0, 0, 0).getTime(); };
/* every notification to schedule, from the data alone (so planning again gives the same list) */
function planAlerts(now = Date.now()) {
  const out = [];
  const push = (key, iso, title, body, carId) => { if (!iso) return; const at = alertAt(iso); if (at > now + 60e3) out.push({ key, at, title, body, carId }); };
  for (const c of S.cars) {
    for (const it of dueItems(c)) {
      const k = `${c.id}:${it.key}`;
      if (it.type === 'svc') {
        const x = it.x;
        if (!x.due) continue;
        /* the date is either the time limit or the day the distance should be reached (an estimate) */
        const at = x.est ? `عند ${nf(x.km.next)} كم` : `في ${fmtDate(x.due)}`;
        if (x.late) {   /* already overdue: one reminder a week after it fell due, if that is still ahead */
          push(k + ':late', addDaysISO(x.due, 7), `فات موعد ${x.sec.name} — ${c.name}`, `كان الموعد ${at}. إن أتممته فسجّله ليُحسب الموعد التالي.`, c.id);
          continue;
        }
        push(k + ':soon', addDaysISO(x.due, -7), `اقترب موعد ${x.sec.name} — ${c.name}`, `بعد أسبوع تقريباً: ${svcWhen(x)}.`, c.id);
        push(k + ':due', x.due, `حان موعد ${x.sec.name} — ${c.name}`,
          x.est ? `حسب معدل قيادتك يصل العداد اليوم تقريباً إلى ${nf(x.km.next)} كم. سجّل العملية بعد إتمامها.` : `اليوم موعد ${x.sec.name} (${fmtDate(x.due)}). سجّل العملية بعد إتمامها.`, c.id);
        push(k + ':late', addDaysISO(x.due, 7), `فات موعد ${x.sec.name} — ${c.name}`,
          `كان الموعد ${at}. إن أتممته فسجّله ليُحسب الموعد التالي${x.est ? '، أو حدّث قراءة العداد' : ''}.`, c.id);
      } else {
        const e = it.e, lead = S.settings.expiryLead;
        push(k + ':lead', addDaysISO(e.date, -lead), `${e.title} — ${c.name}`, `ينتهي بعد ${nDays(lead)}، في ${fmtDate(e.date)}.`, c.id);
        if (lead > 7) push(k + ':week', addDaysISO(e.date, -7), `${e.title} — ${c.name}`, `ينتهي بعد أسبوع، في ${fmtDate(e.date)}.`, c.id);
        push(k + ':day', e.date, `${e.title} — ${c.name}`, `ينتهي اليوم (${fmtDate(e.date)}).`, c.id);
      }
    }
    /* a fresh odometer reading keeps the distance reminders right: ask every N days without one */
    const N = S.settings.odoNudge, last = lastDated(c);
    if (N > 0 && last && c.sections.some(s => s.reminderKm > 0 && colOf(s, 'odometer'))) {
      let k = 1, d = addDaysISO(last.date, N);
      while (alertAt(d) <= now + 60e3 && k < 400) d = addDaysISO(last.date, N * ++k);
      push(`${c.id}:odo`, d, `حدّث قراءة العداد — ${c.name}`, `آخر قراءة مسجّلة كانت قبل ${nDays(N * k)}. تحديثها يُبقي مواعيد الزيت دقيقة.`, c.id);
    }
  }
  out.sort((a, b) => a.at - b.at);
  const list = out.slice(0, 60), used = new Set();
  for (const a of list) { let id = hashId(a.key); while (used.has(id)) id = (id % 2147483646) + 1; used.add(id); a.id = id; }
  return list;
}
function scheduleNotifSync(delay = 1500) {
  if (!notifPlugin()) return;
  clearTimeout(notifTimer);
  notifTimer = setTimeout(syncNotifications, delay);
}
/* one sync at a time, so an older plan can never be scheduled after a newer one */
let notifRun = Promise.resolve();
function syncNotifications(force) {
  notifRun = notifRun.then(() => runNotifSync(force)).catch(e => console.warn('notifications', e));
  return notifRun;
}
async function runNotifSync(force) {
  const ln = notifPlugin(); if (!ln) return;
  try { notifPerm = (await ln.checkPermissions()).display; } catch (e) { notifPerm = 'denied'; }
  const list = S.settings.notify && notifPerm === 'granted' ? planAlerts() : [];
  const sig = JSON.stringify(list.map(a => [a.id, a.at, a.title, a.body]));
  if (sig === notifSig && !force) return;
  try {
    const pend = await ln.getPending();
    const old = ((pend && pend.notifications) || []).map(n => ({ id: n.id }));
    if (old.length) await ln.cancel({ notifications: old });
    if (list.length) await ln.schedule({ notifications: list.map(a => ({
      id: a.id, title: a.title, body: a.body, largeBody: a.body, channelId: 'reminders',
      smallIcon: 'ic_stat_sayyarati', iconColor: '#E9A23B',
      schedule: { at: new Date(a.at), allowWhileIdle: true }, isExactNotification: false,
      extra: { car: a.carId }
    })) });
    notifSig = sig;
  } catch (e) { console.warn('notifications', e); notifSig = null; }
}
/* asks Android for the notification permission (Android 13+), then schedules */
async function enableNotifications(fromHome) {
  const ln = notifPlugin(); if (!ln) return;
  S.settings.notify = true; save();
  try { notifPerm = (await ln.requestPermissions()).display; } catch (e) { notifPerm = 'denied'; }
  if (notifPerm === 'granted') { toast('ستصلك التذكيرات على الهاتف'); await syncNotifications(true); }
  else {
    if (fromHome) S.settings.notify = false;
    save();
    toast('لم يُسمح بالإشعارات. يمكنك تفعيلها من إعدادات الهاتف ← التطبيقات ← سيّارتي');
  }
  render();
}
async function initNotifications() {
  const ln = notifPlugin(); if (!ln) return;
  try {
    await ln.createChannel({ id: 'reminders', name: 'التذكيرات', description: 'مواعيد الزيت والصيانة وانتهاء التأمين والترخيص', importance: 4, visibility: 0, vibration: true });
  } catch (e) { /* older Android: no channels */ }
  ln.addListener('localNotificationActionPerformed', ev => {
    const id = ev && ev.notification && ev.notification.extra && ev.notification.extra.car;
    if (typeof id === 'string' && S.cars.some(c => c.id === id) && S.activeCar !== id) { S.activeCar = id; save(); rolledFor = null; }
    if (!sheetOpen()) go('home'); else render();
  });
  await syncNotifications(true);
  if (view === 'home' || view === 'settings') render();
}

/* ---------- calendar file (.ics) ---------- */
const ICS_CACHE = 'sayyarati-ics', ICS_PATH = 'sayyarati-reminders.ics';
const isIOS = () => /iP(hone|ad|od)/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
/* text values: line breaks become \n and other control characters spaces, so a value can never start a new line */
const icsText = s => String(s).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r\n|\r|\n/g, '\\n').replace(/[\u0000-\u001F\u007F\u2028\u2029]/g, ' ');
/* lines longer than 75 bytes continue on the next line after a space (RFC 5545), never inside a character */
function icsFold(line) {
  const enc = new TextEncoder(), out = [];
  let cur = '', bytes = 0;
  for (const ch of line) {
    const b = enc.encode(ch).length;
    if (bytes + b > (out.length ? 74 : 75)) { out.push(cur); cur = ''; bytes = 0; }
    cur += ch; bytes += b;
  }
  out.push(cur);
  return out.join('\r\n ');
}
/* alarm `days` before an all-day event, at the reminder hour */
function icsTrigger(days) {
  let mins = days * 1440 - S.settings.notifyHour * 60;
  const before = mins > 0; mins = Math.abs(mins);
  const d = Math.floor(mins / 1440), h = Math.floor((mins % 1440) / 60), m = mins % 60;
  return (before ? '-' : '') + 'P' + (d ? d + 'D' : '') + (h || m || !d ? 'T' + (h ? h + 'H' : '') + (m || (!h && !d) ? m + 'M' : '') : '');
}
function buildIcs(cars) {
  const today = todayISO(), lead = S.settings.expiryLead;
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
  const L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Sayyarati//Car log//AR', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:' + icsText('سيّارتي')];
  let count = 0;
  for (const c of cars) for (const it of dueItems(c)) {
    if (!it.date || it.date < today) continue;
    let summary, desc, alarms;
    if (it.type === 'svc') {
      summary = `موعد ${it.x.sec.name} — ${c.name}`;
      desc = `${svcWhen(it.x)}.${it.x.est ? ' التاريخ تقديري حسب معدل قيادتك.' : ''} من تطبيق سيّارتي.`;
      alarms = [7, 0];
    } else {
      summary = `${EXPIRY_RE.test(it.e.title) ? it.e.title : 'ينتهي: ' + it.e.title} — ${c.name}`;
      desc = `تاريخ الانتهاء ${fmtDate(it.e.date)}. من تطبيق سيّارتي.`;
      alarms = [...new Set([lead, 7, 0])].filter(d => d <= lead);
    }
    /* one id per item and date: the iPhone calendar does not take a new date for an event it already has,
       so a moved date (after an oil change, say) comes as a new event */
    L.push('BEGIN:VEVENT', `UID:${c.id}-${hashId(it.key).toString(36)}-${it.date.replace(/-/g, '')}@sayyarati`, `DTSTAMP:${stamp}`, 'SEQUENCE:0',
      `DTSTART;VALUE=DATE:${it.date.replace(/-/g, '')}`, `DTEND;VALUE=DATE:${addDaysISO(it.date, 1).replace(/-/g, '')}`,
      'SUMMARY:' + icsText(summary), 'DESCRIPTION:' + icsText(desc), 'TRANSP:TRANSPARENT');
    for (const d of alarms) L.push('BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:' + icsText(summary), 'TRIGGER:' + icsTrigger(d), 'END:VALARM');
    L.push('END:VEVENT');
    count++;
  }
  L.push('END:VCALENDAR');
  return { text: L.map(icsFold).join('\r\n') + '\r\n', count };
}
function openCalendarSheet() {
  const { count } = buildIcs(S.cars);
  const ios = !isNative && isIOS();
  openSheet('إضافة المواعيد إلى التقويم', `
    <p style="margin:0 0 12px">${count ? `${count === 1 ? 'موعد قادم واحد' : count === 2 ? 'موعدان قادمان' : `${nf(count)} مواعيد قادمة`} يُضاف إلى تقويم هاتفك، مع تنبيه قبل كل موعد بأسبوع وفي يومه${S.settings.expiryLead > 7 ? `، وقبل انتهاء التأمين والترخيص بـ ${esc(nDays(S.settings.expiryLead))}` : ''}.` : 'لا توجد مواعيد قادمة بتاريخ معروف بعد.'}</p>
    <p class="hint-p">المواعيد المحسوبة بالكيلومترات تقديرية حسب معدل قيادتك. بعد كل تبديل زيت أو تجديد أضف المواعيد مرة أخرى، واحذف من التقويم الموعد القديم إن تغيّر تاريخه.</p>
    ${count ? `<button class="btn amber wide" id="calGo">${ios ? 'إضافة إلى تقويم الآيفون' : isNative ? 'مشاركة ملف التقويم' : 'تنزيل ملف التقويم'}</button>
    ${!isNative ? `<button class="btn ghost wide" id="calShare" style="margin-top:10px">مشاركة الملف</button>` : ''}
    ${ios ? `<p class="hint-p" style="margin-top:12px">إن لم تظهر نافذة «إضافة إلى التقويم»: اختر «مشاركة الملف» ← «حفظ في الملفات»، ثم افتح الملف من تطبيق الملفات.</p>` : ''}` : ''}`, b => {
    const run = mode => exportCalendar(mode).catch(e => { console.error(e); toast('تعذّر إنشاء ملف التقويم'); });
    if (b.querySelector('#calGo')) b.querySelector('#calGo').onclick = () => run(ios ? 'ios' : isNative ? 'share' : 'dl');
    if (b.querySelector('#calShare')) b.querySelector('#calShare').onclick = () => run('share');
  });
}
async function exportCalendar(mode) {
  const { text, count } = buildIcs(S.cars);
  if (!count) { toast('لا توجد مواعيد قادمة'); return false; }
  const blob = new Blob([text], { type: 'text/calendar;charset=utf-8' });
  const name = `مواعيد-سيارتي-${todayISO()}.ics`;
  /* iPhone: Safari adds events to the calendar when it opens a calendar file served by the site; the
     service worker serves this one from the app's own storage (nothing leaves the phone) */
  if (mode === 'ios' && window.caches && navigator.serviceWorker && navigator.serviceWorker.controller) {
    try {
      const cache = await caches.open(ICS_CACHE);
      await cache.put(new URL(ICS_PATH, location.href).href, new Response(blob, { headers: {
        'Content-Type': 'text/calendar; charset=utf-8; method=PUBLISH', 'Content-Disposition': 'inline; filename="sayyarati.ics"', 'Cache-Control': 'no-store' } }));
      location.href = ICS_PATH;
      return true;
    } catch (e) { console.warn('ics via service worker', e); }
  }
  return deliver(blob, name, mode === 'share' ? 'share' : 'dl', 'إضافة المواعيد إلى التقويم');
}

/* ======================================================================
   Invoice photos. A photo is made smaller and saved again as a new JPEG (which also drops the
   camera's hidden data, such as where it was taken), then kept on the device: in the app's own
   files on Android, in the browser's storage (IndexedDB) on the web. Entries keep only photo ids.
   ====================================================================== */
const PHOTO_MAX = 1600, PHOTO_Q = 0.78;
const isJpeg = u8 => !!u8 && u8.length > 8 && u8[0] === 0xFF && u8[1] === 0xD8 && u8[2] === 0xFF;
/* width and height from a JPEG's frame header (null when it is not a well-formed JPEG) */
function jpegSize(u8) {
  if (!isJpeg(u8)) return null;
  let i = 2;
  while (i + 9 < u8.length) {
    if (u8[i] !== 0xFF) return null;
    const m = u8[i + 1];
    if (m === 0xFF) { i++; continue; }
    if (m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
    if (m === 0xD9 || m === 0xDA) return null;            // end / image data before any frame header
    const len = (u8[i + 2] << 8) | u8[i + 3];
    if (len < 2) return null;
    if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
      const h = (u8[i + 5] << 8) | u8[i + 6], w = (u8[i + 7] << 8) | u8[i + 8];
      return w > 0 && h > 0 ? { w, h } : null;
    }
    i += 2 + len;
  }
  return null;
}
async function decodeImage(blob) {
  if (window.createImageBitmap) {
    try { const b = await createImageBitmap(blob, { imageOrientation: 'from-image' }); return { src: b, w: b.width, h: b.height, done: () => b.close && b.close() }; }
    catch (e) { /* older browsers: through an <img> */ }
  }
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image(); img.src = url; await img.decode();
    return { src: img, w: img.naturalWidth, h: img.naturalHeight, done: () => URL.revokeObjectURL(url) };
  } catch (e) { URL.revokeObjectURL(url); throw e; }
}
async function compressPhoto(file) {
  if (!file || !file.size || file.size > 60 * 1024 * 1024) throw new Error('not a usable image');
  const im = await decodeImage(file);
  try {
    if (!im.w || !im.h || im.w * im.h > 150e6) throw new Error('image size');
    const k = Math.min(1, PHOTO_MAX / Math.max(im.w, im.h));
    const w = Math.max(1, Math.round(im.w * k)), h = Math.max(1, Math.round(im.h * k));
    const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
    const g = cv.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, w, h);                 // transparent PNGs get a white page
    g.imageSmoothingEnabled = true; g.imageSmoothingQuality = 'high';
    g.drawImage(im.src, 0, 0, w, h);
    const enc = q => new Promise(r => cv.toBlob(r, 'image/jpeg', q));
    let out = await enc(PHOTO_Q);
    if (out && out.size > 700e3) out = await enc(0.6);
    cv.width = cv.height = 0;
    if (!out || !out.size) throw new Error('encode failed');
    return out;
  } finally { im.done(); }
}
/* writes a file in pieces, so a large file never travels to Android as one huge message */
async function writeChunked(fs, path, blob, directory) {
  const CH = 3 * 1024 * 1024;
  let res = null;
  for (let off = 0; off === 0 || off < blob.size; off += CH) {
    const data = bufToB64(await blob.slice(off, off + CH).arrayBuffer());
    if (off === 0) res = await fs.writeFile({ path, data, directory, recursive: true });
    else await fs.appendFile({ path, data, directory });
  }
  return res;
}
const Photos = (() => {
  const urls = new Map();                 /* object URLs shown in the open sheet (web) */
  let dbp = null, base = null;
  const db = () => dbp || (dbp = new Promise((res, rej) => {
    const r = indexedDB.open('sayyarati-photos', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('photos', { keyPath: 'id' });
    r.onsuccess = () => res(r.result);
    r.onerror = () => { dbp = null; rej(r.error); };
  }));
  const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  const store = async mode => (await db()).transaction('photos', mode).objectStore('photos');
  const fs = () => plug('Filesystem');
  const fileUrl = async id => {
    if (!base) base = window.Capacitor.convertFileSrc((await fs().getUri({ path: 'photos', directory: 'DATA' })).uri);
    return `${base}/${id}.jpg`;
  };
  const api = {
    async put(blob, id = uid()) {
      if (!ID_RE.test(id)) throw new Error('bad photo id');
      if (isNative) await writeChunked(fs(), `photos/${id}.jpg`, blob, 'DATA');
      else await req((await store('readwrite')).put({ id, t: Date.now(), size: blob.size, blob }));
      return id;
    },
    async blob(id) {
      if (!ID_RE.test(String(id))) return null;
      if (isNative) {
        try { const r = await fetch(await fileUrl(id)); return r.ok ? await r.blob() : null; } catch (e) { return null; }
      }
      const rec = await req((await store('readonly')).get(id));
      return rec && rec.blob instanceof Blob ? rec.blob : null;
    },
    async url(id, keep = true) {
      if (!ID_RE.test(String(id))) return null;
      if (isNative) return fileUrl(id);
      if (keep && urls.has(id)) return urls.get(id);
      const b = await api.blob(id); if (!b) return null;
      const u = URL.createObjectURL(b);
      if (keep) urls.set(id, u);
      return u;
    },
    async del(id) {
      if (!ID_RE.test(String(id))) return;
      if (isNative) await fs().deleteFile({ path: `photos/${id}.jpg`, directory: 'DATA' }).catch(() => {});
      else await req((await store('readwrite')).delete(id));
    },
    async list() {
      if (isNative) {
        try {
          const r = await fs().readdir({ path: 'photos', directory: 'DATA' });
          return (r.files || []).map(f => ({ id: String(f.name || '').replace(/\.jpg$/, ''), t: Number(f.mtime) || 0, size: Number(f.size) || 0 })).filter(x => ID_RE.test(x.id));
        } catch (e) { return []; }
      }
      const all = await req((await store('readonly')).getAll());
      return all.filter(x => x && typeof x.id === 'string').map(x => ({ id: x.id, t: x.t || 0, size: x.size || 0 }));
    },
    release() { for (const u of urls.values()) URL.revokeObjectURL(u); urls.clear(); }
  };
  return api;
})();
const releaseUrls = () => Photos.release();
const usedPhotoIds = (cars = S.cars) => { const u = new Set(); for (const c of cars) for (const s of c.sections) for (const r of s.rows) for (const id of r.p || []) u.add(id); return u; };
/* photos no entry uses any more (entry deleted, photo removed, restore replaced the data) */
async function gcPhotos() {
  try {
    const last = Number(localStorage.getItem('sayyarati.photoGC')) || 0;
    if (Date.now() - last < 864e5) return;
    const used = usedPhotoIds();
    for (const p of await Photos.list()) if (!used.has(p.id) && Date.now() - p.t > 6 * 3600e3) await Photos.del(p.id);
    localStorage.setItem('sayyarati.photoGC', String(Date.now()));
  } catch (e) { console.warn('photo cleanup', e); }
}
/* photos of one car (or all), oldest entry first: [{id, c, s, r}] */
function photoItems(cars) {
  const out = [];
  for (const c of cars) for (const s of c.sections) for (const r of chrono(s)) for (const id of r.p || []) out.push({ id, c, s, r });
  return out;
}
/* "الزيوت · تبديل زيت · 01.04.2026 · 12,500 كم" */
function photoCaption(s, r) {
  const t = textCol(s), d = colOf(s, 'date'), o = colOf(s, 'odometer');
  const parts = [s.name];
  if (t && r.v[t.id]) parts.push(String(r.v[t.id]).slice(0, 80));
  if (d && r.v[d.id]) parts.push(fmtDate(r.v[d.id]));
  if (o && toNum(r.v[o.id]) != null) parts.push(`${nf(r.v[o.id])} كم`);
  return parts.join(' · ');
}

/* ---------- photo viewer (above the entry form) ---------- */
let viewer = null;
async function openViewer(ids, i, onDelete) {
  if (!ids.length) return;
  viewer = { ids, i: Math.max(0, Math.min(ids.length - 1, i)), onDelete, url: null, arm: 0 };
  $('#viewer').classList.add('on');
  await showViewerImage();
}
async function showViewerImage() {
  const v = viewer; if (!v) return;
  const img = $('#vImg');
  if (v.url && v.url.startsWith('blob:')) URL.revokeObjectURL(v.url);
  v.url = null; img.removeAttribute('src');
  $('#vBox').classList.remove('missing', 'z');
  const many = v.ids.length > 1;
  $('#vCount').textContent = many ? `${v.i + 1} / ${v.ids.length}` : '';
  $('#vPrev').classList.toggle('hidden', !many); $('#vNext').classList.toggle('hidden', !many);
  $('#vDel').classList.toggle('hidden', !v.onDelete);
  clearTimeout(v.arm); v.arm = 0; $('#vDelT').textContent = 'حذف';
  const u = await Photos.url(v.ids[v.i], false).catch(() => null);
  if (viewer !== v) { if (u && u.startsWith('blob:')) URL.revokeObjectURL(u); return; }
  v.url = u;
  if (u) img.src = u; else $('#vBox').classList.add('missing');
}
function closeViewer() {
  if (!viewer) return;
  clearTimeout(viewer.arm);
  if (viewer.url && viewer.url.startsWith('blob:')) URL.revokeObjectURL(viewer.url);
  viewer = null;
  $('#viewer').classList.remove('on');
  $('#vImg').removeAttribute('src');
}
const viewerOpen = () => !!viewer;
function initViewer() {
  $('#vClose').onclick = closeViewer;
  $('#vImg').onclick = () => $('#vBox').classList.toggle('z');
  $('#vImg').onerror = () => { if (viewer && viewer.url) $('#vBox').classList.add('missing'); };
  $('#vPrev').onclick = () => { if (!viewer) return; viewer.i = (viewer.i - 1 + viewer.ids.length) % viewer.ids.length; showViewerImage(); };
  $('#vNext').onclick = () => { if (!viewer) return; viewer.i = (viewer.i + 1) % viewer.ids.length; showViewerImage(); };
  $('#vShare').onclick = async () => {
    if (!viewer) return;
    const b = await Photos.blob(viewer.ids[viewer.i]);
    if (!b) { toast('الصورة غير موجودة على هذا الجهاز'); return; }
    try { await deliver(b, `فاتورة-${todayISO()}-${viewer.i + 1}.jpg`, 'share', 'مشاركة الصورة'); } catch (e) { toast('تعذّرت مشاركة الصورة'); }
  };
  /* delete asks twice: the first tap arms the button for a few seconds */
  $('#vDel').onclick = () => {
    const v = viewer; if (!v || !v.onDelete) return;
    if (!v.arm) { v.arm = setTimeout(() => { if (viewer === v) { v.arm = 0; $('#vDelT').textContent = 'حذف'; } }, 3500); $('#vDelT').textContent = 'تأكيد الحذف'; return; }
    clearTimeout(v.arm); v.arm = 0;
    const id = v.ids[v.i]; v.onDelete(id); v.ids = v.ids.filter(x => x !== id);
    if (!v.ids.length) { closeViewer(); return; }
    v.i = Math.min(v.i, v.ids.length - 1); showViewerImage();
  };
  $('#photoCam').onchange = e => addPhotoFiles(e.target.files);
  $('#photoPick').onchange = e => addPhotoFiles(e.target.files);
}

/* ======================================================================
   Excel — export / import (ExcelJS, loaded on demand from the app itself)
   ====================================================================== */
let excelLoading = null;
const vendorUrl = path => path + '?v=' + String(VENDOR[path] || '').slice(7, 19).replace(/[^A-Za-z0-9]/g, '');
const loaded = {};
function loadScript(path) {
  return loaded[path] || (loaded[path] = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = vendorUrl(path); s.integrity = VENDOR[path];
    s.onload = res; s.onerror = () => { delete loaded[path]; rej(new Error('تعذّر تحميل ملف من ملفات التطبيق: ' + path)); };
    document.head.appendChild(s);
  }));
}
async function loadBytes(path) {
  const r = await fetch(vendorUrl(path), { integrity: VENDOR[path] });
  if (!r.ok) throw new Error('تعذّر تحميل ملف من ملفات التطبيق: ' + path);
  return r.arrayBuffer();
}
function loadExcel() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  return excelLoading || (excelLoading = loadScript('vendor/exceljs.min.js').then(() => window.ExcelJS).finally(() => { excelLoading = null; }));
}
const XL = { ink: 'FF12262E', amber: 'FFE9A23B', amberSoft: 'FFFBEBD0', zebra: 'FFF2F5F4', line: 'FFCAD5D2', white: 'FFFFFFFF', muted: 'FF62747A' };
const thin = { style: 'thin', color: { argb: XL.line } };
const border = { top: thin, left: thin, bottom: thin, right: thin };
function sheetName(base, used) {
  let n = base.replace(/[\[\]:*?\/\\']/g, ' ').trim().slice(0, 31) || 'ورقة';
  if (/^history$/i.test(n)) n += ' 1';
  let k = 2, out = n; while (used.has(out.toLowerCase())) { const suf = ` (${k++})`; out = n.slice(0, 31 - suf.length) + suf; }
  used.add(out.toLowerCase()); return out;
}
const isoToDate = iso => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const colLetter = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const moneyFmt = () => S.settings.decimals ? '#,##0.' + '0'.repeat(S.settings.decimals) : '#,##0';
const argb = hex => 'FF' + (COLOR_RE.test(hex) ? hex.slice(1).toUpperCase() : 'E9A23B');

async function buildWorkbook(cars, withPhotos) {
  const ExcelJS = await loadExcel();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'سيّارتي'; wb.created = new Date();
  const used = new Set(['_sayyarati', '_sayyarati_photos']);
  const multi = cars.length > 1;
  const font = (o = {}) => ({ name: 'Arial', size: 11, ...o });

  for (const c of cars) {
    /* -- card sheet -- */
    const ws = wb.addWorksheet(sheetName(multi ? `بطاقة ${c.name}` : 'بطاقة السيارة', used), { views: [{ rightToLeft: true, showGridLines: false }], properties: { tabColor: { argb: argb(c.color) } } });
    ws.columns = [{ width: 3 }, { width: 26 }, { width: 24 }];
    ws.mergeCells('B2:C2');
    const t = ws.getCell('B2'); t.value = c.name; t.font = font({ size: 20, bold: true, color: { argb: XL.white } }); t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.ink } }; t.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(2).height = 40;
    let r = 3;
    const addKV = (label, val, type, bold) => {
      const a = ws.getCell(r, 2), b = ws.getCell(r, 3);
      a.value = String(label); a.font = font({ bold: true, color: { argb: XL.ink } }); a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.zebra } };
      setCell(b, type, val); b.font = font({ bold: !!bold }); a.border = border; b.border = border; b.alignment = { horizontal: 'center' };
      ws.getRow(r).height = 22; r++;
    };
    for (const f of c.cardFields) addKV(f.label, f.value, f.type);
    r++;
    const sumH = ws.getCell(r, 2); ws.mergeCells(r, 2, r, 3); sumH.value = 'ملخص'; sumH.font = font({ bold: true, color: { argb: XL.white } }); sumH.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.ink } }; sumH.alignment = { horizontal: 'center' }; r++;
    addKV('العداد الحالي', currentOdo(c), 'odometer', true);
    const total = c.sections.reduce((a, s) => a + sectionTotal(s), 0);
    for (const s of c.sections) if (s.columns.some(k => k.type === 'money')) addKV(`مجموع ${s.name}`, sectionTotal(s), 'money');
    addKV('إجمالي المصاريف', total, 'money', true);
    ws.getCell(r - 1, 2).fill = ws.getCell(r - 1, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.amberSoft } };

    /* -- section sheets -- */
    for (const s of c.sections) {
      const sh = wb.addWorksheet(sheetName(multi ? `${s.name} - ${c.name}` : s.name, used), { views: [{ rightToLeft: true, state: 'frozen', ySplit: 2, showGridLines: false }], properties: { tabColor: { argb: argb(c.color) } } });
      const cols = s.columns; const N = cols.length + 1;
      sh.columns = [{ width: 6 }, ...cols.map(k => ({ width: k.type === 'text' ? (/بيان/.test(k.label) ? 40 : 22) : 14 }))];
      sh.mergeCells(1, 1, 1, N);
      const tt = sh.getCell(1, 1); tt.value = `${s.name} — ${c.name}${s.note ? '   (' + s.note + ')' : ''}`;
      tt.font = font({ size: 15, bold: true, color: { argb: XL.white } }); tt.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.ink } }; tt.alignment = { vertical: 'middle', horizontal: 'center' };
      sh.getRow(1).height = 32;
      const hr = sh.getRow(2); hr.height = 24;
      ['م', ...cols.map(k => k.label)].forEach((h, i) => { const cell = hr.getCell(i + 1); cell.value = String(h); cell.font = font({ bold: true, color: { argb: 'FF1B1306' } }); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.amber } }; cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = border; });
      const rows = chrono(s);
      const oIdx = cols.findIndex(k => k.type === 'odometer');
      const dist = distances(s);
      rows.forEach((row, i) => {
        const R = sh.getRow(3 + i); R.height = 20;
        R.getCell(1).value = i + 1;
        cols.forEach((k, j) => {
          const cell = R.getCell(j + 2);
          if (k.type === 'distance') {
            if (oIdx >= 0 && i > 0 && toNum(row.v[cols[oIdx].id]) != null && toNum(rows[i - 1].v[cols[oIdx].id]) != null) {
              const L = colLetter(oIdx + 2); cell.value = { formula: `${L}${3 + i}-${L}${2 + i}`, result: dist.get(row.id) ?? 0 };
            }
            cell.numFmt = '#,##0';
          } else setCell(cell, k.type, row.v[k.id]);
        });
        for (let j = 1; j <= N; j++) { const cell = R.getCell(j); cell.border = border; cell.font = font(); cell.alignment = { horizontal: cols[j - 2] && cols[j - 2].type === 'text' ? 'right' : 'center', vertical: 'middle', wrapText: true };
          if (i % 2) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.zebra } }; }
        R.getCell(1).font = font({ color: { argb: XL.muted } });
      });
      const totR = sh.getRow(3 + rows.length); totR.height = 24;
      let anyMoney = false;
      cols.forEach((k, j) => {
        if (k.type !== 'money') return; anyMoney = true;
        const L = colLetter(j + 2); const cell = totR.getCell(j + 2);
        cell.value = rows.length ? { formula: `SUM(${L}3:${L}${2 + rows.length})`, result: sectionTotal(s) } : 0; cell.numFmt = moneyFmt();
      });
      if (anyMoney) {
        totR.getCell(1).value = 'المجموع';
        for (let j = 1; j <= N; j++) { const cell = totR.getCell(j); cell.font = font({ bold: true }); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.amberSoft } }; cell.border = border; cell.alignment = { horizontal: 'center' }; }
      }
      sh.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    }
  }
  /* -- invoice photos: a sheet to look at them, and a hidden one that ties each photo to its id (for restoring) -- */
  if (withPhotos) await addPhotoSheets(wb, cars, used, font);
  /* -- restore data (hidden sheet) -- */
  const hid = wb.addWorksheet('_sayyarati', { state: 'veryHidden' });
  const payload = JSON.stringify({ app: 'sayyarati', v: 1, exported: new Date().toISOString(), settings: S.settings, cars })
    .replace(/[\u007f-\uffff]/g, ch => '\\u' + ch.charCodeAt(0).toString(16).padStart(4, '0'));
  hid.getCell(1, 1).value = 'sayyarati-backup-v1';
  for (let i = 0, row = 2; i < payload.length; i += 30000, row++) hid.getCell(row, 1).value = payload.slice(i, i + 30000);
  return wb;
}
async function addPhotoSheets(wb, cars, used, font) {
  const items = photoItems(cars);
  if (!items.length) return 0;
  const gal = wb.addWorksheet(sheetName('صور الفواتير', used), { views: [{ rightToLeft: true, state: 'frozen', ySplit: 2, showGridLines: false }], properties: { tabColor: { argb: XL.amber } } });
  const map = wb.addWorksheet('_sayyarati_photos', { state: 'veryHidden' });
  map.getCell(1, 1).value = 'sayyarati-photos-v1';
  gal.columns = [{ width: 5 }, { width: 16 }, { width: 16 }, { width: 30 }, { width: 12 }, { width: 11 }, { width: 31 }];
  gal.mergeCells(1, 1, 1, 7);
  const t = gal.getCell(1, 1); t.value = 'صور الفواتير';
  t.font = font({ size: 15, bold: true, color: { argb: XL.white } }); t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.ink } }; t.alignment = { vertical: 'middle', horizontal: 'center' };
  gal.getRow(1).height = 32;
  const hr = gal.getRow(2); hr.height = 24;
  ['م', 'السيارة', 'القسم', 'البيان', 'التاريخ', 'العداد', 'الصورة'].forEach((h, i) => { const cell = hr.getCell(i + 1); cell.value = h; cell.font = font({ bold: true, color: { argb: 'FF1B1306' } }); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL.amber } }; cell.alignment = { horizontal: 'center', vertical: 'middle' }; cell.border = border; });
  let n = 0;
  for (const it of items) {
    const b = await Photos.blob(it.id).catch(() => null); if (!b) continue;
    const u8 = new Uint8Array(await b.arrayBuffer());
    const sz = jpegSize(u8); if (!sz) continue;
    const imgId = wb.addImage({ buffer: u8.buffer, extension: 'jpeg' });
    const k = Math.min(200 / sz.w, 280 / sz.h), w = Math.round(sz.w * k), h = Math.round(sz.h * k);
    const row = n + 3, R = gal.getRow(row);
    R.height = Math.round(h * 0.75) + 10;
    const tc = textCol(it.s), dc = colOf(it.s, 'date'), oc = colOf(it.s, 'odometer');
    R.getCell(1).value = n + 1;
    R.getCell(2).value = it.c.name;
    R.getCell(3).value = it.s.name;
    setCell(R.getCell(4), 'text', tc ? it.r.v[tc.id] : '');
    setCell(R.getCell(5), 'date', dc ? it.r.v[dc.id] : '');
    setCell(R.getCell(6), 'odometer', oc ? it.r.v[oc.id] : '');
    for (let j = 1; j <= 7; j++) { const cell = R.getCell(j); cell.border = border; cell.font = font(); cell.alignment = { horizontal: j === 4 ? 'right' : 'center', vertical: 'middle', wrapText: true }; }
    gal.addImage(imgId, { tl: { col: 6.04, row: row - 1 + 0.06 }, ext: { width: w, height: h }, editAs: 'oneCell' });
    map.getCell(n + 2, 1).value = it.id;
    map.addImage(imgId, { tl: { col: 1, row: n + 1 }, ext: { width: 4, height: 4 }, editAs: 'oneCell' });
    n++;
  }
  gal.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  return n;
}
/* the photos stored in a backup file: id → JPEG bytes (each checked; anything else is skipped) */
function readBackupPhotos(wb) {
  const out = new Map();
  const ws = wb.getWorksheet('_sayyarati_photos');
  if (!ws || !String(cellVal(ws.getCell(1, 1))).startsWith('sayyarati-photos')) return out;
  for (const im of ws.getImages()) {
    const row = Math.round(Number(im && im.range && im.range.tl && im.range.tl.nativeRow)) + 1;
    if (!(row >= 2)) continue;
    const id = String(cellVal(ws.getCell(row, 1)) || '');
    if (!ID_RE.test(id) || id in Object.prototype || out.has(id)) continue;
    const media = wb.getImage(im.imageId);
    if (!media || !media.buffer || !(media.buffer.length <= LIMITS.photoBytes)) continue;
    const u8 = media.buffer instanceof Uint8Array ? media.buffer : new Uint8Array(media.buffer);
    if (!jpegSize(u8)) continue;
    out.set(id, u8);
    if (out.size > LIMITS.photos) throw new ImportError('too-many');
  }
  return out;
}
/* values are written as typed cells; text always stays text (never a formula) */
function setCell(cell, type, v) {
  if (v === '' || v == null) { cell.value = null; return; }
  if (type === 'money') { const n = toNum(v); cell.value = n == null ? String(v) : n; cell.numFmt = moneyFmt(); }
  else if (type === 'odometer') { const n = toNum(v); cell.value = n == null ? String(v) : n; cell.numFmt = '#,##0'; }
  else if (type === 'number') { const n = toNum(v); cell.value = n == null ? String(v) : n; }
  else if (type === 'date') { const iso = parseDateAny(v); if (iso) { cell.value = isoToDate(iso); cell.numFmt = 'dd.mm.yyyy'; } else cell.value = String(v); }
  else cell.value = String(v);
}

/* ---------- export sheet: PDF report of one car + Excel backup + restore ---------- */
function openBackup(carId) {
  if (!S.cars.length) return;
  let pdfCar = S.cars.find(x => x.id === carId) || car();
  const c = car();
  const mb = n => (n >= 1024 * 1024 ? `${nf(n / 1024 / 1024, n < 10 * 1024 * 1024 ? 1 : 0)} ميجابايت` : `${nf(Math.max(1, n / 1024))} كيلوبايت`);
  openSheet('تصدير ومشاركة', `
    <div class="sub-h" style="margin-top:4px">تقرير PDF لسيارة</div>
    <p style="margin:0 0 12px;color:var(--muted);font-size:14px">تقرير منسّق لسيارة واحدة: بطاقتها وملخص مصاريفها وكل سجلاتها، للطباعة أو لإرساله إلى مشترٍ أو ورشة.</p>
    ${S.cars.length > 1 ? `<div class="pick" aria-label="السيارة">${S.cars.map((x, i) => `<button class="chip" data-pc="${i}" aria-pressed="${x.id === pdfCar.id}"><span class="dot" style="background:${esc(x.color)}"></span>${esc(x.name)}</button>`).join('')}</div>` : ''}
    <label class="check"><input type="checkbox" id="pHide"> إخفاء المبالغ والأسعار (عند مشاركة التقرير مع مشترٍ مثلاً)</label>
    <label class="check" id="pPhWrap"><input type="checkbox" id="pPhotos" checked> إرفاق صور الفواتير (<span id="pPhN"></span>)</label>
    <p class="hint-p hidden" id="pPhHint">صور الفواتير قد تُظهر المبالغ، فلن تُرفق إلا إن اخترتها.</p>
    <button class="btn amber wide" id="pShare">${isNative ? 'إنشاء التقرير ومشاركته' : 'إنشاء التقرير ومشاركته / حفظه'}</button>
    ${!isNative ? `<button class="btn ghost wide" id="pDl" style="margin-top:10px">تنزيل التقرير فقط</button>` : ''}
    <div class="sub-h">نسخة احتياطية Excel</div>
    <p style="margin:0 0 14px;color:var(--muted);font-size:14px">ستظهر قائمة المشاركة: اختر «Drive» للحفظ في جوجل درايف، أو «الملفات»، أو أرسله بالواتساب أو البريد.</p>
    ${S.cars.length > 1 ? `<div class="pick"><button class="chip" data-w="all" aria-pressed="true">كل السيارات (${S.cars.length})</button><button class="chip" data-w="one" aria-pressed="false">${esc(c.name)} فقط</button></div>` : ''}
    <label class="check" id="bPhWrap"><input type="checkbox" id="bPhotos" checked> مع صور الفواتير (<span id="bPhN"></span>)</label>
    <button class="btn wide" id="bShare">${isNative ? 'تصدير ومشاركة' : 'تصدير ومشاركة / حفظ'}</button>
    ${!isNative ? `<button class="btn ghost wide" id="bDl" style="margin-top:10px">تنزيل الملف فقط</button>` : ''}
    <div class="sub-h">الاسترجاع</div>
    <p style="margin:0 0 12px;color:var(--muted);font-size:14px">اختر ملف إكسل صدّرته من التطبيق لاسترجاع كل شيء، أو أي ملف إكسل بنفس ترتيب ملفك القديم لاستيراده كسيارة.${isNative ? ' ويمكنك أيضاً فتح الملف من تطبيق Drive أو الملفات واختيار «فتح باستخدام ← سيّارتي».' : ''}</p>
    <button class="btn ghost wide" id="bImp">استيراد / استرجاع من ملف</button>`, b => {
    let which = 'all', sizes = null;
    const xlCars = () => (which === 'one' ? [c] : S.cars);
    /* how many photos each choice carries (and roughly how big the backup gets) */
    const counts = () => {
      const pn = photoItems([pdfCar]).length;
      b.querySelector('#pPhWrap').classList.toggle('hidden', !pn);
      b.querySelector('#pPhN').textContent = nPhotos(pn);
      const ids = [...usedPhotoIds(xlCars())];
      b.querySelector('#bPhWrap').classList.toggle('hidden', !ids.length);
      const bytes = sizes ? ids.reduce((a, id) => a + (sizes.get(id) || 0), 0) : 0;
      b.querySelector('#bPhN').textContent = `${nPhotos(ids.length)}${bytes ? '، نحو ' + mb(bytes) : ''}`;
    };
    counts();
    Photos.list().then(l => { sizes = new Map(l.map(x => [x.id, x.size])); if (b.isConnected) counts(); }).catch(() => {});
    b.querySelectorAll('[data-pc]').forEach(x => x.onclick = () => { const t = S.cars[+x.dataset.pc]; if (!t) return; pdfCar = t; b.querySelectorAll('[data-pc]').forEach(y => y.setAttribute('aria-pressed', y === x)); counts(); });
    b.querySelectorAll('[data-w]').forEach(x => x.onclick = () => { which = x.dataset.w === 'one' ? 'one' : 'all'; b.querySelectorAll('[data-w]').forEach(y => y.setAttribute('aria-pressed', y === x)); counts(); });
    /* invoices show prices: hiding the amounts also leaves the photos out, unless chosen again */
    b.querySelector('#pHide').onchange = e => { b.querySelector('#pPhotos').checked = !e.target.checked; b.querySelector('#pPhHint').classList.toggle('hidden', !e.target.checked); };
    const busy = async (btn, work) => {
      const old = btn.textContent; btn.textContent = 'جارٍ تجهيز الملف…'; btn.disabled = true;
      try { await work(); } catch (e) { console.error(e); toast('تعذّر إنشاء الملف: ' + String((e && e.message) || e).slice(0, 120)); }
      btn.textContent = old; btn.disabled = false;
    };
    const runPdf = mode => busy(b.querySelector(mode === 'dl' ? '#pDl' : '#pShare'), async () => {
      const withPhotos = !b.querySelector('#pPhWrap').classList.contains('hidden') && b.querySelector('#pPhotos').checked;
      const blob = await makePdf(pdfCar, b.querySelector('#pHide').checked, withPhotos);
      const name = `تقرير-${safeName(pdfCar.name)}-${todayISO()}.pdf`;
      if (await deliver(blob, name, mode, 'مشاركة التقرير')) { closeSheet(); toast('جُهّز تقرير PDF'); }
    });
    const runXlsx = mode => busy(b.querySelector(mode === 'dl' ? '#bDl' : '#bShare'), async () => {
      const cars = xlCars();
      const withPhotos = !b.querySelector('#bPhWrap').classList.contains('hidden') && b.querySelector('#bPhotos').checked;
      const wb = await buildWorkbook(cars, withPhotos);
      /* photos are already compressed: zip them lightly, it is much faster on a phone */
      const buf = await wb.xlsx.writeBuffer({ zip: { compressionOptions: { level: withPhotos ? 1 : 6 } } });
      const name = `سيارتي-${which === 'one' || cars.length === 1 ? safeName(cars[0].name) + '-' : ''}${todayISO()}.xlsx`;
      if (await deliver(new Blob([buf], { type: XLSX_MIME }), name, mode, 'حفظ النسخة الاحتياطية')) { S.lastBackup = Date.now(); save(); closeSheet(); render(); toast('جُهّز ملف الإكسل'); }
    });
    b.querySelector('#pShare').onclick = () => runPdf('share');
    if (b.querySelector('#pDl')) b.querySelector('#pDl').onclick = () => runPdf('dl');
    b.querySelector('#bShare').onclick = () => runXlsx('share');
    if (b.querySelector('#bDl')) b.querySelector('#bDl').onclick = () => runXlsx('dl');
    b.querySelector('#bImp').onclick = pickImport;
  });
}

/* ---------- PDF report ---------- */
const PDF_PHOTOS = 300;
async function makePdf(c, hideMoney, withPhotos) {
  await Promise.all(['vendor/pdf-lib.min.js', 'vendor/fontkit.umd.min.js', 'vendor/bidi.min.js', 'report.js'].map(loadScript));
  const [regular, bold] = await Promise.all([loadBytes('fonts/readex-report-400.ttf'), loadBytes('fonts/readex-report-700.ttf')]);
  const model = reportModel(c, hideMoney);
  if (withPhotos) {
    const all = photoItems([c]), take = all.slice(-PDF_PHOTOS), items = [];
    for (const it of take) {
      const b = await Photos.blob(it.id).catch(() => null); if (!b) continue;
      const jpg = new Uint8Array(await b.arrayBuffer());
      if (jpegSize(jpg)) items.push({ caption: photoCaption(it.s, it.r), jpg });
    }
    if (items.length) model.blocks.push({ type: 'photos', title: 'صور الفواتير', note: take.length < all.length ? `آخر ${nf(take.length)} من ${nf(all.length)}` : '', items });
  }
  const bytes = await window.SayyaratiReport.build(model, {
    PDFLib: window.PDFLib, fontkit: window.fontkit, bidiFactory: window.bidi_js,
    fonts: { regular: new Uint8Array(regular), bold: new Uint8Array(bold) }
  });
  return new Blob([bytes], { type: 'application/pdf' });
}
const ltrIso = t => '⁦' + t + '⁩';
const hexRgb = h => (COLOR_RE.test(h) ? [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16)) : [0xe9, 0xa2, 0x3b]);
/* everything the report shows, as ready-made text (report.js only lays it out) */
function reportModel(c, hide) {
  const moneyT = v => { const n = toNum(v); return n == null ? String(v) : `${nf(n, S.settings.decimals)} ${S.settings.currency}`; };
  const cellText = (type, v) => {
    if (v === '' || v == null) return '';
    if (type === 'money') return moneyT(v);
    if (type === 'odometer') { const n = toNum(v); return n == null ? String(v) : nf(n); }
    if (type === 'date') return fmtDate(v);
    return String(v).slice(0, 400);
  };
  const cur = currentOdo(c), rate = kmPerDay(c), driven = cur - firstOdo(c);
  const total = c.sections.reduce((a, s) => a + sectionTotal(s), 0);
  const sub = [];
  const model = c.cardFields.find(f => f.label.includes('موديل') && f.value !== '');
  if (model) sub.push('موديل ' + model.value);
  const plate = c.cardFields.find(f => f.label.includes('لوحة') && f.value !== '');
  if (plate) sub.push('لوحة ' + plate.value);

  const summary = [{ label: 'العداد الحالي', value: `${nf(cur)} كم` }, { label: 'عدد التسجيلات', value: nf(countRows([c])) }];
  if (!hide) {
    summary.push({ label: 'إجمالي المصاريف', value: moneyT(total) });
    if (driven > 0) summary.push({ label: 'تكلفة الكيلو', value: `${nf(total / driven * 1000, 1)} فلس` });
  }
  if (rate) summary.push({ label: 'معدل القيادة', value: `${nf(rate * 30)} كم/شهر` });
  for (const x of reminders(c)) {
    if (x.none) continue;
    const k = x.km, t = x.time;
    const value = x.by === 'km'
      ? (k.left > 0 ? `عند ${nf(k.next)} كم، باقي ${nf(k.left)} كم` : `متأخر ${nf(-k.left)} كم`) + (t ? `، أو في ${fmtDate(t.due)}` : '')
      : `في ${fmtDate(t.due)}، ${t.days > 0 ? 'باقي ' + nDays(t.days) : t.days === 0 ? 'اليوم' : 'متأخر ' + nDays(-t.days)}` + (k ? `، أو عند ${nf(k.next)} كم` : '');
    summary.push({ label: `${x.sec.name} القادم`, value, accent: x.late });
  }
  for (const e of expiries(c)) summary.push({ label: e.title, value: `${e.days >= 0 ? 'ينتهي' : 'انتهى'} ${fmtDate(e.date)}`, accent: e.days < 30 });

  const blocks = [];
  const card = c.cardFields.filter(f => !(hide && f.type === 'money'));
  if (card.length) blocks.push({ type: 'kv', title: 'بطاقة السيارة', rows: card.map(f => [f.label, (f.type === 'odometer' && toNum(f.value) != null ? `${nf(f.value)} كم` : cellText(f.type, f.value)) || '—']) });
  for (const s of c.sections) {
    const cols = s.columns.filter(k => !(hide && k.type === 'money'));
    const all = chrono(s), dist = distances(s);
    const rows = all.slice(-2000);
    const first = all.length - rows.length;
    const t = {
      type: 'table', title: s.name, note: [s.note, rows.length < all.length ? `آخر ${nf(rows.length)} من ${nf(all.length)}` : ''].filter(Boolean).join(' · '),
      index: true, empty: 'لا توجد تسجيلات',
      columns: [{ label: 'م', align: 'center' }, ...cols.map(k => ({ label: k.label, align: k.type === 'text' ? 'right' : 'center' }))],
      rows: rows.map((r, i) => [nf(first + i + 1), ...cols.map(k => (k.type === 'distance' ? (dist.has(r.id) ? `${ltrIso('+' + nf(dist.get(r.id)))} كم` : '') : cellText(k.type, r.v[k.id])))])
    };
    if (!hide && rows.length && cols.some(k => k.type === 'money')) {
      const label = Math.max(0, cols.findIndex(k => k.type === 'text'));
      t.totals = ['', ...cols.map((k, j) => (k.type === 'money' ? moneyT(rows.reduce((a, r) => a + (toNum(r.v[k.id]) || 0), 0)) : j === label ? 'المجموع' : ''))];
      if (cols[label] && cols[label].type === 'money') t.totals[0] = 'المجموع';
    }
    blocks.push(t);
  }
  if (!hide) {
    const years = {};
    for (const s of c.sections) { const d = colOf(s, 'date'); for (const k of s.columns) if (k.type === 'money') for (const r of s.rows) { const dv = d ? String(r.v[d.id] || '') : ''; const y = ISO_RE.test(dv) ? dv.slice(0, 4) : 'بدون تاريخ'; years[y] = (years[y] || 0) + (toNum(r.v[k.id]) || 0); } }
    const ys = Object.entries(years).sort((a, b) => (/^\d/.test(b[0]) ? +b[0] : -1) - (/^\d/.test(a[0]) ? +a[0] : -1));
    if (ys.length) blocks.push({ type: 'table', title: 'المصاريف حسب السنة', columns: [{ label: 'السنة', align: 'center' }, { label: 'المبلغ', align: 'center' }], rows: ys.map(([y, v]) => [y, moneyT(v)]), totals: ['المجموع', moneyT(total)] });
  }
  return {
    fileTitle: `تقرير ${c.name}`, kicker: 'سيّارتي · تقرير السيارة', title: c.name, subtitle: sub.join('، '),
    dateLine: `تاريخ التقرير ${fmtDate(todayISO())}`, color: hexRgb(c.color), footer: `سيّارتي — ${c.name}`,
    notice: hide ? 'هذا التقرير لا يتضمن المبالغ والأسعار.' : '', summary, blocks
  };
}
const safeName = s => String(s).replace(/[\\/:*?"<>|\u0000-\u001F\s]+/g, '-').replace(/^\.+/, '').slice(0, 40) || 'car';
function bufToB64(buf) { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
function b64ToBuf(b64) { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u.buffer; }
async function deliver(blob, name, mode, dialogTitle = 'مشاركة الملف') {
  if (isNative) {
    const fs = plug('Filesystem'), sh = plug('Share');
    const w = await writeChunked(fs, 'exports/' + name, blob, 'CACHE');
    try { await sh.share({ title: name, dialogTitle, files: [w.uri] }); }
    catch (e) { if (/cancel/i.test(String(e && e.message))) return false; throw e; }
    return true;
  }
  const file = new File([blob], name, { type: blob.type });
  if (mode !== 'dl' && navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return true; }
    catch (e) { if (e && e.name === 'AbortError') return false; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  return true;
}
/* exported files stay in the app's cache only as long as the share needs them */
async function cleanOldExports() {
  const fs = plug('Filesystem'); if (!fs) return;
  try {
    const r = await fs.readdir({ path: 'exports', directory: 'CACHE' });
    for (const f of r.files || []) {
      const name = typeof f === 'string' ? f : f.name; const mtime = typeof f === 'string' ? 0 : Number(f.mtime) || 0;
      if (name && Date.now() - mtime > 864e5) await fs.deleteFile({ path: 'exports/' + name, directory: 'CACHE' }).catch(() => {});
    }
  } catch (e) { /* no exports yet */ }
}

/* ---------- import ---------- */
class ImportError extends Error { constructor(kind, detail) { super(kind); this.kind = kind; this.detail = detail || ''; } }
let importBusy = false;
async function pickImport() {
  const p = xlsxPlugin();
  if (p) {
    if (importBusy) return;
    importBusy = true;
    try {
      const r = await p.pick();
      if (r && !r.cancelled) await importNative(r);
    } catch (e) { showImportError(nativeError(e)); }
    finally { importBusy = false; }
    return;
  }
  const f = $('#fileIn'); f.value = ''; f.click();
}
/* the native picker copies the chosen file into the app's cache; the page reads it from there and deletes it */
async function importNative(r) {
  let buf = null;
  if (r && typeof r.uri === 'string' && /^file:\/\//.test(r.uri)) {
    try {
      const res = await fetch(window.Capacitor.convertFileSrc(r.uri));
      if (!res.ok) throw new Error('HTTP ' + res.status);
      buf = await res.arrayBuffer();
    } catch (e) { throw new ImportError('unreadable', (e && e.message) || ''); }
    finally { if (typeof r.path === 'string' && /^imports\/[\w.-]+$/.test(r.path)) plug('Filesystem').deleteFile({ path: r.path, directory: 'CACHE' }).catch(() => {}); }
  } else if (r && typeof r.data === 'string') buf = b64ToBuf(r.data);
  if (!buf) return;
  if (!buf.byteLength) throw new ImportError('unreadable', 'empty');
  await importBuffer(buf, r.name);
}
function nativeError(e) {
  if (e instanceof ImportError) return e;
  const code = e && e.code, msg = String((e && e.message) || e || '');
  if (code === 'TOO_LARGE') return new ImportError('too-large');
  if (code === 'NO_PERMISSION' || code === 'READ_FAILED' || code === 'TIMEOUT') return new ImportError('unreadable', msg);
  if (code === 'NO_XLSX') return new ImportError('sheets-export', msg);
  return new ImportError('other', msg);
}
$('#fileIn').onchange = async e => {
  const file = e.target.files && e.target.files[0]; if (!file) return;
  try {
    if (file.size > MAX_IMPORT_BYTES) throw new ImportError('too-large');
    let buf;
    try { buf = await file.arrayBuffer(); } catch (err) { throw new ImportError('unreadable', (err && err.name) || ''); }
    if (!buf.byteLength) throw new ImportError('unreadable', 'empty');
    await importBuffer(buf, file.name);
  } catch (err) { showImportError(err); }
};
function showImportError(e) {
  console.warn('import failed', e);
  const kind = e instanceof ImportError ? e.kind : 'other';
  const detail = e instanceof ImportError ? e.detail : String((e && e.message) || e || '');
  const msgs = {
    'too-large': 'الملف أكبر من 150 ميجابايت، وهذا أكبر من أي نسخة احتياطية للتطبيق. تأكد أنك اخترت الملف الصحيح.',
    'unreadable': 'لم يتمكن الهاتف من قراءة الملف من مكانه. يحدث هذا أحياناً مع ملفات الدرايف التي لم تُنزَّل بعد. جرّب أحد الحلين: افتح تطبيق Drive، واضغط ⋮ بجانب الملف ثم «فتح باستخدام» ← «سيّارتي»، أو نزّل الملف إلى الهاتف ثم اختره من «التنزيلات».',
    'not-xlsx': 'هذا الملف ليس بصيغة Excel الحديثة (‎.xlsx‎). إن كان بصيغة ‎.xls‎ القديمة أو جدول Google Sheets، افتحه واحفظه أو نزّله بصيغة ‎.xlsx‎ ثم استورده.',
    'sheets-export': 'هذا جدول Google Sheets ولم يسمح الدرايف بتحويله تلقائياً. افتحه في Google Sheets واختر «تنزيل ← Microsoft Excel (‎.xlsx‎)» ثم استورد الملف الناتج.',
    'bad-backup': 'الملف فيه نسخة احتياطية تالفة أو ناقصة، فلم يُستورد شيء.',
    'too-many': 'الملف فيه تسجيلات أكثر مما يتسع له التطبيق (15,000 تسجيل لكل السيارات معاً)، فلم يُستورد شيء.',
    'empty': 'لم أجد في الملف بيانات يمكن استيرادها.',
    'other': 'تعذّر استيراد الملف.'
  };
  openSheet('تعذّر الاستيراد', `<p style="margin:0 0 14px">${esc(msgs[kind] || msgs.other)}</p>
    ${detail ? `<p style="margin:0 0 16px;color:var(--muted);font-size:12.5px;direction:ltr;text-align:left">${esc(String(detail).slice(0, 200))}</p>` : ''}
    <div class="btn-row"><button class="btn ghost" id="ieClose">إغلاق</button><button class="btn amber" id="ieRetry">اختيار ملف آخر</button></div>`, b => {
    b.querySelector('#ieClose').onclick = closeSheet;
    b.querySelector('#ieRetry').onclick = () => { closeSheet(); pickImport(); };
  });
}
function cellVal(cell) {
  let v = cell.value;
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    if ('result' in v) v = v.result;
    else if (Array.isArray(v.richText)) v = v.richText.map(t => String(t.text ?? '')).join('');
    else if ('text' in v) v = v.text;
    else v = null;
  }
  if (v && typeof v === 'object' && !(v instanceof Date)) v = null;
  return v == null ? '' : v;
}
async function importBuffer(buf, fname) {
  if (buf.byteLength > MAX_IMPORT_BYTES) throw new ImportError('too-large');
  const head = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
  if (!(head[0] === 0x50 && head[1] === 0x4B && head[2] === 0x03 && head[3] === 0x04)) throw new ImportError('not-xlsx');
  const ExcelJS = await loadExcel();
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(buf); } catch (e) { throw new ImportError('not-xlsx', (e && e.message) || ''); }
  const hid = wb.getWorksheet('_sayyarati');
  if (hid && String(cellVal(hid.getCell(1, 1))).startsWith('sayyarati-backup')) {
    let json = '';
    for (let r = 2; r <= Math.min(hid.rowCount, 2000); r++) {
      json += String(cellVal(hid.getCell(r, 1)) || '');
      if (json.length > LIMITS.backupChars) throw new ImportError('too-many');
    }
    let data;
    try { data = JSON.parse(json); } catch (e) { throw new ImportError('bad-backup'); }
    if (!isObj(data) || !Array.isArray(data.cars)) throw new ImportError('bad-backup');
    return askRestore(data, readBackupPhotos(wb));
  }
  const c = sanitizeCar(genericImport(wb, fname), new Set(S.cars.map(x => x.id)));
  if (!c || (!c.sections.some(s => s.rows.length) && !c.cardFields.some(f => f.value !== ''))) throw new ImportError('empty');
  const n = c.sections.reduce((a, s) => a + s.rows.length, 0);
  if (n + countRows(S.cars) > LIMITS.totalRows) throw new ImportError('too-many');
  openSheet('استيراد سيارة', `<p style="margin:0 0 6px">وجدت في الملف:</p>
    <dl class="kv" style="margin-bottom:18px"><dt>السيارة</dt><dd>${esc(c.name)}</dd>${c.sections.map(s => `<dt>${esc(s.name)}</dt><dd>${s.rows.length} تسجيل</dd>`).join('')}<dt>بنود البطاقة</dt><dd>${c.cardFields.length}</dd></dl>
    <label class="field"><span>اسم السيارة</span><input id="giName" maxlength="${LIMITS.label}" value="${esc(c.name)}"></label>
    <button class="btn amber wide" id="giOk">إضافة السيارة (${n} تسجيل)</button>`, b => {
    b.querySelector('#giOk').onclick = () => {
      if (S.cars.length >= LIMITS.cars) { toast('وصلت إلى الحد الأقصى لعدد السيارات'); return; }
      if (n + countRows(S.cars) > LIMITS.totalRows) { showImportError(new ImportError('too-many')); return; }
      c.name = str(b.querySelector('#giName').value, LIMITS.label).trim() || c.name;
      S.cars.push(c); S.activeCar = c.id; save(); closeSheet(); rolledFor = null; go('home'); toast(`استُوردت ${c.name}`);
    };
  });
}
function askRestore(raw, photos = new Map()) {
  const cars = sanitizeCars(raw.cars);
  const settings = sanitizeSettings(raw.settings);
  const n = countRows(cars);
  const when = parseDateAny(String(raw.exported || '').slice(0, 10));
  if (!cars.length) throw new ImportError('bad-backup');
  if (n > LIMITS.totalRows) throw new ImportError('too-many');
  const inFile = [...usedPhotoIds(cars)].filter(id => photos.has(id)).length;
  openSheet('استرجاع نسخة احتياطية', `<p style="margin:0 0 16px">${when ? `النسخة بتاريخ ${dateHTML(when)} ` : 'النسخة '}وفيها ${cars.length} ${cars.length === 1 ? 'سيارة' : 'سيارات'} و${n} تسجيل${inFile ? `، وصور الفواتير: ${nf(inFile)}` : ''}.</p>
    <button class="btn amber wide" id="rMerge">دمج مع بياناتي الحالية</button>
    <p class="hint" style="font-size:13px;color:var(--muted);margin:6px 2px 14px">السيارة الموجودة في النسختين تُستبدل بنسخة الملف، وتبقى السيارات الأخرى.</p>
    <button class="btn danger wide" id="rReplace">استبدال كل بياناتي بالنسخة</button>`, b => {
    b.querySelector('#rMerge').onclick = () => {
      const merged = S.cars.slice();
      for (const c of cars) { const i = merged.findIndex(x => x.id === c.id); if (i >= 0) merged[i] = c; else merged.push(c); }
      if (countRows(merged) > LIMITS.totalRows || merged.length > LIMITS.cars) { showImportError(new ImportError('too-many')); return; }
      finishRestore(merged, settings, photos, cars);
    };
    b.querySelector('#rReplace').onclick = () => confirmBox('ستُحذف كل البيانات الحالية وتحل محلها النسخة. متابعة؟', 'استبدال', () => finishRestore(cars, settings, photos, cars));
  });
}
let restoring = false;
async function finishRestore(cars, settings, photos = new Map(), restored = cars) {
  if (restoring) return;
  restoring = true;
  openSheet('استرجاع نسخة احتياطية', `<p style="margin:0" id="rsProg" role="status">جارٍ الاسترجاع…</p>`);
  let failed = 0;
  try {
    /* the photos the restored entries use, then references to photos that exist nowhere are dropped */
    const need = usedPhotoIds(restored);
    const have = new Set((await Photos.list().catch(() => [])).map(p => p.id));
    const todo = [...photos.keys()].filter(id => need.has(id) && !have.has(id));
    for (let i = 0; i < todo.length; i++) {
      const prog = $('#rsProg'); if (prog) prog.textContent = `جارٍ حفظ صور الفواتير… ${nf(i + 1)} من ${nf(todo.length)}`;
      try { await Photos.put(new Blob([photos.get(todo[i])], { type: 'image/jpeg' }), todo[i]); have.add(todo[i]); } catch (e) { console.warn('photo restore', e); failed++; }
    }
    for (const c of restored) for (const s of c.sections) for (const r of s.rows) if (r.p) { r.p = r.p.filter(id => have.has(id)); if (!r.p.length) delete r.p; }
    S = sanitizeState({ ...S, cars, settings });
    save(); applyTheme(); closeSheet(); rolledFor = null; go('home');
    toast(failed ? `استُرجعت النسخة، وتعذّر حفظ ${nf(failed)} من الصور (مساحة التخزين؟)` : 'استُرجعت النسخة');
  } finally { restoring = false; }
}
/* reads a workbook laid out like the original F150 file: one key/value sheet + table sheets */
function guessType(label) {
  if (/مسافة|المسافه/.test(label)) return 'distance';
  if (/عداد|كيلو|km/i.test(label)) return 'odometer';
  if (/تاريخ|date/i.test(label)) return 'date';
  if (/مبلغ|سعر|تكلفة|قيمة|كلفة|price|cost|amount/i.test(label)) return 'money';
  if (/موديل|سنة|لتر/.test(label)) return 'number';
  return 'text';
}
function genericImport(wb, fname) {
  const c = { id: uid(), name: str(String(fname || '').replace(/\.xlsx?$/i, ''), LIMITS.label) || 'سيارة', color: CAR_COLORS[S.cars.length % CAR_COLORS.length], odoNow: null, created: Date.now(), cardFields: [], sections: [] };
  let gotName = false;
  wb.eachSheet(ws => {
    if (ws.state && ws.state !== 'visible') return;
    if (c.sections.length >= LIMITS.sections) return;
    const grid = [];
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      if (grid.length > LIMITS.rows) return;
      const cells = [];
      row.eachCell({ includeEmpty: false }, (cell, cn) => {
        if (cn > 200) return;
        if (cell.isMerged && cell.master && cell.master.address !== cell.address) return;
        const v = cellVal(cell); if (v !== '' && v != null) cells.push({ cn, v });
      });
      if (cells.length) grid.push({ rn, cells });
    });
    if (!grid.length) return;
    /* header row: first row with >= 3 text cells */
    const hdr = grid.find(r => r.cells.filter(x => typeof x.v === 'string').length >= 3);
    if (!hdr) {
      for (const r of grid) {
        if (r.cells.length === 1 && typeof r.cells[0].v === 'string' && !gotName) { c.name = str(r.cells[0].v, LIMITS.label).trim() || c.name; gotName = true; continue; }
        if (r.cells.length >= 2 && typeof r.cells[0].v === 'string' && c.cardFields.length < LIMITS.cardFields) {
          const label = str(r.cells[0].v, LIMITS.label).trim(); let type = guessType(label); if (type === 'distance') type = 'odometer';
          c.cardFields.push({ id: uid(), label, type, value: cleanValue(type, r.cells[1].v instanceof Date ? parseDateAny(r.cells[1].v) : r.cells[1].v) });
        }
      }
      return;
    }
    const tmpl = /زيت|زيوت/.test(ws.name) ? 'oil' : /صيان/.test(ws.name) ? 'maint' : null;
    const sec = { id: tmpl && !c.sections.some(s => s.id === tmpl) ? tmpl : uid(), name: str(ws.name, LIMITS.label).trim(), icon: tmpl === 'oil' ? 'oil' : tmpl === 'maint' ? 'wrench' : /وقود|بنزين/.test(ws.name) ? 'fuel' : /تأمين|ترخيص/.test(ws.name) ? 'shield' : 'doc', note: '', reminderKm: tmpl === 'oil' ? 7500 : 0, reminderMonths: tmpl === 'oil' ? 6 : 0, columns: [], rows: [] };
    const above = grid.filter(r => r.rn < hdr.rn).flatMap(r => r.cells).filter(x => typeof x.v === 'string');
    if (above.length) sec.note = str(above.map(x => x.v.trim()).join(' '), 200);
    const colMap = [];
    for (const h of hdr.cells) { if (typeof h.v !== 'string' || colMap.length >= LIMITS.columns) continue; const label = str(h.v, LIMITS.label).trim(); const col = { id: uid(), label, type: guessType(label) }; sec.columns.push(col); colMap.push({ cn: h.cn, col }); }
    if (!sec.columns.some(x => x.type === 'odometer')) sec.columns.forEach(x => { if (x.type === 'distance') x.type = 'number'; });
    const stdIds = { 'بيان': 'desc', 'العداد': 'odo', 'المسافة': 'dist', 'التاريخ': 'date', 'المبلغ': 'amount', 'الزيت': 'oil', 'ملاحظات': 'notes' };
    if (tmpl) sec.columns.forEach(x => { if (stdIds[x.label] && !sec.columns.some(y => y.id === stdIds[x.label])) x.id = stdIds[x.label]; });
    for (const r of grid) {
      if (r.rn <= hdr.rn || sec.rows.length >= LIMITS.rows) continue;
      const v = {}; let any = false;
      for (const { cn, col } of colMap) {
        if (col.type === 'distance') continue;
        const cell = r.cells.find(x => x.cn === cn); if (!cell) { v[col.id] = ''; continue; }
        let x = cell.v;
        if (col.type === 'text' && x instanceof Date) x = fmtDate(parseDateAny(x));
        x = cleanValue(col.type, x instanceof Date && col.type !== 'date' ? parseDateAny(x) : x);
        v[col.id] = x; if (x !== '') any = true;
      }
      /* skip the totals row: only money cells filled and nothing else */
      const nonMoney = sec.columns.filter(k => k.type !== 'money' && k.type !== 'distance').some(k => v[k.id] !== '' && v[k.id] != null);
      if (!any || !nonMoney) continue;
      sec.rows.push({ id: uid(), t: Date.now() + sec.rows.length, v });
    }
    c.sections.push(sec);
  });
  if (!c.cardFields.length) c.cardFields = defaultCardFields();
  if (!c.sections.length) c.sections = [SECTION_TEMPLATES.oil(), SECTION_TEMPLATES.maint()];
  return c;
}

/* ---------- files handed to the Android app ("Open with" / "Share") ---------- */
let takingOpened = false;
async function takeOpenedFile() {
  const p = xlsxPlugin(); if (!p || takingOpened) return;
  takingOpened = true;
  try {
    const r = await p.takeOpenedFile();
    if (r && !r.none) await importNative(r);
  } catch (e) { showImportError(nativeError(e)); }
  finally { takingOpened = false; }
}

/* ---------- native back button ---------- */
function initNative() {
  const app = plug('App');
  if (app) app.addListener('backButton', () => {
    if (viewerOpen()) { closeViewer(); return; }
    if (sheetOpen()) { closeSheet(); return; }
    if (view !== 'home') { go('home'); return; }
    app.minimizeApp ? app.minimizeApp() : app.exitApp();
  });
  const p = xlsxPlugin();
  if (p && p.addListener) p.addListener('fileOpened', () => takeOpenedFile());
  takeOpenedFile();
  cleanOldExports();
  initNotifications().catch(e => console.warn('notifications', e));
}

/* ---------- boot ---------- */
document.querySelectorAll('#nav button[data-v]').forEach(b => b.onclick = () => go(b.dataset.v));
$('#fab').onclick = () => { if (!S.cars.length) return; openEntry(); };
$('#carPill').onclick = () => openCarSwitcher();
$('#quickExport').onclick = () => openBackup();
$('#scrim').onclick = closeSheet;
initViewer();
$('#sheetClose').onclick = closeSheet;

(async function boot() {
  let raw = loadLocal();
  if (isNative) { const n = await loadNative(); if (n && (!raw || finite(n.rev) > finite(raw && raw.rev))) raw = n; }
  let loaded = true;
  try { S = sanitizeState(raw); } catch (e) { console.error('state could not be read', e); S = blankState(); loaded = false; }
  applyTheme();
  try { render(); } catch (e) { console.error(e); S = sanitizeState(S); render(); }
  if (isNative) initNative();
  /* photos no entry uses any more are removed once a day — never when the data could not be read */
  if (loaded && raw) setTimeout(gcPhotos, 4000);
  if (!isNative && 'serviceWorker' in navigator && window.isSecureContext) navigator.serviceWorker.register('sw.js').catch(() => {});
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
})();
