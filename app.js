'use strict';
/* ======================================================================
   سيّارتي — سجل السيارة
   البيانات تُحفظ على الجهاز فقط (localStorage، وفي تطبيق أندرويد أيضاً ملف داخل التطبيق).
   قواعد الأمان في هذا الملف (راجع SECURITY.md):
   - كل نص يُدرج في HTML يمر عبر esc()، وكل معرّف/لون/نوع يأتي من ملف يُتحقق منه في sanitize*().
   - لا eval ولا سكربتات خارجية؛ سياسة CSP في index.html تمنع أي سكربت غير ملفات التطبيق نفسه.
   ====================================================================== */
const APP_VERSION = '1.1.0';
const KEY = 'sayyarati.state.v1';
const NATIVE_FILE = 'sayyarati-state.json';
const MAX_IMPORT_BYTES = 20 * 1024 * 1024;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

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

/* ---------- templates ---------- */
const COL = (id, label, type) => ({ id, label, type });
const SECTION_TEMPLATES = {
  oil: () => ({ id: 'oil', name: 'الزيوت', icon: 'oil', note: '', reminderKm: 7500,
    columns: [COL('desc','بيان','text'), COL('odo','العداد','odometer'), COL('dist','المسافة','distance'), COL('date','التاريخ','date'), COL('amount','المبلغ','money'), COL('oil','الزيت','text'), COL('notes','ملاحظات','text')], rows: [] }),
  maint: () => ({ id: 'maint', name: 'الصيانة', icon: 'wrench', note: '', reminderKm: 0,
    columns: [COL('desc','بيان','text'), COL('odo','العداد','odometer'), COL('date','التاريخ','date'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  fuel: () => ({ id: 'fuel', name: 'الوقود', icon: 'fuel', note: '', reminderKm: 0,
    columns: [COL('date','التاريخ','date'), COL('odo','العداد','odometer'), COL('dist','المسافة','distance'), COL('liters','اللترات','number'), COL('amount','المبلغ','money'), COL('station','المحطة','text')], rows: [] }),
  tires: () => ({ id: 'tires', name: 'الإطارات', icon: 'tire', note: '', reminderKm: 0,
    columns: [COL('desc','بيان','text'), COL('odo','العداد','odometer'), COL('date','التاريخ','date'), COL('brand','النوع والمقاس','text'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  ins: () => ({ id: 'ins', name: 'التأمين والترخيص', icon: 'shield', note: '', reminderKm: 0,
    columns: [COL('desc','بيان','text'), COL('date','التاريخ','date'), COL('expiry','تاريخ الانتهاء','date'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  wash: () => ({ id: 'wash', name: 'الغسيل والعناية', icon: 'drop', note: '', reminderKm: 0,
    columns: [COL('desc','بيان','text'), COL('date','التاريخ','date'), COL('amount','المبلغ','money'), COL('notes','ملاحظات','text')], rows: [] }),
  blank: () => ({ id: uid(), name: 'قسم جديد', icon: 'star', note: '', reminderKm: 0,
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
const LIMITS = { cars: 50, cardFields: 60, sections: 40, columns: 40, rows: 20000, text: 2000, label: 80 };
const ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
const COLOR_RE = /^#[0-9A-Fa-f]{6}$/;
const ISO_RE = /^\d{4}-\d{2}-\d{2}$/;
const isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const arr = v => (Array.isArray(v) ? v : []);
const str = (v, max) => (v == null ? '' : typeof v === 'object' ? '' : String(v)).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, max);
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
    rows.push({ id: cleanId(r.id, usedRow), t: finite(r.t, Date.now()), v });
  }
  columns.forEach(c => delete c._src);
  return {
    id: cleanId(s.id, usedSec),
    name: str(s.name, LIMITS.label).trim() || 'قسم',
    icon: typeof s.icon === 'string' && has(ICONS, s.icon) ? s.icon : 'doc',
    note: str(s.note, 200),
    reminderKm: Math.min(1e6, Math.max(0, finite(typeof s.reminderKm === 'string' ? toNum(s.reminderKm) : s.reminderKm, 0))),
    columns, rows
  };
}
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
function sanitizeSettings(x) {
  const s = isObj(x) ? x : {};
  const dec = Number(s.decimals);
  return {
    currency: str(s.currency, 12).trim() || 'د.ك',
    decimals: [0, 2, 3].includes(dec) ? dec : 3,
    theme: ['auto', 'light', 'dark'].includes(s.theme) ? s.theme : 'auto'
  };
}
function sanitizeCars(list) { const used = new Set(); return arr(list).slice(0, LIMITS.cars).map(c => sanitizeCar(c, used)).filter(Boolean); }
function sanitizeState(x) {
  const s = isObj(x) ? x : {};
  const cars = sanitizeCars(s.cars);
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
let saveTimer = null, warnedSave = false;
function save() {
  S.rev = (S.rev || 0) + 1;
  const json = JSON.stringify(S);
  let ok = true;
  try { localStorage.setItem(KEY, json); } catch (e) { ok = false; }
  if (isNative) {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const fs = plug('Filesystem');
      if (fs) fs.writeFile({ path: NATIVE_FILE, data: JSON.stringify(S), directory: 'DATA', encoding: 'utf8', recursive: true }).catch(e => { console.warn(e); saveFailed(); });
    }, 150);
  } else if (!ok) saveFailed();
}
function saveFailed() { if (warnedSave) return; warnedSave = true; toast('تعذّر حفظ البيانات على الجهاز — صدّر نسخة احتياطية الآن'); }
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
function firstOdo(c) { const p = odoPoints(c).map(x => x.odo); return p.length ? Math.min(...p) : 0; }
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
function reminders(c) {
  const out = [];
  const cur = currentOdo(c);
  for (const s of c.sections) {
    const o = colOf(s, 'odometer');
    if (!o || !(s.reminderKm > 0)) continue;
    const rows = chrono(s).filter(r => toNum(r.v[o.id]) != null);
    if (!rows.length) { out.push({ sec: s, none: true }); continue; }
    const lastOdo = toNum(rows[rows.length - 1].v[o.id]);
    const used = Math.max(0, cur - lastOdo);
    out.push({ sec: s, lastOdo, used, next: lastOdo + s.reminderKm, left: lastOdo + s.reminderKm - cur, ratio: used / s.reminderKm });
  }
  return out;
}
/* expiry reminders (e.g. insurance) */
function expiries(c) {
  const out = []; const t = todayISO();
  for (const s of c.sections) for (const col of s.columns) {
    if (col.type !== 'date' || !/انتهاء|ينتهي|صلاحية|expiry/i.test(col.label)) continue;
    let latest = null;
    for (const r of s.rows) { const v = r.v[col.id]; if (ISO_RE.test(v) && (!latest || v > latest.v)) latest = { v, r }; }
    if (latest) out.push({ sec: s, col, date: latest.v, days: Math.round(daysBetween(t, latest.v)), title: (textCol(s) && latest.r.v[textCol(s).id]) || s.name });
  }
  return out;
}

/* ---------- UI helpers ---------- */
let toastT;
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('on'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('on'), 2800); }
function openSheet(title, html, onMount) {
  $('#sheetTitle').textContent = title;
  $('#sheetBody').innerHTML = html;
  $('#sheetBody').scrollTop = 0;
  $('#scrim').classList.add('on'); $('#sheet').classList.add('on');
  if (onMount) onMount($('#sheetBody'));
}
function closeSheet() { $('#scrim').classList.remove('on'); $('#sheet').classList.remove('on'); }
const sheetOpen = () => $('#sheet').classList.contains('on');
function confirmBox(msg, okLabel, onOk, danger = true) {
  openSheet('تأكيد', `<p style="margin:4px 0 20px">${esc(msg)}</p>
    <div class="btn-row"><button class="btn ghost" id="cNo">إلغاء</button><button class="btn ${danger ? 'danger' : ''}" id="cYes">${esc(okLabel)}</button></div>`, b => {
    b.querySelector('#cNo').onclick = closeSheet;
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
  ({ home: renderHome, logs: renderLogs, stats: renderStats, settings: renderSettings })[view]();
  $('#v-' + view).classList.add('on');
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
    let eta = '';
    if (rate && x.left > 0) { const d = Math.round(x.left / rate); eta = d <= 1 ? 'خلال يوم تقريباً' : `بعد نحو ${d} يوماً حسب معدل قيادتك`; }
    const head = x.left > 0 ? `باقي ${km(x.left)}` : `متأخر ${km(-x.left)}`;
    dueHTML += `<div class="due">${gaugeSVG(x.ratio)}<div class="due-txt"><small>${esc(x.sec.name)} القادم عند ${km(x.next)}</small><b>${head}</b><small>${eta || `قطعت ${km(x.used)} من ${km(x.sec.reminderKm)}`}</small></div></div>`;
  }
  for (const e of exp) {
    const head = e.days >= 0 ? `ينتهي بعد ${e.days} يوماً` : `انتهى منذ ${-e.days} يوماً`;
    dueHTML += `<div class="due"><div class="due-txt"><small>${esc(e.title)}</small><b style="color:${e.days < 0 ? 'var(--bad)' : e.days < 30 ? 'var(--warn)' : 'inherit'}">${head}</b><small>${dateHTML(e.date)}</small></div></div>`;
  }

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
    <div class="odo-label">العداد الحالي</div>
    <div class="odo-wrap">
      <div style="display:flex;align-items:flex-end">${odoHTML(cur)}<span class="odo-unit">كم</span></div>
      <button class="odo-btn" id="hOdo">تحديث العداد</button>
    </div>
  </div>
  ${dueHTML ? `<div class="block lift due-list">${dueHTML}</div>` : ''}
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
  $('#v-home').querySelectorAll('[data-sec][data-row]').forEach(b => b.onclick = () => openEntry(b.dataset.sec, b.dataset.row));
}
function recentItem(s, r) {
  const t = textCol(s), d = colOf(s, 'date'), m = s.columns.find(c => c.type === 'money');
  return `<button class="list-item" data-sec="${esc(s.id)}" data-row="${esc(r.id)}"><span class="ic">${icon(s.icon)}</span>
    <span class="grow"><b style="font-weight:500">${esc((t && r.v[t.id]) || s.name)}</b><small>${esc(s.name)}${d && r.v[d.id] ? ` · ${dateHTML(r.v[d.id])}` : ''}</small></span>
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
  draft = { carId: c.id, secId: secId || activeSec || (c.sections[0] && c.sections[0].id), rowId: rowId || null, v: {} };
  if (rowId) { const s = c.sections.find(x => x.id === draft.secId); const r = s && s.rows.find(x => x.id === rowId); if (r) draft.v = clone(r.v); else draft.rowId = null; }
  renderEntry();
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
  openSheet(editing ? `تعديل — ${s ? s.name : ''}` : 'تسجيل جديد', `${carsPick}${secPick}${fields}
    ${s ? `<div class="btn-row" style="margin-top:6px">${editing ? `<button class="btn danger" id="eDel">حذف</button>` : ''}<button class="btn amber" id="eSave" style="flex:2">${editing ? 'حفظ التعديل' : 'حفظ'}</button></div>` : ''}`, b => {
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
  if (!Object.values(v).some(x => x !== '')) { toast('أدخل بياناً واحداً على الأقل'); return; }
  if (draft.rowId) { const r = s.rows.find(x => x.id === draft.rowId); if (r) r.v = v; }
  else { if (s.rows.length >= LIMITS.rows) { toast('وصل القسم إلى الحد الأقصى للتسجيلات'); return; } s.rows.push({ id: uid(), t: Date.now(), v }); }
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
    ${c ? `<button class="btn danger wide" id="ceDel" style="margin-top:10px">حذف السيارة وكل سجلاتها</button>` : ''}`, b => {
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
        const apply = target => { const old = target.cardFields; target.cardFields = work.map(f => { const o = old.find(x => x.id === f.id); return { id: f.id, label: f.label, type: f.type, value: o ? cleanValue(f.type, o.value) : '' }; }); };
        apply(c);
        if (b.querySelector('#cfAll').checked) S.cars.forEach(x => { if (x !== c) apply(x); });
        save(); closeSheet(); render(); toast('حُفظت البنود');
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
/* values of deleted items are removed too, so they do not linger in the data or in backups */
function purgeRemoved(sec) {
  const keep = new Set(sec.columns.map(c => c.id));
  for (const r of sec.rows) for (const k of Object.keys(r.v)) if (!keep.has(k)) delete r.v[k];
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
      ${work.columns.some(x => x.type === 'odometer') ? `<label class="field"><span>تذكير كل (كم) — 0 لإيقافه</span><input id="seRem" class="num" type="text" inputmode="numeric" value="${esc(nf(work.reminderKm || 0).replace(/,/g, ''))}"><div class="hint">يظهر في الصفحة الرئيسية كم باقي على الموعد القادم.</div></label>` : ''}
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
      readItems(b, work.columns);
    };
    bindItems(b, work.columns, sync, () => draw(b));
    b.querySelectorAll('[data-ic]').forEach(x => x.onclick = () => { sync(); if (has(ICONS, x.dataset.ic)) work.icon = x.dataset.ic; draw(b); });
    b.querySelector('#seAdd').onclick = () => { sync(); if (work.columns.length >= LIMITS.columns) return; work.columns.push({ id: uid(), label: 'بند جديد', type: 'text' }); draw(b); };
    b.querySelector('#seSave').onclick = () => {
      sync();
      if (work.columns.some(x => x.type === 'distance') && !work.columns.some(x => x.type === 'odometer')) { toast('بند المسافة يحتاج بنداً من نوع العداد'); return; }
      Object.assign(s, { name: work.name, note: work.note, icon: work.icon, reminderKm: work.reminderKm, columns: clone(work.columns) });
      purgeRemoved(s);
      if (b.querySelector('#seAll').checked) S.cars.forEach(x => {
        if (x === c) return;
        const t = x.sections.find(y => y.id === s.id);
        if (t) { Object.assign(t, { name: s.name, note: s.note, icon: s.icon, reminderKm: s.reminderKm, columns: clone(s.columns) }); purgeRemoved(t); }
        else if (x.sections.length < LIMITS.sections) x.sections.push({ ...clone(s), note: '', rows: [] });
      });
      save(); closeSheet(); render(); toast('حُفظ القسم');
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
    <div class="block"><h2>النسخة الاحتياطية</h2>
      <p style="margin:0 0 14px;color:var(--muted);font-size:14px">ملف إكسل منسّق بكل السيارات والأقسام، يُحفظ في الدرايف أو أي مكان من قائمة المشاركة، ويمكن استرجاعه في أي وقت.${last && !isNaN(last) ? `<br>آخر نسخة: ${dateHTML(last.toISOString().slice(0, 10))}` : ''}</p>
      <div class="btn-row"><button class="btn amber" id="sExp">تصدير إكسل</button><button class="btn ghost" id="sImp">استيراد / استرجاع</button></div>
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
  $('#sExp').onclick = openBackup;
  $('#sImp').onclick = pickImport;
  $('#sCard').onclick = () => openCardEditor(true);
  $('#sAddSec').onclick = openAddSection;
  $('#sAddCar').onclick = () => openCarEditor(null);
  $('#v-settings').querySelectorAll('[data-sec-i]').forEach(b => b.onclick = () => { const s = c.sections[+b.dataset.secI]; if (s) openSectionEditor(s.id); });
  $('#v-settings').querySelectorAll('[data-car-i]').forEach(b => b.onclick = () => { const x = S.cars[+b.dataset.carI]; if (x) openCarEditor(x.id); });
  $('#sCur').onchange = e => { S.settings.currency = str(e.target.value, 12).trim() || 'د.ك'; save(); };
  $('#sDec').onchange = e => { S.settings = sanitizeSettings({ ...S.settings, decimals: +e.target.value }); save(); };
  $('#sTheme').onchange = e => { S.settings = sanitizeSettings({ ...S.settings, theme: e.target.value }); save(); applyTheme(); };
}

/* ======================================================================
   Excel — export / import (ExcelJS, loaded on demand from the app itself)
   ====================================================================== */
let excelLoading = null;
function loadExcel() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  return excelLoading || (excelLoading = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = 'vendor/exceljs.min.js';
    s.onload = () => res(window.ExcelJS); s.onerror = () => { excelLoading = null; rej(new Error('تعذّر تحميل مكتبة الإكسل')); };
    document.head.appendChild(s);
  }));
}
const XL = { ink: 'FF12262E', amber: 'FFE9A23B', amberSoft: 'FFFBEBD0', zebra: 'FFF2F5F4', line: 'FFCAD5D2', white: 'FFFFFFFF', muted: 'FF62747A' };
const thin = { style: 'thin', color: { argb: XL.line } };
const border = { top: thin, left: thin, bottom: thin, right: thin };
function sheetName(base, used) {
  const n = base.replace(/[\[\]:*?\/\\']/g, ' ').trim().slice(0, 31) || 'ورقة';
  let k = 2, out = n; while (used.has(out.toLowerCase())) { const suf = ` (${k++})`; out = n.slice(0, 31 - suf.length) + suf; }
  used.add(out.toLowerCase()); return out;
}
const isoToDate = iso => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const colLetter = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const moneyFmt = () => S.settings.decimals ? '#,##0.' + '0'.repeat(S.settings.decimals) : '#,##0';
const argb = hex => 'FF' + (COLOR_RE.test(hex) ? hex.slice(1).toUpperCase() : 'E9A23B');

async function buildWorkbook(cars) {
  const ExcelJS = await loadExcel();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'سيّارتي'; wb.created = new Date();
  const used = new Set(['_sayyarati']);
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
  /* -- restore data (hidden sheet) -- */
  const hid = wb.addWorksheet('_sayyarati', { state: 'veryHidden' });
  const payload = JSON.stringify({ app: 'sayyarati', v: 1, exported: new Date().toISOString(), settings: S.settings, cars });
  hid.getCell(1, 1).value = 'sayyarati-backup-v1';
  for (let i = 0, row = 2; i < payload.length; i += 30000, row++) hid.getCell(row, 1).value = payload.slice(i, i + 30000);
  return wb;
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

function openBackup() {
  if (!S.cars.length) return;
  const c = car();
  openSheet('نسخة احتياطية إكسل', `
    <p style="margin:0 0 14px;color:var(--muted);font-size:14px">ستظهر قائمة المشاركة: اختر «Drive» للحفظ في جوجل درايف، أو «الملفات»، أو أرسله بالواتساب أو البريد.</p>
    ${S.cars.length > 1 ? `<div class="pick"><button class="chip" data-w="all" aria-pressed="true">كل السيارات (${S.cars.length})</button><button class="chip" data-w="one" aria-pressed="false">${esc(c.name)} فقط</button></div>` : ''}
    <button class="btn amber wide" id="bShare">${isNative ? 'تصدير ومشاركة' : 'تصدير ومشاركة / حفظ'}</button>
    ${!isNative ? `<button class="btn ghost wide" id="bDl" style="margin-top:10px">تنزيل الملف فقط</button>` : ''}
    <div class="sub-h">الاسترجاع</div>
    <p style="margin:0 0 12px;color:var(--muted);font-size:14px">اختر ملف إكسل صدّرته من التطبيق لاسترجاع كل شيء، أو أي ملف إكسل بنفس ترتيب ملفك القديم لاستيراده كسيارة.${isNative ? ' ويمكنك أيضاً فتح الملف من تطبيق Drive أو الملفات واختيار «فتح باستخدام ← سيّارتي».' : ''}</p>
    <button class="btn ghost wide" id="bImp">استيراد / استرجاع من ملف</button>`, b => {
    let which = 'all';
    b.querySelectorAll('[data-w]').forEach(x => x.onclick = () => { which = x.dataset.w === 'one' ? 'one' : 'all'; b.querySelectorAll('[data-w]').forEach(y => y.setAttribute('aria-pressed', y === x)); });
    const run = async mode => {
      const btn = b.querySelector(mode === 'dl' ? '#bDl' : '#bShare'); const old = btn.textContent; btn.textContent = 'جارٍ تجهيز الملف…'; btn.disabled = true;
      try {
        const cars = which === 'one' ? [c] : S.cars;
        const wb = await buildWorkbook(cars);
        const buf = await wb.xlsx.writeBuffer();
        const name = `سيارتي-${which === 'one' || cars.length === 1 ? safeName(cars[0].name) + '-' : ''}${todayISO()}.xlsx`;
        const ok = await deliver(new Blob([buf], { type: XLSX_MIME }), name, mode);
        if (ok) { S.lastBackup = Date.now(); save(); closeSheet(); render(); toast('جُهّز ملف الإكسل'); }
      } catch (e) { console.error(e); toast('تعذّر إنشاء الملف: ' + String((e && e.message) || e).slice(0, 120)); }
      btn.textContent = old; btn.disabled = false;
    };
    b.querySelector('#bShare').onclick = () => run('share');
    if (b.querySelector('#bDl')) b.querySelector('#bDl').onclick = () => run('dl');
    b.querySelector('#bImp').onclick = pickImport;
  });
}
const safeName = s => String(s).replace(/[\\/:*?"<>|\u0000-\u001F\s]+/g, '-').replace(/^\.+/, '').slice(0, 40) || 'car';
function bufToB64(buf) { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
function b64ToBuf(b64) { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u.buffer; }
async function deliver(blob, name, mode) {
  if (isNative) {
    const fs = plug('Filesystem'), sh = plug('Share');
    const w = await fs.writeFile({ path: 'exports/' + name, data: bufToB64(await blob.arrayBuffer()), directory: 'CACHE', recursive: true });
    try { await sh.share({ title: name, dialogTitle: 'حفظ النسخة الاحتياطية', files: [w.uri] }); }
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
      if (r && !r.cancelled && r.data) await importBuffer(b64ToBuf(r.data), r.name);
    } catch (e) { showImportError(nativeError(e)); }
    finally { importBusy = false; }
    return;
  }
  const f = $('#fileIn'); f.value = ''; f.click();
}
function nativeError(e) {
  if (e instanceof ImportError) return e;
  const code = e && e.code, msg = String((e && e.message) || e || '');
  if (code === 'TOO_LARGE') return new ImportError('too-large');
  if (code === 'NO_PERMISSION' || code === 'READ_FAILED') return new ImportError('unreadable', msg);
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
    'too-large': 'الملف أكبر من 20 ميجابايت، وهذا أكبر بكثير من أي سجل سيارة. تأكد أنك اخترت الملف الصحيح.',
    'unreadable': 'لم يتمكن الهاتف من قراءة الملف من مكانه. يحدث هذا أحياناً مع ملفات الدرايف التي لم تُنزَّل بعد. جرّب أحد الحلين: افتح تطبيق Drive، واضغط ⋮ بجانب الملف ثم «فتح باستخدام» ← «سيّارتي»، أو نزّل الملف إلى الهاتف ثم اختره من «التنزيلات».',
    'not-xlsx': 'هذا الملف ليس بصيغة Excel الحديثة (‎.xlsx‎). إن كان بصيغة ‎.xls‎ القديمة أو جدول Google Sheets، افتحه واحفظه أو نزّله بصيغة ‎.xlsx‎ ثم استورده.',
    'sheets-export': 'هذا جدول Google Sheets ولم يسمح الدرايف بتحويله تلقائياً. افتحه في Google Sheets واختر «تنزيل ← Microsoft Excel (‎.xlsx‎)» ثم استورد الملف الناتج.',
    'bad-backup': 'الملف فيه نسخة احتياطية تالفة أو ناقصة، فلم يُستورد شيء.',
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
    for (let r = 2; r <= Math.min(hid.rowCount, 2000); r++) json += String(cellVal(hid.getCell(r, 1)) || '');
    let data;
    try { data = JSON.parse(json); } catch (e) { throw new ImportError('bad-backup'); }
    if (!isObj(data) || !Array.isArray(data.cars)) throw new ImportError('bad-backup');
    return askRestore(data);
  }
  const c = sanitizeCar(genericImport(wb, fname), new Set(S.cars.map(x => x.id)));
  if (!c || (!c.sections.some(s => s.rows.length) && !c.cardFields.some(f => f.value !== ''))) throw new ImportError('empty');
  const n = c.sections.reduce((a, s) => a + s.rows.length, 0);
  openSheet('استيراد سيارة', `<p style="margin:0 0 6px">وجدت في الملف:</p>
    <dl class="kv" style="margin-bottom:18px"><dt>السيارة</dt><dd>${esc(c.name)}</dd>${c.sections.map(s => `<dt>${esc(s.name)}</dt><dd>${s.rows.length} تسجيل</dd>`).join('')}<dt>بنود البطاقة</dt><dd>${c.cardFields.length}</dd></dl>
    <label class="field"><span>اسم السيارة</span><input id="giName" maxlength="${LIMITS.label}" value="${esc(c.name)}"></label>
    <button class="btn amber wide" id="giOk">إضافة السيارة (${n} تسجيل)</button>`, b => {
    b.querySelector('#giOk').onclick = () => {
      if (S.cars.length >= LIMITS.cars) { toast('وصلت إلى الحد الأقصى لعدد السيارات'); return; }
      c.name = str(b.querySelector('#giName').value, LIMITS.label).trim() || c.name;
      S.cars.push(c); S.activeCar = c.id; save(); closeSheet(); rolledFor = null; go('home'); toast(`استُوردت ${c.name}`);
    };
  });
}
function askRestore(raw) {
  const cars = sanitizeCars(raw.cars);
  const settings = sanitizeSettings(raw.settings);
  const n = cars.reduce((a, c) => a + c.sections.reduce((x, s) => x + s.rows.length, 0), 0);
  const when = parseDateAny(String(raw.exported || '').slice(0, 10));
  if (!cars.length) throw new ImportError('bad-backup');
  openSheet('استرجاع نسخة احتياطية', `<p style="margin:0 0 16px">${when ? `النسخة بتاريخ ${dateHTML(when)} ` : 'النسخة '}وفيها ${cars.length} ${cars.length === 1 ? 'سيارة' : 'سيارات'} و${n} تسجيل.</p>
    <button class="btn amber wide" id="rMerge">دمج مع بياناتي الحالية</button>
    <p class="hint" style="font-size:13px;color:var(--muted);margin:6px 2px 14px">السيارة الموجودة في النسختين تُستبدل بنسخة الملف، وتبقى السيارات الأخرى.</p>
    <button class="btn danger wide" id="rReplace">استبدال كل بياناتي بالنسخة</button>`, b => {
    b.querySelector('#rMerge').onclick = () => {
      const merged = S.cars.slice();
      for (const c of cars) { const i = merged.findIndex(x => x.id === c.id); if (i >= 0) merged[i] = c; else merged.push(c); }
      finishRestore(merged.slice(0, LIMITS.cars), settings);
    };
    b.querySelector('#rReplace').onclick = () => confirmBox('ستُحذف كل البيانات الحالية وتحل محلها النسخة. متابعة؟', 'استبدال', () => finishRestore(cars, settings));
  });
}
function finishRestore(cars, settings) {
  S = sanitizeState({ ...S, cars, settings });
  save(); applyTheme(); closeSheet(); rolledFor = null; go('home'); toast('استُرجعت النسخة');
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
    const sec = { id: tmpl && !c.sections.some(s => s.id === tmpl) ? tmpl : uid(), name: str(ws.name, LIMITS.label).trim(), icon: tmpl === 'oil' ? 'oil' : tmpl === 'maint' ? 'wrench' : /وقود|بنزين/.test(ws.name) ? 'fuel' : /تأمين|ترخيص/.test(ws.name) ? 'shield' : 'doc', note: '', reminderKm: tmpl === 'oil' ? 7500 : 0, columns: [], rows: [] };
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
    if (r && !r.none && r.data) await importBuffer(b64ToBuf(r.data), r.name);
  } catch (e) { showImportError(nativeError(e)); }
  finally { takingOpened = false; }
}

/* ---------- native back button ---------- */
function initNative() {
  const app = plug('App');
  if (app) app.addListener('backButton', () => {
    if (sheetOpen()) { closeSheet(); return; }
    if (view !== 'home') { go('home'); return; }
    app.minimizeApp ? app.minimizeApp() : app.exitApp();
  });
  const p = xlsxPlugin();
  if (p && p.addListener) p.addListener('fileOpened', () => takeOpenedFile());
  takeOpenedFile();
  cleanOldExports();
}

/* ---------- boot ---------- */
document.querySelectorAll('#nav button[data-v]').forEach(b => b.onclick = () => go(b.dataset.v));
$('#fab').onclick = () => { if (!S.cars.length) return; openEntry(); };
$('#carPill').onclick = () => openCarSwitcher();
$('#quickExport').onclick = () => openBackup();
$('#scrim').onclick = closeSheet;
$('#sheetClose').onclick = closeSheet;

(async function boot() {
  let raw = loadLocal();
  if (isNative) { const n = await loadNative(); if (n && (!raw || finite(n.rev) > finite(raw && raw.rev))) raw = n; }
  try { S = sanitizeState(raw); } catch (e) { console.error('state could not be read', e); S = blankState(); }
  applyTheme();
  try { render(); } catch (e) { console.error(e); S = sanitizeState(S); render(); }
  if (isNative) initNative();
  if (!isNative && 'serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
})();
