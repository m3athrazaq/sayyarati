/* The web app run as if inside the Android app: window.Capacitor with stand-in plugins (files kept in
   memory here, notifications recorded). Checks the code paths only the phone app uses: imports through
   the native file picker, invoice photos saved as files, files handed to the share sheet in pieces, and
   the reminder notifications (permission, what is scheduled, tapping one).
   The Android side itself (Java/Kotlin) is checked by the build and by the emulator test.
   Run: python3 -m http.server 8765 (in the repo) and `node tests/native.test.js` (needs playwright). */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const BASE = process.env.BASE || 'http://localhost:8765/';
const FIXTURE = path.join(__dirname, '..', 'native', 'selftest', 'fixture-f150-layout.xlsx');
const results = [];
const ok = (cond, msg) => results.push((cond ? 'PASS ' : 'FAIL ') + msg);

/* the phone's files: "DIR/path" → Buffer */
const files = new Map();
const calls = [];
const key = (dir, p) => `${dir}/${String(p).replace(/^\/+/, '')}`;
function fsCall(method, o) {
  calls.push(method);
  const k = key(o.directory, o.path);
  const uri = 'file:///mock/' + k;
  switch (method) {
    case 'writeFile':
    case 'appendFile': {
      const buf = o.encoding ? Buffer.from(o.data, 'utf8') : Buffer.from(o.data, 'base64');
      if (method === 'appendFile') { if (!files.has(k)) throw new Error('File does not exist'); files.set(k, Buffer.concat([files.get(k), buf])); }
      else files.set(k, buf);
      return { uri };
    }
    case 'readFile': {
      if (!files.has(k)) throw new Error('File does not exist');
      return { data: o.encoding ? files.get(k).toString('utf8') : files.get(k).toString('base64') };
    }
    case 'deleteFile': if (!files.delete(k)) throw new Error('File does not exist'); return {};
    case 'readdir': {
      const pre = k.replace(/\/?$/, '/');
      const list = [...files.keys()].filter(f => f.startsWith(pre) && !f.slice(pre.length).includes('/'));
      if (!list.length) throw new Error('Folder does not exist');
      return { files: list.map(f => ({ name: f.slice(pre.length), type: 'file', size: files.get(f).length, mtime: Date.now(), uri: 'file:///mock/' + f })) };
    }
    case 'getUri': return { uri };
    default: throw new Error('unknown ' + method);
  }
}

/* page side: Capacitor and its plugins, before app.js runs */
function fakeCapacitor() {
  const listeners = {};
  const on = name => (ev, cb) => { (listeners[name + ':' + ev] = listeners[name + ':' + ev] || []).push(cb); return Promise.resolve({ remove() {} }); };
  const fsm = {};
  for (const m of ['writeFile', 'appendFile', 'readFile', 'deleteFile', 'readdir', 'getUri']) {
    fsm[m] = async o => { try { return await window.__fs(m, JSON.parse(JSON.stringify(o))); } catch (e) { throw new Error(String(e.message || e).replace(/^Error: /, '')); } };
  }
  const ln = { perm: 'prompt', exact: 'denied', pending: [], scheduled: [], channels: [] };
  window.__ln = ln; window.__listeners = listeners; window.__shared = [];
  window.Capacitor = {
    isNativePlatform: () => true,
    convertFileSrc: u => location.origin + u.replace('file://', '/_capacitor_file_'),
    Plugins: {
      Filesystem: fsm,
      Share: { share: async o => { window.__shared.push(o); return {}; } },
      App: { addListener: on('App'), minimizeApp: async () => {} },
      XlsxImport: {
        addListener: on('XlsxImport'),
        takeOpenedFile: async () => ({ none: true }),
        pick: async () => {
          const r = await window.__fs('writeFile', { path: 'imports/import-1.xlsx', directory: 'CACHE', data: window.__fixture });
          return { name: 'F150.xlsx', size: 1, path: 'imports/import-1.xlsx', uri: r.uri };
        }
      },
      LocalNotifications: {
        addListener: on('LN'),
        checkPermissions: async () => ({ display: ln.perm }),
        checkExactNotificationSetting: async () => ({ exact_alarm: ln.exact }),
        requestPermissions: async () => { ln.perm = 'granted'; return { display: ln.perm }; },
        createChannel: async o => { ln.channels.push(o); },
        getPending: async () => ({ notifications: ln.pending.map(n => ({ id: n.id })) }),
        cancel: async o => { const ids = new Set(o.notifications.map(n => n.id)); ln.pending = ln.pending.filter(n => !ids.has(n.id)); },
        schedule: async o => {
          if (ln.perm !== 'granted') throw new Error('Notifications not enabled');
          const list = JSON.parse(JSON.stringify(o.notifications));   /* as the bridge sends them (dates become text) */
          ln.pending.push(...list); ln.scheduled.push(list.length);
          return { notifications: list.map(n => ({ id: n.id })) };
        }
      }
    }
  };
}

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, locale: 'ar' });
  await ctx.exposeFunction('__fs', (m, o) => fsCall(m, o));
  await ctx.route('**/_capacitor_file_/**', route => {
    const k = decodeURIComponent(new URL(route.request().url()).pathname).replace('/_capacitor_file_/mock/', '');
    const buf = files.get(k);
    if (!buf) return route.fulfill({ status: 404, body: '' });
    return route.fulfill({ status: 200, body: buf, contentType: /\.jpg$/.test(k) ? 'image/jpeg' : 'application/octet-stream' });
  });
  await ctx.addInitScript({ content: `window.__fixture = ${JSON.stringify(fs.readFileSync(FIXTURE).toString('base64'))};(${fakeCapacitor.toString()})();` });
  const p = await ctx.newPage();
  const errors = [], csp = [];
  p.on('pageerror', e => errors.push('page: ' + e.message));
  p.on('console', m => { const t = m.text(); if (/Content Security Policy|Refused to/.test(t)) csp.push(t); else if (m.type() === 'error') errors.push(t); });
  const ln = () => p.evaluate(() => window.__ln);
  const shot = name => (process.env.SHOT_DIR ? p.screenshot({ path: path.join(process.env.SHOT_DIR, name + '.png'), fullPage: true }) : null);
  try {
    await p.goto(BASE + 'index.html'); await p.waitForTimeout(800);
    ok(await p.evaluate(() => isNative), 'app runs in its phone mode');
    ok((await ln()).channels.some(c => c.id === 'reminders' && c.name === 'التذكيرات'), 'notification channel «التذكيرات» created');

    /* import through the native picker: the copy in the cache is read and deleted */
    await p.click('#wImport'); await p.waitForTimeout(1500);
    ok(await p.isVisible('#giOk'), 'file from the native picker shows the import summary');
    ok(!files.has('CACHE/imports/import-1.xlsx'), 'the cached copy of the import was deleted');
    await p.click('#giOk'); await p.waitForTimeout(1000);

    /* reminders: the home card asks for the permission, then the notifications are scheduled */
    ok(await p.isVisible('#hNotifyOn'), 'home asks to turn on notifications');
    await shot('native-home');
    await p.click('#hNotifyOn'); await p.waitForTimeout(800);
    let n = await ln();
    const at = n.pending.map(x => x.schedule && x.schedule.at);
    ok(n.perm === 'granted' && n.pending.length >= 1 && !(await p.isVisible('#hNotifyOn')), `notifications scheduled after permission (${n.pending.length})`);
    ok(n.pending.every(x => Number.isInteger(x.id) && x.id > 0 && x.id < 2 ** 31 && x.channelId === 'reminders' && x.smallIcon === 'ic_stat_sayyarati' && x.isExactNotification === false && x.schedule.allowWhileIdle === true && typeof x.extra.car === 'string')
      && at.every(a => /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(a) && Date.parse(a) > Date.now()), 'each notification: valid id, channel, icon, inexact alarm, future time');
    /* Android 11 and older: exact alarms (they need no permission there) */
    await p.evaluate(async () => { window.__ln.exact = 'granted'; await syncNotifications(true); });
    ok((await ln()).pending.every(x => x.isExactNotification === true), 'exact alarms where Android allows them without a permission');
    await p.evaluate(async () => { window.__ln.exact = 'denied'; await syncNotifications(true); });
    n = await ln();
    /* saving the same data again does not reschedule */
    const before = n.scheduled.length;
    await p.evaluate(() => { save(); }); await p.waitForTimeout(2200);
    ok((await ln()).scheduled.length === before, 'unchanged reminders are not scheduled again');

    /* an invoice photo saved as a file, shown through the app's file address */
    await p.click('#fab'); await p.waitForTimeout(400);
    await p.fill('#sheetBody input[type=text]:not(.num) >> nth=0', 'تبديل زيت');
    const odo = await p.$('#sheetBody input[inputmode=numeric]');
    if (odo) await odo.fill('103500');
    const img = await p.evaluate(async () => { const cv = document.createElement('canvas'); cv.width = 800; cv.height = 1000; cv.getContext('2d').fillRect(10, 10, 300, 300);
      const u = new Uint8Array(await (await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.9))).arrayBuffer()); let s = ''; for (const b of u) s += String.fromCharCode(b); return s; });
    const [fc] = await Promise.all([p.waitForEvent('filechooser'), p.click('#phCam')]);
    ok(await p.evaluate(() => document.querySelector('#photoCam').getAttribute('capture') === 'environment'), '«تصوير» opens the camera');
    await fc.setFiles([{ name: 'cam.jpg', mimeType: 'image/jpeg', buffer: Buffer.from(img, 'latin1') }]);
    await p.waitForFunction(() => { const i = document.querySelector('#phList img'); return i && i.naturalWidth > 0; }, null, { timeout: 15000 }).catch(() => {});
    await p.evaluate(() => document.querySelector('#phList').scrollIntoView());
    await shot('native-entry');
    const photoFiles = [...files.keys()].filter(k => k.startsWith('DATA/photos/'));
    ok(photoFiles.length === 1 && files.get(photoFiles[0])[0] === 0xFF && files.get(photoFiles[0])[1] === 0xD8, 'photo saved as a JPEG file in the app storage');
    ok(await p.evaluate(() => { const i = document.querySelector('#phList img'); return !!i && i.naturalWidth === 800 && i.src.includes('/_capacitor_file_/'); }), 'thumbnail loaded from the app file address');
    await p.click('#eSave'); await p.waitForTimeout(800);
    const stateFile = JSON.parse((files.get('DATA/sayyarati-state.json') || Buffer.from('{}')).toString('utf8'));
    ok(stateFile.cars && stateFile.cars[0].sections.some(s => s.rows.some(r => r.p && r.p.length === 1)), 'entry with its photo saved in the app file');
    await p.waitForTimeout(1800);
    n = await ln();
    ok(n.scheduled.length > before, 'a new oil change reschedules the reminders');

    /* viewer + back button */
    await p.click('#nav [data-v=logs]'); await p.waitForTimeout(300);
    await p.click('.entry .clip'); await p.waitForTimeout(400);
    await p.click('#phList [data-pi="0"]'); await p.waitForTimeout(300);
    await p.evaluate(() => window.__listeners['App:backButton'].forEach(cb => cb()));
    await p.waitForTimeout(200);
    ok(!(await p.isVisible('#viewer.on')) && await p.isVisible('#sheet.on'), 'back button closes the photo viewer first');
    await p.evaluate(() => window.__listeners['App:backButton'].forEach(cb => cb()));
    await p.waitForTimeout(300);

    /* Excel backup handed to the share sheet; a large file is written in pieces */
    await p.click('#quickExport'); await p.waitForTimeout(400);
    await p.click('#bShare'); await p.waitForTimeout(2500);
    const shared = await p.evaluate(() => window.__shared.map(s => s.files[0]));
    const xk = shared.length && shared[0].replace('file:///mock/', '');
    const xlsx = xk && files.get(xk);
    ok(!!xlsx && xlsx[0] === 0x50 && xlsx[1] === 0x4B && xlsx.includes(Buffer.from('xl/media/image1.jpeg')), 'Excel backup with the photo shared from the app cache');
    calls.length = 0;
    const big = await p.evaluate(async () => { const b = new Blob([new Uint8Array(7 * 1024 * 1024 + 5).fill(7)]); await deliver(b, 'big.bin', 'share'); return b.size; });
    ok(files.get('CACHE/exports/big.bin').length === big && calls.filter(c => c === 'appendFile').length === 2, 'a 7 MB file is written in 3 pieces');

    /* tapping a notification opens that car */
    await p.evaluate(() => { S.cars.push(newCar('سيارة ثانية')); S.activeCar = S.cars[1].id; save(); render(); });
    const firstId = await p.evaluate(() => S.cars[0].id);
    await p.evaluate(id => window.__listeners['LN:localNotificationActionPerformed'].forEach(cb => cb({ actionId: 'tap', notification: { id: 1, extra: { car: id } } })), firstId);
    await p.waitForTimeout(400);
    ok(await p.evaluate(id => S.activeCar === id && view === 'home', firstId), 'tapping a notification shows its car');
    await p.evaluate(() => window.__listeners['LN:localNotificationActionPerformed'].forEach(cb => cb({ notification: { extra: { car: '<x>' } } })));
    ok(await p.evaluate(id => S.activeCar === id, firstId), 'a notification for an unknown car changes nothing');

    /* turning notifications off clears them */
    await p.click('#nav [data-v=settings]'); await p.waitForTimeout(300);
    ok(await p.isChecked('#sNotify'), 'settings show notifications on');
    await shot('native-settings');
    await p.uncheck('#sNotify'); await p.waitForTimeout(2200);
    ok((await ln()).pending.length === 0, 'turning notifications off cancels them all');
  } catch (e) {
    results.push('ERROR ' + e.message);
  }
  ok(csp.length === 0, 'no security-policy violations ' + csp.slice(0, 2).join(' | '));
  ok(errors.length === 0, 'no JS errors ' + errors.slice(0, 3).join(' | '));
  console.log(results.join('\n'));
  await browser.close();
  process.exit(results.some(r => !r.startsWith('PASS')) ? 1 : 0);
})();
