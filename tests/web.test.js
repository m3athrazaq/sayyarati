/* Browser test of the web app (Chromium, with the page's security policy active).
   Run: python3 -m http.server 8765 (in the repo) and `node tests/web.test.js` (needs playwright).
   Checks the normal import/export/restore flows and that crafted or broken files cannot run code,
   corrupt the data or lock the app. */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const os = require('os');

const BASE = process.env.BASE || 'http://localhost:8765/';
const FIXTURE = path.join(__dirname, '..', 'native', 'selftest', 'fixture-f150-layout.xlsx');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'sayyarati-test-'));
const results = [];
const ok = (cond, msg) => results.push((cond ? 'PASS ' : 'FAIL ') + msg);

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true, locale: 'ar' });
  const p = await ctx.newPage();
  const errors = [], csp = [];
  p.on('pageerror', e => errors.push('page: ' + e.message));
  p.on('console', m => { const t = m.text(); if (/Content Security Policy|Refused to/.test(t)) csp.push(t); else if (m.type() === 'error') errors.push(t); });
  const state = () => p.evaluate(() => JSON.parse(localStorage.getItem('sayyarati.state.v1') || 'null'));
  const pwned = () => p.evaluate(() => window.__pwned);
  const closeSheet = () => p.evaluate(() => document.querySelector('#sheetClose').click()).then(() => p.waitForTimeout(350));
  const pick = async file => {
    const [fc] = await Promise.all([p.waitForEvent('filechooser'), p.evaluate(() => document.querySelector('#fileIn').click())]);
    await fc.setFiles(file); await p.waitForTimeout(1500);
  };
  /* builds a backup file inside the page with the app's own ExcelJS */
  const backupFile = async (name, payload) => {
    const b64 = await p.evaluate(async json => {
      const wb = new window.ExcelJS.Workbook();
      wb.addWorksheet('x').getCell('A1').value = 'visible';
      const h = wb.addWorksheet('_sayyarati', { state: 'veryHidden' });
      h.getCell(1, 1).value = 'sayyarati-backup-v1';
      for (let i = 0, r = 2; i < json.length; i += 30000, r++) h.getCell(r, 1).value = json.slice(i, i + 30000);
      const u = new Uint8Array(await wb.xlsx.writeBuffer()); let s = '';
      for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
      return btoa(s);
    }, payload);
    const f = path.join(TMP, name); fs.writeFileSync(f, Buffer.from(b64, 'base64')); return f;
  };

  try {
    await p.goto(BASE + 'index.html'); await p.waitForTimeout(600);

    /* 1. a file in the old F150 layout */
    await pick(FIXTURE);
    ok(await p.isVisible('#giOk'), 'old-layout file shows the import summary');
    await p.click('#giOk'); await p.waitForTimeout(1200);
    let st = await state();
    ok(st.cars.length === 1 && st.cars[0].sections.map(s => s.rows.length).join('/') === '8/2', 'imported 8/2 rows');
    ok(st.cars[0].name === 'سيارة تجربة' && st.cars[0].cardFields.length === 4, 'name and 4 card items (merged title cell read once)');
    ok(errors.length === 0, 'no errors after import ' + errors.join(' | '));

    /* 2. the section editor keeps every name with its data when items move */
    await p.click('#nav [data-v=settings]'); await p.waitForTimeout(300);
    await p.click('#v-settings [data-sec-i="0"]'); await p.waitForTimeout(300);
    await p.click('.col-edit[data-i="0"] [data-a=down]'); await p.waitForTimeout(150);
    await p.click('#seSave'); await p.waitForTimeout(400);
    st = await state();
    const cols = st.cars[0].sections[0].columns;
    const row0 = st.cars[0].sections[0].rows[0].v;
    ok(cols[0].type === 'odometer' && cols[1].label === 'بيان' && typeof row0[cols[0].id] === 'number' && /تبديل/.test(row0[cols[1].id]), 'moving an item keeps name and values together');

    /* 3. "apply to all cars" keeps the values of a car imported from Excel */
    await p.click('#sAddCar'); await p.waitForTimeout(300);
    await p.fill('#ceName', 'سيارة ثانية'); await p.click('#ceSave'); await p.waitForTimeout(500);
    await p.click('#nav [data-v=settings]'); await p.waitForTimeout(300);
    await p.click('#sCard'); await p.waitForTimeout(300);
    await p.check('#cfAll'); await p.click('#cfSave'); await p.waitForTimeout(400);
    if (await p.isVisible('#cYes')) await p.click('#cYes');
    await p.waitForTimeout(400);
    st = await state();
    const first = st.cars.find(c => c.name === 'سيارة تجربة');
    ok(first && first.cardFields.find(f => /سعر/.test(f.label)).value === 4500, 'apply-to-all kept the imported car\'s purchase price');

    /* 4. text that looks like a formula stays text in the exported Excel; a section called "History"
          (a name Excel reserves) still exports */
    await p.click('#nav [data-v=settings]'); await p.waitForTimeout(200);
    await p.click('#v-settings [data-sec-i="1"]'); await p.waitForTimeout(300);
    await p.fill('#seName', 'History'); await p.click('#seSave'); await p.waitForTimeout(300);
    await p.click('#nav [data-v=logs]'); await p.waitForTimeout(200);
    await p.click('#fab'); await p.waitForTimeout(300);
    await p.fill('#sheetBody input[type=text]:not(.num) >> nth=0', '=HYPERLINK("http://evil.example","x")');
    await p.click('#eSave'); await p.waitForTimeout(400);
    await p.click('#quickExport'); await p.waitForTimeout(300);
    const [dl] = await Promise.all([p.waitForEvent('download'), p.click('#bDl')]);
    const exported = path.join(TMP, 'export.xlsx'); await dl.saveAs(exported); await p.waitForTimeout(300);
    ok(fs.statSync(exported).size > 2000, 'backup exported');
    const xmls = require('child_process').execSync(`unzip -p "${exported}" 'xl/worksheets/*.xml'`).toString();
    ok(!/<f>[^<]*HYPERLINK/i.test(xmls), 'formula-looking text was not written as a formula');

    /* 4b. PDF report of one car, with and without amounts */
    await p.click('#quickExport'); await p.waitForTimeout(300);
    if (await p.isVisible('[data-pc="0"]')) await p.click('[data-pc="0"]');
    const [pd] = await Promise.all([p.waitForEvent('download', { timeout: 60000 }), p.click('#pDl')]);
    const pdfFile = path.join(TMP, 'report.pdf'); await pd.saveAs(pdfFile); await p.waitForTimeout(300);
    const pdf = fs.readFileSync(pdfFile);
    ok(pdf.slice(0, 5).toString() === '%PDF-' && pdf.length > 8000, 'PDF report created (' + Math.round(pdf.length / 1024) + ' KB)');
    const hidden = await p.evaluate(() => JSON.stringify(reportModel(S.cars[0], true)));
    const shown = await p.evaluate(() => JSON.stringify(reportModel(S.cars[0], false)));
    ok(shown.includes('د.ك') && !hidden.includes('د.ك') && !hidden.includes('4,500') && hidden.includes('لا يتضمن المبالغ'), 'hide-amounts option removes prices and totals');
    const hidOk = await p.evaluate(async () => { const b = await makePdf(S.cars[0], true); return b.size > 8000 && (await b.slice(0, 5).text()) === '%PDF-'; });
    ok(hidOk, 'PDF without amounts created');

    /* 4c. reminders: dates, distance or time (whichever first), expiries, Arabic counts */
    const rem = await p.evaluate(() => {
      const iso = n => addDaysISO(todayISO(), n);
      const out = {};
      out.months = [addMonthsISO('2026-08-31', 6), addMonthsISO('2027-01-31', 1), addMonthsISO('2028-01-31', 1), addMonthsISO('2026-11-15', 12)].join(' ');
      out.days = [1, 2, 3, 10, 11, 30, 100].map(nDays).join('|');
      const carOf = (rows, extra = {}) => sanitizeCar({ id: 'tc', name: 'T', cardFields: extra.card || [], odoNow: extra.odoNow || null, sections: [
        { id: 'oil', name: 'الزيوت', reminderKm: 5000, reminderMonths: 6, columns: [{ id: 'desc', label: 'بيان', type: 'text' }, { id: 'odo', label: 'العداد', type: 'odometer' }, { id: 'date', label: 'التاريخ', type: 'date' }], rows },
        { id: 'ins', name: 'التأمين والترخيص', columns: [{ id: 'desc', label: 'بيان', type: 'text' }, { id: 'exp', label: 'تاريخ الانتهاء', type: 'date' }], rows: extra.ins || [] }] }, new Set());
      /* little driving: the 6 months come first */
      let c = carOf([{ id: 'a', v: { odo: 10000, date: iso(-170) } }], { odoNow: { value: 10500, date: iso(0) } });
      let x = reminders(c)[0];
      out.timeFirst = x.by === 'time' && x.time.due === addMonthsISO(iso(-170), 6) && x.time.days === Math.round(daysBetween(todayISO(), x.time.due)) && !x.late && x.due === x.time.due && !x.est;
      /* a lot of driving: the distance comes first, and its date is estimated from the driving rate */
      c = carOf([{ id: 'a', v: { odo: 10000, date: iso(-30) } }], { odoNow: { value: 14500, date: iso(0) } });
      x = reminders(c)[0];
      out.kmFirst = x.by === 'km' && x.km.left === 500 && x.est && x.due === iso(3) && !x.late;
      /* overdue by distance */
      c = carOf([{ id: 'a', v: { odo: 10000, date: iso(-30) } }], { odoNow: { value: 15200, date: iso(0) } });
      x = reminders(c)[0];
      out.late = x.late && x.by === 'km' && x.km.left === -200;
      /* expiries: insurance and registration in one section are both reminded; the older insurance is replaced
         by the newer one; an item that expired long ago is not shown; a card item named "انتهاء…" counts too */
      c = carOf([], { card: [{ id: 'f1', label: 'انتهاء رخصة القيادة', type: 'date', value: iso(40) }],
        ins: [{ id: 'i1', v: { desc: 'تأمين 2025', exp: iso(-300) } }, { id: 'i2', v: { desc: 'تأمين شامل 2026', exp: iso(65) } },
              { id: 'i3', v: { desc: 'تجديد الدفتر', exp: iso(12) } }, { id: 'i4', v: { desc: 'فحص قديم', exp: iso(-90) } }] });
      out.exp = expiries(c).map(e => e.title + ':' + e.days).join('|');
      /* overdue by an undated reading: due today at the latest, never months ahead */
      c = sanitizeCar({ id: 'tc', name: 'T', cardFields: [], sections: [
        { id: 'oil', name: 'الزيوت', reminderKm: 5000, columns: [{ id: 'odo', label: 'العداد', type: 'odometer' }, { id: 'date', label: 'التاريخ', type: 'date' }], rows: [{ id: 'a', v: { odo: 10000, date: iso(-60) } }, { id: 'b', v: { odo: 11000, date: iso(-30) } }] },
        { id: 'fuel', name: 'الوقود', columns: [{ id: 'o', label: 'العداد', type: 'odometer' }], rows: [{ id: 'f', v: { o: 16500 } }] }] }, new Set());
      x = reminders(c)[0];
      out.undated = x.late && x.due <= todayISO();
      /* dates that do not exist stay text; an item named «تاريخ التجديد» is not an expiry and is filled with today */
      out.badDate = cleanValue('date', '31.02.2027') === '31.02.2027' && cleanValue('date', '2027-02-31') === '31.02.2027' && cleanValue('date', '29.02.2028') === '2028-02-29';
      out.renew = !EXPIRY_RE.test('تاريخ التجديد') && EXPIRY_RE.test('تاريخ الانتهاء');
      /* a section saved before 1.3 (no time interval): oil gets 6 months, an explicit 0 stays 0 */
      const old = sanitizeSection({ id: 'oil', name: 'الزيوت', reminderKm: 7500, columns: [{ id: 'd', label: 'التاريخ', type: 'date' }], rows: [] }, new Set());
      const off = sanitizeSection({ id: 'oil', name: 'الزيوت', reminderKm: 7500, reminderMonths: 0, columns: [], rows: [] }, new Set());
      out.migrate = old.reminderMonths === 6 && off.reminderMonths === 0;
      return out;
    });
    ok(rem.months === '2027-02-28 2027-02-28 2028-02-29 2027-11-15', 'months are added without overflowing (' + rem.months + ')');
    ok(rem.days === 'يوم واحد|يومين|3 أيام|10 أيام|11 يوماً|30 يوماً|100 يوم', 'Arabic day counts (' + rem.days + ')');
    ok(rem.timeFirst, 'reminder by time when the months come first');
    ok(rem.kmFirst, 'reminder by distance when the kilometres come first (date estimated from the driving rate)');
    ok(rem.late, 'overdue reminder');
    ok(rem.exp === 'تجديد الدفتر:12|انتهاء رخصة القيادة:40|تأمين شامل 2026:65', 'expiry dates grouped by what they are (' + rem.exp + ')');
    ok(rem.migrate, 'older oil sections get "every 6 months"; an explicit 0 is kept');
    ok(rem.undated, 'overdue by an undated reading: due now, not months ahead');
    ok(rem.badDate && rem.renew, 'impossible dates stay text; «تجديد» items are not taken for expiry dates');

    /* 4d. phone notifications plan (Android) and the calendar file (iPhone / web) */
    const plan = await p.evaluate(() => {
      const keep = S;
      try {
        const iso = n => addDaysISO(todayISO(), n);
        S = sanitizeState({ settings: { notifyHour: 9, expiryLead: 30, odoNudge: 30 }, cars: [{ id: 'c1', name: 'F150, "الكبيرة"; تجربة', cardFields: [{ id: 'f', label: 'انتهاء التأمين', type: 'date', value: iso(45) }], odoNow: { value: 12000, date: iso(-40) },
          sections: [{ id: 'oil', name: 'الزيوت', reminderKm: 5000, reminderMonths: 6, columns: [{ id: 'odo', label: 'العداد', type: 'odometer' }, { id: 'date', label: 'التاريخ', type: 'date' }], rows: [{ id: 'r', v: { odo: 10000, date: iso(-100) } }] }] }] });
        const a = planAlerts(), b = planAlerts();
        const ids = a.map(x => x.id);
        const nine = a.every(x => new Date(x.at).getHours() === 9 && new Date(x.at).getMinutes() === 0);
        const odo = a.find(x => x.key.includes(':odo:'));
        const ics = buildIcs(S.cars);
        const lines = ics.text.split('\r\n');
        const enc = new TextEncoder();
        return {
          n: a.length, same: JSON.stringify(a) === JSON.stringify(b), future: a.every(x => x.at > Date.now()), nine,
          ids: ids.every(i => Number.isInteger(i) && i > 0 && i < 2 ** 31) && new Set(ids).size === ids.length,
          kinds: a.map(x => (x.key.includes(':odo:') ? 'odo' : x.key.split(':').pop())).sort().join(','),
          odo: odo ? Math.round(daysBetween(addDaysISO(todayISO(), -40), new Date(odo.at - new Date(odo.at).getTimezoneOffset() * 6e4).toISOString().slice(0, 10))) : -1,
          icsCount: ics.count, crlf: !/[^\r]\n/.test(ics.text), short: lines.every(l => enc.encode(l).length <= 75),
          unfold: ics.text.replace(/\r\n /g, ''), alarms: (ics.text.match(/BEGIN:VALARM/g) || []).length
        };
      } finally { S = keep; }
    });
    ok(plan.n >= 5 && plan.same && plan.future && plan.nine && plan.ids, `notifications: ${plan.n} planned, the same each time, at 9:00, unique ids`);
    ok(/day/.test(plan.kinds) && /lead/.test(plan.kinds) && /week/.test(plan.kinds) && /due/.test(plan.kinds) && /odo/.test(plan.kinds), 'notifications for the expiry (30 days, a week, the day), the service and the odometer (' + plan.kinds + ')');
    ok(plan.odo === 60, 'odometer reminder every 30 days after the last reading (' + plan.odo + ' days)');
    ok(plan.icsCount === 2 && plan.crlf && plan.short && plan.alarms === 5, 'calendar file: 2 events, 5 alarms, CRLF, lines folded at 75 bytes');
    ok(plan.unfold.includes('SUMMARY:انتهاء التأمين — F150\\, "الكبيرة"\\; تجربة') && /TRIGGER:-P29DT15H/.test(plan.unfold) && /TRIGGER:PT9H/.test(plan.unfold) && /DTSTART;VALUE=DATE:\d{8}/.test(plan.unfold),
      'calendar text escaped and alarms at the reminder hour');

    /* 4e. invoice photos: added from the gallery, made smaller, hidden camera data removed */
    const mkImg = (w, h, type, orient, secret) => p.evaluate(async ({ w, h, type, orient, secret }) => {
      const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
      const g = cv.getContext('2d'); g.fillStyle = '#e9a23b'; g.fillRect(0, 0, w, h); g.fillStyle = '#12262e'; g.fillRect(0, 0, w / 2, h / 3);
      let u8 = new Uint8Array(await (await new Promise(r => cv.toBlob(r, type, 0.9))).arrayBuffer());
      if (orient || secret) {   /* an EXIF block (as cameras write) with the orientation and a text that must not survive */
        const enc = new TextEncoder(), desc = enc.encode(secret + '\0'), n = 2;
        const ifd = 2 + n * 12 + 4, tiff = new Uint8Array(8 + ifd + desc.length), dv = new DataView(tiff.buffer);
        tiff.set([0x4D, 0x4D, 0, 0x2A, 0, 0, 0, 8]); dv.setUint16(8, n);
        dv.setUint16(10, 0x010E); dv.setUint16(12, 2); dv.setUint32(14, desc.length); dv.setUint32(18, 8 + ifd); tiff.set(desc, 8 + ifd);
        dv.setUint16(22, 0x0112); dv.setUint16(24, 3); dv.setUint32(26, 1); dv.setUint16(30, orient || 1); dv.setUint32(34, 0);
        const body = new Uint8Array(6 + tiff.length); body.set(enc.encode('Exif\0\0')); body.set(tiff, 6);
        const seg = new Uint8Array(4 + body.length); seg.set([0xFF, 0xE1, (body.length + 2) >> 8, (body.length + 2) & 255]); seg.set(body, 4);
        const out = new Uint8Array(u8.length + seg.length); out.set(u8.subarray(0, 2)); out.set(seg, 2); out.set(u8.subarray(2), 2 + seg.length); u8 = out;
      }
      let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
      return s;
    }, { w, h, type, orient, secret }).then(s => Buffer.from(s, 'latin1'));
    await p.click('#nav [data-v=logs]'); await p.waitForTimeout(200);
    await p.click('.seg [data-s] >> nth=0'); await p.waitForTimeout(200);
    await p.click('#fab'); await p.waitForTimeout(400);
    await p.fill('#sheetBody input[type=text]:not(.num) >> nth=0', 'فاتورة الورشة');
    const [ph] = await Promise.all([p.waitForEvent('filechooser'), p.click('#phPick')]);
    await ph.setFiles([
      { name: 'camera.jpg', mimeType: 'image/jpeg', buffer: await mkImg(2400, 1800, 'image/jpeg', 6, 'GPS-SECRET-29.37N') },
      { name: 'receipt.png', mimeType: 'image/png', buffer: await mkImg(500, 900, 'image/png') },
      { name: 'fake.jpg', mimeType: 'image/jpeg', buffer: Buffer.from('not an image at all') }
    ]);
    await p.waitForFunction(() => document.querySelectorAll('#phList .ph img').length === 2 && !document.querySelector('#phList .busy'), null, { timeout: 20000 }).catch(() => {});
    const tiles = await p.$$eval('#phList .ph img', im => im.map(i => i.naturalWidth + 'x' + i.naturalHeight));
    ok(tiles.join(' ') === '1200x1600 500x900', 'two photos added, the broken one refused; camera orientation applied and size reduced (' + tiles.join(' ') + ')');
    /* the viewer: two taps delete a photo from the entry */
    await p.click('#phList [data-pi="1"]'); await p.waitForTimeout(300);
    ok(await p.isVisible('#viewer.on') && (await p.textContent('#vCount')).includes('2'), 'photo viewer opens');
    await p.click('#vDel'); await p.click('#vDel'); await p.waitForTimeout(300);
    ok(!(await p.isVisible('#viewer.on')) || (await p.$$('#phList .ph img')).length === 1, 'photo removed from the entry after confirming');
    if (await p.isVisible('#viewer.on')) await p.click('#vClose');
    await p.click('#eSave'); await p.waitForTimeout(600);
    st = await state();
    const prow = st.cars.flatMap(c => c.sections).flatMap(s => s.rows).find(r => r.p);
    const pinfo = prow ? await p.evaluate(async ids => {
      const b = await Photos.blob(ids[0]); const u8 = new Uint8Array(await b.arrayBuffer());
      const txt = new TextDecoder('latin1').decode(u8);
      return { type: b.type, size: b.size, dim: jpegSize(u8), secret: txt.includes('GPS-SECRET'), exif: txt.includes('Exif') };
    }, prow.p) : null;
    ok(prow && prow.p.length === 1 && pinfo && pinfo.type === 'image/jpeg' && pinfo.dim.w === 1200 && pinfo.dim.h === 1600, 'entry keeps one photo, stored as a 1200×1600 JPEG');
    ok(pinfo && !pinfo.secret && !pinfo.exif, 'hidden camera data (location…) removed from the photo');
    ok(await p.isVisible('.entry .clip'), 'records list marks entries with photos');
    /* photos go into the PDF report when asked */
    const pdfImgs = await p.evaluate(async () => {
      const c = S.cars.find(x => x.sections.some(s => s.rows.some(r => r.p)));
      const count = async b => (new TextDecoder('latin1').decode(new Uint8Array(await b.arrayBuffer())).match(/\/Subtype\s*\/Image/g) || []).length;
      return [await count(await makePdf(c, false, true)), await count(await makePdf(c, false, false))];
    });
    ok(pdfImgs[0] === 1 && pdfImgs[1] === 0, 'PDF report includes the invoice photo only when chosen (' + pdfImgs + ')');

    /* 5. a crafted backup: nothing runs, everything is cleaned */
    const P = 'window.__pwned=';
    const evil = {
      app: 'sayyarati', v: 1, exported: `<img src=x onerror=${P}9>`,
      settings: { decimals: 50, currency: `<img src=x onerror=${P}8>`, theme: 'x' },
      cars: [{
        id: `x" autofocus onfocus="${P}1`, name: `<img src=x onerror="${P}3">`, color: `red;" onmouseover="${P}2`,
        odoNow: { value: `"><img src=x onerror=${P}10>`, date: '<b>' },
        cardFields: [{ id: 'toString', label: `<img src=x onerror=${P}11>`, type: `"><img src=x onerror=${P}12>`, value: `<img src=x onerror=${P}13>` },
                     { id: 'd', label: 'تاريخ', type: 'date', value: `<img src=x onerror=${P}14>` }],
        sections: [{
          id: `"><svg onload=${P}4>`, name: `زيوت<script>${P}15</script>`, icon: `"><img src=x onerror=${P}5>`,
          note: `<img src=x onerror=${P}16>`, reminderKm: `"><img src=x onerror=${P}6>`,
          columns: [{ id: 'odo', label: 'العداد', type: 'odometer' }, { id: 'dt', label: 'التاريخ', type: 'date' },
                    { id: 'constructor', label: 'نص 😀', type: `"><img src=x onerror=${P}17>` }, { id: 'amt', label: 'المبلغ', type: 'money' }],
          rows: [{ id: `"><img src=x onerror=${P}18>`, t: 'x', v: { odo: 1000, dt: `"><img src=x onerror=${P}7>`, constructor: '😀'.repeat(20000), amt: '1e400' } },
                 { id: 'r2', v: { odo: 2000, dt: '2026-01-05', constructor: '@SUM(A1)', amt: 5 } }]
        }]
      }, 'junk', null, { id: 5 }]
    };
    let json = JSON.stringify(evil).replace('"amt":"1e400"', '"amt":"1e400","__proto__":{"polluted":1}').replace('"theme":"x"', '"theme":"x","__proto__":{"polluted":1}');
    await pick(await backupFile('evil.xlsx', json));
    ok(await p.isVisible('#rMerge'), 'crafted backup offers a restore only after it is cleaned');
    await p.click('#rMerge'); await p.waitForTimeout(1200);
    st = await state();
    const ev = st.cars.find(c => c.name.includes('img'));
    for (let i = 0; i < st.cars.length; i++) {
      await p.click('#carPill'); await p.waitForTimeout(200); await p.hover('#sheetBody .list-item >> nth=0');
      await p.click(`#sheetBody [data-car-i="${i}"]`); await p.waitForTimeout(300);
      for (const v of ['home', 'logs', 'stats', 'settings']) { await p.click(`#nav [data-v=${v}]`); await p.waitForTimeout(200); }
      await p.click('#v-settings [data-sec-i="0"]').catch(() => {}); await p.waitForTimeout(200); await closeSheet();
      await p.click('#nav [data-v=logs]'); await p.waitForTimeout(150);
      if (await p.isVisible('.entry')) { await p.click('.entry >> nth=0'); await p.waitForTimeout(200); await closeSheet(); }
      await p.click('#nav [data-v=home]'); await p.waitForTimeout(150); await p.click('#hCard'); await p.waitForTimeout(200); await closeSheet();
    }
    ok((await pwned()) === undefined, 'no injected script ran on any screen');
    ok(st.cars.length === 3 && !!ev, 'junk entries dropped (' + st.cars.length + ' cars)');
    ok(/^[A-Za-z0-9_-]+$/.test(ev.id) && /^#[0-9A-Fa-f]{6}$/.test(ev.color), 'id and colour replaced');
    ok(st.settings.decimals === 3 && st.settings.theme === 'auto' && st.settings.currency.length <= 12, 'settings cleaned');
    ok(await p.evaluate(() => ({}).polluted === undefined), 'no prototype pollution');
    const es = ev.sections[0];
    ok(es.icon === 'doc' && es.reminderKm === 0 && es.columns.every(c => /^[A-Za-z0-9_-]+$/.test(c.id) && c.id !== 'constructor'), 'section cleaned');

    const evilPdf = await p.evaluate(async () => { const c = S.cars.find(x => x.name.includes('img')); const b = await makePdf(c, false); return b.size; });
    ok(evilPdf > 3000, 'PDF of the crafted car still builds (' + Math.round(evilPdf / 1024) + ' KB)');

    /* 6. emoji survive a backup round trip (the backup is cut into cells) */
    await p.click('#quickExport'); await p.waitForTimeout(300);
    await p.click('[data-w=all]').catch(() => {});
    const [dl2] = await Promise.all([p.waitForEvent('download'), p.click('#bDl')]);
    const exported2 = path.join(TMP, 'export2.xlsx'); await dl2.saveAs(exported2); await p.waitForTimeout(300);
    await pick(exported2);
    await p.click('#rReplace'); await p.waitForTimeout(300); await p.click('#cYes'); await p.waitForTimeout(1000);
    st = await state();
    const ev2 = st.cars.find(c => c.name.includes('img'));
    const longText = ev2 && Object.values(ev2.sections[0].rows[0].v).find(v => typeof v === 'string' && v.includes('😀'));
    ok(!!longText && !longText.includes('�') && longText.length === 2000, 'emoji intact after a backup round trip');

    /* 7. broken / fake / oversized files: a clear message, nothing imported */
    fs.writeFileSync(path.join(TMP, 'fake.xlsx'), 'not a spreadsheet');
    const many = { app: 'sayyarati', cars: [{ id: 'big', name: 'big', cardFields: [], sections: [] }] };
    for (let k = 0; k < 4; k++) many.cars[0].sections.push({ id: 's' + k, name: 's' + k, columns: [{ id: 'o', label: 'العداد', type: 'odometer' }], rows: Array.from({ length: 5000 }, (_, i) => ({ id: 'r' + i, v: { o: i } })) });
    const cases = [
      ['broken.xlsx', await backupFile('broken.xlsx', '{"app":"sayyarati","cars":['), 'تالفة'],
      ['no-cars.xlsx', await backupFile('no-cars.xlsx', '{"app":"sayyarati","cars":"no"}'), 'تالفة'],
      ['fake.xlsx', path.join(TMP, 'fake.xlsx'), 'ليس بصيغة'],
      ['too-many.xlsx', await backupFile('too-many.xlsx', JSON.stringify(many)), 'أكثر مما يتسع']
    ];
    for (const [name, file, want] of cases) {
      const before = JSON.stringify(await state());
      await pick(file);
      const t = await p.textContent('#sheetBody').catch(() => '');
      ok((await p.textContent('#sheetTitle')) === 'تعذّر الاستيراد' && t.includes(want) && JSON.stringify(await state()) === before, name + ': message shown, data unchanged');
      if (await p.isVisible('#ieClose')) await p.click('#ieClose');
      await p.waitForTimeout(200);
    }

    /* 8. a stored state that breaks a screen never locks the app */
    await p.evaluate(() => { const s = JSON.parse(localStorage.getItem('sayyarati.state.v1')); s.cars[0].sections[0].rows = Array.from({ length: 100000 }, (_, i) => ({ id: 'r' + i, v: { odo: i } })); localStorage.setItem('sayyarati.state.v1', JSON.stringify(s)); });
    await p.reload(); await p.waitForTimeout(1500);
    ok(await p.isVisible('#nav') && (await state()).cars.length >= 1, 'oversized stored data is trimmed and the app opens');

    /* 9. inside a frame of another page the app does not run */
    const fp = await ctx.newPage();
    await fp.goto('about:blank');
    await fp.setContent(`<iframe src="${BASE}index.html"></iframe>`);
    await fp.waitForTimeout(1500);
    const fr = fp.frames().find(f => f.url().includes('index.html'));
    ok(!fr || await fr.evaluate(() => !document.querySelector('#nav')).catch(() => true), 'refuses to run inside a frame');

    /* 10. a new phone: the backup brings the invoice photos back */
    const ctx2 = await browser.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true, locale: 'ar' });
    const q = await ctx2.newPage();
    q.on('pageerror', e => errors.push('page2: ' + e.message));
    q.on('console', m => { const t = m.text(); if (/Content Security Policy|Refused to/.test(t)) csp.push(t); else if (m.type() === 'error') errors.push(t); });
    await q.goto(BASE + 'index.html'); await q.waitForTimeout(800);
    const pick2 = async file => { const [fc] = await Promise.all([q.waitForEvent('filechooser'), q.evaluate(() => document.querySelector('#fileIn').click())]); await fc.setFiles(file); await q.waitForTimeout(2000); };
    await pick2(exported2);
    ok((await q.textContent('#sheetBody')).includes('صور الفواتير: 1'), 'restore summary counts the invoice photos');
    await q.click('#rReplace'); await q.waitForTimeout(300); await q.click('#cYes'); await q.waitForTimeout(2500);
    const back = await q.evaluate(async () => {
      const ids = [...usedPhotoIds()]; const stored = (await Photos.list()).map(x => x.id);
      const b = ids.length && await Photos.blob(ids[0]);
      return { ids, stored, jpeg: b ? !!jpegSize(new Uint8Array(await b.arrayBuffer())) : false };
    });
    ok(back.ids.length === 1 && back.stored.includes(back.ids[0]) && back.jpeg, 'photos restored on a new device with their entries');

    /* 11. photos in a crafted backup: only real JPEGs with valid ids are taken, references to missing photos dropped */
    const crafted = await q.evaluate(async () => {
      const ExcelJS = await loadExcel();
      const cv = document.createElement('canvas'); cv.width = 40; cv.height = 30; cv.getContext('2d').fillRect(0, 0, 20, 20);
      const bytes = async t => new Uint8Array(await (await new Promise(r => cv.toBlob(r, t, 0.9))).arrayBuffer());
      const jpg = await bytes('image/jpeg'), png = await bytes('image/png');
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('x').getCell('A1').value = 'visible';
      const m = wb.addWorksheet('_sayyarati_photos', { state: 'veryHidden' });
      m.getCell(1, 1).value = 'sayyarati-photos-v1';
      const big = document.createElement('canvas'); big.width = 4200; big.height = 8;
      const huge = new Uint8Array(await (await new Promise(r => big.toBlob(r, 'image/jpeg', 0.5))).arrayBuffer());
      [['okphoto', jpg], ['constructor', jpg], ['pngphoto', png], ['okphoto', jpg], ['bad id!', jpg], ['hugephoto', huge]].forEach(([id, u8], i) => {
        const im = wb.addImage({ buffer: u8.buffer, extension: id === 'pngphoto' ? 'png' : 'jpeg' });
        m.getCell(i + 2, 1).value = id; m.addImage(im, { tl: { col: 1, row: i + 1 }, ext: { width: 4, height: 4 } });
      });
      const h = wb.addWorksheet('_sayyarati', { state: 'veryHidden' });
      h.getCell(1, 1).value = 'sayyarati-backup-v1';
      h.getCell(2, 1).value = JSON.stringify({ app: 'sayyarati', cars: [{ id: 'cc', name: 'crafted', cardFields: [], sections: [{ id: 's', name: 's', columns: [{ id: 'd', label: 'بيان', type: 'text' }],
        rows: [{ id: 'r1', v: { d: 'x' }, p: ['okphoto', 'constructor', 'pngphoto', 'missing1', '<img src=x>', 'okphoto', 'hugephoto'] }] }] }] });
      const buf = await wb.xlsx.writeBuffer();
      await importBuffer(buf, 'crafted.xlsx');
      return true;
    });
    await q.waitForTimeout(500);
    if (crafted && await q.isVisible('#rMerge')) { await q.click('#rMerge'); await q.waitForTimeout(2000); }
    const cr = await q.evaluate(async () => { const c = S.cars.find(x => x.id === 'cc'); return { p: c && c.sections[0].rows[0].p, stored: (await Photos.list()).map(x => x.id).sort() }; });
    ok(JSON.stringify(cr.p) === '["okphoto"]' && cr.stored.includes('okphoto') && !cr.stored.includes('pngphoto') && !cr.stored.includes('constructor') && !cr.stored.includes('hugephoto'), 'crafted photo backup cleaned (' + JSON.stringify(cr) + ')');

    /* 12. photos no entry uses are cleaned up (only once they are hours old) */
    const gc = await q.evaluate(async () => {
      const blob = await Photos.blob('okphoto');
      await Photos.put(blob, 'orphanold'); await Photos.put(blob, 'orphannew');
      const db = await new Promise(r => { const o = indexedDB.open('sayyarati-photos', 1); o.onsuccess = () => r(o.result); });
      await new Promise(r => { const t = db.transaction('photos', 'readwrite'); const st = t.objectStore('photos'); st.get('orphanold').onsuccess = e => { const v = e.target.result; v.t = Date.now() - 864e5; st.put(v); }; t.oncomplete = r; });
      localStorage.removeItem('sayyarati.photoGC');
      await gcPhotos();
      return (await Photos.list()).map(x => x.id).sort().join(',');
    });
    ok(gc.includes('okphoto') && gc.includes('orphannew') && !gc.includes('orphanold'), 'unused photos removed once old (' + gc + ')');

    /* 12b. the cleanup and a restore at the same moment: the restore keeps the photos it brings */
    const race = await q.evaluate(async () => {
      const blob = await Photos.blob('okphoto');
      await Photos.put(blob, 'racephoto');
      const db = await new Promise(r => { const o = indexedDB.open('sayyarati-photos', 1); o.onsuccess = () => r(o.result); });
      await new Promise(r => { const t = db.transaction('photos', 'readwrite'); const st = t.objectStore('photos'); st.get('racephoto').onsuccess = e => { const v = e.target.result; v.t = Date.now() - 3 * 864e5; st.put(v); }; t.oncomplete = r; });
      localStorage.removeItem('sayyarati.photoGC');
      const restored = sanitizeCars([{ id: 'rc', name: 'restored', cardFields: [], sections: [{ id: 's', name: 's', columns: [{ id: 'd', label: 'بيان', type: 'text' }], rows: [{ id: 'r', v: { d: 'x' }, p: ['racephoto'] }] }] }]);
      await Promise.all([gcPhotos(), finishRestore(restored, S.settings, new Map([['racephoto', new Uint8Array(await blob.arrayBuffer())]]), restored)]);
      return { refs: (S.cars[0].sections[0].rows[0].p || []).join(), kept: !!(await Photos.blob('racephoto')) };
    });
    ok(race.refs === 'racephoto' && race.kept, 'cleanup running during a restore leaves its photos alone (' + JSON.stringify(race) + ')');

    /* 13. iPhone "add to calendar": the service worker serves the calendar file the page made */
    await q.waitForFunction(() => !!navigator.serviceWorker.controller, null, { timeout: 15000 }).catch(() => {});
    if (!(await q.evaluate(() => !!navigator.serviceWorker.controller))) { await q.reload(); await q.waitForTimeout(1500); }
    const icsRes = await q.evaluate(async () => {
      const keep = S;
      S = sanitizeState({ cars: [{ id: 'c1', name: 'T', cardFields: [{ id: 'f', label: 'انتهاء التأمين', type: 'date', value: addDaysISO(todayISO(), 20) }], sections: [] }] });
      const { text } = buildIcs(S.cars); S = keep;
      const cache = await caches.open(ICS_CACHE);
      await cache.put(new URL(ICS_PATH, location.href).href, new Response(new Blob([text]), { headers: { 'Content-Type': 'text/calendar; charset=utf-8' } }));
      const r = await fetch(ICS_PATH);
      return { controlled: !!navigator.serviceWorker.controller, type: r.headers.get('content-type'), body: (await r.text()).slice(0, 15) };
    });
    ok(icsRes.controlled && /text\/calendar/.test(icsRes.type) && icsRes.body === 'BEGIN:VCALENDAR', 'calendar file served by the service worker (' + JSON.stringify(icsRes) + ')');
    await ctx2.close();
  } catch (e) {
    results.push('ERROR ' + e.message);
  }
  ok(csp.length === 0, 'no security-policy violations ' + csp.slice(0, 2).join(' | '));
  ok(errors.length === 0, 'no JS errors ' + errors.slice(0, 3).join(' | '));
  console.log(results.join('\n'));
  await browser.close();
  process.exit(results.some(r => !r.startsWith('PASS')) ? 1 : 0);
})();
