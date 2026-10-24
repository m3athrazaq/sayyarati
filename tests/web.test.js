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
  } catch (e) {
    results.push('ERROR ' + e.message);
  }
  ok(csp.length === 0, 'no security-policy violations ' + csp.slice(0, 2).join(' | '));
  ok(errors.length === 0, 'no JS errors ' + errors.slice(0, 3).join(' | '));
  console.log(results.join('\n'));
  await browser.close();
  process.exit(results.some(r => !r.startsWith('PASS')) ? 1 : 0);
})();
