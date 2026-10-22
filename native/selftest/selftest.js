/* Android emulator self-test (CI only — never part of the release APK).
   Drives the real app inside the emulator and writes PASS/FAIL lines to files/selftest-result.txt.
   native/selftest/android-smoke.sh watches for "WAIT …" lines and answers them with adb
   (taps the file in the system picker, sends "Open with" / "Share" intents, presses Back). */
(async () => {
  const log = [];
  const ok = (cond, msg) => log.push((cond ? 'PASS ' : 'FAIL ') + msg);
  const info = msg => log.push('INFO ' + msg);
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const $ = s => document.querySelector(s);
  const P = (window.Capacitor && window.Capacitor.Plugins) || {};
  const FS = P.Filesystem;
  const save = () => FS && FS.writeFile({ path: 'selftest-result.txt', data: log.join('\n') + '\n', directory: 'DATA', encoding: 'utf8' }).catch(() => {});
  const mark = async (m, ms) => { log.push(m); await save(); if (ms) await wait(ms); };
  const until = async (fn, secs) => { for (let i = 0; i < secs * 4; i++) { try { if (await fn()) return true; } catch (e) { /* not yet */ } await wait(250); } return false; };
  const state = () => { try { return JSON.parse(localStorage.getItem('sayyarati.state.v1')) || { cars: [] }; } catch (e) { return { cars: [] }; } };
  const rows = c => c.sections.reduce((a, s) => a + s.rows.length, 0);
  const allRows = () => state().cars.reduce((a, c) => a + rows(c), 0);
  const sheetText = () => ($('#sheet.on') ? ($('#sheetTitle').textContent + ' | ' + $('#sheetBody').innerText).replace(/\s+/g, ' ').slice(0, 200) : '');
  const b64 = buf => { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  try {
    await until(() => $('.view.on'), 20);
    await wait(1500);
    /* later launches (restart checks) only report what they see */
    let earlier = '';
    try { earlier = (await FS.readFile({ path: 'selftest-result.txt', directory: 'DATA', encoding: 'utf8' })).data || ''; } catch (e) { earlier = ''; }
    if (/^(DONE|ERROR)/m.test(earlier)) {
      await until(() => $('#sheet.on'), 12);
      await FS.writeFile({ path: 'selftest-relaunch.txt', directory: 'DATA', encoding: 'utf8',
        data: 'cars ' + state().cars.length + '\nrows ' + allRows() + '\nsheet ' + sheetText() + '\n' }).catch(() => {});
      return;
    }
    info((navigator.userAgent.match(/Chrome\/[\d.]+/) || [''])[0]);
    ok(window.Capacitor.isNativePlatform(), 'running as the native app');
    ok(!!P.XlsxImport && !!P.Filesystem && !!P.Share && !!P.App, 'XlsxImport, Filesystem, Share and App plugins available');
    ok(!!document.querySelector('meta[http-equiv="Content-Security-Policy"]'), 'security policy (CSP) in the page');
    ok(getComputedStyle(document.body).fontFamily.includes('Readex'), 'app font applied');
    let blocked = false;
    try { await fetch('https://example.com/selftest', { mode: 'no-cors', cache: 'no-store' }); } catch (e) { blocked = true; }
    ok(blocked, 'requests to the internet are blocked');
    const a = document.createElement('a'); a.href = 'https://example.com/selftest-nav'; document.body.appendChild(a); a.click(); a.remove();
    await wait(2500);
    ok(location.hostname === 'localhost' && document.visibilityState === 'visible', 'links to other sites are blocked (the app stays on its own pages)');
    ok(!!$('#wImport'), 'welcome screen shown');

    /* files the shell script will hand to the app with intents */
    const fixture = await (await fetch('fixture-f150-layout.xlsx')).arrayBuffer();
    /* (debug builds accept the app's own FileProvider addresses; release builds refuse them) */
    for (const n of ['test-open', 'test-cold']) await FS.writeFile({ path: 'exports/' + n + '.xlsx', data: b64(fixture), directory: 'CACHE', recursive: true });

    /* 1. the system file picker (what the user taps: Import → Drive/Downloads → file) */
    await mark('WAIT picker');
    $('#wImport').click();
    ok(await until(() => $('#giOk'), 120), 'file chosen in the system picker is read and the import summary is shown');
    info('picker sheet: ' + sheetText());
    await mark('STEP picker-summary', 2500);
    if ($('#giOk')) $('#giOk').click();
    await wait(1500);
    let st = state();
    ok(st.cars.length === 1 && rows(st.cars[0]) === 10, `car imported from the picker: ${st.cars.length} car(s), ${st.cars[0] ? rows(st.cars[0]) : 0} rows (expected 1, 10)`);
    const c0 = st.cars[0] || { sections: [], cardFields: [] };
    const oil = c0.sections.find(s => s.id === 'oil');
    ok(!!oil && oil.rows.length === 8 && oil.note === 'Test Oil 5w-30', 'oil sheet: 8 rows and its note');
    ok(c0.name === 'سيارة تجربة', 'car name read from the card sheet (' + c0.name + ')');
    ok(document.body.innerText.includes('الزيوت القادم'), 'home shows the next oil change');
    ok(!document.querySelector('.car-sub') || !document.querySelector('.car-sub').innerHTML.includes('<'), 'card values shown as text');

    /* 2. "Open with" from another app (VIEW content://) */
    await mark('WAIT view-intent');
    ok(await until(() => $('#giOk'), 90), 'file opened with the app (VIEW) shows the import summary');
    await mark('STEP view-summary', 2500);
    if ($('#giOk')) $('#giOk').click();
    await wait(1500);
    ok(state().cars.length === 2, 'second car imported from "Open with" (' + state().cars.length + ')');

    /* 3. export → share sheet */
    $('#quickExport').click();
    await until(() => $('#bShare'), 10);
    $('#bShare').click();
    await mark('WAIT share-sheet');
    let exported = null;
    await until(async () => {
      const r = await FS.readdir({ path: 'exports', directory: 'CACHE' });
      const f = (r.files || []).map(x => (typeof x === 'string' ? x : x.name)).filter(n => /^سيارتي-.*\.xlsx$/.test(n));
      if (f.length) exported = f[f.length - 1];
      return !!exported;
    }, 60);
    ok(!!exported, 'backup written for sharing: ' + exported);
    await mark('STEP after-share', 4000);
    let backup = '';
    if (exported) backup = (await FS.readFile({ path: 'exports/' + exported, directory: 'CACHE' })).data || '';
    ok(backup.startsWith('UEsDB'), 'backup is an .xlsx file (' + Math.round(backup.length * 0.75 / 1024) + ' KB)');
    if (backup) await FS.writeFile({ path: 'exports/test-backup.xlsx', data: backup, directory: 'CACHE', recursive: true });
    if ($('#sheet.on')) $('#sheetClose').click();
    await wait(800);

    /* 4. "Share" the backup to the app (SEND) → restore */
    const before = allRows();
    await mark('WAIT send-intent');
    ok(await until(() => $('#rMerge'), 90), 'backup shared to the app (SEND) offers a restore');
    await mark('STEP restore', 2500);
    if ($('#rMerge')) $('#rMerge').click();
    await wait(1500);
    ok(state().cars.length === 2 && allRows() === before, `restore (merge) keeps 2 cars and ${before} rows (${state().cars.length}, ${allRows()})`);

    /* 5. a new entry is saved to the app's own file too */
    $('#fab').click(); await until(() => $('#eSave'), 10);
    const odo = document.querySelector('#sheetBody input.num[inputmode="numeric"]');
    if (odo) { odo.value = '101000'; odo.dispatchEvent(new Event('input')); }
    const txt = document.querySelector('#sheetBody input[type="text"]:not(.num)');
    if (txt) { txt.value = 'تجربة'; txt.dispatchEvent(new Event('input')); }
    $('#eSave').click(); await wait(1500);
    ok(allRows() === before + 1, 'new entry saved');
    let fileRows = -1;
    try { const f = JSON.parse((await FS.readFile({ path: 'sayyarati-state.json', directory: 'DATA', encoding: 'utf8' })).data); fileRows = f.cars.reduce((a, c) => a + rows(c), 0); } catch (e) { info('state file: ' + e.message); }
    ok(fileRows === before + 1, 'state also saved in the app file (' + fileRows + ')');
    await mark('STEP home', 2500);
    log.push('DONE');
  } catch (e) {
    log.push('ERROR ' + (e && (e.stack || e.message) || e));
  }
  await save();
})();
