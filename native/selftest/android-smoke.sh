#!/usr/bin/env bash
# Runs inside the Android emulator job: installs the test APK (with selftest.js), answers the
# self-test's "WAIT …" steps with adb (system file picker taps, Open with / Share intents, Back),
# and collects screenshots + results in smoke/ (progress.txt shows where time went).
set -u
PKG=com.m3athrazaq.sayyarati
APK=native/android/app/build/outputs/apk/debug/app-debug.apk
XLSX=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
OUT=smoke
mkdir -p "$OUT"
T0=$(date +%s)

log() { echo "[$(( $(date +%s) - T0 ))s] $*" | tee -a "$OUT/progress.txt"; }
note() { log "$*"; echo "::notice::smoke: $*"; }
A() { timeout "${TO:-30}" adb "$@"; }                     # every adb call has a time limit
res() { TO=10 A shell run-as "$PKG" cat files/selftest-result.txt 2>/dev/null | tr -d '\r'; }
shot() { TO=15 A exec-out screencap -p > "$OUT/$1.png" 2>/dev/null || true; }
waitfor() {  # marker, seconds
  local end=$(( $(date +%s) + $2 ))
  while [ "$(date +%s)" -lt "$end" ]; do
    local r; r=$(res)
    if echo "$r" | grep -qE "^($1|DONE|ERROR)"; then
      echo "$r" | grep -qE "^$1" && { log "reached: $1"; return 0; }
      log "stopped before $1"; return 1
    fi
    sleep 1
  done
  log "timed out waiting for $1"; return 1
}
tap_text() {  # regex, tries
  for _ in $(seq 1 "$2"); do
    TO=20 A shell uiautomator dump /sdcard/ui.xml >/dev/null 2>&1
    TO=10 A shell cat /sdcard/ui.xml > "$OUT/ui.xml" 2>/dev/null
    XY=$(python3 native/selftest/find_node.py "$OUT/ui.xml" "$1")
    if [ -n "$XY" ]; then log "tap '$1' at $XY"; TO=10 A shell input tap $XY; return 0; fi
    sleep 1
  done
  log "not found on screen: $1"; return 1
}
top_activity() { TO=15 A shell dumpsys activity activities 2>/dev/null | grep -m1 -E "topResumedActivity|mResumedActivity" | tr -d '\r'; }

note "emulator ready"
TO=10 A shell settings put global window_animation_scale 0
TO=10 A shell settings put global transition_animation_scale 0
TO=10 A shell settings put global animator_duration_scale 0
TO=180 A install -r "$APK" | tail -1 | tee -a "$OUT/progress.txt"
TO=10 A shell getprop ro.build.version.release | sed 's/^/Android /' | tee "$OUT/device.txt"
TO=20 A shell dumpsys package com.google.android.webview | grep -m1 versionName | sed 's/^ */WebView /' | tee -a "$OUT/device.txt"
TO=20 A shell dumpsys package "$PKG" | sed -n '/requested permissions:/,/install permissions:/p' > "$OUT/permissions.txt"

# the file the user picks in the system picker
TO=30 A push native/selftest/fixture-f150-layout.xlsx /sdcard/Download/sayyarati-test.xlsx >/dev/null
TO=60 A shell content call --method scan_volume --uri content://media --arg external_primary >/dev/null 2>&1 || true
TO=20 A shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d file:///sdcard/Download/sayyarati-test.xlsx >/dev/null 2>&1 || true

TO=10 A logcat -c
TO=60 A shell am start -W -n "$PKG/.MainActivity" | tr -d '\r' >> "$OUT/progress.txt"
sleep 6; shot 00-launch
log "top: $(top_activity)"

if waitfor "WAIT picker" 90; then
  sleep 5; shot 01-picker
  log "top: $(top_activity)"
  if ! tap_text "sayyarati-test" 6; then
    tap_text "Show roots" 4 || true
    sleep 2; shot 01b-roots
    tap_text "^Downloads$" 5 || true
    sleep 3; shot 01c-downloads
    if ! tap_text "sayyarati-test" 8; then
      log "file not found in the picker: closing it"
      TO=10 A shell input keyevent KEYCODE_BACK; sleep 1; TO=10 A shell input keyevent KEYCODE_BACK
    fi
  fi
else
  note "the app did not start the test (no WAIT picker) — top: $(top_activity)"
fi
waitfor "STEP picker-summary" 150 && shot 02-picker-summary && note "picker import OK"

if waitfor "WAIT view-intent" 60; then
  TO=60 A shell am start -a android.intent.action.VIEW -t "$XLSX" \
    -d "content://$PKG.fileprovider/exports/test-open.xlsx" -n "$PKG/.MainActivity" | tr -d '\r' >> "$OUT/progress.txt"
fi
waitfor "STEP view-summary" 90 && shot 03-open-with && note "open-with OK"

if waitfor "WAIT share-sheet" 60; then
  sleep 5; shot 04-share-sheet
  log "top: $(top_activity)"
  TO=10 A shell input keyevent KEYCODE_BACK
fi
waitfor "STEP after-share" 90 && shot 05-after-share

if waitfor "WAIT send-intent" 60; then
  TO=60 A shell am start -a android.intent.action.SEND -t "$XLSX" \
    --eu android.intent.extra.STREAM "content://$PKG.fileprovider/exports/test-backup.xlsx" -n "$PKG/.MainActivity" | tr -d '\r' >> "$OUT/progress.txt"
fi
waitfor "STEP restore" 90 && shot 06-restore && note "share/restore OK"
waitfor "STEP home" 60 && shot 07-home
waitfor "DONE" 60
sleep 2
res | tee "$OUT/selftest-result.txt"
TO=20 A exec-out run-as "$PKG" cat files/sayyarati-state.json > "$OUT/state.json" 2>/dev/null
TO=30 A logcat -d > "$OUT/logcat-full.txt" 2>/dev/null || true
note "self-test finished: $(grep -c '^PASS' "$OUT/selftest-result.txt") passed, $(grep -cE '^(FAIL|ERROR)' "$OUT/selftest-result.txt") failed"

# data must survive a full restart of the app
TO=20 A shell am force-stop "$PKG"
TO=60 A shell am start -W -n "$PKG/.MainActivity" >/dev/null
sleep 16; shot 08-relaunch
TO=10 A shell run-as "$PKG" cat files/selftest-relaunch.txt 2>/dev/null | tr -d '\r' | tee "$OUT/relaunch-1.txt"

# an Excel file opened with the app while it is closed (cold start from Drive / Files / WhatsApp)
TO=20 A shell am force-stop "$PKG"
TO=10 A shell run-as "$PKG" rm -f files/selftest-relaunch.txt
TO=60 A shell am start -W -a android.intent.action.VIEW -t "$XLSX" \
  -d "content://$PKG.fileprovider/exports/test-cold.xlsx" -n "$PKG/.MainActivity" >/dev/null
sleep 16; shot 09-cold-open
TO=10 A shell run-as "$PKG" cat files/selftest-relaunch.txt 2>/dev/null | tr -d '\r' | tee "$OUT/relaunch-2.txt"
TO=30 A logcat -d >> "$OUT/logcat-full.txt" 2>/dev/null || true

python3 - <<'PY' 2>&1 | tee -a "$OUT/selftest-result.txt"
import json, re
def check(cond, msg): print(("PASS " if cond else "FAIL ") + msg)
try:
    a = json.load(open("smoke/state.json"))
    n = sum(len(s["rows"]) for c in a["cars"] for s in c["sections"])
    r1 = open("smoke/relaunch-1.txt").read()
    m = re.search(r"rows (\d+)", r1)
    check(bool(m) and int(m.group(1)) == n > 0, f"entries after a full restart: {m.group(1) if m else '?'} (before {n})")
    r2 = open("smoke/relaunch-2.txt").read()
    check("استيراد سيارة" in r2, "Excel file opened while the app was closed shows the import summary: " + r2.replace(chr(10), "; ")[:160])
except Exception as e:
    print("FAIL restart checks:", repr(e))
perms = open("smoke/permissions.txt").read()
check("android.permission.INTERNET" not in perms, "the app has no internet permission")
log = open("smoke/logcat-full.txt", errors="replace").read().splitlines()
web = [l for l in log if "Capacitor" in l or "chromium" in l]
csp = [l for l in web if ("Content Security Policy" in l or "Refused to" in l) and "example.com" not in l]
check(not csp, "no security-policy (CSP) violations by the app itself" + ("" if not csp else ": " + csp[0][-160:]))
crash = [l for l in log if "FATAL EXCEPTION" in l or ("AndroidRuntime" in l and " E " in l)]
check(not crash, "no crash in the log" + ("" if not crash else ": " + crash[0][-160:]))
PY
grep -E "Capacitor|chromium|AndroidRuntime|XlsxImport|ActivityTaskManager" "$OUT/logcat-full.txt" | tail -600 > "$OUT/logcat.txt" || true
rm -f "$OUT/logcat-full.txt" "$OUT/ui.xml"
note "all done in $(( $(date +%s) - T0 ))s"
if grep -qE "^(FAIL|ERROR)" "$OUT/selftest-result.txt" || ! grep -q "^DONE" "$OUT/selftest-result.txt"; then
  echo "SMOKE TEST FAILED"; exit 1
fi
echo "SMOKE TEST OK"
