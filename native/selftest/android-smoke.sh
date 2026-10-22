#!/usr/bin/env bash
# Runs inside the Android emulator job: installs the test APK (with selftest.js), answers the
# self-test's "WAIT …" steps with adb (system file picker taps, Open with / Share intents, Back),
# and collects screenshots + results in smoke/.
set -u
PKG=com.m3athrazaq.sayyarati
APK=native/android/app/build/outputs/apk/debug/app-debug.apk
XLSX=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
OUT=smoke
mkdir -p "$OUT"

res() { adb shell run-as "$PKG" cat files/selftest-result.txt 2>/dev/null | tr -d '\r'; }
shot() { adb exec-out screencap -p > "$OUT/$1.png" 2>/dev/null || true; }
waitfor() {  # marker, seconds
  for _ in $(seq 1 "$2"); do
    if res | grep -qE "^($1|DONE|ERROR)"; then res | grep -qE "^$1" && return 0; echo "stopped before $1"; return 1; fi
    sleep 1
  done
  echo "timed out waiting for $1"; return 1
}
tap_text() {  # regex, seconds
  for _ in $(seq 1 "$2"); do
    adb shell uiautomator dump /sdcard/ui.xml >/dev/null 2>&1
    adb shell cat /sdcard/ui.xml > "$OUT/ui.xml" 2>/dev/null
    XY=$(python3 native/selftest/find_node.py "$OUT/ui.xml" "$1")
    if [ -n "$XY" ]; then echo "tap '$1' at $XY"; adb shell input tap $XY; return 0; fi
    sleep 1
  done
  echo "not found on screen: $1"; return 1
}

adb shell settings put global window_animation_scale 0
adb shell settings put global transition_animation_scale 0
adb shell settings put global animator_duration_scale 0
adb install -r "$APK"
adb shell getprop ro.build.version.release | sed 's/^/Android /' | tee "$OUT/device.txt"
adb shell dumpsys package com.google.android.webview | grep -m1 versionName | sed 's/^ */WebView /' | tee -a "$OUT/device.txt"
adb shell dumpsys package "$PKG" | sed -n '/requested permissions:/,/install permissions:/p' > "$OUT/permissions.txt"

# the file the user picks in the system picker
adb push native/selftest/fixture-f150-layout.xlsx /sdcard/Download/sayyarati-test.xlsx
adb shell content call --method scan_volume --uri content://media --arg external_primary >/dev/null 2>&1 || true
adb shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d file:///sdcard/Download/sayyarati-test.xlsx >/dev/null 2>&1 || true

adb logcat -c
adb shell am start -W -n "$PKG/.MainActivity"
sleep 6; shot 00-launch

if waitfor "WAIT picker" 120; then
  sleep 5; shot 01-picker
  if ! tap_text "sayyarati-test" 8; then
    tap_text "Show roots" 6 || true
    sleep 2; shot 01b-roots
    tap_text "^Downloads$" 8 || true
    sleep 3; shot 01c-downloads
    tap_text "sayyarati-test" 12 || true
  fi
fi
waitfor "STEP picker-summary" 150 && shot 02-picker-summary

if waitfor "WAIT view-intent" 60; then
  adb shell am start -a android.intent.action.VIEW -t "$XLSX" \
    -d "content://$PKG.fileprovider/cache/incoming/open.xlsx" -n "$PKG/.MainActivity"
fi
waitfor "STEP view-summary" 90 && shot 03-open-with

if waitfor "WAIT share-sheet" 60; then
  sleep 5; shot 04-share-sheet
  adb shell input keyevent KEYCODE_BACK
fi
waitfor "STEP after-share" 90 && shot 05-after-share

if waitfor "WAIT send-intent" 60; then
  adb shell am start -a android.intent.action.SEND -t "$XLSX" \
    --eu android.intent.extra.STREAM "content://$PKG.fileprovider/cache/incoming/backup.xlsx" -n "$PKG/.MainActivity"
fi
waitfor "STEP restore" 90 && shot 06-restore
waitfor "STEP home" 90 && shot 07-home
waitfor "DONE" 90
sleep 2
res | tee "$OUT/selftest-result.txt"
adb exec-out run-as "$PKG" cat files/sayyarati-state.json > "$OUT/state.json" 2>/dev/null
adb logcat -d > "$OUT/logcat-full.txt" 2>/dev/null || true

# data must survive a full restart of the app
adb shell am force-stop "$PKG"
adb shell am start -W -n "$PKG/.MainActivity"
sleep 16; shot 08-relaunch
adb shell run-as "$PKG" cat files/selftest-relaunch.txt 2>/dev/null | tr -d '\r' | tee "$OUT/relaunch-1.txt"

# an Excel file opened with the app while it is closed (cold start from Drive / Files / WhatsApp)
adb shell am force-stop "$PKG"
adb shell run-as "$PKG" rm -f files/selftest-relaunch.txt
adb shell am start -W -a android.intent.action.VIEW -t "$XLSX" \
  -d "content://$PKG.fileprovider/cache/incoming/cold.xlsx" -n "$PKG/.MainActivity"
sleep 16; shot 09-cold-open
adb shell run-as "$PKG" cat files/selftest-relaunch.txt 2>/dev/null | tr -d '\r' | tee "$OUT/relaunch-2.txt"
adb logcat -d >> "$OUT/logcat-full.txt" 2>/dev/null || true

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
grep -E "Capacitor|chromium|AndroidRuntime|XlsxImport" "$OUT/logcat-full.txt" | tail -400 > "$OUT/logcat.txt" || true
rm -f "$OUT/logcat-full.txt" "$OUT/ui.xml"
grep -c "^PASS" "$OUT/selftest-result.txt" | sed 's/^/passed: /'
if grep -qE "^(FAIL|ERROR)" "$OUT/selftest-result.txt" || ! grep -q "^DONE" "$OUT/selftest-result.txt"; then
  echo "SMOKE TEST FAILED"; exit 1
fi
echo "SMOKE TEST OK"
