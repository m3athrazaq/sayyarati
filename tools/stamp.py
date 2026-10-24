#!/usr/bin/env python3
"""Keeps the web files consistent after any change to app.js, report.js or the files in vendor/ and fonts/:

- the integrity hashes of the files the app loads on demand (VENDOR in app.js),
- the address + integrity hash of app.js in index.html (app.js?v=<hash>),
- the service worker's cache version and file list (APPV and LAZY in sw.js), so installed copies update.

Run `python3 tools/stamp.py` after editing; CI runs `python3 tools/stamp.py --check`."""
import base64
import hashlib
import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
check = '--check' in sys.argv

# files app.js loads on demand (scripts and binary files)
ON_DEMAND = [
    'vendor/exceljs.min.js',
    'vendor/pdf-lib.min.js',
    'vendor/fontkit.umd.min.js',
    'vendor/bidi.min.js',
    'report.js',
    'fonts/readex-report-400.ttf',
    'fonts/readex-report-700.ttf',
]


def sri(data: bytes) -> str:
    return 'sha384-' + base64.b64encode(hashlib.sha384(data).digest()).decode()


def version_of(integrity: str) -> str:   # must match vendorUrl() in app.js
    return re.sub(r'[^A-Za-z0-9]', '', integrity[7:19])


changed = []


def write(path: pathlib.Path, old: str, new: str):
    if old != new:
        changed.append(path.name)
        if not check:
            path.write_text(new, encoding='utf-8')


vendor = {p: sri((ROOT / p).read_bytes()) for p in ON_DEMAND}

app_p = ROOT / 'app.js'
app = app_p.read_text(encoding='utf-8')
line = 'const VENDOR = ' + json.dumps(vendor, ensure_ascii=False, separators=(', ', ': ')) + ';'
new_app, n = re.subn(r'const VENDOR = \{[^\n]*\};', line, app, count=1)
if n != 1:
    sys.exit('VENDOR line not found in app.js')
write(app_p, app, new_app)

data = (new_app if not check else app).encode('utf-8')
tag = hashlib.sha256(data).hexdigest()[:12]
html_p = ROOT / 'index.html'
html = html_p.read_text(encoding='utf-8')
script = f'<script src="app.js?v={tag}" integrity="{sri(data)}"></script>'
new_html = re.sub(r'<script src="app\.js[^"]*"[^>]*></script>', script, html, count=1)
if script not in new_html:
    sys.exit('app.js script tag not found in index.html')
write(html_p, html, new_html)

sw_p = ROOT / 'sw.js'
sw = sw_p.read_text(encoding='utf-8')
lazy = 'const LAZY = ' + json.dumps([f'./{p}?v={version_of(v)}' for p, v in vendor.items()]) + ';'
new_sw = re.sub(r"const APPV = '[^']*';", f"const APPV = '{tag}';", sw, count=1)
new_sw, n = re.subn(r'const LAZY = \[[^\n]*\];', lazy, new_sw, count=1)
if f"const APPV = '{tag}';" not in new_sw or n != 1:
    sys.exit('APPV / LAZY lines not found in sw.js')
write(sw_p, sw, new_sw)

if check and changed:
    sys.exit('Run python3 tools/stamp.py — out of date: ' + ', '.join(changed))
print(('ok, up to date' if check else 'stamped ' + (', '.join(changed) or 'nothing (already up to date)')) + f' — app.js {tag}')
