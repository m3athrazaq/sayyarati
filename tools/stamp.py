#!/usr/bin/env python3
"""Keeps the web files consistent after any change to app.js or vendor/exceljs.min.js:

- the integrity hash of ExcelJS inside app.js (EXCELJS_SRI),
- the address + integrity hash of app.js in index.html (app.js?v=<hash>),
- the service worker's cache version (APPV in sw.js), so installed copies update.

Run `python3 tools/stamp.py` after editing; CI runs `python3 tools/stamp.py --check`."""
import base64
import hashlib
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
check = '--check' in sys.argv


def sri(data: bytes) -> str:
    return 'sha384-' + base64.b64encode(hashlib.sha384(data).digest()).decode()


changed = []


def write(path: pathlib.Path, old: str, new: str):
    if old != new:
        changed.append(path.name)
        if not check:
            path.write_text(new, encoding='utf-8')


app = ROOT / 'app.js'
src = app.read_text(encoding='utf-8')
ex = sri((ROOT / 'vendor' / 'exceljs.min.js').read_bytes())
new = re.sub(r"const EXCELJS_SRI = '[^']*';", f"const EXCELJS_SRI = '{ex}';", src, count=1)
if new == src and f"const EXCELJS_SRI = '{ex}';" not in src:
    sys.exit('EXCELJS_SRI line not found in app.js')
write(app, src, new)

data = (new if not check else src).encode('utf-8')
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
new_sw = re.sub(r"const APPV = '[^']*';", f"const APPV = '{tag}';", sw, count=1)
if f"const APPV = '{tag}';" not in new_sw:
    sys.exit('APPV line not found in sw.js')
write(sw_p, sw, new_sw)

if check and changed:
    sys.exit('Run python3 tools/stamp.py — out of date: ' + ', '.join(changed))
print(('ok, up to date' if check else 'stamped ' + (', '.join(changed) or 'nothing (already up to date)')) + f' — app.js {tag}')
