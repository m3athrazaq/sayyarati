"""Prints "x y" (centre) of the first UI node whose text or content-desc matches a regex.
Usage: python3 find_node.py ui.xml 'regex'"""
import re
import sys
import xml.etree.ElementTree as ET

try:
    root = ET.parse(sys.argv[1]).getroot()
except Exception:
    sys.exit(0)
pat = re.compile(sys.argv[2])
for node in root.iter('node'):
    label = (node.get('text') or '') + '\n' + (node.get('content-desc') or '')
    if not pat.search(label):
        continue
    m = re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', node.get('bounds') or '')
    if not m:
        continue
    x1, y1, x2, y2 = map(int, m.groups())
    if x2 > x1 and y2 > y1:
        print((x1 + x2) // 2, (y1 + y2) // 2)
        break
