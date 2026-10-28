# -*- coding: utf-8 -*-
"""Эталоны: прогоняет каждый файл из tests/fixtures через cli/excel2arps.py в tests/expected/*.arp.
Страница в браузере должна выдать те же файлы байт в байт (tests/site_test.mjs).

    python tests/make_expected.py
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIX = os.path.join(HERE, "fixtures")
EXP = os.path.join(HERE, "expected")
CLI = os.path.join(ROOT, "cli", "excel2arps.py")
os.makedirs(EXP, exist_ok=True)

env = dict(os.environ, PYTHONUTF8="1")
for name in sorted(os.listdir(FIX)):
    stem, ext = os.path.splitext(name)
    if ext not in (".docx", ".xlsx", ".csv"):
        continue
    out = os.path.join(EXP, name + ".arp")  # example.docx.arp и example.xlsx.arp не затирают друг друга
    r = subprocess.run([sys.executable, CLI, os.path.join(FIX, name), out],
                       capture_output=True, text=True, encoding="utf-8", env=env)
    first = r.stdout.strip().splitlines()[0] if r.stdout.strip() else ""
    print(f"{name:14} {first.split(': ', 1)[-1]}  {r.stderr.strip()}")
