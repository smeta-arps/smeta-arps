# -*- coding: utf-8 -*-
"""Убирает имя автора из файлов Word и Excel (docProps/core.xml, docProps/app.xml).
Содержимое документа не меняется.

    python tests/clean_office_meta.py файл.docx [файл.xlsx ...]
"""
import re
import sys
import zipfile

NAME = "Смета в АРПС"


def clean(path):
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    out = []
    changed = False
    for info, data in items:
        if info.filename == "docProps/core.xml":
            s = data.decode("utf-8")
            new = re.sub(r"(<dc:creator>)[^<]*(</dc:creator>)", r"\g<1>%s\g<2>" % NAME, s)
            new = re.sub(r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)", r"\g<1>%s\g<2>" % NAME, new)
            changed |= new != s
            data = new.encode("utf-8")
        elif info.filename == "docProps/app.xml":
            s = data.decode("utf-8")
            new = re.sub(r"<(Company|Manager)>[^<]*</\1>", r"<\1></\1>", s)
            changed |= new != s
            data = new.encode("utf-8")
        out.append((info, data))
    if changed:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for info, data in out:
                z.writestr(info, data)
    return changed


if __name__ == "__main__":
    for p in sys.argv[1:]:
        print(("очищен  " if clean(p) else "без изменений  ") + p)
