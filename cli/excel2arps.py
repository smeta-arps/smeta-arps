# -*- coding: utf-8 -*-
"""
Word / Excel -> АРПС 1.10 для импорта в Estimate (Самара-КО-Инвест).

Вход:
  .docx - первая таблица с колонками "Шифр" и "Объём"; шапка из строк над таблицей вида "Смета №: ЛС-1",
          "Название: ...", "Объект: ...". Всё, что ниже таблицы (памятка, пример), игнорируется.
  .xlsx - первый лист; шапка в колонках A-B над заголовками.
  .csv  - разделитель ; или табуляция, UTF-8.
  .doc  - не читается, пересохранить в Word как .docx.
Колонки по заголовкам (регистр, порядок и пояснения в скобках не важны, лишние колонки игнорируются):
    [Раздел] | Шифр | Наименование | Ед.изм. | Объём
Раздел: строка с заполненной колонкой "Раздел" или строка, где есть только наименование (нет шифра, единицы и объёма).
Строки "Итого", "Всего", "НДС" пропускаются.
Позиция без шифра, но с объёмом попадает в смету с наименованием и объёмом, шифр ставит сметчица в Estimate.

Выход: текстовый файл АРПС 1.10 в диалекте Estimate 1.9 (cp866, поля через #, CRLF).
На каждую расценку добавляется одна строка ресурса "Оплата труда рабочих" - это триггер:
без неё Estimate при перепривязке не подтягивает расценку из базы (проверено на демо 1.9, 2026-10-21).

Запуск:
    python excel2arps.py smeta.docx out.arp
    python excel2arps.py smeta.xlsx out.arp --num ЛС-1 --name "Ремонт кровли" --object "Жилой дом, ул. Ленина 5"
Без зависимостей: docx и xlsx читаются через zipfile + xml.
"""
import argparse
import csv
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
COL_ALIASES = {
    "section": ("раздел", "section"),
    "code": ("шифр", "код", "обоснование", "шифр расценки", "code"),
    "name": ("наименование", "название", "работы", "наименование работ", "name"),
    "unit": ("ед.изм.", "ед. изм.", "ед.изм", "ед. изм", "единица", "единица измерения", "ед", "unit"),
    "qty": ("объём", "объем", "кол-во", "количество", "qty"),
}


def col_to_idx(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
        # первый лист по порядку в workbook
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid = wb.find("m:sheets/m:sheet", NS).get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = None
        for r in rels:
            if r.get("Id") == rid:
                target = r.get("Target")
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        sheet = ET.fromstring(z.read(target))
        rows = []
        for row in sheet.find("m:sheetData", NS).findall("m:row", NS):
            cells = {}
            for c in row.findall("m:c", NS):
                idx = col_to_idx(c.get("r"))
                t = c.get("t")
                v = c.find("m:v", NS)
                if t == "s" and v is not None:
                    val = shared[int(v.text)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
                elif v is not None:
                    val = v.text
                else:
                    val = ""
                cells[idx] = (val or "").strip()
            if cells:
                width = max(cells) + 1
                rows.append([cells.get(i, "") for i in range(width)])
        return rows


def _w_text(el):
    """Текст элемента Word: абзацы через пробел, табуляции и переносы как пробел."""
    parts = []
    for p in el.iter(W + "p"):
        buf = []
        for node in p.iter():
            if node.tag == W + "t":
                buf.append(node.text or "")
            elif node.tag in (W + "tab", W + "br", W + "cr"):
                buf.append(" ")
        parts.append("".join(buf))
    return re.sub(r"\s+", " ", " ".join(parts).replace("\u00a0", " ")).strip()


def _w_table_rows(tbl):
    rows = []
    for tr in tbl.findall(W + "tr"):
        row = []
        for tc in tr.findall(W + "tc"):
            row.append(_w_text(tc))
            span = tc.find(W + "tcPr/" + W + "gridSpan")
            if span is not None:
                row.extend([""] * (int(span.get(W + "val", "1")) - 1))
        rows.append(row)
    return rows


def read_docx(path):
    """Строки шапки (абзацы вида Ключ: значение до таблицы) + строки первой таблицы с колонками Шифр и Объём."""
    with zipfile.ZipFile(path) as z:
        body = ET.fromstring(z.read("word/document.xml")).find(W + "body")
    meta_rows = []
    for el in body:
        if el.tag == W + "p":
            t = _w_text(el)
            if ":" in t:
                k, v = t.split(":", 1)
                meta_rows.append([k.strip(), v.strip()])
        elif el.tag == W + "tbl":
            rows = _w_table_rows(el)
            try:
                find_header(rows)
            except SystemExit:
                continue  # таблица без нужных колонок
            return meta_rows + rows
    raise SystemExit("В документе нет таблицы с колонками Шифр и Объём")


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        delim = ";" if sample.count(";") >= sample.count("\t") else "\t"
        return [[c.strip() for c in r] for r in csv.reader(f, delimiter=delim)]


def find_header(rows):
    for i, row in enumerate(rows):
        low = [re.sub(r"\s+", " ", re.sub(r"\(.*?\)", "", c)).lower().strip() for c in row]
        mapping = {}
        for key, aliases in COL_ALIASES.items():
            for j, cell in enumerate(low):
                if cell in aliases:
                    mapping[key] = j
                    break
        if "code" in mapping and "qty" in mapping:
            return i, mapping
    raise SystemExit("Не нашёл строку заголовков: нужны колонки 'Шифр' и 'Объём' (плюс Раздел, Наименование, Ед.изм.)")


def parse_qty(s):
    s = s.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def normalize_code(code):
    """ТЕР11-01-011-01, ТЕРр61-1-1, 11-01-011-01 -> как в базе Estimate: без префикса ТЕР у строительных,
    у ремонтных (ТЕРр) префикс 'р' оставляем, т.к. в Estimate это отдельная база ТЕРр."""
    c = code.strip().replace(" ", "")
    c = re.sub(r"^(ТЕР|ФЕР|ГЭСН)(?=\d)", "", c, flags=re.I)
    c = re.sub(r"^(ТЕР|ФЕР|ГЭСН)(р|м|п)(?=\d)", r"\2", c, flags=re.I)
    return c


META_KEYS = {
    "num": ("смета №", "смета n", "номер сметы", "№ сметы", "номер"),
    "name": ("название", "наименование сметы", "название сметы"),
    "obj": ("объект", "наименование объекта", "стройка"),
}


def read_meta(rows, hi):
    """Шапка шаблона над заголовками: подпись в колонке A, значение в колонке B."""
    meta = {}
    for row in rows[:hi]:
        if len(row) < 2:
            continue
        key = row[0].lower().strip().rstrip(":")
        for k, aliases in META_KEYS.items():
            if key in aliases and row[1]:
                meta[k] = row[1]
    return meta


def clean(s):
    """Знак # разделяет поля АРПС, в тексте его быть не должно."""
    return s.replace("#", " ")


def build(rows, num, name, obj, program="SmetaARPS", version="0.7"):
    hi, m = find_header(rows)
    meta = read_meta(rows, hi)
    num = num or meta.get("num", "ЛС-1")
    name = name or meta.get("name", "Локальная смета")
    obj = obj or meta.get("obj", "")
    out = ["1#АРПС 1.10#%s#%s" % (program, version)]
    header = ["", "", clean(obj), clean(num), clean(name), "", "", "", "", "", "", "", "", "", "2001", "", "", "", "", ""]
    out.append("3#" + "#".join(header))
    sec_no = 0
    pos_no = 0
    no_code = []
    have_section = False
    for row in rows[hi + 1:]:
        get = lambda k: row[m[k]] if k in m and m[k] < len(row) else ""
        section, code, nm, unit, qty = get("section"), get("code"), get("name"), get("unit"), get("qty")
        q = parse_qty(qty) if qty else None
        if section and not code:
            sec_no += 1
            out.append("10#0#%d#%s" % (sec_no, clean(section)))
            have_section = True
            continue
        if not code and q is None and not unit:
            if nm.lower().startswith(("итого", "всего", "в том числе", "ндс", "накладные", "сметная прибыль")):
                continue  # итоги в Estimate не нужны, программа считает их сама
            if nm:  # строка-заголовок раздела в колонке наименования
                sec_no += 1
                out.append("10#0#%d#%s" % (sec_no, clean(nm)))
                have_section = True
            continue
        if not code and not nm:
            continue
        if not have_section:
            sec_no += 1
            out.append("10#0#%d#%s" % (sec_no, "Раздел 1"))
            have_section = True
        if q is None:
            print("ПРЕДУПРЕЖДЕНИЕ: нет объёма у '%s %s', ставлю 0" % (code, nm), file=sys.stderr)
            q = 0.0
        pos_no += 1
        if not code:
            no_code.append("%d. %s" % (pos_no, nm))
        fields = ["20", str(pos_no), clean(normalize_code(code)), clean(unit), clean(nm)] + [""] * 20 + ["", "%.4f" % q, "0", "", ""]
        out.append("#".join(fields))
        if code:  # триггер для перепривязки; без шифра заменять нечем, строку не добавляем
            out.append("30#1#чел.-ч#Оплата труда рабочих#0#1.000000###")
    return "\r\n".join(out) + "\r\n", pos_no, sec_no, no_code


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--num", default="", help="номер сметы (иначе из шапки шаблона, ячейка 'Смета №')")
    ap.add_argument("--name", default="", help="название сметы (иначе из шапки 'Название')")
    ap.add_argument("--object", default="", help="объект (иначе из шапки 'Объект')")
    a = ap.parse_args()
    ext = a.src.lower().rsplit(".", 1)[-1]
    if ext == "doc":
        raise SystemExit("Файл .doc старого формата: откройте его в Word и сохраните как .docx")
    readers = {"docx": read_docx, "xlsx": read_xlsx, "csv": read_csv}
    if ext not in readers:
        raise SystemExit("Не знаю формат .%s, нужен .docx, .xlsx или .csv" % ext)
    rows = readers[ext](a.src)
    data, n, s, no_code = build(rows, a.num, a.name, a.object)
    with open(a.dst, "wb") as f:
        f.write(data.encode("cp866", errors="replace"))
    print("%s: %d позиций, %d разделов" % (a.dst, n, s))
    if no_code:
        print("Без шифра (%d), их подберёт сметчица:" % len(no_code))
        for line in no_code:
            print("   " + line)
    if n == 0:
        print("ПРЕДУПРЕЖДЕНИЕ: позиций нет, проверь колонки Шифр и Объём", file=sys.stderr)


if __name__ == "__main__":
    main()
