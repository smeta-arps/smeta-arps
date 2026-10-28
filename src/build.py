# -*- coding: utf-8 -*-
"""Сборка сайта: src/page.html + src/converter.js -> web/index.html
и архив для Windows web/download/smeta-arps-windows.zip.

    python src/build.py                                 для локального просмотра
    python src/build.py https://smeta-arps.ru           для выкладки: абсолютные адреса
                                                        превью ссылки, og:url, canonical, url в JSON-LD
    python src/build.py https://smeta-arps.ru --noindex то же, но поисковикам запрещено индексировать

Запускать из корня репозитория или откуда угодно, пути считаются от этого файла.
"""
import os
import sys
import zipfile

SRC = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SRC)
WEB = os.path.join(ROOT, "web")
OG_VER = "2"  # поднять при новой картинке превью, чтобы Telegram и iMessage не держали старую
GITHUB_URL = ""  # адрес репозитория; пока пусто, на месте кнопки надпись «скоро»

GITHUB_ICON = ('<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 .5a12 12 0 0 0-3.8 23.4c.6.1.8-.3.8-.6v-2.2c-3.3.7-4-1.4-4-1.4-.6-1.4-1.3-1.8-1.3-1.8-1.1-.7.1-.7.1-.7 1.2.1 1.8 1.2 1.8 1.2 1.1 1.8 2.8 1.3 3.5 1 .1-.8.4-1.3.8-1.6-2.7-.3-5.5-1.3-5.5-5.9 0-1.3.5-2.4 1.2-3.2-.1-.3-.5-1.5.1-3.2 0 0 1-.3 3.3 1.2a11.5 11.5 0 0 1 6 0C17.3 4.7 18.3 5 18.3 5c.7 1.7.2 2.9.1 3.2.8.8 1.2 1.9 1.2 3.2 0 4.6-2.8 5.6-5.5 5.9.4.4.8 1.1.8 2.2v3.3c0 .3.2.7.8.6A12 12 0 0 0 12 .5z"/></svg>')

args = [a for a in sys.argv[1:] if not a.startswith("--")]
noindex = "--noindex" in sys.argv
base = args[0].rstrip("/") if args else ""

page = open(os.path.join(SRC, "page.html"), encoding="utf-8").read()
js = open(os.path.join(SRC, "converter.js"), encoding="utf-8").read()

hero = ('<img class="hero-bg" src="img/hero-1920.jpg" srcset="img/hero-1280.jpg 1280w, img/hero-1920.jpg 1920w" '
        'sizes="100vw" width="1920" height="1072" alt="" fetchpriority="high">')

if GITHUB_URL:
    gh_button = '<a class="btn navy" href="%s">%s Код на GitHub</a>' % (GITHUB_URL, GITHUB_ICON)
    gh_foot = '<li><a href="%s">Код на GitHub</a></li>' % GITHUB_URL
else:
    gh_button = '<span class="btn off">%s Код на GitHub: скоро</span>' % GITHUB_ICON
    gh_foot = '<li class="soon">Код на GitHub: скоро</li>'

for mark in ("<!--CONVERTER_SCRIPT-->", "<!--HERO_IMG-->", "<!--GITHUB_BUTTON-->", "<!--GITHUB_FOOT-->"):
    assert mark in page, mark
out = (page.replace("<!--CONVERTER_SCRIPT-->", "<script>\n" + js + "</script>")
           .replace("<!--HERO_IMG-->", hero)
           .replace("<!--GITHUB_BUTTON-->", gh_button)
           .replace("<!--GITHUB_FOOT-->", gh_foot))
assert "—" not in out, "длинное тире в тексте страницы"

# версия для Windows: та же страница с относительными адресами, открывается двойным щелчком без интернета
offline = out.replace('<a class="btn" href="download/smeta-arps-windows.zip" download>', '<a class="btn" href="https://smeta-arps.ru/" hidden>')
offline = offline.replace('<p class="together-note">', '<p class="together-note" hidden>')
os.makedirs(os.path.join(WEB, "download"), exist_ok=True)
zpath = os.path.join(WEB, "download", "smeta-arps-windows.zip")
with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("Смета в АРПС.html", offline.encode("utf-8"))
    for rel in ("img/hero-1280.jpg", "img/hero-1920.jpg", "img/logo-mark.png", "img/mascot-estimator.png",
                "icon-192.png", "favicon-32x32.png", "favicon.ico", "apple-touch-icon.png",
                "templates/smeta-template.docx", "templates/smeta-template.xlsx", "templates/smeta-example.docx"):
        z.write(os.path.join(WEB, rel), rel)
    z.writestr("Как пользоваться.txt", (
        "Смета в АРПС, версия для работы без интернета\r\n\r\n"
        "1. Распакуйте архив в любую папку.\r\n"
        "2. Откройте файл «Смета в АРПС.html» двойным щелчком, он откроется в браузере.\r\n"
        "3. Перетащите на страницу таблицу Word, Excel или CSV и скачайте файл .arp.\r\n\r\n"
        "Файлы никуда не отправляются. Шаблоны лежат в папке templates.\r\n"
        "Сайт: https://smeta-arps.ru  Почта: hello@smeta-arps.ru\r\n").encode("utf-8"))

if base:
    out = out.replace('content="og.jpg"', 'content="%s/og.jpg?v=%s"' % (base, OG_VER))
    out = out.replace('<meta property="og:locale" content="ru_RU">',
                      '<meta property="og:locale" content="ru_RU">\n<meta property="og:url" content="%s/">\n<link rel="canonical" href="%s/">' % (base, base))
    out = out.replace('"name": "Смета в АРПС",\n  "description"',
                      '"name": "Смета в АРПС",\n  "url": "%s/",\n  "image": "%s/logo-512.png",\n  "description"' % (base, base), 1)
    out = out.replace('"@type": "WebSite",\n  "name": "Смета в АРПС",',
                      '"@type": "WebSite",\n  "name": "Смета в АРПС",\n  "url": "%s/",' % base, 1)
if noindex:
    out = out.replace('<meta name="viewport"', '<meta name="robots" content="noindex, nofollow">\n<meta name="viewport"', 1)

open(os.path.join(WEB, "index.html"), "w", encoding="utf-8", newline="\n").write(out)
print("web/index.html:", len(out.encode("utf-8")) // 1024, "KB", ("(адрес %s)" % base) if base else "(локально)",
      "noindex" if noindex else "")
print("web/download/smeta-arps-windows.zip:", os.path.getsize(zpath) // 1024, "KB")
