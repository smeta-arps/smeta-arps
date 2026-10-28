// Конвертер сметы Word / Excel / CSV -> АРПС 1.10 для Estimate.
// Работает в браузере, без сервера и библиотек. Логика совпадает с cli/excel2arps.py,
// проверка: tests/site_test.mjs сравнивает результат с tests/expected байт в байт.
(() => {
  'use strict';

  // ---------- константы ----------
  const PROGRAM = 'SmetaARPS';
  const VERSION = '0.7';
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const TRIGGER = '30#1#чел.-ч#Оплата труда рабочих#0#1.000000###';
  const COL_ALIASES = {
    section: ['раздел', 'section'],
    code: ['шифр', 'код', 'обоснование', 'шифр расценки', 'code'],
    name: ['наименование', 'название', 'работы', 'наименование работ', 'name'],
    unit: ['ед.изм.', 'ед. изм.', 'ед.изм', 'ед. изм', 'единица', 'единица измерения', 'ед', 'unit'],
    qty: ['объём', 'объем', 'кол-во', 'количество', 'qty'],
  };
  const META_KEYS = {
    num: ['смета №', 'смета n', 'номер сметы', '№ сметы', 'номер'],
    name: ['название', 'наименование сметы', 'название сметы'],
    obj: ['объект', 'наименование объекта', 'стройка'],
  };
  const TOTALS = ['итого', 'всего', 'в том числе', 'ндс', 'накладные', 'сметная прибыль'];
  // верхняя половина кодовой страницы DOS 866 (байты 0x80-0xFF)
  const CP866_HI = "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмноп░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀рстуфхцчшщъыьэюяЁёЄєЇїЎў°∙·√№¤■ ";
  const ENC866 = new Map();
  for (let i = 0; i < 128; i++) ENC866.set(CP866_HI[i], 0x80 + i);

  function encode866(text) {
    const bytes = [];
    for (const ch of text) {
      const cp = ch.codePointAt(0);
      if (cp < 0x80) bytes.push(cp);
      else bytes.push(ENC866.has(ch) ? ENC866.get(ch) : 0x3f);
    }
    return new Uint8Array(bytes);
  }

  // ---------- ZIP (docx и xlsx это zip-архивы) ----------
  async function openZip(buf) {
    const u8 = new Uint8Array(buf);
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Файл не похож на .docx или .xlsx. Откройте его в Word или Excel и сохраните заново.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder('utf-8');
    const entries = new Map();
    for (let k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Архив внутри файла повреждён.');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true);
      const xlen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      entries.set(dec.decode(u8.subarray(p + 46, p + 46 + nlen)), { method, csize, off });
      p += 46 + nlen + xlen + clen;
    }
    return {
      has: (name) => entries.has(name),
      async text(name) {
        const e = entries.get(name);
        if (!e) throw new Error('В файле нет части ' + name + '. Сохраните его заново в Word или Excel.');
        const start = e.off + 30 + dv.getUint16(e.off + 26, true) + dv.getUint16(e.off + 28, true);
        const data = u8.subarray(start, start + e.csize);
        let raw;
        if (e.method === 0) raw = data;
        else if (e.method === 8) {
          if (typeof DecompressionStream === 'undefined') throw new Error('Браузер слишком старый. Откройте страницу в свежем Chrome, Edge, Firefox или Safari.');
          const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
          raw = new Uint8Array(await new Response(stream).arrayBuffer());
        } else throw new Error('Неизвестный способ сжатия внутри файла.');
        return dec.decode(raw);
      },
    };
  }

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Не удалось прочитать содержимое файла.');
    return doc;
  }
  const kids = (el, ns, name) => Array.from(el.children).filter((n) => n.namespaceURI === ns && n.localName === name);

  // ---------- Excel ----------
  function colToIdx(ref) {
    const letters = /^[A-Z]+/.exec(ref)[0];
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  async function readXlsx(zip) {
    const shared = [];
    if (zip.has('xl/sharedStrings.xml')) {
      const ss = parseXml(await zip.text('xl/sharedStrings.xml'));
      for (const si of kids(ss.documentElement, NS_S, 'si')) {
        shared.push(Array.from(si.getElementsByTagNameNS(NS_S, 't')).map((t) => t.textContent).join(''));
      }
    }
    const wb = parseXml(await zip.text('xl/workbook.xml'));
    const sheet = wb.getElementsByTagNameNS(NS_S, 'sheet')[0];
    const rid = sheet.getAttributeNS(NS_R, 'id');
    const rels = parseXml(await zip.text('xl/_rels/workbook.xml.rels'));
    let target = null;
    for (const r of Array.from(rels.documentElement.children)) if (r.getAttribute('Id') === rid) target = r.getAttribute('Target');
    target = target.replace(/^\/+/, '');
    if (!target.startsWith('xl/')) target = 'xl/' + target;
    const sh = parseXml(await zip.text(target));
    const sheetData = sh.getElementsByTagNameNS(NS_S, 'sheetData')[0];
    const rows = [];
    for (const row of kids(sheetData, NS_S, 'row')) {
      const cells = new Map();
      for (const c of kids(row, NS_S, 'c')) {
        const idx = colToIdx(c.getAttribute('r'));
        const t = c.getAttribute('t');
        const v = kids(c, NS_S, 'v')[0];
        let val;
        if (t === 's' && v) val = shared[parseInt(v.textContent, 10)];
        else if (t === 'inlineStr') val = Array.from(c.getElementsByTagNameNS(NS_S, 't')).map((x) => x.textContent).join('');
        else if (v) val = v.textContent;
        else val = '';
        cells.set(idx, (val || '').trim());
      }
      if (cells.size) {
        const width = Math.max(...cells.keys()) + 1;
        rows.push(Array.from({ length: width }, (_, i) => (cells.has(i) ? cells.get(i) : '')));
      }
    }
    return rows;
  }

  // ---------- Word ----------
  function collectRun(node, acc) {
    for (const ch of Array.from(node.children)) {
      if (ch.namespaceURI === NS_W) {
        if (ch.localName === 't') acc.push(ch.textContent);
        else if (ch.localName === 'tab' || ch.localName === 'br' || ch.localName === 'cr') acc.push(' ');
      }
      collectRun(ch, acc);
    }
  }

  function wText(el) {
    const paras = (el.namespaceURI === NS_W && el.localName === 'p') ? [el] : Array.from(el.getElementsByTagNameNS(NS_W, 'p'));
    const parts = paras.map((p) => { const acc = []; collectRun(p, acc); return acc.join(''); });
    return parts.join(' ').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  }

  function wTableRows(tbl) {
    const rows = [];
    for (const tr of kids(tbl, NS_W, 'tr')) {
      const row = [];
      for (const tc of kids(tr, NS_W, 'tc')) {
        row.push(wText(tc));
        const pr = kids(tc, NS_W, 'tcPr')[0];
        const span = pr ? kids(pr, NS_W, 'gridSpan')[0] : null;
        if (span) {
          const n = parseInt(span.getAttributeNS(NS_W, 'val') || '1', 10);
          for (let i = 1; i < n; i++) row.push('');
        }
      }
      rows.push(row);
    }
    return rows;
  }

  async function readDocx(zip) {
    const doc = parseXml(await zip.text('word/document.xml'));
    const body = kids(doc.documentElement, NS_W, 'body')[0];
    const metaRows = [];
    for (const el of Array.from(body.children)) {
      if (el.namespaceURI !== NS_W) continue;
      if (el.localName === 'p') {
        const t = wText(el);
        const i = t.indexOf(':');
        if (i >= 0) metaRows.push([t.slice(0, i).trim(), t.slice(i + 1).trim()]);
      } else if (el.localName === 'tbl') {
        const rows = wTableRows(el);
        try { findHeader(rows); } catch (e) { continue; }
        return metaRows.concat(rows);
      }
    }
    throw new Error('В документе нет таблицы с колонками «Шифр» и «Объём».');
  }

  // ---------- CSV ----------
  function readCsv(text) {
    const sample = text.slice(0, 4096);
    const d = (sample.split(';').length - 1) >= (sample.split('\t').length - 1) ? ';' : '\t';
    const rows = [];
    let row = [];
    let f = '';
    let inQ = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQ) {
        if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else inQ = false; } else f += c;
      } else if (c === '"' && f === '') inQ = true;
      else if (c === d) { row.push(f); f = ''; }
      else if (c === '\r' || c === '\n') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(f);
        rows.push(row.length === 1 && row[0] === '' ? [] : row);
        row = []; f = '';
      } else f += c;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows.map((r) => r.map((x) => x.trim()));
  }

  // ---------- общая логика (как в tools/excel2arps.py) ----------
  function findHeader(rows) {
    for (let i = 0; i < rows.length; i++) {
      const low = rows[i].map((c) => c.replace(/\(.*?\)/g, '').replace(/\s+/g, ' ').toLowerCase().trim());
      const m = {};
      for (const [key, aliases] of Object.entries(COL_ALIASES)) {
        const j = low.findIndex((cell) => aliases.includes(cell));
        if (j >= 0) m[key] = j;
      }
      if ('code' in m && 'qty' in m) return [i, m];
    }
    throw new Error('Не нашёл строку заголовков: нужны колонки «Шифр» и «Объём» (плюс «Наименование» и «Ед. изм.»).');
  }

  function parseQty(s) {
    const t = s.replace(/ /g, '').replace(/ /g, '').replace(/,/g, '.');
    return /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t) ? parseFloat(t) : null;
  }

  function normalizeCode(code) {
    let c = code.trim().replace(/ /g, '');
    c = c.replace(/^(ТЕР|ФЕР|ГЭСН)(?=\d)/iu, '');
    c = c.replace(/^(ТЕР|ФЕР|ГЭСН)(р|м|п)(?=\d)/iu, '$2');
    return c;
  }

  const clean = (s) => s.replace(/#/g, ' ');

  function readMeta(rows, hi) {
    const meta = {};
    for (const row of rows.slice(0, hi)) {
      if (row.length < 2) continue;
      const key = row[0].toLowerCase().trim().replace(/:+$/, '');
      for (const [k, aliases] of Object.entries(META_KEYS)) if (aliases.includes(key) && row[1]) meta[k] = row[1];
    }
    return meta;
  }

  function build(rows, over) {
    const [hi, m] = findHeader(rows);
    const meta = readMeta(rows, hi);
    const num = over.num || meta.num || 'ЛС-1';
    const name = over.name || meta.name || 'Локальная смета';
    const obj = over.obj || meta.obj || '';
    const out = ['1#АРПС 1.10#' + PROGRAM + '#' + VERSION];
    out.push('3#' + ['', '', clean(obj), clean(num), clean(name), '', '', '', '', '', '', '', '', '', '2001', '', '', '', '', ''].join('#'));
    const view = [];
    const warnings = [];
    let sec = 0, pos = 0, haveSection = false, noCode = 0;
    const addSection = (title) => { sec++; out.push('10#0#' + sec + '#' + clean(title)); view.push({ type: 'sec', title }); haveSection = true; };
    for (const row of rows.slice(hi + 1)) {
      const get = (k) => (k in m && m[k] < row.length ? row[m[k]] : '');
      const section = get('section'), code = get('code'), nm = get('name'), unit = get('unit'), qty = get('qty');
      let q = qty ? parseQty(qty) : null;
      if (section && !code) { addSection(section); continue; }
      if (!code && q === null && !unit) {
        if (TOTALS.some((t) => nm.toLowerCase().startsWith(t))) continue;
        if (nm) addSection(nm);
        continue;
      }
      if (!code && !nm) continue;
      if (!haveSection) addSection('Раздел 1');
      if (q === null) { warnings.push('Нет объёма у строки «' + (code + ' ' + nm).trim() + '», поставлен 0.'); q = 0; }
      pos++;
      if (!code) noCode++;
      const qs = q.toFixed(4);
      out.push(['20', String(pos), clean(normalizeCode(code)), clean(unit), clean(nm)].concat(Array(20).fill(''), ['', qs, '0', '', '']).join('#'));
      if (code) out.push(TRIGGER);
      view.push({ type: 'pos', n: pos, code: normalizeCode(code), name: nm, unit, qty: qs.replace(/\.?0+$/, '').replace('.', ',') });
    }
    return { text: out.join('\r\n') + '\r\n', positions: pos, sections: sec, noCode, warnings, view, meta: { num, name, obj }, fileMeta: meta };
  }

  // ---------- интерфейс ----------
  const $ = (id) => document.getElementById(id);
  const drop = $('drop'), input = $('file'), result = $('result'), errorBox = $('error');
  let rows = null;
  let baseName = 'smeta';

  const plural = (n, one, few, many) => {
    const a = n % 10, b = n % 100;
    return n + ' ' + (a === 1 && b !== 11 ? one : a >= 2 && a <= 4 && (b < 12 || b > 14) ? few : many);
  };

  function showError(msg) {
    result.hidden = true;
    drop.hidden = false;
    errorBox.textContent = msg;
    errorBox.hidden = false;
  }

  function render() {
    const r = build(rows, { num: $('fNum').value.trim(), name: $('fName').value.trim(), obj: $('fObj').value.trim() });
    const chips = [];
    chips.push(['ok', plural(r.positions, 'позиция', 'позиции', 'позиций')]);
    chips.push(['', plural(r.sections, 'раздел', 'раздела', 'разделов')]);
    if (r.noCode) chips.push(['warn', r.noCode + ' без шифра']);
    $('summary').replaceChildren(...chips.map(([cls, text]) => { const s = document.createElement('span'); s.className = 'chip ' + cls; s.textContent = text; return s; }));
    const w = $('warnings');
    const msgs = r.warnings.slice();
    if (r.positions === 0) msgs.unshift('Позиций не найдено: проверьте, что под заголовками есть строки с шифром или объёмом.');
    w.replaceChildren(...msgs.map((t) => { const li = document.createElement('li'); li.textContent = t; return li; }));
    w.hidden = msgs.length === 0;
    const tb = $('rows');
    tb.replaceChildren(...r.view.map((v) => {
      const tr = document.createElement('tr');
      if (v.type === 'sec') {
        tr.className = 'sec';
        const td = document.createElement('td'); td.colSpan = 5; td.textContent = v.title; tr.append(td);
      } else {
        if (!v.code) tr.className = 'nocode';
        const cells = [['num', String(v.n)], ['code', v.code], ['', v.name], ['', v.unit], ['qty', v.qty]];
        for (const [cls, text] of cells) { const td = document.createElement('td'); if (cls) td.className = cls; td.textContent = text; tr.append(td); }
        if (!v.code) { const tag = document.createElement('span'); tag.className = 'tag'; tag.textContent = 'без шифра'; tr.children[1].append(tag); }
      }
      return tr;
    }));
    $('download').disabled = r.positions === 0;
    return r;
  }

  async function handle(file) {
    errorBox.hidden = true;
    try {
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      if (ext === 'doc' || ext === 'xls') throw new Error('Это старый формат .' + ext + '. Откройте файл в ' + (ext === 'doc' ? 'Word' : 'Excel') + ' и сохраните как .' + ext + 'x.');
      if (!['docx', 'xlsx', 'csv'].includes(ext)) throw new Error('Нужен файл .docx, .xlsx или .csv.');
      const buf = await file.arrayBuffer();
      if (ext === 'csv') rows = readCsv(new TextDecoder('utf-8').decode(buf));
      else {
        const zip = await openZip(buf);
        rows = ext === 'docx' ? await readDocx(zip) : await readXlsx(zip);
      }
      const [hi] = findHeader(rows);
      const meta = readMeta(rows, hi);
      $('fNum').value = meta.num || 'ЛС-1';
      $('fName').value = meta.name || 'Локальная смета';
      $('fObj').value = meta.obj || '';
      baseName = file.name.replace(/\.[^.]+$/, '') || 'smeta';
      $('fileName').textContent = file.name;
      render();
      drop.hidden = true;
      result.hidden = false;
    } catch (e) {
      rows = null;
      showError(e.message || String(e));
    }
  }

  input.addEventListener('change', () => { if (input.files[0]) handle(input.files[0]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) handle(f); });
  // не даём браузеру открыть файл, если его бросили мимо
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
  ['fNum', 'fName', 'fObj'].forEach((id) => $(id).addEventListener('input', () => { if (rows) render(); }));
  $('another').addEventListener('click', () => { result.hidden = true; errorBox.hidden = true; drop.hidden = false; rows = null; input.click(); });
  $('download').addEventListener('click', () => {
    if (!rows) return;
    const r = render();
    const blob = new Blob([encode866(r.text)], { type: 'application/octet-stream' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = baseName + '.arp';
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });

  // для автоматических проверок
  window.__smeta = { build, readCsv, encode866, openZip, readDocx, readXlsx };
})();
