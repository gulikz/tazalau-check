/* Tazalau помощник: версия 1.
   Закладка для tazalau.qoldau.kz (Edge или Chrome). Ничего не устанавливает.
   По файлу Excel (столбец «ИИН/БИН»): открывает реестр с фильтром по ИИН, проверяет, что на странице
   именно результат поиска, сохраняет снимок страницы и пишет ссылку на него в столбец tazalau_file_location.
   Вход на сайт и ЭЦП не нужны. Файл Excel читается и изменяется только внутри браузера и никуда не отправляется. */
(() => {
  'use strict';
  const VERSION = '1';
  if (window.__tazalau) { window.__tazalau.show(); return; }
  if (!/(^|\.)tazalau\.qoldau\.kz$/i.test(location.hostname)) {
    const check = document.getElementById('tazalau-check');
    if (check) {                    /* закладку нажали на странице установки: это проверка */
      check.textContent = '✔ Закладка работает (версия ' + VERSION + '). ' +
        'Теперь откройте tazalau.qoldau.kz и нажмите закладку Tazalau ▶ там.';
      check.className = 'check good';
      return;
    }
    if (confirm('Помощник работает только на сайте tazalau.qoldau.kz.\n\nОткрыть сайт сейчас? ' +
                'Потом нажмите закладку Tazalau ещё раз.')) location.href = 'https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery';
    return;
  }

  /* ======================= НАСТРОЙКИ ======================= */
  const REGISTRIES = [
    { key: 'recovery', group: 'Банкроты и неплатежеспособные', title: 'Восстановление платежеспособности',
      short: 'Восстановление платежеспособности', path: '/ru/list/bankruptcy/recovery', param: 'flApplicantXin' },
    { key: 'bankrupt_judicial', group: 'Банкроты и неплатежеспособные', title: 'Судебное банкротство',
      short: 'Судебное банкротство', path: '/ru/list/bankruptcy/judicial', param: 'flApplicantXin' },
    { key: 'bankrupt_extrajudicial', group: 'Банкроты и неплатежеспособные', title: 'Внесудебное банкротство',
      short: 'Внесудебное банкротство', path: '/ru/list/bankruptcy-and-insolvent', param: 'flApplicantIin' },
    { key: 'debtor_extrajudicial', group: 'Должники', title: 'Внесудебная процедура',
      short: 'Должники: внесудебная процедура', path: '/ru/list/debtor', param: 'flApplicantIin' },
    { key: 'debtor_judicial', group: 'Должники', title: 'Судебная процедура',
      short: 'Должники: судебная процедура', path: '/ru/list/debtor/judicial', param: 'flApplicantXin' },
    { key: 'debtor_terminated', group: 'Должники', title: 'Процедура прекращена (объявления)',
      short: 'Должники: процедура прекращена', path: '/ru/list/debtor/extrajudicial/terminated-and-cancelled', param: 'flApplicantIin' },
  ];
  const DEFAULT_KEYS = ['recovery'];   /* то, что делают вручную сейчас */
  const WIDTH = 1920;                  /* ширина страницы на снимке: при меньшей таблица обрезается справа */
  const HEIGHT = 1000;
  const MAX_WIDTH = 3840;
  const T = Object.assign({
    INTERVAL: 2000,                    /* не чаще одного запроса за это время, мс (сайт допускает около одного в секунду) */
    MAX_INTERVAL: 15000,
    NAV_TIMEOUT: 45000,
    BACKOFF: [3000, 10000, 30000],     /* паузы перед повтором страницы */
    LIMIT_PAUSE: 20000,                /* сколько ждать после «Превышен лимит запросов» */
    SAVE_EVERY: 20,                    /* сохранять Excel каждые столько ИИН (режим папки) */
  }, window.__tazalauTest || {});      /* автоматические проверки укорачивают паузы; в работе не используется */
  const RETRIES = 3;                   /* попыток на одну страницу */
  const LIMIT_RETRIES = 5;
  const MAX_FAILS = 3;                 /* столько ИИН подряд с ошибкой: сайт недоступен, стоп */
  const MAX_RECORDS = 50;              /* больше записей на один ИИН: фильтр не применился */
  const SCAN_ROWS = 30;                /* в скольких верхних строках искать заголовок «ИИН/БИН» */
  const LINK_HEADER = 'tazalau_file_location';
  const RESULT_HEADER = 'tazalau_result';
  const SHOTS_DIR = 'tazalau_screenshots';
  const BACKUP_DIR = 'tazalau_backup';
  /* ========================================================= */

  const START_URL = location.origin + REGISTRIES[0].path;
  const IIN_HEADERS = ['иин/бин', 'бин/иин', 'иин', 'iin/bin', 'iin'];
  const COLORS = { 'НАЙДЕН': '#1a7f37', 'НЕ НАЙДЕН': '#1d2433', 'ОШИБКА': '#cf222e', 'НЕВЕРНЫЙ ИИН': '#b35900' };

  /* Таймер в отдельном потоке (Worker): если окно свёрнуто или закрыто другими окнами, браузер
     замедляет таймеры страницы до раза в минуту, а таймеры Worker: нет. Обычный таймер: запасной. */
  let worker = null, tickId = 0;
  const waiting = new Map();
  try {
    worker = new Worker(URL.createObjectURL(new Blob(['onmessage=e=>setTimeout(()=>postMessage(e.data),e.data[1])'],
      { type: 'text/javascript' })));
    worker.onmessage = e => { const r = waiting.get(e.data[0]); waiting.delete(e.data[0]); if (r) r(); };
    worker.onerror = () => { worker = null; waiting.clear(); };
  } catch (e) { worker = null; }
  const sleep = ms => new Promise(r => {
    setTimeout(r, ms);
    if (worker) { waiting.set(++tickId, r); worker.postMessage([tickId, ms]); }
  });
  const clean = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const pad2 = n => String(n).padStart(2, '0');
  const stampOf = d => d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) + '_' +
    pad2(d.getHours()) + '-' + pad2(d.getMinutes()) + '-' + pad2(d.getSeconds());
  const ruDate = (d, sec) => pad2(d.getDate()) + '.' + pad2(d.getMonth() + 1) + '.' + d.getFullYear() + ' ' +
    pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + (sec ? ':' + pad2(d.getSeconds()) : '');

  /* ---------- ИИН ---------- */
  const W1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const W2 = [3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2];
  function checkDigit(s11) {
    const d = Array.from(s11, Number);
    let k = W1.reduce((a, w, i) => a + w * d[i], 0) % 11;
    if (k === 10) {
      k = W2.reduce((a, w, i) => a + w * d[i], 0) % 11;
      if (k === 10) return null;
    }
    return k;
  }
  const iinOk = s => /^[0-9]{12}$/.test(s) && checkDigit(s.slice(0, 11)) === Number(s[11]);

  /* Значение ячейки -> { iin, padded } | { err } | null (пусто).
     Короткое ЧИСЛО: Excel потерял ведущие нули, дополняем. Короткий ТЕКСТ: опечатка.
     Номер с неверной контрольной цифрой на сайт не отправляется: иначе вышел бы снимок «нет записей»
     по чужому или несуществующему ИИН. */
  function normIin(raw, isNumber) {
    let s = String(raw == null ? '' : raw).replace(/[\s\u00a0\u1680\u2000-\u200f\u2028\u2029\u202f\u205f\u2060\u3000\ufeff]+/g, '');
    s = s.replace(/^['’`]+/, '');
    if (!s) return null;
    if (isNumber) {
      const n = Number(s);
      if (!Number.isFinite(n)) return { err: 'в ячейке не число' };
      if (n < 0 || !Number.isInteger(n) || n > 999999999999999) return { err: 'дробное или отрицательное число вместо ИИН' };
      s = String(n);
    } else if (/^[0-9]+[.,]0+$/.test(s)) s = s.split(/[.,]/)[0];
    if (!/^[0-9]+$/.test(s)) return { err: 'в ИИН есть символы, кроме цифр' };
    let padded = false;
    if (s.length !== 12) {
      if (isNumber && s.length >= 9 && s.length < 12) { s = s.padStart(12, '0'); padded = true; }
      else return { err: 'в ИИН ' + s.length + ' цифр вместо 12' };
    }
    if (!iinOk(s)) {
      return { err: padded ? 'в ИИН меньше 12 цифр и контрольная цифра не сходится'
                           : 'контрольная цифра ИИН не сходится (вероятно, опечатка)' };
    }
    return { iin: s, padded };
  }

  /* ---------- zip: чтение и запись без библиотек ---------- */
  const td = new TextDecoder(), te = new TextEncoder();

  async function pump(stream, u8, tolerant) {
    const w = stream.writable.getWriter();
    w.write(u8).catch(() => {});
    w.close().catch(() => {});
    const r = stream.readable.getReader(), parts = [];
    let len = 0;
    try {
      for (;;) {
        const { done, value } = await r.read();
        if (done) break;
        parts.push(value);
        len += value.length;
      }
    } catch (e) { if (!tolerant) throw e; }   /* мусор после конца данных: распакованное до него годится */
    const out = new Uint8Array(len);
    let p = 0;
    for (const x of parts) { out.set(x, p); p += x.length; }
    return out;
  }
  function stream(kind, format) {
    try { return kind === 'in' ? new DecompressionStream(format) : new CompressionStream(format); }
    catch (e) { throw new Error('браузер устарел: нужен Edge или Chrome версии 103 или новее'); }
  }

  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(u8) {
    let c = 0xffffffff;
    for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  function zipRead(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let end = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { end = i; break; }
    }
    if (end < 0) throw new Error('файл не похож на .xlsx');
    const total = dv.getUint16(end + 10, true);
    let p = dv.getUint32(end + 16, true);
    if (total === 0xffff || p === 0xffffffff) throw new Error('файл слишком большой для помощника');
    const entries = [];
    for (let k = 0; k < total; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('файл повреждён');
      const flags = dv.getUint16(p + 8, true), nlen = dv.getUint16(p + 28, true);
      if (flags & 1) throw new Error('файл защищён паролем: снимите пароль в Excel');
      const off = dv.getUint32(p + 42, true), csize = dv.getUint32(p + 20, true);
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      const nameBytes = buf.slice(p + 46, p + 46 + nlen);
      entries.push({ name: td.decode(nameBytes), nameBytes, flags, method: dv.getUint16(p + 10, true),
        time: dv.getUint16(p + 12, true), date: dv.getUint16(p + 14, true), crc: dv.getUint32(p + 16, true),
        csize, usize: dv.getUint32(p + 24, true), data: buf.subarray(start, start + csize) });
      p += 46 + nlen + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
    }
    return entries;
  }
  async function entryBytes(e) {
    if (e.method === 0) return e.data;
    if (e.method !== 8) throw new Error('неизвестное сжатие внутри .xlsx');
    return pump(stream('in', 'deflate-raw'), e.data, true);
  }
  /* Новая или изменённая часть файла: сжимаем заново */
  async function makeEntry(name, u8, like) {
    const now = new Date();
    const data = await pump(stream('out', 'deflate-raw'), u8, false);
    return { name, nameBytes: te.encode(name), flags: /[^\x00-\x7f]/.test(name) ? 0x0800 : 0, method: 8,
      time: like ? like.time : (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1),
      date: like ? like.date : ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate(),
      crc: crc32(u8), csize: data.length, usize: u8.length, data };
  }
  /* Части, которые не менялись, переносятся в новый файл побайтно, без распаковки: всё, чего помощник
     не касается (формулы, оформление, сводные таблицы, макросы), остаётся в точности как было. */
  function zipWrite(entries) {
    let size = 22;
    for (const e of entries) size += 30 + e.nameBytes.length + e.data.length + 46 + e.nameBytes.length;
    const out = new Uint8Array(size), dv = new DataView(out.buffer);
    let p = 0;
    const offsets = [];
    for (const e of entries) {
      offsets.push(p);
      dv.setUint32(p, 0x04034b50, true); dv.setUint16(p + 4, 20, true); dv.setUint16(p + 6, e.flags & ~8, true);
      dv.setUint16(p + 8, e.method, true); dv.setUint16(p + 10, e.time, true); dv.setUint16(p + 12, e.date, true);
      dv.setUint32(p + 14, e.crc, true); dv.setUint32(p + 18, e.csize, true); dv.setUint32(p + 22, e.usize, true);
      dv.setUint16(p + 26, e.nameBytes.length, true); dv.setUint16(p + 28, 0, true);
      out.set(e.nameBytes, p + 30);
      out.set(e.data, p + 30 + e.nameBytes.length);
      p += 30 + e.nameBytes.length + e.data.length;
    }
    const cdStart = p;
    entries.forEach((e, i) => {
      dv.setUint32(p, 0x02014b50, true); dv.setUint16(p + 4, 20, true); dv.setUint16(p + 6, 20, true);
      dv.setUint16(p + 8, e.flags & ~8, true); dv.setUint16(p + 10, e.method, true);
      dv.setUint16(p + 12, e.time, true); dv.setUint16(p + 14, e.date, true);
      dv.setUint32(p + 16, e.crc, true); dv.setUint32(p + 20, e.csize, true); dv.setUint32(p + 24, e.usize, true);
      dv.setUint16(p + 28, e.nameBytes.length, true);
      dv.setUint32(p + 42, offsets[i], true);
      out.set(e.nameBytes, p + 46);
      p += 46 + e.nameBytes.length;
    });
    dv.setUint32(p, 0x06054b50, true);
    dv.setUint16(p + 8, entries.length, true); dv.setUint16(p + 10, entries.length, true);
    dv.setUint32(p + 12, p - cdStart, true); dv.setUint32(p + 16, cdStart, true);
    return out;
  }

  /* ---------- книга Excel ---------- */
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const HYPERLINK_TYPE = RNS + '/hyperlink';
  /* Что в листе обязано идти после <hyperlinks> (порядок частей листа задан форматом) */
  const AFTER_LINKS = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
    'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF',
    'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

  function parseXml(text, name) {
    const d = new DOMParser().parseFromString(text, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('файл повреждён (' + name + ')');
    return d;
  }
  /* Ровно одна строка-заголовок XML: браузер иногда добавляет свою, а две подряд Excel не примет */
  const xmlOut = d => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    new XMLSerializer().serializeToString(d).replace(/^\s*<\?xml[^>]*\?>\s*/, '');
  const all = (n, tag) => Array.from(n.getElementsByTagNameNS('*', tag));
  const kids = (n, tag) => Array.from(n.children).filter(c => c.localName === tag);
  const colNum = ref => {
    let n = 0;
    for (const ch of ref.replace(/[0-9]+$/, '').toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
  };
  const colName = n => {
    let s = '';
    for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
    return s;
  };
  const normHead = v => String(v == null ? '' : v).replace(/[\s\u00a0\u2000-\u200f\u202f\u2060\ufeff]+/g, '').toLowerCase().replace(/\\/g, '/');
  const rangeOf = ref => {
    const m = /^\$?([A-Z]+)\$?([0-9]+)(?::\$?([A-Z]+)\$?([0-9]+))?$/i.exec(ref || '');
    if (!m) return null;
    const c1 = colNum(m[1]), r1 = Number(m[2]), c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
    return { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) };
  };
  const escTarget = s => s.replace(/%/g, '%25').replace(/#/g, '%23').replace(/ /g, '%20');

  async function openBook(bytes, fileName) {
    if (/\.xlsb$/i.test(fileName)) throw new Error('формат .xlsb не подходит: сохраните файл как .xlsx');
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf) {
      throw new Error('файл защищён паролем или это старый .xls: сохраните его в Excel как .xlsx без пароля');
    }
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('это не файл Excel (.xlsx)');
    const entries = zipRead(bytes);
    const byName = new Map(entries.map(e => [e.name.toLowerCase(), e]));
    const book = { fileName, entries, byName, changed: new Map(), added: new Map(), docs: new Map() };
    book.text = async name => {
      const e = byName.get(name.toLowerCase());
      return e ? td.decode(await entryBytes(e)) : null;
    };
    const wbText = await book.text('xl/workbook.xml'), relText = await book.text('xl/_rels/workbook.xml.rels');
    if (!wbText || !relText) throw new Error('файл не похож на .xlsx');
    const wb = parseXml(wbText, 'workbook'), rels = parseXml(relText, 'rels');
    const relList = all(rels, 'Relationship');
    const path = t => (t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, ''));
    const byType = re => relList.find(r => re.test(r.getAttribute('Type') || ''));
    const ssRel = byType(/\/sharedStrings$/), stRel = byType(/\/styles$/);
    const ssText = ssRel ? await book.text(path(ssRel.getAttribute('Target'))) : null;
    book.strings = ssText ? all(parseXml(ssText, 'sharedStrings'), 'si').map(si => all(si, 't')
      .filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join('')) : [];
    book.stylesPath = stRel ? path(stRel.getAttribute('Target')) : null;
    const sheetEls = all(wb, 'sheet');
    book.sheets = sheetEls.map((s, index) => {
      const rel = relList.find(r => r.getAttribute('Id') === (s.getAttributeNS(RNS, 'id') || s.getAttribute('r:id')));
      return rel && /\/worksheet$/.test(rel.getAttribute('Type') || '')
        ? { name: s.getAttribute('name'), path: path(rel.getAttribute('Target')), index } : null;
    }).filter(Boolean);
    const view = all(wb, 'workbookView')[0];
    book.active = view ? Number(view.getAttribute('activeTab') || 0) : 0;
    return book;
  }

  /* Лист: разбор один раз, дальше правки прямо в его XML */
  async function loadSheet(book, sheet) {
    const text = await book.text(sheet.path);
    if (text == null) throw new Error('в файле нет листа «' + sheet.name + '»');
    const doc = parseXml(text, sheet.name);
    const data = all(doc, 'sheetData')[0];
    if (!data) throw new Error('лист «' + sheet.name + '» пуст');
    const S = { book, sheet, doc, data, rows: new Map(), cells: new Map(), dirty: false };
    for (const row of kids(data, 'row')) {
      const r = Number(row.getAttribute('r'));
      if (!r) throw new Error('непривычный формат листа «' + sheet.name + '»: откройте файл в Excel и сохраните заново');
      S.rows.set(r, row);
      const map = new Map();
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r');
        if (!ref) throw new Error('непривычный формат листа «' + sheet.name + '»: откройте файл в Excel и сохраните заново');
        map.set(colNum(ref), c);
      }
      S.cells.set(r, map);
    }
    S.merges = all(doc, 'mergeCell').map(m => rangeOf(m.getAttribute('ref'))).filter(Boolean);
    S.linkRanges = [];
    const hl = kids(doc.documentElement, 'hyperlinks')[0];
    if (hl) for (const h of kids(hl, 'hyperlink')) { const g = rangeOf(h.getAttribute('ref')); if (g) S.linkRanges.push(g); }
    S.layout = findLayout(S);
    return S;
  }

  function cellValue(S, c) {
    if (!c) return null;
    const t = c.getAttribute('t');
    if (t === 'inlineStr') return { v: all(c, 't').filter(x => x.parentNode.localName !== 'rPh').map(x => x.textContent).join(''), num: false };
    const v = kids(c, 'v')[0];
    if (!v) return kids(c, 'f').length ? { formula: true } : null;
    if (t === 's') return { v: S.book.strings[Number(v.textContent)] || '', num: false };
    if (t === 'e') return { v: '', num: false };
    return { v: v.textContent, num: !t || t === 'n' };
  }
  const textAt = (S, r, c) => { const x = cellValue(S, (S.cells.get(r) || new Map()).get(c)); return x && x.v != null ? x.v : ''; };

  function findLayout(S) {
    let best = null;
    for (const [r, map] of S.cells) {
      if (r > SCAN_ROWS) continue;
      for (const [c, el] of map) {
        const x = cellValue(S, el);
        if (!x || x.v == null) continue;
        const rank = IIN_HEADERS.indexOf(normHead(x.v));
        if (rank < 0) continue;
        if (!best || rank < best.rank || (rank === best.rank && (r < best.row || (r === best.row && c < best.col)))) best = { rank, row: r, col: c };
      }
    }
    if (!best) return null;
    const L = { headerRow: best.row, iinCol: best.col, linkCol: 0, resultCol: 0, lastRow: best.row };
    for (const [c, el] of S.cells.get(best.row)) {
      const x = cellValue(S, el), name = x && x.v != null ? normHead(x.v) : '';
      if (name === LINK_HEADER && (!L.linkCol || c < L.linkCol)) L.linkCol = c;
      if (name === RESULT_HEADER && (!L.resultCol || c < L.resultCol)) L.resultCol = c;
    }
    for (const [r, map] of S.cells) {
      if (r <= best.row) continue;
      const x = cellValue(S, map.get(best.col));
      if (x && (x.formula || clean(x.v))) L.lastRow = Math.max(L.lastRow, r);
    }
    return L;
  }

  function readTasks(S) {
    const L = S.layout, tasks = [];
    const linked = (r, c) => !!clean(textAt(S, r, c)) || S.linkRanges.some(g => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
    for (let r = L.headerRow + 1; r <= L.lastRow; r++) {
      const x = cellValue(S, (S.cells.get(r) || new Map()).get(L.iinCol));
      if (!x) continue;
      const parse = x.formula
        ? { err: 'ИИН задан формулой без сохранённого значения: откройте файл в Excel, сохраните и запустите снова' }
        : normIin(x.v, x.num);
      if (!parse) continue;
      const row = S.rows.get(r);
      tasks.push({ row: r, parse, raw: x.formula ? '' : x.v, hasLink: !!L.linkCol && linked(r, L.linkCol),
        hidden: !!row && /^(1|true)$/.test(row.getAttribute('hidden') || '') });
    }
    return tasks;
  }

  function rowEl(S, r) {
    let row = S.rows.get(r);
    if (row) return row;
    row = S.doc.createElementNS(NS, 'row');
    row.setAttribute('r', r);
    let before = null;
    for (const [k, el] of S.rows) if (k > r && (!before || k < before.k)) before = { k, el };
    S.data.insertBefore(row, before ? before.el : null);
    S.rows.set(r, row);
    S.cells.set(r, new Map());
    return row;
  }
  function cellEl(S, r, col) {
    const row = rowEl(S, r), map = S.cells.get(r);
    let c = map.get(col);
    if (c) return c;
    c = S.doc.createElementNS(NS, 'c');
    c.setAttribute('r', colName(col) + r);
    let before = null;
    for (const [k, el] of map) if (k > col && (!before || k < before.k)) before = { k, el };
    row.insertBefore(c, before ? before.el : null);
    row.removeAttribute('spans');
    map.set(col, c);
    return c;
  }
  const mergedInto = (S, r, col) => S.merges.some(g => r >= g.r1 && r <= g.r2 && col >= g.c1 && col <= g.c2 && !(r === g.r1 && col === g.c1));

  function setText(S, r, col, text, style) {
    if (mergedInto(S, r, col)) throw new Error('строка ' + r + ': ячейка ' + colName(col) + r + ' объединена с соседними');
    const c = cellEl(S, r, col);
    while (c.firstChild) c.firstChild.remove();
    c.setAttribute('t', 'inlineStr');
    if (style != null) c.setAttribute('s', style);
    const is = S.doc.createElementNS(NS, 'is'), t = S.doc.createElementNS(NS, 't');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = text;
    is.appendChild(t);
    c.appendChild(is);
    S.dirty = true;
    return c;
  }

  /* Связи листа (там хранятся адреса гиперссылок) */
  async function sheetRels(S) {
    if (S.rels) return S.rels;
    const i = S.sheet.path.lastIndexOf('/');
    S.relsPath = S.sheet.path.slice(0, i) + '/_rels/' + S.sheet.path.slice(i + 1) + '.rels';
    const text = await S.book.text(S.relsPath);
    S.relsNew = text == null;
    S.rels = parseXml(text || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + PNS + '"/>', 'rels');
    S.relNo = all(S.rels, 'Relationship').reduce((m, r) => Math.max(m, Number((/([0-9]+)$/.exec(r.getAttribute('Id') || '') || [0, 0])[1])), 0);
    return S.rels;
  }

  async function setLink(S, r, col, target) {
    const rels = await sheetRels(S), ref = colName(col) + r, root = S.doc.documentElement;
    let hl = kids(root, 'hyperlinks')[0];
    if (!hl) {
      hl = S.doc.createElementNS(NS, 'hyperlinks');
      root.insertBefore(hl, Array.from(root.children).find(c => AFTER_LINKS.includes(c.localName)) || null);
    }
    for (const h of kids(hl, 'hyperlink')) {
      if (h.getAttribute('ref') !== ref) continue;
      const old = h.getAttributeNS(RNS, 'id');
      h.remove();
      if (old && !kids(hl, 'hyperlink').some(x => x.getAttributeNS(RNS, 'id') === old)) {
        all(rels, 'Relationship').filter(x => x.getAttribute('Id') === old && x.getAttribute('Type') === HYPERLINK_TYPE).forEach(x => x.remove());
      }
    }
    const id = 'rId' + (++S.relNo);
    const h = S.doc.createElementNS(NS, 'hyperlink');
    h.setAttribute('ref', ref);
    h.setAttributeNS(RNS, 'r:id', id);
    hl.appendChild(h);
    const rel = rels.createElementNS(PNS, 'Relationship');
    rel.setAttribute('Id', id);
    rel.setAttribute('Type', HYPERLINK_TYPE);
    rel.setAttribute('Target', escTarget(target));
    rel.setAttribute('TargetMode', 'External');
    rels.documentElement.appendChild(rel);
    S.linkRanges.push({ c1: col, c2: col, r1: r, r2: r });
    S.dirty = true;
  }

  /* Вид ссылки (синяя, подчёркнутая) поверх прежнего оформления ячейки: заливка и рамки остаются */
  async function linkStyle(S, s) {
    const book = S.book;
    if (book.styles === undefined) {
      const text = book.stylesPath ? await book.text(book.stylesPath) : null;
      book.styles = text ? parseXml(text, 'styles') : null;
      book.linkXf = new Map();
    }
    const st = book.styles, key = s == null ? '' : String(s);
    if (!st) return s;
    if (book.linkXf.has(key)) return book.linkXf.get(key);
    const fonts = kids(st.documentElement, 'fonts')[0], xfs = kids(st.documentElement, 'cellXfs')[0];
    const xfList = xfs ? kids(xfs, 'xf') : [], fontList = fonts ? kids(fonts, 'font') : [];
    const xf = xfList[Number(s) || 0] || xfList[0];
    const baseFont = xf && (fontList[Number(xf.getAttribute('fontId')) || 0] || fontList[0]);
    if (!xf || !baseFont) { book.linkXf.set(key, s); return s; }
    const font = baseFont.cloneNode(true);
    kids(font, 'u').concat(kids(font, 'color')).forEach(e => e.remove());
    const u = st.createElementNS(NS, 'u'), color = st.createElementNS(NS, 'color'), sz = kids(font, 'sz')[0];
    color.setAttribute('rgb', 'FF0563C1');
    font.insertBefore(u, sz || font.firstChild);
    font.insertBefore(color, sz ? sz.nextSibling : null);
    fonts.appendChild(font);
    fonts.setAttribute('count', fontList.length + 1);
    const nxf = xf.cloneNode(true);
    nxf.setAttribute('fontId', fontList.length);
    nxf.setAttribute('applyFont', '1');
    xfs.appendChild(nxf);
    xfs.setAttribute('count', xfList.length + 1);
    book.stylesDirty = true;
    book.linkXf.set(key, String(xfList.length));
    return String(xfList.length);
  }

  /* Недостающие столбцы программы: сразу за таблицей (объединённые ячейки шапки тоже считаются шапкой) */
  function ensureColumns(S) {
    const L = S.layout;
    const used = new Set();
    let last = L.iinCol;
    for (const [r, map] of S.cells) for (const [c, el] of map) {
      const x = cellValue(S, el);
      if (!x || !(x.formula || clean(x.v))) continue;
      used.add(c);
      if (r === L.headerRow) last = Math.max(last, c);
    }
    for (const g of S.merges) if (g.r1 <= L.headerRow && g.r2 >= L.headerRow) last = Math.max(last, g.c2);
    const headStyle = ((S.cells.get(L.headerRow) || new Map()).get(L.iinCol) || { getAttribute: () => null }).getAttribute('s');
    const add = (name, width) => {
      let col = Math.max(last, L.linkCol || 0, L.resultCol || 0) + 1;
      while (used.has(col)) col++;
      setText(S, L.headerRow, col, name, headStyle);
      used.add(col);
      setWidth(S, col, width);
      return col;
    };
    if (!L.linkCol) L.linkCol = add(LINK_HEADER, 60);
    if (!L.resultCol) L.resultCol = add(RESULT_HEADER, 46);
  }
  function setWidth(S, col, width) {
    const root = S.doc.documentElement;
    let cols = kids(root, 'cols')[0];
    const list = cols ? kids(cols, 'col') : [];
    if (list.some(c => Number(c.getAttribute('min')) <= col && Number(c.getAttribute('max')) >= col)) return;
    if (!cols) { cols = S.doc.createElementNS(NS, 'cols'); root.insertBefore(cols, S.data); }
    const el = S.doc.createElementNS(NS, 'col');
    el.setAttribute('min', col); el.setAttribute('max', col); el.setAttribute('width', width); el.setAttribute('customWidth', '1');
    cols.insertBefore(el, list.find(c => Number(c.getAttribute('min')) > col) || null);
  }
  function fixDimension(S) {
    const dim = kids(S.doc.documentElement, 'dimension')[0], g = dim && rangeOf(dim.getAttribute('ref'));
    if (!g) return;
    let r2 = g.r2, c2 = g.c2;
    for (const [r, map] of S.cells) { r2 = Math.max(r2, r); for (const c of map.keys()) c2 = Math.max(c2, c); }
    dim.setAttribute('ref', colName(g.c1) + g.r1 + ':' + colName(c2) + r2);
  }

  async function bookBytes(S) {
    const book = S.book;
    fixDimension(S);
    book.changed.set(S.sheet.path.toLowerCase(), xmlOut(S.doc));
    if (S.rels) (S.relsNew ? book.added : book.changed).set(S.relsNew ? S.relsPath : S.relsPath.toLowerCase(), xmlOut(S.rels));
    if (book.stylesDirty) book.changed.set(book.stylesPath.toLowerCase(), xmlOut(book.styles));
    const out = [];
    for (const e of book.entries) {
      const text = book.changed.get(e.name.toLowerCase());
      out.push(text == null ? e : await makeEntry(e.name, te.encode(text), e));
    }
    for (const [name, text] of book.added) out.push(await makeEntry(name, te.encode(text), null));
    return zipWrite(out);
  }

  /* ---------- окно помощника ---------- */
  function h(tag, css, text) {
    const e = document.createElement(tag);
    if (css) e.style.cssText = css;
    if (text != null) e.textContent = text;
    return e;
  }
  function button(text, onClick, primary, tip) {
    const b = h('button', 'margin:0 4px 4px 0;padding:5px 9px;border-radius:5px;cursor:pointer;font:inherit;' +
      (primary ? 'background:#1f6feb;color:#fff;border:1px solid #1f6feb;'
               : 'background:#fff;color:#1d2433;border:1px solid #b9c3d3;'), text);
    b.type = 'button';
    if (tip) b.title = tip;
    b.addEventListener('click', onClick);
    return b;
  }

  /* Узкая полоса слева: файл, запуск, статус и журнал; справа: сайт (уменьшенный, чтобы помещался целиком) */
  const root = h('div', 'position:fixed;inset:0;z-index:2147483647;display:flex;background:#fff;' +
    'font:13px/1.4 Arial,sans-serif;color:#1d2433;text-align:left');
  const panel = h('div', 'width:250px;min-width:250px;height:100%;overflow:auto;box-sizing:border-box;' +
    'padding:8px 10px;background:#f5f7fa;border-right:1px solid #d5dbe5');
  const frameBox = h('div', 'flex:1;height:100%;overflow:hidden;position:relative;background:#e9eef3');
  const frame = h('iframe', 'position:absolute;left:0;top:0;border:0;background:#fff;transform-origin:0 0');
  frame.title = 'Tazalau';
  frameBox.append(frame);
  root.append(panel, frameBox);
  let frameW = WIDTH;
  function layoutFrame() {
    const k = Math.min(1, (frameBox.clientWidth || WIDTH) / frameW);
    frame.style.width = frameW + 'px';
    frame.style.height = HEIGHT + 'px';            /* одна и та же высота на любом экране: снимки одинаковые */
    frame.style.transform = 'scale(' + k + ')';
  }

  const row = (...kidsOf) => { const r = h('div', 'margin:0 0 4px'); r.append(...kidsOf); return r; };
  const head = h('div', 'display:flex;justify-content:space-between;align-items:center;margin-bottom:6px');
  head.append(h('b', '', 'Tazalau v' + VERSION), button('✕', close, false, 'Закрыть помощник'));

  const banner = h('div', 'display:none;position:sticky;top:0;z-index:1;margin:0 0 8px;padding:8px;' +
    'border-radius:6px;background:#fff4d6;border:1px solid #e3b341');
  const bannerText = h('div', 'margin-bottom:6px;font-weight:bold;white-space:pre-line');
  const bannerBtns = h('div');
  banner.append(bannerText, bannerBtns);

  const canPickFolder = typeof window.showDirectoryPicker === 'function';
  const folderB = button('📁 Папка с Excel', pickFolder, true,
    'Папка, где лежит файл Excel. Снимки сохранятся в ней же, ссылки запишутся прямо в файл.');
  const fileSelect = h('select', 'display:none;width:100%;margin:0 0 4px;font-size:12px');
  fileSelect.addEventListener('change', () => openFromFolder(fileSelect.value));
  const fileInput = h('input', 'display:none;width:100%;font-size:12px;margin:4px 0');
  fileInput.type = 'file';
  fileInput.accept = '.xlsx,.xlsm';
  fileInput.title = 'Файл Excel со столбцом «ИИН/БИН». Снимки и копия файла со ссылками сохранятся в «Загрузки».';
  fileInput.addEventListener('change', async () => {
    const f = fileInput.files[0];
    if (f) await openBytes(new Uint8Array(await f.arrayBuffer()), f.name, null);
  });
  const fileLink = h('a', 'color:#1f6feb;cursor:pointer;font-size:12px', canPickFolder ? 'или один файл (результат в «Загрузки»)' : '');
  fileLink.addEventListener('click', () => { fileInput.style.display = fileInput.style.display === 'none' ? 'block' : 'none'; });
  if (!canPickFolder) { folderB.style.display = 'none'; fileInput.style.display = 'block'; }
  const sheetSelect = h('select', 'display:none;width:100%;margin:0 0 4px;font-size:12px');
  sheetSelect.title = 'Лист со столбцом «ИИН/БИН»';
  sheetSelect.addEventListener('change', () => useSheet(sheetSelect.value));

  const summary = h('div', 'margin:6px 0;white-space:pre-line;font-size:12px');
  const badBox = h('details', 'display:none;margin:0 0 6px;color:#9a3412;font-size:12px');
  const badHead = h('summary', 'cursor:pointer');
  const badList = h('div', 'white-space:pre-line;max-height:140px;overflow:auto');
  badBox.append(badHead, badList);

  const regBox = h('details', 'margin:0 0 6px;font-size:12px');
  const regHead = h('summary', 'cursor:pointer');
  regBox.append(regHead);
  const regChecks = new Map();
  let lastGroup = '';
  for (const g of REGISTRIES) {
    if (g.group !== lastGroup) { regBox.append(h('div', 'margin:4px 0 2px;color:#5b6577', g.group)); lastGroup = g.group; }
    const label = h('label', 'display:block;cursor:pointer');
    const box = h('input', 'margin:0 5px 0 0;vertical-align:middle');
    box.type = 'checkbox';
    box.checked = DEFAULT_KEYS.includes(g.key);
    box.addEventListener('change', refresh);
    label.append(box, document.createTextNode(g.title));
    regBox.append(label);
    regChecks.set(g.key, box);
  }
  const visibleLabel = h('label', 'display:block;cursor:pointer;margin:6px 0 0');
  const visibleBox = h('input', 'margin:0 5px 0 0;vertical-align:middle');
  visibleBox.type = 'checkbox';
  visibleBox.addEventListener('change', refresh);
  visibleLabel.append(visibleBox, document.createTextNode('только видимые строки (фильтр Excel)'));
  regBox.append(visibleLabel);

  const startB = button('▶ Старт', run, true);
  const stopB = button('⏹ Стоп', stopRun, false, 'Остановить после текущего ИИН и сохранить');
  const state = h('div', 'margin:2px 0 6px;padding:6px;background:#eef3fb;border-radius:5px;white-space:pre-line;' +
    'font-size:12px', canPickFolder ? 'Выберите папку с файлом Excel.' : 'Выберите файл Excel.');
  state.id = 'tazalau-state';
  const csvB = button('CSV', () => saveBlob(csvBlob(), 'журнал_tazalau_' + stampOf(new Date()) + '.csv'), false, 'Скачать журнал (CSV)');
  const xlsxB = button('Excel ↓', async () => { if (S) saveBlob(new Blob([await bookBytes(S)]), resultName()); }, false,
    'Скачать копию файла Excel со ссылками, как есть сейчас');
  const jbox = h('details', 'font-size:12px');
  jbox.open = true;
  const jhead = h('summary', 'cursor:pointer');
  const table = h('div', 'max-height:45vh;overflow:auto');
  jbox.append(jhead, table);

  panel.append(head, banner, row(folderB), fileSelect, row(fileLink), fileInput, sheetSelect, summary, badBox, regBox,
    row(startB, stopB), state, row(csvB, xlsxB), jbox);
  document.body.appendChild(root);
  const overflowBefore = document.documentElement.style.overflow;
  document.documentElement.style.overflow = 'hidden';
  window.addEventListener('resize', layoutFrame);
  window.addEventListener('beforeunload', e => { if (running) { e.preventDefault(); e.returnValue = ''; } });

  /* ---------- состояние ---------- */
  let book = null, S = null, tasks = [], store = { dir: null, shots: null, handle: null };
  let running = false, stopping = false, interval = T.INTERVAL, lastRequest = 0;
  const journal = [];

  function say(msg) { state.textContent = msg; }
  const chosen = () => REGISTRIES.filter(g => regChecks.get(g.key).checked);

  function plan(recheck) {
    const groups = new Map(), invalid = [];
    let linked = 0, hiddenRows = 0;
    for (const t of tasks) {
      if (visibleBox.checked && t.hidden) { hiddenRows++; continue; }
      if (t.hasLink && !recheck) { linked++; continue; }
      if (t.parse.err) { invalid.push(t); continue; }
      if (!groups.has(t.parse.iin)) groups.set(t.parse.iin, []);
      groups.get(t.parse.iin).push(t.row);
    }
    return { groups, invalid, linked, hiddenRows };
  }

  function refresh() {
    const regs = chosen();
    regHead.textContent = 'Реестры: ' + (regs.length === 1 ? regs[0].short : regs.length ? regs.length + ' выбрано' : 'не выбрано');
    if (!S) { summary.textContent = ''; badBox.style.display = 'none'; updateButtons(); return; }
    const p = plan(false), rows = Array.from(p.groups.values()).reduce((a, x) => a + x.length, 0);
    const skipped = [p.linked ? 'уже со ссылкой ' + p.linked : '', p.hiddenRows ? 'скрыто фильтром ' + p.hiddenRows : ''].filter(Boolean).join(', ');
    summary.textContent = book.fileName + ' · лист «' + S.sheet.name + '»\n' +
      'Строк с ИИН ' + tasks.length + (skipped ? ' · ' + skipped : '') + '\n' +
      (p.groups.size ? 'К проверке: ' + p.groups.size + ' ИИН (строк ' + rows + '), примерно ' + eta(p.groups.size * Math.max(1, regs.length))
                     : (tasks.length ? 'Новых ИИН нет ✔ (Старт спросит, проверять ли заново)' : 'В столбце «ИИН/БИН» нет значений'));
    badBox.style.display = p.invalid.length ? 'block' : 'none';
    badHead.textContent = '⚠ Строк с неверным ИИН: ' + p.invalid.length;
    badList.textContent = p.invalid.map(t => 'строка ' + t.row + ': ' + t.parse.err).join('\n');
    jhead.textContent = 'Журнал (' + journal.length + ')';
    updateButtons();
  }
  function eta(pages) {
    const sec = Math.round(pages * Math.max(1000, interval) / 1000);
    return sec < 60 ? 'меньше минуты' : sec < 5400 ? Math.round(sec / 60) + ' мин' : Math.floor(sec / 3600) + ' ч ' + pad2(Math.round(sec % 3600 / 60)) + ' мин';
  }
  function updateButtons() {
    startB.disabled = running || !S;
    stopB.disabled = !running;
    folderB.disabled = fileInput.disabled = fileSelect.disabled = sheetSelect.disabled = running;
    for (const b of regChecks.values()) b.disabled = running;
    visibleBox.disabled = running;
    xlsxB.style.display = S ? '' : 'none';
  }

  function logRow(r) {
    journal.push(r);
    const line = h('div', 'padding:2px 0;border-bottom:1px solid #eef0f4');
    line.append(h('span', 'font-family:monospace', r.iin + ' '), h('b', 'color:' + (COLORS[r.status] || '#1d2433'), r.status));
    if (r.note) { line.title = r.note; line.append(h('div', 'color:#5b6577;font-size:11px', r.note.slice(0, 110))); }
    table.prepend(line);
    jhead.textContent = 'Журнал (' + journal.length + ')';
  }
  function csvBlob() {
    const safe = v => { const s = clean(v).replace(/[;"]/g, ' '); return /^[=+\-@]/.test(s) ? "'" + s : s; };
    const lines = ['время;ИИН;строки;итог;записей;файл;примечание'].concat(journal.map(r =>
      [safe(r.time), '="' + r.iin.replace(/\D/g, '') + '"', safe(r.rows), safe(r.status), safe(r.counts), safe(r.file), safe(r.note)].join(';')));
    return new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  }

  function beep() {
    try {
      const a = new AudioContext(), o = a.createOscillator();
      o.frequency.value = 880;
      o.connect(a.destination);
      o.start();
      setTimeout(() => { o.stop(); a.close(); }, 250);
    } catch (e) { /* без звука */ }
  }
  let pendingAsk = null;
  function ask(msg, labels) {
    const titleBefore = document.title;
    return new Promise(resolve => {
      bannerText.textContent = msg;
      bannerBtns.textContent = '';
      pendingAsk = l => {
        pendingAsk = null;
        banner.style.display = 'none';
        document.title = titleBefore;
        resolve(l);
      };
      labels.forEach((l, i) => bannerBtns.append(button(l, () => pendingAsk && pendingAsk(l), i === 0)));
      banner.style.display = 'block';
      panel.scrollTop = 0;
      document.title = '⚠ Tazalau: нужно действие';
      beep();
    });
  }

  function close() {
    if (running && !confirm('Идёт проверка. Остановить и закрыть помощник?')) return;
    stopping = true;
    running = false;
    if (pendingAsk) pendingAsk('Отмена');
    root.remove();
    document.documentElement.style.overflow = overflowBefore;
    delete window.__tazalau;
    if (worker) { worker.terminate(); worker = null; }
  }

  /* ---------- файлы: папка (правка на месте) или «Загрузки» ---------- */
  function saveBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
  }
  async function writeFile(dir, name, data) {
    const fh = await dir.getFileHandle(name, { create: true });
    const w = await fh.createWritable();
    try { await w.write(data); await w.close(); }
    catch (e) { try { await w.abort(); } catch (x) { /* уже закрыт */ } throw e; }
  }
  const resultName = () => book.fileName.replace(/(\.[a-z]+)?$/i, m => '_tazalau' + (m || '.xlsx'));

  async function pickFolder() {
    let dir;
    try { dir = await window.showDirectoryPicker({ mode: 'readwrite', id: 'tazalau' }); }
    catch (e) { if (e && e.name !== 'AbortError') say('❌ Папку выбрать не удалось: ' + e.message + '\nНажмите «или один файл» ниже.'); return; }
    await useDir(dir);
  }
  async function useDir(dir) {
    const names = [];
    for await (const [name, handle] of dir.entries()) {
      if (handle.kind === 'file' && /\.(xlsx|xlsm)$/i.test(name) && !name.startsWith('~$')) names.push(name);
    }
    names.sort((a, b) => a.localeCompare(b, 'ru'));
    store = { dir, shots: null, handle: null };
    fileSelect.textContent = '';
    if (!names.length) { book = S = null; tasks = []; fileSelect.style.display = 'none'; refresh(); say('❌ В этой папке нет файлов Excel (.xlsx).'); return; }
    for (const n of names) { const o = h('option', '', n); o.value = n; fileSelect.append(o); }
    fileSelect.style.display = names.length > 1 ? 'block' : 'none';
    await openFromFolder(names[0]);
  }
  async function openFromFolder(name) {
    try {
      const handle = await store.dir.getFileHandle(name);
      const file = await handle.getFile();
      await openBytes(new Uint8Array(await file.arrayBuffer()), name, handle);
    } catch (e) { book = S = null; tasks = []; refresh(); say('❌ ' + (e.message || e)); }
  }

  let sheetsWithIin = [];
  async function openBytes(bytes, name, handle) {
    try {
      book = S = null; tasks = [];
      store.handle = handle;
      if (!handle) store.dir = store.shots = null;
      store.original = bytes;
      store.backedUp = false;
      const b = await openBook(bytes, name);
      sheetsWithIin = [];
      for (const sh of b.sheets) { const s = await loadSheet(b, sh); if (s.layout) sheetsWithIin.push(s); }
      if (!sheetsWithIin.length) throw new Error('в файле не найден столбец с заголовком «ИИН/БИН» (в первых ' + SCAN_ROWS + ' строках)');
      book = b;
      sheetSelect.textContent = '';
      for (const s of sheetsWithIin) { const o = h('option', '', s.sheet.name); o.value = s.sheet.name; sheetSelect.append(o); }
      sheetSelect.style.display = sheetsWithIin.length > 1 ? 'block' : 'none';
      const active = sheetsWithIin.find(s => s.sheet.index === b.active) || sheetsWithIin[0];
      sheetSelect.value = active.sheet.name;
      useSheet(active.sheet.name);
      say(handle ? 'Можно запускать. Снимки: папка ' + SHOTS_DIR + ', ссылки: прямо в файл.'
                 : 'Можно запускать. Снимки и копия файла со ссылками придут в «Загрузки».');
    } catch (e) { book = S = null; tasks = []; refresh(); say('❌ ' + (e.message || e)); }
  }
  function useSheet(name) {
    S = sheetsWithIin.find(s => s.sheet.name === name) || null;
    tasks = S ? readTasks(S) : [];
    refresh();
  }

  /* Сохранение книги. В папке: на место (перед первой записью резервная копия); иначе в «Загрузки» в конце. */
  async function saveBook(final) {
    if (!S || !S.dirty) return true;
    const bytes = await bookBytes(S);
    zipRead(bytes);                               /* самопроверка: собранный файл читается как архив */
    if (!store.handle) {
      if (final) saveBlob(new Blob([bytes]), resultName());
      return true;
    }
    for (let attempt = 0; ; attempt++) {
      try {
        if (!store.backedUp) {
          const bdir = await store.dir.getDirectoryHandle(BACKUP_DIR, { create: true });
          await writeFile(bdir, book.fileName.replace(/(\.[a-z]+)?$/i, m => '__' + stampOf(new Date()) + (m || '.xlsx')), store.original);
          store.backedUp = true;
        }
        const w = await store.handle.createWritable();
        try { await w.write(bytes); await w.close(); }
        catch (e) { try { await w.abort(); } catch (x) { /* уже закрыт */ } throw e; }
        return true;
      } catch (e) {
        if (!final) { say('⚠ Файл Excel занят (открыт в Excel?). Продолжаю, сохраню позже.'); return false; }
        const answer = await ask('Не удалось сохранить «' + book.fileName + '»: похоже, он открыт в Excel.\nЗакройте его и нажмите «Повторить».',
          ['Повторить', 'Сохранить копией']);
        if (answer !== 'Повторить' || attempt > 20) {
          const copy = book.fileName.replace(/(\.[a-z]+)?$/i, m => '_tazalau_' + stampOf(new Date()) + (m || '.xlsx'));
          try { await writeFile(store.dir, copy, bytes); say('⚠ Результат сохранён в копию: ' + copy); }
          catch (x) { saveBlob(new Blob([bytes]), copy); say('⚠ Результат скачан копией: ' + copy); }
          return true;
        }
      }
    }
  }

  /* ---------- работа со страницей сайта (рамка справа) ---------- */
  function fdoc() {
    try {
      const d = frame.contentDocument;
      return d && d.location && d.location.href !== 'about:blank' ? d : null;
    } catch (e) { return null; }
  }
  async function waitFor(fn, ms, every) {
    const end = Date.now() + ms;
    for (;;) {
      let v = null;
      try { v = fn(); } catch (e) { /* страница ещё грузится */ }
      if (v) return v;
      if (Date.now() >= end) return null;
      await sleep(every || 100);
    }
  }
  class SiteError extends Error {}
  class LimitError extends SiteError {}

  /* Готовность страницы определяется по ней самой, а не по событию «загружено»: его задерживает сторонний
     счётчик посещений, и в сети, где он недоступен, каждая страница ждала бы десятки секунд. */
  async function go(url) {
    const old = fdoc();
    try { frame.contentWindow.location.replace(url); } catch (e) { frame.src = url; }
    let broken = false;
    const d = await waitFor(() => {
      let x;
      try { x = frame.contentDocument; void frame.contentWindow.location.href; } catch (e) { broken = true; return true; }
      return x && x !== old && x.location.href !== 'about:blank' && x.readyState !== 'loading' && x.body ? x : null;
    }, T.NAV_TIMEOUT);
    if (!d || broken || d === true) throw new SiteError(d ? 'нет связи с сайтом' : 'сайт не ответил за ' + T.NAV_TIMEOUT / 1000 + ' с');
    /* Сайт сам снимает «loading» через 0,4 с после запуска своих скриптов: это и есть знак, что страница собрана.
       Если вкладка скрыта и таймеры страницы тормозят, не дождавшись, снимаем сами: иначе снимок выйдет пустым. */
    await waitFor(() => !d.body.classList.contains('loading'), 6000);
    d.body.classList.remove('loading');
    void d.body.offsetHeight;                      /* страница стала видимой: теперь она запросит нужные шрифты */
    await sleep(80);
    await waitFor(() => Array.from(d.querySelectorAll('link[rel~="stylesheet"]')).every(l => l.sheet) &&
      Array.from(d.images).every(i => i.complete), 8000);
    await waitFor(() => !d.fonts || d.fonts.status === 'loaded', 6000);
    return d;
  }

  function pageInfo(d, param) {
    const input = d.querySelector('input[name="' + param + '"]');
    const totals = Array.from(d.querySelectorAll('ul.pagination')).map(ul => {
      const span = ul.previousElementSibling;
      if (!span || span.tagName !== 'SPAN') return null;
      let text = '';
      span.childNodes.forEach(n => { if (n.nodeType === 3) text += n.textContent; });
      text = text.replace(/[\s\u00a0\u202f]/g, '');
      return /^[0-9]+$/.test(text) ? parseInt(text, 10) : null;
    });
    const tableEl = d.querySelector('.table-responsive table') || d.querySelector('table.table');
    let dataRows = 0, emptyRow = false;
    if (tableEl) tableEl.querySelectorAll('tbody > tr').forEach(tr => {
      const cells = tr.children;
      if (cells.length === 1 && cells[0].colSpan > 1) emptyRow = true; else dataRows += 1;
    });
    const wrap = tableEl ? tableEl.closest('.table-responsive') : null;
    const menu = d.querySelector('.left-side-menu'), win = d.defaultView;
    return {
      path: win.location.pathname, search: win.location.search, title: d.title,
      notice: d.body ? (d.body.innerText || '').slice(0, 300) : '',
      inputFound: !!input, inputValue: input ? input.value : null, totals, tableFound: !!tableEl, dataRows, emptyRow,
      styled: !!menu && win.getComputedStyle(menu).position === 'fixed',
      overflow: Math.max(0, wrap ? wrap.scrollWidth - wrap.clientWidth : 0, d.documentElement.scrollWidth - win.innerWidth),
    };
  }

  /* Страница принимается, только если это результат поиска именно по нашему ИИН */
  function verify(info, reg, iin) {
    const bannerOf = (info.title + ' ' + info.notice).toLowerCase();
    if (bannerOf.includes('лимит запросов') && !info.inputFound) throw new LimitError('сайт сообщил «Превышен лимит запросов»');
    const path = info.path.replace(/\/+$/, '');
    if (path !== reg.path) throw new SiteError('сайт открыл другую страницу (' + (path || 'неизвестно') + ')');
    if (!info.styled) throw new SiteError('страница загрузилась без оформления');
    if (!info.inputFound) throw new SiteError('на странице нет поля «ИИН»: возможно, сайт изменился');
    if (info.inputValue !== iin) throw new SiteError('сайт не применил фильтр по ИИН');
    if (new URLSearchParams(info.search).getAll(reg.param).join('|') !== iin) throw new SiteError('адрес страницы не соответствует запросу');
    if (!info.tableFound) throw new SiteError('на странице нет таблицы результатов: возможно, сайт изменился');
    const totals = info.totals;
    if (!totals.length || totals.some(t => t == null) || new Set(totals).size !== 1) throw new SiteError('не удалось прочитать счётчик «Всего записей»');
    const count = totals[0], rows = info.dataRows;
    if (count > MAX_RECORDS) throw new SiteError('по одному ИИН показано ' + count + ' записей: похоже, фильтр не сработал');
    if (count === 0 && (rows !== 0 || !info.emptyRow)) throw new SiteError('счётчик записей не совпадает с таблицей');
    if (count > 0 && !(rows >= 1 && rows <= count)) throw new SiteError('счётчик записей не совпадает с таблицей');
    return { count, rows };
  }

  /* ---------- снимок страницы ---------- */
  /* Расширений и прав на «снимок экрана» у закладки нет. Поэтому страница (её HTML и её же стили, со шрифтами
     и картинками внутри) вкладывается в SVG, браузер рисует его тем же движком, а результат переносится на холст.
     Сверено попиксельно с настоящим снимком экрана: заметных отличий нет. */
  const dataCache = new Map(), cssCache = new Map();
  const URL_RE = /url\(\s*(?:"([^"]*)"|'([^']*)'|([^)"'\s]*))\s*\)/g;
  function toDataUrl(url) {
    if (!dataCache.has(url)) dataCache.set(url, (async () => {
      const r = await fetch(url, { cache: 'force-cache', credentials: 'same-origin' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const blob = await r.blob();
      return new Promise((res, rej) => { const f = new FileReader(); f.onload = () => res(f.result); f.onerror = rej; f.readAsDataURL(blob); });
    })());
    return dataCache.get(url);
  }
  const fontKey = f => String(f || '').replace(/["']/g, '').trim().toLowerCase();
  /* @font-face: шрифт вкладывается, только если страница его действительно загрузила, и одним файлом */
  async function inlineFonts(text, base, fonts) {
    const blocks = [];
    text.replace(/@font-face\s*\{[^}]*\}/gi, m => { blocks.push(m); return m; });
    const map = new Map();
    await Promise.all(blocks.map(async block => {
      const family = fontKey((/font-family\s*:\s*([^;}]+)/i.exec(block) || [])[1]);
      if (!fonts.has(family)) { map.set(block, ''); return; }
      const urls = [];
      block.replace(/url\(\s*["']?([^"')]+)["']?\s*\)(?:\s*format\(\s*["']?([^"')]+)["']?\s*\))?/gi, (m, u, f) => { urls.push([u, (f || '').toLowerCase()]); return m; });
      const rank = ([u, f]) => (/woff2/.test(f) || /\.woff2(\?|#|$)/i.test(u) ? 0 : /woff/.test(f) || /\.woff(\?|#|$)/i.test(u) ? 1
        : /truetype|opentype/.test(f) || /\.(ttf|otf)(\?|#|$)/i.test(u) ? 2 : 9);
      const best = urls.filter(x => !x[0].startsWith('data:') && rank(x) < 9).sort((a, b) => rank(a) - rank(b))[0];
      if (!best) return;
      try {
        const data = await toDataUrl(new URL(best[0], base).href.split('#')[0]);
        const rest = block.slice(block.indexOf('{') + 1, block.lastIndexOf('}')).split(';').filter(d => d.trim() && !/^\s*src\s*:/i.test(d)).join(';');
        map.set(block, '@font-face{' + rest + ';src:url("' + data + '")' + (best[1] ? ' format("' + best[1] + '")' : '') + '}');
      } catch (e) { map.set(block, ''); }
    }));
    return text.replace(/@font-face\s*\{[^}]*\}/gi, m => (map.has(m) ? map.get(m) : m));
  }
  async function inlineCss(text, base, loaded, depth) {
    text = await inlineFonts(text, base, loaded.fonts);
    const imports = [];
    text = text.replace(/@import\s+(?:url\(\s*)?["']?([^"')\s;]+)["']?\s*\)?[^;]*;/g, (m, u) => { imports.push(u); return ''; });
    let headCss = '';
    if (depth < 3) for (const u of imports) {
      try {
        const abs = new URL(u, base).href;
        if (new URL(abs).origin === location.origin) headCss += await sheetText(abs, loaded, depth + 1) + '\n';
      } catch (e) { /* чужой или недоступный файл стилей: пропускаем */ }
    }
    const jobs = new Map();
    text.replace(URL_RE, (m, a, b, c) => {
      const raw = a || b || c || '';
      if (!raw || raw.startsWith('data:') || raw.startsWith('#')) return m;
      try { jobs.set(m, new URL(raw, base).href.split('#')[0]); } catch (e) { /* не адрес */ }
      return m;
    });
    const map = new Map();
    await Promise.all(Array.from(jobs, async ([m, abs]) => {
      if (!loaded.urls.has(abs)) { map.set(m, 'url("")'); return; }  /* страница этим не пользовалась */
      try { map.set(m, 'url("' + await toDataUrl(abs) + '")'); } catch (e) { map.set(m, 'url("")'); }
    }));
    return headCss + text.replace(URL_RE, m => (map.has(m) ? map.get(m) : m));
  }
  async function sheetText(href, loaded, depth) {
    if (!cssCache.has(href)) cssCache.set(href, (async () => {
      const r = await fetch(href, { cache: 'force-cache', credentials: 'same-origin' });
      return r.text();
    })());
    return inlineCss(await cssCache.get(href), href, loaded, depth || 0);
  }

  async function buildSvg(d, width) {
    const win = d.defaultView;
    const seen = w => w.performance.getEntriesByType('resource').map(e => e.name.split('#')[0]);
    const loaded = { urls: new Set(seen(win).concat(seen(window))),
      fonts: new Set(Array.from(d.fonts || []).filter(f => f.status === 'loaded').map(f => fontKey(f.family))) };
    const css = [];
    for (const sheet of Array.from(d.styleSheets)) {
      const media = sheet.media && sheet.media.mediaText;
      if (sheet.disabled || (media && !win.matchMedia(media).matches)) continue;
      try {
        if (sheet.href) {
          if (new URL(sheet.href, d.baseURI).origin === win.location.origin) css.push(await sheetText(sheet.href, loaded, 0));
        } else css.push(await inlineCss(Array.from(sheet.cssRules).map(r => r.cssText).join('\n'), d.baseURI, loaded, 0));
      } catch (e) { /* стиль недоступен: без него */ }
    }
    const rootEl = d.documentElement, clone = rootEl.cloneNode(true);
    clone.querySelectorAll('script, noscript, iframe, link, style, base, meta, title, object, embed').forEach(e => e.remove());
    const live = Array.from(rootEl.querySelectorAll('input, textarea, select'));
    Array.from(clone.querySelectorAll('input, textarea, select')).forEach((c, i) => {
      const l = live[i];
      if (!l) return;
      if (c.tagName === 'INPUT') {
        if (/^(checkbox|radio)$/i.test(l.type)) { if (l.checked) c.setAttribute('checked', ''); else c.removeAttribute('checked'); }
        else c.setAttribute('value', l.value);
      } else if (c.tagName === 'TEXTAREA') c.textContent = l.value;
      else Array.from(c.options).forEach((o, j) => {
        if (l.options[j] && l.options[j].selected) o.setAttribute('selected', ''); else o.removeAttribute('selected');
      });
    });
    await Promise.all(Array.from(clone.querySelectorAll('img')).map(async img => {
      const src = img.getAttribute('src');
      img.removeAttribute('srcset');
      img.removeAttribute('loading');
      if (!src || src.startsWith('data:')) return;
      try { img.setAttribute('src', await toDataUrl(new URL(src, d.baseURI).href)); } catch (e) { img.removeAttribute('src'); }
    }));
    for (const el of Array.from(clone.querySelectorAll('[style*="url("]'))) {
      el.setAttribute('style', await inlineCss(el.getAttribute('style'), d.baseURI, loaded, 9));
    }
    const style = d.createElement('style');
    style.textContent = css.join('\n') + '\nhtml,body{overflow:visible!important}::-webkit-scrollbar{display:none!important}';
    (clone.querySelector('head') || clone).appendChild(style);
    const height = Math.max(HEIGHT, rootEl.scrollHeight, d.body ? d.body.scrollHeight : 0);
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + width + '" height="' + height + '" style="font-size:' +
      win.getComputedStyle(rootEl).fontSize + '"><foreignObject x="0" y="0" width="100%" height="100%">' +
      new XMLSerializer().serializeToString(clone) + '</foreignObject></svg>';
    return { svg, width, height, background: win.getComputedStyle(d.body).backgroundColor };
  }

  async function drawSvg(b) {
    const img = new Image();
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(b.svg);
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = b.width;
    canvas.height = b.height;
    const g = canvas.getContext('2d', { willReadFrequently: true });
    g.fillStyle = b.background && b.background !== 'rgba(0, 0, 0, 0)' ? b.background : '#fff';
    g.fillRect(0, 0, b.width, b.height);
    g.drawImage(img, 0, 0);
    const px = g.getImageData(0, 0, b.width, b.height).data;
    let sig = 0, lo = 255, hi = 0;
    for (let i = 0; i < px.length; i += 148) {                /* выборка точек: отпечаток и разброс яркости */
      const y = (px[i] * 3 + px[i + 1] * 6 + px[i + 2]) / 10;
      if (y < lo) lo = y;
      if (y > hi) hi = y;
      sig = (Math.imul(sig, 31) + px[i] + px[i + 1] * 3 + px[i + 2] * 7) | 0;
    }
    return { canvas, sig, spread: hi - lo };
  }
  /* Первый проход иногда рисуется раньше, чем подключились шрифты. Берём кадр, только когда два подряд совпали. */
  async function shoot(d, width) {
    const b = await buildSvg(d, width);
    let prev = await drawSvg(b);
    for (let i = 0; i < 5; i++) {
      const next = await drawSvg(b);
      if (next.sig === prev.sig) { prev = next; break; }
      prev = next;
      await sleep(120);
    }
    if (prev.spread < 60) throw new SiteError('снимок получился пустым');
    return prev.canvas;
  }

  /* Полоса-подпись над снимком: ИИН, реестр, сколько записей, дата и время, адрес. Сама страница не изменяется. */
  const STRIP_FONT = 'Inter, "Segoe UI", Arial, sans-serif';
  const SEP = '   ·   ';
  function wrapText(g, text, max) {
    if (g.measureText(text).width <= max) return [text];
    const lines = [];
    let cur = '';
    for (let seg of text.split(SEP)) {
      const cand = cur ? cur + SEP + seg : seg;
      if (g.measureText(cand).width <= max) { cur = cand; continue; }
      if (cur) lines.push(cur);
      while (seg.length > 1 && g.measureText(seg).width > max) {
        let cut = seg.length - 1;
        while (cut > 1 && g.measureText(seg.slice(0, cut)).width > max) cut--;
        lines.push(seg.slice(0, cut));
        seg = seg.slice(cut);
      }
      cur = seg;
    }
    if (cur) lines.push(cur);
    return lines;
  }
  function stamped(shots) {
    const width = Math.max(...shots.map(s => s.canvas.width)), pad = 14, gap = 7;
    const probe = document.createElement('canvas').getContext('2d');
    const parts = shots.map(s => {
      const off = -s.at.getTimezoneOffset(), zone = 'UTC' + (off < 0 ? '-' : '+') + pad2(Math.floor(Math.abs(off) / 60)) + ':' + pad2(Math.abs(off) % 60);
      const l1 = 'ИИН ' + s.iin + SEP + s.reg.group + ' → ' + s.reg.title + SEP + 'Всего записей: ' + s.count +
        (s.rows < s.count ? ' (на снимке первые ' + s.rows + ')' : '');
      const l2 = 'Проверено: ' + ruDate(s.at, true) + ' (' + zone + ')' + SEP + s.url + SEP + 'Tazalau помощник ' + VERSION;
      const fit = (text, size, weight) => {
        for (; size > 12; size--) { probe.font = weight + ' ' + size + 'px ' + STRIP_FONT; if (probe.measureText(text).width <= width - 2 * pad) break; }
        probe.font = weight + ' ' + size + 'px ' + STRIP_FONT;
        return wrapText(probe, text, width - 2 * pad).map(t => ({ t, font: probe.font, size }));
      };
      const lines = fit(l1, 20, 'bold').map(x => Object.assign(x, { color: '#fff' }))
        .concat(fit(l2, 15, 'normal').map(x => Object.assign(x, { color: '#cbd5e1' })));
      const hStrip = 2 * pad + lines.reduce((a, x) => a + Math.round(x.size * 1.25), 0) + gap * (lines.length - 1);
      return { s, lines, hStrip };
    });
    const out = document.createElement('canvas');
    out.width = width;
    out.height = parts.reduce((a, p) => a + p.hStrip + p.s.canvas.height, 0);
    const g = out.getContext('2d');
    g.fillStyle = '#eef4f8';
    g.fillRect(0, 0, out.width, out.height);
    let y = 0;
    for (const p of parts) {
      g.fillStyle = '#243447';
      g.fillRect(0, y, width, p.hStrip);
      let ty = y + pad;
      g.textBaseline = 'top';
      for (const line of p.lines) {
        g.font = line.font;
        g.fillStyle = line.color;
        g.fillText(line.t, pad, ty);
        ty += Math.round(line.size * 1.25) + gap;
      }
      y += p.hStrip;
      g.drawImage(p.s.canvas, 0, y);
      y += p.s.canvas.height;
    }
    return new Promise((res, rej) => out.toBlob(b => (b ? res(b) : rej(new Error('не удалось собрать снимок'))), 'image/png'));
  }

  /* Одна страница реестра по одному ИИН: запрос, проверка, снимок */
  const timings = [];
  async function check(reg, iin) {
    const wait = lastRequest + interval - Date.now();
    if (wait > 0) await sleep(wait);
    lastRequest = Date.now();
    const url = location.origin + reg.path + '?' + reg.param + '=' + iin;
    const t0 = Date.now();
    try {
      const d = await go(url);
      const t1 = Date.now();
      let info = pageInfo(d, reg.param);
      const { count, rows } = verify(info, reg, iin);
      for (let i = 0; i < 4 && info.overflow > 0 && frameW < MAX_WIDTH; i++) {       /* таблица шире окна: расширяем */
        frameW = Math.min(MAX_WIDTH, frameW + info.overflow + 40);
        layoutFrame();
        await sleep(150);
        info = pageInfo(d, reg.param);
      }
      const at = new Date();
      const canvas = await shoot(d, frameW);
      timings.push({ load: t1 - t0, shot: Date.now() - at.getTime() });
      if (fdoc() !== d) throw new SiteError('страница сменилась во время снимка');
      return { reg, iin, url, count, rows, at, canvas, clipped: info.overflow > 0 };
    } finally {
      if (frameW !== WIDTH) { frameW = WIDTH; layoutFrame(); }
    }
  }

  class Stopped extends Error {}
  async function checkIin(regs, iin) {
    const shots = [];
    for (const reg of regs) {
      let attempt = 0, limited = 0;
      for (;;) {
        if (stopping) throw new Stopped();
        try { shots.push(await check(reg, iin)); break; }
        catch (e) {
          if (e instanceof LimitError) {
            if (++limited > LIMIT_RETRIES) return { error: reg.short + ': ' + e.message };
            const pause = Math.min(T.LIMIT_PAUSE * Math.pow(2, limited - 1), 180000);
            interval = Math.min(T.MAX_INTERVAL, Math.max(interval * 1.5, T.INTERVAL));
            say('⚠ Сайт сообщил «Превышен лимит запросов».\nЖду ' + Math.round(pause / 1000) + ' с, дальше один запрос в ' + (interval / 1000).toFixed(1) + ' с.');
            await sleep(pause);
          } else if (e instanceof SiteError) {
            if (++attempt >= RETRIES) return { error: reg.short + ': ' + e.message };
            await sleep(T.BACKOFF[Math.min(attempt - 1, T.BACKOFF.length - 1)]);
          } else throw e;
        }
      }
    }
    return { shots };
  }

  function describe(shots) {
    const found = shots.filter(s => s.count > 0);
    let text = found.length
      ? 'НАЙДЕН: ' + found.map(s => s.reg.short + ', записей: ' + s.count + (s.rows < s.count ? ' (на снимке первые ' + s.rows + ')' : '')).join('; ')
      : shots.length === 1 ? 'НЕ НАЙДЕН: ' + shots[0].reg.short : 'НЕ НАЙДЕН ни в одном из ' + shots.length + ' реестров';
    if (shots.some(s => s.clipped)) text += ' · внимание: таблица шире снимка';
    return { found: found.length > 0, text: text + ' · проверено ' + ruDate(shots[shots.length - 1].at, false) };
  }

  /* ---------- запуск ---------- */
  async function run() {
    if (running || !S) return;
    const regs = chosen();
    if (!regs.length) { say('❌ Отметьте хотя бы один реестр.'); regBox.open = true; return; }
    let recheck = false, p = plan(false);
    if (!p.groups.size) {
      if (!tasks.length) { say('В столбце «ИИН/БИН» нет значений.'); return; }
      if (p.linked) {                 /* новых ИИН нет (строки с опечатками не в счёт): спрашиваем, проверять ли заново */
        const a = await ask('Новых ИИН нет: ссылка уже записана.\nПроверить все ИИН заново? Старые снимки останутся.', ['Все заново', 'Отмена']);
        if (a !== 'Все заново') return;
        recheck = true;
        p = plan(true);
      } else if (!p.invalid.length) { say('Все строки с ИИН скрыты фильтром Excel: снимите галочку «только видимые строки».'); return; }
    }
    running = true; stopping = false; interval = T.INTERVAL;
    updateButtons();
    say('Проверка началась…');
    let done = 0, found = 0, notFound = 0, errors = 0, fails = 0, unsaved = 0, written = 0, stopNote = '';
    const total = p.groups.size;
    try {
      if (store.handle) {
        store.shots = await store.dir.getDirectoryHandle(SHOTS_DIR, { create: true });
        const w = await store.handle.createWritable({ keepExistingData: true });   /* файл не занят? выясняем до работы */
        await w.close();
      }
      ensureColumns(S);
      const L = S.layout;
      for (const t of p.invalid) {
        try { setText(S, t.row, L.resultCol, 'ОШИБКА: ' + t.parse.err, null); } catch (e) { /* объединённая ячейка */ }
        logRow({ time: ruDate(new Date(), true), iin: clean(t.raw) || '?', rows: String(t.row), status: 'НЕВЕРНЫЙ ИИН', counts: '', file: '', note: t.parse.err });
      }
      for (const [iin, rows] of p.groups) {
        if (stopping) break;
        say('Проверяю ' + (done + 1) + ' из ' + total + ': ' + iin + '\nнайдено ' + found + ' · не найдено ' + notFound + ' · ошибок ' + errors);
        let res;
        try { res = await checkIin(regs, iin); } catch (e) { if (e instanceof Stopped) break; throw e; }
        if (res.error) {
          errors++; fails++;
          for (const r of rows) { try { setText(S, r, L.resultCol, 'ОШИБКА: ' + res.error, null); } catch (e) { /* объединённая ячейка */ } }
          logRow({ time: ruDate(new Date(), true), iin, rows: rows.join(', '), status: 'ОШИБКА', counts: '', file: '', note: res.error });
        } else {
          fails = 0;
          const d = describe(res.shots), name = iin + '_' + stampOf(res.shots[res.shots.length - 1].at) + '.png';
          const blob = await stamped(res.shots);
          if (store.shots) await writeFile(store.shots, name, blob); else saveBlob(blob, name);
          const target = store.shots ? SHOTS_DIR + '/' + name : name;
          let note = '';
          for (const r of rows) {
            try {
              const cur = (S.cells.get(r) || new Map()).get(L.linkCol);
              setText(S, r, L.linkCol, store.shots ? SHOTS_DIR + '\\' + name : name, await linkStyle(S, cur ? cur.getAttribute('s') : null));
              await setLink(S, r, L.linkCol, target);
              setText(S, r, L.resultCol, d.text, null);
              written++;
            } catch (e) { note = e.message; }
          }
          if (d.found) found++; else notFound++;
          logRow({ time: ruDate(new Date(), true), iin, rows: rows.join(', '), status: d.found ? 'НАЙДЕН' : 'НЕ НАЙДЕН',
            counts: res.shots.map(s => s.reg.key + '=' + s.count).join('; '), file: name, note });
        }
        done++; unsaved++;
        if (fails >= MAX_FAILS) { stopNote = 'Сайт не отвечает: ' + fails + ' ИИН подряд не удалось проверить. Попробуйте позже.'; break; }
        if (unsaved >= T.SAVE_EVERY && store.handle) { if (await saveBook(false)) unsaved = 0; }
      }
      await saveBook(true);
    } catch (e) {
      stopNote = 'Ошибка: ' + (e && e.message ? e.message : e);
      try { await saveBook(true); } catch (x) { /* сообщение об ошибке уже есть */ }
    }
    running = false;
    const headLine = stopNote ? '❌ ' + stopNote : stopping ? '⏹ Остановлено.' : '✔ Готово.';
    tasks = readTasks(S);
    refresh();
    say(headLine + '\nПроверено ' + done + ' из ' + total + ' · найдено ' + found + ' · не найдено ' + notFound +
      (errors ? ' · ошибок ' + errors : '') + (p.invalid.length ? '\nНеверных ИИН: ' + p.invalid.length : '') +
      '\nСсылок записано: ' + written + (store.handle ? '' : '\nФайлы пришли в «Загрузки»: держите их в одной папке.') +
      (recheck ? '\nПроверка заново: прежние снимки остались в папке.' : ''));
    if (journal.length && store.shots) {
      try { await writeFile(store.shots, 'журнал_tazalau_' + stampOf(new Date()) + '.csv', csvBlob()); } catch (e) { /* журнал доступен кнопкой CSV */ }
    }
    beep();
  }
  function stopRun() {
    if (!running) return;
    stopping = true;
    stopB.disabled = true;
    say(state.textContent + '\n⏹ Останавливаю после текущего ИИН…');
  }

  window.__tazalau = {
    version: VERSION,
    show: () => { if (!root.isConnected) document.body.appendChild(root); },
    /* для автоматических проверок помощника */
    test: { normIin, checkDigit, openBook, loadSheet, readTasks, ensureColumns, setText, setLink, linkStyle, bookBytes, zipRead,
      useDir, openBytes, run, timings,
      state: () => ({ running, journal, tasks: tasks.length, sheet: S && S.sheet.name, layout: S && S.layout }) },
  };
  layoutFrame();
  lastRequest = Date.now();           /* первая страница в рамке тоже запрос: следующий пойдёт не раньше, чем через паузу */
  frame.src = START_URL;
  refresh();
})();
