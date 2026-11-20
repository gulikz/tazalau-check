/* Tazalau помощник: версия 6.
   Закладка для tazalau.qoldau.kz (Edge или Chrome). Ничего не устанавливает. Ежедневная проверка в три шага:
   1) ДГД: публикации 20 департаментов на gov.kz о процедурах граждан (приказы о финансовых управляющих, объявления о
      возбуждении дел, реестры требований кредиторов) за 31 день: файлы, ИИН из них, проверка каждого региона, свод;
   2) рабочий файл: выгрузка ADW (xlsx или csv) с проблемными кредитами и картами, столбец с ИИН;
   3) по каждому ИИН все шесть реестров tazalau: статус, снимок страницы, ссылки в файл; ИИН из списков ДГД первыми.
   По файлу Excel (столбец «ИИН/БИН»): открывает реестр с фильтром по ИИН, проверяет, что на странице
   именно результат поиска, сохраняет снимок страницы и пишет ссылку на него в столбец tazalau_file_location.
   Снимок каждого человека лежит рядом с файлом Excel в своей папке с именем из столбца FIO и называется «снимок.png».
   В столбец hyperlink пишется полная ссылка: путь к папке на диске + папка человека + «снимок.png».
   Вход на сайт и ЭЦП не нужны. Файл Excel читается и изменяется только внутри браузера и никуда не отправляется. */
(() => {
  'use strict';
  const VERSION = '6';
  if (window.__tazalau) { window.__tazalau.show(); return; }
  if (/^(www\.)?gov\.kz$/i.test(location.hostname)) { relay(); return; }
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

  /* На gov.kz закладка работает связным. Приказы и объявления ДГД всех 20 регионов лежат на gov.kz, а браузер разрешает
     читать страницы и файлы gov.kz только странице самого gov.kz. Поэтому помощник из вкладки tazalau открывает вкладку
     gov.kz, а закладка, нажатая там один раз, по просьбе помощника скачивает страницы и файлы gov.kz и передаёт их
     обратно. Страницу публикации, где ссылку на файл достраивает сам сайт, связной открывает в невидимой рамке (тот же
     сайт) и берёт ссылку оттуда, как человек со страницы. Чужие адреса не скачиваются, просьбы принимаются только от
     вкладки tazalau.qoldau.kz, которая открыла эту вкладку. Всё остальное делает помощник. */
  function relay() {
    const HOME = 'https://tazalau.qoldau.kz';
    if (window.__tazalauRelay) { window.__tazalauRelay(); return; }
    if (!document.body) { document.addEventListener('DOMContentLoaded', relay, { once: true }); return; }   /* страница ещё грузится */
    const box = document.createElement('div');
    box.style.cssText = 'position:fixed;top:10px;right:10px;z-index:2147483647;max-width:420px;padding:10px 12px;border-radius:8px;' +
      'background:#fff8e1;border:2px solid #e3b341;font:14px/1.45 Arial,sans-serif;color:#1d2433;box-shadow:0 2px 12px rgba(0,0,0,.25);white-space:pre-line;text-align:left';
    const say = t => { box.textContent = 'Tazalau ▶ ДГД (версия ' + VERSION + ')\n' + t; };
    document.body.appendChild(box);
    window.__tazalauRelay = () => { if (!box.isConnected) document.body.appendChild(box); };
    let opener = null;
    try { opener = window.opener; } catch (e) { opener = null; }
    if (!opener) {
      say('Эту вкладку открыл не помощник. Откройте tazalau.qoldau.kz, нажмите закладку Tazalau и кнопку «Собрать ДГД с gov.kz»: ' +
        'помощник сам откроет вкладку gov.kz, а в ней останется один раз нажать закладку.');
      return;
    }
    const FILE_HREF = /\/uploads\/|\/api\/v1\/public\/assets\/|\/api\/v2\/(public\/)?files\//i;
    /* Страница публикации в невидимой рамке (тот же сайт): ссылки на файлы, которые показал сам сайт, и запросы страницы */
    const frame = (url, wait) => new Promise(resolve => {
      const f = document.createElement('iframe');
      f.style.cssText = 'position:fixed;left:-12000px;top:0;width:1280px;height:900px;border:0';
      f.setAttribute('aria-hidden', 'true');
      f.src = url;
      document.body.appendChild(f);
      const t0 = Date.now();
      const look = () => {
        let links = [], resources = [], text = '';
        try {
          const d = f.contentDocument;
          if (d && d.body) {
            links = Array.from(d.querySelectorAll('a[href]')).filter(a => FILE_HREF.test(a.href) && new URL(a.href).origin === location.origin)
              .map(a => ({ href: a.href, text: (a.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 300) }));
            text = (d.body.innerText || '').replace(/\s+/g, ' ').slice(0, 400);
            resources = f.contentWindow.performance.getEntriesByType('resource').map(r => r.name)
              .filter(n => n.startsWith(location.origin) && /\/api\//.test(n)).slice(0, 60);
          }
        } catch (e) { text = 'рамка не открылась: ' + ((e && e.message) || e); }
        const done = links.length || Date.now() - t0 > wait || /mainFileNotExist|Файлы не найдены|Файлдар табылмады/i.test(text);
        if (!done) { setTimeout(look, 600); return; }
        setTimeout(() => f.remove(), 50);
        resolve({ links, resources, text });
      };
      setTimeout(look, 800);
    });
    window.addEventListener('message', async e => {
      const m = e.data;
      if (e.origin !== HOME || e.source !== opener || !m || typeof m.tz !== 'string') return;
      if (m.tz === 'say') { say(String(m.text || '')); return; }
      if (m.tz !== 'fetch' && m.tz !== 'frame') return;
      try {
        const u = new URL(String(m.url), location.href);
        if (u.origin !== location.origin) throw new Error('адрес другого сайта: ' + u.origin);
        if (m.tz === 'frame') {
          const r = await frame(u.href, Math.min(60000, +m.wait || 25000));
          opener.postMessage({ tz: 'fetched', id: m.id, ok: true, status: 200, links: r.links, resources: r.resources, text: r.text }, HOME);
          return;
        }
        const head = m.method === 'HEAD';
        const r = await fetch(u.href, { method: head ? 'HEAD' : 'GET', credentials: 'same-origin', cache: 'no-cache', redirect: 'follow' });
        const buf = head ? new ArrayBuffer(0) : await r.arrayBuffer();
        opener.postMessage({ tz: 'fetched', id: m.id, ok: r.ok, status: r.status, type: r.headers.get('content-type') || '',
          disposition: r.headers.get('content-disposition') || '', modified: r.headers.get('last-modified') || '', url: r.url, bytes: buf }, HOME, [buf]);
      } catch (err) {
        opener.postMessage({ tz: 'fetched', id: m.id, ok: false, status: 0, error: String((err && err.message) || err) }, HOME);
      }
    });
    opener.postMessage({ tz: 'hello', href: location.href, host: location.host, version: VERSION }, HOME);
    say('Связь с помощником есть: собираю приказы и объявления ДГД всех 20 регионов. Вкладку не закрывайте, ход сбора виден здесь и во вкладке tazalau.');
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
  const DEFAULT_KEYS = REGISTRIES.map(g => g.key);   /* ежедневная проверка: все шесть реестров */
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
    SETTLE: 80,                        /* пауза между проходами отрисовки снимка, мс: за неё браузер подключает шрифты */
  }, window.__tazalauTest || {});      /* автоматические проверки укорачивают паузы; в работе не используется */
  const RETRIES = 3;                   /* попыток на одну страницу */
  const LIMIT_RETRIES = 5;
  const MAX_FAILS = 3;                 /* столько ИИН подряд с ошибкой: сайт недоступен, стоп */
  const MAX_RECORDS = 50;              /* больше записей на один ИИН: фильтр не применился */
  const SCAN_ROWS = 30;                /* в скольких верхних строках искать заголовок «ИИН/БИН» */
  const LINK_HEADER = 'tazalau_file_location';
  const RESULT_HEADER = 'tazalau_result';
  const STATUS_HEADER = 'tazalau_status';     /* главный статус: самая свежая запись из всех реестров */
  const DETAILS_HEADER = 'tazalau_details';   /* все найденные записи: статусы, даты, суд или услугодатель */
  const ABS_HEADER = 'hyperlink';      /* полная ссылка: путь к папке на диске + папка человека + снимок */
  const SHOT_NAME = 'снимок.png';      /* так называется снимок в папке каждого человека */
  const BACKUP_DIR = 'tazalau_backup'; /* копии файла Excel перед записью и журналы проверок */
  const WIN_PATH = 259;                /* путь к файлу длиннее этого Проводник и Excel не открывают */
  const SWAP = 7;                      /* браузер пишет файл через временный «имя.crswap»: запас на его имя */
  const PATH_FILE = 'путь_к_папке.txt';    /* в папке tazalau_backup: путь к папке на диске, вставленный вручную */
  /* ========================================================= */

  const START_URL = location.origin + REGISTRIES[0].path;
  /* Заголовки столбцов ИИН и ФИО (без пробелов, строчными). Выгрузки из ADW называют их по-разному. */
  const IIN_HEADERS = ['иин/бин', 'бин/иин', 'иин', 'iin/bin', 'iin', 'iin_bin', 'iinbin', 'иин_бин', 'client_iin', 'iin_client',
    'иинклиента', 'иинзаемщика', 'иинзаёмщика', 'иинфл'];
  const FIO_HEADERS = ['fio', 'фио', 'fullname', 'full_name', 'client_fio', 'fio_client', 'client_name', 'clientname',
    'фиоклиента', 'фиозаемщика', 'фиозаёмщика', 'фио_клиента'];     /* столбец с именем человека: из него имя папки */
  const ABS_HEADERS = [ABS_HEADER, 'гиперссылка'];
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

  /* ---------- новая книга Excel с нуля: свод ДГД, книга из CSV ---------- */
  const xmlEsc = v => String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  /* Оформление ячеек: 0 обычная, 1 заголовок, 2 плохо (красная), 3 внимание (жёлтая), 4 в порядке (зелёная), 5 ссылка */
  const STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="6"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/><family val="2"/></font><font><sz val="11"/><color rgb="FF9C5700"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/></font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="4" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const STYLE = { head: 1, bad: 2, warn: 3, good: 4, link: 5 };
  /* sheets: [{ name, rows: [[значение | { v, s, link }]], widths: [ширина столбца] }]. Все значения пишутся текстом:
     ИИН и номера договоров не теряют ведущие нули. Первая строка: заголовок с фильтром, закреплена. */
  async function makeXlsx(sheets) {
    const files = [];
    const add = (name, text) => files.push([name, text]);
    const sheetName = n => String(n).replace(/[\\/?*:\[\]]/g, ' ').slice(0, 31) || 'Лист';
    add('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map((x, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>');
    add('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PNS + '">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    add('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="' + NS + '" xmlns:r="' + RNS + '"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
      sheets.map((x, i) => '<sheet name="' + xmlEsc(sheetName(x.name)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>');
    add('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PNS + '">' +
      sheets.map((x, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + RNS + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="' + RNS + '/styles" Target="styles.xml"/></Relationships>');
    add('xl/styles.xml', STYLES_XML);
    sheets.forEach((x, i) => {
      const width = Math.max(1, ...x.rows.map(r => r.length));
      const links = [];
      const rowsXml = x.rows.map((r, ri) => '<row r="' + (ri + 1) + '">' + r.map((c, ci) => {
        const cell = c != null && typeof c === 'object' ? c : { v: c };
        const ref = colName(ci + 1) + (ri + 1);
        if (cell.link) links.push([ref, cell.link]);
        const st = ri === 0 && x.head !== false ? STYLE.head : cell.s || (cell.link ? STYLE.link : 0);
        const v = cell.v == null ? '' : String(cell.v);
        return '<c r="' + ref + '"' + (st ? ' s="' + st + '"' : '') + (v === '' ? '/>' : ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(v) + '</t></is></c>');
      }).join('') + '</row>').join('');
      const last = colName(width) + Math.max(1, x.rows.length);
      const cols = (x.widths || []).length ? '<cols>' + x.widths.map((w, ci) => '<col min="' + (ci + 1) + '" max="' + (ci + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>' : '';
      add('xl/worksheets/sheet' + (i + 1) + '.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="' + NS + '" xmlns:r="' + RNS + '">' +
        '<dimension ref="A1:' + last + '"/>' + (x.head !== false && x.rows.length > 1 ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
          : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>') + '<sheetFormatPr defaultRowHeight="15"/>' + cols +
        '<sheetData>' + rowsXml + '</sheetData>' + (x.head !== false && x.rows.length > 1 ? '<autoFilter ref="A1:' + last + '"/>' : '') +
        (links.length ? '<hyperlinks>' + links.map(([ref], li) => '<hyperlink ref="' + ref + '" r:id="rId' + (li + 1) + '"/>').join('') + '</hyperlinks>' : '') +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
        '<ignoredErrors><ignoredError sqref="A1:' + last + '" numberStoredAsText="1"/></ignoredErrors></worksheet>');
      if (links.length) {
        add('xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PNS + '">' +
          links.map(([, target], li) => '<Relationship Id="rId' + (li + 1) + '" Type="' + HYPERLINK_TYPE + '" Target="' + xmlEsc(escTarget(target)) + '" TargetMode="External"/>').join('') + '</Relationships>');
      }
    });
    const entries = [];
    for (const [name, text] of files) entries.push(await makeEntry(name, te.encode(text), null));
    return zipWrite(entries);
  }

  /* ---------- текст из файлов: кодировка, CSV ---------- */
  /* UTF-8, а если файл в ней не читается (Excel в русской Windows сохраняет CSV в кодировке 1251), то windows-1251 */
  function decodeText(bytes) {
    let u = bytes;
    if (u[0] === 0xef && u[1] === 0xbb && u[2] === 0xbf) u = u.subarray(3);
    else if (u[0] === 0xff && u[1] === 0xfe) return new TextDecoder('utf-16le').decode(u.subarray(2));
    else if (u[0] === 0xfe && u[1] === 0xff) return new TextDecoder('utf-16be').decode(u.subarray(2));
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u); }
    catch (e) { return new TextDecoder('windows-1251').decode(u); }
  }
  /* CSV по правилам RFC 4180 (кавычки, переносы строк внутри кавычек); разделитель: ; , табуляция или | по первой строке */
  function parseCsv(text) {
    const first = text.slice(0, Math.max(0, text.indexOf('\n')) || text.length);
    const sep = [';', ',', '\t', '|'].map(c => [c, first.split(c).length]).sort((a, b) => b[1] - a[1])[0][0];
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"' && cell === '') q = true;
      else if (ch === sep) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); cell = '';
        if (row.some(c => c !== '')) rows.push(row);
        row = [];
      } else cell += ch;
    }
    row.push(cell);
    if (row.some(c => c !== '')) rows.push(row);
    return rows;
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

  /* Путь к папке на диске в том виде, как его вставили: без кавычек и невидимых знаков, с обратными чертами.
     Вставили путь к самому файлу Excel: берётся его папка. Адрес вида file:/// тоже понимается. */
  function normBase(v) {
    let s = String(v == null ? '' : v).replace(/[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]/g, '').trim()
      .replace(/^"+|"+$/g, '').trim();
    if (/^file:\/\//i.test(s)) {
      s = s.replace(/^file:\/+/i, '');
      try { s = decodeURIComponent(s); } catch (e) { /* оставляем как есть */ }
      if (!/^[a-z]:/i.test(s)) s = '\\\\' + s;
    }
    s = s.replace(/\//g, '\\');
    s = s.slice(0, 2) + s.slice(2).replace(/\\{2,}/g, '\\');
    return s.replace(/\\[^\\]+\.xls[xmb]?$/i, '').replace(/\\+$/, '');
  }
  const baseValid = s => /^[a-z]:(\\[^\\/:*?"<>|]+)*$/i.test(s) || /^\\\\[^\\/:*?"<>|]+(\\[^\\/:*?"<>|]+)+$/.test(s);
  const lastPart = s => s.slice(s.lastIndexOf('\\') + 1);
  const sameName = (a, b) => String(a).normalize('NFC').toLowerCase() === String(b).normalize('NFC').toLowerCase();
  /* Сколько знаков остаётся на имя папки человека, чтобы путь к снимку (и к прежнему снимку с номером) открывался в Windows */
  const folderRoom = base => Math.min(100, WIN_PATH - SWAP - base.length - 2 - (SHOT_NAME.length + 3));

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
    const abs = all(wb, 'absPath')[0];            /* Excel записывает в файл папку, где его сохранили */
    book.absPath = abs ? abs.getAttribute('url') || '' : '';
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
    const L = { headerRow: best.row, iinCol: best.col, fioCol: 0, linkCol: 0, absCol: 0, resultCol: 0, statusCol: 0, detailsCol: 0,
      dgdCol: 0, lastRow: best.row };
    for (const [c, el] of S.cells.get(best.row)) {
      const x = cellValue(S, el), name = x && x.v != null ? normHead(x.v) : '';
      if (name === LINK_HEADER && (!L.linkCol || c < L.linkCol)) L.linkCol = c;
      if (name === RESULT_HEADER && (!L.resultCol || c < L.resultCol)) L.resultCol = c;
      if (name === STATUS_HEADER && (!L.statusCol || c < L.statusCol)) L.statusCol = c;
      if (name === DETAILS_HEADER && (!L.detailsCol || c < L.detailsCol)) L.detailsCol = c;
      if (name === 'дгд' && (!L.dgdCol || c < L.dgdCol)) L.dgdCol = c;
      if (ABS_HEADERS.includes(name) && (!L.absCol || c < L.absCol)) L.absCol = c;
      if (FIO_HEADERS.includes(name.replace(/\./g, '')) && (!L.fioCol || c < L.fioCol)) L.fioCol = c;
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
      let fio = L.fioCol ? clean(textAt(S, r, L.fioCol)) : '';
      if (/^(#N\/A|#Н\/Д|-+|нет|n\/a|null|none)$/i.test(fio)) fio = '';       /* заглушка вместо имени: имени нет */
      tasks.push({ row: r, parse, fio, raw: x.formula ? '' : x.v, hasLink: !!L.linkCol && linked(r, L.linkCol),
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

  /* Прежняя ссылка ячейки (если была) убирается вместе с её адресом */
  async function dropLink(S, r, col) {
    const ref = colName(col) + r, hl = kids(S.doc.documentElement, 'hyperlinks')[0];
    if (!hl || !kids(hl, 'hyperlink').some(h => h.getAttribute('ref') === ref)) return;
    const rels = await sheetRels(S);
    for (const h of kids(hl, 'hyperlink')) {
      if (h.getAttribute('ref') !== ref) continue;
      const old = h.getAttributeNS(RNS, 'id');
      h.remove();
      if (old && !kids(hl, 'hyperlink').some(x => x.getAttributeNS(RNS, 'id') === old)) {
        all(rels, 'Relationship').filter(x => x.getAttribute('Id') === old && x.getAttribute('Type') === HYPERLINK_TYPE).forEach(x => x.remove());
      }
    }
    if (!kids(hl, 'hyperlink').length) hl.remove();          /* пустой список ссылок Excel считает ошибкой */
    S.linkRanges = S.linkRanges.filter(g => !(g.c1 === col && g.c2 === col && g.r1 === r && g.r2 === r));
    S.dirty = true;
  }

  async function setLink(S, r, col, target) {
    await dropLink(S, r, col);
    const rels = await sheetRels(S), ref = colName(col) + r, root = S.doc.documentElement;
    let hl = kids(root, 'hyperlinks')[0];
    if (!hl) {
      hl = S.doc.createElementNS(NS, 'hyperlinks');
      root.insertBefore(hl, Array.from(root.children).find(c => AFTER_LINKS.includes(c.localName)) || null);
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

  /* Полная ссылка записывается формулой =HYPERLINK("путь"): обычную ссылку на файл того же диска Excel при
     сохранении книги переделывает в относительную, и после переноса или отправки книги она перестаёт открываться.
     Формула остаётся как есть и всегда ведёт по полному пути. В ячейке виден сам путь. */
  async function setPathLink(S, r, col, path, style) {
    if (mergedInto(S, r, col)) throw new Error('строка ' + r + ': ячейка ' + colName(col) + r + ' объединена с соседними');
    await dropLink(S, r, col);
    const c = cellEl(S, r, col);
    while (c.firstChild) c.firstChild.remove();
    c.setAttribute('t', 'str');
    if (style != null) c.setAttribute('s', style);
    const f = S.doc.createElementNS(NS, 'f'), v = S.doc.createElementNS(NS, 'v');
    f.textContent = 'HYPERLINK("' + path.replace(/"/g, '""') + '")';
    v.textContent = path;
    c.append(f, v);
    S.dirty = true;
    return c;
  }

  /* Вид ссылки (синяя, подчёркнутая) поверх прежнего оформления ячейки: заливка и рамки остаются.
     general: числовой формат «Общий» (в ячейке с форматом «Текст» формула при правке превратилась бы в текст). */
  async function linkStyle(S, s, general) {
    const book = S.book;
    if (book.styles === undefined) {
      const text = book.stylesPath ? await book.text(book.stylesPath) : null;
      book.styles = text ? parseXml(text, 'styles') : null;
      book.linkXf = new Map();
    }
    const st = book.styles, key = (general ? 'g' : '') + (s == null ? '' : String(s));
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
    if (general) { nxf.setAttribute('numFmtId', '0'); nxf.removeAttribute('applyNumberFormat'); }
    xfs.appendChild(nxf);
    xfs.setAttribute('count', xfList.length + 1);
    book.stylesDirty = true;
    book.linkXf.set(key, String(xfList.length));
    return String(xfList.length);
  }

  /* Недостающие столбцы программы: сразу за таблицей (объединённые ячейки шапки тоже считаются шапкой) */
  function ensureColumns(S, withAbs, withDgd) {
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
      let col = Math.max(last, L.linkCol || 0, L.absCol || 0, L.resultCol || 0, L.statusCol || 0, L.detailsCol || 0, L.dgdCol || 0) + 1;
      while (used.has(col)) col++;
      setText(S, L.headerRow, col, name, headStyle);
      used.add(col);
      setWidth(S, col, width);
      return col;
    };
    if (!L.linkCol) L.linkCol = add(LINK_HEADER, 60);
    if (withAbs && !L.absCol) L.absCol = add(ABS_HEADER, 60);
    if (!L.resultCol) L.resultCol = add(RESULT_HEADER, 46);
    if (!L.statusCol) L.statusCol = add(STATUS_HEADER, 60);
    if (!L.detailsCol) L.detailsCol = add(DETAILS_HEADER, 90);
    if (withDgd && !L.dgdCol) L.dgdCol = add(DGD_HEADER, 50);
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

  /* ---------- текст из PDF (только PDF с текстовым слоем; скан не читается) ---------- */
  /* Разбор без библиотек: объекты файла (в том числе упакованные в потоки объектов), шрифты с таблицами
     перевода в Юникод (ToUnicode), потоки содержимого страниц. Достаточно, чтобы прочитать ИИН и даты. */
  const latin1 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return s; };
  function pdfParse(str) {                    /* значение PDF: словарь, массив, имя, число, ссылка, строка */
    let i = 0;
    const ws = () => { for (;;) { while (i < str.length && /[\s\0]/.test(str[i])) i++; if (str[i] === '%') { while (i < str.length && str[i] !== '\n' && str[i] !== '\r') i++; } else return; } };
    const val = () => {
      ws();
      const c = str[i];
      if (c === '<' && str[i + 1] === '<') {
        i += 2; const d = {};
        for (;;) { ws(); if (str[i] === '>' && str[i + 1] === '>') { i += 2; return d; } if (i >= str.length) return d; const k = val(); if (!k || k.name == null) { i++; continue; } d[k.name] = val(); }
      }
      if (c === '[') { i++; const a = []; for (;;) { ws(); if (str[i] === ']') { i++; return a; } if (i >= str.length) return a; a.push(val()); } }
      if (c === '/') { let j = i + 1; while (j < str.length && !/[\s\0/<>\[\]()%{}]/.test(str[j])) j++; const n = str.slice(i + 1, j).replace(/#([0-9a-f]{2})/gi, (x, hx) => String.fromCharCode(parseInt(hx, 16))); i = j; return { name: n }; }
      if (c === '(') { const r = pdfString(str, i); i = r.end; return { str: r.s }; }
      if (c === '<') { const j = str.indexOf('>', i); const hx = str.slice(i + 1, j).replace(/[^0-9a-f]/gi, ''); i = j + 1; let s = ''; for (let k = 0; k < hx.length; k += 2) s += String.fromCharCode(parseInt((hx.slice(k, k + 2) + '0').slice(0, 2), 16)); return { str: s }; }
      const m = /^[+-]?(\d+\.?\d*|\.\d+)/.exec(str.slice(i, i + 40));
      if (m) {
        i += m[0].length;
        const r = /^\s+(\d+)\s+R(?![A-Za-z])/.exec(str.slice(i, i + 30));
        if (r && /^\d+$/.test(m[0])) { i += r[0].length; return { ref: m[0] }; }
        return Number(m[0]);
      }
      const w = /^[A-Za-z]+/.exec(str.slice(i, i + 20));
      if (w) { i += w[0].length; return w[0] === 'true' ? true : w[0] === 'false' ? false : null; }
      i++; return null;
    };
    return val();
  }
  function pdfString(str, i) {                /* литерал (...) с вложенными скобками и экранированием */
    let depth = 0, s = '', j = i;
    for (; j < str.length; j++) {
      const c = str[j];
      if (c === '\\') {
        const n = str[++j];
        if (n === 'n') s += '\n'; else if (n === 'r') s += '\r'; else if (n === 't') s += '\t'; else if (n === 'b') s += '\b'; else if (n === 'f') s += '\f';
        else if (n === '\r') { if (str[j + 1] === '\n') j++; }
        else if (n === '\n') { /* перенос строки внутри строки */ }
        else if (/[0-7]/.test(n)) { let o = n; while (o.length < 3 && /[0-7]/.test(str[j + 1])) o += str[++j]; s += String.fromCharCode(parseInt(o, 8) & 255); }
        else s += n;
      } else if (c === '(') { if (depth++) s += c; }
      else if (c === ')') { if (--depth === 0) return { s, end: j + 1 }; s += c; }
      else s += c;
    }
    return { s, end: j };
  }
  /* Объекты PDF-файла: словари, потоки (с распаковкой), страницы по порядку. Общая часть для текста и сканов. */
  async function pdfOpen(bytes) {
    const raw = latin1(bytes);
    if (!/^%PDF-/.test(raw.slice(0, 1024).trimStart())) throw new Error('это не PDF');
    if (/\/Encrypt\s*\d+\s+\d+\s+R|\/Encrypt\s*<</.test(raw)) throw new Error('PDF защищён от копирования текста: сохраните его заново без защиты или как Excel');
    const objs = new Map();
    const re = /(?:^|[\s>\]])(\d+)\s+(\d+)\s+obj\b/g;
    let m;
    while ((m = re.exec(raw))) {
      const start = m.index + m[0].length;
      let end = raw.indexOf('endobj', start);
      if (end < 0) end = raw.length;
      let body = raw.slice(start, end), data = null;
      const si = body.search(/stream(?:\r\n|\r|\n)/);           /* после stream: CRLF, LF или (так пишет HP Scan) один CR */
      if (si >= 0 && body.slice(0, si).includes('<<')) {
        const dict = pdfParse(body.slice(0, si));
        const off = start + si + /^stream(?:\r\n|\r|\n)/.exec(body.slice(si))[0].length;
        let len = dict && typeof dict.Length === 'number' ? dict.Length : -1;
        let stop = len >= 0 && raw.slice(off + len, off + len + 30).replace(/^\s+/, '').startsWith('endstream') ? off + len : raw.indexOf('endstream', off);
        if (stop < 0) stop = end;
        if (stop > end) { end = raw.indexOf('endobj', stop); if (end < 0) end = raw.length; }
        data = bytes.subarray(off, stop);
        body = body.slice(0, si);
      }
      objs.set(m[1], { body, data, dict: null });
      re.lastIndex = end;
    }
    const get = o => (o && o.ref != null ? (objs.has(o.ref) ? dictOf(objs.get(o.ref)) : null) : o);
    function dictOf(o) { if (o.dict === null) o.dict = pdfParse(o.body); return o.dict; }
    /* image: для картинок вернуть { data, filter, parms }: данные до последнего фильтра (JPEG, факс), который
       распаковывается отдельно; для остальных потоков: распакованные данные (картинки и редкие сжатия: пусто) */
    async function streamOf(o, image) {
      if (!o || !o.data) return image ? null : new Uint8Array(0);
      const d = dictOf(o) || {};
      let data = o.data;
      const filters = [].concat(get(d.Filter) || []).map(x => x && x.name);
      const parms = [].concat(get(d.DecodeParms) || get(d.DP) || []).map(x => get(x) || null);
      for (let k = 0; k < filters.length; k++) {                                    /* цепочка фильтров по порядку */
        const f = filters[k];
        if (f === 'FlateDecode' || f === 'Fl') {
          data = await pump(stream('in', 'deflate'), data, true);
          if (image && parms[k] && parms[k].Predictor >= 10) data = pngUnpredict(data, parms[k]);
        }
        else if (f === 'ASCII85Decode' || f === 'A85') data = ascii85(data);
        else if (f === 'ASCIIHexDecode' || f === 'AHx') { const hx = latin1(data).replace(/>.*$/s, '').replace(/[^0-9a-f]/gi, ''); data = Uint8Array.from(hx.match(/.{1,2}/g) || [], x => parseInt((x + '0').slice(0, 2), 16)); }
        else if (image) return { data, filter: f, parms: parms[k] || {} };
        else return new Uint8Array(0);           /* картинки и редкие способы сжатия: текста в них нет */
      }
      return image ? { data, filter: '', parms: {} } : data;
    }
    function pngUnpredict(u8, p) {             /* PNG-предсказатели строк (Predictor 10–15) */
      const colors = p.Colors || 1, bpc = p.BitsPerComponent || 8, cols = p.Columns || 1;
      const bpp = Math.max(1, Math.ceil(colors * bpc / 8)), rowLen = Math.ceil(colors * bpc * cols / 8);
      const rows = Math.floor(u8.length / (rowLen + 1)), out = new Uint8Array(rows * rowLen);
      for (let r = 0; r < rows; r++) {
        const t = u8[r * (rowLen + 1)], src = r * (rowLen + 1) + 1, dst = r * rowLen, up = dst - rowLen;
        for (let i = 0; i < rowLen; i++) {
          const a = i >= bpp ? out[dst + i - bpp] : 0, b = r ? out[up + i] : 0, c = r && i >= bpp ? out[up + i - bpp] : 0, x = u8[src + i];
          let v = x;
          if (t === 1) v = x + a; else if (t === 2) v = x + b; else if (t === 3) v = x + ((a + b) >> 1);
          else if (t === 4) { const pp = a + b - c, pa = Math.abs(pp - a), pb = Math.abs(pp - b), pc = Math.abs(pp - c); v = x + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c); }
          out[dst + i] = v & 255;
        }
      }
      return out;
    }
    function ascii85(u8) {
      const t = latin1(u8).replace(/^\s*<~/, '').replace(/~>.*$/s, '').replace(/\s+/g, '');
      const out = [];
      for (let i = 0; i < t.length;) {
        if (t[i] === 'z') { out.push(0, 0, 0, 0); i++; continue; }
        const chunk = t.slice(i, i + 5); i += 5;
        const n = chunk.length, padded = chunk.padEnd(5, 'u');
        let v = 0; for (const ch of padded) v = v * 85 + (ch.charCodeAt(0) - 33);
        const bytes = [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
        out.push(...bytes.slice(0, n - 1));
      }
      return Uint8Array.from(out);
    }
    for (const o of Array.from(objs.values())) {        /* объекты внутри потоков объектов (PDF 1.5+, так сохраняет Word) */
      const d = o.data ? dictOf(o) : null;
      if (!d || !d.Type || d.Type.name !== 'ObjStm') continue;
      const text = latin1(await streamOf(o));
      const first = d.First || 0, head = text.slice(0, first).trim().split(/\s+/).map(Number);
      for (let k = 0; k + 1 < head.length; k += 2) {
        const from = first + head[k + 1], to = k + 3 < head.length ? first + head[k + 3] : text.length;
        if (!objs.has(String(head[k]))) objs.set(String(head[k]), { body: text.slice(from, to), data: null, dict: null });
      }
    }
    const pages = [];
    const seenPages = new Set();
    const walk = node => {                      /* порядок страниц: дерево /Pages */
      if (!node || pages.length > 5000) return;
      if (node.Type && node.Type.name === 'Page') { if (!seenPages.has(node)) { seenPages.add(node); pages.push(node); } return; }
      for (const k of [].concat(get(node.Kids) || [])) walk(get(k));
    };
    for (const o of objs.values()) { const d = dictOf(o); if (d && d.Type && d.Type.name === 'Catalog') walk(get(d.Pages)); }
    if (!pages.length) for (const o of objs.values()) { const d = dictOf(o); if (d && d.Type && d.Type.name === 'Page') pages.push(d); }
    return { objs, get, dictOf, streamOf, pages };
  }
  async function pdfLines(bytes) {
    const { objs, get, streamOf, pages } = await pdfOpen(bytes);
    const cmaps = new Map();
    async function fontMap(font) {
      if (!font) return null;
      const key = font;
      if (cmaps.has(key)) return cmaps.get(key);
      const res = { width: font.Subtype && font.Subtype.name === 'Type0' ? 2 : 1, map: new Map(), w: new Map(), dw: 500 };
      cmaps.set(key, res);
      if (res.width === 2) {                    /* ширины знаков: по ним видно, где между кусками текста пробел */
        const desc = get([].concat(get(font.DescendantFonts) || [])[0]) || {};
        res.dw = typeof desc.DW === 'number' ? desc.DW : 1000;
        const W = get(desc.W) || [];
        for (let k = 0; k < W.length && res.w.size < 70000;) {
          const c = get(W[k]), nx = get(W[k + 1]);
          if (Array.isArray(nx)) { nx.forEach((v, j) => res.w.set(c + j, typeof v === 'number' ? v : res.dw)); k += 2; }
          else { const v = get(W[k + 2]); for (let cc = c; cc <= nx && cc - c < 70000; cc++) res.w.set(cc, v); k += 3; }
        }
      } else {
        const ws = get(font.Widths) || [], fc = typeof font.FirstChar === 'number' ? font.FirstChar : 0;
        ws.forEach((v, j) => { const x = get(v); if (typeof x === 'number') res.w.set(fc + j, x); });
      }
      const tu = font.ToUnicode && font.ToUnicode.ref != null ? objs.get(font.ToUnicode.ref) : null;
      if (tu) {
        const t = latin1(await streamOf(tu));
        const cs = /begincodespacerange\s*<([0-9a-f]+)>/i.exec(t);
        if (cs) res.width = Math.max(1, cs[1].length >> 1);
        const uni = hx => { let s = ''; for (let k = 0; k + 3 < hx.length; k += 4) s += String.fromCharCode(parseInt(hx.slice(k, k + 4), 16)); return s; };
        for (const blk of t.split('beginbfchar').slice(1)) {
          for (const x of blk.split('endbfchar')[0].matchAll(/<([0-9a-f]+)>\s*<([0-9a-f]*)>/gi)) res.map.set(parseInt(x[1], 16), uni(x[2]));
        }
        for (const blk of t.split('beginbfrange').slice(1)) {
          const part = blk.split('endbfrange')[0];
          for (const x of part.matchAll(/<([0-9a-f]+)>\s*<([0-9a-f]+)>\s*(<[0-9a-f]*>|\[[^\]]*\])/gi)) {
            const a = parseInt(x[1], 16), b = parseInt(x[2], 16);
            if (b < a || b - a > 65535) continue;
            if (x[3][0] === '[') { const list = Array.from(x[3].matchAll(/<([0-9a-f]*)>/gi)).map(y => uni(y[1])); for (let c = a; c <= b; c++) res.map.set(c, list[c - a] || ''); }
            else { const hx = x[3].slice(1, -1), base = parseInt(hx.slice(-4), 16), pre = uni(hx.slice(0, -4)); for (let c = a; c <= b; c++) res.map.set(c, pre + String.fromCharCode(base + c - a)); }
          }
        }
      }
      return res;
    }
    const lines = [];
    for (const pg of pages) {
      let resources = get(pg.Resources), up = get(pg.Parent);
      while (!resources && up) { resources = get(up.Resources); up = get(up.Parent); }      /* ресурсы наследуются от родителя */
      const fonts = get((resources || {}).Font) || {};
      let cs = pg.Contents;
      if (cs && cs.ref != null && objs.has(cs.ref) && !objs.get(cs.ref).data) cs = get(cs);  /* ссылка на массив потоков */
      const parts = [];
      for (const c of [].concat(cs || [])) { if (c && c.ref != null && objs.has(c.ref)) parts.push(latin1(await streamOf(objs.get(c.ref)))); }
      const content = parts.join('\n');
      /* Положение текста: где кончился предыдущий кусок и где начался следующий. Разрыв больше 0,15 кегля: пробел,
         другая высота строки: новая строка. Так буквы одного слова не разрываются, а соседние числа не слипаются. */
      let cur = '', font = null, size = 10, scale = 1, tc = 0, twd = 0, tz = 1, tl = 0, lineX = 0, lineY = null, curX = null;
      const flush = () => { if (cur.trim()) lines.push(cur.replace(/\s+/g, ' ').trim()); cur = ''; };
      const moveTo = (x, y) => {
        if (lineY !== null && Math.abs(y - lineY) > 0.3 * size * scale) flush();
        else if (curX !== null && x - curX > 0.15 * size * scale) cur += ' ';
        lineX = x; lineY = y; curX = x;
      };
      const show = s => {
        const width = font ? font.width : 1;
        for (let k = 0; k + width <= s.length; k += width) {
          let code = 0; for (let b = 0; b < width; b++) code = code * 256 + s.charCodeAt(k + b);
          const ch = font && font.map.size ? (font.map.has(code) ? font.map.get(code) : width === 1 ? String.fromCharCode(code) : '') : width === 1 ? String.fromCharCode(code) : '';
          cur += ch;
          const w = font && font.w.has(code) ? font.w.get(code) : font ? font.dw : 500;
          if (curX === null) curX = lineX;
          curX += ((w / 1000) * size + tc + (ch === ' ' ? twd : 0)) * tz * scale;
        }
      };
      const ops = [];
      let i = 0;
      const n = content.length;
      while (i < n) {
        const c = content[i];
        if (/[\s\0]/.test(c)) { i++; continue; }
        if (c === '%') { while (i < n && content[i] !== '\n' && content[i] !== '\r') i++; continue; }
        if (c === '(') { const r = pdfString(content, i); ops.push({ str: r.s }); i = r.end; continue; }
        if (c === '<' && content[i + 1] !== '<') { const j = content.indexOf('>', i); const hx = content.slice(i + 1, j).replace(/[^0-9a-f]/gi, ''); let s = ''; for (let k = 0; k < hx.length; k += 2) s += String.fromCharCode(parseInt((hx.slice(k, k + 2) + '0').slice(0, 2), 16)); ops.push({ str: s }); i = j + 1; continue; }
        if (c === '[') { ops.push({ open: true }); i++; continue; }
        if (c === ']') { const arr = []; while (ops.length && !ops[ops.length - 1].open) arr.unshift(ops.pop()); ops.pop(); ops.push({ arr }); i++; continue; }
        if (c === '<' || c === '>') { i += 2; ops.push(null); continue; }
        if (c === '/') { let j = i + 1; while (j < n && !/[\s\0/<>\[\]()%]/.test(content[j])) j++; ops.push({ name: content.slice(i + 1, j) }); i = j; continue; }
        const num = /^[+-]?(\d+\.?\d*|\.\d+)/.exec(content.slice(i, i + 32));
        if (num) { ops.push(Number(num[0])); i += num[0].length; continue; }
        let j = i; while (j < n && /[A-Za-z*'"]/.test(content[j])) j++;
        if (j === i) { i++; continue; }
        const op = content.slice(i, j);
        i = j;
        if (op === 'BI') { const e = content.indexOf('EI', i); i = e < 0 ? n : e + 2; ops.length = 0; continue; }   /* картинка внутри страницы */
        const a = ops.splice(0);
        const nums = a.filter(x => typeof x === 'number');
        if (op === 'Tf') { const nm = a.find(x => x && x.name); font = nm ? await fontMap(get(fonts[nm.name])) : null; if (nums.length) size = Math.abs(nums[0]) || size; }
        else if (op === 'Tj') { const t = a.find(x => x && x.str != null); if (t) show(t.str); }
        else if (op === "'" || op === '"') {
          if (op === '"' && nums.length >= 2) { twd = nums[0]; tc = nums[1]; }
          moveTo(lineX, (lineY || 0) - tl * scale); const t = a.reverse().find(x => x && x.str != null); if (t) show(t.str);
        }
        else if (op === 'TJ') {
          for (const x of (a.find(y => y && y.arr) || { arr: [] }).arr) {
            if (x && x.str != null) show(x.str);
            else if (typeof x === 'number') { if (curX !== null) curX -= (x / 1000) * size * tz * scale; if (x < -250) cur += ' '; }
          }
        }
        else if ((op === 'Td' || op === 'TD') && nums.length >= 2) { if (op === 'TD') tl = -nums[1]; moveTo(lineX + nums[0] * scale, (lineY || 0) + nums[1] * scale); }
        else if (op === 'Tm' && nums.length >= 6) { scale = Math.hypot(nums[0], nums[1]) || 1; moveTo(nums[4], nums[5]); }
        else if (op === 'T*') moveTo(lineX, (lineY || 0) - tl * scale);
        else if (op === 'TL' && nums.length) tl = nums[0];
        else if (op === 'Tc' && nums.length) tc = nums[0];
        else if (op === 'Tw' && nums.length) twd = nums[0];
        else if (op === 'Tz' && nums.length) tz = nums[0] / 100;
        else if (op === 'BT') { lineX = 0; scale = 1; }
      }
      flush();
      lines.push('');                          /* граница страниц */
    }
    return lines;
  }

  /* ====================== ИИН на сканах: распознавание цифр без библиотек ====================== */
  /* Приказы ДГД публикуются сканами (PDF с картинкой страницы или JPEG). Если у PDF есть текстовый слой (его добавляют
     сканер или ABBYY), ИИН берётся из текста. Иначе помощник читает цифры с картинки сам: страница становится
     чёрно-белой, находятся отдельные знаки, из них цепочки знаков одной высоты с ровным шагом (так выглядит число),
     каждый знак сравнивается с образцами цифр. ИИН засчитывается, только если сходится контрольная цифра и дата
     рождения настоящая; если цифру пришлось подобрать по контрольной цифре, ИИН помечается «проверить». */
  const OCR_DPI = 200;                     /* страница перед распознаванием: точек на дюйм */
  const OCR_PAGES = 3;                     /* ИИН ищется на первых страницах приказа */
  const OCR_GW = 12, OCR_GH = 18;          /* сетка, по которой знак сравнивается с образцом цифры */
  const OCR_SIZES = [16, 22, 30, 40];      /* высоты образцов цифр, точек */
  function canvasOf(w, h) {
    if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(w, h);
    const c = document.createElement('canvas');
    c.width = w;
    c.height = h;
    return c;
  }
  const ctx2d = c => c.getContext('2d', { willReadFrequently: true });
  const mulM = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];

  /* Картинка PDF (XObject /Image) как холст или ImageBitmap */
  async function pdfImage(doc, o) {
    const d = doc.dictOf(o) || {};
    const w = d.Width | 0, h = d.Height | 0;
    if (!w || !h || w * h > 80e6) return null;
    const s = await doc.streamOf(o, true);
    if (!s) return null;
    if (s.filter === 'DCTDecode' || s.filter === 'DCT') return createImageBitmap(new Blob([s.data], { type: 'image/jpeg' }));
    if (s.filter) throw new Error(s.filter === 'JPXDecode' ? 'скан в формате JPEG 2000 (браузер его не открывает)' : 'скан в формате ' + s.filter.replace(/Decode$/, ''));
    const mask = d.ImageMask === true;
    const bpc = mask ? 1 : (d.BitsPerComponent | 0) || 8;
    const cs = doc.get(d.ColorSpace);
    const csName = x => (x && x.name) || (Array.isArray(x) && x[0] && x[0].name) || '';
    const comps = x => { const nm = csName(x); if (/RGB|^Lab$/.test(nm)) return 3; if (/CMYK/.test(nm)) return 4; if (nm === 'ICCBased') { const st = doc.get(x[1]); return (st && st.N) || 3; } return 1; };
    let n = mask ? 1 : comps(cs), lut = null, lutN = 0;
    if (!mask && csName(cs) === 'Indexed') {
      lutN = comps(doc.get(cs[1]));
      const L = cs[3], lo = L && L.ref != null ? doc.objs.get(L.ref) : null;
      if (lo && lo.data) lut = await doc.streamOf(lo);
      else { const v = doc.get(L); if (v && v.str != null) lut = Uint8Array.from(v.str, ch => ch.charCodeAt(0)); }
      if (!lut) return null;
      n = 1;
    }
    const dec = doc.get(d.Decode), inv = Array.isArray(dec) && dec[0] === 1;
    const data = s.data, rowBytes = Math.ceil(w * n * bpc / 8), max = (1 << Math.min(bpc, 8)) - 1;
    if (data.length < rowBytes * h) return null;
    const img = new ImageData(w, h), px = img.data;
    const at = (y, i) => {                   /* i-й отсчёт строки y: целое 0..max */
      if (bpc === 8) return data[y * rowBytes + i];
      if (bpc === 16) return data[y * rowBytes + 2 * i];
      const bit = i * bpc;
      return (data[y * rowBytes + (bit >> 3)] >> (8 - bpc - (bit & 7))) & max;
    };
    const k8 = v => (bpc >= 8 ? v : Math.round(v * 255 / max));
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const p = (y * w + x) * 4;
        let r, g, b, a = 255;
        if (mask) { r = g = b = 0; if (at(y, x) !== (inv ? 1 : 0)) a = 0; }       /* трафарет: закрашено только отмеченное */
        else if (lut) {
          const v = at(y, x) * lutN;
          if (lutN === 1) r = g = b = lut[v];
          else if (lutN === 4) { const kk = lut[v + 3]; r = 255 - Math.min(255, lut[v] + kk); g = 255 - Math.min(255, lut[v + 1] + kk); b = 255 - Math.min(255, lut[v + 2] + kk); }
          else { r = lut[v]; g = lut[v + 1]; b = lut[v + 2]; }
        } else if (n === 1) { const v = k8(at(y, x)); r = g = b = inv ? 255 - v : v; }
        else if (n === 4) {
          const kk = k8(at(y, 4 * x + 3));
          r = 255 - Math.min(255, k8(at(y, 4 * x)) + kk); g = 255 - Math.min(255, k8(at(y, 4 * x + 1)) + kk); b = 255 - Math.min(255, k8(at(y, 4 * x + 2)) + kk);
        } else { r = k8(at(y, n * x)); g = k8(at(y, n * x + 1)); b = k8(at(y, n * x + 2)); }
        px[p] = r; px[p + 1] = g; px[p + 2] = b; px[p + 3] = a;
      }
    }
    const c = canvasOf(w, h);
    ctx2d(c).putImageData(img, 0, 0);
    return c;
  }
  function rotated(src, deg) {                 /* поворот по часовой стрелке на 90, 180 или 270 градусов */
    const w = src.width, h = src.height, swap = deg === 90 || deg === 270;
    const c = canvasOf(swap ? h : w, swap ? w : h), g = ctx2d(c);
    g.translate(c.width / 2, c.height / 2);
    g.rotate(deg * Math.PI / 180);
    g.drawImage(src, -w / 2, -h / 2);
    return c;
  }
  /* Страница PDF картинкой: все картинки страницы на своих местах (матрицы cm), с поворотом /Rotate.
     null: картинок на странице нет (обычный текстовый PDF). */
  async function pdfPageImage(doc, pg, dpi) {
    const inh = key => { let node = pg; for (let i = 0; node && i < 32; i++) { if (node[key] != null) return doc.get(node[key]); node = doc.get(node.Parent); } return null; };
    const box = [].concat(inh('CropBox') || inh('MediaBox') || [0, 0, 595, 842]).map(v => +doc.get(v) || 0);
    const x0 = Math.min(box[0], box[2]), x1 = Math.max(box[0], box[2]), y0 = Math.min(box[1], box[3]), y1 = Math.max(box[1], box[3]);
    if (x1 - x0 < 10 || y1 - y0 < 10) return null;
    const draws = [];
    const scan = async (content, res, ctm, depth) => {
      const xo = doc.get((res || {}).XObject) || {};
      const stack = [], n = content.length;
      let m = ctm, nums = [], name = null, i = 0;
      while (i < n) {
        const c = content[i];
        if (c === ' ' || c === '\n' || c === '\r' || c === '\t' || c === '\0' || c === '\f') { i++; continue; }
        if (c === '%') { while (i < n && content[i] !== '\n' && content[i] !== '\r') i++; continue; }
        if (c === '(') { i = pdfString(content, i).end; nums = []; continue; }
        if (c === '<') { if (content[i + 1] === '<') { i += 2; continue; } const j = content.indexOf('>', i); i = j < 0 ? n : j + 1; continue; }
        if (c === '>' || c === '[' || c === ']' || c === '{' || c === '}') { i++; continue; }
        if (c === '/') { let j = i + 1; while (j < n && !/[\s\0/<>\[\]()%{}]/.test(content[j])) j++; name = content.slice(i + 1, j); i = j; continue; }
        const num = /^[+-]?(?:\d+\.?\d*|\.\d+)/.exec(content.slice(i, i + 32));
        if (num) { nums.push(+num[0]); i += num[0].length; continue; }
        let j = i;
        while (j < n && /[A-Za-z*'"]/.test(content[j])) j++;
        if (j === i) { i++; continue; }
        const op = content.slice(i, j);
        i = j;
        if (op === 'BI') { const e = content.indexOf('EI', i); i = e < 0 ? n : e + 2; }
        else if (op === 'q') stack.push(m);
        else if (op === 'Q') { if (stack.length) m = stack.pop(); }
        else if (op === 'cm' && nums.length >= 6) m = mulM(nums.slice(-6), m);
        else if (op === 'Do' && name && xo[name] && xo[name].ref != null) {
          const o = doc.objs.get(xo[name].ref), d = o ? doc.dictOf(o) : null, sub = d && d.Subtype && d.Subtype.name;
          if (sub === 'Image') draws.push({ o, m });
          else if (sub === 'Form' && depth < 4) {
            const fm = [].concat(doc.get(d.Matrix) || [1, 0, 0, 1, 0, 0]).map(v => +doc.get(v) || 0);
            await scan(latin1(await doc.streamOf(o)), doc.get(d.Resources) || res, mulM(fm, m), depth + 1);
          }
        }
        nums = [];
        name = null;
      }
    };
    let cs = pg.Contents;
    if (cs && cs.ref != null && doc.objs.has(cs.ref) && !doc.objs.get(cs.ref).data) cs = doc.get(cs);
    const parts = [];
    for (const c of [].concat(cs || [])) if (c && c.ref != null && doc.objs.has(c.ref)) parts.push(latin1(await doc.streamOf(doc.objs.get(c.ref))));
    await scan(parts.join('\n'), inh('Resources'), [1, 0, 0, 1, 0, 0], 0);
    if (!draws.length) return null;
    let k = dpi / 72;
    if ((x1 - x0) * (y1 - y0) * k * k > 14e6) k = Math.sqrt(14e6 / ((x1 - x0) * (y1 - y0)));
    const W = Math.round((x1 - x0) * k), H = Math.round((y1 - y0) * k);
    const page = canvasOf(W, H), g = ctx2d(page);
    g.fillStyle = '#fff';
    g.fillRect(0, 0, W, H);
    g.imageSmoothingEnabled = true;
    g.imageSmoothingQuality = 'high';
    const dev = [k, 0, 0, -k, -x0 * k, y1 * k];
    let drawn = 0, err = null;
    for (const dr of draws) {
      let im;
      try { im = await pdfImage(doc, dr.o); } catch (e) { err = e; continue; }
      if (!im) continue;
      const t = mulM(mulM([1 / im.width, 0, 0, -1 / im.height, 0, 1], dr.m), dev);
      if (Math.abs(t[0] * im.width) + Math.abs(t[2] * im.height) > 4) {              /* крошечная картинка (значок): пропустить */
        g.setTransform(t[0], t[1], t[2], t[3], t[4], t[5]);
        g.drawImage(im, 0, 0);
        g.setTransform(1, 0, 0, 1, 0, 0);
        drawn++;
      }
      if (im.close) im.close();
    }
    if (!drawn) { if (err) throw err; return null; }
    const rot = ((Math.round(+inh('Rotate') || 0) % 360) + 360) % 360;
    const out = rot ? rotated(page, rot) : page;
    if (err) out.partial = (err && err.message) || String(err);           /* часть картинок не прочиталась (например, слой текста CCITT) */
    return out;
  }
  /* Файл-картинка (JPEG, PNG): считается листом A4, приводится к OCR_DPI */
  async function imageFileCanvas(bytes, mime) {
    const bm = await createImageBitmap(new Blob([bytes], { type: mime }));
    const inch = bm.height >= bm.width ? 8.27 : 11.69;
    let k = OCR_DPI / (bm.width / inch);
    if (bm.width * bm.height * k * k > 14e6) k = Math.sqrt(14e6 / (bm.width * bm.height));
    const W = Math.max(1, Math.round(bm.width * k)), H = Math.max(1, Math.round(bm.height * k));
    const c = canvasOf(W, H), g = ctx2d(c);
    g.fillStyle = '#fff';
    g.fillRect(0, 0, W, H);
    g.imageSmoothingQuality = 'high';
    g.drawImage(bm, 0, 0, W, H);
    if (bm.close) bm.close();
    return c;
  }
  function grayOf(c) {
    const w = c.width, h = c.height, d = ctx2d(c).getImageData(0, 0, w, h).data, g = new Uint8Array(w * h);
    for (let i = 0, j = 0; j < g.length; i += 4, j++) g[j] = (d[i] * 77 + d[i + 1] * 151 + d[i + 2] * 28) >> 8;
    return { w, h, g };
  }
  /* Связные пятна чёрных точек (8 соседей) по отрезкам строк: { rY, rA, rB, C: [{ x0, y0, x1, y1, n, runs }] } */
  function componentsOf(ink, w, h) {
    const rY = [], rA = [], rB = [], par = [];
    const find = a => { while (par[a] !== a) { par[a] = par[par[a]]; a = par[a]; } return a; };
    let prev0 = 0, prev1 = 0;
    for (let y = 0; y < h; y++) {
      const cur0 = rY.length, o = y * w;
      let p = prev0;
      for (let x = 0; x < w;) {
        if (!ink[o + x]) { x++; continue; }
        const a = x;
        while (x < w && ink[o + x]) x++;
        const b = x - 1, id = rY.length;
        rY.push(y); rA.push(a); rB.push(b); par.push(id);
        while (p < prev1 && rB[p] < a - 1) p++;
        for (let q = p; q < prev1 && rA[q] <= b + 1; q++) {
          const ra = find(q), rb = find(id);
          if (ra !== rb) { if (ra < rb) par[rb] = ra; else par[ra] = rb; }
        }
      }
      prev0 = cur0;
      prev1 = rY.length;
    }
    const C = [], at = new Map();
    for (let r = 0; r < rY.length; r++) {
      const root = find(r);
      let k = at.get(root);
      if (k === undefined) { k = C.length; at.set(root, k); C.push({ x0: rA[r], x1: rB[r], y0: rY[r], y1: rY[r], n: 0, runs: [] }); }
      const c = C[k];
      if (rA[r] < c.x0) c.x0 = rA[r];
      if (rB[r] > c.x1) c.x1 = rB[r];
      if (rY[r] > c.y1) c.y1 = rY[r];
      c.n += rB[r] - rA[r] + 1;
      c.runs.push(r);
    }
    return { rY, rA, rB, C };
  }
  /* Образ знака: насколько темна каждая клетка сетки OCR_GW × OCR_GH (знак растянут на всю сетку), по оттенкам серого:
     так мелкий и размытый скан сравнивается с образцом лучше, чем по чёрно-белой картинке. Без среднего, длина 1. */
  function featOf(G, x0, y0, x1, y1) {
    const { w, h, g } = G, cw = x1 - x0 + 1, ch = y1 - y0 + 1, f = new Float32Array(OCR_GW * OCR_GH);
    let lo = 255, hi = 0;
    for (let y = Math.max(0, y0 - 2); y <= Math.min(h - 1, y1 + 2); y++) {
      for (let x = Math.max(0, x0 - 2); x <= Math.min(w - 1, x1 + 2); x++) { const v = g[y * w + x]; if (v < lo) lo = v; if (v > hi) hi = v; }
    }
    const span = Math.max(40, hi - lo), sx = OCR_GW / cw, sy = OCR_GH / ch;
    for (let y = y0; y <= y1; y++) {
      const gy0 = (y - y0) * sy, gy1 = gy0 + sy;
      for (let x = x0; x <= x1; x++) {
        const ink = Math.min(1, (hi - g[y * w + x]) / span);
        if (ink <= 0.05) continue;
        const gx0 = (x - x0) * sx, gx1 = gx0 + sx;
        for (let gy = Math.floor(gy0); gy < gy1 && gy < OCR_GH; gy++) {
          const wy = Math.min(gy1, gy + 1) - Math.max(gy0, gy);
          if (wy <= 0) continue;
          for (let gx = Math.floor(gx0); gx < gx1 && gx < OCR_GW; gx++) {
            const wx = Math.min(gx1, gx + 1) - Math.max(gx0, gx);
            if (wx > 0) f[gy * OCR_GW + gx] += ink * wx * wy;
          }
        }
      }
    }
    let mean = 0;
    for (const v of f) mean += v;
    mean /= f.length;
    let ss = 0;
    for (let i = 0; i < f.length; i++) { f[i] -= mean; ss += f[i] * f[i]; }
    ss = Math.sqrt(ss) || 1;
    for (let i = 0; i < f.length; i++) f[i] /= ss;
    return f;
  }
  /* Образцы цифр со сканов: средний образ каждой цифры по 1356 цифрам из 104 настоящих приказов ДГД (ИИН, прочитанные
     уверенно), сетка OCR_GW × OCR_GH, по 216 чисел от -127 до 127 на цифру. По среднему восстановить чей-то ИИН нельзя. */
  const SCAN_DIGITS = 'yMnN5gwhIg/pzsnIyMzwMjAHAykx8s7Jyuc5P/HR0Os5O+vL1BlXF9TKytIRVh3X6kNU+M3JyczyUUfvBVhK6MrIyMrjQ1sLH2E+38nIyMnaNWIkMWMx2MnIyMnXL2Q1OWMs1sjIyMjXLmU5OmMq1cjIyMjYL2Q5MmIs1snIx8nZMmQxIGA33MnIx8ngPmIeCFlF5crIyMrqSlkG7UZR9czJyc36VUXs1RtWFNTLy9YaVxrVy+k6O+7S0/RAOejLyc3yMy4IDDMw8M3JyMnO6REjIQvlzcnI0NLX4fILKRvw1M7N5foSLUdZYTz819DO8QIMH0VkZz361tDO1tjZ7CFaZzz51c/Ozs/R4RVWZzv41c/Ozc7R4BRWZjv31M/Nzc7Q4BRWZzr31M/Nzc7Q4BRXZzr31M/Nzc7Q4BVXZzr31M/Nzc7Q4BVXZjr31M/Nzc7Q4BVXZjn21M/Nzc7Q4BZXZjn21M/Nzc7R4RZXZjj21M7Nzc7R4hdYZjj11M/Ozc/S4xhZZjf21dDOz9HV6iNdZz//2tPQ3ur6GEVhZFIpBPHk7AMRHCUpKSchGAz30dPjDDVIPSD019HR0+UvY2prcnBOBdfR3iNXMQ4KKl93WfrU+D4P4tfW4A5jdyfcAhbc09LS0t8xdjzh39/T0dHR0dYYbzTf0tHR0dHR0dcfaRTY0dHR0dHR0uA8Ve/T0dHR0dHS1f9VHdnR0dHR0dHS5DU56dLR0dHR0dLcG0H11NHR0dHR0tkHOwDW0dHR0dHS2AI0/9fS0dHR0dLX+zAD19LR0tTX0tgALgTc19fX3PcJ2/9NRRwXFhYYJUUfD1Z0cGxsbGxsb1n+EjM3NS4tLy8vMhng0NPjBy5CQzQR6dTP0+4rVFdaZGxjN/LS6Co9Gf79FURuaije9Bf12tPT2vtNcUDm1trT0NDQ0eAvZy7g0NDPz8/P0eMzSgLVz8/Pzs/Q2AI5CdrQz8/P0NjoEUk79dfQz8/P3Q87YXJsRwTXz8/P1On+IEtuc0/6z8/Pz9DS2O8ybW0kz8/Pz8/P0NT0S3I7z8/Pzs7P0NHgLGo5z8/Pz8/P0NHfKWAd0tfU0dDQ0dX0Qkb0+RsO8t7a3vItRQrYIV5hTjIkJzYt/NnQ7h03OjYrFv7j09DP0M/Pz9DQ1PQuBdXRz8/P0NDS5TFlHNfRz8/P0NHaG2JtHNfRz8/Q0dYAPVFtHNfRz9DR0/I3EitsG9fR0NDS4Cgn6CNrGtfR0NHZEjfx2yJrGtfS0dP7PAXZ2SJqGtfS0uYyHt3T2CFpGdfS3hov6tTR1yBpF9fSBT0E4Nva4SVqHeDZOVk2LSwrLlJxTS8VP2BcW1lYWmt1a1w19gUDAP7+BEBxOgPx09PT09LT2yZrHNnT0dHR0NDR2SVrG9fS0NDQ0M/Q2SZnGdbSz8/Qz8/P1Qg2/9TRzs7O1QMqMC4vMTANzs7R70VmZ2doaVkRzs/eIz0jGhocHQzhz9QBOf/d2NfX19TP0OgwM/Hc09DOzs3N2BFcX0YsCefUzs3N4yxcaW9vZUoT4tDN1Oj7DCVEY29kLevRztDR09flCUJsaSvhzs7Ozs/Q1e0zalkDzs7Nzc3Oz9TzR2cizc3Nzc3NztDcIWEuzs3Nzc7OztDYEVUezs/Pzs/Pz9HfIUf92uLc1NHQ0toBPR7eFj4o/ujj6wg3I+nQIF9nV0A3OTML39DP5hIvODEhBOnVz87Oy8vLy8zO2/MLGRLly8vLzNP3Kzcc/OjUy8zN2xNMMfPVzszLzMzZHFMv6NDMy8vLzNUVWzjoz8zMy8vLzvhQVADRzc3My8vL3C5mMurr+vvr1czL9lJnOB8sP1JNG+DNFWJgIvTh5xRVXyffL2lF5c7MzdkeZlsKOmk02czMzM7yVGcuOGgx2MzMzMzfQGk9JmY63MzLy8zYMmg3CV5N6M3My8zaNGQe5j9eDNPMzM7oSE320QRSP+/Rz9oUUhnXzNQHRjgJ/x5BGNzNy8zU9x0qJRr82M3M3gY9Q0A/Pz8+P0Ag7UF/enZ2dnZ3f3w1D15PKiEfHh8pZnIRKir14+De3+L7XlDz/+zd2tnY2uAUbibi29rZ2NjY2uk9ZQHc2NjY2NjY2wJkQOfa2NjY2NjZ4CVuFt3Z2NjY2Nna7U5e89rZ2NjY2NjdCmo249jY2NjY2NnlNGkL3NjX2NjY2Nv1W07u2djY2NjY2eAYcCPf2djY2NjY2upDZv7b2NjY2NjZ3QNmQObZ2NjY2Nja5ChvF93Y2NjY2Nja8kxZ9trY2NjY2Nja8C4e4tnY2NjYvsPY+RIaGhH31sO+w+YgLxD8/hU3JefD3y039c7GxtACSSzY/0sk18LAwMPeOU3uCFQ14MPAv8LeO0rq8kxZHNbDwMgATCPS0RtZWyPjz+81JuLCwdkbV19AJSoH1cK/vsHUGF1lYjzpxsC+vsTiIDQ5WmE++MrAxe0uJ+HcEExjSQLK5ThI9sbBy/lDY0jxFls13cK+v8bzRmAfLl8u2cK/vsDNFlowHl075sXAv8DKC1Ab7UFSFtnHxMjjKTrtx/UyRCIA9wUmKPLHvsXdARoaGBT/3MW/y8zX/BwmKRv01MvKzNgXQSAADTZDCdXL1hdSFdrQ0u88UQfS9k9G6c3Ly9IJW0LpH2Uy2szLy8zoSGAPOWot2czKysvaNWcwPmo53svKysvVKmlEMGlP783LysvWLWpFDl5iGNbMy83gPGo74itiUhHn4vEXWWYfzuYkUFI6KhsvZVj+y83Y8AMA7ussZTXiysrMzc7O0f1SVADRysrLy8zO5jNcHNrMysrLzM/iJlIk4c3LysvN0+wlSBvhzcvK1OX4FDAr/NbMy8rK5A4aEfrf0MzLy8rK';
  const SCAN_AR = [0.607, 0.37, 0.63, 0.556, 0.643, 0.556, 0.607, 0.607, 0.556, 0.609];
  /* Образцы цифр: рисуются шрифтами Times New Roman (обычный и жирный) и Arial нужной высоты, как на скане */
  let DIGITS = null;
  function digitTemplates() {
    if (DIGITS) return DIGITS;
    DIGITS = [];
    const fonts = [['', '"Times New Roman", Times, serif', 0.66], ['bold ', '"Times New Roman", Times, serif', 0.66], ['', 'Arial, Helvetica, sans-serif', 0.72]];
    for (const px of OCR_SIZES) {
      for (const [weight, family, k] of fonts) {
        for (let d = 0; d < 10; d++) {
          const c = canvasOf(px * 3, px * 3), x = ctx2d(c);
          x.fillStyle = '#fff';
          x.fillRect(0, 0, c.width, c.height);
          x.fillStyle = '#000';
          x.font = weight + Math.round(px / k) + 'px ' + family;
          x.fillText(String(d), px * 0.5, px * 2);
          const G = grayOf(c);
          let x0 = G.w, y0 = G.h, x1 = -1, y1 = -1;
          for (let y = 0; y < G.h; y++) for (let xx = 0; xx < G.w; xx++) if (G.g[y * G.w + xx] < 160) { if (xx < x0) x0 = xx; if (xx > x1) x1 = xx; if (y < y0) y0 = y; if (y > y1) y1 = y; }
          if (x1 < 0) continue;
          DIGITS.push({ d, px, f: featOf(G, x0, y0, x1, y1), ar: (x1 - x0 + 1) / (y1 - y0 + 1) });
        }
      }
    }
    const raw = Uint8Array.from(atob(SCAN_DIGITS), c => c.charCodeAt(0)), n = OCR_GW * OCR_GH;
    for (let d = 0; d < 10; d++) {
      const f = new Float32Array(n);
      let ss = 0;
      for (let i = 0; i < n; i++) { const v = raw[d * n + i]; f[i] = v > 127 ? v - 256 : v; ss += f[i] * f[i]; }
      ss = Math.sqrt(ss) || 1;
      for (let i = 0; i < n; i++) f[i] /= ss;
      DIGITS.push({ d, px: 0, f, ar: SCAN_AR[d] });                                /* px 0: годится для знака любой высоты */
    }
    return DIGITS;
  }
  /* ИИН: 7-я цифра 3–6 (родился в 1900–2099), дата рождения существует, возраст от 16 до 100 лет */
  function ocrPlausible(s) {
    const c = +s[6], yy = +s.slice(0, 2), mm = +s.slice(2, 4), dd = +s.slice(4, 6);
    if (c < 3 || c > 6) return false;
    const year = (c <= 4 ? 1900 : 2000) + yy, now = new Date().getFullYear();
    if (year > now - 16 || year < now - 100) return false;
    const dt = new Date(year, mm - 1, dd);
    return dt.getMonth() === mm - 1 && dt.getDate() === dd;
  }
  /* ИИН на картинке страницы. G: { w, h, g } (яркость 0..255). [{ iin, conf, score, fixes, x, y }] по порядку на странице */
  function ocrIins(G, dbg) {
    const { w, h, g } = G;
    /* 1. Чёрно-белая: точка тёмная, если темнее средней яркости окна вокруг на 15% (неровная засветка не мешает) */
    const half = Math.max(7, Math.round(Math.min(w, h) / 100)), W1 = w + 1;
    const I = new Uint32Array(W1 * (h + 1));
    for (let y = 0; y < h; y++) {
      let row = 0;
      const o = y * w, r0 = y * W1, r1 = r0 + W1;
      for (let x = 0; x < w; x++) { row += g[o + x]; I[r1 + x + 1] = I[r0 + x + 1] + row; }
    }
    /* два порога: светлые штрихи (размытый мелкий скан) остаются, только если касаются явно тёмных: цифры не рвутся на куски */
    const ink = new Uint8Array(w * h);
    for (let y = 0; y < h; y++) {
      const ya = Math.max(0, y - half), yb = Math.min(h, y + half + 1), ra = ya * W1, rb = yb * W1, o = y * w;
      for (let x = 0; x < w; x++) {
        const v = g[o + x];
        if (v > 232) continue;
        const xa = Math.max(0, x - half), xb = Math.min(w, x + half + 1);
        const m = (I[rb + xb] - I[ra + xb] - I[rb + xa] + I[ra + xa]) / ((xb - xa) * (yb - ya));
        if (v < m * 0.93) ink[o + x] = v < m * 0.82 && v < 205 ? 2 : 1;
      }
    }
    /* 2. Знаки: пятна высотой с цифру; 3. цепочки соседних знаков одной высоты на одной строке */
    const R = componentsOf(ink, w, h);
    R.C = R.C.filter(c => c.runs.some(r => { const o = R.rY[r] * w; for (let x = R.rA[r]; x <= R.rB[r]; x++) if (ink[o + x] === 2) return true; return false; }));
    /* куски одного знака друг над другом (у «2» отрывается нижняя черта, у «5» верхняя): склеить */
    {
      const C = R.C.filter(c => c.y1 - c.y0 < 100 && c.x1 - c.x0 < 100).sort((a, b) => a.x0 - b.x0);
      const up = new Map(), top = c => { while (up.has(c)) c = up.get(c); return c; };
      for (let i = 0; i < C.length; i++) {
        const a = C[i];
        for (let j = i + 1; j < C.length && C[j].x0 <= a.x1; j++) {
          const b = C[j], ov = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0) + 1;
          if (ov < 0.5 * Math.min(a.x1 - a.x0 + 1, b.x1 - b.x0 + 1)) continue;
          const gap = Math.max(a.y0, b.y0) - Math.min(a.y1, b.y1) - 1, H = Math.max(a.y1, b.y1) - Math.min(a.y0, b.y0) + 1;
          if (gap > Math.max(2, 0.12 * H) || H > 1.6 * Math.max(a.y1 - a.y0 + 1, b.y1 - b.y0 + 1)) continue;
          const ra = top(a), rb = top(b);
          if (ra !== rb) up.set(rb, ra);
        }
      }
      if (up.size) {
        for (const c of C) {
          const r = top(c);
          if (r === c) continue;
          r.x0 = Math.min(r.x0, c.x0); r.x1 = Math.max(r.x1, c.x1); r.y0 = Math.min(r.y0, c.y0); r.y1 = Math.max(r.y1, c.y1);
          r.n += c.n; r.runs = r.runs.concat(c.runs); c.gone = true;
        }
        R.C = R.C.filter(c => !c.gone);
      }
    }
    const cand = [];
    for (const c of R.C) {
      const ch = c.y1 - c.y0 + 1, cw = c.x1 - c.x0 + 1;
      if (ch < 9 || ch > 100 || cw < 2 || c.n < ch) continue;
      if (cw <= ch * 1.05) { cand.push(c); continue; }
      if (cw > ch * 2.9) continue;
      /* слипшиеся цифры (мелкий скан): пятно делится на k частей по самым светлым столбцам */
      const k = Math.max(2, Math.round(cw / (ch * 0.62))), prof = new Float32Array(cw);
      for (const r of c.runs) for (let x = R.rA[r]; x <= R.rB[r]; x++) prof[x - c.x0]++;
      const bounds = [0];
      for (let i = 1; i < k; i++) {
        const mid = Math.round(i * cw / k), rad = Math.max(1, Math.round(cw / k * 0.3));
        let best = mid;
        for (let x = Math.max(1, mid - rad); x <= Math.min(cw - 2, mid + rad); x++) if (prof[x] < prof[best]) best = x;
        bounds.push(best);
      }
      bounds.push(cw);
      for (let i = 0; i < k; i++) {
        const xa = c.x0 + bounds[i], xb = c.x0 + bounds[i + 1] - 1;
        const sub = { x0: Infinity, x1: -1, y0: Infinity, y1: -1, n: 0, runs: [] };
        for (const r of c.runs) {
          const a = Math.max(R.rA[r], xa), b = Math.min(R.rB[r], xb);
          if (a > b) continue;
          sub.n += b - a + 1;
          if (a < sub.x0) sub.x0 = a;
          if (b > sub.x1) sub.x1 = b;
          if (R.rY[r] < sub.y0) sub.y0 = R.rY[r];
          if (R.rY[r] > sub.y1) sub.y1 = R.rY[r];
        }
        if (sub.n && sub.y1 - sub.y0 + 1 >= ch * 0.6) cand.push(sub);
      }
    }
    cand.sort((a, b) => a.x0 - b.x0);
    if (dbg && dbg.box) {
      const [bx0, by0, bx1, by1] = dbg.box;
      dbg.glyphs = R.C.filter(c => c.x1 >= bx0 && c.x0 <= bx1 && c.y1 >= by0 && c.y0 <= by1).map(c => [c.x0, c.y0, c.x1 - c.x0 + 1, c.y1 - c.y0 + 1, c.n, cand.includes(c) ? 1 : 0]);
      dbg.cand = cand.filter(c => c.x1 >= bx0 && c.x0 <= bx1 && c.y1 >= by0 && c.y0 <= by1).map(c => [c.x0, c.y0, c.x1 - c.x0 + 1, c.y1 - c.y0 + 1]);
    }
    for (let i = 0; i < cand.length; i++) {
      const a = cand[i], ah = a.y1 - a.y0 + 1;
      let best = null, bestGap = 1e9;
      for (let j = i + 1; j < cand.length; j++) {
        const b = cand[j];
        if (b.x0 > a.x1 + ah) break;
        const bh = b.y1 - b.y0 + 1, hh = Math.max(ah, bh);
        if (bh * 1.3 < ah || ah * 1.3 < bh || b.x1 <= a.x1) continue;
        if (Math.abs(b.y1 - a.y1) > hh * 0.2 || Math.abs(b.y0 - a.y0) > hh * 0.25) continue;
        const gap = b.x0 - a.x1 - 1;
        if (gap < -hh * 0.2) continue;
        if (gap < bestGap) { bestGap = gap; best = b; }
      }
      if (best && (!best.prev || best.prevGap > bestGap)) {
        if (best.prev) best.prev.next = null;
        best.prev = a;
        best.prevGap = bestGap;
        a.next = best;
      }
    }
    const T = digitTemplates();
    const classify = c => {
      if (c.cls) return c.cls;
      const f = featOf(G, c.x0, c.y0, c.x1, c.y1), ch = c.y1 - c.y0 + 1, ar = (c.x1 - c.x0 + 1) / ch;
      const px = OCR_SIZES.reduce((a, s) => (Math.abs(s - ch) < Math.abs(a - ch) ? s : a), OCR_SIZES[0]);
      const best = new Array(10).fill(-9);
      for (const t of T) {
        if (t.px !== px && t.px !== 0) continue;
        let s = 0;
        for (let i = 0; i < f.length; i++) s += f[i] * t.f[i];
        s -= 0.5 * Math.abs(Math.log(ar / t.ar));
        if (s > best[t.d]) best[t.d] = s;
      }
      c.cls = best.map((s, d) => ({ d, s })).sort((a, b) => b.s - a.s);
      return c.cls;
    };
    /* 4. Окна по 12 знаков: все похожи на цифры, шаг ровный; ИИН с верной контрольной цифрой и настоящей датой.
       Не сошлось: одна самая сомнительная цифра заменяется на вторую по сходству (только если та почти так же похожа);
       такой ИИН помечается «сверить с приказом». Две замены не делаются: на настоящих приказах они давали чужие ИИН. */
    /* штрихи, черта, буквы «І», «l» читаются как ряд единиц: в настоящем ИИН одна цифра не стоит 5 раз подряд и единиц не больше 6 */
    const valid = s => iinOk(s) && ocrPlausible(s) && !/(\d)\1{4}/.test(s) && (s.match(/1/g) || []).length <= 6;
    const found = new Map();
    for (const start of cand) {
      if (start.prev) continue;
      const chain = [];
      for (let x = start; x && chain.length < 80; x = x.next) chain.push(x);
      if (dbg && chain.length >= 8) dbg.push({ x: chain[0].x0, y: chain[0].y0, h: chain[0].y1 - chain[0].y0 + 1, n: chain.length,
        str: chain.map(c => { const k = classify(c); return k[0].s >= 0.3 ? String(k[0].d) : '?'; }).join(''),
        s: chain.map(c => Math.round(classify(c)[0].s * 100)), top: chain.map(c => classify(c).slice(0, 3).map(x => x.d + ':' + Math.round(x.s * 100)).join(' ')),
        box: chain.map(c => [c.x0, c.y0, c.x1 - c.x0 + 1, c.y1 - c.y0 + 1]) });
      if (chain.length < 12) continue;
      for (let s0 = 0; s0 + 12 <= chain.length; s0++) {
        const glyphs = chain.slice(s0, s0 + 12);
        const cx = glyphs.map(c => (c.x0 + c.x1) / 2), steps = cx.slice(1).map((v, i) => v - cx[i]);
        const med = steps.slice().sort((a, b) => a - b)[5];
        if (steps.some(v => v > med * 1.9 || v < med * 0.45)) continue;           /* пробел внутри: это не одно число */
        const win = glyphs.map(classify);
        if (win.some(x => x[0].s < 0.3)) continue;
        const base = win.map(x => x[0].d).join(''), sum = win.reduce((a, x) => a + x[0].s, 0);
        let best = valid(base) ? { iin: base, score: sum / 12, fixes: 0 } : null;
        if (!best) {
          const doubt = win.map((x, i) => [x[0].s - x[1].s, i]).filter(x => x[0] < 0.2).sort((a, b) => a[0] - b[0]).slice(0, 4).map(x => x[1]);
          const alts = p => win[p].slice(1, 3).filter(a => a.s >= win[p][0].s - 0.2);
          for (const p of doubt) for (const a of alts(p)) {
            const t = base.slice(0, p) + a.d + base.slice(p + 1), v = (sum - win[p][0].s + a.s) / 12;
            if (valid(t) && (!best || v > best.score)) best = { iin: t, score: v, fixes: 1 };
          }
        }
        if (!best) continue;
        const low = Math.min(...win.map(x => x[0].s));
        best.conf = best.fixes === 0 && low >= 0.45 ? 'high' : 'medium';     /* проверка на 225 приказах: «high» без единой ошибки */
        best.x = glyphs[0].x0;
        best.y = glyphs[0].y0;
        if (dbg && dbg.learn && best.conf === 'high') dbg.learn.push(glyphs.map((c, i) => ({ d: +best.iin[i], h: c.y1 - c.y0 + 1, ar: (c.x1 - c.x0 + 1) / (c.y1 - c.y0 + 1), f: Array.from(featOf(G, c.x0, c.y0, c.x1, c.y1)) })));
        const old = found.get(best.iin);
        if (!old || old.fixes > best.fixes || (old.fixes === best.fixes && old.score < best.score)) found.set(best.iin, best);
      }
    }
    return Array.from(found.values()).sort((a, b) => (Math.abs(a.y - b.y) > 15 ? a.y - b.y : a.x - b.x));
  }
  /* ИИН из файла приказа: текст PDF, иначе распознавание картинки. { iins: [{ iin, conf, how, fixes }], note, pages } */
  async function readScan(bytes, name) {
    const ext = ((/\.([a-z0-9]+)$/i.exec(name || '') || [])[1] || '').toLowerCase();
    const head = latin1(bytes.subarray(0, 8));
    const isPng = head.startsWith('\x89PNG'), isJpg = head.startsWith('\xff\xd8');
    const isPdf = head.startsWith('%PDF') || (!isPng && !isJpg && ext === 'pdf');
    const out = { iins: [], note: '', pages: 0 };
    const take = (list, how) => { for (const x of list) if (!out.iins.some(y => y.iin === x.iin)) out.iins.push(Object.assign({ how }, x)); };
    try {
      const pages = [];
      if (isPdf) {
        const lines = await pdfLines(bytes);
        const ex = extractIins(lines.map((t, i) => ({ row: i + 1, cells: [{ v: clean(t), num: false }] })).filter(r => r.cells[0].v));
        if (ex.items.length) { take(ex.items.map(x => ({ iin: x.iin, conf: 'high', fixes: 0 })), 'текст PDF'); return out; }
        const doc = await pdfOpen(bytes);
        for (const pg of doc.pages.slice(0, OCR_PAGES)) {
          try { const c = await pdfPageImage(doc, pg, OCR_DPI); if (c) pages.push(c); } catch (e) { out.note = (e && e.message) || String(e); }
        }
        if (!pages.length && !out.note) out.note = 'в PDF нет ни ИИН в тексте, ни картинки страницы';
      } else if (isPng || isJpg || /^(jpe?g|png)$/.test(ext)) {
        pages.push(await imageFileCanvas(bytes, isPng ? 'image/png' : 'image/jpeg'));
      } else { out.note = 'не PDF и не картинка'; return out; }
      out.pages = pages.length;
      for (const c of pages) {
        for (const rot of c.width > c.height * 1.15 ? [90, 270, 0, 180] : [0, 180]) {
          const r = ocrIins(grayOf(rot ? rotated(c, rot) : c));
          if (r.length) { take(r, 'распознано со скана'); break; }
          await sleep(0);
        }
        if (out.iins.length) break;
      }
      const partial = pages.map(c => c.partial).filter(Boolean)[0];
      if (!out.iins.length && pages.length) out.note = 'ИИН на скане не распознан' + (partial ? ': ' + partial : '');
    } catch (e) { out.note = (e && e.message) || String(e); }
    return out;
  }

  /* ИИН из любого файла публикации: Word, Excel, страница (текст) или PDF и картинка (текст PDF, иначе скан) */
  async function readOrderFile(bytes, name, type) {
    const head = latin1(bytes.subarray(0, 64));
    let kind = '';
    if (head.startsWith('PK')) {
      try { const names = zipRead(bytes).map(e => e.name); kind = names.includes('word/document.xml') ? 'docx' : names.some(n => /^xl\//.test(n)) ? 'xlsx' : ''; } catch (e) { kind = ''; }
    } else if (/<(!doctype|html|table)/i.test(head) || /text\/html/i.test(type || '')) kind = 'html';
    else if (/^\s*[\wЀ-ӿ"]/.test(head) && /text\/(plain|csv)/i.test(type || '')) kind = 'txt';
    if (!kind) return readScan(bytes, name);
    const parsed = await fileRows('файл.' + kind, bytes);
    const ex = extractIins(parsed.rows), out = { iins: [], note: '', pages: 0 };
    for (const x of ex.items) if (!out.iins.some(y => y.iin === x.iin)) out.iins.push({ iin: x.iin, conf: 'high', how: 'текст файла', fixes: 0 });
    if (!out.iins.length) out.note = 'ИИН в файле не найден';
    return out;
  }

  /* ====================== Шаг 1: списки ДГД по регионам ====================== */
  /* Списки ДГД (финансовые управляющие) за день кладутся в подпапку «ДГД» рабочей папки: файл на регион, в любом
     из форматов xlsx, csv, docx, html, txt или PDF с текстом. Помощник находит в них ИИН (контрольная цифра и дата
     рождения), определяет регион и дату списка, проверяет каждый регион и пишет свод «ДГД_свод.xlsx». */
  const DGD_DIR = 'ДГД';
  const DGD_FILE = 'ДГД_свод.xlsx';
  const DGD_HEADER = 'ДГД';               /* столбец в рабочем файле: где этот ИИН есть в списках ДГД */
  const DGD_STALE_DAYS = 3;               /* самая свежая дата региона старше стольких дней: проверить */
  const DGD_FEW = 2;                      /* столько ИИН и меньше в регионе: подозрительно мало */
  const L_ = 'а-яёәіңғүұқөһa-z';           /* буквы для границ слов (\b в JavaScript кириллицу не понимает) */
  const rx = s => new RegExp(s.replace(/<</g, '(?<![' + L_ + '])').replace(/>>/g, '(?![' + L_ + '])'), 'i');
  const REGION_LIST = [                   /* порядок важен: «Алматинская» раньше «Алматы» */
    ['Абайская обл.', rx('абайск|абай\\s+облыс|област\\S*\\s+абай|<<семей>>'), rx('<<абай>>')],
    ['Акмолинская обл.', rx('акмолин|ақмол')],
    ['Актюбинская обл.', rx('актюбин|ақтөбе|<<актобе>>')],
    ['Алматинская обл.', rx('алматинск|алматы\\s+облыс')],
    ['Атырауская обл.', rx('атырау')],
    ['Восточно-Казахстанская обл.', rx('восточно[-\\s]*казахстан|шығыс\\s+қазақстан|<<вко>>')],
    ['Жамбылская обл.', rx('жамбыл|<<тараз>>')],
    ['Жетысуская обл.', rx('жетысу|жетісу|<<талдыкорган')],
    ['Западно-Казахстанская обл.', rx('западно[-\\s]*казахстан|батыс\\s+қазақстан|<<зко>>|<<уральск')],
    ['Карагандинская обл.', rx('караганд|қарағанд')],
    ['Костанайская обл.', rx('костанай|қостанай')],
    ['Кызылординская обл.', rx('кызылорд|қызылорд')],
    ['Мангистауская обл.', rx('мангист|маңғыст|<<актау>>')],
    ['Павлодарская обл.', rx('павлодар')],
    ['Северо-Казахстанская обл.', rx('северо[-\\s]*казахстан|солтүстік\\s+қазақстан|<<ско>>|петропавл')],
    ['Туркестанская обл.', rx('туркестан|түркістан')],
    ['Улытауская обл.', rx('улытау|ұлытау|жезказган|жезқазған')],
    ['г. Астана', rx('астан|нур[-\\s]*султан')],
    ['г. Алматы', rx('алматы')],
    ['г. Шымкент', rx('шымкент')],
  ];
  /* loose: имя файла или папки, его задаёт человек; там годится и голое «Абай» (в тексте это может быть улица) */
  function regionsIn(text, loose) {
    const t = String(text || '').toLowerCase(), out = [];
    for (const [name, re, extra] of REGION_LIST) {
      if (name === 'г. Алматы' && out.includes('Алматинская обл.') && !/(?:^|[^а-яё])(?:г\.?\s*|город\S*\s+)алматы/i.test(t)) continue;
      if (re.test(t) || (loose && extra && extra.test(t))) out.push(name);
    }
    return out;
  }
  function datesIn(text) {
    const out = [];
    const t = String(text || '');
    for (const m of t.matchAll(/(?<!\d)(\d{1,2})[.\-_/](\d{1,2})[.\-_/](\d{4})(?!\d)/g)) out.push([+m[3], +m[2], +m[1]]);
    for (const m of t.matchAll(/(?<!\d)(\d{4})[.\-_](\d{2})[.\-_](\d{2})(?!\d)/g)) out.push([+m[1], +m[2], +m[3]]);
    return out.filter(([y, mo, d]) => y >= 2000 && y <= 2100 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31)
      .map(([y, mo, d]) => new Date(y, mo - 1, d)).filter(d => d.getDate() >= 1);
  }
  const maxDate = list => list.reduce((a, d) => (d && (!a || d > a) ? d : a), null);
  const dmy = d => (d ? pad2(d.getDate()) + '.' + pad2(d.getMonth() + 1) + '.' + d.getFullYear() : '');
  /* ИИН человека, а не БИН и не случайное число: месяц и день рождения настоящие, 7-я цифра (век и пол) от 0 до 6 */
  const personalIin = s => { const mo = +s.slice(2, 4), d = +s.slice(4, 6), c = +s[6]; return mo >= 1 && mo <= 12 && d >= 1 && d <= 31 && c <= 6; };

  const UP_ = 'А-ЯЁӘІҢҒҮҰҚӨҺ', LO_ = 'а-яёәіңғүұқөһ';
  const NAME_RE = new RegExp('^[' + UP_ + '][' + UP_ + LO_ + '-]+(?:\\s+[' + UP_ + '](?:[' + UP_ + LO_ + '-]+|\\.)){1,3}$');
  const NAME_IN_TEXT = new RegExp('[' + UP_ + '][' + LO_ + '-]+(?:\\s+[' + UP_ + '][' + LO_ + '-]+){1,3}|[' + UP_ + ']{2,}(?:\\s+[' + UP_ + ']{2,}){1,3}', 'g');

  /* ---------- gov.kz: данные страницы (React Router, поток turbo-stream в тегах script) ---------- */
  /* Страница gov.kz приходит с данными внутри: window.__reactRouterContext.streamController.enqueue("…"). Это плоский
     массив значений, где объекты и массивы ссылаются на другие значения по номеру, а отложенные части (P12:[…])
     дописываются в конец. Возвращает { имя запроса: { endpointName, arg, value } } (getFilteredData, getDocumentById…). */
  function govData(html) {
    let text = '';
    for (const m of String(html).matchAll(/streamController\.enqueue\(("(?:[^"\\]|\\.)*")\)/g)) { try { text += JSON.parse(m[1]); } catch (e) { /* не JSON */ } }
    const lines = text.split('\n').filter(l => l.trim());
    if (!lines.length) return {};
    let vals;
    try { vals = JSON.parse(lines[0]); } catch (e) { return {}; }
    if (!Array.isArray(vals)) return {};
    const promises = {};
    for (const l of lines.slice(1)) {
      const m = /^([PE])(\d+):([\s\S]*)$/.exec(l);
      if (!m) continue;
      let arr;
      try { arr = JSON.parse(m[3]); } catch (e) { continue; }
      promises[m[2]] = vals.length;
      for (const x of [].concat(arr)) vals.push(x);
    }
    const memo = new Map();
    const dec = (i, depth) => {
      if (typeof i !== 'number' || depth > 200) return null;
      if (i < 0) return i === -4 ? 0 : null;
      if (memo.has(i)) return memo.get(i);
      const v = vals[i];
      if (Array.isArray(v)) {
        if (v.length && typeof v[0] === 'string' && v[0].length === 1 && 'DRSMYEPZBU'.includes(v[0])) {
          const r = v[0] === 'P' ? (promises[v[1]] != null ? dec(promises[v[1]], depth + 1) : null) : v[0] === 'D' ? v[1] : null;
          memo.set(i, r);
          return r;
        }
        const r = [];
        memo.set(i, r);
        for (const x of v) r.push(dec(x, depth + 1));
        return r;
      }
      if (v && typeof v === 'object') {
        const r = {};
        memo.set(i, r);
        for (const [k, x] of Object.entries(v)) r[k[0] === '_' ? String(vals[+k.slice(1)]) : k] = dec(x, depth + 1);
        return r;
      }
      memo.set(i, v);
      return v;
    };
    const out = {}, seen = new Set();
    const walk = (o, depth) => {
      if (!o || typeof o !== 'object' || seen.has(o) || depth > 60) return;
      seen.add(o);
      if (!Array.isArray(o) && typeof o.endpointName === 'string' && !out[o.endpointName]) out[o.endpointName] = o;
      for (const x of Array.isArray(o) ? o : Object.values(o)) walk(x, depth + 1);
    };
    walk(dec(0, 0), 0);
    return out;
  }

  /* ---------- ДГД на gov.kz: 20 регионов, публикации о банкротстве граждан ---------- */
  /* Адреса департаментов на gov.kz (проверено 20.11.2026: все 20 открываются, у каждого раздел «Документы»). */
  const GOV_SITES = [['kgd-abay', 'Абайская обл.'], ['kgd-akmola', 'Акмолинская обл.'], ['kgd-aktobe', 'Актюбинская обл.'],
    ['kgd-alm', 'Алматинская обл.'], ['kgd-atyrau', 'Атырауская обл.'], ['kgd-vko', 'Восточно-Казахстанская обл.'],
    ['kgd-zhambyl', 'Жамбылская обл.'], ['kgd-zhetysu', 'Жетысуская обл.'], ['kgd-zko', 'Западно-Казахстанская обл.'],
    ['kgd-karaganda', 'Карагандинская обл.'], ['kgd-kostanay', 'Костанайская обл.'], ['kgd-kyzylorda', 'Кызылординская обл.'],
    ['kgd-mangistau', 'Мангистауская обл.'], ['kgd-pavlodar', 'Павлодарская обл.'], ['kgd-sko', 'Северо-Казахстанская обл.'],
    ['kgd-turkistan', 'Туркестанская обл.'], ['kgd-ulytau', 'Улытауская обл.'], ['kgd-astana', 'г. Астана'],
    ['kgd-almaty', 'г. Алматы'], ['kgd-shymkent', 'г. Шымкент']];
  const GOV = 'https://www.gov.kz';
  const GOV_START = GOV + '/org/kgd-almaty/documents/1?lang=ru';
  const GOV_PAGES = 12;                   /* страниц списка документов (по 50) на регион, не больше */
  const DGD_WINDOW = 31;                  /* публикации за столько дней скачиваются и читаются */
  const DGD_SITE_STALE = 14;              /* последняя публикация региона старше стольких дней: проверить */
  const DGD_META = 'публикации_gov_kz.json';   /* в папке региона: что найдено на gov.kz при последнем сборе */
  const DGD_OCR = 'распознано.json';       /* в папке «ДГД»: ИИН, прочитанные из файлов (чтобы не читать заново) */
  const regionFolder = name => name.replace(/ обл\.$/, ' область');   /* Windows не разрешает точку в конце имени папки */
  /* Публикации о процедурах граждан: приказы о финансовых управляющих, объявления о возбуждении дела (судебное банкротство,
     восстановление платежеспособности), о признании банкротом, реестры требований кредиторов (РТК) по должнику.
     Не берутся: юрлица (ТОО, АО, реабилитация, банкротный и временный управляющий, ликвидация), списки управляющих,
     уведомления управляющих о начале деятельности, конкурсы и прочее. */
  const GOV_DOC_RE = rx('финансов\\S*\\s+управляющ|граждан|физическ\\S*\\s+лиц|физ\\.?\\s*лиц|<<фл>>|платежеспособ|судебн\\S*\\s+банкротств|<<ртк>>|' +
    'реестр\\S*\\s+требовани|борышкер|<<должник|төлем\\s+қабілет|азаматтар');
  /* Название публикации: только ФИО человека (так Костанай называет объявления и реестры по гражданам) */
  const PERSON_TITLE_RE = new RegExp('^[«"]?(?:[' + UP_ + '][' + LO_ + 'ʼ’\'-]+(?:\\s+[' + UP_ + '][' + LO_ + 'ʼ’\'-]+){1,3}|[' + UP_ + ']{2,}(?:\\s+[' + UP_ + ']{2,}){1,3})[»"]?$');
  const isGovCitizenDoc = t => (GOV_DOC_RE.test(t) || PERSON_TITLE_RE.test(clean(t))) && !GOV_NOT_RE.test(t);
  /* Сводная публикация за год («Информационные сообщения ФЛ 2026», «Приказы … 2026 год»): файлы в ней дополняют, поэтому
     она перечитывается при каждом сборе, пока идёт этот год */
  const govYearly = t => new RegExp('(?:^|\\D)' + new Date().getFullYear() + '(?:\\D|$)').test(t) && /сообщени|приказы|хабарлам|бұйрықтар/i.test(t);
  const GOV_NOT_RE = rx('реабилитац|банкротн\\S*\\s+управляющ|временн\\S*\\s+управляющ|<<тоо>>|<<ао>>|товариществ|<<llp>>|ликвидац|оценщик|' +
    'конкурс|ваканс|список\\s+(?:лиц|финансов|администратор)|сведени\\S*\\s+о\\s+направленн|собрани\\S*\\s+кредитор|<<ип>>');
  const UPPER_NAME = '[' + UP_ + ']{2,}(?:[\\s-]+[' + UP_ + ']{2,}){1,3}';
  function govKind(title) {
    const t = String(title || ''), low = t.toLowerCase();
    const proc = /восстановлени\S*\s+платежеспособ|(?:^|\s)вп(?:\s|$)|төлем\s+қабілет/i.test(low) ? 'восстановление платежеспособности'
      : /судебн\S*\s+банкрот|(?:^|\s)сб(?:\s|$)/i.test(low) ? 'судебное банкротство' : '';
    let kind = 'публикация';
    if (/приказ|бұйрық/.test(low)) { const k = orderInfo([t], []).kind; kind = 'приказ' + (k ? ' (' + k + ')' : ''); }
    else if (/(?:^|\s)ртк(?:\s|$)|реестр\S*\s+требовани/.test(low)) kind = 'реестр требований кредиторов';
    else if (/объявлени\S*\s+о\s+возбуждени/.test(low)) kind = 'объявление о возбуждении дела';
    else if (/объявлени\S*\s+о\s+признани/.test(low)) kind = 'объявление о признании банкротом';
    else if (/объявлени\S*\s+о\s+(?:применени|завершени|прекращени)/.test(low)) kind = 'объявление о ' + /о\s+(\S+)/.exec(low.slice(low.indexOf('объявлени')))[1].replace(/и$/, 'и процедуры');
    return proc && kind !== 'публикация' ? kind + ', ' + proc : kind;
  }
  /* ФИО должника из названия публикации: «РТК должник ФАМИЛИЯ ИМЯ», «по заявлению должника …», «… ВП Фамилия И.» */
  function govDebtor(title) {
    const t = clean(title);
    if (PERSON_TITLE_RE.test(t)) return t.replace(/[«»"]/g, '');
    for (const re of [new RegExp('(?:должник\\S*|борышкер)\\s+[«"]?(' + UPPER_NAME + '|' + NAME_ANY + ')'), new RegExp('(?:^|\\s)(?:ВП|СБ)\\s+(' + UPPER_NAME + '|' + NAME_ANY + ')')]) {
      const m = re.exec(t);
      if (m) return clean(m[1]);
    }
    const o = orderInfo([t], []);
    if (o.debtor) return o.debtor;
    const tail = new RegExp('(' + UPPER_NAME + '|' + N1 + '\\s+' + N1 + '(?:\\s+' + N1 + ')?)\\s*$').exec(t.replace(/[.,;)»"]+$/, ''));
    return tail && !/^(Республики|Казахстан)/.test(tail[1]) ? clean(tail[1]) : '';
  }
  const N1 = '[' + UP_ + '][' + LO_ + 'ʼ’\'-]+', INI = '[' + UP_ + ']\\.\\s*(?:[' + UP_ + ']\\.?)?';
  const NAME_ANY = N1 + '\\s+' + INI + '|' + N1 + '(?:\\s+' + N1 + '){1,3}';
  const NAME_ONLY = new RegExp('^[«"]?(?:' + NAME_ANY + ')[»"]?$');
  const AFTER = (lead, flags) => new RegExp(lead + '\\s*[«"]?(' + NAME_ANY + ')', flags || '');
  const DEBTOR_RES = [AFTER('[Пп]о\\s+заявлени\\S*\\s+(?:должник\\S*\\s+|гражданин\\S*\\s+|гр\\.\\s*|ФЛ\\s+)?'),
    AFTER('[Вв]\\s+отношении\\s+(?:ФЛ\\s+|физическ\\S+\\s+лиц\\S*\\s+|гражданин\\S*\\s+|должник\\S*\\s+)?'), AFTER('(?:^|\\s)ФЛ\\s+')];
  const MANAGER_RE = AFTER('финансов\\S*\\s+управляющ\\S*\\s+');
  function orderInfo(cells, files) {
    const today = new Date();
    const text = cells.join(' '), dates = [];
    for (const m of text.matchAll(/(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})(?!\d)/g)) {
      const y = m[3].length === 2 ? 2000 + +m[3] : +m[3], d = new Date(y, +m[2] - 1, +m[1]);
      if (d.getMonth() === +m[2] - 1 && d.getDate() === +m[1] && y >= 2020 && y <= 2100) dates.push(d);
    }
    const ok = dates.filter(d => d <= new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1));
    const date = ok.length ? maxDate(ok) : null;
    const low = (text + ' ' + files.map(f => f.name).join(' ')).toLowerCase();
    const kind = /отстран|shettet|шеттет/.test(low) ? 'отстранение' : /освобожд|bosat|босат/.test(low) ? 'освобождение'
      : /назнач|tagay|тағай/.test(low) ? 'назначение' : '';
    let debtor = '', manager = '';
    for (const c of cells) {
      if (!debtor) for (const re of DEBTOR_RES) { const m = re.exec(c); if (m) { debtor = clean(m[1]); break; } }
      if (!manager) { const m = MANAGER_RE.exec(c); if (m) manager = clean(m[1]); }
    }
    for (const c of cells) {                    /* ячейка только с именем: должник, а если он уже есть, то управляющий */
      const v = c.replace(/[«»"]/g, '').trim();
      if (!NAME_ONLY.test(v) || v === debtor || v === manager) continue;
      if (!debtor) debtor = v; else if (!manager) manager = v;
    }
    const num = /(?:№\s*|(?:^|\s)П-\s*)(\d[\d/-]*[А-Яа-я]?)/.exec(text);
    return { date, kind, debtor, manager, number: num ? num[1] : '', future: dates.length > ok.length };
  }

  /* Строки файла: [{ sheet, row, cells: [{ v, num }] }] */
  async function fileRows(name, bytes) {
    const ext = ((/\.([a-z0-9]+)$/i.exec(name) || [])[1] || '').toLowerCase();
    const textRows = lines => lines.map((t, i) => ({ row: i + 1, cells: [{ v: clean(t), num: false }] })).filter(r => r.cells[0].v);
    if (ext === 'xlsx' || ext === 'xlsm') {
      const b = await openBook(bytes, name), rows = [];
      for (const sh of b.sheets) {
        let S2;
        try { S2 = await loadSheet(b, sh); } catch (e) { continue; }
        for (const r of Array.from(S2.cells.keys()).sort((a, c) => a - c)) {
          const map = S2.cells.get(r), cells = [];
          for (const c of Array.from(map.keys()).sort((a, z) => a - z)) {
            const x = cellValue(S2, map.get(c));
            if (x && x.v != null && clean(x.v)) cells[c - 1] = { v: clean(x.v), num: x.num };
          }
          if (cells.length) rows.push({ sheet: sh.name, row: r, cells });
        }
      }
      return { kind: 'Excel', rows };
    }
    if (ext === 'csv') return { kind: 'CSV', rows: parseCsv(decodeText(bytes)).map((r, i) => ({ row: i + 1, cells: r.map(v => ({ v: clean(v), num: false })) })) };
    if (ext === 'txt') return { kind: 'текст', rows: textRows(decodeText(bytes).split(/\r?\n/)) };
    if (ext === 'docx') {
      const entries = zipRead(bytes), e = entries.find(x => x.name === 'word/document.xml');
      if (!e) throw new Error('файл Word повреждён');
      const doc = parseXml(td.decode(await entryBytes(e)), name);
      const textOf = el => Array.from(el.getElementsByTagNameNS('*', 'p')).map(p => all(p, 't').map(t => t.textContent).join('')).join(' ');
      const rows = [];
      let n = 0;
      const body = all(doc, 'body')[0];
      for (const el of body ? Array.from(body.children) : []) {
        if (el.localName === 'tbl') {
          for (const tr of all(el, 'tr')) rows.push({ row: ++n, cells: kids(tr, 'tc').map(tc => ({ v: clean(textOf(tc)), num: false })) });
        } else if (el.localName === 'p') {
          const t = clean(all(el, 't').map(x => x.textContent).join(''));
          if (t) rows.push({ row: ++n, cells: [{ v: t, num: false }] });
        }
      }
      return { kind: 'Word', rows };
    }
    if (ext === 'html' || ext === 'htm') {
      const doc = new DOMParser().parseFromString(decodeText(bytes), 'text/html');
      doc.querySelectorAll('script, style, noscript').forEach(x => x.remove());
      const rows = [];
      let n = 0;
      doc.querySelectorAll('table').forEach(t => {
        t.querySelectorAll('tr').forEach(tr => rows.push({ row: ++n, cells: Array.from(tr.children).map(c => ({ v: clean(c.textContent), num: false })) }));
        t.remove();
      });
      const rest = textRows((doc.body ? doc.body.textContent : '').split(/\n/));
      return { kind: 'страница', rows: rows.concat(rest.map(r => ({ row: ++n, cells: r.cells }))) };
    }
    if (ext === 'pdf') {
      const lines = await pdfLines(bytes);
      const rows = textRows(lines);
      return { kind: 'PDF', rows, note: rows.some(r => /\d{6}/.test(r.cells[0].v)) ? '' : 'в PDF нет текста (скан): ИИН не прочитаны, нужна проверка вручную' };
    }
    const scan = /^(jpe?g|png|tiff?|bmp|gif|heic)$/.test(ext);
    return { kind: ext ? '.' + ext : 'без расширения', rows: [], note: scan ? 'это картинка (скан): ИИН не прочитаны, нужна проверка вручную'
      : ext === 'xls' || ext === 'doc' || ext === 'rtf' ? 'старый формат .' + ext + ': откройте в Excel или Word и сохраните как .xlsx или .docx' : 'формат не поддерживается' };
  }

  /* ИИН из строк файла. Заголовок (строка со словом «ИИН») подсказывает, где ФИО, дата и управляющий. */
  function extractIins(rows) {
    let head = null;
    for (const r of rows.slice(0, 80)) {
      if (r.cells.length > 1 && r.cells.some(c => c && /иин|iin/i.test(c.v) && !/\d{6}/.test(c.v))) { head = r; break; }
    }
    const role = [];
    if (head) head.cells.forEach((c, i) => {
      const t = c ? c.v.toLowerCase() : '';
      if (!t) return;
      const mgr = /управляющ/.test(t);
      if (/иин|iin/.test(t)) role[i] = mgr ? 'mgrIin' : 'iin';
      else if (mgr) role[i] = 'mgr';
      else if (/фио|ф\.\s*и\.\s*о|должник|граждан|заем|заём|наименован|фамили/.test(t)) role[i] = role.includes('fio') ? 'x' : 'fio';
      else if (/дата/.test(t)) role[i] = role.includes('date') ? 'x' : 'date';
    });
    const items = [], invalid = [];
    for (const r of rows) {
      if (r === head) continue;
      const found = new Map();
      r.cells.forEach((c, i) => {
        if (!c || role[i] === 'mgrIin') return;
        const before = found.size;
        const whole = normIin(c.v, c.num);
        if (whole && whole.iin && personalIin(whole.iin)) found.set(whole.iin, i);
        for (const m of c.v.matchAll(/(?<!\d)\d{12}(?!\d)/g)) if (iinOk(m[0]) && personalIin(m[0]) && !found.has(m[0])) found.set(m[0], i);
        if (found.size === before) {          /* ИИН, разбитый пробелами: 8709 1630 1642, 870 916 301 642, 870916 301642, по цифре */
          for (const m of c.v.matchAll(/(?<!\d)(?:\d{4}[ \xa0]\d{4}[ \xa0]\d{4}|\d{3}(?:[ \xa0]\d{3}){3}|\d{6}[ \xa0]\d{6}|(?:\d[ \xa0]){11}\d)(?!\d)/g)) {
            const x = m[0].replace(/\D/g, '');
            if (iinOk(x) && personalIin(x) && !found.has(x)) found.set(x, i);
          }
        }
        if (found.size === before && whole && whole.err && (role[i] === 'iin' || (!head && /^\d{12}$/.test(c.v)))) {
          invalid.push({ where: (r.sheet ? r.sheet + ', ' : '') + 'строка ' + r.row, raw: c.v, err: whole.err });
        }
      });
      if (!found.size) continue;
      const by = want => r.cells.findIndex((c, i) => c && role[i] === want);
      const fioCol = by('fio'), mgrCol = by('mgr'), dateCol = by('date');
      let fio = fioCol >= 0 ? r.cells[fioCol].v : '';
      if (!fio) { const c = r.cells.find((x, i) => x && i !== mgrCol && NAME_RE.test(x.v)); fio = c ? c.v : ''; }
      if (!fio && r.cells.length === 1) {                                 /* строка текста (PDF, Word): имя прямо перед ИИН */
        const at = r.cells[0].v.search(/\d{4}/);
        const names = Array.from((at > 0 ? r.cells[0].v.slice(0, at) : '').matchAll(NAME_IN_TEXT));
        fio = names.length ? names[names.length - 1][0] : '';
      }
      const serial = c => (c && c.num && +c.v > 30000 && +c.v < 60000 ? [new Date(1899, 11, 30 + Math.floor(+c.v))] : []);   /* дата Excel: число дней */
      const rowDates = dateCol >= 0 ? datesIn(r.cells[dateCol].v).concat(serial(r.cells[dateCol])) : datesIn(r.cells.filter(Boolean).map(c => c.v).join(' '));
      for (const iin of found.keys()) {
        items.push({ iin, fio: fio.slice(0, 120), mgr: mgrCol >= 0 ? r.cells[mgrCol].v.slice(0, 120) : '', date: maxDate(rowDates),
          where: (r.sheet ? 'лист «' + r.sheet + '», ' : '') + 'строка ' + r.row });
      }
    }
    return { items, invalid };
  }

  /* Все файлы подпапки «ДГД» (и её подпапок): публикации, собранные с gov.kz (в папке региона файл DGD_META), и файлы,
     положенные вручную (списки xlsx, csv, docx, html, PDF, сканы). ИИН из сканов читаются один раз и запоминаются
     в «ДГД/распознано.json». */
  async function readDgd(dir) {
    let root;
    try { root = await dir.getDirectoryHandle(DGD_DIR); } catch (e) { return null; }
    const files = [], metas = [];
    const walk = async (d, prefix, depth) => {
      for await (const [name, hnd] of d.entries()) {
        if (hnd.kind === 'directory') { if (depth < 2) await walk(hnd, prefix + name + '/', depth + 1); continue; }
        if (/^(~\$|\.)/.test(name) || /\.(crswap|tmp|ini|db)$/i.test(name) || name === DGD_OCR) continue;
        if (name === DGD_META) { const data = await readJsonFile(d, name); if (data && data.region) metas.push({ dir: d, folder: prefix.slice(DGD_DIR.length + 1).replace(/\/$/, ''), data }); continue; }
        files.push({ name, path: prefix + name, handle: hnd });
      }
    };
    await walk(root, DGD_DIR + '/', 0);
    files.sort((a, b) => a.path.localeCompare(b.path, 'ru'));
    const cache = (await readJsonFile(root, DGD_OCR)) || {};
    let cacheChanged = false;
    /* ИИН из файла-скана: из памяти или распознаванием (и запомнить) */
    const scanOf = async (key, handle, name) => {
      const file = await handle.getFile();
      const hit = cache[key];
      if (hit && hit.size === file.size && Array.isArray(hit.iins)) return hit;
      const r = await readOrderFile(new Uint8Array(await file.arrayBuffer()), name, '');
      cache[key] = { size: file.size, at: new Date().toISOString(), note: r.note || '', iins: r.iins.map(x => ({ iin: x.iin, conf: x.conf, how: x.how, fixes: x.fixes || 0 })) };
      cacheChanged = true;
      return cache[key];
    };
    const out = { files: [], items: [], orders: [], sites: [], at: new Date() };
    const used = new Set();
    for (const m of metas.sort((a, b) => a.folder.localeCompare(b.folder, 'ru'))) {
      const meta = m.data, region = meta.region;
      const site = { region, host: meta.host || '', state: meta.state || 'ok', note: meta.note || '', collectedAt: meta.collectedAt ? new Date(meta.collectedAt) : null,
        failedAt: meta.failedAt ? new Date(meta.failedAt) : null, window: meta.window || DGD_WINDOW, pages: meta.pages || [], rows: [] };
      for (const r of meta.rows || []) {
        const row = { region, date: r.date ? new Date(r.date + 'T00:00:00') : null, dateFrom: r.dateFrom || '', kind: r.kind || '', debtor: r.debtor || '',
          manager: r.manager || '', number: r.number || '', text: r.text || '', page: r.page || '', future: !!r.future, firstSeen: r.firstSeen || '',
          files: [], iins: [], gov: (r.files || []).some(f => f.gov), notes: [].concat(r.notes || []) };
        for (const f of r.files || []) {
          const file = { url: f.url, name: f.name, gov: !!f.gov, saved: f.saved || '', err: f.err || '', path: '' };
          row.files.push(file);
          if (!f.saved) continue;
          file.path = DGD_DIR + '/' + m.folder + '/' + f.saved;
          used.add(file.path);
          let res;
          try { res = await scanOf(m.folder + '/' + f.saved, await m.dir.getFileHandle(f.saved), f.saved); }
          catch (e) { res = { iins: [], note: 'файла нет в папке' }; }
          if (res.note && !res.iins.length) row.notes.push(res.note);
          for (const x of res.iins) if (!row.iins.some(y => y.iin === x.iin)) row.iins.push(Object.assign({ path: file.path }, x));
        }
        site.rows.push(row);
        out.orders.push(row);
        for (const x of row.iins) {
          out.items.push({ iin: x.iin, fio: row.debtor, mgr: row.manager, date: row.date, listDate: row.date, region, path: x.path, kind: row.kind,
            conf: x.conf, how: x.how, page: row.page, where: (row.kind || 'публикация') + (row.number ? ' № ' + row.number : '') + (row.date ? ' от ' + dmy(row.date) : '') });
        }
      }
      out.sites.push(site);
    }
    for (const f of files) {
      if (used.has(f.path)) continue;
      const rec = { path: f.path, name: f.name, region: '', kind: '', listDate: null, lastDate: null, count: 0, unique: 0, invalid: [], note: '' };
      const fromPath = regionsIn(f.path.slice(DGD_DIR.length + 1), true);      /* регион по имени: известен, даже если файл не прочитан */
      if (fromPath.length === 1) rec.region = fromPath[0];
      try {
        const file = await f.handle.getFile();
        const bytes = new Uint8Array(await file.arrayBuffer());
        const parsed = await fileRows(f.name, bytes);
        rec.kind = parsed.kind;
        rec.note = parsed.note || '';
        const ex = extractIins(parsed.rows);
        rec.invalid = ex.invalid;
        const head = parsed.rows.slice(0, 40).map(r => r.cells.filter(Boolean).map(c => c.v).join(' ')).join(' ');
        const fromText = fromPath.length ? [] : regionsIn(head);
        const regs = fromPath.length ? fromPath : fromText;
        rec.region = regs.length === 1 ? regs[0] : '';
        if (regs.length > 1) rec.note = (rec.note ? rec.note + '; ' : '') + 'несколько регионов в названии: ' + regs.join(', ') + ' (переименуйте файл)';
        const nameDates = datesIn(f.path);
        rec.listDate = maxDate(nameDates) || null;
        if (!ex.items.length && /скан|картинка/.test(rec.note)) {            /* скан или фото приказа: ИИН читается с картинки */
          const r = await scanOf(f.path.slice(DGD_DIR.length + 1), f.handle, f.name);
          for (const x of r.iins) ex.items.push({ iin: x.iin, fio: '', mgr: '', date: rec.listDate, where: x.how + (x.conf !== 'high' ? ' (сверить с приказом)' : ''), conf: x.conf, how: x.how });
          rec.note = r.iins.length ? (rec.kind === 'PDF' ? 'скан PDF' : 'картинка (скан)') + ': ИИН ' + (r.iins.every(x => x.conf === 'high') ? 'распознан' : 'распознан неуверенно, сверить с приказом')
            : rec.note.replace(/: ИИН не прочитаны, нужна проверка вручную/, '') + ': ИИН не распознан, нужна проверка вручную' + (r.note && !/не распознан/.test(r.note) ? ' (' + r.note + ')' : '');
        }
        rec.lastDate = maxDate(ex.items.map(x => x.date));
        if (!rec.listDate) rec.listDate = rec.lastDate || null;
        if (!rec.listDate) { rec.listDate = new Date(file.lastModified); rec.dateNote = 'дата списка не найдена: взято время файла'; }
        rec.count = ex.items.length;
        rec.unique = new Set(ex.items.map(x => x.iin)).size;
        for (const x of ex.items) out.items.push(Object.assign({ region: rec.region, listDate: rec.listDate, path: f.path }, x));
        if (!ex.items.length && !rec.note && parsed.rows.length) rec.note = 'ИИН в файле не найдены';
        if (!parsed.rows.length && !rec.note) rec.note = 'файл пустой';
      } catch (e) { rec.note = 'не прочитан: ' + (e && e.message ? e.message : e); rec.failed = true; }
      out.files.push(rec);
    }
    if (cacheChanged) { try { await writeJsonFile(root, DGD_OCR, cache); } catch (e) { /* прочитается снова */ } }
    out.byIin = new Map();
    for (const x of out.items) { if (!out.byIin.has(x.iin)) out.byIin.set(x.iin, []); out.byIin.get(x.iin).push(x); }
    out.regions = dgdRegions(out);
    return out;
  }

  /* Проверка регионов: всегда все 20, чтобы пустые были видны. По сайту: открылся ли, есть ли страница приказов,
     сколько приказов за DGD_WINDOW дней и за год, свежий ли последний, прочитаны ли ИИН. По файлам: как раньше. */
  function dgdRegions(dg) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const days = d => Math.floor((today - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 86400000);
    const names = REGION_LIST.map(x => x[0]).concat(dg.files.some(f => !f.region) ? ['Регион не определён'] : []);
    return names.map(name => {
      const files = dg.files.filter(f => (f.region || 'Регион не определён') === name);
      const items = dg.items.filter(x => (x.region || 'Регион не определён') === name);
      const fileItems = items.filter(x => !x.page);
      const unique = new Set(items.map(x => x.iin)).size;
      const invalid = files.reduce((a, f) => a + f.invalid.length, 0);
      const listDate = maxDate(files.map(f => f.listDate)), lastDate = maxDate(fileItems.map(x => x.date));
      const site = (dg.sites || []).find(s => s.region === name) || null;
      const problems = [];
      let level = 'good';
      const bad = t => { problems.push(t); level = 'bad'; };
      const warn = t => { problems.push(t); if (level !== 'bad') level = 'warn'; };
      const r = { name, files, items: items.length, unique, invalid, listDate, lastDate, site, level, text: '' };
      if (site) {
        const since = new Date(today); since.setDate(since.getDate() - (site.window || DGD_WINDOW));
        const win = site.rows.filter(x => x.date && x.date >= since), year = site.rows.filter(x => x.date && x.date.getFullYear() === today.getFullYear());
        const last = maxDate(site.rows.map(x => x.date));
        Object.assign(r, { win: win.length, year: year.length, lastOrder: last, read: win.filter(x => x.iins.length).length,
          unread: win.filter(x => !x.iins.length && x.files.some(f => f.saved)).length, pages: site.pages.length });
        const when = site.collectedAt ? ' (собрано ' + dmy(site.collectedAt) + ')' : '';
        if (site.state === 'skip') bad('сайт не открылся' + (site.rows.length ? '; приказы с прошлого сбора' + when : ''));
        else if (site.state !== 'ok') bad((site.note || 'сбор не удался') + (site.rows.length ? '; приказы с прошлого сбора' + when : ''));
        if (site.state !== 'ok') { /* причина уже названа */ }
        else if (!site.pages.some(p => p.read)) bad('раздел «Документы» на gov.kz не открылся');
        else if (!win.length) bad('0 публикаций за ' + (site.window || DGD_WINDOW) + ' дн.');
        else if (win.length <= DGD_FEW) warn('всего ' + win.length + ' публикаци' + (win.length === 1 ? 'я' : 'и') + ' за ' + (site.window || DGD_WINDOW) + ' дн.: проверить на gov.kz');
        if (last && days(last) > DGD_SITE_STALE) warn('последняя публикация ' + dmy(last) + ': ' + days(last) + ' дн. назад');
        if (r.unread) warn('ИИН не прочитан в публикациях: ' + r.unread + ' (открыть файл, сверить вручную)');
        const nofile = win.filter(x => !x.files.some(f => f.saved)).length;
        if (nofile) warn('без файла: ' + nofile + ' (открыть публикацию на gov.kz)');
        const doubt = win.filter(x => x.iins.some(y => y.conf !== 'high')).length;
        if (doubt) warn('ИИН прочитан неуверенно: ' + doubt + ' (сверить с приказом)');
        const undated = site.rows.filter(x => !x.date).length;
        if (undated) warn('строк без даты: ' + undated);
        const future = site.rows.filter(x => x.future).length;
        if (future) warn('дата приказа в будущем: ' + future);
        if (site.state === 'ok' && site.rows.length && !year.length && today.getMonth() > 0) warn('публикаций за ' + today.getFullYear() + ' год на gov.kz нет');
      }
      if (files.length || !site) {
        const newest = maxDate([listDate, lastDate]);
        const age = newest ? days(newest) : null;
        const fileUnique = new Set(fileItems.map(x => x.iin)).size;
        if (!files.length) bad(site ? '' : 'нет списка');
        else if (files.every(f => f.failed)) bad('файл не прочитан');
        else if (!fileUnique) bad('0 ИИН' + (files.some(f => f.note) ? ': ' + files.map(f => f.note).filter(Boolean)[0] : ''));
        else if (fileUnique <= DGD_FEW && !site) warn('всего ' + fileUnique + ' ИИН: проверить вручную');
        if (files.length && age !== null && age > DGD_STALE_DAYS) warn('самая свежая дата ' + dmy(newest) + ': ' + age + ' дн. назад');
        if (files.some(f => f.dateNote)) warn('дата списка не найдена');
        if (invalid) warn('неверных ИИН: ' + invalid);
        const unread = files.filter(f => f.failed || /не прочитан|не распознан|старый формат|не поддерживается/.test(f.note || '')).length;
        if (unread && files.length > unread) warn('не прочитано файлов: ' + unread);
        if (name === 'Регион не определён') warn('укажите регион в имени файла');
      }
      r.level = level;
      r.text = problems.filter(Boolean).length ? problems.filter(Boolean).join('; ') : 'в порядке';
      return r;
    });
  }

  async function dgdWorkbook(dg) {
    const lvl = { bad: STYLE.bad, warn: STYLE.warn, good: STYLE.good };
    const link = (v, to) => (to ? { v, link: to } : v);
    const regionRows = [['Регион', 'Проверка', 'gov.kz: сбор', 'Последняя публикация', 'Публикаций за ' + DGD_WINDOW + ' дн.', 'Публикаций за год', 'ИИН прочитан',
      'ИИН не прочитан', 'Файлов вручную', 'Уникальных ИИН', 'Неверных ИИН', 'Дата списка (файлы)', 'Раздел «Документы» на gov.kz', 'Файлы']]
      .concat(dg.regions.map(r => [r.name, { v: r.text, s: lvl[r.level] },
        r.site ? (r.site.state === 'ok' ? dmy(r.site.collectedAt) : 'не открылся ' + dmy(r.site.failedAt || r.site.collectedAt)) : 'нет',
        dmy(r.lastOrder), r.site ? String(r.win) : '', r.site ? String(r.year) : '', r.site ? String(r.read) : '', r.site ? String(r.unread) : '',
        String(r.files.length), String(r.unique), String(r.invalid), dmy(r.listDate),
        r.site ? (r.site.pages[0] ? r.site.pages[0].url.replace(/\/documents\/\d+\?.*$/, '/documents/1?lang=ru') : '') : '', r.files.map(f => f.path).join('; ')]));
    const conf = x => (x.conf === 'high' || !x.conf ? '' : ' (сверить)');
    const orderRowsOut = [['Регион', 'Дата публикации', 'Вид', 'Должник (как в названии)', 'Финансовый управляющий', 'ИИН', 'Как прочитан ИИН', 'Файл', 'Публикация на gov.kz', 'Впервые увиден', 'Название публикации']]
      .concat((dg.orders || []).slice().sort((a, b) => (b.date || 0) - (a.date || 0)).map(o => {
        const f = o.files.find(x => x.saved) || o.files[0] || {};
        return [o.region, dmy(o.date) + (o.dateFrom ? ' (' + o.dateFrom + ')' : ''), o.kind, o.debtor, o.manager,
          o.iins.length ? { v: o.iins.map(x => x.iin).join(', '), s: o.iins.some(x => x.conf !== 'high') ? STYLE.warn : 0 } : (f.saved ? { v: 'не прочитан', s: STYLE.warn } : ''),
          o.iins.map(x => x.how + conf(x)).join('; ') || o.notes.join('; ') || f.err || (f.saved ? '' : o.date && (new Date() - o.date) / 86400000 > DGD_WINDOW ? 'файл не скачивался (старше ' + DGD_WINDOW + ' дн.)' : 'файла нет'),
          f.saved ? link(f.path.slice(DGD_DIR.length + 1), f.path) : link(f.url || '', f.url), link(o.page, o.page), o.firstSeen, o.text];
      }));
    const iinRows = [['ИИН', 'ФИО в списке', 'Регион', 'Дата в строке', 'Финансовый управляющий', 'Дата списка', 'Файл', 'Где в файле']]
      .concat(dg.items.map(x => [x.iin, x.fio, x.region || 'не определён', dmy(x.date), x.mgr, dmy(x.listDate), { v: x.path, link: x.path }, x.where + conf(x)]));
    const fileRowsOut = [['Файл', 'Регион', 'Формат', 'Дата списка', 'Строк с ИИН', 'Уникальных ИИН', 'Неверных ИИН', 'Примечание']]
      .concat(dg.files.map(f => [{ v: f.path, link: f.path }, f.region || 'не определён', f.kind, dmy(f.listDate), String(f.count), String(f.unique),
        String(f.invalid.length), { v: [f.note, f.dateNote].concat(f.invalid.slice(0, 5).map(x => x.where + ': ' + x.raw + ' (' + x.err + ')')).filter(Boolean).join('; '),
          s: f.failed || (f.note && !/распознан$/.test(f.note)) ? STYLE.warn : 0 }]));
    return makeXlsx([
      { name: 'Регионы', rows: regionRows, widths: [30, 60, 16, 14, 12, 12, 11, 12, 10, 12, 11, 14, 70, 50] },
      { name: 'Публикации', rows: orderRowsOut, widths: [26, 14, 30, 34, 30, 15, 26, 46, 50, 12, 70] },
      { name: 'ИИН', rows: iinRows, widths: [15, 36, 28, 14, 36, 13, 40, 30] },
      { name: 'Файлы', rows: fileRowsOut, widths: [44, 28, 10, 13, 13, 16, 14, 70] },
    ]);
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
  const root = h('div', 'position:fixed;inset:0;z-index:2147483647;display:flex;background:#fff;visibility:visible;' +
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
    'Папка, где лежит файл Excel. В ней же появятся папки людей со снимками, ссылки запишутся прямо в файл.');
  const fileSelect = h('select', 'display:none;width:100%;margin:0 0 4px;font-size:12px');
  fileSelect.addEventListener('change', () => openFromFolder(fileSelect.value));
  const fileInput = h('input', 'display:none;width:100%;font-size:12px;margin:4px 0');
  fileInput.type = 'file';
  fileInput.accept = '.xlsx,.xlsm,.csv';
  fileInput.title = 'Файл Excel со столбцом «ИИН/БИН». Снимки и копия файла со ссылками сохранятся в «Загрузки».';
  fileInput.addEventListener('change', async () => {
    const f = fileInput.files[0];
    if (!f) return;
    if (/\.csv$/i.test(f.name)) {
      try { await openBytes(await csvToXlsx(new Uint8Array(await f.arrayBuffer())), f.name.replace(/\.csv$/i, '.xlsx'), null); }
      catch (e) { say('❌ ' + (e.message || e)); }
    } else await openBytes(new Uint8Array(await f.arrayBuffer()), f.name, null);
  });
  const fileLink = h('a', 'color:#1f6feb;cursor:pointer;font-size:12px', canPickFolder ? 'или один файл (результат в «Загрузки»)' : '');
  fileLink.addEventListener('click', () => { fileInput.style.display = fileInput.style.display === 'none' ? 'block' : 'none'; });
  if (!canPickFolder) { folderB.style.display = 'none'; fileInput.style.display = 'block'; }
  const sheetSelect = h('select', 'display:none;width:100%;margin:0 0 4px;font-size:12px');
  sheetSelect.title = 'Лист со столбцом «ИИН/БИН»';
  sheetSelect.addEventListener('change', () => useSheet(sheetSelect.value));

  /* Путь к выбранной папке на диске. Браузер его не сообщает, а для полной ссылки (столбец hyperlink) он нужен.
     Берётся из самого файла Excel (Excel записывает туда папку, где файл сохранили) или вставляется сюда один раз. */
  const baseBox = h('div', 'display:none;margin:2px 0 6px;font-size:12px');
  const baseInput = h('textarea', 'display:block;width:100%;height:34px;box-sizing:border-box;margin:2px 0;padding:4px;' +
    'font:11px/1.3 Arial,sans-serif;color:#1d2433;background:#fff;border:1px solid #b9c3d3;border-radius:4px;' +
    'resize:vertical;word-break:break-all');
  baseInput.id = 'tazalau-base';
  baseInput.spellcheck = false;
  baseInput.placeholder = 'W:\\Отдел\\Папка';
  baseInput.title = 'Откройте эту папку в Проводнике, щёлкните по адресной строке, скопируйте путь и вставьте сюда';
  const baseNote = h('div', 'white-space:pre-line');
  baseNote.id = 'tazalau-base-state';
  const baseAutoB = h('a', 'display:none;color:#1f6feb;cursor:pointer', 'Взять путь из файла Excel');
  baseBox.append(h('div', 'color:#5b6577', 'Путь к этой папке на диске (для столбца ' + ABS_HEADER + '):'), baseInput, baseNote, baseAutoB);
  baseInput.addEventListener('input', () => refresh());
  baseInput.addEventListener('change', () => { baseInput.value = normBase(baseInput.value); refresh(); });
  baseAutoB.addEventListener('click', () => { baseInput.value = baseAuto; refresh(); });

  const summary = h('div', 'margin:6px 0;white-space:pre-line;font-size:12px');
  /* Шаг 1: списки ДГД из подпапки «ДГД» */
  const dgdBox = h('div', 'display:none;margin:0 0 6px;font-size:12px');
  dgdBox.id = 'tazalau-dgd';
  const dgdDetails = h('details');
  const dgdHead = h('summary', 'cursor:pointer');
  const dgdList = h('div', 'white-space:pre-line;max-height:200px;overflow:auto;margin:4px 0');
  const dgdB = button('Свод ДГД ↓', () => writeDgd(false), false, 'Записать свод списков ДГД (все ИИН и проверку регионов) в файл ' + DGD_FILE);
  /* ИИН из списков ДГД готовым списком для SQL (шаг 2: выгрузка ADW только по этим ИИН) */
  const sqlB = button('ИИН для SQL ⧉', copySql, false, 'Скопировать ИИН из списков ДГД в виде VALUES (...) для запроса к ADW');
  dgdDetails.append(dgdHead, dgdList);
  /* Сбор публикаций ДГД 20 регионов с gov.kz: вкладка gov.kz + закладка в ней (см. relay) */
  const collectB = button('🌐 Собрать ДГД с gov.kz', () => collectDgd(), true,
    'Помощник откроет вкладку gov.kz: нажмите в ней закладку Tazalau один раз. Он соберёт публикации ДГД всех 20 регионов о процедурах граждан ' +
    '(приказы о финансовых управляющих, объявления о возбуждении дел, реестры требований) за ' + DGD_WINDOW + ' дней в подпапку «' + DGD_DIR + '» и прочитает ИИН.');
  const collectBox = h('div', 'display:none;white-space:pre-line;max-height:240px;overflow:auto;margin:2px 0 4px;padding:4px 6px;' +
    'background:#fff;border:1px solid #d5dbe5;border-radius:4px;font-size:12px');
  collectBox.id = 'tazalau-collect';
  const skipB = button('Пропустить регион', () => skipSite(), false, 'Регион долго не отвечает: бросить его и перейти к следующему');
  const stopCollectB = button('Закончить сбор', () => finishCollect('остановлено'), false, 'Остановить сбор: всё собранное сохранится');
  const collectCtl = row(skipB, stopCollectB);
  collectCtl.style.display = 'none';
  dgdBox.append(row(collectB), collectBox, collectCtl, dgdDetails, row(dgdB, sqlB));
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

  panel.append(head, banner, row(folderB), fileSelect, row(fileLink), fileInput, sheetSelect, baseBox, summary, dgdBox, badBox, regBox,
    row(startB, stopB), state, row(csvB, xlsxB), jbox);
  document.body.appendChild(root);
  const overflowBefore = document.documentElement.style.overflow;
  document.documentElement.style.overflow = 'hidden';
  window.addEventListener('resize', layoutFrame);
  window.addEventListener('beforeunload', e => { if (running) { e.preventDefault(); e.returnValue = ''; } });

  /* ---------- состояние ---------- */
  let book = null, S = null, tasks = [], store = { dir: null, handle: null };
  let running = false, stopping = false, interval = T.INTERVAL, lastRequest = 0;
  let baseAuto = '';                   /* путь к выбранной папке, найденный в самом файле Excel */
  let dgd = null;                      /* списки ДГД папки: null (нет подпапки), 'reading', { files, items, regions, byIin } */
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

  /* Имя папки человека: ФИО из столбца FIO, а если его нет, ИИН. Знаки, недопустимые в именах папок Windows
     (и «#», «%», с которыми Excel не открывает ссылки), заменяются пробелом. Однофамильцам с разными ИИН
     к имени добавляется ИИН. Считается по всем строкам листа, поэтому имя одно и то же при любом запуске. */
  const safeName = v => {
    const t = clean(String(v || '').replace(/[\\/:*?"<>|#%\u0000-\u001f]/g, ' ')).slice(0, 100).replace(/[. ]+$/, '');
    return /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(t) ? '_' + t : t;
  };
  const cutName = (name, n) => (n > 0 ? name.slice(0, n).replace(/[. ]+$/, '') : '');
  /* maxLen: сколько знаков можно отдать имени папки (путь к папке на диске длинный: имя укорачивается) */
  function folderNames(maxLen) {
    const fio = new Map(), count = new Map(), out = new Map();
    for (const t of tasks) if (t.parse.iin && !fio.get(t.parse.iin)) fio.set(t.parse.iin, cutName(safeName(t.fio), maxLen));
    for (const [iin, name] of fio) { const key = (name || iin).toLowerCase(); count.set(key, (count.get(key) || 0) + 1); }
    for (const [iin, name] of fio) {
      out.set(iin, !name ? iin : count.get(name.toLowerCase()) > 1 ? (cutName(name, maxLen - iin.length - 1) + ' ' + iin).trim() : name);
    }
    return out;
  }
  /* Прежний снимок не затирается: получает дату в имени, а «снимок.png» всегда самый свежий.
     room: сколько знаков можно отдать имени файла; если имя с датой не помещается в путь Windows, берётся короткое с номером. */
  async function keepOld(dir, room) {
    let old;
    try { old = await (await dir.getFileHandle(SHOT_NAME)).getFile(); } catch (e) { return; }
    const data = await old.arrayBuffer();
    const dated = SHOT_NAME.replace(/\.png$/, '_' + stampOf(new Date(old.lastModified)) + '.png');
    if (dated.length <= room) {
      try { await writeFile(dir, dated, data); return; }
      catch (e) { try { await dir.removeEntry(dated); } catch (x) { /* не создан */ } }
    }
    for (let n = 1; n < 100; n++) {
      const name = SHOT_NAME.replace(/\.png$/, '_' + n + '.png');
      try { await dir.getFileHandle(name); continue; } catch (e) { /* имя свободно */ }
      await writeFile(dir, name, data);
      return;
    }
    throw new Error('в папке уже 99 прежних снимков');
  }

  /* Путь к папке, вставленный вручную, помощник помнит в самой папке: файл «путь_к_папке.txt» в tazalau_backup.
     Так путь не остаётся в браузере (хранилище браузера читает и сам сайт) и общий для всех, кто работает в этой папке.
     В файле путь или слово «нет» (на вопрос ответили «без hyperlink», больше не спрашивать). */
  async function readSaved() {
    try {
      const dir = await store.dir.getDirectoryHandle(BACKUP_DIR);
      const text = await (await (await dir.getFileHandle(PATH_FILE)).getFile()).text();
      const line = text.replace(/^\ufeff/, '').split(/\r?\n/)[0].trim();
      return /^нет$/i.test(line) ? '' : normBase(line) || undefined;
    } catch (e) { return undefined; }
  }
  /* v: путь, '' (без hyperlink) или null (забыть: путь снова берётся из файла Excel) */
  async function writeSaved(v) {
    try {
      if (v == null) {
        if (store.saved === undefined) return;
        await (await store.dir.getDirectoryHandle(BACKUP_DIR)).removeEntry(PATH_FILE);
      } else {
        if (store.saved === v) return;
        const dir = await store.dir.getDirectoryHandle(BACKUP_DIR, { create: true });
        await writeFile(dir, PATH_FILE, new Blob([(v || 'нет') + '\r\n'], { type: 'text/plain' }));
      }
      store.saved = v == null ? undefined : v;
    } catch (e) { /* не записалось: в следующий раз путь вставят заново */ }
  }
  /* Что сейчас в поле пути: base (годный путь или '') и пояснение для окна */
  function baseState() {
    if (!store.handle) return { base: '', text: '' };
    const raw = normBase(baseInput.value), folder = String(store.dir.name);
    if (!raw) {
      return { base: '', text: 'Столбец ' + ABS_HEADER + ' сейчас не заполняется. Чтобы заполнялся, вставьте сюда путь к папке из адресной строки Проводника.' };
    }
    if (!baseValid(raw)) return { base: '', bad: true, text: '⚠ Это не путь Windows. Нужен вид W:\\Отдел\\Папка или \\\\сервер\\папка.' };
    const room = folderRoom(raw);
    if (room < 12) {
      return { base: '', bad: true, text: '❌ Путь слишком длинный (' + raw.length + ' знаков): Windows не откроет снимки. ' +
        'Нужна папка с путём не длиннее ' + (raw.length + room - 12) + ' знаков.' };
    }
    const mismatch = !/^[a-z]:$/i.test(raw) && !sameName(lastPart(raw), folder);
    const cut = new Set(tasks.filter(t => t.parse.iin && safeName(t.fio).length > room).map(t => t.parse.iin)).size;
    const lines = [mismatch
      ? '⚠ Путь кончается папкой «' + lastPart(raw) + '», а выбрана папка «' + folder + '». Ссылки откроются, только если это одна и та же папка.'
      : '✔ Ссылка в столбце ' + ABS_HEADER + ': этот путь + папка человека + ' + SHOT_NAME + (raw === baseAuto ? '\nПуть взят из самого файла Excel.' : '')];
    if (cut) lines.push('Длинных имён папок будет укорочено: ' + cut + ' (до ' + room + ' знаков: длиннее путь Windows не откроет).');
    return { base: raw, mismatch, room, text: lines.join('\n') };
  }

  function refresh() {
    const regs = chosen();
    regHead.textContent = 'Реестры: ' + (regs.length === 1 ? regs[0].short : regs.length ? regs.length + ' выбрано' : 'не выбрано');
    const b = baseState();
    baseBox.style.display = S && store.handle ? 'block' : 'none';
    baseNote.textContent = b.text;
    baseNote.style.color = b.bad || b.mismatch ? '#9a3412' : b.base ? '#1a7f37' : '#5b6577';
    baseAutoB.style.display = baseAuto && normBase(baseInput.value) !== baseAuto ? 'inline' : 'none';
    baseAutoB.title = baseAuto;
    baseInput.style.height = 'auto';                /* поле по высоте текста: длинный путь виден целиком */
    baseInput.style.height = Math.min(150, Math.max(34, baseInput.scrollHeight + 2)) + 'px';
    showDgd();
    if (!S) { summary.textContent = ''; badBox.style.display = 'none'; updateButtons(); return; }
    const p = plan(false), rows = Array.from(p.groups.values()).reduce((a, x) => a + x.length, 0);
    const inDgd = dgd && dgd.byIin ? Array.from(p.groups.keys()).filter(i => dgd.byIin.has(i)).length : 0;
    const skipped = [p.linked ? 'уже со ссылкой ' + p.linked : '', p.hiddenRows ? 'скрыто фильтром ' + p.hiddenRows : ''].filter(Boolean).join(', ');
    summary.textContent = book.fileName + ' · лист «' + S.sheet.name + '»\n' +
      'Строк с ИИН ' + tasks.length + (skipped ? ' · ' + skipped : '') + '\n' +
      (S.layout.fioCol ? 'Папки снимков: по столбцу FIO' : 'Столбца FIO нет: папки снимков по ИИН') + '\n' +
      (p.groups.size ? 'К проверке: ' + p.groups.size + ' ИИН (строк ' + rows + '), примерно ' + eta(p.groups.size * Math.max(1, regs.length)) +
        (inDgd ? '\nИз них в списках ДГД: ' + inDgd + ' (проверяются первыми)' : '')
                     : (tasks.length ? 'Новых ИИН нет ✔ (Старт спросит, проверять ли заново)' : 'В столбце «ИИН/БИН» нет значений'));
    badBox.style.display = p.invalid.length ? 'block' : 'none';
    badHead.textContent = '⚠ Строк с неверным ИИН: ' + p.invalid.length;
    badList.textContent = p.invalid.map(t => 'строка ' + t.row + ': ' + t.parse.err).join('\n');
    jhead.textContent = 'Журнал (' + journal.length + ')';
    updateButtons();
  }
  function showDgd() {
    dgdBox.style.display = store.dir ? 'block' : 'none';
    dgdB.style.display = sqlB.style.display = dgd && dgd.files ? '' : 'none';
    if (dgd === 'reading') { dgdHead.textContent = 'ДГД: читаю списки…'; dgdList.textContent = ''; return; }
    if (!dgd) {
      dgdHead.textContent = 'ДГД: подпапки «' + DGD_DIR + '» нет';
      dgdList.textContent = 'Нажмите «Собрать ДГД с gov.kz»: помощник сам соберёт публикации ДГД всех 20 регионов о процедурах граждан в подпапку «' +
        DGD_DIR + '» и прочитает ИИН. Списки, скачанные вручную, тоже можно положить в «' + DGD_DIR + '»: файл на регион, в имени регион и дата.';
      return;
    }
    const withList = dgd.regions.filter(r => (r.files.length || r.site) && r.name !== 'Регион не определён').length;
    const issues = dgd.regions.filter(r => r.level !== 'good');
    dgdHead.textContent = 'ДГД: регионов со списком ' + withList + ' из ' + REGION_LIST.length + ' · ИИН ' + dgd.byIin.size +
      (issues.length ? ' · ⚠ проверить: ' + issues.length : ' ✔');
    const icon = { bad: '❌ ', warn: '⚠ ', good: '✔ ' };
    if (issues.some(r => r.level === 'bad') && !dgdBox.dataset.shown) { dgdDetails.open = true; dgdBox.dataset.shown = '1'; }   /* красное видно сразу */
    dgdList.textContent = issues.map(r => icon[r.level] + r.name + ': ' + r.text).join('\n') +
      (issues.length < dgd.regions.length ? (issues.length ? '\n' : '') + '✔ в порядке: ' + dgd.regions.filter(r => r.level === 'good').map(r => r.name).join(', ') : '');
  }
  async function loadDgd() {
    if (!store.dir) { dgd = null; return; }
    const dir = store.dir;
    dgd = 'reading';
    refresh();
    let res = null;
    try { res = await readDgd(dir); } catch (e) { res = null; say('❌ Списки ДГД не прочитаны: ' + (e.message || e)); }
    if (store.dir === dir) { dgd = res; refresh(); }
  }
  async function writeDgd(quiet) {
    if (!dgd || !dgd.files || !store.dir) return false;
    try {
      await writeFile(store.dir, DGD_FILE, await dgdWorkbook(dgd));
      if (!quiet) say('✔ Свод ДГД записан: ' + DGD_FILE + ' (листы «Регионы», «ИИН», «Файлы»).');
      return true;
    } catch (e) {
      say('⚠ Свод ДГД не записан (' + DGD_FILE + ' открыт в Excel?): ' + (e.message || e));
      return false;
    }
  }
  async function copySql() {
    if (!dgd || !dgd.byIin) return;
    const list = Array.from(dgd.byIin.keys()).sort();
    const text = list.length ? 'VALUES\n' + list.map(i => "  ('" + i + "')").join(',\n') : "VALUES ('000000000000')";
    try { await navigator.clipboard.writeText(text); }
    catch (e) {
      const a = document.createElement('textarea');
      a.value = text;
      document.body.appendChild(a);
      a.select();
      document.execCommand('copy');
      a.remove();
    }
    say('✔ Скопировано ИИН: ' + list.length + '. Вставьте в запрос ADW вместо строки «VALUES (...)» в блоке dgd(iin).');
  }
  /* Один ли человек: имя с сайта ДГД (в любом падеже, иногда с инициалами) и ФИО из рабочего файла */
  const KZ_RU = { 'ә': 'а', 'і': 'и', 'ң': 'н', 'ғ': 'г', 'ү': 'у', 'ұ': 'у', 'қ': 'к', 'ө': 'о', 'һ': 'х', 'ё': 'е', 'й': 'и' };
  const nameParts = v => String(v || '').toLowerCase().replace(/[әіңғүұқөһёй]/g, ch => KZ_RU[ch]).replace(/[^а-яa-z\s.-]/g, ' ').split(/[\s.]+/).filter(Boolean);
  function samePerson(a, b) {
    const A = nameParts(a), B = nameParts(b);
    if (A.length < 2 || B.length < 2) return false;
    const stem = w => (w.length > 5 ? w.slice(0, -2) : w.slice(0, Math.max(3, w.length - 1)));
    const sa = stem(A[0]), sb = stem(B[0]);
    if (Math.min(sa.length, sb.length) < 3 || !(sa.startsWith(sb) || sb.startsWith(sa))) return false;
    if (A[1][0] !== B[1][0] || (A[1].length > 2 && B[1].length > 2 && A[1].slice(0, 3) !== B[1].slice(0, 3))) return false;
    return !(A[2] && B[2] && A[2][0] !== B[2][0]);
  }
  const dgdLine = x => (x.page
    ? x.region + ', ' + x.where + (x.mgr ? ', управляющий: ' + x.mgr : '') + (x.conf && x.conf !== 'high' ? ', ИИН прочитан со скана неуверенно: сверить с приказом' : '')
    : (x.region || 'регион не определён') + ', список от ' + dmy(x.listDate) + (x.date ? ', дата ' + dmy(x.date) : '') + (x.mgr ? ', управляющий: ' + x.mgr : '') +
      (x.conf && x.conf !== 'high' ? ', ИИН прочитан со скана неуверенно' : ''));
  /* Столбец «ДГД» рабочего файла: регион, приказ или список, финансовый управляющий; ссылка на файл ДГД.
     Приказ, где ИИН не прочитан, сверяется по ФИО: такая строка помечается «возможно» и жёлтым. */
  async function fillDgd(S) {
    const L = S.layout;
    const blind = (dgd.orders || []).filter(o => !o.iins.length && o.debtor);
    let marked = 0;
    for (const t of tasks) {
      const hits = t.parse.iin ? dgd.byIin.get(t.parse.iin) || [] : [];
      try {
        const maybe = !hits.length && t.fio ? blind.filter(o => samePerson(o.debtor, t.fio)) : [];
        if (maybe.length) {
          const o = maybe[0], f = o.files.find(x => x.saved) || o.files[0] || {};
          const cur = (S.cells.get(t.row) || new Map()).get(L.dgdCol);
          setText(S, t.row, L.dgdCol, ('⚠ возможно (по ФИО, ИИН в файле не прочитан): ' + maybe.map(m => m.region + ', «' + m.debtor + '», ' + (m.kind || 'публикация') +
            (m.date ? ' от ' + dmy(m.date) : '') + (m.manager ? ', управляющий: ' + m.manager : '')).join('; ')).slice(0, 32000),
            f.path || f.url ? await linkStyle(S, cur ? cur.getAttribute('s') : null) : null);
          if (f.path || f.url) await setLink(S, t.row, L.dgdCol, f.path || f.url);
          marked++;
          continue;
        }
        if (!hits.length) {
          const cur = (S.cells.get(t.row) || new Map()).get(L.dgdCol);
          if (cur && clean(textAt(S, t.row, L.dgdCol))) { setText(S, t.row, L.dgdCol, '', null); await dropLink(S, t.row, L.dgdCol); }
          continue;
        }
        const seen = new Set(), parts = [];
        for (const x of hits) {
          const k = x.region + '|' + dmy(x.listDate) + '|' + x.path;
          if (seen.has(k)) continue;
          seen.add(k);
          parts.push(dgdLine(x));
        }
        const cur = (S.cells.get(t.row) || new Map()).get(L.dgdCol);
        setText(S, t.row, L.dgdCol, parts.join('; ').slice(0, 32000), await linkStyle(S, cur ? cur.getAttribute('s') : null));
        await setLink(S, t.row, L.dgdCol, hits[0].path);
        marked++;
      } catch (e) { /* объединённая ячейка */ }
    }
    return marked;
  }
  /* ---------- сбор с gov.kz: одна вкладка gov.kz, закладка в ней нажимается один раз ---------- */
  let collect = null;                  /* идёт сбор: { sites, popup, origin, pending, seq, cur, ocr, cache, … } */
  async function readJsonFile(dir, name) {
    try { return JSON.parse((await (await (await dir.getFileHandle(name)).getFile()).text()).replace(/^﻿/, '')); } catch (e) { return null; }
  }
  async function writeJsonFile(dir, name, obj) { await writeFile(dir, name, new Blob([JSON.stringify(obj, null, 1)], { type: 'application/json' })); }
  const isoDay = d => (d ? d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) : '');
  const dayOf = v => { const d = new Date(v); return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
  function collectDgd() {
    if (running) return;
    if (!store.dir) { say('Сначала выберите папку (📁 Папка с Excel): в её подпапку «' + DGD_DIR + '» помощник сложит приказы и объявления ДГД.'); return; }
    if (collect) { try { collect.popup.focus(); } catch (e) { /* вкладка закрыта */ } return; }
    /* Вкладка gov.kz открывается сразу по нажатию кнопки (иначе браузер её заблокирует); это обычная вкладка с панелью закладок */
    const popup = window.open(GOV_START, 'tazalau_dgd');
    if (!popup) { say('❌ Браузер не дал открыть вкладку gov.kz. Разрешите всплывающие окна для tazalau.qoldau.kz (значок справа в адресной строке) и нажмите кнопку ещё раз.'); return; }
    const sites = GOV_SITES.map(([slug, name]) => ({ slug, key: slug, name, state: 'wait', note: '' }));
    collect = { sites, popup, origin: '', pending: new Map(), seq: 0, cur: null, ocr: Promise.resolve(), ocrLeft: 0, cache: null, started: new Date(), connected: false };
    collect.watch = setInterval(() => { if (collect && collect.popup && collect.popup.closed && !collect.finishing) finishCollect('вкладка gov.kz закрыта'); }, 1500);
    showCollect();
    updateButtons();
  }
  function showCollect(final) {
    const c = collect;
    collectCtl.style.display = c && !c.finishing ? '' : 'none';
    collectB.style.display = c ? 'none' : '';
    if (!c && !final) { collectBox.style.display = 'none'; return; }
    collectBox.style.display = 'block';
    if (final) { collectBox.textContent = final; return; }
    if (!c.connected) { collectBox.textContent = 'Открыта вкладка gov.kz: нажмите в ней закладку Tazalau ▶ (один раз). Дальше помощник соберёт все 20 регионов сам.'; return; }
    const icon = { wait: '·', work: '⏬', done: '✔', fail: '❌', skip: '⏭' };
    const lines = c.sites.map(s => icon[s.state] + ' ' + s.name + (s.note ? ': ' + s.note : ''));
    const done = c.sites.filter(s => /done|fail|skip/.test(s.state)).length;
    collectBox.textContent = 'Сбор ДГД с gov.kz: ' + done + ' из ' + c.sites.length + (c.ocrLeft ? ' · читаю ИИН: ' + c.ocrLeft : '') + '\n' + lines.join('\n');
  }
  function tell(text) { try { if (collect && collect.origin) collect.popup.postMessage({ tz: 'say', text }, collect.origin); } catch (e) { /* вкладка закрыта */ } }
  function relayAsk(msg, ms) {
    const c = collect;
    return new Promise((resolve, reject) => {
      if (!c || c.finishing) { reject(new Error('сбор остановлен')); return; }
      const id = ++c.seq;
      const t = setTimeout(() => { c.pending.delete(id); reject(new Error('вкладка gov.kz не ответила')); }, ms || 90000);
      c.pending.set(id, m => { clearTimeout(t); resolve(m); });
      try { c.popup.postMessage(Object.assign({ id }, msg), c.origin); } catch (e) { clearTimeout(t); c.pending.delete(id); reject(e); }
    });
  }
  const relayFetch = (url, ms, method) => relayAsk({ tz: 'fetch', url, method: method || 'GET' }, ms);
  const relayFrame = (url, wait) => relayAsk({ tz: 'frame', url, wait: wait || 25000 }, (wait || 25000) + 20000);
  function onRelay(e) {
    const c = collect, m = e.data;
    if (!c || !m || typeof m.tz !== 'string' || e.source !== c.popup) return;
    if (!/^https:\/\/(www\.)?gov\.kz$/i.test(e.origin)) return;
    if (m.tz === 'fetched') { const p = c.pending.get(m.id); if (p) { c.pending.delete(m.id); p(m); } return; }
    if (m.tz === 'hello' && !c.connected) { c.origin = e.origin; c.connected = true; harvestAll(); }
  }
  window.addEventListener('message', onRelay);
  async function harvestAll() {
    const c = collect;
    showCollect();
    for (const s of c.sites) {
      if (collect !== c || c.finishing) return;
      c.cur = s;
      s.state = 'work';
      s.note = 'читаю список публикаций';
      showCollect();
      try { await harvestGov(s); s.state = 'done'; }
      catch (err) { s.state = 'fail'; s.note = (err && err.message) || String(err); await saveSiteFail(s); }
      showCollect();
    }
    if (collect === c && !c.finishing) finishCollect('');
  }
  function skipSite() {
    const c = collect;
    if (!c || !c.cur || c.cur.state !== 'work') return;
    c.cur.skip = true;                                     /* текущий регион бросить после ближайшего запроса */
  }
  async function regionDirOf(name) {
    const root = await store.dir.getDirectoryHandle(DGD_DIR, { create: true });
    return root.getDirectoryHandle(regionFolder(name), { create: true });
  }
  /* Регион не собран: в сводке это видно, а публикации прошлого сбора остаются (с пометкой) */
  async function saveSiteFail(s) {
    try {
      const dir = await regionDirOf(s.name);
      const old = (await readJsonFile(dir, DGD_META)) || { region: s.name, slug: s.slug, pages: [], rows: [] };
      old.state = s.state === 'skip' ? 'skip' : 'fail';
      old.note = s.note;
      old.failedAt = new Date().toISOString();
      await writeJsonFile(dir, DGD_META, old);
    } catch (e) { /* папка недоступна: в окне видно */ }
  }
  /* Один регион: список документов gov.kz (новые сначала) → публикации о процедурах граждан за DGD_WINDOW дней →
     страница каждой новой публикации в невидимой рамке → ссылки на файлы → файлы в папку региона → ИИН */
  async function harvestGov(s) {
    const base = GOV + '/org/' + s.slug;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const since = new Date(today); since.setDate(since.getDate() - DGD_WINDOW);
    const dir = await regionDirOf(s.name);
    const old = (await readJsonFile(dir, DGD_META)) || { rows: [] };
    const before = new Map((old.rows || []).map(r => [r.key, r]));
    const list = [], pages = [];
    for (let page = 1; page <= GOV_PAGES; page++) {
      if (s.skip) throw new Error('пропущен');
      tell(s.name + ': список документов, страница ' + page);
      const url = base + '/documents/' + page + '?lang=ru&size=50';
      const r = await relayFetch(url);
      if (!r.ok) {
        if (page === 1) throw new Error('раздел «Документы» на gov.kz не открылся' + (r.status ? ' (код ' + r.status + ')' : r.error ? ': ' + r.error : ''));
        break;
      }
      const fd = (govData(decodeText(new Uint8Array(r.bytes))).getFilteredData || {}).value;
      if (!fd || !Array.isArray(fd.data)) { if (page === 1) throw new Error('в разделе «Документы» на gov.kz нет списка'); break; }
      pages.push({ url, read: true, rows: fd.data.length });
      for (const x of fd.data) list.push({ id: String(x.id), title: clean(x.title || ''), date: dayOf(x.created_date) });
      const last = fd.data[fd.data.length - 1];
      if (!fd.data.length || !last || (dayOf(last.created_date) || today) < since || page * 50 >= (+fd.count || 0)) break;
    }
    const mine = list.filter(x => isGovCitizenDoc(x.title));
    const rows = [];
    let got = 0, n = 0;
    const fresh = mine.filter(x => x.date && x.date >= since);
    for (const x of mine) {
      const key = 'gov:' + x.id, was = before.get(key);
      const row = { key, id: x.id, date: x.date, kind: govKind(x.title), debtor: govDebtor(x.title), manager: (orderInfo([x.title], []).manager || ''),
        text: x.title, page: base + '/documents/details/' + x.id + '?lang=ru', firstSeen: was ? was.firstSeen : isoDay(new Date()), files: [], notes: [] };
      rows.push(row);
      const saved = was ? (was.files || []).filter(f => f.saved) : [];
      const yearly = govYearly(x.title);
      let ok = saved.length > 0 && !yearly;
      for (const f of saved) { try { await dir.getFileHandle(f.saved); } catch (e) { ok = false; } }
      if (ok) { row.files = saved; continue; }
      if (!(x.date && x.date >= since) && !yearly) continue;               /* старше окна: только считается */
      if (yearly) row.yearly = true;
      if (s.skip) throw new Error('пропущен');
      n++;
      s.note = 'публикация ' + n + ' из ' + fresh.length;
      showCollect();
      tell(s.name + ': публикация ' + n + ' из ' + fresh.length + ': ищу файл');
      let links = [];
      try {
        const fr = await relayFrame(row.page, 25000);
        links = (fr.links || []).map(l => l.href);
        if (!links.length) {                                            /* ссылки нет: адрес файла в ответах сайта странице */
          for (const api of (fr.resources || []).filter(u => !/\.(js|css|svg|png|woff2?)(\?|$)/i.test(u)).slice(0, 8)) {
            const j = await relayFetch(api, 30000);
            if (!j.ok || !j.bytes) continue;
            for (const mm of decodeText(new Uint8Array(j.bytes)).matchAll(/\/(?:uploads|api\/v1\/public\/assets|api\/v2\/(?:public\/)?files)\/[^"'\s\\<>]+/g)) links.push(GOV + mm[0]);
          }
        }
        if (!links.length && fr.text) row.notes.push('на странице нет файла' + (/mainFileNotExist|не найден/i.test(fr.text) ? ' (сайт пишет, что файла нет)' : ''));
      } catch (e) { row.notes.push('страница публикации не открылась: ' + e.message); }
      links = Array.from(new Set(links)).filter(u => !/\.(svg|png|jpe?g)$/i.test(u) || /original/i.test(u)).slice(0, 3);
      for (const href of links) {
        let r;
        try { r = await relayFetch(href, 180000); } catch (e) { row.files.push({ url: href, name: '', saved: '', err: e.message }); continue; }
        if (!r.ok || !r.bytes || !r.bytes.byteLength) { row.files.push({ url: href, name: '', saved: '', err: 'файл не скачался' + (r.status ? ' (код ' + r.status + ')' : '') }); continue; }
        const bytes = new Uint8Array(r.bytes);
        let fname = '';
        const disp = /filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i.exec(r.disposition || '');
        if (disp) { try { fname = decodeURIComponent(disp[1] || disp[2]); } catch (e) { fname = disp[1] || disp[2]; } }
        if (!fname) fname = decodeURIComponent(href.split('?')[0].split('/').pop() || '');
        const ext = (/\.([a-z0-9]{2,4})$/i.exec(fname) || [])[1] || (/pdf/i.test(r.type) ? 'pdf' : /png/i.test(r.type) ? 'png' : /jpe?g/i.test(r.type) ? 'jpg'
          : /wordprocessing|msword/i.test(r.type) ? 'docx' : /sheet|excel/i.test(r.type) ? 'xlsx' : 'pdf');
        const name = orderFileName(row, { name: fname.replace(/\.[a-z0-9]{2,4}$/i, '') + '.' + ext }, r.type);
        try { await writeFile(dir, name, bytes); } catch (e) { row.files.push({ url: href, name: fname, saved: '', err: 'не записан: ' + e.message }); continue; }
        row.files.push({ url: href, name: fname, saved: name, err: '' });
        got++;
        queueOcr(regionFolder(s.name) + '/' + name, bytes, name, r.type);
      }
      await sleep(150);
    }
    const meta = {
      version: VERSION, source: 'gov.kz', region: s.name, slug: s.slug, state: 'ok', note: '', collectedAt: new Date().toISOString(), window: DGD_WINDOW,
      pages: pages.map(p => ({ url: p.url, title: 'Документы ' + s.slug, year: null, rows: p.rows, read: p.read, err: '' })),
      rows: rows.map(r => ({ key: r.key, id: r.id, date: isoDay(r.date), dateFrom: r.yearly ? 'сводная за год, перечитана ' + dmy(new Date()) : '', kind: r.kind, debtor: r.debtor, manager: r.manager, number: '',
        text: r.text, page: r.page, future: false, firstSeen: r.firstSeen, notes: r.notes,
        files: r.files.map(f => ({ url: f.url, name: f.name || '', gov: false, saved: f.saved || '', err: f.err || '' })) })),
    };
    await writeJsonFile(dir, DGD_META, meta);
    s.note = 'публикаций за ' + DGD_WINDOW + ' дн. ' + fresh.length + ', новых файлов ' + got;
    const left = collect ? collect.sites.filter(x => x.state === 'wait').length : 0;
    tell(s.name + ': готово (' + s.note + ').' + (left ? ' Дальше: ещё ' + left + ' регион(ов).' : ' Это был последний регион: вкладку можно закрыть.'));
  }
  function orderFileName(row, f, type) {
    const ext = (/\.([a-z0-9]{2,4})$/i.exec(f.name) || [])[1] || (/pdf/i.test(type) ? 'pdf' : /png/i.test(type) ? 'png' : 'jpg');
    const base = (isoDay(row.date) || 'без-даты') + '_' + (safeName(row.debtor).slice(0, 40) || safeName(row.kind).slice(0, 30) || 'публикация') + '_' +
      (safeName(f.name.replace(/\.[a-z0-9]{2,4}$/i, '')).slice(0, 40) || row.id || 'файл');
    return clean(base).replace(/\s+/g, ' ') + '.' + ext.toLowerCase();
  }
  /* ИИН из скачанных файлов читаются по очереди, пока скачиваются следующие */
  function queueOcr(key, bytes, name, type) {
    const c = collect;
    c.ocrLeft++;
    c.ocr = c.ocr.then(async () => {
      try {
        if (!c.cache) c.cache = (await readJsonFile(await store.dir.getDirectoryHandle(DGD_DIR, { create: true }), DGD_OCR)) || {};
        const r = await readOrderFile(bytes, name, type);
        c.cache[key] = { size: bytes.length, at: new Date().toISOString(), note: r.note || '', iins: r.iins.map(x => ({ iin: x.iin, conf: x.conf, how: x.how, fixes: x.fixes || 0 })) };
      } catch (e) { if (c.cache) c.cache[key] = { size: bytes.length, at: new Date().toISOString(), note: (e && e.message) || String(e), iins: [] }; }
      c.ocrLeft--;
      showCollect();
    });
  }
  async function finishCollect(reason) {
    const c = collect;
    if (!c || c.finishing) return;
    c.finishing = true;
    clearInterval(c.watch);
    for (const p of c.pending.values()) p({ ok: false, status: 0, error: 'сбор остановлен' });
    c.pending.clear();
    try { if (!c.popup.closed) c.popup.close(); } catch (e) { /* уже закрыто */ }
    for (const s of c.sites) if (s.state === 'work' || s.state === 'wait') { s.state = 'skip'; s.note = 'не собран' + (reason ? ' (' + reason + ')' : ''); }
    showCollect();
    say('ДГД: дочитываю ИИН из скачанных файлов…');
    await c.ocr;
    try { if (c.cache) await writeJsonFile(await store.dir.getDirectoryHandle(DGD_DIR, { create: true }), DGD_OCR, c.cache); } catch (e) { /* перечитается */ }
    collect = null;
    await loadDgd();
    const wrote = await writeDgd(true);
    const ok = c.sites.filter(s => s.state === 'done').length;
    const lines = c.sites.map(s => ({ done: '✔ ', fail: '❌ ', skip: '⏭ ' }[s.state] || '· ') + s.name + (s.note ? ': ' + s.note : ''));
    showCollect('Сбор ДГД с gov.kz ' + ruDate(new Date()) + ': регионов собрано ' + ok + ' из ' + c.sites.length + (reason ? ' (' + reason + ')' : '') + '\n' + lines.join('\n'));
    say((ok === c.sites.length ? '✔ ' : '⚠ ') + 'ДГД собраны: ' + ok + ' из ' + c.sites.length + ' регионов. ' +
      (wrote ? 'Свод: ' + DGD_FILE + '. ' : '') + 'Проверка регионов ниже, в блоке «ДГД».' + (S ? ' Дальше: кнопка «Старт».' : ''));
    updateButtons();
  }

  function eta(pages) {
    const sec = Math.round(pages * Math.max(1000, interval) / 1000);
    return sec < 60 ? 'меньше минуты' : sec < 5400 ? Math.round(sec / 60) + ' мин' : Math.floor(sec / 3600) + ' ч ' + pad2(Math.round(sec % 3600 / 60)) + ' мин';
  }
  function updateButtons() {
    startB.disabled = running || !S || !!collect;
    collectB.disabled = running;
    stopB.disabled = !running;
    folderB.disabled = fileInput.disabled = fileSelect.disabled = sheetSelect.disabled = running;
    for (const b of regChecks.values()) b.disabled = running;
    visibleBox.disabled = baseInput.disabled = running;
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
    if (collect) { clearInterval(collect.watch); try { collect.popup.close(); } catch (e) { /* уже закрыто */ } collect = null; }
    window.removeEventListener('message', onRelay);
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
    const found = [];
    for await (const [name, handle] of dir.entries()) {
      if (handle.kind !== 'file' || !/\.(xlsx|xlsm|csv)$/i.test(name) || name.startsWith('~$') || name === DGD_FILE) continue;
      let at = 0;
      try { at = (await handle.getFile()).lastModified; } catch (e) { /* файл занят */ }
      found.push({ name, at });
    }
    found.sort((a, b) => b.at - a.at || a.name.localeCompare(b.name, 'ru'));      /* самый свежий файл первым */
    const names = found.map(x => x.name);
    store = { dir, handle: null };
    dgd = null;
    fileSelect.textContent = '';
    if (!names.length) {
      book = S = null; tasks = []; fileSelect.style.display = 'none';
      await loadDgd();
      say('❌ В этой папке нет файла с ИИН (.xlsx или .csv из ADW).' + (dgd && dgd.files ? ' Списки ДГД прочитаны: «Свод ДГД ↓».' : ''));
      return;
    }
    for (const n of names) { const o = h('option', '', n); o.value = n; fileSelect.append(o); }
    fileSelect.style.display = names.length > 1 ? 'block' : 'none';
    await openFromFolder(names[0]);
    await loadDgd();
  }
  /* Выгрузка ADW в CSV: рядом создаётся книга .xlsx с тем же именем, работа идёт в ней. Если CSV новее книги
     (новая выгрузка за день), прежняя книга уходит в tazalau_backup, а новая собирается из CSV заново. */
  async function csvToXlsx(bytes) {
    const rows = parseCsv(decodeText(bytes));
    if (!rows.length) throw new Error('файл CSV пустой');
    const widths = rows[0].map((x, i) => Math.min(60, Math.max(12, ...rows.slice(0, 200).map(r => String(r[i] || '').length + 2))));
    return makeXlsx([{ name: 'Лист1', rows, widths }]);
  }
  async function openFromFolder(name) {
    if (/\.csv$/i.test(name)) {
      try {
        const csv = await (await store.dir.getFileHandle(name)).getFile();
        const target = name.replace(/\.csv$/i, '.xlsx');
        let old = null;
        try { old = await (await store.dir.getFileHandle(target)).getFile(); } catch (e) { /* книги ещё нет */ }
        if (!old || old.lastModified < csv.lastModified) {
          const bytes = await csvToXlsx(new Uint8Array(await csv.arrayBuffer()));
          if (old) {
            const bdir = await store.dir.getDirectoryHandle(BACKUP_DIR, { create: true });
            await writeFile(bdir, target.replace(/\.xlsx$/i, '__' + stampOf(new Date(old.lastModified)) + '.xlsx'), await old.arrayBuffer());
          }
          await writeFile(store.dir, target, bytes);
        }
        if (!Array.from(fileSelect.options).some(o => o.value === target)) { const o = h('option', '', target); o.value = target; fileSelect.append(o); }
        fileSelect.value = target;
        return openFromFolder(target);
      } catch (e) { book = S = null; tasks = []; refresh(); say('❌ ' + (e.message || e)); return; }
    }
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
      if (!handle) { store.dir = null; dgd = null; }
      baseAuto = '';
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
      if (handle) {                    /* путь к этой папке: вставленный раньше, иначе записанный в файле самим Excel */
        const fromFile = normBase(b.absPath);
        if (baseValid(fromFile) && sameName(lastPart(fromFile), store.dir.name)) baseAuto = fromFile;
        store.saved = await readSaved();
        baseInput.value = store.saved || baseAuto;
      }
      useSheet(active.sheet.name);
      say(handle ? 'Можно запускать. В этой папке появится папка на каждого человека, в ней ' + SHOT_NAME + '. Ссылки: прямо в файл.'
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

  /* Найденные записи: строка таблицы результатов как { заголовок столбца: значение } */
  function readRecords(d) {
    const table = d.querySelector('.table-responsive table') || d.querySelector('table.table');
    if (!table) return [];
    const heads = Array.from(table.querySelectorAll('thead th')).map(th => clean(th.textContent));
    const out = [];
    table.querySelectorAll('tbody > tr').forEach(tr => {
      const cells = Array.from(tr.children);
      if (cells.length === 1 && cells[0].colSpan > 1) return;       /* «Нет записей» */
      const rec = {};
      cells.forEach((td, i) => { const v = clean(td.innerText || td.textContent); if (v) rec[heads[i] || 'Столбец ' + (i + 1)] = v; });
      out.push(rec);
    });
    return out;
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
  /* @font-face: шрифт вкладывается, только если страница его действительно загрузила, и одним файлом.
     Из перечисленных в правиле файлов берётся тот, который загрузила сама страница, иначе первый годный по порядку
     (так выбирает и браузер). Форматы embedded-opentype (.eot) и svg браузер не использует: их не берём. */
  async function inlineFonts(text, base, loaded) {
    const blocks = [];
    text.replace(/@font-face\s*\{[^}]*\}/gi, m => { blocks.push(m); return m; });
    const map = new Map();
    await Promise.all(blocks.map(async block => {
      const family = fontKey((/font-family\s*:\s*([^;}]+)/i.exec(block) || [])[1]);
      if (!loaded.fonts.has(family)) { map.set(block, ''); return; }
      const urls = [];
      block.replace(/url\(\s*["']?([^"')]+)["']?\s*\)(?:\s*format\(\s*["']?([^"')]+)["']?\s*\))?/gi, (m, u, f) => { urls.push([u, (f || '').toLowerCase()]); return m; });
      const abs = x => new URL(x[0], base).href.split('#')[0];
      const usable = urls.filter(([u, f]) => !u.startsWith('data:') &&
        (f ? /^(woff2?|truetype|opentype|collection)(-variations)?$/.test(f) : !/\.(eot|svg)(\?|#|$)/i.test(u)));
      const best = usable.find(x => { try { return loaded.urls.has(abs(x)); } catch (e) { return false; } }) || usable[0];
      if (!best) return;
      try {
        const data = await toDataUrl(abs(best));
        const rest = block.slice(block.indexOf('{') + 1, block.lastIndexOf('}')).split(';').filter(d => d.trim() && !/^\s*src\s*:/i.test(d)).join(';');
        map.set(block, '@font-face{' + rest + ';src:url("' + data + '")' + (best[1] ? ' format("' + best[1] + '")' : '') + '}');
      } catch (e) { map.set(block, ''); }
    }));
    return text.replace(/@font-face\s*\{[^}]*\}/gi, m => (map.has(m) ? map.get(m) : m));
  }
  async function inlineCss(text, base, loaded, depth) {
    text = await inlineFonts(text, base, loaded);
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
    const dots = new Uint32Array(px.buffer, px.byteOffset, px.length >> 2);
    let sig = 0, lo = 255, hi = 0;
    for (let i = 0; i < dots.length; i++) sig = (Math.imul(sig, 31) + dots[i]) | 0;      /* отпечаток по всем точкам: видна и мелочь */
    for (let i = 0; i < px.length; i += 148) {                /* выборка точек: разброс яркости */
      const y = (px[i] * 3 + px[i + 1] * 6 + px[i + 2]) / 10;
      if (y < lo) lo = y;
      if (y > hi) hi = y;
    }
    return { canvas, sig, spread: hi - lo };
  }
  /* Шрифты внутри SVG браузер подключает не сразу. Первый проход после запуска помощника рисуется без части из них
     (на сайте так пропадал значок в меню), и проходы подряд, без паузы, выходят такими же. Поэтому между проходами
     пауза, а кадр берётся, только когда совпали два прохода, разделённые паузой. */
  async function shoot(d, width) {
    const b = await buildSvg(d, width);
    let prev = await drawSvg(b);
    for (let i = 0; i < 6; i++) {
      await sleep(T.SETTLE);
      const next = await drawSvg(b);
      const same = next.sig === prev.sig;
      prev = next;
      if (same) break;
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
      const st = s.records && s.records.length ? recordInfo(s.reg, s.records[0]).status : '';
      const l1 = 'ИИН ' + s.iin + SEP + s.reg.group + ' → ' + s.reg.title + SEP + 'Всего записей: ' + s.count +
        (s.rows < s.count ? ' (на снимке первые ' + s.rows + ')' : '') + (st ? SEP + 'Статус: ' + st : '');
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
      const records = count ? readRecords(d) : [];
      const at = new Date();
      const canvas = await shoot(d, frameW);
      timings.push({ load: t1 - t0, shot: Date.now() - at.getTime() });
      if (fdoc() !== d) throw new SiteError('страница сменилась во время снимка');
      return { reg, iin, url, count, rows, records, at, canvas, clipped: info.overflow > 0 };
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

  /* Статус записи: какой столбец что значит, по заголовкам таблицы реестра (у судебных и внесудебных реестров они разные) */
  const DATE_LABELS = [
    [/иска/i, 'иск'], [/принятия к производству/i, 'принято к производству'], [/решения суда/i, 'решение суда'],
    [/вступления в силу/i, 'вступило в силу'], [/завершения дела/i, 'дело завершено'],
    [/подачи заявления.*прекращ/i, 'заявление о прекращении'], [/подачи заявления/i, 'заявление подано'],
    [/начала процедуры/i, 'начало процедуры'], [/завершения процедуры/i, 'процедура завершена'],
    [/прекращения процедуры/i, 'процедура прекращена'],
  ];
  const keyOf = (rec, re) => Object.keys(rec).find(k => re.test(k));
  const valOf = (rec, re) => { const k = keyOf(rec, re); return k ? rec[k] : ''; };
  const dateOf = s => { const m = /(\d{2})\.(\d{2})\.(\d{4})/.exec(s || ''); return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null; };
  function recordInfo(reg, rec) {
    let status = valOf(rec, /статус дела tazalau/i) || valOf(rec, /статус заявителя/i) || valOf(rec, /статус дела/i);
    if (reg.key === 'debtor_terminated') {
      const who = valOf(rec, /инициатор/i);
      status = 'процедура прекращена' + (who ? ' (инициатор: ' + who + ')' : '');
    }
    const events = [];
    for (const k of Object.keys(rec)) {
      const d = /^дата/i.test(k) && !/передачи судье/i.test(k) ? dateOf(rec[k]) : null;
      if (!d) continue;
      const label = (DATE_LABELS.find(([re]) => re.test(k)) || [null, k.toLowerCase()])[1];
      events.push({ d, label, text: /\d{2}\.\d{2}\.\d{4}/.exec(rec[k])[0] });
    }
    events.sort((a, b) => a.d - b.d);
    /* Возвращённый, отклонённый, прекращённый иск процедуры не открывает: такая запись важна меньше, чем идущее дело */
    const closed = /возвращ|отказ|без рассмотрения|прекращ|отмен/i.test(status);
    return { reg, rec, status: status || 'запись есть', events, last: events[events.length - 1] || null, first: events[0] || null, closed,
      category: valOf(rec, /категория дела/i), place: [valOf(rec, /^регион/i), valOf(rec, /наименование суда|услугодатель/i)].filter(Boolean).join(', '),
      creditors: valOf(rec, /список кредиторов/i), sums: Object.keys(rec).filter(k => /[сc]умма задолженности/i.test(k)).map(k => rec[k]),
      fio: valOf(rec, /^фио/i) };
  }
  function recordLine(x) {
    const parts = [x.reg.short + ': ' + x.status];
    if (x.category) parts.push(x.category);
    if (x.events.length) parts.push(x.events.map(e => e.label + ' ' + e.text).join(', '));
    if (x.place) parts.push(x.place);
    if (x.creditors) parts.push('кредиторы: ' + x.creditors.slice(0, 400) + (x.sums.length ? '; сумма, тг: ' + x.sums.join(' / ') : ''));
    return parts.join(' · ');
  }
  /* Фамилия в файле и на сайте (на сайте «Фамилия И. О.»): разные фамилии стоит перепроверить */
  const surname = s => String(s || '').trim().split(/\s+/)[0].toLowerCase().replace(/ё/g, 'е')
    .replace(/[әӘ]/g, 'а').replace(/[іІ]/g, 'и').replace(/[ңҢ]/g, 'н').replace(/[ғҒ]/g, 'г').replace(/[үұҮҰ]/g, 'у')
    .replace(/[қҚ]/g, 'к').replace(/[өӨ]/g, 'о').replace(/[һҺ]/g, 'х').replace(/[^a-zа-я-]/g, '');

  function describe(shots, fioInFile) {
    const found = shots.filter(s => s.count > 0);
    let text = found.length
      ? 'НАЙДЕН: ' + found.map(s => s.reg.short + ', записей: ' + s.count + (s.rows < s.count ? ' (на снимке первые ' + s.rows + ')' : '')).join('; ')
      : shots.length === 1 ? 'НЕ НАЙДЕН: ' + shots[0].reg.short : 'НЕ НАЙДЕН ни в одном из ' + shots.length + ' реестров';
    if (shots.some(s => s.clipped)) text += ' · внимание: таблица шире снимка';
    const infos = [];
    for (const s of found) for (const rec of s.records || []) infos.push(recordInfo(s.reg, rec));
    let status = found.length ? 'НАЙДЕН (статус не прочитан)' : 'нет записей ни в одном из проверенных реестров';
    if (infos.length) {
      /* Главная запись: идущее дело важнее закрытого; среди равных самое новое дело (по дате подачи), затем последнее событие */
      const t = e => (e ? e.d.getTime() : 0);
      const rank = x => [x.closed ? 0 : 1, t(x.first), t(x.last)];
      const better = (a, b) => { const ra = rank(a), rb = rank(b); for (let k = 0; k < 3; k++) if (ra[k] !== rb[k]) return ra[k] > rb[k] ? a : b; return a; };
      const latest = infos.reduce(better);
      status = latest.reg.short + ': ' + latest.status + (latest.last ? ' · ' + latest.last.label + ' ' + latest.last.text : '') +
        (infos.length > 1 ? ' (всего записей: ' + infos.length + ')' : '');
    }
    const notes = [];
    const own = surname(fioInFile);
    const other = infos.map(x => x.fio).find(f => f && own && surname(f) && surname(f) !== own);
    if (other) notes.push('внимание: на сайте фамилия «' + other.split(/\s+/)[0] + '», в файле «' + String(fioInFile).trim().split(/\s+/)[0] + '»');
    const details = infos.map(recordLine).concat(notes).join(' | ').slice(0, 32000);
    return { found: found.length > 0, text: text + ' · проверено ' + ruDate(shots[shots.length - 1].at, false), status, details, infos };
  }

  /* ---------- запуск ---------- */
  /* Путь для столбца hyperlink перед запуском: годный путь, '' (без столбца) или null (запуск отменён, путь надо поправить) */
  async function baseForRun() {
    if (!store.handle) return '';
    const b = baseState();
    if (!b.base) {
      if (!b.bad && store.saved === '') return '';                     /* уже отвечали «без hyperlink» */
      const a = await ask(b.bad ? b.text + '\nСтолбец ' + ABS_HEADER + ' останется пустым.'
        : 'Путь к папке на диске не указан: столбец ' + ABS_HEADER + ' останется пустым.\n' +
          'Вставьте путь из адресной строки Проводника в поле «Путь к этой папке» и нажмите Старт ещё раз.',
        ['Указать путь', 'Без ' + ABS_HEADER]);
      if (a !== 'Без ' + ABS_HEADER) { baseInput.focus(); return null; }
      if (!b.bad) await writeSaved('');
      return '';
    }
    if (b.mismatch) {
      const a = await ask('Путь кончается папкой «' + lastPart(b.base) + '», а выбрана папка «' + store.dir.name + '».\n' +
        'Если это разные папки, ссылки в столбце ' + ABS_HEADER + ' не откроются.', ['Исправить путь', 'Записать так']);
      if (a !== 'Записать так') { baseInput.focus(); return null; }
    }
    await writeSaved(b.base === baseAuto ? null : b.base);      /* путь из файла Excel запоминать незачем */
    return b.base;
  }

  async function run() {
    if (running || !S || pendingAsk) return;
    const regs = chosen();
    if (!regs.length) { say('❌ Отметьте хотя бы один реестр.'); regBox.open = true; return; }
    let recheck = false, p = plan(false);
    if (!p.groups.size) {
      if (!tasks.length) { say('В столбце «ИИН/БИН» нет значений.'); return; }
      if (!p.linked && !p.invalid.length) { say('Все строки с ИИН скрыты фильтром Excel: снимите галочку «только видимые строки».'); return; }
    }
    const base = await baseForRun();
    if (base == null) return;
    if (!p.groups.size && p.linked) { /* новых ИИН нет (строки с опечатками не в счёт): спрашиваем, проверять ли заново */
      const a = await ask('Новых ИИН нет: ссылка уже записана.\nПроверить все ИИН заново? Старые снимки останутся.', ['Все заново', 'Отмена']);
      if (a !== 'Все заново') return;
      recheck = true;
      p = plan(true);
    }
    running = true; stopping = false; interval = T.INTERVAL;
    updateButtons();
    say('Проверка началась…');
    let done = 0, found = 0, notFound = 0, errors = 0, fails = 0, diskFails = 0, unsaved = 0, written = 0, stopNote = '';
    const total = p.groups.size;
    try {
      if (store.handle) {
        const w = await store.handle.createWritable({ keepExistingData: true });   /* файл не занят? выясняем до работы */
        await w.close();
      }
      const useDgd = !!(dgd && dgd.byIin);
      ensureColumns(S, !!base, useDgd);
      const L = S.layout;
      if (useDgd) {
        await writeDgd(true);
        await fillDgd(S);
        p.groups = new Map(Array.from(p.groups).sort((a, b) => (dgd.byIin.has(b[0]) ? 1 : 0) - (dgd.byIin.has(a[0]) ? 1 : 0)));
      }
      for (const t of p.invalid) {
        try { setText(S, t.row, L.resultCol, 'ОШИБКА: ' + t.parse.err, null); } catch (e) { /* объединённая ячейка */ }
        logRow({ time: ruDate(new Date(), true), iin: clean(t.raw) || '?', rows: String(t.row), status: 'НЕВЕРНЫЙ ИИН', counts: '', file: '', note: t.parse.err });
      }
      const names = folderNames(base ? folderRoom(base) : 100);
      const fioByRow = new Map(tasks.map(t => [t.row, t.fio]));
      for (const [iin, rows] of p.groups) {
        if (stopping) break;
        say('Проверяю ' + (done + 1) + ' из ' + total + ': ' + iin + '\nнайдено ' + found + ' · не найдено ' + notFound + ' · ошибок ' + errors);
        let res;
        try { res = await checkIin(regs, iin); } catch (e) { if (e instanceof Stopped) break; throw e; }
        if (res.error) {
          errors++; fails++;
          for (const r of rows) {
            try { setText(S, r, L.resultCol, 'ОШИБКА: ' + res.error, null); setText(S, r, L.statusCol, '', null); setText(S, r, L.detailsCol, '', null); }
            catch (e) { /* объединённая ячейка */ }
          }
          logRow({ time: ruDate(new Date(), true), iin, rows: rows.join(', '), status: 'ОШИБКА', counts: '', file: '', note: res.error });
        } else {
          fails = 0;
          const d = describe(res.shots, fioByRow.get(rows[0])), folder = names.get(iin) || iin;
          const blob = await stamped(res.shots);
          let target = '', lost = '';
          if (store.handle) {                      /* папка человека рядом с файлом Excel, в ней «снимок.png» */
            try {
              const personDir = await store.dir.getDirectoryHandle(folder, { create: true });
              await keepOld(personDir, base ? WIN_PATH - SWAP - base.length - folder.length - 2 : 1000);
              await writeFile(personDir, SHOT_NAME, blob);
              target = folder + '/' + SHOT_NAME;
            } catch (e) { lost = 'снимок не сохранён в папку «' + folder + '»: ' + (e && e.message ? e.message : e); }
          } else {                                 /* в «Загрузки» папку создать нельзя: имя человека в начале имени файла */
            target = folder + ' ' + (recheck ? SHOT_NAME.replace(/\.png$/, ' ' + stampOf(res.shots[res.shots.length - 1].at) + '.png') : SHOT_NAME);
            saveBlob(blob, target);
          }
          if (lost) {                              /* без снимка ссылку не пишем: строка останется на следующий запуск */
            errors++; diskFails++;
            for (const r of rows) {
              try { setText(S, r, L.resultCol, 'ОШИБКА: ' + lost, null); setText(S, r, L.statusCol, d.status, null); setText(S, r, L.detailsCol, d.details, null); }
              catch (e) { /* объединённая ячейка */ }
            }
            logRow({ time: ruDate(new Date(), true), iin, rows: rows.join(', '), status: 'ОШИБКА', counts: '', file: '', note: lost });
          } else {
            diskFails = 0;
            const full = base ? base + '\\' + folder + '\\' + SHOT_NAME : '';      /* полный путь к снимку на диске */
            let note = '';
            for (const r of rows) {
              try {
                const cur = (S.cells.get(r) || new Map()).get(L.linkCol);
                setText(S, r, L.linkCol, target.replace(/\//g, '\\'), await linkStyle(S, cur ? cur.getAttribute('s') : null));
                await setLink(S, r, L.linkCol, target);
                if (full) {
                  const was = (S.cells.get(r) || new Map()).get(L.absCol);
                  await setPathLink(S, r, L.absCol, full, await linkStyle(S, was ? was.getAttribute('s') : null, true));
                }
                setText(S, r, L.resultCol, d.text, null);
                setText(S, r, L.statusCol, d.status, null);
                setText(S, r, L.detailsCol, d.details, null);
                written++;
              } catch (e) { note = e.message; }
            }
            if (d.found) found++; else notFound++;
            logRow({ time: ruDate(new Date(), true), iin, rows: rows.join(', '), status: d.found ? 'НАЙДЕН' : 'НЕ НАЙДЕН',
              counts: res.shots.map(s => s.reg.key + '=' + s.count).join('; '), file: full || target, note: note || (d.found ? d.status : '') });
          }
        }
        done++; unsaved++;
        if (fails >= MAX_FAILS) { stopNote = 'Сайт не отвечает: ' + fails + ' ИИН подряд не удалось проверить. Попробуйте позже.'; break; }
        if (diskFails >= MAX_FAILS) { stopNote = 'Снимки не сохраняются в папку: ' + diskFails + ' раза подряд. Проверьте доступ к папке и место на диске.'; break; }
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
      '\nСсылок записано: ' + written + (base && written ? ' (полный путь: столбец ' + ABS_HEADER + ')' : '') +
      (store.handle ? '' : '\nФайлы пришли в «Загрузки»: держите их в одной папке.') +
      (recheck ? '\nПроверка заново: прежние снимки остались в тех же папках, с датой или номером в имени.' : ''));
    if (journal.length && store.handle) {
      try {
        const bdir = await store.dir.getDirectoryHandle(BACKUP_DIR, { create: true });
        await writeFile(bdir, 'журнал_tazalau_' + stampOf(new Date()) + '.csv', csvBlob());
      } catch (e) { /* журнал доступен кнопкой CSV */ }
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
    test: { normIin, checkDigit, openBook, loadSheet, readTasks, ensureColumns, setText, setLink, setPathLink, linkStyle, bookBytes, zipRead,
      useDir, openBytes, run, timings, normBase, baseValid, folderRoom, go, buildSvg, drawSvg,
      setRegs: keys => { for (const [k, box] of regChecks) box.checked = keys.includes(k); refresh(); },
      describe, recordInfo, readRecords, makeXlsx, parseCsv, decodeText, pdfLines, fileRows, extractIins, regionsIn, readDgd, dgdWorkbook,
      readScan, readOrderFile, ocrIins, grayOf, ocrPlausible, orderInfo, samePerson, GOV_SITES, govData, govKind, govDebtor, GOV_DOC_RE, GOV_NOT_RE, isGovCitizenDoc, govYearly,
      collecting: () => collect && { sites: collect.sites.map(x => ({ key: x.key, name: x.name, state: x.state, note: x.note })), ocrLeft: collect.ocrLeft, finishing: !!collect.finishing },
      relayFetch: url => relayFetch(url, 15000),
      learnPage: async (bytes, name) => { const c = await window.__tazalau.test.debugPage(bytes, name); const dbg = []; dbg.learn = []; if (c) ocrIins(grayOf(c), dbg); return dbg.learn; },
      debugPage: async (bytes, name) => { if (/\.pdf$/i.test(name)) { const doc = await pdfOpen(bytes); return pdfPageImage(doc, doc.pages[0], OCR_DPI); } return imageFileCanvas(bytes, 'image/jpeg'); },
      state: () => ({ running, journal, tasks: tasks.length, sheet: S && S.sheet.name, layout: S && S.layout, base: baseState().base, baseAuto,
        dgd: dgd === 'reading' ? 'reading' : dgd && { files: dgd.files.length, iins: dgd.byIin.size, orders: (dgd.orders || []).length,
          regions: dgd.regions.map(r => ({ name: r.name, level: r.level, text: r.text, unique: r.unique, win: r.win, year: r.year, read: r.read, unread: r.unread })) } }) },
  };
  layoutFrame();
  lastRequest = Date.now();           /* первая страница в рамке тоже запрос: следующий пойдёт не раньше, чем через паузу */
  frame.src = START_URL;
  refresh();
})();
