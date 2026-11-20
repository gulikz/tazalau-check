# Собирает ../Tazalau_helper.html (страница установки закладки + краткая инструкция) из tazalau_helper.js.
# usage: python3 build.py            (нужен node/npx для сжатия кода; без них закладка собирается из несжатого кода)
import html
import re
import shutil
import subprocess
import sys
import urllib.parse
from pathlib import Path

DEV = Path(__file__).resolve().parent
ROOT = DEV.parent
SOURCE = ROOT / "tazalau_helper.js"
MIN = DEV / "tazalau_helper.min.js"
OUT = ROOT / "Tazalau_helper.html"          # имя латиницей: так файл без искажений скачивается с GitHub

source = SOURCE.read_text(encoding="utf-8")
version = re.search(r"const VERSION = '([^']+)'", source).group(1)

def light_minify(text: str) -> str:
    """Сжатие без сторонних программ: убирает комментарии и отступы, сам код не трогает.

    Разбирает текст посимвольно, различая строки, шаблоны и регулярные выражения, поэтому «/*» внутри
    строки или выражения комментарием не считается. Комментарии вида // в коде помощника не используются.
    """
    out, i, n = [], 0, len(text)
    prev = ""                                   # последний значимый символ: по нему отличаем деление от выражения
    while i < n:
        ch = text[i]
        if ch in "\"'`":                        # строка или шаблон: копируем целиком
            j = i + 1
            while j < n and text[j] != ch:
                j += 2 if text[j] == "\\" else 1
            out.append(text[i:j + 1]); prev = ch; i = j + 1
        elif ch == "/" and text[i + 1:i + 2] == "*":          # комментарий
            i = text.index("*/", i + 2) + 2
        elif ch == "/" and (prev == "" or prev in "(,=:[!&|?{};\n" or prev in "+-*%<>~^"):   # регулярное выражение
            j, in_class = i + 1, False
            while j < n and (text[j] != "/" or in_class):
                if text[j] == "\\":
                    j += 1
                elif text[j] == "[":
                    in_class = True
                elif text[j] == "]":
                    in_class = False
                j += 1
            out.append(text[i:j + 1]); prev = "/"; i = j + 1
        else:
            out.append(ch)
            if not ch.isspace():
                prev = ch
            elif ch == "\n" and prev not in ")]}":
                prev = "\n" if prev in "({[,;=" else prev
            i += 1
    lines = [line.strip() for line in "".join(out).splitlines()]
    return "\n".join(line for line in lines if line)


MINIFY = "--no-minify" not in sys.argv      # --no-minify: ничего не скачивать и не запускать, собрать из несжатого кода
if MINIFY and shutil.which("node"):
    subprocess.run(["node", "--check", str(SOURCE)], check=True)
code = None
if MINIFY and shutil.which("npx"):
    done = subprocess.run(["npx", "--yes", "terser@5.31.6", str(SOURCE), "--ecma", "2020", "--compress", "--mangle",
                           "--comments", "false", "-o", str(MIN)], capture_output=True, text=True)
    if done.returncode == 0:
        subprocess.run(["node", "--check", str(MIN)], check=True)
        code = MIN.read_text(encoding="utf-8").strip()
    else:
        print("сжать код не удалось, беру несжатый:", done.stderr.strip()[:200])
if code is None:
    code = light_minify(source)
    MIN.write_text(code, encoding="utf-8")
    if shutil.which("node"):
        subprocess.run(["node", "--check", str(MIN)], check=True)     # сжатый код обязан остаться корректным
if "—" in source:
    sys.exit("в коде помощника есть длинное тире: его быть не должно")
bookmarklet = "javascript:" + urllib.parse.quote(code, safe="-_.!~*'()")

page = """<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tazalau помощник</title>
<style>
  :root { --ink:#1d2433; --muted:#5b6577; --line:#dde3ec; --bg:#f5f7fa; --card:#fff; --accent:#1f6feb; }
  body { margin:0; background:var(--bg); color:var(--ink); font:15px/1.5 Arial, sans-serif; }
  main { max-width:720px; margin:0 auto; padding:20px 16px 32px; }
  h1 { font-size:22px; margin:0 0 2px; }
  .sub, .small { color:var(--muted); font-size:13px; margin:0 0 12px; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:14px 18px; margin:10px 0; }
  ol, ul { padding-left:20px; margin:4px 0; } li { margin:6px 0; }
  .bm { display:inline-block; padding:10px 20px; border-radius:8px; background:var(--accent); color:#fff;
        font-weight:bold; font-size:17px; text-decoration:none; cursor:grab; margin:6px 0; }
  button { font:inherit; padding:6px 12px; border-radius:6px; border:1px solid #b9c3d3; background:#fff; cursor:pointer; }
  .ok { color:#1a7f37; font-weight:bold; margin-left:6px; }
  .hint { color:#9a3412; font-weight:bold; margin-left:8px; }
  .check { margin:6px 0 0; padding:8px 10px; border-radius:8px; background:#eef3fb; }
  .check.good { background:#dafbe1; color:#1a7f37; font-weight:bold; }
  .bad { background:#ffebe9; border-color:#cf222e; font-weight:bold; }
  details { margin-top:8px; } summary { cursor:pointer; color:var(--accent); }
  table { border-collapse:collapse; width:100%; font-size:14px; }
  td { border-top:1px solid var(--line); padding:5px 6px; vertical-align:top; }
  td:first-child { font-weight:bold; white-space:nowrap; }
  kbd { border:1px solid #b9c3d3; border-bottom-width:2px; border-radius:4px; padding:0 4px; font-size:13px; background:#fff; white-space:nowrap; }
  code { background:#eef2f6; padding:1px 5px; border-radius:4px; }
  textarea { position:absolute; left:-9999px; }
</style>
</head>
<body>
<main id="tazalau-installer">
  <h1>Tazalau помощник</h1>
  <p class="sub">Версия __VERSION__ · Edge или Chrome · ничего не устанавливается · ЭЦП и вход на сайт не нужны</p>
  <div class="card bad" id="browserWarn" hidden></div>

  <div class="card">
    <ol>
      <li>Перетащите кнопку на панель избранного (если панели нет, нажмите <kbd data-mac="⌘ Cmd">Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>B</kbd>):<br>
        <a class="bm" id="bm" href="__BOOKMARKLET__" title="Перетащите на панель избранного">Tazalau ▶</a><span class="hint" id="bmHint"></span></li>
      <li>Проверка: нажмите эту закладку здесь же.
        <p class="check" id="tazalau-check">Здесь должна появиться «✔ Закладка работает».</p></li>
      <li>Каждый день, три шага:
        <ol type="a">
          <li>положите выгрузку ADW (xlsx или csv, клиенты с признаком «проблема», столбцы <code>ИИН/БИН</code> и <code>FIO</code>) в рабочую папку,
            откройте <a href="https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery" target="_blank" rel="noopener">tazalau.qoldau.kz</a> →
            закладка <b>Tazalau ▶</b> → <b>Папка с Excel</b>;</li>
          <li><b>🌐 Собрать ДГД с gov.kz</b>: откроется вкладка gov.kz, нажмите в ней закладку <b>Tazalau ▶</b> один раз. Помощник сам
            соберёт публикации ДГД всех 20 регионов о процедурах граждан за 31 день (приказы о финансовых управляющих, объявления о
            возбуждении дел, реестры требований кредиторов) и прочитает ИИН из файлов. Затем посмотрите блок <b>ДГД</b>: красные и жёлтые регионы;</li>
          <li><b>▶ Старт</b>: каждый ИИН выгрузки проверяется на tazalau (сначала те, что есть в публикациях ДГД): снимок, статус, ссылки.</li>
        </ol>
        На вопрос браузера о доступе к папке нажмите «Изменить файлы» или «Разрешить». Файл Excel перед этим закройте.</li>
    </ol>
    <p class="small"><b>Столбец hyperlink.</b> В нём полная ссылка: путь к папке на диске + папка человека + <code>снимок.png</code>.
      Путь к папке помощник берёт из самого файла Excel, если файл сохранён программой Excel в этой же папке. Если поле
      «Путь к этой папке» в окне помощника пустое, вставьте путь один раз: откройте папку в Проводнике, щёлкните по адресной
      строке, скопируйте (<kbd>Ctrl</kbd>+<kbd>C</kbd>) и вставьте в поле. Помощник запомнит его в этой же папке (в <code>tazalau_backup</code>), для всех, кто в ней работает.</p>
    <details>
      <summary>Что получится</summary>
      <table>
        <tr><td>папки людей</td><td>рядом с файлом Excel появится папка на каждого человека с именем из столбца <code>FIO</code>,
          внутри файл <code>снимок.png</code>: вверху снимка ИИН, реестр, дата, время и адрес. Если столбца FIO нет или имя не указано,
          папка называется по ИИН. Если папка человека уже есть, снимок кладётся в неё</td></tr>
        <tr><td>hyperlink</td><td>столбец в файле Excel: полная ссылка на снимок, например
          <code>W:\\Отдел\\Папка\\Иванов Иван\\снимок.png</code>. Открывается щелчком из любого места, пока снимок лежит по этому пути</td></tr>
        <tr><td>tazalau_file_location</td><td>столбец в файле Excel: короткая ссылка на тот же снимок (от папки с файлом Excel).
          Работает и после переноса всей папки в другое место</td></tr>
        <tr><td>tazalau_result</td><td>столбец в файле Excel: НАЙДЕН (в каких реестрах), НЕ НАЙДЕН или ОШИБКА с причиной и датой проверки</td></tr>
        <tr><td>tazalau_status</td><td>главный статус: реестр, статус дела или заявителя и последнее событие с датой. Идущее дело важнее
          возвращённого или отклонённого иска</td></tr>
        <tr><td>tazalau_details</td><td>все найденные записи: статус, даты (иск, решение, начало и завершение процедуры), суд или ЦОН,
          кредиторы и суммы. Если фамилия на сайте другая, здесь будет предупреждение</td></tr>
        <tr><td>ДГД</td><td>столбец в файле Excel: в публикации какого региона есть этот ИИН (приказ, объявление, реестр требований), дата,
          финансовый управляющий; ссылка открывает сам файл. Если ИИН в файле не прочитан, а ФИО совпадает: «⚠ возможно (по ФИО)», сверьте вручную</td></tr>
        <tr><td>ДГД/&lt;регион&gt;</td><td>скачанные с gov.kz файлы (PDF, сканы, Word) с датой и ФИО в имени. ИИН помощник читает сам, со скана тоже;
          прочитанное с подбором цифры помечено «сверить»</td></tr>
        <tr><td>ДГД_свод.xlsx</td><td>свод: лист «Регионы» (все 20: раздел на gov.kz не открылся, 0 или 1–2 публикации за 31 день, последняя давно,
          ИИН не прочитан, нет файла), лист «Публикации» (каждая со ссылками на файл и на gov.kz), «ИИН» и «Файлы»</td></tr>
        <tr><td>tazalau_backup</td><td>копия файла Excel, сделанная перед записью, журнал проверки и файл с путём к папке, если его вставляли вручную</td></tr>
      </table>
      <p class="small">Проверяются все шесть реестров сайта (снять лишние можно в блоке «Реестры»). Скорость: одна страница примерно
        за 2 секунды, быстрее сайт не разрешает; шесть реестров на ИИН: около 12 секунд, около 300 ИИН в час. Можно остановить и продолжить позже:
        строки, где ссылка уже есть, пропускаются. При проверке заново прежний снимок остаётся в той же папке, с датой или номером в имени.
        Windows не открывает файлы с путём длиннее 259 знаков: если путь к папке длинный, слишком длинные имена папок людей укорачиваются.</p>
    </details>
    <details>
      <summary>Нет кнопки «Папка с Excel» или папку выбрать не даёт</summary>
      <p>Нажмите «или один файл» и выберите файл Excel. Снимки и копия файла со ссылками (имя заканчивается на <code>_tazalau</code>)
        придут в «Загрузки»; на вопрос о скачивании нескольких файлов нажмите «Разрешить». Держите их в одной папке.
        Папки в «Загрузках» создать нельзя, поэтому там снимок называется «ФИО снимок.png».</p>
    </details>
    <details>
      <summary>Не перетаскивается</summary>
      <p><button id="copyBm" type="button">Скопировать закладку</button><span class="ok" id="okBm"></span>
        → <kbd data-mac="⌘ Cmd">Ctrl</kbd>+<kbd>D</kbd> → «Ещё» → в поле URL <kbd data-mac="⌘ Cmd">Ctrl</kbd>+<kbd>A</kbd>,
        <kbd data-mac="⌘ Cmd">Ctrl</kbd>+<kbd>V</kbd> → «Сохранить». Не в адресную строку.</p>
    </details>
    <details>
      <summary>Если не работает</summary>
      <ul>
        <li>Проверка молчит: добавьте закладку ещё раз. Если не помогло, такие закладки блокирует ИТ банка: нужна заявка в ИТ (ниже),
          обходить не нужно.</li>
        <li>«ОШИБКА: контрольная цифра ИИН не сходится»: в ИИН опечатка, исправьте его в файле. Такой номер помощник на сайт не отправляет.</li>
        <li>«Превышен лимит запросов»: помощник сам подождёт и продолжит медленнее. Не запускайте его на нескольких компьютерах сразу.</li>
        <li>«Не удалось сохранить»: файл открыт в Excel. Закройте его и нажмите «Повторить».</li>
        <li>Регион долго не собирается (gov.kz молчит): нажмите в окне помощника «Пропустить регион», он будет красным; соберите позже.
          Файлы, скачанные вручную, можно положить в папку <code>ДГД/&lt;регион&gt;</code>: помощник прочитает и их.</li>
        <li>В блоке ДГД регион красный: раздел «Документы» на gov.kz не открылся или 0 публикаций за 31 день. Жёлтый: 1–2 публикации,
          последняя старше 14 дней, ИИН в файле не прочитан (откройте файл по ссылке в листе «Публикации»), у публикации нет файла.</li>
        <li>Ссылка в столбце hyperlink не открывается: в поле «Путь к этой папке» стоит не тот путь. Вставьте путь из адресной строки
          Проводника и запустите проверку заново (Старт, затем «Все заново»).</li>
      </ul>
    </details>
    <details>
      <summary>Текст заявки в ИТ (если закладка заблокирована)</summary>
      <p id="itText" style="white-space:pre-line">Прошу в Microsoft Edge (Google Chrome) на моём рабочем ПК разрешить для сайтов tazalau.qoldau.kz и www.gov.kz запуск закладки-помощника
(автоматизация проверки ИИН по открытым реестрам tazalau и публикациям ДГД на gov.kz):
1) если закладки javascript: запрещены политикой URLBlocklist, исключение для этой закладки;
2) доступ сайта к выбранной папке: FileSystemWriteAskForUrls = [*.]qoldau.kz;
3) открытие вкладки gov.kz из помощника: PopupsAllowedForUrls = [*.]qoldau.kz;
4) по возможности: AutomaticDownloadsAllowedForUrls = [*.]qoldau.kz и SleepingTabsBlockedForUrls = [*.]qoldau.kz.</p>
      <button id="copyIt" type="button">Скопировать текст заявки</button><span class="ok" id="okIt"></span>
    </details>
  </div>
  <p class="small">Файл Excel и снимки не покидают ваш компьютер: помощник обращается только к сайтам tazalau.qoldau.kz и www.gov.kz,
    так же как при ручной проверке. Не запускайте вдвоём.</p>
</main>
<textarea id="bmText" readonly aria-hidden="true">__BOOKMARKLET_TEXT__</textarea>
<script>
  function copy(id, okId) {
    var t = document.getElementById(id), ok = document.getElementById(okId);
    var text = t.value != null ? t.value : t.innerText;
    var done = function () { ok.textContent = '✓ скопировано'; setTimeout(function () { ok.textContent = ''; }, 3000); };
    var fallback = function () {
      var a = document.createElement('textarea');
      a.value = text;
      document.body.appendChild(a);
      a.select();
      document.execCommand('copy');
      a.remove();
      done();
    };
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(text).then(done, fallback);
    else fallback();
  }
  document.getElementById('copyBm').addEventListener('click', function () { copy('bmText', 'okBm'); });
  document.getElementById('copyIt').addEventListener('click', function () { copy('itText', 'okIt'); });

  /* Кнопку перетаскивают, а не нажимают: нажатие здесь ничего не установит */
  document.getElementById('bm').addEventListener('click', function (e) {
    e.preventDefault();
    var hint = document.getElementById('bmHint');
    hint.textContent = '← не нажимайте, а перетащите на панель избранного';
    setTimeout(function () { hint.textContent = ''; }, 8000);
  });

  /* Сочетания клавиш для Mac и предупреждение про неподходящий браузер */
  function browserWarning(ua) {
    var ver = +((ua.match(/Chrome\\/(\\d+)/) || [])[1] || 0);
    if (/Electron|Code\\//.test(ua)) return 'Это окно редактора, а не браузер. Откройте файл двойным щелчком в Microsoft Edge или Google Chrome.';
    if (/Firefox\\//.test(ua)) return 'Помощник проверен в Edge и Chrome. Откройте этот файл в Microsoft Edge или Google Chrome.';
    if (!ver) return 'Этот браузер не подходит (например, Safari). Откройте файл в Microsoft Edge или Google Chrome.';
    if (ver < 103) return 'Браузер устарел (версия ' + ver + '). Нужен Edge или Chrome версии 103 или новее, обратитесь в ИТ.';
    return '';
  }
  if (/Mac/.test(navigator.platform || navigator.userAgent)) {
    Array.prototype.forEach.call(document.querySelectorAll('[data-mac]'), function (k) { k.textContent = k.getAttribute('data-mac'); });
  }
  var warn = browserWarning(navigator.userAgent);
  if (warn) {
    var w = document.getElementById('browserWarn');
    w.textContent = warn;
    w.hidden = false;
  }
</script>
</body>
</html>
"""
if "—" in page:
    sys.exit("на странице установки есть длинное тире: его быть не должно")
page = (page.replace("__VERSION__", html.escape(version))
            .replace("__BOOKMARKLET__", html.escape(bookmarklet, quote=True))
            .replace("__BOOKMARKLET_TEXT__", html.escape(bookmarklet, quote=False)))
OUT.write_text(page, encoding="utf-8")
print(f"версия {version}: код {len(source) // 1024} КБ -> сжатый {len(code) // 1024} КБ -> закладка {len(bookmarklet) // 1024} КБ")
print(f"готово: {OUT} ({OUT.stat().st_size // 1024} КБ)")
