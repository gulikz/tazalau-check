# Проверки закладки-помощника. Сайт заменён местным макетом: ни один запрос не уходит в сеть.
# usage: python3 test_helper.py            (нужны playwright, openpyxl, pillow; Google Chrome или Edge на компьютере)
#
# Макет повторяет то, на что помощник опирается на настоящем сайте: <body class="loading"> с задержкой 0,4 с,
# поле фильтра с ИИН, счётчик «Всего записей», таблицу результатов, страницу «Превышен лимит запросов» с кодом 200.
import base64
import html
import io
import json
import re
import sys
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image
from playwright.sync_api import sync_playwright

DEV = Path(__file__).resolve().parent
ROOT = DEV.parent.parent
sys.path.insert(0, str(ROOT))
from tazalau.config import Settings                      # noqa: E402  (настольная программа: перекрёстная проверка)
from tazalau.excel_io import ExcelBook                   # noqa: E402
from tazalau.iin import check_digit                      # noqa: E402

INSTALLER = DEV.parent / "Tazalau_helper.html"
SITE = "https://tazalau.qoldau.kz"
REGISTRIES = {
    "/ru/list/bankruptcy/recovery": "flApplicantXin", "/ru/list/bankruptcy/judicial": "flApplicantXin",
    "/ru/list/bankruptcy-and-insolvent": "flApplicantIin", "/ru/list/debtor": "flApplicantIin",
    "/ru/list/debtor/judicial": "flApplicantXin", "/ru/list/debtor/extrajudicial/terminated-and-cancelled": "flApplicantIin",
}
source = INSTALLER.read_text(encoding="utf-8")
HREF = html.unescape(re.search(r'<a class="bm" id="bm" href="([^"]+)"', source).group(1))
VERSION = re.search(r"const VERSION = '([^']+)'", (DEV.parent / "tazalau_helper.js").read_text(encoding="utf-8")).group(1)

results, problems = {}, []


def expect(name, condition, detail=""):
    results[name] = bool(condition)
    if not condition:
        problems.append(f"{name}: {detail}")


def fake(prefix11):
    """ИИН с верной контрольной цифрой и невозможной датой рождения (13-й месяц): такого человека нет."""
    return prefix11 + str(check_digit(prefix11))


A, B, C, D, E = (fake(p) for p in ("99139990001", "00139990002", "05139990003", "77139990004", "66139990005"))


class Site:
    """Макет сайта: отвечает по таблице, считает обращения."""

    def __init__(self):
        self.found = {}          # ИИН -> записей (или {путь реестра: записей})
        self.limit = {}          # ИИН -> сколько раз ответить «Превышен лимит запросов»
        self.fail = {}           # ИИН -> сколько раз ответить ошибкой 500 (-1 = всегда)
        self.wide = set()        # ИИН с очень широкой таблицей
        self.ignore = set()      # ИИН, для которых сайт «не применил фильтр»
        self.calls = []

    def page(self, path, param, iin):
        count = self.found.get(iin, 0)
        if isinstance(count, dict):
            count = count.get(path, 0)
        shown = "" if iin in self.ignore else iin
        rows = "".join(f"<tr><td>ТЕСТ Т. {i + 1}</td><td>Восстановление платежеспособности</td><td>16.10.2026</td>"
                       f"<td>Тестовый городской суд</td><td>Дело на рассмотрении</td></tr>" for i in range(min(count, 15)))
        if not count:
            rows = '<tr><td colspan="5" class="text-center">Нет записей</td></tr>'
        pager = f'<div class="text-nowrap"><span><small>Всего записей</small> {count}</span><ul class="pagination"><li>1</li></ul></div>'
        wide = "min-width:2600px;" if iin in self.wide else ""
        return f"""<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Cписок банкротов</title>
<style>
 body{{margin:0;font:14px Arial,sans-serif;background:#eef4f8;color:#1d2433}} body.loading{{visibility:hidden}}
 .left-side-menu{{position:fixed;left:0;top:0;bottom:0;width:320px;background:#fff;padding:20px;box-sizing:border-box}}
 .content-page{{margin-left:320px;padding:24px}} .card{{background:#fff;border-radius:12px;padding:20px;margin:12px 0}}
 .table-responsive{{overflow-x:auto}} table{{border-collapse:collapse;width:100%;{wide}}} td,th{{padding:8px;text-align:left}}
 ul.pagination{{display:inline-flex;list-style:none;margin:0 0 0 8px;padding:0}} input{{padding:6px;width:300px}}
</style></head><body class="loading">
<div class="left-side-menu"><b>TAZALAU</b><p>Реестры</p></div>
<div class="content-page"><h4 class="page-title">Список банкротов ({html.escape(path)})</h4>
<form method="get"><div class="card"><label>ИИН</label><br><input type="text" name="{param}" value="{html.escape(shown)}"></div>
{pager}<div class="card"><div class="table-responsive"><table class="table"><thead><tr><th>ФИО</th><th>Категория дела</th>
<th>Дата иска</th><th>Наименование суда</th><th>Статус дела</th></tr></thead><tbody>{rows}</tbody></table></div></div>{pager}</form></div>
<script>setTimeout(function () {{ document.body.classList.remove('loading'); }}, 400);</script></body></html>"""

    def handle(self, route, request):
        url = urllib.parse.urlsplit(request.url)
        if request.url.startswith("file:"):
            return route.continue_()
        if url.hostname == "example.org":
            return route.fulfill(body="<!doctype html><title>other</title><p>other site</p>", headers={"content-type": "text/html"})
        if url.hostname != "tazalau.qoldau.kz":
            return route.abort()                                   # наружу не выходит ничего
        param = REGISTRIES.get(url.path.rstrip("/"))
        if not param:
            return route.fulfill(status=404, body="not found", headers={"content-type": "text/plain"})
        iin = (urllib.parse.parse_qs(url.query).get(param) or [""])[0]
        if iin:
            self.calls.append((url.path, iin))
        if self.limit.get(iin, 0) > 0:
            self.limit[iin] -= 1
            return route.fulfill(headers={"content-type": "text/html; charset=utf-8"}, body=(
                "<!doctype html><html><head><meta charset='utf-8'><title>Превышен лимит запросов!</title></head>"
                "<body>Превышен лимит запросов! Попробуйте вернуться на страницу позже.</body></html>"))
        left = self.fail.get(iin, 0)
        if left:
            if left > 0:
                self.fail[iin] = left - 1
            return route.fulfill(status=500, headers={"content-type": "text/html; charset=utf-8"},
                                 body="<!doctype html><title>500</title><h1>Internal Server Error</h1>")
        return route.fulfill(headers={"content-type": "text/html; charset=utf-8"}, body=self.page(url.path, param, iin))


def make_book(rows, title="восстановление", extra=None):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(["ФИО", "ИИН/БИН", "Ссудный счет"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="C6E0B4")
    for row in rows:
        ws.append(list(row))
    ws["A2"].fill = PatternFill("solid", fgColor="FFFF00")
    if extra:
        extra(wb, ws)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


BASE_ROWS = [("первый", A, "KZ1"), ("второй", int(B), "KZ2"), ("первый ещё раз", f" {A} ", "KZ3"),
             ("опечатка", A[:11] + str((int(A[11]) + 1) % 10), "KZ4"), ("третий", "000000000000", "KZ5")]

FAST = {"INTERVAL": 0, "BACKOFF": [20, 20, 20], "LIMIT_PAUSE": 40, "NAV_TIMEOUT": 6000, "SAVE_EVERY": 2}
DONE = "() => /^(✔ Готово\\.|⏹ Остановлено\\.|❌ )/.test(document.querySelector('#tazalau-state').textContent)"
STATE = "() => document.querySelector('#tazalau-state').textContent"


def open_helper(ctx, site, viewport=None, overrides=None):
    page = ctx.new_page()
    if viewport:
        page.set_viewport_size(viewport)
    page.errors = []
    page.on("pageerror", lambda e: page.errors.append(str(e)))
    page.goto(SITE + "/ru/list/bankruptcy/recovery")
    page.evaluate("cfg => { window.__tazalauTest = cfg; }", dict(FAST, **(overrides or {})))
    ctx.new_cdp_session(page).send("Page.navigate", {"url": HREF})        # как нажатие на закладку
    page.wait_for_function("() => !!window.__tazalau")
    return page


def opfs_put(page, name, data, folder=""):
    page.evaluate("""async ([folder, name, b64]) => {
        let dir = await navigator.storage.getDirectory();
        if (folder) dir = await dir.getDirectoryHandle(folder, { create: true });
        const w = await (await dir.getFileHandle(name, { create: true })).createWritable();
        await w.write(Uint8Array.from(atob(b64), c => c.charCodeAt(0)));
        await w.close();
    }""", [folder, name, base64.b64encode(data).decode()])


def opfs_tree(page, folder=""):
    """{путь: байты} всего, что лежит в папке (у помощника это «папка с Excel»)."""
    listing = page.evaluate("""async folder => {
        let root = await navigator.storage.getDirectory();
        if (folder) root = await root.getDirectoryHandle(folder);
        const out = {};
        const walk = async (dir, prefix) => {
            for await (const [name, h] of dir.entries()) {
                if (h.kind === 'directory') await walk(h, prefix + name + '/');
                else {
                    const bytes = new Uint8Array(await (await h.getFile()).arrayBuffer());
                    let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
                    out[prefix + name] = btoa(s);
                }
            }
        };
        await walk(root, '');
        return out;
    }""", folder)
    return {name: base64.b64decode(data) for name, data in listing.items()}


def use_folder(page, folder):
    page.evaluate("""async folder => {
        const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle(folder);
        await window.__tazalau.test.useDir(dir);
    }""", folder)
    page.wait_for_function("() => window.__tazalau.test.state().sheet")


def run_and_wait(page, timeout=90000):
    page.get_by_text("▶ Старт").click()
    page.wait_for_function(DONE, timeout=timeout)
    return page.evaluate(STATE)


def check_xlsx(name, original, patched):
    """Независимая проверка книги, записанной помощником."""
    before, after = zipfile.ZipFile(io.BytesIO(original)), zipfile.ZipFile(io.BytesIO(patched))
    expect(f"{name}: архив цел", after.testzip() is None)
    bad = []
    for part in after.namelist():
        if part.endswith((".xml", ".rels", ".vml")):
            try:
                ET.fromstring(after.read(part))
            except ET.ParseError as exc:
                bad.append(f"{part}: {exc}")
    expect(f"{name}: каждая часть файла корректный XML", not bad, "; ".join(bad))
    changed = sorted(n for n in after.namelist() if n not in before.namelist() or before.read(n) != after.read(n))
    lost = sorted(set(before.namelist()) - set(after.namelist()))
    expect(f"{name}: ни одна часть файла не потеряна", not lost, lost)
    allowed = re.compile(r"^xl/(worksheets/sheet\d+\.xml|worksheets/_rels/sheet\d+\.xml\.rels|styles\.xml)$")
    expect(f"{name}: изменены только лист, его связи и стили", all(allowed.match(n) for n in changed), changed)
    return load_workbook(io.BytesIO(patched))


with sync_playwright() as p:
    try:
        browser = p.chromium.launch(channel="chrome")
    except Exception:
        browser = p.chromium.launch(channel="msedge")
    site = Site()
    ctx = browser.new_context(locale="ru-RU", accept_downloads=True)
    ctx.route("**/*", site.handle)

    # ---------- 1. страница установки: проверка закладки и чужой сайт ----------
    page = ctx.new_page()
    page.goto(INSTALLER.resolve().as_uri())
    expect("установка: предупреждения о браузере нет", page.locator("#browserWarn").is_hidden())
    bm_text = html.unescape(re.search(r'id="bmText"[^>]*>([^<]*)</textarea>', source).group(1))
    expect("установка: кнопка и текст для копирования совпадают", HREF == bm_text and HREF.startswith("javascript:"))
    expect("установка: на странице нет длинных тире", "—" not in page.inner_text("body"))
    cdp = ctx.new_cdp_session(page)
    cdp.send("Page.navigate", {"url": HREF})
    page.wait_for_function("() => document.getElementById('tazalau-check').className.includes('good')")
    expect("установка: закладка сообщает, что работает", f"✔ Закладка работает (версия {VERSION})" in page.inner_text("#tazalau-check"))
    page.click("#bm")
    expect("установка: щелчок по кнопке подсказывает перетащить", "перетащите" in page.inner_text("#bmHint"))
    dialogs = []
    page.on("dialog", lambda d: (dialogs.append(d.message), d.dismiss()))
    page.goto("https://example.org/")
    ctx.new_cdp_session(page).send("Page.navigate", {"url": HREF})
    page.wait_for_timeout(400)
    expect("чужой сайт: помощник предлагает открыть tazalau.qoldau.kz", dialogs and "только на сайте tazalau.qoldau.kz" in dialogs[0], dialogs)
    expect("чужой сайт: окно помощника не появляется", page.evaluate("() => !window.__tazalau"))
    page.close()

    # ---------- 2. один файл: результат приходит в «Загрузки» ----------
    site.found = {A: 1}
    page = open_helper(ctx, site, {"width": 1366, "height": 768})
    downloads = []
    page.on("download", lambda d: downloads.append(d))
    original = make_book(BASE_ROWS)
    page.get_by_text("или один файл").click()
    page.set_input_files("input[type=file]", files=[{"name": "отчет.xlsx", "mimeType": "application/octet-stream", "buffer": original}])
    page.wait_for_function("() => window.__tazalau.test.state().tasks > 0")
    panel = page.inner_text("body")
    expect("файл: сводка по файлу", "Строк с ИИН 5" in panel and "К проверке: 3 ИИН (строк 4)" in panel, panel[:400])
    expect("файл: неверный ИИН показан с причиной", "Строк с неверным ИИН: 1" in panel)
    final = run_and_wait(page)
    page.wait_for_timeout(800)
    got = {d.suggested_filename: Path(d.path()).read_bytes() for d in downloads}
    pngs = sorted(n for n in got if n.endswith(".png"))
    expect("файл: итог", "✔ Готово" in final and "Проверено 3 из 3 · найдено 1 · не найдено 2" in final and "Ссылок записано: 4" in final, final)
    expect("файл: один запрос на ИИН, дубль не запрашивается", [c[1] for c in site.calls] == [A, B, "000000000000"], site.calls)
    expect("файл: три снимка и копия книги", len(pngs) == 3 and "отчет_tazalau.xlsx" in got, sorted(got))
    expect("файл: в «Загрузках» снимок называется «ФИО снимок.png»", pngs == ["второй снимок.png", "первый снимок.png", "третий снимок.png"], pngs)
    wb = check_xlsx("файл", original, got["отчет_tazalau.xlsx"])
    ws = wb["восстановление"]
    expect("файл: заголовки столбцов", (ws["D1"].value, ws["E1"].value) == ("tazalau_file_location", "tazalau_result"))
    expect("файл: заголовки оформлены как соседние", ws["D1"].font.b and ws["D1"].fill.fgColor.rgb.endswith("C6E0B4"))
    link = ws["D2"].hyperlink
    expect("файл: ссылка ведёт на снимок рядом с книгой", link and urllib.parse.unquote(link.target) == "первый снимок.png"
           and ws["D2"].value == "первый снимок.png", link and link.target)
    expect("файл: ссылка синяя и подчёркнутая", ws["D2"].font.color.rgb.endswith("0563C1") and ws["D2"].font.underline == "single")
    expect("файл: дубль ссылается на тот же снимок", ws["D4"].hyperlink.target == link.target)
    expect("файл: найден / не найден", str(ws["E2"].value).startswith("НАЙДЕН: Восстановление платежеспособности, записей: 1 · проверено")
           and str(ws["E3"].value).startswith("НЕ НАЙДЕН: Восстановление платежеспособности · проверено"), (ws["E2"].value, ws["E3"].value))
    expect("файл: опечатка в ИИН остаётся без ссылки, с причиной", ws["D5"].hyperlink is None
           and ws["E5"].value == "ОШИБКА: контрольная цифра ИИН не сходится (вероятно, опечатка)", ws["E5"].value)
    expect("файл: прежние ячейки не тронуты", ws["B2"].value == A and ws["B3"].value == int(B) and ws["A2"].fill.fgColor.rgb.endswith("FFFF00"))
    expect("файл: в текстах нет длинных тире", "—" not in panel + final + "".join(str(c.value) for row in ws.iter_rows() for c in row))
    shot = Image.open(io.BytesIO(got["первый снимок.png"])).convert("RGB")
    expect("снимок: ширина 1920, сверху полоса-подпись", shot.width == 1920 and shot.height > 1000 and shot.getpixel((3, 3)) == (36, 52, 71), shot.size)
    expect("снимок: не пустой", len(shot.getcolors(maxcolors=1 << 22) or range(999)) > 20)
    strip = shot.height - 1000
    expect("снимок: страница на снимке видимая (не «loading»)", shot.getpixel((100, strip + 500)) == (255, 255, 255), shot.getpixel((100, strip + 500)))
    expect("файл: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 3. папка: книга правится на месте ----------
    site.calls.clear()
    site.found = {A: 2, D: {"/ru/list/debtor/judicial": 1}}
    page = open_helper(ctx, site, {"width": 1920, "height": 1080})
    rows = BASE_ROWS + [("четвёртый", D, "KZ6")]
    original = make_book(rows)
    opfs_put(page, "отчет по банкротству.xlsx", original, "work")
    opfs_put(page, "~$отчет по банкротству.xlsx", b"lock", "work")           # служебный файл Excel: не книга
    use_folder(page, "work")
    expect("папка: служебный файл Excel не предлагается", page.evaluate("() => window.__tazalau.test.state().sheet") == "восстановление")
    final = run_and_wait(page)
    tree = opfs_tree(page, "work")
    shots = sorted(n for n in tree if n.startswith("tazalau_screenshots/") and n.endswith(".png"))
    backups = [n for n in tree if n.startswith("tazalau_backup/")]
    expect("папка: итог", "✔ Готово" in final and "найдено 1 · не найдено 3" in final and "Ссылок записано: 5" in final, final)
    expect("папка: у каждого человека своя папка, в ней «снимок.png»", shots == [
        f"tazalau_screenshots/{name}/снимок.png" for name in ("второй", "первый", "третий", "четвёртый")], shots)
    expect("папка: резервная копия совпадает с исходным файлом байт в байт", len(backups) == 1 and tree[backups[0]] == original, backups)
    expect("папка: журнал сохранён рядом со снимками", any(re.match(r"tazalau_screenshots/журнал_tazalau_.*\.csv$", n) for n in tree), sorted(tree))
    patched = tree["отчет по банкротству.xlsx"]
    wb = check_xlsx("папка", original, patched)
    ws = wb["восстановление"]
    target = ws["D2"].hyperlink.target
    expect("папка: ссылка относительная, в подпапку", target in shots and ws["D2"].value == target.replace("/", "\\"), (target, ws["D2"].value))
    expect("папка: записей 2", "записей: 2" in str(ws["E2"].value), ws["E2"].value)
    temp = DEV / "out"
    temp.mkdir(exist_ok=True)
    (temp / "папка.xlsx").write_bytes(patched)
    desk = ExcelBook.open(temp / "папка.xlsx", Settings())                    # читает настольная программа
    layout = desk.use_sheet()
    tasks = desk.read_tasks()
    expect("папка: настольная программа видит те же столбцы", (layout.link_col, layout.result_col) == (4, 5), layout)
    expect("папка: настольная программа считает строки обработанными", [t.has_link for t in tasks] == [True, True, True, False, True, True],
           [t.has_link for t in tasks])
    desk.close()

    # повторный запуск: новых ИИН нет, помощник спрашивает
    calls_before = len(site.calls)
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /ссылка уже записана/.test(document.body.innerText)")
    page.get_by_text("Отмена").click()
    page.wait_for_timeout(300)
    expect("повтор: без подтверждения сайт не запрашивается", len(site.calls) == calls_before)
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /ссылка уже записана/.test(document.body.innerText)")
    page.get_by_text("Все заново").click()
    page.wait_for_function(DONE)
    tree = opfs_tree(page, "work")
    again = sorted(n for n in tree if n.startswith("tazalau_screenshots/первый/") and n.endswith(".png"))
    expect("повтор «все заново»: «снимок.png» свежий, прежний остался с датой в имени",
           len(again) == 2 and again[0] == "tazalau_screenshots/первый/снимок.png"
           and re.fullmatch(r"tazalau_screenshots/первый/снимок_\d{4}-\d\d-\d\d_\d\d-\d\d-\d\d\.png", again[1])
           and len(site.calls) == calls_before + 4, (again, len(site.calls), calls_before))
    wb2 = check_xlsx("повтор", original, tree["отчет по банкротству.xlsx"])
    rels = zipfile.ZipFile(io.BytesIO(tree["отчет по банкротству.xlsx"])).read("xl/worksheets/_rels/sheet1.xml.rels").decode()
    sheet_xml = zipfile.ZipFile(io.BytesIO(tree["отчет по банкротству.xlsx"])).read("xl/worksheets/sheet1.xml").decode()
    expect("повтор: ссылки заменены, а не удвоены", rels.count("relationships/hyperlink") == 5 and sheet_xml.count("<hyperlink ") == 5,
           (rels.count("relationships/hyperlink"), sheet_xml.count("<hyperlink ")))
    expect("папка: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 3а. папки по столбцу FIO ----------
    site.calls.clear()
    site.found = {}
    page = open_helper(ctx, site, {"width": 1366, "height": 768})

    def fio_book(header, people):
        wb = Workbook()
        ws = wb.active
        ws.title = "ИИН"
        ws.append([header, "ИИН/БИН", "tazalau_file_location"] if header else ["ИИН/БИН", "tazalau_file_location"])
        for name, iin in people:
            ws.append([name, iin] if header else [iin])
        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    people = [("ИВАНОВ ИВАН ИВАНОВИЧ", A), ("Иванов Иван Иванович", int(B)),      # однофамильцы: разные ИИН
              ("Петров Пётр", int(C)), ("Петров Пётр", int(C)),                   # один человек в двух строках
              (None, D),                                                           # имя не указано
              ('ТОО "Ромашка": филиал/отдел №1?', E)]                               # знаки, недопустимые в имени папки
    original = fio_book("FIO", people)
    opfs_put(page, "iin.xlsx", original, "fio")
    use_folder(page, "fio")
    expect("FIO: помощник сообщает, что папки будут по столбцу FIO", "Папки снимков: по столбцу FIO" in page.inner_text("body"))
    final = run_and_wait(page)
    tree = opfs_tree(page, "fio")
    folders = sorted({n.split("/")[1] for n in tree if n.startswith("tazalau_screenshots/") and n.endswith(".png")})
    expect("FIO: папка на человека, однофамильцам добавлен ИИН, без имени: по ИИН, лишние знаки убраны", folders == sorted([
        f"ИВАНОВ ИВАН ИВАНОВИЧ {A}", f"Иванов Иван Иванович {B}", "Петров Пётр", D, "ТОО Ромашка филиал отдел №1"]), folders)
    expect("FIO: в каждой папке один файл «снимок.png»", sorted(n for n in tree if n.endswith(".png")) == sorted(
        f"tazalau_screenshots/{f}/снимок.png" for f in folders), sorted(n for n in tree if n.endswith(".png")))
    ws = check_xlsx("FIO", original, tree["iin.xlsx"])["ИИН"]
    expect("FIO: столбцы на месте, результат добавлен", [ws.cell(1, c).value for c in range(1, 5)] == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result"])
    expect("FIO: ссылка ведёт в папку человека (пробелы в адресе закодированы)",
           ws["C2"].hyperlink.target == f"tazalau_screenshots/ИВАНОВ%20ИВАН%20ИВАНОВИЧ%20{A}/снимок.png"
           and ws["C2"].value == f"tazalau_screenshots\\ИВАНОВ ИВАН ИВАНОВИЧ {A}\\снимок.png", (ws["C2"].hyperlink.target, ws["C2"].value))
    expect("FIO: адрес ссылки после раскодирования совпадает с настоящим файлом",
           all(urllib.parse.unquote(ws.cell(r, 3).hyperlink.target) in tree for r in range(2, 8)),
           [ws.cell(r, 3).hyperlink.target for r in range(2, 8)])
    expect("FIO: один человек в двух строках: один запрос, одна папка, две ссылки",
           [c[1] for c in site.calls].count(C) == 1 and ws["C4"].hyperlink.target == ws["C5"].hyperlink.target == "tazalau_screenshots/Петров%20Пётр/снимок.png")
    for header in ("ФИО", "Ф.И.О.", "fio"):
        opfs_put(page, "iin.xlsx", fio_book(header, [("Сидоров Сидор", A)]), "fio_" + header.replace(".", ""))
        use_folder(page, "fio_" + header.replace(".", ""))
        run_and_wait(page)
        expect(f"FIO: заголовок «{header}» тоже подходит",
               "tazalau_screenshots/Сидоров Сидор/снимок.png" in opfs_tree(page, "fio_" + header.replace(".", "")))
    opfs_put(page, "iin.xlsx", fio_book(None, [(None, A), (None, int(B))]), "nofio")
    use_folder(page, "nofio")
    expect("без FIO: помощник предупреждает, что папки будут по ИИН", "Столбца FIO нет: папки снимков по ИИН" in page.inner_text("body"))
    run_and_wait(page)
    expect("без FIO: папки называются по ИИН", sorted(n for n in opfs_tree(page, "nofio") if n.endswith(".png")) == sorted(
        f"tazalau_screenshots/{i}/снимок.png" for i in (A, B)))
    expect("FIO: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 4. несколько реестров: один снимок на ИИН ----------
    site.calls.clear()
    site.found = {D: {"/ru/list/debtor/judicial": 1, "/ru/list/bankruptcy/recovery": 2}}
    page = open_helper(ctx, site, {"width": 1600, "height": 900})
    original = make_book([("четвёртый", D, "KZ6")])
    opfs_put(page, "два реестра.xlsx", original, "multi")
    use_folder(page, "multi")
    page.locator("details", has_text="Реестры").locator("summary").click()
    page.get_by_label("Судебная процедура", exact=True).check()
    final = run_and_wait(page)
    tree = opfs_tree(page, "multi")
    shots = [n for n in tree if n.endswith(".png")]
    ws = check_xlsx("реестры", original, tree["два реестра.xlsx"])["восстановление"]
    image = Image.open(io.BytesIO(tree[shots[0]]))
    expect("реестры: два запроса, один файл", sorted(c[0] for c in site.calls) == ["/ru/list/bankruptcy/recovery", "/ru/list/debtor/judicial"]
           and len(shots) == 1, (site.calls, shots))
    expect("реестры: снимки склеены друг под другом", image.height > 2000 and image.width == 1920, image.size)
    expect("реестры: итог перечисляет оба", str(ws["E2"].value).startswith(
        "НАЙДЕН: Восстановление платежеспособности, записей: 2; Должники: судебная процедура, записей: 1"), ws["E2"].value)
    page.close()

    # ---------- 5. сайт просит не торопиться, ошибается, отвечает не тем ----------
    site.calls.clear()
    site.found, site.limit, site.fail, site.ignore, site.wide = {}, {B: 2}, {C: -1}, {D}, {E}
    page = open_helper(ctx, site, {"width": 1366, "height": 768})
    original = make_book([("а", A, ""), ("б", int(B), ""), ("в", int(C), ""), ("г", D, ""), ("д", E, "")])
    opfs_put(page, "сбои.xlsx", original, "errors")
    use_folder(page, "errors")
    final = run_and_wait(page)
    tree = opfs_tree(page, "errors")
    ws = check_xlsx("сбои", original, tree["сбои.xlsx"])["восстановление"]
    expect("лимит: после «Превышен лимит запросов» ИИН всё же проверен", ws["D3"].hyperlink is not None
           and [c[1] for c in site.calls].count(B) == 3, [c[1] for c in site.calls])
    expect("ошибка 500: три попытки, ссылки нет, причина записана", [c[1] for c in site.calls].count(C) == 3 and ws["D4"].hyperlink is None
           and str(ws["E4"].value).startswith("ОШИБКА: Восстановление платежеспособности:"), ws["E4"].value)
    expect("фильтр не применён: это ошибка, а не «не найден»", ws["D5"].hyperlink is None and "не применил фильтр" in str(ws["E5"].value), ws["E5"].value)
    wide = Image.open(io.BytesIO(tree["tazalau_screenshots/д/снимок.png"]))
    expect("широкая таблица: снимок расширен, ничего не обрезано", wide.width > 2600 and "шире снимка" not in str(ws["E6"].value), (wide.size, ws["E6"].value))
    normal = Image.open(io.BytesIO(tree["tazalau_screenshots/а/снимок.png"]))
    expect("широкая таблица: остальные снимки обычной ширины", normal.width == 1920, normal.size)
    expect("сбои: итог с числом ошибок", "ошибок 2" in final and "найдено 0 · не найдено 3" in final, final)

    # сайт лежит: помощник останавливается сам
    site.calls.clear()
    site.fail = {A: -1, B: -1, C: -1, D: -1, E: -1}
    site.limit, site.ignore, site.wide = {}, set(), set()
    opfs_put(page, "сайт лежит.xlsx", make_book([("а", A, ""), ("б", int(B), ""), ("в", int(C), ""), ("г", D, ""), ("д", E, "")]), "down")
    use_folder(page, "down")
    final = run_and_wait(page)
    expect("сайт недоступен: остановка после трёх ИИН подряд", "Сайт не отвечает: 3 ИИН подряд" in final and len(site.calls) == 9, (final, len(site.calls)))
    expect("сбои: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 6. «Стоп» и продолжение ----------
    site.calls.clear()
    site.found, site.fail = {}, {}
    page = open_helper(ctx, site, {"width": 1366, "height": 768}, {"INTERVAL": 700})
    opfs_put(page, "стоп.xlsx", make_book([("а", A, ""), ("б", int(B), ""), ("в", int(C), ""), ("г", D, ""), ("д", E, "")]), "stop")
    use_folder(page, "stop")
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => window.__tazalau.test.state().journal.length >= 2")
    page.get_by_text("⏹ Стоп").click()
    page.wait_for_function(DONE)
    final = page.evaluate(STATE)
    done_first = len(site.calls)
    ws = load_workbook(io.BytesIO(opfs_tree(page, "stop")["стоп.xlsx"]))["восстановление"]
    linked = [r for r in range(2, 7) if ws.cell(r, 4).hyperlink]
    expect("стоп: сделанное сохранено, остальное нет", "⏹ Остановлено" in final and 2 <= len(linked) < 5 and len(linked) == done_first, (final, linked))
    use_folder(page, "stop")
    run_and_wait(page)
    ws = load_workbook(io.BytesIO(opfs_tree(page, "stop")["стоп.xlsx"]))["восстановление"]
    expect("стоп: следующий запуск доделывает только оставшиеся", all(ws.cell(r, 4).hyperlink for r in range(2, 7))
           and len(site.calls) == 5, len(site.calls))
    page.close()

    # ---------- 7. книга «как из Excel»: таблица, свои ссылки, объединённый заголовок, формула ----------
    try:
        import xlsxwriter
    except ImportError:
        xlsxwriter = None
    if xlsxwriter:
        site.calls.clear()
        site.found = {A: 1}
        buffer = io.BytesIO()
        book = xlsxwriter.Workbook(buffer, {"in_memory": True})
        sheet = book.add_worksheet("восстановление")
        sheet.merge_range("A1:E1", "Отчёт", book.add_format({"bold": True}))
        heads = ["ФИО", "ИИН/БИН", "Сумма", "путь файла", "Дата"]
        data = [("первый", int(A)), ("второй", B), ("третий", int(C))]
        for i, (name, iin) in enumerate(data, start=3):
            sheet.write(i, 0, name)
            sheet.write(i, 1, iin)
            sheet.write_number(i, 2, 1000.5 * i, book.add_format({"num_format": "#,##0.00"}))
            sheet.write_url(i, 3, r"external:W:\Отдел\{}\Снимок.JPG".format(name), string="Снимок.JPG")
        sheet.add_table(2, 0, 2 + len(data), 4, {"columns": [{"header": h} for h in heads]})
        sheet.write_formula("C8", "=SUM(C4:C6)", None, 0)
        sheet.write_comment("A4", "примечание")
        sheet.freeze_panes(3, 0)
        sheet.set_row(5, None, None, {"hidden": True})
        book.add_worksheet("карты").write_row(0, 0, ["не", "трогать", 42])
        book.close()
        original = buffer.getvalue()
        page = open_helper(ctx, site, {"width": 1366, "height": 768})
        opfs_put(page, "отчет.xlsx", original, "excel")
        use_folder(page, "excel")
        layout = page.evaluate("() => window.__tazalau.test.state().layout")
        expect("книга: заголовки найдены в третьей строке", (layout["headerRow"], layout["iinCol"]) == (3, 2), layout)
        final = run_and_wait(page)
        tree = opfs_tree(page, "excel")
        wb = check_xlsx("книга", original, tree["отчет.xlsx"])
        ws = wb["восстановление"]
        expect("книга: новые столбцы сразу за таблицей", [ws.cell(3, c).value for c in (6, 7)] == ["tazalau_file_location", "tazalau_result"],
               [ws.cell(3, c).value for c in range(1, 9)])
        expect("книга: прежние ссылки на месте", ws["D4"].hyperlink.target.endswith("Снимок.JPG") and ws["D4"].value == "Снимок.JPG")
        expect("книга: новая ссылка рядом, а не вместо", ws["F4"].hyperlink.target == "tazalau_screenshots/первый/снимок.png")
        expect("книга: таблица, формула, примечание, закрепление не тронуты", list(ws.tables.values())[0].ref == "A3:E6"
               and ws["C8"].value == "=SUM(C4:C6)" and ws["A4"].comment is not None and ws["C4"].number_format == "#,##0.00"
               and ws.freeze_panes == "A4" and wb["карты"]["C1"].value == 42)
        expect("книга: скрытая строка осталась скрытой и обработана", ws.row_dimensions[6].hidden and ws["F6"].hyperlink is not None)
        expect("книга: числа с потерянными нулями", ws["F6"].hyperlink.target == "tazalau_screenshots/третий/снимок.png"
               and [c[1] for c in site.calls].count(C) == 1, site.calls)
        (temp / "книга.xlsx").write_bytes(tree["отчет.xlsx"])
        desk = ExcelBook.open(temp / "книга.xlsx", Settings())
        expect("книга: настольная программа читает её так же", desk.use_sheet().link_col == 6 and all(t.has_link for t in desk.read_tasks()))
        desk.close()
        # «только видимые строки»
        opfs_put(page, "отчет.xlsx", original, "visible")
        use_folder(page, "visible")
        page.locator("details", has_text="Реестры").locator("summary").click()
        page.get_by_label("только видимые строки").check()
        site.calls.clear()
        run_and_wait(page)
        expect("книга: скрытые фильтром строки пропускаются по галочке", sorted(c[1] for c in site.calls) == sorted([A, B]), site.calls)
        expect("книга: страница без ошибок", not page.errors, page.errors[:3])
        page.close()

    # ---------- 8. неподходящие файлы ----------
    page = open_helper(ctx, site)
    page.get_by_text("или один файл").click()
    for label, name, data, needle in [
        ("нет столбца ИИН/БИН", "без иин.xlsx", make_book([], extra=lambda wb, ws: ws.cell(1, 2, "Счёт")), "не найден столбец с заголовком «ИИН/БИН»"),
        ("старый .xls или пароль", "старый.xlsx", b"\xd0\xcf\x11\xe0" + b"\0" * 64, "защищён паролем или это старый .xls"),
        ("не Excel", "мусор.xlsx", b"hello, this is not a workbook", "не файл Excel"),
    ]:
        page.set_input_files("input[type=file]", files=[{"name": name, "mimeType": "application/octet-stream", "buffer": data}])
        page.wait_for_function("needle => document.querySelector('#tazalau-state').textContent.includes(needle)", arg=needle, timeout=5000)
        expect(f"плохой файл: {label}", needle in page.evaluate(STATE) and page.get_by_text("▶ Старт").is_disabled())
    page.close()
    browser.close()

print(json.dumps(results, ensure_ascii=False, indent=1))
print(f"\n{sum(results.values())} из {len(results)} проверок пройдено")
if problems:
    print("\nПРОБЛЕМЫ:\n  " + "\n  ".join(problems))
    sys.exit(1)
