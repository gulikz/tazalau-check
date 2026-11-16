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
from PIL import Image, ImageChops
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
# Любой файл шрифта с компьютера: макет сайта отдаёт его как свой значковый шрифт (macOS, Windows, Linux)
FONT_CANDIDATES = ["/System/Library/Fonts/Supplemental/Webdings.ttf", "/System/Library/Fonts/Supplemental/Wingdings.ttf",
                   r"C:\Windows\Fonts\webdings.ttf", r"C:\Windows\Fonts\wingding.ttf",
                   "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf", "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf"]
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
        self.icon_font = None    # файл шрифта: в меню появляется значок, нарисованный «значковым шрифтом сайта»
        self.font_calls = []
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
        icons = '<link rel="stylesheet" href="/icons/style.css">' if self.icon_font else ""
        icon = '<p><i class="ic ic-a"></i> <span>Qoldau.kz</span></p>' if self.icon_font else ""
        return f"""<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Cписок банкротов</title>{icons}
<style>
 body{{margin:0;font:14px Arial,sans-serif;background:#eef4f8;color:#1d2433}} body.loading{{visibility:hidden}}
 .left-side-menu{{position:fixed;left:0;top:0;bottom:0;width:320px;background:#fff;padding:20px;box-sizing:border-box}}
 .content-page{{margin-left:320px;padding:24px}} .card{{background:#fff;border-radius:12px;padding:20px;margin:12px 0}}
 .table-responsive{{overflow-x:auto}} table{{border-collapse:collapse;width:100%;{wide}}} td,th{{padding:8px;text-align:left}}
 ul.pagination{{display:inline-flex;list-style:none;margin:0 0 0 8px;padding:0}} input{{padding:6px;width:300px}}
</style></head><body class="loading">
<div class="left-side-menu"><b>TAZALAU</b><p>Реестры</p>{icon}</div>
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
        if url.path == "/icons/style.css":                         # значковый шрифт описан так же, как на настоящем сайте
            return route.fulfill(headers={"content-type": "text/css"}, body=(
                "@font-face {\n  font-family: 'probeicons';\n  src:  url('fonts/probe.eot?x1');\n"
                "  src:  url('fonts/probe.eot?x1#iefix') format('embedded-opentype'),\n    url('fonts/probe.ttf?x1') format('truetype'),\n"
                "    url('fonts/probe.svg?x1#probe') format('svg');\n  font-weight: normal;\n  font-style: normal;\n  font-display: block;\n}\n"
                ".ic { font-family: 'probeicons' !important; font-style: normal; font-size: 40px; line-height: 1; }\n"
                '.ic-a:before { content: "abc"; }\n'))
        if url.path.startswith("/icons/fonts/"):
            self.font_calls.append(url.path)
            if url.path.endswith(".ttf") and self.icon_font:
                return route.fulfill(headers={"content-type": "font/ttf"}, body=Path(self.icon_font).read_bytes())
            return route.fulfill(status=404, body="not found", headers={"content-type": "text/plain"})
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

FAST = {"INTERVAL": 0, "BACKOFF": [20, 20, 20], "LIMIT_PAUSE": 40, "NAV_TIMEOUT": 6000, "SAVE_EVERY": 2, "SETTLE": 20}
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


def use_folder(page, folder, base=None):
    """Открыть папку в помощнике. Путь к папке на диске (для столбца hyperlink) помощник помнит в самой папке,
    в файле tazalau_backup/путь_к_папке.txt:
    None: на вопрос о пути уже ответили «без hyperlink» (в файле слово «нет»); "": ничего не запомнено, как при
    первом запуске; "keep": что запомнено, то и остаётся; иначе путь вставляется в поле."""
    page.evaluate("""async ([folder, base]) => {
        const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle(folder);
        if (base === null) {
            const service = await dir.getDirectoryHandle('tazalau_backup', { create: true });
            const w = await (await service.getFileHandle('путь_к_папке.txt', { create: true })).createWritable();
            await w.write('нет\\r\\n');
            await w.close();
        } else if (base !== 'keep') {
            try { await (await dir.getDirectoryHandle('tazalau_backup')).removeEntry('путь_к_папке.txt'); } catch (e) { /* и не было */ }
        }
        await window.__tazalau.test.useDir(dir);
    }""", [folder, base])
    page.wait_for_function("() => window.__tazalau.test.state().sheet")
    if base and base != "keep":
        page.fill("#tazalau-base", base)
        page.locator("#tazalau-base").dispatch_event("change")


def with_abs_path(data, folder_url):
    """Книга «как из Excel»: Excel записывает в файл папку, где его сохранили (x15ac:absPath)."""
    source, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    mark = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15">'
            f'<x15ac:absPath url="{html.escape(folder_url)}" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>'
            '</mc:Choice></mc:AlternateContent>')
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            body = source.read(item.filename)
            if item.filename == "xl/workbook.xml":
                text, count = re.subn(r"(<workbookPr\b[^>]*/>)", lambda m: m.group(1) + mark, body.decode("utf-8"), count=1)
                assert count == 1, "в книге нет <workbookPr/>"
                body = text.encode("utf-8")
            target.writestr(item, body)
    return out.getvalue()


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
    expect("файл: поля пути к папке нет (куда придут файлы, неизвестно)", page.locator("#tazalau-base").is_hidden())
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
    shots = sorted(n for n in tree if n.endswith(".png"))
    backups = [n for n in tree if n.startswith("tazalau_backup/") and n.endswith(".xlsx")]
    expect("папка: итог", "✔ Готово" in final and "найдено 1 · не найдено 3" in final and "Ссылок записано: 5" in final, final)
    expect("папка: у каждого человека своя папка рядом с файлом Excel, в ней «снимок.png»", shots == [
        f"{name}/снимок.png" for name in ("второй", "первый", "третий", "четвёртый")], shots)
    expect("папка: резервная копия совпадает с исходным файлом байт в байт", len(backups) == 1 and tree[backups[0]] == original, backups)
    expect("папка: журнал сохранён в служебную папку", any(re.match(r"tazalau_backup/журнал_tazalau_.*\.csv$", n) for n in tree), sorted(tree))
    expect("папка: кроме папок людей и служебной папки, ничего лишнего не появилось",
           {n.split("/")[0] for n in tree} == {"отчет по банкротству.xlsx", "~$отчет по банкротству.xlsx", "tazalau_backup",
                                               "второй", "первый", "третий", "четвёртый"}, sorted(tree))
    patched = tree["отчет по банкротству.xlsx"]
    wb = check_xlsx("папка", original, patched)
    ws = wb["восстановление"]
    target = ws["D2"].hyperlink.target
    expect("папка: ссылка относительная, в папку человека", target in shots and ws["D2"].value == target.replace("/", "\\"), (target, ws["D2"].value))
    expect("папка: без пути к папке столбец hyperlink не создаётся", [ws.cell(1, c).value for c in range(4, 7)] == ["tazalau_file_location", "tazalau_result", None],
           [ws.cell(1, c).value for c in range(1, 7)])
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
    again = sorted(n for n in tree if n.startswith("первый/") and n.endswith(".png"))
    expect("повтор «все заново»: «снимок.png» свежий, прежний остался с датой в имени",
           len(again) == 2 and again[0] == "первый/снимок.png"
           and re.fullmatch(r"первый/снимок_\d{4}-\d\d-\d\d_\d\d-\d\d-\d\d\.png", again[1])
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
    folders = sorted({n.split("/")[0] for n in tree if n.endswith(".png")})
    expect("FIO: папка на человека, однофамильцам добавлен ИИН, без имени: по ИИН, лишние знаки убраны", folders == sorted([
        f"ИВАНОВ ИВАН ИВАНОВИЧ {A}", f"Иванов Иван Иванович {B}", "Петров Пётр", D, "ТОО Ромашка филиал отдел №1"]), folders)
    expect("FIO: в каждой папке один файл «снимок.png»", sorted(n for n in tree if n.endswith(".png")) == sorted(
        f"{f}/снимок.png" for f in folders), sorted(n for n in tree if n.endswith(".png")))
    ws = check_xlsx("FIO", original, tree["iin.xlsx"])["ИИН"]
    expect("FIO: столбцы на месте, результат добавлен", [ws.cell(1, c).value for c in range(1, 5)] == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result"])
    expect("FIO: ссылка ведёт в папку человека (пробелы в адресе закодированы)",
           ws["C2"].hyperlink.target == f"ИВАНОВ%20ИВАН%20ИВАНОВИЧ%20{A}/снимок.png"
           and ws["C2"].value == f"ИВАНОВ ИВАН ИВАНОВИЧ {A}\\снимок.png", (ws["C2"].hyperlink.target, ws["C2"].value))
    expect("FIO: адрес ссылки после раскодирования совпадает с настоящим файлом",
           all(urllib.parse.unquote(ws.cell(r, 3).hyperlink.target) in tree for r in range(2, 8)),
           [ws.cell(r, 3).hyperlink.target for r in range(2, 8)])
    expect("FIO: один человек в двух строках: один запрос, одна папка, две ссылки",
           [c[1] for c in site.calls].count(C) == 1 and ws["C4"].hyperlink.target == ws["C5"].hyperlink.target == "Петров%20Пётр/снимок.png")
    for header in ("ФИО", "Ф.И.О.", "fio"):
        opfs_put(page, "iin.xlsx", fio_book(header, [("Сидоров Сидор", A)]), "fio_" + header.replace(".", ""))
        use_folder(page, "fio_" + header.replace(".", ""))
        run_and_wait(page)
        expect(f"FIO: заголовок «{header}» тоже подходит",
               "Сидоров Сидор/снимок.png" in opfs_tree(page, "fio_" + header.replace(".", "")))
    opfs_put(page, "iin.xlsx", fio_book(None, [(None, A), (None, int(B))]), "nofio")
    use_folder(page, "nofio")
    expect("без FIO: помощник предупреждает, что папки будут по ИИН", "Столбца FIO нет: папки снимков по ИИН" in page.inner_text("body"))
    run_and_wait(page)
    expect("без FIO: папки называются по ИИН", sorted(n for n in opfs_tree(page, "nofio") if n.endswith(".png")) == sorted(
        f"{i}/снимок.png" for i in (A, B)))
    expect("FIO: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 3б. столбец hyperlink: полная ссылка «путь к папке на диске + папка человека + снимок.png» ----------
    site.calls.clear()
    site.found = {A: 1}
    page = open_helper(ctx, site, {"width": 1366, "height": 768})
    for label, given, want in [
        ("путь к самому файлу в кавычках", '"W:\\Отдел\\Папка\\iin.xlsx"', r"W:\Отдел\Папка"),
        ("черта в конце", "W:\\Отдел\\Папка\\", r"W:\Отдел\Папка"),
        ("прямые черты и пробелы по краям", "  W:/Отдел/Папка/ ", r"W:\Отдел\Папка"),
        ("невидимый знак в начале", "\u202aW:\\Отдел\\Папка", r"W:\Отдел\Папка"),
        ("двойная черта внутри", "W:\\Отдел\\\\Папка", r"W:\Отдел\Папка"),
        ("адрес file:///", "file:///W:/Отдел/Папка%20с%20пробелом", r"W:\Отдел\Папка с пробелом"),
        ("сетевой путь", "\\\\fileserver\\share\\Общая папка\\", r"\\fileserver\share\Общая папка"),
        ("сетевой адрес file://", "file://fileserver/share/dir", r"\\fileserver\share\dir"),
    ]:
        got = page.evaluate("v => window.__tazalau.test.normBase(v)", given)
        expect(f"путь к папке: {label}", got == want, got)
    valid = page.evaluate("list => list.map(v => window.__tazalau.test.baseValid(window.__tazalau.test.normBase(v)))",
                          ["W:\\Отдел", "W:", "\\\\srv\\share", "\\\\srv", "Отдел\\Папка", "W:\\От|дел", "/Users/x", "просто текст"])
    expect("путь к папке: годится только путь Windows", valid == [True, True, True, False, False, False, False, False], valid)

    def template(people, heads=("FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink")):
        wb = Workbook()
        ws = wb.active
        ws.title = "ИИН"
        ws.append(list(heads))
        for name, iin in people:
            ws.append([name, iin])
        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    def heads_of(ws, count=6):
        return [ws.cell(1, c).value for c in range(1, count + 1)]

    SAVED = "tazalau_backup/путь_к_папке.txt"            # здесь помощник помнит путь, вставленный вручную

    def set_base(page, text):
        page.fill("#tazalau-base", text)
        page.locator("#tazalau-base").dispatch_event("change")

    def path_of(cell):
        """Путь из формулы =HYPERLINK("путь") в столбце hyperlink."""
        found = re.fullmatch(r'=HYPERLINK\("(.*)"\)', str(cell.value or ""))
        return found.group(1) if found else None

    def shown(data, ref):
        """Что видно в ячейке, пока Excel не пересчитал формулу (сохранённое значение)."""
        return load_workbook(io.BytesIO(data), data_only=True)["ИИН"][ref].value

    # (а) путь к папке записан в файле самим Excel: вводить ничего не надо
    BASE = r"W:\Отдел\Архив 2026\Проверка платежеспособности"
    FOLDER = "Проверка платежеспособности"
    people = [("Иванов Иван Иванович", A), ("Петров Пётр", int(B)), ("Петров Пётр", int(B))]
    original = with_abs_path(fio_book("FIO", people), BASE + "\\")
    opfs_put(page, "iin.xlsx", original, FOLDER)
    use_folder(page, FOLDER, "")
    note = page.inner_text("#tazalau-base-state")
    expect("hyperlink: путь к папке взят из самого файла Excel", page.input_value("#tazalau-base") == BASE
           and "✔" in note and "Путь взят из самого файла Excel" in note, (page.input_value("#tazalau-base"), note))
    final = run_and_wait(page)
    tree = opfs_tree(page, FOLDER)
    ws = check_xlsx("hyperlink", original, tree["iin.xlsx"])["ИИН"]
    expect("hyperlink: столбец добавлен рядом со ссылкой", heads_of(ws) == ["FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "tazalau_result", None], heads_of(ws))
    full = BASE + "\\Иванов Иван Иванович\\снимок.png"
    expect("hyperlink: ссылка записана формулой с полным путём: папка на диске + папка человека + снимок.png",
           ws["D2"].value == '=HYPERLINK("' + full + '")' and ws["D2"].data_type == "f", ws["D2"].value)
    expect("hyperlink: в ячейке виден сам полный путь", shown(tree["iin.xlsx"], "D2") == full, shown(tree["iin.xlsx"], "D2"))
    expect("hyperlink: путь в формуле короче предела Excel для текста в формуле (255 знаков)", len(full) < 255, len(full))
    expect("hyperlink: ссылка синяя и подчёркнутая", ws["D2"].font.color.rgb.endswith("0563C1") and ws["D2"].font.underline == "single")
    sheet_xml = zipfile.ZipFile(io.BytesIO(tree["iin.xlsx"])).read("xl/worksheets/sheet1.xml").decode()
    expect("hyperlink: формула в файле записана как положено", '<c r="D2"' in sheet_xml and re.search(
        r'<c r="D2"[^>]*t="str"[^>]*><f>HYPERLINK\("W:\\Отдел\\Архив 2026\\[^"<]*\\снимок\.png"\)</f><v>W:\\[^<]*</v></c>', sheet_xml) is not None,
        re.findall(r'<c r="D2".*?</c>', sheet_xml))
    expect("hyperlink: обычная ссылка (относительная) остаётся только в столбце tazalau_file_location",
           sheet_xml.count("<hyperlink ") == 3 and all(ws.cell(r, 4).hyperlink is None and ws.cell(r, 3).hyperlink for r in (2, 3, 4)),
           sheet_xml.count("<hyperlink "))
    expect("hyperlink: полный путь ведёт к тому же файлу, что и короткая ссылка, и файл есть",
           all(path_of(ws.cell(r, 4)) == BASE + "\\" + urllib.parse.unquote(ws.cell(r, 3).hyperlink.target).replace("/", "\\")
               and urllib.parse.unquote(ws.cell(r, 3).hyperlink.target) in tree for r in (2, 3, 4)),
           [(ws.cell(r, 3).hyperlink.target, ws.cell(r, 4).value) for r in (2, 3, 4)])
    expect("hyperlink: один человек в двух строках: одна и та же ссылка", ws["D3"].value == ws["D4"].value
           and path_of(ws["D3"]).endswith("\\Петров Пётр\\снимок.png"), ws["D3"].value)
    expect("hyperlink: итог упоминает столбец", "столбец hyperlink" in final and "Ссылок записано: 3" in final, final)
    expect("hyperlink: путь из файла Excel нигде отдельно не запоминается", SAVED not in tree, sorted(tree))
    expect("hyperlink: в текстах нет длинных тире", "—" not in note + final + page.inner_text("body"))

    # (б) в файле уже есть столбец hyperlink, пути в файле нет: помощник спрашивает, путь вставляют один раз
    site.calls.clear()
    original = template([("Сидоров Сидор", A), ("Кузнецова Анна", int(C))])
    opfs_put(page, "iin.xlsx", original, "manual")
    use_folder(page, "manual", "")
    note = page.inner_text("#tazalau-base-state")
    expect("hyperlink: пути нет: помощник сообщает, что столбец не заполняется", page.input_value("#tazalau-base") == ""
           and "сейчас не заполняется" in note, note)
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /Путь к папке на диске не указан/.test(document.body.innerText)")
    page.get_by_role("button", name="Указать путь").click()
    page.wait_for_timeout(300)
    expect("hyperlink: без пути помощник спрашивает и до ответа сайт не запрашивает",
           not site.calls and page.evaluate("() => document.activeElement.id") == "tazalau-base", site.calls)
    set_base(page, '"W:\\Отдел\\manual\\iin.xlsx"')
    note = page.inner_text("#tazalau-base-state")
    expect("hyperlink: вставленный путь к файлу в кавычках приведён к пути папки",
           page.input_value("#tazalau-base") == r"W:\Отдел\manual" and "✔" in note, (page.input_value("#tazalau-base"), note))
    final = run_and_wait(page)
    tree = opfs_tree(page, "manual")
    ws = check_xlsx("hyperlink вручную", original, tree["iin.xlsx"])["ИИН"]
    expect("hyperlink: готовый столбец заполнен, второй такой же не создан",
           heads_of(ws) == ["FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "tazalau_result", None], heads_of(ws))
    expect("hyperlink: путь, вставленный вручную", path_of(ws["D2"]) == r"W:\Отдел\manual\Сидоров Сидор\снимок.png"
           and shown(tree["iin.xlsx"], "D2") == r"W:\Отдел\manual\Сидоров Сидор\снимок.png"
           and ws["C2"].hyperlink.target == "Сидоров%20Сидор/снимок.png" and "Сидоров Сидор/снимок.png" in tree, ws["D2"].value)
    expect("hyperlink: вставленный путь запомнен в самой папке (tazalau_backup), а не в браузере",
           tree.get(SAVED, b"").decode("utf-8").strip() == r"W:\Отдел\manual"
           and page.evaluate("() => Object.keys(localStorage).concat(Object.keys(sessionStorage)).filter(k => /tazalau/i.test(k)).length") == 0,
           tree.get(SAVED))
    use_folder(page, "manual", "keep")
    expect("hyperlink: при следующем открытии папки путь уже стоит в поле", page.input_value("#tazalau-base") == r"W:\Отдел\manual",
           page.input_value("#tazalau-base"))

    # в столбце hyperlink уже стояла своя ссылка, а ячейка была в формате «Текст»
    site.calls.clear()

    def old_link(wb, ws):
        ws["D2"].value = "старая ссылка"
        ws["D2"].hyperlink = "old/1.png"
        ws["D2"].number_format = "@"
        ws["E2"].value = "заметка"
        ws["E2"].hyperlink = "old/2.png"

    wb0 = load_workbook(io.BytesIO(template([("Сидоров Сидор", A)], heads=("FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "Заметки"))))
    old_link(wb0, wb0["ИИН"])
    buffer = io.BytesIO()
    wb0.save(buffer)
    original = buffer.getvalue()
    opfs_put(page, "iin.xlsx", original, "старое")
    use_folder(page, "старое", r"W:\Отдел\старое")
    final = run_and_wait(page)
    patched = opfs_tree(page, "старое")["iin.xlsx"]
    ws = check_xlsx("hyperlink поверх старой ссылки", original, patched)["ИИН"]
    archive = zipfile.ZipFile(io.BytesIO(patched))
    sheet_xml, rels = archive.read("xl/worksheets/sheet1.xml").decode(), archive.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
    expect("hyperlink: прежняя ссылка ячейки убрана вместе с адресом, чужая ссылка рядом цела",
           path_of(ws["D2"]) == r"W:\Отдел\старое\Сидоров Сидор\снимок.png" and ws["D2"].hyperlink is None
           and '<hyperlink ref="D2"' not in sheet_xml and "old/1.png" not in rels and ws["E2"].hyperlink.target == "old/2.png"
           and ws["E2"].value == "заметка", (ws["D2"].value, rels))
    expect("hyperlink: формат «Текст» заменён на «Общий», результат записан за последним столбцом",
           ws["D2"].number_format == "General" and heads_of(ws, 7) == ["FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "Заметки", "tazalau_result", None],
           (ws["D2"].number_format, heads_of(ws, 7)))

    # (в) путь кончается другой папкой: предупреждение и вопрос; сетевой путь
    site.calls.clear()
    original = template([("Сидоров Сидор", A)])
    opfs_put(page, "iin.xlsx", original, "unc")
    use_folder(page, "unc", "\\\\fileserver\\share\\Общая папка\\другая")
    note = page.inner_text("#tazalau-base-state")
    expect("hyperlink: путь кончается не той папкой: предупреждение", "Путь кончается папкой «другая», а выбрана папка «unc»" in note, note)
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /Если это разные папки/.test(document.body.innerText)")
    page.get_by_role("button", name="Исправить путь").click()
    page.wait_for_timeout(300)
    expect("hyperlink: «Исправить путь» ничего не запускает", not site.calls, site.calls)
    set_base(page, "\\\\fileserver\\share\\Общая папка\\unc\\")
    final = run_and_wait(page)
    ws = check_xlsx("hyperlink сеть", original, opfs_tree(page, "unc")["iin.xlsx"])["ИИН"]
    expect("hyperlink: сетевой путь", ws["D2"].value == r'=HYPERLINK("\\fileserver\share\Общая папка\unc\Сидоров Сидор\снимок.png")', ws["D2"].value)
    opfs_put(page, "iin.xlsx", original, "other")
    use_folder(page, "other", r"W:\Отдел\не та")
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /Если это разные папки/.test(document.body.innerText)")
    page.get_by_role("button", name="Записать так").click()
    page.wait_for_function(DONE)
    ws = load_workbook(io.BytesIO(opfs_tree(page, "other")["iin.xlsx"]))["ИИН"]
    expect("hyperlink: «Записать так»: записан путь, который указал человек", path_of(ws["D2"]) == r"W:\Отдел\не та\Сидоров Сидор\снимок.png", ws["D2"].value)

    # (г) «Без hyperlink»: столбец не создаётся, второй раз вопроса нет
    site.calls.clear()
    original = template([("Сидоров Сидор", A), ("Кузнецова Анна", int(C))], heads=("FIO", "ИИН/БИН"))
    opfs_put(page, "iin.xlsx", original, "decline")
    use_folder(page, "decline", "")
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /Путь к папке на диске не указан/.test(document.body.innerText)")
    page.get_by_role("button", name="Без hyperlink").click()
    page.wait_for_function(DONE)
    tree = opfs_tree(page, "decline")
    ws = check_xlsx("без hyperlink", original, tree["iin.xlsx"])["ИИН"]
    expect("hyperlink: «Без hyperlink»: проверка идёт, столбца нет, ответ запомнен в папке",
           heads_of(ws) == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result", None, None]
           and ws["C2"].hyperlink.target == "Сидоров%20Сидор/снимок.png" and tree.get(SAVED, b"").decode("utf-8").strip() == "нет",
           (heads_of(ws), tree.get(SAVED)))
    use_folder(page, "decline", "keep")
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /ссылка уже записана/.test(document.body.innerText)")
    expect("hyperlink: ответ «без hyperlink» запомнен, второй раз о пути не спрашивают", "Путь к папке на диске не указан" not in page.inner_text("body"))
    page.get_by_role("button", name="Отмена").click()
    page.wait_for_timeout(200)

    # (д) длинный путь, как на сетевом диске: длинные имена папок укорачиваются, путь к снимку открывается в Windows
    def long_base(total, tail):
        text = "W:"
        while len(text) + len(tail) < total:
            text += "\\" + "Каталог_с_длинным_именем"[: total - len(tail) - len(text) - 1]
        return text + tail

    LONG = long_base(195, "\\длинный путь")
    room = page.evaluate("v => window.__tazalau.test.folderRoom(v)", LONG)
    cut = lambda name, n: name[:n].rstrip(". ")
    fits, long1, twin = ("Нурмагамбетова Айгерим Бахытжановна", "Абдрахманова-Нурмагамбетова Гульбаршын Сейткамзаевна",
                         "Константинопольская Александра Владиславовна")
    expect("длинный путь: на имя папки остаётся 42 знака", len(LONG) == 195 and room == 42 and len(fits) <= room < min(len(long1), len(twin)), (len(LONG), room))
    site.calls.clear()
    original = template([(fits, A), (long1, int(B)), (twin, int(C)), (twin, D)])
    opfs_put(page, "iin.xlsx", original, "длинный путь")
    use_folder(page, "длинный путь", LONG)
    note = page.inner_text("#tazalau-base-state")
    expect("длинный путь: помощник предупреждает, сколько имён укоротит", "✔" in note and "будет укорочено: 3 (до 42 знаков" in note, note)
    final = run_and_wait(page)
    tree = opfs_tree(page, "длинный путь")
    ws = check_xlsx("длинный путь", original, tree["iin.xlsx"])["ИИН"]
    want = [fits, cut(long1, room), cut(cut(twin, room), room - 13) + " " + C, cut(cut(twin, room), room - 13) + " " + D]
    expect("длинный путь: имена папок укорочены, одинаковым после укорочения добавлен ИИН",
           [path_of(ws.cell(r, 4)) for r in range(2, 6)] == [LONG + "\\" + f + "\\снимок.png" for f in want]
           and all(f + "/снимок.png" in tree for f in want), [str(ws.cell(r, 4).value)[len(LONG):] for r in range(2, 6)])
    longest = max(len(path_of(ws.cell(r, 4)) or "") for r in range(2, 6))
    expect("длинный путь: полный путь к снимку с запасом короче 259 знаков", longest <= 259 - 7 - 3 and all(len(f) <= room for f in want), longest)
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /ссылка уже записана/.test(document.body.innerText)")
    page.get_by_text("Все заново").click()
    page.wait_for_function(DONE)
    tree = opfs_tree(page, "длинный путь")
    pictures = sorted(n for n in tree if n.endswith(".png"))
    expect("длинный путь: при проверке заново прежний снимок получает короткое имя с номером (имя с датой не поместилось бы)",
           pictures == sorted(f + "/" + name for f in want for name in ("снимок.png", "снимок_1.png")), pictures)
    expect("длинный путь: и прежний снимок открывается в Windows", max(len(LONG) + 1 + len(n) for n in pictures) + 7 <= 259,
           max(len(LONG) + 1 + len(n) for n in pictures))

    # (е) слишком длинный путь и не путь вовсе
    site.calls.clear()
    original = template([("Сидоров Сидор", A)], heads=("FIO", "ИИН/БИН"))
    opfs_put(page, "iin.xlsx", original, "слишком")
    use_folder(page, "слишком", long_base(230, "\\слишком"))
    note = page.inner_text("#tazalau-base-state")
    expect("путь к папке: слишком длинный путь назван прямо", "Путь слишком длинный (230 знаков)" in note and "не длиннее 225 знаков" in note, note)
    set_base(page, "просто текст")
    note = page.inner_text("#tazalau-base-state")
    expect("путь к папке: не путь Windows", "Это не путь Windows" in note, note)
    page.get_by_text("▶ Старт").click()
    page.wait_for_function("() => /останется пустым/.test(document.body.innerText)")
    expect("путь к папке: при неверном пути помощник спрашивает перед запуском", not site.calls
           and "Это не путь Windows" in page.inner_text("body"), site.calls)
    page.get_by_role("button", name="Без hyperlink").click()
    page.wait_for_function(DONE)
    tree = opfs_tree(page, "слишком")
    ws = load_workbook(io.BytesIO(tree["iin.xlsx"]))["ИИН"]
    expect("путь к папке: с неверным путём столбец hyperlink не пишется, а неверный путь не запоминается как отказ",
           heads_of(ws) == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result", None, None] and SAVED not in tree, (heads_of(ws), sorted(tree)))

    # (ж) снимок одного человека не сохранился (на месте его папки лежит файл): ошибка только в его строке
    site.calls.clear()
    original = template([("Сидоров Сидор", A), ("Кузнецова Анна", int(C))], heads=("FIO", "ИИН/БИН"))
    opfs_put(page, "iin.xlsx", original, "занято")
    opfs_put(page, "Сидоров Сидор", b"file, not a folder", "занято")
    use_folder(page, "занято", r"W:\Отдел\занято")
    final = run_and_wait(page)
    ws = check_xlsx("сбой записи", original, opfs_tree(page, "занято")["iin.xlsx"])["ИИН"]
    expect("сбой записи: ошибка в строке этого человека, ссылки нет, остальные проверены",
           ws["C2"].hyperlink is None and ws["D2"].value is None
           and str(ws["E2"].value).startswith("ОШИБКА: снимок не сохранён в папку «Сидоров Сидор»")
           and ws["C3"].hyperlink is not None and path_of(ws["D3"]) == r"W:\Отдел\занято\Кузнецова Анна\снимок.png"
           and "✔ Готово" in final and "ошибок 1" in final, (ws["E2"].value, final))
    expect("hyperlink: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 3в. значковый шрифт сайта попадает на снимок (и на самый первый после запуска помощника) ----------
    font_file = next((f for f in map(Path, FONT_CANDIDATES) if f.is_file()), None)
    if font_file is None:
        print("пропущено: на этом компьютере не нашлось файла шрифта для проверки значков")
    else:
        site.calls.clear()
        site.found = {}
        site.icon_font = font_file
        page = open_helper(ctx, site, {"width": 1366, "height": 768})          # новый запуск помощника: шрифты ещё не прогреты
        opfs_put(page, "iin.xlsx", make_book([("значок", A, "")]), "icons")
        use_folder(page, "icons")
        final = run_and_wait(page)
        mine = Image.open(io.BytesIO(opfs_tree(page, "icons")["значок/снимок.png"])).convert("RGB")
        real = ctx.new_page()
        real.set_viewport_size({"width": 1920, "height": 1000})
        real.goto(f"{SITE}/ru/list/bankruptcy/recovery?flApplicantXin={A}")
        real.wait_for_function("() => !document.body.classList.contains('loading') && document.fonts.status === 'loaded'")
        real.wait_for_timeout(300)
        reference = Image.open(io.BytesIO(real.screenshot(full_page=True))).convert("RGB")
        plain = real.evaluate("""() => { const i = document.querySelector('.ic'); i.style.setProperty('font-family', 'Arial', 'important');
            const r = i.getBoundingClientRect(); return [r.x, r.y, r.width, r.height]; }""")
        fallback = Image.open(io.BytesIO(real.screenshot(full_page=True))).convert("RGB")     # так значок выглядит БЕЗ шрифта сайта
        real.close()
        strip = mine.height - reference.height
        box = (0, 0, 320, 160)

        def strong(a, b):
            """Сколько точек заметно отличаются, с допуском сдвига на одну точку."""
            diff = None
            for dx, dy in ((0, 0), (0, 1), (0, -1), (1, 0), (-1, 0)):
                step = ImageChops.difference(a, ImageChops.offset(b, dx, dy)).convert("L")
                diff = step if diff is None else ImageChops.darker(diff, step)
            return sum(diff.crop((1, 1, a.width - 1, a.height - 1)).histogram()[60:])

        got = mine.crop((0, strip, mine.width, mine.height)).crop(box)
        expect("значки: проверка осмысленна (со шрифтом сайта и без него значок выглядит по-разному)",
               strong(reference.crop(box), fallback.crop(box)) > 150, strong(reference.crop(box), fallback.crop(box)))
        expect("значки: на первом же снимке значок нарисован шрифтом сайта, как на настоящем снимке экрана",
               "✔ Готово" in final and mine.size == (1920, reference.height + strip) and strong(got, reference.crop(box)) <= 10,
               (final[:60], mine.size, strong(got, reference.crop(box)), strong(got, fallback.crop(box))))
        expect("значки: файлы шрифта в форматах, которые браузер не использует (.eot, .svg), помощник не запрашивает",
               not [u for u in site.font_calls if not u.endswith(".ttf")], site.font_calls)
        expect("значки: страница без ошибок", not page.errors, page.errors[:3])
        site.icon_font = None
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
    wide = Image.open(io.BytesIO(tree["д/снимок.png"]))
    expect("широкая таблица: снимок расширен, ничего не обрезано", wide.width > 2600 and "шире снимка" not in str(ws["E6"].value), (wide.size, ws["E6"].value))
    normal = Image.open(io.BytesIO(tree["а/снимок.png"]))
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
        expect("книга: новая ссылка рядом, а не вместо", ws["F4"].hyperlink.target == "первый/снимок.png")
        expect("книга: таблица, формула, примечание, закрепление не тронуты", list(ws.tables.values())[0].ref == "A3:E6"
               and ws["C8"].value == "=SUM(C4:C6)" and ws["A4"].comment is not None and ws["C4"].number_format == "#,##0.00"
               and ws.freeze_panes == "A4" and wb["карты"]["C1"].value == 42)
        expect("книга: скрытая строка осталась скрытой и обработана", ws.row_dimensions[6].hidden and ws["F6"].hyperlink is not None)
        expect("книга: числа с потерянными нулями", ws["F6"].hyperlink.target == "третий/снимок.png"
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
