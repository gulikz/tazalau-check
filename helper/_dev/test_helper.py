# Проверки закладки-помощника. Сайт заменён местным макетом: ни один запрос не уходит в сеть.
# usage: python3 test_helper.py            (нужны playwright, openpyxl, pillow; Google Chrome или Edge на компьютере)
#
# Макет повторяет то, на что помощник опирается на настоящем сайте: <body class="loading"> с задержкой 0,4 с,
# поле фильтра с ИИН, счётчик «Всего записей», таблицу результатов, страницу «Превышен лимит запросов» с кодом 200.
import base64
import datetime
import html
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import docx                                                # python-docx: файл Word для проверки списков ДГД
from PIL import Image, ImageChops
from playwright.sync_api import sync_playwright
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas

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


def plausible(prefix7):
    """Синтетический ИИН с правдоподобной датой рождения (для списков ДГД: там отбрасываются номера с невозможной датой)."""
    for tail in range(1000, 10000):
        body = prefix7 + str(tail)
        if check_digit(body) is not None:
            return body + str(check_digit(body))


REGISTRY_ORDER = ("/ru/list/bankruptcy/recovery", "/ru/list/bankruptcy/judicial", "/ru/list/bankruptcy-and-insolvent",
                  "/ru/list/debtor", "/ru/list/debtor/judicial", "/ru/list/debtor/extrajudicial/terminated-and-cancelled")


# Заголовки таблиц результатов: как на настоящем сайте (сверено 20.11.2026), в том числе «Cумма» с латинской C
JUDICIAL = ["ФИО должника", "Категория дела", "Дата Иска", "Дата передачи судье", "Дата принятия к производству", "Дата решения суда",
            "Дата вступления в силу решения суда", "Дата завершения дела", "Наименование суда", "Регион", "Статус дела Torelik"]
HEADS = {
    "/ru/list/bankruptcy/recovery": JUDICIAL,
    "/ru/list/bankruptcy/judicial": JUDICIAL + ["Статус дела Tazalau"],
    "/ru/list/debtor/judicial": JUDICIAL,
    "/ru/list/bankruptcy-and-insolvent": ["ФИО заявителя", "Услугодатель", "Входящий номер заявления",
        "Дата подачи заявления от фронт системы на применение процедуры", "Дата начала процедуры внесудебного банкротства",
        "Дата завершения процедуры внесудебного банкротства", "Статус заявителя", "Список кредиторов, указанный заявителем",
        "Cумма задолженности, указанная заявителем, тг.", "Список кредиторов по данным кредитного бюро", "Cумма задолженности по кредитному бюро, тг."],
    "/ru/list/debtor": ["ФИО заявителя", "Услугодатель", "Входящий номер заявления", "Дата подачи заявления от фронт системы на применение процедуры",
        "Дата начала процедуры внесудебного банкротства", "Статус заявителя", "Список кредиторов, данные ПКБ/ГКБ", "Cумма задолженности, данные ПКБ/ГКБ, тг."],
    "/ru/list/debtor/extrajudicial/terminated-and-cancelled": ["ФИО заявителя", "Услугодатель", "Входящий номер заявления",
        "Дата подачи заявления от фронт системы на применение процедуры", "Дата начала процедуры внесудебного банкротства", "Инициатор процедуры",
        "Дата прекращения процедуры внесудебного банкротства", "Входящий номер заявления на прекращение",
        "Дата подачи заявления от фронт системы на прекращение процедуры", "Список кредиторов", "Cумма задолженности, указанная заявителем, тг."],
}


def default_value(head, path):
    """Правдоподобное значение столбца для записи макета."""
    rules = [(r"^ФИО", "ТЕСТ Т. Т."), (r"Категория", "Восстановление платежеспособности"), (r"Дата Иска|Дата передачи", "16.10.2026"),
             (r"принятия к производству", "20.10.2026"), (r"Дата (решения|вступления|завершения дела)", ""),
             (r"Наименование суда", "Тестовый городской суд"), (r"^Регион", "Алматы"), (r"Статус дела Torelik", "Дело на рассмотрении"),
             (r"Статус дела Tazalau", "Судебное банкротство (тест)"), (r"Услугодатель", "Тестовый ЦОН"), (r"номер заявления", "202600000000001"),
             (r"Дата подачи", "01.05.2026 10:00:00"), (r"Дата начала", "18.05.2026 15:00:00"), (r"Дата завершения процедуры", "19.11.2026 17:00:00"),
             (r"Дата прекращения", "01.07.2026 12:00:00"), (r"Статус заявителя", "Банкрот" if "insolvent" in path else "Должник"),
             (r"Инициатор", "Заявитель"), (r"Список кредиторов", '1) АО "Тестовый банк" 2) ТОО "Тестовое МФО"'), (r"умма", "1 234 567")]
    for pattern, value in rules:
        if re.search(pattern, head):
            return value
    return ""


class Site:
    """Макет сайта: отвечает по таблице, считает обращения."""

    def __init__(self):
        self.found = {}          # ИИН -> записей (или {путь реестра: записей})
        self.limit = {}          # ИИН -> сколько раз ответить «Превышен лимит запросов»
        self.fail = {}           # ИИН -> сколько раз ответить ошибкой 500 (-1 = всегда)
        self.wide = set()        # ИИН с очень широкой таблицей
        self.ignore = set()      # ИИН, для которых сайт «не применил фильтр»
        self.records = {}        # (ИИН, путь реестра) -> [{заголовок: значение}]: свои значения записей вместо обычных
        self.icon_font = None    # файл шрифта: в меню появляется значок, нарисованный «значковым шрифтом сайта»
        self.font_calls = []
        self.calls = []

    def page(self, path, param, iin):
        count = self.found.get(iin, 0)
        if isinstance(count, dict):
            count = count.get(path, 0)
        shown = "" if iin in self.ignore else iin
        heads = HEADS.get(path, JUDICIAL)
        custom = self.records.get((iin, path), [])
        rows = ""
        for i in range(min(count, 15)):
            values = dict({h: default_value(h, path) for h in heads}, **(custom[i] if i < len(custom) else {}))
            rows += "<tr>" + "".join(f"<td>{html.escape(values[h])}</td>" for h in heads) + "</tr>"
        if not count:
            rows = f'<tr><td colspan="{len(heads)}" class="text-center">Нет записей</td></tr>'
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
{pager}<div class="card"><div class="table-responsive"><table class="table"><thead><tr>{"".join(f"<th>{html.escape(h)}</th>" for h in heads)}</tr></thead>
<tbody>{rows}</tbody></table></div></div>{pager}</form></div>
<script>setTimeout(function () {{ document.body.classList.remove('loading'); }}, 400);</script></body></html>"""

    def handle(self, route, request):
        url = urllib.parse.urlsplit(request.url)
        if request.url.startswith("file:"):
            return route.continue_()
        if url.hostname == "example.org":
            return route.fulfill(body="<!doctype html><title>other</title><p>other site</p>", headers={"content-type": "text/html"})
        if url.hostname and url.hostname.endswith(".kgd.gov.kz") and getattr(self, "kgd", None):
            return self.kgd.handle(route, request)
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


TNR = next((f for f in ["/System/Library/Fonts/Supplemental/Times New Roman.ttf", r"C:\Windows\Fonts\times.ttf",
                         "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"] if Path(f).exists()), None)


def order_scan(iin, fio, kind="pdf", dpi=200, blank=False):
    """Скан приказа ДГД с ИИН (картинка страницы A4, как с настоящего сканера): PDF с JPEG внутри или сам JPEG."""
    from PIL import ImageDraw, ImageFont, ImageFilter
    w, h = int(8.27 * dpi), int(11.69 * dpi)
    img = Image.new("L", (w, h), 246)
    if not blank:
        d = ImageDraw.Draw(img)
        k = dpi / 72
        big, body = ImageFont.truetype(TNR, int(14 * k)), ImageFont.truetype(TNR, int(14 * k))
        d.text((int(0.9 * dpi), int(1.0 * dpi)), "ПРИКАЗ № 123", font=big, fill=20)
        d.text((int(0.9 * dpi), int(2.0 * dpi)), "О назначении финансового управляющего", font=body, fill=25)
        d.text((int(1.2 * dpi), int(3.0 * dpi)), f"1. Назначить финансовым управляющим Тестова Т.Т. в отношении", font=body, fill=25)
        d.text((int(0.9 * dpi), int(3.35 * dpi)), f"гражданина {fio}, ИИН {iin}, проживающего в г. Тестовск.", font=body, fill=25)
        d.text((int(1.2 * dpi), int(4.2 * dpi)), "2. Контроль за исполнением приказа оставляю за собой.", font=body, fill=25)
        img = img.filter(ImageFilter.GaussianBlur(0.6))
    buf = io.BytesIO()
    if kind == "jpg":
        img.save(buf, "JPEG", quality=80, dpi=(dpi, dpi))
    else:
        img.convert("RGB").save(buf, "PDF", resolution=dpi)
    return buf.getvalue()


class KgdSite:
    """Макет сайтов ДГД 20 регионов (один Drupal): меню «Физическим лицам», страницы приказов в разных видах,
    как на настоящих сайтах (сверено с копиями страниц 2024–2025 гг.), файлы-сканы. Имена и ИИН выдуманы."""
    HOSTS = ["abai", "akm", "akb", "alm", "atr", "vko", "zhmb", "zhts", "zko", "krg", "kst", "kzl", "mng", "pvl", "sko", "trk", "ult",
             "astana", "almaty", "shymkent"]

    def __init__(self, today):
        self.today = today
        self.down = set()                  # сайты, которые не отвечают
        self.no_page = set()               # в меню нет страницы приказов
        self.calls = []                    # (метод, хост, путь)
        self.files = {}                    # (хост, путь) -> (байты, тип, дата загрузки)
        self.rows = {}                     # хост -> [строки таблицы: (ячейки html, [пути файлов])]
        self.slug = {}
        self.iins = {}                     # хост -> [(ИИН, ФИО)] по порядку приказов

    def day(self, n):
        return (self.today - datetime.timedelta(days=n)).strftime("%d.%m.%Y")

    def add_file(self, host, path, data, ctype, ago):
        self.files[(host, path)] = (data, ctype, self.today - datetime.timedelta(days=ago))

    def menu(self, host):
        slug = self.slug.get(host, f"prikazy-o-naznacheniiotstranenii-osvobozhdenii-finansovogo-upravlyayushchego-{host}")
        y = self.today.year
        items = [f'<li><a href="/ru/depsection/fizicheskim-licam">Физическим лицам</a><ul>',
                 f'<li><a href="/ru/content/{y - 1}-god-{host}">{y - 1} год. Восстановление платежеспособности и банкротство граждан</a><ul>',
                 (f'<li><a href="/ru/content/{slug}-old">Приказы о назначении/отстранении (освобождении) финансового управляющего</a></li>'
                  if host not in self.no_page else '') + '</ul></li>',
                 f'<li><a href="/ru/content/{y}-god-{host}">{y} год. Восстановление платежеспособности и банкротство граждан</a><ul>']
        if host not in self.no_page:
            title = ("Приказы о назначении и отстранении реабилитационных, банкротных управляющих" if host == "kst"
                     else "Приказы о назначении/отстранении (освобождении) финансового управляющего")
            items.append(f'<li><a href="/ru/content/{slug}">{title}</a></li>')
        items.append('<li><a href="/ru/content/reestry-trebovaniy">Реестры требований кредиторов</a></li></ul></li></ul></li>'
                     '<li><a href="/ru/depsection/reabilitaciya">Реабилитация и банкротство</a><ul><li><a href="/ru/content/prikazy-reabil">'
                     'Приказы о назначении/отстранении (освобождении) реабилитационного и банкротного управляющих</a></li></ul></li>')
        return '<ul class="menu sm">' + "".join(items) + "</ul>"

    def page(self, host, body):
        return (f'<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>ДГД {host}</title></head>'
                f'<body class="html not-front domain-{host}-kgd-gov-kz"><div id="sidebar">{self.menu(host)}</div>'
                f'<div id="content"><div class="node"><div class="field field-name-body"><div class="field-items"><div class="field-item">'
                f'{body}</div></div></div></div></div></body></html>')

    def table(self, host):
        trs = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>" for cells, _ in self.rows.get(host, []))
        return f"<table><tbody><tr><td>Наименование</td><td>ФИО</td><td>Дата</td></tr>{trs}</tbody></table>"

    def handle(self, route, request):
        url = urllib.parse.urlsplit(request.url)
        host = url.hostname.split(".")[0]
        self.calls.append((request.method, host, url.path))
        if host in self.down:
            return route.abort("timedout")
        path = urllib.parse.unquote(url.path)
        if (host, path) in self.files:
            data, ctype, when = self.files[(host, path)]
            stamp = when.strftime("%a, %d %b %Y 10:00:00 GMT")
            return route.fulfill(status=200, headers={"content-type": ctype, "last-modified": stamp}, body=b"" if request.method == "HEAD" else data)
        slug = self.slug.get(host, f"prikazy-o-naznacheniiotstranenii-osvobozhdenii-finansovogo-upravlyayushchego-{host}")
        if path == f"/ru/content/{slug}":
            return route.fulfill(headers={"content-type": "text/html; charset=utf-8"}, body=self.page(host, self.table(host)))
        if path.startswith("/ru/"):
            return route.fulfill(headers={"content-type": "text/html; charset=utf-8"}, body=self.page(host, "<p>Раздел</p>"))
        return route.fulfill(status=404, headers={"content-type": "text/html"}, body="<h1>404</h1>")


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


def open_helper(ctx, site, viewport=None, overrides=None, regs=("recovery",)):
    """Помощник на макете сайта. regs: какие реестры отмечены (None: как у пользователя, все шесть)."""
    page = ctx.new_page()
    if viewport:
        page.set_viewport_size(viewport)
    page.errors = []
    page.on("pageerror", lambda e: page.errors.append(str(e)))
    page.goto(SITE + "/ru/list/bankruptcy/recovery")
    page.evaluate("cfg => { window.__tazalauTest = cfg; }", dict(FAST, **(overrides or {})))
    ctx.new_cdp_session(page).send("Page.navigate", {"url": HREF})        # как нажатие на закладку
    page.wait_for_function("() => !!window.__tazalau")
    if regs is not None:
        page.evaluate("keys => window.__tazalau.test.setRegs(keys)", list(regs))
    return page


def opfs_put(page, name, data, folder=""):
    page.evaluate("""async ([folder, name, b64]) => {
        let dir = await navigator.storage.getDirectory();
        for (const part of folder ? folder.split('/') : []) dir = await dir.getDirectoryHandle(part, { create: true });
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
    page.get_by_role("button", name="▶ Старт").click()
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
    expect("папка: без пути к папке столбец hyperlink не создаётся, статус и подробности добавлены",
           [ws.cell(1, c).value for c in range(4, 9)] == ["tazalau_file_location", "tazalau_result", "tazalau_status", "tazalau_details", None],
           [ws.cell(1, c).value for c in range(1, 9)])
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
    page.get_by_role("button", name="▶ Старт").click()
    page.wait_for_function("() => /ссылка уже записана/.test(document.body.innerText)")
    page.get_by_text("Отмена").click()
    page.wait_for_timeout(300)
    expect("повтор: без подтверждения сайт не запрашивается", len(site.calls) == calls_before)
    page.get_by_role("button", name="▶ Старт").click()
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
    expect("hyperlink: столбец добавлен рядом со ссылкой", heads_of(ws, 8) == ["FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "tazalau_result",
           "tazalau_status", "tazalau_details", None], heads_of(ws, 8))
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
    page.get_by_role("button", name="▶ Старт").click()
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
           heads_of(ws, 8) == ["FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "tazalau_result", "tazalau_status", "tazalau_details", None],
           heads_of(ws, 8))
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
           ws["D2"].number_format == "General" and heads_of(ws, 9) == ["FIO", "ИИН/БИН", "tazalau_file_location", "hyperlink", "Заметки", "tazalau_result",
           "tazalau_status", "tazalau_details", None], (ws["D2"].number_format, heads_of(ws, 9)))

    # (в) путь кончается другой папкой: предупреждение и вопрос; сетевой путь
    site.calls.clear()
    original = template([("Сидоров Сидор", A)])
    opfs_put(page, "iin.xlsx", original, "unc")
    use_folder(page, "unc", "\\\\fileserver\\share\\Общая папка\\другая")
    note = page.inner_text("#tazalau-base-state")
    expect("hyperlink: путь кончается не той папкой: предупреждение", "Путь кончается папкой «другая», а выбрана папка «unc»" in note, note)
    page.get_by_role("button", name="▶ Старт").click()
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
    page.get_by_role("button", name="▶ Старт").click()
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
    page.get_by_role("button", name="▶ Старт").click()
    page.wait_for_function("() => /Путь к папке на диске не указан/.test(document.body.innerText)")
    page.get_by_role("button", name="Без hyperlink").click()
    page.wait_for_function(DONE)
    tree = opfs_tree(page, "decline")
    ws = check_xlsx("без hyperlink", original, tree["iin.xlsx"])["ИИН"]
    expect("hyperlink: «Без hyperlink»: проверка идёт, столбца нет, ответ запомнен в папке",
           heads_of(ws) == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result", "tazalau_status", "tazalau_details"]
           and ws["C2"].hyperlink.target == "Сидоров%20Сидор/снимок.png" and tree.get(SAVED, b"").decode("utf-8").strip() == "нет",
           (heads_of(ws), tree.get(SAVED)))
    use_folder(page, "decline", "keep")
    page.get_by_role("button", name="▶ Старт").click()
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
    page.get_by_role("button", name="▶ Старт").click()
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
    page.get_by_role("button", name="▶ Старт").click()
    page.wait_for_function("() => /останется пустым/.test(document.body.innerText)")
    expect("путь к папке: при неверном пути помощник спрашивает перед запуском", not site.calls
           and "Это не путь Windows" in page.inner_text("body"), site.calls)
    page.get_by_role("button", name="Без hyperlink").click()
    page.wait_for_function(DONE)
    tree = opfs_tree(page, "слишком")
    ws = load_workbook(io.BytesIO(tree["iin.xlsx"]))["ИИН"]
    expect("путь к папке: с неверным путём столбец hyperlink не пишется, а неверный путь не запоминается как отказ",
           heads_of(ws) == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result", "tazalau_status", "tazalau_details"] and SAVED not in tree,
           (heads_of(ws), sorted(tree)))

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

    # ---------- 3г. статусы: все шесть реестров, главный статус и подробности ----------
    P_REC, P_JUD, P_EXT, P_DEB, P_DJ, P_TERM = REGISTRY_ORDER
    site.calls.clear()
    site.found = {A: {P_REC: 1, P_EXT: 1}, B: {P_TERM: 1}}
    site.records = {(A, P_REC): [{"Дата Иска": "10.03.2026", "Статус дела Torelik": "Дело на рассмотрении"}],
                    (A, P_EXT): [{"Статус заявителя": "Банкрот"}],
                    (B, P_TERM): [{"ФИО заявителя": "ДРУГОВ Д. Д."}]}
    page = open_helper(ctx, site, {"width": 1366, "height": 768}, regs=None)
    page.locator("details", has_text="Реестры").locator("summary").click()
    expect("статусы: по умолчанию отмечены все шесть реестров", "Реестры: 6 выбрано" in page.inner_text("body")
           and all(page.get_by_label(t, exact=True).is_checked() for t in ("Восстановление платежеспособности", "Внесудебная процедура", "Процедура прекращена (объявления)")))
    original = fio_book("FIO", [("Тест Тимур", A), ("Иванов Иван", int(B)), ("Петров Пётр", int(C))])
    opfs_put(page, "iin.xlsx", original, "status")
    use_folder(page, "status")
    final = run_and_wait(page, timeout=240000)
    tree = opfs_tree(page, "status")
    ws = check_xlsx("статусы", original, tree["iin.xlsx"])["ИИН"]
    heads = [ws.cell(1, c).value for c in range(1, 8)]
    expect("статусы: столбцы итог, статус, подробности", heads == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result", "tazalau_status", "tazalau_details", None], heads)
    expect("статусы: каждый ИИН проверен во всех шести реестрах по порядку", [path for path, i in site.calls if i == A] == list(REGISTRY_ORDER)
           and len(site.calls) == 18, ([path for path, i in site.calls if i == A], len(site.calls)))
    expect("статусы: итог перечисляет реестры с записями", str(ws["D2"].value).startswith(
        "НАЙДЕН: Восстановление платежеспособности, записей: 1; Внесудебное банкротство, записей: 1 · проверено"), ws["D2"].value)
    expect("статусы: главный статус: самая свежая запись из всех реестров",
           ws["E2"].value == "Внесудебное банкротство: Банкрот · процедура завершена 19.11.2026 (всего записей: 2)", ws["E2"].value)
    details = str(ws["F2"].value)
    expect("статусы: подробности по каждой записи: статус, даты, суд или услугодатель, кредиторы и суммы",
           "Восстановление платежеспособности: Дело на рассмотрении · Восстановление платежеспособности · иск 10.03.2026, принято к производству 20.10.2026 · Алматы, Тестовый городской суд" in details
           and 'Внесудебное банкротство: Банкрот · заявление подано 01.05.2026, начало процедуры 18.05.2026, процедура завершена 19.11.2026 · Тестовый ЦОН · кредиторы: 1) АО "Тестовый банк" 2) ТОО "Тестовое МФО"; сумма, тг: 1 234 567 / 1 234 567' in details,
           details)
    expect("статусы: прекращённая процедура и другая фамилия на сайте отмечены",
           str(ws["E3"].value).startswith("Должники: процедура прекращена: процедура прекращена (инициатор: Заявитель) · процедура прекращена 01.07.2026")
           and "внимание: на сайте фамилия «ДРУГОВ», в файле «Иванов»" in str(ws["F3"].value), (ws["E3"].value, ws["F3"].value))
    expect("статусы: не найден нигде", str(ws["D4"].value).startswith("НЕ НАЙДЕН ни в одном из 6 реестров")
           and ws["E4"].value == "нет записей ни в одном из проверенных реестров" and not ws["F4"].value, (ws["D4"].value, ws["E4"].value, ws["F4"].value))
    tall = Image.open(io.BytesIO(tree["Тест Тимур/снимок.png"]))
    expect("статусы: на одном снимке все шесть страниц друг под другом", tall.width == 1920 and tall.height > 6 * 1000, tall.size)
    expect("статусы: в журнале статус найденного", any(r["iin"] == A and "Банкрот" in r["note"] for r in page.evaluate("() => window.__tazalau.test.state().journal")))
    rule = page.evaluate("""() => {
        const t = window.__tazalau.test, reg = { key: 'recovery', short: 'Восстановление платежеспособности' }, at = new Date();
        const rec = (status, filed, done) => ({ 'Дата Иска': filed, 'Дата решения суда': done || '', 'Статус дела Torelik': status });
        const shot = recs => ({ reg, count: recs.length, rows: recs.length, records: recs, at });
        return [
          t.describe([shot([rec('Исковое заявление возвращено', '05.10.2026', '23.10.2026'), rec('Дело на рассмотрении', '21.10.2026')])], '').status,
          t.describe([shot([rec('Дело на рассмотрении', '01.03.2025'), rec('Отказано в принятии искового заявления', '01.11.2026', '05.11.2026')])], '').status,
          t.describe([shot([rec('Дело на рассмотрении', '01.03.2025'), rec('Дело на рассмотрении', '01.11.2026')])], '').status,
        ];
    }""")
    expect("статусы: главная запись: идущее дело важнее возвращённого или отклонённого иска, среди равных самое новое", rule == [
        "Восстановление платежеспособности: Дело на рассмотрении · иск 21.10.2026 (всего записей: 2)",
        "Восстановление платежеспособности: Дело на рассмотрении · иск 01.03.2025 (всего записей: 2)",
        "Восстановление платежеспособности: Дело на рассмотрении · иск 01.11.2026 (всего записей: 2)"], rule)
    expect("статусы: страница без ошибок", not page.errors, page.errors[:3])
    site.records = {}
    page.close()

    # ---------- 3д. выгрузка ADW в CSV (pgAdmin и DBeaver сохраняют CSV) ----------
    site.calls.clear()
    site.found = {A: 1}
    page = open_helper(ctx, site, {"width": 1366, "height": 768})
    adw = ("IIN_BIN;FIO_CLIENT;CONTRACT;PRODUCT\r\n" f"{A};Тестов Тест;007-0001;Кредитная карта\r\n"
           f'{B};"Иванова; Анна";007-0002;Кредит\r\n').encode("cp1251")
    opfs_put(page, "ADW 20.11.2026.csv", adw, "csv")
    use_folder(page, "csv")
    tree = opfs_tree(page, "csv")
    expect("CSV: рядом собрана книга .xlsx, работа идёт в ней", "ADW 20.11.2026.xlsx" in tree and page.evaluate("() => window.__tazalau.test.state().tasks") == 2,
           sorted(tree))
    final = run_and_wait(page)
    tree = opfs_tree(page, "csv")
    ws = load_workbook(io.BytesIO(tree["ADW 20.11.2026.xlsx"]))["Лист1"]
    heads = [ws.cell(1, c).value for c in range(1, 9)]
    expect("CSV: столбцы выгрузки на месте (заголовки IIN_BIN и FIO_CLIENT узнаны), результаты добавлены", heads == ["IIN_BIN", "FIO_CLIENT", "CONTRACT", "PRODUCT",
           "tazalau_file_location", "tazalau_result", "tazalau_status", "tazalau_details"], heads)
    expect("CSV: кодировка 1251, кавычки и разделитель «;» прочитаны; ИИН и номера договоров остались текстом",
           (ws["A2"].value, ws["B3"].value, ws["C2"].value, ws["D2"].value) == (A, "Иванова; Анна", "007-0001", "Кредитная карта"), [c.value for c in ws[2]])
    expect("CSV: папки по ФИО из выгрузки", "Тестов Тест/снимок.png" in tree and "Иванова; Анна/снимок.png" in tree, sorted(n for n in tree if n.endswith(".png")))
    expect("CSV: сам файл выгрузки не тронут", tree["ADW 20.11.2026.csv"] == adw)
    worked = tree["ADW 20.11.2026.xlsx"]
    use_folder(page, "csv")
    expect("CSV: повторное открытие продолжает ту же книгу", opfs_tree(page, "csv")["ADW 20.11.2026.xlsx"] == worked
           and "уже со ссылкой 2" in page.inner_text("body"))
    page.wait_for_timeout(1200)
    opfs_put(page, "ADW 20.11.2026.csv", ("IIN_BIN;FIO_CLIENT\r\n" f"{C};Новый Клиент\r\n").encode("utf-8"), "csv")
    use_folder(page, "csv")
    tree = opfs_tree(page, "csv")
    backups = [n for n in tree if n.startswith("tazalau_backup/ADW 20.11.2026__")]
    ws = load_workbook(io.BytesIO(tree["ADW 20.11.2026.xlsx"]))["Лист1"]
    expect("CSV: новая выгрузка за день: книга собрана заново, прежняя сохранена в tazalau_backup",
           len(backups) == 1 and tree[backups[0]] == worked and ws["A2"].value == C and ws.max_row == 2 and ws["C1"].value is None, (backups, ws["A2"].value))
    page.get_by_text("или один файл").click()
    page.set_input_files("input[type=file]", files=[{"name": "выгрузка.csv", "mimeType": "text/csv", "buffer": adw}])
    page.wait_for_function("() => window.__tazalau.test.state().tasks === 2")
    expect("CSV: и в режиме «один файл»", "выгрузка.xlsx" in page.inner_text("body"))
    sample = (ROOT / "sql" / "ADW_пример_выгрузки.csv").read_bytes()
    opfs_put(page, "ADW_пример_выгрузки.csv", sample, "sample")
    use_folder(page, "sample")
    page.wait_for_function("() => window.__tazalau.test.state().tasks === 6")
    ws = load_workbook(io.BytesIO(opfs_tree(page, "sample")["ADW_пример_выгрузки.xlsx"]))["Лист1"]
    body = page.inner_text("body")
    expect("пример выгрузки ADW (sql/ADW_пример_выгрузки.csv): 6 строк, 5 ИИН, ведущий ноль цел, опечатка показана",
           ws["A5"].value == "001316000001" and "К проверке: 4 ИИН (строк 5)" in body and "Строк с неверным ИИН: 1" in body, body[:600])
    expect("CSV: страница без ошибок", not page.errors, page.errors[:3])
    page.close()

    # ---------- 3е. шаг 1: списки ДГД по регионам ----------
    today = datetime.date.today()
    d0, d1, old = (today.strftime("%d.%m.%Y"), (today - datetime.timedelta(days=1)).strftime("%d.%m.%Y"),
                   (today - datetime.timedelta(days=50)).strftime("%d.%m.%Y"))
    P1, P2, P3 = plausible("8501013"), plausible("9203154"), plausible("7712305")        # Павлодар (PDF)
    C1, C2, C3 = plausible("0101015"), plausible("8807074"), plausible("6605053")        # Алматы (Excel; C1 с ведущим нулём)
    D1, D2, D3 = plausible("9902023"), plausible("8303034"), plausible("7404045")        # Шымкент (CSV в 1251)
    E1, F1, G1, H1, M1 = plausible("9105064"), plausible("8206073"), plausible("7307084"), plausible("9408093"), plausible("8009103")
    K1, K2, K3 = plausible("8610113"), plausible("9711124"), plausible("6812013")        # Костанай (PDF через Ghostscript)
    site.calls.clear()
    site.found = {}
    page = open_helper(ctx, site, {"width": 1366, "height": 768})

    def order_pdf(region, rows):
        maker = ctx.new_page()
        maker.set_content("<html><head><meta charset='utf-8'><style>body{font-family:Arial} td{border:1px solid #000;padding:4px}</style></head><body>"
                          f"<h3>Департамент государственных доходов по {region}</h3><p>Приказы о назначении финансового управляющего, {d0}</p>"
                          "<table><tr><td>№</td><td>ФИО должника</td><td>ИИН должника</td><td>Дата приказа</td><td>Финансовый управляющий</td></tr>" +
                          "".join(f"<tr><td>{n}</td><td>{fio}</td><td>{iin}</td><td>{date}</td><td>Петров Пётр Петрович</td></tr>" for n, (fio, iin, date) in enumerate(rows, 1)) +
                          "</table></body></html>")
        data = maker.pdf()
        maker.close()
        return data

    put = lambda name, data: opfs_put(page, name.split("/")[-1], data, "dgd/" + "/".join(name.split("/")[:-1]) if "/" in name else "dgd")
    put(f"ДГД/Павлодарская {d0}.pdf", order_pdf("Павлодарской области", [("Иванов Иван Иванович", P1, d1), ("Сидорова Анна", P2, d0), ("Ахметов Нурлан", P3, d0)]))
    wb = Workbook()
    sh = wb.active
    sh.append(["Департамент государственных доходов по городу Алматы"])
    sh.append(["№", "ФИО должника", "ИИН должника", "Дата приказа", "ФИО финансового управляющего", "ИИН финансового управляющего"])
    for n, (fio, iin) in enumerate([("Нулев Ноль", int(C1)), ("Второв Второй", C2), ("Третьев Третий", int(C3))], 1):
        sh.append([n, fio, iin, today, "Петров Пётр Петрович", M1])
        sh.cell(sh.max_row, 4).number_format = "DD.MM.YYYY"
    buf = io.BytesIO(); wb.save(buf)
    put(f"ДГД/г. Алматы {d0}.xlsx", buf.getvalue())
    put(f"ДГД/~$г. Алматы {d0}.xlsx", b"lock")
    put(f"ДГД/Шымкент/список {d1}.csv", ("ИИН;ФИО;Дата приказа\r\n" + "".join(f"{i};{n};{d1}\r\n" for i, n in
        [(D1, "Шымкентов Ш."), (D2, "Второва В."), (D3, "Третьякова Т."), (D1[:11] + str((int(D1[11]) + 1) % 10), "Опечаткин О.")])).encode("cp1251"))
    doc = docx.Document()
    doc.add_paragraph("Департамент государственных доходов по области Абай")
    table = doc.add_table(rows=2, cols=4)
    for cell, text in zip(table.rows[0].cells + table.rows[1].cells, ["№", "ФИО должника", "ИИН", "Дата приказа", "1", "Абаев Абай", E1, d0]):
        cell.text = text
    buf = io.BytesIO(); doc.save(buf)
    put(f"ДГД/Абай {d0}.docx", buf.getvalue())
    put("ДГД/Улытау.html", ("<html><head><meta charset='utf-8'></head><body><h3>ДГД по области Ұлытау</h3><table><tr><th>ФИО</th><th>ИИН</th><th>Дата</th></tr>"
                            f"<tr><td>Улытаев Улан</td><td>{F1}</td><td>{old}</td></tr></table></body></html>").encode("utf-8"))
    png = io.BytesIO(); Image.new("RGB", (40, 30), "white").save(png, "PNG")
    put(f"ДГД/Жетысу {d0}.png", png.getvalue())
    plain = PdfReader(io.BytesIO(order_pdf("Атырауской области", [("Закрытов З.", G1, d0)])))
    locked = PdfWriter(clone_from=plain)
    locked.encrypt(user_password="", owner_password="secret")
    buf = io.BytesIO(); locked.write(buf)
    put(f"ДГД/Атырау {d0}.pdf", buf.getvalue())
    wb = Workbook(); wb.active.append(["ФИО", "ИИН"]); wb.active.append(["Безрегионов Б.", G1]); buf = io.BytesIO(); wb.save(buf)
    put(f"ДГД/прочее/список {d0}.xlsx", buf.getvalue())
    gs = shutil.which("gs")
    if gs:
        with tempfile.TemporaryDirectory() as tmp:
            src, dst = Path(tmp) / "in.pdf", Path(tmp) / "out.pdf"
            src.write_bytes(order_pdf("Костанайской области", [("Костанаев К.", K1, d0), ("Второв К.", K2, d0), ("Третьев К.", K3, d0)]))
            subprocess.run([gs, "-q", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.7", "-o", str(dst), str(src)], check=True)
            expect("ДГД: проверка осмысленна: Ghostscript упаковал объекты PDF в потоки", b"/ObjStm" in dst.read_bytes())
            put(f"ДГД/Костанай {d0}.pdf", dst.read_bytes())
    work = fio_book("FIO", [("Нигде Ни", H1), ("Павлодаров П.", P1), ("Нулев Ноль", C1), ("Шымкентов Ш.", D1)])
    opfs_put(page, "работа.xlsx", work, "dgd")
    use_folder(page, "dgd")
    page.wait_for_function("() => { const d = window.__tazalau.test.state().dgd; return d && d !== 'reading'; }", timeout=30000)
    state = page.evaluate("() => window.__tazalau.test.state().dgd")
    regions = {r["name"]: r for r in state["regions"]}
    expect("ДГД: в проверке все 20 регионов и отдельная строка для файлов без региона", len(regions) == 21 and "Регион не определён" in regions, sorted(regions))
    want = {"Павлодарская обл.": ("good", "в порядке", 3), "г. Алматы": ("good", "в порядке", 3), "г. Шымкент": ("warn", "неверных ИИН: 1", 3),
            "Абайская обл.": ("warn", "всего 1 ИИН: проверить вручную", 1), "Жетысуская обл.": ("bad", "0 ИИН: это картинка (скан)", 0),
            "Атырауская обл.": ("bad", "файл не прочитан", 0), "Акмолинская обл.": ("bad", "нет списка", 0),
            "Регион не определён": ("warn", "укажите регион в имени файла", 1)}
    if gs:
        want["Костанайская обл."] = ("good", "в порядке", 3)
    for name, (level, text, unique) in want.items():
        r = regions[name]
        expect(f"ДГД: {name}: {text}", r["level"] == level and text in r["text"] and r["unique"] == unique, r)
    ul = regions["Улытауская обл."]
    expect("ДГД: старый список отмечен (самая свежая дата 50 дней назад) и мало ИИН", ul["level"] == "warn" and "всего 1 ИИН" in ul["text"]
           and f"самая свежая дата {old}: 50 дн. назад" in ul["text"], ul)
    expect("ДГД: ИИН управляющего не считается ИИН должника, файл-замок Excel пропущен", state["iins"] == 3 + 3 + 3 + 1 + 1 + 1 + (3 if gs else 0)
           and state["files"] == 8 + (1 if gs else 0), state)
    panel = page.inner_text("#tazalau-dgd")
    expect("ДГД: окно показывает проверку регионов", "ДГД: регионов со списком" in panel and "⚠ проверить:" in panel, panel[:300])
    expect("ДГД: окно сообщает, сколько ИИН файла есть в списках ДГД", "Из них в списках ДГД: 3 (проверяются первыми)" in page.inner_text("body"))
    final = run_and_wait(page)
    tree = opfs_tree(page, "dgd")
    expect("ДГД: ИИН из списков ДГД проверены первыми", [c[1] for c in site.calls] == [P1, C1, D1, H1], [c[1] for c in site.calls])
    ws = check_xlsx("ДГД работа", work, tree["работа.xlsx"])["ИИН"]
    heads = [ws.cell(1, c).value for c in range(1, 8)]
    expect("ДГД: столбец «ДГД» добавлен в рабочий файл", heads == ["FIO", "ИИН/БИН", "tazalau_file_location", "tazalau_result", "tazalau_status", "tazalau_details", "ДГД"], heads)
    expect("ДГД: регион, дата списка, дата строки, управляющий; ссылка на файл ДГД",
           ws["G3"].value == f"Павлодарская обл., список от {d0}, дата {d1}" and urllib.parse.unquote(ws["G3"].hyperlink.target) == f"ДГД/Павлодарская {d0}.pdf"
           and ws["G4"].value == f"г. Алматы, список от {d0}, дата {d0}, управляющий: Петров Пётр Петрович"
           and ws["G5"].value == f"г. Шымкент, список от {d1}, дата {d1}" and ws["G2"].value in (None, ""),
           [ws.cell(r, 7).value for r in range(2, 6)])
    summary = load_workbook(io.BytesIO(tree["ДГД_свод.xlsx"]))
    reg, iins, files = summary["Регионы"], summary["ИИН"], summary["Файлы"]
    rows = {reg.cell(r, 1).value: [reg.cell(r, c).value for c in range(1, 15)] for r in range(2, reg.max_row + 1)}
    expect("ДГД свод: лист «Регионы»: 21 строка, проверка с цветом", summary.sheetnames == ["Регионы", "Приказы", "ИИН", "Файлы"] and len(rows) == 21
           and rows["Акмолинская обл."][1] == "нет списка" and reg.cell(list(rows).index("Акмолинская обл.") + 2, 2).fill.fgColor.rgb.endswith("FFC7CE")
           and rows["Павлодарская обл."][1] == "в порядке" and reg.cell(list(rows).index("Павлодарская обл.") + 2, 2).fill.fgColor.rgb.endswith("C6EFCE")
           and rows["Павлодарская обл."][2] == "нет" and rows["Павлодарская обл."][11] == d0, rows.get("Павлодарская обл."))
    got = {iins.cell(r, 1).value: [iins.cell(r, c).value for c in range(1, 9)] for r in range(2, iins.max_row + 1)}
    expect("ДГД свод: лист «ИИН»: ИИН текстом (ведущий ноль цел), ФИО, регион, управляющий, файл со ссылкой",
           got[C1][:5] == [C1, "Нулев Ноль", "г. Алматы", d0, "Петров Пётр Петрович"] and got[P1][1] == "Иванов Иван Иванович"
           and iins.cell(list(got).index(P1) + 2, 7).hyperlink is not None and M1 not in got, got.get(C1))
    notes = {files.cell(r, 1).value: files.cell(r, 8).value for r in range(2, files.max_row + 1)}
    expect("ДГД свод: лист «Файлы»: причины по нечитаемым файлам", "PDF защищён" in str(notes.get(f"ДГД/Атырау {d0}.pdf"))
           and "картинка (скан)" in str(notes.get(f"ДГД/Жетысу {d0}.png")) and "контрольная цифра" in str(notes.get(f"ДГД/Шымкент/список {d1}.csv")), notes)
    page.get_by_role("button", name="ИИН для SQL ⧉").click()
    page.wait_for_function("() => /Скопировано ИИН:/.test(document.querySelector('#tazalau-state').textContent)")
    expect("ДГД: «ИИН для SQL» копирует все ИИН списков", f"Скопировано ИИН: {state['iins']}." in page.evaluate(STATE), page.evaluate(STATE))
    options = page.evaluate("() => Array.from(document.querySelectorAll('select option')).map(o => o.value)")
    expect("ДГД: свод не предлагается как рабочий файл", "ДГД_свод.xlsx" not in options, options)
    expect("ДГД: страница без ошибок", not page.errors, page.errors[:3])
    for name, data in [("Павлодарская.pdf", order_pdf("Павлодарской области", [("Один О.", P1, d0)]))]:
        opfs_put(page, name, data, "dgdonly/ДГД")
    page.evaluate("""async () => { const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle('dgdonly');
        await window.__tazalau.test.useDir(dir); }""")
    page.wait_for_function("() => { const d = window.__tazalau.test.state().dgd; return d && d !== 'reading'; }", timeout=30000)
    expect("ДГД без рабочего файла: помощник объясняет", "нет файла с ИИН" in page.evaluate(STATE))
    page.get_by_role("button", name="Свод ДГД ↓").click()
    page.wait_for_function("() => /Свод ДГД записан/.test(document.querySelector('#tazalau-state').textContent)")
    expect("ДГД без рабочего файла: свод пишется кнопкой", "ДГД_свод.xlsx" in opfs_tree(page, "dgdonly"))
    page.close()

    # ---------- 3ж. чтение PDF: простые шрифты, скан, защита ----------
    page = open_helper(ctx, site, {"width": 1366, "height": 768})
    buf = io.BytesIO()
    canvas = rl_canvas.Canvas(buf)
    canvas.setFont("Helvetica", 11)
    canvas.drawString(72, 760, f"IIN {P1}  date {d0}")
    canvas.drawString(72, 740, f"IIN {P2[:4]} {P2[4:8]} {P2[8:]}")
    canvas.save()
    simple = buf.getvalue()
    buf = io.BytesIO()
    canvas = rl_canvas.Canvas(buf)
    canvas.drawImage(ImageReader(Image.new("RGB", (200, 80), "gray")), 72, 600, width=200, height=80)
    canvas.save()
    scanned = buf.getvalue()
    probe = page.evaluate("""async ([a, b]) => { const t = window.__tazalau.test, u = x => Uint8Array.from(atob(x), c => c.charCodeAt(0));
        const r1 = await t.fileRows('простой.pdf', u(a)), r2 = await t.fileRows('скан.pdf', u(b));
        return { simple: t.extractIins(r1.rows).items.map(x => x.iin), note: r2.note }; }""", [base64.b64encode(simple).decode(), base64.b64encode(scanned).decode()])
    expect("PDF: простой шрифт без таблицы символов и ИИН группами по 4 цифры", probe["simple"] == [P1, P2], probe)
    expect("PDF: скан без текста распознан как скан", "скан" in probe["note"], probe)
    page.close()

    # ---------- 3з. ДГД: сбор приказов с сайтов 20 регионов (окно сайта + закладка в нём) ----------
    if TNR:
        now = datetime.datetime.now()
        kgd = KgdSite(now)
        site.kgd = kgd
        seq = iter(range(100, 999))
        made = {}                                                     # ИИН -> (хост, ФИО)

        def person(host, fio, ago, style="kzl", kind="pdf", readable=True, text_layer=False):
            n = next(seq)
            born = now - datetime.timedelta(days=365 * (30 + n % 25) + n)
            iin = plausible(born.strftime("%y%m%d") + str(3 + n % 2))
            fname = f"/sites/default/files/u{n}/prikaz_{n}.{'jpg' if kind == 'jpg' else 'pdf'}"
            if text_layer:
                buf = io.BytesIO(); cv = rl_canvas.Canvas(buf); cv.setFont("Helvetica", 12)
                cv.drawString(72, 760, f"Order {n}: IIN {iin}"); cv.save(); data = buf.getvalue()
            else:
                data = order_scan(iin, fio, kind, blank=not readable)
            kgd.add_file(host, fname, data, "image/jpeg" if kind == "jpg" else "application/pdf", ago)
            link = f'<a href="https://{host}.kgd.gov.kz{fname}">'
            if style == "kzl":
                cells = [f"{link}Приказ о назначении финансовым управляющим Тестов Т.Т. в отношении ФЛ {fio} от {kgd.day(ago)} г.</a>"]
            elif style == "almaty":
                cells = [f"{link}Приказ о назначении финансового управляющего по заявлению должника {fio}</a>", "Сидорова Анна Сергеевна", kgd.day(ago)]
            elif style == "atr":
                cells = [f"{link}{fio}</a>", str(n)]
            else:
                cells = [f"{link}{fio}</a>", kgd.day(ago)]
            kgd.rows.setdefault(host, []).append((cells, [fname]))
            if readable:
                made[iin] = (host, fio)
            return iin

        K1, K2, K3 = (person("kzl", f, a) for f, a in [("Кызылова Кайрата Ивановича", 1), ("Сейтовой Алмы Ериковны", 2), ("Жумабаева Нурлана", 3)])
        person("kzl", "Старова Олега Петровича", 40)
        A1, A2 = person("almaty", "Ахметова Айгерима", 1, "almaty", "jpg"), person("almaty", "Борисовой Веры", 2, "almaty", "jpg")
        person("almaty", "Нечитаемовой Нины Николаевны", 3, "almaty", "jpg", readable=False)
        kgd.rows["almaty"].append((['<a href="https://www.gov.kz/memleket/entities/kgd-almaty/documents/details/123">Приказ о назначении финансового '
                                    'управляющего по заявлению должника Гостевой Галины</a>', "Сидорова Анна Сергеевна", kgd.day(4)], []))
        for f, a in [("Атыраев Арман Нурланович", 1), ("Бекова Дана Маратовна", 2), ("Сагиев Ерлан Кайратович", 3), ("Давний Дмитрий Иванович", 90)]:
            person("atr", f, a, "atr")
        kgd.slug["kst"] = "prikazy-o-naznachenii-i-otstranenii-finansovyh-upravlyayushchih-12"
        for f, a in [("Костин К.К.", 1), ("Лаврова Л.Л.", 4), ("Маратов М.М.", 6)]:
            person("kst", f, a, "plain")
        person("abai", "Абаев А.А.", 2, "plain")
        for f, a in [("Старый С.С.", 60), ("Старая С.С.", 70)]:
            person("akm", f, a, "plain")
        for f, a in [("Павлов П.П.", 20), ("Павлова П.П.", 25)]:
            person("pvl", f, a, "plain")
        V1 = person("vko", "Восточный В.В.", 1, "plain", text_layer=True)
        for f, a in [("Восточная В.В.", 2), ("Восточнов В.В.", 3)]:
            person("vko", f, a, "plain")
        kgd.down.add("mng")
        kgd.no_page.add("zko")
        for host in ["akb", "alm", "zhmb", "zhts", "krg", "sko", "trk", "ult", "astana", "shymkent"]:
            for i in range(3):
                person(host, f"Тестов{host.capitalize()} Т{i}.Т.", 1 + i, "plain")
        N1, X1 = plausible("8801014"), plausible("7702023")
        work = fio_book("FIO", [("ПОСТОРОННИЙ ПЁТР ПЕТРОВИЧ", X1), ("НЕЧИТАЕМОВА НИНА НИКОЛАЕВНА", N1), ("КЫЗЫЛОВ КАЙРАТ ИВАНОВИЧ", K1)])
        site.calls.clear()
        site.found = {}
        page = open_helper(ctx, site, {"width": 1366, "height": 900})
        opfs_put(page, "ADW.xlsx", work, "kgd")
        use_folder(page, "kgd")
        page.wait_for_function("() => { const d = window.__tazalau.test.state().dgd; return d !== 'reading'; }", timeout=30000)
        expect("сбор ДГД: кнопка видна, пока в папке нет ДГД", page.get_by_role("button", name="🌐 Собрать ДГД с сайтов").is_visible())

        def collect_all(stop_after=None):
            with page.expect_popup() as info:
                page.get_by_role("button", name="🌐 Собрать ДГД с сайтов").click()
            popup = info.value
            done = []
            for _ in range(60):
                st = page.evaluate("() => window.__tazalau.test.collecting()")
                if not st or st["finishing"]:
                    break
                cur = next((x for x in st["sites"] if x["state"] in ("open", "work")), None)
                if not cur:
                    break
                host = cur["key"]
                if stop_after is not None and len(done) >= stop_after:
                    page.get_by_role("button", name="Закончить сбор").click()
                    break
                if cur["state"] == "open":
                    if host in kgd.down:
                        page.wait_for_timeout(300)
                        page.get_by_role("button", name="Пропустить регион").click()
                        done.append(host)
                        continue
                    popup.wait_for_url(f"https://{host}.kgd.gov.kz/**", timeout=20000)
                    popup.wait_for_load_state("domcontentloaded")
                    ctx.new_cdp_session(popup).send("Page.navigate", {"url": HREF})        # закладка в окне сайта
                page.wait_for_function("""key => { const c = window.__tazalau.test.collecting(); if (!c) return true;
                    const s = c.sites.find(x => x.key === key); return s.state !== 'open' && s.state !== 'work'; }""", arg=host, timeout=90000)
                done.append(host)
            page.wait_for_function("() => !window.__tazalau.test.collecting() && /ДГД собраны/.test(document.querySelector('#tazalau-state').textContent)",
                                   timeout=180000)
            return done

        done = collect_all()
        said = page.evaluate(STATE)
        expect("сбор ДГД: пройдены все 20 сайтов по очереди, итог в окне", len(done) == 20 and "ДГД собраны: 18 из 20" in said, (done, said))
        tree = opfs_tree(page, "kgd")
        kz = [k for k in tree if k.startswith("ДГД/Кызылординская область/") and not k.endswith(".json")]
        expect("сбор ДГД: скачаны только приказы за 31 день (Кызылорда: 3 из 4), в папку региона", len(kz) == 3
               and "ДГД/Кызылординская область/приказы_с_сайта.json" in tree, sorted(kz))
        al = [k for k in tree if k.startswith("ДГД/г. Алматы/") and not k.endswith(".json")]
        expect("сбор ДГД: JPEG-сканы скачаны, ссылка на gov.kz не скачивается", len(al) == 3 and all(k.endswith(".jpg") for k in al), al)
        at = [k for k in tree if k.startswith("ДГД/Атырауская область/") and not k.endswith(".json")]
        expect("сбор ДГД: строки без даты: дата загрузки файла на сайт (Атырау: 3 свежих из 4)", len(at) == 3, at)
        heads = [c for c in kgd.calls if c[0] == "HEAD"]
        expect("сбор ДГД: для строк без даты сначала запрос даты (HEAD), без скачивания", len(heads) == 4 and all(c[1] == "atr" for c in heads), heads)
        expect("сбор ДГД: распознанные ИИН запомнены в ДГД/распознано.json", "ДГД/распознано.json" in tree
               and K1 in tree["ДГД/распознано.json"].decode("utf-8"))
        regions = {r["name"]: r for r in page.evaluate("() => window.__tazalau.test.state().dgd.regions")}
        want = {"Кызылординская обл.": ("good", "в порядке"), "Костанайская обл.": ("good", "в порядке"),
                "Атырауская обл.": ("good", "в порядке"), "Восточно-Казахстанская обл.": ("good", "в порядке"),
                "г. Алматы": ("warn", "ИИН не прочитан в приказах: 1"), "Абайская обл.": ("warn", "всего 1 приказ за 31 дн."),
                "Акмолинская обл.": ("bad", "0 приказов за 31 дн."), "Павлодарская обл.": ("warn", "всего 2 приказа за 31 дн."),
                "Мангистауская обл.": ("bad", "сайт не открылся"), "Западно-Казахстанская обл.": ("bad", "в меню сайта нет страницы")}
        for name, (level, text) in want.items():
            r = regions[name]
            expect(f"сбор ДГД: проверка региона: {name}: {text}", r["level"] == level and text in r["text"], r)
        expect("сбор ДГД: свежесть: последний приказ 20 дн. назад отмечен", "последний приказ " + kgd.day(20) + ": 20 дн. назад" in regions["Павлодарская обл."]["text"],
               regions["Павлодарская обл."])
        expect("сбор ДГД: приказ на gov.kz отмечен для ручной загрузки", "приказы на gov.kz: 1" in regions["г. Алматы"]["text"], regions["г. Алматы"])
        expect("сбор ДГД: Костанай найден по адресу страницы, хотя пункт меню назван иначе", regions["Костанайская обл."]["win"] == 3, regions["Костанайская обл."])
        dg = page.evaluate("() => window.__tazalau.test.state().dgd")
        read = {i for i in made}
        expect("сбор ДГД: ИИН прочитаны со сканов PDF и JPEG и из текста PDF: все читаемые за 31 день", dg["iins"] >= 40, dg)
        summary = load_workbook(io.BytesIO(tree["ДГД_свод.xlsx"]))
        orders = summary["Приказы"]
        got = {}
        for r in range(2, orders.max_row + 1):
            for iin in str(orders.cell(r, 6).value or "").split(", "):
                got[iin] = [orders.cell(r, c).value for c in range(1, 11)] + [orders.cell(r, 8).hyperlink]
        expect("сбор ДГД: свод, лист «Приказы»: регион, дата, вид, должник, управляющий, ИИН, как прочитан, файл со ссылкой",
               got.get(K1, [None])[0] == "Кызылординская обл." and got[K1][1] == kgd.day(1) and got[K1][2] == "назначение"
               and got[K1][3] == "Кызылова Кайрата Ивановича" and got[K1][4] == "Тестов Т.Т." and got[K1][6] == "распознано со скана"
               and got[K1][10] is not None and got.get(V1, [None] * 7)[6] == "текст PDF" and got.get(A1, [None])[0] == "г. Алматы", got.get(K1))
        wrong = [i for i in got if i and i.isdigit() and i not in made]
        expect("сбор ДГД: ни одного чужого ИИН (только те, что в приказах)", not wrong, wrong)
        missed = [i for i in made if i not in got and kgd.files]
        old_ok = {i for i, (h, f) in made.items() if h in ("akm", "pvl") or f in ("Старова Олега Петровича", "Давний Дмитрий Иванович")}
        expect("сбор ДГД: все читаемые приказы за 31 день прочитаны", not [i for i in missed if i not in old_ok], [made[i] for i in missed if i not in old_ok])
        reg = summary["Регионы"]
        row = {reg.cell(r, 1).value: [reg.cell(r, c).value for c in range(1, 15)] for r in range(2, reg.max_row + 1)}
        expect("сбор ДГД: свод, лист «Регионы»: сбор, последний приказ, приказов за 31 день, прочитано, страница на сайте",
               row["Кызылординская обл."][2] == now.strftime("%d.%m.%Y") and row["Кызылординская обл."][3] == kgd.day(1)
               and row["Кызылординская обл."][4] == "3" and row["Кызылординская обл."][6] == "3" and "kzl.kgd.gov.kz/ru/content/" in str(row["Кызылординская обл."][12])
               and str(row["Мангистауская обл."][2]).startswith("не открылся"), row["Кызылординская обл."])
        final = run_and_wait(page)
        tree = opfs_tree(page, "kgd")
        expect("сбор ДГД: ИИН из приказов ДГД проверяются на tazalau первыми", [c[1] for c in site.calls][:1] == [K1], site.calls[:3])
        ws = load_workbook(io.BytesIO(tree["ADW.xlsx"]))["ИИН"]
        col = [ws.cell(1, c).value for c in range(1, 12)].index("ДГД") + 1
        vals = {ws.cell(r, 2).value: (ws.cell(r, col).value, ws.cell(r, col).hyperlink) for r in range(2, 5)}
        expect("сбор ДГД: в рабочем файле столбец «ДГД»: регион, приказ, дата, управляющий, ссылка на скан",
               str(vals[K1][0]).startswith(f"Кызылординская обл., приказ о назначении от {kgd.day(1)}, управляющий: Тестов Т.Т.") and vals[K1][1] is not None
               and vals[X1][0] in (None, ""), vals)
        expect("сбор ДГД: приказ без прочитанного ИИН сверен по ФИО: «возможно» со ссылкой на скан",
               str(vals[N1][0]).startswith("⚠ возможно (по ФИО, ИИН в приказе не прочитан): г. Алматы, «Нечитаемовой Нины Николаевны»") and vals[N1][1] is not None,
               vals[N1])
        # связной: чужие адреса не скачивает; без помощника только подсказывает
        before = len([c for c in kgd.calls if c[2].startswith("/sites/")])
        done2 = collect_all(stop_after=1)
        after = len([c for c in kgd.calls if c[2].startswith("/sites/")])
        expect("сбор ДГД: повторный сбор не скачивает уже сохранённые приказы", after == before and done2[:1] == ["abai"], (before, after, done2))
        meta = json.loads(opfs_tree(page, "kgd")["ДГД/Абайская область/приказы_с_сайта.json"].decode("utf-8"))
        expect("сбор ДГД: дата «впервые увиден» сохраняется между сборами", meta["rows"][0]["firstSeen"] == now.strftime("%Y-%m-%d"), meta["rows"][0])
        lone = ctx.new_page()
        lone.goto("https://abai.kgd.gov.kz/ru/depsection/fizicheskim-licam")
        ctx.new_cdp_session(lone).send("Page.navigate", {"url": HREF})
        lone.wait_for_function("() => document.body.innerText.includes('Это окно открыл не помощник')", timeout=10000)
        expect("связной: окно сайта, открытое не помощником, только объясняет, что делать", True)
        lone.close()
        with page.expect_popup() as info:
            page.evaluate("() => { window.__probe = window.open('https://kzl.kgd.gov.kz/ru/depsection/fizicheskim-licam', 'probe'); }")
        probe = info.value
        probe.wait_for_load_state("domcontentloaded")
        ctx.new_cdp_session(probe).send("Page.navigate", {"url": HREF})
        probe.wait_for_function("() => !!window.__tazalauRelay")
        answers = page.evaluate("""() => new Promise(done => {
            const got = {};
            const on = e => { if (e.data && e.data.tz === 'fetched') { got[e.data.id] = { ok: e.data.ok, err: e.data.error || '', size: e.data.bytes ? e.data.bytes.byteLength : 0 };
              if (got[1] && got[2]) { window.removeEventListener('message', on); done(got); } } };
            window.addEventListener('message', on);
            window.__probe.postMessage({ tz: 'fetch', id: 1, url: 'https://example.org/secret' }, 'https://kzl.kgd.gov.kz');
            window.__probe.postMessage({ tz: 'fetch', id: 2, url: '/ru/depsection/fizicheskim-licam' }, 'https://kzl.kgd.gov.kz');
            setTimeout(() => done(got), 8000); })""")
        expect("связной: чужой адрес не скачивает, адрес своего сайта скачивает",
               answers.get("1", {}).get("ok") is False and "другого сайта" in answers["1"]["err"] and answers.get("2", {}).get("ok") is True
               and answers["2"]["size"] > 100, answers)
        probe.close()
        expect("сбор ДГД: страница без ошибок", not page.errors, page.errors[:3])
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
    page.get_by_role("button", name="▶ Старт").click()
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
        expect(f"плохой файл: {label}", needle in page.evaluate(STATE) and page.get_by_role("button", name="▶ Старт").is_disabled())
    page.close()
    browser.close()

print(json.dumps(results, ensure_ascii=False, indent=1))
print(f"\n{sum(results.values())} из {len(results)} проверок пройдено")
if problems:
    print("\nПРОБЛЕМЫ:\n  " + "\n  ".join(problems))
    sys.exit(1)
