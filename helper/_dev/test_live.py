# Проверка закладки-помощника на НАСТОЯЩЕМ сайте tazalau.qoldau.kz (несколько запросов, с обычной паузой 2 с).
# usage: python3 test_live.py
# На сайт уходит только заведомо несуществующий ИИН из нулей и ИИН с невозможной датой рождения.
# Чтобы проверить и случай «найден», задайте переменную окружения TAZALAU_TEST_IIN.
import base64
import html
import io
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook, load_workbook
from PIL import Image, ImageChops
from playwright.sync_api import sync_playwright

DEV = Path(__file__).resolve().parent
sys.path.insert(0, str(DEV.parent.parent))
from tazalau.iin import check_digit                      # noqa: E402

HREF = html.unescape(re.search(r'<a class="bm" id="bm" href="([^"]+)"',
                               (DEV.parent / "Tazalau_helper.html").read_text(encoding="utf-8")).group(1))
START = "https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery"
ZERO = "000000000000"
FAKE = "00139990002" + str(check_digit("00139990002"))      # 13-й месяц: такого человека нет; как число теряет два нуля
KNOWN = os.environ.get("TAZALAU_TEST_IIN", "")
BASE = r"W:\Отдел\Проверка связи\live"
DONE = "() => /^(✔ Готово\\.|⏹ Остановлено\\.|❌ )/.test(document.querySelector('#tazalau-state').textContent)"
problems = []


def expect(name, condition, detail=""):
    print(("  ok   " if condition else "  FAIL ") + name + ("" if condition else f"  [{detail}]"))
    if not condition:
        problems.append(name)


wb = Workbook()
ws = wb.active
ws.title = "восстановление"
ws.append(["ФИО", "ИИН/БИН"])
rows = [("нет такого", ZERO), ("потерянные нули", int(FAKE))] + ([("известный", int(KNOWN))] if KNOWN else [])
for row in rows:
    ws.append(list(row))
buffer = io.BytesIO()
wb.save(buffer)
original = buffer.getvalue()

with sync_playwright() as p:
    try:
        browser = p.chromium.launch(channel="chrome")
    except Exception:
        browser = p.chromium.launch(channel="msedge")
    ctx = browser.new_context(viewport={"width": 1366, "height": 768}, locale="ru-RU", service_workers="block")
    page = ctx.new_page()
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.goto(START, wait_until="domcontentloaded", timeout=60000)
    ctx.new_cdp_session(page).send("Page.navigate", {"url": HREF})        # как нажатие на закладку
    page.wait_for_function("() => !!window.__tazalau", timeout=20000)
    page.evaluate("""async b64 => {
        const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle('live', { create: true });
        const w = await (await dir.getFileHandle('проверка.xlsx', { create: true })).createWritable();
        await w.write(Uint8Array.from(atob(b64), c => c.charCodeAt(0)));
        await w.close();
        await window.__tazalau.test.useDir(dir);
    }""", base64.b64encode(original).decode())
    page.wait_for_function("() => window.__tazalau.test.state().sheet")
    page.fill("#tazalau-base", BASE)                      # путь к папке на диске: для столбца hyperlink
    page.locator("#tazalau-base").dispatch_event("change")
    started = time.time()
    page.get_by_text("▶ Старт").click()
    page.wait_for_function(DONE, timeout=240000)
    took = time.time() - started
    final = page.evaluate("() => document.querySelector('#tazalau-state').textContent")
    timings = page.evaluate("() => window.__tazalau.test.timings")
    tree = page.evaluate("""async () => {
        const root = await (await navigator.storage.getDirectory()).getDirectoryHandle('live');
        const out = {};
        const walk = async (dir, prefix) => { for await (const [name, h] of dir.entries()) {
            if (h.kind === 'directory') await walk(h, prefix + name + '/');
            else { const b = new Uint8Array(await (await h.getFile()).arrayBuffer()); let s = '';
                   for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); out[prefix + name] = btoa(s); } } };
        await walk(root, '');
        return out;
    }""")
    tree = {k: base64.b64decode(v) for k, v in tree.items()}
    print(f"\nПомощник на настоящем сайте: {len(rows)} ИИН за {took:.1f} с; по страницам, мс: {timings}")
    print(final.replace("\n", " | "))
    expect("итог без ошибок", final.startswith("✔ Готово.") and "ошибок" not in final, final)
    shots = sorted(n for n in tree if n.endswith(".png"))
    expect("по снимку на каждый ИИН", len(shots) == len(rows), shots)
    patched = tree["проверка.xlsx"]
    archive = zipfile.ZipFile(io.BytesIO(patched))
    for part in archive.namelist():
        if part.endswith((".xml", ".rels")):
            ET.fromstring(archive.read(part))
    sheet = load_workbook(io.BytesIO(patched))["восстановление"]
    heads = [sheet.cell(1, c).value for c in range(1, 6)]
    expect("столбцы: ссылка, полная ссылка, итог", heads == ["ФИО", "ИИН/БИН", "tazalau_file_location", "hyperlink", "tazalau_result"], heads)
    expect("ссылки и итог записаны в книгу", all(sheet.cell(r, 3).hyperlink for r in range(2, 2 + len(rows)))
           and str(sheet["E2"].value).startswith("НЕ НАЙДЕН: Восстановление платежеспособности"), sheet["E2"].value)
    expect("папка названа по столбцу ФИО, снимок называется «снимок.png»",
           sheet["C3"].hyperlink.target == "потерянные%20нули/снимок.png"
           and "потерянные нули/снимок.png" in tree, sheet["C3"].hyperlink.target)
    expect("столбец hyperlink: полный путь = путь к папке + папка человека + снимок.png",
           sheet["D3"].value == '=HYPERLINK("' + BASE + '\\потерянные нули\\снимок.png")', sheet["D3"].value)
    if KNOWN:
        expect("известный ИИН найден", str(sheet["E4"].value).startswith("НАЙДЕН:"), sheet["E4"].value)
    expect("в окне помощника нет ошибок", not errors, errors[:3])

    # Снимок помощника против настоящего снимка экрана той же страницы
    time.sleep(2.2)
    target = KNOWN or ZERO
    real = ctx.new_page()
    real.set_viewport_size({"width": 1920, "height": 1000})
    real.goto(f"{START}?flApplicantXin={target}", wait_until="domcontentloaded", timeout=60000)
    real.wait_for_function("() => !document.body.classList.contains('loading') && document.fonts.status === 'loaded'", timeout=30000)
    real.wait_for_timeout(700)
    reference = Image.open(io.BytesIO(real.screenshot(full_page=True))).convert("RGB")
    mine = Image.open(io.BytesIO(tree[f"{'известный' if KNOWN else 'нет такого'}/снимок.png"])).convert("RGB")
    strip = mine.height - reference.height
    expect("снимок: ширина 1920 и полоса-подпись сверху", mine.width == 1920 and strip > 40 and mine.getpixel((3, 3)) == (36, 52, 71),
           (mine.size, reference.size))
    if strip > 0 and mine.width == reference.width:
        body = mine.crop((0, strip, mine.width, mine.height))
        # нижнюю строку со счётчиком посещений сайта не сравниваем: она меняется сама
        box = (0, 0, reference.width, reference.height - 110)
        a, b = body.crop(box), reference.crop(box)
        # Сравнение с допуском в одну точку: движок рисует страницу внутри SVG и на экране с разным округлением
        # дробных координат, и блок целиком может лечь на точку выше или ниже. Это не отличие в содержимом.
        diff = None
        for dx, dy in ((0, 0), (0, 1), (0, -1), (1, 0), (-1, 0)):
            step = ImageChops.difference(a, ImageChops.offset(b, dx, dy)).convert("L")
            diff = step if diff is None else ImageChops.darker(diff, step)
        diff = diff.crop((1, 1, box[2] - 1, box[3] - 1))
        histogram = diff.histogram()
        total = diff.width * diff.height
        strong = sum(histogram[60:])
        print(f"  отличий от настоящего снимка экрана (с допуском в одну точку): заметных {strong} точек из {total}, любых {100 * (total - histogram[0]) / total:.2f}%")
        # первый снимок после запуска помощника: именно он раньше выходил без значка в меню (шрифт не успевал подключиться)
        expect("первый снимок сессии совпадает с настоящим снимком экрана, значки и шрифты на месте", strong <= 10, f"{strong} точек")
        out = DEV / "out"
        out.mkdir(exist_ok=True)
        mine.save(out / "live_helper.png")
        reference.save(out / "live_reference.png")
    browser.close()

print("\nВСЁ В ПОРЯДКЕ" if not problems else "\nПРОБЛЕМЫ: " + "; ".join(problems))
sys.exit(1 if problems else 0)
