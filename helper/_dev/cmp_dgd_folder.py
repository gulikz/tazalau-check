# Одна и та же папка «ДГД» глазами закладки (readDgd в браузере) и пакета tazalau_daily (Python): проверка регионов
# должна совпасть до строки. Сеть не нужна. Печатает только числа и тексты проверки (без имён и ИИН).
# usage: python3 cmp_dgd_folder.py <рабочая папка, внутри которой папка ДГД>
import base64
import re
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

DEV = Path(__file__).resolve().parent
sys.path.insert(0, str(DEV.parent.parent))
from tazalau_daily.dgd import read_dgd          # noqa: E402

WORK = Path(sys.argv[1])
SRC = (DEV.parent / "Tazalau_helper.html").read_text(encoding="utf-8")
import html as _html                             # noqa: E402
HREF = _html.unescape(re.search(r'<a class="bm" id="bm" href="([^"]+)"', SRC).group(1))
STUB = '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>tazalau</title></head><body><p>stub</p></body></html>'

py = {r["name"]: r for r in read_dgd(WORK, use_ocr=False, log=lambda t: None)["regions"]}
with sync_playwright() as p:
    b = p.chromium.launch(channel="chrome", headless=True)
    ctx = b.new_context()
    ctx.route("**/*", lambda route, req: route.fulfill(status=200, headers={"content-type": "text/html; charset=utf-8"}, body=STUB))
    page = ctx.new_page()
    page.goto("https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery")
    ctx.new_cdp_session(page).send("Page.navigate", {"url": HREF})
    page.wait_for_function("() => !!window.__tazalau")
    files = [f for f in (WORK / "ДГД").rglob("*") if f.is_file() and f.name != "распознано.json"]
    for f in files:
        rel = f.relative_to(WORK).as_posix()
        page.evaluate("""async ([rel, b64]) => {
            let d = await (await navigator.storage.getDirectory()).getDirectoryHandle('cmp', { create: true });
            const parts = rel.split('/');
            for (const x of parts.slice(0, -1)) d = await d.getDirectoryHandle(x, { create: true });
            const w = await (await d.getFileHandle(parts[parts.length - 1], { create: true })).createWritable();
            await w.write(Uint8Array.from(atob(b64), c => c.charCodeAt(0))); await w.close(); }""", [rel, base64.b64encode(f.read_bytes()).decode()])
    page.evaluate("""async () => { const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle('cmp');
        await window.__tazalau.test.useDir(dir); }""")
    page.wait_for_function("() => { const d = window.__tazalau.test.state().dgd; return d && d !== 'reading' && d.regions; }", timeout=900000)
    js = page.evaluate("""() => window.__tazalau.test.state().dgd.regions.map(r => ({ name: r.name, level: r.level, text: r.text,
        rows: r.annRows, uniq: r.annUnique, posted: r.annPosted }))""")
    b.close()

diff = 0
for r in js:
    q = py.get(r["name"])
    pv = (q["annRows"], q["annUnique"], q["annPosted"].isoformat() if q["annPosted"] else "", q["level"]) if q else None
    jv = (r["rows"], r["uniq"], r["posted"], r["level"])
    same = pv == jv and q["text"] == r["text"]
    diff += not same
    print(f"{r['name'][:28]:28} {'OK  ' if same else 'DIFF'} закладка {jv} | Python {pv}")
    if not same and q:
        print(f"     закладка: {r['text'][:160]}\n     Python:   {q['text'][:160]}")
print("регионов с расхождением:", diff, "из", len(js))
