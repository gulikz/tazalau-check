# Разбор настоящих страниц приказов ДГД (сохранённых локально, например из веб-архива) кодом помощника.
# Печатает только сводные числа: имён и ИИН не выводит. Страницы с персональными данными в репозиторий не кладутся.
# usage: python3 eval_dgd_pages.py <папка со страницами: <регион>/*.html>
import sys, json
from pathlib import Path
from playwright.sync_api import sync_playwright

DEV = Path(__file__).resolve().parent
SRC = (DEV.parent / "tazalau_helper.js").read_text(encoding="utf-8")
STUB = '<!doctype html><html><head><meta charset="utf-8"></head><body><p>stub</p></body></html>'
root = Path(sys.argv[1])
with sync_playwright() as p:
    b = p.chromium.launch(channel="chrome", headless=True)
    ctx = b.new_context()
    ctx.route("**/*", lambda r, q: r.fulfill(status=200, content_type="text/html; charset=utf-8", body=STUB))
    page = ctx.new_page()
    page.goto("https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery")
    page.add_script_tag(content=SRC)
    page.wait_for_function("() => window.__tazalau && window.__tazalau.test")
    total = {"pages": 0, "rows": 0, "dated": 0, "files": 0, "debtor": 0, "kind": 0, "manager": 0}
    for region in sorted(x for x in root.iterdir() if x.is_dir()):
        for f in sorted(region.glob("*.html")):
            html = f.read_text(encoding="utf-8", errors="replace")
            if len(html) < 2000:
                continue
            host = region.name + ".kgd.gov.kz"
            r = page.evaluate("""([html, url]) => { const t = window.__tazalau.test;
                const pages = t.orderPages(html, url), rows = t.orderRows(html, url);
                return { menu: pages.map(p => ({ year: p.year, title: p.title.slice(0, 60) })),
                  rows: rows.length, dated: rows.filter(r => r.date).length, files: rows.filter(r => r.files.length).length,
                  gov: rows.filter(r => r.files.some(f => f.gov)).length, debtor: rows.filter(r => r.debtor).length,
                  manager: rows.filter(r => r.manager).length, kind: rows.filter(r => r.kind).length,
                  years: rows.reduce((a, r) => { const y = r.date ? r.date.getFullYear() : 'нет'; a[y] = (a[y] || 0) + 1; return a; }, {}) }; }""",
                [html, f"https://{host}/ru/content/{f.stem.split('__')[0]}"])
            total["pages"] += 1
            for k in ("rows", "dated", "files", "debtor", "kind", "manager"):
                total[k] += r[k]
            print(f"{region.name:9} {f.stem.split('__')[0][-28:]:28} строк {r['rows']:4} с датой {r['dated']:4} с файлом {r['files']:4} gov.kz {r['gov']:3}"
                  f" должник {r['debtor']:4} управляющий {r['manager']:4} вид {r['kind']:4} годы {r['years']}  меню: {len(r['menu'])} стр. {[m['year'] for m in r['menu']]}")
    print("ИТОГО", json.dumps(total, ensure_ascii=False))
    b.close()
