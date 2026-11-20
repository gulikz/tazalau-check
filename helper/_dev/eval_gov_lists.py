# Проверка на настоящем gov.kz: какие публикации ДГД за 31 день помощник взял бы в каждом из 20 регионов (по видам).
# Скачивает только списки «Документы» (по запросу раз в 4 секунды) и разбирает их кодом помощника. Имён и ИИН не печатает.
# usage: python3 eval_gov_lists.py [дней]
import sys, json, time, re, urllib.request, collections
from pathlib import Path
from playwright.sync_api import sync_playwright

DEV = Path(__file__).resolve().parent
SRC = (DEV.parent / "tazalau_helper.js").read_text(encoding="utf-8")
DAYS = int(sys.argv[1]) if len(sys.argv) > 1 else 31
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
STUB = '<!doctype html><html><head><meta charset="utf-8"></head><body><p>stub</p></body></html>'


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ru"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("utf-8", "replace")


with sync_playwright() as p:
    b = p.chromium.launch(channel="chrome", headless=True)
    ctx = b.new_context()
    ctx.route("**/*", lambda r, q: r.fulfill(status=200, content_type="text/html; charset=utf-8", body=STUB))
    page = ctx.new_page()
    page.goto("https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery")
    page.add_script_tag(content=SRC)
    page.wait_for_function("() => window.__tazalau && window.__tazalau.test")
    sites = page.evaluate("() => window.__tazalau.test.GOV_SITES")
    total = collections.Counter()
    for slug, region in sites:
        items, err = [], ""
        for n in range(1, 13):
            try:
                html = get(f"https://www.gov.kz/org/{slug}/documents/{n}?lang=ru&size=50")
            except Exception as e:
                err = str(e)[:60]
                break
            time.sleep(4)
            r = page.evaluate("""([html, days]) => { const t = window.__tazalau.test, fd = (t.govData(html).getFilteredData || {}).value;
                if (!fd || !Array.isArray(fd.data)) return null;
                const since = Date.now() - days * 86400000;
                return { count: fd.count, last: fd.data.length ? fd.data[fd.data.length - 1].created_date : null,
                  items: fd.data.filter(x => new Date(x.created_date) >= since).map(x => ({ mine: t.isGovCitizenDoc(x.title),
                    kind: t.govKind(x.title), debtor: !!t.govDebtor(x.title) })) }; }""", [html, DAYS])
            if r is None:
                err = "нет списка"
                break
            items += r["items"]
            if not r["last"] or time.time() * 1000 - DAYS * 86400000 > time.mktime(time.strptime(r["last"][:19], "%Y-%m-%dT%H:%M:%S")) * 1000 or n * 50 >= (r["count"] or 0):
                break
        mine = [x for x in items if x["mine"]]
        kinds = collections.Counter(x["kind"] for x in mine)
        total.update(kinds)
        print(f"{region:28} {slug:14} всего за {DAYS} дн.: {len(items):4}  о гражданах: {len(mine):3}  с ФИО в названии: {sum(x['debtor'] for x in mine):3}  "
              + (f"ОШИБКА {err}  " if err else "") + "; ".join(f"{k}: {v}" for k, v in kinds.most_common()), flush=True)
    print("ИТОГО по видам:", json.dumps(dict(total.most_common()), ensure_ascii=False))
    b.close()
