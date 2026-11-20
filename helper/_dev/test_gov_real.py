# Сбор ДГД кодом помощника с НАСТОЯЩЕГО gov.kz (все 20 регионов). Браузер к gov.kz не ходит: его запросы к gov.kz
# выполняет этот скрипт обычными запросами (не чаще одного в 1,5 секунды) и отдаёт ответы браузеру. tazalau заменён заглушкой.
# Итог: папка с файлами ДГД и свод ДГД_свод.xlsx в каталоге, переданном первым аргументом (там персональные данные:
# в репозиторий не класть). Печатает только числа.
# usage: python3 test_gov_real.py <папка для итога>
import base64, html, json, re, sys, time, urllib.request
from pathlib import Path
from playwright.sync_api import sync_playwright

DEV = Path(__file__).resolve().parent
OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
SRC = (DEV.parent / "Tazalau_helper.html").read_text(encoding="utf-8")
HREF = html.unescape(re.search(r'<a class="bm" id="bm" href="([^"]+)"', SRC).group(1))
UA = "TazalauCheck/1.0 (test of the bookmark helper)"
STUB = '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>tazalau</title></head><body><p>stub</p></body></html>'
last = [0.0]
stats = {"requests": 0, "errors": 0, "bytes": 0}


def gov(route, request):
    if request.method not in ("GET", "HEAD"):
        return route.abort()
    wait = 1.5 - (time.time() - last[0])
    if wait > 0:
        time.sleep(wait)
    last[0] = time.time()
    stats["requests"] += 1
    for attempt in range(3):
        try:
            req = urllib.request.Request(request.url, method=request.method, headers={"User-Agent": UA, "Accept-Language": "ru"})
            with urllib.request.urlopen(req, timeout=90) as r:
                body = r.read()
                stats["bytes"] += len(body)
                hdr = {k.lower(): v for k, v in r.headers.items() if k.lower() in ("content-type", "content-disposition", "last-modified")}
                return route.fulfill(status=r.status, headers=hdr, body=body)
        except urllib.error.HTTPError as e:
            return route.fulfill(status=e.code, headers={"content-type": "text/plain"}, body=b"")
        except Exception:
            time.sleep(8)
    stats["errors"] += 1
    return route.abort("timedout")


GOV_STUB = '<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>gov.kz</title></head><body><p>gov.kz</p></body></html>'


def handler(route, request):
    host = request.url.split("/")[2]
    if host in ("www.gov.kz", "gov.kz"):
        if request.resource_type == "document":          # помощник 7 не открывает gov.kz во фреймах: документ = сама вкладка
            # вкладка gov.kz: помощнику нужен только адрес www.gov.kz, тяжёлую страницу с десятками файлов не грузим
            return route.fulfill(status=200, headers={"content-type": "text/html; charset=utf-8"}, body=GOV_STUB)
        if request.resource_type in ("image", "font", "stylesheet", "script", "media"):
            return route.abort()
        return gov(route, request)
    if host == "tazalau.qoldau.kz":
        return route.fulfill(status=200, headers={"content-type": "text/html; charset=utf-8"}, body=STUB)
    return route.abort()


with sync_playwright() as p:
    b = p.chromium.launch(channel="chrome", headless=True)
    ctx = b.new_context(viewport={"width": 1366, "height": 900})
    ctx.route("**/*", handler)
    page = ctx.new_page()
    print("старт", flush=True)
    page.goto("https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery")
    ctx.new_cdp_session(page).send("Page.navigate", {"url": HREF})
    page.wait_for_function("() => !!window.__tazalau")
    page.evaluate("""async () => { const root = await navigator.storage.getDirectory(); const dir = await root.getDirectoryHandle('real', { create: true });
        await window.__tazalau.test.useDir(dir); }""")
    page.wait_for_function("() => { const d = window.__tazalau.test.state().dgd; return d !== 'reading'; }", timeout=60000)
    page.get_by_role("button", name="🌐 Собрать ДГД с gov.kz").click()
    with page.expect_popup() as info:
        page.get_by_role("button", name="Открыть gov.kz").click()
    tab = info.value
    tab.wait_for_load_state("domcontentloaded", timeout=120000)
    print("вкладка gov.kz открыта:", tab.url[:60], flush=True)
    ctx.new_cdp_session(tab).send("Page.navigate", {"url": HREF})
    t0 = time.time()
    while True:
        st = page.evaluate("() => window.__tazalau.test.collecting()")
        if not st:
            break
        done = [x for x in st["sites"] if x["state"] in ("done", "fail", "skip")]
        print(f"{int(time.time() - t0):5}s регионов {len(done)}/20 запросов {stats['requests']} ошибок {stats['errors']}", flush=True)
        if time.time() - t0 > 7200:
            page.get_by_role("button", name="Закончить сбор").click()
        page.wait_for_timeout(30000)                      # ждать через Playwright: так он обслуживает запросы страницы
    page.wait_for_function("() => /ДГД собраны/.test(document.querySelector('#tazalau-state').textContent)", timeout=600000)
    print("ИТОГ:", page.evaluate("() => document.querySelector('#tazalau-state').textContent"))
    print(page.inner_text("#tazalau-collect"))
    files = page.evaluate("""async () => { const out = {}; const root = await (await navigator.storage.getDirectory()).getDirectoryHandle('real');
        const walk = async (d, pre) => { for await (const [n, h] of d.entries()) { if (h.kind === 'directory') await walk(h, pre + n + '/');
          else { const f = await h.getFile(); const b = new Uint8Array(await f.arrayBuffer()); let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); out[pre + n] = btoa(s); } } };
        await walk(root, ''); return out; }""")
    for name, b64 in files.items():
        dst = OUT / name
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(base64.b64decode(b64))
    print("файлов сохранено:", len(files), "запросов к gov.kz:", stats["requests"], "МБ:", round(stats["bytes"] / 1e6, 1))
    b.close()
