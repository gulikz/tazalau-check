"""Проверка «браузер запускается и рисует страницу» без обращения к сайту.

    python tools/ci_browser_check.py msedge chrome chromium

Для каждого названного браузера: запустить его так же, как это делает программа, отрисовать
страницу с кириллицей и убедиться, что снимок не пустой. Нужна, чтобы на настоящей Windows
отделить вопрос «работает ли браузер под управлением программы» от вопроса «доступен ли сайт».
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from tazalau.capture import BrowserStartError, TazalauBrowser, looks_blank   # noqa: E402
from tazalau.config import Settings                                           # noqa: E402

PAGE = ("<html><body style='font:40px sans-serif;background:#fff;color:#000'>"
        "<h1>Проверка браузера</h1><p>ИИН 000000000000</p></body></html>")


def check(channel: str) -> bool:
    try:
        with TazalauBrowser(Settings(browser_channels=[channel])) as browser:
            page = browser._page
            page.set_content(PAGE)
            png = page.screenshot(full_page=True)
            ok = not looks_blank(png)
            print(f"{channel:9} {'OK ' if ok else 'ПУСТОЙ СНИМОК'}  {browser.description}, снимок {len(png)} байт")
            return ok
    except BrowserStartError as exc:
        print(f"{channel:9} НЕ ЗАПУСТИЛСЯ: {str(exc).splitlines()[-1]}")
        return False


if __name__ == "__main__":
    results = [check(channel) for channel in (sys.argv[1:] or ["msedge", "chrome", "chromium"])]
    sys.exit(0 if all(results) else 1)
