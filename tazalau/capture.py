"""Браузер: открыть реестр с фильтром по ИИН, убедиться, что на экране именно результат поиска
по этому ИИН, и снять снимок всей страницы.

Используется установленный в системе Microsoft Edge или Google Chrome (в скрытом режиме).
Входа на сайт и ЭЦП не требуется: реестры открытые.

Что здесь важно и почему (всё найдено при проверке сайта, а не придумано):
  * Сайт 0,4 с после загрузки держит ``<body class="loading">`` невидимым. Снимок, сделанный
    в этот момент, — пустой серый лист. Поэтому ждём снятия класса и проверяем, что снимок не пуст.
  * При ширине окна 1600 таблица не помещается и правые столбцы уходят под прокрутку. Если
    таблица шире своей области, окно расширяется до тех пор, пока она не поместится целиком.
  * Сайт ставит service worker, который при обрыве связи подсовывает страницу «офлайн» с кодом
    200. Он отключён, а каждая страница проверяется по содержимому: поле фильтра должно
    содержать наш ИИН, счётчик «Всего записей» — совпадать с таблицей.
  * При частых запросах (чаще ~1 в секунду) сайт вместо реестра отдаёт страницу «Превышен
    лимит запросов!» — тоже с кодом 200. Поэтому запросы идут не чаще заданного интервала,
    а такая страница распознаётся отдельно: программа ждёт и дальше работает медленнее.
"""
from __future__ import annotations

import io
import os
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Callable
from urllib.parse import parse_qs, urlsplit

from PIL import Image

from .config import Settings, bundled_browsers_dir, load_prefs, update_prefs
from .registries import Registry

# Сторонние счётчики посещений: на снимок не влияют, а в закрытой сети их ожидание
# задерживает загрузку каждой страницы. Блокируются только они, сам сайт — нет.
_BLOCKED_URLS = ["*googletagmanager.com*", "*google-analytics.com*", "*doubleclick.net*"]

_READY_JS = """() => document.readyState === 'complete'
    && !!document.body
    && !document.body.classList.contains('loading')
    && getComputedStyle(document.body).visibility === 'visible'
    && (!document.fonts || document.fonts.status === 'loaded')"""

_INFO_JS = """(param) => {
    const input = document.querySelector('input[name="' + param + '"]');
    const totals = Array.from(document.querySelectorAll('ul.pagination')).map(ul => {
        const span = ul.previousElementSibling;
        if (!span || span.tagName !== 'SPAN') return null;
        let text = '';
        span.childNodes.forEach(n => { if (n.nodeType === Node.TEXT_NODE) text += n.textContent; });
        text = text.replace(/[\\s\\u00a0\\u202f]/g, '');
        return /^\\d+$/.test(text) ? parseInt(text, 10) : null;
    });
    const table = document.querySelector('.table-responsive table') || document.querySelector('table.table');
    let headerCols = 0, dataRows = 0, emptyRow = false;
    if (table) {
        const head = table.querySelector('thead tr');
        headerCols = head ? head.children.length : 0;
        table.querySelectorAll('tbody > tr').forEach(tr => {
            const cells = tr.children;
            if (cells.length === 1 && cells[0].colSpan > 1) emptyRow = true; else dataRows += 1;
        });
    }
    const wrap = table ? table.closest('.table-responsive') : null;
    const menu = document.querySelector('.left-side-menu');
    const doc = document.documentElement;
    const tableOverflow = wrap ? wrap.scrollWidth - wrap.clientWidth : 0;
    const pageOverflow = doc.scrollWidth - window.innerWidth;
    return {
        path: location.pathname, search: location.search, title: document.title,
        notice: document.body ? (document.body.innerText || '').slice(0, 300) : '',
        inputFound: !!input, inputValue: input ? input.value : null,
        totals, tableFound: !!table, headerCols, dataRows, emptyRow,
        styled: !!menu && getComputedStyle(menu).position === 'fixed',
        overflow: Math.max(0, tableOverflow, pageOverflow),
    };
}"""

_SETTLE_JS = "() => new Promise(done => requestAnimationFrame(() => requestAnimationFrame(() => done(true))))"


class CaptureError(Exception):
    """Страницу не удалось получить или она не похожа на результат поиска. Текст — для человека."""


class RateLimitedError(CaptureError):
    """Сайт ответил «Превышен лимит запросов» — нужно подождать и работать медленнее."""


class BrowserStartError(Exception):
    """Не удалось запустить браузер — продолжать бессмысленно."""


_LIMIT_MARK = "лимит запросов"


@dataclass(frozen=True)
class PageShot:
    registry: Registry
    iin: str
    url: str
    count: int                  # «Всего записей» на странице
    rows_on_page: int           # строк в таблице (меньше count, если записей больше одной страницы)
    png: bytes
    checked_at: datetime        # местное время с часовым поясом
    clipped: bool = False       # таблица не поместилась даже при максимальной ширине


def looks_blank(png: bytes) -> bool:
    """Пустой (одноцветный) снимок: на настоящей странице есть и тёмный текст, и белый фон."""
    with Image.open(io.BytesIO(png)) as image:
        low, high = image.convert("L").getextrema()
    return high - low < 60


def verify_page(info: dict, registry: Registry, iin: str, max_plausible: int) -> tuple[int, int]:
    """Проверяет, что страница — результат поиска по данному ИИН. Возвращает (всего, строк)."""
    banner = f"{info.get('title') or ''} {info.get('notice') or ''}".lower()
    if _LIMIT_MARK in banner and not info.get("inputFound"):
        raise RateLimitedError("сайт сообщил «Превышен лимит запросов»")
    path = str(info.get("path") or "").rstrip("/")
    if path != registry.path.rstrip("/"):
        raise CaptureError(f"сайт открыл другую страницу ({path or 'неизвестно'})")
    if not info.get("styled"):
        raise CaptureError("страница загрузилась без оформления")
    if not info.get("inputFound"):
        raise CaptureError("на странице нет поля «ИИН» — возможно, сайт изменился")
    if info.get("inputValue") != iin:
        raise CaptureError("сайт не применил фильтр по ИИН")
    query = parse_qs(str(info.get("search") or "").lstrip("?"))
    if query.get(registry.param) != [iin]:
        raise CaptureError("адрес страницы не соответствует запросу")
    if not info.get("tableFound"):
        raise CaptureError("на странице нет таблицы результатов — возможно, сайт изменился")
    totals = info.get("totals") or []
    if not totals or any(t is None for t in totals) or len(set(totals)) != 1:
        raise CaptureError("не удалось прочитать счётчик «Всего записей»")
    count, rows = int(totals[0]), int(info.get("dataRows") or 0)
    if count > max_plausible:
        raise CaptureError(f"по одному ИИН показано {count} записей — похоже, фильтр не сработал")
    if count == 0 and (rows != 0 or not info.get("emptyRow")):
        raise CaptureError("счётчик записей не совпадает с таблицей")
    if count > 0 and not (1 <= rows <= count):
        raise CaptureError("счётчик записей не совпадает с таблицей")
    return count, rows


class TazalauBrowser:
    """Один скрытый браузер на весь запуск. Использовать из одного потока."""

    def __init__(self, settings: Settings, log: Callable[[str], None] | None = None,
                 sleep: Callable[[float], None] | None = None):
        self.settings = settings
        self._log = log or (lambda text: None)
        self._sleep = sleep or time.sleep            # подменяется на прерываемое ожидание
        self._playwright = self._browser = self._context = self._page = None
        self.description = ""
        self.interval = settings.min_request_interval_s
        self._last_request: float | None = None

    # ------------------------------------------------------------------ темп запросов
    def _pace(self) -> None:
        """Выдерживает паузу между началами соседних запросов."""
        if self._last_request is not None:
            wait = self._last_request + self.interval - time.monotonic()
            if wait > 0:
                self._sleep(wait)
        self._last_request = time.monotonic()

    def slow_down(self) -> float:
        """Сайт попросил не торопиться: до конца запуска работаем в полтора раза медленнее."""
        self.interval = min(self.settings.max_request_interval_s,
                            max(self.interval * 1.5, self.settings.min_request_interval_s, 1.0))
        return self.interval

    def __enter__(self) -> "TazalauBrowser":
        self.start()
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()

    # ------------------------------------------------------------------ запуск
    def start(self) -> str:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise BrowserStartError(
                "Не установлен компонент Playwright (pip install playwright).") from exc
        settings = self.settings
        browsers = bundled_browsers_dir()
        if browsers is not None:          # встроенный браузер — до запуска драйвера Playwright
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
        try:
            self._playwright = sync_playwright().start()
        except Exception as exc:
            raise BrowserStartError(f"Не удалось запустить Playwright: {_short(exc)}") from exc

        problems: list[tuple[str, str]] = []
        for channel, headless in launch_plan(settings, load_prefs()):
            options: dict = {"headless": headless,
                             "timeout": settings.browser_launch_timeout_s * 1000}
            if channel != "chromium":
                options["channel"] = channel
            if settings.proxy_server:
                options["proxy"] = {"server": settings.proxy_server}
            try:
                self._browser = self._playwright.chromium.launch(**options)
            except Exception as exc:
                problems.append((channel, _short(exc)))
                continue
            self.description = f"{_CHANNEL_NAMES.get(channel, channel)} {self._browser.version}"
            if not headless and settings.headless:
                self.description += " (в видимом окне)"
                self._log("Скрытый режим браузера на этом компьютере недоступен, поэтому браузер "
                          "открыт обычным окном. Не закрывайте его до конца проверки.")
            # В следующий раз начнём сразу с того, что сработало, без ожидания неудачных попыток.
            update_prefs({"browser_channel": channel, "browser_headless": headless})
            break
        if self._browser is None:
            self.close()
            raise BrowserStartError(explain_launch_failure(problems))

        context_options: dict = {
            "viewport": {"width": settings.viewport_width, "height": settings.viewport_height},
            "device_scale_factor": settings.device_scale_factor,
            "locale": "ru-RU",
            "service_workers": "block",
            "accept_downloads": False,
        }
        if settings.user_agent:
            context_options["user_agent"] = settings.user_agent
        try:
            self._context = self._browser.new_context(**context_options)
            self._page = self._context.new_page()
        except Exception as exc:
            self.close()
            raise BrowserStartError(f"Не удалось открыть окно браузера: {_short(exc)}") from exc
        try:
            session = self._context.new_cdp_session(self._page)
            session.send("Network.enable")
            session.send("Network.setBlockedURLs", {"urls": _BLOCKED_URLS})
        except Exception as exc:      # не критично: без блокировки просто медленнее
            self._log(f"Не удалось отключить сторонние счётчики: {_short(exc)}")
        return self.description

    def close(self) -> None:
        for resource, method in ((self._context, "close"), (self._browser, "close"),
                                 (self._playwright, "stop")):
            try:
                if resource is not None:
                    getattr(resource, method)()
            except Exception:
                pass
        self._playwright = self._browser = self._context = self._page = None

    # ------------------------------------------------------------------ проверка одного ИИН
    def check(self, registry: Registry, iin: str) -> PageShot:
        from playwright.sync_api import Error as PlaywrightError
        from playwright.sync_api import TimeoutError as PlaywrightTimeout

        if self._page is None:
            raise BrowserStartError("Браузер не запущен.")
        settings, page = self.settings, self._page
        url = registry.url(iin, settings.base_url)
        widened = False
        self._pace()
        try:
            try:
                response = page.goto(url, wait_until="load", timeout=settings.nav_timeout_s * 1000)
            except PlaywrightTimeout as exc:
                raise CaptureError(
                    f"сайт не ответил за {settings.nav_timeout_s:.0f} с") from exc
            if response is None:
                raise CaptureError("сайт не вернул ответ")
            if response.status != 200:
                raise CaptureError(f"сайт ответил кодом {response.status}")
            self._wait_ready(page)

            info = page.evaluate(_INFO_JS, registry.param)
            count, rows = verify_page(info, registry, iin, settings.max_plausible_records)

            width = settings.viewport_width
            for _ in range(4):
                overflow = int(info.get("overflow") or 0)
                if overflow <= 0 or width >= settings.max_viewport_width:
                    break
                width = min(settings.max_viewport_width, width + overflow + 40)
                page.set_viewport_size({"width": width, "height": settings.viewport_height})
                widened = True
                self._wait_ready(page)
                info = page.evaluate(_INFO_JS, registry.param)
            clipped = int(info.get("overflow") or 0) > 0

            checked_at = datetime.now().astimezone()
            png = page.screenshot(full_page=True, type="png", animations="disabled", caret="hide")
            if looks_blank(png):
                raise CaptureError("снимок получился пустым")
            # Страница не должна была смениться, пока делался снимок.
            if urlsplit(page.url).path.rstrip("/") != registry.path.rstrip("/"):
                raise CaptureError("страница сменилась во время снимка")
            return PageShot(registry, iin, url, count, rows, png, checked_at, clipped)
        except PlaywrightTimeout as exc:
            raise CaptureError("страница не успела загрузиться") from exc
        except PlaywrightError as exc:
            raise CaptureError(f"нет связи с сайтом ({_short(exc)})") from exc
        finally:
            if widened:
                try:
                    page.set_viewport_size({"width": settings.viewport_width,
                                            "height": settings.viewport_height})
                except Exception:
                    pass

    def _wait_ready(self, page) -> None:
        page.wait_for_function(_READY_JS, timeout=self.settings.ready_timeout_s * 1000)
        page.evaluate(_SETTLE_JS)


_CHANNEL_NAMES = {"msedge": "Microsoft Edge", "chrome": "Google Chrome",
                  "chromium": "встроенный Chromium"}
_NOT_INSTALLED = ("is not found", "doesn't exist", "does not exist")


def launch_plan(settings: Settings, prefs: dict) -> list[tuple[str, bool]]:
    """В каком порядке пробовать браузеры: (канал, скрытый ли режим).

    Сначала — что успешно сработало в прошлый раз. Затем каждый браузер в скрытом режиме и,
    если скрытый режим на компьютере запрещён политикой организации, — обычным окном.
    У встроенного Chromium видимого окна нет: в комплекте только его «безоконная» сборка.
    """
    plan: list[tuple[str, bool]] = []
    for channel in settings.browser_channels:
        plan.append((channel, settings.headless))
        if settings.headless and channel != "chromium":
            plan.append((channel, False))
    remembered = (prefs.get("browser_channel"), prefs.get("browser_headless"))
    if remembered in plan:
        plan.remove(remembered)
        plan.insert(0, remembered)
    return plan


def explain_launch_failure(problems: list[tuple[str, str]]) -> str:
    """Понятное объяснение, почему ни один браузер не запустился."""
    details = "\n".join(f"{_CHANNEL_NAMES.get(channel, channel)}: {text}"
                        for channel, text in dict(problems).items())
    if all(any(mark in text for mark in _NOT_INSTALLED) for _, text in problems):
        return ("Не найден браузер Microsoft Edge или Google Chrome.\n"
                "Установите один из них и запустите проверку снова.\n\n" + details)
    return ("Браузер на компьютере есть, но программе не удалось его запустить.\n\n"
            "На рабочих компьютерах так бывает, когда управление браузером запрещено настройками "
            "безопасности организации. Обратитесь в ИТ-поддержку (нужно разрешение на «удалённую "
            "отладку» Edge/Chrome — политика RemoteDebuggingAllowed) либо возьмите комплект "
            "программы со встроенным браузером.\n\n" + details)


def _short(exc: BaseException) -> str:
    """Первая содержательная строка сообщения (у Playwright они многострочные)."""
    for line in str(exc).splitlines():
        line = line.strip()
        if line and not line.startswith(("Call log", "===", "-")):
            return line[:200]
    return type(exc).__name__
