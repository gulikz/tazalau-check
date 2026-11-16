"""Проверка страницы по её содержимому, обнаружение пустого снимка и сборка итоговой картинки."""
from __future__ import annotations

import io
from datetime import datetime, timedelta, timezone

import pytest
from PIL import Image

from tazalau.capture import (CaptureError, PageShot, RateLimitedError, TazalauBrowser, looks_blank,
                             verify_page)
from tazalau.config import Settings
from tazalau.registries import BY_KEY, REGISTRIES, resolve
from tazalau.stamp import build_image, save_png, stamp_lines

from conftest import IIN_A, make_png

RECOVERY = BY_KEY["recovery"]


def page_info(**changes) -> dict:
    """Сведения о странице в том виде, в каком их возвращает браузер для «найдена 1 запись»."""
    info = {
        "path": RECOVERY.path, "search": f"?{RECOVERY.param}={IIN_A}",
        "inputFound": True, "inputValue": IIN_A, "totals": [1, 1],
        "tableFound": True, "headerCols": 11, "dataRows": 1, "emptyRow": False,
        "styled": True, "overflow": 0,
    }
    info.update(changes)
    return info


def test_found_and_not_found_pages_are_accepted():
    assert verify_page(page_info(), RECOVERY, IIN_A, 50) == (1, 1)
    assert verify_page(page_info(totals=[0, 0], dataRows=0, emptyRow=True), RECOVERY, IIN_A, 50) == (0, 0)
    assert verify_page(page_info(totals=[20, 20], dataRows=15), RECOVERY, IIN_A, 50) == (20, 15)
    assert verify_page(page_info(path=RECOVERY.path + "/"), RECOVERY, IIN_A, 50) == (1, 1)


@pytest.mark.parametrize("changes,fragment", [
    ({"path": "/ru/offline"}, "другую страницу"),                            # заглушка «нет связи»
    ({"path": "/ru/list/bankruptcy/judicial"}, "другую страницу"),           # не тот реестр
    ({"styled": False}, "без оформления"),
    ({"inputFound": False, "inputValue": None}, "нет поля «ИИН»"),           # сайт изменил вёрстку
    ({"inputValue": ""}, "не применил фильтр"),                               # фильтр проигнорирован
    ({"inputValue": "000000000000"}, "не применил фильтр"),
    ({"search": ""}, "адрес страницы"),
    ({"search": "?flApplicantXin=000000000000"}, "адрес страницы"),
    ({"tableFound": False}, "нет таблицы"),
    ({"totals": []}, "Всего записей"),
    ({"totals": [None, None]}, "Всего записей"),
    ({"totals": [1, 2]}, "Всего записей"),                                   # счётчики расходятся
    ({"totals": [29911, 29911], "dataRows": 15}, "фильтр не сработал"),       # показан весь реестр
    ({"totals": [0, 0], "dataRows": 1}, "не совпадает"),
    ({"totals": [0, 0], "dataRows": 0, "emptyRow": False}, "не совпадает"),
    ({"totals": [1, 1], "dataRows": 0, "emptyRow": True}, "не совпадает"),
    ({"totals": [1, 1], "dataRows": 3}, "не совпадает"),
])
def test_anything_else_is_an_error_not_a_result(changes, fragment):
    with pytest.raises(CaptureError, match=fragment):
        verify_page(page_info(**changes), RECOVERY, IIN_A, 50)


def test_rate_limit_page_is_recognised_as_such():
    """Сайт отдаёт её с кодом 200 по тому же адресу — отличить можно только по содержимому."""
    limit_page = page_info(
        title="Превышен лимит запросов!", inputFound=False, inputValue=None, tableFound=False,
        notice="Превышен лимит запросов! Попробуйте вернуться на страницу позже.",
        styled=False, totals=[], headerCols=0, dataRows=0)
    with pytest.raises(RateLimitedError):
        verify_page(limit_page, RECOVERY, IIN_A, 50)
    # Обычная страница с результатом, где эти слова встретились в тексте, остаётся результатом.
    assert verify_page(page_info(notice="... лимит запросов ..."), RECOVERY, IIN_A, 50) == (1, 1)


def test_requests_are_paced_and_can_slow_down():
    waits: list[float] = []
    browser = TazalauBrowser(Settings(min_request_interval_s=2.0, max_request_interval_s=5.0),
                             sleep=waits.append)
    browser._pace()                                    # первый запрос — без ожидания
    assert waits == []
    browser._pace()                                    # второй — почти полный интервал
    assert len(waits) == 1 and 1.9 < waits[0] <= 2.0
    assert browser.slow_down() == 3.0 and browser.slow_down() == 4.5
    assert browser.slow_down() == 5.0 and browser.slow_down() == 5.0      # не больше предела
    browser._pace()
    assert 4.9 < waits[-1] <= 5.0


def test_blank_screenshot_is_detected():
    blank = io.BytesIO()
    Image.new("RGB", (1920, 1000), (238, 244, 248)).save(blank, format="PNG")   # цвет фона сайта
    assert looks_blank(blank.getvalue())
    assert not looks_blank(make_png())


def test_registry_table_is_consistent():
    assert len(REGISTRIES) == 6 and len({r.key for r in REGISTRIES}) == 6
    assert len({r.path for r in REGISTRIES}) == 6 and len({r.short for r in REGISTRIES}) == 6
    assert RECOVERY.url(IIN_A) == f"https://tazalau.qoldau.kz/ru/list/bankruptcy/recovery?flApplicantXin={IIN_A}"
    assert [r.key for r in resolve(["recovery", "debtor_extrajudicial"])] == ["debtor_extrajudicial", "recovery"]
    with pytest.raises(ValueError):
        resolve(["unknown"])


def _shot(count=1, rows=None, width=640, key="recovery") -> PageShot:
    registry = BY_KEY[key]
    when = datetime(2026, 11, 15, 21, 30, 5, tzinfo=timezone(timedelta(hours=5)))
    return PageShot(registry, IIN_A, registry.url(IIN_A), count, count if rows is None else rows,
                    make_png(width=width), when)


def test_stamp_text():
    line1, line2 = stamp_lines(_shot())
    assert line1 == (f"ИИН {IIN_A}   ·   Банкроты и неплатежеспособные → Восстановление "
                     "платежеспособности   ·   Всего записей: 1")
    assert line2.startswith("Проверено: 15.11.2026 21:30:05 (UTC+05:00)   ·   https://tazalau.qoldau.kz/")
    assert f"flApplicantXin={IIN_A}" in line2
    assert "(на снимке первые 15)" in stamp_lines(_shot(count=20, rows=15))[0]


def test_image_keeps_the_page_untouched_under_the_stamp():
    shot = _shot()
    page = Image.open(io.BytesIO(shot.png)).convert("RGB")
    image = build_image([shot])
    strip = image.height - page.height
    assert strip > 30 and image.width == page.width
    assert image.crop((0, strip, image.width, image.height)).tobytes() == page.tobytes()
    assert build_image([shot], stamp=False).tobytes() == page.tobytes()


def test_stamp_wraps_on_a_narrow_image_instead_of_cutting_the_address():
    wide = build_image([_shot(width=1920)])
    narrow = build_image([_shot(width=420)])
    assert narrow.height - 360 > wide.height - 360                       # подпись заняла больше строк
    strip = narrow.crop((0, 0, 420, narrow.height - 360))
    right_edge = strip.crop((412, 0, 420, strip.height))
    assert right_edge.getextrema() == ((36, 36), (52, 52), (71, 71))     # текст не упирается в край


def test_wrapping_loses_no_characters():
    from PIL import ImageDraw
    from tazalau.stamp import _font, _wrap
    draw = ImageDraw.Draw(Image.new("RGB", (8, 8)))
    font = _font(12)
    for text in stamp_lines(_shot()):
        pieces = _wrap(draw, text, font, 300)
        assert len(pieces) > 1 and all(draw.textlength(piece, font=font) <= 300 for piece in pieces)
        assert "".join(pieces).replace("   ·   ", "") == text.replace("   ·   ", "")


def test_several_pages_of_different_width_are_stacked():
    image = build_image([_shot(width=640), _shot(width=900, key="debtor_judicial")])
    assert image.width == 900 and image.height > 2 * 360
    assert image.getpixel((899, image.height - 1)) == (255, 255, 255)   # нижняя страница целиком
    with pytest.raises(ValueError):
        build_image([])


def test_png_is_written_atomically(tmp_path):
    target = tmp_path / "снимок.png"
    save_png(build_image([_shot()]), target)
    assert Image.open(target).format == "PNG"
    assert [p.name for p in tmp_path.iterdir()] == ["снимок.png"]


# ------------------------------------------------------------------ запуск браузера на «строгих» компьютерах
def test_launch_plan_tries_hidden_then_visible_and_remembers_what_worked():
    from tazalau.capture import launch_plan
    settings = Settings(browser_channels=["msedge", "chrome", "chromium"])
    assert launch_plan(settings, {}) == [
        ("msedge", True), ("msedge", False), ("chrome", True), ("chrome", False), ("chromium", True)]
    # Что сработало в прошлый раз — первым, остальное в прежнем порядке.
    remembered = {"browser_channel": "chrome", "browser_headless": False}
    assert launch_plan(settings, remembered)[:2] == [("chrome", False), ("msedge", True)]
    assert len(launch_plan(settings, remembered)) == 5
    # Устаревшая или чужая запись в настройках ничего не ломает.
    assert launch_plan(settings, {"browser_channel": "firefox", "browser_headless": True})[0] == ("msedge", True)
    # Если разработчик сам включил видимый режим, скрытый не пробуем.
    assert launch_plan(Settings(browser_channels=["msedge"], headless=False), {}) == [("msedge", False)]


def test_launch_failure_is_explained_by_cause():
    from tazalau.capture import explain_launch_failure
    missing = explain_launch_failure([
        ("msedge", "BrowserType.launch: Chromium distribution 'msedge' is not found at C:\\Program Files"),
        ("chromium", "BrowserType.launch: Executable doesn't exist at C:\\app\\browsers\\chrome.exe")])
    assert missing.startswith("Не найден браузер Microsoft Edge или Google Chrome.")
    blocked = explain_launch_failure([
        ("msedge", "BrowserType.launch: Timeout 30000ms exceeded."),
        ("msedge", "BrowserType.launch: Target page, context or browser has been closed"),
        ("chromium", "BrowserType.launch: Executable doesn't exist at C:\\app\\browsers\\chrome.exe")])
    assert blocked.startswith("Браузер на компьютере есть, но программе не удалось его запустить.")
    assert "RemoteDebuggingAllowed" in blocked and "встроенным браузером" in blocked
    assert blocked.count("Microsoft Edge:") == 1                      # по одной строке на браузер


class _FakePlaywright:
    """Подмена Playwright: каждый (браузер, режим) либо запускается, либо падает с заданной ошибкой."""

    def __init__(self, failures):
        self.failures, self.attempts, self.stopped = failures, [], False
        self.chromium = self

    def start(self):
        return self

    def stop(self):
        self.stopped = True

    def launch(self, **options):
        key = (options.get("channel", "chromium"), options["headless"])
        self.attempts.append(key)
        assert options["timeout"] == 30_000
        if key in self.failures:
            raise RuntimeError(self.failures[key])
        return self

    version = "130.0"

    def new_context(self, **options):
        assert options["service_workers"] == "block" and options["locale"] == "ru-RU"
        return self

    def new_page(self):
        return self

    def new_cdp_session(self, page):
        return self

    def send(self, method, params=None):
        return None

    def close(self):
        return None


def _start_with(monkeypatch, failures, **settings_overrides):
    import playwright.sync_api as api
    fake = _FakePlaywright(failures)
    monkeypatch.setattr(api, "sync_playwright", lambda: fake)
    notes: list[str] = []
    browser = TazalauBrowser(Settings(browser_channels=["msedge", "chrome", "chromium"],
                                      **settings_overrides), log=notes.append)
    return browser, fake, notes


def test_headless_forbidden_by_policy_falls_back_to_a_visible_window(monkeypatch):
    from tazalau.config import load_prefs
    browser, fake, notes = _start_with(monkeypatch, {
        ("msedge", True): "BrowserType.launch: Target page, context or browser has been closed"})
    assert browser.start() == "Microsoft Edge 130.0 (в видимом окне)"
    assert fake.attempts == [("msedge", True), ("msedge", False)]
    assert any("Не закрывайте его" in note for note in notes)
    assert load_prefs() == {"browser_channel": "msedge", "browser_headless": False}

    # Следующий запуск начинается сразу с рабочего варианта — без ожидания неудачной попытки.
    browser, fake, notes = _start_with(monkeypatch, {
        ("msedge", True): "BrowserType.launch: Target page, context or browser has been closed"})
    browser.start()
    assert fake.attempts == [("msedge", False)]


def test_browser_control_forbidden_falls_back_to_bundled_chromium(monkeypatch):
    blocked = "BrowserType.launch: Timeout 30000ms exceeded."
    browser, fake, _ = _start_with(monkeypatch, {
        ("msedge", True): blocked, ("msedge", False): blocked,
        ("chrome", True): "BrowserType.launch: Chromium distribution 'chrome' is not found at C:\\...",
        ("chrome", False): "BrowserType.launch: Chromium distribution 'chrome' is not found at C:\\..."})
    assert browser.start() == "встроенный Chromium 130.0"
    assert fake.attempts[-1] == ("chromium", True) and len(fake.attempts) == 5


def test_no_usable_browser_is_reported_and_cleaned_up(monkeypatch):
    from tazalau.capture import BrowserStartError
    from tazalau.config import load_prefs
    blocked = "BrowserType.launch: Timeout 30000ms exceeded."
    failures = {(channel, mode): blocked for channel in ("msedge", "chrome") for mode in (True, False)}
    failures[("chromium", True)] = "BrowserType.launch: Executable doesn't exist at C:\\x\\chrome.exe"
    browser, fake, _ = _start_with(monkeypatch, failures)
    with pytest.raises(BrowserStartError, match="RemoteDebuggingAllowed"):
        browser.start()
    assert fake.stopped and len(fake.attempts) == 5 and load_prefs() == {}


def test_bundled_browser_folder_is_used_when_present(tmp_path, monkeypatch):
    import os

    import tazalau.config as config
    app = tmp_path / "TazalauCheck" / "app"
    (app.parent / "browsers" / "chromium_headless_shell-1200").mkdir(parents=True)
    app.mkdir()
    monkeypatch.setattr(config, "app_dir", lambda: app)
    assert config.bundled_browsers_dir() == app.parent / "browsers"
    monkeypatch.delenv("PLAYWRIGHT_BROWSERS_PATH", raising=False)
    browser, fake, _ = _start_with(monkeypatch, {})
    monkeypatch.setattr("tazalau.capture.bundled_browsers_dir", config.bundled_browsers_dir)
    browser.start()
    assert os.environ["PLAYWRIGHT_BROWSERS_PATH"] == str(app.parent / "browsers")
    monkeypatch.delenv("PLAYWRIGHT_BROWSERS_PATH", raising=False)
    (app.parent / "browsers" / "chromium_headless_shell-1200").rmdir()
    assert config.bundled_browsers_dir() is None                        # пустая папка не считается


def test_preferences_are_merged_not_overwritten():
    from tazalau.config import load_prefs, update_prefs
    update_prefs({"xlsx": "a.xlsx", "registries": ["recovery"]})
    update_prefs({"browser_channel": "msedge"})
    assert load_prefs() == {"xlsx": "a.xlsx", "registries": ["recovery"], "browser_channel": "msedge"}


def test_png_save_survives_a_momentary_lock(tmp_path, monkeypatch):
    import tazalau.fsutil as fsutil
    real_replace, attempts = fsutil.os.replace, []

    def flaky(source, target):
        attempts.append(target)
        if len(attempts) == 1:
            raise PermissionError(13, "файл занят другим процессом")
        return real_replace(source, target)

    monkeypatch.setattr(fsutil.os, "replace", flaky)
    target = tmp_path / "снимок.png"
    save_png(build_image([_shot()]), target)
    assert len(attempts) == 2 and Image.open(target).format == "PNG"
    assert [p.name for p in tmp_path.iterdir()] == ["снимок.png"]


def test_lock_that_never_clears_is_reported_after_all_attempts(tmp_path, monkeypatch):
    import tazalau.fsutil as fsutil
    attempts = []

    def locked(source, target):
        attempts.append(target)
        raise PermissionError(13, "файл занят другим процессом")

    monkeypatch.setattr(fsutil.os, "replace", locked)
    source = tmp_path / "a.tmp"
    source.write_bytes(b"x")
    with pytest.raises(PermissionError):
        fsutil.replace_with_retry(source, tmp_path / "a.png")
    assert len(attempts) == fsutil.RETRY_ATTEMPTS
