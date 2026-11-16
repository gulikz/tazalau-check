"""Проверка против настоящего сайта: ``pytest -m live``.

Это «датчик изменений сайта»: если tazalau.qoldau.kz поменяет адреса, имена полей или вёрстку,
эти тесты упадут раньше, чем коллеги получат неверные снимки.

На сайт отправляется только заведомо несуществующий ИИН из двенадцати нулей. Чтобы проверить
ещё и случай «найден», задайте переменную окружения TAZALAU_TEST_IIN (ИИН, который точно есть
в реестре «Восстановление платежеспособности»).
"""
from __future__ import annotations

import io
import os
import socket
import threading

import pytest
from PIL import Image

from tazalau.capture import CaptureError, TazalauBrowser, looks_blank
from tazalau.config import Settings
from tazalau.registries import BY_KEY, REGISTRIES

pytestmark = pytest.mark.live

ZERO = "000000000000"


@pytest.fixture(scope="module")
def browser():
    with TazalauBrowser(Settings()) as instance:
        yield instance


@pytest.mark.parametrize("registry", REGISTRIES, ids=[r.key for r in REGISTRIES])
def test_every_registry_answers_and_is_recognised(browser, registry):
    shot = browser.check(registry, ZERO)
    assert shot.count == 0 and shot.rows_on_page == 0 and not shot.clipped
    assert shot.url == f"https://tazalau.qoldau.kz{registry.path}?{registry.param}={ZERO}"
    assert not looks_blank(shot.png)
    with Image.open(io.BytesIO(shot.png)) as image:
        assert image.width == 1920 and image.height >= 1000
    assert shot.checked_at.tzinfo is not None


def test_repeated_navigation_never_gives_a_blank_image(browser):
    """Сайт на 0,4 с прячет страницу после загрузки; раньше 5 снимков из 6 выходили пустыми."""
    for _ in range(6):
        shot = browser.check(BY_KEY["recovery"], ZERO)
        with Image.open(io.BytesIO(shot.png)) as image:
            colors = image.convert("RGB").getcolors(maxcolors=1_000_000)
        assert colors is None or len(colors) > 500


def test_viewport_is_restored_after_each_page(browser):
    browser.check(BY_KEY["bankrupt_extrajudicial"], ZERO)
    assert browser._page.viewport_size == {"width": 1920, "height": 1000}


@pytest.mark.skipif(not os.environ.get("TAZALAU_TEST_IIN"), reason="не задан TAZALAU_TEST_IIN")
def test_known_iin_is_found(browser):
    shot = browser.check(BY_KEY["recovery"], os.environ["TAZALAU_TEST_IIN"])
    assert shot.count >= 1 and 1 <= shot.rows_on_page <= shot.count


def _check_in_own_thread(settings: Settings):
    """Отдельный браузер в отдельном потоке (как в программе: один браузер на рабочий поток).
    Возвращает исключение, которым закончилась проверка, или None."""
    outcome: list = [None]

    def work():
        try:
            with TazalauBrowser(settings) as other:
                other.check(BY_KEY["recovery"], ZERO)
        except Exception as exc:                        # noqa: BLE001 — исключение и есть результат
            outcome[0] = exc

    thread = threading.Thread(target=work)
    thread.start()
    thread.join(timeout=120)
    assert not thread.is_alive()
    return outcome[0]


def test_visible_browser_window_gives_the_same_result():
    """Запасной режим для компьютеров, где скрытый запуск браузера запрещён политикой: браузер
    открывается обычным окном (на пару секунд появится на экране). Снимок должен быть таким же."""
    outcome: list = []

    def work():
        try:
            with TazalauBrowser(Settings(headless=False)) as visible:
                outcome.append(visible.check(BY_KEY["recovery"], ZERO))
        except Exception as exc:                        # noqa: BLE001
            outcome.append(exc)

    thread = threading.Thread(target=work)
    thread.start()
    thread.join(timeout=120)
    shot = outcome[0]
    assert not isinstance(shot, Exception), shot
    assert shot.count == 0 and not shot.clipped and not looks_blank(shot.png)
    with Image.open(io.BytesIO(shot.png)) as image:
        # В видимом окне полоса прокрутки сайта (10 px) занимает место, поэтому снимок чуть уже.
        assert 1890 <= image.width <= 1920 and image.height >= 1000


def test_unreachable_site_is_an_error_not_a_result():
    with socket.socket() as probe:                      # свободный порт, на котором никто не слушает
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]
    error = _check_in_own_thread(Settings(base_url=f"http://127.0.0.1:{port}", nav_timeout_s=8))
    assert isinstance(error, CaptureError) and "нет связи с сайтом" in str(error)


def test_wrong_page_is_an_error_not_a_result():
    """Сайт отвечает, но страница не та (здесь — несуществующий раздел): ошибка, а не «не найден»."""
    error = _check_in_own_thread(Settings(base_url="https://tazalau.qoldau.kz/no-such-section"))
    assert isinstance(error, CaptureError), error
