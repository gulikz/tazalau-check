"""Общие заготовки для тестов.

Обычные тесты (``pytest``) в сеть не выходят: сайт заменяет FakeBrowser.
Тесты с пометкой ``live`` обращаются к настоящему сайту и запускаются отдельно:
``pytest -m live``.
"""
from __future__ import annotations

import io
import sys
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import tazalau.config as config_module                   # noqa: E402
import tazalau.fsutil as fsutil_module                    # noqa: E402
from tazalau.capture import CaptureError, PageShot      # noqa: E402
from tazalau.config import Settings                      # noqa: E402
from tazalau.iin import check_digit                      # noqa: E402


@pytest.fixture(autouse=True)
def _isolated(tmp_path_factory, monkeypatch):
    """Ни один тест не трогает личные настройки пользователя и не ждёт пауз при повторах."""
    monkeypatch.setattr(config_module, "PREFS_FILE",
                        tmp_path_factory.mktemp("prefs") / "tazalau_prefs.json")
    monkeypatch.setattr(fsutil_module, "RETRY_PAUSE_S", 0.0)


def fake_iin(prefix11: str) -> str:
    """ИИН с верной контрольной цифрой. В тестах используются только префиксы с невозможной
    датой рождения (13-й месяц, 99-е число) — такой номер не может принадлежать человеку."""
    digit = check_digit(prefix11)
    assert digit is not None, prefix11
    return prefix11 + str(digit)


IIN_A = fake_iin("99139990001")
IIN_B = fake_iin("00139990002")      # начинается с «00»: Excel-число теряет два нуля
IIN_C = fake_iin("05139990003")      # начинается с «0»
IIN_D = fake_iin("77139990004")


def make_png(width: int = 640, height: int = 360, text: str = "page") -> bytes:
    image = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, width, 40), fill=(20, 30, 40))
    draw.text((10, 60), text, fill=(0, 0, 0))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


class FakeBrowser:
    """Заменяет настоящий браузер: отвечает по заранее заданной таблице и считает обращения."""

    description = "Fake Browser 1.0"

    def __init__(self, found: dict | None = None, failures: dict | None = None, on_check=None):
        self.found = found or {}            # ИИН -> записей (или {код реестра: записей})
        self.failures = dict(failures or {})  # ИИН -> сколько раз подряд ответить ошибкой (-1 = всегда)
        self.on_check = on_check
        self.calls: list[tuple[str, str]] = []
        self.entered = self.exited = False

    def __enter__(self):
        self.entered = True
        return self

    def __exit__(self, *exc_info):
        self.exited = True

    def check(self, registry, iin: str) -> PageShot:
        self.calls.append((registry.key, iin))
        if self.on_check:
            self.on_check(registry, iin)
        left = self.failures.get(iin, 0)
        if left:
            if left > 0:
                self.failures[iin] = left - 1
            raise CaptureError("сайт не ответил за 45 с")
        count = self.found.get(iin, 0)
        if isinstance(count, dict):
            count = count.get(registry.key, 0)
        return PageShot(registry, iin, registry.url(iin), count, min(count, 15),
                        make_png(text=f"{registry.key} {iin} {count}"), datetime.now().astimezone())


@pytest.fixture
def settings() -> Settings:
    """Настройки без пауз, чтобы тесты шли мгновенно."""
    return Settings(min_request_interval_s=0.0, rate_limit_pause_s=0.0,
                    retry_backoff_s=[0.0], save_every=2)


@pytest.fixture
def book_path(tmp_path: Path) -> Path:
    """Книга с типичными «трудными» значениями в столбце ИИН/БИН."""
    wb = Workbook()
    ws = wb.active
    ws.title = "восстановление"
    ws.append(["ФИО", "ИИН/БИН", "Счёт"])
    ws.append(["первый", IIN_A, "KZ1"])                 # 2  текст
    ws.append(["второй", int(IIN_B), "KZ2"])            # 3  число без двух ведущих нулей
    ws.append(["первый (дубль)", f" {IIN_A} ", "KZ3"])  # 4  тот же ИИН с пробелами
    ws.append(["опечатка", IIN_A[:11] + str((int(IIN_A[11]) + 1) % 10), "KZ4"])  # 5 контрольная цифра
    ws.append([None, None, None])                       # 6  пустая строка
    ws.append(["третий", float(IIN_C), "KZ5"])          # 7  число с плавающей точкой
    ws.append(["четвёртый", IIN_D, "KZ6"])              # 8
    wb.create_sheet("прочее").append(["нет", "нужных", "столбцов"])
    path = tmp_path / "отчет по банкротству.xlsx"
    wb.save(path)
    return path


def pytest_configure(config):
    config.addinivalue_line("markers", "live: обращается к настоящему сайту tazalau.qoldau.kz")
    config.addinivalue_line("markers", "gui: открывает окно программы (нужен экран)")
