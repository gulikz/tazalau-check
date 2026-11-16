"""Настройки программы.

Обычному пользователю менять ничего не нужно. Разработчик может положить рядом с программой
файл ``tazalau_config.json`` и переопределить любое поле ``Settings`` (например, задержку
между запросами или адрес прокси).
"""
from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field, fields
from pathlib import Path

from .registries import BASE_URL

CONFIG_FILE_NAME = "tazalau_config.json"
PREFS_FILE = Path.home() / ".tazalau_check.json"

LINK_HEADER = "tazalau_file_location"
RESULT_HEADER = "tazalau_result"
SCREENSHOT_DIR_NAME = "tazalau_screenshots"
BACKUP_DIR_NAME = "tazalau_backup"
JOURNAL_FILE_NAME = "tazalau_journal.csv"
DEBUG_LOG_NAME = "tazalau_debug.log"


def _default_channels() -> list[str]:
    # Используется браузер, уже установленный в системе: ничего скачивать не требуется.
    if sys.platform.startswith("win"):
        return ["msedge", "chrome", "chromium"]
    return ["chrome", "msedge", "chromium"]


@dataclass
class Settings:
    base_url: str = BASE_URL

    # --- браузер и снимок -------------------------------------------------------------
    browser_channels: list[str] = field(default_factory=_default_channels)
    headless: bool = True
    viewport_width: int = 1920          # при 1600 таблицы пяти реестров из шести обрезаются справа
    viewport_height: int = 1000
    max_viewport_width: int = 3840      # предел авто-расширения, если таблица всё же не помещается
    device_scale_factor: float = 1.0
    stamp: bool = True                  # полоса над снимком: ИИН, реестр, дата/время, адрес
    proxy_server: str | None = None     # например "http://proxy.bank.local:8080"
    user_agent: str | None = None
    browser_launch_timeout_s: float = 30.0   # сколько ждать запуска одного браузера

    # --- вежливость к сайту и устойчивость ------------------------------------------------
    # Сайт допускает примерно один запрос в секунду с одного адреса (замерено 15.11.2026:
    # при интервале 1,0 с отказов нет, при 0,85 с — каждый седьмой запрос получает страницу
    # «Превышен лимит запросов»). Весь офис обычно выходит в интернет с одного адреса, поэтому
    # программа берёт только половину лимита, оставляя место коллегам, работающим вручную.
    min_request_interval_s: float = 2.0     # не чаще одного запроса за это время
    max_request_interval_s: float = 15.0    # до какой паузы программа сама замедляется
    rate_limit_pause_s: float = 20.0        # сколько ждать после «Превышен лимит запросов»
    rate_limit_retries: int = 5             # сколько раз ждать и повторять один и тот же ИИН
    nav_timeout_s: float = 45.0
    ready_timeout_s: float = 20.0
    retries: int = 3                    # попыток на одну страницу
    retry_backoff_s: list[float] = field(default_factory=lambda: [3.0, 10.0, 30.0])
    max_consecutive_failures: int = 3   # столько ИИН подряд с ошибкой — сайт недоступен, стоп
    max_plausible_records: int = 50     # больше записей на один ИИН = фильтр не применился

    # --- Excel ----------------------------------------------------------------------------
    save_every: int = 20                # сохранять файл каждые N проверенных ИИН
    header_scan_rows: int = 30          # в скольких верхних строках искать заголовок «ИИН/БИН»
    link_mode: str = "absolute"         # "absolute" — полный путь; "relative" — относительно файла
    allow_bad_checksum: bool = False    # True — отправлять на сайт ИИН с неверной контрольной цифрой

    def validate(self) -> None:
        if self.link_mode not in ("absolute", "relative"):
            raise ValueError("link_mode: допустимо 'absolute' или 'relative'")
        if self.viewport_width < 1200:
            raise ValueError("viewport_width: не меньше 1200, иначе сайт прячет боковое меню")
        if self.retries < 1 or self.save_every < 1:
            raise ValueError("retries и save_every должны быть не меньше 1")
        if self.min_request_interval_s < 0:
            raise ValueError("min_request_interval_s не может быть отрицательным")
        if self.max_request_interval_s < self.min_request_interval_s:
            raise ValueError("max_request_interval_s меньше min_request_interval_s")
        if not self.browser_channels:
            raise ValueError("browser_channels: список браузеров пуст")


def app_dir() -> Path:
    """Папка программы: рядом с .exe в собранном виде, иначе — корень проекта."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def asset_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "tazalau" / "assets" / name


def load_settings(path: Path | None = None) -> Settings:
    """Настройки по умолчанию + необязательный tazalau_config.json. Ошибка в файле — исключение
    с понятным текстом: молча игнорировать опечатку в настройке хуже, чем остановиться."""
    settings = Settings()
    path = path or (app_dir() / CONFIG_FILE_NAME)
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError) as exc:
            raise ValueError(f"Не удалось прочитать {path.name}: {exc}") from exc
        if not isinstance(data, dict):
            raise ValueError(f"{path.name}: ожидается объект JSON с настройками")
        known = {f.name: f for f in fields(Settings)}
        for key, value in data.items():
            if key.startswith("_"):
                continue                      # "_comment": "..." — пометки разработчика
            if key not in known:
                raise ValueError(f"{path.name}: неизвестная настройка «{key}»")
            default = getattr(settings, key)
            if default is not None and value is not None and not _same_kind(default, value):
                raise ValueError(f"{path.name}: у настройки «{key}» неверный тип значения")
            setattr(settings, key, value)
    settings.validate()
    return settings


def _same_kind(default: object, value: object) -> bool:
    if isinstance(default, bool):
        return isinstance(value, bool)
    if isinstance(default, (int, float)):
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    return isinstance(value, type(default))


def load_prefs() -> dict:
    """Последние выбранные файл/папка/реестры. Удобство, не более: любая ошибка — пустой словарь."""
    try:
        data = json.loads(PREFS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_prefs(prefs: dict) -> None:
    try:
        PREFS_FILE.write_text(json.dumps(prefs, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def update_prefs(changes: dict) -> None:
    """Дописать отдельные значения, не затирая остальные (окно и браузер хранят своё)."""
    prefs = load_prefs()
    prefs.update(changes)
    save_prefs(prefs)


def bundled_browsers_dir() -> Path | None:
    """Папка ``browsers`` со встроенным браузером Playwright рядом с программой, если она есть.

    Нужна на компьютерах, где управлять установленным Edge/Chrome запрещено политикой
    организации: встроенный Chromium под эти политики не подпадает.
    """
    for folder in (app_dir() / "browsers", app_dir().parent / "browsers"):
        try:
            if folder.is_dir() and any(folder.iterdir()):
                return folder
        except OSError:
            continue
    return None
