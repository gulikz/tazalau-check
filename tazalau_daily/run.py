"""Весь день целиком: ДГД с gov.kz → свод и проверка регионов → ADW → сверка → tazalau → итоговый Excel.
Каждый шаг: отдельная функция (ноутбук вызывает их по одной и показывает промежуточные таблицы);
run_daily() вызывает все подряд (запуск по расписанию: python -m tazalau_daily)."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from . import VERSION
from .adw import Adw, latest_export
from .adw import load as load_adw_source
from .checks import check_iins, in_worker, plan, read_journal
from .dgd import DGD_FILE, read_dgd, write_dgd_book
from .govkz import GovClient, harvest_all
from .match import match
from .report import write_report


@dataclass
class Config:
    work_dir: Path                               # рабочая папка дня: ADW, ДГД/, снимки, итог (персональные данные!)
    year: str | None = None                      # год публикаций ДГД; None: текущий
    regions: list[str] | None = None             # только эти регионы (имена или kgd-…); None: все 20
    collect_dgd: bool = True                     # False: не ходить на gov.kz, читать то, что уже в папке ДГД
    use_ocr: bool = True                         # распознавать сканы, если установлен Tesseract
    gov_spacing: float = 1.5                     # секунд между запросами к gov.kz
    proxy: str | None = None                     # прокси банка, например "http://proxy.bank.local:8080"
    ca_file: str | None = None                   # сертификат прокси банка (если он подменяет HTTPS)
    adw_file: Path | None = None                 # выгрузка ADW; None: самая свежая ADW*.xlsx/csv в рабочей папке
    adw_sql: Path | None = None                  # или SQL-файл к PostgreSQL (строка подключения в ADW_DSN)
    adw_dsn_env: str = "ADW_DSN"
    check_scope: str = "all"                     # "dgd": только клиенты из публикаций ДГД; "all": все (ДГД первыми)
    check_limit: int | None = None               # не больше стольких ИИН за запуск
    check_minutes: float | None = None           # не дольше стольких минут на tazalau
    recheck_days: int = 1                        # ИИН, проверенный без ошибки раньше, чем столько дней назад, пропускается
    headless: bool = True                        # False: видно окно браузера
    log: Callable[[str], None] = field(default=print, repr=False)

    def __post_init__(self) -> None:
        self.work_dir = Path(self.work_dir).expanduser()
        self.work_dir.mkdir(parents=True, exist_ok=True)


def step_collect(cfg: Config) -> list[dict]:
    """Шаг 1. Публикации ДГД со всех департаментов на gov.kz в <рабочая папка>/ДГД/<регион>/."""
    client = GovClient(spacing=cfg.gov_spacing, proxy=cfg.proxy, ca_file=cfg.ca_file, log=cfg.log)
    return harvest_all(cfg.work_dir, year=cfg.year, regions=cfg.regions, client=client, log=cfg.log)


def step_dgd(cfg: Config) -> dict | None:
    """Шаг 1б. Все файлы папки ДГД: ИИН, проверка 20 регионов, «ДГД_свод.xlsx»."""
    dg = read_dgd(cfg.work_dir, use_ocr=cfg.use_ocr, log=cfg.log)
    if dg is None:
        cfg.log("Папки ДГД нет: сначала шаг 1 (сбор с gov.kz) или положите файлы ДГД в папку «ДГД».")
        return None
    path = write_dgd_book(dg, cfg.work_dir)
    bad = [r for r in dg["regions"] if r["level"] == "bad"]
    warn = [r for r in dg["regions"] if r["level"] == "warn"]
    cfg.log(f"ДГД: ИИН {len(dg['by_iin'])}, строк объявлений {len(dg['announcements'])}, публикаций {len(dg['orders'])}, "
            f"файлов вручную {len(dg['files'])}. Регионы: красных {len(bad)}, жёлтых {len(warn)}. Свод: {path}")
    return dg


def step_adw(cfg: Config) -> Adw:
    """Шаг 2. Клиенты с признаком «проблема»: выгрузка или SQL."""
    if cfg.adw_sql:
        dsn = os.environ.get(cfg.adw_dsn_env)
        if not dsn:
            raise RuntimeError(f"Строка подключения к ADW не задана: переменная окружения {cfg.adw_dsn_env}.")
        adw = load_adw_source(sql=cfg.adw_sql, dsn=dsn)
    else:
        source = cfg.adw_file or latest_export(cfg.work_dir)
        if not source:
            raise FileNotFoundError(f"В папке {cfg.work_dir} нет выгрузки ADW (файл ADW*.xlsx или ADW*.csv).")
        adw = load_adw_source(source)
    ch = adw.checks
    cfg.log(f"ADW ({adw.source}): строк {ch['строк']}, ИИН {ch['ИИН']}, неверных ИИН {ch['неверных ИИН']}, без ФИО {ch['без ФИО']}, "
            f"ИИН с разными ФИО {ch['ИИН с разными ФИО']}.")
    return adw


def step_match(cfg: Config, adw: Adw, dg: dict | None) -> dict:
    """Шаг 3. Клиенты в публикациях ДГД."""
    m = match(adw, dg)
    cfg.log(f"Сверка: в публикациях ДГД по ИИН {len(m['exact'])} клиент(ов), возможно по ФИО {len(m['maybe'])}; "
            f"ИИН из ДГД не клиенты банка: {m['dgd_not_clients']}.")
    return m


def step_tazalau(cfg: Config, adw: Adw, m: dict, *, settings=None) -> tuple[dict, dict]:
    """Шаг 4. Снимки и статусы на tazalau. Возвращает (результаты этого запуска, журнал всех проверок)."""
    from tazalau.config import Settings

    settings = settings or Settings()
    settings.headless = cfg.headless
    if cfg.proxy:
        settings.proxy_server = cfg.proxy
    journal = read_journal(cfg.work_dir)
    priority = list(m["exact"]) + list(m["maybe"])
    clients = adw.clients if cfg.check_scope == "all" else {i: adw.clients[i] for i in priority}
    due = plan(clients, priority, journal, recheck_days=cfg.recheck_days)
    queue = due[:cfg.check_limit] if cfg.check_limit else due
    cfg.log(f"tazalau: в очереди {len(queue)} ИИН (из ДГД первыми: {sum(1 for i in queue if i in set(priority))}); "
            f"уже проверены недавно: {len(clients) - len(due)}"
            + (f"; отложены до следующего запуска (лимит {cfg.check_limit}): {len(due) - len(queue)}" if len(due) > len(queue) else "") + ".")
    results = in_worker(check_iins, queue, adw.clients, cfg.work_dir, settings=settings, log=cfg.log, minutes=cfg.check_minutes)
    return results, read_journal(cfg.work_dir)


def step_report(cfg: Config, adw: Adw, dg: dict | None, m: dict, results: dict, journal: dict, extra: dict | None = None) -> Path:
    """Шаг 5. Итоговый Excel."""
    found = sum(1 for r in results.values() if r.get("found"))
    errors = sum(1 for r in results.values() if r.get("error"))
    summary = {
        "Дата": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "Версия": f"tazalau_daily {VERSION}",
        "Рабочая папка": str(cfg.work_dir),
        "ADW: источник": adw.source,
        "ADW: клиентов (ИИН)": len(adw.clients),
        "ADW: неверных ИИН": len(adw.invalid),
        "ДГД: ИИН в публикациях": len(dg["by_iin"]) if dg else 0,
        "ДГД: строк объявлений о возбуждении дел": len(dg["announcements"]) if dg else 0,
        "ДГД: красных регионов": sum(1 for r in dg["regions"] if r["level"] == "bad") if dg else "",
        "ДГД: жёлтых регионов": sum(1 for r in dg["regions"] if r["level"] == "warn") if dg else "",
        "Клиентов в публикациях ДГД (по ИИН)": len(m["exact"]),
        "Клиентов в публикациях ДГД (возможно, по ФИО)": len(m["maybe"]),
        "tazalau: проверено в этом запуске": len(results),
        "tazalau: найдено в реестрах (в этом запуске)": found,
        "tazalau: ошибок (в этом запуске)": errors,
        "tazalau: клиентов без единой проверки": sum(1 for i in adw.clients if i not in journal),
        "Свод ДГД": str(cfg.work_dir / DGD_FILE),
    }
    summary.update(extra or {})
    path = cfg.work_dir / f"Итог_{date.today().isoformat()}.xlsx"
    write_report(path, cfg.work_dir, adw, dg, m, results, journal, summary)
    cfg.log(f"Итог: {path}")
    return path


def run_daily(cfg: Config) -> dict:
    """Все шаги подряд. Ошибка одного региона gov.kz не останавливает день; без ADW итог не строится."""
    out: dict = {"started": datetime.now()}
    if cfg.collect_dgd:
        out["collect"] = step_collect(cfg)
    out["dgd"] = step_dgd(cfg)
    out["adw"] = step_adw(cfg)
    out["match"] = step_match(cfg, out["adw"], out["dgd"])
    out["results"], out["journal"] = step_tazalau(cfg, out["adw"], out["match"])
    out["report"] = step_report(cfg, out["adw"], out["dgd"], out["match"], out["results"], out["journal"])
    out["finished"] = datetime.now()
    return out
