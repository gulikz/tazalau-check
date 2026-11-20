"""Запуск всего дня из командной строки (для Планировщика заданий Windows или cron):

    python -m tazalau_daily --work "D:\\Tazalau\\работа"                       # всё: gov.kz, ADW*.xlsx из папки, tazalau, итог
    python -m tazalau_daily --work ... --adw-sql adw_problem.sql             # ADW запросом (строка подключения в ADW_DSN)
    python -m tazalau_daily --work ... --scope dgd --minutes 120             # tazalau только для клиентов из ДГД, до 2 часов
    python -m tazalau_daily --work ... --no-gov                              # без gov.kz: только то, что уже в папке ДГД
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .run import Config, run_daily


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="python -m tazalau_daily", description="Ежедневная проверка: ДГД (gov.kz), ADW, tazalau, итоговый Excel.")
    p.add_argument("--work", required=True, help="рабочая папка (ADW, ДГД/, снимки, итог)")
    p.add_argument("--adw", help="файл выгрузки ADW (по умолчанию самый свежий ADW*.xlsx/csv в рабочей папке)")
    p.add_argument("--adw-sql", help="SQL-файл к ADW (строка подключения в переменной ADW_DSN)")
    p.add_argument("--no-gov", action="store_true", help="не собирать gov.kz, читать папку ДГД как есть")
    p.add_argument("--regions", nargs="*", help="только эти регионы (kgd-mangistau …)")
    p.add_argument("--scope", choices=("all", "dgd"), default="all", help="кого проверять на tazalau")
    p.add_argument("--limit", type=int, help="не больше стольких ИИН на tazalau")
    p.add_argument("--minutes", type=float, help="не дольше стольких минут на tazalau")
    p.add_argument("--proxy", help="прокси, например http://proxy.bank.local:8080")
    p.add_argument("--show-browser", action="store_true", help="показать окно браузера")
    a = p.parse_args(argv)
    work = Path(a.work).expanduser()
    work.mkdir(parents=True, exist_ok=True)
    log_path = work / "tazalau_daily.log"

    def log(text: str) -> None:
        line = f"{datetime.now():%d.%m.%Y %H:%M:%S} {text}"
        print(line, flush=True)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")

    cfg = Config(work_dir=work, adw_file=Path(a.adw) if a.adw else None, adw_sql=Path(a.adw_sql) if a.adw_sql else None,
                 collect_dgd=not a.no_gov, regions=a.regions, check_scope=a.scope, check_limit=a.limit, check_minutes=a.minutes,
                 proxy=a.proxy, headless=not a.show_browser, log=log)
    try:
        out = run_daily(cfg)
    except Exception as exc:
        log(f"ОСТАНОВЛЕНО: {exc}")
        return 1
    log(f"Готово: {out['report']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
