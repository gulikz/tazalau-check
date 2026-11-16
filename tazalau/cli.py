"""Запуск: без аргументов — окно программы; с аргументами — работа из командной строки.

Командная строка нужна разработчику (проверка, автоматизация). Коллегам достаточно окна.
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import APP_NAME, __version__
from .config import SCREENSHOT_DIR_NAME, load_settings
from .registries import DEFAULT_KEYS, REGISTRIES


def _parser() -> argparse.ArgumentParser:
    keys = ", ".join(r.key for r in REGISTRIES)
    parser = argparse.ArgumentParser(
        prog="tazalau_app", description=f"{APP_NAME} {__version__} — проверка ИИН по реестрам "
                                        "tazalau.qoldau.kz со снимком страницы.")
    parser.add_argument("--xlsx", type=Path, help="книга Excel со столбцом «ИИН/БИН»")
    parser.add_argument("--out", type=Path, help="папка для снимков (по умолчанию рядом с книгой)")
    parser.add_argument("--sheet", help="имя листа (по умолчанию — лист со столбцом «ИИН/БИН»)")
    parser.add_argument("--registries", default=",".join(DEFAULT_KEYS),
                        help=f"коды реестров через запятую или all. Доступны: {keys}")
    parser.add_argument("--recheck", action="store_true", help="проверить и строки, где ссылка уже есть")
    parser.add_argument("--no-result", action="store_true", help="не писать столбец tazalau_result")
    parser.add_argument("--only-visible", action="store_true", help="пропускать строки, скрытые фильтром")
    parser.add_argument("--copy", action="store_true", help="не менять исходный файл, результат — в копию")
    parser.add_argument("--plan", action="store_true", help="только показать, что будет сделано")
    parser.add_argument("--selftest", action="store_true",
                        help="проверить браузер и связь с сайтом (один запрос) и выйти")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {__version__}")
    return parser


def _registry_keys(text: str) -> tuple[str, ...]:
    if text.strip().lower() == "all":
        return tuple(r.key for r in REGISTRIES)
    return tuple(part.strip() for part in text.split(",") if part.strip())


def selftest(settings) -> int:
    """Запускает браузер и делает один запрос с заведомо несуществующим ИИН из нулей."""
    from .capture import BrowserStartError, CaptureError, TazalauBrowser
    from .registries import BY_KEY

    import platform

    print(f"{APP_NAME} {__version__} — проверка компьютера")
    print(f"Система: {platform.system()} {platform.release()}, Python {platform.python_version()}")
    print("Запускаю браузер (до минуты)…", flush=True)
    try:
        with TazalauBrowser(settings, log=lambda text: print(f"  {text}")) as browser:
            print(f"Браузер: {browser.description}")
            print("Открываю сайт tazalau.qoldau.kz…", flush=True)
            shot = browser.check(BY_KEY["recovery"], "000000000000")
    except BrowserStartError as exc:
        print(f"\nОШИБКА: {exc}")
        print("\nИТОГ: программа НЕ сможет работать на этом компьютере, пока не решён вопрос "
              "с браузером.\nПокажите это окно разработчику или ИТ-поддержке.")
        return 1
    except CaptureError as exc:
        print(f"\nОШИБКА: {exc}")
        print("\nИТОГ: браузер работает, но сайт tazalau.qoldau.kz с этого компьютера открыть "
              "не удалось.\nПроверьте, открывается ли сайт в обычном браузере. Если да — покажите "
              "это окно разработчику.")
        return 1
    print(f"Сайт отвечает, страница распознана (записей: {shot.count}, снимок {len(shot.png)} байт).")
    if shot.count != 0:
        print("\nИТОГ: неожиданный ответ сайта — покажите это окно разработчику.")
        return 1
    print("\nИТОГ: всё в порядке, программой можно пользоваться на этом компьютере.")
    return 0


def run_cli(args, settings) -> int:
    from .runner import RunOptions, Runner, analyze, plan_from_tasks

    # abspath, а не resolve(): в Windows resolve() превращает сетевой диск «W:\» в «\\сервер\папка»,
    # и ссылки в Excel получились бы не такими, к каким привыкли коллеги.
    xlsx = Path(os.path.abspath(args.xlsx.expanduser()))
    out_dir = (Path(os.path.abspath(args.out.expanduser())) if args.out
               else xlsx.parent / SCREENSHOT_DIR_NAME)
    options = RunOptions(xlsx=xlsx, out_dir=out_dir, sheet=args.sheet,
                         registries=_registry_keys(args.registries), recheck=args.recheck,
                         write_result=not args.no_result, only_visible=args.only_visible,
                         save_copy=args.copy)
    if args.plan:
        from .excel_io import ExcelError
        try:
            info = analyze(xlsx, args.sheet, settings)
        except ExcelError as exc:
            print(f"ОШИБКА: {exc}")
            return 1
        plan = plan_from_tasks(info.layout, info.tasks, info.risks, options)
        print(f"Листы со столбцом ИИН: {', '.join(info.sheets)}")
        print(f"Лист «{plan.layout.sheet}»: строк с ИИН {plan.rows_with_value}, к проверке "
              f"{plan.rows_to_check} (разных ИИН {len(plan.groups)}), уже со ссылкой "
              f"{plan.skipped_linked}, скрыто {plan.skipped_hidden}, неверных ИИН {len(plan.invalid)}")
        for task in plan.invalid:
            print(f"  строка {task.row}: {task.parse.error}")
        if plan.risks:
            print("В книге есть элементы, которые могут пострадать при записи: " + ", ".join(plan.risks))
        return 0

    def on_log(level: str, text: str) -> None:
        print(f"{'!' if level in ('warn', 'error') else ' '} {text}", flush=True)

    summary = Runner(options, settings, on_log=on_log).run()
    print("-" * 60)
    print(summary.text())
    if summary.fatal:
        return 1
    return 2 if (summary.errors or summary.aborted) else 0


def main(argv: list[str] | None = None) -> int:
    # В собранной оконной программе под Windows потоков вывода нет вовсе.
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name)
        if stream is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
        elif hasattr(stream, "reconfigure"):
            # Консоль Windows в кодировке cp866/cp1251 не знает некоторых знаков («→», «≈»):
            # пусть вместо падения программы там будет «?».
            try:
                stream.reconfigure(errors="replace")
            except (OSError, ValueError):
                pass
    args = _parser().parse_args(argv)
    try:
        settings = load_settings()
    except ValueError as exc:
        if args.xlsx or args.selftest:
            print(f"ОШИБКА: {exc}")
            return 1
        from .gui import show_fatal
        show_fatal(str(exc))
        return 1
    if args.selftest:
        return selftest(settings)
    if args.xlsx:
        return run_cli(args, settings)
    try:
        from .gui import run_gui
    except ImportError as exc:
        print("ОШИБКА: в этой установке Python нет оконной библиотеки tkinter "
              f"({exc}). Установите Python с сайта python.org либо работайте из командной "
              "строки: tazalau_app.py --help")
        return 1
    return run_gui(settings)
