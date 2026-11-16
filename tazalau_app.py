"""Tazalau Check — точка запуска.

    python tazalau_app.py                 окно программы
    python tazalau_app.py --help          работа из командной строки

Коллегам программа запускается через pythonw.exe — без окна консоли. Поэтому любая ошибка
при старте показывается отдельным окном и записывается в файл: иначе при сбое «просто ничего
не происходит», и понять причину невозможно.
"""
import os
import sys
import traceback

ERROR_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tazalau_startup_error.log")


def _report_fatal(text):
    """Сообщить об ошибке запуска всеми доступными способами."""
    try:
        with open(ERROR_LOG, "w", encoding="utf-8") as handle:
            handle.write(text)
        where = "\n\nПодробности записаны в файл:\n" + ERROR_LOG
    except OSError:
        where = ""
    message = "Программа не смогла запуститься.\n\n" + text.strip().splitlines()[-1] + where
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, message, "Tazalau Check", 0x10)
            return
        except Exception:
            pass
    try:
        sys.stderr.write(text + "\n")
    except Exception:
        pass


def run():
    if sys.version_info < (3, 9):
        _report_fatal("Нужен Python 3.9 или новее, а запущен %d.%d." % sys.version_info[:2])
        return 1
    try:
        from tazalau.cli import main
        return main()
    except SystemExit as exit_request:        # --help, --version и ошибки в аргументах
        code = exit_request.code
        return code if isinstance(code, int) else (0 if code is None else 1)
    except KeyboardInterrupt:
        return 130
    except BaseException:
        _report_fatal(traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(run())
