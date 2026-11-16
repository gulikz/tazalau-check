"""Окно программы: ``pytest -m gui`` (нужен экран; окна на несколько секунд появляются на нём).

Сайт заменён FakeBrowser, диалоги — записью их текста, поэтому тест проходит весь путь
пользователя: выбрать файл -> увидеть, что будет сделано -> «Старт» -> итог -> повторный запуск.

Окно показывается по-настоящему, как у пользователя. Это не прихоть: в macOS запись в журнал
окна, которое ни разу не было показано, заставляет Tk бесконечно перерисовывать его внутри
``update()`` — тест зависает, хотя у видимого (и у свёрнутого) окна такого нет.
"""
from __future__ import annotations

import gc
import threading
import time
import zipfile
from pathlib import Path

import pytest

tk = pytest.importorskip("tkinter", reason="в этой установке Python нет tkinter")

from openpyxl import Workbook, load_workbook

import tazalau.config as config_module
from tazalau.gui import App

from conftest import IIN_A, FakeBrowser

pytestmark = pytest.mark.gui


def pump(root: tk.Tk, condition, timeout: float = 15.0) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        root.update()
        if condition():
            return
        time.sleep(0.01)
    raise AssertionError("окно не дождалось ожидаемого состояния")


@pytest.fixture(scope="session")
def tk_root():
    """Один корень Tk на все тесты: несколько корней подряд в одном процессе Tk переносит плохо."""
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("нет экрана")
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def window(tk_root, tmp_path, settings, monkeypatch):
    monkeypatch.setattr(config_module, "PREFS_FILE", tmp_path / "prefs.json")   # не трогаем настройки пользователя
    root = tk.Toplevel(tk_root)
    browser = FakeBrowser(found={IIN_A: 1})
    app = App(root, settings, browser_factory=lambda: browser)
    app.dialogs = []
    app.answer = True
    app.ask = lambda title, text: (app.dialogs.append(("ask", title, text)), app.answer)[1]
    app.info = lambda title, text: app.dialogs.append(("info", title, text))
    app.warn = lambda title, text: app.dialogs.append(("warn", title, text))
    app.error = lambda title, text: app.dialogs.append(("error", title, text))
    app.browser = browser
    app.base = tk_root
    yield app
    release(app)
    tk_root.update()


def release(app: App) -> None:
    """Закрывает окно и освобождает переменные Tk здесь, в главном потоке. Иначе сборщик мусора
    может добраться до них в рабочем потоке следующего теста, и тот на секунды застрянет."""
    if app.root.winfo_exists():
        app.root.destroy()
    for name, value in list(vars(app).items()):
        if isinstance(value, tk.Variable) or (
                isinstance(value, dict) and any(isinstance(v, tk.Variable) for v in value.values())):
            setattr(app, name, None)
    gc.collect()


def choose(app: App, path: Path) -> None:
    app._set_xlsx(path)
    pump(app.root, lambda: app.analysis is not None or "Читаю" not in app.var_info.get())


def finish(app: App) -> None:
    pump(app.root, lambda: app.worker is None and any(kind != "ask" for kind, *_ in app.dialogs))


def test_choose_file_shows_what_will_be_done(window, book_path):
    assert window.var_info.get() == "Выберите файл Excel."
    choose(window, book_path)
    info = window.var_info.get()
    assert "Столбец «ИИН/БИН»: B (заголовки в строке 1). Строк с ИИН: 6." in info
    assert "К проверке: 5 строк, разных ИИН — 4." in info and "с неверным ИИН — 1" in info
    assert window.var_sheet.get() == "восстановление"
    assert window.var_out.get() == str(book_path.parent / "tazalau_screenshots")
    assert str(window.cmb_sheet.cget("state")) == "readonly"
    # По умолчанию отмечен только реестр с фотографий.
    assert window._selected_registries() == ["recovery"]
    assert not window.var_copy.get() and window.var_result.get() and not window.var_recheck.get()


def test_full_run_from_the_window(window, book_path):
    choose(window, book_path)
    window._start()
    kind, title, question = window.dialogs[0]
    assert kind == "ask" and "Будет проверено ИИН: 4 (строк: 5)" in question
    assert "резервная копия" in question and "меньше минуты" in question
    from tazalau.gui import _duration
    assert (_duration(45), _duration(125), _duration(200 * 60)) == ("45 с", "2 мин 05 с", "3 ч 20 мин")
    assert str(window.btn_start.cget("state")) == "disabled" and str(window.btn_stop.cget("state")) == "normal"
    finish(window)

    kind, title, text = window.dialogs[-1]
    assert kind == "warn" and title == "Проверка завершена с замечаниями"      # в файле есть неверный ИИН
    assert "Проверено ИИН: 4 из 4" in text and "найдено в реестрах: 1" in text
    assert "Ссылок записано в Excel: 5" in text
    assert window.progress["value"] == 4 and window.var_status.get() == "Проверка завершена."
    assert str(window.btn_start.cget("state")) == "normal" and str(window.btn_stop.cget("state")) == "disabled"
    log = window.txt_log.get("1.0", "end")
    assert f"[1/4] ИИН {IIN_A}: НАЙДЕН" in log and "Резервная копия" in log

    ws = load_workbook(book_path)["восстановление"]
    assert ws.cell(2, 4).hyperlink is not None and Path(ws.cell(2, 4).value).is_file()

    # После запуска окно само перечитывает файл: проверять больше нечего.
    pump(window.root, lambda: window.analysis is not None)
    assert "К проверке: 0 строк" in window.var_info.get() and "уже со ссылкой — 5" in window.var_info.get()
    calls = len(window.browser.calls)
    window._start()
    finish(window)
    assert len(window.browser.calls) == calls                                # сайт больше не запрашивался


def test_nothing_to_do_is_said_plainly(window, book_path):
    wb = load_workbook(book_path)
    wb.active.delete_rows(5)
    wb.save(book_path)
    choose(window, book_path)
    window._start()
    finish(window)
    window.dialogs.clear()
    pump(window.root, lambda: window.analysis is not None)
    window._start()
    assert window.dialogs[-1][:2] == ("info", "Нечего проверять") and window.worker is None


def test_declining_the_question_starts_nothing(window, book_path):
    choose(window, book_path)
    window.answer = False
    window._start()
    assert window.worker is None and window.browser.calls == []
    assert load_workbook(book_path).active.max_column == 3                   # файл не тронут


def test_start_without_file_or_registry(window, book_path):
    window._start()
    assert window.dialogs[-1][:2] == ("warn", "Нет файла")
    choose(window, book_path)
    for var in window.var_registry.values():
        var.set(False)
    window._start()
    assert window.dialogs[-1][:2] == ("warn", "Не выбран реестр") and window.worker is None


def test_stop_button(window, book_path):
    gate = threading.Event()
    window.browser.on_check = lambda registry, iin: gate.wait(10)             # «сайт думает»
    choose(window, book_path)
    window._start()
    pump(window.root, lambda: len(window.browser.calls) == 1)
    window._stop()
    assert "Останавливаю" in window.var_status.get()
    gate.set()
    finish(window)
    assert "Проверка остановлена." in window.dialogs[-1][2]
    assert len(window.browser.calls) == 1                                     # текущий ИИН доведён, дальше не пошли
    assert load_workbook(book_path)["восстановление"].cell(2, 4).hyperlink is not None


def test_closing_the_window_during_a_run_saves_first(window, book_path):
    gate = threading.Event()
    window.browser.on_check = lambda registry, iin: gate.wait(10)
    choose(window, book_path)
    window._start()
    pump(window.root, lambda: len(window.browser.calls) == 1)
    window._on_close()
    assert window.dialogs[-1][0] == "ask" and "Остановить её, сохранить" in window.dialogs[-1][2]
    gate.set()
    pump(window.base, lambda: not window.root.winfo_exists())                 # окно закрылось само
    assert load_workbook(book_path)["восстановление"].cell(2, 4).hyperlink is not None


def test_risky_workbook_switches_to_copy_mode(window, book_path):
    original = book_path.read_bytes()
    with zipfile.ZipFile(book_path, "a") as archive:
        archive.writestr("xl/pivotTables/pivotTable1.xml", "<x/>")
    original = book_path.read_bytes()
    choose(window, book_path)
    assert window.var_copy.get()
    assert "⚠ В файле есть: сводные таблицы" in window.var_info.get()
    window._start()
    assert "в КОПИЮ файла" in window.dialogs[0][2]
    finish(window)
    assert book_path.read_bytes() == original                                 # исходный файл не изменён
    copy_path = Path(window.var_xlsx.get())
    assert copy_path != book_path and "_tazalau_" in copy_path.name           # окно переключилось на копию
    assert load_workbook(copy_path)["восстановление"].cell(2, 4).hyperlink is not None
    pump(window.root, lambda: window.analysis is not None)
    assert not window.var_copy.get() and "⚠" not in window.var_info.get()     # «копию копии» не предлагаем


def test_unusable_file_is_explained_in_the_window(window, tmp_path):
    wb = Workbook()
    wb.active.append(["ФИО", "Счёт"])
    path = tmp_path / "без иин.xlsx"
    wb.save(path)
    choose(window, path)
    assert window.analysis is None and "не найден столбец с заголовком «ИИН/БИН»" in window.var_info.get()
    window._start()
    assert window.dialogs[-1][:2] == ("warn", "Файл не прочитан")


def test_options_change_the_plan_shown(window, book_path):
    wb = load_workbook(book_path)
    wb.active.row_dimensions[3].hidden = True
    wb.save(book_path)
    choose(window, book_path)
    window.var_visible.set(True)
    window._refresh_info()
    assert "К проверке: 4 строк, разных ИИН — 3." in window.var_info.get()
    assert "скрыто фильтром — 1" in window.var_info.get()


def test_settings_are_remembered_between_launches(window, book_path, settings):
    choose(window, book_path)
    window.var_registry["debtor_judicial"].set(True)
    window._start()
    finish(window)
    top = tk.Toplevel(window.base)
    second = App(top, settings, browser_factory=lambda: FakeBrowser())
    try:
        assert second.var_xlsx.get() == str(book_path)
        assert second._selected_registries() == ["debtor_judicial", "recovery"]
        pump(top, lambda: second.analysis is not None)
        assert "К проверке: 0 строк" in second.var_info.get()               # вчерашняя работа видна сразу
    finally:
        release(second)


def test_window_survives_an_internal_error_in_message_handling(window, book_path):
    """Сбой при обработке одного сообщения не должен «заморозить» окно навсегда."""
    choose(window, book_path)
    real = window.progress.configure
    calls = {"n": 0}

    def flaky(**options):
        calls["n"] += 1
        if calls["n"] == 2:
            raise ValueError("сбой отрисовки")
        return real(**options)

    window.progress.configure = flaky
    window._start()
    finish(window)
    assert "Проверка завершена." in window.dialogs[-1][2]
    assert "Внутренняя ошибка окна: ValueError: сбой отрисовки" in window.txt_log.get("1.0", "end")
    assert str(window.btn_start.cget("state")) == "normal"


def test_crash_in_the_worker_thread_is_reported(window, book_path, monkeypatch):
    import tazalau.gui as gui_module

    def exploding_runner(*args, **kwargs):
        raise MemoryError("не хватило памяти")

    monkeypatch.setattr(gui_module, "Runner", exploding_runner)
    choose(window, book_path)
    window._start()
    finish(window)
    assert window.dialogs[-1][0] == "error" and "MemoryError" in window.dialogs[-1][2]
    assert window.worker is None and str(window.btn_start.cget("state")) == "normal"


def _clipped(app: App, widget=None) -> list[str]:
    """Элементы, которые видны не полностью: им выделено меньше места, чем нужно содержимому, либо
    они вылезли за край окна. Журнал не в счёт: это единственная «резиновая» часть окна."""
    root = app.root
    widget = widget or root
    flexible = app.txt_log.master
    bottom, right = root.winfo_rooty() + root.winfo_height(), root.winfo_rootx() + root.winfo_width()
    problems = []
    for child in widget.winfo_children():
        if child is flexible:
            continue
        if child.winfo_manager() and child.winfo_viewable() and child is not flexible.master:
            name = f"{child.winfo_class()} «{str(child.cget('text'))[:25]}»" if "text" in child.keys() \
                else child.winfo_class()
            if (child.winfo_width() + 1 < child.winfo_reqwidth()
                    or child.winfo_height() + 1 < child.winfo_reqheight()):
                problems.append(f"{name}: дано {child.winfo_width()}x{child.winfo_height()}, "
                                f"нужно {child.winfo_reqwidth()}x{child.winfo_reqheight()}")
            elif (child.winfo_rooty() + child.winfo_height() > bottom + 1
                  or child.winfo_rootx() + child.winfo_width() > right + 1):
                problems.append(f"{name}: выходит за край окна")
        problems.extend(_clipped(app, child))
    return problems


def _settle(root) -> None:
    for _ in range(25):
        root.update()
        time.sleep(0.02)


def _longest_info(window, book_path) -> None:
    """Самый длинный текст сведений о файле: книга «с рисками» и к тому же открыта в Excel."""
    (book_path.parent / f"~${book_path.name}").write_bytes(b"")
    with zipfile.ZipFile(book_path, "a") as archive:
        archive.writestr("xl/pivotTables/pivotTable1.xml", "<x/>")
        archive.writestr("xl/slicers/slicer1.xml", "<x/>")
    choose(window, book_path)
    _settle(window.root)
    assert "⚠ В файле есть" in window.var_info.get() and "открыт в Excel" in window.var_info.get()


def test_nothing_is_cut_off_in_the_window(window, book_path):
    root = window.root
    _longest_info(window, book_path)
    initial = (root.winfo_width(), root.winfo_height())
    assert _clipped(window) == []
    assert window.txt_log.winfo_height() >= 60                            # журналу хватает места
    assert root.winfo_height() <= root.winfo_screenheight() - 100
    # Самое маленькое окно, до которого можно сжать: всё, кроме журнала, по-прежнему видно целиком.
    min_w, min_h = root.minsize()
    root.geometry(f"{min_w}x{min_h}")
    _settle(root)
    assert (root.winfo_width(), root.winfo_height()) == (min_w, min_h)
    assert _clipped(window) == []
    assert window.txt_log.winfo_height() >= 15
    print(f"\n   окно: начальное {initial[0]}x{initial[1]}, минимальное {min_w}x{min_h}, "
          f"экран {root.winfo_screenwidth()}x{root.winfo_screenheight()}")


def test_window_fits_a_small_laptop_screen(window, book_path):
    """Экран 1366×768 — обычный офисный ноутбук. Окно должно поместиться целиком, а журнал остаться виден."""
    root = window.root
    window._fit_to_screen(1366, 768)
    _settle(root)
    assert root.winfo_height() <= 768 - 110 and root.winfo_width() <= 1366 - 40
    choose(window, book_path)                                             # обычные сведения, две строки
    _settle(root)
    typical_log = window.txt_log.winfo_height()
    assert _clipped(window) == [] and typical_log >= 40
    _longest_info(window, book_path)
    worst_log = window.txt_log.winfo_height()
    print(f"\n   экран 1366x768: окно {root.winfo_width()}x{root.winfo_height()}, журнал "
          f"{typical_log} px обычно и {worst_log} px при самом длинном тексте сведений")
    assert _clipped(window) == []
