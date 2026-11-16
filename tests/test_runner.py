"""Ход проверки целиком — без сети: сайт заменён FakeBrowser."""
from __future__ import annotations

import csv
import io
import sys
import threading
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook
from PIL import Image

import tazalau.runner as runner_module
from tazalau.capture import BrowserStartError, PageShot, RateLimitedError
from tazalau.config import BACKUP_DIR_NAME, JOURNAL_FILE_NAME, LINK_HEADER, RESULT_HEADER
from tazalau.excel_io import ExcelBook, ExcelLockedError
from tazalau.registries import BY_KEY
from tazalau.runner import RunOptions, Runner, analyze, describe, estimate_minutes, plan_from_tasks

from conftest import IIN_A, IIN_B, IIN_C, IIN_D, FakeBrowser, make_png

LINK_COL, RESULT_COL = 4, 5


def run(book_path, settings, browser=None, *, out_dir=None, stop_event=None, **options):
    browser = browser if browser is not None else FakeBrowser()
    logs: list[tuple[str, str]] = []
    progress: list[tuple[int, int]] = []
    opts = RunOptions(xlsx=book_path, out_dir=out_dir or book_path.parent / "снимки", **options)
    runner = Runner(opts, settings, browser_factory=lambda: browser, stop_event=stop_event,
                    on_log=lambda level, text: logs.append((level, text)),
                    on_progress=lambda done, total, summary: progress.append((done, total)))
    summary = runner.run()
    return summary, browser, logs, progress


def sheet(path):
    return load_workbook(path)["восстановление"]


def links(path) -> dict[int, str]:
    ws = sheet(path)
    return {row: ws.cell(row, LINK_COL).value for row in range(2, 9) if ws.cell(row, LINK_COL).hyperlink}


def results(path) -> dict[int, str]:
    ws = sheet(path)
    return {row: ws.cell(row, RESULT_COL).value for row in range(2, 9) if ws.cell(row, RESULT_COL).value}


# ------------------------------------------------------------------ обычный запуск
def test_full_run(book_path, settings):
    summary, browser, logs, progress = run(book_path, settings, FakeBrowser(found={IIN_A: 1}))

    assert summary.fatal is None and not summary.stopped and summary.aborted is None
    assert (summary.iins_total, summary.iins_done) == (4, 4)
    assert (summary.found, summary.not_found, summary.errors, summary.invalid) == (1, 3, 0, 1)
    assert summary.rows_linked == 5 and summary.workbook == book_path
    # Каждый ИИН запрошен один раз, хотя IIN_A стоит в двух строках.
    assert browser.calls == [("recovery", IIN_A), ("recovery", IIN_B), ("recovery", IIN_C), ("recovery", IIN_D)]
    assert browser.entered and browser.exited
    assert progress[0] == (0, 4) and progress[-1] == (4, 4)

    written = links(book_path)
    assert sorted(written) == [2, 3, 4, 7, 8]
    assert written[2] == written[4]                               # дубль ссылается на тот же снимок
    for text in written.values():
        assert Path(text).is_file() and Path(text).parent == book_path.parent / "снимки"
    assert Path(written[3]).name.startswith(IIN_B + "_")          # ведущие нули восстановлены

    texts = results(book_path)
    assert texts[2].startswith("НАЙДЕН: Восстановление платежеспособности — записей: 1 · проверено ")
    assert texts[3].startswith("НЕ НАЙДЕН: Восстановление платежеспособности · проверено ")
    assert texts[5] == "ОШИБКА: контрольная цифра ИИН не сходится (вероятно, опечатка)"
    assert 5 not in written and 6 not in texts
    ws = sheet(book_path)
    assert (ws.cell(1, LINK_COL).value, ws.cell(1, RESULT_COL).value) == (LINK_HEADER, RESULT_HEADER)

    assert summary.backup.parent.name == BACKUP_DIR_NAME and summary.backup.is_file()
    assert ExcelBook.open(summary.backup, settings).use_sheet().link_col is None   # копия — «до»


def test_screenshot_is_stamped_image_of_the_page(book_path, settings):
    run(book_path, settings, FakeBrowser(found={IIN_A: 2}))
    with Image.open(links(book_path)[2]) as image:
        page = Image.open(io.BytesIO(make_png()))
        assert image.width == page.width and image.height > page.height     # сверху добавлена полоса
        strip = image.height - page.height
        assert image.getpixel((5, strip + 5)) == (20, 30, 40)               # сама страница не сдвинута вбок
        assert image.getpixel((2, 2)) == (36, 52, 71)                       # цвет полосы-подписи


def test_journal_lists_every_iin(book_path, settings):
    summary, *_ = run(book_path, settings, FakeBrowser(found={IIN_A: 1}))
    assert summary.journal.name == JOURNAL_FILE_NAME
    with open(summary.journal, encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle, delimiter=";"))
    assert rows[0][:5] == ["Дата и время", "Файл Excel", "Лист", "Строки", "ИИН"]
    by_iin = {row[4]: row for row in rows[1:]}
    assert by_iin[IIN_A][3] == "2, 4" and by_iin[IIN_A][5] == "НАЙДЕН" and by_iin[IIN_A][6] == "recovery=1"
    assert by_iin[IIN_B][5] == "НЕ НАЙДЕН"
    assert any(row[5] == "НЕВЕРНЫЙ ИИН" and row[3] == "5" for row in rows[1:])


# ------------------------------------------------------------------ повторные запуски
def test_second_run_does_nothing_and_does_not_touch_the_file(book_path, settings):
    run(book_path, settings)
    before = book_path.read_bytes()
    backups_before = sorted((book_path.parent / BACKUP_DIR_NAME).iterdir())

    summary, browser, logs, _ = run(book_path, settings)
    assert browser.calls == [] and not browser.entered
    assert summary.fatal is None and summary.iins_total == 0 and summary.invalid == 1
    assert book_path.read_bytes() == before                       # файл не переписывался
    assert sorted((book_path.parent / BACKUP_DIR_NAME).iterdir()) == backups_before


def test_fully_processed_book_reports_nothing_to_do(book_path, settings):
    wb = load_workbook(book_path)
    wb.active.delete_rows(5)                                       # убираем строку с опечаткой
    wb.save(book_path)
    run(book_path, settings)
    summary, browser, *_ = run(book_path, settings)
    assert browser.calls == [] and "ссылка на снимок уже записана" in summary.text()


def test_empty_and_fully_hidden_columns_are_explained(tmp_path, settings):
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.append(["ФИО", "ИИН/БИН"])
    empty = tmp_path / "пусто.xlsx"
    wb.save(empty)
    summary, browser, *_ = run(empty, settings)
    assert "нет ни одного значения" in summary.text() and not browser.entered
    assert not (tmp_path / BACKUP_DIR_NAME).exists()               # файл не трогали

    wb.active.append(["x", IIN_A])
    wb.active.row_dimensions[2].hidden = True
    hidden = tmp_path / "скрыто.xlsx"
    wb.save(hidden)
    summary, browser, *_ = run(hidden, settings, only_visible=True)
    assert "скрыты фильтром" in summary.text() and not browser.entered


def test_recheck_processes_everything_again_with_new_files(book_path, settings):
    run(book_path, settings)
    first = links(book_path)
    summary, browser, *_ = run(book_path, settings, FakeBrowser(found={IIN_B: 3}), recheck=True)
    assert len(browser.calls) == 4 and summary.found == 1
    second = links(book_path)
    assert sorted(second) == sorted(first)
    assert all(Path(first[row]).is_file() for row in first)        # старые снимки не удаляются
    assert "записей: 3" in results(book_path)[3]


def test_only_visible_rows(book_path, settings):
    wb = load_workbook(book_path)
    wb.active.row_dimensions[3].hidden = True                      # IIN_B
    wb.active.row_dimensions[8].hidden = True                      # IIN_D
    wb.save(book_path)
    summary, browser, *_ = run(book_path, settings, only_visible=True)
    assert [iin for _, iin in browser.calls] == [IIN_A, IIN_C]
    assert sorted(links(book_path)) == [2, 4, 7]


def test_without_result_column(book_path, settings):
    run(book_path, settings, write_result=False)
    ws = sheet(book_path)
    assert ws.max_column == LINK_COL                    # столбец результата не создавался
    assert ws.cell(1, LINK_COL).value == LINK_HEADER


def test_sheet_can_be_chosen(book_path, settings):
    wb = load_workbook(book_path)
    extra = wb.create_sheet("карты")
    extra.append(["ИИН/БИН"])
    extra.append([IIN_D])
    wb.save(book_path)
    summary, browser, *_ = run(book_path, settings, sheet="карты")
    assert browser.calls == [("recovery", IIN_D)]
    assert load_workbook(book_path)["карты"].cell(2, 2).hyperlink is not None
    assert links(book_path) == {}                                  # другой лист не тронут


# ------------------------------------------------------------------ ошибки сайта
def test_site_error_never_becomes_not_found(book_path, settings):
    summary, browser, logs, _ = run(book_path, settings, FakeBrowser(failures={IIN_B: -1}))
    assert (summary.errors, summary.not_found, summary.found) == (1, 3, 0)
    assert 3 not in links(book_path)                               # ссылки нет -> будет повтор
    assert results(book_path)[3] == ("ОШИБКА: Восстановление платежеспособности: "
                                     "сайт не ответил за 45 с")
    assert browser.calls.count(("recovery", IIN_B)) == settings.retries
    assert not any(p.name.startswith(IIN_B) for p in (book_path.parent / "снимки").iterdir())

    summary, browser, *_ = run(book_path, settings)                # сайт ожил
    assert browser.calls == [("recovery", IIN_B)] and summary.not_found == 1
    assert 3 in links(book_path) and results(book_path)[3].startswith("НЕ НАЙДЕН")


def test_retry_then_success(book_path, settings):
    summary, browser, logs, _ = run(book_path, settings, FakeBrowser(failures={IIN_B: 2}))
    assert summary.errors == 0 and summary.not_found == 4
    assert browser.calls.count(("recovery", IIN_B)) == 3
    assert sum("Повтор через" in text for _, text in logs) == 2


def test_run_aborts_when_the_site_is_down(book_path, settings):
    settings.max_consecutive_failures = 2
    down = FakeBrowser(failures={IIN_A: -1, IIN_B: -1, IIN_C: -1, IIN_D: -1})
    summary, browser, *_ = run(book_path, settings, down)
    assert summary.aborted and "Сайт не отвечает" in summary.aborted
    assert (summary.iins_done, summary.errors) == (2, 2)
    assert len(browser.calls) == 2 * settings.retries              # дальше сайт не дёргаем
    assert links(book_path) == {}


def test_failure_counter_resets_after_a_success(book_path, settings):
    settings.max_consecutive_failures = 2
    assert type(settings)().max_consecutive_failures == 3          # значение по умолчанию
    summary, *_ = run(book_path, settings, FakeBrowser(failures={IIN_A: -1, IIN_C: -1}))
    assert summary.aborted is None and (summary.errors, summary.not_found) == (2, 2)


class LimitedBrowser(FakeBrowser):
    """Сайт отвечает «Превышен лимит запросов» заданное число раз на один ИИН."""

    def __init__(self, iin: str, hits: int):
        super().__init__()
        self.limited_iin, self.hits, self.slowdowns = iin, hits, 0

    def check(self, registry, iin):
        if iin == self.limited_iin and self.hits:
            self.hits -= 1
            self.calls.append((registry.key, iin))
            raise RateLimitedError("сайт сообщил «Превышен лимит запросов»")
        return super().check(registry, iin)

    def slow_down(self) -> float:
        self.slowdowns += 1
        return 3.0


def test_rate_limit_waits_slows_down_and_recovers(book_path, settings):
    settings.retries = 1                                   # обычных повторов нет вовсе
    summary, browser, logs, _ = run(book_path, settings, LimitedBrowser(IIN_B, hits=3))
    assert (summary.errors, summary.not_found) == (0, 4)   # лимит не «съел» попытку и не стал ошибкой
    assert browser.slowdowns == 3
    assert browser.calls.count(("recovery", IIN_B)) == 4
    warnings = [text for level, text in logs if "Превышен лимит запросов" in text]
    assert len(warnings) == 3 and "один запрос в 3.0 с" in warnings[0]
    assert 3 in links(book_path)


def test_rate_limit_that_never_clears_is_an_error_not_a_result(book_path, settings):
    settings.rate_limit_retries = 2
    summary, browser, *_ = run(book_path, settings, LimitedBrowser(IIN_B, hits=10**6))
    assert summary.errors == 1 and 3 not in links(book_path)
    assert browser.calls.count(("recovery", IIN_B)) == 3   # первая попытка + 2 повтора
    assert "лимит запросов" in results(book_path)[3].lower()


def test_browser_that_cannot_start_is_fatal_but_harmless(book_path, settings):
    def broken():
        raise BrowserStartError("Не найден браузер Microsoft Edge или Google Chrome.")

    opts = RunOptions(xlsx=book_path, out_dir=book_path.parent / "снимки")
    summary = Runner(opts, settings, browser_factory=broken).run()
    assert summary.fatal.startswith("Не найден браузер")
    assert links(book_path) == {} and load_workbook(book_path) is not None


def test_image_cannot_be_saved_stops_the_run(book_path, settings, monkeypatch):
    def no_space(image, path):
        raise OSError(28, "No space left on device")

    monkeypatch.setattr(runner_module, "save_png", no_space)
    summary, browser, *_ = run(book_path, settings)
    assert "Не удалось сохранить снимок" in summary.fatal
    assert len(browser.calls) == 1 and links(book_path) == {}


# ------------------------------------------------------------------ остановка
def test_stop_keeps_finished_work(book_path, settings):
    stop = threading.Event()
    seen: list[str] = []

    def on_check(registry, iin):
        seen.append(iin)
        if len(seen) == 2:
            stop.set()                                             # «Стоп» во время второго ИИН

    summary, browser, *_ = run(book_path, settings, FakeBrowser(on_check=on_check), stop_event=stop)
    assert summary.stopped and summary.iins_done == 2
    assert sorted(links(book_path)) == [2, 3, 4]                   # текущий ИИН доведён до конца
    assert len(browser.calls) == 2

    summary, browser, *_ = run(book_path, settings)                # продолжение с места остановки
    assert [iin for _, iin in browser.calls] == [IIN_C, IIN_D]


def test_stop_before_start_does_not_open_browser(book_path, settings):
    stop = threading.Event()
    stop.set()
    summary, browser, *_ = run(book_path, settings, stop_event=stop)
    assert summary.stopped and not browser.entered and browser.calls == []


# ------------------------------------------------------------------ файл занят
def test_copy_mode(book_path, settings):
    original = book_path.read_bytes()
    summary, *_ = run(book_path, settings, save_copy=True)
    assert book_path.read_bytes() == original
    assert summary.workbook != book_path and "_tazalau_" in summary.workbook.name
    assert summary.backup is None and not (book_path.parent / BACKUP_DIR_NAME).exists()
    assert sorted(links(summary.workbook)) == [2, 3, 4, 7, 8]


@pytest.mark.skipif(sys.platform.startswith("win"), reason="в Windows занятость определяется блокировкой")
def test_file_open_in_excel_is_refused_before_any_work(book_path, settings):
    (book_path.parent / f"~${book_path.name}").write_bytes(b"")
    summary, browser, *_ = run(book_path, settings)
    assert "открыт в Excel" in summary.fatal and not browser.entered
    assert not (book_path.parent / BACKUP_DIR_NAME).exists()


def test_file_opened_in_excel_during_the_run_goes_to_a_rescue_file(book_path, settings, monkeypatch):
    """Пользователь открыл файл в Excel посреди проверки: результат не должен пропасть."""
    real_save = ExcelBook.save
    calls = {"n": 0}

    def flaky_save(self):
        calls["n"] += 1
        if calls["n"] > 1 and self.save_path == self.path:
            raise ExcelLockedError("Файл занят — скорее всего, открыт в Excel. Закройте его.")
        return real_save(self)

    monkeypatch.setattr(ExcelBook, "save", flaky_save)
    monkeypatch.setattr(runner_module, "FINAL_SAVE_WAIT_S", 0)
    monkeypatch.setattr(runner_module.time, "sleep", lambda seconds: None)
    summary, _, logs, _ = run(book_path, settings)

    assert summary.fatal is None and summary.rows_linked == 5
    assert summary.workbook != book_path and "_tazalau_unsaved_" in summary.workbook.name
    assert any("результат сохранён отдельно" in note for note in summary.notes)
    assert sorted(links(summary.workbook)) == [2, 3, 4, 7, 8]
    assert links(book_path) == {}                                  # исходный остался как был
    assert any("сохраню позже" in text for _, text in logs)        # промежуточное сохранение отложено


# ------------------------------------------------------------------ проверки до запуска
@pytest.mark.skipif(sys.platform.startswith("win"), reason="права доступа POSIX")
def test_read_only_file_and_folder_are_explained(book_path, settings):
    book_path.chmod(0o444)
    try:
        summary, browser, *_ = run(book_path, settings)
        assert "только для чтения" in summary.fatal and not browser.entered
        summary, browser, *_ = run(book_path, settings, save_copy=True)     # в копию — можно
        assert summary.fatal is None and summary.rows_linked == 5
    finally:
        book_path.chmod(0o644)

    folder = book_path.parent
    for leftover in folder.glob("*_tazalau_*.xlsx"):
        leftover.unlink()
    out_dir = folder.parent / "снимки отдельно"
    folder.chmod(0o555)                                            # в папку с файлом писать нельзя
    try:
        summary, browser, *_ = run(book_path, settings, out_dir=out_dir)
        assert "резервную копию" in summary.fatal and not browser.entered
    finally:
        folder.chmod(0o755)
    assert links(book_path) == {}                                  # без копии файл не тронут


def test_folder_with_hash_is_rejected(book_path, settings):
    summary, browser, *_ = run(book_path, settings, out_dir=book_path.parent / "снимки #1")
    assert "«#»" in summary.fatal and not browser.entered


def test_no_registry_and_unknown_registry(book_path, settings):
    assert "ни один реестр" in run(book_path, settings, registries=())[0].fatal
    assert "Неизвестный реестр" in run(book_path, settings, registries=("нет такого",))[0].fatal


def test_book_without_iin_column(tmp_path, settings):
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.append(["ФИО"])
    path = tmp_path / "x.xlsx"
    wb.save(path)
    summary, browser, *_ = run(path, settings)
    assert "ИИН/БИН" in summary.fatal and not browser.entered


def test_journal_problem_does_not_stop_the_run(book_path, settings):
    out_dir = book_path.parent / "снимки"
    (out_dir / JOURNAL_FILE_NAME).mkdir(parents=True)              # вместо файла — папка: запись невозможна
    summary, *_ = run(book_path, settings)
    assert summary.fatal is None and summary.rows_linked == 5
    assert summary.journal.name.startswith("tazalau_journal_") and summary.journal.is_file()


# ------------------------------------------------------------------ несколько реестров
def test_several_registries_make_one_image_per_iin(book_path, settings):
    keys = ("recovery", "debtor_judicial", "bankrupt_judicial")
    found = {IIN_A: {"recovery": 1, "debtor_judicial": 2}}
    summary, browser, *_ = run(book_path, settings, FakeBrowser(found=found), registries=keys)
    assert (summary.found, summary.not_found) == (1, 3)
    # Порядок обращения — как в меню сайта, а не как перечислил пользователь.
    assert [key for key, iin in browser.calls if iin == IIN_A] == [
        "debtor_judicial", "bankrupt_judicial", "recovery"]
    texts = results(book_path)
    assert texts[2].startswith("НАЙДЕН: Должники: судебная процедура — записей: 2; "
                               "Восстановление платежеспособности — записей: 1 · проверено")
    assert texts[3].startswith("НЕ НАЙДЕН ни в одном из 3 реестров · проверено")
    single = Image.open(io.BytesIO(make_png()))
    with Image.open(links(book_path)[2]) as image:
        assert image.height > 3 * single.height and image.width == single.width
    assert len(list((book_path.parent / "снимки").glob("*.png"))) == 4     # один файл на ИИН


def test_failure_in_one_registry_fails_the_whole_iin(book_path, settings):
    class HalfBroken(FakeBrowser):
        def check(self, registry, iin):
            if registry.key == "recovery" and iin == IIN_A:
                self.calls.append((registry.key, iin))
                from tazalau.capture import CaptureError
                raise CaptureError("сайт ответил кодом 503")
            return super().check(registry, iin)

    summary, *_ = run(book_path, settings, HalfBroken(), registries=("debtor_judicial", "recovery"))
    assert summary.errors == 1 and 2 not in links(book_path)       # половинчатый снимок не пишется
    assert "код" in results(book_path)[2]


# ------------------------------------------------------------------ мелкие функции
def _shot(key: str, count: int, clipped: bool = False, rows: int | None = None) -> PageShot:
    registry = BY_KEY[key]
    return PageShot(registry, IIN_A, registry.url(IIN_A), count, count if rows is None else rows,
                    make_png(), datetime(2026, 11, 15, 21, 30).astimezone(), clipped)


def test_describe():
    assert describe([_shot("recovery", 1)]) == (
        True, "НАЙДЕН: Восстановление платежеспособности — записей: 1 · проверено 15.11.2026 21:30")
    assert describe([_shot("recovery", 0)]) == (
        False, "НЕ НАЙДЕН: Восстановление платежеспособности · проверено 15.11.2026 21:30")
    assert describe([_shot("recovery", 0), _shot("debtor_judicial", 0)])[1].startswith(
        "НЕ НАЙДЕН ни в одном из 2 реестров")
    assert "таблица шире снимка" in describe([_shot("recovery", 1, clipped=True)])[1]
    assert "записей: 20 (на снимке первые 15)" in describe([_shot("recovery", 20, rows=15)])[1]


def test_plan_and_estimate(book_path, settings):
    info = analyze(book_path, None, settings)
    assert info.sheets == ["восстановление"] and info.risks == [] and not info.locked
    options = RunOptions(xlsx=book_path, out_dir=book_path.parent)
    plan = plan_from_tasks(info.layout, info.tasks, info.risks, options)
    assert (plan.rows_with_value, plan.rows_to_check, len(plan.groups), len(plan.invalid)) == (6, 5, 4, 1)
    assert list(plan.groups) == [IIN_A, IIN_B, IIN_C, IIN_D] and plan.groups[IIN_A] == [2, 4]
    assert estimate_minutes(600, 1, settings) == pytest.approx(10.0)       # 600 ИИН без паузы ≈ 10 мин


def test_file_names_never_collide(book_path, settings, tmp_path):
    runner = Runner(RunOptions(xlsx=book_path, out_dir=tmp_path), settings)
    first = runner._free_name("x")
    first.write_bytes(b"1")
    second = runner._free_name("x")
    second.write_bytes(b"2")
    assert (first.name, second.name, runner._free_name("x").name) == ("x.png", "x_2.png", "x_3.png")
