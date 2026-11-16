"""Ход проверки: книга Excel -> запросы к сайту -> снимки -> ссылки в книге.

Правила, на которых всё держится:
  * Ссылка пишется только после того, как снимок проверен и сохранён на диск. Пустая ячейка
    ссылки означает «не сделано» — при следующем запуске такая строка будет обработана заново.
  * Одинаковые ИИН в разных строках запрашиваются один раз.
  * Ошибка сайта никогда не превращается в «не найден»: либо достоверный снимок, либо ошибка.
  * Файл Excel сохраняется периодически и обязательно в конце, даже при остановке или сбое.
"""
from __future__ import annotations

import csv
import os
import threading
import time
import traceback
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

from .capture import (BrowserStartError, CaptureError, PageShot, RateLimitedError,
                      TazalauBrowser)
from .config import DEBUG_LOG_NAME, JOURNAL_FILE_NAME, Settings
from .excel_io import ExcelBook, ExcelError, ExcelLockedError, Layout, RowTask, is_locked
from .fsutil import remove_quietly
from .registries import DEFAULT_KEYS, Registry, resolve
from .stamp import build_image, save_png

FINAL_SAVE_WAIT_S = 180          # сколько ждать, пока пользователь закроет файл в Excel

_JOURNAL_HEADER = ["Дата и время", "Файл Excel", "Лист", "Строки", "ИИН", "Итог",
                   "Записей по реестрам", "Файл снимка", "Примечание"]


class RunError(Exception):
    """Проверку нельзя начать или продолжить. Текст показывается пользователю."""


class _Stopped(Exception):
    """Пользователь нажал «Стоп» посреди проверки ИИН."""


@dataclass
class RunOptions:
    xlsx: Path
    out_dir: Path
    sheet: str | None = None
    registries: tuple[str, ...] = DEFAULT_KEYS
    recheck: bool = False            # обрабатывать и строки, где ссылка уже есть
    write_result: bool = True        # писать итог в столбец tazalau_result
    only_visible: bool = False       # пропускать строки, скрытые фильтром Excel
    save_copy: bool = False          # не менять исходный файл, результат — в копию


@dataclass
class Plan:
    """Что будет сделано — считается до запуска и показывается пользователю."""

    layout: Layout
    rows_with_value: int
    groups: "OrderedDict[str, list[int]]"        # ИИН -> строки, в порядке первого появления
    invalid: list[RowTask]
    skipped_linked: int
    skipped_hidden: int
    risks: list[str]

    @property
    def rows_to_check(self) -> int:
        return sum(len(rows) for rows in self.groups.values())


@dataclass
class Summary:
    iins_total: int = 0
    iins_done: int = 0
    found: int = 0
    not_found: int = 0
    errors: int = 0
    invalid: int = 0
    rows_linked: int = 0
    stopped: bool = False
    aborted: str | None = None       # проверка прервана программой (сайт недоступен и т. п.)
    fatal: str | None = None         # проверка не состоялась
    workbook: Path | None = None     # куда записан результат
    backup: Path | None = None
    out_dir: Path | None = None
    journal: Path | None = None
    elapsed_s: float = 0.0
    notes: list[str] = field(default_factory=list)

    def text(self) -> str:
        if self.fatal:
            return self.fatal
        head = "Проверка остановлена." if self.stopped else (
            "Проверка прервана." if self.aborted else "Проверка завершена.")
        lines = [head]
        if self.aborted:
            lines.append(self.aborted)
        lines.append(f"Проверено ИИН: {self.iins_done} из {self.iins_total}")
        lines.append(f"  найдено в реестрах: {self.found}")
        lines.append(f"  не найдено: {self.not_found}")
        if self.errors:
            lines.append(f"  не удалось проверить (ошибка сайта): {self.errors}")
        if self.invalid:
            lines.append(f"Строк с неверным ИИН (пропущены): {self.invalid}")
        lines.append(f"Ссылок записано в Excel: {self.rows_linked}")
        if self.workbook:
            lines.append(f"Файл: {self.workbook}")
        lines.extend(self.notes)
        return "\n".join(lines)


@dataclass
class Analysis:
    """Результат быстрого чтения файла — для окна программы и для режима --plan."""

    path: Path
    sheets: list[str]                # листы, где найден столбец «ИИН/БИН»
    layout: Layout
    tasks: list[RowTask]
    risks: list[str]
    locked: bool


def analyze(path: Path, sheet: str | None, settings: Settings) -> Analysis:
    """Открывает книгу только для чтения сведений: что в ней есть и сколько предстоит работы."""
    path = Path(path)
    book = ExcelBook.open(path, settings)
    try:
        sheets = [layout.sheet for layout in book.layouts()]
        layout = book.use_sheet(sheet if sheet in sheets else None)
        return Analysis(path, sheets, layout, book.read_tasks(), list(book.risks), is_locked(path))
    finally:
        book.close()


def make_plan(book: ExcelBook, options: RunOptions) -> Plan:
    layout = book.use_sheet(options.sheet)
    return plan_from_tasks(layout, book.read_tasks(), list(book.risks), options)


def plan_from_tasks(layout: Layout, tasks: list[RowTask], risks: list[str],
                    options: RunOptions) -> Plan:
    groups: "OrderedDict[str, list[int]]" = OrderedDict()
    invalid: list[RowTask] = []
    skipped_linked = skipped_hidden = 0
    for task in tasks:
        if options.only_visible and task.hidden:
            skipped_hidden += 1
            continue
        if task.has_link and not options.recheck:
            skipped_linked += 1
            continue
        if not task.parse.ok:
            invalid.append(task)
            continue
        groups.setdefault(task.parse.iin, []).append(task.row)
    return Plan(layout, len(tasks), groups, invalid, skipped_linked, skipped_hidden, risks)


def nothing_to_do(plan: Plan) -> str:
    """Почему проверять нечего — своими словами, а не общим «готово»."""
    if plan.rows_with_value == 0:
        return "В столбце «ИИН/БИН» нет ни одного значения."
    if plan.skipped_hidden and not plan.skipped_linked:
        return ("Все строки с ИИН скрыты фильтром Excel. Снимите галочку «Только видимые строки» "
                "или измените фильтр в файле.")
    return "Во всех строках ссылка на снимок уже записана — новых ИИН для проверки нет."


def estimate_minutes(iins: int, registries: int, settings: Settings) -> float:
    """Грубая оценка длительности: один запрос за интервал, но не быстрее ~1 с на страницу."""
    return iins * registries * max(1.0, settings.min_request_interval_s) / 60.0


def describe(shots: list[PageShot]) -> tuple[bool, str]:
    """(найден ли, текст для столбца tazalau_result)."""
    found = [shot for shot in shots if shot.count > 0]
    when = shots[-1].checked_at.strftime("%d.%m.%Y %H:%M")
    if found:
        text = "НАЙДЕН: " + "; ".join(
            f"{shot.registry.short} — записей: {shot.count}"
            + (f" (на снимке первые {shot.rows_on_page})" if shot.rows_on_page < shot.count else "")
            for shot in found)
    elif len(shots) == 1:
        text = f"НЕ НАЙДЕН: {shots[0].registry.short}"
    else:
        text = f"НЕ НАЙДЕН ни в одном из {len(shots)} реестров"
    if any(shot.clipped for shot in shots):
        text += " · внимание: таблица шире снимка"
    return bool(found), f"{text} · проверено {when}"


class Runner:
    def __init__(self, options: RunOptions, settings: Settings, *,
                 on_log: Callable[[str, str], None] | None = None,
                 on_progress: Callable[[int, int, Summary], None] | None = None,
                 stop_event: threading.Event | None = None,
                 browser_factory: Callable[[], object] | None = None):
        self.options = options
        self.settings = settings
        self._on_log = on_log or (lambda level, text: None)
        self._on_progress = on_progress or (lambda done, total, summary: None)
        self.stop_event = stop_event or threading.Event()
        self._browser_factory = browser_factory or (
            lambda: TazalauBrowser(settings, log=lambda text: self.log("warn", text),
                                   sleep=self._sleep))
        self._out_dir = Path(options.out_dir)
        self._journal_path: Path | None = None
        self._dirty = False

    # ------------------------------------------------------------------ служебное
    def log(self, level: str, text: str) -> None:
        self._debug(f"[{level}] {text}")
        self._on_log(level, text)

    def _debug(self, text: str) -> None:
        try:
            with open(self._out_dir / DEBUG_LOG_NAME, "a", encoding="utf-8") as handle:
                handle.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {text}\n")
        except OSError:
            pass

    def _sleep(self, seconds: float) -> None:
        if seconds > 0:
            self.stop_event.wait(seconds)

    def _journal(self, layout: Layout, rows: list[int], iin: str, status: str,
                 counts: str = "", file_name: str = "", note: str = "") -> None:
        """Строка журнала. Журнал — вспомогательный: его сбой проверку не останавливает."""
        record = [datetime.now().strftime("%d.%m.%Y %H:%M:%S"), Path(self.options.xlsx).name,
                  layout.sheet, ", ".join(str(r) for r in rows), iin, status, counts, file_name, note]
        for attempt in range(2):
            path = self._journal_path or (self._out_dir / JOURNAL_FILE_NAME)
            try:
                new_file = not path.exists()
                with open(path, "a", encoding="utf-8-sig", newline="") as handle:
                    writer = csv.writer(handle, delimiter=";")
                    if new_file:
                        writer.writerow(_JOURNAL_HEADER)
                    writer.writerow(record)
                self._journal_path = path
                return
            except OSError:
                # Журнал открыт в Excel — пишем в отдельный файл этого запуска.
                if attempt == 0:
                    self._journal_path = self._out_dir / (
                        f"tazalau_journal_{datetime.now():%Y%m%d_%H%M%S}.csv")

    def _prepare_out_dir(self) -> None:
        folder = self._out_dir
        if "#" in str(folder.resolve()):
            raise RunError("В пути к папке для снимков есть символ «#» — Excel не открывает "
                           "ссылки с таким символом. Выберите другую папку.")
        try:
            folder.mkdir(parents=True, exist_ok=True)
            probe = folder / f"~tazalau_write_test_{time.time_ns()}"
            probe.write_bytes(b"")
        except OSError as exc:
            raise RunError(f"Нет доступа на запись в папку для снимков:\n{folder}\n({exc})") from exc
        remove_quietly(probe)

    # ------------------------------------------------------------------ запуск
    def run(self) -> Summary:
        summary = Summary(out_dir=self._out_dir)
        started = time.monotonic()
        try:
            self._run(summary)
        except (RunError, ExcelError, BrowserStartError, ValueError) as exc:
            summary.fatal = str(exc)
            self.log("error", str(exc))
        except Exception as exc:                      # непредвиденное — подробности в отладочный журнал
            self._debug(traceback.format_exc())
            summary.fatal = f"Непредвиденная ошибка: {type(exc).__name__}: {exc}"
            self.log("error", summary.fatal)
        summary.journal = self._journal_path
        summary.elapsed_s = time.monotonic() - started
        return summary

    def _run(self, summary: Summary) -> None:
        options, settings = self.options, self.settings
        registries = resolve(options.registries)
        if not registries:
            raise RunError("Не выбран ни один реестр.")
        xlsx = Path(options.xlsx)
        self._prepare_out_dir()
        if not options.save_copy and xlsx.is_file():
            if not os.access(xlsx, os.W_OK):
                raise RunError(f"Файл «{xlsx.name}» доступен только для чтения — записать в него "
                               "ссылки нельзя.\nПолучите права на изменение файла либо отметьте "
                               "«Не менять исходный файл — записать результат в копию».")
            if is_locked(xlsx):
                raise RunError(f"Файл «{xlsx.name}» открыт в Excel.\n"
                               "Закройте его и нажмите «Старт» ещё раз.")

        self.log("info", f"Открываю файл «{xlsx.name}»…")
        book = ExcelBook.open(xlsx, settings)
        try:
            plan = make_plan(book, options)
            layout = plan.layout
            summary.iins_total = len(plan.groups)
            self.log("info", f"Лист «{layout.sheet}»: строк с ИИН — {plan.rows_with_value}, "
                             f"к проверке — {plan.rows_to_check} (разных ИИН: {len(plan.groups)}).")
            if plan.skipped_linked:
                self.log("info", f"Пропущено строк, где ссылка уже есть: {plan.skipped_linked}.")
            if plan.skipped_hidden:
                self.log("info", f"Пропущено скрытых строк: {plan.skipped_hidden}.")
            if not plan.groups and not plan.invalid:
                summary.notes.append(nothing_to_do(plan))
                self.log("info", summary.notes[-1])
                return

            if options.save_copy:
                book.use_copy()
            before = book.layout
            layout = book.ensure_columns(options.write_result)
            self._dirty = layout != before                 # добавлены служебные столбцы
            for task in plan.invalid:
                summary.invalid += 1
                self.log("warn", f"Строка {task.row}: {task.parse.error} — пропущена.")
                self._write_result(book, task.row, f"ОШИБКА: {task.parse.error}")
                self._journal(layout, [task.row], str(task.parse.raw), "НЕВЕРНЫЙ ИИН",
                              note=task.parse.error or "")

            if not plan.groups:
                # Проверять на сайте нечего. Файл трогаем, только если в нём что-то изменилось.
                if self._dirty:
                    self._protect_original(book, summary)
                    self._save(book, summary, final=True)
                return

            self._protect_original(book, summary)
            # Сохраняем сразу: если файл нельзя записать, это выяснится до часа работы с сайтом.
            book.save()
            summary.workbook = book.save_path
            self._dirty = False
            try:
                if self.stop_event.is_set():
                    summary.stopped = True
                else:
                    self.log("info", "Запускаю браузер…")
                    with self._browser_factory() as browser:
                        self.log("info", f"Браузер: {getattr(browser, 'description', '')}. "
                                         f"Реестры: {', '.join(r.short for r in registries)}.")
                        self._loop(book, browser, registries, plan, summary)
            finally:
                if self._dirty:
                    self._save(book, summary, final=True)
        finally:
            book.close()

    def _protect_original(self, book: ExcelBook, summary: Summary) -> None:
        """Перед первой записью: резервная копия либо сообщение, что пишем в копию."""
        if book.save_path != book.path:
            self.log("info", "Исходный файл не изменяется. Результат — в копии "
                             f"«{book.save_path.name}».")
        else:
            summary.backup = book.backup()
            if summary.backup:
                self.log("info", f"Резервная копия: {summary.backup}")

    # ------------------------------------------------------------------ основной цикл
    def _loop(self, book: ExcelBook, browser, registries: list[Registry],
              plan: Plan, summary: Summary) -> None:
        settings, layout = self.settings, book.layout
        total = len(plan.groups)
        consecutive_failures = unsaved = 0
        self._on_progress(0, total, summary)
        for index, (iin, rows) in enumerate(plan.groups.items(), start=1):
            if self.stop_event.is_set():
                summary.stopped = True
                break
            try:
                shots, error = self._check_iin(browser, registries, iin)
            except _Stopped:
                summary.stopped = True
                break

            if shots is None:
                summary.errors += 1
                consecutive_failures += 1
                self.log("error", f"[{index}/{total}] ИИН {iin}: не удалось проверить — {error}")
                for row in rows:
                    self._write_result(book, row, f"ОШИБКА: {error}")
                self._journal(layout, rows, iin, "ОШИБКА", note=error or "")
            else:
                consecutive_failures = 0
                found, text = describe(shots)
                stamp = shots[-1].checked_at.strftime("%Y-%m-%d_%H-%M-%S")
                path = self._free_name(f"{iin}_{stamp}")
                try:
                    save_png(build_image(shots, stamp=settings.stamp,
                                         scale=settings.device_scale_factor), path)
                except OSError as exc:
                    raise RunError(f"Не удалось сохранить снимок в папку:\n{self._out_dir}\n({exc})") from exc
                for row in rows:
                    try:
                        book.write_link(row, path)
                        summary.rows_linked += 1
                    except ExcelError as exc:
                        self.log("warn", str(exc))
                    self._write_result(book, row, text)
                self._dirty = True
                if found:
                    summary.found += 1
                else:
                    summary.not_found += 1
                self.log("found" if found else "info",
                         f"[{index}/{total}] ИИН {iin}: {text.split(' · ')[0]}")
                self._journal(layout, rows, iin, "НАЙДЕН" if found else "НЕ НАЙДЕН",
                              "; ".join(f"{s.registry.key}={s.count}" for s in shots), path.name)

            summary.iins_done = index
            unsaved += 1
            self._on_progress(index, total, summary)
            if consecutive_failures >= settings.max_consecutive_failures:
                summary.aborted = (f"Сайт не отвечает: {consecutive_failures} ИИН подряд не удалось "
                                   "проверить. Попробуйте позже — уже сделанное сохранено.")
                self.log("error", summary.aborted)
                break
            if unsaved >= settings.save_every and self._dirty:
                self._save(book, summary, final=False)
                unsaved = 0

    def _check_iin(self, browser, registries: list[Registry], iin: str):
        """Снимки всех выбранных реестров по одному ИИН либо причина неудачи."""
        settings = self.settings
        shots: list[PageShot] = []
        for registry in registries:
            error: str | None = None
            attempt = limited = 0
            while True:
                if self.stop_event.is_set():
                    raise _Stopped()
                try:
                    shots.append(browser.check(registry, iin))
                    error = None
                    break
                except RateLimitedError as exc:
                    # Не ошибка ИИН, а просьба сайта не торопиться: ждём и замедляемся.
                    limited += 1
                    error = str(exc)
                    if limited > settings.rate_limit_retries:
                        break
                    pause = min(settings.rate_limit_pause_s * 2 ** (limited - 1), 180.0)
                    slower = getattr(browser, "slow_down", None)
                    pace = f" Дальше — один запрос в {slower():.1f} с." if slower else ""
                    self.log("warn", f"Сайт сообщил «Превышен лимит запросов». "
                                     f"Жду {pause:.0f} с и повторяю ИИН {iin}.{pace}")
                    self._sleep(pause)
                except CaptureError as exc:
                    attempt += 1
                    error = str(exc)
                    if attempt >= settings.retries:
                        break
                    backoff = settings.retry_backoff_s or [5.0]
                    pause = backoff[min(attempt - 1, len(backoff) - 1)]
                    self.log("warn", f"ИИН {iin}: {error}. Повтор через {pause:.0f} с "
                                     f"(попытка {attempt + 1} из {settings.retries}).")
                    self._sleep(pause)
            if error is not None:
                return None, f"{registry.short}: {error}"
        return shots, None

    # ------------------------------------------------------------------ запись
    def _free_name(self, stem: str) -> Path:
        path = self._out_dir / f"{stem}.png"
        counter = 2
        while path.exists():
            path = self._out_dir / f"{stem}_{counter}.png"
            counter += 1
        return path

    def _write_result(self, book: ExcelBook, row: int, text: str) -> None:
        if book.layout is None or book.layout.result_col is None:
            return
        try:
            if book.write_result(row, text):
                self._dirty = True
        except ExcelError as exc:
            self.log("warn", str(exc))

    def _save(self, book: ExcelBook, summary: Summary, *, final: bool) -> None:
        """Сохранение книги. Если файл в этот момент открыт в Excel: промежуточное сохранение
        откладывается, а итоговое ждёт, после чего результат уходит в отдельный файл."""
        deadline = time.monotonic() + FINAL_SAVE_WAIT_S
        warned = False
        while True:
            try:
                book.save()
                self._dirty = False
                summary.workbook = book.save_path
                if warned:
                    self.log("info", "Файл сохранён.")
                return
            except ExcelLockedError as exc:
                if not final:
                    self.log("warn", f"{exc} Продолжаю работу, сохраню позже.")
                    return
                if not warned:
                    self.log("warn", f"{exc} Жду до {FINAL_SAVE_WAIT_S // 60} минут…")
                    warned = True
                if time.monotonic() >= deadline:
                    break
                time.sleep(3)
            except ExcelError as exc:
                self.log("error", str(exc))
                if not final:
                    return
                break
        # Не удалось записать в исходный файл — результат не должен пропасть.
        rescue = book.path.with_name(
            f"{book.path.stem}_tazalau_unsaved_{datetime.now():%Y%m%d_%H%M%S}{book.path.suffix}")
        book.save_path = rescue
        try:
            book.save()
            self._dirty = False
            summary.workbook = rescue
            note = (f"ВНИМАНИЕ: исходный файл был занят, результат сохранён отдельно:\n{rescue}")
        except ExcelError as exc:
            note = (f"ВНИМАНИЕ: не удалось сохранить результат в Excel ({exc}). "
                    f"Снимки и журнал находятся в папке:\n{self._out_dir}")
        summary.notes.append(note)
        self.log("error", note)
