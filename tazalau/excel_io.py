"""Книга Excel: чтение ИИН и запись ссылок на снимки.

Главный принцип — не навредить рабочему файлу:
  * перед первой записью создаётся резервная копия;
  * файл сохраняется во временный файл рядом и только после проверки заменяет исходный;
  * если в книге есть элементы, которые библиотека openpyxl не умеет сохранять без потерь
    (фигуры, срезы, сводные таблицы, макросы, метка конфиденциальности…), программа сообщает
    об этом, и результат можно записать в копию, не трогая исходный файл.
"""
from __future__ import annotations

import copy
import os
import re
import shutil
import sys
import warnings
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .config import BACKUP_DIR_NAME, LINK_HEADER, RESULT_HEADER, Settings
from .fsutil import remove_quietly, replace_with_retry
from .iin import IinParse, normalize_iin

LINK_COLOR = "0563C1"            # стандартный синий цвет гиперссылки Excel
SUPPORTED_SUFFIXES = (".xlsx", ".xlsm")

# Допустимые заголовки столбца с ИИН — по убыванию приоритета (после нормализации).
_IIN_HEADERS = ("иин/бин", "бин/иин", "иин", "iin/bin", "iin")

_WS = re.compile(r"[\s\u00a0\u2000-\u200f\u202f\u2060\ufeff]+")

# Части файла .xlsx, которые openpyxl при сохранении теряет или переписывает неточно.
_RISKY_PARTS = (
    ("xl/revisions/", "совместный доступ к книге (журнал изменений)"),
    ("xl/pivotTables/", "сводные таблицы"),
    ("xl/slicers/", "срезы"),
    ("xl/slicerCaches/", "срезы"),
    ("xl/timelines/", "временные шкалы"),
    ("xl/charts/", "диаграммы"),
    ("xl/chartsheets/", "листы-диаграммы"),
    ("xl/diagrams/", "схемы SmartArt"),
    ("xl/ctrlProps/", "кнопки и элементы управления"),
    ("xl/activeX/", "элементы ActiveX"),
    ("xl/embeddings/", "вложенные документы"),
    ("xl/model/", "модель данных Power Pivot"),
    ("xl/queryTables/", "запросы к внешним данным"),
    ("xl/connections.xml", "подключения к внешним данным"),
    ("xl/threadedComments/", "примечания-обсуждения"),
    ("xl/vbaProject.bin", "макросы"),
    ("docMetadata/", "метка конфиденциальности"),
    ("customXml/", "служебные свойства документа"),
)
_SHAPE_MARKERS = (b"<xdr:sp>", b"<xdr:sp ", b"<xdr:cxnSp", b"<xdr:grpSp")

# Предупреждения openpyxl при чтении -> понятное описание того, что будет потеряно.
_WARNING_TEXT = (
    ("Data Validation extension", "расширенная проверка данных (списки)"),
    ("Conditional Formatting extension", "расширенное условное форматирование"),
    ("Slicer List extension", "срезы"),
    ("Sparkline Group extension", "спарклайны"),
    ("Unknown extension", "дополнительные возможности листа"),
    ("DrawingML support is incomplete", "фигуры и надписи"),
    ("header or footer", "колонтитулы"),
)


class ExcelError(Exception):
    """Ошибка с текстом, который можно показать пользователю как есть."""


class ExcelLockedError(ExcelError):
    """Файл занят (обычно открыт в Excel)."""


@dataclass(frozen=True)
class Layout:
    sheet: str
    header_row: int
    iin_col: int
    link_col: int | None
    result_col: int | None
    last_row: int               # последняя строка со значением в столбце ИИН


@dataclass(frozen=True)
class RowTask:
    row: int
    parse: IinParse
    has_link: bool
    hidden: bool


@dataclass(frozen=True)
class SheetStats:
    rows_with_value: int
    invalid: int
    already_linked: int
    hidden: int


def _plain(value: object) -> object:
    """CellRichText (ячейка с частичным форматированием) -> обычная строка."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):          # CellRichText — наследник list
        return "".join(str(part) for part in value)
    return value


def _norm_header(value: object) -> str:
    value = _plain(value)
    if value is None:
        return ""
    return _WS.sub("", str(value)).lower().replace("\\", "/")


def _is_blank(value: object) -> bool:
    value = _plain(value)
    return value is None or (isinstance(value, str) and not _WS.sub("", value))


def _cell_items(ws):
    """Существующие ячейки листа без создания новых (iter_rows создаёт пустые ячейки, а на
    листе с форматированием «до самого низа» это миллион объектов)."""
    cells = getattr(ws, "_cells", None)
    if isinstance(cells, dict):
        return list(cells.items())
    return [((c.row, c.column), c) for row in ws.iter_rows() for c in row]


def find_layout(ws, scan_rows: int = 30) -> Layout | None:
    """Ищет строку заголовков со столбцом «ИИН/БИН» и служебные столбцы программы."""
    items = _cell_items(ws)
    best: tuple[int, int, int] | None = None
    for (row, col), cell in items:
        if row > scan_rows:
            continue
        name = _norm_header(cell.value)
        if name in _IIN_HEADERS:
            rank = (_IIN_HEADERS.index(name), row, col)
            if best is None or rank < best:
                best = rank
    if best is None:
        return None
    _, header_row, iin_col = best

    link_col = result_col = None
    last_row = header_row
    for (row, col), cell in items:
        if row == header_row:
            name = _norm_header(cell.value)
            if name == LINK_HEADER and (link_col is None or col < link_col):
                link_col = col
            elif name == RESULT_HEADER and (result_col is None or col < result_col):
                result_col = col
        elif row > header_row and col == iin_col and not _is_blank(cell.value):
            last_row = max(last_row, row)
    return Layout(ws.title, header_row, iin_col, link_col, result_col, last_row)


def scan_risky_features(path: Path) -> list[str]:
    """Что в книге может пострадать при сохранении через openpyxl (по содержимому архива)."""
    found: list[str] = []
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            for prefix, label in _RISKY_PARTS:
                if any(n == prefix or n.startswith(prefix) for n in names) and label not in found:
                    found.append(label)
            for name in names:
                if name.startswith("xl/drawings/drawing") and name.endswith(".xml"):
                    data = archive.read(name)
                    if any(marker in data for marker in _SHAPE_MARKERS):
                        if "фигуры и надписи" not in found:
                            found.append("фигуры и надписи")
                        break
    except (OSError, zipfile.BadZipFile):
        pass                     # сам файл проверяется при открытии, здесь только подсказка
    return found


def is_locked(path: Path) -> bool:
    """Открыт ли файл в Excel. В Windows Excel блокирует файл на запись; в macOS — только
    кладёт рядом служебный файл «~$имя»."""
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        return True
    except OSError:
        return False
    if not sys.platform.startswith("win"):
        return (path.parent / f"~${path.name}").exists()
    return False


def link_target(file_path: Path, workbook_path: Path, mode: str = "absolute") -> tuple[str, str]:
    """(адрес гиперссылки, текст ячейки) для файла снимка.

    Адрес записывается в том же виде, в каком его пишет сам Excel: ``file:///W:\\папка\\файл``,
    пробелы и служебные символы «%», «#» закодированы, кириллица — как есть.
    """
    file_path = Path(os.path.abspath(file_path))
    if mode == "relative":
        try:
            rel = os.path.relpath(file_path, os.path.abspath(workbook_path.parent))
        except ValueError:           # разные диски в Windows — относительного пути не существует
            rel = None
        if rel is not None:
            return _escape(rel.replace("\\", "/")), rel
    text = str(file_path)
    return absolute_target(text), text


def absolute_target(path_text: str) -> str:
    r"""Адрес для полного пути — в том виде, в каком его записывает сам Excel:

    ``/Users/a b/f.png``       -> ``file:///Users/a%20b/f.png``
    ``W:\a b\f.png``           -> ``file:///W:\a%20b\f.png``
    ``\\srv\share\a b\f.png``  -> ``file:///\\srv\share\a%20b\f.png``
    """
    if path_text.startswith("/"):
        return "file://" + _escape(path_text)
    return "file:///" + _escape(path_text)


def _escape(path_text: str) -> str:
    return path_text.replace("%", "%25").replace("#", "%23").replace(" ", "%20")


class ExcelBook:
    """Открытая книга: один лист с ИИН, чтение заданий и запись результатов."""

    def __init__(self, path: Path, settings: Settings, workbook, load_warnings: list[str]):
        self.path = Path(path)
        self.settings = settings
        self.wb = workbook
        self.save_path = self.path          # куда сохранять (исходный файл или копия)
        self.layout: Layout | None = None
        self._values_wb = None              # та же книга с вычисленными значениями формул
        self._backup_done = False
        self.risks = scan_risky_features(self.path)
        for label in load_warnings:
            if label not in self.risks:
                self.risks.append(label)

    # ---------------------------------------------------------------- открытие
    @classmethod
    def open(cls, path: Path | str, settings: Settings) -> "ExcelBook":
        path = Path(path)
        workbook, notes = _load(path, data_only=False)
        return cls(path, settings, workbook, notes)

    def close(self) -> None:
        for wb in (self.wb, self._values_wb):
            try:
                if wb is not None:
                    wb.close()
                    macros = getattr(wb, "vba_archive", None)     # openpyxl сам его не закрывает
                    if macros is not None:
                        macros.close()
            except Exception:
                pass

    # ---------------------------------------------------------------- листы
    def layouts(self) -> list[Layout]:
        """Листы, на которых найден столбец «ИИН/БИН»."""
        result = []
        for ws in self.wb.worksheets:
            layout = find_layout(ws, self.settings.header_scan_rows)
            if layout is not None:
                result.append(layout)
        return result

    def default_sheet(self) -> str | None:
        layouts = self.layouts()
        if not layouts:
            return None
        names = [l.sheet for l in layouts]
        active = getattr(self.wb.active, "title", None)
        return active if active in names else names[0]

    def use_sheet(self, name: str | None = None) -> Layout:
        layouts = self.layouts()
        if not layouts:
            raise ExcelError(
                "В файле не найден столбец с заголовком «ИИН/БИН».\n"
                "Заголовок должен находиться в одной из первых "
                f"{self.settings.header_scan_rows} строк листа."
            )
        name = name or self.default_sheet()
        for layout in layouts:
            if layout.sheet == name:
                self.layout = layout
                return layout
        raise ExcelError(f"На листе «{name}» нет столбца «ИИН/БИН».")

    @property
    def ws(self):
        if self.layout is None:
            raise ExcelError("Лист не выбран.")
        return self.wb[self.layout.sheet]

    # ---------------------------------------------------------------- чтение
    def read_tasks(self) -> list[RowTask]:
        """Все строки с непустым значением в столбце ИИН, по порядку."""
        layout, ws = self._require_layout(), self.ws
        cells = getattr(ws, "_cells", None)
        tasks: list[RowTask] = []
        for row in range(layout.header_row + 1, layout.last_row + 1):
            cell = cells.get((row, layout.iin_col)) if isinstance(cells, dict) \
                else ws.cell(row=row, column=layout.iin_col)
            if cell is None:
                continue
            value = _plain(cell.value)
            if getattr(cell, "data_type", None) == "f":
                value = self._formula_value(row, layout.iin_col)
                if value is None:
                    parse = IinParse(raw=cell.value, error=(
                        "ИИН задан формулой без сохранённого значения — откройте файл в Excel, "
                        "сохраните и запустите проверку снова"))
                    tasks.append(RowTask(row, parse, self._has_link(row), self._hidden(row)))
                    continue
            parse = normalize_iin(value, allow_bad_checksum=self.settings.allow_bad_checksum)
            if parse.empty:
                continue
            tasks.append(RowTask(row, parse, self._has_link(row), self._hidden(row)))
        return tasks

    def stats(self, tasks: list[RowTask]) -> SheetStats:
        return SheetStats(
            rows_with_value=len(tasks),
            invalid=sum(1 for t in tasks if not t.parse.ok),
            already_linked=sum(1 for t in tasks if t.has_link),
            hidden=sum(1 for t in tasks if t.hidden),
        )

    def _formula_value(self, row: int, col: int):
        if self._values_wb is None:
            self._values_wb, _ = _load(self.path, data_only=True)
        return _plain(self._values_wb[self.layout.sheet].cell(row=row, column=col).value)

    def _has_link(self, row: int) -> bool:
        col = self.layout.link_col
        if col is None:
            return False
        cells = getattr(self.ws, "_cells", None)
        cell = cells.get((row, col)) if isinstance(cells, dict) else self.ws.cell(row=row, column=col)
        if cell is None:
            return False
        return not _is_blank(cell.value) or getattr(cell, "hyperlink", None) is not None

    def _hidden(self, row: int) -> bool:
        dimension = self.ws.row_dimensions.get(row)     # .get не создаёт запись о строке
        return bool(dimension is not None and dimension.hidden)

    def _require_layout(self) -> Layout:
        if self.layout is None:
            raise ExcelError("Лист не выбран.")
        return self.layout

    # ---------------------------------------------------------------- запись
    def ensure_columns(self, want_result: bool) -> Layout:
        """Создаёт недостающие столбцы tazalau_file_location / tazalau_result справа от таблицы."""
        layout = self._require_layout()
        link_col, result_col = layout.link_col, layout.result_col
        if link_col is None:
            link_col = self._new_column(LINK_HEADER, width=70)
        if want_result and result_col is None:
            result_col = self._new_column(RESULT_HEADER, width=46, taken=(link_col,))
        self.layout = Layout(layout.sheet, layout.header_row, layout.iin_col,
                             link_col, result_col, layout.last_row)
        return self.layout

    def _new_column(self, header: str, width: float, taken: tuple[int, ...] = ()) -> int:
        layout, ws = self.layout, self.ws
        used_cols, header_last = set(taken), layout.iin_col
        for (row, col), cell in _cell_items(ws):
            if _is_blank(cell.value):
                continue
            used_cols.add(col)
            if row == layout.header_row:
                header_last = max(header_last, col)
        # Объединённые ячейки в строке заголовков (даже пустые) — тоже часть шапки таблицы.
        for merged in ws.merged_cells.ranges:
            if merged.min_row <= layout.header_row <= merged.max_row:
                header_last = max(header_last, merged.max_col)
        col = max([header_last, *taken]) + 1
        while col in used_cols:                # столбец занят данными без заголовка — правее
            col += 1
        source = ws.cell(row=layout.header_row, column=layout.iin_col)
        target = ws.cell(row=layout.header_row, column=col)
        try:
            target.value = header
        except AttributeError as exc:          # попали в объединённую ячейку
            raise ExcelError(
                f"Не удалось добавить столбец «{header}»: ячейка "
                f"{get_column_letter(col)}{layout.header_row} объединена с соседними.") from exc
        if source.has_style:
            for attr in ("font", "fill", "border", "alignment", "protection"):
                setattr(target, attr, copy.copy(getattr(source, attr)))
        ws.column_dimensions[get_column_letter(col)].width = width
        return col

    def write_link(self, row: int, file_path: Path) -> str:
        """Пишет в tazalau_file_location путь к снимку как рабочую гиперссылку."""
        layout = self._require_layout()
        if layout.link_col is None:
            raise ExcelError("Столбец tazalau_file_location не создан.")
        target, text = link_target(file_path, self.save_path, self.settings.link_mode)
        cell = self.ws.cell(row=row, column=layout.link_col)
        try:
            cell.value = text
            cell.hyperlink = target
            font = copy.copy(cell.font)
            font.color = LINK_COLOR
            font.underline = "single"
            cell.font = font
        except AttributeError as exc:
            raise ExcelError(f"Строка {row}: ячейка для ссылки объединена с соседними.") from exc
        return text

    def write_result(self, row: int, text: str) -> bool:
        """Пишет итог в tazalau_result. Возвращает True, если содержимое ячейки изменилось."""
        layout = self._require_layout()
        if layout.result_col is None:
            return False
        cell = self.ws.cell(row=row, column=layout.result_col)
        if _plain(cell.value) == text:
            return False
        try:
            cell.value = text
        except AttributeError as exc:
            raise ExcelError(f"Строка {row}: ячейка для результата объединена с соседними.") from exc
        return True

    # ---------------------------------------------------------------- сохранение
    def use_copy(self) -> Path:
        """Не трогать исходный файл: результат пойдёт в новую копию рядом с ним."""
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.save_path = self.path.with_name(f"{self.path.stem}_tazalau_{stamp}{self.path.suffix}")
        return self.save_path

    def backup(self) -> Path | None:
        """Резервная копия исходного файла (один раз за запуск, только при записи «на место»)."""
        if self._backup_done or self.save_path != self.path:
            return None
        folder = self.path.parent / BACKUP_DIR_NAME
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destination = folder / f"{self.path.stem}__{stamp}{self.path.suffix}"
        try:
            folder.mkdir(exist_ok=True)
            shutil.copy2(self.path, destination)
        except OSError as exc:
            # Без резервной копии исходный файл не трогаем вообще.
            raise ExcelError(
                f"Не удалось создать резервную копию в папке «{self.path.parent}» ({exc}).\n"
                "Проверьте права на запись в эту папку.") from exc
        self._backup_done = True
        return destination

    def save(self) -> Path:
        """Сохраняет книгу: временный файл рядом -> проверка -> замена. Исходный файл либо
        остаётся прежним, либо заменяется целым новым — половинчатого состояния не бывает."""
        target = self.save_path
        temp = target.with_name(f"~tazalau_tmp_{os.getpid()}_{target.name}")
        try:
            self.wb.save(temp)
            try:
                with zipfile.ZipFile(temp) as archive:
                    intact = (archive.testzip() is None
                              and "[Content_Types].xml" in archive.namelist())
            except zipfile.BadZipFile:
                intact = False
            if not intact:
                raise ExcelError(f"Не удалось сохранить файл «{target.name}»: "
                                 "результат записи повреждён, исходный файл не изменён.")
            replace_with_retry(temp, target)
        except ExcelError:
            raise
        except PermissionError as exc:
            raise ExcelLockedError(
                f"Файл «{target.name}» занят или недоступен для записи — скорее всего, "
                "он открыт в Excel. Закройте его.") from exc
        except Exception as exc:         # любая ошибка записи: исходный файл остаётся прежним
            raise ExcelError(f"Не удалось сохранить файл «{target.name}»: "
                             f"{type(exc).__name__}: {exc}") from exc
        finally:
            remove_quietly(temp)
        return target


def _load(path: Path, *, data_only: bool):
    """load_workbook с понятными сообщениями об ошибках. Возвращает (книга, [что будет потеряно])."""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise ExcelError("Старый формат .xls не поддерживается.\n"
                         "Откройте файл в Excel и сохраните как «Книга Excel (*.xlsx)».")
    if suffix == ".xlsb":
        raise ExcelError("Двоичный формат .xlsb не поддерживается.\n"
                         "Сохраните файл как «Книга Excel (*.xlsx)».")
    if suffix not in SUPPORTED_SUFFIXES:
        raise ExcelError("Нужен файл Excel с расширением .xlsx.")
    if not path.is_file():
        raise ExcelError(f"Файл не найден: {path}")
    try:
        with open(path, "rb") as handle:
            signature = handle.read(8)
    except PermissionError as exc:
        raise ExcelLockedError(f"Нет доступа к файлу «{path.name}» — возможно, он открыт в Excel.") from exc
    except OSError as exc:
        raise ExcelError(f"Не удалось прочитать файл «{path.name}»: {exc}") from exc
    if signature.startswith(b"\xd0\xcf\x11\xe0"):
        raise ExcelError(f"Файл «{path.name}» защищён паролем (зашифрован). "
                         "Снимите пароль в Excel и повторите.")
    try:
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            try:
                workbook = load_workbook(path, keep_vba=(suffix == ".xlsm"),
                                         data_only=data_only, rich_text=not data_only)
            except (TypeError, ValueError, AttributeError):
                if data_only:
                    raise
                # Редкие книги не читаются с сохранением форматирования внутри ячеек.
                workbook = load_workbook(path, keep_vba=(suffix == ".xlsm"), data_only=False)
    except PermissionError as exc:
        raise ExcelLockedError(f"Нет доступа к файлу «{path.name}» — возможно, он открыт в Excel.") from exc
    except Exception as exc:
        raise ExcelError(f"Не удалось открыть «{path.name}»: файл повреждён "
                         f"или это не книга Excel ({type(exc).__name__}).") from exc
    notes: list[str] = []
    for item in caught:
        message = str(item.message)
        label = next((ru for key, ru in _WARNING_TEXT if key in message), None)
        if label and label not in notes:
            notes.append(label)
    return workbook, notes
