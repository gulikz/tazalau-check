"""Запись книг Excel: первая строка листа: заголовок с фильтром, закреплена; цвет ячейки по уровню (плохо, внимание,
в порядке); ссылки на файлы относительные (папку можно переносить целиком), на сайты: полные. ИИН и номера договоров
пишутся текстом: ведущие нули не теряются."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FILLS = {
    "bad": (PatternFill("solid", fgColor="FFC7CE"), Font(color="9C0006")),
    "warn": (PatternFill("solid", fgColor="FFEB9C"), Font(color="9C5700")),
    "good": (PatternFill("solid", fgColor="C6EFCE"), Font(color="006100")),
}
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
LINK_FONT = Font(color="0563C1", underline="single")


@dataclass
class X:
    """Ячейка с оформлением: v: значение, s: уровень (bad, warn, good), link: адрес или путь к файлу."""
    v: object = ""
    s: str = ""
    link: str = ""


def _plain(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date, int, float)) and not isinstance(v, bool):
        return v
    text = str(v)
    return text[:32000] if len(text) > 32000 else text


def file_link(target: Path | str, book: Path) -> str:
    """Путь к файлу относительно книги (ссылка работает, пока папка переносится целиком)."""
    target = Path(target)
    try:
        return os.path.relpath(target, Path(book).parent).replace("\\", "/")
    except ValueError:                     # другой диск в Windows: только полный путь
        return target.resolve().as_uri()


def write_book(path: Path, sheets: list[dict]) -> Path:
    """sheets: [{name, rows: [[значение | X]], widths: [ширина]}]. Пишет через временный файл: открытая в Excel
    старая книга не портится (PermissionError объясняет, что файл открыт)."""
    wb = Workbook()
    wb.remove(wb.active)
    for sh in sheets:
        ws = wb.create_sheet(str(sh["name"])[:31].translate(str.maketrans("\\/?*:[]", "       ")))
        rows = sh["rows"]
        for r, line in enumerate(rows, start=1):
            for c, value in enumerate(line, start=1):
                x = value if isinstance(value, X) else X(value)
                cell = ws.cell(row=r, column=c, value=_plain(x.v))
                if isinstance(x.v, (datetime, date)) and not isinstance(x.v, datetime):
                    cell.number_format = "DD.MM.YYYY"
                elif isinstance(x.v, datetime):
                    cell.number_format = "DD.MM.YYYY HH:MM"
                if r == 1:
                    cell.font = Font(bold=True)
                    cell.fill = HEAD_FILL
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    continue
                if x.s in FILLS:
                    cell.fill, cell.font = FILLS[x.s]
                if x.link:
                    cell.hyperlink = x.link
                    if x.s not in FILLS:
                        cell.font = LINK_FONT
        for i, w in enumerate(sh.get("widths") or [], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        if rows:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(max(len(rows[0]), 1))}{max(len(rows), 1)}"
    path = Path(path)
    tmp = path.with_name("~tmp_" + path.name)
    wb.save(tmp)
    try:
        os.replace(tmp, path)
    except PermissionError as exc:
        tmp.unlink(missing_ok=True)
        raise PermissionError(f"Файл {path.name} открыт в Excel: закройте его и повторите.") from exc
    return path
