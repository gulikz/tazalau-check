"""Книга Excel: поиск столбцов, чтение ИИН, запись ссылок и сохранность файла."""
from __future__ import annotations

import os
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from tazalau.config import BACKUP_DIR_NAME, LINK_HEADER, RESULT_HEADER, Settings
from tazalau.excel_io import (ExcelBook, ExcelError, ExcelLockedError, absolute_target, find_layout,
                              is_locked, link_target, scan_risky_features)

from conftest import IIN_A, IIN_B, IIN_C, IIN_D


def _open(path: Path, settings: Settings | None = None) -> ExcelBook:
    book = ExcelBook.open(path, settings or Settings())
    book.use_sheet()
    return book


def _patch_zip(path: Path, edit) -> None:
    """Переписывает архив .xlsx, пропуская каждую часть через edit(имя, данные) -> данные | None."""
    temp = path.with_suffix(".patched")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = edit(item.filename, source.read(item.filename))
            if data is not None:
                target.writestr(item, data)
    os.replace(temp, path)


# ------------------------------------------------------------------ поиск столбцов
def test_layout_in_plain_book(book_path):
    book = _open(book_path)
    layout = book.layout
    assert (layout.sheet, layout.header_row, layout.iin_col) == ("восстановление", 1, 2)
    assert layout.link_col is None and layout.result_col is None
    assert layout.last_row == 8
    assert [l.sheet for l in book.layouts()] == ["восстановление"]


@pytest.mark.parametrize("header", ["ИИН/БИН", " иин / бин ", "ИИН\\БИН", "ИИН/\nБИН", "ИИН/БИН\u00a0",
                                    "БИН/ИИН", "ИИН", "IIN"])
def test_header_spellings(tmp_path, header):
    wb = Workbook()
    ws = wb.active
    ws.append(["Отчёт по банкротству"])
    ws.append([])
    ws.append(["№", "ФИО", header])
    ws.append([1, "кто-то", IIN_A])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    layout = _open(path).layout
    assert (layout.header_row, layout.iin_col, layout.last_row) == (3, 3, 4)


def test_exact_header_wins_over_similar(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН", "ФИО", "ИИН/БИН"])
    ws.append(["не то", "x", IIN_A])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    assert _open(path).layout.iin_col == 3


def test_header_not_found_is_explained(tmp_path):
    wb = Workbook()
    wb.active.append(["ФИО", "Счёт"])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    with pytest.raises(ExcelError, match="ИИН/БИН"):
        _open(path)


def test_header_below_scan_window_is_not_used(tmp_path):
    wb = Workbook()
    ws = wb.active
    for _ in range(40):
        ws.append(["x"])
    ws.append(["ИИН/БИН"])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    assert find_layout(load_workbook(path).active, scan_rows=30) is None


def test_sheet_choice_and_default_is_active_sheet(tmp_path):
    wb = Workbook()
    first = wb.active
    first.title = "карты"
    first.append(["ИИН/БИН"])
    first.append([IIN_A])
    second = wb.create_sheet("восстановление")
    second.append(["ФИО", "ИИН/БИН"])
    second.append(["x", IIN_B])
    wb.active = 1
    path = tmp_path / "h.xlsx"
    wb.save(path)
    book = ExcelBook.open(path, Settings())
    assert [l.sheet for l in book.layouts()] == ["карты", "восстановление"]
    assert book.default_sheet() == "восстановление"
    assert book.use_sheet("карты").iin_col == 1
    with pytest.raises(ExcelError):
        book.use_sheet("нет такого")


# ------------------------------------------------------------------ чтение заданий
def test_read_tasks(book_path):
    tasks = _open(book_path).read_tasks()
    assert [t.row for t in tasks] == [2, 3, 4, 5, 7, 8]               # пустая строка 6 пропущена
    assert [t.parse.iin for t in tasks] == [IIN_A, IIN_B, IIN_A, None, IIN_C, IIN_D]
    assert "контрольная цифра" in tasks[3].parse.error
    assert tasks[1].parse.padded and tasks[4].parse.padded
    assert not any(t.has_link or t.hidden for t in tasks)


def test_hidden_rows_are_flagged(book_path):
    wb = load_workbook(book_path)
    wb.active.row_dimensions[3].hidden = True
    wb.save(book_path)
    assert [t.row for t in _open(book_path).read_tasks() if t.hidden] == [3]


def test_rich_text_cell_is_read_as_plain(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН"])
    ws["A2"] = CellRichText(IIN_A[:4], TextBlock(InlineFont(b=True), IIN_A[4:]))
    path = tmp_path / "rich.xlsx"
    wb.save(path)
    tasks = _open(path).read_tasks()
    assert [t.parse.iin for t in tasks] == [IIN_A]


def test_formula_iin_uses_saved_value_or_explains(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН", "источник"])
    ws.append(["=B2", IIN_A])          # значение формулы будет вписано ниже, как это делает Excel
    ws.append(["=B3", IIN_B])          # формула без сохранённого значения
    path = tmp_path / "formula.xlsx"
    wb.save(path)

    def edit(name, data):
        if name == "xl/worksheets/sheet1.xml":
            text = data.decode("utf-8")
            patched = re.sub(r'(<c r="A2"[^>]*>)<f>B2</f><v\s*/>', rf'\1<f>B2</f><v>{IIN_A}</v>', text)
            patched = patched.replace('<c r="A2" t="str">', '<c r="A2" t="str">')
            if patched == text:
                patched = re.sub(r'<c r="A2"([^>]*)><f>B2</f>(<v></v>|<v/>)?',
                                 rf'<c r="A2"\1><f>B2</f><v>{IIN_A}</v>', text)
            assert patched != text, text[:600]
            return patched.encode("utf-8")
        return data

    _patch_zip(path, edit)
    tasks = _open(path).read_tasks()
    assert tasks[0].parse.iin == IIN_A
    assert not tasks[1].parse.ok and "формулой" in tasks[1].parse.error


def test_sheet_formatted_to_the_last_row_is_read_quickly(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН"])
    ws.append([IIN_A])
    ws.cell(row=1_048_576, column=1).fill = PatternFill("solid", fgColor="EEEEEE")
    path = tmp_path / "tall.xlsx"
    wb.save(path)
    started = time.perf_counter()
    book = _open(path)
    tasks = book.read_tasks()
    assert time.perf_counter() - started < 3
    assert book.layout.last_row == 2 and [t.row for t in tasks] == [2]


# ------------------------------------------------------------------ столбцы программы
def test_new_columns_go_right_after_the_table(book_path):
    book = _open(book_path)
    layout = book.ensure_columns(want_result=True)
    assert (layout.link_col, layout.result_col) == (4, 5)
    ws = book.ws
    assert (ws.cell(1, 4).value, ws.cell(1, 5).value) == (LINK_HEADER, RESULT_HEADER)
    assert book.ensure_columns(want_result=True) == layout               # повторный вызов ничего не меняет


def test_new_column_skips_columns_holding_data_without_header(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН", "ФИО"])
    ws.append([IIN_A, "x", "заметка без заголовка"])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    book = _open(path)
    assert book.ensure_columns(want_result=False).link_col == 4
    assert book.layout.result_col is None


def test_existing_columns_are_reused_whatever_the_case(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН", " Tazalau_File_Location ", "прочее", "TAZALAU_RESULT"])
    ws.append([IIN_A])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    layout = _open(path).ensure_columns(want_result=True)
    assert (layout.link_col, layout.result_col) == (2, 4)


def test_new_header_copies_the_look_of_existing_header(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН"])
    ws["A1"].font = Font(bold=True, size=9)
    ws["A1"].fill = PatternFill("solid", fgColor="C6E0B4")
    ws["A1"].alignment = Alignment(wrap_text=True)
    ws.append([IIN_A])
    path = tmp_path / "h.xlsx"
    wb.save(path)
    book = _open(path)
    book.ensure_columns(want_result=False)
    cell = book.ws.cell(1, 2)
    assert cell.font.b and cell.fill.fgColor.rgb.endswith("C6E0B4") and cell.alignment.wrap_text


# ------------------------------------------------------------------ ссылки
@pytest.mark.parametrize("path_text,expected", [
    ("/Users/user/Desktop/снимки отдела/1.png", "file:///Users/user/Desktop/снимки%20отдела/1.png"),
    (r"W:\Отдел\Папка с пробелами\Банкротство\1.png",
     r"file:///W:\Отдел\Папка%20с%20пробелами\Банкротство\1.png"),
    (r"\\fileserver\share\Общая папка\1.png", r"file:///\\fileserver\share\Общая%20папка\1.png"),
    (r"C:\100% готово\1.png", r"file:///C:\100%25%20готово\1.png"),
])
def test_absolute_target_matches_excel_own_format(path_text, expected):
    assert absolute_target(path_text) == expected


def test_link_roundtrip_points_to_the_real_file(book_path, tmp_path):
    shots = tmp_path / "папка со снимками"
    shots.mkdir()
    picture = shots / f"{IIN_A}_2026-11-15_10-00-00.png"
    picture.write_bytes(b"png")
    book = _open(book_path)
    book.ensure_columns(want_result=True)
    text = book.write_link(2, picture)
    book.write_result(2, "НАЙДЕН")
    book.save()

    ws = load_workbook(book_path)["восстановление"]
    cell = ws.cell(2, 4)
    assert cell.value == text == str(picture)
    assert cell.hyperlink.target == absolute_target(str(picture))
    assert cell.font.color.rgb.endswith("0563C1") and cell.font.underline == "single"
    assert ws.cell(2, 5).value == "НАЙДЕН"
    # Адрес, раскодированный обратно, — это существующий файл.
    from urllib.parse import unquote
    assert Path(unquote(cell.hyperlink.target[len("file://"):])).is_file() or sys.platform.startswith("win")
    # Теперь строка считается обработанной.
    tasks = {t.row: t for t in _open(book_path).read_tasks()}
    assert tasks[2].has_link and not tasks[3].has_link


def test_relative_link_mode(book_path, tmp_path):
    picture = tmp_path / "tazalau_screenshots" / "1 2.png"
    target, text = link_target(picture, book_path, "relative")
    assert target == "tazalau_screenshots/1%202.png"
    assert text == os.path.join("tazalau_screenshots", "1 2.png")
    assert link_target(picture, book_path, "absolute")[1] == str(picture)


def test_link_keeps_the_cell_formatting(book_path, tmp_path):
    wb = load_workbook(book_path)
    ws = wb.active
    ws.cell(1, 4).value = LINK_HEADER
    ws.cell(2, 4).fill = PatternFill("solid", fgColor="FFFF00")
    ws.cell(2, 4).font = Font(name="Calibri", size=9, italic=True)
    wb.save(book_path)
    book = _open(book_path)
    book.ensure_columns(want_result=False)
    book.write_link(2, tmp_path / "x.png")
    book.save()
    cell = load_workbook(book_path).active.cell(2, 4)
    assert cell.fill.fgColor.rgb.endswith("FFFF00")
    assert (cell.font.name, cell.font.sz, cell.font.i) == ("Calibri", 9, True)


def test_recheck_replaces_the_link_and_repeated_saves_do_not_duplicate(book_path, tmp_path):
    book = _open(book_path)
    book.ensure_columns(want_result=False)
    book.write_link(2, tmp_path / "old.png")
    book.save()
    book.write_link(2, tmp_path / "new.png")
    book.write_link(3, tmp_path / "other.png")
    book.save()
    book.save()
    with zipfile.ZipFile(book_path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        rels = archive.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
    assert sheet.count("<hyperlink ") == 2 and rels.count("relationships/hyperlink") == 2
    assert "old.png" not in rels and "new.png" in rels


def test_merged_target_cell_is_reported_not_crashed(book_path, tmp_path):
    wb = load_workbook(book_path)
    wb.active.merge_cells("D2:D3")
    wb.save(book_path)
    book = _open(book_path)
    book.ensure_columns(want_result=False)
    assert book.layout.link_col == 4
    with pytest.raises(ExcelError, match="объединена"):
        book.write_link(3, tmp_path / "x.png")


# ------------------------------------------------------------------ сохранность файла
def test_backup_is_an_exact_copy_and_made_once(book_path):
    original = book_path.read_bytes()
    book = _open(book_path)
    backup = book.backup()
    assert backup.parent.name == BACKUP_DIR_NAME and backup.read_bytes() == original
    assert book.backup() is None
    assert len(list(backup.parent.iterdir())) == 1


def test_save_leaves_no_temporary_files(book_path, tmp_path):
    book = _open(book_path)
    book.ensure_columns(want_result=True)
    book.write_link(2, tmp_path / "x.png")
    book.save()
    leftovers = [p.name for p in book_path.parent.iterdir() if p.name.startswith("~")]
    assert leftovers == []
    assert zipfile.ZipFile(book_path).testzip() is None


def test_failed_save_keeps_the_original_intact(book_path, tmp_path, monkeypatch):
    original = book_path.read_bytes()
    book = _open(book_path)
    book.ensure_columns(want_result=True)
    book.write_link(2, tmp_path / "x.png")

    def broken_save(target):
        Path(target).write_bytes(b"half written")
        raise OSError("диск переполнен")

    monkeypatch.setattr(book.wb, "save", broken_save)
    with pytest.raises(ExcelError, match="Не удалось сохранить"):
        book.save()
    assert book_path.read_bytes() == original
    assert [p.name for p in book_path.parent.iterdir() if p.name.startswith("~")] == []


def test_corrupt_output_never_replaces_the_original(book_path, monkeypatch):
    original = book_path.read_bytes()
    book = _open(book_path)
    monkeypatch.setattr(book.wb, "save", lambda target: Path(target).write_bytes(b"not a zip"))
    with pytest.raises(ExcelError):
        book.save()
    assert book_path.read_bytes() == original


def test_file_open_in_excel_gives_locked_error(book_path, monkeypatch):
    original = book_path.read_bytes()
    book = _open(book_path)

    def locked(source, target):
        raise PermissionError(13, "файл занят другим процессом")

    monkeypatch.setattr("tazalau.fsutil.os.replace", locked)
    with pytest.raises(ExcelLockedError, match="открыт в Excel"):
        book.save()
    assert book_path.read_bytes() == original
    assert [p.name for p in book_path.parent.iterdir() if p.name.startswith("~")] == []


def test_save_survives_a_momentary_lock_by_antivirus(book_path, tmp_path, monkeypatch):
    """Windows: свежий файл на миг занят антивирусом — первая попытка переименования падает."""
    import tazalau.fsutil as fsutil
    real_replace, attempts = fsutil.os.replace, []

    def flaky(source, target):
        attempts.append(target)
        if len(attempts) <= 2:
            raise PermissionError(13, "файл занят другим процессом")
        return real_replace(source, target)

    monkeypatch.setattr(fsutil.os, "replace", flaky)
    book = _open(book_path)
    book.ensure_columns(want_result=False)
    book.write_link(2, tmp_path / "x.png")
    book.save()
    assert len(attempts) == 3
    assert load_workbook(book_path).active.cell(2, 4).hyperlink is not None


def test_copy_mode_never_touches_the_original(book_path, tmp_path):
    original = book_path.read_bytes()
    book = _open(book_path)
    copy_path = book.use_copy()
    assert book.backup() is None                         # копия сама себе страховка
    book.ensure_columns(want_result=False)
    book.write_link(2, tmp_path / "x.png")
    book.save()
    assert book_path.read_bytes() == original
    assert copy_path.parent == book_path.parent and "_tazalau_" in copy_path.name
    assert load_workbook(copy_path).active.cell(2, 4).hyperlink is not None


def test_everything_else_in_the_book_survives(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "восстановление"
    ws.append(["ФИО", "ИИН/БИН", "Сумма", "Дата"])
    ws.append(["первый", IIN_A, 1234.5, "01.10.2026"])
    ws.append(["второй", IIN_B, "=C2*2", "02.10.2026"])
    ws["A2"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["C2"].number_format = "#,##0.00"
    ws["A3"].comment = Comment("проверить", "автор")
    ws.merge_cells("F1:G1")
    ws.auto_filter.ref = "A1:D3"
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 33
    ws.row_dimensions[3].hidden = True
    validation = DataValidation(type="list", formula1='"да,нет"')
    validation.add("E2:E3")
    ws.add_data_validation(validation)
    ws.conditional_formatting.add("C2:C3", CellIsRule(operator="greaterThan", formula=["1000"],
                                                      fill=PatternFill("solid", fgColor="FFC7CE")))
    other = wb.create_sheet("карты")
    other.append(["не трогать", 42])
    path = tmp_path / "full.xlsx"
    wb.save(path)

    book = _open(path)
    book.ensure_columns(want_result=True)
    book.write_link(2, tmp_path / "x.png")
    book.write_result(2, "НАЙДЕН")
    book.save()

    after = load_workbook(path)
    ws = after["восстановление"]
    assert ws["A2"].fill.fgColor.rgb.endswith("FFFF00")
    assert ws["C2"].value == 1234.5 and ws["C2"].number_format == "#,##0.00"
    assert ws["C3"].value == "=C2*2"                                     # формула осталась формулой
    assert ws["A3"].comment.text == "проверить"
    assert "F1:G1" in [str(r) for r in ws.merged_cells.ranges]
    assert ws.auto_filter.ref == "A1:D3" and ws.freeze_panes == "A2"
    assert ws.column_dimensions["A"].width == 33 and ws.row_dimensions[3].hidden
    assert len(ws.data_validations.dataValidation) == 1
    assert len(ws.conditional_formatting) == 1
    assert after["карты"]["A1"].value == "не трогать" and after["карты"]["B1"].value == 42
    assert after.sheetnames == ["восстановление", "карты"]
    # Служебные столбцы встали правее объединённого заголовка F1:G1, ничего не затронув.
    assert (ws.cell(1, 8).value, ws.cell(1, 9).value) == (LINK_HEADER, RESULT_HEADER)


# ------------------------------------------------------------------ рискованные книги и плохие файлы
def test_plain_book_has_no_risks(book_path):
    assert scan_risky_features(book_path) == []
    assert _open(book_path).risks == []


def test_risky_parts_are_named_in_plain_words(book_path):
    with zipfile.ZipFile(book_path, "a") as archive:
        archive.writestr("xl/slicers/slicer1.xml", "<x/>")
        archive.writestr("xl/pivotTables/pivotTable1.xml", "<x/>")
        archive.writestr("docMetadata/LabelInfo.xml", "<x/>")
        archive.writestr("xl/drawings/drawing9.xml", "<xdr:wsDr><xdr:sp macro=''/></xdr:wsDr>")
    risks = scan_risky_features(book_path)
    assert risks == ["сводные таблицы", "срезы", "метка конфиденциальности", "фигуры и надписи"]


def test_openpyxl_loss_warnings_become_risks(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ИИН/БИН"])
    ws.append([IIN_A])
    path = tmp_path / "ext.xlsx"
    wb.save(path)
    extension = (
        '<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations '
        'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">'
        '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>$A$1:$A$2</xm:f>'
        '</x14:formula1><xm:sqref>C1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>')

    def edit(name, data):
        if name == "xl/worksheets/sheet1.xml":
            return data.replace(b"</worksheet>", extension.encode("utf-8") + b"</worksheet>")
        return data

    _patch_zip(path, edit)
    assert "расширенная проверка данных (списки)" in ExcelBook.open(path, Settings()).risks


@pytest.mark.parametrize("name,content,fragment", [
    ("старый.xls", b"\xd0\xcf\x11\xe0whatever", ".xls не поддерживается"),
    ("двоичный.xlsb", b"PK", ".xlsb не поддерживается"),
    ("список.csv", b"a;b", "расширением .xlsx"),
    ("зашифрован.xlsx", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64, "защищён паролем"),
    ("мусор.xlsx", b"this is not a workbook at all", "повреждён"),
])
def test_unusable_files_are_explained(tmp_path, name, content, fragment):
    path = tmp_path / name
    path.write_bytes(content)
    with pytest.raises(ExcelError, match=re.escape(fragment)):
        ExcelBook.open(path, Settings())


def test_missing_file_is_explained(tmp_path):
    with pytest.raises(ExcelError, match="не найден"):
        ExcelBook.open(tmp_path / "нет.xlsx", Settings())


@pytest.mark.skipif(sys.platform.startswith("win"), reason="в Windows занятость определяется блокировкой файла")
def test_excel_owner_file_means_open_in_excel(book_path):
    assert not is_locked(book_path)
    marker = book_path.parent / f"~${book_path.name}"
    marker.write_bytes(b"")
    assert is_locked(book_path)


def test_xlsm_keeps_its_macro_project(tmp_path, book_path):
    macro = tmp_path / "с макросом.xlsm"
    shutil.copy(book_path, macro)
    with zipfile.ZipFile(macro, "a") as archive:
        archive.writestr("xl/vbaProject.bin", b"fake vba project bytes")
    book = _open(macro)
    assert "макросы" in book.risks
    book.ensure_columns(want_result=False)
    book.write_link(2, tmp_path / "x.png")
    book.save()
    book.close()
    with zipfile.ZipFile(macro) as archive:
        assert archive.read("xl/vbaProject.bin") == b"fake vba project bytes"


def test_workbook_written_by_another_program_survives(tmp_path):
    """Книга, созданная не openpyxl (здесь XlsxWriter — по устройству файл близок к тому, что
    пишет сам Excel: общие строки, «умная» таблица, существующие ссылки, даты, формулы)."""
    xlsxwriter = pytest.importorskip("xlsxwriter")
    import datetime as dt
    import xml.etree.ElementTree as ET

    path = tmp_path / "отчет по банкротству.xlsx"
    book = xlsxwriter.Workbook(str(path))
    sheet = book.add_worksheet("восстановление")
    head = book.add_format({"bold": True, "bg_color": "#C6E0B4", "text_wrap": True, "border": 1, "font_size": 9})
    money = book.add_format({"num_format": "#,##0.00"})
    date = book.add_format({"num_format": "dd.mm.yyyy"})
    yellow = book.add_format({"bg_color": "#FFFF00"})
    sheet.merge_range("A1:G1", "Отчёт по банкротству", book.add_format({"bold": True, "font_size": 14}))
    headers = ["жалоба/\nзаявление", "ФИО", "ИИН/БИН", "Ссудный счет", "дата решения суда", "Сумма", "путь файла"]
    people = [("ТЕСТ ПЕРВЫЙ", int(IIN_A)), ("ТЕСТ ВТОРОЙ", IIN_B), ("ТЕСТ ТРЕТИЙ", int(IIN_C)),
              ("ТЕСТ ПЕРВЫЙ", IIN_A), ("ТЕСТ ПЯТЫЙ", IIN_D)]
    for row, (name, iin) in enumerate(people, start=3):
        sheet.write(row, 0, "заявление")
        sheet.write(row, 1, name, yellow if row == 3 else None)
        sheet.write(row, 2, iin)
        sheet.write(row, 3, f"KZ{row:03d}")
        sheet.write_datetime(row, 4, dt.datetime(2026, 10, row), date)
        sheet.write_number(row, 5, 1000.5 * row, money)
        sheet.write_url(row, 6, r"external:W:\Отдел\Банкротство\{}\Снимок.JPG".format(name), string="Снимок.JPG")
    sheet.add_table(2, 0, 2 + len(people), 6, {"columns": [{"header": h, "header_format": head} for h in headers],
                                               "style": "Table Style Medium 2"})
    sheet.write_formula("F9", "=SUM(F4:F8)", money, 0)
    sheet.freeze_panes(3, 0)
    sheet.set_column("B:B", 30)
    sheet.set_column("D:D", None, None, {"hidden": True})
    sheet.conditional_format("F4:F8", {"type": "cell", "criteria": ">", "value": 5000, "format": yellow})
    sheet.data_validation("A4:A8", {"validate": "list", "source": ["жалоба", "заявление"]})
    sheet.write_comment("B4", "проверить адрес")
    book.define_name("Долги", "=восстановление!$F$4:$F$8")
    other = book.add_worksheet("карты")
    other.write_row(0, 0, ["не", "трогать", 42])
    book.close()

    opened = _open(path)
    assert opened.risks == []                                             # обычная книга — пишем на место
    layout = opened.layout
    assert (layout.header_row, layout.iin_col, layout.last_row) == (3, 3, 8)
    tasks = opened.read_tasks()
    assert [t.parse.iin for t in tasks] == [IIN_A, IIN_B, IIN_C, IIN_A, IIN_D]
    layout = opened.ensure_columns(want_result=True)
    assert (layout.link_col, layout.result_col) == (8, 9)                 # сразу за «умной» таблицей
    shots = tmp_path / "снимки"
    shots.mkdir()
    for task in tasks:
        picture = shots / f"{task.parse.iin}.png"
        picture.write_bytes(b"png")
        opened.write_link(task.row, picture)
        opened.write_result(task.row, "НЕ НАЙДЕН")
    opened.backup()
    opened.save()
    opened.close()

    with zipfile.ZipFile(path) as archive:                                # каждая часть файла — корректный XML
        assert archive.testzip() is None
        for name in archive.namelist():
            if name.endswith((".xml", ".rels", ".vml")):
                ET.fromstring(archive.read(name))
    after = load_workbook(path)
    ws = after["восстановление"]
    assert after.sheetnames == ["восстановление", "карты"] and after["карты"]["C1"].value == 42
    assert ws["A1"].value == "Отчёт по банкротству" and "A1:G1" in [str(r) for r in ws.merged_cells.ranges]
    assert [ws.cell(3, c).value for c in range(1, 10)] == headers + [LINK_HEADER, RESULT_HEADER]
    assert ws["B4"].fill.fgColor.rgb.endswith("FFFF00") and ws["B4"].comment.text == "проверить адрес"
    assert ws["C4"].value == int(IIN_A) and ws["C5"].value == IIN_B          # сами ИИН не тронуты
    assert ws["E4"].value == dt.datetime(2026, 10, 3) and ws["E4"].number_format == "dd.mm.yyyy"
    assert ws["F4"].value == 3001.5 and ws["F4"].number_format == "#,##0.00"
    assert ws["F9"].value == "=SUM(F4:F8)"
    assert ws["G4"].hyperlink.target.endswith("Снимок.JPG") and ws["G4"].value == "Снимок.JPG"   # старые ссылки целы
    assert ws["H4"].hyperlink.target == absolute_target(str(shots / f"{IIN_A}.png"))
    assert ws["H4"].hyperlink.target != ws["G4"].hyperlink.target
    assert ws["I8"].value == "НЕ НАЙДЕН"
    assert list(ws.tables.values())[0].ref == "A3:G8"                         # таблица осталась прежней
    assert ws.freeze_panes == "A4" and ws.column_dimensions["D"].hidden
    assert ws.column_dimensions["B"].width == pytest.approx(30, abs=1)
    assert len(ws.conditional_formatting) == 1 and len(ws.data_validations.dataValidation) == 1
    assert "Долги" in after.defined_names
    # Повторное чтение программой: все строки считаются обработанными.
    again = _open(path)
    assert all(task.has_link for task in again.read_tasks())
    assert (again.layout.link_col, again.layout.result_col) == (8, 9)
