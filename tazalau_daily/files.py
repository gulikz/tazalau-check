"""Чтение файлов ДГД: таблица Excel, CSV, Word, страница, текст, PDF с текстом. Строки приводятся к одному виду
(лист, номер строки, ячейки), дальше их разбирают parse_announcements (таблица «Объявление о возбуждении производства…»)
и extract_iins (любой список или приказ: ИИН, ФИО, дата, управляющий). Повторяет fileRows, parseAnnouncements и
extractIins закладки-помощника."""
from __future__ import annotations

import csv
import io
import logging
import re
import warnings
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree

from .common import (LO, UP, clean, dates_in, excel_day, iin_ok, max_date, norm_iin, personal_iin, strip_tags,
                     SPACED_IIN, IIN_IN_TEXT)


@dataclass
class Cell:
    v: str            # текст ячейки без лишних пробелов
    num: bool         # в ячейке число (Excel мог потерять ведущие нули ИИН)
    raw: object = None


@dataclass
class Row:
    row: int
    cells: list       # Cell | None по номеру столбца (с нуля)
    sheet: str = ""


@dataclass
class FileRows:
    kind: str
    rows: list[Row] = field(default_factory=list)
    note: str = ""


def _cell(value: object) -> Cell | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return Cell(str(value), False, value)
    if isinstance(value, datetime):
        return Cell(value.strftime("%d.%m.%Y"), False, value)
    if isinstance(value, date):
        return Cell(value.strftime("%d.%m.%Y"), False, value)
    if isinstance(value, (int, float)):
        text = str(int(value)) if float(value).is_integer() else str(value)
        return Cell(text, True, value)
    text = clean(value)
    if text in EXCEL_ERRORS:                    # ошибка формулы в ячейке: значения нет (как у закладки)
        return None
    return Cell(text, False, value) if text else None


EXCEL_ERRORS = {"#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#GETTING_DATA"}


def _text_rows(lines) -> list[Row]:
    out = []
    for i, t in enumerate(lines, start=1):
        t = clean(t)
        if t:
            out.append(Row(i, [Cell(t, False, t)]))
    return out


def decode_text(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def file_rows(path: Path) -> FileRows:
    """Строки файла. Неподдерживаемый формат и скан: строк нет, причина в note."""
    ext = path.suffix.lower().lstrip(".")
    if ext in ("xlsx", "xlsm"):
        # Потоковое чтение: у некоторых файлов ДГД оформление растянуто на миллион пустых строк, и обычное чтение
        # создаёт миллионы ячеек (гигабайты памяти). Размер листа из файла не используется (бывает неверным).
        from openpyxl import load_workbook
        with warnings.catch_warnings():                 # «расширение не поддерживается»: к данным не относится
            warnings.simplefilter("ignore")
            wb = load_workbook(path, read_only=True, data_only=True)
        rows: list[Row] = []
        try:
            for ws in wb.worksheets:
                if not hasattr(ws, "iter_rows"):        # лист-диаграмма
                    continue
                ws.reset_dimensions()
                empty = 0
                for r, values in enumerate(ws.iter_rows(values_only=True), start=1):
                    cells = [_cell(v) for v in values]
                    while cells and cells[-1] is None:
                        cells.pop()
                    if not cells:
                        empty += 1
                        if empty > 5000:                # дальше только оформление без данных
                            break
                        continue
                    empty = 0
                    rows.append(Row(r, cells, ws.title))
        finally:
            wb.close()
        return FileRows("Excel", rows)
    if ext == "xls":
        try:
            import xlrd
        except ImportError:
            return FileRows(".xls", note="старый формат .xls: установите xlrd или сохраните файл как .xlsx")
        book = xlrd.open_workbook(str(path))
        rows = []
        for sh in book.sheets():
            for r in range(sh.nrows):
                cells = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        value = None
                    else:
                        value = cell.value
                    cells.append(_cell(value))
                if any(cells):
                    rows.append(Row(r + 1, cells, sh.name))
        return FileRows("Excel (.xls)", rows)
    data = path.read_bytes()
    if ext == "csv":
        text = decode_text(data)
        try:
            dialect = csv.Sniffer().sniff(text[:4000], delimiters=";,\t")
        except csv.Error:
            dialect = csv.excel
            dialect.delimiter = ";" if text[:4000].count(";") > text[:4000].count(",") else ","
        rows = [Row(i, [_cell(v) for v in line]) for i, line in enumerate(csv.reader(io.StringIO(text), dialect), start=1)]
        return FileRows("CSV", [r for r in rows if any(r.cells)])
    if ext == "txt":
        return FileRows("текст", _text_rows(decode_text(data).splitlines()))
    if ext == "docx":
        return FileRows("Word", _docx_rows(data))
    if ext in ("html", "htm"):
        return FileRows("страница", _html_rows(decode_text(data)))
    if ext == "pdf":
        lines, note = pdf_lines(data)
        rows = _text_rows(lines)
        if not any(re.search(r"\d{6}", r.cells[0].v) for r in rows):
            note = note or "в PDF нет текста (скан): ИИН не прочитаны, нужна проверка вручную"
        return FileRows("PDF", rows, note)
    if ext in ("jpg", "jpeg", "png", "tif", "tiff", "bmp", "gif", "heic"):
        return FileRows("." + ext, note="это картинка (скан): ИИН не прочитаны, нужна проверка вручную")
    if ext in ("doc", "rtf"):
        return FileRows("." + ext, note=f"старый формат .{ext}: откройте в Word и сохраните как .docx")
    return FileRows("." + ext if ext else "без расширения", note="формат не поддерживается")


def _docx_rows(data: bytes) -> list[Row]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        xml = z.read("word/document.xml")
    root = ElementTree.fromstring(xml)
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    body = root.find(W + "body")
    rows: list[Row] = []

    def text_of(el) -> str:
        return " ".join("".join(t.text or "" for t in p.iter(W + "t")) for p in el.iter(W + "p"))

    n = 0
    for el in list(body) if body is not None else []:
        if el.tag == W + "tbl":
            for tr in el.iter(W + "tr"):
                n += 1
                rows.append(Row(n, [_cell(clean(text_of(tc))) for tc in tr.findall(W + "tc")]))
        elif el.tag == W + "p":
            t = clean("".join(x.text or "" for x in el.iter(W + "t")))
            if t:
                n += 1
                rows.append(Row(n, [Cell(t, False, t)]))
    return rows


def _html_rows(text: str) -> list[Row]:
    text = re.sub(r"(?is)<(script|style|noscript)\b.*?</\1>", " ", text)
    rows: list[Row] = []
    n = 0
    for table in re.findall(r"(?is)<table\b.*?</table>", text):
        for tr in re.findall(r"(?is)<tr\b.*?</tr>", table):
            n += 1
            rows.append(Row(n, [_cell(strip_tags(c)) for c in re.findall(r"(?is)<t[dh]\b[^>]*>(.*?)</t[dh]>", tr)]))
    rest = re.sub(r"(?is)<table\b.*?</table>", "\n", text)
    rest = re.sub(r"(?i)<(br|/p|/div|/li|/h\d)\b[^>]*>", "\n", rest)
    for line in rest.split("\n"):
        t = strip_tags(line)
        if t:
            n += 1
            rows.append(Row(n, [Cell(t, False, t)]))
    return rows


def pdf_lines(data: bytes) -> tuple[list[str], str]:
    """Строки текстового слоя PDF. Без pypdf или для скана: пусто и причина."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return [], "PDF не прочитан: установите pypdf (pip install pypdf)"
    logging.getLogger("pypdf").setLevel(logging.ERROR)     # предупреждения о неаккуратных PDF не нужны в журнале
    try:
        reader = PdfReader(io.BytesIO(data))
        lines: list[str] = []
        for page in reader.pages:
            lines.extend((page.extract_text() or "").splitlines())
        return lines, ""
    except Exception as exc:  # повреждённый или зашифрованный PDF
        return [], f"PDF не прочитан: {exc}"


# ---------------------------------------------------------------- «Объявление о возбуждении производства…»
_IIN_HEAD = re.compile(r"идентификацион|(^|[^а-яё])иин([^а-яё]|$)|жсн|жеке\s+сәйкестендіру|сәйкестендіру\s+нөмір")
_FIO_HEAD = re.compile(r"фамили|фио|ф\.\s*и\.\s*о|аты-?\s*жөн|тегі")
_MGR_HEAD = re.compile(r"управляющ|басқарушы")
_RUS_HEAD = re.compile(r"фамили|фио|(^|[^а-яё])иин([^а-яё]|$)|должник")


def _parse_sheet(rows: list[Row]) -> dict | None:
    """Одна таблица объявлений (один лист)."""
    hi = -1
    for i, r in enumerate(rows[:40]):
        t = [c.v.lower() for c in r.cells if c]
        if len(t) >= 3 and any(_IIN_HEAD.search(x) for x in t) and any(re.search(r"фамили|фио|ф\.\s*и\.\s*о|аты|тегі", x) for x in t):
            hi = i
            break
    if hi < 0:
        return None
    head = rows[hi].cells
    cols = len(head) + 2

    def text(i: int) -> str:
        return head[i].v.lower() if 0 <= i < len(head) and head[i] else ""

    role: dict[str, int] = {}
    used: set[int] = set()

    def pick(key: str, test) -> None:
        for i in range(cols):
            if i not in used and text(i) and test(text(i)):
                role[key] = i
                used.add(i)
                return

    pick("n", lambda t: re.match(r"^\s*(№|n\s*п|п/п)", t))
    pick("iin", lambda t: _IIN_HEAD.search(t) and not _MGR_HEAD.search(t))
    pick("fio", lambda t: _FIO_HEAD.search(t) and not _MGR_HEAD.search(t))
    pick("manager", lambda t: _FIO_HEAD.search(t) and _MGR_HEAD.search(t))
    pick("claimsAddress", lambda t: re.search(r"адрес|мекен\s*-?\s*жай", t) and re.search(r"прием|приём|қабылда", t))
    pick("address", lambda t: re.search(r"адрес|мекен\s*-?\s*жай|орналасқан\s+жер", t))
    pick("court", lambda t: re.search(r"суд|сот", t) and re.search(r"наименован|атау", t))
    pick("ruling", lambda t: re.search(r"определени|ұйғарым|вынесени|шешім", t))
    pick("claims", lambda t: re.search(r"срок|мерзім", t) and re.search(r"требован|талап", t))
    pick("contacts", lambda t: re.search(r"контакт|байланыс|телефон", t))
    pick("posted", lambda t: re.search(r"размещени|орналастыр|опубликов|жариял", t))
    if "manager" not in role:
        pick("manager", lambda t: _FIO_HEAD.search(t))           # второй столбец с ФИО: управляющий
    if "claims" in role and not text(role["claims"] + 1) and role["claims"] + 1 not in used:
        role["to"] = role["claims"] + 1                          # «срок приёма требований»: «с» и «до» рядом
        used.add(role["to"])
    if "iin" not in role:
        return None
    start = hi + 1
    while start < len(rows) and start <= hi + 3:                # «с / до» и строка с номерами столбцов
        vals = [clean(c.v) for c in rows[start].cells if c]
        if vals and (all(re.fullmatch(r"\d{1,2}", v) for v in vals)
                     or all(re.fullmatch(r"(с|до|от|по|бастап|дейін|\d{1,2})", v, re.I) for v in vals)):
            start += 1
        else:
            break

    def cell(r: Row, key: str):
        i = role.get(key)
        return r.cells[i] if i is not None and i < len(r.cells) else None

    def val(r: Row, key: str) -> str:
        c = cell(r, key)
        return clean(c.v) if c else ""

    def day(c) -> date | None:
        if not c or not clean(c.v):
            return None
        return excel_day(c.raw if c.raw is not None else c.v)

    records, invalid = [], []
    for r in rows[start:]:
        raw = cell(r, "iin")
        iin, err = norm_iin(raw.raw if raw is not None and raw.num else (raw.v if raw else None))
        if not iin:
            for c in r.cells:
                m = c and IIN_IN_TEXT.search(c.v)
                if m and iin_ok(m.group(1)):
                    iin = m.group(1)
                    break
        fio = val(r, "fio")
        if not iin and not fio:
            continue
        if not iin and err:
            invalid.append({"where": (r.sheet + ", " if r.sheet else "") + f"строка {r.row}", "raw": raw.v if raw else "", "err": err})
        records.append({"n": val(r, "n"), "fio": fio, "iin": iin, "address": val(r, "address"), "court": val(r, "court"),
                        "ruling": day(cell(r, "ruling")), "manager": val(r, "manager"), "from": day(cell(r, "claims")),
                        "to": day(cell(r, "to")), "claimsAddress": val(r, "claimsAddress"), "contacts": val(r, "contacts"),
                        "posted": day(cell(r, "posted")), "row": r.row, "sheet": r.sheet})
    russian = any(_RUS_HEAD.search(text(i)) for i in range(len(head)))
    return {"records": records, "invalid": invalid, "roles": role, "headerRow": rows[hi].row, "russian": russian}


def parse_announcements(rows: list[Row]) -> dict | None:
    """Таблица объявлений о возбуждении дел (строка на должника): столбцы узнаются по смыслу заголовка
    (у регионов разный порядок, русский или казахский). Каждый лист отдельно; лист-перевод (казахский и русский
    с теми же должниками) не удваивает строки: повтор (тот же ИИН или ФИО, те же даты) берётся один раз, русский лист
    первым. None: это не такая таблица."""
    sheets: dict[str, list[Row]] = {}
    for r in rows:
        sheets.setdefault(r.sheet, []).append(r)
    parsed = [p for p in (_parse_sheet(part) for part in sheets.values()) if p]
    if not parsed:
        return None
    parsed.sort(key=lambda p: not p["russian"])                  # русский лист первым (стабильно для равных)
    records, invalid, seen = [], [], set()
    for p in parsed:
        bad_rows = {x["where"] for x in p["invalid"]}
        for rec in p["records"]:
            key = (rec["iin"] or clean(rec["fio"]).lower(), rec["ruling"], rec["posted"])
            if key in seen:
                bad_rows.discard((rec["sheet"] + ", " if rec["sheet"] else "") + f"строка {rec['row']}")
                continue
            seen.add(key)
            records.append(rec)
        invalid += [x for x in p["invalid"] if x["where"] in bad_rows]
    first = parsed[0]
    return {"records": records, "invalid": invalid, "roles": first["roles"], "headerRow": first["headerRow"], "sheets": len(parsed)}


# ---------------------------------------------------------------- любой список или приказ
NAME_RE = re.compile(f"^[{UP}][{UP}{LO}-]+(?:\\s+[{UP}](?:[{UP}{LO}-]+|\\.)){{1,3}}$")
NAME_IN_TEXT = re.compile(f"[{UP}][{LO}-]+(?:\\s+[{UP}][{LO}-]+){{1,3}}|[{UP}]{{2,}}(?:\\s+[{UP}]{{2,}}){{1,3}}")


def extract_iins(rows: list[Row]) -> dict:
    """ИИН людей в строках файла. Заголовок (строка со словом «ИИН») подсказывает, где ФИО, дата и управляющий."""
    head = None
    for r in rows[:80]:
        if len(r.cells) > 1 and any(c and re.search(r"иин|iin", c.v, re.I) and not re.search(r"\d{6}", c.v) for c in r.cells):
            head = r
            break
    role: list[str] = []
    if head is not None:
        for c in head.cells:
            t = c.v.lower() if c else ""
            if not t:
                role.append("")
                continue
            mgr = "управляющ" in t
            if re.search(r"иин|iin", t):
                role.append("mgrIin" if mgr else "iin")
            elif mgr:
                role.append("mgr")
            elif re.search(r"фио|ф\.\s*и\.\s*о|должник|граждан|заем|заём|наименован|фамили", t):
                role.append("x" if "fio" in role else "fio")
            elif "дата" in t:
                role.append("x" if "date" in role else "date")
            else:
                role.append("")
    items, invalid = [], []
    for r in rows:
        if r is head:
            continue
        found: dict[str, int] = {}
        for i, c in enumerate(r.cells):
            if not c or (i < len(role) and role[i] == "mgrIin"):
                continue
            before = len(found)
            whole, err = norm_iin(c.raw if c.num else c.v)
            if whole and personal_iin(whole):
                found.setdefault(whole, i)
            for m in IIN_IN_TEXT.finditer(c.v):
                x = m.group(1)
                if iin_ok(x) and personal_iin(x):
                    found.setdefault(x, i)
            if len(found) == before:
                for m in SPACED_IIN.finditer(c.v):
                    x = re.sub(r"\D", "", m.group(0))
                    if iin_ok(x) and personal_iin(x):
                        found.setdefault(x, i)
            if len(found) == before:                  # в PDF ИИН слился со следующим числом: «ИИН 87091630164224000»
                for m in re.finditer(r"(?:ИИН|ЖСН)\D{0,12}?(\d{13,})", c.v, re.I):
                    x = m.group(1)[:12]
                    if iin_ok(x) and personal_iin(x):
                        found.setdefault(x, i)
            if len(found) == before and err and ((i < len(role) and role[i] == "iin") or (head is None and re.fullmatch(r"\d{12}", c.v))):
                invalid.append({"where": (r.sheet + ", " if r.sheet else "") + f"строка {r.row}", "raw": c.v, "err": err})
        if not found:
            continue

        def by(want: str) -> int:
            for i, c in enumerate(r.cells):
                if c and i < len(role) and role[i] == want:
                    return i
            return -1

        fio_col, mgr_col, date_col = by("fio"), by("mgr"), by("date")
        fio = r.cells[fio_col].v if fio_col >= 0 else ""
        if not fio:
            fio = next((c.v for i, c in enumerate(r.cells) if c and i != mgr_col and NAME_RE.match(c.v)), "")
        if not fio and len(r.cells) == 1:                    # строка текста (PDF, Word): имя прямо перед ИИН
            at = re.search(r"\d{4}", r.cells[0].v)
            names = NAME_IN_TEXT.findall(r.cells[0].v[:at.start()] if at else "")
            fio = names[-1] if names else ""
        if date_col >= 0:
            c = r.cells[date_col]
            row_dates = dates_in(c.v) + ([excel_day(c.raw)] if c.num and excel_day(c.raw) else [])   # текст или дата Excel
        else:
            row_dates = dates_in(" ".join(c.v for c in r.cells if c))
        for iin in found:
            items.append({"iin": iin, "fio": fio[:120], "mgr": r.cells[mgr_col].v[:120] if mgr_col >= 0 else "",
                          "date": max_date(row_dates), "where": (f"лист «{r.sheet}», " if r.sheet else "") + f"строка {r.row}"})
    return {"items": items, "invalid": invalid}
