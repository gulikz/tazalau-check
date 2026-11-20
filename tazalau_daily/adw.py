"""Шаг 2: клиенты банка с признаком «проблема» (кредиты и кредитные карты) из ADW.

Два пути:
  * выгрузка (xlsx или csv) с обязательным столбцом «ИИН/БИН» и желательным «FIO»: как sql/ADW_пример_выгрузки.csv;
  * запрос к PostgreSQL (нужен psycopg2 или psycopg): файл SQL выполняется до строки «блок 3» (проверки), результатом
    считается последний SELECT. Пароль не хранится в ноутбуке: строка подключения берётся из переменной окружения
    ADW_DSN или вводится при запуске.

Проверки на выходе: число строк и ИИН, неверные ИИН, пустые ФИО, один ИИН с разными ФИО, разбивка по продуктам."""
from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from pathlib import Path

from .common import clean, name_key, norm_iin
from .files import decode_text

IIN_HEAD = re.compile(r"^(иин(/бин)?|бин/иин|iin(/bin)?|client_iin|иин клиента|иин заемщика|иин заёмщика)$", re.I)
FIO_HEAD = re.compile(r"^(fio|фио|ф\.?и\.?о\.?|клиент|заемщик|заёмщик|наименование клиента|фио клиента)$", re.I)
PRODUCT_HEAD = re.compile(r"продукт|product", re.I)
CONTRACT_HEAD = re.compile(r"договор|contract", re.I)
DPD_HEAD = re.compile(r"просрочк|dpd", re.I)
DEBT_HEAD = re.compile(r"задолженност|debt|сумма", re.I)


@dataclass
class Adw:
    source: str
    columns: list[str]
    rows: list[dict]
    iin_col: str
    fio_col: str
    clients: dict = field(default_factory=dict)       # ИИН -> {iin, fio, products, contracts, dpd, debt, rows}
    invalid: list = field(default_factory=list)      # строки с неверным ИИН
    checks: dict = field(default_factory=dict)


def _header_index(header: list[str], pattern: re.Pattern) -> int:
    for i, h in enumerate(header):
        if pattern.search(clean(h)):
            return i
    return -1


def _table_from_rows(raw: list[list], source: str) -> tuple[list[str], list[dict]]:
    """Ищет строку заголовка со столбцом ИИН в первых 30 строках; ниже неё: данные."""
    for hi, line in enumerate(raw[:30]):
        header = [clean(x) for x in line]
        if any(IIN_HEAD.match(h) for h in header):
            break
    else:
        raise ValueError(f"{source}: нет столбца «ИИН/БИН» в первых 30 строках")
    names, seen = [], {}
    for i, h in enumerate(header):
        h = h or f"Столбец {i + 1}"
        if h in seen:
            seen[h] += 1
            h = f"{h} ({seen[h]})"
        else:
            seen[h] = 1
        names.append(h)
    rows = []
    for line in raw[hi + 1:]:
        if not any(x not in (None, "") for x in line):
            continue
        rows.append({names[i]: (line[i] if i < len(line) else None) for i in range(len(names))})
    return names, rows


def read_export(path: Path) -> tuple[list[str], list[dict]]:
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:                     # первый лист, где нашёлся столбец ИИН
                raw = [list(r) for r in ws.iter_rows(values_only=True)]
                try:
                    return _table_from_rows(raw, f"{path.name}, лист «{ws.title}»")
                except ValueError:
                    continue
        finally:
            wb.close()
        raise ValueError(f"{path.name}: ни на одном листе нет столбца «ИИН/БИН»")
    if ext in (".csv", ".txt"):
        text = decode_text(path.read_bytes())
        sample = text[:4000]
        delimiter = max(";,\t", key=sample.count)
        raw = [list(r) for r in csv.reader(io.StringIO(text), delimiter=delimiter)]
        return _table_from_rows(raw, path.name)
    raise ValueError(f"{path.name}: нужен файл .xlsx или .csv")


def run_sql(sql_path: Path, dsn: str, stop_marker: str = "блок 3") -> tuple[list[str], list[dict]]:
    """Выполняет SQL-файл до строки со stop_marker (одним пакетом), возвращает результат последнего SELECT."""
    text = Path(sql_path).read_text(encoding="utf-8")
    lines = text.splitlines()
    cut = next((i for i, line in enumerate(lines) if stop_marker and stop_marker in line and line.lstrip().startswith("--")), len(lines))
    script = "\n".join(lines[:cut])
    try:
        import psycopg2 as driver
    except ImportError:
        try:
            import psycopg as driver
        except ImportError as exc:
            raise ImportError("Для запроса к ADW нужен драйвер PostgreSQL: pip install psycopg2-binary "
                              "(или сохраните выгрузку в xlsx/csv и укажите её файл).") from exc
    conn = driver.connect(dsn)
    try:
        conn.autocommit = False
        with conn.cursor() as cur:
            cur.execute(script)                           # временные функция и представление живут до отката
            last = None
            while True:                                   # psycopg2 видит последний результат, psycopg 3 идёт по всем
                if cur.description is not None:
                    last = ([d[0] for d in cur.description], cur.fetchall())
                try:
                    more = cur.nextset()
                except Exception:
                    more = False
                if not more:
                    break
        conn.rollback()                                   # ничего не сохраняется в базе
    finally:
        conn.close()
    if last is None:
        raise ValueError("SQL не вернул строк (нужен SELECT в конце, до строки «блок 3»)")
    names, data = last
    return names, [dict(zip(names, r)) for r in data]


def prepare(columns: list[str], rows: list[dict], source: str) -> Adw:
    """ИИН в порядке, клиенты (ИИН -> договоры), проверки качества выгрузки."""
    i_iin = _header_index(columns, IIN_HEAD)
    i_fio = _header_index(columns, FIO_HEAD)
    if i_iin < 0:
        raise ValueError(f"{source}: нет столбца «ИИН/БИН»")
    iin_col, fio_col = columns[i_iin], (columns[i_fio] if i_fio >= 0 else "")
    product_col = next((c for c in columns if PRODUCT_HEAD.search(c)), "")
    contract_col = next((c for c in columns if CONTRACT_HEAD.search(c)), "")
    dpd_col = next((c for c in columns if DPD_HEAD.search(c)), "")
    debt_col = next((c for c in columns if DEBT_HEAD.search(c)), "")
    adw = Adw(source, columns, rows, iin_col, fio_col)
    for n, r in enumerate(rows, start=1):
        iin, err = norm_iin(r.get(iin_col))
        if not iin:
            adw.invalid.append({"row": n, "raw": clean(r.get(iin_col)), "err": err or "пусто", "fio": clean(r.get(fio_col)) if fio_col else ""})
            continue
        c = adw.clients.setdefault(iin, {"iin": iin, "fio": "", "fios": set(), "products": [], "contracts": [], "dpd": None, "debt": 0.0, "rows": []})
        fio = clean(r.get(fio_col)) if fio_col else ""
        if fio:
            c["fios"].add(fio)
            c["fio"] = c["fio"] or fio
        if product_col and clean(r.get(product_col)) and clean(r.get(product_col)) not in c["products"]:
            c["products"].append(clean(r.get(product_col)))
        if contract_col and clean(r.get(contract_col)):
            c["contracts"].append(clean(r.get(contract_col)))
        if dpd_col:
            try:
                d = int(float(str(r.get(dpd_col)).replace(",", ".")))
                c["dpd"] = d if c["dpd"] is None else max(c["dpd"], d)
            except (TypeError, ValueError):
                pass
        if debt_col:
            try:
                c["debt"] += float(str(r.get(debt_col)).replace(" ", "").replace("\xa0", "").replace(",", "."))
            except (TypeError, ValueError):
                pass
        c["rows"].append(n)
    by_product: dict[str, set] = {}
    if product_col:
        for c in adw.clients.values():
            for p in c["products"]:
                by_product.setdefault(p, set()).add(c["iin"])
    adw.checks = {
        "строк": len(rows), "ИИН": len(adw.clients), "неверных ИИН": len(adw.invalid),
        "без ФИО": sum(1 for c in adw.clients.values() if not c["fio"]),
        "ИИН с разными ФИО": sum(1 for c in adw.clients.values() if len({name_key(f) or f for f in c["fios"]}) > 1),
        "по продуктам (ИИН)": {p: len(s) for p, s in sorted(by_product.items())},
    }
    return adw


def load(source: Path | str | None = None, *, sql: Path | str | None = None, dsn: str | None = None) -> Adw:
    """Выгрузка из файла (source) или запрос SQL (sql + dsn)."""
    if sql:
        if not dsn:
            raise ValueError("для SQL нужна строка подключения (переменная окружения ADW_DSN)")
        columns, rows = run_sql(Path(sql), dsn)
        return prepare(columns, rows, f"SQL {Path(sql).name}")
    if not source:
        raise ValueError("укажите файл выгрузки ADW или SQL")
    columns, rows = read_export(Path(source))
    return prepare(columns, rows, Path(source).name)


def latest_export(folder: Path, pattern: str = "ADW*") -> Path | None:
    """Самая свежая выгрузка в папке (по времени файла): ADW_2026-11-20.xlsx, ADW_20.11.2026.csv…"""
    files = [p for p in Path(folder).glob(pattern) if p.suffix.lower() in (".xlsx", ".xlsm", ".csv") and not p.name.startswith("~$")]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None
