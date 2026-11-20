"""Папка «ДГД»: публикации, собранные с gov.kz (папка региона с «публикации_gov_kz.json»), и файлы, положенные вручную.
Читает все файлы, находит ИИН (таблицы объявлений построчно, приказы и реестры по тексту, сканы распознаванием),
проверяет каждый из 20 регионов и пишет свод «ДГД_свод.xlsx». Повторяет readDgd, dgdRegions и dgdWorkbook закладки."""
from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from . import ocr
from .common import REGION_NAMES, UNKNOWN_REGION, dates_in, dmy, max_date, regions_in
from .files import extract_iins, file_rows, parse_announcements
from .govkz import ANN_TITLE_RE, DGD_DIR, DGD_META, write_json
from .xlsx import X, file_link, write_book

DGD_FILE = "ДГД_свод.xlsx"
DGD_OCR = "распознано.json"        # ИИН, прочитанные со сканов (общий с закладкой)
DGD_FEW = 2                        # столько ИИН и меньше в регионе: подозрительно мало
DGD_SITE_STALE = 14                # последнее объявление региона старше стольких дней: проверить
DGD_STALE_DAYS = 3                 # файлы вручную: самая свежая дата старше стольких дней
SKIP_SUFFIX = (".crswap", ".tmp", ".ini", ".db", ".part")


def _day(value) -> date | None:
    if not value:
        return None
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _when(value) -> datetime | None:
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")) if value else None
    except ValueError:
        return None


def read_dgd(work_dir: Path, *, use_ocr: bool = True, log: Callable[[str], None] = print) -> dict | None:
    work_dir = Path(work_dir)
    root = work_dir / DGD_DIR
    if not root.is_dir():
        return None
    files, metas = [], []

    def walk(d: Path, depth: int) -> None:
        for p in sorted(d.iterdir(), key=lambda x: x.name):
            if p.is_dir():
                if depth < 2:
                    walk(p, depth + 1)
                continue
            if p.name.startswith(("~$", ".", "~tmp_")) or p.name.lower().endswith(SKIP_SUFFIX) or p.name == DGD_OCR:
                continue
            if p.name == DGD_META:
                try:
                    data = json.loads(p.read_text(encoding="utf-8"))
                except (OSError, ValueError):
                    continue
                if data.get("region"):
                    metas.append({"dir": p.parent, "folder": p.parent.relative_to(root).as_posix(), "data": data})
                continue
            files.append({"name": p.name, "path": p.relative_to(work_dir).as_posix(), "full": p})

    walk(root, 0)
    cache_path = root / DGD_OCR
    try:
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        cache = {}
    changed = [False]

    def scan(key: str, full: Path) -> dict:
        """ИИН со скана: из памяти (если файл не менялся) или распознаванием."""
        size = full.stat().st_size
        hit = cache.get(key)
        if hit and hit.get("size") == size and isinstance(hit.get("iins"), list):
            return hit
        if not use_ocr:
            return {"iins": [], "note": "скан: ИИН не прочитан, нужна проверка вручную"}
        found, note = ocr.read_scan(full)
        cache[key] = {"size": size, "at": datetime.now().astimezone().isoformat(timespec="seconds"), "note": note,
                      "iins": [{"iin": x, "conf": "medium", "how": "распознавание скана (Tesseract)", "fixes": 0} for x in found]}
        changed[0] = True
        return cache[key]

    out = {"files": [], "items": [], "orders": [], "sites": [], "announcements": [], "at": datetime.now()}
    used: set[str] = set()
    for m in sorted(metas, key=lambda x: x["folder"]):
        meta, region = m["data"], m["data"]["region"]
        site = {"region": region, "state": meta.get("state") or "ok", "note": meta.get("note") or "",
                "collectedAt": _when(meta.get("collectedAt")), "failedAt": _when(meta.get("failedAt")), "pages": meta.get("pages") or [],
                "rows": [], "via": meta.get("via") or "", "source": meta.get("version") or "",
                "path": "; ".join(dict.fromkeys(x.get("path") for x in meta.get("rows") or [] if ANN_TITLE_RE.search(x.get("text") or "") and x.get("path")))}
        for r in meta.get("rows") or []:
            row = {"region": region, "date": _day(r.get("date")), "kind": r.get("kind") or "", "debtor": r.get("debtor") or "",
                   "manager": r.get("manager") or "", "number": r.get("number") or "", "text": r.get("text") or "", "page": r.get("page") or "",
                   "firstSeen": r.get("firstSeen") or "", "files": [], "iins": [], "notes": list(r.get("notes") or []), "path": r.get("path") or "",
                   "annRecords": 0, "annInvalid": 0, "tableRecords": 0}
            for f in r.get("files") or []:
                file = {"url": f.get("url") or "", "name": f.get("name") or "", "saved": f.get("saved") or "", "err": f.get("err") or "", "path": ""}
                row["files"].append(file)
                if not file["saved"]:
                    continue
                full = m["dir"] / file["saved"]
                file["path"] = full.relative_to(work_dir).as_posix()
                used.add(file["path"])
                if not full.exists():
                    row["notes"].append("файла нет в папке")
                    continue
                try:
                    parsed = file_rows(full)
                except Exception as exc:
                    row["notes"].append(f"файл не прочитан: {exc}")
                    continue
                if full.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                    ann = parse_announcements(parsed.rows)
                    if ann and ann["records"]:          # таблица ДГД: строка на должника
                        # о возбуждении дела (есть срок приёма требований или так названа): в общий лист объявлений;
                        # о применении, завершении процедуры, признании банкротом: только ИИН с видом из названия
                        initiation = "claims" in ann["roles"] or bool(ANN_TITLE_RE.search(row["text"]))
                        kind = "объявление о возбуждении дела" if initiation else (row["kind"] if row["kind"] != "публикация" else "таблица ДГД")
                        if initiation:
                            row["annRecords"] += len(ann["records"])
                            row["annInvalid"] += len(ann["invalid"])
                        else:
                            row["tableRecords"] = row.get("tableRecords", 0) + len(ann["records"])
                        for rec in ann["records"]:
                            if initiation:
                                out["announcements"].append({"region": region, "page": row["page"], "path": file["path"], "title": row["text"], **rec})
                            if rec["iin"] and not any(y["iin"] == rec["iin"] for y in row["iins"]):
                                row["iins"].append({"iin": rec["iin"], "conf": "high", "how": "таблица ДГД", "path": file["path"], "rec": rec, "kind": kind})
                        continue
                found = extract_iins(parsed.rows)["items"]
                several = len({x["iin"] for x in found}) > 1 and "реестр требований" in row["kind"]
                for n_found, x in enumerate(found):
                    if not any(y["iin"] == x["iin"] for y in row["iins"]):
                        # реестр требований: по одному должнику: первый ИИН (в шапке) его, следующие ниже по тексту
                        # обычно кредитор-гражданин или управляющий: такие помечаются «сверить»
                        later = several and any(y["iin"] != x["iin"] for y in found[:n_found])
                        row["iins"].append({"iin": x["iin"], "conf": "medium" if later else "high",
                                            "how": "текст файла" + (" (ниже по тексту: возможно, кредитор или управляющий)" if later else ""),
                                            "path": file["path"], "later": later})
                if not found and (parsed.note and "скан" in parsed.note or full.suffix.lower() in (".jpg", ".jpeg", ".png")):
                    res = scan(f"{m['folder']}/{file['saved']}", full)
                    if res.get("note") and not res.get("iins"):
                        row["notes"].append(res["note"])
                    for x in res.get("iins") or []:
                        if not any(y["iin"] == x["iin"] for y in row["iins"]):
                            row["iins"].append({"iin": x["iin"], "conf": x.get("conf") or "medium", "how": x.get("how") or "скан", "path": file["path"]})
                elif not found and parsed.note:
                    row["notes"].append(parsed.note)
            for iin in r.get("titleIins") or []:
                if not any(y["iin"] == iin for y in row["iins"]):
                    row["iins"].append({"iin": iin, "conf": "high", "how": "название публикации", "path": ""})
            site["rows"].append(row)
            out["orders"].append(row)
            for x in row["iins"]:
                rec = x.get("rec")
                if rec:
                    initiation = x["kind"] == "объявление о возбуждении дела"
                    where = (x["kind"] + (f", {'определение' if initiation else 'решение'} суда от {dmy(rec['ruling'])}" if rec["ruling"] else "")
                             + (f", размещено {dmy(rec['posted'])}" if rec["posted"] else "") + (f", требования до {dmy(rec['to'])}" if rec["to"] else ""))
                    out["items"].append({"iin": x["iin"], "fio": rec["fio"], "mgr": rec["manager"], "date": rec["posted"] or rec["ruling"] or row["date"],
                                         "listDate": rec["posted"] or row["date"], "region": region, "path": x["path"] or row["page"],
                                         "kind": x["kind"], "conf": "high", "how": x["how"], "page": row["page"],
                                         "court": rec["court"], "contacts": rec["contacts"], "claimsTo": rec["to"], "where": where})
                else:
                    out["items"].append({"iin": x["iin"], "fio": row["debtor"], "mgr": row["manager"], "date": row["date"], "listDate": row["date"],
                                         "region": region, "path": x["path"] or row["page"], "kind": row["kind"], "conf": x["conf"], "how": x["how"],
                                         "page": row["page"], "where": (row["kind"] or "публикация") + (f" № {row['number']}" if row["number"] else "")
                                         + (f" от {dmy(row['date'])}" if row["date"] else "")})
        # Один и тот же должник в нескольких файлах региона (обновлённая таблица рядом с прежней, лист-перевод):
        # в общем списке объявлений один раз: из самой свежей публикации
        mine = [a for a in out["announcements"] if a["region"] == region]
        if mine:
            order = {r["page"]: i for i, r in enumerate(sorted(site["rows"], key=lambda r: r["date"] or date.min, reverse=True))}
            mine.sort(key=lambda a: order.get(a["page"], len(order)))
            seen, keep = set(), []
            for a in mine:
                key = (a["iin"] or a["fio"].lower(), a["ruling"], a["posted"])
                if key not in seen:
                    seen.add(key)
                    keep.append(a)
            site["annDuplicates"] = len(mine) - len(keep)
            out["announcements"] = [a for a in out["announcements"] if a["region"] != region] + keep
        out["sites"].append(site)

    for f in files:                                    # файлы, положенные вручную
        if f["path"] in used:
            continue
        rec = {"path": f["path"], "name": f["name"], "region": "", "kind": "", "listDate": None, "lastDate": None, "count": 0, "unique": 0,
               "invalid": [], "note": "", "failed": False, "dateNote": ""}
        from_path = regions_in(f["path"][len(DGD_DIR) + 1:], True)
        if len(from_path) == 1:
            rec["region"] = from_path[0]
        try:
            parsed = file_rows(f["full"])
            rec["kind"], rec["note"] = parsed.kind, parsed.note
            ex = extract_iins(parsed.rows)
            rec["invalid"] = ex["invalid"]
            head = " ".join(" ".join(c.v for c in r.cells if c) for r in parsed.rows[:40])
            regs = from_path or regions_in(head)
            rec["region"] = regs[0] if len(regs) == 1 else ""
            if len(regs) > 1:
                rec["note"] = (rec["note"] + "; " if rec["note"] else "") + "несколько регионов в названии: " + ", ".join(regs) + " (переименуйте файл)"
            rec["listDate"] = max_date(dates_in(f["path"]))
            items = ex["items"]
            if not items and ("скан" in (parsed.note or "") or "картинка" in (parsed.note or "")):
                res = scan(f["path"][len(DGD_DIR) + 1:], f["full"])
                items = [{"iin": x["iin"], "fio": "", "mgr": "", "date": rec["listDate"], "conf": x.get("conf") or "medium",
                          "how": x.get("how") or "скан", "where": (x.get("how") or "скан") + " (сверить с приказом)"} for x in res.get("iins") or []]
                rec["note"] = ("скан: ИИН распознан, сверить с приказом" if items
                               else "скан: ИИН не распознан, нужна проверка вручную" + (f" ({res['note']})" if res.get("note") and "не распознан" not in res["note"] else ""))
            rec["lastDate"] = max_date(x["date"] for x in items)
            rec["listDate"] = rec["listDate"] or rec["lastDate"]
            if not rec["listDate"]:
                rec["listDate"] = date.fromtimestamp(f["full"].stat().st_mtime)
                rec["dateNote"] = "дата списка не найдена: взято время файла"
            rec["count"] = len(items)
            rec["unique"] = len({x["iin"] for x in items})
            for x in items:
                out["items"].append({"region": rec["region"], "listDate": rec["listDate"], "path": f["path"], "conf": "high", "how": "текст файла",
                                     "kind": "", "page": "", **x})
            if not items and not rec["note"] and parsed.rows:
                rec["note"] = "ИИН в файле не найдены"
            if not parsed.rows and not rec["note"]:
                rec["note"] = "файл пустой"
        except Exception as exc:
            rec["note"], rec["failed"] = f"не прочитан: {exc}", True
        out["files"].append(rec)
    if changed[0]:
        try:
            write_json(cache_path, cache)
        except OSError:
            pass
    out["by_iin"] = {}
    for x in out["items"]:
        out["by_iin"].setdefault(x["iin"], []).append(x)
    out["regions"] = audit_regions(out)
    return out


def audit_regions(dg: dict, today: date | None = None) -> list[dict]:
    """Проверка всех 20 регионов (пустые тоже видны): собран ли gov.kz, есть ли объявления о возбуждении дел за год,
    сколько строк и ИИН, свежее ли последнее размещение, всё ли прочитано. Уровень: bad, warn, good."""
    today = today or date.today()
    year = today.year

    def days(d: date) -> int:
        return (today - d).days

    names = REGION_NAMES + ([UNKNOWN_REGION] if any(not f["region"] for f in dg["files"]) else [])
    out = []
    for name in names:
        files = [f for f in dg["files"] if (f["region"] or UNKNOWN_REGION) == name]
        items = [x for x in dg["items"] if (x.get("region") or UNKNOWN_REGION) == name]
        file_items = [x for x in items if not x.get("page")]
        site = next((s for s in dg["sites"] if s["region"] == name), None)
        problems: list[str] = []
        level = ["good"]

        def bad(t: str) -> None:
            if t:
                problems.append(t)
            level[0] = "bad"

        def warn(t: str) -> None:
            problems.append(t)
            if level[0] != "bad":
                level[0] = "warn"

        r = {"name": name, "files": files, "items": len(items), "unique": len({x["iin"] for x in items}),
             "invalid": sum(len(f["invalid"]) for f in files), "listDate": max_date(f["listDate"] for f in files),
             "lastDate": max_date(x.get("date") for x in file_items), "site": site, "year": 0, "win": 0, "annRows": 0,
             "annUnique": 0, "annPosted": None, "path": ""}
        if site:
            anns = [x for x in dg["announcements"] if x["region"] == name]
            # объявление о возбуждении дела: по названию или по содержимому (г. Алматы кладёт таблицу объявлений
            # в «Информационное сообщение ФЛ 2026»)
            ann_docs = [x for x in site["rows"] if ANN_TITLE_RE.search(x["text"]) or x.get("annRecords")]
            # ИИН объявлений: строки таблиц и ИИН из объявлений-документов (Word или PDF на одного должника)
            ann_iins = {x["iin"] for x in anns if x["iin"]} | {y["iin"] for x in ann_docs for y in x["iins"]}
            uniq = len(ann_iins)
            posted = max_date(x["posted"] for x in anns if x["posted"] and x["posted"] <= today)
            invalid_ann = sum(x.get("annInvalid") or 0 for x in site["rows"])
            future = sum(1 for x in anns if x["posted"] and x["posted"] > today)
            unread = sum(1 for x in site["rows"] if not x["iins"] and any(f["saved"] for f in x["files"]))
            r.update(win=len(site["rows"]), year=len(ann_docs), annRows=len(anns), annUnique=uniq, annPosted=posted, path=site["path"], unread=unread)
            when = f" (собрано {dmy(site['collectedAt'])})" if site["collectedAt"] else ""
            if site["state"] == "skip":
                bad("gov.kz не ответил" + (f": {site['note']}" if site["note"] else "") + (f"; данные прошлого сбора{when}" if site["rows"] else ""))
            elif site["state"] != "ok":
                bad((site["note"] or "сбор не удался") + (f"; данные прошлого сбора{when}" if site["rows"] else ""))
            if site["state"] != "ok":
                pass
            elif not site["rows"]:
                bad(f"публикаций о гражданах за {year} год на gov.kz не найдено")
            elif not ann_docs:
                warn(f"объявления о возбуждении дел за {year} год нет (есть другие публикации: {len(site['rows'])})")
            elif not anns and not ann_iins:
                bad("в объявлениях о возбуждении дел 0 строк (файл не скачался или не прочитан)")
            elif uniq <= DGD_FEW:
                warn(f"всего {uniq} ИИН в объявлениях: проверить на gov.kz")
            if posted and days(posted) > DGD_SITE_STALE:
                warn(f"последнее объявление размещено {dmy(posted)}: {days(posted)} дн. назад")
            if future:
                warn(f"дата размещения в будущем: {future} (ошибка в файле ДГД, сверить)")
            if invalid_ann:
                warn(f"ИИН с ошибкой в таблицах объявлений: {invalid_ann} (так в файле ДГД)")
            if unread:
                warn(f"ИИН не прочитан в файлах: {unread} (открыть файл, сверить вручную)")
            nofile = sum(1 for x in site["rows"] if not any(f["saved"] for f in x["files"]) and not x["iins"])
            if nofile:
                warn(f"без файла и без ИИН в названии: {nofile}")
            doubt = sum(1 for x in site["rows"] if any(y["conf"] != "high" and not y.get("later") for y in x["iins"]))
            if doubt:
                warn(f"ИИН прочитан со скана неуверенно: {doubt} (сверить)")
            later = sum(1 for x in site["rows"] if any(y.get("later") for y in x["iins"]))
            if later:
                warn(f"в реестре требований больше одного ИИН: {later} (второй, возможно, кредитор или управляющий: сверить)")
        if files or not site:
            newest = max_date([r["listDate"], r["lastDate"]])
            age = days(newest) if newest else None
            file_unique = len({x["iin"] for x in file_items})
            if not files:
                bad("" if site else "нет списка")
            elif all(f["failed"] for f in files):
                bad("файл не прочитан")
            elif not file_unique:
                note = next((f["note"] for f in files if f["note"]), "")
                bad("0 ИИН" + (f": {note}" if note else ""))
            elif file_unique <= DGD_FEW and not site:
                warn(f"всего {file_unique} ИИН: проверить вручную")
            if files and age is not None and age > DGD_STALE_DAYS:
                warn(f"самая свежая дата {dmy(newest)}: {age} дн. назад")
            if any(f["dateNote"] for f in files):
                warn("дата списка не найдена")
            if r["invalid"]:
                warn(f"неверных ИИН: {r['invalid']}")
            unread_files = sum(1 for f in files if f["failed"] or any(w in (f["note"] or "") for w in ("не прочитан", "не распознан", "старый формат", "не поддерживается")))
            if unread_files and len(files) > unread_files:
                warn(f"не прочитано файлов: {unread_files}")
            if name == UNKNOWN_REGION:
                warn("укажите регион в имени файла")
        r["level"] = level[0]
        r["text"] = "; ".join(p for p in problems if p) or "в порядке"
        out.append(r)
    return out


def write_dgd_book(dg: dict, work_dir: Path, path: Path | None = None) -> Path:
    """ДГД_свод.xlsx: «Регионы» (проверка каждого региона), «Объявления ГГГГ» (все таблицы объявлений одной таблицей),
    «Публикации», «ИИН», «Файлы» (положенные вручную)."""
    work_dir = Path(work_dir)
    path = Path(path) if path else work_dir / DGD_FILE
    year = date.today().year

    def flink(rel: str) -> str:
        return file_link(work_dir / rel, path) if rel else ""

    def link(v, to: str) -> X:
        return X(v, link=to) if to else X(v)

    regions = [["Регион", "Проверка", "gov.kz: сбор", "Где на gov.kz", "Объявлений о возбуждении (файлов)", "Строк в объявлениях",
                "Уникальных ИИН", "Последнее размещение", "Других публикаций", "Файлов вручную", "Уникальных ИИН (файлы вручную)",
                "Неверных ИИН (файлы вручную)", "Дата списка (файлы)", "Файлы"]]
    for r in dg["regions"]:
        s = r["site"]
        regions.append([r["name"], X(r["text"], r["level"]),
                        (dmy(s["collectedAt"]) if s["state"] == "ok" else "не открылся " + dmy(s["failedAt"] or s["collectedAt"])) if s else "нет",
                        (s["path"] or s["via"]) if s else "", r["year"] if s else "", r["annRows"] if s else "", r["annUnique"] if s else "",
                        r["annPosted"] if s else "", (r["win"] - r["year"]) if s else "", len(r["files"]),
                        len({x["iin"] for x in dg["items"] if (x.get("region") or UNKNOWN_REGION) == r["name"] and not x.get("page")}),
                        r["invalid"], r["listDate"], "; ".join(f["path"] for f in r["files"])])
    anns = sorted(dg["announcements"], key=lambda x: (x["region"], -(x["posted"] or date.min).toordinal()))
    ann_rows = [["Регион", "№", "ФИО должника", "ИИН должника", "Адрес должника", "Суд", "Дата определения суда", "Финансовый управляющий",
                 "Приём требований с", "Приём требований до", "Адрес приёма требований", "Контакты управляющего", "Дата размещения",
                 "Публикация на gov.kz", "Файл"]]
    for x in anns:
        ann_rows.append([x["region"], x["n"], x["fio"], x["iin"] or X("нет ИИН", "warn"), x["address"], x["court"], x["ruling"], x["manager"],
                         x["from"], x["to"], x["claimsAddress"], x["contacts"], x["posted"], link(x["title"] or x["page"], x["page"]),
                         link(x["path"][len(DGD_DIR) + 1:], flink(x["path"]))])

    def conf(x: dict) -> str:
        return "" if x.get("conf") in ("high", None, "") else " (сверить)"

    orders = [["Регион", "Дата публикации", "Вид", "Должник (как в названии)", "ИИН", "Как прочитан ИИН", "Где на gov.kz", "Файл",
               "Публикация на gov.kz", "Впервые увиден", "Название публикации"]]
    for o in sorted(dg["orders"], key=lambda o: -(o["date"] or date.min).toordinal()):
        f = next((x for x in o["files"] if x["saved"]), o["files"][0] if o["files"] else {})
        if o["annRecords"]:
            iin_cell, how = X(f"строк: {o['annRecords']}, ИИН: {len(o['iins'])} (лист «Объявления {year}»)"), "таблица объявлений"
        elif o["tableRecords"]:
            iin_cell, how = X(f"строк: {o['tableRecords']}, ИИН: {len(o['iins'])} (лист «ИИН»)"), "таблица ДГД"
        elif o["iins"]:
            iin_cell = X(", ".join(x["iin"] for x in o["iins"]), "warn" if any(x["conf"] != "high" for x in o["iins"]) else "")
            how = "; ".join(x["how"] + conf(x) for x in o["iins"])
        else:
            iin_cell = X("не прочитан", "warn") if f.get("saved") else X("")
            how = "; ".join(o["notes"]) or f.get("err") or ("" if f.get("saved") else "файла нет")
        file_cell = link(f["path"][len(DGD_DIR) + 1:], flink(f["path"])) if f.get("saved") else link(f.get("url") or "", f.get("url") or "")
        orders.append([o["region"], o["date"], o["kind"], o["debtor"], iin_cell, how, o["path"], file_cell, link(o["page"], o["page"]),
                       o["firstSeen"], o["text"]])
    iins = [["ИИН", "ФИО", "Регион", "Дата", "Финансовый управляющий", "Дата списка", "Файл", "Где"]]
    for x in dg["items"]:
        target = x["path"] if x["path"].startswith("http") else flink(x["path"])
        iins.append([x["iin"], x.get("fio") or "", x.get("region") or "не определён", x.get("date"), x.get("mgr") or "", x.get("listDate"),
                     link(x["path"], target), x.get("where", "") + conf(x)])
    files = [["Файл", "Регион", "Формат", "Дата списка", "Строк с ИИН", "Уникальных ИИН", "Неверных ИИН", "Примечание"]]
    for f in dg["files"]:
        note = "; ".join([t for t in (f["note"], f["dateNote"]) if t] + [f"{x['where']}: {x['raw']} ({x['err']})" for x in f["invalid"][:5]])
        files.append([link(f["path"], flink(f["path"])), f["region"] or "не определён", f["kind"], f["listDate"], f["count"], f["unique"],
                      len(f["invalid"]), X(note, "warn" if f["failed"] or (f["note"] and not f["note"].endswith("сверить с приказом")) else "")])
    return write_book(path, [
        {"name": "Регионы", "rows": regions, "widths": [30, 60, 16, 60, 14, 12, 12, 14, 12, 10, 14, 12, 14, 50]},
        {"name": f"Объявления {year}", "rows": ann_rows, "widths": [26, 6, 34, 15, 40, 30, 14, 32, 13, 13, 36, 30, 13, 50, 40]},
        {"name": "Публикации", "rows": orders, "widths": [26, 14, 30, 34, 20, 26, 50, 46, 50, 12, 70]},
        {"name": "ИИН", "rows": iins, "widths": [15, 36, 28, 14, 36, 13, 40, 50]},
        {"name": "Файлы", "rows": files, "widths": [44, 28, 10, 13, 13, 16, 14, 70]},
    ])
