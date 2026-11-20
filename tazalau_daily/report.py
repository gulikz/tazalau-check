"""Шаг 5: итоговый Excel дня. Строка на клиента ADW: ДГД (где и когда опубликован), статус на tazalau, снимок
(ссылка), ссылки на шесть реестров с фильтром по ИИН, когда проверен. Плюс листы: совпадения с ДГД, проверка регионов,
проверка выгрузки ADW, сводка. Клиент, не проверенный сегодня (лимит или время), показан с последней проверкой из журнала."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from tazalau.registries import BASE_URL

from .adw import Adw
from .checks import REGISTRIES
from .common import dmy
from .govkz import DGD_DIR
from .match import dgd_text, maybe_text
from .xlsx import X, file_link, write_book


def _when(text: str) -> datetime | None:
    try:
        return datetime.strptime(text, "%d.%m.%Y %H:%M:%S")
    except (TypeError, ValueError):
        return None


def write_report(path: Path, work_dir: Path, adw: Adw, dg: dict | None, m: dict, results: dict, journal: dict,
                 summary: dict | None = None) -> Path:
    work_dir, path = Path(work_dir), Path(path)

    def flink(rel: str) -> str:
        if not rel:
            return ""
        return rel if rel.startswith("http") else file_link(work_dir / rel, path)

    regs = REGISTRIES
    head = ["ИИН", "ФИО (ADW)", "Продукты", "Договоры", "Дней просрочки", "Задолженность, тг", "ДГД: публикация", "ДГД: файл или страница",
            "ДГД: возможно (по ФИО)", "tazalau: найден", "tazalau: статус", "tazalau: результат", "tazalau: все записи", "Снимок",
            "Проверено"] + [r.short for r in regs] + ["Ошибка проверки"]
    lines = []
    for iin, c in adw.clients.items():
        exact, maybe = m["exact"].get(iin) or [], m["maybe"].get(iin) or []
        res = results.get(iin)
        if res is None and iin in journal:                    # не проверен сегодня: последняя проверка из журнала
            j = journal[iin]
            res = {"error": j.get("Ошибка") or "", "checked_at": _when(j.get("Дата и время")), "found": j.get("Найден") == "да",
                   "text": j.get("Результат") or "", "status": j.get("Статус") or "", "details": j.get("Подробности") or "",
                   "png": (work_dir / j["Снимок"]) if j.get("Снимок") else None, "old": True}
        first = exact[0] if exact else {}
        dgd_link = first.get("path") or first.get("page") or ""
        row = [iin, c["fio"], ", ".join(c["products"]), ", ".join(c["contracts"][:20]), c["dpd"] if c["dpd"] is not None else "",
               round(c["debt"], 2) if c["debt"] else "",
               X(dgd_text(exact), "warn") if exact else "", X("открыть", link=flink(dgd_link)) if dgd_link else "",
               X(maybe_text(maybe), "warn") if maybe else ""]
        if res and not res.get("error"):
            found, png = res.get("found"), res.get("png")
            png_rel = ""
            if png:
                png = Path(png)
                png_rel = png.relative_to(work_dir).as_posix() if png.is_relative_to(work_dir) else str(png)
            row += [X("да", "bad") if found else "нет", X(res.get("status") or "", "bad" if found else ""), res.get("text") or "",
                    res.get("details") or "", X("снимок", link=flink(png_rel)) if png_rel else "", res.get("checked_at") or ""]
        else:
            row += [X("ошибка", "bad") if res else X("не проверен", "warn"), "", "", "", "", (res or {}).get("checked_at") or ""]
        row += [X("открыть", link=r.url(iin, BASE_URL)) for r in regs]
        row.append(X(res["error"], "bad") if res and res.get("error") else "")
        rank = (0 if exact else 1 if maybe else 2, 0 if res and res.get("found") else 1, iin)
        lines.append((rank, row))
    main = [head] + [row for _, row in sorted(lines, key=lambda x: x[0])]

    matches = [["ИИН", "ФИО (ADW)", "ФИО в публикации", "Регион", "Вид", "Где", "Суд", "Требования принимаются до", "Дата",
                "Финансовый управляющий", "Контакты", "Файл", "Публикация на gov.kz", "Как прочитан ИИН"]]
    for iin, items in m["exact"].items():
        for x in items:
            matches.append([iin, adw.clients[iin]["fio"], x.get("fio") or "", x.get("region") or "", x.get("kind") or "", x.get("where") or "",
                            x.get("court") or "", x.get("claimsTo") or "", x.get("date") or "", x.get("mgr") or "", x.get("contacts") or "",
                            X("файл", link=flink(x["path"])) if x.get("path") and not x["path"].startswith("http") else "",
                            X("открыть", link=x["page"]) if x.get("page") else "", x.get("how") or ""])
    for iin, hits in m["maybe"].items():
        for h in hits:
            matches.append([iin, adw.clients[iin]["fio"], h["fio"], h["region"], X("возможно (по ФИО)", "warn"), h["where"], "", "", "", "", "",
                            X("файл", link=flink(h["path"])) if h.get("path") else "", X("открыть", link=h["page"]) if h.get("page") else "", "по ФИО"])

    regions = [["Регион", "Проверка", "Объявлений о возбуждении (файлов)", "Строк в объявлениях", "Уникальных ИИН", "Последнее размещение",
                "Других публикаций", "Где на gov.kz"]]
    for r in (dg or {}).get("regions") or []:
        s = r["site"]
        regions.append([r["name"], X(r["text"], r["level"]), r["year"] if s else "", r["annRows"] if s else "", r["annUnique"] if s else r["unique"],
                        r["annPosted"] if s else r["listDate"], (r["win"] - r["year"]) if s else "", (s["path"] or s["via"]) if s else ""])

    quality = [["Проверка выгрузки ADW", "Значение"]]
    for k, v in adw.checks.items():
        if isinstance(v, dict):
            for p, n in v.items():
                quality.append([f"{k}: {p}", n])
        else:
            quality.append([k, X(v, "warn") if k in ("неверных ИИН", "ИИН с разными ФИО") and v else v])
    quality.append(["", ""])
    quality.append(["Строка выгрузки", "Значение ИИН / причина"])
    for x in adw.invalid[:2000]:
        quality.append([x["row"], X(f"{x['raw']}: {x['err']}" + (f" ({x['fio']})" if x["fio"] else ""), "bad")])

    summary_rows = [["Показатель", "Значение"]] + [[k, v] for k, v in (summary or {}).items()]
    return write_book(path, [
        {"name": "Итог", "rows": main, "widths": [15, 34, 18, 22, 10, 14, 50, 10, 40, 10, 45, 45, 60, 10, 17] + [12] * len(regs) + [40]},
        {"name": "Совпадения с ДГД", "rows": matches, "widths": [15, 30, 30, 26, 26, 50, 30, 14, 12, 30, 26, 10, 12, 20]},
        {"name": "Регионы ДГД", "rows": regions, "widths": [30, 70, 14, 12, 12, 14, 12, 60]},
        {"name": "Проверка ADW", "rows": quality, "widths": [40, 60]},
        {"name": "Сводка", "rows": summary_rows, "widths": [50, 60]},
    ])
