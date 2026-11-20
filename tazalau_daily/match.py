"""Шаг 3: кто из клиентов ADW есть в публикациях ДГД. Точно: по ИИН. Возможно: по фамилии и первой букве имени,
если в публикации ИИН не прочитан (скан, нет файла, пустая ячейка ИИН): такие строки помечаются «⚠ возможно»."""
from __future__ import annotations

from .adw import Adw
from .common import dmy, name_key


def match(adw: Adw, dg: dict | None) -> dict:
    if not dg:
        return {"exact": {}, "maybe": {}, "dgd_iins": 0, "dgd_not_clients": 0}
    exact = {iin: dg["by_iin"][iin] for iin in adw.clients if iin in dg["by_iin"]}
    # Публикации без прочитанного ИИН: строки объявлений с пустым ИИН и публикации, где ИИН не найден ни в файле, ни в названии.
    unread = [{"fio": a["fio"], "region": a["region"], "where": "объявление о возбуждении дела (ИИН в строке пустой или с ошибкой)"
               + (f", размещено {dmy(a['posted'])}" if a.get("posted") else ""), "page": a["page"], "path": a["path"]}
              for a in dg["announcements"] if a["fio"] and not a["iin"]]
    unread += [{"fio": o["debtor"], "region": o["region"], "where": (o["kind"] or "публикация") + (f" от {dmy(o['date'])}" if o["date"] else "")
                + " (ИИН в файле не прочитан)", "page": o["page"], "path": next((f["path"] for f in o["files"] if f["saved"]), "")}
               for o in dg["orders"] if o["debtor"] and not o["iins"]]
    by_key: dict[str, list] = {}
    for u in unread:
        k = name_key(u["fio"])
        if k:
            by_key.setdefault(k, []).append(u)
    maybe = {}
    for iin, c in adw.clients.items():
        if iin in exact:
            continue
        hits = [u for f in (c["fios"] or {c["fio"]}) for u in by_key.get(name_key(f), [])]
        if hits:
            maybe[iin] = hits
    return {"exact": exact, "maybe": maybe, "dgd_iins": len(dg["by_iin"]),
            "dgd_not_clients": len(set(dg["by_iin"]) - set(adw.clients))}


def dgd_text(items: list[dict]) -> str:
    """Где ИИН есть в публикациях ДГД (для столбца итогового файла)."""
    parts = []
    for x in items:
        conf = "" if x.get("conf") in ("high", None, "") else " (сверить)"
        parts.append(f"{x.get('region') or 'регион не определён'}: {x.get('where') or x.get('kind') or 'публикация'}{conf}")
    return "; ".join(dict.fromkeys(parts))


def maybe_text(hits: list[dict]) -> str:
    return "⚠ возможно (по ФИО, ИИН в файле не прочитан): " + "; ".join(dict.fromkeys(f"{h['region']}: {h['where']}" for h in hits))
