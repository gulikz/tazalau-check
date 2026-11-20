"""Шаг 4: каждый ИИН на tazalau.qoldau.kz во всех шести реестрах: снимок страницы (один файл на человека, реестры
друг под другом, над каждым полоса с ИИН, реестром, датой и адресом), статус и даты записи из таблицы реестра.

Браузер и проверка страницы: из программы Tazalau Check (пакет tazalau): тот же темп запросов (сайт допускает около
одного запроса в секунду с адреса, берётся половина), те же проверки, что страница показывает именно этот ИИН.
Статус читается так же, как в закладке: главная запись: идущее дело, среди равных: самое новое.

Журнал «tazalau_журнал.csv» в рабочей папке: каждый проверенный ИИН с датой. Повторный запуск в тот же день
продолжает с непроверенных; запуск по расписанию каждый день проверяет сначала ИИН из публикаций ДГД, затем
остальных клиентов, начиная с давно не проверявшихся."""
from __future__ import annotations

import asyncio
import csv
import re
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from tazalau.capture import BrowserStartError, CaptureError, RateLimitedError, TazalauBrowser
from tazalau.config import Settings
from tazalau.registries import BY_KEY, Registry
from tazalau.stamp import build_image, save_png

from .common import clean, safe_name, surname

ORDER = ["recovery", "bankrupt_judicial", "bankrupt_extrajudicial", "debtor_extrajudicial", "debtor_judicial", "debtor_terminated"]
REGISTRIES = [BY_KEY[k] for k in ORDER]
JOURNAL = "tazalau_журнал.csv"
JOURNAL_FIELDS = ["Дата и время", "ИИН", "ФИО", "Найден", "Результат", "Статус", "Подробности", "Снимок", "Ошибка"]

RECORDS_JS = """() => {
  const clean = s => String(s == null ? '' : s).replace(/\\s+/g, ' ').trim();
  const table = document.querySelector('.table-responsive table') || document.querySelector('table.table');
  if (!table) return [];
  const heads = Array.from(table.querySelectorAll('thead th')).map(th => clean(th.textContent));
  const out = [];
  table.querySelectorAll('tbody > tr').forEach(tr => {
    const cells = Array.from(tr.children);
    if (cells.length === 1 && cells[0].colSpan > 1) return;
    const rec = {};
    cells.forEach((td, i) => { const v = clean(td.innerText || td.textContent); if (v) rec[heads[i] || 'Столбец ' + (i + 1)] = v; });
    out.push(rec);
  });
  return out;
}"""


class StatusBrowser(TazalauBrowser):
    """Браузер Tazalau Check, который после снимка ещё и читает строки таблицы реестра."""

    last_records: list = []

    def check(self, registry: Registry, iin: str):
        shot = super().check(registry, iin)
        try:
            self.last_records = self._page.evaluate(RECORDS_JS) if shot.count else []
        except Exception:
            self.last_records = []
        return shot


# ---------------------------------------------------------------- статус записи
DATE_LABELS = [
    (r"иска", "иск"), (r"принятия к производству", "принято к производству"), (r"решения суда", "решение суда"),
    (r"вступления в силу", "вступило в силу"), (r"завершения дела", "дело завершено"),
    (r"подачи заявления.*прекращ", "заявление о прекращении"), (r"подачи заявления", "заявление подано"),
    (r"начала процедуры", "начало процедуры"), (r"завершения процедуры", "процедура завершена"),
    (r"прекращения процедуры", "процедура прекращена"),
]


def _val(rec: dict, pattern: str) -> str:
    k = next((k for k in rec if re.search(pattern, k, re.I)), None)
    return rec[k] if k else ""


def _date(s: str) -> date | None:
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", s or "")
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def record_info(reg: Registry, rec: dict) -> dict:
    status = _val(rec, r"статус дела tazalau") or _val(rec, r"статус заявителя") or _val(rec, r"статус дела")
    if reg.key == "debtor_terminated":
        who = _val(rec, r"инициатор")
        status = "процедура прекращена" + (f" (инициатор: {who})" if who else "")
    events = []
    for k, v in rec.items():
        d = _date(v) if re.match(r"^дата", k, re.I) and not re.search(r"передачи судье", k, re.I) else None
        if d:
            label = next((lab for pat, lab in DATE_LABELS if re.search(pat, k, re.I)), k.lower())
            events.append({"d": d, "label": label, "text": d.strftime("%d.%m.%Y")})
    events.sort(key=lambda e: e["d"])
    closed = bool(re.search(r"возвращ|отказ|без рассмотрения|прекращ|отмен", status, re.I))   # иск без процедуры
    return {"reg": reg, "rec": rec, "status": status or "запись есть", "events": events, "last": events[-1] if events else None,
            "first": events[0] if events else None, "closed": closed, "category": _val(rec, r"категория дела"),
            "place": ", ".join(x for x in (_val(rec, r"^регион"), _val(rec, r"наименование суда|услугодатель")) if x),
            "creditors": _val(rec, r"список кредиторов"), "sums": [v for k, v in rec.items() if re.search(r"[сc]умма задолженности", k, re.I)],
            "fio": _val(rec, r"^фио")}


def record_line(x: dict) -> str:
    parts = [f"{x['reg'].short}: {x['status']}"]
    if x["category"]:
        parts.append(x["category"])
    if x["events"]:
        parts.append(", ".join(f"{e['label']} {e['text']}" for e in x["events"]))
    if x["place"]:
        parts.append(x["place"])
    if x["creditors"]:
        parts.append("кредиторы: " + x["creditors"][:400] + (f"; сумма, тг: {' / '.join(x['sums'])}" if x["sums"] else ""))
    return " · ".join(parts)


def describe(shots: list, fio_in_file: str = "") -> dict:
    """shots: [(PageShot, записи таблицы)]. Итог по ИИН: найден ли, текст, главный статус, все записи."""
    found = [(s, recs) for s, recs in shots if s.count > 0]
    if found:
        text = "НАЙДЕН: " + "; ".join(f"{s.registry.short}, записей: {s.count}" + (f" (на снимке первые {s.rows_on_page})" if s.rows_on_page < s.count else "")
                                      for s, _ in found)
    else:
        text = f"НЕ НАЙДЕН: {shots[0][0].registry.short}" if len(shots) == 1 else f"НЕ НАЙДЕН ни в одном из {len(shots)} реестров"
    if any(s.clipped for s, _ in shots):
        text += " · внимание: таблица шире снимка"
    infos = [record_info(s.registry, rec) for s, recs in found for rec in recs]
    status = "НАЙДЕН (статус не прочитан)" if found else "нет записей ни в одном из проверенных реестров"
    if infos:
        def rank(x):
            return (0 if x["closed"] else 1, x["first"]["d"].toordinal() if x["first"] else 0, x["last"]["d"].toordinal() if x["last"] else 0)
        latest = max(infos, key=rank)                    # при равенстве: первая (порядок реестров)
        status = f"{latest['reg'].short}: {latest['status']}" + (f" · {latest['last']['label']} {latest['last']['text']}" if latest["last"] else "") \
            + (f" (всего записей: {len(infos)})" if len(infos) > 1 else "")
    notes = []
    own = surname(fio_in_file)
    other = next((x["fio"] for x in infos if x["fio"] and own and surname(x["fio"]) and surname(x["fio"]) != own), "")
    if other:
        notes.append(f"внимание: на сайте фамилия «{other.split()[0]}», в файле «{clean(fio_in_file).split()[0]}»")
    details = " | ".join([record_line(x) for x in infos] + notes)[:32000]
    at = shots[-1][0].checked_at
    return {"found": bool(found), "text": f"{text} · проверено {at:%d.%m.%Y %H:%M}", "status": status, "details": details, "infos": infos}


# ---------------------------------------------------------------- журнал
def read_journal(work_dir: Path) -> dict[str, dict]:
    """Последняя проверка каждого ИИН."""
    path = Path(work_dir) / JOURNAL
    out: dict[str, dict] = {}
    if not path.exists():
        return out
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            if row.get("ИИН"):
                out[row["ИИН"]] = row
    return out


def _append_journal(work_dir: Path, row: dict) -> None:
    path = Path(work_dir) / JOURNAL
    new = not path.exists()
    with path.open("a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=JOURNAL_FIELDS, delimiter=";")
        if new:
            w.writeheader()
        w.writerow({k: row.get(k, "") for k in JOURNAL_FIELDS})


def plan(clients: dict, priority: list[str], journal: dict, *, today: date | None = None, recheck_days: int = 1,
         limit: int | None = None) -> list[str]:
    """Очередь на сегодня: сначала ИИН из публикаций ДГД (priority), потом остальные, давно не проверявшиеся первыми.
    ИИН, проверенные без ошибки меньше recheck_days дней назад, пропускаются."""
    today = today or date.today()

    def last(iin: str) -> date | None:
        row = journal.get(iin)
        if not row or row.get("Ошибка"):
            return None
        try:
            return datetime.strptime(row["Дата и время"][:10], "%d.%m.%Y").date()
        except (KeyError, ValueError):
            return None

    def due(iin: str) -> bool:
        d = last(iin)
        return d is None or (today - d).days >= recheck_days

    first = [i for i in dict.fromkeys(priority) if i in clients and due(i)]
    rest = sorted((i for i in clients if i not in set(first) and due(i)), key=lambda i: (last(i) or date.min, i))
    queue = first + rest
    return queue[:limit] if limit else queue


def _check_one(browser: StatusBrowser, regs: list[Registry], iin: str, settings: Settings, log: Callable[[str], None]):
    """Снимки всех реестров одного ИИН с записями таблиц либо (None, причина). Как Runner._check_iin программы."""
    shots = []
    for reg in regs:
        attempt = limited = 0
        while True:
            try:
                shot = browser.check(reg, iin)
                shots.append((shot, list(browser.last_records)))
                break
            except RateLimitedError as exc:
                limited += 1
                if limited > settings.rate_limit_retries:
                    return None, f"{reg.short}: {exc}"
                pause = min(settings.rate_limit_pause_s * 2 ** (limited - 1), 180.0)
                log(f"   сайт просит не торопиться: жду {pause:.0f} с, дальше один запрос в {browser.slow_down():.1f} с")
                time.sleep(pause)
            except CaptureError as exc:
                attempt += 1
                if attempt >= settings.retries:
                    return None, f"{reg.short}: {exc}"
                backoff = settings.retry_backoff_s or [5.0]
                time.sleep(backoff[min(attempt - 1, len(backoff) - 1)])
    return shots, None


def check_iins(queue: list[str], clients: dict, work_dir: Path, *, settings: Settings | None = None, registries: list[Registry] | None = None,
               log: Callable[[str], None] = print, minutes: float | None = None, shots_dir: Path | None = None,
               browser_factory=None) -> dict[str, dict]:
    """Проверяет ИИН из очереди (не дольше minutes минут, если задано). Возвращает {ИИН: результат}; каждый результат
    сразу пишется в журнал, так что прерванный запуск ничего не теряет."""
    settings = settings or Settings()
    regs = registries or REGISTRIES
    work_dir = Path(work_dir)
    shots_dir = Path(shots_dir) if shots_dir else work_dir / "Снимки_tazalau" / date.today().isoformat()
    results: dict[str, dict] = {}
    if not queue:
        return results
    started = time.monotonic()
    fails = 0
    log(f"tazalau: к проверке {len(queue)} ИИН × {len(regs)} реестров, примерно "
        f"{round(len(queue) * len(regs) * max(settings.min_request_interval_s, 1.0) / 60)} мин.")
    browser_factory = browser_factory or (lambda: StatusBrowser(settings, log=lambda t: None))   # тесты подставляют свой
    with browser_factory() as browser:
        for n, iin in enumerate(queue, start=1):
            if minutes and time.monotonic() - started > minutes * 60:
                log(f"tazalau: время вышло ({minutes} мин), проверено {n - 1} из {len(queue)}; остальные: при следующем запуске")
                break
            fio = clients.get(iin, {}).get("fio", "")
            shots, error = _check_one(browser, regs, iin, settings, log)
            now = datetime.now()
            row = {"Дата и время": now.strftime("%d.%m.%Y %H:%M:%S"), "ИИН": iin, "ФИО": fio}
            if error:
                fails += 1
                row.update({"Ошибка": error})
                results[iin] = {"iin": iin, "error": error, "checked_at": now}
                _append_journal(work_dir, row)
                log(f"{n}/{len(queue)} {iin}: ОШИБКА {error}")
                if fails >= settings.max_consecutive_failures:
                    log("tazalau: несколько ИИН подряд с ошибкой: сайт недоступен, проверка остановлена")
                    break
                continue
            fails = 0
            d = describe(shots, fio)
            folder = shots_dir / (f"{safe_name(fio)[:60]} {iin}".strip() if fio else iin)
            folder.mkdir(parents=True, exist_ok=True)
            png = folder / f"{iin}_{now:%Y-%m-%d_%H-%M-%S}.png"
            save_png(build_image([s for s, _ in shots], stamp=settings.stamp), png)
            res = {"iin": iin, "error": "", "checked_at": now, "png": png, "found": d["found"], "text": d["text"], "status": d["status"],
                   "details": d["details"], "urls": {s.registry.key: s.url for s, _ in shots},
                   "counts": {s.registry.key: s.count for s, _ in shots}}
            results[iin] = res
            row.update({"Найден": "да" if d["found"] else "нет", "Результат": d["text"], "Статус": d["status"], "Подробности": d["details"],
                        "Снимок": png.relative_to(work_dir).as_posix() if png.is_relative_to(work_dir) else str(png)})
            _append_journal(work_dir, row)
            if n % 10 == 0 or d["found"] or n == len(queue):
                log(f"{n}/{len(queue)} {'НАЙДЕН' if d['found'] else 'нет записей'}" + (f": {d['status']}" if d["found"] else ""))
    return results


def in_worker(fn, *args, **kwargs):
    """Playwright (синхронный) не работает внутри цикла asyncio, а в Jupyter такой цикл есть: тогда: в отдельном потоке."""
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return fn(*args, **kwargs)
    with ThreadPoolExecutor(max_workers=1) as pool:
        return pool.submit(fn, *args, **kwargs).result()


__all__ = ["REGISTRIES", "JOURNAL", "StatusBrowser", "describe", "read_journal", "plan", "check_iins", "in_worker", "BrowserStartError"]
