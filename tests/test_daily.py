"""Ежедневный цикл (пакет tazalau_daily) без сети: gov.kz и tazalau заменены заготовками.
ИИН в тестах: с невозможной датой рождения (31 апреля, 13-й месяц): такой номер не может принадлежать человеку."""
from __future__ import annotations

import csv
import json
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from tazalau.capture import CaptureError, PageShot
from tazalau.registries import BY_KEY
from tazalau_daily import adw as adw_mod
from tazalau_daily import checks, dgd, files, govkz, match, report
from tazalau_daily.common import name_key, norm_iin, regions_in

from conftest import fake_iin, make_png

# 31 апреля не бывает: ИИН проходит проверку «похож на ИИН человека», но человеку принадлежать не может.
P1 = fake_iin("99043130001")
P2 = fake_iin("98043140002")
P3 = fake_iin("01043150003")      # начинается с «0»: Excel-число теряет ведущий ноль
P4 = fake_iin("97043130004")
YEAR = "2026"


# ---------------------------------------------------------------- gov.kz: заготовка сайта
def turbo(obj) -> str:
    """Страница gov.kz с данными внутри: плоский массив значений turbo-stream (ключи: ссылки «_номер»)."""
    vals: list = []

    def enc(o) -> int:
        i = len(vals)
        vals.append(None)
        if isinstance(o, dict):
            d = {}
            for k, v in o.items():
                vals.append(k)
                key = "_" + str(len(vals) - 1)          # номер ключа: до кодирования значения (оно дописывает в vals)
                d[key] = enc(v)
            vals[i] = d
        elif isinstance(o, list):
            vals[i] = [enc(x) for x in o]
        else:
            vals[i] = o
        return i

    enc({"loaderData": {"route": {"queries": obj}}})
    return ("<html><body><script>window.__reactRouterContext.streamController.enqueue("
            + json.dumps(json.dumps(vals, ensure_ascii=False)) + ");</script></body></html>")


def query(endpoint: str, value) -> dict:
    return {"endpointName": endpoint, "arg": {}, "value": value}


def announcement_xlsx(rows: list[list]) -> bytes:
    """Таблица «Объявление о возбуждении производства…» как у ДГД: шапка в 4-й строке, «с / до», номера столбцов."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Объявление о возбуждении производства по делу о применении процедуры восстановления платежеспособности"])
    ws.append([])
    ws.append([])
    ws.append(["№", "ФИО должника", "ИИН должника", "Адрес должника", "Наименование суда", "Дата вынесения определения",
               "ФИО финансового управляющего", "Срок приема требований", None, "Адрес приема требований", "Контактные данные",
               "Дата размещения объявления"])
    ws.append([None, None, None, None, None, None, None, "с", "до"])
    ws.append(list(range(1, 13)))
    for r in rows:
        ws.append(r)
    import io
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


ANN_ROWS = [
    [1, "Тестов Тест Тестович", P1, "г. Актау", "СМЭС Мангистауской области", datetime(2026, 11, 2), "Управляющая А. Б.",
     datetime(2026, 11, 5), datetime(2027, 1, 5), "г. Актау, 1", "+7 700 000 00 00", datetime(2026, 11, 5)],
    [2, "Примеров Пример", int(P3), "г. Жанаозен", "СМЭС Мангистауской области", "03.11.2026", "Управляющий В. Г.",
     "06.11.2026", "06.01.2027", "г. Актау, 2", "", "06.11.2026"],
    [3, "Безиинов Без Иинович", "", "г. Актау", "СМЭС", "04.11.2026", "Управляющий В. Г.", "", "", "", "", "07.11.2026"],
    [4, "Опечаткин Опечатка", P2[:11] + str((int(P2[11]) + 1) % 10), "г. Актау", "СМЭС", "04.11.2026", "", "", "", "", "", "07.11.2026"],
]

ANN_TITLE = ("Объявление о возбуждении производства по делу о применении процедуры восстановления платежеспособности "
             "и порядке заявления требований кредиторами")
RTK_TITLE = f"Реестр требований кредиторов в процедуре восстановления платежеспособности по должнику: Четвёртов Четвёрт, ИИН {P4}"


class FakeGov(govkz.GovClient):
    """gov.kz по таблице адресов. Считает запросы."""

    def __init__(self, pages: dict):
        super().__init__(spacing=0, log=lambda t: None)
        self.pages = pages
        self.calls: list[str] = []

    def get(self, url, *, timeout=None, limit=govkz.MAX_FILE):
        self.calls.append(url)
        self.requests += 1
        hit = self.pages.get(url)
        if hit is None:
            return govkz.Response(404, b"", "", "код 404")
        body, ctype = hit
        return govkz.Response(200, body if isinstance(body, bytes) else body.encode("utf-8"), ctype)


def mangistau_site(xlsx: bytes) -> dict:
    base = "https://www.gov.kz/org/kgd-mangistau"
    api = "https://www.gov.kz/api/v1/public/content-manager/documents?id=eq%3A{}&projects=eq%3Akgd-mangistau"
    html = "text/html; charset=utf-8"
    return {
        base + "/activities/directions?lang=ru": (turbo([query("getActivityDirections", {"directions": [
            {"id": 10, "title": "Восстановление и платежеспособности граждан Республики Казахстан", "parent_direction": None},
            {"id": 11, "title": "Реабилитация и банкротство юридических лиц", "parent_direction": None},
            {"id": 12, "title": "Налоговое администрирование", "parent_direction": None},
            {"id": 20, "title": "2026 год", "parent_direction": {"id": 10}}]})]), html),
        base + "/activities/10?lang=ru": (turbo([query("fetchArticleByDirection", {"title": "Восстановление", "documents": [], "article": {"content": ""}}),
                                                 query("getChildDirections", {"directions": [{"id": 20, "title": "2026 год"}, {"id": 21, "title": "2025 год"}]})]), html),
        base + "/activities/20?lang=ru": (turbo([query("fetchArticleByDirection", {"title": "2026 год", "documents": [], "article": {
            "content": '<p>Реестр требований кредиторов по должнику Пятов Пять</p><a href="/uploads/2026/7/2/feed01_original.5.pdf">скачать</a>'}}),
            query("getChildDirections", {"directions": [{"id": 30, "title": "2026 год. Восстановление платежеспособности. Информационные сообщения"}]})]), html),
        base + "/activities/30?lang=ru": (turbo([query("fetchArticleByDirection", {"title": "Информационные сообщения", "article": {"content": ""}, "documents": [
            {"id": 1001, "title": ANN_TITLE, "date": "2026-11-07T10:00:00+05:00"},
            {"id": 1002, "title": RTK_TITLE, "date": "2026-07-01T10:00:00+05:00"},
            {"id": 1003, "title": "Объявление о возбуждении производства по делу о реабилитации ТОО «Ромашка»", "date": "2026-07-02"}]}),
            query("getChildDirections", {"directions": []})]), html),
        api.format(1001): (json.dumps({"data": [{"full_text": [{"document": "/uploads/2026/11/7/abc123_original.234.xlsx"}]}]}), "application/json"),
        api.format(1002): (json.dumps({"data": [{"full_text": [{"document": "/uploads/2026/7/1/def456_original.10.pdf"}]}]}), "application/json"),
        "https://www.gov.kz/uploads/2026/11/7/abc123_original.234.xlsx": (xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "https://www.gov.kz/uploads/2026/7/2/feed01_original.5.pdf": (b"%PDF-1.4 scan without text", "application/pdf"),
    }


def test_gov_data_decodes_stream_and_promises():
    page = turbo([query("getActivityDirections", {"directions": [{"id": 1, "title": "Банкротство", "parent_direction": None}]})])
    data = govkz.gov_data(page)
    assert data["getActivityDirections"]["value"]["directions"][0]["title"] == "Банкротство"
    assert govkz.gov_data("<html>нет данных</html>") == {}
    # отложенная часть «P1:[…]» дописывается в конец массива
    vals = [{"_1": 2}, "q", ["P", 1]]
    text = json.dumps(vals) + "\nP1:" + json.dumps([{"_4": 5}, "endpointName", "getX"]) + "\n"
    html = "<script>streamController.enqueue(" + json.dumps(text) + ")</script>"
    assert govkz.gov_data(html) == {"getX": {"endpointName": "getX"}}


def test_titles_kind_debtor_and_iins():
    assert govkz.ann_citizen(ANN_TITLE)
    assert not govkz.ann_citizen("Объявление о возбуждении производства по делу о реабилитации ТОО «Ромашка»")
    assert govkz.title_iins(RTK_TITLE) == [P4]
    assert govkz.title_iins("БИН 991340000001 и номер 123") == []
    assert govkz.gov_kind(ANN_TITLE) == "объявление о возбуждении дела, восстановление платежеспособности"
    assert govkz.gov_kind(RTK_TITLE).startswith("реестр требований кредиторов")
    assert govkz.gov_debtor(RTK_TITLE) == "Четвёртов Четвёрт"
    assert govkz.gov_debtor("РТК должник ИВАНОВ ИВАН") == "ИВАНОВ ИВАН"
    assert govkz.gov_debtor("Сидорова Анна Петровна") == "Сидорова Анна Петровна"


def test_parse_announcements_reads_every_row(tmp_path):
    path = tmp_path / "ann.xlsx"
    path.write_bytes(announcement_xlsx(ANN_ROWS))
    ann = files.parse_announcements(files.file_rows(path).rows)
    assert ann["headerRow"] == 4
    assert {"n", "iin", "fio", "manager", "court", "ruling", "claims", "to", "claimsAddress", "contacts", "posted", "address"} <= set(ann["roles"])
    recs = ann["records"]
    assert [r["iin"] for r in recs] == [P1, P3, "", ""]                 # число без ведущего нуля восстановлено
    assert recs[0]["ruling"] == date(2026, 11, 2) and recs[0]["to"] == date(2027, 1, 5) and recs[0]["posted"] == date(2026, 11, 5)
    assert recs[1]["from"] == date(2026, 11, 6) and recs[1]["manager"] == "Управляющий В. Г."
    assert len(ann["invalid"]) == 1 and "контрольная" in ann["invalid"][0]["err"]
    assert files.parse_announcements([files.Row(1, [files.Cell("просто текст", False)])]) is None


def test_extract_iins_from_text_and_table():
    rows = [files.Row(1, [files.Cell(f"Приказ о назначении финансового управляющего по заявлению Иванова Ивана, ИИН {P1}", False)]),
            files.Row(2, [files.Cell("ИИН разбит: " + " ".join([P2[:4], P2[4:8], P2[8:]]), False)]),
            files.Row(3, [files.Cell("номер 123456789012 без контрольной цифры", False)])]
    found = [x["iin"] for x in files.extract_iins(rows)["items"]]
    assert found == [P1, P2]
    table = [files.Row(1, [files.Cell("ФИО", False), files.Cell("ИИН", False), files.Cell("Дата", False)]),
             files.Row(2, [files.Cell("Петров Пётр", False), files.Cell(P3[1:], True, int(P3)), files.Cell("05.11.2026", False)]),
             files.Row(3, [files.Cell("Ошибкин", False), files.Cell("12345", False), None])]
    ex = files.extract_iins(table)
    assert [(x["iin"], x["fio"], x["date"]) for x in ex["items"]] == [(P3, "Петров Пётр", date(2026, 11, 5))]
    assert len(ex["invalid"]) == 1


def test_harvest_reads_tree_downloads_once_and_audits_regions(tmp_path):
    site = FakeGov(mangistau_site(announcement_xlsx(ANN_ROWS)))
    res = govkz.harvest_region(site, "kgd-mangistau", "Мангистауская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    assert res["state"] == "ok" and res["announcements"] == 1 and res["publications"] == 3   # ТОО отброшено
    assert not any("/activities/11" in u or "/activities/21" in u for u in site.calls)        # юрлица и 2025 год не обходятся
    assert not any("def456" in u for u in site.calls)                                         # ИИН уже в названии: файл не нужен
    folder = tmp_path / "ДГД" / "Мангистауская область"
    meta = json.loads((folder / govkz.DGD_META).read_text(encoding="utf-8"))
    assert meta["state"] == "ok" and meta["slug"] == "kgd-mangistau"
    assert sorted(p.suffix for p in folder.iterdir() if p.name != govkz.DGD_META) == [".pdf", ".xlsx"]

    again = FakeGov(mangistau_site(announcement_xlsx(ANN_ROWS)))
    res2 = govkz.harvest_region(again, "kgd-mangistau", "Мангистауская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    assert res2["new_files"] == 0 and not any("/uploads/" in u for u in again.calls)         # тот же адрес = тот же файл

    dg = dgd.read_dgd(tmp_path, use_ocr=False, log=lambda t: None)
    assert len(dg["announcements"]) == 4
    assert set(dg["by_iin"]) == {P1, P3, P4}
    regions = {r["name"]: r for r in dg["regions"]}
    assert len(regions) == 20
    mg = regions["Мангистауская обл."]
    assert mg["annRows"] == 4 and mg["annUnique"] == 2 and mg["level"] == "warn"
    assert "всего 2 ИИН" in mg["text"] and "ИИН с ошибкой" in mg["text"] and "ИИН не прочитан в файлах: 1" in mg["text"]
    assert regions["Атырауская обл."]["level"] == "bad" and regions["Атырауская обл."]["text"] == "нет списка"

    book = dgd.write_dgd_book(dg, tmp_path)
    wb = load_workbook(book)
    assert wb.sheetnames == ["Регионы", f"Объявления {date.today().year}", "Публикации", "ИИН", "Файлы"]
    ws = wb[f"Объявления {date.today().year}"]
    assert ws.max_row == 5 and ws.cell(2, 15).hyperlink is not None


def test_failed_region_is_marked_and_keeps_old_rows(tmp_path):
    site = FakeGov(mangistau_site(announcement_xlsx(ANN_ROWS)))
    govkz.harvest_region(site, "kgd-mangistau", "Мангистауская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    broken = FakeGov({})
    out = govkz.harvest_all(tmp_path, year=YEAR, regions=["kgd-mangistau"], client=broken, log=lambda t: None)
    assert out[0]["state"] == "fail"
    dg = dgd.read_dgd(tmp_path, use_ocr=False, log=lambda t: None)
    mg = next(r for r in dg["regions"] if r["name"] == "Мангистауская обл.")
    assert mg["level"] == "bad" and "данные прошлого сбора" in mg["text"] and mg["annRows"] == 4


def test_documents_search_when_tree_has_no_announcement(tmp_path):
    base = "https://www.gov.kz/org/kgd-sko"
    html = "text/html; charset=utf-8"
    pages = {
        base + "/activities/directions?lang=ru": (turbo([query("getActivityDirections", {"directions": [
            {"id": 5, "title": "Банкротство физических лиц", "parent_direction": None}]})]), html),
        base + "/activities/5?lang=ru": (turbo([query("fetchArticleByDirection", {"documents": [], "article": {"content": ""}}),
                                                query("getChildDirections", {"directions": []})]), html),
        base + "/documents/1?lang=ru&size=50&title=%D0%B2%D0%BE%D0%B7%D0%B1%D1%83%D0%B6%D0%B4%D0%B5%D0%BD%D0%B8%D0%B8%20%D0%BF%D1%80%D0%BE%D0%B8%D0%B7%D0%B2%D0%BE%D0%B4%D1%81%D1%82%D0%B2%D0%B0":
            (turbo([query("getFilteredData", {"data": [{"id": 77, "title": ANN_TITLE, "created_date": "2026-10-01T09:00:00"},
                                                         {"id": 78, "title": ANN_TITLE, "created_date": "2025-10-01T09:00:00"}]})]), html),
    }
    res = govkz.harvest_region(FakeGov(pages), "kgd-sko", "Северо-Казахстанская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    assert res["via"] == "Документы (поиск)" and res["publications"] == 1                    # прошлогоднее не взято


# ---------------------------------------------------------------- ADW и сверка
def adw_csv(tmp_path: Path) -> Path:
    path = tmp_path / "ADW_2026-11-20.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ИИН/БИН", "FIO", "Продукт", "Договор", "Дней просрочки", "Задолженность, тг"])
        w.writerow([P1, "ТЕСТОВ ТЕСТ ТЕСТОВИЧ", "Кредит", "K-1", "45", "1000,50"])
        w.writerow([P1, "ТЕСТОВ ТЕСТ ТЕСТОВИЧ", "Кредитная карта", "C-1", "12", "200"])
        w.writerow([P3[1:], "ПРИМЕРОВ ПРИМЕР", "Кредит", "K-2", "30", "300"])       # ведущий ноль потерян: 11 цифр текстом
        w.writerow([P2, "БЕЗИИНОВ БЕЗ", "Кредит", "K-3", "5", "10"])
        w.writerow(["12345", "ОШИБКИН", "Кредит", "K-4", "1", "1"])
    return path


def test_adw_export_checks_and_match(tmp_path):
    a = adw_mod.load(adw_csv(tmp_path))
    assert a.checks["строк"] == 5 and a.checks["ИИН"] == 2 and a.checks["неверных ИИН"] == 2   # 11 цифр текстом и «12345»
    assert a.clients[P1]["products"] == ["Кредит", "Кредитная карта"] and a.clients[P1]["dpd"] == 45
    assert abs(a.clients[P1]["debt"] - 1200.5) < 1e-6
    site = FakeGov(mangistau_site(announcement_xlsx(ANN_ROWS)))
    govkz.harvest_region(site, "kgd-mangistau", "Мангистауская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    dg = dgd.read_dgd(tmp_path, use_ocr=False, log=lambda t: None)
    m = match.match(a, dg)
    assert list(m["exact"]) == [P1] and list(m["maybe"]) == [P2]       # P2: в объявлении ИИН пустой, совпали фамилия и имя
    assert "возможно" in match.maybe_text(m["maybe"][P2])
    assert adw_mod.latest_export(tmp_path).name == "ADW_2026-11-20.csv"


def test_name_key_and_regions():
    assert name_key("ТЕСТОВ ТЕСТ ТЕСТОВИЧ") == name_key("Тестов Т.Т.") == "тестов т"
    assert regions_in("Департамент по Алматинской области") == ["Алматинская обл."]
    assert regions_in("г. Алматы") == ["г. Алматы"]
    assert norm_iin(int(P3)) == (P3, "")


# ---------------------------------------------------------------- tazalau: статус, очередь, итог
def shot(key: str, iin: str, count: int) -> PageShot:
    reg = BY_KEY[key]
    return PageShot(reg, iin, reg.url(iin), count, count, make_png(), datetime(2026, 11, 20, 10, 0).astimezone())


def test_describe_prefers_running_case():
    recs = [{"ФИО": "Тестов Т. Т.", "Статус дела": "Возвращено", "Дата подачи иска": "01.02.2026"},
            {"ФИО": "Тестов Т. Т.", "Статус дела": "Принято к производству", "Дата подачи иска": "01.01.2026",
             "Дата принятия к производству": "10.01.2026"}]
    d = checks.describe([(shot("recovery", P1, 2), recs), (shot("bankrupt_judicial", P1, 0), [])], "ТЕСТОВ ТЕСТ")
    assert d["found"] and d["text"].startswith("НАЙДЕН: Восстановление платежеспособности, записей: 2")
    assert d["status"] == "Восстановление платежеспособности: Принято к производству · принято к производству 10.01.2026 (всего записей: 2)"
    d2 = checks.describe([(shot("recovery", P1, 1), [{"ФИО": "Другой Д. Д.", "Статус дела": "Идёт"}])], "ТЕСТОВ ТЕСТ")
    assert "фамилия «Другой»" in d2["details"]
    d3 = checks.describe([(shot("recovery", P1, 0), []), (shot("debtor_judicial", P1, 0), [])])
    assert not d3["found"] and d3["text"].startswith("НЕ НАЙДЕН ни в одном из 2 реестров")


def test_plan_puts_dgd_first_and_skips_recent():
    clients = {i: {} for i in (P1, P2, P3, P4)}
    journal = {P2: {"Дата и время": date.today().strftime("%d.%m.%Y") + " 09:00:00", "Ошибка": ""},
               P3: {"Дата и время": "01.11.2026 09:00:00", "Ошибка": ""},
               P4: {"Дата и время": date.today().strftime("%d.%m.%Y") + " 09:00:00", "Ошибка": "нет связи"}}
    # P3 из ДГД: первым; P2 проверен сегодня: пропущен; P4 (ошибка) и P1 (не проверялся): оба «давно», по порядку ИИН
    assert checks.plan(clients, [P3], journal) == [P3, P4, P1]
    assert checks.plan(clients, [P3], journal, limit=2) == [P3, P4]
    assert checks.plan(clients, [], journal, recheck_days=30) == [P4, P1]   # P3 проверен 01.11: в пределах 30 дней


class FakeStatusBrowser:
    def __init__(self, found: dict, fail: set | None = None):
        self.found, self.fail, self.last_records, self.calls = found, fail or set(), [], []

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def slow_down(self):
        return 2.0

    def check(self, registry, iin):
        self.calls.append((registry.key, iin))
        if iin in self.fail:
            raise CaptureError("сайт не ответил за 45 с")
        n = self.found.get((registry.key, iin), 0)
        self.last_records = [{"ФИО": "Тестов Т. Т.", "Статус дела": "Принято к производству", "Дата подачи иска": "01.11.2026"}] * n
        return shot(registry.key, iin, n)


def test_check_iins_writes_shots_journal_and_report(tmp_path, settings):
    a = adw_mod.load(adw_csv(tmp_path))
    fake = FakeStatusBrowser({("recovery", P1): 1})
    res = checks.check_iins([P1, P2], a.clients, tmp_path, settings=settings, log=lambda t: None, browser_factory=lambda: fake)
    assert res[P1]["found"] and not res[P2]["found"] and len(fake.calls) == 12
    assert res[P1]["png"].exists() and res[P1]["png"].parent.name == f"ТЕСТОВ ТЕСТ ТЕСТОВИЧ {P1}"
    journal = checks.read_journal(tmp_path)
    assert journal[P1]["Найден"] == "да" and journal[P2]["Найден"] == "нет"

    failing = FakeStatusBrowser({}, fail={P2})
    res2 = checks.check_iins([P2], a.clients, tmp_path, settings=settings, log=lambda t: None, browser_factory=lambda: failing)
    assert res2[P2]["error"].startswith("Восстановление платежеспособности: сайт не ответил")

    site = FakeGov(mangistau_site(announcement_xlsx(ANN_ROWS)))
    govkz.harvest_region(site, "kgd-mangistau", "Мангистауская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    dg = dgd.read_dgd(tmp_path, use_ocr=False, log=lambda t: None)
    m = match.match(a, dg)
    out = report.write_report(tmp_path / "Итог.xlsx", tmp_path, a, dg, m, {P1: res[P1]}, checks.read_journal(tmp_path), {"Дата": "20.11.2026"})
    wb = load_workbook(out)
    assert wb.sheetnames == ["Итог", "Совпадения с ДГД", "Регионы ДГД", "Проверка ADW", "Сводка"]
    ws = wb["Итог"]
    head = [c.value for c in ws[1]]
    assert ws.cell(2, 1).value == P1                                    # клиент из публикации ДГД: первым
    assert ws.cell(2, head.index("tazalau: найден") + 1).value == "да"
    assert ws.cell(2, head.index("Снимок") + 1).hyperlink.target.endswith(".png")
    assert "flApplicant" in ws.cell(2, head.index("Восстановление платежеспособности") + 1).hyperlink.target
    p2_row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == P2)
    assert ws.cell(p2_row, head.index("Ошибка проверки") + 1).value.startswith("Восстановление")   # из журнала
    assert wb["Совпадения с ДГД"].max_row == 3                          # одно точное и одно «возможно»


# ---------------------------------------------------------------- то, что показали настоящие файлы ДГД (20.11.2026)
def two_sheet_xlsx(path: Path, rows: list[list]) -> None:
    """Как у Астаны и Улытау: первый лист на казахском (без слова «ИИН»), второй: тот же список на русском."""
    wb = Workbook()
    kz = wb.active
    kz.title = "ВП (гос.яз)"
    kz.append(["Төлем қабілеттілігін қалпына келтіру рәсімін қолдану туралы іс бойынша іс жүргізуді қозғау туралы хабарландыру"])
    kz.append([])
    kz.append(["№", "Борышкердің тегі, аты, әкесінің аты (егер ол жеке басын куәландыратын құжатта көрсетілсе)",
               "Борышкердің жеке сәйкестендіру нөмірі", "Борышкердің мекен жайы", "Іс жүргізуді қозғау туралы ұйғарым шығарған соттың атауы",
               "Ұйғарымның шығарылған күні", "Қаржылық басқарушының тегі, аты, әкесінің аты",
               "Кредиторлар талаптарын қабылдау мерзімі", None, "Талаптарды қабылдау мекенжайы", "Қаржылық басқарушының байланыс деректері",
               "Хабарландыруды орналастыру күні"])
    kz.append([None, None, None, None, None, None, None, "бастап", "дейін"])
    ru = wb.create_sheet("ВП (рус.яз)")
    ru.append(["Объявление о возбуждении производства по делу о применении процедуры восстановления платежеспособности"])
    ru.append([])
    ru.append(["№", "ФИО должника", "ИИН должника", "Адрес должника", "Наименование суда", "Дата вынесения определения",
               "ФИО финансового управляющего", "Срок приема требований", None, "Адрес приема требований", "Контактные данные",
               "Дата размещения объявления"])
    ru.append([None, None, None, None, None, None, None, "с", "до"])
    for r in rows:
        kz.append(r)
        ru.append([r[0], r[1], r[2], r[3], "СМЭС г. Астаны"] + r[5:])
    wb.save(path)


def test_kazakh_first_sheet_and_translation_counted_once(tmp_path):
    path = tmp_path / "ann2.xlsx"
    two_sheet_xlsx(path, [[1, "Тестов Тест", P1, "Астана", "Астана қаласының соты", "02.11.2026", "Басқарушы А.", "05.11.2026", "05.01.2027", "Астана", "8 700", "05.11.2026"],
                          [2, "Примеров Пример", P3, "Астана", "Астана қаласының соты", "03.11.2026", "Басқарушы А.", "06.11.2026", "06.01.2027", "Астана", "8 700", "06.11.2026"]])
    ann = files.parse_announcements(files.file_rows(path).rows)
    assert ann["sheets"] == 2 and len(ann["records"]) == 2                     # лист-перевод не удваивает строки
    assert {r["sheet"] for r in ann["records"]} == {"ВП (рус.яз)"}             # русский лист первым
    assert ann["records"][0]["court"] == "СМЭС г. Астаны" and ann["records"][0]["to"] == date(2027, 1, 5)
    kz_only = [r for r in files.file_rows(path).rows if r.sheet == "ВП (гос.яз)"]
    kz_ann = files.parse_announcements(kz_only)                               # казахский заголовок узнаётся сам
    assert kz_ann and {"iin", "fio", "manager", "address", "claimsAddress", "posted", "claims", "to"} <= set(kz_ann["roles"])


def site_with(table: bytes, title: str, extra: list | None = None) -> dict:
    """gov.kz одного региона: одна публикация с таблицей (и, по желанию, ещё публикации того же раздела)."""
    base = "https://www.gov.kz/org/kgd-zhetysu"
    api = "https://www.gov.kz/api/v1/public/content-manager/documents?id=eq%3A{}&projects=eq%3Akgd-zhetysu"
    html = "text/html; charset=utf-8"
    docs = [{"id": 501, "title": title, "date": "2026-11-12T10:00:00+05:00"}] + [d for d, _ in (extra or [])]
    pages = {
        base + "/activities/directions?lang=ru": (turbo([query("getActivityDirections", {"directions": [
            {"id": 1, "title": "Восстановление платежеспособности и банкротство граждан", "parent_direction": None}]})]), html),
        base + "/activities/1?lang=ru": (turbo([query("fetchArticleByDirection", {"documents": docs, "article": {"content": ""}}),
                                                query("getChildDirections", {"directions": []})]), html),
        api.format(501): (json.dumps({"full_text": [{"document": "/uploads/2026/11/12/aaa111_original.1.xlsx"}]}), "application/json"),
        "https://www.gov.kz/uploads/2026/11/12/aaa111_original.1.xlsx": (table, "application/octet-stream"),
    }
    for d, (path, body) in (extra or []):
        pages[api.format(d["id"])] = (json.dumps({"full_text": [{"document": path}]}), "application/json")
        pages["https://www.gov.kz" + path] = (body, "application/octet-stream")
    return pages


def test_bankrupt_table_is_not_an_announcement_and_duplicates_across_files(tmp_path):
    bankrupt = Workbook()
    ws = bankrupt.active
    ws.append(["№", "ФИО банкрота", "ИИН банкрота", "Адрес", "Наименование суда", "Дата решения суда о признании банкротом", "Дата размещения"])
    ws.append([1, "Банкротов Банк", P2, "Талдыкорган", "СМЭС", "10.11.2026", "12.11.2026"])
    import io
    buf = io.BytesIO()
    bankrupt.save(buf)
    same_again = announcement_xlsx(ANN_ROWS[:2] + [[9, "Будущев Будущий", P4, "г. Актау", "СМЭС", "04.11.2026", "", "", "", "", "", "07.11.2027"]])
    pages = site_with(buf.getvalue(), "Объявление о признании банкротом", extra=[
        ({"id": 502, "title": ANN_TITLE, "date": "2026-11-05T10:00:00+05:00"}, ("/uploads/2026/11/5/bbb222_original.2.xlsx", announcement_xlsx(ANN_ROWS[:2]))),
        ({"id": 503, "title": ANN_TITLE + " (обновлено)", "date": "2026-11-07T10:00:00+05:00"}, ("/uploads/2026/11/7/ccc333_original.3.xlsx", same_again))])
    govkz.harvest_region(FakeGov(pages), "kgd-zhetysu", "Жетысуская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    dg = dgd.read_dgd(tmp_path, use_ocr=False, log=lambda t: None)
    assert [x["kind"] for x in dg["by_iin"][P2]] == ["объявление о признании банкротом"]      # не «о возбуждении»
    assert P2 not in {a["iin"] for a in dg["announcements"]}
    assert len(dg["announcements"]) == 3                                                    # P1 и P3 из двух файлов: по разу, плюс P4
    zh = next(r for r in dg["regions"] if r["name"] == "Жетысуская обл.")
    assert "дата размещения в будущем: 1" in zh["text"]
    assert zh["year"] == 2                                                                  # таблица о признании банкротом не объявление


def test_register_with_two_iins_marks_the_second(tmp_path):
    Document = pytest.importorskip("docx").Document          # python-docx: только для этого теста
    d = Document()
    d.add_paragraph(f"Реестр требований кредиторов по должнику: Тестов Тест Тестович, ИИН {P1}")
    for _ in range(30):
        d.add_paragraph("Кредитор: АО «Банк», сумма 100 000 тенге")
    d.add_paragraph(f"Финансовый управляющий Управляев У. У., ИИН {P2}")
    import io
    buf = io.BytesIO()
    d.save(buf)
    pages = site_with(announcement_xlsx(ANN_ROWS[:1]), ANN_TITLE, extra=[
        ({"id": 600, "title": "Реестр требований кредиторов в процедуре восстановления платежеспособности", "date": "2026-11-10"},
         ("/uploads/2026/11/10/ddd444_original.4.docx", buf.getvalue()))])
    govkz.harvest_region(FakeGov(pages), "kgd-zhetysu", "Жетысуская обл.", tmp_path / "ДГД", YEAR, log=lambda t: None)
    dg = dgd.read_dgd(tmp_path, use_ocr=False, log=lambda t: None)
    assert dg["by_iin"][P1][-1]["conf"] == "high" and dg["by_iin"][P2][0]["conf"] == "medium"
    zh = next(r for r in dg["regions"] if r["name"] == "Жетысуская обл.")
    assert "в реестре требований больше одного ИИН: 1" in zh["text"]


def test_glued_iin_error_cells_and_kazakhstan_dates(tmp_path):
    rows = [files.Row(1, [files.Cell(f"Должник Тестов Т.Т., ИИН {P1}2400000 тенге", False)])]
    assert [x["iin"] for x in files.extract_iins(rows)["items"]] == [P1]
    from tazalau_daily.common import parse_day
    assert parse_day("2026-04-30T19:00:00Z") == date(2026, 5, 1) and parse_day("2026-04-30T18:59:00Z") == date(2026, 4, 30)
    assert parse_day("2026-06-11") == date(2026, 6, 11)
    wb = Workbook()
    wb.active.append(["#REF!", "текст"])
    wb.save(tmp_path / "e.xlsx")
    cells = files.file_rows(tmp_path / "e.xlsx").rows[0].cells
    assert cells[0] is None and cells[1].v == "текст"


def test_gov_client_waits_and_retries_busy_server(monkeypatch):
    import io
    import urllib.error
    client = govkz.GovClient(spacing=0, log=lambda t: None)
    answers = [urllib.error.HTTPError("https://www.gov.kz/x", 503, "busy", {}, io.BytesIO(b"")), None]

    class Ok:
        status = 200
        headers = {"Content-Type": "application/pdf"}

        def read(self, n):
            return b"%PDF"

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_open(req, timeout=None):
        a = answers.pop(0)
        if a is not None:
            raise a
        return Ok()

    monkeypatch.setattr(client._opener, "open", fake_open)
    monkeypatch.setattr(govkz.time, "sleep", lambda s: None)
    r = client.get("https://www.gov.kz/uploads/2026/1/1/x.pdf")
    assert r.ok and r.body == b"%PDF" and client.requests == 2
    with pytest.raises(ValueError):
        client.get("https://example.org/secret")                              # только www.gov.kz


@pytest.mark.parametrize("name", ["Tazalau_ежедневная_проверка.ipynb", "Tazalau_ежедневная_проверка_один_файл.ipynb"])
def test_notebook_is_saved_without_outputs(name):
    nb = json.loads((Path(__file__).resolve().parent.parent / "notebooks" / name).read_text(encoding="utf-8"))
    code = [c for c in nb["cells"] if c["cell_type"] == "code"]
    assert code and all(not c.get("outputs") and c.get("execution_count") is None for c in code)


def test_single_file_notebook_carries_current_code():
    """Код внутри ноутбука «одним файлом» совпадает с пакетами проекта (иначе: python tools/build_notebook.py)."""
    import base64
    import re as _re
    import zlib
    root = Path(__file__).resolve().parent.parent
    nb = json.loads((root / "notebooks" / "Tazalau_ежедневная_проверка_один_файл.ipynb").read_text(encoding="utf-8"))
    loader = "".join(nb["cells"][1]["source"])
    src = json.loads(zlib.decompress(base64.b64decode(_re.search(r'PAYLOAD = "([^"]+)"', loader).group(1))).decode("utf-8"))
    assert "tazalau_daily/run.py" in src and "tazalau/capture.py" in src
    stale = [rel for rel, text in src.items() if (root / rel).read_text(encoding="utf-8") != text]
    missing = [f"tazalau_daily/{p.name}" for p in (root / "tazalau_daily").glob("*.py") if f"tazalau_daily/{p.name}" not in src]
    assert not stale and not missing, (stale, missing)
