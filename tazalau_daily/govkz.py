"""gov.kz: публикации ДГД о банкротстве и восстановлении платежеспособности граждан, все 20 департаментов.

Путь, как у человека: www.gov.kz/org/kgd-<регион> › «Деятельность» › раздел о восстановлении платежеспособности и
банкротстве граждан (названия у регионов разные, у кого-то он внутри «Реабилитация и банкротство») › год ›
«Информационные сообщения» › «Объявление о возбуждении производства…» › файл xlsx. Файлы лежат на любом уровне:
документом раздела или ссылкой прямо в тексте раздела (/uploads/…). Если в «Деятельности» объявлений нет, они ищутся
в разделе «Документы».

Данные страницы gov.kz приходят внутри неё (React Router, поток turbo-stream), файлы документа отдаёт открытый
справочник /api/v1/public/content-manager/documents. Адрес файла содержит отпечаток содержимого: тот же адрес = тот же
файл, повторно он не скачивается. Запросы идут по одному, не чаще одного в 1,5 секунды.

Складывает так же, как закладка-помощник: <рабочая папка>/ДГД/<регион>/файлы + «публикации_gov_kz.json»."""
from __future__ import annotations

import json
import os
import re
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from . import VERSION
from .common import (LO, UP, clean, iin_ok, iso, max_date, parse_day, personal_iin, region_folder, rx, safe_name,
                     strip_tags)

GOV = "https://www.gov.kz"
# Адреса департаментов на gov.kz (проверено 20.11.2026: все 20 открываются).
GOV_SITES = [
    ("kgd-abay", "Абайская обл."), ("kgd-akmola", "Акмолинская обл."), ("kgd-aktobe", "Актюбинская обл."),
    ("kgd-alm", "Алматинская обл."), ("kgd-atyrau", "Атырауская обл."), ("kgd-vko", "Восточно-Казахстанская обл."),
    ("kgd-zhambyl", "Жамбылская обл."), ("kgd-zhetysu", "Жетысуская обл."), ("kgd-zko", "Западно-Казахстанская обл."),
    ("kgd-karaganda", "Карагандинская обл."), ("kgd-kostanay", "Костанайская обл."), ("kgd-kyzylorda", "Кызылординская обл."),
    ("kgd-mangistau", "Мангистауская обл."), ("kgd-pavlodar", "Павлодарская обл."), ("kgd-sko", "Северо-Казахстанская обл."),
    ("kgd-turkistan", "Туркестанская обл."), ("kgd-ulytau", "Улытауская обл."), ("kgd-astana", "г. Астана"),
    ("kgd-almaty", "г. Алматы"), ("kgd-shymkent", "г. Шымкент"),
]
DGD_DIR = "ДГД"
DGD_META = "публикации_gov_kz.json"     # в папке региона: что найдено на gov.kz при последнем сборе
MAX_FILE = 60 * 1024 * 1024

# ---------------------------------------------------------------- что за публикация (по названию)
ANN_TITLE_RE = rx(r"объявлени\S*\s+о\s+возбуждени|іс\s+қозғау|хабарландыру")
GOV_NOT_RE = rx(r"реабилитац|банкротн\S*\s+управляющ|временн\S*\s+управляющ|<<тоо>>|<<ао>>|товариществ|<<llp>>|ликвидац|"
                r"оценщик|конкурс|ваканс|список\s+(?:лиц|финансов|администратор)|сведени\S*\s+о\s+направленн|"
                r"собрани\S*\s+кредитор|<<ип>>")
GOV_ROOT_RE = rx(r"банкрот|платеж|оңалту|төлем\s+қабілет")
GOV_ROOT_CITIZEN_RE = rx(r"платежеспособ|граждан|физическ|азамат|төлем")
GOV_CITIZEN_RE = rx(r"платежеспособ|граждан|физическ|физ\.?\s*лиц|азамат|төлем|информацион|хабарлама|объявлени|реестр|ртк")
GOV_LEGAL_RE = rx(r"юридическ|заңды\s+тұлға")
UPLOAD_RE = re.compile(r"/uploads/\d{4}/\d{1,2}/\d{1,2}/[^\"'\s<>\\)]+\.(?:xlsx|xlsm|xls|docx|doc|pdf|jpe?g|png|zip|rar|csv)", re.I)

_N1 = f"[{UP}][{LO}ʼ’'-]+"
_INI = f"[{UP}]\\.\\s*(?:[{UP}]\\.?)?"
NAME_ANY = f"{_N1}\\s+{_INI}|{_N1}(?:\\s+{_N1}){{1,3}}"
UPPER_NAME = f"[{UP}]{{2,}}(?:[\\s-]+[{UP}]{{2,}}){{1,3}}"
PERSON_TITLE_RE = re.compile(f"^[«\"]?(?:[{UP}][{LO}ʼ’'-]+(?:\\s+[{UP}][{LO}ʼ’'-]+){{1,3}}|[{UP}]{{2,}}(?:\\s+[{UP}]{{2,}}){{1,3}})[»\"]?$")
NAME_ONLY = re.compile(f"^[«\"]?(?:{NAME_ANY})[»\"]?$")


def _after(lead: str) -> re.Pattern:
    return re.compile(lead + f"\\s*[«\"]?({NAME_ANY})")


DEBTOR_RES = [_after(r"[Пп]о\s+заявлени\S*\s+(?:должник\S*\s+|гражданин\S*\s+|гр\.\s*|ФЛ\s+)?"),
              _after(r"[Вв]\s+отношении\s+(?:ФЛ\s+|физическ\S+\s+лиц\S*\s+|гражданин\S*\s+|должник\S*\s+)?"),
              _after(r"(?:^|\s)ФЛ\s+")]
MANAGER_RE = _after(r"финансов\S*\s+управляющ\S*\s+")


def ann_citizen(title: str) -> bool:
    """Объявление о возбуждении дела гражданина (не юрлица)."""
    return bool(ANN_TITLE_RE.search(title) and not GOV_NOT_RE.search(title)
                and re.search(r"платежеспособ|граждан|судебн\S*\s+банкрот|физ|төлем|азамат", title, re.I))


def title_iins(title: str) -> list[str]:
    """ИИН прямо в названии: «Реестр требований кредиторов … по должнику: ФИО, ИИН 000000000000»."""
    out: list[str] = []
    for m in re.finditer(r"(?<!\d)(\d{12})(?!\d)", title or ""):
        if iin_ok(m.group(1)) and personal_iin(m.group(1)) and m.group(1) not in out:
            out.append(m.group(1))
    return out


def order_info(cells: list[str]) -> dict:
    """Дата, вид, должник, управляющий и номер из текста приказа или названия публикации."""
    today = date.today()
    text = " ".join(cells)
    dates = []
    for m in re.finditer(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})(?!\d)", text):
        y = 2000 + int(m.group(3)) if len(m.group(3)) == 2 else int(m.group(3))
        try:
            d = date(y, int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        if 2020 <= y <= 2100:
            dates.append(d)
    ok = [d for d in dates if d <= date.fromordinal(today.toordinal() + 1)]
    low = text.lower()
    kind = ("отстранение" if re.search(r"отстран|shettet|шеттет", low) else "освобождение" if re.search(r"освобожд|bosat|босат", low)
            else "назначение" if re.search(r"назнач|tagay|тағай", low) else "")
    debtor = manager = ""
    for c in cells:
        if not debtor:
            for pattern in DEBTOR_RES:
                m = pattern.search(c)
                if m:
                    debtor = clean(m.group(1))
                    break
        if not manager:
            m = MANAGER_RE.search(c)
            if m:
                manager = clean(m.group(1))
    for c in cells:                       # ячейка только с именем: должник, а если он уже есть, то управляющий
        v = re.sub(r"[«»\"]", "", c).strip()
        if not NAME_ONLY.match(v) or v in (debtor, manager):
            continue
        if not debtor:
            debtor = v
        elif not manager:
            manager = v
    num = re.search(r"(?:№\s*|(?:^|\s)П-\s*)(\d[\d/-]*[А-Яа-я]?)", text)
    return {"date": max_date(ok), "kind": kind, "debtor": debtor, "manager": manager,
            "number": num.group(1) if num else "", "future": len(dates) > len(ok)}


def gov_kind(title: str) -> str:
    t = str(title or "")
    low = t.lower()
    proc = ("восстановление платежеспособности" if re.search(r"восстановлени\S*\s+платежеспособ|(?:^|\s)вп(?:\s|$)|төлем\s+қабілет", low)
            else "судебное банкротство" if re.search(r"судебн\S*\s+банкрот|(?:^|\s)сб(?:\s|$)", low) else "")
    kind = "публикация"
    if re.search(r"приказ|бұйрық", low):
        k = order_info([t])["kind"]
        kind = "приказ" + (f" ({k})" if k else "")
    elif re.search(r"(?:^|\s)ртк(?:\s|$)|реестр\S*\s+требовани", low):
        kind = "реестр требований кредиторов"
    elif re.search(r"объявлени\S*\s+о\s+возбуждени", low):
        kind = "объявление о возбуждении дела"
    elif re.search(r"объявлени\S*\s+о\s+признани", low):
        kind = "объявление о признании банкротом"
    elif re.search(r"объявлени\S*\s+о\s+(?:применени|завершени|прекращени)", low):
        word = re.search(r"о\s+(\S+)", low[low.index("объявлени"):]).group(1)
        kind = "объявление о " + re.sub(r"и$", "и процедуры", word)
    return f"{kind}, {proc}" if proc and kind != "публикация" else kind


def gov_debtor(title: str) -> str:
    """ФИО должника из названия: «РТК должник ФАМИЛИЯ ИМЯ», «по заявлению должника …», «… ВП Фамилия И.»."""
    t = clean(title)
    if PERSON_TITLE_RE.match(t):
        return re.sub(r"[«»\"]", "", t)
    for pattern in (re.compile(f"(?:должник\\S*|борышкер)\\s+[«\"]?({UPPER_NAME}|{NAME_ANY})"),
                    re.compile(f"(?:^|\\s)(?:ВП|СБ)\\s+({UPPER_NAME}|{NAME_ANY})")):
        m = pattern.search(t)
        if m:
            return clean(m.group(1))
    o = order_info([t])
    if o["debtor"]:
        return o["debtor"]
    tail = re.search(f"({UPPER_NAME}|{_N1}\\s+{_N1}(?:\\s+{_N1})?)\\s*$", re.sub(r"[.,;)»\"]+$", "", t))
    return clean(tail.group(1)) if tail and not re.match(r"^(Республики|Казахстан)", tail.group(1)) else ""


# ---------------------------------------------------------------- данные страницы gov.kz
def gov_data(html: str) -> dict:
    """Данные, встроенные в страницу gov.kz: { имя запроса: { endpointName, arg, value } }.
    Поток turbo-stream: плоский массив значений, объекты ссылаются на другие значения по номеру,
    отложенные части («P12:[…]») дописываются в конец."""
    text = ""
    for m in re.finditer(r'streamController\.enqueue\(("(?:[^"\\]|\\.)*")\)', html or ""):
        try:
            text += json.loads(m.group(1))
        except ValueError:
            pass
    lines = [line for line in text.split("\n") if line.strip()]
    if not lines:
        return {}
    try:
        vals = json.loads(lines[0])
    except ValueError:
        return {}
    if not isinstance(vals, list):
        return {}
    promises: dict[str, int] = {}
    for line in lines[1:]:
        m = re.match(r"^([PE])(\d+):(.*)$", line, re.S)
        if not m:
            continue
        try:
            arr = json.loads(m.group(3))
        except ValueError:
            continue
        promises[m.group(2)] = len(vals)
        vals.extend(arr if isinstance(arr, list) else [arr])
    memo: dict[int, object] = {}

    def dec(i, depth):
        if not isinstance(i, int) or isinstance(i, bool) or depth > 200:
            return None
        if i < 0:
            return 0 if i == -4 else None
        if i in memo:
            return memo[i]
        if i >= len(vals):
            return None
        v = vals[i]
        if isinstance(v, list):
            if v and isinstance(v[0], str) and len(v[0]) == 1 and v[0] in "DRSMYEPZBU":
                if v[0] == "P":
                    r = dec(promises[str(v[1])], depth + 1) if str(v[1]) in promises else None
                elif v[0] == "D":
                    r = v[1]
                else:
                    r = None
                memo[i] = r
                return r
            r = []
            memo[i] = r
            r.extend(dec(x, depth + 1) for x in v)
            return r
        if isinstance(v, dict):
            r = {}
            memo[i] = r
            for k, x in v.items():
                key = str(vals[int(k[1:])]) if k.startswith("_") and k[1:].isdigit() and int(k[1:]) < len(vals) else k
                r[key] = dec(x, depth + 1)
            return r
        memo[i] = v
        return v

    out: dict = {}
    seen: set[int] = set()

    def walk(o, depth):
        if not isinstance(o, (dict, list)) or id(o) in seen or depth > 60:
            return
        seen.add(id(o))
        if isinstance(o, dict) and isinstance(o.get("endpointName"), str) and o["endpointName"] not in out:
            out[o["endpointName"]] = o
        for x in (o if isinstance(o, list) else o.values()):
            walk(x, depth + 1)

    walk(dec(0, 0), 0)
    return out


def _value(data: dict, endpoint: str) -> dict:
    v = (data.get(endpoint) or {}).get("value")
    return v if isinstance(v, dict) else {}


# ---------------------------------------------------------------- запросы
class GovBlocked(Exception):
    """gov.kz перестал отвечать этому адресу (ограничение частоты запросов)."""


@dataclass
class Response:
    status: int
    body: bytes = b""
    content_type: str = ""
    error: str = ""

    @property
    def ok(self) -> bool:
        return 200 <= self.status < 300

    def text(self) -> str:
        return self.body.decode("utf-8", "replace")


class GovClient:
    """Обычные запросы к www.gov.kz по одному: не чаще одного в spacing секунд, повтор при обрыве,
    длинная пауза, если сайт ограничил частоту (403/429), и остановка, если ограничение не снимается."""

    def __init__(self, spacing: float = 1.5, timeout: float = 90, retries: int = 3, proxy: str | None = None,
                 ca_file: str | None = None, user_agent: str | None = None, log: Callable[[str], None] = print):
        self.spacing, self.timeout, self.retries, self.log = spacing, timeout, retries, log
        handlers: list = []
        if proxy:
            handlers.append(urllib.request.ProxyHandler({"http": proxy, "https": proxy}))
        context = ssl.create_default_context(cafile=ca_file) if ca_file else None
        if context is not None:
            handlers.append(urllib.request.HTTPSHandler(context=context))
        self._opener = urllib.request.build_opener(*handlers)
        self._headers = {"Accept-Language": "ru"}
        if user_agent:
            self._headers["User-Agent"] = user_agent
        self._last = 0.0
        self.requests = 0
        self.bytes = 0

    def _pace(self) -> None:
        wait = self.spacing - (time.monotonic() - self._last)
        if wait > 0:
            time.sleep(wait)
        self._last = time.monotonic()

    def get(self, url: str, *, timeout: float | None = None, limit: int = MAX_FILE) -> Response:
        if not url.startswith(GOV + "/"):
            raise ValueError("запросы только к www.gov.kz")
        blocked = 0
        attempt = 0
        while True:
            self._pace()
            self.requests += 1
            try:
                req = urllib.request.Request(url, headers=self._headers)
                with self._opener.open(req, timeout=timeout or self.timeout) as r:
                    body = r.read(limit + 1)
                    if len(body) > limit:
                        return Response(r.status, b"", r.headers.get("Content-Type", ""), "файл больше %d МБ" % (limit // 2 ** 20))
                    self.bytes += len(body)
                    return Response(r.status, body, r.headers.get("Content-Type", ""))
            except urllib.error.HTTPError as exc:
                if exc.code in (403, 429):
                    blocked += 1
                    if blocked > 2:
                        raise GovBlocked("gov.kz ограничил запросы с этого адреса (код %d). Повторите сбор через 15-30 минут." % exc.code)
                    pause = 60 * blocked
                    self.log(f"   gov.kz ответил кодом {exc.code}: жду {pause} с и повторяю")
                    time.sleep(pause)
                    continue
                if exc.code >= 500 and attempt < self.retries:
                    attempt += 1
                    time.sleep(8 * attempt)
                    continue
                return Response(exc.code, b"", "", f"код {exc.code}")
            except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as exc:
                attempt += 1
                if attempt > self.retries:
                    return Response(0, b"", "", f"нет связи с gov.kz ({getattr(exc, 'reason', exc)})")
                time.sleep(8 * attempt)

    def page(self, url: str) -> dict:
        r = self.get(url)
        return gov_data(r.text()) if r.ok else {}


# ---------------------------------------------------------------- обход «Деятельности»
@dataclass
class Tree:
    roots: list[str] = field(default_factory=list)
    nodes: list[dict] = field(default_factory=list)
    items: list[dict] = field(default_factory=list)


def tree_docs(client: GovClient, slug: str, year: str, log: Callable[[str], None] = print) -> Tree:
    """Все документы и файлы ветки о гражданах этого года в «Деятельности» департамента."""
    base = f"{GOV}/org/{slug}"
    r = client.get(base + "/activities/directions?lang=ru")
    if not r.ok:
        raise RuntimeError("раздел «Деятельность» на gov.kz не открылся" + (f" ({r.error})" if r.error else ""))
    dirs = _value(gov_data(r.text()), "getActivityDirections").get("directions") or []
    top = [d for d in dirs if not ((d.get("parent_direction") or {}).get("id")) and GOV_ROOT_RE.search(d.get("title") or "")
           and not GOV_LEGAL_RE.search(d.get("title") or "")]
    citizen = [d for d in top if GOV_ROOT_CITIZEN_RE.search(d.get("title") or "")]

    def year_ok(title: str) -> bool:
        years = re.findall(r"20\d\d", title or "")
        return not years or year in years

    roots = [d for d in (citizen or top) if year_ok(d.get("title") or "")]
    tree = Tree(roots=[clean(d.get("title")) for d in roots])
    items: dict[str, dict] = {}
    seen: set[str] = set()

    def visit(did, path: list[str], depth: int) -> None:
        if str(did) in seen or depth > 6 or len(seen) > 60:
            return
        seen.add(str(did))
        data = client.page(f"{base}/activities/{did}?lang=ru")
        if not data:
            return
        art = _value(data, "fetchArticleByDirection")
        kids = _value(data, "getChildDirections").get("directions") or []
        tree.nodes.append({"id": str(did), "path": path})
        for x in art.get("documents") or []:
            key = f"doc:{x.get('id')}"
            if key not in items:
                items[key] = {"key": key, "id": str(x.get("id")), "title": clean(x.get("title")), "date": parse_day(x.get("date")),
                              "path": path, "doc": True}
        content = ((art.get("article") or {}).get("content")) or ""
        for m in UPLOAD_RE.finditer(content):              # файл ссылкой прямо в тексте раздела
            key = "file:" + m.group(0)
            near = strip_tags(content[max(0, m.start() - 600):m.start()])[-160:]
            if key not in items:
                title = clean(art.get("title") or path[-1]) + (f" › {near}" if len(near) > 8 else "")
                items[key] = {"key": key, "id": str(did), "title": title, "date": None, "path": path, "url": GOV + m.group(0)}
        for k in kids:
            t = k.get("title") or ""
            years = re.findall(r"20\d\d", t)
            if years and year not in years:
                continue
            if GOV_LEGAL_RE.search(t):
                continue
            if not (year in years or GOV_CITIZEN_RE.search(t) or GOV_CITIZEN_RE.search(" ".join(path))):
                continue
            visit(k.get("id"), path + [clean(t)], depth + 1)

    for d in roots:
        visit(d.get("id"), [clean(d.get("title"))], 0)
    tree.items = list(items.values())
    return tree


def doc_files(client: GovClient, slug: str, doc_id: str) -> list[str]:
    """Адреса файлов документа из открытого справочника gov.kz (content-manager)."""
    q = urllib.parse.urlencode({"id": f"eq:{doc_id}", "projects": f"eq:{slug}"})
    r = client.get(f"{GOV}/api/v1/public/content-manager/documents?{q}", timeout=60)
    out: list[str] = []
    if not r.ok:
        return out
    strings: list[str] = []
    try:
        def walk(o):
            if isinstance(o, str):
                strings.append(o)
            elif isinstance(o, dict):
                for x in o.values():
                    walk(x)
            elif isinstance(o, list):
                for x in o:
                    walk(x)
        walk(json.loads(r.text()))
    except ValueError:
        strings = [r.text().replace("\\/", "/")]
    for s in strings:
        for m in UPLOAD_RE.finditer(s):
            url = GOV + m.group(0)
            if url not in out:
                out.append(url)
    return out


def search_documents(client: GovClient, slug: str, year: str, have: set[str]) -> list[dict]:
    """Запасной путь: объявления о возбуждении дел граждан этого года в разделе «Документы» (поиск по названию)."""
    base = f"{GOV}/org/{slug}"
    for q in ("возбуждении производства", "возбуждении"):
        r = client.get(f"{base}/documents/1?lang=ru&size=50&title=" + urllib.parse.quote(q))
        data = _value(gov_data(r.text()), "getFilteredData") if r.ok else {}
        found = []
        for x in data.get("data") or []:
            t, d = clean(x.get("title")), parse_day(x.get("created_date"))
            key = f"doc:{x.get('id')}"
            if ann_citizen(t) and ((d and str(d.year) == year) or year in t) and key not in have:
                found.append({"key": key, "id": str(x.get("id")), "title": t, "date": d, "path": ["Документы"], "doc": True,
                              "found": "Документы (поиск)"})
        if found:
            return found
    return []


# ---------------------------------------------------------------- один регион
def _read_json(path: Path) -> dict | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _write_atomic(path: Path, data: bytes) -> None:
    tmp = path.with_name(path.name + ".part")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def write_json(path: Path, data) -> None:
    _write_atomic(path, json.dumps(data, ensure_ascii=False, indent=1, default=str).encode("utf-8"))


def harvest_region(client: GovClient, slug: str, region: str, dgd_dir: Path, year: str | None = None,
                   log: Callable[[str], None] = print, download_all: bool = False) -> dict:
    """Публикации о гражданах одного департамента за год: файлы в ДГД/<регион>/, список в «публикации_gov_kz.json»."""
    year = year or str(date.today().year)
    base = f"{GOV}/org/{slug}"
    folder = dgd_dir / region_folder(region)
    folder.mkdir(parents=True, exist_ok=True)
    old = _read_json(folder / DGD_META) or {"rows": []}
    saved_before = {f["url"]: f["saved"] for r in old.get("rows") or [] for f in r.get("files") or [] if f.get("saved") and f.get("url")}
    first_seen = {r.get("key"): r.get("firstSeen") for r in old.get("rows") or []}
    tree = tree_docs(client, slug, year, log)
    via = "Деятельность"
    items = [x for x in tree.items if not GOV_NOT_RE.search(x["title"]) or title_iins(x["title"])]
    if not any(ANN_TITLE_RE.search(x["title"]) for x in items):
        extra = search_documents(client, slug, year, {x["key"] for x in items})
        if extra:
            items += extra
            via = "Деятельность + Документы (поиск)" if tree.items else "Документы (поиск)"
    rows, got = [], 0
    for x in items:
        row = {"key": x["key"], "id": x["id"], "date": x["date"], "kind": gov_kind(x["title"]), "debtor": gov_debtor(x["title"]),
               "text": x["title"], "page": f"{base}/documents/details/{x['id']}?lang=ru" if x.get("doc") else f"{base}/activities/{x['id']}?lang=ru",
               "firstSeen": first_seen.get(x["key"]) or date.today().isoformat(), "files": [], "notes": [],
               "path": " › ".join(x["path"]), "via": x.get("found") or "Деятельность", "titleIins": title_iins(x["title"])}
        rows.append(row)
        urls = [x["url"]] if x.get("url") else []
        if x.get("doc"):
            urls = doc_files(client, slug, x["id"])
            if not urls:
                row["notes"].append("у публикации нет файла")
        for href in urls[:4]:
            was = saved_before.get(href)
            if was and (folder / was).exists():
                row["files"].append({"url": href, "name": was, "saved": was, "err": ""})
                continue
            if row["titleIins"] and not ANN_TITLE_RE.search(x["title"]) and not download_all:
                row["files"].append({"url": href, "name": "", "saved": "", "err": ""})      # ИИН уже в названии
                continue
            r = client.get(href, timeout=180)
            if not r.ok or not r.body:
                row["files"].append({"url": href, "name": "", "saved": "", "err": "файл не скачался" + (f" ({r.error})" if r.error else "")})
                continue
            if "text/html" in r.content_type.lower() and not href.lower().endswith((".htm", ".html")):
                row["files"].append({"url": href, "name": "", "saved": "", "err": "вместо файла пришла страница"})
                continue
            m = re.search(r"\.([a-z0-9]{2,4})$", href, re.I)
            ext = (m.group(1) if m else "bin").lower()
            stem = safe_name(row["debtor"])[:40] or safe_name(x["title"])[:60] or "публикация"
            name = f"{iso(x['date']) or year}_{stem}_{href.rsplit('/', 1)[-1][:10]}.{ext}"
            try:
                _write_atomic(folder / name, r.body)
            except OSError as exc:
                row["files"].append({"url": href, "name": "", "saved": "", "err": f"не записан: {exc}"})
                continue
            row["files"].append({"url": href, "name": name, "saved": name, "err": ""})
            got += 1
    meta = {
        "version": "py-" + VERSION, "source": "gov.kz", "region": region, "slug": slug, "state": "ok", "note": "",
        "collectedAt": datetime.now().astimezone().isoformat(timespec="seconds"), "year": year,
        "via": via if items else "", "roots": tree.roots, "nodes": len(tree.nodes),
        "pages": [{"url": base + "/activities/directions?lang=ru", "title": "Деятельность", "year": int(year), "rows": len(items), "read": True, "err": ""}],
        "rows": [{"key": r["key"], "id": r["id"], "date": iso(r["date"]), "dateFrom": "", "kind": r["kind"], "debtor": r["debtor"],
                  "manager": "", "number": "", "text": r["text"], "page": r["page"], "future": False, "firstSeen": r["firstSeen"],
                  "notes": r["notes"], "path": r["path"], "via": r["via"], "titleIins": r["titleIins"],
                  "files": [{"url": f["url"], "name": f["name"], "gov": False, "saved": f["saved"], "err": f["err"]} for f in r["files"]]}
                 for r in rows],
    }
    write_json(folder / DGD_META, meta)
    anns = sum(1 for r in rows if ANN_TITLE_RE.search(r["text"]))
    return {"region": region, "slug": slug, "state": "ok", "note": "", "publications": len(rows), "announcements": anns, "new_files": got,
            "via": meta["via"], "nodes": len(tree.nodes), "roots": len(tree.roots),
            "file_errors": sum(1 for r in rows for f in r["files"] if f["err"])}


def save_failure(dgd_dir: Path, slug: str, region: str, note: str, state: str = "fail") -> None:
    """Регион не собран: в своде это видно, публикации прошлого сбора остаются (с пометкой)."""
    folder = dgd_dir / region_folder(region)
    folder.mkdir(parents=True, exist_ok=True)
    old = _read_json(folder / DGD_META) or {"region": region, "slug": slug, "source": "gov.kz", "pages": [], "rows": []}
    old.update(state=state, note=note, failedAt=datetime.now().astimezone().isoformat(timespec="seconds"))
    write_json(folder / DGD_META, old)


def harvest_all(work_dir: Path, *, year: str | None = None, regions: list[str] | None = None, client: GovClient | None = None,
                log: Callable[[str], None] = print, download_all: bool = False) -> list[dict]:
    """Все 20 департаментов по очереди. regions: только эти (имена или адреса kgd-…). Возвращает итог по каждому."""
    dgd_dir = Path(work_dir) / DGD_DIR
    dgd_dir.mkdir(parents=True, exist_ok=True)
    client = client or GovClient(log=log)
    sites = [s for s in GOV_SITES if not regions or s[0] in regions or s[1] in regions]
    out: list[dict] = []
    started = time.monotonic()
    stop = ""
    for n, (slug, region) in enumerate(sites, start=1):
        if stop:
            save_failure(dgd_dir, slug, region, stop, "skip")
            out.append({"region": region, "slug": slug, "state": "skip", "note": stop})
            continue
        t0, r0 = time.monotonic(), client.requests
        try:
            res = harvest_region(client, slug, region, dgd_dir, year, log, download_all)
        except GovBlocked as exc:
            stop = str(exc)
            save_failure(dgd_dir, slug, region, stop, "skip")
            res = {"region": region, "slug": slug, "state": "skip", "note": stop}
        except Exception as exc:  # один регион не должен останавливать остальные
            note = f"сбор не удался: {exc}"
            save_failure(dgd_dir, slug, region, note)
            res = {"region": region, "slug": slug, "state": "fail", "note": note}
        res["seconds"] = round(time.monotonic() - t0)
        res["requests"] = client.requests - r0
        out.append(res)
        if res["state"] == "ok":
            log(f"{n:2}/{len(sites)} {region}: публикаций {res['publications']}, объявлений о возбуждении {res['announcements']}, "
                f"новых файлов {res['new_files']}" + (f", ошибок файлов {res['file_errors']}" if res["file_errors"] else "")
                + (f" ({res['via']})" if res["via"] and res["via"] != "Деятельность" else "") + f" · {res['seconds']} с")
        else:
            log(f"{n:2}/{len(sites)} {region}: {res['note']}")
    log(f"Готово за {round((time.monotonic() - started) / 60, 1)} мин: запросов к gov.kz {client.requests}, "
        f"скачано {round(client.bytes / 1e6, 1)} МБ.")
    return out
