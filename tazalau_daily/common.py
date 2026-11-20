"""Общие мелочи: чистка текста, шаблоны с границами слов для кириллицы и казахских букв, даты, ИИН, имена файлов.
Всё повторяет закладку-помощник (helper/tazalau_helper.js), чтобы оба инструмента одинаково понимали одни и те же файлы."""
from __future__ import annotations

import html as _html
import re
from datetime import date, datetime, timedelta, timezone

from tazalau.iin import checksum_ok, normalize_iin

# Буквы для границ слов: в шаблоне «<<» = перед этим местом нет буквы, «>>» = после него нет буквы.
LETTERS = "а-яёәіңғүұқөһa-z"
UP = "А-ЯЁӘІҢҒҮҰҚӨҺ"
LO = "а-яёәіңғүұқөһ"


def rx(pattern: str) -> re.Pattern:
    return re.compile(pattern.replace("<<", f"(?<![{LETTERS}])").replace(">>", f"(?![{LETTERS}])"), re.I)


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def strip_tags(fragment: str) -> str:
    return clean(_html.unescape(re.sub(r"<[^>]+>", " ", fragment or "")))


# ---------------------------------------------------------------- ИИН
def iin_ok(s: str) -> bool:
    return len(s) == 12 and s.isdigit() and checksum_ok(s)


def personal_iin(s: str) -> bool:
    """ИИН человека, а не БИН и не случайное число: месяц и день рождения настоящие, 7-я цифра (век и пол) от 0 до 6."""
    mo, d, c = int(s[2:4]), int(s[4:6]), int(s[6])
    return 1 <= mo <= 12 and 1 <= d <= 31 and c <= 6


def norm_iin(value: object) -> tuple[str, str]:
    """Значение ячейки -> (ИИН, '') | ('', причина) | ('', '') для пустой. Короткое число дополняется нулями слева."""
    p = normalize_iin(value)
    if p.empty:
        return "", ""
    return (p.iin, "") if p.iin else ("", p.error or "не ИИН")


IIN_IN_TEXT = re.compile(r"(?<!\d)(\d{12})(?!\d)")
SPACED_IIN = re.compile(r"(?<!\d)(?:\d{4}[ \xa0]\d{4}[ \xa0]\d{4}|\d{3}(?:[ \xa0]\d{3}){3}|\d{6}[ \xa0]\d{6}|(?:\d[ \xa0]){11}\d)(?!\d)")


def iins_in_text(text: str, *, spaced: bool = True) -> list[str]:
    """Все ИИН людей в тексте (контрольная цифра сходится), в порядке появления, без повторов.
    spaced: ещё и разбитые пробелами («8709 1630 1642»), если цельных нет."""
    out: list[str] = []
    for m in IIN_IN_TEXT.finditer(text or ""):
        if iin_ok(m.group(1)) and personal_iin(m.group(1)) and m.group(1) not in out:
            out.append(m.group(1))
    if spaced and not out:
        for m in SPACED_IIN.finditer(text or ""):
            x = re.sub(r"\D", "", m.group(0))
            if iin_ok(x) and personal_iin(x) and x not in out:
                out.append(x)
    return out


# ---------------------------------------------------------------- даты
def dates_in(text: str) -> list[date]:
    out: list[date] = []
    t = str(text or "")
    found = [(int(m.group(3)), int(m.group(2)), int(m.group(1)))
             for m in re.finditer(r"(?<!\d)(\d{1,2})[.\-_/](\d{1,2})[.\-_/](\d{4})(?!\d)", t)]
    found += [(int(m.group(1)), int(m.group(2)), int(m.group(3)))
              for m in re.finditer(r"(?<!\d)(\d{4})[.\-_](\d{2})[.\-_](\d{2})(?!\d)", t)]
    for y, mo, d in found:
        if 2000 <= y <= 2100:
            try:
                out.append(date(y, mo, d))
            except ValueError:
                pass
    return out


def max_date(values) -> date | None:
    good = [v for v in values if v]
    return max(good) if good else None


def excel_day(value: object) -> date | None:
    """Дата из ячейки: дата Excel, число дней Excel (20000..80000) или текст «дд.мм.гггг»."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 < value < 80000:
            return date(1899, 12, 30) + timedelta(days=int(value))
        return None
    return max_date(dates_in(str(value)))


KZ_TIME = timezone(timedelta(hours=5))      # Казахстан с 1 марта 2024 года живёт по UTC+5


def parse_day(value: object) -> date | None:
    """День публикации по времени Казахстана из ответа gov.kz: «2026-04-30T19:00:00Z»: это уже 1 мая в Алматы."""
    text = str(value or "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", text)
    if not m:
        return None
    try:
        when = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if when.tzinfo is not None:
            return when.astimezone(KZ_TIME).date()
    except ValueError:
        pass
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def dmy(d: date | None) -> str:
    return d.strftime("%d.%m.%Y") if d else ""


def iso(d: date | None) -> str:
    return d.isoformat() if d else ""


# ---------------------------------------------------------------- имена файлов
def safe_name(value: object) -> str:
    """Имя файла или папки, годное для Windows."""
    t = clean(re.sub(r'[\\/:*?"<>|#%\x00-\x1f]', " ", str(value or "")))[:100].rstrip(". ")
    return "_" + t if re.fullmatch(r"(con|prn|aux|nul|com[1-9]|lpt[1-9])", t, re.I) else t


def region_folder(name: str) -> str:
    """Windows не разрешает точку в конце имени папки: «Абайская обл.» -> «Абайская область»."""
    return re.sub(r" обл\.$", " область", name)


# ---------------------------------------------------------------- регионы
REGION_LIST = [  # порядок важен: «Алматинская» раньше «Алматы»
    ("Абайская обл.", rx(r"абайск|абай\s+облыс|област\S*\s+абай|<<семей>>"), rx(r"<<абай>>")),
    ("Акмолинская обл.", rx(r"акмолин|ақмол"), None),
    ("Актюбинская обл.", rx(r"актюбин|ақтөбе|<<актобе>>"), None),
    ("Алматинская обл.", rx(r"алматинск|алматы\s+облыс"), None),
    ("Атырауская обл.", rx(r"атырау"), None),
    ("Восточно-Казахстанская обл.", rx(r"восточно[-\s]*казахстан|шығыс\s+қазақстан|<<вко>>"), None),
    ("Жамбылская обл.", rx(r"жамбыл|<<тараз>>"), None),
    ("Жетысуская обл.", rx(r"жетысу|жетісу|<<талдыкорган"), None),
    ("Западно-Казахстанская обл.", rx(r"западно[-\s]*казахстан|батыс\s+қазақстан|<<зко>>|<<уральск"), None),
    ("Карагандинская обл.", rx(r"караганд|қарағанд"), None),
    ("Костанайская обл.", rx(r"костанай|қостанай"), None),
    ("Кызылординская обл.", rx(r"кызылорд|қызылорд"), None),
    ("Мангистауская обл.", rx(r"мангист|маңғыст|<<актау>>"), None),
    ("Павлодарская обл.", rx(r"павлодар"), None),
    ("Северо-Казахстанская обл.", rx(r"северо[-\s]*казахстан|солтүстік\s+қазақстан|<<ско>>|петропавл"), None),
    ("Туркестанская обл.", rx(r"туркестан|түркістан"), None),
    ("Улытауская обл.", rx(r"улытау|ұлытау|жезказган|жезқазған"), None),
    ("г. Астана", rx(r"астан|нур[-\s]*султан"), None),
    ("г. Алматы", rx(r"алматы"), None),
    ("г. Шымкент", rx(r"шымкент"), None),
]
REGION_NAMES = [r[0] for r in REGION_LIST]
UNKNOWN_REGION = "Регион не определён"


def regions_in(text: str, loose: bool = False) -> list[str]:
    """Регионы, названные в тексте. loose: имя файла или папки (там годится и голое «Абай»)."""
    t = str(text or "").lower()
    out: list[str] = []
    for name, pattern, extra in REGION_LIST:
        if (name == "г. Алматы" and "Алматинская обл." in out
                and not re.search(r"(?:^|[^а-яё])(?:г\.?\s*|город\S*\s+)алматы", t)):
            continue
        if pattern.search(t) or (loose and extra is not None and extra.search(t)):
            out.append(name)
    return out


# ---------------------------------------------------------------- фамилии
_KZ = str.maketrans({"ә": "а", "і": "и", "ң": "н", "ғ": "г", "ү": "у", "ұ": "у", "қ": "к", "ө": "о", "һ": "х", "ё": "е"})


def surname(value: object) -> str:
    """Фамилия для сравнения: строчные, казахские буквы как русские, без знаков."""
    parts = clean(value).split(" ")
    return re.sub(r"[^a-zа-я-]", "", parts[0].lower().translate(_KZ)) if parts and parts[0] else ""


def name_key(value: object) -> str:
    """Фамилия и первая буква имени: «Иванов Иван Иванович» и «ИВАНОВ И.И.» дают одно и то же."""
    parts = [p for p in re.split(r"[\s.]+", clean(value).lower().translate(_KZ)) if p]
    if len(parts) < 2:
        return ""
    return re.sub(r"[^a-zа-я-]", "", parts[0]) + " " + parts[1][:1]
