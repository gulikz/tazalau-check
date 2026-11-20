"""Собирает два ноутбука (без результатов: в репозиторий не попадают ИИН):
  notebooks/Tazalau_ежедневная_проверка.ipynb          : берёт пакеты из папки проекта;
  notebooks/Tazalau_ежедневная_проверка_один_файл.ipynb: тот же ноутбук с кодом проекта внутри (скрытая ячейка):
                                                          его можно скопировать на другой компьютер одним файлом.
usage: python tools/build_notebook.py"""
from __future__ import annotations

import base64
import copy
import json
import zlib
from pathlib import Path

import nbformat
from nbformat.v4 import new_code_cell, new_markdown_cell, new_notebook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "notebooks" / "Tazalau_ежедневная_проверка.ipynb"
OUT_ONE = ROOT / "notebooks" / "Tazalau_ежедневная_проверка_один_файл.ipynb"
# Код, который нужен ноутбуку: проверка tazalau (браузер, снимок) и весь ежедневный цикл
EMBED = ["tazalau/__init__.py", "tazalau/iin.py", "tazalau/config.py", "tazalau/registries.py", "tazalau/capture.py",
         "tazalau/stamp.py", "tazalau/fsutil.py"] + sorted(f"tazalau_daily/{p.name}" for p in (ROOT / "tazalau_daily").glob("*.py"))

cells = []
md = lambda s: cells.append(new_markdown_cell(s.strip()))
code = lambda s: cells.append(new_code_cell(s.strip()))

md("""
# Ежедневная проверка должников: ДГД → ADW → tazalau → итоговый Excel

Ноутбук делает весь день сам:

1. **ДГД.** Со всех 20 департаментов на gov.kz скачивает публикации о восстановлении платежеспособности и банкротстве граждан за год: «Деятельность» › раздел о банкротстве граждан › год › «Информационные сообщения» › «Объявление о возбуждении производства…» › xlsx. Если там объявлений нет, ищет их в «Документах». Затем проверяет каждый регион: нет публикаций, 1–2 ИИН, давно не обновлялся, файл не прочитан. Пишет свод `ДГД_свод.xlsx`.
2. **ADW.** Клиенты с признаком «проблема» (кредиты и кредитные карты): выгрузка `ADW*.xlsx/csv` или SQL к PostgreSQL. Проверяет выгрузку: неверные ИИН, пустые ФИО, разные ФИО у одного ИИН.
3. **Сверка.** Кто из клиентов есть в публикациях ДГД: по ИИН, а если ИИН в файле не прочитан, то «возможно» по ФИО.
4. **tazalau.qoldau.kz.** Каждый ИИН во всех шести реестрах (сначала клиенты из ДГД): снимок страницы, статус и даты записи.
5. **Итог.** `Итог_ГГГГ-ММ-ДД.xlsx`: строка на клиента со статусом, ДГД, ссылкой на снимок и ссылками на реестры.

**Как запустить (3 шага):**
1. В ячейке «Настройки» укажите рабочую папку `WORK_DIR` и положите в неё выгрузку ADW (или укажите `ADW_SQL`).
2. Меню **Run → Run All Cells** (или Kernel → Restart & Run All).
3. Откройте итоговый файл: ссылка будет в последней ячейке.

Повторный запуск в тот же день ничего не скачивает и не проверяет заново: файлы gov.kz узнаются по адресу, а проверенные ИИН записаны в журнал `tazalau_журнал.csv`.

> В рабочей папке и в результатах ячеек персональные данные (ИИН, ФИО). Не кладите их в общий доступ. Перед тем как отправить ноутбук, очистите результаты: Edit → Clear All Outputs.
""")

code("""
# Пакеты проекта: ноутбук лежит в папке notebooks/, пакеты: на уровень выше
import sys
from pathlib import Path

ROOT = Path.cwd().parent if Path.cwd().name == "notebooks" else Path.cwd()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import importlib.util
need = {"openpyxl": "openpyxl", "PIL": "pillow", "playwright": "playwright", "pandas": "pandas"}
missing = [pip for mod, pip in need.items() if importlib.util.find_spec(mod) is None]
optional = {"pypdf": "PDF с текстом", "xlrd": "старые .xls", "psycopg2": "SQL к ADW (или psycopg)"}
print("Не хватает пакетов: pip install " + " ".join(missing) if missing else "Обязательные пакеты на месте.")
for mod, why in optional.items():
    if importlib.util.find_spec(mod) is None:
        print(f"  по желанию: {mod}: {why}")

import pandas as pd
from tazalau_daily import VERSION
from tazalau_daily.run import Config, step_collect, step_dgd, step_adw, step_match, step_tazalau, step_report
pd.set_option("display.max_colwidth", 120)
print("tazalau_daily", VERSION, "· папка проекта:", ROOT)
""")

md("""
## Настройки

Меняется только эта ячейка. Пути можно писать как в Проводнике: `Path(r"D:\\Работа\\Tazalau")`.
""")

code("""
from datetime import date

WORK_DIR = Path.home() / "Tazalau_работа"   # рабочая папка: ADW, папка ДГД, снимки, итог
ADW_FILE = None             # None: самая свежая выгрузка ADW*.xlsx / ADW*.csv в рабочей папке; или Path(r"…\\ADW_2026-11-20.xlsx")
ADW_SQL = None              # или Path(r"…\\adw_problem.sql"): запрос к ADW (PostgreSQL); пароль не пишите сюда (спросит ниже)

COLLECT_GOV = True          # False: не ходить на gov.kz, взять то, что уже в папке ДГД
REGIONS = None              # None: все 20; или ["kgd-mangistau", "Павлодарская обл."]

CHECK_SCOPE = "all"         # "all": все клиенты ADW (из ДГД первыми); "dgd": только клиенты из публикаций ДГД
CHECK_LIMIT = None          # не больше стольких ИИН за запуск (None: без ограничения)
CHECK_MINUTES = None        # не дольше стольких минут на tazalau (остальные: при следующем запуске)
RECHECK_DAYS = 1            # ИИН, проверенный без ошибки меньше стольких дней назад, не проверяется снова

PROXY = None                # прокси банка, если нужен: "http://proxy.bank.local:8080"
SHOW_BROWSER = False        # True: видно окно браузера во время проверки tazalau

cfg = Config(work_dir=WORK_DIR, adw_file=ADW_FILE, adw_sql=ADW_SQL, collect_dgd=COLLECT_GOV, regions=REGIONS,
             check_scope=CHECK_SCOPE, check_limit=CHECK_LIMIT, check_minutes=CHECK_MINUTES, recheck_days=RECHECK_DAYS,
             proxy=PROXY, headless=not SHOW_BROWSER)
print("Рабочая папка:", cfg.work_dir)
""")

md("""
## Шаг 1. Публикации ДГД с gov.kz (20 регионов)

Около 20–30 минут при первом запуске (запросы идут по одному, не чаще одного в 1,5 секунды); при повторном: быстрее, уже скачанные файлы не скачиваются. Файлы ложатся в `ДГД/<регион>/`: так же, как у закладки-помощника, поэтому собирать можно и закладкой, и ноутбуком.
""")

code("""
collected = step_collect(cfg) if cfg.collect_dgd else []
if collected:
    display(pd.DataFrame(collected).reindex(columns=["region", "state", "publications", "announcements", "new_files", "via", "seconds", "note"])
            .rename(columns={"region": "Регион", "state": "Сбор", "publications": "Публикаций", "announcements": "Объявлений (по названию)",
                             "new_files": "Новых файлов", "via": "Где найдено", "seconds": "Секунд", "note": "Примечание"}).fillna(""))
""")

md("""
### Проверка регионов

Красный: регион не собран, публикаций за год нет или таблица объявлений не прочитана. Жёлтый: мало ИИН (1–2), последнее объявление старше 14 дней, ИИН с ошибкой в файле ДГД, файл без ИИН. Подробности и все строки объявлений: в `ДГД_свод.xlsx`.
""")

code("""
dg = step_dgd(cfg)
if dg:
    regions = pd.DataFrame([{
        "Регион": r["name"], "Проверка": r["text"], "Объявлений (файлов)": r["year"], "Строк": r["annRows"],
        "Уникальных ИИН": r["annUnique"] or r["unique"], "Последнее размещение": r["annPosted"] or r["listDate"], "_level": r["level"]}
        for r in dg["regions"]])
    colors = {"bad": "background-color:#FFC7CE", "warn": "background-color:#FFEB9C", "good": "background-color:#C6EFCE"}
    display(regions.drop(columns="_level").style.apply(
        lambda row: [colors[regions.loc[row.name, "_level"]] if c == "Проверка" else "" for c in row.index], axis=1))
""")

md("""
## Шаг 2. Клиенты ADW с признаком «проблема»

Выгрузка: столбец `ИИН/БИН` (12 цифр текстом) обязателен, `FIO` желателен, остальные (продукт, договор, просрочка, сумма) по желанию: как в `sql/ADW_пример_выгрузки.csv`. Для запроса к базе: `ADW_SQL` и строка подключения в переменной окружения `ADW_DSN` (или ячейка спросит её).
""")

code("""
import os, getpass
if cfg.adw_sql and not os.environ.get(cfg.adw_dsn_env):
    os.environ[cfg.adw_dsn_env] = getpass.getpass("Строка подключения к ADW (postgresql://пользователь:пароль@сервер:5432/база): ")
adw = step_adw(cfg)
display(pd.Series({k: v for k, v in adw.checks.items() if not isinstance(v, dict)}, name="Выгрузка ADW").to_frame())
if adw.checks.get("по продуктам (ИИН)"):
    display(pd.Series(adw.checks["по продуктам (ИИН)"], name="ИИН по продуктам").to_frame())
if adw.invalid:
    print("Неверные ИИН (первые 20; все: на листе «Проверка ADW» итогового файла):")
    display(pd.DataFrame(adw.invalid[:20]).rename(columns={"row": "Строка", "raw": "ИИН в выгрузке", "err": "Причина", "fio": "ФИО"}))
""")

md("""
## Шаг 3. Кто из клиентов есть в публикациях ДГД
""")

code("""
m = step_match(cfg, adw, dg)
rows = [{"ИИН": iin, "ФИО (ADW)": adw.clients[iin]["fio"], "Регион": x.get("region"), "Публикация": x.get("where"),
         "Требования до": x.get("claimsTo")} for iin, items in m["exact"].items() for x in items]
rows += [{"ИИН": iin, "ФИО (ADW)": adw.clients[iin]["fio"], "Регион": h["region"], "Публикация": "⚠ возможно, по ФИО: " + h["where"]}
         for iin, hits in m["maybe"].items() for h in hits]
display(pd.DataFrame(rows) if rows else "Совпадений с публикациями ДГД нет.")
""")

md("""
## Шаг 4. tazalau.qoldau.kz: снимки и статусы

Сайт допускает около одного запроса в секунду с адреса; программа берёт половину, чтобы не мешать коллегам: примерно 12–15 секунд на ИИН (шесть реестров). Очередь: сначала клиенты из ДГД, затем остальные, давно не проверявшиеся первыми. Прервать можно кнопкой ■ (Interrupt): всё проверенное уже записано в журнал, следующий запуск продолжит.
""")

code("""
results, journal = step_tazalau(cfg, adw, m)
found = [r for r in results.values() if r.get("found")]
print(f"Проверено сейчас: {len(results)}, найдено в реестрах: {len(found)}, ошибок: {sum(1 for r in results.values() if r.get('error'))}")
if found:
    display(pd.DataFrame([{"ИИН": r["iin"], "ФИО": adw.clients[r["iin"]]["fio"], "Статус": r["status"]} for r in found]))
""")

md("""
## Шаг 5. Итоговый Excel
""")

code("""
report = step_report(cfg, adw, dg, m, results, journal)
from IPython.display import FileLink
display(FileLink(os.path.relpath(report, Path.cwd())))
print("Итог:", report)
print("Свод ДГД:", cfg.work_dir / "ДГД_свод.xlsx")
print("Снимки:", cfg.work_dir / "Снимки_tazalau")
""")

md("""
## Запуск каждый день без ноутбука

Те же шаги из командной строки (`run_daily`), журнал работы: в `tazalau_daily.log` рабочей папки:

```
python -m tazalau_daily --work "D:\\Tazalau\\работа"
python -m tazalau_daily --work "D:\\Tazalau\\работа" --scope dgd --minutes 120      # tazalau только для клиентов из ДГД, до 2 часов
python -m tazalau_daily --work "D:\\Tazalau\\работа" --adw-sql "D:\\Tazalau\\adw_problem.sql"   # ADW запросом (строка подключения в ADW_DSN)
```

Планировщик заданий Windows (прав администратора не нужно), каждый будний день в 8:00: команда выполняется один раз в окне «Командная строка» из папки проекта:

```
schtasks /Create /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 08:00 /TN "Tazalau ежедневно" /TR "cmd /c cd /d C:\\путь\\к\\tazalau_bankrotstvo && python -m tazalau_daily --work D:\\Tazalau\\работа"
```

Или выполнить этот ноутбук целиком: `jupyter nbconvert --to notebook --execute notebooks/Tazalau_ежедневная_проверка.ipynb --output-dir D:\\Tazalau\\работа` (копия с результатами ляжет в рабочую папку).
""")

META = {"kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"}, "language_info": {"name": "python"}}
nb = new_notebook(cells=cells, metadata=META)
OUT.parent.mkdir(exist_ok=True)
nbformat.write(nb, OUT)
print("записан", OUT)


def payload() -> str:
    """Исходный код проекта одной строкой: JSON {путь: текст} → zlib → base64."""
    src = {rel: (ROOT / rel).read_text(encoding="utf-8") for rel in EMBED}
    return base64.b64encode(zlib.compress(json.dumps(src, ensure_ascii=False).encode("utf-8"), 9)).decode("ascii")


LOADER = """# Код проекта (пакеты tazalau и tazalau_daily) внутри ноутбука: распаковывается во временную папку компьютера.
# Ячейку можно не открывать. Собрана командой python tools/build_notebook.py из папки проекта.
import base64, hashlib, json, sys, tempfile, zlib
from pathlib import Path

PAYLOAD = "%s"
_src = json.loads(zlib.decompress(base64.b64decode(PAYLOAD)).decode("utf-8"))
CODE_DIR = Path(tempfile.gettempdir()) / ("tazalau_daily_" + hashlib.sha256(PAYLOAD.encode()).hexdigest()[:12])
for _rel, _text in _src.items():
    _f = CODE_DIR / _rel
    _f.parent.mkdir(parents=True, exist_ok=True)
    if not _f.exists() or _f.read_text(encoding="utf-8") != _text:
        _f.write_text(_text, encoding="utf-8")
for _name in [m for m in sys.modules if m in ("tazalau", "tazalau_daily") or m.startswith(("tazalau.", "tazalau_daily."))]:
    del sys.modules[_name]                     # при повторном запуске: код именно из этого ноутбука
if str(CODE_DIR) not in sys.path:
    sys.path.insert(0, str(CODE_DIR))
print("Код проекта:", CODE_DIR, "·", len(_src), "файлов")"""

one = copy.deepcopy(nb)
intro = one.cells[0]
intro.source = intro.source.replace(
    "Ноутбук делает весь день сам:",
    "**Вариант одним файлом:** код проекта уже внутри (скрытая ячейка ниже), папка проекта не нужна: достаточно этого файла, "
    "Python с пакетами `openpyxl pillow playwright pandas pypdf xlrd` и браузера Chrome или Edge.\n\nНоутбук делает весь день сам:")
loader = new_code_cell(LOADER % payload())
loader.metadata = {"jupyter": {"source_hidden": True}, "tags": ["код-проекта"]}
packages = one.cells[1]
packages.source = packages.source.replace(
    "# Пакеты проекта: ноутбук лежит в папке notebooks/, пакеты: на уровень выше\n"
    "import sys\nfrom pathlib import Path\n\n"
    "ROOT = Path.cwd().parent if Path.cwd().name == \"notebooks\" else Path.cwd()\n"
    "if str(ROOT) not in sys.path:\n    sys.path.insert(0, str(ROOT))\n\n", "# Пакеты для работы ноутбука\n")
packages.source = packages.source.replace('print("tazalau_daily", VERSION, "· папка проекта:", ROOT)', 'print("tazalau_daily", VERSION)')
assert "ROOT" not in packages.source, packages.source
one.cells.insert(1, loader)
last = one.cells[-1]
last.source = last.source.replace("notebooks/Tazalau_ежедневная_проверка.ipynb", "Tazalau_ежедневная_проверка_один_файл.ipynb")
last.source = last.source.split("Те же шаги из командной строки")[0] + (
    "Этот файл выполняется целиком командой (например, из Планировщика заданий Windows каждый будний день):\n\n"
    "```\njupyter nbconvert --to notebook --execute Tazalau_ежедневная_проверка_один_файл.ipynb --output-dir D:\\Tazalau\\работа\n```\n\n"
    "Копия с результатами ляжет в рабочую папку, сам файл останется пустым. Запуск `python -m tazalau_daily`: из папки проекта (README).")
nbformat.write(one, OUT_ONE)
print("записан", OUT_ONE, f"({OUT_ONE.stat().st_size // 1024} КБ, код {len(EMBED)} файлов)")
