"""Собирает переносимый комплект Tazalau Check для Windows.

    python tools/make_windows_package.py                   # основной комплект
    python tools/make_windows_package.py --browser-addon   # плюс отдельное дополнение с браузером

Запускать можно на любой системе (macOS, Linux, Windows): ничего под Windows здесь не
выполняется, комплект только складывается из готовых частей:

  * Python для Windows (сборка python-build-standalone, та же, что ставит менеджер uv) —
    скачивается по закреплённой ссылке и сверяется с закреплённой контрольной суммой;
  * библиотеки для Windows закреплённых версий — pip ставит их в комплект «для чужой системы»;
  * сама программа, файлы запуска, инструкция и образец.

Результат: dist_windows/TazalauCheck_Windows.zip. На рабочем компьютере его достаточно
распаковать и запустить TazalauCheck.bat — устанавливать ничего не нужно, прав администратора
не требуется.

Дополнение (TazalauCheck_Windows_browser_addon.zip) нужно только там, где программам запрещено
управлять установленным Edge/Chrome: оно распаковывается в ту же папку и добавляет в комплект
собственный браузер Chromium, на который политики организации не распространяются.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import os
import shutil
import struct
import subprocess
import sys
import tarfile
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CACHE = ROOT / ".cache_windows"
OUT = ROOT / "dist_windows"
BUNDLE = OUT / "TazalauCheck"
ZIP_PATH = OUT / "TazalauCheck_Windows.zip"
ADDON = OUT / "browser_addon"
ADDON_ZIP = OUT / "TazalauCheck_Windows_browser_addon.zip"

# Чтобы обновить Python: взять новую ссылку и сумму из файла SHA256SUMS нужного выпуска
# https://github.com/astral-sh/python-build-standalone/releases и поменять три строки ниже.
PYTHON_URL = ("https://github.com/astral-sh/python-build-standalone/releases/download/20261110/"
              "cpython-3.13.18%2B20261110-x86_64-pc-windows-msvc-install_only_stripped.tar.gz")
PYTHON_SHA256 = "2e0fed09d48989f64c9305da9ff2d37a41c0d13ab2aa6283405b2b620dfbf805"
PYTHON_TAG = "3.13"

# Версии, на которых программа проверена. Менять только вместе с прогоном тестов.
PACKAGES = ["playwright==1.65.2", "openpyxl==3.1.5", "pillow==13.0.0"]

PRUNE = ["include", "libs", "Scripts", "Lib/test", "Lib/idlelib", "Lib/turtledemo",
         "Lib/ensurepip", "Lib/site-packages/pip", "Lib/site-packages/setuptools"]

LAUNCHER = """@echo off
rem Tazalau Check - start the program.
start "" "%~dp0python\\pythonw.exe" -E -s "%~dp0app\\tazalau_app.py" %*
"""

SELFTEST = """@echo off
rem Tazalau Check - one-minute check that this computer can run the program:
rem starts the browser and opens the site once. The result is printed in this window.
"%~dp0python\\python.exe" -E -s "%~dp0app\\tazalau_app.py" --selftest
echo.
pause
"""

README_TXT = """Tazalau Check {version} — проверка ИИН по реестрам tazalau.qoldau.kz

КАК НАЧАТЬ
  1. Скопируйте папку TazalauCheck целиком на свой компьютер (например, на Рабочий стол).
  2. Один раз запустите «Проверка связи.bat». Через минуту в окне появится строка ИТОГ:
     «всё в порядке» — можно работать; иначе покажите это окно разработчику.
  3. Для работы запускайте TazalauCheck.bat.
  4. Как пользоваться — в файле ИНСТРУКЦИЯ.html (откроется в браузере).

ЕСЛИ WINDOWS СПРАШИВАЕТ, МОЖНО ЛИ ЗАПУСКАТЬ ФАЙЛ
  Это обычное предупреждение для файлов, полученных с другого компьютера. Чтобы оно не
  появлялось: ДО распаковки щёлкните архив правой кнопкой — «Свойства» — поставьте галочку
  «Разблокировать» — «ОК», и только потом распаковывайте.

ЧТО ВНУТРИ
  TazalauCheck.bat       запуск программы
  Проверка связи.bat     проверка, что на этом компьютере программа сможет работать
  ИНСТРУКЦИЯ.html        как пользоваться
  пример.xlsx            образец файла Excel
  app\\                   программа
  python\\                Python {python} для Windows — используется только этой программой,
                         в систему ничего не устанавливается

Программе нужен Microsoft Edge или Google Chrome, установленный на компьютере.
ЭЦП и вход на сайт не требуются.

ЕСЛИ «ПРОВЕРКА СВЯЗИ» ПИШЕТ, ЧТО БРАУЗЕР НЕ УДАЛОСЬ ЗАПУСТИТЬ
  На некоторых рабочих компьютерах программам запрещено управлять Edge и Chrome. Тогда
  попросите у разработчика дополнение TazalauCheck_Windows_browser_addon.zip и распакуйте
  его в ту же папку, где лежит папка TazalauCheck: внутри появится папка browsers со
  встроенным браузером. После этого снова запустите «Проверка связи.bat».

Собрано: {built}
"""

NOTICES = r"""Tazalau Check - third-party components / сторонние компоненты

The package bundles the following software unchanged. Each component stays under its own
licence; the licence texts are inside the package at the paths shown.
В комплект без изменений входят перечисленные ниже программы. Каждая распространяется на
условиях своей лицензии; тексты лицензий лежат внутри комплекта по указанным путям.

  CPython {python} for Windows (python-build-standalone build), Tcl/Tk
      python\LICENSE.txt
  Playwright for Python, with the Playwright driver and Node.js
      python\Lib\site-packages\playwright-*.dist-info\licenses\LICENSE
      python\Lib\site-packages\playwright\driver\LICENSE
      python\Lib\site-packages\playwright\driver\package\ThirdPartyNotices.txt
  openpyxl, et_xmlfile, Pillow, greenlet, pyee, typing_extensions
      python\Lib\site-packages\<name>-*.dist-info\
  DejaVu fonts
      app\tazalau\assets\LICENSE_DEJAVU.txt
  Chromium headless shell and FFmpeg (only in the optional browser add-on)
      browsers\chromium_headless_shell-*\chrome-headless-shell-win64\LICENSE.headless_shell
      browsers\ffmpeg-*\COPYING.LGPLv2.1
"""

HTML_PAGE = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<title>Tazalau Check — как пользоваться</title>
<style>
 body {{ font: 16px/1.55 "Segoe UI", Arial, sans-serif; color: #1f2933; max-width: 820px;
        margin: 32px auto; padding: 0 20px; }}
 h1 {{ font-size: 28px; }} h2 {{ font-size: 21px; margin-top: 34px; border-bottom: 1px solid #d9e2ec;
        padding-bottom: 6px; }}
 code {{ background: #eef2f6; padding: 1px 5px; border-radius: 4px; font-family: Consolas, monospace; }}
 table {{ border-collapse: collapse; width: 100%; margin: 14px 0; }}
 th, td {{ border: 1px solid #cbd2d9; padding: 8px 10px; text-align: left; vertical-align: top; }}
 th {{ background: #f0f4f8; }} hr {{ border: 0; border-top: 1px solid #d9e2ec; margin: 26px 0; }}
 li {{ margin: 4px 0; }}
</style></head><body>
{body}
</body></html>
"""


def log(text: str) -> None:
    print(text, flush=True)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fetch_python() -> Path:
    """Скачивает (или берёт из кэша) архив Python для Windows и сверяет контрольную сумму."""
    CACHE.mkdir(exist_ok=True)
    archive = CACHE / PYTHON_URL.rsplit("/", 1)[1].replace("%2B", "+")
    if not archive.exists() or sha256(archive) != PYTHON_SHA256:
        log("Скачиваю Python для Windows…")
        temp = archive.with_suffix(".part")
        with urllib.request.urlopen(PYTHON_URL, timeout=120) as response, open(temp, "wb") as out:
            shutil.copyfileobj(response, out)
        os.replace(temp, archive)
    actual = sha256(archive)
    if actual != PYTHON_SHA256:
        archive.unlink()
        raise SystemExit(f"Контрольная сумма Python не совпала: ожидалась {PYTHON_SHA256}, "
                         f"получена {actual}. Архив удалён, сборка остановлена.")
    return archive


def extract_python(archive: Path) -> None:
    log("Распаковываю Python…")
    with tarfile.open(archive, "r:gz") as tar:
        base = BUNDLE.resolve()
        for member in tar.getmembers():          # архив не должен писать за пределы комплекта
            target = (BUNDLE / member.name).resolve()
            if base not in target.parents and target != base:
                raise SystemExit(f"Подозрительный путь в архиве Python: {member.name}")
            if not (member.isfile() or member.isdir()):
                raise SystemExit(f"Неожиданный элемент в архиве Python: {member.name}")
        try:
            tar.extractall(BUNDLE, filter="data")
        except TypeError:                        # старый Python без параметра filter
            tar.extractall(BUNDLE)
    python = BUNDLE / "python"
    for name in PRUNE:                           # средства разработки программе не нужны
        shutil.rmtree(python / name, ignore_errors=True)
    for junk in list(python.glob("DLLs/_test*.pyd")) + list(python.glob("DLLs/xxlimited*.pyd")) \
            + list(python.glob("tcl/*.lib")) + list(python.glob("tcl/*.sh")):
        junk.unlink()
    for leftover in list(python.glob("Lib/site-packages/pip-*.dist-info")) \
            + list(python.glob("Lib/site-packages/setuptools-*.dist-info")):
        shutil.rmtree(leftover, ignore_errors=True)
    for cache in list(python.rglob("__pycache__")):
        shutil.rmtree(cache, ignore_errors=True)


def install_packages() -> None:
    """Ставит библиотеки для Windows внутрь комплекта (pip умеет ставить «для чужой системы»)."""
    log("Ставлю библиотеки для Windows: " + ", ".join(PACKAGES))
    site = BUNDLE / "python" / "Lib" / "site-packages"
    site.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
         "--no-compile", "--target", str(site), "--platform", "win_amd64",
         "--python-version", PYTHON_TAG, "--implementation", "cp", "--only-binary=:all:",
         "--no-warn-conflicts", *PACKAGES],
        check=True)
    shutil.rmtree(site / "bin", ignore_errors=True)       # запускалки для системы сборщика


def copy_app() -> str:
    log("Копирую программу…")
    app = BUNDLE / "app"
    shutil.copytree(ROOT / "tazalau", app / "tazalau",
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    shutil.copy2(ROOT / "tazalau_app.py", app / "tazalau_app.py")
    shutil.copy2(ROOT / "tazalau_config.example.json", app / "tazalau_config.example.json")
    shutil.copy2(ROOT / "пример.xlsx", BUNDLE / "пример.xlsx")
    namespace: dict = {}
    exec((ROOT / "tazalau" / "__init__.py").read_text(encoding="utf-8"), namespace)
    return namespace["__version__"]


def write_bat(path: Path, text: str) -> None:
    """Файлы .bat — только ASCII и окончания строк CRLF, иначе cmd.exe может прочитать их неверно."""
    if not text.isascii():
        raise SystemExit(f"В {path.name} есть не-ASCII символы")
    path.write_bytes(text.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))


def write_docs(version: str) -> None:
    write_bat(BUNDLE / "TazalauCheck.bat", LAUNCHER)
    write_bat(BUNDLE / "Проверка связи.bat", SELFTEST)
    source = (ROOT / "ИНСТРУКЦИЯ.md").read_text(encoding="utf-8")
    try:
        import markdown
        body = markdown.markdown(source, extensions=["tables"])
    except ImportError:                          # без библиотеки markdown — как есть, но читаемо
        import html
        body = "<pre style='white-space: pre-wrap'>" + html.escape(source) + "</pre>"
    (BUNDLE / "ИНСТРУКЦИЯ.html").write_text(HTML_PAGE.format(body=body), encoding="utf-8")
    readme = README_TXT.format(version=version, python=PYTHON_TAG,
                               built=datetime.now().strftime("%d.%m.%Y %H:%M"))
    # UTF-8 с пометкой BOM и CRLF: «Блокнот» любой версии Windows покажет кириллицу правильно.
    (BUNDLE / "ПРОЧТИ.txt").write_bytes(("\ufeff" + readme.replace("\n", "\r\n")).encode("utf-8"))
    notices = NOTICES.format(python=PYTHON_TAG)
    (BUNDLE / "THIRD_PARTY_NOTICES.txt").write_bytes(
        ("\ufeff" + notices.replace("\n", "\r\n")).encode("utf-8"))


def make_browser_addon() -> None:
    """Дополнение со встроенным Chromium для Windows — на случай запрета управлять Edge/Chrome."""
    wanted = next(p.split("==")[1] for p in PACKAGES if p.startswith("playwright=="))
    try:
        have = importlib.metadata.version("playwright")
    except importlib.metadata.PackageNotFoundError:
        have = "не установлен"
    if have != wanted:
        raise SystemExit(f"Для загрузки браузера нужен playwright=={wanted} (сейчас: {have}). "
                         "Версия браузера жёстко привязана к версии Playwright в комплекте.")
    log("Скачиваю встроенный Chromium для Windows (около 120 МБ)…")
    shutil.rmtree(ADDON, ignore_errors=True)
    browsers = ADDON / "TazalauCheck" / "browsers"
    browsers.mkdir(parents=True)
    env = dict(os.environ, PLAYWRIGHT_BROWSERS_PATH=str(browsers),
               PLAYWRIGHT_HOST_PLATFORM_OVERRIDE="win64")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium-headless-shell"],
                   check=True, env=env, stdout=subprocess.DEVNULL)
    shutil.rmtree(browsers / ".links", ignore_errors=True)   # служебные пути компьютера сборщика
    shells = list(browsers.glob("chromium_headless_shell-*/*/chrome-headless-shell.exe"))
    if len(shells) != 1 or pe_machine(shells[0]) != 0x8664:
        raise SystemExit("Встроенный браузер для Windows не скачался.")
    ADDON_ZIP.unlink(missing_ok=True)
    with zipfile.ZipFile(ADDON_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(ADDON.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(ADDON))
    log(f"Дополнение: {ADDON_ZIP} ({ADDON_ZIP.stat().st_size / 2 ** 20:.0f} МБ)")
    log(f"  sha256: {sha256(ADDON_ZIP)}")


def pe_machine(path: Path) -> int:
    """Тип процессора из заголовка .exe/.dll/.pyd (0x8664 — 64-разрядный Windows)."""
    with open(path, "rb") as handle:
        header = handle.read(0x40)
        if header[:2] != b"MZ":
            return 0
        handle.seek(struct.unpack_from("<I", header, 0x3C)[0])
        signature = handle.read(6)
    return struct.unpack_from("<H", signature, 4)[0] if signature[:4] == b"PE\0\0" else 0


def verify() -> None:
    """Проверка комплекта «на глаз машины»: всё ли на месте и всё ли действительно для Windows x64.
    Запустить программы Windows здесь нельзя — это проверка состава, а не работы."""
    log("Проверяю состав комплекта…")
    python, site = BUNDLE / "python", BUNDLE / "python" / "Lib" / "site-packages"
    required = [
        "TazalauCheck.bat", "Проверка связи.bat", "ИНСТРУКЦИЯ.html", "ПРОЧТИ.txt", "пример.xlsx",
        "THIRD_PARTY_NOTICES.txt", "python/LICENSE.txt",
        "python/Lib/site-packages/playwright/driver/LICENSE",
        "app/tazalau_app.py", "app/tazalau/gui.py", "app/tazalau/assets/DejaVuSans.ttf",
        "python/python.exe", "python/pythonw.exe", "python/python313.dll", "python/vcruntime140.dll",
        "python/DLLs/_tkinter.pyd", "python/DLLs/tcl86t.dll", "python/DLLs/tk86t.dll",
        "python/tcl/tcl8.6/init.tcl", "python/tcl/tk8.6/tk.tcl", "python/Lib/tkinter/__init__.py",
        "python/Lib/site-packages/playwright/driver/node.exe",
        "python/Lib/site-packages/playwright/driver/package/cli.js",
        "python/Lib/site-packages/openpyxl/__init__.py",
        "python/Lib/site-packages/PIL/__init__.py",
        "python/Lib/site-packages/greenlet/__init__.py",
    ]
    missing = [name for name in required if not (BUNDLE / name).exists()]
    if missing:
        raise SystemExit("В комплекте не хватает файлов:\n  " + "\n  ".join(missing))
    binaries = [python / "python.exe", python / "pythonw.exe", python / "python313.dll",
                python / "DLLs" / "_tkinter.pyd", site / "playwright" / "driver" / "node.exe"]
    binaries += list((site / "PIL").glob("_imaging.*.pyd")) + list((site / "greenlet").glob("_greenlet.*.pyd"))
    if len(binaries) < 7:
        raise SystemExit("Не найдены двоичные модули Pillow или greenlet для Windows.")
    wrong = [str(b.relative_to(BUNDLE)) for b in binaries if pe_machine(b) != 0x8664]
    if wrong:
        raise SystemExit("Файлы не для 64-разрядной Windows:\n  " + "\n  ".join(wrong))
    tag = f"cp{PYTHON_TAG.replace('.', '')}-win_amd64"
    for module in list((site / "PIL").glob("*.pyd")) + list((site / "greenlet").glob("*.pyd")):
        if tag not in module.name:
            raise SystemExit(f"Модуль собран не под этот Python: {module.name} (нужен {tag})")
    foreign = [str(p.relative_to(BUNDLE)) for p in site.rglob("*") if p.suffix in (".so", ".dylib")]
    if foreign:
        raise SystemExit("В комплект попали библиотеки для другой системы:\n  " + "\n  ".join(foreign))
    for bat in BUNDLE.glob("*.bat"):
        data = bat.read_bytes()
        if not data.isascii() or b"\r\n" not in data or data.replace(b"\r\n", b"").count(b"\n"):
            raise SystemExit(f"{bat.name}: нужны только ASCII и окончания строк CRLF")
    for source in (BUNDLE / "app").rglob("*.py"):       # синтаксис программы
        compile(source.read_text(encoding="utf-8"), str(source), "exec")


def make_zip() -> None:
    log("Упаковываю архив…")
    ZIP_PATH.unlink(missing_ok=True)
    # zipfile помечает имена с кириллицей как UTF-8 — Проводник Windows 10/11 читает их правильно.
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(BUNDLE.rglob("*")):
            if path.is_file():
                archive.write(path, Path("TazalauCheck") / path.relative_to(BUNDLE))


def main() -> None:
    parser = argparse.ArgumentParser(description="Сборка переносимого комплекта Tazalau Check для Windows.")
    parser.add_argument("--browser-addon", action="store_true",
                        help="дополнительно собрать архив со встроенным Chromium (для компьютеров, "
                             "где управлять Edge/Chrome запрещено политикой организации)")
    args = parser.parse_args()

    archive = fetch_python()
    shutil.rmtree(BUNDLE, ignore_errors=True)
    BUNDLE.mkdir(parents=True)
    extract_python(archive)
    install_packages()
    version = copy_app()
    write_docs(version)
    verify()
    make_zip()

    files = sum(1 for p in BUNDLE.rglob("*") if p.is_file())
    size = sum(p.stat().st_size for p in BUNDLE.rglob("*") if p.is_file()) / 2 ** 20
    log(f"\nГотово: {ZIP_PATH}")
    log(f"  архив {ZIP_PATH.stat().st_size / 2 ** 20:.0f} МБ; в распакованном виде {size:.0f} МБ, файлов: {files}")
    log(f"  sha256 архива: {sha256(ZIP_PATH)}")
    if args.browser_addon:
        make_browser_addon()


if __name__ == "__main__":
    main()
