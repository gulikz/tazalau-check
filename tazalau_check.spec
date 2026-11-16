# -*- mode: python ; coding: utf-8 -*-
"""Сборка Tazalau Check в папку с запускаемым файлом:

    pyinstaller --noconfirm --clean tazalau_check.spec

Результат: dist/TazalauCheck/TazalauCheck(.exe). Папку целиком можно передавать коллегам.
Браузер в сборку не входит — используется Microsoft Edge или Google Chrome, уже установленный
на компьютере (поэтому сборка небольшая и ничего не скачивает из интернета).
"""
from PyInstaller.utils.hooks import collect_data_files

datas = [
    ("tazalau/assets/DejaVuSans.ttf", "tazalau/assets"),
    ("tazalau/assets/DejaVuSans-Bold.ttf", "tazalau/assets"),
    ("tazalau/assets/LICENSE_DEJAVU.txt", "tazalau/assets"),
]
# Драйвер Playwright (node + сценарии управления браузером) — без него браузер не запустить.
datas += collect_data_files("playwright")

a = Analysis(
    ["tazalau_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=["playwright.sync_api", "PIL._tkinter_finder"],
    hookspath=[],
    runtime_hooks=[],
    # Тяжёлые библиотеки, которые могут оказаться в окружении разработчика, программе не нужны.
    excludes=["numpy", "pandas", "matplotlib", "scipy", "IPython", "pytest", "notebook",
              "jupyter", "sphinx", "PyQt5", "PySide6", "PyQt6", "PySide2"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="TazalauCheck",
    debug=False,
    strip=False,
    upx=False,                 # сжатие UPX часто вызывает ложные срабатывания антивирусов
    console=False,             # оконная программа, без чёрного окна консоли
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="TazalauCheck")
