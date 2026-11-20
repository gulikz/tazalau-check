"""ИИН со сканов (PDF без текста, фото): Tesseract, если он установлен (на рабочих ПК его обычно нет: тогда скан
помечается «проверить вручную»). Результаты кладутся в «ДГД/распознано.json» в том же виде, что у закладки-помощника:
что прочитал один инструмент, второй не читает заново."""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from .common import iins_in_text

PAGES = 5


def available() -> bool:
    return shutil.which("tesseract") is not None


def _tesseract(image: Path) -> str:
    r = subprocess.run(["tesseract", str(image), "stdout", "--psm", "6"], capture_output=True, timeout=120)
    return r.stdout.decode("utf-8", "replace")


def read_scan(path: Path) -> tuple[list[str], str]:
    """(ИИН, примечание). ИИН берутся только с верной контрольной цифрой и настоящей датой рождения."""
    if not available():
        return [], "скан: ИИН не прочитан (распознавание не установлено), нужна проверка вручную"
    ext = path.suffix.lower()
    text = ""
    try:
        if ext == ".pdf":
            if not shutil.which("pdftoppm"):
                return [], "скан PDF: нет pdftoppm для распознавания, нужна проверка вручную"
            with tempfile.TemporaryDirectory() as tmp:
                subprocess.run(["pdftoppm", "-r", "300", "-l", str(PAGES), "-png", str(path), str(Path(tmp) / "p")],
                               capture_output=True, timeout=300, check=True)
                for image in sorted(Path(tmp).glob("p*.png")):
                    text += _tesseract(image) + "\n"
        else:
            text = _tesseract(path)
    except (subprocess.SubprocessError, OSError) as exc:
        return [], f"скан не распознан ({exc})"
    found = iins_in_text(text)
    return found, "" if found else "скан: ИИН не распознан, нужна проверка вручную"
