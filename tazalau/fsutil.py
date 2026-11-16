"""Мелочи файловой системы, которые в Windows ведут себя иначе, чем везде."""
from __future__ import annotations

import os
import time
from pathlib import Path

# В тестах обнуляется, чтобы не ждать.
RETRY_ATTEMPTS = 6
RETRY_PAUSE_S = 0.25


def replace_with_retry(source: Path, target: Path) -> None:
    """``os.replace`` с повторами.

    В Windows только что записанный файл на доли секунды бывает занят антивирусом или
    индексатором поиска — переименование падает с PermissionError, а через мгновение проходит.
    Без повторов долгая проверка на рабочем компьютере могла бы оборваться на случайном ИИН.
    Если файл занят всерьёз (книга открыта в Excel), исключение уходит наверх после всех попыток.
    """
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            os.replace(source, target)
            return
        except PermissionError:
            if attempt == RETRY_ATTEMPTS:
                raise
            time.sleep(RETRY_PAUSE_S * attempt)


def remove_quietly(path: Path) -> None:
    """Удалить служебный файл; если он на мгновение занят — не страшно, это не данные."""
    for attempt in range(1, 4):
        try:
            if path.exists():
                path.unlink()
            return
        except OSError:
            time.sleep(RETRY_PAUSE_S * attempt)
