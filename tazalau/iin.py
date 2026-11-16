"""ИИН/БИН: приведение значения ячейки Excel к 12 цифрам и проверка контрольной цифры.

Зачем это нужно:
  * Excel хранит ИИН как число и теряет ведущие нули (ИИН родившихся в 2000-х начинается с «0»).
  * Опечатка в ИИН дала бы снимок «Нет записей» по чужому/несуществующему номеру — то есть
    ложное «не найден». Поэтому номер с неверной контрольной цифрой на сайт не отправляется.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_WEIGHTS_1 = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
_WEIGHTS_2 = (3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2)

# Пробелы всех видов, неразрывные пробелы, символы нулевой ширины, BOM.
_JUNK = re.compile(r"[\s\u00a0\u1680\u2000-\u200f\u2028\u2029\u202f\u205f\u2060\u3000\ufeff]+")

# Сколько ведущих нулей Excel мог «съесть» у числа: 000101… (родился 01.01.2000) -> 9 цифр.
_MIN_NUMERIC_DIGITS = 9


def check_digit(first11: str) -> int | None:
    """Контрольная (12-я) цифра по первым 11. None — номер с таким началом не выдаётся."""
    digits = [int(ch) for ch in first11]
    rem = sum(d * w for d, w in zip(digits, _WEIGHTS_1)) % 11
    if rem == 10:
        rem = sum(d * w for d, w in zip(digits, _WEIGHTS_2)) % 11
        if rem == 10:
            return None
    return rem


def checksum_ok(iin: str) -> bool:
    if len(iin) != 12 or not iin.isdigit():
        return False
    expected = check_digit(iin[:11])
    return expected is not None and expected == int(iin[11])


@dataclass(frozen=True)
class IinParse:
    """Результат разбора одной ячейки столбца «ИИН/БИН»."""

    raw: object
    iin: str | None = None      # 12 цифр, если значение пригодно для запроса
    empty: bool = False         # ячейка пустая — строка просто пропускается
    error: str | None = None    # понятная причина, почему значение не годится
    padded: bool = False        # были восстановлены ведущие нули

    @property
    def ok(self) -> bool:
        return self.iin is not None


def normalize_iin(value: object, *, allow_bad_checksum: bool = False) -> IinParse:
    """Приводит значение ячейки к ИИН/БИН из 12 цифр либо объясняет, почему это невозможно."""
    if value is None:
        return IinParse(raw=value, empty=True)
    if isinstance(value, bool):
        return IinParse(raw=value, error="в ячейке не ИИН, а логическое значение")

    numeric = False
    if isinstance(value, int):
        if value < 0:
            return IinParse(raw=value, error="отрицательное число вместо ИИН")
        digits, numeric = str(value), True
    elif isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return IinParse(raw=value, error="в ячейке не число")
        if value < 0 or not value.is_integer():
            return IinParse(raw=value, error="дробное или отрицательное число вместо ИИН")
        digits, numeric = str(int(value)), True
    else:
        text = _JUNK.sub("", str(value))
        text = text.lstrip("'’`")                      # апостроф «текстового» формата
        if not text:
            return IinParse(raw=value, empty=True)
        if re.fullmatch(r"\d+[.,]0+", text):            # «123456789012.0» после выгрузки
            text = re.split(r"[.,]", text)[0]
        if not text.isdigit() or not text.isascii():
            return IinParse(raw=value, error="в ИИН есть символы, кроме цифр")
        digits = text

    padded = False
    if len(digits) != 12:
        # Короткое ЧИСЛО — почти наверняка потерянные ведущие нули; короткий ТЕКСТ — опечатка.
        if numeric and _MIN_NUMERIC_DIGITS <= len(digits) < 12:
            digits, padded = digits.zfill(12), True
        else:
            return IinParse(raw=value, error=f"в ИИН {len(digits)} цифр вместо 12")

    if not checksum_ok(digits):
        if padded:
            return IinParse(raw=value, error="в ИИН меньше 12 цифр и контрольная цифра не сходится")
        if not allow_bad_checksum:
            return IinParse(
                raw=value,
                error="контрольная цифра ИИН не сходится (вероятно, опечатка)",
            )
    return IinParse(raw=value, iin=digits, padded=padded)
