"""Итоговое изображение: снимок страницы + полоса-подпись над ним.

Подпись (ИИН, реестр, сколько записей, дата и время проверки, адрес страницы) добавляется
отдельной полосой СВЕРХУ снимка — сама страница сайта не изменяется и не перекрывается.
В обычном снимке экрана адресной строки и времени нет, а для подтверждения проверки они нужны.
Если выбрано несколько реестров, их снимки склеиваются в один файл друг под другом.
"""
from __future__ import annotations

import io
import os
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from . import APP_NAME, __version__
from .capture import PageShot
from .config import asset_path
from .fsutil import remove_quietly, replace_with_retry

_STRIP_BG = (36, 52, 71)
_TEXT_MAIN = (255, 255, 255)
_TEXT_SUB = (203, 213, 225)
_PAGE_BG = (238, 244, 248)       # фон страниц сайта — им заполняются поля при разной ширине

_SYSTEM_FONTS = {
    False: ("C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf", "/Library/Fonts/Arial Unicode.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    True: ("C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/arialbd.ttf",
           "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
           "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
}


def _font(size: int, bold: bool = False):
    """Шрифт с кириллицей: вложенный DejaVu, затем системные. None — кириллицы нет вовсе."""
    bundled = asset_path("DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf")
    for candidate in (str(bundled), *_SYSTEM_FONTS[bold]):
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    return None


def stamp_lines(shot: PageShot) -> tuple[str, str]:
    offset = shot.checked_at.strftime("%z")                      # +0500
    zone = f"UTC{offset[:3]}:{offset[3:]}" if offset else "местное время"
    note = ""
    if shot.rows_on_page < shot.count:
        note = f" (на снимке первые {shot.rows_on_page})"
    line1 = f"ИИН {shot.iin}   ·   {shot.registry.label}   ·   Всего записей: {shot.count}{note}"
    line2 = (f"Проверено: {shot.checked_at.strftime('%d.%m.%Y %H:%M:%S')} ({zone})   ·   "
             f"{shot.url}   ·   {APP_NAME} {__version__}")
    return line1, line2


def _ascii_lines(shot: PageShot) -> tuple[str, str]:
    """Запасной вариант подписи, если в системе нет ни одного шрифта с кириллицей."""
    line1 = f"IIN {shot.iin}   |   {shot.registry.key}   |   records: {shot.count}"
    line2 = f"checked: {shot.checked_at.isoformat(timespec='seconds')}   |   {shot.url}"
    return line1, line2


_SEPARATOR = "   ·   "
_MIN_FONT_PX = 12                 # мельче текст подписи уже плохо читается — тогда переносим строки


def _fit(draw: ImageDraw.ImageDraw, text: str, size: int, bold: bool, max_width: int):
    """Шрифт, при котором строка помещается по ширине, но не мельче читаемого."""
    floor = min(size, _MIN_FONT_PX)
    while True:
        font = _font(size, bold)
        if font is None:
            return None
        if draw.textlength(text, font=font) <= max_width or size <= floor:
            return font
        size -= 1


def _wrap(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    """Перенос строки подписи: сначала по разделителям, слишком длинный кусок (адрес) — по символам.
    Ничего не обрезается: на узком снимке подпись просто занимает больше строк."""
    if draw.textlength(text, font=font) <= max_width:
        return [text]
    lines: list[str] = []
    current = ""
    for segment in text.split(_SEPARATOR):
        candidate = segment if not current else current + _SEPARATOR + segment
        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        while len(segment) > 1 and draw.textlength(segment, font=font) > max_width:
            cut = len(segment) - 1
            while cut > 1 and draw.textlength(segment[:cut], font=font) > max_width:
                cut -= 1
            lines.append(segment[:cut])
            segment = segment[cut:]
        current = segment
    if current:
        lines.append(current)
    return lines


def _strip(width: int, shot: PageShot, scale: float) -> Image.Image:
    pad = round(14 * scale)
    gap = round(7 * scale)
    inner = max(40, width - 2 * pad)
    probe = ImageDraw.Draw(Image.new("RGB", (8, 8)))
    line1, line2 = stamp_lines(shot)
    font1 = _fit(probe, line1, round(20 * scale), True, inner)
    font2 = _fit(probe, line2, round(15 * scale), False, inner)
    if font1 is None or font2 is None:
        line1, line2 = _ascii_lines(shot)
        font1 = font2 = ImageFont.load_default()
    rows = [(text, font1, _TEXT_MAIN) for text in _wrap(probe, line1, font1, inner)]
    rows += [(text, font2, _TEXT_SUB) for text in _wrap(probe, line2, font2, inner)]
    heights = [_line_height(probe, font) for _, font, _ in rows]
    strip = Image.new("RGB", (width, 2 * pad + sum(heights) + gap * (len(rows) - 1)), _STRIP_BG)
    draw = ImageDraw.Draw(strip)
    y = pad
    for (text, font, color), height in zip(rows, heights):
        draw.text((pad, y), text, font=font, fill=color)
        y += height + gap
    return strip


def _line_height(draw: ImageDraw.ImageDraw, font) -> int:
    top, bottom = draw.textbbox((0, 0), "ЙИу|", font=font)[1::2]
    return bottom - min(0, top)


def build_image(shots: list[PageShot], *, stamp: bool = True, scale: float = 1.0) -> Image.Image:
    """Один итоговый кадр: для каждого реестра — полоса-подпись и снимок страницы под ней."""
    if not shots:
        raise ValueError("Нет снимков для сохранения.")
    pages = [Image.open(io.BytesIO(shot.png)).convert("RGB") for shot in shots]
    width = max(page.width for page in pages)
    parts: list[Image.Image] = []
    for shot, page in zip(shots, pages):
        if stamp:
            parts.append(_strip(width, shot, scale))
        parts.append(page)
    canvas = Image.new("RGB", (width, sum(part.height for part in parts)), _PAGE_BG)
    y = 0
    for part in parts:
        canvas.paste(part, (0, y))
        y += part.height
    return canvas


def save_png(image: Image.Image, path: Path) -> None:
    """Запись через временный файл: в папке не остаётся недописанных картинок."""
    path = Path(path)
    temp = path.with_name(f"~{path.name}.part")
    try:
        image.save(temp, format="PNG")
        replace_with_retry(temp, path)
    finally:
        remove_quietly(temp)


def open_in_system(path: Path) -> None:
    """Открыть файл или папку программой по умолчанию (для кнопок окна)."""
    import subprocess

    if sys.platform.startswith("win"):
        os.startfile(str(path))                       # noqa: S606 — путь выбран пользователем
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
