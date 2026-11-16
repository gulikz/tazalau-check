"""Открытые реестры tazalau.qoldau.kz (раздел «Реестры»).

Все шесть списков доступны без входа и фильтруются по ИИН параметром в адресе страницы
(проверено 15.11.2026). Если сайт изменит адрес или имя параметра — правится только эта таблица.
"""
from __future__ import annotations

from dataclasses import dataclass

BASE_URL = "https://tazalau.qoldau.kz"
SITE_HOST = "tazalau.qoldau.kz"

GROUP_DEBTORS = "Должники"
GROUP_BANKRUPTS = "Банкроты и неплатежеспособные"


@dataclass(frozen=True)
class Registry:
    key: str        # короткий код для настроек и журнала
    group: str      # пункт меню «Реестры»
    title: str      # вкладка на странице
    short: str      # короткое имя для столбца результата и галочек в окне
    path: str       # адрес страницы
    param: str      # имя поля фильтра «ИИН»

    @property
    def label(self) -> str:
        return f"{self.group} → {self.title}"

    def url(self, iin: str, base_url: str = BASE_URL) -> str:
        return f"{base_url.rstrip('/')}{self.path}?{self.param}={iin}"


REGISTRIES: tuple[Registry, ...] = (
    Registry("debtor_extrajudicial", GROUP_DEBTORS, "Внесудебная процедура",
             "Должники: внесудебная процедура",
             "/ru/list/debtor", "flApplicantIin"),
    Registry("debtor_judicial", GROUP_DEBTORS, "Судебная процедура",
             "Должники: судебная процедура",
             "/ru/list/debtor/judicial", "flApplicantXin"),
    Registry("debtor_terminated", GROUP_DEBTORS, "Процедура прекращена (объявления)",
             "Должники: процедура прекращена",
             "/ru/list/debtor/extrajudicial/terminated-and-cancelled", "flApplicantIin"),
    Registry("bankrupt_extrajudicial", GROUP_BANKRUPTS, "Внесудебное банкротство",
             "Внесудебное банкротство",
             "/ru/list/bankruptcy-and-insolvent", "flApplicantIin"),
    Registry("bankrupt_judicial", GROUP_BANKRUPTS, "Судебное банкротство",
             "Судебное банкротство",
             "/ru/list/bankruptcy/judicial", "flApplicantXin"),
    Registry("recovery", GROUP_BANKRUPTS, "Восстановление платежеспособности",
             "Восстановление платежеспособности",
             "/ru/list/bankruptcy/recovery", "flApplicantXin"),
)

BY_KEY: dict[str, Registry] = {r.key: r for r in REGISTRIES}

# То, что коллеги делают вручную сейчас (см. фото): «Восстановление платежеспособности».
DEFAULT_KEYS: tuple[str, ...] = ("recovery",)


def resolve(keys) -> list[Registry]:
    """Реестры по кодам, в порядке меню сайта. Неизвестный код — ошибка, а не тихий пропуск."""
    wanted = list(keys)
    unknown = [k for k in wanted if k not in BY_KEY]
    if unknown:
        raise ValueError(f"Неизвестный реестр: {', '.join(unknown)}")
    return [r for r in REGISTRIES if r.key in wanted]
