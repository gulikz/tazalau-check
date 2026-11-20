-- =====================================================================================================
-- ЧЕРНОВИК. Шаг 2 ежедневной проверки: клиенты с признаком «проблема» по кредитам и кредитным картам (ADW, PostgreSQL).
-- Настоящие схемы, таблицы и поля банка здесь НЕ указаны: все места для замены помечены «ЗАМЕНИТЬ».
-- Итоговый запрос даст владелец выгрузки; черновик задаёт, ЧТО должно получиться на выходе, чтобы помощник
-- Tazalau прочитал файл без настройки.
--
-- На выходе (имена столбцов важны, порядок нет):
--   "ИИН/БИН"  12 цифр ТЕКСТОМ, с ведущими нулями      обязательно: по нему идёт проверка
--   "FIO"      ФИО клиента                             желательно: из него имя папки со снимком
--   остальное  продукт, договор, просрочка, сумма       по желанию: помощник их не меняет
--
-- Порядок: выполнить блок 1 (функция и временное представление), затем блок 2 (выгрузка) и сохранить результат:
--   DBeaver: результат -> «Экспорт данных» -> XLSX или CSV;  pgAdmin: «Download as CSV».
--   Файл положить в рабочую папку дня, например ADW_2026-11-20.csv. Блок 3 (проверки) по желанию.
--   Как должен выглядеть результат: sql/ADW_пример_выгрузки.csv (вымышленные люди; один клиент с кредитом и картой,
--   ИИН с ведущим нулём, ИИН с опечаткой, которую помощник покажет как «НЕВЕРНЫЙ ИИН»).
-- Варианты блока 2:
--   А) весь проблемный портфель (по умолчанию): помощник сам отметит ИИН из списков ДГД и проверит их первыми;
--   Б) только ИИН из списков ДГД: в помощнике кнопка «ИИН для SQL ⧉» копирует их списком VALUES (...);
--      вставьте его вместо VALUES ('000000000000') и раскомментируйте строку WHERE.
-- Временные функция и представление живут до конца сеанса и прав на создание объектов в схемах не требуют.
-- =====================================================================================================

-- ---------------------------------------- блок 1 ----------------------------------------------------
-- Проверка ИИН: 12 цифр и контрольная цифра (тот же алгоритм, что в помощнике и на сайте)
CREATE OR REPLACE FUNCTION pg_temp.iin_ok(p text) RETURNS boolean LANGUAGE sql IMMUTABLE AS $$
  WITH d AS (
    SELECT array_agg(substr(p, i, 1)::int ORDER BY i) AS a
    FROM generate_series(1, 12) AS i
    WHERE p ~ '^[0-9]{12}$'
  ), k AS (
    SELECT a,
           (SELECT sum(a[i] * i) FROM generate_series(1, 11) AS i) % 11 AS k1,
           (SELECT sum(a[i] * (ARRAY[3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2])[i]) FROM generate_series(1, 11) AS i) % 11 AS k2
    FROM d
  )
  SELECT COALESCE((SELECT CASE WHEN k1 <> 10 THEN k1 = a[12]
                               WHEN k2 <> 10 THEN k2 = a[12]
                               ELSE false END
                   FROM k WHERE a IS NOT NULL), false)
$$;

CREATE OR REPLACE TEMP VIEW adw_problem AS
WITH params AS (
    SELECT (current_date - 1)::date AS report_date          -- ЗАМЕНИТЬ при необходимости: дата среза витрины (обычно вчера)
),
loans AS (                                                 -- кредиты с признаком «проблема»
    SELECT l.client_iin::text            AS iin_raw,       -- ЗАМЕНИТЬ: поле ИИН клиента
           l.client_fio::text            AS fio,           -- ЗАМЕНИТЬ: ФИО клиента
           'Кредит'::text                AS product,
           l.contract_number::text       AS contract,      -- ЗАМЕНИТЬ: номер договора
           l.days_past_due::int          AS dpd,           -- ЗАМЕНИТЬ: дни просрочки
           l.total_debt::numeric         AS debt,          -- ЗАМЕНИТЬ: общая задолженность, тг
           l.branch_name::text           AS branch         -- ЗАМЕНИТЬ: филиал
    FROM   adw_schema.loan_contracts l                     -- ЗАМЕНИТЬ: витрина кредитов в ADW
    JOIN   params p ON l.report_date = p.report_date       -- ЗАМЕНИТЬ: поле даты среза
    WHERE  l.problem_flag = 1                              -- ЗАМЕНИТЬ: как в ЕДВ обозначен признак «проблема»
),
cards AS (                                                 -- кредитные карты с признаком «проблема»
    SELECT c.client_iin::text, c.client_fio::text, 'Кредитная карта'::text, c.card_contract_number::text,
           c.days_past_due::int, c.total_debt::numeric, c.branch_name::text
    FROM   adw_schema.card_contracts c                     -- ЗАМЕНИТЬ: витрина кредитных карт
    JOIN   params p ON c.report_date = p.report_date
    WHERE  c.problem_flag = 1                              -- ЗАМЕНИТЬ
)
SELECT lpad(regexp_replace(u.iin_raw, '[^0-9]', '', 'g'), 12, '0') AS iin,   -- ИИН числом теряет ведущие нули: возвращаем
       u.fio, u.product, u.contract, u.dpd, u.debt, u.branch,
       (SELECT report_date FROM params) AS report_date
FROM  (SELECT * FROM loans UNION ALL SELECT * FROM cards) AS u;

-- ---------------------------------------- блок 2: выгрузка ------------------------------------------
SELECT iin                    AS "ИИН/БИН",
       fio                    AS "FIO",
       product                AS "Продукт",
       contract               AS "Договор",
       dpd                    AS "Дней просрочки",
       debt                   AS "Задолженность, тг",
       branch                 AS "Филиал",
       CASE WHEN pg_temp.iin_ok(iin) THEN 'да' ELSE 'НЕТ: проверить ИИН' END AS "ИИН верный",
       to_char(report_date, 'DD.MM.YYYY') AS "Дата среза"
FROM   adw_problem
-- WHERE  iin IN (VALUES ('000000000000'))                 -- вариант Б: вставить список из «ИИН для SQL ⧉»
ORDER  BY iin, product, contract;

-- ---------------------------------------- блок 3: проверки (по желанию) -----------------------------
-- 1) Объём и состав (сравнить со вчерашним днём: резкий скачок или ноль = вопрос к витрине)
SELECT product AS "Продукт", count(*) AS "Строк", count(DISTINCT iin) AS "ИИН" FROM adw_problem GROUP BY product ORDER BY product;
-- 2) Неверные ИИН (должно быть 0: по испорченному ИИН сайт ничего не найдёт)
SELECT iin AS "ИИН", fio AS "ФИО", contract AS "Договор" FROM adw_problem WHERE NOT pg_temp.iin_ok(iin);
-- 3) Пустые ФИО (желательно 0: без ФИО папка снимка называется по ИИН)
SELECT count(*) AS "Без ФИО" FROM adw_problem WHERE coalesce(trim(fio), '') = '';
-- 4) Один ИИН с разными ФИО (возможна ошибка привязки клиента)
SELECT iin AS "ИИН", array_agg(DISTINCT fio) AS "ФИО" FROM adw_problem GROUP BY iin HAVING count(DISTINCT fio) > 1;
