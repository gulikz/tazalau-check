"""Synthetic demo data for the soft-collection segmentation notebook.

``make_demo_tables()`` returns a dict of pandas DataFrames that reproduce the
schema of the real tables (table names, column names incl. the Cyrillic ones,
pandas-3 dtypes, string-typed dates, SCD sentinels, duplicates and other
quirks) so the notebook can be run end-to-end without the bank's data.

Nothing here is real: clients, IINs, names, phones, contracts and amounts are
random.  A latent client risk drives every table, so a model built on proper
as-of features finds a realistic (moderate) signal.

Tables (keys of the returned dict)::

    contr, ver_df_all_2024_2026, bankrupt, ocrm_call, vneshpt_ot, vnpt, excl,
    vneshpt, ovrd, geo, dm, dm_rej, hes_pkb_agg

Usage::

    from demo_data import make_demo_tables
    tables = make_demo_tables(n_contracts=4000, seed=42, cut_date="2026-11-11")
    globals().update(tables)        # same names as in the real notebook

Run ``python demo_data.py`` for a self-test (shapes, dtypes, quirks, rates).
"""
from __future__ import annotations

import datetime as _dt
from types import SimpleNamespace

import numpy as np
import pandas as pd

__all__ = ["make_demo_tables", "TABLE_NAMES", "PKB_RULE_NAMES"]

TABLE_NAMES = (
    "contr", "ver_df_all_2024_2026", "bankrupt", "ocrm_call", "vneshpt_ot", "vnpt", "excl",
    "vneshpt", "ovrd", "geo", "dm", "dm_rej", "hes_pkb_agg",
)

# --------------------------------------------------------------------------------------
# Column lists (real schema)
# --------------------------------------------------------------------------------------
VER_TEXT_COLS = [
    "IIN", "BRANCH", "GENDER", "CITIZENSHIP", "TRUST_PHONE_NUMBER", "FAMILY_STATUS", "CLIENT_SEGMENT",
    "RESIDENCE_COUNTRY", "CLIENT_SEGMENT_RB", "CLIENT_CATEGORY", "EDUCATION", "FLG_DEATH",
    "FLG_LOAN_OVERDUE_MORE_1", "MOBILE_OS", "MOBILE_MODEL", "MOBILE_BRAND", "MAU_FIREBASE",
    "MAU_AUTORIZATION", "LAST_EVENT_COUNTRY", "LAST_EVENT_CITY", "FLG_BLACK_LIST", "TYPE_BLACKLIST",
    "EPG_1_0_OR_NEW_2", "SALARY_CATEGORY_BY_MODEL", "TRANSACTION_CLIENT_SEGMENT", "MCG_TOP1", "MCG_TOP2",
    "MCG_TOP3", "DIGITAL_SEGMENT", "Full_name",
]
VER_DATE_COLS = ["DATE_VALUE", "DOC_EXPIRY_DATE"]
VER_NUM_COLS = (
    [
        "AGE", "CNT_OPEN_DEPOSIT", "CNT_OPEN_DEPOSIT_ONLINE", "CNT_DEPOSIT", "CNT_CLOSE_DEPO",
        "AMOUNT_REPLENISH_DEPO", "AMOUNT_WITHDRAWAL_DEPO", "CNT_LOAN", "CNT_LOAN_ONLINE",
        "AMOUNT_LOAN_MAIN_DEBT", "AMOUNT_LOAN_OVERDUE", "MONTH_LOAN_PAY_AMOUNT", "AMOUNT_OUT_MONEY",
        "AMOUNT_SOC_PAYOUT", "AMOUNT_PENSION", "AMOUNT_CASHIN_ATM", "VOLUME_NONCASH_TRANS",
        "AMOUNT_CASHOUT_ATM", "CNT_INTER_TRANSFER_HB", "AMOUNT_INTER_TRANSFER_HB", "CNT_TRAN_OTHER_BANK",
        "AMOUNT_TRAN_OTHER_BANK", "CNT_TRANSFER_KASPI", "AMOUNT_TRANSFER_KASPI", "CNT_KASPI_CARD_TRANSFER",
        "CNT_INBOX_TRANSF_ANOT_BANK", "VOLUME_INBOX_TRANSF_ANOT_BANK", "CNT_INBOX_TRANSF_KASPI",
        "VOLUME_INBOX_TRANSF_KASPI", "CNT_INBOX_TRANSF_HALYK", "VOLUME_INBOX_TRANSF_HALYK", "CNT_PAYMENT",
        "PAYMENT_AMOUNT", "CNT_BANK_PRODUCTS_DEBIT_CARD", "CNT_MOBILE_PAY_HB", "AMOUNT_MOBILE_PAY_HB",
        "CNT_COMMUNAL_PAY_HB", "AMOUNT_COMMUNAL_PAY_HB", "CNT_PARKING_PAY", "AMOUNT_PARKING_PAY",
        "CNT_FINE_PAY", "AMOUNT_FINE_PAY", "CNT_TRANSPORT_PAY", "AMOUNT_TRANSPORT_PAY", "CNT_EDUCATION_PAY",
        "AMOUNT_EDUCATION_PAY", "CNT_UNIQUE_PAY_SERVICES", "USETIME_USER_HB", "AVG_USETIME_DAY",
        "CNT_SESSIONS", "CNT_SUCCESS_SESSIONS", "CNT_UNIQUE_DAY_IN_MONTH", "HALYK_EMPLOYEE",
        "CNT_LOAN_PKB_WITHOUT_HB", "LOAN_BALANCE_KASPI", "ADDICTION_TO_BZK", "ADDICTION_TO_TOVARKA",
        "WALLET_TO_INSTALLMENT", "WALLET_TO_DEPOSIT", "WALLET_TO_CREDIT", "WALLET_TO_MARKET",
        "CNT_BANK_PRODUCTS",
    ]
    + [f"PKB_REQ_{w}" for w in ("7D", "30D", "90D", "120D", "180D", "360D")]
    + [f"PKB_{k}_CONTRACT_CNT_{c}" for k in ("ACT", "COMPL", "RETURN") for c in ("KZT", "USD")]
    + ["PKB_SUM_UNPAY_ACT", "SUM_PKB_OVRD_ACT", "SUM_PKB_UNPAY_COMPL", "SUM_PKB_OVRD_COMPL",
       "SUM_PKB_UNPAY_RECALL"]
    + ["CRED_CNT", "CRED_CNT_ACTIVE", "CRED_CNT_ACTIVE_KZT", "CRED_CNT_ACTIVE_USD", "CRED_CNT_NON_ACTIVE",
       "CRED_CNT_NON_ACTIVE_KZT", "CRED_CNT_NON_ACTIVE_USD"]
    + [f"CRED_CNT_REPAYMENT_DAY_{i}_BZK" for i in (1, 2, 3)]
    + ["CRED_DOGOVOR_AVG_360D", "CRED_DOGOVOR_AVG_60D", "CRED_DOGOVOR_SUM", "CRED_DOGOVOR_SUM_30D",
       "CRED_DOGOVOR_SUM_60D", "CRED_DOGOVOR_SUM_90D"]
    + ["CRED_GIVEN_AVG_180D", "CRED_GIVEN_AVG_360D", "CRED_GIVEN_AVG_60D", "CRED_GIVEN_SUM",
       "CRED_GIVEN_SUM_30D", "CRED_GIVEN_SUM_60D", "CRED_GIVEN_SUM_90D"]
    + ["CRED_PAID_PERCENT_LOAN_SUM", "CRED_PAID_PERCENT_LOAN_SUM_90D"]
    + ["CRED_REMAIN_LOAN_AVG_30D", "CRED_REMAIN_LOAN_AVG_60D", "CRED_REMAIN_LOAN_SUM",
       "CRED_REMAIN_LOAN_SUM_30D", "CRED_REMAIN_LOAN_SUM_60D", "CRED_REMAIN_LOAN_SUM_90D"]
    + ["DEPO_CNT", "DEPO_CNT_EURO", "DEPO_CNT_LCY", "DEPO_CNT_USD"]
    + [f"DEPO_{k}{c}_90D" for k in ("IN", "OUT", "REWARD", "SALDO") for c in ("", "_EURO", "_LCY", "_USD")]
    + ["DEPO_SUM", "DEPO_SUM_90D", "DEPO_SUM_LCY", "DEPO_SUM_LCY_90D", "DEPO_SUM_USD"]
    + [f"CRED_REQ_{p}_CNT_{w}" for p in ("AVTO", "BNPL", "BNPL_QR", "BZK", "IPOTEKA")
       for w in ("60D", "60D_APPR", "7D", "7D_APPR", "DAYS_FIRST", "DAYS_LAST")]
    + ["TOTAL_OPERATIONS", "COUNT_NON_DIGITAL_OPERATIONS", "COUNT_DIGITAL_OPERATIONS", "DIGITAL_RATIO",
       "IN_COUNT_WALLET", "OUT_COUNT_WALLET"]
)
_VER_HEAD = ["DATE_VALUE", "IIN", "BRANCH", "GENDER", "AGE", "CITIZENSHIP", "TRUST_PHONE_NUMBER",
             "FAMILY_STATUS", "CLIENT_SEGMENT", "RESIDENCE_COUNTRY"]
VER_COLS = (_VER_HEAD
            + [c for c in VER_TEXT_COLS if c not in _VER_HEAD and c != "Full_name"]
            + ["DOC_EXPIRY_DATE"]
            + [c for c in VER_NUM_COLS if c not in _VER_HEAD]
            + ["Full_name"])

CONTR_COLS = ["process_ref", "iin", "open_date", "close_date", "end_date", "contract_number", "status_name",
              "contract_amount", "partner_name", "restructing_fl", "product_id", "repayment_kind", "repayment_day"]
OVRD_COLS = ["rep_month", "act_date", "source_id", "contract_number", "ovd_days_cnt"]
BANKRUPT_COLS = ["date_value", "gm_system_code", "date_change", "processid", "bankruptcytype", "statementstatus",
                 "lastmodificationdate", "processcreate", "iin", "clientid", "orderid", "credit_product_type",
                 "request_amount", "amount_duration", "fio", "mobile", "city", "gender", "process_ref"]
OCRM_COLS = [
    "date_value1", "date_change1", "callid", "oper_ocrm_subject_id", "call_direct", "call_durat", "ocrm_subject_id",
    "tr_oper_ocrm_subject_id", "tnumb_from", "tnumb_to", "result_predl", "call_comment", "createdon_date",
    "modifiedon_date", "skill_gr_tsname", "talktime", "queue_name", "created_by_ocrm_subject_id", "num_dbz",
    "tsptp_date", "consult_theme", "tsnumber", "tstransfernumber", "tsclientmark", "before_conn_time",
    "tsresultid", "ocrm_caseid", "integrationid", "tslineid", "tslanguageid", "tsskillgroupid",
    "isinternetcall_fl", "iscallback_fl", "tstypeid", "date_value", "date_change", "activityid",
    "act_creat_date", "title_call_dl", "activ_sost", "company_sc", "fact_zvr_date", "client_ocrm_subject_id",
    "activ_result", "activ_result_dtl", "otv_sotr_ocrm_subject_id", "tab_num_sotr", "ocrm_tsdebt_id",
]
K2_COLS = ["dwh_id", "source_id", "gm_system_code", "date_change", "dwh_code", "date_begin", "date_end",
           "is_actual", "req_id", "code", "longname", "doc_num", "sender_k2_bank_id", "recip_k2_bank_id",
           "currency_code", "k2_knp_id", "k2_abonent_id", "accept_date", "opertype_id", "recip_client_iin",
           "close_date", "k2_req_status_id", "bs_gm_opertype_id", "k2_employee_id", "k2_requirement_group_id",
           "amount_sum", "descr", "opertype_name", "doc_date", "dogovor", "iin_bin", "system_bs"]
EXCL_COLS = ["k2_exclude_id", "client_id", "client_longname", "iin_bin", "is_active_fl", "exclude_title",
             "is_actual", "date_begin", "date_end", "updated_date"]
VNESHPT_COLS = ["created_date", "date_change", "РЕФЕРЕНС", "k2_loan_id", "k2_mt_reason_id", "k2_abonent_id",
                "k2_extreq_stat_id", "longname", "status_name", "status_code", "ext_dea_code", "ИИН",
                "Тип бизнеса", "Причина", "dwh_id", "contract_number", "code", "debt", "date_begin", "date_end"]
GEO_COLS = ["date_value", "date_change", "public_id", "login", "event_datetime", "app_name", "action",
            "hb_authorized_fl", "response_message", "ip_address", "latitude", "longitude", "client_iin",
            "device_id", "device_name", "operationsystem", "operationsystemversion", "app_version_name",
            "device_name_new"]
PKB_AGG_COLS = ["iin", "rule_name", "cnt_7d", "cnt_30d", "cnt_60d", "cnt_all", "first_ts", "last_ts"]

PKB_RULE_NAMES = [
    "Мониторинг кредитной заявки (онлайн заявки)",
    "Пролонгация по займу",
    "Любой поставщик (кроме ПКБ) искал информацию по субъекту",
    "Совокупная сумма просрочки увеличилась/уменьшилась",
    "Мониторинг появление нового займа",
    "Мониторинг факта урегулирование задолженности",
    "Мониторинг статуса «Лудоман»",
    "Отписка субъектов без активных контрактов",
    "Мониторинг статуса «Призывник»",
    "Изменения по залогу займа",
    "Мониторинг запросов на кредитный отчет",
    "У субъекта появился новый контракт",
    "Мониторинг изменения по кредиту - (Реструктуризация)",
    "Появление нового займа с целью Рефинансирования",
    "Мониторинг статуса «Стоп-Кредит»",
    "Мониторинг статуса по кредиту изменился – Суд",
    "Частичная просрочка по кредиторам (заемщик допустил просрочку)",
    "Мониторинг статуса «Банкротства»",
    "Мониторинг закрытие займа (погашение)",
    "Мониторинг закрытие займа – переуступка",
    "Сделать сравнение, у кого до погашения не менее 3 месяцев",
    "Совокупная сумма ежемесячного платежа изменилась",
    "Мониторинг изменения дней просрочки займа",
    "Совокупная сумма основного долга увеличилась/уменьшилась",
]

# dm / dm_rej: 159 columns; only request_amount, amount_duration, response_result are numeric dtype.
_DM_HEAD = ["date_value", "gm_system_code", "date_change", "processid", "processcreate", "lastmodificationdate",
            "iin", "clientid", "orderid", "credit_product_type", "request_amount", "amount_duration", "fio",
            "mobile", "city", "gender", "process_ref", "response_result"]
_DM_NUMTEXT = ["monthly_income", "interest_rate", "approved_amount", "approved_duration", "monthly_payment",
               "down_payment", "car_price", "car_year", "work_experience_months", "dependents_cnt", "age",
               "pkb_score", "pkb_active_loans_cnt", "pkb_active_loans_sum", "pkb_overdue_sum", "pkb_max_dpd",
               "pkb_req_30d", "gcvp_avg_income", "gcvp_months_cnt", "dti", "pti", "app_score",
               "behavior_score", "fraud_score", "employer_bin"]
_DM_CAT = ["channel", "partner_name", "product_code", "product_name", "currency", "car_brand", "car_model",
           "car_condition", "income_source", "employment_type", "position_category", "family_status",
           "education", "birth_date", "doc_type", "doc_issue_date", "doc_expiry_date", "reg_region",
           "fact_region", "phone_verified", "pkb_report_id", "gcvp_last_payment_date", "risk_grade",
           "decision_reason", "decision_code", "underwriter_id", "stop_factor_flag", "blacklist_flag",
           "is_salary_client", "is_halyk_employee", "insurance_flag", "kasko_flag", "branch_code", "manager_id",
           "dealer_bin", "sms_consent", "pep_flag", "status_name", "comment_text"]
_DM_N_FILLER = 159 - len(_DM_HEAD) - len(_DM_NUMTEXT) - len(_DM_CAT)
DM_COLS = _DM_HEAD + _DM_NUMTEXT + _DM_CAT + [f"attr_{i:03d}" for i in range(1, _DM_N_FILLER + 1)]

# --------------------------------------------------------------------------------------
# Reference values
# --------------------------------------------------------------------------------------
# short name, long name, city ru, city en, lat, lon, contract prefix, weight
_BRANCHES = [
    ("Алматинский ГФ", "Алматинский ГФ", "Алматы", "Almaty", 43.238, 76.946, "26", 0.24),
    ("Астанинский ГФ", "Астанинский ГФ", "Астана", "Astana", 51.128, 71.430, "29", 0.20),
    ("Шымкентский ГФ", "Шымкентский ГФ", "Шымкент", "Shymkent", 42.341, 69.590, "31", 0.09),
    ("Акмолинский ОФ", "Акмолинский Областной Филиал", "Кокшетау", "Kokshetau", 53.283, 69.396, "11", 0.05),
    ("Туркестанский ОФ", "Туркестанский Областной Филиал", "Туркестан", "Turkestan", 42.430, 68.960, "39", 0.06),
    ("Карагандинский ОФ", "Карагандинский Областной Филиал", "Караганда", "Karaganda", 49.806, 73.085, "16", 0.08),
    ("Абайский ОФ", "Абайский Областной Филиал", "Семей", "Semey", 50.430, 80.240, "22", 0.05),
    ("Атырауский ОФ", "Атырауский Областной Филиал", "Атырау", "Atyrau", 47.094, 51.923, "15", 0.06),
    ("Актюбинский ОФ", "Актюбинский Областной Филиал", "Актобе", "Aktobe", 50.283, 57.167, "12", 0.05),
    ("Павлодарский ОФ", "Павлодарский Областной Филиал", "Павлодар", "Pavlodar", 52.287, 76.967, "19", 0.04),
    ("Костанайский ОФ", "Костанайский Областной Филиал", "Костанай", "Kostanay", 53.214, 63.624, "17", 0.04),
    ("Мангистауский ОФ", "Мангистауский Областной Филиал", "Актау", "Aktau", 43.650, 51.170, "13", 0.04),
]
_PARTNERS = [  # name, weight, small risk effect
    ("ТОО «СарыаркаАвтоПром»", 0.30, 0.00),
    ("ТОО «Hyundai Auto Kazakhstan»", 0.20, -0.10),
    ("ТОО «Astana Motors»", 0.12, -0.05),
    ("ТОО «Kia Qazaqstan»", 0.11, 0.05),
    ("ТОО «Toyota Motor Kazakhstan»", 0.08, -0.25),
    ("ТОО «Chery Auto KZ»", 0.08, 0.20),
    ("ТОО «Haval Motor Kazakhstan»", 0.06, 0.15),
    ("ТОО «Автоцентр Жетысу»", 0.05, 0.25),
]
_ABROAD = [("Turkey", "Istanbul", 41.01, 28.98), ("Russia", "Moscow", 55.76, 37.62),
           ("United Arab Emirates", "Dubai", 25.20, 55.27), ("Korea, Republic of", "Seoul", 37.57, 126.98),
           ("Germany", "Berlin", 52.52, 13.40)]
_OTHER_BANKS = [("Kaspi Bank", 4), ("Евразийский банк", 16), ("Bank RBK", 21), ("Freedom Bank", 27),
                ("Банк ЦентрКредит", 8), ("ForteBank", 12)]
_SURN = ["АХМЕТОВ", "БЕКОВ", "ЖУМАБАЕВ", "КАСЫМОВ", "НУРЛАНОВ", "СЕРИКОВ", "ТУЛЕГЕНОВ", "ОМАРОВ", "ИСКАКОВ",
         "САПАРОВ", "ИВАНОВ", "ПЕТРОВ", "СМИРНОВ", "КИМ", "ПАК", "ЕРМЕКОВ", "АБДРАХМАНОВ", "КАЛИЕВ", "МУКАНОВ",
         "ДЖАКСЫБЕКОВ", "ОСПАНОВ", "БАЙМУХАНОВ", "КУЗНЕЦОВ", "ВОРОБЬЕВ", "ТЕН", "ЕСЕНОВ", "УТЕПОВ",
         "ЖАКСЫЛЫКОВ", "ТОКТАРОВ", "СУЛТАНОВ"]
_FIRST_M = ["ЕРЛАН", "АСХАТ", "НУРЛАН", "ДАНИЯР", "АРМАН", "БАУЫРЖАН", "ЕРЖАН", "ТИМУР", "РУСЛАН", "АЛЕКСАНДР",
            "СЕРГЕЙ", "ДМИТРИЙ", "АЙДОС", "ДАУРЕН", "МАРАТ", "ЕРБОЛ", "АСЛАН", "КАНАТ", "АЛИХАН", "ДАМИР"]
_FIRST_F = ["АЙГЕРИМ", "ДИНАРА", "ГУЛЬНАРА", "АСЕЛЬ", "ЖАНАР", "САУЛЕ", "МАДИНА", "АЛИЯ", "ЕЛЕНА", "ОЛЬГА",
            "НАТАЛЬЯ", "КАМИЛА", "ДАНА", "АЙЖАН", "ИНДИРА", "ЛЯЗЗАТ"]
_PATR = ["СЕРИК", "НУРЛАН", "БОЛАТ", "КАНАТ", "ЕРЛАН", "МАРАТ", "АСКАР", "ЖАНАТ", "ВИКТОР", "ИВАН", "ТАЛГАТ",
         "БАХЫТ", "РУСТАМ", "ЕРМЕК", "САМАТ", "АСЕТ"]
# device name, os, brand, premium
_DEVICES = [("iPhone", "iOS", "Apple", 0), ("iPhone 17 Pro", "iOS", "Apple", 1), ("iPhone 16 Plus", "iOS", "Apple", 0),
            ("iPhone 14 Pro", "iOS", "Apple", 1), ("iPhone 16 Pro Max", "iOS", "Apple", 1),
            ("iPhone 13", "iOS", "Apple", 0), ("iPhone 11", "iOS", "Apple", 0),
            ("SM-A356E", "android", "Samsung", 0), ("23078PND5G", "android", "Xiaomi", 0),
            ("CPH2815", "android", "OPPO", 0), ("SM-S928B", "android", "Samsung", 1),
            ("RMX3834", "android", "realme", 0), ("SM-A155F", "android", "Samsung", 0),
            ("2312DRA50G", "android", "Xiaomi", 0), ("ALI-NX1", "android", "HONOR", 0)]
_IOS_VER = ["17.6.1", "18.5", "18.7", "26.0.1", "26.6.2"]
_AND_VER = ["13", "14", "15", "16"]
_APP_VER = ["6.11.0", "6.14.2", "6.18.0", "6.22.1", "6.25.0", "7.0.3", "7.4.0", "7.8.2", "7.12.0", "7.15.1",
            "7.18.0", "7.21.1"]
_MCG = ["Продуктовые магазины", "Рестораны и кафе", "АЗС", "Онлайн-покупки", "Аптеки", "Такси",
        "Одежда и обувь", "Электроника", "Коммунальные услуги", "Азартные игры"]

_GEO_START = "2025-09-18"
_PKB_START = "2026-09-07"
_VNPT_START = "2026-01-01"
_VNESHPT_START = "2026-03-01"
_OPEN_Q = (0.0, 0.25, 0.50, 0.75, 1.0)
_OPEN_D = ("2024-04-01", "2024-09-03", "2024-12-13", "2025-09-08", "2026-03-31")
_BIG = 10_000_000          # "never" in day units
_GEO_MAX_ROWS = 580_000
_TERM_MONTHS = 84

# Delinquency model parameters (logit scale). Entry: per due date; roll: P(episode reaches 90+ DPD).
_P = dict(
    entry_a0=-4.15, entry_r=0.95, entry_recent=1.2, entry_neps=0.25, entry_eps_sd=0.35,
    roll_b0=-4.45, roll_r=1.45, roll_neps=0.2, roll_eng=-0.25, roll_eps_sd=0.55, roll_fpd=0.65,
    fpd_a0=-3.75, fpd_z=0.7, cure_late=0.35, writeoff=0.65,
)


# --------------------------------------------------------------------------------------
# Small helpers (dates are int days since 1970-01-01 internally)
# --------------------------------------------------------------------------------------
def _day(s) -> int:
    return int(np.datetime64(str(s)[:10], "D").astype("int64"))


def _month_of(days):
    return np.asarray(days, dtype="int64").astype("datetime64[D]").astype("datetime64[M]").astype("int64")


def _month_start(m):
    return np.asarray(m, dtype="int64").astype("datetime64[M]").astype("datetime64[D]").astype("int64")


def _month_end(m):
    return _month_start(np.asarray(m, dtype="int64") + 1) - 1


def _dstr(days) -> np.ndarray:
    """int days -> 'YYYY-MM-DD' (numpy unicode array)."""
    return np.asarray(days, dtype="int64").astype("datetime64[D]").astype(str)


def _ts(days, secs=0, unit="us", nat=None) -> np.ndarray:
    """int days (+ seconds) -> numpy datetime64[unit]; ``nat`` = boolean mask of NaT."""
    s = np.asarray(days, dtype="int64") * 86400 + np.asarray(secs, dtype="int64")
    out = s.astype("datetime64[s]").astype(f"datetime64[{unit}]")
    if nat is not None:
        out = out.copy()
        out[np.asarray(nat, dtype=bool)] = np.datetime64("NaT", unit)
    return out


def _const_ts(value: str, n: int, unit="us") -> np.ndarray:
    return np.full(n, np.datetime64(value, unit), dtype=f"datetime64[{unit}]")


def _S(values, na=None) -> pd.api.extensions.ExtensionArray:
    """values -> pandas 3 'str' array; ``na`` = boolean mask of missing."""
    a = np.asarray(values, dtype=object)
    if na is not None:
        a = a.copy()
        a[np.asarray(na, dtype=bool)] = None
    return pd.array(a, dtype="str")


def _S_na(n: int) -> pd.api.extensions.ExtensionArray:
    return pd.array(np.full(n, None, dtype=object), dtype="str")


def _sig(x):
    return 1.0 / (1.0 + np.exp(-x))


def _hex(rng, n: int, nchar: int, upper=True) -> list:
    if n == 0:
        return []
    nb = (nchar + 1) // 2
    raw = rng.integers(0, 256, size=n * nb, dtype=np.uint8).tobytes().hex()
    if upper:
        raw = raw.upper()
    st = 2 * nb
    return [raw[i * st:i * st + nchar] for i in range(n)]


def _guid(rng, n: int, braces=True, upper=True) -> np.ndarray:
    fmt = "{{{}-{}-{}-{}-{}}}" if braces else "{}-{}-{}-{}-{}"
    return np.array([fmt.format(s[:8], s[8:12], s[12:16], s[16:20], s[20:])
                     for s in _hex(rng, n, 32, upper)], dtype=object)


def _zfill(arr, width: int) -> np.ndarray:
    a = np.asarray(arr, dtype="int64")
    if a.size == 0:
        return np.array([], dtype=f"U{width}")
    return np.char.zfill(a.astype(str), width)


def _pick(rng, values, n, p=None) -> np.ndarray:
    vals = np.asarray(values, dtype=object)
    return vals[rng.choice(len(vals), size=n, p=p)]


def _fmt_num(x, fmt="%.2f", na=None) -> np.ndarray:
    out = np.char.mod(fmt, np.nan_to_num(np.asarray(x, dtype=float))).astype(object)
    if na is not None:
        out[np.asarray(na, dtype=bool)] = None
    return out


def _expand_months(start, last):
    """Episode [start, last] -> one item per calendar month: (episode idx, month, act day, dpd)."""
    start = np.asarray(start, dtype="int64")
    last = np.asarray(last, dtype="int64")
    m0, m1 = _month_of(start), _month_of(last)
    nm = (m1 - m0 + 1).astype("int64")
    rep = np.repeat(np.arange(len(start)), nm)
    off = np.arange(nm.sum()) - np.repeat(np.cumsum(nm) - nm, nm)
    m = m0[rep] + off
    at = np.minimum(_month_end(m), last[rep])
    dpd = at - start[rep] + 1
    return rep, m, at, dpd


def _uniq_ints(rng, n, lo, hi) -> np.ndarray:
    """n distinct ints in [lo, hi)."""
    hi = max(hi, lo + n + 1)
    return lo + rng.choice(hi - lo, size=n, replace=False)


# --------------------------------------------------------------------------------------
# World simulation: clients, contracts, latent risk, delinquency episodes
# --------------------------------------------------------------------------------------
def _simulate_world(n_contracts: int, seed: int, cut_date: str) -> SimpleNamespace:
    W = SimpleNamespace()
    names = ["clients", "contracts", "dpd", "ver", "bankrupt", "calls", "k2ot", "vnpt", "excl", "vneshpt",
             "ovrd", "geo", "dm", "pkb", "misc"]
    W.rng = {k: np.random.default_rng(s) for k, s in zip(names, np.random.SeedSequence(seed).spawn(len(names)))}
    W.cut = _day(cut_date)
    cut = W.cut

    # ---------------- clients
    rng = W.rng["clients"]
    n_cl = max(20, int(round(n_contracts / 1.05)))
    n_multi = max(0, n_contracts - n_cl)
    W.n_cl = n_cl
    z = rng.standard_normal(n_cl)                     # static latent risk
    eng = rng.standard_normal(n_cl)                   # digital engagement
    male = rng.random(n_cl) < 0.64
    young = rng.random(n_cl) < 0.055                  # born 2000+ -> IIN with leading zero
    age26 = np.clip(rng.normal(40, 9, n_cl), 27, 66)
    by = np.where(young, rng.integers(2000, 2005, n_cl), np.floor(2026 - age26)).astype(int)
    bdoy = rng.integers(0, 365, n_cl)
    birth = np.array([_day(f"{y}-01-01") for y in by]) + bdoy
    bstr = _dstr(birth)
    yy = np.array([s[2:4] for s in bstr]); mm = np.array([s[5:7] for s in bstr]); dd = np.array([s[8:10] for s in bstr])
    cg = np.where(by >= 2000, np.where(male, "5", "6"), np.where(male, "3", "4"))
    while True:
        ser = _zfill(rng.integers(0, 10000, n_cl), 4)
        chk = rng.integers(0, 10, n_cl).astype(str)
        iin = np.array([a + b + c + d + e + f for a, b, c, d, e, f in zip(yy, mm, dd, cg, ser, chk)], dtype=object)
        if len(set(iin)) == n_cl:
            break
    W.z, W.eng, W.male, W.birth, W.iin = z, eng, male, birth, iin

    sur = _pick(rng, _SURN, n_cl)
    fem_sur = np.array([s + "А" if s.endswith(("ОВ", "ЕВ", "ИН")) else s for s in sur], dtype=object)
    first = np.where(male, _pick(rng, _FIRST_M, n_cl), _pick(rng, _FIRST_F, n_cl))
    proot = _pick(rng, _PATR, n_cl)
    patr = np.where(male, np.char.add(proot.astype(str), "ОВИЧ"), np.char.add(proot.astype(str), "ОВНА"))
    kaz = rng.random(n_cl) < 0.06
    patr = np.where(kaz, np.char.add(proot.astype(str), np.where(male, "ҰЛЫ", "ҚЫЗЫ")), patr)
    nopatr = rng.random(n_cl) < 0.04
    surname = np.where(male, sur, fem_sur)
    W.fio = np.array([f"{s} {f}" if npt else f"{s} {f} {p}"
                      for s, f, p, npt in zip(surname, first, patr, nopatr)], dtype=object)
    ops = np.array(["701", "702", "705", "707", "708", "747", "771", "775", "776", "777", "778"])
    W.phone = np.char.add(np.char.add("7", ops[rng.integers(0, len(ops), n_cl)]),
                          _zfill(rng.integers(0, 10_000_000, n_cl), 7)).astype(object)
    bw = np.array([b[7] for b in _BRANCHES]); bw = bw / bw.sum()
    W.branch = rng.choice(len(_BRANCHES), n_cl, p=bw)
    W.branch_long = rng.random(n_cl) < 0.35        # spelling variant 'Областной Филиал'
    W.crm_id = _uniq_ints(rng, n_cl, 20_000_000, 30_000_000).astype(float)
    W.abs_id = _uniq_ints(rng, n_cl, 1_000_000, 9_000_000)

    # time-varying hardship shock (job loss etc.)
    has_shock = rng.random(n_cl) < _sig(-1.3 + 0.8 * z)
    W.shock_s = np.where(has_shock, rng.integers(_day("2024-02-01"), cut + 150, n_cl), _BIG)
    W.shock_d = rng.integers(150, 600, n_cl)
    W.shock_a = rng.uniform(0.8, 1.8, n_cl)
    dead = rng.random(n_cl) < 0.003
    W.death = np.where(dead, rng.integers(_day("2024-08-01"), cut - 20, n_cl), _BIG)
    abroad = rng.random(n_cl) < _sig(-4.0 + 0.7 * z)
    W.abroad_from = np.where(abroad, rng.integers(_day("2025-06-01"), cut - 40, n_cl), _BIG)
    W.nu = rng.standard_normal((4, n_cl))           # client-level nuisance (unrelated to risk) per feature group
    W.abroad_k = rng.integers(0, len(_ABROAD), n_cl)
    W.doc_exp = rng.integers(_day("2024-03-01"), _day("2034-12-31"), n_cl)
    W.doc_exp_nat = rng.random(n_cl) < 0.02
    W.log_inc = np.log(420_000) - 0.22 * z + 0.45 * rng.standard_normal(n_cl)
    W.has_depo = rng.random(n_cl) < _sig(0.3 - 0.7 * z + 0.3 * rng.standard_normal(n_cl))
    W.log_depo = np.log(250_000) - 0.5 * z + 1.1 * rng.standard_normal(n_cl)
    W.app_user = rng.random(n_cl) < _sig(1.9 + 0.7 * eng - 0.3 * z)
    W.salary = rng.random(n_cl) < _sig(-0.2 - 0.35 * z)
    W.employee = rng.random(n_cl) < 0.012
    ios = rng.random(n_cl) < _sig(-0.3 + 0.8 * (W.log_inc - np.log(420_000)))
    ios_idx = np.array([i for i, d in enumerate(_DEVICES) if d[1] == "iOS"])
    and_idx = np.array([i for i, d in enumerate(_DEVICES) if d[1] == "android"])
    W.dev1 = np.where(ios, ios_idx[rng.integers(0, len(ios_idx), n_cl)], and_idx[rng.integers(0, len(and_idx), n_cl)])
    W.dev2 = np.where(rng.random(n_cl) < 0.2, rng.integers(0, len(_DEVICES), n_cl), -1)
    # other banks the client holds accounts in (for vneshpt)
    W.banks = [np.unique(np.concatenate([[0] if rng.random() < 0.85 else [],
                                         rng.choice(np.arange(1, len(_OTHER_BANKS)), rng.integers(0, 3), replace=False)])
                         ).astype(int) for _ in range(n_cl)]
    W.acc_num = np.array([f"{_BRANCHES[b][6]}77770000{s:06d}" for b, s in
                          zip(W.branch, _uniq_ints(rng, n_cl, 400_000, 999_999))], dtype=object)

    # ---------------- contracts
    rng = W.rng["contracts"]
    n_c = n_cl + n_multi
    cl = np.concatenate([np.arange(n_cl), rng.choice(n_cl, n_multi, replace=False)])
    open_lo, open_hi = _day(_OPEN_D[0]), min(_day(_OPEN_D[-1]), cut - 60)
    od = np.interp(rng.random(n_c), _OPEN_Q, [_day(d) for d in _OPEN_D]).round().astype("int64")
    od = np.clip(od, open_lo, open_hi)
    for j in range(n_cl, n_c):                     # second contract of the same client opens later
        i = cl[j]
        a, b = sorted((od[i], od[j]))
        if b - a < 45:
            b = a + 45 + int(rng.integers(0, 200))
        if b > open_hi:
            a, b = a - (b - open_hi), open_hi
        od[i], od[j] = max(a, open_lo), b
    W.n_c, W.c_cl, W.c_open = n_c, cl, od
    order = np.argsort(od, kind="stable")
    rank = np.empty(n_c, dtype=int); rank[order] = np.arange(n_c)
    pref_num = np.sort(_uniq_ints(rng, n_c, 140_000, 140_000 + 30 * n_c))[rank]
    W.c_pref = np.array([f"DGA{p:012d}" for p in pref_num], dtype=object)
    cn_ser = np.sort(_uniq_ints(rng, n_c, 60_000, 300_000))[rank]
    W.c_num = np.array([f"{_BRANCHES[W.branch[c]][6]}77770000{s:06d}" for c, s in zip(cl, cn_ser)], dtype=object)
    inc_rank = (W.log_inc[cl] - W.log_inc.mean()) / W.log_inc.std()
    W.c_amount = np.round(8e6 + 4e6 * np.clip(rng.beta(2, 2.2, n_c) + 0.08 * inc_rank, 0, 1), -3)
    rate = rng.choice([0.04, 0.14, 0.17, 0.199], n_c, p=[0.15, 0.35, 0.35, 0.15]) / 12
    W.c_inst = W.c_amount * rate / (1 - (1 + rate) ** -_TERM_MONTHS)
    W.c_rate = rate
    oday = od.astype("datetime64[D]").astype(object)
    W.c_repday = np.where(rng.random(n_c) < 0.9, np.clip([d.day for d in oday], 1, 28), rng.integers(1, 29, n_c))
    W.c_end = np.array([_day(f"{d.year + 7}-{d.month:02d}-{d.day:02d}") if not (d.month == 2 and d.day == 29)
                        else _day(f"{d.year + 7}-03-01") for d in oday])
    pw = np.array([p[1] for p in _PARTNERS]); pw = pw / pw.sum()
    W.c_partner = rng.choice(len(_PARTNERS), n_c, p=pw)
    W.c_repkind = np.where(rng.random(n_c) < 0.96, 1, 2)
    early = (rng.random(n_c) < 0.03) & (z[cl] < 0.5)
    W.c_close = np.where(early, od + rng.integers(150, 700, n_c), _BIG)
    W.c_close = np.where(W.c_close > cut, _BIG, W.c_close)
    W.c_status = np.where(W.c_close < _BIG, "Закрыт", "Актуален").astype(object)

    _simulate_dpd(W)
    return W


def _risk_at(W, ci, t):
    """Latent risk of client(s) ``ci`` at day(s) ``t`` (static z + temporary hardship shock)."""
    on = (t >= W.shock_s[ci]) & (t < W.shock_s[ci] + W.shock_d[ci])
    return W.z[ci] + W.shock_a[ci] * on


def _simulate_dpd(W):
    rng = W.rng["dpd"]
    cut, n_c, cl = W.cut, W.n_c, W.c_cl
    open_m = _month_of(W.c_open)
    kmax = int(_month_of(cut) - open_m.min())
    busy = np.full(n_c, -_BIG, dtype="int64")       # last day of the current/previous episode
    last_end = np.full(n_c, -_BIG, dtype="int64")
    n_eps = np.zeros(n_c, dtype="int64")
    P = _P
    fpd = rng.random(n_c) < _sig(P["fpd_a0"] + P["fpd_z"] * W.z[cl])
    eps_c = rng.normal(0, P["entry_eps_sd"], n_c)
    part_eff = np.array([p[2] for p in _PARTNERS])[W.c_partner]
    E = {k: [] for k in ("c", "start", "end", "roll", "restr", "fpd")}
    for k in range(1, kmax + 1):
        m = open_m + k
        due = _month_start(m) + W.c_repday - 1
        elig = (due + 1 <= cut) & (due > busy) & (due < W.c_close) & (k <= _TERM_MONTHS)
        idx = np.nonzero(elig)[0]
        if idx.size == 0:
            continue
        ci, t = cl[idx], due[idx]
        r = _risk_at(W, ci, t)
        recent = (t - last_end[idx]) < 180
        mon = (_month_of(t) % 12) + 1
        season = np.where(np.isin(mon, (1, 2, 9)), 0.15, 0.0)
        logit = (P["entry_a0"] + P["entry_r"] * r + P["entry_recent"] * recent
                 + P["entry_neps"] * np.minimum(n_eps[idx], 3) + part_eff[idx] + eps_c[idx] + season)
        p = _sig(logit)
        is_fpd = fpd[idx] & (k == 1)
        dead = t >= W.death[ci]
        p = np.where(is_fpd | dead, 1.0, p)
        enter = rng.random(idx.size) < p
        if not enter.any():
            continue
        idx, ci, t, r = idx[enter], ci[enter], t[enter], r[enter]
        is_fpd, dead = is_fpd[enter], dead[enter]
        nn = idx.size
        lr = (P["roll_b0"] + P["roll_r"] * r + P["roll_neps"] * np.minimum(n_eps[idx], 3)
              + P["roll_eng"] * W.eng[ci] + rng.normal(0, P["roll_eps_sd"], nn))
        p90 = np.where(is_fpd, P["roll_fpd"], _sig(lr))
        p90 = np.where(dead, 1.0, p90)
        roll = rng.random(nn) < p90
        # duration of non-rolling episodes: fast (1-7), mid (8-30), long (31-89)
        w = np.stack([np.full(nn, 0.55), np.full(nn, 0.30), 0.15 * np.exp(0.4 * r)], 1)
        w = w / w.sum(1, keepdims=True)
        u = rng.random(nn)
        bucket = (u > w[:, 0]).astype(int) + (u > w[:, 0] + w[:, 1]).astype(int)
        L = np.select([bucket == 0, bucket == 1], [rng.integers(1, 8, nn), rng.integers(8, 31, nn)],
                      rng.integers(31, 90, nn))
        restr = (bucket == 2) & (rng.random(nn) < 0.10)
        cure_late = rng.random(nn) < P["cure_late"]
        L = np.where(roll, np.where(cure_late & ~dead, rng.integers(95, 270, nn), _BIG), L)
        restr = np.where(roll, cure_late & ~dead & (rng.random(nn) < 0.5), restr)
        start = t + 1
        end = np.where(L >= _BIG, _BIG, start + L - 1)
        for key, val in (("c", idx), ("start", start), ("end", end), ("roll", roll), ("restr", restr),
                         ("fpd", is_fpd)):
            E[key].append(val)
        busy[idx] = end
        last_end[idx] = end
        n_eps[idx] += 1
    E = {k: np.concatenate(v) if v else np.array([], dtype="int64") for k, v in E.items()}
    # write-offs: long never-cured episodes, contract leaves the book
    never = (E["end"] >= _BIG) & (E["start"] + 400 < cut)
    wo = never & (rng.random(E["c"].size) < P["writeoff"])
    wo_day = E["start"] + rng.integers(390, 450, E["c"].size)
    for i in np.nonzero(wo)[0]:
        c = E["c"][i]
        if wo_day[i] <= cut and W.c_close[c] >= _BIG:
            W.c_close[c] = wo_day[i]
            W.c_status[c] = "Списан"
    W.c_restr = np.zeros(n_c, dtype="int64")
    W.c_restr[E["c"][E["restr"].astype(bool)]] = 1
    # episode last delinquent day as observed in data (cut / write-off)
    E["last"] = np.minimum(np.minimum(E["end"], cut), W.c_close[E["c"]] - 1)
    keep = E["last"] >= E["start"]
    E = {k: v[keep] for k, v in E.items()}
    E["cl"] = cl[E["c"]]
    E["roll"] = E["roll"].astype(bool)
    o = np.lexsort((E["start"], E["c"]))
    W.E = {k: v[o] for k, v in E.items()}
    W.E["crm_debt_id"] = _uniq_ints(rng, len(o), 4_000_000, 9_999_999).astype(float)
    # first two rolling ("bad") episode starts per client -> drives pre-delinquency signals
    W.bad1 = np.full(W.n_cl, _BIG, dtype="int64")
    W.bad2 = np.full(W.n_cl, _BIG, dtype="int64")
    rs = np.nonzero(W.E["roll"])[0]
    for i in rs[np.argsort(W.E["start"][rs], kind="stable")]:
        c, s = W.E["cl"][i], W.E["start"][i]
        if W.bad1[c] >= _BIG:
            W.bad1[c] = s
        elif W.bad2[c] >= _BIG and s != W.bad1[c]:
            W.bad2[c] = s
    # pre-default deterioration is visible in the data only for part of the bad episodes
    u1, u2 = rng.random(W.n_cl), rng.random(W.n_cl)
    W.vbad1 = np.where(u1 < 0.3, W.bad1, _BIG); W.vbad2 = np.where(u2 < 0.3, W.bad2, _BIG)
    W.gbad1 = np.where(u1 < 0.6, W.bad1, _BIG); W.gbad2 = np.where(u2 < 0.6, W.bad2, _BIG)


def _near_bad(W, ci, t, before=90, after=150):
    b1, b2 = W.vbad1[ci], W.vbad2[ci]
    return (((t >= b1 - before) & (t <= b1 + after)) | ((t >= b2 - before) & (t <= b2 + after)))


def _dpd_at(W, ci, t):
    """Current max DPD of client(s) at day(s) t (0 if none). Vectorised via merge."""
    E = W.E
    left = pd.DataFrame({"row": np.arange(len(ci)), "cl": ci, "t": t})
    right = pd.DataFrame({"cl": E["cl"], "s": E["start"], "l": E["last"]})
    m = left.merge(right, on="cl", how="inner")
    m = m[(m["s"] <= m["t"]) & (m["t"] <= m["l"])]
    out = np.zeros(len(ci), dtype="int64")
    if len(m):
        g = (m["t"] - m["s"] + 1).groupby(m["row"]).max()
        out[g.index.to_numpy()] = g.to_numpy()
    return out


# --------------------------------------------------------------------------------------
# Table builders
# --------------------------------------------------------------------------------------
def _build_contr(W) -> pd.DataFrame:
    rng = W.rng["contracts"]
    n = W.n_c
    dup = np.nonzero(rng.random(n) < 0.01)[0]          # duplicated contracts with another end_date
    idx = np.concatenate([np.arange(n), dup])
    end = W.c_end[idx].copy()
    end[n:] = end[n:] - rng.integers(5, 25, dup.size)
    close = W.c_close[idx]
    df = pd.DataFrame({
        "process_ref": _S(W.c_pref[idx]),
        "iin": _S(W.iin[W.c_cl[idx]]),
        "open_date": _ts(W.c_open[idx], unit="ns"),
        "close_date": _ts(np.where(close >= _BIG, 0, close), unit="ns", nat=close >= _BIG),
        "end_date": _ts(end, unit="ns"),
        "contract_number": _S(W.c_num[idx]),
        "status_name": _S(W.c_status[idx]),
        "contract_amount": W.c_amount[idx].astype("float64"),
        "partner_name": _S(np.array([p[0] for p in _PARTNERS], dtype=object)[W.c_partner[idx]]),
        "restructing_fl": W.c_restr[idx].astype("int64"),
        "product_id": _S(np.full(idx.size, "16158", dtype=object)),
        "repayment_kind": W.c_repkind[idx].astype("int64"),
        "repayment_day": _S(W.c_repday[idx].astype(str)),
    })
    df = df.iloc[rng.permutation(len(df))]
    df.index = np.sort(rng.choice(int(len(df) * 1.7), len(df), replace=False))   # pre-filtered, non-contiguous
    return df[CONTR_COLS]


def _build_ovrd(W) -> pd.DataFrame:
    rng = W.rng["ovrd"]
    E = W.E
    rep, m, at, dpd = _expand_months(E["start"], E["last"])
    c = E["c"][rep]
    d = pd.DataFrame({"c": c, "m": m, "at": at, "dpd": dpd})
    d = d.sort_values(["c", "m", "dpd"]).drop_duplicates(["c", "m"], keep="last")   # one row per contract-month
    cn = W.c_num[d["c"].to_numpy()]
    # other products of the same / other clients (not in contr)
    n_oth = max(1, int(round(len(d) * 0.03 / 0.97 / 2.5)))
    o_start = rng.integers(_day("2024-05-01"), W.cut, n_oth)
    o_len = np.where(rng.random(n_oth) < 0.8, rng.integers(1, 40, n_oth), rng.integers(40, 400, n_oth))
    o_last = np.minimum(o_start + o_len - 1, W.cut)
    o_rep, o_m, o_at, o_dpd = _expand_months(o_start, o_last)
    pre = np.array(["29998400000", "40080110006", "22991810001", "31999840000"])
    o_cn = np.char.add(pre[rng.integers(0, len(pre), n_oth)], _zfill(rng.integers(0, 100000, n_oth), 5))[o_rep]
    months = np.concatenate([d["m"].to_numpy(), o_m])
    acts = np.concatenate([d["at"].to_numpy(), o_at])
    dpds = np.concatenate([d["dpd"].to_numpy(), o_dpd]).astype("float64")
    cns = np.concatenate([cn.astype(object), o_cn.astype(object)])
    o = np.lexsort((rng.random(months.size), months))
    df = pd.DataFrame({
        "rep_month": _S(_dstr(_month_start(months[o]))),
        "act_date": _S(_dstr(acts[o])),
        "source_id": (3_600_000_000 + rng.integers(0, 99_999_999, months.size)).astype("int64"),
        "contract_number": _S(cns[o]),
        "ovd_days_cnt": dpds[o],
    })
    return df[OVRD_COLS].reset_index(drop=True)


def _snapshot_days(cut: int) -> np.ndarray:
    last_snap = cut - 10
    m_last = _month_of(cut) - 3
    me = _month_end(np.arange(_month_of(_day("2024-04-30")), m_last + 1))
    me = me[me < last_snap]
    extra = sorted({last_snap - k for k in range(6)} | {last_snap - k for k in (15, 20, 35, 50, 60)})
    extra = [d for d in extra if d > me[-1]] if len(me) else extra
    return np.concatenate([me, np.array(extra, dtype="int64")])


def _build_ver(W) -> pd.DataFrame:
    rng = W.rng["ver"]
    snaps = _snapshot_days(W.cut)
    first_open = np.full(W.n_cl, _BIG, dtype="int64")
    np.minimum.at(first_open, W.c_cl, W.c_open)
    start_snap = _month_end(_month_of(first_open) - 2)
    first_i = np.searchsorted(snaps, start_snap)
    cnt = len(snaps) - first_i
    ci = np.repeat(np.arange(W.n_cl), cnt)
    si = np.arange(cnt.sum()) - np.repeat(np.cumsum(cnt) - cnt, cnt) + first_i[ci]
    t = snaps[si]
    o = np.lexsort((W.iin[ci].astype(str), t))            # final row order: DATE_VALUE, IIN
    ci, t = ci[o], t[o]
    R = ci.size
    N = lambda: rng.standard_normal(R)  # noqa: E731
    U = lambda: rng.random(R)  # noqa: E731

    r = _risk_at(W, ci, t)
    shock = ((t >= W.shock_s[ci]) & (t < W.shock_s[ci] + W.shock_d[ci])).astype(float)
    dpd = _dpd_at(W, ci, t)
    nearbad = _near_bad(W, ci, t, before=120, after=60).astype(float)
    dead = t >= W.death[ci]
    abroad = t >= W.abroad_from[ci]
    age = np.floor((t - W.birth[ci]) / 365.25)
    rf = r + 0.8 * W.nu[0][ci] + 0.4 * N()          # risk as seen through financial behaviour
    rp = r + 0.8 * W.nu[1][ci] + 0.4 * N()          # risk as seen through the credit bureau
    hunger = 0.5 * rf + 0.4 * shock + 0.4 * nearbad + 0.6 * N()
    eng = W.eng[ci] - 0.2 * rf - 0.3 * nearbad + 0.35 * N()
    app = W.app_user[ci] & ~dead

    def ln(mu, s=0.6):
        return np.exp(mu + s * N())

    def zi(x, pz):
        x = np.asarray(x, dtype=float).copy()
        x[U() < pz] = 0.0
        return x

    def pois(lam):
        return rng.poisson(np.broadcast_to(np.clip(lam, 0, 1e7), (R,))).astype(float)

    # ---- loan outstanding on the auto loan(s)
    debt = np.zeros(R); inst = np.zeros(R)
    cdf = pd.DataFrame({"cl": W.c_cl, "c": np.arange(W.n_c)})
    m = pd.DataFrame({"row": np.arange(R), "cl": ci, "t": t}).merge(cdf, on="cl")
    cc = m["c"].to_numpy(); tt = m["t"].to_numpy()
    open_ = (tt >= W.c_open[cc]) & (tt < W.c_close[cc])
    k = np.clip((tt - W.c_open[cc]) // 30.4, 0, _TERM_MONTHS)
    i_ = W.c_rate[cc]
    bal = W.c_amount[cc] * ((1 + i_) ** _TERM_MONTHS - (1 + i_) ** k) / ((1 + i_) ** _TERM_MONTHS - 1)
    np.add.at(debt, m["row"].to_numpy(), np.where(open_, bal, 0.0))
    np.add.at(inst, m["row"].to_numpy(), np.where(open_, W.c_inst[cc], 0.0))

    inc = np.exp(W.log_inc[ci] - 0.4 * shock - 0.25 * (dpd > 0) - 0.15 * nearbad + 0.3 * N())
    inc = np.where(dead, 0.0, inc)
    depo = np.where(W.has_depo[ci], np.exp(W.log_depo[ci] - 0.5 * shock - 0.3 * nearbad + 0.4 * N()), 0.0)
    depo = np.where(U() < 0.06, 0.0, depo)
    sal = W.salary[ci]

    X = {}
    X["AGE"] = age
    X["CNT_DEPOSIT"] = np.where(depo > 0, 1 + pois(0.4 + 0 * depo), 0.0)
    X["CNT_OPEN_DEPOSIT"] = np.minimum(X["CNT_DEPOSIT"], pois(0.15 * (depo > 0)))
    X["CNT_OPEN_DEPOSIT_ONLINE"] = np.minimum(X["CNT_OPEN_DEPOSIT"], pois(0.1 * (depo > 0)))
    X["CNT_CLOSE_DEPO"] = pois(0.08 + 0.15 * shock)
    X["AMOUNT_REPLENISH_DEPO"] = zi(depo * 0.15 * ln(0, 0.8), 0.45 + 0.2 * shock)
    X["AMOUNT_WITHDRAWAL_DEPO"] = zi(depo * 0.12 * ln(0.3 * shock, 0.8), 0.55)
    X["CNT_LOAN"] = 1 + pois(0.35 + 0.25 * np.maximum(hunger, 0))
    X["CNT_LOAN_ONLINE"] = np.minimum(X["CNT_LOAN"] - 1, pois(0.3 + 0.2 * np.maximum(hunger, 0)))
    other_loan = np.where(X["CNT_LOAN"] > 1, ln(np.log(900_000), 0.8) * (X["CNT_LOAN"] - 1), 0.0)
    X["AMOUNT_LOAN_MAIN_DEBT"] = debt + other_loan
    X["AMOUNT_LOAN_OVERDUE"] = np.where(dpd > 0, inst * np.ceil(dpd / 30.0) * rng.uniform(0.9, 1.1, R), 0.0)
    X["MONTH_LOAN_PAY_AMOUNT"] = inst + other_loan * 0.045
    X["AMOUNT_OUT_MONEY"] = inc * rng.uniform(0.6, 1.1, R)
    X["AMOUNT_SOC_PAYOUT"] = zi(ln(np.log(55_000), 0.3), 0.93)
    pens_age = np.where(W.male[ci], 63, 61)
    X["AMOUNT_PENSION"] = np.where(age >= pens_age, ln(np.log(140_000), 0.25), zi(ln(np.log(90_000), 0.3), 0.99))
    X["AMOUNT_CASHIN_ATM"] = zi(inc * 0.2 * ln(0, 0.8), 0.6)
    X["VOLUME_NONCASH_TRANS"] = inc * 0.45 * ln(0, 0.4)
    X["AMOUNT_CASHOUT_ATM"] = zi(inc * 0.25 * ln(0.2 * (1 - sal), 0.8), 0.35)
    X["CNT_INTER_TRANSFER_HB"] = pois(3 + 2 * np.maximum(eng + 1, 0))
    X["AMOUNT_INTER_TRANSFER_HB"] = X["CNT_INTER_TRANSFER_HB"] * ln(np.log(25_000), 0.7)
    X["CNT_TRAN_OTHER_BANK"] = pois(2 + 0.5 * np.maximum(hunger, 0))
    X["AMOUNT_TRAN_OTHER_BANK"] = X["CNT_TRAN_OTHER_BANK"] * ln(np.log(40_000), 0.8)
    X["CNT_TRANSFER_KASPI"] = pois(4 + 1.5 * np.maximum(hunger, 0))
    X["AMOUNT_TRANSFER_KASPI"] = X["CNT_TRANSFER_KASPI"] * ln(np.log(30_000), 0.8)
    X["CNT_KASPI_CARD_TRANSFER"] = pois(2 + 1.0 * np.maximum(hunger, 0))
    sh_h = np.where(sal, 0.65, 0.15) * rng.uniform(0.7, 1.3, R)
    sh_k = rng.uniform(0.1, 0.4, R)
    sh_o = np.clip(1 - sh_h - sh_k, 0.02, None)
    X["VOLUME_INBOX_TRANSF_HALYK"] = zi(inc * sh_h, 0.05)
    X["VOLUME_INBOX_TRANSF_KASPI"] = zi(inc * sh_k, 0.15)
    X["VOLUME_INBOX_TRANSF_ANOT_BANK"] = zi(inc * sh_o, 0.25)
    for a, b in (("HALYK", 120_000), ("KASPI", 60_000), ("ANOT_BANK", 90_000)):
        vol = X[f"VOLUME_INBOX_TRANSF_{a}"]
        X[f"CNT_INBOX_TRANSF_{a}"] = np.where(vol > 0, 1 + pois(vol / b), 0.0)
    X["CNT_PAYMENT"] = pois(6 + 3 * np.maximum(eng + 1, 0))
    X["PAYMENT_AMOUNT"] = X["CNT_PAYMENT"] * ln(np.log(9_000), 0.6)
    X["CNT_BANK_PRODUCTS_DEBIT_CARD"] = 1 + pois(0.6)
    nsvc = 0
    for nm, lam, amt in (("MOBILE", 1.5, 3_000), ("COMMUNAL", 1.0, 25_000), ("PARKING", 0.4, 1_500),
                         ("FINE", 0.15, 15_000), ("TRANSPORT", 0.8, 2_000), ("EDUCATION", 0.15, 60_000)):
        cn = pois(lam * np.exp(0.3 * eng)) * app
        sfx = "_HB" if nm in ("MOBILE", "COMMUNAL") else ""
        X[f"CNT_{nm}_PAY{sfx}"] = cn
        X[f"AMOUNT_{nm}_PAY{sfx}"] = cn * ln(np.log(amt), 0.5)
        nsvc = nsvc + (cn > 0)
    X["CNT_UNIQUE_PAY_SERVICES"] = nsvc.astype(float) + pois(0.5 * app)
    sess = pois(np.exp(np.log(38) + 0.45 * eng - 0.7 * (dpd > 30) + 0.3 * N())) * app
    X["CNT_SESSIONS"] = sess
    X["CNT_SUCCESS_SESSIONS"] = rng.binomial(sess.astype(int), 0.95).astype(float)
    X["USETIME_USER_HB"] = sess * rng.uniform(1.5, 6.0, R)
    X["AVG_USETIME_DAY"] = X["USETIME_USER_HB"] / 30.0
    X["CNT_UNIQUE_DAY_IN_MONTH"] = np.minimum(31, rng.binomial(31, np.clip(sess / 40, 0, 0.98)).astype(float))
    X["HALYK_EMPLOYEE"] = W.employee[ci].astype(float)
    X["CNT_LOAN_PKB_WITHOUT_HB"] = pois(0.6 + 0.6 * np.maximum(hunger, 0))
    X["LOAN_BALANCE_KASPI"] = zi(ln(np.log(350_000) + 0.3 * hunger, 0.9), _sig(0.6 - 0.9 * hunger))
    X["ADDICTION_TO_BZK"] = np.clip(_sig(-1.2 + 0.8 * hunger + 0.5 * N()), 0, 1)
    X["ADDICTION_TO_TOVARKA"] = np.clip(_sig(-1.6 + 0.6 * hunger + 0.6 * N()), 0, 1)
    wl = rng.dirichlet([2, 3, 2, 4], R)
    for j, nm in enumerate(("INSTALLMENT", "DEPOSIT", "CREDIT", "MARKET")):
        X[f"WALLET_TO_{nm}"] = wl[:, j]
    X["CNT_BANK_PRODUCTS"] = 2 + pois(2 + 0.5 * (depo > 0))
    lam = np.exp(-3.4 + 0.45 * rp + 0.5 * shock + 0.4 * nearbad + 0.5 * N())   # bureau inquiries per day
    acc = np.zeros(R); prev = 0
    for w in (7, 30, 90, 120, 180, 360):
        acc = acc + pois(lam * (w - prev)); prev = w
        X[f"PKB_REQ_{w}D"] = acc.copy()
    X["PKB_ACT_CONTRACT_CNT_KZT"] = X["CNT_LOAN"] + X["CNT_LOAN_PKB_WITHOUT_HB"]
    X["PKB_ACT_CONTRACT_CNT_USD"] = pois(0.002 * np.ones(R))
    X["PKB_COMPL_CONTRACT_CNT_KZT"] = pois(2.5 + 0.5 * np.maximum(hunger, 0))
    X["PKB_COMPL_CONTRACT_CNT_USD"] = pois(0.01 * np.ones(R))
    X["PKB_RETURN_CONTRACT_CNT_KZT"] = pois(0.05 + 0.1 * np.maximum(rp, 0))
    X["PKB_RETURN_CONTRACT_CNT_USD"] = np.zeros(R)
    other_bank_debt = np.where(X["CNT_LOAN_PKB_WITHOUT_HB"] > 0,
                               X["CNT_LOAN_PKB_WITHOUT_HB"] * ln(np.log(700_000), 0.9), 0.0)
    X["PKB_SUM_UNPAY_ACT"] = X["AMOUNT_LOAN_MAIN_DEBT"] + other_bank_debt + X["LOAN_BALANCE_KASPI"]
    p_ovr = _sig(-2.7 + 0.7 * rp + 2.0 * (dpd > 0) + 0.6 * shock + 0.4 * nearbad)
    X["SUM_PKB_OVRD_ACT"] = np.where(U() < p_ovr, X["AMOUNT_LOAN_OVERDUE"] + ln(np.log(60_000) + 0.3 * rp, 1.0), 0.0)
    X["SUM_PKB_UNPAY_COMPL"] = zi(ln(np.log(1_500_000), 1.0), 0.3)
    X["SUM_PKB_OVRD_COMPL"] = zi(ln(np.log(40_000) + 0.3 * rp, 1.0), _sig(1.6 - 0.5 * rp))
    X["SUM_PKB_UNPAY_RECALL"] = zi(ln(np.log(200_000), 1.0), 0.985 - 0.01 * (rp > 1))
    X["CRED_CNT_ACTIVE_KZT"] = X["CNT_LOAN"]
    X["CRED_CNT_ACTIVE_USD"] = np.zeros(R)
    X["CRED_CNT_ACTIVE"] = X["CRED_CNT_ACTIVE_KZT"]
    X["CRED_CNT_NON_ACTIVE_KZT"] = pois(1.2 + 0.3 * np.maximum(hunger, 0))
    X["CRED_CNT_NON_ACTIVE_USD"] = np.zeros(R)
    X["CRED_CNT_NON_ACTIVE"] = X["CRED_CNT_NON_ACTIVE_KZT"]
    X["CRED_CNT"] = X["CRED_CNT_ACTIVE"] + X["CRED_CNT_NON_ACTIVE"]
    for i in (1, 2, 3):
        X[f"CRED_CNT_REPAYMENT_DAY_{i}_BZK"] = pois(0.05 * i + 0.1 * np.maximum(hunger, 0))
    new_amt = ln(np.log(800_000), 0.9)
    pnew = _sig(-3.2 + 0.7 * hunger)
    g30 = (U() < pnew) * new_amt
    g60 = g30 + (U() < pnew) * new_amt
    g90 = g60 + (U() < pnew) * new_amt
    g360 = g90 + (U() < 3 * pnew) * new_amt * 1.5
    X["CRED_DOGOVOR_SUM_30D"], X["CRED_DOGOVOR_SUM_60D"], X["CRED_DOGOVOR_SUM_90D"] = g30, g60, g90
    X["CRED_DOGOVOR_SUM"] = g360
    X["CRED_DOGOVOR_AVG_60D"] = g60 / 2
    X["CRED_DOGOVOR_AVG_360D"] = g360 / 12
    X["CRED_GIVEN_SUM_30D"], X["CRED_GIVEN_SUM_60D"], X["CRED_GIVEN_SUM_90D"] = g30 * 0.97, g60 * 0.97, g90 * 0.97
    X["CRED_GIVEN_SUM"] = g360 * 0.97
    X["CRED_GIVEN_AVG_60D"], X["CRED_GIVEN_AVG_180D"], X["CRED_GIVEN_AVG_360D"] = g60 / 2, g90 / 6, g360 / 12
    X["CRED_PAID_PERCENT_LOAN_SUM"] = X["AMOUNT_LOAN_MAIN_DEBT"] * 0.012 * (dpd == 0)
    X["CRED_PAID_PERCENT_LOAN_SUM_90D"] = X["CRED_PAID_PERCENT_LOAN_SUM"] * rng.uniform(2.4, 3.0, R)
    rem = X["AMOUNT_LOAN_MAIN_DEBT"]
    X["CRED_REMAIN_LOAN_SUM"] = rem
    X["CRED_REMAIN_LOAN_SUM_30D"] = rem * 1.012
    X["CRED_REMAIN_LOAN_SUM_60D"] = rem * 1.024
    X["CRED_REMAIN_LOAN_SUM_90D"] = rem * 1.036
    X["CRED_REMAIN_LOAN_AVG_30D"] = rem * 1.006
    X["CRED_REMAIN_LOAN_AVG_60D"] = rem * 1.012
    usd = (depo > 0) & (U() < 0.12)
    eur = (depo > 0) & (U() < 0.03)
    X["DEPO_CNT_LCY"] = np.where(depo > 0, X["CNT_DEPOSIT"], 0.0)
    X["DEPO_CNT_USD"] = usd.astype(float)
    X["DEPO_CNT_EURO"] = eur.astype(float)
    X["DEPO_CNT"] = X["DEPO_CNT_LCY"] + X["DEPO_CNT_USD"] + X["DEPO_CNT_EURO"]
    d_usd = np.where(usd, depo * rng.uniform(0.2, 1.5, R), 0.0)
    d_eur = np.where(eur, depo * rng.uniform(0.1, 0.8, R), 0.0)
    X["DEPO_SUM_LCY"] = depo
    X["DEPO_SUM_USD"] = d_usd
    X["DEPO_SUM"] = depo + d_usd + d_eur
    X["DEPO_SUM_LCY_90D"] = depo * rng.uniform(0.85, 1.25, R) * np.exp(0.4 * shock)
    X["DEPO_SUM_90D"] = X["DEPO_SUM_LCY_90D"] + d_usd + d_eur
    for cur, base, f in (("LCY", depo, 1.0), ("USD", d_usd, 0.3), ("EURO", d_eur, 0.2)):
        din = zi(base * 0.3 * ln(0, 0.7), 0.4)
        dout = zi(base * 0.25 * ln(0.4 * shock, 0.7), 0.5)
        X[f"DEPO_IN_{cur}_90D"], X[f"DEPO_OUT_{cur}_90D"] = din, dout
        X[f"DEPO_REWARD_{cur}_90D"] = base * 0.03 * f
        X[f"DEPO_SALDO_{cur}_90D"] = din - dout
    for kk in ("IN", "OUT", "REWARD", "SALDO"):
        X[f"DEPO_{kk}_90D"] = X[f"DEPO_{kk}_LCY_90D"] + X[f"DEPO_{kk}_USD_90D"] + X[f"DEPO_{kk}_EURO_90D"]
    for p_, base, coef in (("AVTO", 0.03, 0.1), ("BNPL", 0.25, 0.5), ("BNPL_QR", 0.15, 0.5), ("BZK", 0.12, 0.9),
                           ("IPOTEKA", 0.01, 0.0)):
        lam60 = base * np.exp(coef * hunger)
        n60 = pois(lam60)
        n7 = rng.binomial(n60.astype(int), 0.12).astype(float)
        appr = np.clip(0.75 - 0.15 * hunger, 0.05, 0.95)
        X[f"CRED_REQ_{p_}_CNT_60D"] = n60
        X[f"CRED_REQ_{p_}_CNT_7D"] = n7
        X[f"CRED_REQ_{p_}_CNT_60D_APPR"] = rng.binomial(n60.astype(int), appr).astype(float)
        X[f"CRED_REQ_{p_}_CNT_7D_APPR"] = np.minimum(n7, X[f"CRED_REQ_{p_}_CNT_60D_APPR"])
        ever = (n60 > 0) | (U() < 6 * lam60 / (1 + 6 * lam60))
        last = np.where(n60 > 0, rng.integers(0, 60, R), rng.integers(60, 700, R)).astype(float)
        X[f"CRED_REQ_{p_}_CNT_DAYS_LAST"] = np.where(ever, last, np.nan)
        X[f"CRED_REQ_{p_}_CNT_DAYS_FIRST"] = np.where(ever, last + rng.integers(0, 900, R), np.nan)
    qr_launch = _day("2025-06-01")
    for w in ("60D", "60D_APPR", "7D", "7D_APPR", "DAYS_FIRST", "DAYS_LAST"):
        X[f"CRED_REQ_BNPL_QR_CNT_{w}"] = np.where(t < qr_launch, np.nan, X[f"CRED_REQ_BNPL_QR_CNT_{w}"])
    dig = pois(sess * 2.2 + X["CNT_PAYMENT"])
    ndig = pois(4 + 2 * (1 - app) + 2 * (1 - sal))
    X["COUNT_DIGITAL_OPERATIONS"] = dig
    X["COUNT_NON_DIGITAL_OPERATIONS"] = ndig
    X["TOTAL_OPERATIONS"] = dig + ndig
    with np.errstate(invalid="ignore", divide="ignore"):
        X["DIGITAL_RATIO"] = np.where(dig + ndig > 0, dig / (dig + ndig), np.nan)
    X["IN_COUNT_WALLET"] = X["CNT_INBOX_TRANSF_HALYK"] + X["CNT_INBOX_TRANSF_KASPI"] + X["CNT_INBOX_TRANSF_ANOT_BANK"]
    X["OUT_COUNT_WALLET"] = X["CNT_INTER_TRANSFER_HB"] + X["CNT_TRAN_OTHER_BANK"] + X["CNT_TRANSFER_KASPI"]
    wallet_launch = _day("2025-01-01")
    for nm in ("WALLET_TO_INSTALLMENT", "WALLET_TO_DEPOSIT", "WALLET_TO_CREDIT", "WALLET_TO_MARKET",
               "IN_COUNT_WALLET", "OUT_COUNT_WALLET"):
        X[nm] = np.where(t < wallet_launch, np.nan, X[nm])
    # bureau fields missing when no PKB report was pulled
    nopkb = U() < 0.04
    for nm in VER_NUM_COLS:
        if nm.startswith(("PKB_", "SUM_PKB")):
            X[nm] = np.where(nopkb, np.nan, X[nm])
    for nm in ("AMOUNT_SOC_PAYOUT", "AMOUNT_PENSION", "LOAN_BALANCE_KASPI", "USETIME_USER_HB", "AVG_USETIME_DAY"):
        X[nm] = np.where(U() < 0.03, np.nan, X[nm])
    for nm in VER_NUM_COLS:
        if nm.startswith(("AMOUNT_", "VOLUME_", "DEPO_SUM", "DEPO_IN", "DEPO_OUT", "DEPO_REWARD", "DEPO_SALDO",
                          "CRED_DOGOVOR", "CRED_GIVEN", "CRED_REMAIN", "CRED_PAID", "PKB_SUM", "SUM_PKB",
                          "PAYMENT_AMOUNT", "LOAN_BALANCE", "MONTH_LOAN")):
            X[nm] = np.round(X[nm], 2)
    missing = set(VER_NUM_COLS) - set(X)
    assert not missing, missing

    # ---- text columns
    br = W.branch[ci]
    long_v = W.branch_long[ci] ^ (U() < 0.08)
    b_short = np.array([b[0] for b in _BRANCHES], dtype=object)
    b_long = np.array([b[1] for b in _BRANCHES], dtype=object)
    branch = np.where(long_v, b_long[br], b_short[br])
    branch = np.where(U() < 0.02, np.char.add(branch.astype(str), " ").astype(object), branch)
    city_en = np.array([b[3] for b in _BRANCHES], dtype=object)[br]
    cit = np.where(U() < 0.98, np.where(U() < 0.6, "КАЗАХСТАН", "Казахстан"),
                   _pick(rng, ["РОССИЯ", "УЗБЕКИСТАН", "Кыргызстан"], R))
    seg = np.select([age < 25, age < 35, age >= pens_age, W.log_inc[ci] > np.log(900_000)],
                    ["Молодежь1", "Молодежь2", "Пенсионеры", "Премиальный"], "Массовый").astype(object)
    seg_rb = np.select([W.log_inc[ci] > np.log(1_000_000), W.log_inc[ci] > np.log(500_000)],
                       ["AFFLUENT", "MASS_AFFLUENT"], "MASS").astype(object)
    ab = W.abroad_k[ci]
    ab_country = np.array([a[0] for a in _ABROAD], dtype=object)[ab]
    ab_city = np.array([a[1] for a in _ABROAD], dtype=object)[ab]
    dev = np.array(_DEVICES, dtype=object)[W.dev1[ci]]
    has_app = app & (sess > 0)
    mcg_p = np.ones(len(_MCG)); mcg_p[-1] = 0.0
    m1 = rng.choice(len(_MCG) - 1, R)
    m2 = (m1 + 1 + rng.choice(len(_MCG) - 2, R)) % (len(_MCG) - 1)
    m3 = (m2 + 1 + rng.choice(len(_MCG) - 3, R)) % (len(_MCG) - 1)
    gamb = U() < _sig(-4.8 + 0.8 * rf + 0.5 * nearbad)
    mcg = np.array(_MCG, dtype=object)
    mcg3 = np.where(gamb, mcg[-1], mcg[m3])
    txt = {
        "IIN": W.iin[ci],
        "BRANCH": branch,
        "GENDER": np.where(W.male[ci], "M", "F"),
        "CITIZENSHIP": cit,
        "TRUST_PHONE_NUMBER": np.where(U() < 0.9, W.phone[ci], None),
        "FAMILY_STATUS": np.where(U() < 0.7, None, _pick(rng, ["Женат/Замужем", "Холост/Не замужем", "Разведен(а)",
                                                             "Вдовец/Вдова"], R, p=[0.6, 0.28, 0.1, 0.02])),
        "CLIENT_SEGMENT": seg,
        "RESIDENCE_COUNTRY": np.where(U() < 0.98, "KZ", None),
        "CLIENT_SEGMENT_RB": seg_rb,
        "CLIENT_CATEGORY": np.where(W.employee[ci], "Сотрудник банка",
                                    np.where(W.z[ci] > 1.8, "ИП", "Физическое лицо")),
        "EDUCATION": _pick(rng, [None, "Высшее", "Среднее специальное", "Среднее"], R, p=[0.4, 0.35, 0.18, 0.07]),
        "FLG_DEATH": np.where(dead, "1", "0"),
        "FLG_LOAN_OVERDUE_MORE_1": np.where(U() < 0.03, None, np.where(dpd > 1, "1", "0")),
        "MOBILE_OS": np.where(W.app_user[ci], dev[:, 1], None),
        "MOBILE_MODEL": np.where(W.app_user[ci], dev[:, 0], None),
        "MOBILE_BRAND": np.where(W.app_user[ci], dev[:, 2], None),
        "MAU_FIREBASE": np.where(has_app, "1", "0"),
        "MAU_AUTORIZATION": np.where(has_app, "Y", "N"),
        "LAST_EVENT_COUNTRY": np.where(~W.app_user[ci], None, np.where(abroad & (U() < 0.7), ab_country, "Kazakhstan")),
        "LAST_EVENT_CITY": np.where(~W.app_user[ci], None, np.where(abroad & (U() < 0.7), ab_city, city_en)),
        "FLG_BLACK_LIST": np.where(U() < _sig(-5.5 + 1.0 * rp), "1", "0"),
        "EPG_1_0_OR_NEW_2": _pick(rng, ["0", "1", "2"], R, p=[0.5, 0.3, 0.2]),
        "SALARY_CATEGORY_BY_MODEL": np.where(age >= pens_age, "Пенсионер",
                                             np.where(sal, "Зарплатный", "Незарплатный")),
        "TRANSACTION_CLIENT_SEGMENT": np.select([sess > 40, sess > 10, sess > 0], ["Активный", "Транзакционный",
                                                                                 "Пассивный"], "Спящий").astype(object),
        "MCG_TOP1": mcg[m1], "MCG_TOP2": mcg[m2], "MCG_TOP3": mcg3,
        "DIGITAL_SEGMENT": np.select([sess > 30, sess > 5], ["Digital", "Hybrid"], "Offline").astype(object),
        "Full_name": W.fio[ci],
    }
    bl = txt["FLG_BLACK_LIST"] == "1"
    txt["TYPE_BLACKLIST"] = np.where(bl, _pick(rng, ["Мошенничество", "Негативная информация",
                                                     "Просроченная задолженность"], R), None)
    data = {}
    for c in VER_COLS:
        if c == "DATE_VALUE":
            data[c] = _ts(t, unit="us")
        elif c == "DOC_EXPIRY_DATE":
            data[c] = _ts(W.doc_exp[ci], unit="us", nat=W.doc_exp_nat[ci])
        elif c in txt:
            data[c] = _S(txt.pop(c))
        else:
            data[c] = np.asarray(X.pop(c), dtype="float64")
    return pd.DataFrame(data)


def _build_bankrupt(W) -> pd.DataFrame:
    rng = W.rng["bankrupt"]
    n = W.n_cl
    lo = _day("2024-03-26")
    t_ref = np.where(W.bad1 < _BIG, W.bad1, np.where(W.shock_s < _BIG, W.shock_s, _BIG))
    p = _sig(-5.6 + 1.0 * W.z + 1.4 * (t_ref < _BIG))
    who = np.nonzero(rng.random(n) < p)[0]
    nrep = 1 + (rng.random(who.size) < 0.12)
    ci = np.repeat(who, nrep)
    k = ci.size
    base = np.where(t_ref[ci] < _BIG, t_ref[ci] + rng.integers(-40, 300, k), rng.integers(lo, W.cut + 1, k))
    base = base + np.where(np.r_[False, ci[1:] == ci[:-1]], rng.integers(30, 200, k), 0)
    keep = base <= W.cut
    ci, base = ci[keep], base[keep]
    k = ci.size
    pc_day = np.maximum(base, lo)
    pc_sec = rng.integers(8 * 3600, 20 * 3600, k)
    cref = W.c_pref[ci]                     # contract i (i < n_cl) is the first contract of client i
    city = np.array([b[2] for b in _BRANCHES], dtype=object)[W.branch[ci]]
    df = pd.DataFrame({
        "date_value": pd.Series([_dt.date(1970, 1, 1) + _dt.timedelta(days=int(d)) for d in pc_day], dtype=object),
        "gm_system_code": _S(np.full(k, "CMD", dtype=object)),
        "date_change": _ts(pc_day + 1, rng.integers(5 * 3600, 6 * 3600, k), unit="ns"),
        "processid": _S(np.array(["hl" + h for h in _hex(rng, k, 30, upper=False)], dtype=object)),
        "bankruptcytype": _S(np.full(k, "VnesudebnoeBankrotstvo", dtype=object)),
        "statementstatus": _S(np.full(k, "Otkazano", dtype=object)),
        "lastmodificationdate": _ts(pc_day + rng.integers(1, 20, k), rng.integers(0, 86400, k), unit="ns"),
        "processcreate": _ts(pc_day, pc_sec, unit="ns"),
        "iin": _S(W.iin[ci]),
        "clientid": _S(W.abs_id[ci].astype(str)),
        "orderid": _S(_zfill(rng.integers(1, 99_999_999, k), 9)),
        "credit_product_type": _S(_pick(rng, ["Автокредит", "Беззалоговый кредит", "Кредитная карта", None], k,
                                        p=[0.45, 0.3, 0.15, 0.1])),
        "request_amount": np.where(rng.random(k) < 0.7, np.nan, np.round(rng.uniform(5e5, 1.2e7, k), -3)),
        "amount_duration": np.where(rng.random(k) < 0.7, np.nan, rng.choice([12.0, 24.0, 36.0, 60.0, 84.0], k)),
        "fio": _S(W.fio[ci]),
        "mobile": _S(np.char.add("8", np.array([p_[1:] for p_ in W.phone[ci]], dtype=str))),
        "city": _S(city),
        "gender": _S(np.where(W.male[ci], "M", "F")),
        "process_ref": _S(cref, na=rng.random(k) < 0.4),
    })
    return df[BANKRUPT_COLS].sort_values("processcreate").reset_index(drop=True)


def _build_ocrm(W) -> pd.DataFrame:
    rng = W.rng["calls"]
    E = W.E
    w0 = E["start"] + 1
    w1 = np.minimum(E["last"], E["start"] + 75)
    ok = w1 >= w0
    ep = np.nonzero(ok)[0]
    span = (w1 - w0 + 1)[ep]
    n_calls = np.minimum(1 + rng.poisson(0.3 * span), 35)
    n_acts = np.minimum(1 + rng.poisson(span / 12.0), 8)
    cl_e = E["cl"][ep]
    r_e = _risk_at(W, cl_e, E["start"][ep])
    roll_e = E["roll"][ep]
    # ---------------- activities (one CRM task per attempt, belongs to a delinquency episode)
    a_ep = np.repeat(np.arange(ep.size), n_acts)
    A = a_ep.size
    a_day = w0[ep][a_ep] + (rng.random(A) * span[a_ep]).astype(int)
    a_sec = rng.integers(8 * 3600, 19 * 3600, A)
    a_cl = cl_e[a_ep]
    abroad_a = a_day >= W.abroad_from[a_cl]
    p_contact_a = _sig(0.2 - 0.55 * r_e[a_ep] + 0.45 * W.eng[a_cl] - 1.2 * abroad_a - 0.5 * roll_e[a_ep])
    canceled = rng.random(A) < 0.25
    ua = rng.random(A)
    res = np.where(ua < p_contact_a * 0.55, "Обещание получено",
                   np.where(ua < p_contact_a * 0.75, "Оставили информацию 3 лицу",
                            np.where(ua < p_contact_a, "Отказ от оплаты", "Нет ответа"))).astype(object)
    res = np.where(canceled, None, res)
    ptp_a = a_day + rng.integers(1, 8, A)
    dtl = np.full(A, None, dtype=object)
    for i in range(A):
        if res[i] is None:
            continue
        if abroad_a[i] and rng.random() < 0.6:
            dtl[i] = rng.choice(["Оставлено сообщение супруге, заемщик уехал заграницу",
                                 "Со слов матери заемщик уехал за границу на заработки"])
        elif res[i] == "Обещание получено":
            d = _dt.date(1970, 1, 1) + _dt.timedelta(days=int(ptp_a[i]))
            c = _dt.date(1970, 1, 1) + _dt.timedelta(days=int(a_day[i]))
            dtl[i] = (f"PTP {d:%d.%m.%Y} ЕСТЬ ДОЗВОН РОБОТА. Звонок от {c.day}-{c.month}-{c.year} "
                      f"{a_sec[i] // 3600}:{(a_sec[i] % 3600) // 60}")
        elif res[i] == "Оставили информацию 3 лицу" and rng.random() < 0.5:
            dtl[i] = rng.choice(["Оставлено сообщение супруге", "Трубку взял брат, обещал передать",
                                 "Ответила коллега по работе"])
        elif rng.random() < 0.3:
            dtl[i] = rng.choice(["Абонент недоступен", "Сбросил звонок", "Клиент сообщил о задержке заработной платы"])
    a_guid = _guid(rng, A)
    a_off_ep = np.cumsum(n_acts) - n_acts                   # activities of an episode are contiguous
    # client pool (activities ordered by client because episodes are ordered by contract ~ client)
    o_cl = np.argsort(a_cl, kind="stable")
    cl_sorted = a_cl[o_cl]
    cl_first = np.searchsorted(cl_sorted, np.arange(W.n_cl), "left")
    cl_cnt = np.searchsorted(cl_sorted, np.arange(W.n_cl), "right") - cl_first
    campaigns = np.array(["Информативная РБ SC Звонки 4-30 Модель", "SC_ШЫМКЕНТ_Операторы_5",
                          "Информативная РБ SC 6-30", "Сервис SC РБ  Звонки 4-30 Модель"], dtype=object)
    a_camp = campaigns[rng.choice(4, A, p=[0.4, 0.15, 0.25, 0.2])]
    opers = _uniq_ints(rng, 60, 29_000_000, 29_999_999).astype(float)
    a_oper = np.where(rng.random(A) < 0.15, -1.0, opers[rng.integers(0, 60, A)])

    # ---------------- calls
    c_ep = np.repeat(np.arange(ep.size), n_calls)
    C = c_ep.size
    c_day = w0[ep][c_ep] + (rng.random(C) * span[c_ep]).astype(int)
    c_sec = rng.integers(9 * 3600, 20 * 3600, C)
    c_cl = cl_e[c_ep]
    c_con = E["c"][ep][c_ep]
    abroad_c = c_day >= W.abroad_from[c_cl]
    p_contact = _sig(0.1 - 0.55 * r_e[c_ep] + 0.45 * W.eng[c_cl] - 1.2 * abroad_c - 0.5 * roll_e[c_ep])
    contact = rng.random(C) < p_contact
    uc = rng.random(C)
    ptp_share = np.where(roll_e[c_ep], 0.35, 0.5)
    result = np.where(contact,
                      np.where(uc < ptp_share, "PTP",
                               np.where(uc < ptp_share + 0.12, "Отказ",
                                        np.where(uc < ptp_share + 0.27, "Третье лицо",
                                                 np.where(uc < ptp_share + 0.37, "Оплачено", "Перезвонить")))),
                      np.where(uc < 0.85, "Нет ответа", np.where(uc < 0.92, "Абонент недоступен",
                                                                  "Неверный номер"))).astype(object)
    talk = np.where(contact, rng.integers(15, 300, C), 0).astype(float)
    durat = talk + rng.integers(5, 60, C)
    ptp_day = c_day + rng.integers(1, 8, C)
    is_ptp = result == "PTP"
    phone = W.phone[c_cl]
    fmt_u = rng.random(C)
    tnumb_to = np.array([("8" + p[1:]) if u < 0.6 else (p[1:] if u < 0.95 else "7" + p[1:])
                         for p, u in zip(phone, fmt_u)], dtype=object)
    comm = np.empty(C, dtype=object)
    for i in range(C):
        c = _dt.date(1970, 1, 1) + _dt.timedelta(days=int(c_day[i]))
        hh, mi = c_sec[i] // 3600, (c_sec[i] % 3600) // 60
        rr = result[i]
        if rr == "PTP":
            d = _dt.date(1970, 1, 1) + _dt.timedelta(days=int(ptp_day[i]))
            comm[i] = f"PTP {d:%d.%m.%Y} обещ погасит. " + ("ЕСТЬ ДОЗВОН РОБОТ" if rng.random() < 0.5 else
                                                            f"Звонок от {c.day}-{c.month}-{c.year} {hh}:{mi}")
        elif rr in ("Нет ответа", "Абонент недоступен", "Неверный номер"):
            comm[i] = f"{phone[i]}{rr}. Звонок от {c.day}-{c.month}-{c.year} {hh}:{mi}"
        elif rr == "Третье лицо":
            comm[i] = ("Трубку взяла мать, заемщик уехал за границу" if abroad_c[i] else
                       "Ответила супруга, обещала передать информацию")
        elif rr == "Отказ":
            comm[i] = "Клиент отказывается оплачивать, ссылается на отсутствие работы"
        elif rr == "Оплачено":
            comm[i] = "Клиент сообщил что оплатил"
        else:
            comm[i] = "Просил перезвонить позже" if rng.random() < 0.7 else None
    c_guid = _guid(rng, C)
    c_oper = opers[rng.integers(0, 60, C)]
    skill = np.array(["SC_Операторы_2", "SoftCollection_GO_6_30", "SC_Операторы_1", "SoftCollection_GO_4_30"],
                     dtype=object)[rng.choice(4, C, p=[0.35, 0.35, 0.15, 0.15])]
    queue = _pick(rng, [None, "Soft Collection Очередь №2", "Soft Collection Очередь №1"], C, p=[0.5, 0.3, 0.2])
    incoming = rng.random(C) < 0.04

    # ---------------- calls x activities (the real extract repeats each call for several activities)
    k = 1 + np.minimum(rng.poisson(0.9, C), 4)
    pc = np.repeat(np.arange(C), k)
    same = rng.random(pc.size) < 0.8
    e_of = c_ep[pc]
    a_same = a_off_ep[e_of] + (rng.random(pc.size) * n_acts[e_of]).astype(int)
    cli = c_cl[pc]
    a_other = o_cl[cl_first[cli] + (rng.random(pc.size) * cl_cnt[cli]).astype(int)]
    ai = np.where(same, a_same, a_other)
    pair = pd.DataFrame({"pc": pc, "ai": ai}).drop_duplicates()
    pc, ai = pair["pc"].to_numpy(), pair["ai"].to_numpy()
    n = pc.size
    tech = lambda p: rng.random(n) < p  # noqa: E731
    df = pd.DataFrame({
        "date_value1": _S(_dstr(c_day[pc])),
        "date_change1": _ts(c_day[pc] + 1, rng.integers(2 * 3600, 3 * 3600, C)[pc], unit="us"),
        "callid": _S(c_guid[pc]),
        "oper_ocrm_subject_id": c_oper[pc],
        "call_direct": _S(np.where(incoming[pc], "Incoming", "Outgoing")),
        "call_durat": durat[pc],
        "ocrm_subject_id": W.crm_id[c_cl[pc]],
        "tr_oper_ocrm_subject_id": np.full(n, -1.0),
        "tnumb_from": _S(np.array([str(int(x) % 90000 + 10000) for x in c_oper], dtype=object)[pc]),
        "tnumb_to": _S(tnumb_to[pc]),
        "result_predl": _S(result[pc]),
        "call_comment": _S(comm[pc]),
        "createdon_date": _ts(c_day[pc], c_sec[pc], unit="us"),
        "modifiedon_date": _ts(c_day[pc], c_sec[pc] + durat[pc].astype(int) + rng.integers(30, 4000, C)[pc],
                               unit="us"),
        "skill_gr_tsname": _S(skill[pc]),
        "talktime": talk[pc],
        "queue_name": _S(queue[pc]),
        "created_by_ocrm_subject_id": c_oper[pc],
        "num_dbz": _S(W.c_num[c_con[pc]]),
        "tsptp_date": _ts(ptp_day[pc], unit="us", nat=~is_ptp[pc]),
        "consult_theme": _S(np.where(incoming[pc] & tech(0.5), "Погашение кредита", None)),
        "tsnumber": _S(np.where(tech(0.01), _zfill(rng.integers(0, 10**7, n), 7), None)),
        "tstransfernumber": _S_na(n),
        "tsclientmark": np.where(tech(0.005), rng.integers(1, 6, n).astype(float), np.nan),
        "before_conn_time": np.where(tech(0.01), rng.integers(1, 30, n).astype(float), np.nan),
        "tsresultid": _S(np.where(tech(0.01), "{5A1B2C3D-0000-4000-8000-000000000001}", None)),
        "ocrm_caseid": _S_na(n),
        "integrationid": np.full(n, np.nan),
        "tslineid": _S(np.where(tech(0.01), "1", None)),
        "tslanguageid": _S(np.where(tech(0.01), "ru", None)),
        "tsskillgroupid": _S_na(n),
        "isinternetcall_fl": np.where(tech(0.02), 0.0, np.nan),
        "iscallback_fl": np.where(tech(0.02), 0.0, np.nan),
        "tstypeid": _S(np.where(tech(0.01), "1", None)),
        "date_value": _S(_dstr(a_day[ai])),
        "date_change": _ts(a_day[ai] + 1, rng.integers(3600, 4 * 3600, A)[ai], unit="us"),
        "activityid": _S(a_guid[ai]),
        "act_creat_date": _ts(a_day[ai], a_sec[ai] - 1800, unit="us"),
        "title_call_dl": _S(np.full(n, "Выполнить звонок должнику", dtype=object)),
        "activ_sost": _S(np.where(canceled[ai], "Canceled", "Completed")),
        "company_sc": _S(a_camp[ai]),
        "fact_zvr_date": _ts(a_day[ai], a_sec[ai], unit="us", nat=canceled[ai]),
        "client_ocrm_subject_id": W.crm_id[a_cl[ai]],
        "activ_result": _S(res[ai]),
        "activ_result_dtl": _S(dtl[ai]),
        "otv_sotr_ocrm_subject_id": a_oper[ai],
        "tab_num_sotr": _S_na(n),
        "ocrm_tsdebt_id": E["crm_debt_id"][ep][a_ep][ai],
    })
    o = np.lexsort((ai, c_sec[pc], c_day[pc]))
    df = df.iloc[o].reset_index(drop=True)
    return df[OCRM_COLS]


def _k2_frame(**cols) -> pd.DataFrame:
    return pd.DataFrame(cols)[K2_COLS]


def _build_vneshpt_ot(W) -> pd.DataFrame:
    rng = W.rng["k2ot"]
    cut = W.cut
    n = W.n_cl
    lo = _day("2017-01-01")
    p = _sig(-1.6 + 0.9 * W.z)
    who = rng.random(n) < p
    n_base = np.where(who, 1 + rng.poisson(1.8, n), 0)
    shocked = (W.shock_s < cut) & (rng.random(n) < 0.6)
    n_post = shocked * (1 + rng.poisson(1.0, n))          # new enforcement after the hardship shock
    ndoc = n_base + n_post
    ci = np.repeat(np.arange(n), ndoc)
    k = ci.size
    pos = np.arange(k) - np.repeat(np.cumsum(ndoc) - ndoc, ndoc)
    is_post = pos >= n_base[ci]
    u = rng.random(k) ** 0.6                               # history back to 2017, denser recently
    acc = np.where(is_post, W.shock_s[ci] + rng.integers(20, 260, k), (lo + u * (cut - lo)).astype("int64"))
    keep = acc <= cut                                      # post-shock orders after the data cut do not exist yet
    ci, acc = ci[keep], acc[keep]
    k = ci.size
    # extra docs for a few non-portfolio IINs
    n_np = max(1, int(0.05 * k))
    np_iin = np.array([f"{rng.integers(60, 99):02d}{rng.integers(1, 13):02d}{rng.integers(1, 29):02d}"
                       f"{rng.integers(3, 5)}{rng.integers(0, 100000):05d}" for _ in range(n_np)], dtype=object)
    np_acc = (lo + rng.random(n_np) ** 0.6 * (cut - lo)).astype("int64")
    # orders per doc (one per arrested account), same requirement group & amount
    per = 1 + rng.poisson(0.8, k + n_np)
    d_iin = np.concatenate([W.iin[ci], np_iin])
    d_acc = np.concatenate([acc, np_acc])
    d_z = np.concatenate([_risk_at(W, ci, acc), rng.standard_normal(n_np)])
    d_dog = np.concatenate([W.acc_num[ci], np.array([f"0977770000{x:06d}" for x in rng.integers(0, 999999, n_np)],
                                                    dtype=object)])
    d_amt = np.clip(np.round(np.exp(rng.normal(np.log(150_000), 1.3, d_iin.size)), -1), 17_250, 11_022_000)
    d_grp = _uniq_ints(rng, d_iin.size, 40_000_000, 160_000_000)
    p_close = _sig(1.0 - 0.7 * d_z)
    d_closed = rng.random(d_iin.size) < p_close
    d_lag = np.round(rng.exponential(60, d_iin.size)).astype(int) + rng.integers(0, 5, d_iin.size)
    d_close = np.where(d_closed, d_acc + d_lag, _BIG)
    d_close = np.where(d_close > cut, _BIG, d_close)
    ix = np.repeat(np.arange(d_iin.size), per)
    m = ix.size
    sec = rng.integers(8 * 3600, 19 * 3600, m)
    close = d_close[ix]
    is_open = close >= _BIG
    ir4 = rng.random(m) < 0.45
    descr_opts = np.array(['На основании п.2 ст.55 Закона РК "Об исполнительном производстве и статусе судебных '
                           'исполнителей"', "Согласно исполнительного документа Иртышского районного суда",
                           "Согласно исполнительной надписи нотариуса", "Взыскание по исполнительному листу",
                           "Взыскание налоговой задолженности (ДГД)"], dtype=object)
    descr = descr_opts[rng.choice(5, m, p=[0.5, 0.15, 0.15, 0.1, 0.1])]
    last_change = np.where(is_open, d_acc[ix], close)
    begin_old = rng.random(m) < 0.7
    code = np.array([("HSBKMJ" + f"{x:015d}") if u_ < 0.5 else ("HSBKMA" + f"{x % 10**10:010d}")
                     for x, u_ in zip(rng.integers(10**8, 10**9, m), rng.random(m))], dtype=object)
    req = _uniq_ints(rng, m, 10_000_000, 90_000_000)
    df = _k2_frame(
        dwh_id=_uniq_ints(rng, m, 1_000_000_000, 2_000_000_000).astype("int64"),
        source_id=np.full(m, 51, dtype="int64"),
        gm_system_code=_S(np.full(m, "K2", dtype=object)),
        date_change=_ts(last_change + 1, rng.integers(0, 6 * 3600, m), unit="us"),
        dwh_code=_S(req.astype(str)),
        date_begin=np.where(begin_old, np.datetime64("1900-01-01", "us"), _ts(last_change, unit="us")),
        date_end=_const_ts("9999-12-31", m),
        is_actual=_S(np.full(m, "A", dtype=object)),
        req_id=req.astype("int64"),
        code=_S(code),
        longname=_S_na(m),
        doc_num=_S(np.where(rng.random(m) < 0.85,
                            np.char.add(np.char.add(rng.integers(1000, 9999, m).astype(str), "-"),
                                        rng.integers(1000, 9999, m).astype(str)),
                            rng.integers(1, 99, m).astype(str)).astype(object)),
        sender_k2_bank_id=np.where(rng.random(m) < 0.7, 36, -1).astype("int64"),
        recip_k2_bank_id=np.where(rng.random(m) < 0.5, 1, 36).astype("int64"),
        currency_code=_S(np.full(m, "KZT", dtype=object)),
        k2_knp_id=np.where(rng.random(m) < 0.7, 54, 11).astype("int64"),
        k2_abonent_id=np.where(rng.random(m) < 0.6, 10, -1).astype("int64"),
        accept_date=_ts(d_acc[ix], sec, unit="us"),
        opertype_id=np.where(ir4, 224, 223).astype("int64"),
        recip_client_iin=_S(d_iin[ix]),
        close_date=_ts(np.where(is_open, 0, close), rng.integers(9 * 3600, 22 * 3600, m), unit="us", nat=is_open),
        k2_req_status_id=np.where(is_open, rng.choice([10, 11], m), rng.choice([3, 5], m, p=[0.85, 0.15])).astype("int64"),
        bs_gm_opertype_id=np.where(rng.random(m) < 0.5, np.nan, -1.0),
        k2_employee_id=np.where(rng.random(m) < 0.6, np.nan, 7618.0),
        k2_requirement_group_id=d_grp[ix].astype("int64"),
        amount_sum=d_amt[ix].astype("float64"),
        descr=_S(descr),
        opertype_name=_S(np.where(ir4, "ИР4", "ИР3")),
        doc_date=_ts(d_acc[ix] - rng.integers(0, 30, m), unit="us"),
        dogovor=_S(d_dog[ix]),
        iin_bin=_S(d_iin[ix]),
        system_bs=_S(np.full(m, "RB", dtype=object)),
    )
    return df.sort_values("accept_date").reset_index(drop=True)


def _build_vnpt(W) -> pd.DataFrame:
    rng = W.rng["vnpt"]
    cut = W.cut
    E = W.E
    start = _day(_VNPT_START)
    rep, mth, at, dpd = _expand_months(E["start"], E["last"])
    sel = (_month_start(mth) >= start)
    rep, mth, at = rep[sel], mth[sel], at[sel]
    c = E["c"][rep]
    ms = _month_start(mth)
    first_day = np.where(E["start"][rep] >= ms, E["start"][rep],
                         np.minimum(ms + W.c_repday[c], at))
    nd = 1 + rng.poisson(0.7, rep.size)
    ix = np.repeat(np.arange(rep.size), nd)
    j = np.arange(ix.size) - np.repeat(np.cumsum(nd) - nd, nd)
    acc = np.minimum(first_day[ix] + np.where(j == 0, 0, rng.integers(1, 12, ix.size)), at[ix])
    e = rep[ix]
    cc = c[ix]
    cured = (E["end"][e] < _BIG) & (E["end"][e] + 1 <= cut)
    r = _risk_at(W, W.c_cl[cc], acc)
    part = rng.random(ix.size) < _sig(-0.8 - 0.6 * r)
    lag_part = rng.integers(0, 4, ix.size)
    close = np.where(cured, np.maximum(acc, E["end"][e] + 1), _BIG)      # paid when the episode is cured
    close = np.where(part, acc + lag_part, close)
    close = np.where(close > cut, _BIG, close)
    amt = W.c_inst[cc] * np.ceil(np.maximum(dpd[sel][ix], 1) / 30.0) * np.where(part, rng.uniform(0.01, 0.6, ix.size),
                                                                                 rng.uniform(0.6, 1.0, ix.size))
    # same-day demands on due dates for non-delinquent contracts (auto-debit from another own account)
    mlist = np.arange(_month_of(start), _month_of(cut) + 1)
    cm = np.repeat(np.arange(W.n_c), mlist.size)
    mm = np.tile(mlist, W.n_c)
    due = _month_start(mm) + W.c_repday[cm] - 1
    ok = (due >= W.c_open[cm] + 25) & (due <= cut) & (due < W.c_close[cm])
    p_sd = _sig(-2.9 + 0.5 * W.z[W.c_cl[cm]])
    ok &= rng.random(cm.size) < p_sd
    ok &= _dpd_contract_on(W, cm, due) == 0
    cm, due = cm[ok], due[ok]
    acc2 = due
    close2 = due
    amt2 = W.c_inst[cm] * rng.uniform(0.3, 1.0, cm.size)
    cc_all = np.concatenate([cc, cm])
    acc_all = np.concatenate([acc, acc2])
    close_all = np.concatenate([close, close2])
    amt_all = np.round(np.concatenate([amt, amt2]), 2)
    m = cc_all.size
    sec_a = rng.integers(4 * 3600, 7 * 3600, m)
    sec_c = np.where(close_all == acc_all, sec_a + rng.integers(600, 12 * 3600, m), rng.integers(8 * 3600, 23 * 3600, m))
    sec_c = np.minimum(sec_c, 86399)
    is_open = close_all >= _BIG
    cn = W.c_num[cc_all]
    iin = W.iin[W.c_cl[cc_all]]
    last_change = np.where(is_open, acc_all, close_all)
    req = _uniq_ints(rng, m, 100_000_000, 400_000_000)
    df = _k2_frame(
        dwh_id=_uniq_ints(rng, m, 3_000_000_000, 4_000_000_000).astype("int64"),
        source_id=np.full(m, 51, dtype="int64"),
        gm_system_code=_S(np.full(m, "K2", dtype=object)),
        date_change=_ts(last_change + 1, rng.integers(0, 6 * 3600, m), unit="us"),
        dwh_code=_S(req.astype(str)),
        date_begin=np.where(rng.random(m) < 0.6, np.datetime64("1900-01-01", "us"), _ts(last_change, unit="us")),
        date_end=_const_ts("9999-12-31", m),
        is_actual=_S(np.full(m, "A", dtype=object)),
        req_id=req.astype("int64"),
        code=_S_na(m),
        longname=_S_na(m),
        doc_num=_S(cn),
        sender_k2_bank_id=np.full(m, 36, dtype="int64"),
        recip_k2_bank_id=np.full(m, 36, dtype="int64"),
        currency_code=_S(np.full(m, "KZT", dtype=object)),
        k2_knp_id=np.full(m, 14, dtype="int64"),
        k2_abonent_id=np.full(m, 6, dtype="int64"),
        accept_date=_ts(acc_all, sec_a, unit="us"),
        opertype_id=np.full(m, 238, dtype="int64"),
        recip_client_iin=_S(iin),
        close_date=_ts(np.where(is_open, 0, close_all), sec_c, unit="us", nat=is_open),
        k2_req_status_id=np.where(is_open, 10, rng.choice([11, 3], m)).astype("int64"),
        bs_gm_opertype_id=np.full(m, -1.0),
        k2_employee_id=np.where(rng.random(m) < 0.5, np.nan, 7618.0),
        k2_requirement_group_id=_uniq_ints(rng, m, 143_000_000, 160_000_000).astype("int64"),
        amount_sum=amt_all.astype("float64"),
        descr=_S(np.char.add("Погашение задолженности по договору ", cn.astype(str)).astype(object)),
        opertype_name=_S(np.full(m, "ВнПТ", dtype=object)),
        doc_date=_ts(acc_all, unit="us"),
        dogovor=_S(cn),
        iin_bin=_S(iin),
        system_bs=_S(np.full(m, "RB", dtype=object)),
    )
    return df.sort_values("accept_date").reset_index(drop=True)


def _dpd_contract_on(W, c, t):
    E = W.E
    left = pd.DataFrame({"row": np.arange(len(c)), "c": c, "t": t})
    right = pd.DataFrame({"c": E["c"], "s": E["start"], "l": E["last"]})
    mm = left.merge(right, on="c")
    mm = mm[(mm["s"] <= mm["t"]) & (mm["t"] <= mm["l"])]
    out = np.zeros(len(c), dtype="int64")
    if len(mm):
        g = (mm["t"] - mm["s"] + 1).groupby(mm["row"]).max()
        out[g.index.to_numpy()] = g.to_numpy()
    return out


def _build_excl(W) -> pd.DataFrame:
    rng = W.rng["excl"]
    n = W.n_cl
    titles_soc = ["Письмо от УВЕК 02.04.2020", "согласно списка на отсрочку СУСН и не СУСН от 15.04.2020",
                  "упращенка", "Список БизБиргемиз"]
    bad_cl = W.bad1 < _BIG
    p_bk = _sig(-6.6 + 0.8 * W.z + 2.2 * bad_cl)
    bk = rng.random(n) < p_bk
    soc = (rng.random(n) < 0.004) & ~bk
    who = np.nonzero(bk | soc)[0]
    title = np.where(bk[who], "банкротство фл", _pick(rng, titles_soc, who.size, p=[0.35, 0.3, 0.2, 0.15]))
    upd = np.where(bk[who] & bad_cl[who], W.bad1[who] + rng.integers(60, 300, who.size),
                   np.where(np.char.startswith(title.astype(str), "банкрот"), rng.integers(_day("2023-06-01"), W.cut, who.size),
                            _day("2020-04-02") + rng.integers(0, 90, who.size)))
    upd_nat = (rng.random(who.size) < 0.4) | (upd > W.cut)
    iin = W.iin[who]
    name = W.fio[who]
    cid = W.abs_id[who].astype(float)
    # non-portfolio rows (bank-wide list)
    m_np = int(n * 0.015) + 1
    np_iin = np.array([f"{rng.integers(50, 99):02d}{rng.integers(1, 13):02d}{rng.integers(1, 29):02d}"
                       f"{rng.integers(3, 5)}{rng.integers(0, 100000):05d}" for _ in range(m_np)], dtype=object)
    np_name = np.array([f"{s} {f}" for s, f in zip(_pick(rng, _SURN, m_np), _pick(rng, _FIRST_M, m_np))], dtype=object)
    np_title = _pick(rng, titles_soc + ["банкротство фл"], m_np, p=[0.3, 0.3, 0.15, 0.1, 0.15])
    np_upd = rng.integers(_day("2020-04-01"), W.cut, m_np)
    iin = np.concatenate([iin, np_iin]); name = np.concatenate([name, np_name])
    title = np.concatenate([title, np_title]); upd = np.concatenate([upd, np_upd])
    upd_nat = np.concatenate([upd_nat, rng.random(m_np) < 0.5])
    cid = np.concatenate([cid, rng.integers(1_000_000, 9_000_000, m_np).astype(float)])
    active = np.where(rng.random(iin.size) < 0.85, 1.0, 0.0)
    # duplicates: same IIN again, name in another case, sometimes another title
    dup = np.nonzero(rng.random(iin.size) < 0.12)[0]
    ix = np.concatenate([np.arange(iin.size), dup])
    nm = name[ix].copy()
    nm[iin.size:] = np.array([s.title() if rng.random() < 0.7 else s.lower() for s in nm[iin.size:]], dtype=object)
    tt = title[ix].copy()
    swap = rng.random(dup.size) < 0.3
    tt[iin.size:] = np.where(swap, _pick(rng, titles_soc, dup.size), tt[iin.size:])
    m = ix.size
    cid_m = cid[ix]
    cid_m[rng.random(m) < 0.05] = np.nan
    df = pd.DataFrame({
        "k2_exclude_id": (1000 + np.arange(m) * 3 + rng.integers(0, 3, m)).astype("int64"),
        "client_id": cid_m,
        "client_longname": _S(nm),
        "iin_bin": _S(iin[ix]),
        "is_active_fl": np.where(rng.random(m) < 0.03, np.nan, active[ix]),
        "exclude_title": _S(tt),
        "is_actual": _S(np.full(m, "A", dtype=object)),
        "date_begin": _const_ts("1900-01-01", m, unit="ns"),
        "date_end": _const_ts("9999-12-31", m, unit="us"),
        "updated_date": _ts(np.where(upd_nat[ix], 0, upd[ix]), rng.integers(0, 86400, m), unit="ns", nat=upd_nat[ix]),
    })
    df = df.iloc[rng.permutation(m)].reset_index(drop=True)
    return df[EXCL_COLS]


def _build_vneshpt(W) -> pd.DataFrame:
    rng = W.rng["vneshpt"]
    cut = W.cut
    E = W.E
    start = _day(_VNESHPT_START)
    rep, mth, at, dpd = _expand_months(E["start"], E["last"])
    sel = (_month_start(mth) >= start) & (dpd >= 3)
    rep, mth, at, dpd = rep[sel], mth[sel], at[sel], dpd[sel]
    c = E["c"][rep]
    cl = W.c_cl[c]
    nb = np.array([len(W.banks[x]) for x in cl], dtype="int64")
    ix = np.repeat(np.arange(rep.size), nb)
    bank = np.concatenate([W.banks[x] for x in cl]) if ix.size else np.array([], dtype=int)
    dbl = 1 + (rng.random(ix.size) < 0.25)                     # sometimes 2 demands per bank-month
    ix, bank = np.repeat(ix, dbl), np.repeat(bank, dbl)
    k = ix.size
    ms = _month_start(mth[ix])
    s0 = np.maximum(ms, E["start"][rep[ix]] + 2)
    cday = np.minimum(s0 + (rng.random(k) * np.maximum(at[ix] - s0 + 1, 1)).astype(int), at[ix])
    csec = rng.integers(6 * 3600, 20 * 3600, k)
    ref = np.array(["PTPH" + f"{x:012d}" for x in _uniq_ints(rng, k, 10_000_000, 99_999_999)], dtype=object)
    r = _risk_at(W, cl[ix], cday)
    # status rows per demand: accepted (always) + executed / rejected / partially executed
    p_exec = _sig(0.2 - 0.8 * r)
    u = rng.random(k)
    second = np.where(u < p_exec * 0.7, 1, np.where(u < p_exec, 2, np.where(u < 0.85, 3, 0)))
    st = [("Принята банком", "PC1", "01 - Документ принят банком", 3),
          ("Исполнена", "PC2", "02 - Документ исполнен", 4),
          ("Частично исполнена", "PC3", "03 - Частичное исполнение", 5),
          ("Отклонена банком", "RJ1", "04 - Недостаточно средств на счете", 6)]
    nrow = 1 + (second > 0)
    jx = np.repeat(np.arange(k), nrow)
    pos = np.arange(jx.size) - np.repeat(np.cumsum(nrow) - nrow, nrow)
    sidx = np.where(pos == 0, 0, second[jx])
    stt = np.array(st, dtype=object)
    chg_day = cday[jx] + np.where(pos == 0, 0, rng.integers(0, 4, jx.size))
    ok = chg_day <= cut
    jx, pos, sidx, chg_day = jx[ok], pos[ok], sidx[ok], chg_day[ok]
    m = jx.size
    cc = c[ix][jx]
    loan_id = (400_000 + cc * 7 + 3).astype("int64")
    ext = W.c_num[cc].copy()
    oth = rng.random(m) < 0.05
    ext[oth] = np.array([f"229918100{x:07d}" for x in rng.integers(0, 10**7, oth.sum())], dtype=object)
    debt = np.round(W.c_inst[cc] * np.ceil(dpd[ix][jx] / 30.0) * rng.uniform(0.9, 1.1, m), 2)
    debt = np.where(rng.random(m) < 0.4, np.nan, debt)
    has_db = rng.random(m) < 0.3
    db = chg_day + rng.integers(0, 3, m)
    banks = np.array(_OTHER_BANKS, dtype=object)[bank[jx]]
    df = pd.DataFrame({
        "created_date": _ts(cday[jx], csec[jx], unit="us"),
        "date_change": _ts(chg_day, csec[jx] + pos * 3600 + 60, unit="us"),
        "РЕФЕРЕНС": _S(ref[jx]),
        "k2_loan_id": loan_id,
        "k2_mt_reason_id": np.full(m, 20, dtype="int64"),
        "k2_abonent_id": banks[:, 1].astype("int64"),
        "k2_extreq_stat_id": stt[sidx, 3].astype("int64"),
        "longname": _S(banks[:, 0]),
        "status_name": _S(stt[sidx, 0]),
        "status_code": _S(stt[sidx, 1]),
        "ext_dea_code": _S(ext),
        "ИИН": _S(W.iin[W.c_cl[cc]]),
        "Тип бизнеса": _S(np.full(m, "RB", dtype=object)),
        "Причина": _S(stt[sidx, 2]),
        "dwh_id": loan_id.astype("float64"),
        "contract_number": _S(W.c_num[cc], na=rng.random(m) > 0.03),
        "code": _S(np.where(rng.random(m) < 0.02, "K2EXT", None)),
        "debt": debt,
        "date_begin": _ts(db, unit="us", nat=~has_db),
        "date_end": _ts(db, unit="us", nat=~has_db),
    })
    # heavy exact duplication (x2..x5)
    rep_n = rng.integers(2, 6, m)
    df = df.iloc[np.repeat(np.arange(m), rep_n)]
    df = df.sort_values(["created_date", "РЕФЕРЕНС", "date_change"], kind="stable").reset_index(drop=True)
    return df[VNESHPT_COLS]


def _build_geo(W) -> pd.DataFrame:
    rng = W.rng["geo"]
    cut = W.cut
    g0 = _day(_GEO_START)
    if cut < g0:                                            # app-login history starts later than the cut
        e = np.array([], dtype=object)
        return pd.DataFrame({c: (np.array([], "float64") if c in ("hb_authorized_fl", "latitude", "longitude") else
                                 np.array([], "datetime64[ns]") if c in ("date_change", "event_datetime") else
                                 _S(e)) for c in GEO_COLS})
    days = np.arange(g0, cut + 1)
    nd = days.size
    n = W.n_cl
    base = W.app_user * np.exp(np.log(0.3) + 0.45 * W.eng + 0.3 * rng.standard_normal(n))
    T = days[None, :]
    lam = np.repeat(base[:, None], nd, axis=1).astype("float32")
    sh = (T >= W.shock_s[:, None]) & (T < (W.shock_s + W.shock_d)[:, None])
    lam = lam * np.where(sh, np.float32(0.67), np.float32(1.0))
    nb = (((T >= W.gbad1[:, None] - 75) & (T <= W.gbad1[:, None] + 150))
          | ((T >= W.gbad2[:, None] - 75) & (T <= W.gbad2[:, None] + 150)))
    lam = lam * np.where(nb, np.float32(0.35), np.float32(1.0))
    lam = np.where(T >= W.death[:, None], np.float32(0.0), lam)
    tot = float(lam.sum(dtype="float64"))
    if tot > _GEO_MAX_ROWS:                                 # keep geo <= ~600k rows for any n_contracts
        lam = lam * np.float32(_GEO_MAX_ROWS / tot)
    cnt = rng.poisson(lam)
    ci, di = np.nonzero(cnt)
    rep = cnt[ci, di]
    ci = np.repeat(ci, rep); dd = days[np.repeat(di, rep)]
    m = ci.size
    night = rng.random(m) < 0.05 + 0.03 * (W.z[ci] > 1)
    hrs = np.where(night, rng.uniform(0, 6, m), np.clip(rng.normal(14, 4, m), 6, 23.99))
    sec = (hrs * 3600).astype("int64")
    o = np.lexsort((sec, dd))
    ci, dd, sec = ci[o], dd[o], sec[o]
    use2 = (W.dev2[ci] >= 0) & (rng.random(m) < 0.25)
    devi = np.where(use2, W.dev2[ci], W.dev1[ci])
    devs = np.array(_DEVICES, dtype=object)
    dname, dos = devs[devi, 0], devs[devi, 1]
    ios = dos == "iOS"
    dev_id_cl = _hex(rng, 2 * n, 16, upper=False)
    dev_id = np.array(dev_id_cl, dtype=object)[np.where(use2, n + ci, ci)]
    osv = np.where(ios, np.array(_IOS_VER, dtype=object)[rng.integers(0, len(_IOS_VER), 2 * n)][np.where(use2, n + ci, ci)],
                   np.array(_AND_VER, dtype=object)[rng.integers(0, len(_AND_VER), 2 * n)][np.where(use2, n + ci, ci)])
    action = np.where(ios, np.where(rng.random(m) < 0.65, "face_id", "pin"),
                      np.where(rng.random(m) < 0.45, "touch_id", "pin"))
    vstep = (nd + len(_APP_VER) - 1) // len(_APP_VER)
    lag = rng.integers(0, 3, n)[ci]
    vi = np.clip((dd - g0) // vstep - lag, 0, len(_APP_VER) - 1)
    appv = np.array(_APP_VER, dtype=object)[vi]
    pool = np.array(["149.27", "37.150", "95.56", "2.134", "178.89", "92.46", "5.34"])
    ipc = np.array([f"{pool[rng.integers(0, len(pool))]}.{rng.integers(0, 256)}.{rng.integers(1, 255)}"
                    for _ in range(3 * n)], dtype=object)
    ip = ipc[ci * 3 + rng.integers(0, 3, m)]
    abroad = (dd >= W.abroad_from[ci]) & (rng.random(m) < 0.75)
    ab = W.abroad_k[ci]
    abr = np.array(_ABROAD, dtype=object)
    ab_i = np.nonzero(abroad)[0]
    ip = ip.copy()
    ip[ab_i] = np.array([f"85.105.{x}.{y}" for x, y in zip(rng.integers(0, 256, ab_i.size),
                                                           rng.integers(1, 255, ab_i.size))], dtype=object)
    lat0 = np.array([b[4] for b in _BRANCHES])[W.branch[ci]]
    lon0 = np.array([b[5] for b in _BRANCHES])[W.branch[ci]]
    lat = np.where(abroad, abr[ab, 2].astype(float), lat0) + rng.normal(0, 0.03, m)
    lon = np.where(abroad, abr[ab, 3].astype(float), lon0) + rng.normal(0, 0.04, m)
    miss = rng.random(m) < 0.5
    lat = np.round(np.where(miss, 0.0, lat), 6)
    lon = np.round(np.where(miss, 0.0, lon), 6)
    login = _guid(rng, n, braces=False, upper=False)
    df = pd.DataFrame({
        "date_value": _S(_dstr(dd)),
        "date_change": _ts(dd, sec + rng.integers(3600, 30 * 3600, m), unit="ns"),
        "public_id": _S(_guid(rng, m, braces=False, upper=False)),
        "login": _S(login[ci]),
        "event_datetime": _ts(dd, sec, unit="ns"),
        "app_name": _S(np.where(ios, "hb_ios", "hb_android")),
        "action": _S(action),
        "hb_authorized_fl": np.where(rng.random(m) < 0.3, 1.0, 0.0),
        "response_message": _S(np.where(rng.random(m) < 0.003, "Неверный ПИН-код", None)),
        "ip_address": _S(ip),
        "latitude": lat.astype("float64"),
        "longitude": lon.astype("float64"),
        "client_iin": _S(W.iin[ci]),
        "device_id": _S(dev_id),
        "device_name": _S(dname),
        "operationsystem": _S(dos),
        "operationsystemversion": _S(osv),
        "app_version_name": _S(appv),
        "device_name_new": _S(np.where(ios, "iPhone", dname)),
    })
    df.index = pd.RangeIndex(891_919, 891_919 + m)
    return df[GEO_COLS]


def _dm_frame(rng, iin, fio, phone, city, male, birth, t_day, z_at, kind, process_ref, product, req_amt,
              dur, partner, approved) -> pd.DataFrame:
    n = iin.size
    sec = rng.integers(8 * 3600, 22 * 3600, n)
    pc = pd.Series(_ts(t_day, sec, unit="s")).dt.strftime("%Y-%m-%d %H:%M:%S").to_numpy(dtype=object)
    lm = pd.Series(_ts(t_day + rng.integers(0, 3, n), sec + 600, unit="s")).dt.strftime(
        "%Y-%m-%d %H:%M:%S.%f").to_numpy(dtype=object)
    age = np.floor((t_day - birth) / 365.25)
    inc = np.exp(np.log(420_000) - 0.22 * z_at + 0.5 * rng.standard_normal(n))
    pkb_loans = rng.poisson(np.clip(1.0 + 0.5 * z_at, 0.1, None))
    dti = np.clip(0.25 + 0.05 * z_at + 0.12 * rng.standard_normal(n), 0.02, 0.95)
    app_score = np.round(620 - 30 * z_at + 60 * rng.standard_normal(n))
    beh = np.round(600 - 20 * z_at + 70 * rng.standard_normal(n))
    pkb_score = np.round(np.clip(700 - 30 * z_at + 80 * rng.standard_normal(n), 300, 999))
    X = {c: np.full(n, None, dtype=object) for c in DM_COLS}
    X["date_value"] = _dstr(t_day).astype(object)
    X["gm_system_code"] = np.full(n, "CMD", dtype=object)
    X["date_change"] = pd.Series(_ts(t_day + 1, rng.integers(3600, 5 * 3600, n), unit="s")).dt.strftime(
        "%Y-%m-%d %H:%M:%S").to_numpy(dtype=object)
    X["processid"] = np.array(["hl" + h for h in _hex(rng, n, 30, upper=False)], dtype=object)
    X["processcreate"] = pc
    X["lastmodificationdate"] = lm
    X["iin"] = iin
    X["clientid"] = rng.integers(1_000_000, 9_000_000, n).astype(str).astype(object)
    X["orderid"] = _zfill(rng.integers(1, 99_999_999, n), 9).astype(object)
    X["credit_product_type"] = product
    X["fio"] = fio
    X["mobile"] = phone
    X["city"] = city
    X["gender"] = np.where(male, "M", "F").astype(object)
    X["process_ref"] = process_ref
    X["monthly_income"] = _fmt_num(np.round(inc, -2), "%.1f", na=rng.random(n) < 0.05)
    X["interest_rate"] = _fmt_num(rng.choice([4.0, 14.0, 17.0, 19.9, 24.5, 34.9], n), "%.1f")
    X["approved_amount"] = _fmt_num(np.where(approved, req_amt, 0.0), "%.2f", na=~approved)
    X["approved_duration"] = _fmt_num(dur, "%.0f", na=~approved)
    X["monthly_payment"] = _fmt_num(req_amt / np.maximum(dur, 1) * 1.15, "%.2f")
    car = product == "Автокредит"
    X["down_payment"] = _fmt_num(req_amt * rng.uniform(0.1, 0.4, n), "%.2f", na=~car)
    X["car_price"] = _fmt_num(req_amt / rng.uniform(0.6, 0.9, n), "%.2f", na=~car)
    X["car_year"] = _fmt_num(rng.choice([2023, 2024, 2025, 2026], n), "%.0f", na=~car)
    X["work_experience_months"] = _fmt_num(np.clip(rng.gamma(2, 30, n) - 6 * z_at, 1, 480), "%.0f",
                                           na=rng.random(n) < 0.08)
    X["dependents_cnt"] = _fmt_num(rng.poisson(1.2, n), "%.0f")
    X["age"] = _fmt_num(age, "%.0f")
    X["pkb_score"] = _fmt_num(pkb_score, "%.0f", na=rng.random(n) < 0.04)
    X["pkb_active_loans_cnt"] = _fmt_num(pkb_loans, "%.0f")
    X["pkb_active_loans_sum"] = _fmt_num(pkb_loans * np.exp(rng.normal(np.log(900_000), 0.8, n)), "%.2f")
    X["pkb_overdue_sum"] = _fmt_num(np.where(rng.random(n) < _sig(-2.5 + 0.9 * z_at),
                                             np.exp(rng.normal(np.log(50_000), 1, n)), 0.0), "%.2f")
    X["pkb_max_dpd"] = _fmt_num(np.where(rng.random(n) < _sig(-1.8 + 0.8 * z_at), rng.integers(1, 120, n), 0),
                                "%.0f")
    X["pkb_req_30d"] = _fmt_num(rng.poisson(np.exp(0.2 + 0.4 * z_at)), "%.0f")
    X["gcvp_avg_income"] = _fmt_num(np.round(inc * rng.uniform(0.6, 1.0, n), 2), "%.2f", na=rng.random(n) < 0.15)
    X["gcvp_months_cnt"] = _fmt_num(rng.integers(0, 13, n), "%.0f", na=rng.random(n) < 0.15)
    X["dti"] = _fmt_num(dti, "%.4f")
    X["pti"] = _fmt_num(np.clip(dti * rng.uniform(0.5, 0.9, n), 0, 1), "%.4f")
    X["app_score"] = _fmt_num(app_score, "%.0f")
    X["behavior_score"] = _fmt_num(beh, "%.0f", na=rng.random(n) < 0.3)
    X["fraud_score"] = _fmt_num(rng.uniform(0, 1, n), "%.3f")
    X["employer_bin"] = _zfill(rng.integers(0, 10**12, n), 12).astype(object)
    X["employer_bin"][rng.random(n) < 0.2] = None
    X["channel"] = _pick(rng, ["Дилер", "Homebank", "Отделение"], n, p=[0.6, 0.3, 0.1])
    X["partner_name"] = partner
    X["product_code"] = np.where(car, "16158", _pick(rng, ["10101", "12001", "13550"], n)).astype(object)
    X["product_name"] = product
    X["currency"] = np.full(n, "KZT", dtype=object)
    X["car_brand"] = np.where(car, _pick(rng, ["Hyundai", "Kia", "Toyota", "Chery", "Haval", "JAC", "Lada"], n), None)
    X["car_model"] = np.where(car, _pick(rng, ["Tucson", "Elantra", "Sportage", "Camry", "Tiggo 7", "Jolion",
                                              "Granta"], n), None)
    X["car_condition"] = np.where(car, _pick(rng, ["Новый", "С пробегом"], n, p=[0.85, 0.15]), None)
    X["income_source"] = _pick(rng, ["ГЦВП", "Справка с работы", "Выписка", None], n, p=[0.6, 0.2, 0.1, 0.1])
    X["employment_type"] = _pick(rng, ["Наемный работник", "ИП", "Госслужащий", "Пенсионер"], n,
                                 p=[0.7, 0.12, 0.13, 0.05])
    X["position_category"] = _pick(rng, ["Специалист", "Руководитель", "Рабочий", None], n, p=[0.5, 0.15, 0.25, 0.1])
    X["family_status"] = _pick(rng, ["Женат/Замужем", "Холост/Не замужем", "Разведен(а)"], n, p=[0.6, 0.3, 0.1])
    X["education"] = _pick(rng, ["Высшее", "Среднее специальное", "Среднее"], n, p=[0.55, 0.3, 0.15])
    X["birth_date"] = _dstr(birth).astype(object)
    X["doc_type"] = np.full(n, "Удостоверение личности", dtype=object)
    X["doc_issue_date"] = _dstr(t_day - rng.integers(100, 3600, n)).astype(object)
    X["doc_expiry_date"] = _dstr(t_day + rng.integers(100, 3600, n)).astype(object)
    X["reg_region"] = city
    X["fact_region"] = np.where(rng.random(n) < 0.9, city, _pick(rng, [b[2] for b in _BRANCHES], n))
    X["phone_verified"] = _pick(rng, ["1", "0", "true"], n, p=[0.8, 0.1, 0.1])
    X["pkb_report_id"] = _zfill(rng.integers(0, 10**10, n), 10).astype(object)
    X["gcvp_last_payment_date"] = _dstr(t_day - rng.integers(5, 60, n)).astype(object)
    X["risk_grade"] = np.select([app_score > 680, app_score > 600, app_score > 540], ["A", "B", "C"], "D").astype(object)
    if kind == "approved":
        X["decision_reason"] = np.where(rng.random(n) < 0.8, "Одобрено", None).astype(object)
        X["decision_code"] = np.full(n, "APPROVE", dtype=object)
        X["status_name"] = np.full(n, "Одобрено", dtype=object)
    else:
        X["decision_reason"] = _pick(rng, ["Высокая долговая нагрузка", "Негативная кредитная история",
                                           "Недостаточный доход", "Стоп-фактор: текущая просрочка", "Отказ клиента"],
                                     n, p=[0.35, 0.25, 0.2, 0.15, 0.05])
        X["decision_code"] = _pick(rng, ["DECLINE_DTI", "DECLINE_PKB", "DECLINE_INCOME", "DECLINE_STOP"], n)
        X["status_name"] = np.full(n, "Отказано", dtype=object)
    X["underwriter_id"] = np.where(rng.random(n) < 0.3, rng.integers(1000, 1100, n).astype(str), None).astype(object)
    X["stop_factor_flag"] = np.where(rng.random(n) < _sig(-3 + z_at), "1", "0").astype(object)
    X["blacklist_flag"] = np.where(rng.random(n) < 0.01, "Y", "N").astype(object)
    X["is_salary_client"] = np.where(rng.random(n) < 0.45, "1", "0").astype(object)
    X["is_halyk_employee"] = np.where(rng.random(n) < 0.01, "1", "0").astype(object)
    X["insurance_flag"] = _pick(rng, ["1", "0", None], n, p=[0.6, 0.35, 0.05])
    X["kasko_flag"] = np.where(car, _pick(rng, ["1", "0"], n, p=[0.7, 0.3]), None)
    X["branch_code"] = _pick(rng, [b[6] for b in _BRANCHES], n)
    X["manager_id"] = rng.integers(10_000, 99_999, n).astype(str).astype(object)
    X["dealer_bin"] = np.where(car, _zfill(rng.integers(0, 10**12, 8), 12)[rng.integers(0, 8, n)], None).astype(object)
    X["sms_consent"] = _pick(rng, ["Да", "Нет"], n, p=[0.9, 0.1])
    X["pep_flag"] = np.full(n, "0", dtype=object)
    X["comment_text"] = np.where(rng.random(n) < 0.1, "Клиент предоставил доп. документы", None).astype(object)
    fill = [c for c in DM_COLS if c.startswith("attr_")]
    for j, c in enumerate(fill):
        mode = j % 7
        if mode == 0:      # numbers-as-strings, >=90 % parsable, pure noise
            X[c] = _fmt_num(rng.normal(100, 30, n), "%.3f", na=rng.random(n) < 0.03)
        elif mode == 1:    # mostly numeric with junk tokens (~60 % parsable)
            v = _fmt_num(rng.integers(0, 50, n), "%.0f")
            v[rng.random(n) < 0.4] = "н/д"
            X[c] = v
        elif mode in (2, 3):
            X[c] = _pick(rng, ["A", "B", "C", None], n, p=[0.4, 0.3, 0.1, 0.2])
        elif mode == 4:
            X[c] = _pick(rng, ["Y", "N"], n, p=[0.2, 0.8])
        else:              # (almost) empty technical columns
            X[c] = np.where(rng.random(n) < 0.01, "x", None).astype(object)
    data = {c: _S(X[c]) for c in DM_COLS if c not in ("request_amount", "amount_duration", "response_result")}
    data["request_amount"] = np.asarray(req_amt, dtype="float64")
    data["amount_duration"] = np.asarray(dur, dtype="float64")
    data["response_result"] = (np.ones(n) if kind == "approved" else rng.choice([0.0, 2.0], n, p=[0.85, 0.15]))
    return pd.DataFrame(data)[DM_COLS]


def _build_dm(W):
    rng = W.rng["dm"]
    cut = W.cut
    city_ru = np.array([b[2] for b in _BRANCHES], dtype=object)
    pn = np.array([p[0] for p in _PARTNERS], dtype=object)

    def mobile(ci):
        return np.char.add("8", np.array([p[1:] for p in W.phone[ci]], dtype=str)).astype(object)

    # ---- approved: the auto loan application of each contract
    c = np.arange(W.n_c)
    ci = W.c_cl[c]
    t1 = W.c_open - rng.integers(1, 21, W.n_c)
    req1 = np.round(W.c_amount * rng.uniform(1.0, 1.12, W.n_c), -3)
    # ---- approved: other products of portfolio clients
    k2 = rng.poisson(0.6, W.n_cl)
    ci2 = np.repeat(np.arange(W.n_cl), k2)
    t2 = rng.integers(_day("2024-01-01"), cut + 1, ci2.size)
    prod2 = _pick(rng, ["Беззалоговый кредит", "Кредитная карта", "Товарный кредит"], ci2.size, p=[0.5, 0.3, 0.2])
    req2 = np.round(np.exp(rng.normal(np.log(1_200_000), 0.8, ci2.size)), -3)
    dur2 = rng.choice([12.0, 24.0, 36.0, 48.0, 60.0], ci2.size)
    pref2 = np.array([f"DGA{x:012d}" for x in _uniq_ints(rng, ci2.size, 500_000_000, 600_000_000)], dtype=object)
    ci_a = np.concatenate([ci, ci2])
    t_a = np.concatenate([t1, t2])
    z_a = _risk_at(W, ci_a, t_a)
    appr = _dm_frame(
        rng, W.iin[ci_a], W.fio[ci_a], mobile(ci_a), city_ru[W.branch[ci_a]], W.male[ci_a], W.birth[ci_a], t_a,
        z_a, "approved", np.concatenate([W.c_pref, pref2]),
        np.concatenate([np.full(W.n_c, "Автокредит", dtype=object), prod2]),
        np.concatenate([req1, req2]), np.concatenate([np.full(W.n_c, 84.0), dur2]),
        np.concatenate([pn[W.c_partner], np.full(ci2.size, None, dtype=object)]), np.ones(ci_a.size, bool))
    # ---- non-portfolio approved/rejected applications (other clients of the bank)
    def outsiders(m):
        iin = np.array([f"{rng.integers(60, 99):02d}{rng.integers(1, 13):02d}{rng.integers(1, 29):02d}"
                        f"{rng.integers(3, 5)}{rng.integers(0, 100000):05d}" for _ in range(m)], dtype=object)
        male = rng.random(m) < 0.6
        fio = np.array([f"{s} {f}" for s, f in zip(_pick(rng, _SURN, m), _pick(rng, _FIRST_M, m))], dtype=object)
        ph = np.char.add("8707", _zfill(rng.integers(0, 10**7, m), 7)).astype(object)
        city = _pick(rng, [b[2] for b in _BRANCHES], m)
        birth = rng.integers(_day("1960-01-01"), _day("2004-01-01"), m)
        t = rng.integers(_day("2024-04-01"), cut + 1, m)
        return iin, fio, ph, city, male, birth, t, rng.standard_normal(m)

    m_out = int(0.3 * W.n_c)
    o = outsiders(m_out)
    out_a = _dm_frame(rng, *o, "approved",
                      np.array([f"DGA{x:012d}" for x in _uniq_ints(rng, m_out, 600_000_000, 700_000_000)], dtype=object),
                      _pick(rng, ["Беззалоговый кредит", "Кредитная карта", "Автокредит"], m_out),
                      np.round(np.exp(rng.normal(np.log(1_500_000), 0.9, m_out)), -3),
                      rng.choice([12.0, 24.0, 36.0, 60.0, 84.0], m_out), np.full(m_out, None, dtype=object),
                      np.ones(m_out, bool))
    dm = pd.concat([appr, out_a], ignore_index=True)
    dm = dm.iloc[rng.permutation(len(dm))].reset_index(drop=True)

    # ---- rejected: credit hunger, more for risky clients and before bad episodes
    months = np.arange(_month_of(_day("2024-04-01")), _month_of(cut) + 1)
    cm = np.repeat(np.arange(W.n_cl), months.size)
    mm = np.tile(months, W.n_cl)
    tt = _month_start(mm) + 14
    lam = np.exp(-3.5 + 0.5 * _risk_at(W, cm, tt) + 0.6 * W.nu[2][cm] + 0.9 * _near_bad(W, cm, tt, before=150, after=30))
    lam = np.where(tt >= W.death[cm], 0.0, lam)
    k = rng.poisson(lam)
    ix = np.repeat(np.arange(cm.size), k)
    rc = cm[ix]
    rt = np.minimum(_month_start(mm[ix]) + rng.integers(0, 28, ix.size), cut)
    prod = _pick(rng, ["Беззалоговый кредит", "Кредитная карта", "Товарный кредит", "Автокредит"], rc.size,
                 p=[0.55, 0.25, 0.15, 0.05])
    req = np.round(np.exp(rng.normal(np.log(1_000_000), 0.8, rc.size)), -3)
    rej = _dm_frame(rng, W.iin[rc], W.fio[rc], mobile(rc), city_ru[W.branch[rc]], W.male[rc], W.birth[rc], rt,
                    _risk_at(W, rc, rt), "rejected",
                    np.array([f"DGA{x:012d}" for x in _uniq_ints(rng, rc.size, 700_000_000, 800_000_000)], dtype=object),
                    prod, req, rng.choice([12.0, 24.0, 36.0, 60.0], rc.size), np.full(rc.size, None, dtype=object),
                    np.zeros(rc.size, bool))
    m_out2 = int(1.2 * W.n_c)
    o = outsiders(m_out2)
    out_r = _dm_frame(rng, *o, "rejected",
                      np.array([f"DGA{x:012d}" for x in _uniq_ints(rng, m_out2, 800_000_000, 900_000_000)], dtype=object),
                      _pick(rng, ["Беззалоговый кредит", "Кредитная карта", "Автокредит"], m_out2),
                      np.round(np.exp(rng.normal(np.log(1_500_000), 0.9, m_out2)), -3),
                      rng.choice([12.0, 24.0, 36.0, 60.0], m_out2), np.full(m_out2, None, dtype=object),
                      np.zeros(m_out2, bool))
    dm_rej = pd.concat([rej, out_r], ignore_index=True)
    dm_rej = dm_rej.iloc[rng.permutation(len(dm_rej))].reset_index(drop=True)
    return dm, dm_rej


def _build_pkb(W) -> pd.DataFrame:
    rng = W.rng["pkb"]
    cut = W.cut
    w0 = _day(_PKB_START)
    as_of = (cut + 1) * 86400
    t0 = w0 * 86400 + 17 * 3600 + 2 * 60 + 46
    if as_of <= t0:
        return pd.DataFrame({"iin": _S_na(0), "rule_name": _S_na(0), **{c: np.array([], "int64") for c in
                             ("cnt_7d", "cnt_30d", "cnt_60d", "cnt_all")},
                             "first_ts": np.array([], "datetime64[us]"), "last_ts": np.array([], "datetime64[us]")})
    n = W.n_cl
    ci = np.arange(n)
    r = _risk_at(W, ci, np.full(n, cut))
    dl = _dpd_at(W, ci, np.full(n, cut)) > 0
    E = W.E
    had = np.zeros(n, bool)
    had[E["cl"][(E["last"] >= w0)]] = True
    cured = np.zeros(n, bool)
    cured[E["cl"][(E["end"] < cut) & (E["end"] >= w0)]] = True
    alive = W.death > w0
    young_m = W.male & ((cut - W.birth) / 365.25 < 27)
    e = lambda b, a=0.0: b * np.exp(a * r)  # noqa: E731
    lam = np.stack([
        e(0.8, 0.6), e(0.05, 0.8), e(1.0, 0.5), e(0.3, 0.9) + 3 * had, e(0.4, 0.4), 0.1 + 1.0 * cured,
        e(0.004, 1.2), np.full(n, 0.01), 0.15 * young_m, e(0.02, 0.6), e(1.5, 0.4), e(0.5, 0.3), e(0.02, 0.7),
        e(0.05, 0.3), e(0.006, 0.8), e(0.01, 1.0), e(0.2, 0.9) + 2 * had, e(0.003, 1.2), np.full(n, 0.3),
        e(0.005, 1.0) + 0.05 * dl, np.full(n, 0.2), np.full(n, 1.2), e(0.3, 0.9) + 4 * had, np.full(n, 2.0),
    ], axis=1) * alive[:, None]
    cnt = rng.poisson(lam)
    cc, rr = np.nonzero(cnt)
    k = cnt[cc, rr]
    ev_c = np.repeat(cc, k); ev_r = np.repeat(rr, k)
    ts = rng.integers(t0, as_of, ev_c.size)
    d = pd.DataFrame({"c": ev_c, "r": ev_r, "ts": ts,
                      "w7": ts >= as_of - 7 * 86400, "w30": ts >= as_of - 30 * 86400, "w60": ts >= as_of - 60 * 86400})
    g = d.groupby(["c", "r"], sort=True).agg(cnt_7d=("w7", "sum"), cnt_30d=("w30", "sum"), cnt_60d=("w60", "sum"),
                                              cnt_all=("ts", "size"), first_ts=("ts", "min"), last_ts=("ts", "max"))
    g = g.reset_index()
    rules = np.array(PKB_RULE_NAMES, dtype=object)
    df = pd.DataFrame({
        "iin": _S(W.iin[g["c"].to_numpy()]),
        "rule_name": _S(rules[g["r"].to_numpy()]),
        "cnt_7d": g["cnt_7d"].to_numpy().astype("int64"),
        "cnt_30d": g["cnt_30d"].to_numpy().astype("int64"),
        "cnt_60d": g["cnt_60d"].to_numpy().astype("int64"),
        "cnt_all": g["cnt_all"].to_numpy().astype("int64"),
        "first_ts": g["first_ts"].to_numpy().astype("datetime64[s]").astype("datetime64[us]"),
        "last_ts": g["last_ts"].to_numpy().astype("datetime64[s]").astype("datetime64[us]"),
    })
    return df[PKB_AGG_COLS]


# --------------------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------------------
def make_demo_tables(n_contracts: int = 4000, seed: int = 42, cut_date: str = "2026-11-11") -> dict:
    """Generate the 13 demo tables with the real schema. Deterministic for a given seed.

    n_contracts : number of auto-loan contracts (clients ~ n_contracts / 1.05; ~1 % duplicated contract rows).
    seed        : random seed (same seed -> identical tables).
    cut_date    : data cut ('YYYY-MM-DD', 2025-07-01 .. 2030-12-31); ovrd's last month is partial up to this day.
    """
    if not _day("2025-07-01") <= _day(cut_date) <= _day("2030-12-31"):
        raise ValueError("cut_date must be between 2025-07-01 and 2030-12-31 (contracts open 2024-04..2026-03)")
    if int(n_contracts) < 50:
        raise ValueError("n_contracts must be >= 50")
    W = _simulate_world(int(n_contracts), int(seed), str(cut_date))
    dm, dm_rej = _build_dm(W)
    tables = {
        "contr": _build_contr(W),
        "ver_df_all_2024_2026": _build_ver(W),
        "bankrupt": _build_bankrupt(W),
        "ocrm_call": _build_ocrm(W),
        "vneshpt_ot": _build_vneshpt_ot(W),
        "vnpt": _build_vnpt(W),
        "excl": _build_excl(W),
        "vneshpt": _build_vneshpt(W),
        "ovrd": _build_ovrd(W),
        "geo": _build_geo(W),
        "dm": dm,
        "dm_rej": dm_rej,
        "hes_pkb_agg": _build_pkb(W),
    }
    return {k: tables[k] for k in TABLE_NAMES}


# --------------------------------------------------------------------------------------
# Self-test
# --------------------------------------------------------------------------------------
def _episode_stats(contr: pd.DataFrame, ovrd: pd.DataFrame, bad_dpd=90, horizon_months=6, gap=5):
    """Notebook-like episode logic -> (ever-delinquent share, n entries, labeled, bad rate)."""
    o = ovrd[ovrd["contract_number"].isin(set(contr["contract_number"]))].copy()
    o["act"] = pd.to_datetime(o["act_date"])
    o["dpd"] = o["ovd_days_cnt"]
    o = o[o["dpd"] > 0]
    o["start"] = o["act"] - pd.to_timedelta(o["dpd"] - 1, unit="D")
    o = o.sort_values(["contract_number", "start"])
    prev = o.groupby("contract_number")["start"].shift()
    o["new"] = prev.isna() | ((o["start"] - prev).dt.days > gap)
    o["ep"] = o["new"].cumsum()
    ep = o.groupby("ep").agg(contract_number=("contract_number", "first"), entry=("start", "min"),
                             max_dpd=("dpd", "max"))
    o["d90"] = o["start"] + pd.to_timedelta(bad_dpd - 1, unit="D")
    b = o.loc[o["dpd"] >= bad_dpd, ["contract_number", "d90"]]
    data_cut = o["act"].max()
    ep["end_h"] = ep["entry"] + pd.DateOffset(months=horizon_months)
    m = ep.reset_index().merge(b, on="contract_number", how="left")
    m["hit"] = (m["d90"] > m["entry"]) & (m["d90"] <= m["end_h"])
    bad = m.groupby("ep")["hit"].any()
    ep["bad"] = bad.reindex(ep.index).fillna(False).astype(int)
    ep["labeled"] = ep["end_h"] <= data_cut
    ever = ep["contract_number"].nunique() / contr["contract_number"].nunique()
    lab = ep[ep["labeled"]]
    return ever, len(ep), len(lab), lab["bad"].mean(), ep


def _self_test(n_contracts=4000, seed=42):
    import resource
    import time
    rss0 = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    t0 = time.time()
    T = make_demo_tables(n_contracts=n_contracts, seed=seed)
    el = time.time() - t0
    rss1 = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    print(f"make_demo_tables(n_contracts={n_contracts}, seed={seed}) -> {el:.1f}s, peak RSS {rss1 / 2**20:.2f} GB "
          f"(+{(rss1 - rss0) / 2**20:.2f} GB during generation)")
    assert el < 60, "generation too slow"
    tot = 0
    print(f"{'table':24s} {'rows':>9s} {'cols':>5s} {'MB':>8s}  dtypes")
    for k, df in T.items():
        mb = df.memory_usage(deep=True).sum() / 2**20
        tot += mb
        dts = df.dtypes.astype(str).value_counts().to_dict()
        print(f"{k:24s} {len(df):9d} {df.shape[1]:5d} {mb:8.1f}  {dts}")
    print(f"total in-memory size: {tot:.0f} MB")
    # ---- schema checks
    exp_cols = {"contr": 13, "ver_df_all_2024_2026": 203, "bankrupt": 19, "ocrm_call": 48, "vneshpt_ot": 32,
                "vnpt": 32, "excl": 10, "vneshpt": 20, "ovrd": 5, "geo": 19, "dm": 159, "dm_rej": 159,
                "hes_pkb_agg": 8}
    for k, nc in exp_cols.items():
        assert T[k].shape[1] == nc, (k, T[k].shape)
    v = T["ver_df_all_2024_2026"]
    is_txt = lambda s: isinstance(s.dtype, pd.StringDtype) or s.dtype == object  # noqa: E731
    n_txt = sum(is_txt(v[c]) for c in v.columns)
    n_dt = sum(pd.api.types.is_datetime64_any_dtype(v[c]) for c in v.columns)
    n_num = sum(pd.api.types.is_numeric_dtype(v[c]) for c in v.columns)
    print(f"ver_df: text={n_txt} date={n_dt} num={n_num}")
    assert (n_txt, n_dt, n_num) == (30, 2, 171)
    oc = T["ocrm_call"]
    print("ocrm_call: text=%d num=%d date=%d" % (sum(is_txt(oc[c]) for c in oc), sum(
        pd.api.types.is_numeric_dtype(oc[c]) for c in oc), sum(pd.api.types.is_datetime64_any_dtype(oc[c]) for c in oc)))
    for k in ("vneshpt_ot", "vnpt"):
        d = T[k]
        print(f"{k}: text={sum(is_txt(d[c]) for c in d)} num={sum(pd.api.types.is_numeric_dtype(d[c]) for c in d)} "
              f"date={sum(pd.api.types.is_datetime64_any_dtype(d[c]) for c in d)}  date_end={d['date_end'].dtype} "
              f"max={d['date_end'].max()}  open(close NaT)={d['close_date'].isna().mean():.1%}  "
              f"accept {d['accept_date'].min():%Y-%m-%d}..{d['accept_date'].max():%Y-%m-%d}")
    dm = T["dm"]
    print("dm numeric cols:", [c for c in dm if pd.api.types.is_numeric_dtype(dm[c])])
    assert [c for c in dm if pd.api.types.is_numeric_dtype(dm[c])] == ["request_amount", "amount_duration",
                                                                        "response_result"]
    assert list(T["vneshpt"].columns[[2, 11, 12, 13]]) == ["РЕФЕРЕНС", "ИИН", "Тип бизнеса", "Причина"]
    ex = T["excl"]
    print("excl dtypes:", {c: str(ex[c].dtype) for c in ("date_begin", "date_end", "updated_date", "is_active_fl")},
          "date_end max", ex["date_end"].max())
    for k, cols in (("ovrd", ("rep_month", "act_date")), ("ocrm_call", ("date_value1", "date_value")),
                    ("geo", ("date_value",)), ("dm", ("processcreate", "date_value"))):
        for c in cols:
            assert isinstance(T[k][c].dtype, pd.StringDtype), (k, c, T[k][c].dtype)
    print("bankrupt.date_value dtype:", T["bankrupt"]["date_value"].dtype, type(T["bankrupt"]["date_value"].iloc[0]))
    # ---- quirks
    c = T["contr"]
    dup = c.duplicated(["process_ref", "contract_number"]).sum()
    print(f"contr: {len(c)} rows, {c['contract_number'].nunique()} contracts, duplicated contracts={dup}, "
          f"IIN leading zero={c['iin'].str.startswith('0').mean():.1%}, clients with 2 contracts="
          f"{(c.drop_duplicates('contract_number').groupby('iin').size() > 1).sum()}, status="
          f"{c['status_name'].value_counts().to_dict()}")
    print("contr.open_date quartiles:", c["open_date"].quantile([0, .25, .5, .75, 1]).dt.strftime("%Y-%m-%d").tolist())
    print("ver_df BRANCH sample:", v["BRANCH"].value_counts().head(6).to_dict())
    print("ver_df CITIZENSHIP:", v["CITIZENSHIP"].value_counts().head(4).to_dict())
    print("ver_df DATE_VALUE quantiles:", v["DATE_VALUE"].quantile([0, .25, .5, .75, 1]).dt.strftime("%Y-%m-%d").tolist())
    print(f"ver_df FLG_DEATH=='1': {(v['FLG_DEATH'] == '1').sum()} rows; MAU_AUTORIZATION "
          f"{v['MAU_AUTORIZATION'].value_counts().to_dict()}")
    print(f"ocrm_call: rows={len(oc)}, distinct calls={oc['callid'].nunique()}, distinct activities="
          f"{oc['activityid'].nunique()}, rows/call={len(oc) / oc['callid'].nunique():.2f}")
    print("ocrm_call result_predl:", oc.drop_duplicates("callid")["result_predl"].value_counts().to_dict())
    vn = T["vneshpt"]
    print(f"vneshpt: rows={len(vn)}, exact-dup share={vn.duplicated().mean():.1%}, "
          f"created {vn['created_date'].min():%Y-%m-%d}..{vn['created_date'].max():%Y-%m-%d}, "
          f"banks={vn['longname'].value_counts().to_dict()}")
    g = T["geo"]
    zero = ((g["latitude"] == 0) & (g["longitude"] == 0)).mean()
    lat, lon = g["latitude"], g["longitude"]
    val = ~((lat == 0) & (lon == 0))
    abroad = val & ~(lat.between(40.5, 55.5) & lon.between(46.4, 87.4))
    print(f"geo: rows={len(g)}, dates {g['date_value'].min()}..{g['date_value'].max()}, zero-coords={zero:.1%}, "
          f"abroad events={abroad.mean():.2%}, clients abroad={g.loc[abroad, 'client_iin'].nunique()}")
    print(f"excl: rows={len(ex)}, dup iin={ex['iin_bin'].duplicated().sum()}, titles="
          f"{ex['exclude_title'].value_counts().to_dict()}")
    print(f"bankrupt: rows={len(T['bankrupt'])}, processcreate {T['bankrupt']['processcreate'].min():%Y-%m-%d}.."
          f"{T['bankrupt']['processcreate'].max():%Y-%m-%d}")
    pk = T["hes_pkb_agg"]
    print(f"hes_pkb_agg: rows={len(pk)}, iins={pk['iin'].nunique()}, rules={pk['rule_name'].nunique()}, "
          f"first_ts min={pk['first_ts'].min()}")
    # ---- delinquency
    ov = T["ovrd"]
    in_c = ov["contract_number"].isin(set(c["contract_number"]))
    print(f"ovrd: rows={len(ov)}, not-in-contr share={1 - in_c.mean():.1%}, months "
          f"{ov['rep_month'].min()}..{ov['rep_month'].max()}, last-month act_date max={ov['act_date'].max()}")
    print("ovrd rows per rep_month (tail):", ov["rep_month"].value_counts().sort_index().tail(8).to_dict())
    ever, n_ep, n_lab, bad, ep = _episode_stats(c.drop_duplicates("contract_number"), ov)
    print(f"ever delinquent={ever:.1%}, entry episodes={n_ep}, labeled={n_lab}, bad90-in-6m among labeled={bad:.1%}, "
          f"self-cure30={(ep['max_dpd'] <= 30).mean():.1%}, max dpd={ep['max_dpd'].max():.0f}")
    by_m = ep[ep["labeled"]].groupby(ep["entry"].dt.to_period("Q"))["bad"].agg(["size", "mean"])
    print("labeled entries / bad rate by quarter:\n", by_m.round(3).to_string())
    # determinism
    T2 = make_demo_tables(n_contracts=min(n_contracts, 600), seed=seed)
    T3 = make_demo_tables(n_contracts=min(n_contracts, 600), seed=seed)
    for k in T2:
        pd.testing.assert_frame_equal(T2[k], T3[k])
    print("determinism check (same seed -> identical tables): OK")
    print("SELF-TEST OK")


if __name__ == "__main__":
    _self_test()
