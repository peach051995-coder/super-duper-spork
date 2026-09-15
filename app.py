# -*- coding: utf-8 -*-
"""
CRM ФОП — Майстер-Дашборд
==========================
Веб-додаток для бухгалтера, який веде облік кількох ФОП-клієнтів.
Базою даних слугує Google Таблиця "CRM_FOP_Master" з аркушами:
    - Клієнти (з параметрами: вартість супроводу, ставки ЄП/ВЗ/ЄСВ, пільга ЄСВ,
      посилання на особисту таблицю ФОПа для синхронізації доходу)
    - Доходи (Дата, ФОП, Клієнт, Сума, Валюта, Комісія, Джерело, Послуга, №документа, ID_транзакції)
    - Контроль_Місяця (нараховано/сплачено по супроводу, ЄСВ, ЄП, ВЗ + звітність)
    - Історія_Груп
    - Історія_Пільг_ЄСВ

Стек: Python + Streamlit (інтерфейс) + gspread (робота з Google Sheets).
Ключі доступу НЕ зашиті в код — вони беруться зі st.secrets
(файл .streamlit/secrets.toml локально, або "Secrets" у Streamlit Cloud).
"""

import streamlit as st
import pandas as pd
from datetime import date, datetime
import re
import hashlib
import gspread
from gspread.utils import rowcol_to_a1, ValueRenderOption
from google.oauth2.service_account import Credentials

import mono_api  # банківський модуль — синхронізація з Monobank

# ============================================================
# 1. НАЛАШТУВАННЯ ПРОГРАМИ
# ============================================================
SPREADSHEET_NAME = "CRM_FOP_Master"
SHEET_CLIENTS = "Клієнти"
SHEET_INCOME = "Доходи"
SHEET_CONTROL = "Контроль_Місяця"
SHEET_GROUP_HISTORY = "Історія_Груп"

GROUP_OPTIONS = ["2 група", "3 група 5%", "3 група 3% + ПДВ", "Загальна система"]

SOURCE_OPTIONS = [
    "Готівка", "ПРРО (готівка)", "Ощадбанк", "Приватбанк",
    "Термінал (Ощадбанк)", "Термінал (Приватбанк)", "Monobank", "Інше",
]
CURRENCY_OPTIONS = ["UAH", "USD", "EUR"]

# Порядок колонок на аркуші "Доходи" — має збігатись із заголовками в Google Таблиці.
INCOME_COLUMNS = [
    "Дата операції", "ПІБ ФОП", "Клієнт", "Сума, грн", "Валюта",
    "Комісія, грн", "Джерело", "Послуга/Коментар", "№ документа", "ID_транзакції",
]

SHEET_ESV_BENEFIT_HISTORY = "Історія_Пільг_ЄСВ"
ESV_BENEFIT_OPTIONS = ["Немає", "Декрет", "Офіційне працевлаштування", "Інвалідність", "Інше"]
EP_RATE_TYPE_OPTIONS = ["Відсоток від доходу", "Фіксована сума"]

# Порядок колонок на аркуші "Контроль_Місяця".
CONTROL_COLUMNS = [
    "Звітний Місяць", "ПІБ ФОП",
    "Супровід нараховано", "Супровід сплачено",
    "ЄСВ нараховано", "ЄСВ сплачено",
    "Єдиний податок нараховано", "Єдиний податок сплачено",
    "ВЗ нараховано", "ВЗ сплачено",
    "Звітність подано",
]

# Значення за замовчуванням для параметрів ФОПа, якщо на аркуші "Клієнти"
# ще немає відповідних колонок або клітинка порожня.
CLIENT_PARAM_DEFAULTS = {
    "Вартість супроводу, грн/міс": 0.0,
    "Тип ставки ЄП": "Фіксована сума",
    "Ставка ЄП": 0.0,
    "Тип ставки ВЗ": "Фіксована сума",
    "Ставка ВЗ": 0.0,
    "Ставка ЄСВ, грн/міс": 0.0,
    "Пільга ЄСВ (поточна)": "Немає",
    "Посилання на таблицю ФОПа": "",
    "Аркуш операцій": "",
    "Виключити фрази (через кому)": "",
    "Виключити клієнтів (через кому)": "",
}
# Ключі параметрів ФОПа, які зберігаються як текст (решта — числа).
CLIENT_PARAM_TEXT_KEYS = {
    "Тип ставки ЄП", "Тип ставки ВЗ", "Пільга ЄСВ (поточна)",
    "Посилання на таблицю ФОПа", "Аркуш операцій",
    "Виключити фрази (через кому)", "Виключити клієнтів (через кому)",
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

UKR_MONTHS = [
    "Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень",
    "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень",
]

STATUS_YES = "Так"
STATUS_NO = "Ні"
STATUS_NA = "Не актуально"

st.set_page_config(page_title="CRM ФОП — Майстер-Дашборд", page_icon="📊", layout="wide")


# ============================================================
# 2. ПІДКЛЮЧЕННЯ ДО GOOGLE SHEETS
# ============================================================
@st.cache_resource(show_spinner=False)
def get_client():
    """Авторизація через сервісний акаунт (Service Account)."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    return get_client().open(SPREADSHEET_NAME)


def get_ws(sheet_name: str):
    return get_spreadsheet().worksheet(sheet_name)


@st.cache_data(ttl=30, show_spinner=False)
def load_df(sheet_name: str) -> pd.DataFrame:
    """Читає аркуш і повертає pandas DataFrame. Кешується на 30 секунд."""
    records = get_ws(sheet_name).get_all_records()
    return pd.DataFrame(records)


def refresh_data():
    """Скидає кеш, щоб наступне читання підтягнуло свіжі дані з таблиці."""
    st.cache_data.clear()


def append_row(sheet_name: str, row: list):
    get_ws(sheet_name).append_row(row, value_input_option="USER_ENTERED")


# ============================================================
# 3. ДОПОМІЖНІ ФУНКЦІЇ
# ============================================================
def parse_ukr_date(s) -> "date | None":
    """Перетворює рядок 'ДД.ММ.РРРР' на date. Повертає None, якщо не вдалось."""
    try:
        return datetime.strptime(str(s).strip(), "%d.%m.%Y").date()
    except Exception:
        return None


def sync_history_sheet(history_sheet: str, value_col: str, target_col: str) -> int:
    """
    Універсальна синхронізація «запланованих змін»: бере аркуш історії
    (колонки: ПІБ ФОП | <value_col> | Діє з), для кожного ФОПа знаходить
    найновіший запис, чия дата «Діє з» вже настала, і якщо значення там
    відрізняється від поточного в «Клієнти» (колонка target_col) — оновлює
    клітинку. Повертає кількість застосованих змін. Безпечно викликати
    повторно — вже застосовані зміни просто нічого не міняють.
    """
    history_df = load_df(history_sheet)
    if history_df.empty:
        return 0
    required = {"ПІБ ФОП", value_col, "Діє з"}
    if not required.issubset(history_df.columns):
        return 0

    clients_df = load_df(SHEET_CLIENTS)
    if clients_df.empty or "ПІБ ФОП" not in clients_df.columns or target_col not in clients_df.columns:
        return 0

    today = date.today()
    history_df = history_df.copy()
    history_df["_дата"] = history_df["Діє з"].apply(parse_ukr_date)
    due = history_df[history_df["_дата"].notna() & (history_df["_дата"] <= today)]
    if due.empty:
        return 0

    col_index = list(clients_df.columns).index(target_col) + 1
    applied = 0
    ws_clients = get_ws(SHEET_CLIENTS)
    for name, rows in due.groupby("ПІБ ФОП"):
        latest = rows.sort_values("_дата").iloc[-1]
        new_value = str(latest[value_col]).strip()
        match = clients_df[clients_df["ПІБ ФОП"] == name]
        if match.empty:
            continue
        row_idx = match.index[0]
        current_value = str(clients_df.loc[row_idx, target_col]).strip()
        if new_value and new_value != current_value:
            row_number = row_idx + 2  # +1 з нуля, +1 заголовок
            cell = rowcol_to_a1(row_number, col_index)
            ws_clients.update(cell, [[new_value]])
            applied += 1

    if applied:
        refresh_data()
    return applied


def sync_group_changes() -> int:
    return sync_history_sheet(SHEET_GROUP_HISTORY, "Нова група", "Група платника")


def sync_esv_benefit_changes() -> int:
    return sync_history_sheet(SHEET_ESV_BENEFIT_HISTORY, "Пільга ЄСВ", "Пільга ЄСВ (поточна)")


def upcoming_change(name: str, history_df: pd.DataFrame, value_col: str) -> str:
    """Повертає короткий текст про найближчу ще НЕ застосовану заплановану зміну."""
    if history_df.empty or "ПІБ ФОП" not in history_df.columns:
        return ""
    today = date.today()
    rows = history_df[history_df["ПІБ ФОП"] == name].copy()
    if rows.empty:
        return ""
    rows["_дата"] = rows["Діє з"].apply(parse_ukr_date)
    future = rows[rows["_дата"].notna() & (rows["_дата"] > today)].sort_values("_дата")
    if future.empty:
        return ""
    first = future.iloc[0]
    return f"→ {first[value_col]} з {first['Діє з']}"


def find_amount_column(columns):
    """Серед кількох колонок 'сума...' (напр. 'в валюті рахунку' і 'в валюті операції')
    обирає ту, що стосується валюти РАХУНКУ — вона і є те, що реально надійшло ФОПу."""
    candidates = [c for c in columns if "сума" in str(c).lower()]
    if not candidates:
        return None
    for c in candidates:
        if "рахун" in str(c).lower():
            return c
    return candidates[0]


def find_currency_column(columns):
    """
    Шукає колонку валюти. Виключає колонки типу 'Сума в валюті рахунку' —
    вони теж містять слово 'валюті', але це не колонка валюти, а суми.
    """
    for col in columns:
        low = str(col).lower()
        if "валют" in low and "сума" not in low:
            return col
    return None


def load_uploaded_statement(uploaded_file) -> pd.DataFrame:
    """
    Читає завантажений файл банківської виписки (.xlsx/.xls/.csv). Автоматично
    знаходить рядок заголовків (він не завжди перший рядок файлу — часто вище є
    назва рахунку й період), шукаючи перший рядок, де є одночасно щось на кшталт
    "дата" і "сума".
    """
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file, header=None, dtype=object)
    else:
        raw = pd.read_excel(uploaded_file, header=None, dtype=object)

    header_row_idx = None
    for i in range(min(10, len(raw))):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        has_date = any("дата" in v for v in row_vals)
        has_amount = any("сума" in v for v in row_vals)
        if has_date and has_amount:
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(
            "Не вдалось знайти рядок заголовків (з колонками «Дата» і «Сума») "
            "серед перших 10 рядків файлу."
        )

    headers = [str(v).strip() for v in raw.iloc[header_row_idx].tolist()]
    data = raw.iloc[header_row_idx + 1:].copy()
    data.columns = headers
    data = data.dropna(how="all")
    return data


def extract_sheet_id(url_or_id: str) -> str:
    """Витягує ID Google-таблиці з повного посилання. Якщо це вже ID — повертає як є."""
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", str(url_or_id))
    return match.group(1) if match else str(url_or_id).strip()


def find_column(columns, keywords):
    """Шукає серед назв колонок ту, що містить одне з ключових слів (без урахування регістру)."""
    for col in columns:
        low = str(col).lower()
        if any(kw in low for kw in keywords):
            return col
    return None


def row_hash(*parts) -> str:
    """Короткий хеш для дедуплікації рядків з особистих таблиць ФОПів (де немає власного ID)."""
    raw = "|".join(str(p) for p in parts)
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:16]


def parse_sheet_date(val):
    """
    Розбирає дату, яка могла прийти або текстом ("27.07.2026" / "2026-08-05 18:10:25"),
    або "сирим" серійним номером дати Google Sheets (число днів від 30.12.1899) —
    останнє трапляється, якщо в колонці справжні дати Google, а не текст.
    Формат "РРРР-ММ-ДД" розпізнається окремо від "ДД.ММ.РРРР" — інакше
    dayfirst=True може переплутати місяць і день у форматі з роком спочатку.
    """
    if val is None or val == "":
        return pd.NaT
    if isinstance(val, (int, float)):
        try:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(val))
        except Exception:
            return pd.NaT
    s = str(val).strip()
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s):
        return pd.to_datetime(s, dayfirst=False, errors="coerce")
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def parse_ua_number(val):
    """
    Розбирає число, яке могло прийти як текст із комою замість крапки
    (українська локаль, напр. "345,45") — без цього pandas/gspread можуть
    сплутати кому з розділювачем тисяч і "з'їсти" її, перетворивши 345,45 на 34545.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip().replace(" ", "").replace("\u00a0", "")
    if s == "":
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    return pd.to_numeric(s, errors="coerce")


def fetch_ledger_df(sheet_ref: str, tab_name: str) -> pd.DataFrame:
    """
    Читає вказаний аркуш з ОКРЕМОЇ особистої Google Таблиці ФОПа (не CRM_FOP_Master).
    value_render_option=unformatted — щоб отримати справжні числа, а не рядок з
    комою за локаллю показу. numericise_ignore=["all"] — щоб gspread не пробував
    сам "вгадувати" числа (саме це псує коми на 345,45 -> 34545); розбір чисел
    робимо самі через parse_ua_number, це надійніше.
    """
    sheet_id = extract_sheet_id(sheet_ref)
    sh = get_client().open_by_key(sheet_id)
    ws = sh.worksheet(tab_name)
    records = ws.get_all_records(
        value_render_option=ValueRenderOption.unformatted,
        numericise_ignore=["all"],
    )
    return pd.DataFrame(records)


def parse_month_label(label: str):
    """'Вересень 2026' -> (9, 2026). Повертає (None, None), якщо не вдалось розпізнати."""
    try:
        month_name, year_str = str(label).rsplit(" ", 1)
        month_idx = UKR_MONTHS.index(month_name) + 1
        return month_idx, int(year_str)
    except Exception:
        return None, None


def get_client_exclusions(name: str) -> dict:
    """
    Повертає правила виключення доходу для клієнта: фрази, які якщо є в
    «Послуга/Коментар» — рядок не рахується як дохід (напр. "власними коштами"),
    і конкретні значення «Клієнт», які теж повністю виключаються (напр. назва
    фонду/установи, чиї надходження не є оподатковуваним доходом ФОПа).
    """
    result = {"phrases": [], "payers": []}
    clients_df = load_df(SHEET_CLIENTS)
    if clients_df.empty or "ПІБ ФОП" not in clients_df.columns:
        return result
    match = clients_df[clients_df["ПІБ ФОП"] == name]
    if match.empty:
        return result
    row = match.iloc[0]
    phrases_raw = str(row.get("Виключити фрази (через кому)", "")).strip()
    if phrases_raw:
        result["phrases"] = [p.strip().lower() for p in phrases_raw.split(",") if p.strip()]
    payers_raw = str(row.get("Виключити клієнтів (через кому)", "")).strip()
    if payers_raw:
        result["payers"] = [p.strip() for p in payers_raw.split(",") if p.strip()]
    return result


def is_income_row_excluded(service, payer, exclusions: dict) -> bool:
    """True, якщо рядок доходу треба виключити з оподатковуваної суми за правилами клієнта."""
    service_l = str(service or "").lower()
    for phrase in exclusions.get("phrases", []):
        if phrase and phrase in service_l:
            return True
    payer_s = str(payer or "").strip()
    if payer_s and payer_s in exclusions.get("payers", []):
        return True
    return False


def compute_month_income(client_name: str, month_label_str: str) -> float:
    """
    Сумує ОПОДАТКОВУВАНИЙ дохід (з аркуша «Доходи», лише UAH) конкретного ФОПа
    за вказаний звітний місяць — з урахуванням правил виключення цього клієнта
    (див. get_client_exclusions): власні перекази, нецільові надходження тощо
    не входять у суму, з якої рахується єдиний податок / ВЗ.
    """
    income_df = load_df(SHEET_INCOME)
    if income_df.empty or "Дата операції" not in income_df.columns or "ПІБ ФОП" not in income_df.columns:
        return 0.0
    m, y = parse_month_label(month_label_str)
    if m is None:
        return 0.0
    amount_col = "Сума, грн" if "Сума, грн" in income_df.columns else "Сума доходу (UAH)"
    if amount_col not in income_df.columns:
        return 0.0

    exclusions = get_client_exclusions(client_name)

    def matches(row):
        d = parse_ukr_date(row.get("Дата операції"))
        if d is None or d.month != m or d.year != y:
            return False
        if row.get("ПІБ ФОП") != client_name:
            return False
        currency = str(row.get("Валюта", "UAH")).strip().upper()
        if currency not in ("UAH", ""):
            return False
        if is_income_row_excluded(row.get("Послуга/Коментар"), row.get("Клієнт"), exclusions):
            return False
        return True

    mask = income_df.apply(matches, axis=1)
    return float(pd.to_numeric(income_df.loc[mask, amount_col], errors="coerce").sum())


def get_client_params(name: str) -> dict:
    """Повертає параметри ФОПа (вартість супроводу, ставки, пільга ЄСВ) з розумними значеннями за замовчуванням."""
    result = dict(CLIENT_PARAM_DEFAULTS)
    clients_df = load_df(SHEET_CLIENTS)
    if clients_df.empty or "ПІБ ФОП" not in clients_df.columns:
        return result
    match = clients_df[clients_df["ПІБ ФОП"] == name]
    if match.empty:
        return result
    row = match.iloc[0]
    for key in result:
        if key not in row or str(row[key]).strip() == "":
            continue
        if key in CLIENT_PARAM_TEXT_KEYS:
            result[key] = str(row[key]).strip()
        else:
            val = pd.to_numeric(row[key], errors="coerce")
            if pd.notna(val):
                result[key] = float(val)
    return result


def month_label(d: date) -> str:
    """Напр. date(2026, 9, 13) -> 'Вересень 2026'."""
    return f"{UKR_MONTHS[d.month - 1]} {d.year}"


def shifted_month_label(offset: int) -> str:
    """Назва місяця зі зсувом offset (може бути від'ємним) відносно поточного."""
    today = date.today()
    total = today.month - 1 + offset
    year = today.year + total // 12
    month = total % 12
    return f"{UKR_MONTHS[month]} {year}"


def days_until(day_of_month: int, today: date) -> int:
    """Скільки днів лишилось до найближчого дедлайну (число day_of_month)."""
    try:
        target = today.replace(day=day_of_month)
    except ValueError:
        target = today.replace(day=28)
    if target < today:
        if today.month == 12:
            target = target.replace(year=today.year + 1, month=1)
        else:
            target = target.replace(month=today.month + 1)
    return (target - today).days


def status_badge(value: str) -> str:
    """HTML-бейдж кольорового статусу (зелений/червоний/сірий)."""
    value = (value or "").strip()
    colors = {
        STATUS_YES: "#1a7f37",
        STATUS_NO: "#c92a2a",
        STATUS_NA: "#868e96",
    }
    bg = colors.get(value, "#868e96")
    return f'<span style="background:{bg};color:white;padding:2px 10px;border-radius:10px;font-size:13px;">{value or "—"}</span>'


def money_badge(accrued: float, paid: float) -> str:
    """HTML-бейдж 'нараховано/сплачено' з кольором за станом боргу."""
    debt = round(accrued - paid, 2)
    if abs(accrued) < 0.01 and abs(paid) < 0.01:
        bg, text = "#868e96", "не нараховано"
    elif debt > 0.01:
        bg, text = "#c92a2a", f"борг {debt:,.2f} грн".replace(",", " ")
    elif debt < -0.01:
        bg, text = "#1a7f37", f"переплата {abs(debt):,.2f} грн".replace(",", " ")
    else:
        bg, text = "#1a7f37", "сплачено повністю"
    detail = f"{paid:,.2f} / {accrued:,.2f} грн".replace(",", " ")
    return (
        f'<span style="background:{bg};color:white;padding:2px 10px;border-radius:10px;font-size:12px;">{text}</span>'
        f'<div style="font-size:11px;color:#888;margin-top:2px;">{detail}</div>'
    )


def render_control_table(df: pd.DataFrame):
    """Малює таблицю статусів контролю: нараховано/сплачено/борг по кожній категорії."""
    if df.empty:
        st.info("Дані контролю за цей місяць відсутні.")
        return

    def num(row, col):
        val = pd.to_numeric(row.get(col, 0), errors="coerce")
        return float(val) if pd.notna(val) else 0.0

    headers = ["ПІБ ФОП", "Супровід", "ЄСВ", "Єдиний податок", "ВЗ", "Звітність"]
    html = "<table style='width:100%;border-collapse:collapse;'>"
    html += "<tr>" + "".join(
        f"<th style='text-align:left;padding:6px;border-bottom:2px solid #ddd;'>{h}</th>" for h in headers
    ) + "</tr>"
    for _, row in df.iterrows():
        html += "<tr>"
        html += f"<td style='padding:8px 6px;border-bottom:1px solid #eee;vertical-align:top;'>{row.get('ПІБ ФОП', '')}</td>"
        html += f"<td style='padding:8px 6px;border-bottom:1px solid #eee;'>{money_badge(num(row, 'Супровід нараховано'), num(row, 'Супровід сплачено'))}</td>"
        html += f"<td style='padding:8px 6px;border-bottom:1px solid #eee;'>{money_badge(num(row, 'ЄСВ нараховано'), num(row, 'ЄСВ сплачено'))}</td>"
        html += f"<td style='padding:8px 6px;border-bottom:1px solid #eee;'>{money_badge(num(row, 'Єдиний податок нараховано'), num(row, 'Єдиний податок сплачено'))}</td>"
        html += f"<td style='padding:8px 6px;border-bottom:1px solid #eee;'>{money_badge(num(row, 'ВЗ нараховано'), num(row, 'ВЗ сплачено'))}</td>"
        report = str(row.get("Звітність подано", "")).strip() or STATUS_NO
        html += f"<td style='padding:8px 6px;border-bottom:1px solid #eee;vertical-align:top;'>{status_badge(report)}</td>"
        html += "</tr>"
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# 4. СТОРІНКА: ГОЛОВНИЙ ДАШБОРД
# ============================================================
def page_dashboard():
    st.title("📊 Головний Дашборд")
    today = date.today()
    current_month = month_label(today)

    # --- Блок дедлайнів ---
    d_esv = days_until(19, today)
    d_ep = days_until(20, today)
    c1, c2, c3 = st.columns(3)
    with c1:
        if d_esv <= 3:
            st.error(f"⏰ ЄСВ (до 19-го числа): {d_esv} дн.")
        else:
            st.info(f"ЄСВ (до 19-го числа): {d_esv} дн.")
    with c2:
        if d_ep <= 3:
            st.error(f"⏰ Єдиний податок (до 20-го числа): {d_ep} дн.")
        else:
            st.info(f"Єдиний податок (до 20-го числа): {d_ep} дн.")
    with c3:
        st.success(f"Поточний звітний період: {current_month}")

    st.divider()

    # --- Сумарний дохід за поточний місяць ---
    income_df = load_df(SHEET_INCOME)
    total_income = 0.0
    amount_col = "Сума, грн" if "Сума, грн" in income_df.columns else "Сума доходу (UAH)"
    if not income_df.empty and "Дата операції" in income_df.columns and amount_col in income_df.columns:
        def is_this_month(d_str):
            try:
                d = datetime.strptime(str(d_str), "%d.%m.%Y").date()
                return d.month == today.month and d.year == today.year
            except Exception:
                return False
        mask = income_df["Дата операції"].apply(is_this_month)
        if "Валюта" in income_df.columns:
            # Сумуємо лише гривневі операції, щоб не змішувати валюти в одній сумі
            mask &= income_df["Валюта"].astype(str).str.upper().isin(["UAH", ""])
        filtered_income = income_df.loc[mask].copy()
        if not filtered_income.empty and "ПІБ ФОП" in filtered_income.columns:
            # Виключаємо рядки за правилами кожного клієнта (власні перекази тощо)
            exclusions_cache = {}

            def is_excluded(row):
                name = row.get("ПІБ ФОП")
                if name not in exclusions_cache:
                    exclusions_cache[name] = get_client_exclusions(name)
                return is_income_row_excluded(
                    row.get("Послуга/Коментар"), row.get("Клієнт"), exclusions_cache[name]
                )

            keep_mask = ~filtered_income.apply(is_excluded, axis=1)
            filtered_income = filtered_income.loc[keep_mask]
        total_income = pd.to_numeric(
            filtered_income[amount_col], errors="coerce"
        ).sum() if not filtered_income.empty else 0.0

    st.metric(
        f"Сумарний дохід усіх активних ФОП за {current_month}",
        f"{total_income:,.2f} грн".replace(",", " "),
    )

    st.divider()

    # --- Таблиця статусів поточного місяця ---
    st.subheader(f"Статуси за {current_month}")
    control_df = load_df(SHEET_CONTROL)
    if not control_df.empty and "Звітний Місяць" in control_df.columns:
        month_df = control_df[control_df["Звітний Місяць"] == current_month]
    else:
        month_df = pd.DataFrame()
    render_control_table(month_df)

    if st.button("🔄 Оновити дані"):
        refresh_data()
        st.rerun()


# ============================================================
# 5. СТОРІНКА: ВНЕСЕННЯ ДАНИХ
# ============================================================
def page_data_entry():
    st.title("✏️ Внесення даних")
    clients_df = load_df(SHEET_CLIENTS)
    active_clients = []
    if not clients_df.empty and "Статус" in clients_df.columns:
        active_clients = clients_df.loc[
            clients_df["Статус"] == "Активний", "ПІБ ФОП"
        ].tolist()

    tab1, tab2, tab3, tab4 = st.tabs([
        "💰 Дохід ФОПа", "✅ Статуси за місяць", "🔄 З таблиці ФОПа", "📁 З файлу виписки",
    ])

    # --- Форма 1: дохід ---
    with tab1:
        if not active_clients:
            st.warning("Спочатку додайте хоча б одного активного ФОПа на сторінці «База Клієнтів».")
        with st.form("income_form", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                d = st.date_input("Дата операції", value=date.today(), format="DD.MM.YYYY")
                fop = st.selectbox("ФОП", active_clients) if active_clients else None
                payer = st.text_input("Клієнт (хто оплатив)", placeholder="напр. Гончарко Анна Павлівна")
                source = st.selectbox("Джерело", SOURCE_OPTIONS)
            with col_b:
                amount = st.number_input("Сума, грн", min_value=0.0, step=100.0, format="%.2f")
                currency = st.selectbox("Валюта", CURRENCY_OPTIONS, index=0)
                commission = st.number_input("Комісія, грн", min_value=0.0, step=1.0, format="%.2f")
                doc_number = st.text_input("№ документа (необов'язково)")
            service = st.text_input("Послуга / коментар", placeholder="напр. За освітні послуги")
            submitted = st.form_submit_button("Зберегти")
            if submitted:
                if not fop:
                    st.error("Немає жодного активного ФОПа для вибору.")
                else:
                    append_row(SHEET_INCOME, [
                        d.strftime("%d.%m.%Y"), fop, payer, amount, currency,
                        commission, source, service, doc_number, "",
                    ])
                    refresh_data()
                    st.success(f"Дохід {amount:.2f} грн для {fop} збережено.")

    # --- Форма 2: статуси місяця ---
    with tab2:
        if not active_clients:
            st.warning("Спочатку додайте хоча б одного активного ФОПа на сторінці «База Клієнтів».")
        else:
            client2 = st.selectbox("ФОП", active_clients, key="control_client")
            month_options = [shifted_month_label(o) for o in (-1, 0, 1, 2)]
            month_sel = st.selectbox("Звітний місяць", month_options, index=1, key="control_month")

            params = get_client_params(client2)
            month_income = compute_month_income(client2, month_sel)
            esv_benefit = params["Пільга ЄСВ (поточна)"]
            ep_rate_type = params["Тип ставки ЄП"]
            vz_rate_type = params["Тип ставки ВЗ"]

            # Авто-розрахунок нарахувань за параметрами клієнта (можна змінити нижче вручну)
            default_service = params["Вартість супроводу, грн/міс"]
            default_esv = 0.0 if esv_benefit != "Немає" else params["Ставка ЄСВ, грн/міс"]
            if ep_rate_type == "Відсоток від доходу":
                default_ep = round(month_income * params["Ставка ЄП"] / 100, 2)
            else:
                default_ep = params["Ставка ЄП"]
            if vz_rate_type == "Відсоток від доходу":
                default_vz = round(month_income * params["Ставка ВЗ"] / 100, 2)
            else:
                default_vz = params["Ставка ВЗ"]

            info_bits = [
                f"Пільга ЄСВ: **{esv_benefit}**",
                f"Тариф ЄП: **{ep_rate_type}, {params['Ставка ЄП']}**",
                f"Тариф ВЗ: **{vz_rate_type}, {params['Ставка ВЗ']}**",
            ]
            if "Відсоток від доходу" in (ep_rate_type, vz_rate_type):
                info_bits.append(f"дохід за місяць: **{month_income:,.2f} грн**".replace(",", " "))
            st.caption(" · ".join(info_bits))

            # Підтягуємо вже збережений запис за цей місяць (якщо є) — для редагування
            control_df = load_df(SHEET_CONTROL)
            existing_idx, existing_row = None, None
            if not control_df.empty and "Звітний Місяць" in control_df.columns:
                match = control_df[
                    (control_df["Звітний Місяць"] == month_sel) & (control_df["ПІБ ФОП"] == client2)
                ]
                if not match.empty:
                    existing_idx = match.index[0]
                    existing_row = control_df.loc[existing_idx]

            def prefill(col, default):
                if existing_row is not None:
                    val = pd.to_numeric(existing_row.get(col), errors="coerce")
                    if pd.notna(val):
                        return float(val)
                return default

            form_key = f"control_form_{client2}_{month_sel}"
            with st.form(form_key):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Нараховано**")
                    service_accrued = st.number_input("Супровід, грн", min_value=0.0, step=50.0,
                                                        value=prefill("Супровід нараховано", default_service),
                                                        format="%.2f", key=f"sa_{form_key}")
                    esv_accrued = st.number_input("ЄСВ, грн", min_value=0.0, step=10.0,
                                                    value=prefill("ЄСВ нараховано", default_esv),
                                                    format="%.2f", key=f"ea_{form_key}")
                    ep_accrued = st.number_input("Єдиний податок, грн", min_value=0.0, step=10.0,
                                                   value=prefill("Єдиний податок нараховано", default_ep),
                                                   format="%.2f", key=f"epa_{form_key}")
                    vz_accrued = st.number_input("ВЗ, грн", min_value=0.0, step=10.0,
                                                   value=prefill("ВЗ нараховано", default_vz),
                                                   format="%.2f", key=f"vza_{form_key}")
                with col2:
                    st.markdown("**Сплачено**")
                    service_paid = st.number_input("Супровід, грн ", min_value=0.0, step=50.0,
                                                     value=prefill("Супровід сплачено", 0.0),
                                                     format="%.2f", key=f"sp_{form_key}")
                    esv_paid = st.number_input("ЄСВ, грн ", min_value=0.0, step=10.0,
                                                 value=prefill("ЄСВ сплачено", 0.0),
                                                 format="%.2f", key=f"ep_{form_key}")
                    ep_paid = st.number_input("Єдиний податок, грн ", min_value=0.0, step=10.0,
                                                value=prefill("Єдиний податок сплачено", 0.0),
                                                format="%.2f", key=f"epp_{form_key}")
                    vz_paid = st.number_input("ВЗ, грн ", min_value=0.0, step=10.0,
                                                value=prefill("ВЗ сплачено", 0.0),
                                                format="%.2f", key=f"vzp_{form_key}")
                report_default = existing_row is not None and str(existing_row.get("Звітність подано", "")).strip() == STATUS_YES
                report_filed = st.checkbox("Звітність подано", value=report_default, key=f"rf_{form_key}")
                submitted2 = st.form_submit_button("Зберегти статуси")
                if submitted2:
                    row_values = [
                        month_sel, client2,
                        service_accrued, service_paid,
                        esv_accrued, esv_paid,
                        ep_accrued, ep_paid,
                        vz_accrued, vz_paid,
                        STATUS_YES if report_filed else STATUS_NO,
                    ]
                    ws = get_ws(SHEET_CONTROL)
                    if existing_idx is not None:
                        row_number = existing_idx + 2
                        ws.update(f"A{row_number}:K{row_number}", [row_values])
                        st.success(f"Статуси для {client2} за {month_sel} оновлено.")
                    else:
                        ws.append_row(row_values, value_input_option="USER_ENTERED")
                        st.success(f"Статуси для {client2} за {month_sel} додано.")
                    refresh_data()

    # --- Форма 3: синхронізація з особистою таблицею ФОПа ---
    with tab3:
        if not active_clients:
            st.warning("Спочатку додайте хоча б одного активного ФОПа на сторінці «База Клієнтів».")
        else:
            st.caption(
                "Підтягує рядки з аркуша «Перелік операцій» ОСОБИСТОЇ Google Таблиці цього "
                "ФОПа (тієї, яку ти вже ведеш) — щоб не вносити дохід двічі. Таблицю треба "
                "один раз розшарити на службовий email (ролі «Читач» достатньо) і вказати "
                "посилання в «База Клієнтів» → «Редагувати параметри ФОПа»."
            )
            ledger_client = st.selectbox("ФОП", active_clients, key="ledger_client_select")
            clients_df3 = load_df(SHEET_CLIENTS)
            row3 = clients_df3[clients_df3["ПІБ ФОП"] == ledger_client]
            sheet_ref = ""
            tab_name = ""
            if not row3.empty:
                if "Посилання на таблицю ФОПа" in row3.columns:
                    sheet_ref = str(row3.iloc[0].get("Посилання на таблицю ФОПа", "")).strip()
                if "Аркуш операцій" in row3.columns:
                    tab_name = str(row3.iloc[0].get("Аркуш операцій", "")).strip()
            if not tab_name:
                tab_name = f"Перелік операцій {date.today().year}"

            if not sheet_ref:
                st.warning(
                    f"Для {ledger_client} не вказано посилання на особисту таблицю. "
                    "Додай його в «База Клієнтів» → «Редагувати параметри ФОПа»."
                )
            else:
                st.caption(f"Джерело: аркуш «{tab_name}» у вказаній таблиці.")
                if st.button("📥 Завантажити дохід з таблиці ФОПа"):
                    try:
                        raw_df = fetch_ledger_df(sheet_ref, tab_name)
                        cols = list(raw_df.columns)
                        date_col = find_column(cols, ["дата"])
                        amount_col = find_column(cols, ["сума"])
                        if not date_col or not amount_col:
                            st.error(
                                "Не вдалось знайти колонки «Дата» і «Сума» на аркуші. "
                                f"Знайдені колонки: {', '.join(str(c) for c in cols)}."
                            )
                        else:
                            currency_col = find_currency_column(cols)
                            commission_col = find_column(cols, ["комісі", "комиси"])
                            payer_col = find_column(cols, ["клієнт", "контрагент"])
                            service_col = find_column(cols, ["послуг", "деталі", "призначення"])
                            doc_col = find_column(cols, ["документ", "накладн", "акт"])
                            source_col = find_column(cols, ["джерел", "банк"])

                            parsed_rows = []
                            for _, r in raw_df.iterrows():
                                d = parse_sheet_date(r.get(date_col))
                                if pd.isna(d):
                                    continue
                                amount = parse_ua_number(r.get(amount_col))
                                if amount is None or pd.isna(amount) or amount <= 0:
                                    continue
                                date_str = d.strftime("%d.%m.%Y")
                                payer = str(r.get(payer_col, "")).strip() if payer_col else ""
                                service = str(r.get(service_col, "")).strip() if service_col else ""
                                currency = str(r.get(currency_col, "UAH")).strip() if currency_col else "UAH"
                                commission_val = parse_ua_number(r.get(commission_col)) if commission_col else None
                                commission = float(commission_val) if commission_val is not None and pd.notna(commission_val) else 0.0
                                doc_number = str(r.get(doc_col, "")).strip() if doc_col else ""
                                source = str(r.get(source_col, "")).strip() if source_col else ""
                                h = row_hash(date_str, amount, payer, service)
                                parsed_rows.append({
                                    "Дата": date_str, "Клієнт": payer, "Сума": round(float(amount), 2),
                                    "Валюта": currency or "UAH", "Комісія": commission,
                                    "Джерело": source, "Послуга": service, "№ документа": doc_number,
                                    "hash": h,
                                })
                            st.session_state["ledger_rows"] = parsed_rows
                            st.session_state["ledger_sync_client"] = ledger_client
                            if not parsed_rows:
                                st.info("Не знайдено рядків з коректними датою й сумою на цьому аркуші.")
                    except gspread.exceptions.SpreadsheetNotFound:
                        st.error("Таблицю не знайдено — перевір посилання.")
                    except gspread.exceptions.WorksheetNotFound:
                        st.error(f"Аркуш «{tab_name}» не знайдено в цій таблиці — перевір назву (вона часто відрізняється за роком).")
                    except gspread.exceptions.APIError:
                        st.error(
                            "Немає доступу до цієї таблиці. Розшар її на службовий email "
                            "(той самий, що й для CRM_FOP_Master) з роллю «Читач»."
                        )
                    except Exception as e:
                        st.error(str(e))

            ledger_rows = st.session_state.get("ledger_rows")
            if ledger_rows and st.session_state.get("ledger_sync_client") == ledger_client:
                existing_df = load_df(SHEET_INCOME)
                existing_ids = set()
                if not existing_df.empty and "ID_транзакції" in existing_df.columns:
                    existing_ids = set(existing_df["ID_транзакції"].astype(str))

                preview_rows = []
                for r in ledger_rows:
                    already = r["hash"] in existing_ids
                    preview_rows.append({
                        "Імпортувати": not already,
                        "Вже імпортовано": already,
                        "Дата": r["Дата"], "Клієнт": r["Клієнт"], "Сума, грн": r["Сума"],
                        "Комісія, грн": r["Комісія"],
                        "Валюта": r["Валюта"], "Джерело": r["Джерело"], "Послуга": r["Послуга"],
                        "№ документа": r["№ документа"],
                        "hash": r["hash"],
                    })
                preview_df = pd.DataFrame(preview_rows)
                st.subheader("Перегляд перед імпортом")
                st.caption("Зніми позначку з операцій, які не треба імпортувати (напр. вже занесені вручну раніше).")
                edited = st.data_editor(
                    preview_df,
                    column_config={
                        "Імпортувати": st.column_config.CheckboxColumn(),
                        "Вже імпортовано": st.column_config.CheckboxColumn(disabled=True),
                    },
                    disabled=["Вже імпортовано", "Дата", "Клієнт", "Сума, грн", "Комісія, грн",
                              "Валюта", "Джерело", "Послуга", "№ документа", "hash"],
                    hide_index=True,
                    use_container_width=True,
                    key="ledger_editor",
                    column_order=["Імпортувати", "Вже імпортовано", "Дата", "Клієнт", "Сума, грн",
                                  "Комісія, грн", "Валюта", "Джерело", "Послуга", "№ документа"],
                )
                to_import = edited[edited["Імпортувати"] & ~edited["Вже імпортовано"]]
                st.write(f"Обрано до імпорту: **{len(to_import)}** операцій.")
                if st.button("✅ Імпортувати обрані у «Доходи»", key="ledger_import_btn", disabled=to_import.empty):
                    ws = get_ws(SHEET_INCOME)
                    rows_to_write = [
                        [r["Дата"], ledger_client, r["Клієнт"], r["Сума, грн"], r["Валюта"],
                         r["Комісія, грн"], r["Джерело"], r["Послуга"], r["№ документа"], r["hash"]]
                        for _, r in to_import.iterrows()
                    ]
                    # Один запит на весь пакет одразу — окремий запит на кожен рядок
                    # швидко впирається в ліміт Google Sheets API (записів/хв).
                    ws.append_rows(rows_to_write, value_input_option="USER_ENTERED")
                    refresh_data()
                    st.success(f"Імпортовано {len(to_import)} операцій у «Доходи».")
                    st.session_state["ledger_rows"] = None
                    st.rerun()

    # --- Форма 4: імпорт із завантаженого файлу виписки (без токенів і доступів) ---
    with tab4:
        if not active_clients:
            st.warning("Спочатку додайте хоча б одного активного ФОПа на сторінці «База Клієнтів».")
        else:
            st.caption(
                "Найпростіший спосіб без жодних токенів чи доступів: просто завантаж файл "
                "виписки (.xlsx/.xls/.csv), який тобі дав клієнт або сам банк. Застосунок сам "
                "знайде потрібні колонки за заголовками."
            )
            file_client = st.selectbox("ФОП", active_clients, key="file_client_select")
            file_source_default = st.text_input(
                "Джерело коштів (застосується до всіх рядків цього файлу)",
                placeholder="напр. Ощадбанк, Приватбанк, Monobank...",
                key="file_source_default",
            )
            uploaded_file = st.file_uploader(
                "Файл виписки", type=["xlsx", "xls", "csv"], key="statement_uploader"
            )

            if uploaded_file is not None and st.button("📥 Розпізнати файл"):
                try:
                    raw_df = load_uploaded_statement(uploaded_file)
                    cols = list(raw_df.columns)
                    date_col = find_column(cols, ["дата"])
                    amount_col = find_amount_column(cols)
                    if not date_col or not amount_col:
                        st.error(
                            "Не вдалось знайти колонки «Дата» і «Сума» у файлі. "
                            f"Знайдені колонки: {', '.join(str(c) for c in cols)}."
                        )
                    else:
                        currency_col = find_currency_column(cols)
                        commission_col = find_column(cols, ["комісі", "комиси"])
                        payer_col = find_column(cols, ["контрагент", "клієнт"])
                        service_col = find_column(cols, ["деталі", "послуг", "призначення"])
                        doc_col = find_column(cols, ["документ", "накладн", "акт"])
                        source_col = find_column(cols, ["джерел"])

                        parsed_rows = []
                        for _, r in raw_df.iterrows():
                            d = parse_sheet_date(r.get(date_col))
                            if pd.isna(d):
                                continue
                            amount = parse_ua_number(r.get(amount_col))
                            if amount is None or pd.isna(amount) or amount <= 0:
                                continue
                            date_str = d.strftime("%d.%m.%Y")
                            payer = str(r.get(payer_col, "")).strip() if payer_col else ""
                            service = str(r.get(service_col, "")).strip() if service_col else ""
                            currency = str(r.get(currency_col, "UAH")).strip() if currency_col else "UAH"
                            commission_val = parse_ua_number(r.get(commission_col)) if commission_col else None
                            commission = float(commission_val) if commission_val is not None and pd.notna(commission_val) else 0.0
                            doc_number = str(r.get(doc_col, "")).strip() if doc_col else ""
                            source = str(r.get(source_col, "")).strip() if source_col else file_source_default.strip()
                            h = row_hash(date_str, amount, payer, service)
                            parsed_rows.append({
                                "Дата": date_str, "Клієнт": payer, "Сума": round(float(amount), 2),
                                "Валюта": currency or "UAH", "Комісія": commission,
                                "Джерело": source, "Послуга": service, "№ документа": doc_number,
                                "hash": h,
                            })
                        st.session_state["file_rows"] = parsed_rows
                        st.session_state["file_sync_client"] = file_client
                        if not parsed_rows:
                            st.info("Не знайдено рядків з коректними датою й сумою у файлі.")
                        else:
                            st.success(f"Розпізнано {len(parsed_rows)} операцій із надходженнями.")
                except Exception as e:
                    st.error(str(e))

            file_rows = st.session_state.get("file_rows")
            if file_rows and st.session_state.get("file_sync_client") == file_client:
                existing_df = load_df(SHEET_INCOME)
                existing_ids = set()
                if not existing_df.empty and "ID_транзакції" in existing_df.columns:
                    existing_ids = set(existing_df["ID_транзакції"].astype(str))

                preview_rows = []
                for r in file_rows:
                    already = r["hash"] in existing_ids
                    preview_rows.append({
                        "Імпортувати": not already,
                        "Вже імпортовано": already,
                        "Дата": r["Дата"], "Клієнт": r["Клієнт"], "Сума, грн": r["Сума"],
                        "Комісія, грн": r["Комісія"],
                        "Валюта": r["Валюта"], "Джерело": r["Джерело"], "Послуга": r["Послуга"],
                        "№ документа": r["№ документа"],
                        "hash": r["hash"],
                    })
                preview_df = pd.DataFrame(preview_rows)
                st.subheader("Перегляд перед імпортом")
                st.caption(
                    "Зніми позначку з операцій, які не треба імпортувати (напр. власні "
                    "перекази між рахунками чи поповнення картки)."
                )
                edited = st.data_editor(
                    preview_df,
                    column_config={
                        "Імпортувати": st.column_config.CheckboxColumn(),
                        "Вже імпортовано": st.column_config.CheckboxColumn(disabled=True),
                    },
                    disabled=["Вже імпортовано", "Дата", "Клієнт", "Сума, грн", "Комісія, грн",
                              "Валюта", "Джерело", "Послуга", "№ документа", "hash"],
                    hide_index=True,
                    use_container_width=True,
                    key="file_editor",
                    column_order=["Імпортувати", "Вже імпортовано", "Дата", "Клієнт", "Сума, грн",
                                  "Комісія, грн", "Валюта", "Джерело", "Послуга", "№ документа"],
                )
                to_import = edited[edited["Імпортувати"] & ~edited["Вже імпортовано"]]
                st.write(f"Обрано до імпорту: **{len(to_import)}** операцій.")
                if st.button("✅ Імпортувати обрані у «Доходи»", key="file_import_btn", disabled=to_import.empty):
                    ws = get_ws(SHEET_INCOME)
                    rows_to_write = [
                        [r["Дата"], file_client, r["Клієнт"], r["Сума, грн"], r["Валюта"],
                         r["Комісія, грн"], r["Джерело"], r["Послуга"], r["№ документа"], r["hash"]]
                        for _, r in to_import.iterrows()
                    ]
                    ws.append_rows(rows_to_write, value_input_option="USER_ENTERED")
                    refresh_data()
                    st.success(f"Імпортовано {len(to_import)} операцій у «Доходи».")
                    st.session_state["file_rows"] = None
                    st.rerun()


# ============================================================
# 6. СТОРІНКА: БАЗА КЛІЄНТІВ
# ============================================================
def page_clients():
    st.title("👥 База Клієнтів")

    # Автоматично підтягуємо групи й пільги ЄСВ, для яких настала запланована дата
    applied_group = sync_group_changes()
    applied_esv = sync_esv_benefit_changes()
    if applied_group:
        st.info(f"Автоматично оновлено групу платника для {applied_group} ФОП(а) — настала запланована дата зміни.")
    if applied_esv:
        st.info(f"Автоматично оновлено пільгу ЄСВ для {applied_esv} ФОП(а) — настала запланована дата зміни.")

    with st.expander("➕ Додати нового ФОПа", expanded=False):
        with st.form("new_client_form", clear_on_submit=True):
            name = st.text_input("ПІБ ФОП")
            group = st.selectbox("Група платника", GROUP_OPTIONS)
            status = st.selectbox("Статус", ["Активний", "На паузі"])
            st.markdown("**Параметри для розрахунків** (можна змінити пізніше)")
            col1, col2, col3 = st.columns(3)
            with col1:
                fee = st.number_input("Вартість супроводу, грн/міс", min_value=0.0, step=50.0, format="%.2f")
                esv_rate = st.number_input("Ставка ЄСВ, грн/міс", min_value=0.0, step=50.0, format="%.2f")
            with col2:
                ep_type = st.selectbox("Тип ставки ЄП", EP_RATE_TYPE_OPTIONS)
                ep_rate = st.number_input("Ставка ЄП (% доходу або грн)", min_value=0.0, step=0.5, format="%.2f")
            with col3:
                vz_type = st.selectbox("Тип ставки ВЗ", EP_RATE_TYPE_OPTIONS)
                vz_rate = st.number_input("Ставка ВЗ (% доходу або грн)", min_value=0.0, step=0.5, format="%.2f")
            submitted = st.form_submit_button("Додати")
            if submitted:
                if not name.strip():
                    st.error("Вкажіть ПІБ ФОП.")
                else:
                    clients_df = load_df(SHEET_CLIENTS)
                    next_id = 1
                    if not clients_df.empty and "ID" in clients_df.columns:
                        numeric_ids = pd.to_numeric(clients_df["ID"], errors="coerce")
                        if numeric_ids.notna().any():
                            next_id = int(numeric_ids.max()) + 1
                    append_row(SHEET_CLIENTS, [
                        next_id, name.strip(), group, status,
                        fee, ep_type, ep_rate, vz_type, vz_rate, esv_rate, "Немає",
                    ])
                    refresh_data()
                    st.success(f"Клієнта {name} додано (ID {next_id}).")

    clients_df = load_df(SHEET_CLIENTS)
    all_names = clients_df["ПІБ ФОП"].tolist() if not clients_df.empty and "ПІБ ФОП" in clients_df.columns else []
    param_cols_present = all(
        c in clients_df.columns for c in CLIENT_PARAM_DEFAULTS
    ) if not clients_df.empty else False

    with st.expander("✏️ Редагувати параметри ФОПа (вартість супроводу, ставки)", expanded=False):
        if not all_names:
            st.warning("Спочатку додай хоча б одного ФОПа вище.")
        elif not param_cols_present:
            st.warning(
                "На аркуші «Клієнти» ще немає потрібних колонок. Додай їх (див. SETUP.md, "
                "Крок 1): «Вартість супроводу, грн/міс», «Тип ставки ЄП», «Ставка ЄП», "
                "«Тип ставки ВЗ», «Ставка ВЗ», «Ставка ЄСВ, грн/міс», «Пільга ЄСВ (поточна)», "
                "«Посилання на таблицю ФОПа», «Аркуш операцій», «Виключити фрази (через кому)», "
                "«Виключити клієнтів (через кому)»."
            )
        else:
            edit_client = st.selectbox("ФОП", all_names, key="edit_client_select")
            p = get_client_params(edit_client)
            with st.form("edit_client_form"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    new_fee = st.number_input("Вартість супроводу, грн/міс", min_value=0.0, step=50.0,
                                               value=p["Вартість супроводу, грн/міс"], format="%.2f")
                    new_esv_rate = st.number_input("Ставка ЄСВ, грн/міс", min_value=0.0, step=50.0,
                                                    value=p["Ставка ЄСВ, грн/міс"], format="%.2f")
                with col2:
                    new_ep_type = st.selectbox(
                        "Тип ставки ЄП", EP_RATE_TYPE_OPTIONS,
                        index=EP_RATE_TYPE_OPTIONS.index(p["Тип ставки ЄП"]) if p["Тип ставки ЄП"] in EP_RATE_TYPE_OPTIONS else 0,
                    )
                    new_ep_rate = st.number_input("Ставка ЄП (% доходу або грн)", min_value=0.0, step=0.5,
                                                   value=p["Ставка ЄП"], format="%.2f")
                with col3:
                    new_vz_type = st.selectbox(
                        "Тип ставки ВЗ", EP_RATE_TYPE_OPTIONS,
                        index=EP_RATE_TYPE_OPTIONS.index(p["Тип ставки ВЗ"]) if p["Тип ставки ВЗ"] in EP_RATE_TYPE_OPTIONS else 0,
                    )
                    new_vz_rate = st.number_input("Ставка ВЗ (% доходу або грн)", min_value=0.0, step=0.5,
                                                   value=p["Ставка ВЗ"], format="%.2f")
                st.markdown("**Особиста таблиця ФОПа** (для синхронізації доходу)")
                col4, col5 = st.columns(2)
                with col4:
                    new_sheet_ref = st.text_input(
                        "Посилання на Google Таблицю ФОПа",
                        value=p["Посилання на таблицю ФОПа"],
                        placeholder="https://docs.google.com/spreadsheets/d/...",
                    )
                with col5:
                    new_tab_name = st.text_input(
                        "Назва аркуша з операціями",
                        value=p["Аркуш операцій"],
                        placeholder=f"напр. Перелік операцій {date.today().year}",
                    )
                st.markdown("**Виключення з оподатковуваного доходу** (необов'язково)")
                st.caption(
                    "Рядки, що підпадають під ці правила, залишаються в «Доходи», але не "
                    "враховуються в сумі, з якої рахується єдиний податок / ВЗ, і не входять "
                    "у сумарний дохід на Дашборді."
                )
                col6, col7 = st.columns(2)
                with col6:
                    new_excl_phrases = st.text_input(
                        "Виключити фрази в описі (через кому)",
                        value=p["Виключити фрази (через кому)"],
                        placeholder="напр. власними коштами",
                    )
                with col7:
                    new_excl_payers = st.text_input(
                        "Виключити платників (через кому)",
                        value=p["Виключити клієнтів (через кому)"],
                        placeholder="напр. Полтавський ОЦЗ",
                    )
                save = st.form_submit_button("Зберегти параметри")
                if save:
                    match = clients_df[clients_df["ПІБ ФОП"] == edit_client]
                    if match.empty:
                        st.error("Клієнта не знайдено.")
                    else:
                        row_number = match.index[0] + 2
                        ws_c = get_ws(SHEET_CLIENTS)
                        cols = list(clients_df.columns)
                        updates = {
                            "Вартість супроводу, грн/міс": new_fee,
                            "Тип ставки ЄП": new_ep_type,
                            "Ставка ЄП": new_ep_rate,
                            "Тип ставки ВЗ": new_vz_type,
                            "Ставка ВЗ": new_vz_rate,
                            "Ставка ЄСВ, грн/міс": new_esv_rate,
                            "Посилання на таблицю ФОПа": new_sheet_ref.strip(),
                            "Аркуш операцій": new_tab_name.strip(),
                            "Виключити фрази (через кому)": new_excl_phrases.strip(),
                            "Виключити клієнтів (через кому)": new_excl_payers.strip(),
                        }
                        for col_name, val in updates.items():
                            col_index = cols.index(col_name) + 1
                            cell = rowcol_to_a1(row_number, col_index)
                            ws_c.update(cell, [[val]])
                        refresh_data()
                        st.success(f"Параметри {edit_client} оновлено.")

    with st.expander("📅 Запланувати зміну групи", expanded=False):
        st.caption(
            "Наприклад: ФОП зараз на 3 групі, а з 01.10.2026 переходить на 2 групу. "
            "Додай цю зміну заздалегідь — програма сама підставить нову групу в "
            "«Клієнти», щойно настане вказана дата (навіть якщо ти зайдеш в програму пізніше)."
        )
        if not all_names:
            st.warning("Спочатку додай хоча б одного ФОПа вище.")
        else:
            with st.form("group_change_form", clear_on_submit=True):
                gc_client = st.selectbox("ФОП", all_names, key="gc_client")
                gc_group = st.selectbox("Нова група платника", GROUP_OPTIONS, key="gc_group")
                gc_date = st.date_input("Діє з", value=date.today(), format="DD.MM.YYYY", key="gc_date")
                gc_submit = st.form_submit_button("Запланувати")
                if gc_submit:
                    append_row(SHEET_GROUP_HISTORY, [gc_client, gc_group, gc_date.strftime("%d.%m.%Y")])
                    refresh_data()
                    st.success(f"Заплановано: {gc_client} → «{gc_group}» з {gc_date.strftime('%d.%m.%Y')}.")

    with st.expander("📅 Запланувати зміну пільги ЄСВ", expanded=False):
        st.caption(
            "Напр.: ФОП іде в декрет з 01.11.2026 — заплануй тут, і з цієї дати пільга "
            "автоматично з'явиться в «Клієнти» (нарахування ЄСВ у «Внесення даних» "
            "теж стане нульовим для періодів з цієї дати)."
        )
        if not all_names:
            st.warning("Спочатку додай хоча б одного ФОПа вище.")
        elif not param_cols_present:
            st.warning("Спочатку додай колонку «Пільга ЄСВ (поточна)» на аркуш «Клієнти» (див. SETUP.md).")
        else:
            with st.form("esv_benefit_form", clear_on_submit=True):
                eb_client = st.selectbox("ФОП", all_names, key="eb_client")
                eb_benefit = st.selectbox("Пільга ЄСВ", ESV_BENEFIT_OPTIONS, key="eb_benefit")
                eb_date = st.date_input("Діє з", value=date.today(), format="DD.MM.YYYY", key="eb_date")
                eb_submit = st.form_submit_button("Запланувати")
                if eb_submit:
                    append_row(SHEET_ESV_BENEFIT_HISTORY, [eb_client, eb_benefit, eb_date.strftime("%d.%m.%Y")])
                    refresh_data()
                    st.success(f"Заплановано: {eb_client} → пільга «{eb_benefit}» з {eb_date.strftime('%d.%m.%Y')}.")

    st.subheader("Список клієнтів")
    group_history_df = load_df(SHEET_GROUP_HISTORY)
    esv_history_df = load_df(SHEET_ESV_BENEFIT_HISTORY)
    display_df = clients_df.copy()
    if not display_df.empty:
        display_df["Заплановані зміни групи"] = display_df["ПІБ ФОП"].apply(
            lambda n: upcoming_change(n, group_history_df, "Нова група")
        )
        display_df["Заплановані зміни пільги ЄСВ"] = display_df["ПІБ ФОП"].apply(
            lambda n: upcoming_change(n, esv_history_df, "Пільга ЄСВ")
        )
    st.dataframe(display_df, use_container_width=True, hide_index=True)


# ============================================================
# 7. СТОРІНКА: СИНХРОНІЗАЦІЯ З MONOBANK
# ============================================================
def get_mono_tokens() -> dict:
    """
    Токени Monobank по ФОП-клієнтах, задані в secrets.toml у розділі [mono_tokens],
    напр.:
        [mono_tokens]
        "Іванов Іван Іванович" = "uXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Ключ має ЗБІГАТИСЯ з полем "ПІБ ФОП" на аркуші "Клієнти".
    """
    return dict(st.secrets.get("mono_tokens", {}))


def page_mono_sync():
    st.title("🏦 Синхронізація з Monobank")

    tokens = get_mono_tokens()
    if not tokens:
        st.warning(
            "Токени Monobank не налаштовані. Додай розділ [mono_tokens] у "
            "secrets.toml — див. SETUP.md, розділ «Банківський модуль (Monobank)»."
        )
        return

    client = st.selectbox("ФОП", list(tokens.keys()))
    token = tokens[client]
    accounts_key = f"mono_accounts_{client}"

    col1, col2 = st.columns(2)
    with col1:
        date_from = st.date_input("Період з", value=date.today().replace(day=1), key="mono_from")
    with col2:
        date_to = st.date_input("Період по", value=date.today(), key="mono_to")

    info_wait = mono_api.seconds_left(token, "client-info")
    load_label = "🔍 Завантажити рахунки" if accounts_key not in st.session_state else "🔄 Оновити рахунки й баланси"
    if st.button(load_label, disabled=info_wait > 0):
        try:
            info = mono_api.get_client_info(token)
            st.session_state[accounts_key] = {
                "name": info.get("name", "—"),
                "accounts": info.get("accounts", []),
            }
        except Exception as e:
            st.error(str(e))
    if info_wait > 0:
        st.caption(f"⏳ Рахунки можна оновити ще раз через {info_wait} сек. (ліміт Monobank на цей запит — 1/60 сек).")

    cached = st.session_state.get(accounts_key)
    if not cached:
        st.info("Натисни «Завантажити рахунки», щоб побачити список рахунків цього токена.")
        return

    st.success(f"Клієнт Monobank: {cached['name']}")
    accounts = cached["accounts"]
    if not accounts:
        st.error("У цього токена немає жодного доступного рахунку.")
        return

    account_labels = {mono_api.format_account_label(acc): acc for acc in accounts}
    st.caption(
        f"У токена {len(accounts)} рахунок(ів)/картку(ок) — обери саме той, що стосується "
        "діяльності ФОПа (зазвичай гривневий, тип «fop» чи «black», а не доларовий чи інший)."
    )
    selected_label = st.selectbox("Рахунок для виписки", list(account_labels.keys()))
    selected_account = account_labels[selected_label]
    account_id = selected_account["id"]

    stmt_wait = mono_api.seconds_left(token, f"statement-{account_id}")
    if stmt_wait > 0:
        st.caption(f"⏳ Виписку для цього рахунку можна запросити ще раз через {stmt_wait} сек.")

    if st.button("📥 Отримати виписку по обраному рахунку", disabled=stmt_wait > 0):
        try:
            txs = mono_api.get_statement(token, account_id, date_from, date_to)
            income_txs = [t for t in txs if t.get("amount", 0) > 0]
            st.session_state["mono_transactions"] = income_txs
            st.session_state["mono_client"] = client
            st.session_state["mono_account_label"] = selected_label
            st.session_state["mono_account_currency"] = mono_api.CURRENCY_CODES.get(
                selected_account.get("currencyCode"), "UAH"
            )
            if not income_txs:
                st.info("За обраний період надходжень на цей рахунок не знайдено.")
        except Exception as e:
            st.error(str(e))

    txs = st.session_state.get("mono_transactions")
    if txs and st.session_state.get("mono_client") == client:
        acc_currency = st.session_state.get("mono_account_currency", "UAH")
        st.caption(f"Виписка по рахунку: {st.session_state.get('mono_account_label', '')} · валюта {acc_currency}")
        existing_df = load_df(SHEET_INCOME)
        existing_ids = set()
        if not existing_df.empty and "ID_транзакції" in existing_df.columns:
            existing_ids = set(existing_df["ID_транзакції"].astype(str))
        elif not existing_df.empty:
            st.warning(
                "На аркуші «Доходи» немає колонки «ID_транзакції» — додай її останньою колонкою "
                "(див. SETUP.md), інакше програма не зможе відсіювати вже імпортовані операції."
            )

        rows = []
        for t in txs:
            tx_id = str(t.get("id"))
            rows.append({
                "Імпортувати": tx_id not in existing_ids,
                "Вже імпортовано": tx_id in existing_ids,
                "Дата": datetime.fromtimestamp(t.get("time")).strftime("%d.%m.%Y"),
                f"Сума ({acc_currency})": round(t.get("amount", 0) / 100, 2),
                "Опис": t.get("description", ""),
                "ID": tx_id,
            })
        preview_df = pd.DataFrame(rows)

        st.subheader("Перегляд надходжень перед імпортом")
        st.caption(
            "Зніми позначку «Імпортувати» з операцій, які НЕ треба вносити як дохід "
            "(наприклад, власні перекази між рахунками або повернення коштів)."
        )
        edited = st.data_editor(
            preview_df,
            column_config={
                "Імпортувати": st.column_config.CheckboxColumn(),
                "Вже імпортовано": st.column_config.CheckboxColumn(disabled=True),
            },
            disabled=["Вже імпортовано", "Дата", f"Сума ({acc_currency})", "Опис", "ID"],
            hide_index=True,
            use_container_width=True,
            key="mono_editor",
        )

        to_import = edited[edited["Імпортувати"] & ~edited["Вже імпортовано"]]
        st.write(f"Обрано до імпорту: **{len(to_import)}** операцій.")

        if st.button("✅ Імпортувати обрані у «Доходи»", disabled=to_import.empty):
            ws = get_ws(SHEET_INCOME)
            amount_key = f"Сума ({acc_currency})"
            rows_to_write = [
                [r["Дата"], client, "", r[amount_key], acc_currency, 0, "Monobank", r["Опис"], "", r["ID"]]
                for _, r in to_import.iterrows()
            ]
            ws.append_rows(rows_to_write, value_input_option="USER_ENTERED")
            refresh_data()
            st.success(f"Імпортовано {len(to_import)} операцій у «Доходи».")
            st.session_state["mono_transactions"] = None
            st.rerun()


# ============================================================
# 8. ГОЛОВНА ФУНКЦІЯ / НАВІГАЦІЯ
# ============================================================
def main():
    st.sidebar.title("📁 CRM ФОП")
    page = st.sidebar.radio(
        "Навігація",
        ["Головний Дашборд", "Внесення даних", "База Клієнтів", "Монобанк"],
    )

    if "gcp_service_account" not in st.secrets:
        st.error(
            "Не знайдено налаштувань доступу до Google Sheets.\n\n"
            "Додайте розділ [gcp_service_account] у файл .streamlit/secrets.toml "
            "(див. інструкцію в SETUP.md)."
        )
        st.stop()

    if page == "Головний Дашборд":
        page_dashboard()
    elif page == "Внесення даних":
        page_data_entry()
    elif page == "База Клієнтів":
        page_clients()
    elif page == "Монобанк":
        page_mono_sync()


if __name__ == "__main__":
    main()
