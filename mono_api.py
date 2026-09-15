# -*- coding: utf-8 -*-
"""
Інтеграція з Monobank Personal API.
Документація: https://api.monobank.ua/docs/

Кожен ФОП-клієнт самостійно генерує свій особистий токен на сторінці
https://api.monobank.ua/ (розділ "Створити токен") і передає його бухгалтеру.
Токени зберігаються ЛИШЕ в secrets.toml (ніколи не в Google Таблиці і не в коді) —
див. SETUP.md, розділ "Банківський модуль (Monobank)".

Обмеження Monobank API, які враховані нижче (кожне — окремий "секундомір"):
- /personal/client-info: 1 запит на 60 секунд для токена;
- /personal/statement: 1 запит на 60 секунд ДЛЯ КОЖНОГО РАХУНКУ окремо
  (тобто виписку по одному рахунку токена можна запитувати незалежно від того,
  коли востаннє питали баланс чи виписку по іншому рахунку цього ж токена);
- виписка видається максимум за 31 добу + 1 годину за один виклик.

Мапа поширених валют (ISO 4217, числовий код -> літерний).
"""

import time
import requests
import streamlit as st

MONO_BASE = "https://api.monobank.ua"
RATE_LIMIT_SECONDS = 60
MAX_RANGE_SECONDS = 31 * 24 * 60 * 60 + 3600  # 31 доба + 1 година, ліміт самого API

CURRENCY_CODES = {
    980: "UAH", 840: "USD", 978: "EUR", 826: "GBP",
    985: "PLN", 124: "CAD", 756: "CHF", 392: "JPY",
}


def _key(token: str, suffix: str) -> str:
    # Використовуємо останні символи токена (не сам токен) як частину ключа стану
    return f"mono_last_call_{token[-8:]}_{suffix}"


def seconds_left(token: str, suffix: str = "client-info") -> int:
    """
    Скільки секунд лишилось чекати перед наступним запитом.
    suffix розрізняє ліміти різних ендпоінтів/рахунків — напр. "client-info"
    для перевірки рахунку, або f"statement-{account_id}" для виписки
    конкретного рахунку.
    """
    last = st.session_state.get(_key(token, suffix))
    if last is None:
        return 0
    remaining = RATE_LIMIT_SECONDS - (time.time() - last)
    return max(0, int(remaining) + 1)


def _mark_called(token: str, suffix: str):
    st.session_state[_key(token, suffix)] = time.time()


def _check_rate_limit(token: str, suffix: str):
    wait = seconds_left(token, suffix)
    if wait > 0:
        raise RuntimeError(
            f"Ліміт Monobank API — 1 запит на {RATE_LIMIT_SECONDS} сек. "
            f"Зачекай ще {wait} сек. і спробуй знову."
        )


def get_client_info(token: str) -> dict:
    """Повертає ім'я клієнта та список його рахунків (з балансами)."""
    _check_rate_limit(token, "client-info")
    resp = requests.get(f"{MONO_BASE}/personal/client-info", headers={"X-Token": token}, timeout=15)
    _mark_called(token, "client-info")
    if resp.status_code == 429:
        raise RuntimeError("Monobank тимчасово обмежив запити. Спробуй за хвилину.")
    if resp.status_code == 403:
        raise RuntimeError("Токен недійсний або відкликаний клієнтом.")
    resp.raise_for_status()
    return resp.json()


def get_statement(token: str, account_id: str, date_from, date_to) -> list:
    """
    Повертає список транзакцій рахунку account_id за період [date_from, date_to]
    (об'єкти datetime.date, включно). Максимум 31 доба за один виклик.
    Ліміт рахується окремо для кожного account_id.
    """
    from_ts = int(time.mktime(date_from.timetuple()))
    to_ts = int(time.mktime(date_to.timetuple())) + 86399  # включно до кінця дня date_to

    if to_ts - from_ts > MAX_RANGE_SECONDS:
        raise ValueError("Період не може перевищувати 31 добу. Обери менший діапазон і повтори для решти періоду.")
    if to_ts < from_ts:
        raise ValueError("Дата «по» не може бути раніше дати «з».")

    suffix = f"statement-{account_id}"
    _check_rate_limit(token, suffix)
    url = f"{MONO_BASE}/personal/statement/{account_id}/{from_ts}/{to_ts}"
    resp = requests.get(url, headers={"X-Token": token}, timeout=20)
    _mark_called(token, suffix)
    if resp.status_code == 429:
        raise RuntimeError("Monobank тимчасово обмежив запити. Спробуй за хвилину.")
    if resp.status_code == 403:
        raise RuntimeError("Токен недійсний або відкликаний клієнтом.")
    resp.raise_for_status()
    return resp.json()


def format_account_label(acc: dict) -> str:
    """Людяний опис рахунку для селектора: ідентифікатор · валюта · тип · баланс."""
    currency = CURRENCY_CODES.get(acc.get("currencyCode"), str(acc.get("currencyCode", "?")))
    balance = acc.get("balance", 0) / 100
    masked = acc.get("maskedPan") or []
    ident = masked[0] if masked else acc.get("iban", acc.get("id", "рахунок"))
    acc_type = acc.get("type", "")
    return f"{ident} · {currency} · {acc_type} · {balance:,.2f} грн".replace(",", " ")
