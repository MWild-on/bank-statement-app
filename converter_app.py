# converter_app.py — Анализ банковской выписки

import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

from ui_common import section_header, apply_global_css


# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

def extract_bank_account(text: str) -> str:
    """Извлечь 20-значный счёт из строки."""
    match = re.search(r"\b\d{20}\b", str(text))
    return match.group(0) if match else ""


def extract_is_from_bailiff(text: str) -> str:
    """
    Определить, пришёл ли платёж от ФССП / УФК и т.п.
    Возвращает "Y" / "N".
    """
    txt = str(text).lower().replace("\n", " ")
    keywords = [
        "уфк", "росп", "осп", "уфссп", "гуссп", "гуфссп",
        "фссп", "фссп россии", "государственная служба судебных приставов",
    ]
    return "Y" if any(kw in txt for kw in keywords) else "N"


def extract_court_order_number(text: str) -> str:
    """Извлечь номер судебного приказа / ИД (текущая версия с приоритетами и исключениями)."""
    text_l = str(text).lower()

    # Приоритет: ВС/ФС + 9 цифр
    priority_match = re.search(r"\b(вс|фс)\s?(\d{9})\b", text_l)
    if priority_match:
        return f"{priority_match.group(1).upper()} {priority_match.group(2)}"

    # Прямой шаблон ИД с 'ид'
    match_id_direct = re.search(r"\bид\s+([\d\-]+/\d{4}(?:-\d{1,3})?)\b", text_l)
    if match_id_direct:
        return match_id_direct.group(1)

    patterns = [
        r"№[а-яa-z]+[\d\-]*-([\d\-]+/\d{4}(?:-\d{1,3})?)",
        r"(?:судебный приказ|суд\.? приказ|с/пр)[^\d]{0,3}([\d]{1,2}-\d{1,4}-\d{1,5}/\d{4})",
        r"(?:судебный приказ|суд\.? приказ|с/пр)\s*(?:№|:)?\s*([\d\-/]+)",
        r"взыскание по ид от \d{2}\.\d{2}\.\d{4} ?№([\d\-/]+)",
        r"по и/д\s*№?\s*([\d\-/]+)",
        r"\bи/д\s*№?\s*([\d\-/]+)",
        r"(?:по\s+)?и/л\s*(?:№|n)?\s*([\d\-/]+)",
        r"\b(?:ид n|ид|n)\s*(?:№|n)?\s*([\d\-]+/\d{4}(?:-\d{1,3})?)\b",
        r"№\s*([\d\-]+/\d{4}(?:-\d{1,3})?)",
        r"суд\.пр\s*([\d\-]+/[\d\-]+)",
        r"исполнительный лист\s*([\d\-]+/\d{4})",
        r"\bил\s+([\d\-]+/\d{4})",
        r"и/л\s*(?:№|n)?\s*([\w\-]+/\d{4})",
        r"по документу\s+([\d\-]+/\d{4})",
        r"с/п\s*([\d\-]+/\d{4})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text_l)
        if not match:
            continue

        value = match.group(1)
        if len(value.strip()) < 5:
            continue

        # Проверяем, что это не номер ИП
        before = text_l[: match.start()]
        if re.search(r"(ип|\bисп\w*)\s*$", before.strip()[-20:]):
            continue
        if re.search(r"-ип$", value):
            continue

        return value

    return ""


def extract_court_order_date(text: str, court_number: str) -> str:
    """Дата судебного приказа вблизи номера приказа."""
    txt = str(text).lower()
    cn = court_number.strip().lower()
    if not cn or len(cn) < 5:
        return ""

    txt_clean = re.sub(r"[()\[\]]", " ", txt)
    pos = txt_clean.find(cn)
    if pos == -1:
        return ""

    context = txt_clean[max(0, pos - 50): pos + 50]
    date_patterns = [
        r"от\s*(\d{2}\.\d{2}\.\d{4})",
        r"от\s*(\d{4}-\d{2}-\d{2})",
        r"(\d{2}\.\d{2}\.\d{4})",
        r"(\d{4}-\d{2}-\d{2})",
    ]
    for pattern in date_patterns:
        m = re.search(pattern, context)
        if m:
            return m.group(1)
    return ""


def extract_ip_number(text: str) -> str:
    """Извлечь номер ИП (исполнительного производства)."""
    t = str(text).lower()

    m1 = re.search(r"(?:и/п|ип)?[ №:]*([0-9]{4,8}/[0-9]{2}/[0-9]{4,8}-ип)\b", t)
    if m1:
        return m1.group(1)

    m2 = re.search(r"(?:и/п|ип)?[ №:]*([0-9]{4,8}/[0-9]{2}/[0-9]{4,8})\b", t)
    if m2:
        before = t[: m2.start()]
        if "ид" not in before[-20:]:
            return m2.group(1)

    m3 = re.search(r"\(ип\s+([\w\-\/]+)", t)
    if m3:
        return m3.group(1)

    return ""


def extract_fio(text: str) -> str:
    """Попытка вытащить ФИО из текста назначения."""
    txt = str(text)
    patterns = [
        r"\bс\s+([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bдолг[а]?:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bдолжника:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"с должника\s+([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"долга взыскателю\s*:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bс:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
    ]
    for pattern in patterns:
        m = re.search(pattern, txt, flags=re.IGNORECASE)
        if m:
            return m.group(1).title().strip()
    return ""


def process_bank_statement(df: pd.DataFrame) -> pd.DataFrame:
    """Основная обработка выписки -> итоговая таблица."""
    # Берём только кредитовые операции
    df = df[pd.to_numeric(df["Сумма по кредиту"], errors="coerce") > 0].copy()

    res = pd.DataFrame()
    res["CaseID"] = ""
    res["TransactionType"] = "Оплата"
    res["Sum"] = df["Сумма по кредиту"]

    res["PaymentDate"] = (
        pd.to_datetime(df["Дата проводки"], errors="coerce").dt.date
    )
    res["BookingDate"] = datetime.now().date()

    res["BankAccount"] = df["Кредит"].apply(extract_bank_account)
    res["InvoiceNum"] = ""
    res["InvoiceID"] = ""
    res["PaymentProvider"] = ""

    res["IsFromBailiff"] = df["Счет"].apply(extract_is_from_bailiff)

    res["CourtOrderNumber"] = df["Назначение платежа"].apply(
        extract_court_order_number
    )
    res["Дата приказа"] = df.apply(
        lambda row: extract_court_order_date(
            row["Назначение платежа"],
            extract_court_order_number(row["Назначение платежа"]),
        ),
        axis=1,
    )

    res["Номер ИП"] = df["Назначение платежа"].apply(extract_ip_number)
    res["ФИО"] = df["Назначение платежа"].apply(extract_fio)
    res["Назначение платежа"] = df["Назначение платежа"]

    return res


# ===== ОСНОВНАЯ ФУНКЦИЯ МОДУЛЯ =====

def run():
    # ← единый CSS, как на остальных вкладках!
    apply_global_css()

    section_header(
        "Анализ банковской выписки",
        "Загрузите файл выписки. Я выделю только нужные операции и соберу таблицу..."
    )

    uploaded_file = st.file_uploader(
        "Загрузите файл выписки (Excel)", type=["xlsx", "xls"]
    )

    if not uploaded_file:
        return

    try:
        # Читаем выписку — как и раньше, пропуская первые 2 строки
        df_raw = pd.read_excel(uploaded_file, skiprows=2)

        # Переименовываем нужные столбцы по фиксированным индексам
        df_raw.columns.values[1] = "Дата проводки"
        df_raw.columns.values[4] = "Счет"
        df_raw.columns.values[6] = "Дебет"
        df_raw.columns.values[8] = "Кредит"
        df_raw.columns.values[13] = "Сумма по кредиту"
        df_raw.columns.values[14] = "№ документа"
        df_raw.columns.values[20] = "Назначение платежа"

        df = df_raw.copy()

        st.success("Файл успешно загружен и распознан.")

        df_result = process_bank_statement(df)

        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader("Результат обработки")
        with col2:
            output = BytesIO()
            df_result.to_excel(output, index=False, engine="openpyxl")
            st.download_button(
                "📥 Скачать результат (Excel)",
                data=output.getvalue(),
                file_name="результат_выписки.xlsx",
            )

        st.dataframe(df_result)

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
