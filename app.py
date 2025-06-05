import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ===== Вспомогательные функции =====

def extract_bank_account(text):
    match = re.search(r'\b\d{20}\b', str(text))
    return match.group(0) if match else ""

def extract_is_from_bailiff(text):
    text = str(text).lower().replace('\n', ' ')
    keywords = [
        "уфк", "росп", "осп", "уфссп", "гуссп", "гуфссп",
        "фссп", "фссп россии", "государственная служба судебных приставов"
    ]
    return "Y" if any(kw in text for kw in keywords) else "N"

def extract_court_order_number(text):
    text = str(text).lower()

    # Вариант 1: "Судебный приказ" или его сокращения
    match1 = re.search(r'(?:судебный приказ|суд\.? приказ|с/пр)\s*(?:№|:)?\s*([\d\-\/]+)', text)
    if match1:
        return match1.group(1)


    # Вариант 2: "Взыскание по ИД от <дата> №..."
    match2 = re.search(r'взыскание по ид от \d{2}\.\d{2}\.\d{4} ?№([\d\-\/]+)', text)
    if match2:
        return match2.group(1)

    # Вариант 3: "и/д" + номер
    match3 = re.search(r'по и/д\s*№?\s*([\d\-\/]+)', text)
    if match3:
        return match3.group(1)

    # Вариант 4: "по и/л" + номер
    match4 = re.search(r'по и/л\s*([\d\-\/]+)', text)
    if match4:
        return match4.group(1)

        
    # Вариант 5: "ид", "ид n", или "n", затем где-то в тексте номер с годом
    match5 = re.search(r'(?:ид n|ид|n)\s*№?\s*.*?([\d\-]+/(?:1[0-9]|2[0-9]|30|201[0-9]|202[0-9]|2030))', text)
    if match5:
        return match5.group(1)


    return ""

    
def extract_ip_number(text):
    text = str(text).lower()

    # Вариант 1: ИП с дефисом и годом
    match1 = re.search(r'(?:и/п|ип)[ №:]*([\d\-]+/(?:20|2[1-9]|30|201[0-9]|202[0-9]|2030))\b', text)
    if match1:
        return match1.group(1)

    # Вариант 2: 6+ цифр (старый простой)
    match2 = re.search(r'(?:и/п|ип)[ №:]*([0-9]{6,})', text)
    if match2:
        return match2.group(1)

    # Вариант 3: формат типа 12345/24/00001 с -ип или без него
    match3 = re.search(r'(\d{4,6}/\d{2}/\d{4,6})(?:-ип)?', text)
    if match3:
        return match3.group(1)

    return ""


# ===== Обработка данных =====

def process_bank_statement(df):
    result = pd.DataFrame()
    result["CaseID"] = ""
    result["TransactionType"] = "Оплата"
    result["Sum"] = df["Сумма по кредиту"]
    result["PaymentDate"] = pd.to_datetime(df["Дата проводки"], errors="coerce").dt.date
    result["BookingDate"] = datetime.now().date()
    result["BankAccount"] = df["Кредит"].apply(extract_bank_account)
    result["InvoiceNum"] = ""
    result["InvoiceID"] = ""
    result["PaymentProvider"] = ""
    result["IsFromBailiff"] = df["Счет"].apply(extract_is_from_bailiff)
    result["CourtOrderNumber"] = df["Назначение платежа"].apply(extract_court_order_number)
    result["№ документа"] = df["№ документа"]
    result["Номер ИП"] = df["Назначение платежа"].apply(extract_ip_number)
    #result["ФИО"] = df["Счет"].apply(extract_fio_from_account)
    result["Назначение платежа"] = df["Назначение платежа"]  # <-- добавлено
    return result

# ===== Интерфейс Streamlit =====

st.set_page_config(page_title="Обработка выписки", layout="centered")
st.title("📄 Анализ банковской выписки")

import os
import json

# Путь к истории
history_file = "history.json"

# Загрузка истории, если она есть
history = []
if os.path.exists(history_file):
    with open(history_file, "r") as f:
        try:
            history = json.load(f)
        except:
            history = []

# Вывод истории на главной (всегда)
st.markdown("### 🧾 История обработок (последние 5):")
if history:
    for record in history[-5:][::-1]:
        court = record.get("court_count", record.get("count", 0))  # поддержка старого 'count'
        ip = record.get("ip_count", 0)

        st.markdown(
            f"• **{record['timestamp']}** — "
            f"`CourtOrderNumber`: **{court}**, "
            f"`Номер ИП`: **{ip}**, "

        )
else:
    st.markdown("_История пока пуста_")


uploaded_file = st.file_uploader("Загрузите файл выписки (Excel)", type=["xlsx", "xls"])

if uploaded_file:

    try:
        # Чтение с правильного места, ручная установка заголовков
        df_raw = pd.read_excel(uploaded_file, skiprows=2)

        # Назначим имена колонок вручную
        df_raw.columns.values[1] = "Дата проводки"
        df_raw.columns.values[4] = "Счет"
        df_raw.columns.values[6] = "Дебет"
        df_raw.columns.values[8] = "Кредит"
        df_raw.columns.values[13] = "Сумма по кредиту"
        df_raw.columns.values[14] = "№ документа"
        df_raw.columns.values[20] = "Назначение платежа"

        df = df_raw.copy()

        st.success("Файл успешно загружен и распознан!")
        df_result = process_bank_statement(df)
        # Подсчёт записей с CourtOrderNumber
        count_valid_court_numbers = df_result["CourtOrderNumber"].astype(str).str.strip().ne("").sum()
        st.info(f"🔎 Найдено записей с CourtOrderNumber: **{count_valid_court_numbers}**")
        # Подсчёт записей с Номером ИП
        count_valid_ip_number = df_result["Номер ИП"].astype(str).str.strip().ne("").sum()
        st.info(f"🔎 Найдено записей с Номер ИП: **{count_valid_ip_number}**")


        # Сохраняем результат в историю
        history.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "court_count": int(count_valid_court_numbers),
            "ip_count": int(count_valid_ip_number)

        })

        with open(history_file, "w") as f:
            json.dump(history[-20:], f, ensure_ascii=False, indent=2)

       
        st.subheader("✅ Результат")
        st.dataframe(df_result)

        output = BytesIO()
        df_result.to_excel(output, index=False, engine='openpyxl')
        st.download_button("📥 Скачать результат (Excel)", data=output.getvalue(), file_name="результат.xlsx")

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")

