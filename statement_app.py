# statement_app.py — раздел "Создание выписки"

import io
import zipfile
from copy import deepcopy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from docx import Document   # pip install python-docx


# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

def _format_date(d) -> str:
    if pd.isna(d) or d is None:
        return ""
    if isinstance(d, pd.Timestamp):
        d = d.date()
    return d.strftime("%d.%m.%Y")


def _format_amount(x) -> str:
    if pd.isna(x) or x == 0:
        return ""
    s = f"{float(x):,.2f}"          # 12345.6 -> 12,345.60
    s = s.replace(",", "X").replace(".", ",").replace("X", " ")
    return s


def _load_caseid_df(xls_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(xls_bytes), sheet_name="CaseID")


def _load_payments_raw(xls_bytes: bytes) -> pd.DataFrame:
    # двухуровневая шапка, как в твоём примере
    return pd.read_excel(io.BytesIO(xls_bytes), sheet_name="Payments", header=[0, 1])


def _flatten_payments(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.iloc[1:].copy()  # пропускаем вторую строку шапки
    df.columns = [
        "_".join([str(c) for c in col if str(c) != "nan"]).strip()
        for col in df.columns
    ]
    return df


def _prepare_payments(df_flat: pd.DataFrame) -> pd.DataFrame:
    df = df_flat.copy()
    # точные имена колонок из твоего файла
    date_col = "ДАННЫЕ ПО ВЫПИСКЕ_Дата проводки"
    reg_col = "ДАННЫЕ ПО ВЫПИСКЕ_Рег.номер"
    debit_col = "ДАННЫЕ ПО ВЫПИСКЕ_Сумма по дебету"
    credit_col = "ДАННЫЕ ПО ВЫПИСКЕ_Сумма по кредиту"

    df[reg_col] = pd.to_numeric(df[reg_col], errors="coerce").astype("Int64")
    df[debit_col] = pd.to_numeric(df[debit_col], errors="coerce").fillna(0)
    df[credit_col] = pd.to_numeric(df[credit_col], errors="coerce").fillna(0)
    df["СумаПлатежа"] = df[debit_col] + df[credit_col]

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    return df


def _fill_template_r(
    template_bytes: bytes,
    header_data: dict,
    payments_case_df: pd.DataFrame,
) -> bytes:
    """
    Сформировать один docx по шаблону Template R (для одного Рег.номера).
    """
    doc = Document(io.BytesIO(template_bytes))

    # ---------- ВЕРХНЯЯ ТАБЛИЦА ----------
    top = doc.tables[0]

    stmt_date_str = header_data["stmt_date"].strftime("%d.%m.%Y")
    stmt_time_str = header_data["stmt_time"].strftime("%H:%M:%S")

    # период в простом формате, чтобы не заморачиваться с русскими месяцами
    period_str = (
        f"за период с {header_data['period_from'].strftime('%d.%m.%Y')} "
        f"по {header_data['period_to'].strftime('%d.%m.%Y')}"
    )

    # row 0, col 0: дата (оставляем только дату)
    top.cell(0, 0).text = stmt_date_str

    # row 1, col 0: банк
    top.cell(1, 0).text = header_data["bank_name"]

    # row 2: "Дата формирования выписки ..."
    msg = f"Дата формирования выписки {stmt_date_str} в {stmt_time_str}"
    top.cell(2, 0).text = msg
    top.cell(2, 1).text = msg

    # row 3: заголовок и счёт
    top.cell(3, 0).text = "ВЫПИСКА ОПЕРАЦИЙ ПО ЛИЦЕВОМУ СЧЕТУ"
    top.cell(3, 1).text = "ВЫПИСКА ОПЕРАЦИЙ ПО ЛИЦЕВОМУ СЧЕТУ"
    top.cell(3, 2).text = header_data["account_number"]

    # row 4, col 2: наименование компании
    top.cell(4, 2).text = header_data["company_name"]

    # row 5, col 2 и 3: период
    top.cell(5, 2).text = period_str
    top.cell(5, 3).text = period_str

    # row 6, col 1: валюта
    top.cell(6, 1).text = header_data["currency"]

    # ---------- НИЖНЯЯ ТАБЛИЦА (ОПЕРАЦИИ) ----------
    tbl = doc.tables[1]

    # row 0 – заголовки, row 1 – дубли шапки, row 2 – шаблон строки
    template_row_tr = tbl.rows[2]._tr
    tbl_el = tbl._tbl

    # очищаем строки ниже 2-й, чтобы не накопилось мусора
    while len(tbl.rows) > 2:
        tbl_el.remove(tbl.rows[-1]._tr)

    date_col = "ДАННЫЕ ПО ВЫПИСКЕ_Дата проводки"
    acc_debit_col = "ДАННЫЕ ПО ВЫПИСКЕ_Счет"
    acc_credit_col = "ДАННЫЕ ПО ВЫПИСКЕ_Счет.1"
    debit_col = "ДАННЫЕ ПО ВЫПИСКЕ_Сумма по дебету"
    credit_col = "ДАННЫЕ ПО ВЫПИСКЕ_Сумма по кредиту"
    docnum_col = "ДАННЫЕ ПО ВЫПИСКЕ_№ документа"
    vo_col = "ДАННЫЕ ПО ВЫПИСКЕ_ВО"
    bank_col = "ДАННЫЕ ПО ВЫПИСКЕ_Банк (БИК и наименование)"
    purpose_col = "ДАННЫЕ ПО ВЫПИСКЕ_Назначение платежа"

    for _, r in payments_case_df.iterrows():
        new_tr = deepcopy(template_row_tr)
        tbl_el.append(new_tr)
        new_row = tbl.rows[-1]
        cells = new_row.cells

        cells[0].text = _format_date(r[date_col])
        cells[1].text = str(r.get(acc_debit_col, "") or "")
        cells[2].text = str(r.get(acc_credit_col, "") or "")
        cells[3].text = _format_amount(r.get(debit_col, 0))
        cells[4].text = _format_amount(r.get(credit_col, 0))
        cells[5].text = str(r.get(docnum_col, "") or "")
        cells[6].text = str(r.get(vo_col, "") or "")
        cells[7].text = str(r.get(bank_col, "") or "")
        cells[8].text = str(r.get(purpose_col, "") or "")

       # удаляем исходную шаблонную строку, если добавили реальные
    if len(tbl.rows) > 3:
        try:
            tbl_el.remove(tbl.rows[2]._tr)  # удаляем строку по текущему индексу 2
        except ValueError:
            # если строки уже нет в таблице — просто игнорируем
            pass


    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


def _update_caseid_with_sums(case_df: pd.DataFrame, payments_df: pd.DataFrame) -> pd.DataFrame:
    reg_col_payments = "ДАННЫЕ ПО ВЫПИСКЕ_Рег.номер"
    sums = payments_df.groupby(reg_col_payments)["СумаПлатежа"].sum()
    result = case_df.copy()
    result["Сума платежей"] = result["Рег.номер"].map(sums).fillna(0)
    return result


def _build_result_excel(case_df: pd.DataFrame, payments_raw: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        case_df.to_excel(writer, sheet_name="CaseID", index=False)
        payments_raw.to_excel(writer, sheet_name="Payments", index=False)
    out.seek(0)
    return out


# ===== ОСНОВНАЯ ФУНКЦИЯ РАЗДЕЛА =====

def run():
    st.header("Создание выписки (шаблон R)")

    uploaded_xlsx = st.file_uploader(
        "Загрузите Excel с листами CaseID и Payments",
        type=["xlsx"],
    )

    template_file = st.file_uploader(
        "Загрузите Word-шаблон выписки (Template R.docx)",
        type=["docx"],
    )

    with st.form("stmt_header_form"):
        st.subheader("Параметры шапки выписки")

        today = date.today()
        now = datetime.now().time()

        stmt_date = st.date_input("Дата формирования выписки", value=today)
        stmt_time = st.time_input("Время формирования выписки", value=now)

        bank_name = st.text_input("Банк", value="ПАО СБЕРБАНК")
        account_number = st.text_input("Номер счета", value="40702810738000100334")
        company_name = st.text_input(
            "Наименование компании",
            value='ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "РУССКИЙ ИНФОРМАЦИОННЫЙ СЕРВИС"',
        )

        period_from = st.date_input("Период с", value=today.replace(month=1, day=1))
        period_to = st.date_input("Период по", value=today)

        currency = st.text_input("Валюта", value="Российский рубль")

        submitted = st.form_submit_button("Сформировать выписки")

    if not submitted:
        return

    if uploaded_xlsx is None or template_file is None:
        st.error("Нужно загрузить и Excel, и шаблон Template R.docx.")
        return

    xls_bytes = uploaded_xlsx.read()
    template_bytes = template_file.read()

    # читаем и готовим данные
    case_df = _load_caseid_df(xls_bytes)
    payments_raw = _load_payments_raw(xls_bytes)
    payments_flat = _flatten_payments(payments_raw)
    payments = _prepare_payments(payments_flat)

    header_data = {
        "stmt_date": stmt_date,
        "stmt_time": stmt_time,
        "bank_name": bank_name,
        "account_number": account_number,
        "company_name": company_name,
        "period_from": period_from,
        "period_to": period_to,
        "currency": currency,
    }

    reg_col_payments = "ДАННЫЕ ПО ВЫПИСКЕ_Рег.номер"

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for _, row in case_df.iterrows():
            reg_num = row["Рег.номер"]
            template_code = str(row["Шалон"]).strip()  # в файле так и написано: 'Шалон'

            if template_code != "R":
                # пока работаем только с шаблоном R
                continue

            case_payments = payments[
                payments[reg_col_payments] == int(reg_num)
            ]

            if case_payments.empty:
                continue

            docx_bytes = _fill_template_r(template_bytes, header_data, case_payments)

            # Сейчас сохраняем DOCX, позже можно добавить конвертацию в PDF
            filename = f"{reg_num}.docx"
            zf.writestr(filename, docx_bytes)

    zip_buf.seek(0)

    # обновляем CaseID суммой платежей
    case_updated = _update_caseid_with_sums(case_df, payments)
    result_excel = _build_result_excel(case_updated, payments_raw)

    st.success("Выписки сформированы.")

    st.download_button(
        "⬇️ Скачать архив выписок (ZIP)",
        data=zip_buf.getvalue(),
        file_name="statements_R.zip",
        mime="application/zip",
    )

    st.download_button(
        "⬇️ Скачать обновлённый Excel (с колонкой «Сума платежей»)",
        data=result_excel.getvalue(),
        file_name=uploaded_xlsx.name.replace(".xlsx", "_with_sums.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
