import io
import calendar
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP
import zipfile

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------- Шрифт с кириллицей ----------

FONT_NAME = "DejaVuSans"
pdfmetrics.registerFont(TTFont(FONT_NAME, "DejaVuSans.ttf"))


# ---------- Форматирование ----------

def fmt_money(value: Decimal | float | int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    # обнуляем микроскопические остатки типа -1.3E-12
    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{value:.2f}"
    parts = s.split(".")
    parts[0] = "{:,}".format(int(parts[0])).replace(",", " ")
    return f"{parts[0]}.{parts[1]} руб."


def fmt_date(d: dt.date | None) -> str:
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")


# =========================
#  Вспомогательные функции
# =========================

def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def next_month(year: int, month: int):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def load_cpi_from_sheet(xls: pd.ExcelFile, sheet_name: str = "Индекс"):
    """
    Лист 'Индекс':
      Год | Месяц | Индексы потребительских цен
    """
    df = pd.read_excel(xls, sheet_name=sheet_name)
    cpi_map = {}
    for _, row in df.iterrows():
        y = int(row["Год"])
        m = int(row["Месяц"])
        idx_val = Decimal(str(row["Индексы потребительских цен"]))
        cpi_map[(y, m)] = idx_val
    return cpi_map, df


def compute_indexation_for_period(
    amount: Decimal,
    start_date: dt.date,
    end_date: dt.date,
    cpi_map,
) -> Decimal:
    """
    Индексация суммы `amount` за период [start_date, end_date] включительно.
    - первый и последний месяц пропорционально дням;
    - месяцы между ними берутся целиком.
    """
    if amount <= 0 or start_date > end_date:
        return Decimal("0.00")

    product = Decimal("1.0")
    s = start_date
    e = end_date

    y, m = s.year, s.month

    while True:
        key = (y, m)
        if key not in cpi_map:
            raise ValueError(
                f"Нет ИПЦ для {y}-{m:02d} — добавь этот месяц на лист 'Индекс'."
            )

        cpi = cpi_map[key]  # например 100.84
        delta = cpi - Decimal("100")
        dim = days_in_month(y, m)

        # Период целиком в одном месяце
        if y == s.year and m == s.month and y == e.year and m == e.month:
            days = e.day - s.day + 1
            prop = Decimal(days) / Decimal(dim)
            factor = (delta * prop + Decimal("100")) / Decimal("100")
            product *= factor
            break

        # Первый месяц (не последний)
        elif y == s.year and m == s.month:
            days = dim - s.day + 1
            prop = Decimal(days) / Decimal(dim)
            factor = (delta * prop + Decimal("100")) / Decimal("100")
            product *= factor

        # Последний месяц (не первый)
        elif y == e.year and m == e.month:
            days = e.day
            prop = Decimal(days) / Decimal(dim)
            factor = (delta * prop + Decimal("100")) / Decimal("100")
            product *= factor
            break

        # Полный месяц внутри периода
        else:
            factor = cpi / Decimal("100")
            product *= factor

        y, m = next_month(y, m)

    ind = (amount * product - amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return ind


# =========================
#  Расчёт по одному долгу
# =========================

def calculate_indexation_for_debt(
    reg_num: int,
    main_row: pd.Series,
    payments_df: pd.DataFrame,
    cpi_map,
    cutoff_date: dt.date,
):
    """
    По одному долгу:
    - индексация считается только до cutoff_date (включительно);
    - по каждому платежу отдельный период;
    - если после последнего платежа до cutoff_date есть дни,
      считаем индексацию на остаток долга до cutoff_date.
    """

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    initial_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))

    pays = payments_df[payments_df["Рег. номер"] == reg_num].copy()
    if pays.empty:
        # только хвост до cutoff_date, если приказ раньше
        if order_date <= cutoff_date:
            ind = compute_indexation_for_period(
                initial_debt, order_date, cutoff_date, cpi_map
            )
            return ind, [{
                "period_start": order_date,
                "period_end": cutoff_date,
                "debt_before": initial_debt,
                "indexation": ind,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": initial_debt,
            }]
        return Decimal("0.00"), []

    # даты платежей
    pays["date"] = pd.to_datetime(pays["Дата платежа"]).dt.date
    # учитываем только платежи не позже крайней даты
    pays = pays[pays["date"] <= cutoff_date]
    pays = pays[pays["Сумма платежа"] > 0]

    if pays.empty:
        # нет платежей в периоде расчёта – одна сплошная индексация до cutoff_date
        if order_date <= cutoff_date:
            ind = compute_indexation_for_period(
                initial_debt, order_date, cutoff_date, cpi_map
            )
            return ind, [{
                "period_start": order_date,
                "period_end": cutoff_date,
                "debt_before": initial_debt,
                "indexation": ind,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": initial_debt,
            }]
        return Decimal("0.00"), []

    grouped = (
        pays.groupby("date")["Сумма платежа"]
        .sum()
        .reset_index()
        .sort_values("date")
    )

    total_indexation = Decimal("0.00")
    remaining_debt = initial_debt
    current_start = order_date
    periods = []

    # периоды к платежам
    for _, row in grouped.iterrows():
        pay_date = row["date"]
        pay_amount = Decimal(str(row["Сумма платежа"]))

        if remaining_debt <= 0:
            break
        if current_start > cutoff_date:
            break

        period_end = min(pay_date, cutoff_date)

        ind = compute_indexation_for_period(
            remaining_debt, current_start, period_end, cpi_map
        )
        total_indexation += ind

        remaining_after_payment = remaining_debt - pay_amount

        periods.append(
            {
                "period_start": current_start,
                "period_end": period_end,
                "debt_before": remaining_debt,
                "indexation": ind,
                "payment_date": pay_date,
                "payment_amount": pay_amount,
                "debt_after_payment": remaining_after_payment,
            }
        )

        remaining_debt = remaining_after_payment
        current_start = pay_date + dt.timedelta(days=1)

    # хвост до крайней даты, если долг ещё есть
    if remaining_debt > 0 and current_start <= cutoff_date:
        ind = compute_indexation_for_period(
            remaining_debt, current_start, cutoff_date, cpi_map
        )
        total_indexation += ind
        periods.append(
            {
                "period_start": current_start,
                "period_end": cutoff_date,
                "debt_before": remaining_debt,
                "indexation": ind,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": remaining_debt,
            }
        )

    return total_indexation, periods


# =========================
#  PDF по долгу (в память)
# =========================

def generate_pdf_bytes_for_debt(
    reg_num: int,
    main_row: pd.Series,
    total_indexation: Decimal,
    periods,
    cutoff_date: dt.date,
) -> bytes:
    """
    Возвращает байты PDF для одного долга.
    Имя файла потом будет <Рег номер>.pdf в ZIP.
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    x_margin = 20 * mm
    y = height - 20 * mm
    line_step = 6 * mm

    def draw_line(text: str, bold: bool = False):
        nonlocal y
        c.setFont(FONT_NAME, 12 if bold else 10)
        # простейший перенос строки, если очень длинно
        max_chars = 95
        lines = [text[i:i+max_chars] for i in range(0, len(text), max_chars)] or [""]
        for ln in lines:
            c.drawString(x_margin, y, ln)
            y -= line_step
            if y < 20 * mm:
                c.showPage()
                y = height - 20 * mm

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    base_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))

    # Заголовок
    draw_line(
        f"Расчёт индексации присуждённой денежной суммы. Рег. номер {reg_num}",
        bold=True,
    )
    draw_line("")
    draw_line(f"Крайняя дата расчёта: {fmt_date(cutoff_date)}")
    draw_line(
        f"Сумма долга на дату вынесения приказа ({fmt_date(order_date)}): "
        f"{fmt_money(base_debt)}"
    )
    draw_line(f"Итоговая сумма индексации: {fmt_money(total_indexation)}")
    draw_line("")

    # Детализация по периодам
    for i, p in enumerate(periods, start=1):
        if p["payment_date"] is not None and p["payment_amount"] != Decimal("0.00"):
            draw_line(
                f"Платёж #{i}: дата {fmt_date(p['payment_date'])}, "
                f"сумма {fmt_money(p['payment_amount'])}",
                bold=True,
            )
        else:
            draw_line(f"Период без платежа #{i}", bold=True)

        draw_line(
            f"  Период индексации: {fmt_date(p['period_start'])} – "
            f"{fmt_date(p['period_end'])}"
        )
        draw_line(
            f"  Остаток долга на начало периода: {fmt_money(p['debt_before'])}"
        )
        draw_line(
            f"  Индексация за период: {fmt_money(p['indexation'])}"
        )
        draw_line(
            f"  Остаток долга после периода: {fmt_money(p['debt_after_payment'])}"
        )
        draw_line("")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# =========================
#  Основная обработка файла
# =========================

def process_workbook(uploaded_file, cutoff_date: dt.date):
    """
    Принимает загруженный Excel-файл и крайнюю дату расчёта.
    Возвращает:
      - bytes итогового Excel
      - bytes ZIP с PDF по каждому долгу
      - DataFrame с основным листом
    """
    xls = pd.ExcelFile(uploaded_file)

    main_df = pd.read_excel(xls, sheet_name="Основной")
    payments_df = pd.read_excel(xls, sheet_name="Платежи")
    cpi_map, cpi_df = load_cpi_from_sheet(xls, sheet_name="Индекс")

    # ограничим крайнюю дату последним месяцем, который есть в индексации
    max_year = int(cpi_df["Год"].max())
    max_month = int(cpi_df["Месяц"].max())
    last_cpi_date = dt.date(max_year, max_month, days_in_month(max_year, max_month))

    effective_cutoff = min(cutoff_date, last_cpi_date)

    new_index_col = []
    pdf_files = []  # (file_name, bytes)

    for _, row in main_df.iterrows():
        reg_num = int(row["Рег номер"])

        total_indexation, periods = calculate_indexation_for_debt(
            reg_num=reg_num,
            main_row=row,
            payments_df=payments_df,
            cpi_map=cpi_map,
            cutoff_date=effective_cutoff,
        )

        new_index_col.append(float(total_indexation))

        if periods:
            pdf_bytes = generate_pdf_bytes_for_debt(
                reg_num=reg_num,
                main_row=row,
                total_indexation=total_indexation,
                periods=periods,
                cutoff_date=effective_cutoff,
            )
            pdf_files.append((f"{reg_num}.pdf", pdf_bytes))

    main_df["Сумма индексации (расчёт)"] = new_index_col

    # ----- Excel в память -----
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        main_df.to_excel(writer, sheet_name="Основной", index=False)
        payments_df.to_excel(writer, sheet_name="Платежи", index=False)
        cpi_df.to_excel(writer, sheet_name="Индекс", index=False)
    excel_buffer.seek(0)
    excel_bytes = excel_buffer.getvalue()

    # ----- ZIP с PDF в память -----
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, fbytes in pdf_files:
            zf.writestr(fname, fbytes)
    zip_buffer.seek(0)
    zip_bytes = zip_buffer.getvalue()

    return excel_bytes, zip_bytes, main_df, effective_cutoff


# =========================
#  Streamlit entry point для вкладки
# =========================

def run():
    st.title("Расчёт индексации присуждённых денежных сумм")

    st.write(
        "Загрузите Excel-файл со страницами **«Основной»**, "
        "**«Платежи»** и **«Индекс»**. "
        "Крайняя дата расчёта задаётся ниже — после неё индексация не считается."
    )

    # крайняя дата по умолчанию — последнее число прошлого месяца
    today = dt.date.today()
    first_day_this_month = today.replace(day=1)
    default_cutoff = first_day_this_month - dt.timedelta(days=1)

    cutoff_date = st.date_input(
        "Крайняя дата для расчёта",
        value=default_cutoff,
        format="DD.MM.YYYY",
    )

    uploaded_file = st.file_uploader(
        "Загрузите файл формата .xlsx",
        type=["xlsx"],
    )

    if uploaded_file is not None:
        st.success("Файл загружен. Нажмите кнопку ниже для расчёта.")

        if st.button("Рассчитать индексацию"):
            try:
                excel_bytes, zip_bytes, main_df, effective_cutoff = process_workbook(
                    uploaded_file, cutoff_date
                )

                st.info(f"Фактическая крайняя дата расчёта: {effective_cutoff}")

                st.subheader("Результат (фрагмент таблицы)")
                cols_to_show = [
                    col for col in main_df.columns
                    if col in ("Рег номер", "Сумма индексации (расчёт)", "Сумма платежей с декабря 2024")
                ]
                if cols_to_show:
                    st.dataframe(main_df[cols_to_show].head(50))
                else:
                    st.dataframe(main_df.head(50))

                st.download_button(
                    label="Скачать Excel с индексацией",
                    data=excel_bytes,
                    file_name="Файл для расчета_с_индексацией.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.download_button(
                    label="Скачать ZIP с PDF-отчётами",
                    data=zip_bytes,
                    file_name="pdf_otchety.zip",
                    mime="application/zip",
                )

            except Exception as e:
                st.error(f"Ошибка при обработке файла: {e}")
