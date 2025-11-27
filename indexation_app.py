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
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# Регистрация обычного и жирного шрифта
pdfmetrics.registerFont(TTFont("DejaVuSans", str(BASE_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", str(BASE_DIR / "DejaVuSans-Bold.ttf")))

# Семейство — чтобы <b> / <strong> автоматически включали Bold
pdfmetrics.registerFontFamily(
    "DejaVuSans",
    normal="DejaVuSans",
    bold="DejaVuSans-Bold",
)

FONT_NAME = "DejaVuSans"
FONT_NAME_BOLD = "DejaVuSans-Bold"


# ------------------------------
#   Форматирование чисел и дат
# ------------------------------

def fmt_money(value: Decimal | float | int) -> str:
    """Форматирование вида: 12 345,67 руб."""
    if not isinstance(value, Decimal):
        value = Decimal(str(value))

    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")

    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    s = f"{value:,.2f}"
    s = s.replace(",", " ").replace(".", ",")
    return f"{s} руб."


def fmt_plain(value: Decimal | float | int) -> str:
    """Форматирование чисел для формул: 12345.67 → 12 345,67"""
    if not isinstance(value, Decimal):
        value = Decimal(str(value))

    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")

    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    s = f"{value:.2f}"
    integer, frac = s.split(".")

    integer = "{:,}".format(int(integer)).replace(",", " ")
    return f"{integer},{frac}"


def fmt_date(d: dt.date | None) -> str:
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")


# ------------------------------
#   Вспомогательные функции
# ------------------------------

def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def next_month(year: int, month: int):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def load_cpi_from_sheet(xls: pd.ExcelFile, sheet_name: str = "Индекс"):
    """Загрузка CPI словаря {(год, месяц): индекс}"""
    df = pd.read_excel(xls, sheet_name=sheet_name)
    cpi_map = {}
    for _, row in df.iterrows():
        y = int(row["Год"])
        m = int(row["Месяц"])
        idx_val = Decimal(str(row["Индексы потребительских цен"]))
        cpi_map[(y, m)] = idx_val
    return cpi_map, df

# ============================================================
#      НОВАЯ ТОЧНАЯ МЕТОДИКА РАСЧЁТА (как в Word)
# ============================================================

def monthly_factor(cpi_value: Decimal) -> Decimal:
    """Преобразует ИПЦ (например 100.84) → множитель (1.0084)"""
    return (cpi_value / Decimal("100")).quantize(Decimal("0.0001"))


def compute_period_exact(amount: Decimal,
                         start_date: dt.date,
                         end_date: dt.date,
                         cpi_map) -> tuple[Decimal, list]:
    """
    Точный расчёт индексации суммы за период.
    Возвращает:
        - итоговая индексация
        - список разбивки по месяцам:
            [
                {
                    'year': 2022,
                    'month': 2,
                    'days': 12,
                    'total_days': 28,
                    'cpi': Decimal('100.55'),
                    'factor': Decimal('1.0024'),
                    'increment': Decimal('152.44')
                },
                ...
            ]
    """
    if amount <= 0 or start_date > end_date:
        return Decimal("0.00"), []

    results = []
    current_amount = amount
    s = start_date
    e = end_date
    y, m = s.year, s.month
    total_increment = Decimal("0.00")

    while True:
        dim = days_in_month(y, m)
        cpi = cpi_map[(y, m)]
        factor_full = monthly_factor(cpi)

        # границы внутри месяца
        start_d = s.day if (y == s.year and m == s.month) else 1
        end_d = e.day if (y == e.year and m == e.month) else dim

        days = end_d - start_d + 1
        proportion = Decimal(days) / Decimal(dim)

        # итоговый множитель месяца
        effective_factor = (Decimal("1.0") +
                            (factor_full - Decimal("1.0")) * proportion)

        new_amount = (current_amount * effective_factor).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        increment = new_amount - current_amount

        results.append({
            "year": y,
            "month": m,
            "days": days,
            "total_days": dim,
            "cpi": cpi,
            "factor": effective_factor,
            "increment": increment
        })

        total_increment += increment
        current_amount = new_amount

        # переход к следующему месяцу
        if y == e.year and m == e.month:
            break

        y, m = next_month(y, m)

    return total_increment, results


# ============================================================
#      Расчёт одного долга через новую модель
# ============================================================

def calculate_indexation_for_debt(
    reg_num: int,
    main_row: pd.Series,
    payments_df: pd.DataFrame,
    cpi_map,
    cutoff_date: dt.date,
):
    """
    Новый расчёт:
    - каждый период (между платежами) считается через compute_period_exact(),
      а не через старую формулу.
    - общая сумма остаётся прежней
    """

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    initial_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))

    pays = payments_df[payments_df["Рег. номер"] == reg_num].copy()

    # фильтруем платежи
    pays["date"] = pd.to_datetime(pays["Дата платежа"]).dt.date
    pays = pays[pays["date"] <= cutoff_date]
    pays = pays[pays["Сумма платежа"] > 0]

    # если платежей нет → 1 большой период
    if pays.empty:
        if order_date <= cutoff_date:
            inc, months = compute_period_exact(
                initial_debt, order_date, cutoff_date, cpi_map
            )
            return inc, [{
                "period_start": order_date,
                "period_end": cutoff_date,
                "debt_before": initial_debt,
                "indexation": inc,
                "monthly_breakdown": months,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": initial_debt,
            }]
        return Decimal("0.00"), []

    # группировка по датам платежей
    grouped = (
        pays.groupby("date")["Сумма платежа"]
        .sum()
        .reset_index()
        .sort_values("date")
    )

    periods = []
    total_inc = Decimal("0.00")
    remaining = initial_debt
    current_start = order_date

    for _, pay in grouped.iterrows():
        pay_date = pay["date"]
        pay_amount = Decimal(str(pay["Сумма платежа"]))

        if remaining <= 0:
            break
        if current_start > cutoff_date:
            break

        period_end = min(pay_date, cutoff_date)

        inc, months = compute_period_exact(
            remaining, current_start, period_end, cpi_map
        )

        total_inc += inc

        after_payment = remaining - pay_amount

        periods.append({
            "period_start": current_start,
            "period_end": period_end,
            "debt_before": remaining,
            "indexation": inc,
            "monthly_breakdown": months,
            "payment_date": pay_date,
            "payment_amount": pay_amount,
            "debt_after_payment": after_payment
        })

        remaining = after_payment
        current_start = pay_date + dt.timedelta(days=1)

    # хвост после последнего платежа
    if remaining > 0 and current_start <= cutoff_date:
        inc, months = compute_period_exact(
            remaining, current_start, cutoff_date, cpi_map
        )
        total_inc += inc

        periods.append({
            "period_start": current_start,
            "period_end": cutoff_date,
            "debt_before": remaining,
            "indexation": inc,
            "monthly_breakdown": months,
            "payment_date": None,
            "payment_amount": Decimal("0.00"),
            "debt_after_payment": remaining
        })

    return total_inc, periods
# ============================================================
#                 PDF ГЕНЕРАЦИЯ (НОВАЯ)
# ============================================================

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors


def generate_pdf_bytes_for_debt(
    reg_num: int,
    main_row: pd.Series,
    total_indexation: Decimal,
    periods,
    cutoff_date: dt.date,
) -> bytes:

    buffer = io.BytesIO()

    # поля уже автоматически нормализованы
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=25*mm,
        rightMargin=25*mm,
        topMargin=18*mm,
        bottomMargin=18*mm,
        title=f"Расчёт индексации {reg_num}",
    )

    story = []
    styles = getSampleStyleSheet()

    # --- Стили ---
    style_title = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontName=FONT_NAME_BOLD,
        fontSize=14,
        leading=18,
        alignment=TA_CENTER,
        spaceAfter=12,
    )

    style_label = ParagraphStyle(
        "Label",
        parent=styles["Normal"],
        fontName=FONT_NAME_BOLD,
        fontSize=11,
        leading=14,
        spaceAfter=4,
    )

    style_value = ParagraphStyle(
        "Value",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=14,
        spaceAfter=10,
    )

    style_text = ParagraphStyle(
        "Text",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=15,
        spaceAfter=8,
    )

    style_h2 = ParagraphStyle(
        "H2",
        parent=styles["Normal"],
        fontName=FONT_NAME_BOLD,
        fontSize=12,
        leading=16,
        spaceAfter=10,
    )

    # ---------------------------------------------------------
    #                      Заголовок
    # ---------------------------------------------------------

    story.append(Paragraph(
        "Расчёт индексации присуждённых денежных сумм",
        style_title
    ))
    story.append(Spacer(1, 6))

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    base_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))
    total_days = (cutoff_date - order_date).days + 1

    # ------------------------------
    # Блок основных параметров
    # ------------------------------

    def labeled(label: str, value: str):
        story.append(Paragraph(f"{label}", style_label))
        story.append(Paragraph(value, style_value))

    labeled(
        f"Взысканная сумма на дату начала периода индексации ({fmt_date(order_date)}):",
        fmt_money(base_debt)
    )

    labeled(
        "Период индексации:",
        f"{fmt_date(order_date)} – {fmt_date(cutoff_date)} ({total_days} дней)"
    )

    labeled(
        "Регион:",
        "Российская Федерация"
    )

    labeled(
        "Сумма индексации:",
        fmt_money(total_indexation)
    )

    story.append(Spacer(1, 12))

    # ---------------------------------------------------------
    #                Порядок расчёта (как в Word)
    # ---------------------------------------------------------

    story.append(Paragraph("Порядок расчёта:", style_h2))

    story.append(Paragraph(
        "Индексация рассчитывается отдельно по каждому месяцу периода.<br/>"
        "Для полного месяца применяется множитель (ИПЦ / 100).<br/>"
        "Для неполного месяца ИПЦ умножается на долю дней в месяце.",
        style_text
    ))

    story.append(Spacer(1, 12))


    # ---------------------------------------------------------
    #             Вывод всех периодов и месяцев
    # ---------------------------------------------------------

    for i, p in enumerate(periods, start=1):

        # Заголовок периода
        story.append(Paragraph(
            f"Индексация за период: {fmt_date(p['period_start'])} – {fmt_date(p['period_end'])}",
            style_h2
        ))

        # Если есть платеж
        if p["payment_date"] and p["payment_amount"] != Decimal("0.00"):
            story.append(Paragraph(
                f"Платёж {fmt_date(p['payment_date'])}: {fmt_money(p['payment_amount'])}",
                style_text
            ))
        else:
            story.append(Paragraph(
                "Платёж в данном периоде отсутствует",
                style_text
            ))

        # Остаток долга
        story.append(Paragraph(
            f"Остаток долга на начало периода: {fmt_money(p['debt_before'])}",
            style_text
        ))

        # ------------------------------
        # Таблица помесячной индексации
        # ------------------------------

        table_data = [
            ["Месяц", "Дней", "ИПЦ", "Множитель", "Прирост"]
        ]

        for m in p["monthly_breakdown"]:
            dt_display = f"{m['month']:02d}.{m['year']}"
            table_data.append([
                dt_display,
                f"{m['days']} / {m['total_days']}",
                f"{m['cpi']}",
                f"{m['factor']}",
                fmt_money(m['increment'])
            ])

        tbl = Table(
            table_data,
            colWidths=[30*mm, 22*mm, 25*mm, 28*mm, 30*mm]
        )

        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ]))

        story.append(tbl)
        story.append(Spacer(1, 8))

        story.append(Paragraph(
            f"Индексация за период: {fmt_money(p['indexation'])}",
            style_text
        ))

        story.append(Spacer(1, 12))


    # ---------------------------------------------------------
    #                          Итог
    # ---------------------------------------------------------

    story.append(Paragraph(
        f"Итоговая индексация: {fmt_money(total_indexation)}",
        style_h2
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================
#       НОВАЯ ЛОГИКА — ПОМЕСЯЧНЫЙ РАСЧЁТ КАК В WORD
# ============================================================

def compute_month_factor(cpi_value: Decimal, prop: Decimal) -> Decimal:
    """
    Рассчитывает коэффициент месяца:
    factor = 1 + ((ИПЦ - 100) / 100) * prop
    """
    delta = (cpi_value - Decimal("100")) / Decimal("100")
    return Decimal("1") + delta * prop


def compute_indexation_monthly(amount: Decimal, start_date: dt.date, end_date: dt.date, cpi_map) -> Decimal:
    """
    Индексация за период [start_date; end_date] помесячно.
    Пропорция первого и последнего месяца = (число дней / дней в месяце).
    """
    y, m = start_date.year, start_date.month
    curr = start_date
    total_factor = Decimal("1.0")

    while True:
        dim = days_in_month(y, m)
        cpi = cpi_map[(y, m)]

        # месяц полностью внутри периода
        month_first = dt.date(y, m, 1)
        month_last = dt.date(y, m, dim)

        if curr == start_date and curr.year == end_date.year and curr.month == end_date.month:
            # один месяц
            prop = Decimal(end_date.day - start_date.day + 1) / Decimal(dim)
            F = compute_month_factor(cpi, prop)
            total_factor *= F
            break

        if curr.year == start_date.year and curr.month == start_date.month:
            # первый месяц
            prop = Decimal(dim - start_date.day + 1) / Decimal(dim)
            F = compute_month_factor(cpi, prop)
            total_factor *= F

        elif y == end_date.year and m == end_date.month:
            # последний месяц
            prop = Decimal(end_date.day) / Decimal(dim)
            F = compute_month_factor(cpi, prop)
            total_factor *= F
            break

        else:
            # полный месяц
            F = compute_month_factor(cpi, Decimal("1"))
            total_factor *= F

        # переход к следующему месяцу
        y, m = next_month(y, m)
        curr = dt.date(y, m, 1)

    return (amount * total_factor - amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def compute_period_indexation(amount: Decimal, start_date: dt.date, end_date: dt.date, cpi_map) -> Decimal:
    """ Обёртка над compute_indexation_monthly() """
    if amount <= 0 or start_date > end_date:
        return Decimal("0.00")
    return compute_indexation_monthly(amount, start_date, end_date, cpi_map)


def compute_debt_periods(
    reg_num: int,
    order_date: dt.date,
    base_amount: Decimal,
    payments_df: pd.DataFrame,
    cpi_map,
    cutoff_date: dt.date,
):
    """
    Генерирует список периодов между платежами:
    [
        {
            period_start: ...
            period_end: ...
            debt_before: ...
            payment_date: ...
            payment_amount: ...
            debt_after_payment: ...
            indexation: ...
        }
    ]
    Полностью соответствует Word-образцу.
    """

    # только платежи по текущему рег. номеру
    pays = payments_df[payments_df["Рег. номер"] == reg_num].copy()
    pays["date"] = pd.to_datetime(pays["Дата платежа"]).dt.date
    pays = pays[pays["date"] <= cutoff_date]
    pays = pays[pays["Сумма платежа"] > 0]
    pays = pays.sort_values("date")

    periods = []
    remaining = base_amount
    current_start = order_date

    if pays.empty:
        # один период — от приказа до cutoff_date
        ind = compute_period_indexation(remaining, order_date, cutoff_date, cpi_map)
        periods.append({
            "period_start": order_date,
            "period_end": cutoff_date,
            "debt_before": remaining,
            "payment_date": None,
            "payment_amount": Decimal("0.00"),
            "debt_after_payment": remaining,
            "indexation": ind,
        })
        return periods

    # если есть платежи
    for _, row in pays.iterrows():
        pay_date = row["date"]
        pay_amount = Decimal(str(row["Сумма платежа"]))

        period_end = min(pay_date, cutoff_date)

        ind = compute_period_indexation(remaining, current_start, period_end, cpi_map)

        new_remaining = remaining - pay_amount

        periods.append({
            "period_start": current_start,
            "period_end": period_end,
            "debt_before": remaining,
            "payment_date": pay_date,
            "payment_amount": pay_amount,
            "debt_after_payment": new_remaining,
            "indexation": ind,
        })

        remaining = new_remaining
        current_start = pay_date + dt.timedelta(days=1)

        if current_start > cutoff_date:
            return periods

    # хвост после последнего платежа
    if remaining > 0 and current_start <= cutoff_date:
        ind = compute_period_indexation(remaining, current_start, cutoff_date, cpi_map)

        periods.append({
            "period_start": current_start,
            "period_end": cutoff_date,
            "debt_before": remaining,
            "payment_date": None,
            "payment_amount": Decimal("0.00"),
            "debt_after_payment": remaining,
            "indexation": ind,
        })

    return periods



# ============================================================
#      PROCESS WORKBOOK: Excel → расчёт → Excel + ZIP(PDF)
# ============================================================

def process_workbook(uploaded_file, cutoff_date: dt.date):
    """
    Принимает загруженный Excel-файл и крайнюю дату расчёта.
    Возвращает:
      - bytes итогового Excel
      - bytes ZIP с PDF по каждому долгу
      - DataFrame с основным листом
      - фактическую конечную дату расчёта (ограниченную последним индексом)
    """

    xls = pd.ExcelFile(uploaded_file)

    main_df = pd.read_excel(xls, sheet_name="Основной")
    payments_df = pd.read_excel(xls, sheet_name="Платежи")
    cpi_map, cpi_df = load_cpi_from_sheet(xls, sheet_name="Индекс")

    # ограничим крайний период максимальными ИПЦ
    max_year = int(cpi_df["Год"].max())
    max_month = int(cpi_df["Месяц"].max())
    last_cpi_date = dt.date(max_year, max_month, days_in_month(max_year, max_month))

    effective_cutoff = min(cutoff_date, last_cpi_date)

    # новые столбцы (итог индексации)
    new_index_col = []
    pdf_files = []  # (filename, bytes)

    for _, row in main_df.iterrows():
        reg_num = int(row["Рег номер"])

        # -------------------------
        # новый расчёт по месяцам
        # -------------------------
        base_amount = Decimal(str(row["Сумма платежей с декабря 2024"]))
        order_date = pd.to_datetime(row["Дата вынесения приказа"]).date()

        periods = compute_debt_periods(
            reg_num=reg_num,
            order_date=order_date,
            base_amount=base_amount,
            payments_df=payments_df,
            cpi_map=cpi_map,
            cutoff_date=effective_cutoff,
        )

        total_indexation = sum(p["indexation"] for p in periods)
        new_index_col.append(float(total_indexation))

        # PDF формируем только если есть периоды
        if periods:
            pdf_bytes = generate_pdf_bytes_for_debt(
                reg_num=reg_num,
                main_row=row,
                total_indexation=Decimal(str(total_indexation)),
                periods=periods,
                cutoff_date=effective_cutoff,
            )
            pdf_files.append((f"{reg_num}.pdf", pdf_bytes))

    # -------------------------
    # Формируем Excel
    # -------------------------
    main_df["Сумма индексации (расчёт)"] = new_index_col

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        main_df.to_excel(writer, sheet_name="Основной", index=False)
        payments_df.to_excel(writer, sheet_name="Платежи", index=False)
        cpi_df.to_excel(writer, sheet_name="Индекс", index=False)
    excel_buffer.seek(0)
    excel_bytes = excel_buffer.getvalue()

    # -------------------------
    # Формируем ZIP
    # -------------------------
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, fb in pdf_files:
            zf.writestr(fname, fb)
    zip_buffer.seek(0)
    zip_bytes = zip_buffer.getvalue()

    return excel_bytes, zip_bytes, main_df, effective_cutoff


# ============================================================
#                Streamlit UI: Вкладка Индексация
# ============================================================

def run():
    st.title("Расчёт индексации присуждённых денежных сумм")

    st.write(
        "Загрузите Excel-файл со страницами **«Основной»**, "
        "**«Платежи»** и **«Индекс»**. "
        "Крайняя дата расчёта задаётся ниже — после неё индексация не считается."
    )

    # крайняя дата по умолчанию — последнее число прошлого месяца
    today = dt.date.today()
    first_day = today.replace(day=1)
    default_cutoff = first_day - dt.timedelta(days=1)

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
                excel_bytes, zip_bytes, main_df, eff_cutoff = process_workbook(
                    uploaded_file, cutoff_date
                )

                st.info(f"Фактическая крайняя дата расчёта: {eff_cutoff}")

                st.subheader("Результат:")
                st.dataframe(
                    main_df[["Рег номер", "Сумма платежей с декабря 2024", "Сумма индексации (расчёт)"]]
                )

                st.download_button(
                    label="Скачать Excel с индексацией",
                    data=excel_bytes,
                    file_name="индексация.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.download_button(
                    label="Скачать ZIP с PDF отчётами",
                    data=zip_bytes,
                    file_name="indexation_reports.zip",
                    mime="application/zip",
                )

            except Exception as e:
                st.error(f"Ошибка: {e}")
