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

# -----------------------------
#   Шрифты
# -----------------------------

BASE_DIR = Path(__file__).resolve().parent

pdfmetrics.registerFont(TTFont("DejaVuSans", str(BASE_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", str(BASE_DIR / "DejaVuSans-Bold.ttf")))

pdfmetrics.registerFontFamily(
    "DejaVuSans",
    normal="DejaVuSans",
    bold="DejaVuSans-Bold",
)

FONT_NAME = "DejaVuSans"
FONT_NAME_BOLD = "DejaVuSans-Bold"


# -----------------------------
#   Форматирование
# -----------------------------

def fmt_money(value: Decimal | float | int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))

    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")

    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{value:,.2f}".replace(",", " ").replace(".", ",")
    return f"{s} руб."


def fmt_plain(value: Decimal | float | int) -> str:
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
    return "" if d is None else d.strftime("%d.%m.%Y")


# -----------------------------
#   Вспомогательные функции
# -----------------------------

def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def next_month(year: int, month: int):
    if month == 12:
        return year + 1, 1
    return year, month + 1


# -----------------------------
#   Загрузка CPI
# -----------------------------

def load_cpi_from_sheet(xls: pd.ExcelFile, sheet_name: str = "Индекс"):
    df = pd.read_excel(xls, sheet_name=sheet_name)
    cpi_map = {}

    for _, row in df.iterrows():
        y = int(row["Год"])
        m = int(row["Месяц"])
        idx_val = Decimal(str(row["Индексы потребительских цен"]))
        cpi_map[(y, m)] = idx_val

    return cpi_map, df


# -----------------------------
#   Модель расчёта как в Word
# -----------------------------

def monthly_factor(cpi_value: Decimal) -> Decimal:
    """ИПЦ 100.84 → 1.0084"""
    return (cpi_value / Decimal("100")).quantize(Decimal("0.0001"))


def compute_period_exact(amount: Decimal,
                         start_date: dt.date,
                         end_date: dt.date,
                         cpi_map) -> tuple[Decimal, list]:
    """
    Возвращает:
        total_increment — итог индексации
        monthly — разбивка по месяцам:
            [
                {
                    "year": ...,
                    "month": ...,
                    "days": ...,
                    "total_days": ...,
                    "cpi": Decimal,
                    "factor": Decimal,
                    "increment": Decimal
                }
            ]
    """
    if amount <= 0 or start_date > end_date:
        return Decimal("0.00"), []

    results = []
    curr_amount = amount
    s = start_date
    e = end_date
    y, m = s.year, s.month

    total_increment = Decimal("0.00")

    while True:
        dim = days_in_month(y, m)
        cpi = cpi_map[(y, m)]
        factor_full = monthly_factor(cpi)

        start_d = s.day if (y == s.year and m == s.month) else 1
        end_d = e.day if (y == e.year and m == e.month) else dim

        days = end_d - start_d + 1
        prop = Decimal(days) / Decimal(dim)

        effective_factor = Decimal("1.0") + (factor_full - Decimal("1.0")) * prop

        new_amount = (curr_amount * effective_factor).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        increment = new_amount - curr_amount

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
        curr_amount = new_amount

        if y == e.year and m == e.month:
            break

        y, m = next_month(y, m)

    return total_increment, results

# ============================================================
#      ЛОГИКА РАСЧЁТА ДОЛГА (РАЗБИВКА ПО ПЕРИОДАМ)
#      Полностью соответствует Word-методике
# ============================================================

def compute_month_factor(cpi_value: Decimal, prop: Decimal) -> Decimal:
    """
    Коэффициент месяца:
    F = 1 + ((ИПЦ - 100) / 100) * prop
    """
    delta = (cpi_value - Decimal("100")) / Decimal("100")
    return Decimal("1.0") + delta * prop


def compute_indexation_monthly(amount: Decimal,
                                start_date: dt.date,
                                end_date: dt.date,
                                cpi_map) -> Decimal:
    """
    Индексация за период [start_date; end_date] с помесячным расчётом.
    Пропорция для неполных месяцев = (число дней / дней в месяце).
    """
    y, m = start_date.year, start_date.month
    curr = start_date
    total_factor = Decimal("1.0")

    while True:
        dim = days_in_month(y, m)
        cpi = cpi_map[(y, m)]

        # один месяц
        if curr.year == end_date.year and curr.month == end_date.month and curr == start_date:
            prop = Decimal(end_date.day - start_date.day + 1) / Decimal(dim)
            F = compute_month_factor(cpi, prop)
            total_factor *= F
            break

        # первый месяц
        if curr.year == start_date.year and curr.month == start_date.month:
            prop = Decimal(dim - start_date.day + 1) / Decimal(dim)
            F = compute_month_factor(cpi, prop)
            total_factor *= F

        # последний месяц
        elif y == end_date.year and m == end_date.month:
            prop = Decimal(end_date.day) / Decimal(dim)
            F = compute_month_factor(cpi, prop)
            total_factor *= F
            break

        # полный месяц
        else:
            F = compute_month_factor(cpi, Decimal("1"))
            total_factor *= F

        # переход к следующему месяцу
        y, m = next_month(y, m)
        curr = dt.date(y, m, 1)

    return (amount * total_factor - amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def compute_period_indexation(amount: Decimal,
                              start_date: dt.date,
                              end_date: dt.date,
                              cpi_map) -> Decimal:
    """
    Обёртка над compute_indexation_monthly
    """
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
    Формирует список периодов расчёта Word-формата:
    [
        {
            period_start,
            period_end,
            debt_before,
            payment_date,
            payment_amount,
            debt_after_payment,
            indexation
        }
    ]
    """

    pays = payments_df[payments_df["Рег. номер"] == reg_num].copy()
    pays["date"] = pd.to_datetime(pays["Дата платежа"]).dt.date
    pays = pays[pays["date"] <= cutoff_date]
    pays = pays[pays["Сумма платежа"] > 0]
    pays = pays.sort_values("date")

    periods = []
    remaining = base_amount
    current_start = order_date

    # --------------------------
    #   Если платежей нет
    # --------------------------
    if pays.empty:
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

    # --------------------------
    #   Если платежи есть
    # --------------------------
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

    # --------------------------
    #   Хвост после последнего платежа
    # --------------------------
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

    # поля выровнены под Word
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

    # ------------------------------
    #  Стили
    # ------------------------------

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
        leading=15,
        spaceAfter=2,
    )

    style_value = ParagraphStyle(
        "Value",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=15,
        spaceAfter=12,
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

    # ============================================================
    #                         Заголовок
    # ============================================================

    story.append(Paragraph(
        "Расчёт индексации присуждённых денежных сумм!",
        style_title
    ))

    story.append(Spacer(1, 6))

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    base_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))
    total_days = (cutoff_date - order_date).days + 1

    def labeled(label: str, value: str):
        story.append(Paragraph(label, style_label))
        story.append(Paragraph(value, style_value))

    # ------------------------------
    # Основные параметры
    # ------------------------------
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

    # ============================================================
    #                  Порядок расчёта
    # ============================================================

    story.append(Paragraph("Порядок расчёта:", style_h2))

    story.append(Paragraph(
        "Индексация рассчитывается помесячно.<br/>"
        "В неполных месяцах применяется пропорция дней:<br/>"
        "Формула: <b>1 + ((ИПЦ - 100) / 100) × пропорция</b>",
        style_text
    ))

    story.append(Spacer(1, 12))

    # ============================================================
    #                Вывод Периодов
    # ============================================================

    for i, p in enumerate(periods, start=1):

        story.append(Paragraph(
            f"Индексация за период: {fmt_date(p['period_start'])} – {fmt_date(p['period_end'])}",
            style_h2
        ))

        # Платёж
        if p["payment_date"] and p["payment_amount"] > 0:
            story.append(Paragraph(
                f"Платёж: {fmt_date(p['payment_date'])}, сумма {fmt_money(p['payment_amount'])}",
                style_text
            ))
        else:
            story.append(Paragraph("Платёж в данном периоде отсутствует", style_text))

        # Остаток долга
        story.append(Paragraph(
            f"Остаток долга на начало периода: {fmt_money(p['debt_before'])}",
            style_text
        ))

        story.append(Paragraph(
            f"Индексация за период: {fmt_money(p['indexation'])}",
            style_text
        ))

        story.append(Spacer(1, 10))

    # ============================================================
    #                      Итоговая строка
    # ============================================================

    story.append(Paragraph(
        f"Итоговая индексация: {fmt_money(total_indexation)}",
        style_h2
    ))

    # ============================================================
    #                     Сборка PDF
    # ============================================================

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

import io
import calendar
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP
import zipfile
import pandas as pd
import streamlit as st

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors

from pathlib import Path

# ============================================================
#           ШРИФТЫ
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

pdfmetrics.registerFont(TTFont("DejaVuSans", str(BASE_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", str(BASE_DIR / "DejaVuSans-Bold.ttf")))

pdfmetrics.registerFontFamily(
    "DejaVuSans",
    normal="DejaVuSans",
    bold="DejaVuSans-Bold"
)

FONT_NAME = "DejaVuSans"
FONT_NAME_BOLD = "DejaVuSans-Bold"


# ============================================================
#              ФОРМАТИРОВАНИЕ
# ============================================================

def fmt_money(value: Decimal | float | int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    s = f"{value:,.2f}"
    s = s.replace(",", " ").replace(".", ",")
    return f"{s} руб."


def fmt_plain(value: Decimal | float | int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{value:.2f}"
    integer, frac = s.split(".")
    integer = "{:,}".format(int(integer)).replace(",", " ")
    return f"{integer},{frac}"


def fmt_date(d: dt.date | None) -> str:
    return d.strftime("%d.%m.%Y") if d else ""


# ============================================================
#              ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def next_month(year: int, month: int):
    return (year + 1, 1) if month == 12 else (year, month + 1)


def load_cpi_from_sheet(xls: pd.ExcelFile, sheet_name="Индекс"):
    df = pd.read_excel(xls, sheet_name=sheet_name)
    cpi_map = {
        (int(row["Год"]), int(row["Месяц"])): Decimal(str(row["Индексы потребительских цен"]))
        for _, row in df.iterrows()
    }
    return cpi_map, df


# ============================================================
#              ПОМЕСЯЧНЫЙ РАСЧЁТ (как в Word)
# ============================================================

def compute_month_factor(cpi_value: Decimal, prop: Decimal) -> Decimal:
    delta = (cpi_value - Decimal("100")) / Decimal("100")
    return Decimal("1") + delta * prop


def compute_indexation_monthly(amount: Decimal,
                               start_date: dt.date,
                               end_date: dt.date,
                               cpi_map) -> Decimal:
    y, m = start_date.year, start_date.month
    curr = start_date
    total_factor = Decimal("1.0")

    while True:
        dim = days_in_month(y, m)
        cpi = cpi_map[(y, m)]

        if y == start_date.year and m == start_date.month and y == end_date.year and m == end_date.month:
            prop = Decimal(end_date.day - start_date.day + 1) / Decimal(dim)
            total_factor *= compute_month_factor(cpi, prop)
            break

        if y == start_date.year and m == start_date.month:
            prop = Decimal(dim - start_date.day + 1) / Decimal(dim)
            total_factor *= compute_month_factor(cpi, prop)

        elif y == end_date.year and m == end_date.month:
            prop = Decimal(end_date.day) / Decimal(dim)
            total_factor *= compute_month_factor(cpi, prop)
            break

        else:
            total_factor *= compute_month_factor(cpi, Decimal("1"))

        y, m = next_month(y, m)
        curr = dt.date(y, m, 1)

    return (amount * total_factor - amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def compute_period_indexation(amount, start_date, end_date, cpi_map):
    if amount <= 0 or start_date > end_date:
        return Decimal("0.00")
    return compute_indexation_monthly(amount, start_date, end_date, cpi_map)


def compute_debt_periods(reg_num, order_date, base_amount, payments_df, cpi_map, cutoff_date):
    pays = payments_df[payments_df["Рег. номер"] == reg_num].copy()
    pays["date"] = pd.to_datetime(pays["Дата платежа"]).dt.date
    pays = pays[(pays["date"] <= cutoff_date) & (pays["Сумма платежа"] > 0)]
    pays = pays.sort_values("date")

    periods = []
    remaining = base_amount
    current_start = order_date

    if pays.empty:
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
#              PDF-ГЕНЕРАЦИЯ (новая)
# ============================================================

def generate_pdf_bytes_for_debt(reg_num, main_row, total_indexation, periods, cutoff_date):
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=25*mm,
        rightMargin=25*mm,
        topMargin=18*mm,
        bottomMargin=18*mm,
    )

    story = []
    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        "Title", parent=styles["Normal"],
        fontName=FONT_NAME_BOLD, fontSize=14,
        leading=18, alignment=TA_CENTER, spaceAfter=12
    )

    style_label = ParagraphStyle(
        "Label", parent=styles["Normal"],
        fontName=FONT_NAME_BOLD, fontSize=11,
        leading=15, spaceAfter=2
    )

    style_value = ParagraphStyle(
        "Value", parent=styles["Normal"],
        fontName=FONT_NAME, fontSize=11,
        leading=15, spaceAfter=10
    )

    style_text = ParagraphStyle(
        "Text", parent=styles["Normal"],
        fontName=FONT_NAME, fontSize=11,
        leading=15, spaceAfter=6
    )

    style_h2 = ParagraphStyle(
        "H2", parent=styles["Normal"],
        fontName=FONT_NAME_BOLD, fontSize=12,
        leading=16, spaceAfter=10
    )

    # Заголовок
    story.append(Paragraph("Расчёт индексации присуждённых денежных сумм", style_title))

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    base_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))
    total_days = (cutoff_date - order_date).days + 1

    def labeled(label, value):
        story.append(Paragraph(label, style_label))
        story.append(Paragraph(value, style_value))

    labeled(
        f"Взысканная сумма на дату начала периода индексации ({fmt_date(order_date)}):",
        fmt_money(base_debt)
    )

    labeled(
        "Период индексации:",
        f"{fmt_date(order_date)} – {fmt_date(cutoff_date)} ({total_days} дней)"
    )

    labeled("Регион:", "Российская Федерация")
    labeled("Сумма индексации:", fmt_money(total_indexation))

    story.append(Spacer(1, 12))

    # Порядок расчёта
    story.append(Paragraph("Порядок расчёта:", style_h2))
    story.append(Paragraph(
        "Индексация рассчитывается помесячно с применением пропорции дней<br/>"
        "для первого и последнего месяцев периода.",
        style_text
    ))

    story.append(Spacer(1, 12))

    # Периоды
    for p in periods:
        story.append(Paragraph(
            f"Индексация за период: {fmt_date(p['period_start'])} – {fmt_date(p['period_end'])}",
            style_h2
        ))

        if p["payment_date"]:
            story.append(Paragraph(
                f"Платёж: {fmt_date(p['payment_date'])}, сумма {fmt_money(p['payment_amount'])}",
                style_text
            ))
        else:
            story.append(Paragraph("Платёж в данном периоде отсутствует", style_text))

        story.append(Paragraph(
            f"Остаток долга на начало периода: {fmt_money(p['debt_before'])}",
            style_text
        ))

        story.append(Paragraph(
            f"Индексация за период: {fmt_money(p['indexation'])}",
            style_text
        ))

        story.append(Spacer(1, 10))

    # Итог
    story.append(Paragraph(
        f"Итоговая индексация: {fmt_money(total_indexation)}",
        style_h2
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
#        ОБРАБОТКА EXCEL → расчёт → Excel + ZIP(PDF)
# ============================================================

def process_workbook(uploaded_file, cutoff_date):
    xls = pd.ExcelFile(uploaded_file)

    main_df = pd.read_excel(xls, sheet_name="Основной")
    payments_df = pd.read_excel(xls, sheet_name="Платежи")
    cpi_map, cpi_df = load_cpi_from_sheet(xls)

    max_year = int(cpi_df["Год"].max())
    max_month = int(cpi_df["Месяц"].max())
    last_cpi_date = dt.date(max_year, max_month, days_in_month(max_year, max_month))

    effective_cutoff = min(cutoff_date, last_cpi_date)

    new_index_col = []
    pdf_files = []

    for _, row in main_df.iterrows():
        reg_num = int(row["Рег номер"])
        base_amount = Decimal(str(row["Сумма платежей с декабря 2024"]))
        order_date = pd.to_datetime(row["Дата вынесения приказа"]).date()

        periods = compute_debt_periods(
            reg_num, order_date, base_amount,
            payments_df, cpi_map, effective_cutoff
        )

        total_indexation = sum(p["indexation"] for p in periods)
        new_index_col.append(float(total_indexation))

        pdf_bytes = generate_pdf_bytes_for_debt(
            reg_num, row, Decimal(str(total_indexation)), periods, effective_cutoff
        )
        pdf_files.append((f"{reg_num}.pdf", pdf_bytes))

    main_df["Сумма индексации (расчёт)"] = new_index_col

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        main_df.to_excel(writer, "Основной", index=False)
        payments_df.to_excel(writer, "Платежи", index=False)
        cpi_df.to_excel(writer, "Индекс", index=False)

    excel_buffer.seek(0)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, bytes_data in pdf_files:
            zf.writestr(name, bytes_data)

    zip_buffer.seek(0)

    return excel_buffer.getvalue(), zip_buffer.getvalue(), main_df, effective_cutoff


# ============================================================
#                 STREAMLIT UI
# ============================================================

def run():
    st.title("Расчёт индексации присуждённых денежных сумм")

    today = dt.date.today()
    default_cutoff = today.replace(day=1) - dt.timedelta(days=1)

    cutoff_date = st.date_input(
        "Крайняя дата для расчёта",
        value=default_cutoff,
        format="DD.MM.YYYY"
    )

    uploaded_file = st.file_uploader("Загрузите файл .xlsx", type=["xlsx"])

    if uploaded_file is not None:
        st.success("Файл загружен.")

        if st.button("Рассчитать индексацию"):
            try:
                excel_bytes, zip_bytes, df, eff_cutoff = process_workbook(
                    uploaded_file, cutoff_date
                )

                st.info(f"Фактическая крайняя дата расчёта: {eff_cutoff}")

                st.dataframe(df[["Рег номер", "Сумма платежей с декабря 2024", "Сумма индексации (расчёт)"]])

                st.download_button(
                    "Скачать Excel",
                    excel_bytes,
                    "индексация.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    "Скачать ZIP (PDF)",
                    zip_bytes,
                    "reports.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"Ошибка: {e}")
