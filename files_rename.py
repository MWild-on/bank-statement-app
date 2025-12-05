# files_rename.py

import os
import csv
from pathlib import Path

import streamlit as st
from PyPDF2 import PdfReader   # библиотека добавлена в requirements.txt

KEYWORD = "судебный приказ"


def first_page_contains_keyword(pdf_path: Path, keyword: str) -> bool:
    """Проверка: есть ли на первой странице указанный текст."""
    try:
        reader = PdfReader(str(pdf_path))
        if not reader.pages:
            return False
        text = reader.pages[0].extract_text() or ""
        return keyword.lower() in text.lower()
    except Exception as e:
        # лог выведем в интерфейс Streamlit
        st.write(f"Ошибка при чтении {pdf_path.name}: {e}")
        return False


def process_root_folder(root: Path):
    """
    Обходит все папки от root, ищет pdf-файлы с 'doc' в названии.
    Если на первой странице есть 'судебный приказ' — переименовывает
    в '<название_папки>_СП(.pdf)' и формирует отчёт.
    Возвращает (results, report_path).
    """
    results = []  # список [folder_name, "Да"/"Нет"]

    for dirpath, dirnames, filenames in os.walk(root):
        dirpath = Path(dirpath)
        folder_name = dirpath.name
        renamed_in_folder = False

        for fname in filenames:
            lower = fname.lower()
            # условия: PDF + в имени есть "doc"
            if "doc" in lower and lower.endswith(".pdf"):
                pdf_path = dirpath / fname

                if first_page_contains_keyword(pdf_path, KEYWORD):
                    base_new_name = f"{folder_name}_СП"
                    new_name = base_new_name + ".pdf"
                    new_path = dirpath / new_name

                    # если имя занято — добавляем индекс
                    i = 1
                    while new_path.exists():
                        new_name = f"{base_new_name}_{i}.pdf"
                        new_path = dirpath / new_name
                        i += 1

                    pdf_path.rename(new_path)
                    renamed_in_folder = True
                    st.write(f"Переименовано в папке '{folder_name}': {fname} → {new_name}")

        results.append([folder_name, "Да" if renamed_in_folder else "Нет"])

    # сохраняем отчёт в корне
    report_path = root / "report_sp.csv"
    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Название папки", "Переименование (Да/Нет)"])
        writer.writerows(results)

    return results, report_path


def run():
    """UI-обёртка для Streamlit."""
    st.title("Переименование файлов (судебные приказы)")

    st.markdown(
        "Инструмент обходит все подкаталоги выбранной папки, "
        "ищет PDF-файлы с `doc` в имени и проверяет первую страницу на текст "
        "`судебный приказ`. Если найдено — файл переименуется в "
        "`<название_папки>_СП(.pdf)`."
    )

    root_str = st.text_input(
        "Укажите путь к корневой папке на сервере (там, где лежат папки с PDF):",
        value=".",
        help="Например: /mount/src/w001-app/data или D:\\\\docs",
    )

    if st.button("Запустить обработку"):
        if not root_str:
            st.error("Укажите путь к папке.")
            return

        root = Path(root_str)
        if not root.exists():
            st.error(f"Папка не найдена: {root}")
            return

        with st.spinner("Обработка папок..."):
            results, report_path = process_root_folder(root)

        st.success("Готово.")
        st.write(f"CSV-отчёт сохранён в файле: `{report_path}`")

        # выводим мини-табличку результата
        st.subheader("Результат по папкам")
        st.table(
            [{"Название папки": r[0], "Переименование": r[1]} for r in results]
        )
