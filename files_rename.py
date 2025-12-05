
import os
import csv
from pathlib import Path

from PyPDF2 import PdfReader   # pip install PyPDF2


KEYWORD = "судебный приказ"

def run():
    st.title("Переименование файлов (Судебные приказы)")
    folder = st.text_input("Укажите путь к папке:")

    if st.button("Запустить"):
        if folder and Path(folder).exists():
            process_root_folder(Path(folder))
            st.success("Готово. Проверьте CSV-отчёт в указанной папке.")
        else:
            st.error("Папка не найдена.")


def first_page_contains_keyword(pdf_path: Path, keyword: str) -> bool:
    """Проверка: есть ли на первой странице указанный текст."""
    try:
        reader = PdfReader(str(pdf_path))
        if not reader.pages:
            return False
        text = reader.pages[0].extract_text() or ""
        return keyword.lower() in text.lower()
    except Exception as e:
        print(f"Ошибка при чтении {pdf_path}: {e}")
        return False


def process_root_folder(root: Path):
    """
    Обходит все папки от root, ищет pdf-файлы с 'doc' в названии.
    Если на первой странице есть 'судебный приказ' — переименовывает
    в '<название_папки>_СП(.pdf)', и пишет отчёт CSV.
    """
    results = []  # список строк для итогового CSV

    for dirpath, dirnames, filenames in os.walk(root):
        dirpath = Path(dirpath)
        folder_name = dirpath.name
        renamed_in_folder = False

        for fname in filenames:
            lower = fname.lower()
            # ваши условия: PDF + в имени есть "doc"
            if "doc" in lower and lower.endswith(".pdf"):
                pdf_path = dirpath / fname

                if first_page_contains_keyword(pdf_path, KEYWORD):
                    base_new_name = f"{folder_name}_СП"
                    new_name = base_new_name + ".pdf"
                    new_path = dirpath / new_name

                    # если в папке уже есть такой файл — добавляем индекс
                    i = 1
                    while new_path.exists():
                        new_name = f"{base_new_name}_{i}.pdf"
                        new_path = dirpath / new_name
                        i += 1

                    print(f"Переименовываю: {pdf_path.name} -> {new_name}")
                    pdf_path.rename(new_path)
                    renamed_in_folder = True

        results.append([folder_name, "Да" if renamed_in_folder else "Нет"])

    # сохраняем отчёт
    report_path = root / "report_sp.csv"
    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Название папки", "Переименование (Да/Нет)"])
        writer.writerows(results)

    print(f"\nГотово. Отчёт: {report_path}")


if __name__ == "__main__":
    # Замените путь ниже на нужный вам корневой каталог
    root_folder = Path(r"D:\path\to\your\folders")
    process_root_folder(root_folder)
