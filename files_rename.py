import os
import csv
from pathlib import Path

from PyPDF2 import PdfReader   # внешняя библиотека

# GUI для выбора папки
import tkinter as tk
from tkinter import filedialog, messagebox

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
        print(f"Ошибка при чтении {pdf_path}: {e}")
        return False


def process_root_folder(root: Path):
    """
    Обходит все папки от root, ищет pdf-файлы с 'doc' в названии.
    Если на первой странице есть 'судебный приказ' — переименовывает
    в '<название_папки>_СП(.pdf)' и пишет отчет CSV.
    """
    results = []  # [folder_name, "Да"/"Нет"]

    for dirpath, dirnames, filenames in os.walk(root):
        dirpath = Path(dirpath)
        folder_name = dirpath.name
        renamed_in_folder = False

        for fname in filenames:
            lower = fname.lower()
            if "doc" in lower and lower.endswith(".pdf"):
                pdf_path = dirpath / fname

                if first_page_contains_keyword(pdf_path, KEYWORD):
                    base_new_name = f"{folder_name}_СП"
                    new_name = base_new_name + ".pdf"
                    new_path = dirpath / new_name

                    i = 1
                    while new_path.exists():
                        new_name = f"{base_new_name}_{i}.pdf"
                        new_path = dirpath / new_name
                        i += 1

                    print(f"[OK] {pdf_path} -> {new_path.name}")
                    pdf_path.rename(new_path)
                    renamed_in_folder = True

        results.append([folder_name, "Да" if renamed_in_folder else "Нет"])

    report_path = root / "report_sp.csv"
    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Название папки", "Переименование (Да/Нет)"])
        writer.writerows(results)

    return results, report_path


def main():
    root_tk = tk.Tk()
    root_tk.withdraw()
    root_tk.update()

    messagebox.showinfo(
        "Переименование файлов",
        "Сейчас нужно выбрать ПАПКУ, внутри которой лежат все дела с PDF-файлами.",
    )

    folder = filedialog.askdirectory(title="Выберите корневую папку с PDF-файлами")
    if not folder:
        messagebox.showwarning("Отмена", "Папка не выбрана. Работа завершена.")
        return

    root_path = Path(folder)
    if not root_path.exists():
        messagebox.showerror("Ошибка", f"Папка не найдена:\n{root_path}")
        return

    results, report_path = process_root_folder(root_path)

    lines = [
        "Готово.",
        f"Отчёт сохранён в файле:\n{report_path}",
        "",
        "Результат по папкам:",
    ]
    for folder_name, flag in results:
        lines.append(f"{folder_name}: {flag}")

    messagebox.showinfo("Работа завершена", "\n".join(lines))


if __name__ == "__main__":
    main()
