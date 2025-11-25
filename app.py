# app.py — ГЛАВНЫЙ ФАЙЛ

import streamlit as st
import converter_app
import indexation_app

def main():
    st.set_page_config(
        page_title="Bank tools",
        layout="wide",
    )

    st.sidebar.title("Навигация")
    page = st.sidebar.radio(
        "Выберите раздел:",
        ("Конвертер", "Индексация"),
    )

    if page == "Конвертер":
        converter_app.run()
    elif page == "Индексация":
        indexation_app.run()

if __name__ == "__main__":
    main()
