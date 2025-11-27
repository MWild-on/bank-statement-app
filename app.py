# app.py — ГЛАВНЫЙ ФАЙЛ

import streamlit as st
import converter_app
import indexation_app
import statement_app  # 🔹 НОВЫЙ МОДУЛЬ

# ===== Простая авторизация =====
CREDENTIALS = {
    "Mariam": "Mariam4321",
    "MM": "MM5432",
    "MO": "1",
}

def login():
    st.title("🔐 Авторизация")
    with st.form("login_form"):
        username = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти")
        if submitted:
            if username in CREDENTIALS and CREDENTIALS[username] == password:
                st.session_state["auth"] = True
                st.session_state["user"] = username
            else:
                st.error("Неверный логин или пароль")

if "auth" not in st.session_state or not st.session_state["auth"]:
    login()
    st.stop()


def main():
    st.set_page_config(
        page_title="Bank tools",
        layout="wide",
    )

    st.sidebar.title("Навигация")
    page = st.sidebar.radio(
        "Выберите раздел:",
        ("Конвертер", "Индексация", "Создание выписки"),
    )

    if page == "Конвертер":
        converter_app.run()
    elif page == "Индексация":
        indexation_app.run()
    elif page == "Создание выписки":
        statement_app.run()   # 🔹 Вызов нового раздела

if __name__ == "__main__":
    main()
