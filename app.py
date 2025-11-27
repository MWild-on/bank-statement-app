# app.py — ГЛАВНЫЙ ФАЙЛ

import streamlit as st
import converter_app
import indexation_app
import statement_app  # 🔹 НОВЫЙ МОДУЛЬ

# ===== Простая авторизация =====
import streamlit as st

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"] 
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Пароль:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Пароль:", type="password", on_change=password_entered, key="password")
        st.error("Неверный пароль")
        return False
    else:
        return True

if not check_password():
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
        statement_app.run()  
if __name__ == "__main__":
    main()
