# app.py — ГЛАВНЫЙ ФАЙЛ

import streamlit as st
import converter_app
import indexation_app
import statement_app  # 🔹 НОВЫЙ МОДУЛЬ

# ===== Простая авторизация =====
import streamlit as st

import streamlit as st

def check_password():
    # Функция проверки пароля
    def password_entered():
        entered = st.session_state.get("password", "")
        if entered == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
        else:
            st.session_state["password_correct"] = False

    # Первичная инициализация флага
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    # Если пароль ещё не введён или он неверный — показываем форму входа
    if not st.session_state["password_correct"]:
        st.markdown("### Добро пожаловать в W001-app")

        st.text_input("Пароль:", type="password", key="password")

        if st.button("Войти"):
            password_entered()

        if st.session_state.get("password") and not st.session_state["password_correct"]:
            st.error("Неверный пароль")

        return False

    # Если пароль корректный — доступ разрешён
    return True


# Блок защиты
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
