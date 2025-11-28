# app.py — ГЛАВНЫЙ ФАЙЛ

import streamlit as st
import converter_app
import indexation_app
import statement_app

from ui_common import apply_global_css  # общий стиль для всех страниц

# Сначала настраиваем страницу (до любых других st.*)
st.set_page_config(
    page_title="Bank tools",
    layout="wide",
)

# ===== Простая авторизация =====


# ===== Простая авторизация =====

def check_auth():
    """Проверка логина и пароля по st.secrets['users']."""
    users = st.secrets["users"]  # это dict: {login: password}

    def login_entered():
        login = st.session_state.get("login", "").strip()
        password = st.session_state.get("password", "")

        if login in users and password == users[login]:
            st.session_state["auth_ok"] = True
            st.session_state["current_user"] = login
        else:
            st.session_state["auth_ok"] = False

    # первичная инициализация флагов
    if "auth_ok" not in st.session_state:
        st.session_state["auth_ok"] = False
        st.session_state["current_user"] = None

    # если ещё не залогинен — показываем форму
    if not st.session_state["auth_ok"]:
        st.markdown("### Добро пожаловать в W001-app")

        st.text_input("Логин:", key="login")
        st.text_input("Пароль:", type="password", key="password")

        if st.button("Войти"):
            login_entered()

        if (
            st.session_state.get("login")
            and st.session_state.get("password")
            and not st.session_state["auth_ok"]
        ):
            st.error("Неверный логин или пароль")

        return False

    # если авторизация прошла успешно
    return True



# Блок защиты
if not check_auth():
    st.stop()



def main():
    # применяем единый CSS для всех разделов
    apply_global_css()
    user = st.session_state.get("current_user", "—")
    st.sidebar.title("Навигация")
    st.sidebar.caption(f"Пользователь: {user}")

    
    st.sidebar.title("Навигация")
    page = st.sidebar.radio(
        "",
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
