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
    """Авторизация по логину и паролю."""
    users = st.secrets["users"]  # {login: password}

    # Инициализация
    if "auth_ok" not in st.session_state:
        st.session_state["auth_ok"] = False
        st.session_state["current_user"] = None

    # Если уже авторизован — сразу пропускаем
    if st.session_state["auth_ok"]:
        return True

    # --- Форма логина ---
    st.markdown("### Добро пожаловать в W001-app")

    login = st.text_input("Логин:", key="login")
    password = st.text_input("Пароль:", type="password", key="password")

    if st.button("Войти"):
        if login in users and password == users[login]:
            st.session_state["auth_ok"] = True
            st.session_state["current_user"] = login

            # Ключевой момент — сразу перезапускаем приложение
            st.rerun()
        else:
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
