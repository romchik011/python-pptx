import streamlit as st
from code_parser import parse_python_code, collect_leaves
from pptx_generator import build_presentation_from_code

st.set_page_config(page_title="Python Код → PPTX", page_icon="📊", layout="wide")

st.title("📊 Python Код → Обычная Презентация")
st.caption("Вставьте Python-код. Приложение создаст слайды из описаний (docstring) и структуры, БЕЗ отображения самого кода.")

with st.sidebar:
    st.header("⚙️ Настройки")
    title = st.text_input("Название презентации", "Документация проекта")
    author = st.text_input("Автор", "")
    theme = st.selectbox("Тема оформления", ["light", "dark"], format_func=lambda x: "☀️ Светлая" if x == "light" else "🌙 Тёмная")
    module_name = st.text_input("Имя модуля", "my_project")

EXAMPLE_CODE = '''"""
Модуль для работы с пользователями базы данных.
Содержит классы для CRUD операций и валидации данных.
"""

def validate_email(email: str) -> bool:
    """
    Проверяет корректность формата электронной почты.
    Возвращает True, если email валиден, иначе False.
    """
    return "@" in email and "." in email

class UserManager:
    """Управляет жизненным циклом пользователей в системе."""

    def __init__(self, db_connection):
        """Инициализирует менеджер с подключением к базе данных."""
        self.db = db_connection
        self.cache = {}

    def create_user(self, username: str, email: str) -> dict:
        """
        Создает нового пользователя в системе.
        
        Args:
            username: Уникальное имя пользователя
            email: Адрес электронной почты
            
        Returns:
            Словарь с данными созданного пользователя
        """
        if not validate_email(email):
            raise ValueError("Некорректный email")
        return {"username": username, "email": email, "status": "active"}
'''

# ВАЖНО: здесь НЕТ параметра language="python", чтобы не было ошибки
code = st.text_area(
    "Вставьте ваш Python-код сюда:",
    value=EXAMPLE_CODE,
    height=400,
    help="Приложение извлечет docstring и названия функций/классов для создания текстовых слайдов."
)

generate_btn = st.button("🚀 Создать презентацию", type="primary", use_container_width=True)

if generate_btn:
    if not code.strip():
        st.warning("⚠️ Пожалуйста, введите Python-код!")
    else:
        try:
            with st.spinner("🔍 Анализирую код..."):
                root = parse_python_code(code, module_name=module_name)
                elements = collect_leaves(root)

            if not elements:
                st.info("ℹ️ В коде не найдено функций или классов. Будет создан только титульный слайд.")
            else:
                classes_count = sum(1 for e in elements if e.kind == "class")
                funcs_count = sum(1 for e in elements if e.kind in ("function", "method"))
                st.success(f"✅ Найдено: {len(elements)} элементов ({classes_count} классов, {funcs_count} функций)")

            with st.spinner("📊 Генерирую обычные текстовые слайды..."):
                pptx_buffer = build_presentation_from_code(
                    root=root,
                    elements=elements,
                    title=title,
                    author=author,
                    theme_name=theme,
                )

            st.download_button(
                label="⬇️ Скачать обычную презентацию (.pptx)",
                data=pptx_buffer,
                file_name=f"{module_name}_docs.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

        except ValueError as e:
            st.error(f"❌ Ошибка парсинга: {e}")
        except Exception as e:
            st.error(f"💥 Неожиданная ошибка: {e}")
