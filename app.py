import streamlit as st
from code_parser import parse_python_code, collect_leaves
from pptx_generator import build_presentation

st.set_page_config(page_title="Python to PPTX", page_icon="", layout="wide")
st.title(" Python -> PowerPoint конвертер")
st.caption("Вставьте Python-код - получите презентацию с подсветкой синтаксиса")

with st.sidebar:
    st.header("Настройки")
    title = st.text_input("Название презентации", "Python Code Overview")
    author = st.text_input("Автор", "")
    theme = st.selectbox("Тема оформления", ["dark", "light"])
    module_name = st.text_input("Имя модуля", "module")

EXAMPLE = '''"""Пример модуля."""


def greet(name: str) -> str:
    """Возвращает приветствие."""
    return f"Hello, {name}!"


class Calculator:
    """Простой калькулятор."""

    def __init__(self):
        self.history = []

    def add(self, a: float, b: float) -> float:
        """Складывает два числа."""
        result = a + b
        self.history.append(("add", a, b, result))
        return result
'''

code = st.text_area("Python-код", value=EXAMPLE, height=400)

if st.button("🚀 Сгенерировать презентацию", type="primary"):
    if not code.strip():
        st.warning("Введите Python-код!")
    else:
        try:
            with st.spinner("Парсим код..."):
                root = parse_python_code(code, module_name=module_name)
                elements = collect_leaves(root)

            if not elements:
                st.warning("В коде не найдено функций или классов.")
            else:
                st.success(f"Найдено элементов: {len(elements)}")

            with st.spinner("Собираем презентацию..."):
                pptx_buffer = build_presentation(
                    root=root, elements=elements,
                    title=title, author=author, theme_name=theme,
                )

            st.download_button(
                label="Скачать .pptx",
                data=pptx_buffer,
                file_name=f"{module_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
            )
        except ValueError as e:
            st.error(f"Ошибка: {e}")
        except Exception as e:
            st.error(f"Неожиданная ошибка: {e}")
