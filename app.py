import streamlit as st
from code_parser import parse_python_code, collect_leaves
from pptx_generator import build_presentation

st.set_page_config(page_title="Python → PPTX Конвертер", page_icon="🐍", layout="wide")

st.title("🐍 Python → PowerPoint конвертер")
st.caption("Вставьте любой Python-код — получите готовую презентацию с подсветкой синтаксиса на слайдах")

with st.sidebar:
    st.header("⚙️ Настройки презентации")
    title = st.text_input("Название презентации", "Python Code Overview")
    author = st.text_input("Автор", "")
    theme = st.selectbox("Тема оформления", ["dark", "light"], 
                         format_func=lambda x: "🌙 Тёмная (Monokai)" if x == "dark" else "☀️ Светлая")
    module_name = st.text_input("Имя модуля", "module")

EXAMPLE_CODE = '''"""Пример модуля с классом и функциями."""


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

    @staticmethod
    def pi() -> float:
        return 3.14159265
'''

# Исправленная строка (убран параметр language="python")
code = st.text_area(
    "Вставьте ваш Python-код сюда:",
    value=EXAMPLE_CODE,
    height=400,
    help="Приложение автоматически извлечёт классы, функции, docstring и декораторы."
)

col1, col2 = st.columns([1, 4])
with col1:
    generate_btn = st.button("🚀 Сгенерировать .pptx", type="primary", use_container_width=True)

if generate_btn:
    if not code.strip():
        st.warning("⚠️ Пожалуйста, введите Python-код!")
    else:
        try:
            with st.spinner("🔍 Анализирую структуру кода..."):
                root = parse_python_code(code, module_name=module_name)
                elements = collect_leaves(root)

            if not elements:
                st.info("ℹ️ В коде не найдено функций или классов. Будет создан только титульный слайд.")
            else:
                classes_count = sum(1 for e in elements if e.kind == "class")
                funcs_count = sum(1 for e in elements if e.kind in ("function", "method"))
                st.success(f"✅ Найдено: {len(elements)} элементов ({classes_count} классов, {funcs_count} функций)")

            with st.spinner("🎨 Рисую слайды и добавляю подсветку синтаксиса..."):
                pptx_buffer = build_presentation(
                    root=root,
                    elements=elements,
                    title=title,
                    author=author,
                    theme_name=theme,
                )

            st.download_button(
                label="⬇️ Скачать презентацию (.pptx)",
                data=pptx_buffer,
                file_name=f"{module_name}_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

            with st.expander("👁️ Предпросмотр структуры"):
                for el in elements:
                    icon = "🧩" if el.kind == "class" else "⚙️"
                    st.markdown(f"{icon} **{el.name}** — *{el.kind}*")
                    if el.docstring:
                        st.caption(el.docstring.splitlines()[0])

        except ValueError as e:
            st.error(f"❌ Ошибка парсинга: {e}")
        except Exception as e:
            st.error(f"💥 Неожиданная ошибка при генерации: {e}")
