import streamlit as st
from code_parser import parse_python_code, collect_leaves
from pptx_generator import build_presentation_from_code

st.set_page_config(page_title="Python Код → PPTX", page_icon="📊", layout="wide")

st.title("📊 Python Код → Презентация (неограниченное количество слайдов)")
st.caption("Вставьте Python-код. Приложение создаст столько слайдов, сколько нужно!")

with st.sidebar:
    st.header("⚙️ Настройки")
    title = st.text_input("Название презентации", "Документация проекта")
    author = st.text_input("Автор", "")
    theme = st.selectbox("Тема оформления", ["light", "dark"], 
                         format_func=lambda x: "☀️ Светлая" if x == "light" else " Тёмная")
    module_name = st.text_input("Имя модуля", "my_project")
    
    st.divider()
    st.subheader("📏 Настройки слайдов")
    max_chars = st.slider("Максимум символов на слайд", 200, 800, 400, 
                          help="Если описание длинное, оно разобьется на несколько слайдов")

EXAMPLE_CODE = '''"""
Большой модуль для демонстрации.
Содержит много классов и функций для создания презентации.
"""

def function_one():
    """Первая функция с описанием."""
    pass

def function_two():
    """Вторая функция с описанием."""
    pass

def function_three():
    """Третья функция с описанием."""
    pass

class ClassOne:
    """Первый класс."""
    
    def method_one(self):
        """Метод первого класса."""
        pass
    
    def method_two(self):
        """Еще один метод."""
        pass

class ClassTwo:
    """Второй класс."""
    
    def method_one(self):
        """Метод второго класса."""
        pass

def function_four():
    """Четвертая функция."""
    pass

def function_five():
    """Пятая функция."""
    pass
'''

code = st.text_area(
    "Вставьте ваш Python-код сюда:",
    value=EXAMPLE_CODE,
    height=400,
    help="Чем больше функций и классов, тем больше слайдов будет создано!"
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
                st.info("ℹ️ В коде не найдено функций или классов.")
            else:
                classes_count = sum(1 for e in elements if e.kind == "class")
                funcs_count = sum(1 for e in elements if e.kind in ("function", "method"))
                st.success(f"✅ Найдено: {len(elements)} элементов ({classes_count} классов, {funcs_count} функций)")
                st.info(f"📊 Будет создано примерно {len(elements) + 3} слайдов")

            with st.spinner("📊 Генерирую презентацию..."):
                pptx_buffer = build_presentation_from_code(
                    root=root,
                    elements=elements,
                    title=title,
                    author=author,
                    theme_name=theme,
                    max_chars_per_slide=max_chars,
                )

            st.download_button(
                label="⬇️ Скачать презентацию (.pptx)",
                data=pptx_buffer,
                file_name=f"{module_name}_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

        except ValueError as e:
            st.error(f"❌ Ошибка парсинга: {e}")
        except Exception as e:
            st.error(f"💥 Неожиданная ошибка: {e}")
