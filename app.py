import streamlit as st

from text_parser import parse_text
from pptx_generator import build_presentation

st.set_page_config(page_title="Текст → Презентация", page_icon="📊", layout="wide")

st.title("📊 Текст → готовая презентация")
st.caption("Введите тему и план обычным текстом — получите оформленный файл .pptx без единой строки кода.")

with st.sidebar:
    st.header("⚙️ Настройки")
    author = st.text_input("Автор", "")
    theme = st.selectbox(
        "Тема оформления",
        ["light", "dark"],
        format_func=lambda x: "☀️ Светлая" if x == "light" else "🌙 Тёмная",
    )
    max_items = st.slider(
        "Максимум пунктов на слайд", 3, 10, 6,
        help="Если пунктов больше, слайд автоматически разделится на несколько.",
    )

    st.divider()
    st.markdown(
        "**Как писать текст**\n\n"
        "- `# Название` — название презентации\n"
        "- `## Заголовок` — новый слайд\n"
        "- `- пункт` — пункт списка\n"
        "- обычные строки — абзац текста\n\n"
        "Без заголовков `##` каждый блок, отделённый пустой строкой, станет отдельным слайдом."
    )

EXAMPLE_TEXT = """# Как начать заниматься спортом
Практический план на первый месяц

## Зачем это нужно
- Больше энергии в течение дня
- Крепче сон и лучше настроение
- Меньше риск хронических болезней

## С чего начать
- Выберите вид активности, который нравится
- Начните с 20–30 минут три раза в неделю
- Запланируйте тренировки в календаре

## Типичные ошибки
- Слишком большая нагрузка в первые дни
- Отсутствие разминки
- Ожидание быстрых результатов

## Итоги
Регулярность важнее интенсивности. Маленькие шаги каждую неделю дают заметный результат уже через месяц.
"""

text = st.text_area(
    "Текст презентации:",
    value=EXAMPLE_TEXT,
    height=420,
)

if st.button("🚀 Создать презентацию", type="primary", use_container_width=True):
    try:
        deck = parse_text(text)
        pptx_buffer = build_presentation(
            deck,
            author=author.strip(),
            theme_name=theme,
            max_items_per_slide=max_items,
        )
        st.success(f"✅ Готово: {len(deck.slides)} слайдов с материалом + титульный и финальный.")
        safe_name = "".join(c for c in deck.title if c.isalnum() or c in " _-").strip().replace(" ", "_")[:50]
        st.download_button(
            label="⬇️ Скачать презентацию (.pptx)",
            data=pptx_buffer,
            file_name=f"{safe_name or 'presentation'}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            use_container_width=True,
        )
    except ValueError as e:
        st.error(f"❌ {e}")
    except Exception as e:
        st.error(f"💥 Неожиданная ошибка: {e}")
