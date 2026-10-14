import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io

# ТЕКСТ ДЛЯ СЛАЙДОВ (здесь НЕТ кода, только обычные слова)
SLIDES_DATA = {
    "🐍 О языке Python": [
        ("Что такое Python?", [
            "Python — это язык программирования",
            "Создан в 1991 году",
            "Один из самых популярных в мире",
            "Используется в веб-разработке и AI"
        ]),
        ("Преимущества Python", [
            "Простой и понятный синтаксис",
            "Большое сообщество разработчиков",
            "Много готовых библиотек",
            "Работает на всех системах"
        ]),
        ("Где применяется", [
            "Создание сайтов",
            "Искусственный интеллект",
            "Анализ данных",
            "Автоматизация задач"
        ])
    ],
    "🚀 О космосе": [
        ("Солнечная система", [
            "Солнце — звезда в центре",
            "Восемь планет вокруг Солнца",
            "Земля — наш дом",
            "Есть астероиды и кометы"
        ]),
        ("Исследование космоса", [
            "Первый спутник запущен в 1957",
            "Гагарин полетел в космос в 1961",
            "Сегодня работает МКС",
            "Планы — полёт на Марс"
        ]),
        ("Звёзды и галактики", [
            "Звёзды — огромные шары газа",
            "Наша галактика — Млечный Путь",
            "Во Вселенной триллионы галактик",
            "Космос бесконечен"
        ])
    ],
    "💼 О бизнесе": [
        ("Наша компания", [
            "Работаем с 2020 года",
            "Команда из 50 человек",
            "Офисы в трёх городах",
            "Более 100 клиентов"
        ]),
        ("Наши услуги", [
            "Разработка сайтов",
            "Мобильные приложения",
            "Дизайн и маркетинг",
            "Консультации"
        ]),
        ("Контакты", [
            "Email: info@example.com",
            "Телефон: +7 999 123-45-67",
            "Адрес: Москва, ул. Примерная, 1",
            "Сайт: example.com"
        ])
    ]
}

def create_presentation(title, slides):
    prs = Presentation()
    
    # Титульный слайд
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Автоматическая презентация"
    
    # Слайды с текстом
    for slide_title, points in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_title
        
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()
        
        for point in points:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = Pt(22)
    
    # Финальный слайд
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Спасибо за внимание!"
    slide.placeholders[1].text = "Вопросы?"
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# === ИНТЕРФЕЙС САЙТА ===
st.set_page_config(page_title="Генератор презентаций", page_icon="📊")

st.title("📊 Генератор презентаций")
st.write("Выберите тему и скачайте готовую презентацию")

st.markdown("---")

# Три кнопки для выбора темы
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("🐍 Python", use_container_width=True):
        st.session_state.topic = "🐍 О языке Python"

with col2:
    if st.button("🚀 Космос", use_container_width=True):
        st.session_state.topic = "🚀 О космосе"

with col3:
    if st.button("💼 Бизнес", use_container_width=True):
        st.session_state.topic = "💼 О бизнесе"

# Если тема выбрана - показываем кнопку скачивания
if "topic" in st.session_state:
    topic = st.session_state.topic
    st.markdown("---")
    st.success(f"✅ Выбрана тема: **{topic}**")
    
    # Показываем что будет в презентации
    st.write("**Слайды в презентации:**")
    for slide_title, _ in SLIDES_DATA[topic]:
        st.write(f"• {slide_title}")
    
    # Кнопка скачивания
    if st.button("📥 СКАЧАТЬ ПРЕЗЕНТАЦИЮ", type="primary", use_container_width=True):
        pptx_data = create_presentation(topic, SLIDES_DATA[topic])
        
        st.download_button(
            label="💾 Скачать файл .pptx",
            data=pptx_data,
            file_name=f"{topic}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
