import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io
import re

def parse_template(text):
    slides = []
    slide_blocks = re.split(r'===\s*СЛАЙД\s*\d+\s*===', text)
    
    for block in slide_blocks:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        
        if not lines:
            continue
        
        title = lines[0]
        
        points = []
        for line in lines[1:]:
            clean_line = re.sub(r'^[-*•]\s*', '', line)
            points.append(clean_line)
        
        slides.append((title, points))
    
    return slides

def create_presentation(slides):
    prs = Presentation()
    
    if slides:
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = slides[0][0]
        title_slide.placeholders[1].text = "Презентация"
    
    for slide_title, points in slides[1:] if len(slides) > 1 else slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_title
        
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()
        
        for point in points:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = Pt(22)
    
    end_slide = prs.slides.add_slide(prs.slide_layouts[0])
    end_slide.shapes.title.text = "Спасибо!"
    end_slide.placeholders[1].text = ""
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

st.set_page_config(page_title="Шаблон презентаций", page_icon="📝")

st.title("📝 Генератор презентаций из шаблона")

st.markdown("""
### 📋 Как писать шаблон:

1. Каждый слайд начинается с `=== СЛАЙД 1 ===`, `=== СЛАЙД 2 ===` и т.д.
2. Первая строка после маркера — **заголовок слайда**
3. Остальные строки — **пункты** (можно начинать с `-` или `*`)
""")

st.markdown("---")

default_template = """=== СЛАЙД 1 ===
Добро пожаловать
Презентация о нашей компании

=== СЛАЙД 2 ===
О нас
- Работаем с 2020 года
- Команда из 50 профессионалов
- Более 100 довольных клиентов

=== СЛАЙД 3 ===
Наши услуги
- Разработка сайтов
- Мобильные приложения
- Дизайн и маркетинг
- Консультации

=== СЛАЙД 4 ===
Контакты
- Email: info@example.com
- Телефон: +7 999 123-45-67
- Адрес: Москва, ул. Примерная, 1
"""

user_template = st.text_area(
    "Введите ваш шаблон здесь:",
    value=default_template,
    height=400
)

if st.button("🎨 Создать презентацию", type="primary", use_container_width=True):
    if user_template.strip():
        with st.spinner('Анализирую шаблон...'):
            slides = parse_template(user_template)
            
            if not slides:
                st.error("❌ Не удалось найти слайды. Проверьте формат шаблона!")
            else:
                st.success(f"✅ Найдено слайдов: {len(slides)}")
                
                st.write("**Превью слайдов:**")
                for i, (title, points) in enumerate(slides, 1):
                    with st.expander(f"Слайд {i}: {title}"):
                        if points:
                            for point in points:
                                st.write(f"• {point}")
                        else:
                            st.write("(только заголовок)")
                
                with st.spinner('Создаю файл...'):
                    pptx_data = create_presentation(slides)
                
                st.download_button(
                    label="📥 Скачать презентацию .pptx",
                    data=pptx_data,
                    file_name="presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
    else:
        st.error("❌ Введите текст шаблона!")
