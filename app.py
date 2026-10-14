import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io

def generate_pptx(text_content):
    """Функция создания презентации из текста"""
    prs = Presentation()
    
    # Разбиваем текст на слайды по двойному переносу строки
    slides_raw = text_content.split('\n\n')
    
    for block in slides_raw:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        
        if not lines:
            continue
            
        # Используем макет "Заголовок и объект" (индекс 1)
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        # Заголовок слайда
        title = slide.shapes.title
        title.text = lines[0]
        
        # Тело слайда (список)
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        tf.clear() 
        
        if len(lines) > 1:
            for point in lines[1:]:
                p = tf.add_paragraph()
                p.text = point
                p.level = 0
                p.font.size = Pt(24)

    # Сохраняем в байтовый поток
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- Интерфейс Streamlit ---

st.set_page_config(page_title="Генератор Презентаций", page_icon="📊")

st.title("📊 Генератор PowerPoint из текста")

st.markdown("""
**Как пользоваться:**
1. Введите текст ниже.
2. Первая строка блока — **Заголовок слайда**.
3. Остальные строки — **пункты списка**.
4. Оставьте **пустую строку** между слайдами.
""")

default_text = """Пример слайда 1
Пункт 1
Пункт 2

Пример слайда 2
Текст о чем-то важном
Еще один пункт
"""

user_text = st.text_area("Введите текст для презентации:", value=default_text, height=300)

if st.button("Сгенерировать и скачать .pptx"):
    if user_text:
        with st.spinner('Генерация презентации...'):
            pptx_file = generate_pptx(user_text)
            
        st.download_button(
            label="📥 Скачать презентацию",
            data=pptx_file,
            file_name="my_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    else:
        st.warning("Пожалуйста, введите текст.")
