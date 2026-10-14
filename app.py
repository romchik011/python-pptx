import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import re

def parse_color(color_str):
    """Преобразует цвет из строки в RGBColor"""
    color_str = color_str.strip()
    if color_str.startswith('#'):
        color_str = color_str[1:]
    if len(color_str) == 6:
        r = int(color_str[0:2], 16)
        g = int(color_str[2:4], 16)
        b = int(color_str[4:6], 16)
        return RGBColor(r, g, b)
    return None

def parse_template(text):
    """Разбирает текстовый шаблон на слайды с настройками"""
    slides = []
    slide_blocks = re.split(r'===\s*СЛАЙД\s*\d+\s*===', text)
    
    for block in slide_blocks:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        
        if not lines:
            continue
        
        # Настройки по умолчанию
        settings = {
            'bg': None,
            'font': 'Calibri',
            'size': 22,
            'title_size': 32,
            'text_color': None,
            'title_color': None
        }
        
        # Читаем мета-теги (строки начинающиеся с @)
        content_lines = []
        for line in lines:
            if line.startswith('@bg:'):
                settings['bg'] = line.split('@bg:')[1].strip()
            elif line.startswith('@font:'):
                settings['font'] = line.split('@font:')[1].strip()
            elif line.startswith('@size:'):
                try:
                    settings['size'] = int(line.split('@size:')[1].strip())
                except:
                    pass
            elif line.startswith('@title_size:'):
                try:
                    settings['title_size'] = int(line.split('@title_size:')[1].strip())
                except:
                    pass
            elif line.startswith('@color:'):
                settings['text_color'] = line.split('@color:')[1].strip()
            elif line.startswith('@title_color:'):
                settings['title_color'] = line.split('@title_color:')[1].strip()
            else:
                content_lines.append(line)
        
        if not content_lines:
            continue
        
        # Первая строка - заголовок
        title = content_lines[0]
        
        # Остальные строки - пункты
        points = []
        for line in content_lines[1:]:
            clean_line = re.sub(r'^[-*•]\s*', '', line)
            points.append(clean_line)
        
        slides.append({
            'title': title,
            'points': points,
            'settings': settings
        })
    
    return slides

def set_slide_background(slide, color_hex):
    """Устанавливает цвет фона слайда"""
    if not color_hex:
        return
    
    color = parse_color(color_hex)
    if not color:
        return
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def create_presentation(slides):
    """Создаёт презентацию из разобранных данных"""
    prs = Presentation()
    prs.slide_width = 12192000  # 16:9
    prs.slide_height = 6858000
    
    for i, slide_data in enumerate(slides):
        settings = slide_data['settings']
        
        # Выбираем макет
        if i == 0:
            slide_layout = prs.slide_layouts[6]  # Пустой макет
        else:
            slide_layout = prs.slide_layouts[6]  # Пустой макет
        
        slide = prs.slides.add_slide(slide_layout)
        
        # Устанавливаем фон
        set_slide_background(slide, settings['bg'])
        
        # Цвета
        title_color = parse_color(settings['title_color']) if settings['title_color'] else RGBColor(255, 255, 255)
        text_color = parse_color(settings['text_color']) if settings['text_color'] else RGBColor(255, 255, 255)
        
        if i == 0:
            # Титульный слайд - большой заголовок по центру
            from pptx.util import Emu
            left = Emu(1000000)
            top = Emu(2000000)
            width = Emu(10000000)
            height = Emu(2000000)
            
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slide_data['title']
            p.font.size = Pt(settings['title_size'] + 10)
            p.font.name = settings['font']
            p.font.color.rgb = title_color
            p.alignment = PP_ALIGN.CENTER
            
            # Подзаголовок
            if slide_data['points']:
                top2 = Emu(4200000)
                txBox2 = slide.shapes.add_textbox(left, top2, width, Emu(1500000))
                tf2 = txBox2.text_frame
                tf2.word_wrap = True
                p2 = tf2.paragraphs[0]
                p2.text = slide_data['points'][0]
                p2.font.size = Pt(settings['size'])
                p2.font.name = settings['font']
                p2.font.color.rgb = text_color
                p2.alignment = PP_ALIGN.CENTER
        else:
            # Обычный слайд
            from pptx.util import Emu
            
            # Заголовок
            left = Emu(500000)
            top = Emu(300000)
            width = Emu(11000000)
            height = Emu(1200000)
            
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slide_data['title']
            p.font.size = Pt(settings['title_size'])
            p.font.name = settings['font']
            p.font.color.rgb = title_color
            p.font.bold = True
            
            # Пункты
            if slide_data['points']:
                top_points = Emu(1600000)
                height_points = Emu(5000000)
                txBox2 = slide.shapes.add_textbox(left, top_points, width, height_points)
                tf2 = txBox2.text_frame
                tf2.word_wrap = True
                
                for j, point in enumerate(slide_data['points']):
                    if j == 0:
                        p = tf2.paragraphs[0]
                    else:
                        p = tf2.add_paragraph()
                    p.text = f"• {point}"
                    p.font.size = Pt(settings['size'])
                    p.font.name = settings['font']
                    p.font.color.rgb = text_color
                    p.space_after = Pt(12)
    
    # Сохраняем
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# === ИНТЕРФЕЙС ===
st.set_page_config(page_title="Генератор презентаций Pro", page_icon="🎨", layout="wide")

st.title("🎨 Генератор презентаций Pro")

st.markdown("""
### 📋 Как использовать настройки:

Добавьте **мета-теги** в начало блока слайда:

| Тег | Описание | Пример |
|-----|----------|--------|
| `@bg:` | Цвет фона | `@bg: #1a1a2e` |
| `@font:` | Шрифт | `@font: Arial` |
| `@size:` | Размер текста | `@size: 24` |
| `@title_size:` | Размер заголовка | `@title_size: 36` |
| `@color:` | Цвет текста | `@color: #ffffff` |
| `@title_color:` | Цвет заголовка | `@title_color: #ffd700` |
""")

# Палитра цветов
st.markdown("---")
st.subheader("🎨 Популярные цвета фона:")

colors = {
    "🌙 Тёмно-синий": "#1a1a2e",
    " Морской": "#16213e",
    "🌲 Изумрудный": "#0f3460",
    "🔴 Бордовый": "#533483",
    " Чёрный": "#000000",
    "⚪ Белый": "#ffffff",
    "🌸 Розовый": "#e94560",
    "🌿 Зелёный": "#2d6a4f",
    "🌅 Оранжевый": "#f4a261",
    " Фиолетовый": "#7209b7",
    "☕ Кофейный": "#6f4e37",
    "🌊 Голубой": "#48cae4"
}

cols = st.columns(4)
for i, (name, hex_code) in enumerate(colors.items()):
    with cols[i % 4]:
        st.markdown(f"**{name}**")
        st.code(hex_code)

st.markdown("---")
st.subheader("🔤 Популярные шрифты:")
st.markdown("`Arial` • `Calibri` • `Times New Roman` • `Georgia` • `Verdana` • `Trebuchet MS` • `Impact` • `Comic Sans MS`")

st.markdown("---")

default_template = """=== СЛАЙД 1 ===
@bg: #1a1a2e
@font: Georgia
@title_size: 44
@title_color: #ffd700
@color: #ffffff

Вольфганг Амадей Моцарт
Великий композитор • 1756 — 1791

=== СЛАЙД 2 ===
@bg: #16213e
@font: Arial
@size: 24
@title_size: 36
@title_color: #48cae4
@color: #ffffff

Ранние годы
- Родился 27 января 1756 года в Зальцбурге
- Отец — Леопольд Моцарт, композитор и скрипач
- Мать — Анна Мария Моцарт
- С детства проявлял необычные музыкальные способности

=== СЛАЙД 3 ===
@bg: #0f3460
@font: Verdana
@size: 22
@title_size: 34
@title_color: #f4a261
@color: #ffffff

Музыкальный вундеркинд
- Начал играть на клавесине в 3 года
- В 5 лет уже сочинял маленькие пьесы
- В 6 лет выступал перед императорским двором
- К 8 годам написал свои первые симфонии

=== СЛАЙД 4 ===
@bg: #533483
@font: Georgia
@size: 24
@title_size: 36
@title_color: #ffd700
@color: #ffffff

Европейские гастроли
- С 1762 года путешествовал по Европе
- Выступал в Вене, Париже, Лондоне, Италии
- Познакомился с великими музыкантами эпохи
- Впитал лучшие традиции европейской музыки

=== СЛАЙД 5 ===
@bg: #2d6a4f
@font: Arial
@size: 22
@title_size: 34
@title_color: #ffffff
@color: #d8f3dc

Зрелые годы в Вене
- В 1781 году переехал в Вену
- Стал независимым композитором
- Давал уроки, концерты, писал на заказ
- Создал свои главные шедевры

=== СЛАЙД 6 ===
@bg: #7209b7
@font: Trebuchet MS
@size: 24
@title_size: 36
@title_color: #ffd700
@color: #ffffff

Великие оперы
- «Свадьба Фигаро» (1786)
- «Дон Жуан» (1787)
- «Так поступают все» (1790)
- «Волшебная флейта» (1791)
- Реформатор оперного искусства

=== СЛАЙД 7 ===
@bg: #1a1a2e
@font: Georgia
@size: 22
@title_size: 34
@title_color: #48cae4
@color: #ffffff

Симфонии и концерты
- Написал 41 симфонию
- 27 фортепианных концертов
- 5 скрипичных концертов
- Симфония №40 и №41 — вершины жанра

=== СЛАЙД 8 ===
@bg: #6f4e37
@font: Arial
@size: 24
@title_size: 36
@title_color: #ffd700
@color: #ffffff

Камерная музыка
- Более 20 струнных квартетов
- Сонаты для фортепиано и скрипки
- Квинтеты и трио
- «Маленькая ночная серенада» — самое известное произведение

=== СЛАЙД 9 ===
@bg: #000000
@font: Georgia
@size: 22
@title_size: 34
@title_color: #e94560
@color: #cccccc

Последние годы жизни
- Финансовые трудности
- Болезнь и упадок сил
- Продолжал работать несмотря ни на что
- Умер 5 декабря 1791 года в возрасте 35 лет

=== СЛАЙД 10 ===
@bg: #16213e
@font: Verdana
@size: 22
@title_size: 34
@title_color: #ffd700
@color: #ffffff

«Реквием» — последнее произведение
- Заказан анонимным меценатом
- Моцарт чувствовал, что пишет его для себя
- Не успел завершить — дописал ученик Зюсмайер
- Одно из самых трагичных произведений в истории музыки

=== СЛАЙД 11 ===
@bg: #0f3460
@font: Arial
@size: 24
@title_size: 36
@title_color: #48cae4
@color: #ffffff

Наследие Моцарта
- Более 600 произведений за короткую жизнь
- Вершина венского классицизма
- Влияние на Бетховена, Шуберта, Чайковского
- Музыка актуальна спустя более 200 лет

=== СЛАЙД 12 ===
@bg: #533483
@font: Georgia
@size: 24
@title_size: 38
@title_color: #ffd700
@color: #ffffff

Интересные факты
- Обладал абсолютным музыкальным слухом
- Мог сочинять музыку в уме, не записывая
- Писал музыку с невероятной скоростью
- Его имя стало символом музыкального гения
"""

user_template = st.text_area(
    "Введите ваш шаблон здесь:",
    value=default_template,
    height=500
)

if st.button("🎨 Создать презентацию", type="primary", use_container_width=True):
    if user_template.strip():
        with st.spinner('Анализирую шаблон...'):
            slides = parse_template(user_template)
            
            if not slides:
                st.error("❌ Не удалось найти слайды. Проверьте формат шаблона!")
            else:
                st.success(f"✅ Найдено слайдов: {len(slides)}")
                
                # Показываем превью с цветами
                st.write("**Превью слайдов:**")
                for i, slide_data in enumerate(slides, 1):
                    settings = slide_data['settings']
                    bg_color = settings['bg'] or '#ffffff'
                    
                    with st.expander(f"Слайд {i}: {slide_data['title']}"):
                        st.markdown(f"**Фон:** {bg_color} | **Шрифт:** {settings['font']} | **Размер:** {settings['size']}")
                        if slide_data['points']:
                            for point in slide_data['points']:
                                st.write(f"• {point}")
                        else:
                            st.write("(только заголовок)")
                
                with st.spinner('Создаю файл...'):
                    pptx_data = create_presentation(slides)
                
                st.download_button(
                    label="📥 Скачать презентацию .pptx",
                    data=pptx_data,
                    file_name="mozart_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
    else:
        st.error("❌ Введите текст шаблона!")
