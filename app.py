import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import re
import requests
import json

# ============================================================
# ФУНКЦИИ ДЛЯ СОЗДАНИЯ ПРЕЗЕНТАЦИИ
# ============================================================

def parse_color(color_str):
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
    slides = []
    slide_blocks = re.split(r'===\s*СЛАЙД\s*\d+\s*===', text)
    
    for block in slide_blocks:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        if not lines:
            continue
        
        settings = {
            'bg': '#0a192f',
            'font': 'Calibri',
            'size': 22,
            'title_size': 32,
            'text_color': '#ccd6f6',
            'title_color': '#64ffda'
        }
        
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
        
        title = content_lines[0]
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
    from pptx.util import Emu
    prs = Presentation()
    prs.slide_width = 12192000
    prs.slide_height = 6858000
    
    for i, slide_data in enumerate(slides):
        settings = slide_data['settings']
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        
        set_slide_background(slide, settings['bg'])
        
        title_color = parse_color(settings['title_color']) if settings['title_color'] else RGBColor(255, 255, 255)
        text_color = parse_color(settings['text_color']) if settings['text_color'] else RGBColor(255, 255, 255)
        
        if i == 0:
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
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ============================================================
# ИНТЕГРАЦИЯ С OPENAI API
# ============================================================

def call_openai(api_key, prompt, model="gpt-3.5-turbo"):
    """Вызов OpenAI API"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 2000
    }
    
    try:
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=data,
            timeout=30
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"Ошибка API: {str(e)}"

def generate_presentation_with_ai(api_key, topic, num_slides=5, color_theme="blue"):
    """Генерация презентации через ИИ"""
    
    color_themes = {
        "blue": {"bg1": "#0a192f", "bg2": "#112240", "bg3": "#1a2f4a", "title": "#64ffda", "text": "#ccd6f6"},
        "red": {"bg1": "#2d0a0a", "bg2": "#4a0f0f", "bg3": "#6b1515", "title": "#ff6b6b", "text": "#ffd4d4"},
        "green": {"bg1": "#0a2d1a", "bg2": "#0f4a2a", "bg3": "#156b3d", "title": "#6bff9e", "text": "#d4ffd4"},
        "purple": {"bg1": "#1a0a2d", "bg2": "#2a0f4a", "bg3": "#3d156b", "title": "#c56bff", "text": "#e8d4ff"},
        "dark": {"bg1": "#1a1a1a", "bg2": "#2d2d2d", "bg3": "#404040", "title": "#ffffff", "text": "#cccccc"},
    }
    
    colors = color_themes.get(color_theme, color_themes["blue"])
    bg_colors = [colors["bg1"], colors["bg2"], colors["bg3"]]
    
    prompt = f"""Создай структуру презентации на тему "{topic}" из {num_slides} слайдов.

Для каждого слайда укажи:
1. Заголовок слайда
2. 3-4 пункта с конкретной информацией (не общие фразы, а реальные факты)

Формат ответа (строго следуй):

СЛАЙД 1: [Заголовок титульного слайда]
ПОДЗАГОЛОВОК: [Краткое описание]

СЛАЙД 2: [Заголовок]
- [Пункт 1]
- [Пункт 2]
- [Пункт 3]

СЛАЙД 3: [Заголовок]
- [Пункт 1]
- [Пункт 2]
- [Пункт 3]

И так далее для всех {num_slides} слайдов.

ВАЖНО: Пиши конкретные факты, даты, имена, цифры. Не используй общие фразы типа "Ключевое понятие 1"."""
    
    response = call_openai(api_key, prompt)
    
    if response.startswith("Ошибка API:"):
        return None, response
    
    # Парсим ответ ИИ
    slides_data = []
    current_slide = None
    
    for line in response.split('\n'):
        line = line.strip()
        
        if line.startswith('СЛАЙД') and ':' in line:
            if current_slide:
                slides_data.append(current_slide)
            title = line.split(':', 1)[1].strip()
            current_slide = {'title': title, 'points': []}
        elif line.startswith('ПОДЗАГОЛОВОК') and ':' in line:
            if current_slide:
                subtitle = line.split(':', 1)[1].strip()
                current_slide['points'].append(subtitle)
        elif line.startswith('-') and current_slide:
            point = line[1:].strip()
            current_slide['points'].append(point)
    
    if current_slide:
        slides_data.append(current_slide)
    
    # Генерируем шаблон
    template_lines = []
    
    # Титульный слайд
    if slides_data:
        template_lines.append("=== СЛАЙД 1 ===")
        template_lines.append(f"@bg: {colors['bg1']}")
        template_lines.append("@font: Georgia")
        template_lines.append(f"@title_size: 44")
        template_lines.append(f"@title_color: {colors['title']}")
        template_lines.append(f"@color: {colors['text']}")
        template_lines.append("")
        template_lines.append(slides_data[0]['title'])
        if slides_data[0]['points']:
            template_lines.append(slides_data[0]['points'][0])
        template_lines.append("")
        
        # Остальные слайды
        for i, slide in enumerate(slides_data[1:], 2):
            bg = bg_colors[(i - 2) % 3]
            template_lines.append(f"=== СЛАЙД {i} ===")
            template_lines.append(f"@bg: {bg}")
            template_lines.append("@font: Arial")
            template_lines.append("@size: 22")
            template_lines.append("@title_size: 34")
            template_lines.append(f"@title_color: {colors['title']}")
            template_lines.append(f"@color: {colors['text']}")
            template_lines.append("")
            template_lines.append(slide['title'])
            for point in slide['points']:
                template_lines.append(f"- {point}")
            template_lines.append("")
    
    return "\n".join(template_lines), None

def edit_text_with_ai(api_key, current_text, instruction):
    """Редактирование текста через ИИ"""
    
    prompt = f"""У тебя есть текст презентации в специальном формате:

{current_text}

Задача пользователя: {instruction}

Измени текст согласно задаче. Сохрани формат:
- === СЛАЙД X === для разделения слайдов
- @bg:, @font:, @size: и т.д. для настроек
- Первая строка после настроек - заголовок
- Остальные строки с - в начале - пункты

Верни полный обновлённый текст."""
    
    response = call_openai(api_key, prompt)
    return response

# ============================================================
# ИНТЕРФЕЙС
# ============================================================

st.set_page_config(page_title="ИИ Генератор презентаций", page_icon="🤖", layout="wide")

st.title("🤖 ИИ Генератор презентаций")

# Настройка API ключа
with st.sidebar:
    st.header("⚙️ Настройки")
    api_key = st.text_input("OpenAI API Key:", type="password", help="Введите ваш API ключ OpenAI")
    
    if not api_key:
        st.warning("️ Введите API ключ для работы ИИ")
        st.info("Получить ключ: platform.openai.com/api-keys")

if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'edit_history' not in st.session_state:
    st.session_state.edit_history = []
if 'generated_template' not in st.session_state:
    st.session_state.generated_template = ""

# Три колонки
col_gen, col_editor, col_edit = st.columns([1, 1.2, 1])

# ============================================================
# КОЛОНКА 1: ГЕНЕРАЦИЯ
# ============================================================
with col_gen:
    st.subheader("🎨 Генерация презентации")
    
    chat_container = st.container(height=300)
    with chat_container:
        for msg in st.session_state.chat_history:
            if msg['role'] == 'user':
                st.chat_message("user").write(msg['content'])
            else:
                st.chat_message("assistant").write(msg['content'])
    
    user_input = st.chat_input("Тема и кол-во слайдов")
    
    if user_input:
        if not api_key:
            st.session_state.chat_history.append({'role': 'user', 'content': user_input})
            st.session_state.chat_history.append({'role': 'assistant', 'content': '❌ Сначала введите API ключ в настройках слева!'})
            st.rerun()
        
        st.session_state.chat_history.append({'role': 'user', 'content': user_input})
        
        # Парсим запрос
        msg_lower = user_input.lower()
        num_match = re.search(r'(\d+)\s*слайд', msg_lower)
        num_slides = int(num_match.group(1)) if num_match else 5
        
        color_theme = "blue"
        if any(w in msg_lower for w in ['красн', 'red']):
            color_theme = "red"
        elif any(w in msg_lower for w in ['зелён', 'green', 'зелен']):
            color_theme = "green"
        elif any(w in msg_lower for w in ['фиолет', 'purple']):
            color_theme = "purple"
        elif any(w in msg_lower for w in ['тёмн', 'dark', 'темн', 'чёрн', 'черн']):
            color_theme = "dark"
        
        topic = user_input
        remove_words = ['сделай', 'создай', 'генерируй', 'презентацию', 'презентация', 'слайдов', 'слайд', 'про', 'на', 'тему']
        for word in remove_words:
            topic = topic.replace(word, '')
        topic = re.sub(r'\d+', '', topic).strip(' ,.-–—!?:;')
        topic = ' '.join(topic.split())
        if not topic:
            topic = "Презентация"
        topic = topic[0].upper() + topic[1:]
        
        with st.spinner('ИИ генерирует презентацию...'):
            template, error = generate_presentation_with_ai(api_key, topic, num_slides, color_theme)
        
        if error:
            st.session_state.chat_history.append({'role': 'assistant', 'content': f'❌ {error}'})
        else:
            st.session_state.generated_template = template
            st.session_state.chat_history.append({'role': 'assistant', 'content': f'✅ Готово! Создал презентацию «{topic}» на {num_slides} слайдов. Шаблон в редакторе справа!'})
        
        st.rerun()

# ============================================================
# КОЛОНКА 2: РЕДАКТОР
# ============================================================
with col_editor:
    st.subheader("📝 Редактор шаблона")
    
    if st.session_state.generated_template:
        edited_template = st.text_area(
            "Редактируйте текст:",
            value=st.session_state.generated_template,
            height=600
        )
        
        # Кнопки действий
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("💾 Сохранить изменения", use_container_width=True):
                st.session_state.generated_template = edited_template
                st.success("✅ Сохранено!")
        
        with col_btn2:
            if st.button("🎨 Создать презентацию", type="primary", use_container_width=True):
                with st.spinner('Создаю презентацию...'):
                    slides = parse_template(edited_template)
                    
                    if not slides:
                        st.error("❌ Не удалось разобрать шаблон!")
                    else:
                        st.success(f"✅ Создано слайдов: {len(slides)}")
                        pptx_data = create_presentation(slides)
                        
                        st.download_button(
                            label=" Скачать .pptx",
                            data=pptx_data,
                            file_name="presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
    else:
        st.text_area("Здесь появится шаблон. Напишите тему слева.", value="", height=600, disabled=True)

# ============================================================
# КОЛОНКА 3: ИИ-РЕДАКТОР
# ============================================================
with col_edit:
    st.subheader("✏️ ИИ-редактор текста")
    
    edit_container = st.container(height=300)
    with edit_container:
        for msg in st.session_state.edit_history:
            if msg['role'] == 'user':
                st.chat_message("user").write(msg['content'])
            else:
                st.chat_message("assistant").write(msg['content'])
    
    edit_input = st.chat_input("Как изменить текст?")
    
    if edit_input:
        if not api_key:
            st.session_state.edit_history.append({'role': 'user', 'content': edit_input})
            st.session_state.edit_history.append({'role': 'assistant', 'content': '❌ Сначала введите API ключ!'})
            st.rerun()
        
        if not st.session_state.generated_template:
            st.session_state.edit_history.append({'role': 'user', 'content': edit_input})
            st.session_state.edit_history.append({'role': 'assistant', 'content': '❌ Сначала создайте презентацию в левой колонке!'})
            st.rerun()
        
        st.session_state.edit_history.append({'role': 'user', 'content': edit_input})
        
        with st.spinner('ИИ редактирует текст...'):
            new_text = edit_text_with_ai(api_key, st.session_state.generated_template, edit_input)
        
        if new_text.startswith("Ошибка API:"):
            st.session_state.edit_history.append({'role': 'assistant', 'content': f'❌ {new_text}'})
        else:
            st.session_state.generated_template = new_text
            st.session_state.edit_history.append({'role': 'assistant', 'content': '✅ Текст обновлён! Изменения появились в редакторе.'})
        
        st.rerun()
    
    # Быстрые команды редактирования
    st.markdown("**⚡ Быстрые команды:**")
    
    if st.button("Сделай текст короче", use_container_width=True):
        if api_key and st.session_state.generated_template:
            st.session_state.edit_history.append({'role': 'user', 'content': 'Сделай текст короче'})
            new_text = edit_text_with_ai(api_key, st.session_state.generated_template, 'Сократи текст на каждом слайде, оставь только самое важное')
            st.session_state.generated_template = new_text
            st.session_state.edit_history.append({'role': 'assistant', 'content': '✅ Текст сокращён!'})
            st.rerun()
    
    if st.button("Добавь больше фактов", use_container_width=True):
        if api_key and st.session_state.generated_template:
            st.session_state.edit_history.append({'role': 'user', 'content': 'Добавь больше фактов'})
            new_text = edit_text_with_ai(api_key, st.session_state.generated_template, 'Добавь больше конкретных фактов, цифр и деталей на каждый слайд')
            st.session_state.generated_template = new_text
            st.session_state.edit_history.append({'role': 'assistant', 'content': '✅ Добавлены факты!'})
            st.rerun()
    
    if st.button("Смени на красный цвет", use_container_width=True):
        if st.session_state.generated_template:
            new_text = st.session_state.generated_template
            new_text = new_text.replace('#0a192f', '#2d0a0a').replace('#112240', '#4a0f0f').replace('#1a2f4a', '#6b1515').replace('#64ffda', '#ff6b6b').replace('#ccd6f6', '#ffd4d4')
            st.session_state.generated_template = new_text
            st.success('✅ Цвет изменён на красный!')
            st.rerun()

st.markdown("---")
with st.expander("📖 Как пользоваться"):
    st.markdown("""
    ### Шаг 1: Введите API ключ
    В панели слева введите ваш OpenAI API ключ.
    
    ### Шаг 2: Создайте презентацию
    В левой колонке напишите тему, например: "квантовая физика 7 слайдов"
    
    ### Шаг 3: Отредактируйте текст
    - В средней колонке можно редактировать вручную
    - В правой колонке попросите ИИ изменить текст
    
    ### Примеры запросов ИИ-редактору:
    - "Сделай текст короче"
    - "Добавь больше фактов"
    - "Переведи на английский"
    - "Добавь слайд про историю"
    """)
