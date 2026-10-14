import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import re

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
# ИИ-ГЕНЕРАТОР ШАБЛОНОВ (без API ключей)
# ============================================================

# База знаний для генерации презентаций
KNOWLEDGE_BASE = {
    "моцарт": {
        "title": "Вольфганг Амадей Моцарт",
        "slides": [
            ("Ранние годы", [
                "Родился 27 января 1756 года в Зальцбурге",
                "Отец — Леопольд Моцарт, композитор и скрипач",
                "Мать — Анна Мария Моцарт",
                "С детства проявлял необычные музыкальные способности"
            ]),
            ("Музыкальный вундеркинд", [
                "Начал играть на клавесине в 3 года",
                "В 5 лет уже сочинял маленькие пьесы",
                "В 6 лет выступал перед императорским двором",
                "К 8 годам написал свои первые симфонии"
            ]),
            ("Европейские гастроли", [
                "С 1762 года путешествовал по Европе",
                "Выступал в Вене, Париже, Лондоне, Италии",
                "Познакомился с великими музыкантами эпохи",
                "Впитал лучшие традиции европейской музыки"
            ]),
            ("Зрелые годы в Вене", [
                "В 1781 году переехал в Вену",
                "Стал независимым композитором",
                "Давал уроки, концерты, писал на заказ",
                "Создал свои главные шедевры"
            ]),
            ("Великие оперы", [
                "«Свадьба Фигаро» (1786)",
                "«Дон Жуан» (1787)",
                "«Так поступают все» (1790)",
                "«Волшебная флейта» (1791)"
            ]),
            ("Наследие", [
                "Более 600 произведений за короткую жизнь",
                "Вершина венского классицизма",
                "Влияние на Бетховена, Шуберта, Чайковского",
                "Музыка актуальна спустя более 200 лет"
            ])
        ]
    },
    "python": {
        "title": "Язык программирования Python",
        "slides": [
            ("Что такое Python?", [
                "Высокоуровневый язык программирования",
                "Создан Гвидо ван Россумом в 1991 году",
                "Назван в честь шоу «Монти Пайтон»",
                "Один из самых популярных языков в мире"
            ]),
            ("Преимущества", [
                "Простой и понятный синтаксис",
                "Огромное сообщество разработчиков",
                "Тысячи готовых библиотек",
                "Кроссплатформенность"
            ]),
            ("Применение", [
                "Веб-разработка (Django, Flask)",
                "Искусственный интеллект и машинное обучение",
                "Анализ данных (Pandas, NumPy)",
                "Автоматизация задач"
            ]),
            ("Популярные библиотеки", [
                "NumPy — математические вычисления",
                "Pandas — анализ данных",
                "TensorFlow — машинное обучение",
                "Requests — работа с HTTP"
            ]),
            ("Как начать изучение", [
                "Установите Python с python.org",
                "Выберите редактор (VS Code, PyCharm)",
                "Изучите основы: переменные, циклы, функции",
                "Практикуйтесь на простых задачах"
            ])
        ]
    },
    "космос": {
        "title": "Космос и Вселенная",
        "slides": [
            ("Солнечная система", [
                "Солнце — звезда в центре системы",
                "8 планет вращаются вокруг Солнца",
                "Земля — единственная известная планета с жизнью",
                "Есть астероиды, кометы, карликовые планеты"
            ]),
            ("Звёзды и галактики", [
                "Звёзды — огромные шары раскалённого газа",
                "Наша галактика — Млечный Путь",
                "Во Вселенной триллионы галактик",
                "Самая близкая звезда — Проксима Центавра"
            ]),
            ("Исследование космоса", [
                "1957 — первый спутник (СССР)",
                "1961 — Гагарин в космосе",
                "1969 — высадка на Луну",
                "Сегодня работает МКС"
            ]),
            ("Чёрные дыры", [
                "Области с огромной гравитацией",
                "Ничто не может их покинуть, даже свет",
                "Образуются при гибели звёзд",
                "Первое фото получено в 2019 году"
            ]),
            ("Будущее космонавтики", [
                "Полёты на Марс",
                "Космический туризм",
                "Поиск внеземной жизни",
                "Колонизация других планет"
            ])
        ]
    }
}

# Универсальные шаблоны для неизвестных тем
UNIVERSAL_TEMPLATES = {
    "default": [
        ("Введение", [
            "Краткое описание темы",
            "Актуальность и важность",
            "Цели и задачи"
        ]),
        ("Основные понятия", [
            "Ключевое понятие 1",
            "Ключевое понятие 2",
            "Ключевое понятие 3"
        ]),
        ("История", [
            "Начало развития",
            "Ключевые этапы",
            "Современное состояние"
        ]),
        ("Применение", [
            "Область применения 1",
            "Область применения 2",
            "Область применения 3"
        ]),
        ("Преимущества", [
            "Положительная сторона 1",
            "Положительная сторона 2",
            "Положительная сторона 3"
        ]),
        ("Примеры", [
            "Пример 1",
            "Пример 2",
            "Пример 3"
        ]),
        ("Проблемы и вызовы", [
            "Существующая проблема 1",
            "Существующая проблема 2",
            "Пути решения"
        ]),
        ("Будущее", [
            "Тенденции развития",
            "Перспективы",
            "Прогнозы экспертов"
        ]),
        ("Заключение", [
            "Основные выводы",
            "Значение темы",
            "Призыв к действию"
        ])
    ],
    "person": [
        ("Биография", [
            "Дата и место рождения",
            "Семья и детство",
            "Образование"
        ]),
        ("Начало пути", [
            "Первые достижения",
            "Ключевые события",
            "Важные встречи"
        ]),
        ("Главные достижения", [
            "Достижение 1",
            "Достижение 2",
            "Достижение 3"
        ]),
        ("Вклад в развитие", [
            "Вклад в область 1",
            "Вклад в область 2",
            "Влияние на современников"
        ]),
        ("Интересные факты", [
            "Факт 1",
            "Факт 2",
            "Факт 3"
        ]),
        ("Наследие", [
            "Влияние на потомков",
            "Память о человеке",
            "Значение для истории"
        ])
    ],
    "company": [
        ("О компании", [
            "Год основания",
            "Миссия и ценности",
            "Количество сотрудников"
        ]),
        ("Услуги и продукты", [
            "Услуга 1",
            "Услуга 2",
            "Услуга 3"
        ]),
        ("Достижения", [
            "Награды и премии",
            "Ключевые клиенты",
            "Статистика успеха"
        ]),
        ("Команда", [
            "Руководство",
            "Специалисты",
            "Корпоративная культура"
        ]),
        ("Партнёры", [
            "Ключевые партнёры",
            "Сотрудничество",
            "Совместные проекты"
        ]),
        ("Контакты", [
            "Адрес",
            "Телефон и email",
            "Социальные сети"
        ])
    ]
}

def detect_topic_type(topic):
    """Определяет тип темы для выбора шаблона"""
    topic_lower = topic.lower()
    
    # Ключевые слова для определения типа
    person_keywords = ['человек', 'личность', 'биография', 'писатель', 'учёный', 'композитор', 'художник', 'политик']
    company_keywords = ['компания', 'бизнес', 'организация', 'фирма', 'стартап']
    
    for kw in person_keywords:
        if kw in topic_lower:
            return "person"
    for kw in company_keywords:
        if kw in topic_lower:
            return "company"
    
    return "default"

def generate_template_from_topic(topic, num_slides=8, color_theme="blue"):
    """Генерирует шаблон презентации по теме"""
    
    # Цветовые темы
    color_themes = {
        "blue": {"bg1": "#0a192f", "bg2": "#112240", "bg3": "#1a2f4a", "title": "#64ffda", "text": "#ccd6f6"},
        "red": {"bg1": "#2d0a0a", "bg2": "#4a0f0f", "bg3": "#6b1515", "title": "#ff6b6b", "text": "#ffd4d4"},
        "green": {"bg1": "#0a2d1a", "bg2": "#0f4a2a", "bg3": "#156b3d", "title": "#6bff9e", "text": "#d4ffd4"},
        "purple": {"bg1": "#1a0a2d", "bg2": "#2a0f4a", "bg3": "#3d156b", "title": "#c56bff", "text": "#e8d4ff"},
        "dark": {"bg1": "#1a1a1a", "bg2": "#2d2d2d", "bg3": "#404040", "title": "#ffffff", "text": "#cccccc"},
    }
    
    colors = color_themes.get(color_theme, color_themes["blue"])
    
    # Проверяем базу знаний
    topic_lower = topic.lower()
    matched_topic = None
    for key in KNOWLEDGE_BASE:
        if key in topic_lower:
            matched_topic = key
            break
    
    if matched_topic:
        data = KNOWLEDGE_BASE[matched_topic]
        slides_data = data["slides"]
        main_title = data["title"]
    else:
        # Используем универсальный шаблон
        topic_type = detect_topic_type(topic)
        slides_data = UNIVERSAL_TEMPLATES[topic_type]
        main_title = topic
    
    # Ограничиваем количество слайдов
    slides_data = slides_data[:num_slides]
    
    # Генерируем текст шаблона
    bg_colors = [colors["bg1"], colors["bg2"], colors["bg3"]]
    
    template_lines = []
    
    # Титульный слайд
    template_lines.append("=== СЛАЙД 1 ===")
    template_lines.append(f"@bg: {colors['bg1']}")
    template_lines.append("@font: Georgia")
    template_lines.append(f"@title_size: 44")
    template_lines.append(f"@title_color: {colors['title']}")
    template_lines.append(f"@color: {colors['text']}")
    template_lines.append("")
    template_lines.append(main_title)
    template_lines.append(f"Презентация на тему: {topic}")
    template_lines.append("")
    
    # Остальные слайды
    for i, (slide_title, points) in enumerate(slides_data, 2):
        bg = bg_colors[(i - 2) % 3]
        template_lines.append(f"=== СЛАЙД {i} ===")
        template_lines.append(f"@bg: {bg}")
        template_lines.append("@font: Arial")
        template_lines.append("@size: 22")
        template_lines.append("@title_size: 34")
        template_lines.append(f"@title_color: {colors['title']}")
        template_lines.append(f"@color: {colors['text']}")
        template_lines.append("")
        template_lines.append(slide_title)
        for point in points:
            template_lines.append(f"- {point}")
        template_lines.append("")
    
    return "\n".join(template_lines)

def ai_chat_response(user_message, chat_history):
    """Имитация ИИ-ответа (умный генератор)"""
    message_lower = user_message.lower()
    
    # Приветствие
    if any(word in message_lower for word in ['привет', 'здравствуй', 'hi', 'hello']):
        return "Привет! 👋 Я помогу вам создать презентацию. Напишите тему, и я сгенерирую шаблон. Например: «Сделай презентацию про космос на 8 слайдов»"
    
    # Помощь
    if any(word in message_lower for word in ['помощь', 'help', 'как']):
        return """📋 **Как пользоваться:**

1. Напишите тему презентации
2. Укажите количество слайдов (по умолчанию 8)
3. Выберите цветовую тему (синий, красный, зелёный, фиолетовый, тёмный)

**Примеры запросов:**
- «Сделай презентацию про Моцарта»
- «Презентация про Python на 10 слайдов»
- «Создай презентацию про космос, красная тема»
- «Презентация про компанию, 6 слайдов, зелёная»"""
    
    # Генерация презентации
    if any(word in message_lower for word in ['сделай', 'создай', 'генерируй', 'презентацию', 'слайд']):
        # Извлекаем тему
        topic = user_message
        
        # Убираем служебные слова
        remove_words = ['сделай', 'создай', 'генерируй', 'презентацию', 'про', 'на', 'тему', 'тема']
        for word in remove_words:
            topic = topic.replace(word, '')
        
        # Извлекаем количество слайдов
        num_match = re.search(r'(\d+)\s*слайд', message_lower)
        num_slides = int(num_match.group(1)) if num_match else 8
        
        # Извлекаем цветовую тему
        color_theme = "blue"
        if 'красн' in message_lower:
            color_theme = "red"
        elif 'зелён' in message_lower or 'green' in message_lower:
            color_theme = "green"
        elif 'фиолет' in message_lower or 'purple' in message_lower:
            color_theme = "purple"
        elif 'тёмн' in message_lower or 'dark' in message_lower or 'чёрн' in message_lower:
            color_theme = "dark"
        
        topic = topic.strip()
        if not topic:
            topic = "Выбранная тема"
        
        # Генерируем шаблон
        template = generate_template_from_topic(topic, num_slides, color_theme)
        
        return f"✅ Готово! Создал шаблон презентации **«{topic}»** на {num_slides} слайдов ({color_theme} тема).\n\nШаблон появился в поле ввода ниже. Можете отредактировать его и нажать **«Создать презентацию»**."
    
    # По умолчанию
    return "Я вас не совсем понял. Напишите, например: «Сделай презентацию про космос на 8 слайдов» или напишите «помощь» для списка команд."

# ============================================================
# ИНТЕРФЕЙС STREAMLIT
# ============================================================

st.set_page_config(page_title="ИИ Генератор презентаций", page_icon="🤖", layout="wide")

st.title("🤖 ИИ Генератор презентаций")

# Инициализация session_state
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'generated_template' not in st.session_state:
    st.session_state.generated_template = ""

# Две колонки: чат слева, редактор справа
col_chat, col_editor = st.columns([1, 1])

# ============================================================
# ЛЕВАЯ КОЛОНКА: ЧАТ
# ============================================================
with col_chat:
    st.subheader("💬 Чат с ИИ-помощником")
    
    # Отображаем историю чата
    chat_container = st.container(height=400)
    with chat_container:
        for msg in st.session_state.chat_history:
            if msg['role'] == 'user':
                st.chat_message("user").write(msg['content'])
            else:
                st.chat_message("assistant").write(msg['content'])
    
    # Поле ввода для чата
    user_input = st.chat_input("Напишите тему презентации...")
    
    if user_input:
        # Добавляем сообщение пользователя
        st.session_state.chat_history.append({'role': 'user', 'content': user_input})
        
        # Получаем ответ ИИ
        response = ai_chat_response(user_input, st.session_state.chat_history)
        st.session_state.chat_history.append({'role': 'assistant', 'content': response})
        
        # Если это запрос на генерацию - создаём шаблон
        if any(word in user_input.lower() for word in ['сделай', 'создай', 'генерируй', 'презентацию']):
            # Извлекаем параметры
            topic = user_input
            remove_words = ['сделай', 'создай', 'генерируй', 'презентацию', 'про', 'на', 'тему', 'тема']
            for word in remove_words:
                topic = topic.replace(word, '')
            
            num_match = re.search(r'(\d+)\s*слайд', user_input.lower())
            num_slides = int(num_match.group(1)) if num_match else 8
            
            color_theme = "blue"
            if 'красн' in user_input.lower():
                color_theme = "red"
            elif 'зелён' in user_input.lower() or 'green' in user_input.lower():
                color_theme = "green"
            elif 'фиолет' in user_input.lower() or 'purple' in user_input.lower():
                color_theme = "purple"
            elif 'тёмн' in user_input.lower() or 'dark' in user_input.lower():
                color_theme = "dark"
            
            topic = topic.strip()
            if not topic:
                topic = "Презентация"
            
            # Генерируем шаблон и сохраняем
            template = generate_template_from_topic(topic, num_slides, color_theme)
            st.session_state.generated_template = template
        
        st.rerun()
    
    # Кнопки быстрых команд
    st.markdown("**⚡ Быстрые команды:**")
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    
    with col_btn1:
        if st.button("🎵 Моцарт", use_container_width=True):
            st.session_state.chat_history.append({'role': 'user', 'content': 'Сделай презентацию про Моцарта на 8 слайдов'})
            response = ai_chat_response('Сделай презентацию про Моцарта на 8 слайдов', st.session_state.chat_history)
            st.session_state.chat_history.append({'role': 'assistant', 'content': response})
            template = generate_template_from_topic("Моцарт", 8, "blue")
            st.session_state.generated_template = template
            st.rerun()
    
    with col_btn2:
        if st.button("🐍 Python", use_container_width=True):
            st.session_state.chat_history.append({'role': 'user', 'content': 'Сделай презентацию про Python на 6 слайдов'})
            response = ai_chat_response('Сделай презентацию про Python на 6 слайдов', st.session_state.chat_history)
            st.session_state.chat_history.append({'role': 'assistant', 'content': response})
            template = generate_template_from_topic("Python", 6, "green")
            st.session_state.generated_template = template
            st.rerun()
    
    with col_btn3:
        if st.button(" Космос", use_container_width=True):
            st.session_state.chat_history.append({'role': 'user', 'content': 'Сделай презентацию про космос на 7 слайдов'})
            response = ai_chat_response('Сделай презентацию про космос на 7 слайдов', st.session_state.chat_history)
            st.session_state.chat_history.append({'role': 'assistant', 'content': response})
            template = generate_template_from_topic("космос", 7, "purple")
            st.session_state.generated_template = template
            st.rerun()
    
    # Кнопка очистки чата
    if st.button("🗑️ Очистить чат", use_container_width=True):
        st.session_state.chat_history = []
        st.rerun()

# ============================================================
# ПРАВАЯ КОЛОНКА: РЕДАКТОР ШАБЛОНА
# ============================================================
with col_editor:
    st.subheader("📝 Редактор шаблона")
    
    # Если есть сгенерированный шаблон - показываем его
    if st.session_state.generated_template:
        edited_template = st.text_area(
            "Отредактируйте шаблон при необходимости:",
            value=st.session_state.generated_template,
            height=500
        )
    else:
        edited_template = st.text_area(
            "Здесь появится сгенерированный шаблон. Напишите тему в чате слева.",
            value="",
            height=500,
            disabled=True
        )
    
    # Кнопка создания презентации
    if st.button(" Создать презентацию", type="primary", use_container_width=True):
        if edited_template.strip():
            with st.spinner('Создаю презентацию...'):
                slides = parse_template(edited_template)
                
                if not slides:
                    st.error("❌ Не удалось разобрать шаблон!")
                else:
                    st.success(f"✅ Создано слайдов: {len(slides)}")
                    
                    pptx_data = create_presentation(slides)
                    
                    st.download_button(
                        label="📥 Скачать .pptx",
                        data=pptx_data,
                        file_name="presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
        else:
            st.warning("⚠️ Сначала сгенерируйте шаблон в чате!")

# ============================================================
# ИНСТРУКЦИЯ ВНИЗУ
# ============================================================
st.markdown("---")
with st.expander("📖 Как пользоваться"):
    st.markdown("""
    ### 🤖 Работа с ИИ-чатом:
    
    1. **Напишите тему** в чат слева, например:
       - «Сделай презентацию про Моцарта»
       - «Создай презентацию про Python на 10 слайдов»
       - «Презентация про компанию, красная тема»
    
    2. **ИИ сгенерирует шаблон** — он появится в редакторе справа
    
    3. **Отредактируйте** текст при необходимости
    
    4. **Нажмите «Создать презентацию»** и скачайте файл
    
    ###  Цветовые темы:
    - **Синяя** (по умолчанию) — `@bg: #0a192f`
    - **Красная** — напишите «красная тема»
    - **Зелёная** — напишите «зелёная тема»
    - **Фиолетовая** — напишите «фиолетовая тема»
    - **Тёмная** — напишите «тёмная тема»
    
    ### ️ Ручное редактирование:
    Вы можете менять текст, цвета, шрифты и размеры прямо в редакторе!
    """)
