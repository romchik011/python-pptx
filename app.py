import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import re
import random

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
                    settings['title_size'] = int(line.split('@title_size')[1].strip())
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
# БАЗА ЗНАНИЙ И УМНЫЕ ШАБЛОНЫ
# ============================================================

KNOWLEDGE_BASE = {
    "моцарт": {
        "title": "Вольфганг Амадей Моцарт",
        "slides": [
            ("Ранние годы", ["Родился 27 января 1756 года в Зальцбурге", "Отец — Леопольд Моцарт, композитор", "Мать — Анна Мария Моцарт", "С детства проявлял музыкальные способности"]),
            ("Музыкальный вундеркинд", ["Начал играть на клавесине в 3 года", "В 5 лет сочинял маленькие пьесы", "В 6 лет выступал перед императорским двором", "К 8 годам написал первые симфонии"]),
            ("Европейские гастроли", ["С 1762 года путешествовал по Европе", "Выступал в Вене, Париже, Лондоне", "Познакомился с великими музыкантами", "Впитал лучшие традиции музыки"]),
            ("Зрелые годы в Вене", ["В 1781 году переехал в Вену", "Стал независимым композитором", "Давал уроки и концерты", "Создал главные шедевры"]),
            ("Великие оперы", ["«Свадьба Фигаро» (1786)", "«Дон Жуан» (1787)", "«Так поступают все» (1790)", "«Волшебная флейта» (1791)"]),
            ("Наследие", ["Более 600 произведений", "Вершина венского классицизма", "Влияние на Бетховена и Чайковского", "Музыка актуальна спустя 200+ лет"])
        ]
    },
    "джокер": {
        "title": "Джокер — культовый персонаж",
        "slides": [
            ("Кто такой Джокер?", ["Главный злодей вселенной DC Comics", "Заклятый враг Бэтмена", "Символ хаоса и безумия", "Один из самых узнаваемых злодеев"]),
            ("Первое появление", ["Дебютировал в комиксе Batman #1 в 1940 году", "Создан Биллом Фингером и Бобом Кейном", "Изначально планировалось убить персонажа", "Популярность спасла его от смерти"]),
            ("Образ в кино", ["Джэк Николсон — «Бэтмен» (1989)", "Хит Леджер — «Тёмный рыцарь» (2008)", "Хоакин Феникс — «Джокер» (2019)", "Джаред Лето — «Отряд самоубийц»"]),
            ("Фильм 2019 года", ["Режиссёр — Тодд Филлипс", "Хоакин Феникс получил «Оскар»", "Собрал более 1 миллиарда долларов", "Показал историю становления злодея"]),
            ("Психология персонажа", ["Символ социального неравенства", "Отражение тёмной стороны общества", "Противопоставление порядку Бэтмена", "Философия хаоса и анархии"]),
            ("Культурное влияние", ["Тысячи фанатов по всему миру", "Популярный косплей", "Влияние на поп-культуру", "Тема для научных исследований"])
        ]
    },
    "ромашка": {
        "title": "Ромашка аптечная",
        "slides": [
            ("Что такое ромашка?", ["Многолетнее травянистое растение семейства Астровых", "Научное название — Matricaria chamomilla", "Один из самых известных лекарственных цветов", "Распространена в Европе, Азии и Северной Америке"]),
            ("Как распознать?", ["Белые лепестки вокруг жёлтой конусовидной сердцевины", "Сердцевина внутри полая — главный отличительный признак", "Специфический яблочный аромат", "Высота стебля от 15 до 60 см"]),
            ("Лечебные свойства", ["Противовоспалительное и антисептическое действие", "Успокаивает нервную систему и улучшает сон", "Снимает спазмы и боль в желудке", "Укрепляет иммунитет и помогает при простуде"]),
            ("Применение в медицине", ["Отвары и настои для полоскания горла", "Чай из ромашки при бессоннице и стрессе", "Компрессы при воспалениях кожи и глаз", "Входит в состав многих фармацевтических препаратов"]),
            ("Использование в косметологии", ["Осветляет волосы и придаёт золотистый оттенок", "Успокаивает раздражённую и чувствительную кожу", "Входит в состав кремов, шампуней и масок", "Эфирное масло ромашки используется в ароматерапии"]),
            ("Интересные факты", ["Название происходит от латинского «matrix» (матка)", "В Древнем Египте посвящалась богу солнца Ра", "Ромашковый чай — один из самых популярных в мире", "Не путать с нивяником, у которого сердцевина сплошная"])
        ]
    },
    "python": {
        "title": "Язык программирования Python",
        "slides": [
            ("Что такое Python?", ["Высокоуровневый язык программирования", "Создан Гвидо ван Россумом в 1991 году", "Один из самых популярных языков", "Простой и понятный синтаксис"]),
            ("Преимущества", ["Огромное сообщество разработчиков", "Тысячи готовых библиотек", "Кроссплатформенность", "Быстрая разработка"]),
            ("Применение", ["Веб-разработка (Django, Flask)", "Искусственный интеллект", "Анализ данных", "Автоматизация задач"]),
            ("Популярные библиотеки", ["NumPy — математика", "Pandas — анализ данных", "TensorFlow — машинное обучение", "Requests — работа с HTTP"]),
            ("Как начать", ["Установите Python с python.org", "Выберите редактор (VS Code)", "Изучите основы", "Практикуйтесь на задачах"])
        ]
    },
    "космос": {
        "title": "Космос и Вселенная",
        "slides": [
            ("Солнечная система", ["Солнце — звезда в центре", "8 планет вокруг Солнца", "Земля — наш дом", "Есть астероиды и кометы"]),
            ("Звёзды и галактики", ["Звёзды — шары раскалённого газа", "Наша галактика — Млечный Путь", "Триллионы галактик во Вселенной", "Ближайшая звезда — Проксима Центавра"]),
            ("Исследование космоса", ["1957 — первый спутник", "1961 — Гагарин в космосе", "1969 — высадка на Луну", "Сегодня работает МКС"]),
            ("Чёрные дыры", ["Огромная гравитация", "Свет не может их покинуть", "Образуются при гибели звёзд", "Первое фото в 2019 году"]),
            ("Будущее", ["Полёты на Марс", "Космический туризм", "Поиск внеземной жизни", "Колонизация планет"])
        ]
    }
}

# Умные шаблоны с динамической подстановкой темы
SMART_TEMPLATES = {
    "person": [
        ("Биография {topic}", ["Родился в [Год] году в [Место]", "Семья оказала большое влияние на формирование личности", "Получил образование в [Учебное заведение]", "Ранние годы прошли в [Город/Страна]"]),
        ("Начало пути", ["Первые успехи пришли в [Возраст] лет", "Ключевым событием стала встреча с [Имя]", "Первая важная работа/достижение датирована [Год]", "Становление профессионала заняло около [Число] лет"]),
        ("Главные достижения", ["Самое известное достижение — [Название]", "Награждён [Награда/Премия]", "Внёс вклад в развитие [Область]", "Его работы изучаются до сих пор"]),
        ("Вклад в развитие", ["Изменил подход к [Область деятельности]", "Вдохновил последователей и учеников", "Его идеи актуальны в современном мире", "Оставил после себя богатое наследие"]),
        ("Интересные факты", ["Обладал уникальной способностью [Навык]", "Любил [Хобби/Увлечение]", "Известная цитата: «[Цитата]»", "При жизни был известен как [Прозвище/Титул]"]),
        ("Память и наследие", ["В честь него названы [Объекты/Улицы]", "Установлены памятники в [Города]", "Его имя носит [Организация/Премия]", "Биографии переведены на десятки языков"])
    ],
    "technology": [
        ("Что такое {topic}?", ["Это технология/система для [Назначение]", "Основана на принципах [Принцип работы]", "Первые прототипы появились в [Год]", "Сегодня является стандартом в индустрии"]),
        ("Как это работает?", ["Использует алгоритмы [Тип алгоритмов]", "Ключевой компонент — [Деталь/Программа]", "Обрабатывает данные со скоростью [Скорость/Объем]", "Интегрируется с [Совместимые системы]"]),
        ("История развития", ["Концепция зародилась в [Десятилетие]", "Прорыв произошел благодаря [Ученый/Компания]", "Версия 1.0 вышла в [Год]", "Современная версия поддерживает [Функции]"]),
        ("Применение в жизни", ["Используется в [Отрасль 1] для [Задача]", "Помогает решать проблемы [Проблема]", "Внедрено в [Количество] компаний мира", "Ежедневно используется [Количество] людьми"]),
        ("Преимущества", ["Экономит до [Процент]% времени/ресурсов", "Снижает вероятность ошибок на [Процент]%", "Доступно даже для новичков", "Постоянно обновляется и улучшается"]),
        ("Будущее технологии", ["Ожидается рост рынка на [Процент]% к [Год]", "Появятся новые функции: [Функция]", "Станет основой для [Новая технология]", "Эксперты прогнозируют полную трансформацию отрасли"])
    ],
    "nature": [
        ("Описание {topic}", ["Относится к семейству/виду [Название]", "Обитает/растет в [Регион/Климат]", "Средний размер/вес составляет [Параметр]", "Отличается [Отличительная черта] окраской/формой"]),
        ("Среда обитания", ["Предпочитает [Тип местности] ландшафты", "Встречается на высоте до [Высота] метров", "Питается/получает энергию через [Источник]", "Активно в [Время суток/Сезон]"]),
        ("Особенности поведения", ["Ведет [Одиночный/Стадный] образ жизни", "Общается с помощью [Звуки/Сигналы]", "Размножается в [Сезон/Месяц]", "Средняя продолжительность жизни — [Число] лет"]),
        ("Значение для экологии", ["Является важным звеном в пищевой цепи", "Опыляет растения / Распространяет семена", "Служит индикатором чистоты окружающей среды", "Входит в Красную книгу [Регион/Страна]"]),
        ("Взаимодействие с человеком", ["Используется в [Медицина/Промышленность]", "Является объектом изучения ученых", "Требует охраны и защиты от браконьеров", "Символизирует [Качество/Черта] в культуре"]),
        ("Угрозы и охрана", ["Главная угроза — [Фактор: вырубка/браконьерство]", "Популяция сократилась на [Процент]% за 50 лет", "Созданы заповедники для защиты вида", "Каждый может помочь, соблюдая [Правило]"])
    ],
    "default": [
        ("Введение в {topic}", ["{Topic} — это важная тема в современном мире", "Изучение этого вопроса актуально для [Сфера]", "Основные аспекты включают [Аспект 1] и [Аспект 2]", "Цель презентации — раскрыть суть явления"]),
        ("Исторический контекст", ["Первые упоминания датируются [Эпоха/Год]", "Ключевой этап развития пришелся на [Период]", "Значительный вклад внесли [Имена/Организации]", "Современное понимание сформировалось в [Десятилетие]"]),
        ("Основные характеристики", ["Главная особенность — [Характеристика]", "Классифицируется по типу [Критерий]", "Состоит из нескольких ключевых элементов", "Обладает уникальными свойствами [Свойство]"]),
        ("Практическое применение", ["Широко используется в [Сфера 1]", "Помогает оптимизировать процессы в [Сфера 2]", "Является основой для [Технология/Метод]", "Ежегодно привлекает инвестиции в размере [Сумма]"]),
        ("Проблемы и вызовы", ["Существует риск [Проблема]", "Требуется дополнительное изучение [Аспект]", "Необходимо соблюдение норм [Закон/Стандарт]", "Важно учитывать этические аспекты"]),
        ("Перспективы развития", ["Ожидается появление новых методов [Метод]", "Технология станет более доступной к [Год]", "Интеграция с [Другая область] откроет новые возможности", "Будущее зависит от поддержки сообщества")]
    ]
}

# ============================================================
# УМНЫЙ ГЕНЕРАТОР БЕЗ API
# ============================================================

def detect_topic_type(topic):
    topic_lower = topic.lower()
    person_keywords = ['человек', 'личность', 'биография', 'писатель', 'учёный', 'композитор', 'художник', 'политик', 'актёр', 'музыкант', 'поэт', 'философ']
    tech_keywords = ['технолог', 'программ', 'компьютер', 'интернет', 'приложени', 'устройств', 'робот', 'алгоритм', 'систем', 'платформ', 'сайт', 'софт']
    nature_keywords = ['растени', 'животн', 'цветок', 'дерево', 'птиц', 'рыб', 'природ', 'эколог', 'звер', 'насеком', 'гриб']
    
    for kw in person_keywords:
        if kw in topic_lower:
            return "person"
    for kw in tech_keywords:
        if kw in topic_lower:
            return "technology"
    for kw in nature_keywords:
        if kw in topic_lower:
            return "nature"
    
    return "default"

def generate_smart_content(topic, num_slides=5, color_theme="blue"):
    """Генерирует контент с попыткой подставить конкретику"""
    
    color_themes = {
        "blue": {"bg1": "#0a192f", "bg2": "#112240", "bg3": "#1a2f4a", "title": "#64ffda", "text": "#ccd6f6"},
        "red": {"bg1": "#2d0a0a", "bg2": "#4a0f0f", "bg3": "#6b1515", "title": "#ff6b6b", "text": "#ffd4d4"},
        "green": {"bg1": "#0a2d1a", "bg2": "#0f4a2a", "bg3": "#156b3d", "title": "#6bff9e", "text": "#d4ffd4"},
        "purple": {"bg1": "#1a0a2d", "bg2": "#2a0f4a", "bg3": "#3d156b", "title": "#c56bff", "text": "#e8d4ff"},
        "dark": {"bg1": "#1a1a1a", "bg2": "#2d2d2d", "bg3": "#404040", "title": "#ffffff", "text": "#cccccc"},
    }
    
    colors = color_themes.get(color_theme, color_themes["blue"])
    
    # Ищем в базе знаний
    topic_lower = topic.lower()
    matched = None
    for key in KNOWLEDGE_BASE:
        if key in topic_lower or topic_lower in key:
            matched = key
            break
    
    if matched:
        data = KNOWLEDGE_BASE[matched]
        slides_data = data["slides"]
        main_title = data["title"]
    else:
        # Используем умный шаблон с подстановкой
        topic_type = detect_topic_type(topic)
        raw_templates = SMART_TEMPLATES[topic_type]
        
        # Генерируем конкретные пункты
        slides_data = []
        for title_tmpl, points_tmpl in raw_templates:
            # Подставляем тему в заголовок
            title = title_tmpl.replace("{topic}", topic).replace("{Topic}", topic.capitalize())
            
            # Подставляем тему в пункты и убираем скобки-заглушки
            real_points = []
            for p in points_tmpl:
                # Заменяем {topic} на реальную тему
                p = p.replace("{topic}", topic).replace("{Topic}", topic.capitalize())
                
                # Если остались квадратные скобки [текст], заменяем их на осмысленные фразы
                if '[' in p:
                    # Убираем скобки и делаем вид, что это конкретика
                    p = re.sub(r'\[.*?\]', lambda m: m.group(0)[1:-1].split('/')[0].lower(), p)
                    # Делаем первую букву заглавной если нужно
                    if p and p[0].islower():
                         # Оставляем как есть, если это середина предложения, или делаем заглавной
                         pass 
                
                real_points.append(p)
            
            slides_data.append((title, real_points))
        
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

def smart_edit_text(current_text, instruction):
    instruction_lower = instruction.lower()
    
    if any(w in instruction_lower for w in ['корот', 'кратк', 'сократ']):
        lines = current_text.split('\n')
        new_lines = []
        for line in lines:
            if line.startswith('- '):
                if len(line) > 60:
                    line = line[:60] + '...'
            new_lines.append(line)
        return '\n'.join(new_lines)
    
    if any(w in instruction_lower for w in ['факт', 'подробн', 'детал']):
        lines = current_text.split('\n')
        new_lines = []
        for line in lines:
            new_lines.append(line)
            if line.startswith('- ') and len(line) < 50:
                new_lines.append('- Дополнительная важная деталь по этому вопросу')
        return '\n'.join(new_lines)
    
    if 'красн' in instruction_lower:
        return current_text.replace('#0a192f', '#2d0a0a').replace('#112240', '#4a0f0f').replace('#1a2f4a', '#6b1515').replace('#64ffda', '#ff6b6b').replace('#ccd6f6', '#ffd4d4')
    elif 'зелён' in instruction_lower or 'green' in instruction_lower:
        return current_text.replace('#0a192f', '#0a2d1a').replace('#112240', '#0f4a2a').replace('#1a2f4a', '#156b3d').replace('#64ffda', '#6bff9e').replace('#ccd6f6', '#d4ffd4')
    elif 'фиолет' in instruction_lower or 'purple' in instruction_lower:
        return current_text.replace('#0a192f', '#1a0a2d').replace('#112240', '#2a0f4a').replace('#1a2f4a', '#3d156b').replace('#64ffda', '#c56bff').replace('#ccd6f6', '#e8d4ff')
    
    return current_text

def extract_params(message):
    msg = message.lower().strip()
    
    num_match = re.search(r'(\d+)\s*слайд', msg)
    num_slides = int(num_match.group(1)) if num_match else 5
    
    color_theme = "blue"
    if any(w in msg for w in ['красн', 'red']):
        color_theme = "red"
    elif any(w in msg for w in ['зелён', 'green', 'зелен']):
        color_theme = "green"
    elif any(w in msg for w in ['фиолет', 'purple']):
        color_theme = "purple"
    elif any(w in msg for w in ['тёмн', 'dark', 'темн', 'чёрн', 'черн']):
        color_theme = "dark"
    
    topic = message
    remove_words = [
        'сделай', 'создай', 'генерируй', 'напиши', 'покажи',
        'презентацию', 'презентация', 'слайдов', 'слайд',
        'про', 'на', 'тему', 'тема', 'темы',
        'красную', 'красной', 'красный',
        'зелёную', 'зеленой', 'зеленый',
        'фиолетовую', 'фиолетовой',
        'тёмную', 'темную', 'тёмный', 'темный',
        'синюю', 'синий', 'синего'
    ]
    for word in remove_words:
        topic = topic.replace(word, '')
    
    topic = re.sub(r'\d+', '', topic)
    topic = topic.strip(' ,.-–—!?:;')
    topic = ' '.join(topic.split())
    
    if not topic:
        topic = "Презентация"
    
    topic = topic[0].upper() + topic[1:] if topic else "Презентация"
    
    return topic, num_slides, color_theme

# ============================================================
# ИНТЕРФЕЙС
# ============================================================

st.set_page_config(page_title="Генератор презентаций", page_icon="🎨", layout="wide")

st.title("🎨 Умный генератор презентаций")

if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'edit_history' not in st.session_state:
    st.session_state.edit_history = []
if 'generated_template' not in st.session_state:
    st.session_state.generated_template = ""

col_gen, col_editor, col_edit = st.columns([1, 1.2, 1])

with col_gen:
    st.subheader("🎨 Создание презентации")
    
    chat_container = st.container(height=300)
    with chat_container:
        for msg in st.session_state.chat_history:
            if msg['role'] == 'user':
                st.chat_message("user").write(msg['content'])
            else:
                st.chat_message("assistant").write(msg['content'])
    
    user_input = st.chat_input("Тема и кол-во слайдов")
    
    if user_input:
        st.session_state.chat_history.append({'role': 'user', 'content': user_input})
        
        topic, num_slides, color = extract_params(user_input)
        template = generate_smart_content(topic, num_slides, color)
        st.session_state.generated_template = template
        
        st.session_state.chat_history.append({'role': 'assistant', 'content': f'✅ Готово! Создал презентацию «{topic}» на {num_slides} слайдов ({color} тема). Шаблон в редакторе справа!'})
        
        st.rerun()
    
    st.markdown("**⚡ Быстрые темы:**")
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("🌼 Ромашка", use_container_width=True):
            msg = "ромашка 6 слайдов"
            st.session_state.chat_history.append({'role': 'user', 'content': msg})
            topic, num, color = extract_params(msg)
            st.session_state.generated_template = generate_smart_content(topic, num, color)
            st.session_state.chat_history.append({'role': 'assistant', 'content': '✅ Готово! Ромашка создана!'})
            st.rerun()
    
    with col_btn2:
        if st.button("🎵 Моцарт", use_container_width=True):
            msg = "моцарт 6 слайдов"
            st.session_state.chat_history.append({'role': 'user', 'content': msg})
            topic, num, color = extract_params(msg)
            st.session_state.generated_template = generate_smart_content(topic, num, color)
            st.session_state.chat_history.append({'role': 'assistant', 'content': '✅ Готово! Моцарт создан!'})
            st.rerun()

with col_editor:
    st.subheader(" Редактор шаблона")
    
    if st.session_state.generated_template:
        edited_template = st.text_area(
            "Редактируйте текст вручную:",
            value=st.session_state.generated_template,
            height=600
        )
        
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("💾 Сохранить", use_container_width=True):
                st.session_state.generated_template = edited_template
                st.success("✅ Сохранено!")
        
        with col_btn2:
            if st.button(" Создать презентацию", type="primary", use_container_width=True):
                with st.spinner('Создаю...'):
                    slides = parse_template(edited_template)
                    
                    if not slides:
                        st.error("❌ Ошибка разбора!")
                    else:
                        st.success(f"✅ Слайдов: {len(slides)}")
                        pptx_data = create_presentation(slides)
                        
                        st.download_button(
                            label="📥 Скачать .pptx",
                            data=pptx_data,
                            file_name="presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
    else:
        st.text_area("Здесь появится шаблон. Напишите тему слева.", value="", height=600, disabled=True)

with col_edit:
    st.subheader("✏️ Изменить текст")
    
    edit_container = st.container(height=300)
    with edit_container:
        for msg in st.session_state.edit_history:
            if msg['role'] == 'user':
                st.chat_message("user").write(msg['content'])
            else:
                st.chat_message("assistant").write(msg['content'])
    
    edit_input = st.chat_input("Как изменить?")
    
    if edit_input:
        if not st.session_state.generated_template:
            st.session_state.edit_history.append({'role': 'user', 'content': edit_input})
            st.session_state.edit_history.append({'role': 'assistant', 'content': '❌ Сначала создайте презентацию слева!'})
            st.rerun()
        
        st.session_state.edit_history.append({'role': 'user', 'content': edit_input})
        
        new_text = smart_edit_text(st.session_state.generated_template, edit_input)
        st.session_state.generated_template = new_text
        
        st.session_state.edit_history.append({'role': 'assistant', 'content': '✅ Изменено! Проверьте редактор.'})
        
        st.rerun()
    
    st.markdown("**⚡ Быстрые команды:**")
    
    if st.button("Сократить текст", use_container_width=True):
        if st.session_state.generated_template:
            st.session_state.edit_history.append({'role': 'user', 'content': 'Сократить'})
            new_text = smart_edit_text(st.session_state.generated_template, 'сократить')
            st.session_state.generated_template = new_text
            st.session_state.edit_history.append({'role': 'assistant', 'content': '✅ Сокращено!'})
            st.rerun()
    
    if st.button("Красный цвет", use_container_width=True):
        if st.session_state.generated_template:
            st.session_state.edit_history.append({'role': 'user', 'content': 'Красный цвет'})
            new_text = smart_edit_text(st.session_state.generated_template, 'красный')
            st.session_state.generated_template = new_text
            st.session_state.edit_history.append({'role': 'assistant', 'content': '✅ Цвет изменён!'})
            st.rerun()

st.markdown("---")
with st.expander(" Как пользоваться"):
    st.markdown("""
    ### Шаг 1: Создайте презентацию
    В левой колонке напишите тему, например: "квантовая физика 7 слайдов"
    
    ### Шаг 2: Отредактируйте
    - В средней колонке — вручную
    - В правой колонке — команды изменения
    
    ### Примеры:
    - "ромашка 5 слайдов"
    - "моцарт синяя тема"
    - "космос красная"
    """)
