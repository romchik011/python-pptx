from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from code_parser import CodeElement

THEMES = {
    "light": {
        "bg": RGBColor(255, 255, 255),
        "title": RGBColor(44, 62, 80),
        "text": RGBColor(52, 73, 94),
        "accent": RGBColor(52, 152, 219),
        "muted": RGBColor(127, 140, 141),
    },
    "dark": {
        "bg": RGBColor(44, 62, 80),
        "title": RGBColor(236, 240, 241),
        "text": RGBColor(189, 195, 199),
        "accent": RGBColor(52, 152, 219),
        "muted": RGBColor(149, 165, 166),
    },
}


def set_slide_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, 
                color=None, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    p.alignment = align
    return txBox


def build_presentation_from_code(root: CodeElement, elements: list, 
                                  title: str = "Документация проекта", 
                                  author: str = "", theme_name: str = "light") -> BytesIO:
    theme = THEMES[theme_name]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # 1. Титульный слайд
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide, theme["bg"])
    
    main_title = title if title != "Документация проекта" else f"Проект: {root.name}"
    add_textbox(slide, Inches(1), Inches(2.5), Inches(11), Inches(1.5), 
                main_title, font_size=44, bold=True, color=theme["title"], align=PP_ALIGN.CENTER)
    
    if root.docstring:
        add_textbox(slide, Inches(1.5), Inches(4.2), Inches(10.3), Inches(1), 
                    root.docstring.strip(), font_size=20, color=theme["muted"], align=PP_ALIGN.CENTER)
    
    if author:
        add_textbox(slide, Inches(1), Inches(5.8), Inches(11), Inches(0.6), 
                    f"Автор: {author}", font_size=16, color=theme["muted"], align=PP_ALIGN.CENTER)

    # 2. Оглавление (Структура)
    if elements:
        slide = prs.slides.add_slide(blank_layout)
        set_slide_bg(slide, theme["bg"])
        add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(1), 
                    "📋 Структура проекта", font_size=36, bold=True, color=theme["accent"])
        
        tf_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11), Inches(5))
        tf = tf_box.text_frame
        tf.word_wrap = True
        
        for i, el in enumerate(elements):
            prefix = "🧩 Класс:" if el.kind == "class" else "⚙️ Функция:"
            line = f"{prefix} {el.name}"
            
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(18)
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(10)

    # 3. Слайды для каждого элемента (ОБЫЧНЫЙ ТЕКСТ, БЕЗ КОДА)
    for el in elements:
        slide = prs.slides.add_slide(blank_layout)
        set_slide_bg(slide, theme["bg"])

        kind_label = {"class": " Класс", "function": "⚙️ Функция", "method": "🔧 Метод"}.get(el.kind, el.kind)
        header = f"{kind_label}: {el.name}"
        
        # Добавляем аргументы в заголовок, если есть
        if el.args and el.kind != "class":
            header += f"({el.args})"
            
        add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(1), 
                    header, font_size=32, bold=True, color=theme["accent"])

        y_pos = Inches(1.8)
        
        # Декораторы как текст
        if el.decorators:
            dec_text = "Декораторы: @" + ", @".join(el.decorators)
            add_textbox(slide, Inches(1), y_pos, Inches(11), Inches(0.5), 
                        dec_text, font_size=16, color=theme["muted"])
            y_pos = Inches(2.4)

        # Docstring как основное содержимое слайда
        content_text = el.docstring.strip() if el.docstring else "Нет описания."
        
        # Разбиваем длинный docstring на абзацы
        paragraphs = content_text.split('\n\n')
        for para in paragraphs:
            add_textbox(slide, Inches(1), y_pos, Inches(11), Inches(4), 
                        para.strip(), font_size=20, color=theme["text"])
            y_pos += Inches(len(para) / 30 + 0.5)  # Примерная высота

    # 4. Финальный слайд
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide, theme["bg"])
    add_textbox(slide, Inches(1), Inches(3), Inches(11), Inches(1.5), 
                "✨ Спасибо за внимание!", font_size=40, bold=True, color=theme["title"], align=PP_ALIGN.CENTER)

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer
