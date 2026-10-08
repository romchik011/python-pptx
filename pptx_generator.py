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

def add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, color=None, align=PP_ALIGN.LEFT):
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

def split_text_into_chunks(text, max_chars=400):
    """Разбивает длинный текст на части по max_chars символов."""
    if len(text) <= max_chars:
        return [text]
    
    chunks = []
    paragraphs = text.split('\n\n')
    current_chunk = ""
    
    for para in paragraphs:
        if len(current_chunk) + len(para) + 2 <= max_chars:
            current_chunk += (para + "\n\n") if current_chunk else para
        else:
            if current_chunk:
                chunks.append(current_chunk.strip())
            current_chunk = para
    
    if current_chunk:
        chunks.append(current_chunk.strip())
    
    return chunks

def build_presentation_from_code(root, elements, title="Документация проекта", 
                                  author="", theme_name="light", max_chars_per_slide=400):
    """
    Создает презентацию с неограниченным количеством слайдов.
    Каждый элемент кода получает свой слайд.
    Длинные описания разбиваются на несколько слайдов.
    """
    theme = THEMES[theme_name]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    slide_count = 0

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
    
    slide_count += 1

    # 2. Оглавление (Структура проекта)
    if elements:
        slide = prs.slides.add_slide(blank_layout)
        set_slide_bg(slide, theme["bg"])
        add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(1), 
                    "📋 Структура проекта", font_size=36, bold=True, color=theme["accent"])
        
        tf_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11), Inches(5))
        tf = tf_box.text_frame
        tf.word_wrap = True
        
        for i, el in enumerate(elements):
            prefix = "🧩 Класс: " if el.kind == "class" else "️ Функция: "
            line = f"{prefix}{el.name}"
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(18)
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(10)
        
        slide_count += 1

    # 3. Слайды для каждого элемента (НЕОГРАНИЧЕННОЕ КОЛИЧЕСТВО!)
    for el in elements:
        kind_label = {"class": " Класс", "function": "⚙️ Функция", "method": "🔧 Метод"}.get(el.kind, "Элемент")
        header = f"{kind_label}: {el.name}"
        if el.args and el.kind != "class":
            header += f"({el.args})"
        
        # Получаем текст описания
        content_text = el.docstring.strip() if el.docstring else "Описание отсутствует. Добавьте строку документации (docstring) в код."
        
        # Разбиваем длинный текст на части
        text_chunks = split_text_into_chunks(content_text, max_chars_per_slide)
        
        # Создаем слайд для каждой части текста
        for chunk_idx, chunk in enumerate(text_chunks):
            slide = prs.slides.add_slide(blank_layout)
            set_slide_bg(slide, theme["bg"])
            
            # Заголовок слайда
            title_text = header if len(text_chunks) == 1 else f"{header} (часть {chunk_idx + 1}/{len(text_chunks)})"
            add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(1), 
                        title_text, font_size=32, bold=True, color=theme["accent"])
            
            # Декораторы (только на первом слайде элемента)
            y_pos = Inches(1.8)
            if el.decorators and chunk_idx == 0:
                add_textbox(slide, Inches(1), y_pos, Inches(11), Inches(0.5), 
                            f"Декораторы: @{', @'.join(el.decorators)}", font_size=16, color=theme["muted"])
                y_pos = Inches(2.4)
            
            # Содержимое
            add_textbox(slide, Inches(1), y_pos, Inches(11), Inches(4.5), 
                        chunk, font_size=20, color=theme["text"])
            
            slide_count += 1

    # 4. Финальный слайд
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide, theme["bg"])
    add_textbox(slide, Inches(1), Inches(3), Inches(11), Inches(1.5), 
                "✨ Спасибо за внимание!", font_size=40, bold=True, color=theme["title"], align=PP_ALIGN.CENTER)
    
    slide_count += 1
    
    print(f"✅ Создано слайдов: {slide_count}")

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer
