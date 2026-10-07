from io import BytesIO
from typing import List
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pygments import lex
from pygments.lexers import PythonLexer
from pygments.token import Token
from code_parser import CodeElement

THEMES = {
    "dark": {
        "bg": RGBColor(0x27, 0x28, 0x22),
        "title": RGBColor(0xF8, 0xF8, 0xF2),
        "text": RGBColor(0xF8, 0xF8, 0xF2),
        "accent": RGBColor(0xA6, 0xE2, 0x2E),
        "muted": RGBColor(0x75, 0x71, 0x5E),
        "code_bg": RGBColor(0x1E, 0x1F, 0x1C),
        "token_colors": {
            Token.Keyword: RGBColor(0xF9, 0x26, 0x72),
            Token.Name.Builtin: RGBColor(0x66, 0xD9, 0xEF),
            Token.Name.Function: RGBColor(0xA6, 0xE2, 0x2E),
            Token.Name.Class: RGBColor(0x66, 0xD9, 0xEF),
            Token.Name.Decorator: RGBColor(0xA6, 0xE2, 0x2E),
            Token.String: RGBColor(0xE6, 0xDB, 0x74),
            Token.Comment: RGBColor(0x75, 0x71, 0x5E),
            Token.Number: RGBColor(0xAE, 0x81, 0xFF),
            Token.Operator: RGBColor(0xF9, 0x26, 0x72),
            Token.Punctuation: RGBColor(0xF8, 0xF8, 0xF2),
        },
    },
    "light": {
        "bg": RGBColor(0xFF, 0xFF, 0xFF),
        "title": RGBColor(0x2D, 0x2D, 0x2D),
        "text": RGBColor(0x33, 0x33, 0x33),
        "accent": RGBColor(0x00, 0x7A, 0xCC),
        "muted": RGBColor(0x88, 0x88, 0x88),
        "code_bg": RGBColor(0xF5, 0xF5, 0xF5),
        "token_colors": {
            Token.Keyword: RGBColor(0xD7, 0x3A, 0x49),
            Token.Name.Builtin: RGBColor(0x00, 0x5C, 0xC1),
            Token.Name.Function: RGBColor(0x6F, 0x42, 0xC1),
            Token.Name.Class: RGBColor(0x00, 0x5C, 0xC1),
            Token.Name.Decorator: RGBColor(0xE3, 0x62, 0x09),
            Token.String: RGBColor(0x03, 0x2F, 0x62),
            Token.Comment: RGBColor(0x6A, 0x73, 0x7D),
            Token.Number: RGBColor(0x00, 0x5C, 0xC1),
            Token.Operator: RGBColor(0xD7, 0x3A, 0x49),
            Token.Punctuation: RGBColor(0x24, 0x29, 0x2E),
        },
    },
}


def _set_slide_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_textbox(slide, left, top, width, height, text,
                 font_size=18, bold=False, color=None, align=PP_ALIGN.LEFT):
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


def _colorize_code(tf, code, theme, font_size=Pt(10)):
    tf.word_wrap = True
    default_color = theme["text"]
    lexer = PythonLexer()
    first = True
    for ttype, value in lex(code, lexer):
        color = default_color
        tt = ttype
        while tt and tt not in theme["token_colors"]:
            tt = tt.parent
        if tt:
            color = theme["token_colors"][tt]
        for line_idx, line in enumerate(value.split("\n")):
            if not first and line_idx == 0 and "\n" in value:
                p = tf.add_paragraph()
                p.font.size = font_size
            if first:
                p = tf.paragraphs[0]
                p.font.size = font_size
                first = False
            if line:
                run = p.add_run()
                run.text = line
                run.font.size = font_size
                run.font.name = "Consolas"
                run.font.color.rgb = color


def _add_code_block(slide, left, top, width, height, code, theme):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = theme["code_bg"]
    shape.line.fill.background()
    tf = shape.text_frame
    tf.margin_left = Pt(10)
    tf.margin_right = Pt(10)
    tf.margin_top = Pt(8)
    tf.margin_bottom = Pt(8)
    tf.word_wrap = True
    _colorize_code(tf, code, theme, font_size=Pt(11))
    return shape


def build_presentation(root, elements, title="Python Code Overview",
                       author="", theme_name="dark"):
    theme = THEMES[theme_name]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(blank)
    _set_slide_bg(slide, theme["bg"])
    _add_textbox(slide, Inches(1), Inches(2.2), Inches(11), Inches(1.2),
                 title, font_size=44, bold=True, color=theme["title"],
                 align=PP_ALIGN.CENTER)
    subtitle = f"{root.name}.py" if root.name != "module" else ""
    if subtitle:
        _add_textbox(slide, Inches(1), Inches(3.5), Inches(11), Inches(0.8),
                     subtitle, font_size=24, color=theme["accent"],
                     align=PP_ALIGN.CENTER)
    if author:
        _add_textbox(slide, Inches(1), Inches(5.5), Inches(11), Inches(0.6),
                     f"Автор: {author}", font_size=18, color=theme["muted"],
                     align=PP_ALIGN.CENTER)

    if elements:
        slide = prs.slides.add_slide(blank)
        _set_slide_bg(slide, theme["bg"])
        _add_textbox(slide, Inches(0.7), Inches(0.4), Inches(11), Inches(0.8),
                     "Структура проекта", font_size=32, bold=True,
                     color=theme["title"])
        tf_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.4),
                                          Inches(11), Inches(5.5))
        tf = tf_box.text_frame
        tf.word_wrap = True
        for i, el in enumerate(elements):
            prefix = "class: " if el.kind == "class" else "func: "
            line = f"{prefix}{el.name}"
            if el.args and el.kind != "class":
                line += f"({el.args})"
            if el.docstring:
                short = el.docstring.splitlines()[0][:80]
                line += f"  -  {short}"
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(16)
            p.font.name = "Consolas"
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(6)

    for el in elements:
        slide = prs.slides.add_slide(blank)
        _set_slide_bg(slide, theme["bg"])
        kind_label = {"class": "Класс", "function": "Функция",
                      "method": "Метод"}.get(el.kind, el.kind)
        header = f"{kind_label}: {el.name}"
        if el.args and el.kind != "class":
            header += f"({el.args})"
        _add_textbox(slide, Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                     header, font_size=26, bold=True, color=theme["title"])
        y = Inches(1.0)
        if el.decorators:
            _add_textbox(slide, Inches(0.6), y, Inches(12), Inches(0.4),
                         " @" + " @".join(el.decorators),
                         font_size=14, color=theme["accent"])
            y = Inches(1.4)
        if el.docstring:
            _add_textbox(slide, Inches(0.6), y, Inches(12), Inches(1.2),
                         el.docstring.strip(), font_size=14,
                         color=theme["muted"])
            y = y + Inches(1.3)
        code_height = Inches(7.5) - y - Inches(0.3)
        if code_height < Inches(1):
            code_height = Inches(1)
        _add_code_block(slide, Inches(0.5), y, Inches(12.3), code_height,
                        el.source, theme)

    slide = prs.slides.add_slide(blank)
    _set_slide_bg(slide, theme["bg"])
    _add_textbox(slide, Inches(1), Inches(3), Inches(11), Inches(1),
                 "Спасибо за внимание!", font_size=40, bold=True,
                 color=theme["title"], align=PP_ALIGN.CENTER)

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

