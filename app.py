from flask import Flask, render_template, request, send_file
from pptx import Presentation
from pptx.util import Pt
import io

app = Flask(__name__)

def generate_pptx(text_content):
    """Функция создания презентации из текста"""
    prs = Presentation()
    
    # Разбиваем текст на слайды. 
    # Логика: Два переноса строки (\n\n) разделяют слайды.
    # Первая строка блока - заголовок, остальные - пункты списка.
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
        tf.clear() # Очистить дефолтный текст
        
        # Если есть только заголовок и нет пунктов, пропускаем добавление пунктов
        if len(lines) > 1:
            for point in lines[1:]:
                p = tf.add_paragraph()
                p.text = point
                p.level = 0
                p.font.size = Pt(24)

    # Сохраняем в байтовый поток (в память), а не на диск
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        text_data = request.form.get('presentation_text')
        
        if text_data:
            # Генерируем файл
            pptx_file = generate_pptx(text_data)
            
            # Отправляем файл пользователю
            return send_file(
                pptx_file,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name='my_presentation.pptx'
            )
            
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
