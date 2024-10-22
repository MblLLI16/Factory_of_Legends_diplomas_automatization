import os
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from io import BytesIO
import pandas as pd

# Функция для создания временного PDF с ФИО участника в памяти
def create_diploma_template_in_memory(participant_name):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)

    # Подключение шрифта Arial Unicode MS
    pdfmetrics.registerFont(TTFont('ArialUnicode', 'ARIALUNI.TTF'))
    
    # Устанавливаем шрифт Arial Unicode и размер 21
    c.setFont("ArialUnicode", 21)

    # Настройка расположения текста. Меняй координаты (185, 540) по необходимости
    c.drawString(185, 540, f"{participant_name}")  # Вставка ФИО участника
    
    c.showPage()
    c.save()

    packet.seek(0)
    return packet

# Функция для объединения PDF-шаблона и временного PDF с именем
def merge_pdfs(template_pdf, overlay_stream, output_pdf):
    with open(template_pdf, "rb") as template_file:
        template_reader = PyPDF2.PdfReader(template_file)
        overlay_reader = PyPDF2.PdfReader(overlay_stream)

        template_page = template_reader.pages[0]
        overlay_page = overlay_reader.pages[0]

        # Наложение текста на шаблон
        template_page.merge_page(overlay_page)

        # Сохраняем итоговый PDF
        writer = PyPDF2.PdfWriter()
        writer.add_page(template_page)
        with open(output_pdf, "wb") as output_file:
            writer.write(output_file)

# Функция для чтения данных из диапазона Excel
def read_participants_from_excel(excel_file):
    # Считываем данные из столбца C (3-й столбец), начиная с 6-й строки
    df = pd.read_excel(excel_file, usecols="C", skiprows=4, nrows=70)  # Чтение ячеек C6:C75
    participants = df.iloc[:, 0].tolist()  # Преобразуем столбец в список
    return participants

if __name__ == "__main__":
    # Имя файла твоего шаблона PDF
    template_pdf = "template.pdf"  # Укажи правильное имя файла шаблона
    
    # Чтение списка участников из Excel файла
    excel_file = "./input/Tulun_Registry_2024.xlsx"  # Укажи путь к твоему файлу Excel
    
    # Получаем имя файла Excel без расширения для создания папки
    excel_filename = os.path.splitext(os.path.basename(excel_file))[0]

    # Путь для сохранения дипломов
    output_dir = os.path.join("./output", excel_filename)
    
    # Создаем папку, если она не существует
    os.makedirs(output_dir, exist_ok=True)

    # Чтение участников из Excel файла
    participants = read_participants_from_excel(excel_file)

    # Цикл для обработки всех участников
    for participant in participants:
        # Создаем PDF с ФИО участника в памяти
        overlay_stream = create_diploma_template_in_memory(participant)

        # Имя для выходного файла с путём в нужной папке
        output_pdf = os.path.join(output_dir, f"{participant}.pdf")

        # Объединяем шаблон и временный PDF с ФИО
        merge_pdfs(template_pdf, overlay_stream, output_pdf)

        print(f"Диплом создан для {participant}")
