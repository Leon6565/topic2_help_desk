import pdfplumber
import re
import io
import os
import pandas as pd

def clean_text(text):
    # Регулярные выражения для удаления указанных надписей
    patterns = [
        r'АО «Гринатом»',  # АО «Гринатом»
        r'Автор',  # Автор
        r'Инструкция пользователей «[\s\S]*?»',  # Инструкция пользователей «<текст>»
        r'Страниц\s+\d+',  # Страниц <число>
        r'Версия\s+\d+',  # Версия <число>
        r'Рисунок\s+\d+(:|\.|)'  # Рисунок <число>, Рисунок <число>:, Рисунок <число>.
    ]
    
    for pattern in patterns:
        text = re.sub(pattern, '', text, flags=re.DOTALL)
    
    return text.strip()

def extract_text_by_chapters(pdf_path):
    chapters = {}
    current_chapter = None
    current_subchapter = None

    # Регулярные выражения для глав и подглав
    chapter_pattern = re.compile(r'^\s*(\d+\.\d+)\s+(.*)$')
    subchapter_pattern = re.compile(r'^\s*(\d+\.\d+\.\d+)\s+(.*)$')

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line in lines:
                    # Очистка линии от ненужных надписей
                    cleaned_line = clean_text(line)
                    
                    chapter_match = chapter_pattern.match(cleaned_line)
                    subchapter_match = subchapter_pattern.match(cleaned_line)

                    if subchapter_match:
                        current_subchapter = subchapter_match.group(1) + " " + subchapter_match.group(2)
                        chapters[current_subchapter] = ""
                    elif chapter_match:
                        current_chapter = chapter_match.group(1) + " " + chapter_match.group(2)
                        chapters[current_chapter] = ""
                    elif current_subchapter:
                        chapters[current_subchapter] += cleaned_line + "\n"
                    elif current_chapter:
                        chapters[current_chapter] += cleaned_line + "\n"
    
    # Удаление пустых строк и лишних пробелов
    for key in chapters:
        chapters[key] = chapters[key].strip()
    # Создаем строковый буфер
    output_buffer = io.StringIO()

    # Записываем данные в буфер вместо вывода на экран
    for key, value in chapters.items():
        output_buffer.write(f"{key}:\n{value}\n{'-'*80}\n")

    # Получаем строку из буфера
    output_string = output_buffer.getvalue()

    # Закрываем буфер
    output_buffer.close()
    return output_string

# Пример использования функции
#pdf_path = r'C:\Users\user\Desktop\Data_Science_study\_projects\Система технической поддержки\dataset\Инструкция_Учет_внутрихозяйственных_расчетов.pdf'
#print(extract_text_by_chapters(pdf_path))


def extract_pages_to_excel(text, output_file, output_path='.', input_filename='input.txt'):
    try:
        # Проверяем, существует ли директория и есть ли права на запись
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        if not os.access(output_path, os.W_OK):
            raise PermissionError(f"No write permission for the directory: {output_path}")

        # Регулярное выражение для поиска страниц
        pattern = re.compile(r'Страница (\d+)')
        
        # Разбиваем текст на части, используя регулярное выражение
        parts = pattern.split(text)
        
        # Инициализируем список для хранения страниц
        data = []
        
        # Перебираем части и извлекаем номер страницы и контент
        for i in range(1, len(parts), 2):
            page_number = int(parts[i])
            content = parts[i + 1].strip().replace('\n', ' ')
            data.append({
                'Filename': input_filename,
                'Page Number': page_number,
                'Content': content
            })
        
        # Создаем DataFrame
        df = pd.DataFrame(data)

        # Создаем полный путь к выходному файлу
        full_output_path = os.path.join(output_path, output_file)
        
        # Записываем DataFrame в Excel-файл
        df.to_excel(full_output_path, index=False)

        print(f"File successfully saved to {full_output_path}")

    except PermissionError as e:
        print(e)
    except Exception as e:
        print(f"An error occurred: {e}")     
