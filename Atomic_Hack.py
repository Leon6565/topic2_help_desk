from parser_pdf_to_exel import clean_text, extract_text_by_chapters, extract_pages_to_excel
import os
import pandas as pd
import re
def sanitize_filename(filename):
    return "".join([c if c.isalnum() else "_" for c in filename])

def process_pdfs_to_excel(input_folder, output_folder):
    # Создаем выходную папку, если она не существует
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Проходим по всем файлам в указанной папке
    for filename in os.listdir(input_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(input_folder, filename)
            sanitized_filename = sanitize_filename(os.path.splitext(filename)[0])
            output_file = os.path.join(output_folder, f"{sanitized_filename}.xlsx")
            print(filename)
            try:
                # Используем вашу функцию для извлечения текста
                text_data = extract_text_by_chapters(pdf_path)
                
                # Создаем Excel файл используя вашу функцию
                extract_pages_to_excel(text_data, output_file, input_filename = filename)
            except Exception as e:
                print(f"An error occurred while processing {filename}: {e}")

# Пример использования функции
input_folder = r'C:\Users\user\Desktop\Data_Science_study\_projects\Система технической поддержки\dataset\Для Хакатона'
output_folder = r'C:\Users\user\Desktop\Data_Science_study\_projects\Система технической поддержки\dataset\excel'
combined_output_file = r'C:\Users\user\Desktop\Data_Science_study\_projects\Система технической поддержки\dataset\excel\combined\combined_output_file.xlsx'
#process_pdfs_to_excel(input_folder, output_folder)

def combine_excel_files(output_folder, combined_output_file):
    all_data = []
    for filename in os.listdir(output_folder):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(output_folder, filename)
            df = pd.read_excel(file_path)
            all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)
    combined_df.to_excel(combined_output_file, index=False)
    
#combine_excel_files(output_folder, combined_output_file)

import pandas as pd
import string

# Загрузка Excel файла
file_path = r'C:\Users\user\Desktop\Data_Science_study\_projects\Система технической поддержки\dataset\excel\combined\combined_output_file.xlsx'
df = pd.read_excel(file_path)

# Функция для удаления знаков препинания, тире, скобок и кавычек
def remove_special_characters(text):
    return re.sub(r'[^A-Za-zА-Яа-я0-9\s]', '', text)

# Применение функции к третьему столбцу
df.iloc[:, 2] = df.iloc[:, 2].apply(remove_special_characters)

# Сохранение обработанного файла
output_path = r'C:\Users\user\Desktop\Data_Science_study\_projects\Система технической поддержки\dataset\excel\combined\combined_output_file_clear.xlsx'
df.to_excel(output_path, index=False)

print("Знаки препинания удалены, файл сохранен как:", output_path)
