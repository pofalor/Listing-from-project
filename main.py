import os
import docx
from docx.shared import Pt
import re

def get_file_content(file_path):
    """Получает содержимое файла."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        try:
          with open(file_path, 'r', encoding='latin-1') as f:
            return f.read()
        except Exception as e:
           return f"Ошибка чтения файла: {e}"
    except Exception as e:
        return f"Ошибка чтения файла: {e}"

def sanitize_text(text):
    """Удаляет из текста символы, несовместимые с XML."""
    if text is None:
      return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)

def add_file_info_to_doc(directory, doc):
    """Рекурсивно проходит по директориям, получает имена и содержимое файлов и добавляет их в документ."""
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_content = get_file_content(file_path)
            sanitized_content = sanitize_text(file_content)

            p = doc.add_paragraph()
            run = p.add_run(f"{file}:")
            run.bold = True
            run.font.size = Pt(12)

            p = doc.add_paragraph()
            run = p.add_run(sanitized_content)
            run.font.size = Pt(10)

            doc.add_paragraph() # добавляем пустую строку для разделения файлов

def main(target_directory, output_file):
    """Главная функция, создает документ, добавляет в него информацию о файлах и сохраняет его."""
    doc = docx.Document()
    add_file_info_to_doc(target_directory, doc)
    doc.save(output_file)
    print(f"Информация о файлах сохранена в '{output_file}'")

if __name__ == "__main__":
    target_directory = input("Введите путь к директории: ")  # Запрашиваем путь к директории у пользователя
    output_file = input("Введите имя выходного файла Word (.docx): ") # Запрашиваем имя выходного файла
    main(target_directory, output_file)