import os
import docx
from docx.shared import Pt, Inches
import re


def load_ignore_patterns(ignore_file_path='.docignore'):
    # Загружает паттерны игнорирования из файла
    ignore_patterns = []
    if not os.path.exists(ignore_file_path):
        raise FileNotFoundError(
            f"Файл {ignore_file_path} не найден. Создайте файл .docignore для настройки игнорирования")

    try:
        with open(ignore_file_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):  # Игнорируем пустые строки и комментарии
                    ignore_patterns.append(line)
        print(f"Загружено {len(ignore_patterns)} правил игнорирования из {ignore_file_path}")
    except Exception as e:
        print(f"Ошибка чтения файла {ignore_file_path}: {e}")

    return ignore_patterns


def should_ignore(path, ignore_patterns, is_directory=False):
    # Проверяет, нужно ли игнорировать файл или папку
    if not ignore_patterns:
        return False

    # Нормализуем путь для сравнения
    normalized_path = os.path.normpath(path)

    for pattern in ignore_patterns:
        # Если паттерн заканчивается на /, это относится только к директориям
        if pattern.endswith('/'):
            if not is_directory:
                continue
            pattern = pattern[:-1]  # Убираем trailing slash

        # Простое сравнение имен
        if pattern in normalized_path:
            return True

        # Поддержка wildcard *
        if '*' in pattern:
            import fnmatch
            if fnmatch.fnmatch(normalized_path, pattern):
                return True

    return False


def get_file_content(file_path):
    # Получает содержимое файла
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
    # Удаляет из текста символы, несовместимые с XML
    if text is None:
        return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)


def add_file_to_column(cell, file_name, content, file_number):
    # Добавляет содержимое файла в ячейку таблицы
    # Добавляем заголовок файла как Heading 1
    p = cell.add_paragraph()
    p.style = 'Heading 1'
    run = p.add_run(f"ПРИЛОЖЕНИЕ {file_number}. {file_name}")
    run.font.size = Pt(14)  # Уменьшенный шрифт для кода
    run.bold = True
    run.font.name = 'Times New Roman'

    # Добавляем содержимое с уменьшенным шрифтом
    p = cell.add_paragraph()
    run = p.add_run(content)
    run.font.size = Pt(6)  # Уменьшенный шрифт для кода

    # Добавляем разделитель
    # cell.add_paragraph()


def add_file_info_to_doc(directory, doc, use_two_columns=False, ignore_patterns=None):
    # Рекурсивно проходит по директориям, получает имена и содержимое файлов и добавляет их в документ
    if ignore_patterns is None:
        ignore_patterns = []

    all_files = []

    # Сначала собираем все файлы с учетом игнорирования
    for root, dirs, files in os.walk(directory):
        # Фильтруем директории для игнорирования
        dirs[:] = [d for d in dirs if not should_ignore(os.path.join(root, d), ignore_patterns, True)]

        for file in files:
            file_path = os.path.join(root, file)

            # Проверяем, нужно ли игнорировать файл
            if should_ignore(file_path, ignore_patterns, False):
                print(f"Игнорируем: {file_path}")
                continue

            file_content = get_file_content(file_path)
            sanitized_content = sanitize_text(file_content)
            all_files.append((file, sanitized_content))

    if not all_files:
        doc.add_paragraph("Не найдено файлов для обработки (возможно, все игнорируются)")
        return

    if use_two_columns:
        # Создаем таблицу с двумя колонками
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Inches(3)
        table.columns[1].width = Inches(3)

        current_col = 0  # 0 - левая колонка, 1 - правая колонка
        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        # Очищаем начальные ячейки
        left_cell.paragraphs[0].clear()
        right_cell.paragraphs[0].clear()

        file_number = 1
        for file_name, content in all_files:
            if current_col == 0:
                # Добавляем в левую колонку
                add_file_to_column(left_cell, file_name, content, file_number)
                current_col = 1
            else:
                # Добавляем в правую колонку
                add_file_to_column(right_cell, file_name, content, file_number)
                current_col = 0
            file_number += 1

        # Если осталась нечетное количество файлов, добавляем пустую строку для симметрии
        if len(all_files) % 2 != 0:
            add_file_to_column(right_cell, "", "", file_number)

    else:
        # Обычный режим - одна колонка
        file_number = 1
        for file_name, content in all_files:
            # Добавляем заголовок файла как Heading 1
            p = doc.add_paragraph()
            p.style = 'Heading 1'
            run = p.add_run(f"ПРИЛОЖЕНИЕ {file_number}. {file_name}")
            run.bold = True

            # Добавляем содержимое с уменьшенным шрифтом
            p = doc.add_paragraph()
            run = p.add_run(content)
            run.font.size = Pt(8)  # Уменьшенный шрифт для кода

            doc.add_paragraph()  # добавляем пустую строку для разделения файлов
            file_number += 1


def main(target_directory, output_file, use_two_columns=False):
    # Главная функция, создает документ, добавляет в него информацию о файлах и сохраняет его
    # Загружаем правила игнорирования
    try:
        ignore_patterns = load_ignore_patterns()
    except FileNotFoundError as e:
        print(f"Внимание: {e}")
        print("Продолжаем без игнорирования файлов")
        ignore_patterns = []

    doc = docx.Document()

    # Добавляем заголовок с информацией о настройках
    title = doc.add_paragraph()
    title_run = title.add_run("Листинг проекта")
    title_run.bold = True
    title_run.font.size = Pt(16)

    info = doc.add_paragraph()
    info.add_run(f"Директория: {target_directory}\n")
    info.add_run(f"Колонки: {'2' if use_two_columns else '1'}\n")
    info.add_run(f"Правил игнорирования: {len(ignore_patterns)}\n")
    doc.add_paragraph()

    add_file_info_to_doc(target_directory, doc, use_two_columns, ignore_patterns)
    doc.save(output_file)
    print(f"Информация о файлах сохранена в '{output_file}'")
    if use_two_columns:
        print("Режим: две колонки (интеллектуальное заполнение)")
    else:
        print("Режим: одна колонка")


def ask_yes_no_question(question):
    # Задает вопрос да/нет и возвращает булево значение
    while True:
        answer = input(f"{question} (да/нет): ").lower().strip()
        if answer in ['да', 'д', 'yes', 'y']:
            return True
        elif answer in ['нет', 'н', 'no', 'n']:
            return False
        else:
            print("Пожалуйста, введите 'да' или 'нет'")


if __name__ == "__main__":
    target_directory = input("Введите путь к директории: ")  # Запрашиваем путь к директории у пользователя
    output_file = input("Введите имя выходного файла Word (.docx): ")  # Запрашиваем имя выходного файла

    # Спрашиваем о двухколоночном режиме
    use_two_columns = ask_yes_no_question("Использовать двухколоночный режим для экономии места?")

    main(target_directory, output_file, use_two_columns)