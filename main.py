import logging
import os
import re
import sys

import docx
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


LOG_FILE = "listing_from_project.log"
LOGGER = logging.getLogger(__name__)


def configure_logging(log_file=LOG_FILE):
    if LOGGER.handlers:
        return

    LOGGER.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(log_file, encoding='utf-8-sig')
    file_handler.setFormatter(formatter)
    LOGGER.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    LOGGER.addHandler(stream_handler)

    LOGGER.propagate = False


def inches_to_twips(value):
    return int(value * 1440)


def set_section_columns(section, columns=1, spacing_inches=0.25):
    cols = section._sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        section._sectPr.append(cols)

    cols.set(qn('w:num'), str(columns))
    cols.set(qn('w:equalWidth'), '1')
    if columns > 1:
        cols.set(qn('w:space'), str(inches_to_twips(spacing_inches)))


def set_run_font(run, font_name, size=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:ascii'), font_name)
    run._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if size is not None:
        run.font.size = Pt(size)


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
        LOGGER.info("Загружено %s правил игнорирования из %s", len(ignore_patterns), ignore_file_path)
    except Exception as e:
        LOGGER.exception("Ошибка чтения файла %s: %s", ignore_file_path, e)

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
    except UnicodeDecodeError as e:
        LOGGER.warning("Файл %s не прочитан как UTF-8: %s. Пробуем latin-1", file_path, e)
        try:
            with open(file_path, 'r', encoding='latin-1') as f:
                return f.read()
        except Exception as e:
            LOGGER.exception("Ошибка чтения файла %s: %s", file_path, e)
            return f"Ошибка чтения файла {file_path}: {e}"
    except Exception as e:
        LOGGER.exception("Ошибка чтения файла %s: %s", file_path, e)
        return f"Ошибка чтения файла {file_path}: {e}"


def sanitize_text(text):
    # Удаляет из текста символы, несовместимые с XML
    if text is None:
        return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)


def add_file_to_doc(doc, file_name, content, file_number, code_font_size=8, heading_font_size=14):
    # Добавляет содержимое файла в документ.
    # Добавляем заголовок файла как Heading 1
    p = doc.add_paragraph()
    p.style = 'Heading 1'
    p.paragraph_format.keep_with_next = True
    run = p.add_run(f"ПРИЛОЖЕНИЕ {file_number}. {file_name}")
    set_run_font(run, 'Times New Roman', heading_font_size)
    run.bold = True

    # Добавляем содержимое с уменьшенным шрифтом
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    run = p.add_run(content)
    set_run_font(run, 'Courier New', code_font_size)

    # Добавляем разделитель
    doc.add_paragraph()


def log_walk_error(error):
    path = getattr(error, 'filename', '<неизвестно>')
    LOGGER.error("Ошибка обхода директории %s: %s", path, error)


def add_file_info_to_doc(directory, doc, use_two_columns=False, ignore_patterns=None):
    # Рекурсивно проходит по директориям, получает имена и содержимое файлов и добавляет их в документ
    if ignore_patterns is None:
        ignore_patterns = []

    all_files = []

    # Сначала собираем все файлы с учетом игнорирования
    for root, dirs, files in os.walk(directory, onerror=log_walk_error):
        # Фильтруем директории для игнорирования
        try:
            dirs[:] = [d for d in dirs if not should_ignore(os.path.join(root, d), ignore_patterns, True)]
        except Exception as e:
            LOGGER.exception("Ошибка фильтрации директорий в %s: %s", root, e)
            raise

        for file in files:
            file_path = os.path.join(root, file)

            try:
                # Проверяем, нужно ли игнорировать файл
                if should_ignore(file_path, ignore_patterns, False):
                    LOGGER.info("Игнорируем: %s", file_path)
                    continue

                file_content = get_file_content(file_path)
                sanitized_content = sanitize_text(file_content)
                all_files.append((file_path, file, sanitized_content))
            except Exception as e:
                LOGGER.exception("Ошибка обработки файла %s: %s", file_path, e)
                all_files.append((file_path, file, f"Ошибка обработки файла {file_path}: {e}"))

    if not all_files:
        doc.add_paragraph("Не найдено файлов для обработки (возможно, все игнорируются)")
        return

    code_font_size = 6 if use_two_columns else 8
    heading_font_size = 10 if use_two_columns else 14

    for file_number, (file_path, file_name, content) in enumerate(all_files, start=1):
        try:
            add_file_to_doc(doc, file_name, content, file_number, code_font_size, heading_font_size)
        except Exception as e:
            LOGGER.exception("Ошибка добавления файла %s в документ: %s", file_path, e)
            add_file_to_doc(
                doc,
                file_name,
                f"Ошибка добавления файла {file_path} в документ: {e}",
                file_number,
                code_font_size,
                heading_font_size,
            )


def main(target_directory, output_file, use_two_columns=False):
    configure_logging()
    LOGGER.info("Начинаем обработку директории %s", target_directory)

    # Главная функция, создает документ, добавляет в него информацию о файлах и сохраняет его
    # Загружаем правила игнорирования
    try:
        ignore_patterns = load_ignore_patterns()
    except FileNotFoundError as e:
        LOGGER.warning("Внимание: %s", e)
        LOGGER.warning("Продолжаем без игнорирования файлов")
        ignore_patterns = []

    try:
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

        if use_two_columns:
            content_section = doc.add_section(WD_SECTION.CONTINUOUS)
            set_section_columns(content_section, 2)

        add_file_info_to_doc(target_directory, doc, use_two_columns, ignore_patterns)
        doc.save(output_file)
        LOGGER.info("Информация о файлах сохранена в '%s'", output_file)
        if use_two_columns:
            LOGGER.info("Режим: две колонки")
        else:
            LOGGER.info("Режим: одна колонка")
    except Exception as e:
        LOGGER.exception(
            "Обработка прервана. Директория: %s, выходной файл: %s. Причина: %s",
            target_directory,
            output_file,
            e,
        )
        raise


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
