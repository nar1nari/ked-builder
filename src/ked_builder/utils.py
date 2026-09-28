import re

from docx import Document


def normalize_text(text: str) -> str:
    """
    Приводит текст к виду, удобному для нечёткого сравнения заголовков:
    нижний регистр, "ё" -> "е", любые знаки препинания/спецсимволы/
    неразрывные пробелы превращаются в пробел, а повторяющиеся пробелы
    схлопываются. Так "Сведения  о дисциплинах, образовательной программе."
    и "Сведения о дисциплинах образовательной программе" считаются одинаковыми.
    """
    text = text.lower().replace("ё", "е")
    text = re.sub(r"[\W_]+", " ", text)
    return text.strip()


def text_contains(text: str, needle: str) -> bool:
    """Нечёткая (по нормализованному тексту) проверка вхождения needle в text."""
    normalized_needle = normalize_text(needle)
    return bool(normalized_needle) and normalized_needle in normalize_text(text)


def find_paragraph(doc: Document, needle: str):  # ty: ignore[invalid-type-form]
    """Возвращает первый абзац, содержащий needle (см. text_contains), или None."""
    for paragraph in doc.paragraphs:
        if text_contains(paragraph.text, needle):
            return paragraph
    return None


def find_table_after_heading(doc: Document, heading_text: str):  # ty: ignore[invalid-type-form]
    """
    Возвращает первую таблицу, идущую сразу после абзаца с заданным заголовком.
    Заголовок ищется не по точному совпадению, а через text_contains -
    лишние пробелы, запятые, точки и регистр значения не имеют.
    """
    heading_found = False
    for element in doc.element.body:
        tag = element.tag.split("}")[-1]  # 'p' или 'tbl'

        if tag == "p":
            text = element.text or ""
            if text_contains(text, heading_text):
                heading_found = True

        elif tag == "tbl" and heading_found:
            for table in doc.tables:
                if table._element is element:
                    return table

    return None


def _replace_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    """
    Заменяет плейсхолдеры внутри абзаца run-по-run. В реальных документах
    Word текст абзаца обычно разбит на несколько run'ов (из-за форматирования
    отдельных слов), но сами плейсхолдеры вида "[TAG]" в проверенных местах
    этого шаблона всегда целиком лежат в одном run'е, поэтому простой замены
    текста внутри каждого run'а достаточно.
    """
    for run in paragraph.runs:
        for tag, value in replacements.items():
            if tag in run.text:
                run.text = run.text.replace(tag, value)


def replace_placeholders(doc: Document, replacements: dict[str, str]) -> None:  # ty: ignore[invalid-type-form]
    """
    Заменяет плейсхолдеры вида "[TAG]" на значения из replacements по всему
    документу - в обычных абзацах и во всех ячейках всех таблиц.
    """
    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_in_paragraph(paragraph, replacements)
