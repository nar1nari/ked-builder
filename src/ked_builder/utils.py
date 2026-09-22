from docx import Document


def find_table_after_heading(doc: Document, heading_text: str):  # ty: ignore[invalid-type-form]
    """Возвращает первую таблицу, идущую сразу после абзаца с заданным заголовком."""
    heading_found = False
    for element in doc.element.body:
        tag = element.tag.split("}")[-1]  # 'p' или 'tbl'

        if tag == "p":
            text = element.text or ""
            if heading_text in text:
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
