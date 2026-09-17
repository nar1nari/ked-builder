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
