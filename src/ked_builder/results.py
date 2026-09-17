from docx import Document

from ked_builder.utils import find_table_after_heading

HEADING_RU = "Формируемые результаты обучения"
HEADING_KK = "Білім беру бағдарламасын оқытудың қалыптастырылатын нәтижелері"


def get_results(docx_path: str, language: str = "ru") -> dict[str, str]:
    results = {}
    doc = Document(docx_path)

    table = find_table_after_heading(
        doc, HEADING_KK if language == "kz" else HEADING_RU
    )

    for row in table.rows:
        cells = row.cells
        results[cells[0].text.strip()] = cells[1].text.strip()

    return results
