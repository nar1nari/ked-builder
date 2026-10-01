import re
from dataclasses import dataclass
from os import PathLike
from typing import BinaryIO

import openpyxl

# Сведения об ОП лежат в шапке учебного плана (до таблицы с дисциплинами),
# поэтому достаточно просмотреть первые строки листа.
_HEADER_ROWS = 15

# "Білім беру бағдарламасы: 6В06150 Ақпараттық технологиялар ... /
#  Образовательная программа: 6В06150 Информационные технологии ..."
_PROGRAM_PATTERN = re.compile(
    r"Білім\s+беру\s+бағдарламасы\s*:?\s*(?P<code_kk>\S+)\s+(?P<name_kk>.+?)\s*/\s*"
    r"Образовательная\s+программа\s*:?\s*(?P<code_ru>\S+)\s+(?P<name_ru>.+)",
    re.IGNORECASE | re.DOTALL,
)

# "... на 2025-2029 ..." - берём первый диапазон годов в шапке
_YEARS_PATTERN = re.compile(r"(?<!\d)(\d{4})\s*[-–—]\s*(\d{4})(?!\d)")

ExcelSource = str | PathLike | BinaryIO


@dataclass
class ProgramInfo:
    code: str  # код ОП, например "6В06150"
    name: str  # название ОП на русском
    name_kk: str  # название ОП на казахском
    year: int  # год начала обучения (год набора)


def _read_header_texts(xlsx: ExcelSource) -> list[str]:
    """Возвращает тексты всех непустых строковых ячеек шапки УП по порядку."""
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    try:
        ws = wb.active
        texts: list[str] = []
        for row in ws.iter_rows(min_row=1, max_row=_HEADER_ROWS):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    texts.append(re.sub(r"\s+", " ", cell.value).strip())
        return texts
    finally:
        wb.close()


def _parse_program(texts: list[str]) -> re.Match:
    for text in texts:
        match = _PROGRAM_PATTERN.search(text)
        if match:
            return match
    raise ValueError(
        "В шапке учебного плана не найдена строка вида "
        '"Білім беру бағдарламасы: <код> <название> / Образовательная программа: <код> <название>"'
    )


def _parse_year(texts: list[str]) -> int:
    for text in texts:
        match = _YEARS_PATTERN.search(text)
        if match:
            return int(match.group(1))
    raise ValueError(
        'В шапке учебного плана не найден период обучения вида "2025-2029"'
    )


def get_program_info(xlsx: ExcelSource) -> ProgramInfo:
    """
    Читает шапку учебного плана и возвращает код, название (рус./каз.) и
    год начала обучения. Принимает путь к файлу или файловый объект
    (например, загруженный пользователем файл или io.BytesIO).

    Если нужных данных в шапке нет (файл не похож на УП), бросает ValueError
    с понятным сообщением.
    """
    texts = _read_header_texts(xlsx)
    match = _parse_program(texts)

    return ProgramInfo(
        code=match.group("code_ru").strip(),
        name=match.group("name_ru").strip(),
        name_kk=match.group("name_kk").strip(),
        year=_parse_year(texts),
    )


def get_program_code(xlsx: ExcelSource) -> str:
    """Код образовательной программы, например "6В06150"."""
    return get_program_info(xlsx).code


def get_program_name(xlsx: ExcelSource) -> str:
    """Наименование образовательной программы на русском."""
    return get_program_info(xlsx).name


def get_program_name_kk(xlsx: ExcelSource) -> str:
    """Наименование образовательной программы на казахском."""
    return get_program_info(xlsx).name_kk


def get_program_year(xlsx: ExcelSource) -> int:
    """Год начала обучения (год набора), например 2025."""
    return get_program_info(xlsx).year
