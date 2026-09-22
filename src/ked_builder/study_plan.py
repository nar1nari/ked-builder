import re

import openpyxl
from docx.document import Document as DocumentObject
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from ked_builder.courses import _CODE_ROW_PATTERN, _split_code, _split_trilingual

FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(10)

# ширины столбцов (EMU) - пропорции взяты из таблицы образца "на печать КЭД",
# столбец "Цикл" чуть расширен, чтобы в нём помещалась буква "л"
COLUMN_WIDTHS = [520000, 1477710, 628015, 2535555, 795655]

HEADER_ROW = [
    "Цикл",
    "Модулі/Модуль",
    "Код",
    "Пән атауы /\nНаименование дисциплины",
    "Академиялық кредиттер саны/Количество академических кредитов.",
]

# соответствие последних букв русского сокращения цикла (берётся из самого
# учебного плана, например "ООД ОК", "БД ВК", "ООД КВ") типу компонента -
# без каких-либо сопоставлений конкретных дисциплин, только по этим
# универсальным двухбуквенным кодам
_COMPONENT_BY_CYCLE_SUFFIX = [
    ("ОК", "Обязательный компонент", "Міндетті компонент"),
    ("КВ", "Компонент по выбору", "Таңдау компоненті"),
    ("ВК", "Вузовский компонент", "ЖОО компоненті"),
    ("ИА", "Итоговая аттестация", "Қорытынды аттестация"),
]


def _component_from_cycle(cycle_ru: str) -> tuple[str, str]:
    """Определяет (компонент на русском, компонент на казахском) по сокращению цикла."""
    cycle_ru = cycle_ru.strip()
    for suffix, ru, kk in _COMPONENT_BY_CYCLE_SUFFIX:
        if cycle_ru.endswith(suffix):
            return ru, kk
    return cycle_ru, cycle_ru  # неизвестное сокращение - оставляем как есть


def get_study_plan(xlsx_path: str) -> dict[int, list[dict]]:
    """
    Читает учебный план (.xlsx) и возвращает {курс (1-4): список строк
    таблицы плана обучения для этого курса}, уже в порядке печати.
    Каждая строка - один из трёх видов:
        {"type": "semester", "number": int, "credits": int|None}
        {"type": "component", "label_ru": str, "label_kk": str, "credits": int}
        {"type": "discipline", "cycle_ru", "module_kk", "module_ru",
         "code_kk", "code_ru", "name_kk", "name_ru", "credits"}

    Итог по кредитам внутри компонента считается суммой кредитов всех его
    дисциплин, КРОМЕ "Компонента по выбору" - там несколько альтернативных
    дисциплин имеют одинаковое количество кредитов (в файле отмечено "*" -
    "как у первой дисциплины этой группы"), и в итог идёт не сумма всех
    альтернатив, а кредиты одного выбора.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    # --- проход 1: строки листа -> события ---
    events: list[dict] = []
    last_module_raw = None
    last_credits = None

    for row in ws.iter_rows(min_row=7):
        cycle_raw = row[0].value
        module_raw = row[1].value
        code_raw = row[2].value
        name_raw = row[3].value
        credits_raw = row[4].value

        if module_raw:
            last_module_raw = module_raw

        if code_raw and name_raw and _CODE_ROW_PATTERN.search(str(code_raw)):
            if credits_raw in (None, "*"):
                credits = last_credits
            else:
                credits = credits_raw
                last_credits = credits_raw

            _, cycle_ru, _ = _split_trilingual(cycle_raw or "")
            module_kk, module_ru, _ = _split_trilingual(last_module_raw or "")
            name_kk, name_ru, _ = _split_trilingual(name_raw)
            code_ru, code_kk = _split_code(code_raw)
            # для этой таблицы код даётся без номера - как в образце "на печать"
            code_ru = re.sub(r"\s*\d.*$", "", code_ru).strip()
            code_kk = re.sub(r"\s*\d.*$", "", code_kk).strip()

            events.append(
                {
                    "type": "discipline",
                    "cycle_ru": cycle_ru,
                    "module_kk": module_kk,
                    "module_ru": module_ru,
                    "code_kk": code_kk,
                    "code_ru": code_ru,
                    "name_kk": name_kk,
                    "name_ru": name_ru,
                    "credits": credits,
                }
            )
            continue

        raw_text = str(cycle_raw or "").strip()
        if not raw_text:
            continue

        _, ru_part, _ = _split_trilingual(raw_text)
        ru_part = ru_part.strip()

        year_match = re.match(r"^(\d+)\s*курс", ru_part, re.IGNORECASE)
        if year_match:
            events.append({"type": "year", "year": int(year_match.group(1))})
            continue

        if re.search(r"семестр", ru_part, re.IGNORECASE):
            if re.match(r"^(всего|итого)", ru_part, re.IGNORECASE):
                events.append({"type": "semester_total", "credits": credits_raw})
            else:
                events.append({"type": "semester_start"})
            continue

        # прочие строки (итого по курсу, практики, комплексные экзамены
        # без явного номера семестра и т.п.) в план 1.1-1.4 не входят
        continue

    # --- проход 2: события -> план по курсам ---
    plan: dict[int, list[dict]] = {1: [], 2: [], 3: [], 4: []}
    current_year: int | None = None
    semester_number = 0
    semester_groups: list[dict] = []
    current_group: dict | None = None

    for ev in events:
        if ev["type"] == "year":
            current_year = ev["year"]

        elif ev["type"] == "semester_start":
            semester_number += 1
            semester_groups = []
            current_group = None

        elif ev["type"] == "discipline":
            component_ru, component_kk = _component_from_cycle(ev["cycle_ru"])
            is_elective = component_ru == "Компонент по выбору"

            if current_group is None or current_group["label_ru"] != component_ru:
                if current_group is not None:
                    semester_groups.append(current_group)
                current_group = {
                    "label_ru": component_ru,
                    "label_kk": component_kk,
                    "credits": 0,
                    "is_elective": is_elective,
                    "rows": [],
                }

            if is_elective:
                current_group["credits"] = ev["credits"]
            else:
                current_group["credits"] += ev["credits"] or 0
            current_group["rows"].append(ev)

        elif ev["type"] == "semester_total":
            if current_group is not None:
                semester_groups.append(current_group)
                current_group = None

            if current_year is not None:
                plan[current_year].append(
                    {
                        "type": "semester",
                        "number": semester_number,
                        "credits": ev["credits"],
                    }
                )
                for group in semester_groups:
                    plan[current_year].append(
                        {
                            "type": "component",
                            "label_ru": group["label_ru"],
                            "label_kk": group["label_kk"],
                            "credits": group["credits"],
                        }
                    )
                    plan[current_year].extend(group["rows"])
            semester_groups = []

    return plan


# =====================================================================
# Рендеринг таблицы в Word
# =====================================================================


def _insert_table_after(doc: DocumentObject, paragraph: Paragraph, cols: int) -> Table:
    """
    Создаёт таблицу и вставляет её сразу ПОСЛЕ заданного абзаца (как и
    _insert_paragraph_after для абзацев - python-docx не даёт метода
    "вставить после" из коробки, поэтому таблица создаётся в конце
    документа через doc.add_table(), а затем её XML-элемент переносится
    соседним с paragraph).
    """
    empty_p = OxmlElement("w:p")
    paragraph._p.addnext(empty_p)

    table = doc.add_table(rows=1, cols=cols)
    empty_p.addnext(table._tbl)
    return table


def _apply_font(run, bold: bool) -> None:
    run.bold = bold
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT_NAME)


def _set_cell_text(
    cell, text: str | None, bold: bool = False, align=WD_ALIGN_PARAGRAPH.CENTER
) -> None:
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # очищаем все абзацы/раны в ячейке (после merge() в ней могут остаться
    # пустые раны от объединяемых ячеек) и оставляем один чистый абзац
    paragraphs = cell.paragraphs
    p = paragraphs[0]
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    for extra_p in paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)

    p.alignment = align
    run = p.add_run(text or "")
    _apply_font(run, bold)


def _set_table_borders(table: Table) -> None:
    """Простая чёрная сетка 0.5pt по всем границам - как в образце."""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.append(borders)


def _set_column_widths(table: Table) -> None:
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(COLUMN_WIDTHS):
            if idx < len(row.cells):
                row.cells[idx].width = Emu(width)
    for idx, width in enumerate(COLUMN_WIDTHS):
        table.columns[idx].width = Emu(width)


def render_study_plan_table(
    doc: DocumentObject, anchor: Paragraph, course_rows: list[dict]
) -> Table:
    """Строит таблицу учебного плана для одного курса и вставляет её сразу после anchor."""
    table = _insert_table_after(doc, anchor, cols=5)
    _set_table_borders(table)

    for idx, text in enumerate(HEADER_ROW):
        _set_cell_text(table.rows[0].cells[idx], text, bold=False)

    for row_data in course_rows:
        cells = table.add_row().cells

        if row_data["type"] == "semester":
            label = f"{row_data['number']} СЕМЕСТР"
            merged = cells[0]
            for c in cells[1:4]:
                merged = merged.merge(c)
            _set_cell_text(merged, label, bold=True)
            _set_cell_text(cells[4], str(row_data["credits"] or ""), bold=True)

        elif row_data["type"] == "component":
            label = f"{row_data['label_kk']} / {row_data['label_ru']}"
            merged = cells[0]
            for c in cells[1:4]:
                merged = merged.merge(c)
            _set_cell_text(merged, label, bold=True)
            _set_cell_text(cells[4], str(row_data["credits"] or ""), bold=True)

        else:  # discipline
            _set_cell_text(cells[0], row_data["cycle_ru"])
            _set_cell_text(cells[1], f"{row_data['module_kk']}/{row_data['module_ru']}")
            _set_cell_text(cells[2], f"{row_data['code_kk']}/{row_data['code_ru']}")
            _set_cell_text(cells[3], f"{row_data['name_kk']} / {row_data['name_ru']}")
            _set_cell_text(cells[4], str(row_data["credits"] or ""))

    _set_column_widths(table)
    return table


def insert_study_plan_tables(doc: DocumentObject, xlsx_path: str) -> None:
    """
    Находит в документе заголовки "1.1 ... 1 курс", "1.2 ... 2 курс",
    "1.3 ... 3 курс", "1.4 ... 4 курс" и вставляет сразу после каждого
    таблицу учебного плана соответствующего курса.
    """
    plan = get_study_plan(xlsx_path)

    for course_year in (1, 2, 3, 4):
        heading_needle = f"{course_year} курсына арналған оқу жоспары"
        heading_paragraph = None
        for p in doc.paragraphs:
            if heading_needle in p.text:
                heading_paragraph = p
                break
        if heading_paragraph is None:
            raise ValueError(f"Заголовок для {course_year} курса не найден в шаблоне")

        render_study_plan_table(doc, heading_paragraph, plan[course_year])
