from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from ked_builder.courses import Course, get_courses
from ked_builder.results import get_results

XLSX_DOC = "./documents/--УП-6В06150-ИТИИИ набор-2025 ОБНОВЛЕННЫЙ (1).xlsx"
RU_DOC = "./documents/ОП_ИТиИИ[2025-2026] проверено 17.09.2025.docx"
KZ_DOC = "./documents/ОП_ИТиИИ_КАЗ[2025-2026].docx"
TEMPLATE_DOC = "./documents/ked_template.docx"
OUTPUT_DOC = "./documents/КЭД_ИТиИИ_2025-2026.docx"

# Заголовок, после которого нужно вставить описание элективных дисциплин.
# Ищем по вхождению подстроки, поэтому регистр и точная пунктуация не важны.
HEADING_ELECTIVES = "Элективті пәндердің сипаттамасы"

FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(11)


def format_requirements(
    requirements: list[str] | None,
    courses_by_number: dict[str, Course],
    language: str = "ru",
) -> str | None:
    if requirements is None:
        return None

    result = []

    for number in requirements:
        if "0" in number:
            result.append("Мектеп курсы" if language == "kz" else "Школьный курс")
        elif "9" in number:
            result.append(
                "Кәсіби қызмет" if language == "kz" else "Профессиональная деятельность"
            )
        elif number in courses_by_number:
            result.append(
                (
                    courses_by_number[number].name_kk
                    if language == "kz"
                    else courses_by_number[number].name
                )
                or number
            )
        else:
            result.append(number)

    return ", ".join(result)


def _insert_paragraph_after(paragraph: Paragraph) -> Paragraph:
    """
    Создаёт новый пустой абзац сразу ПОСЛЕ заданного - в отличие от
    doc.add_paragraph(), который всегда добавляет абзац в конец документа.
    python-docx не даёт такой метод "из коробки", поэтому абзац создаётся
    напрямую через XML и вставляется соседним элементом.
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)


def _set_run_font(run) -> None:
    """
    Ставит шрифт Times New Roman 11 на run. Обычного run.font.name
    недостаточно - Word хранит шрифт отдельно для обычных (ascii/hAnsi),
    кириллических/CJK (eastAsia) и специальных (cs) символов, поэтому
    прописываем явно все варианты, иначе кириллица может остаться
    в шрифте абзаца по умолчанию из шаблона.
    """
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT_NAME)


def _add_title_paragraph(anchor: Paragraph, text: str) -> Paragraph:
    """Название дисциплины: жирный текст, выровнен по центру."""
    p = _insert_paragraph_after(anchor)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    _set_run_font(run)
    return p


def _add_blank_paragraph(anchor: Paragraph) -> Paragraph:
    return _insert_paragraph_after(anchor)


def _add_label_paragraph(anchor: Paragraph, label: str, text: str | None) -> Paragraph:
    """
    Строка вида "[TAB] **Метка:** текст" - таб-отступ, жирная метка,
    обычный текст после неё, весь абзац выровнен по ширине.
    """
    p = _insert_paragraph_after(anchor)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    tab_run = p.add_run("\t")
    _set_run_font(tab_run)

    label_run = p.add_run(label)
    label_run.bold = True
    _set_run_font(label_run)

    if text:
        text_run = p.add_run(" " + text)
        _set_run_font(text_run)

    return p


def _render_course_block(
    anchor: Paragraph,
    course: Course,
    courses_by_number: dict[str, Course],
    results_ru: dict[str, str],
    results_kz: dict[str, str],
) -> Paragraph:
    """
    Вставляет одну дисциплину в документ (казахский блок, затем русский)
    и возвращает последний вставленный абзац - новый "якорь" для
    следующей дисциплины.
    """
    # --- казахский блок ---
    code_kk = course.code_kk.split(" ")[0] if course.code_kk else ""
    anchor = _add_title_paragraph(anchor, f"{code_kk} {course.name_kk or ''}".strip())
    anchor = _add_blank_paragraph(anchor)
    anchor = _add_label_paragraph(
        anchor,
        "Пререквизеттері:",
        format_requirements(course.prerequisites, courses_by_number, "kz"),
    )
    anchor = _add_label_paragraph(
        anchor,
        "Постреквизиттері:",
        format_requirements(course.postrequisites, courses_by_number, "kz"),
    )
    anchor = _add_label_paragraph(anchor, "Оқу мақсаты:", course.goal_kk)
    anchor = _add_label_paragraph(anchor, "Курстың қысқаша мазмұны:", course.summary_kk)
    if course.results is not None:
        anchor = _add_label_paragraph(
            anchor,
            "Оқыту нәтижесі:",
            " ".join(results_kz.get(r, r) for r in course.results),
        )
    anchor = _add_label_paragraph(anchor, "Бағдарлама жетекшісі:", course.mentor)
    anchor = _add_label_paragraph(anchor, "Кафедра:", course.department)
    anchor = _add_blank_paragraph(anchor)

    # --- русский блок ---
    code_ru = course.code.split(" ")[0] if course.code else ""
    anchor = _add_title_paragraph(anchor, f"{code_ru} {course.name or ''}".strip())
    anchor = _add_blank_paragraph(anchor)
    anchor = _add_label_paragraph(
        anchor,
        "Пререквизиты:",
        format_requirements(course.prerequisites, courses_by_number),
    )
    anchor = _add_label_paragraph(
        anchor,
        "Постреквизиты:",
        format_requirements(course.postrequisites, courses_by_number),
    )
    anchor = _add_label_paragraph(anchor, "Цель изучения:", course.goal)
    anchor = _add_label_paragraph(anchor, "Краткое содержание курса:", course.summary)
    if course.results is not None:
        anchor = _add_label_paragraph(
            anchor,
            "Результаты обучения:",
            " ".join(results_ru.get(r, r) for r in course.results),
        )
    anchor = _add_label_paragraph(anchor, "Руководитель программы:", course.mentor)
    anchor = _add_label_paragraph(anchor, "Кафедра:", course.department)
    anchor = _add_blank_paragraph(anchor)

    return anchor


def build_ked(
    xlsx_doc: str = XLSX_DOC,
    ru_doc: str = RU_DOC,
    kz_doc: str = KZ_DOC,
    template_doc: str = TEMPLATE_DOC,
    output_doc: str = OUTPUT_DOC,
) -> None:
    courses = get_courses(xlsx_doc, ru_doc, kz_doc)
    results_ru = get_results(ru_doc)
    results_kz = get_results(kz_doc, "kz")

    courses_by_number = {c.number: c for c in courses if c.number is not None}
    # в каталог элективных дисциплин попадают только дисциплины
    # "Компонента по выбору" (см. Course.is_elective)
    electives = [c for c in courses if c.is_elective]

    missing_kk = [c.name for c in electives if not c.name_kk]
    missing_goal = [c.name for c in electives if not c.goal]
    if missing_kk:
        print("Внимание: нет казахского названия для дисциплин:")
        for name in missing_kk:
            print(f"  - {name}")
    if missing_goal:
        print("Внимание: не найдено описание (goal/summary) для дисциплин:")
        for name in missing_goal:
            print(f"  - {name}")

    doc = Document(template_doc)

    heading_paragraph = None
    for p in doc.paragraphs:
        if HEADING_ELECTIVES in p.text:
            heading_paragraph = p
            break
    if heading_paragraph is None:
        raise ValueError(f'Заголовок "{HEADING_ELECTIVES}" не найден в шаблоне')

    anchor = _add_blank_paragraph(heading_paragraph)

    for course in electives:
        anchor = _render_course_block(
            anchor, course, courses_by_number, results_ru, results_kz
        )

    doc.save(output_doc)
    print(f"Сохранено: {output_doc} ({len(electives)} дисциплин)")


if __name__ == "__main__":
    build_ked()
