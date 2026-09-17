import difflib
import re

from docx import Document

from ked_builder.utils import find_table_after_heading


class Course:
    def __init__(
        self,
        name: str | None = None,
        name_kk: str | None = None,
        code: str | None = None,
        code_kk: str | None = None,
        number: str | None = None,
        prerequisites: list[str] | None = None,
        postrequisites: list[str] | None = None,
        goal: str | None = None,
        goal_kk: str | None = None,
        summary: str | None = None,
        results: list[str] | None = None,
        mentor: str | None = None,
        department: str | None = None,
        component: str | None = None,
    ):
        self.name = name
        self.name_kk = name_kk
        self.code = code
        self.code_kk = code_kk
        self.number = number
        self.prerequisites = prerequisites
        self.postrequisites = postrequisites
        self.goal = goal
        self.goal_kk = goal_kk
        self.summary = summary
        self.results = results
        self.mentor = mentor
        self.department = department
        # тип компонента дисциплины, например "Обязательный компонент",
        # "Вузовский компонент" или "Компонент по выбору" - берётся из
        # строки-разделителя раздела, под которой стоит дисциплина в
        # таблице "Сведения о дисциплинах образовательной программы"
        self.component = component

    def __repr__(self):
        return (
            f"Course(name={self.name!r}, name_kk={self.name_kk!r}, code={self.code!r}, "
            f"code_kk={self.code_kk!r}, number={self.number!r}, component={self.component!r}, "
            f"prerequisites={self.prerequisites!r}, postrequisites={self.postrequisites!r})"
        )

    @property
    def is_elective(self) -> bool:
        """Является ли дисциплина компонентом по выбору (элективом)."""
        return bool(self.component) and "выбор" in self.component.lower()


HEADING_COURSES = "Сведения о дисциплинах образовательной программы"
HEADING_COURSES_KK = "Білім беру бағдарламасының пәндері туралы мәліметтер"
HEADING_CODES = "Содержание образовательной программы"
HEADING_CODES_KK = "Білім беру бағдарламасының мазмұны"
HEADING_PREREQUISITES = "Карта соответствия пререквизитов и постреквизитов дисциплин"


def _normalize_name(name: str) -> str:
    """
    Приводит название дисциплины к упрощённому виду для сравнения:
    нижний регистр, "," -> ".", убираем приставку "Minor 1/2 ..." и
    лишние пробелы. Нужно, потому что в разных таблицах документа одно
    и то же название дисциплины иногда записано по-разному
    (например "Психология, Культурология" и "Психология. Культурология",
    или "Minor 1 Организация бизнеса" и просто "Организация бизнеса").
    """
    n = name.lower().strip()
    n = n.replace(",", ".")
    n = re.sub(r"^minor\s*\d+\.?\s*", "", n)
    n = re.sub(r"\s+", " ", n)
    return n.strip(" .")


def _lookup(name: str, exact: dict, normalized: dict, normalized_keys: list[str]):
    """
    Ищет значение для названия дисциплины в справочнике, полученном из
    другой таблицы документа. Порядок попыток:
    1. точное совпадение названия,
    2. совпадение после нормализации (регистр, "," vs ".", "Minor ..."),
    3. ближайшее по написанию название (спасает от опечаток в документе).
    """
    if name in exact:
        return exact[name]

    norm = _normalize_name(name)
    if norm in normalized:
        return normalized[norm]

    match = difflib.get_close_matches(norm, normalized_keys, n=1, cutoff=0.75)
    if match:
        return normalized[match[0]]

    return None


def get_course_codes(
    docx_path: str, heading_text: str = HEADING_CODES
) -> dict[str, str]:
    """
    Из таблицы под заголовком "Содержание образовательной программы"
    (или её казахского аналога, если передать heading_text=HEADING_CODES_KK)
    собирает словарь {название дисциплины: код дисциплины}.

    Структура таблицы: 2 строки заголовка (первая - с объединёнными
    ячейками "Компоненты модуля", вторая - с реальными подзаголовками),
    данные начинаются с 3-й строки (индекс 2).
    Нужные столбцы: cells[4] - код дисциплины, cells[5] - название дисциплины.

    Если название встречается несколько раз (например "Физическая культура"
    по семестрам с разными кодами) - берётся код из первого найденного
    вхождения.
    """
    doc = Document(docx_path)

    table = find_table_after_heading(doc, heading_text)
    if table is None:
        raise ValueError("Таблица после заголовка не найдена")

    codes: dict[str, str] = {}
    for row in table.rows[2:]:  # пропускаем 2 строки заголовка
        cells = row.cells
        if len(cells) < 6:
            continue

        name = cells[5].text.strip()
        code = cells[4].text.strip()

        if name and name not in codes:
            codes[name] = code

    return codes


def _code_suffix(code: str) -> str:
    """
    Извлекает языконезависимую часть кода дисциплины - цифры (и, если есть,
    номер в скобках), без учёта пробелов. Буквенная часть кода - это
    транслитерация названия дисциплины, поэтому она разная в русской и
    казахской версиях (например "IK 1101" и "KT 1101" - это одна и та же
    дисциплина "История Казахстана"/"Қазақстан тарихы", но код у неё в
    каждой версии свой - см. Course.code / Course.code_kk). А вот числовая
    часть кода одинакова в обеих версиях - на ней и сопоставляем дисциплины
    между языками.
    """
    match = re.search(r"(\d+)\s*(\([^)]*\))?\s*$", code)
    if not match:
        return code.strip()
    number = match.group(1)
    parens = re.sub(r"\s+", "", match.group(2) or "")
    return number + parens


# строка-номер дисциплины вида "1-1", "4-3", либо просто "1" (для
# "Школьного курса" из нулевого семестра, у которого номер в таблице
# не содержит дефиса)
_NUMBER_PATTERN = re.compile(r"^\d+(-\d+)?$")


def get_course_numbers(docx_path: str) -> dict[str, dict]:
    """
    Из таблицы "Карта соответствия пререквизитов и постреквизитов дисциплин"
    собирает для каждой дисциплины:
        - number: её номер в карте (например "1-1")
        - prerequisites: список номеров дисциплин-пререквизитов
        - postrequisites: список номеров дисциплин-постреквизитов

    Возвращает словарь {название дисциплины: {number, prerequisites, postrequisites}}.
    """
    doc = Document(docx_path)

    table = find_table_after_heading(doc, HEADING_PREREQUISITES)
    if table is None:
        raise ValueError("Таблица после заголовка не найдена")

    entries: dict[str, dict] = {}
    for row in table.rows:
        cells = row.cells
        if len(cells) != 10:
            continue  # необычная строка (например строка про диплом) - пропускаем

        raw_number = cells[2].text.strip()
        if not _NUMBER_PATTERN.match(raw_number):
            continue  # это строка-разделитель семестра, а не дисциплина

        semester = cells[0].text.strip()
        number = raw_number if "-" in raw_number else f"{semester}-{raw_number}"

        name = cells[3].text.strip()
        prereq_raw = cells[6].text.strip()
        postreq_raw = cells[8].text.strip()

        def parse_list(raw: str) -> list[str]:
            if raw in ("", "-"):
                return []
            return [item.strip() for item in re.split(r",\s*", raw) if item.strip()]

        if name and name not in entries:
            entries[name] = {
                "number": number,
                "prerequisites": parse_list(prereq_raw),
                "postrequisites": parse_list(postreq_raw),
            }

    return entries


def get_course_texts(docx_path: str, heading_text: str) -> dict[str, str]:
    """
    Из таблицы "Сведения о дисциплинах ..." (на любом языке, определяется
    параметром heading_text) собирает словарь {название дисциплины: цель/описание}.
    """
    doc = Document(docx_path)

    table = find_table_after_heading(doc, heading_text)
    if table is None:
        raise ValueError("Таблица после заголовка не найдена")

    texts: dict[str, str] = {}
    for row in table.rows:
        cells = row.cells
        if len(cells) < 4:
            continue

        row_number = cells[0].text.strip()
        if not row_number.isdigit():
            continue  # строка-разделитель раздела, а не дисциплина

        name = cells[2].text.strip()
        goal = cells[3].text.strip()
        if name and name not in texts:
            texts[name] = goal

    return texts


def get_courses(
    docx_path: str,
    docx_path_kk: str | None = None,
    kk_overrides: dict[str, str] | None = None,
) -> list[Course]:
    """
    :param docx_path: путь к документу образовательной программы на русском языке
    :param docx_path_kk: (опционально) путь к тому же документу на казахском
                         языке - если передан, у каждой дисциплины дополнительно
                         заполняются name_kk, goal_kk и code_kk
    :param kk_overrides: (опционально) словарь {русское название: казахское
                         название} для дисциплин, которые не удаётся сопоставить
                         автоматически по коду. Это в первую очередь актуально
                         для "Компонента по выбору": там несколько дисциплин
                         могут делить один номер слота в коде, а в казахском
                         документе у части кодов бывают ошибки (например,
                         пропущена числовая часть кода). Автоматическое
                         сопоставление в таких случаях не рискует подставлять
                         перевод наугад - вместо этого можно явно указать
                         правильные пары вручную через этот параметр.
    """
    doc = Document(docx_path)

    target_table = find_table_after_heading(doc, HEADING_COURSES)
    if target_table is None:
        raise ValueError("Таблица после заголовка не найдена")

    codes = get_course_codes(docx_path)
    normalized_codes = {_normalize_name(name): code for name, code in codes.items()}
    normalized_code_keys = list(normalized_codes.keys())

    numbers = get_course_numbers(docx_path)
    normalized_numbers = {
        _normalize_name(name): entry for name, entry in numbers.items()
    }
    normalized_number_keys = list(normalized_numbers.keys())

    # казахские название и код по числовому суффиксу кода, и название -> цель.
    # Сопоставляем дисциплины между языками не по названию (они разные)
    # и не по позиции в таблице (порядок местами отличается между языками),
    # а по числовой части кода дисциплины - она одинакова в обеих версиях
    # документа для большинства дисциплин, даже когда буквенная часть кода
    # (транслитерация) разная - поэтому у казахской дисциплины СВОЙ код
    # (code_kk), а не тот же самый, что и в русской версии (code).
    #
    # Исключение - "Компонент по выбору": там несколько РАЗНЫХ дисциплин
    # могут делить один и тот же числовой суффикс кода (он обозначает номер
    # слота выбора, а не саму дисциплину), и состав вариантов в русской и
    # казахской версиях документа может не совпадать по количеству и
    # порядку. Такие неоднозначные случаи сознательно НЕ сопоставляются
    # автоматически (name_kk/goal_kk/code_kk останутся None), чтобы не
    # подставить неверный перевод - их нужно сверить вручную.
    kk_entry_by_suffix: dict[str, tuple[str, str]] = {}
    kk_codes: dict[str, str] = {}
    normalized_kk_codes: dict[str, str] = {}
    normalized_kk_code_keys: list[str] = []
    kk_texts_by_name: dict[str, str] = {}
    if docx_path_kk:
        kk_codes = get_course_codes(docx_path_kk, HEADING_CODES_KK)
        normalized_kk_codes = {
            _normalize_name(name): code for name, code in kk_codes.items()
        }
        normalized_kk_code_keys = list(normalized_kk_codes.keys())

        suffix_candidates: dict[str, list[tuple[str, str]]] = {}
        for kk_name, kk_code in kk_codes.items():
            suffix_candidates.setdefault(_code_suffix(kk_code), []).append(
                (kk_name, kk_code)
            )
        kk_entry_by_suffix = {
            suffix: entries[0]
            for suffix, entries in suffix_candidates.items()
            if len(entries) == 1
        }
        kk_texts_by_name = get_course_texts(docx_path_kk, HEADING_COURSES_KK)

    normalized_kk_overrides = (
        {_normalize_name(name): kk_name for name, kk_name in kk_overrides.items()}
        if kk_overrides
        else {}
    )
    normalized_kk_override_keys = list(normalized_kk_overrides.keys())

    # Разбираем строки таблицы. "Широкие" строки-разделители (например
    # "Цикл базовых дисциплин / Компонент по выбору") - это одна
    # объединённая на всю ширину ячейка, поэтому первый столбец (№) там
    # НЕ является числом - по этому признаку их и отличаем от обычных
    # строк с дисциплиной. Текст такой строки запоминаем в current_component
    # и приписываем всем следующим за ней дисциплинам, пока не встретится
    # новый разделитель.
    courses = []
    current_component = None
    for row in target_table.rows:
        cells = row.cells
        if len(cells) < 4:
            continue

        row_number = cells[0].text.strip()
        if not row_number.isdigit():
            # строка-разделитель раздела: например
            # "Цикл базовых дисциплин\nКомпонент по выбору" -
            # компонент - это последняя строка текста
            current_component = cells[0].text.strip().split("\n")[-1]
            continue

        results = cells[1].text.strip().split()
        name = cells[2].text.strip()
        goal = cells[3].text.strip()

        code = _lookup(name, codes, normalized_codes, normalized_code_keys)

        number_entry = _lookup(
            name, numbers, normalized_numbers, normalized_number_keys
        )
        number = number_entry["number"] if number_entry else None
        prerequisites = number_entry["prerequisites"] if number_entry else None
        postrequisites = number_entry["postrequisites"] if number_entry else None

        name_kk = None
        code_kk = None
        goal_kk = None

        # 1. ручное исключение (самое надёжное - указано человеком)
        override_kk_name = None
        if kk_overrides:
            override_kk_name = _lookup(
                name, kk_overrides, normalized_kk_overrides, normalized_kk_override_keys
            )

        if override_kk_name:
            name_kk = override_kk_name
            code_kk = _lookup(
                name_kk, kk_codes, normalized_kk_codes, normalized_kk_code_keys
            )
            goal_kk = kk_texts_by_name.get(name_kk)
        # 2. иначе - автоматическое сопоставление по числовому суффиксу кода
        elif code and kk_entry_by_suffix:
            kk_entry = kk_entry_by_suffix.get(_code_suffix(code))
            if kk_entry:
                name_kk, code_kk = kk_entry
                goal_kk = kk_texts_by_name.get(name_kk)

        courses.append(
            Course(
                results=results,
                name=name,
                goal=goal,
                name_kk=name_kk,
                goal_kk=goal_kk,
                code=code,
                code_kk=code_kk,
                number=number,
                prerequisites=prerequisites,
                postrequisites=postrequisites,
                component=current_component,
            )
        )

    return courses
