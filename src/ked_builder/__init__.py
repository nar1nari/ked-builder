import argparse

from ked_builder.generate_ked import build_ked


def main() -> None:
    parser = argparse.ArgumentParser(description="Сгенерировать КЭД из ОП и УП")

    parser.add_argument(
        "--year",
        type=int,
        required=True,
        help="Год начала образовательной программы.",
    )
    parser.add_argument(
        "--program-number",
        required=True,
        help="Код образовательной программы.",
    )
    parser.add_argument(
        "--program-name",
        required=True,
        help="Наименование образовательной программы на русском.",
    )
    parser.add_argument(
        "--program-name-kk",
        required=True,
        help="Наименование образовательной программы на казахском.",
    )
    parser.add_argument(
        "--xlsx-doc",
        required=True,
        help="Путь до xlsx документа УП",
    )
    parser.add_argument(
        "--ru-doc",
        required=True,
        help="Путь до docx документа ОП.",
    )
    parser.add_argument(
        "--kz-doc",
        required=True,
        help="Путь до документа ОП на казахском.",
    )
    parser.add_argument(
        "--output-doc",
        required=True,
        help="Путь для сохранения КЭД.",
    )
    parser.add_argument(
        "--template-doc",
        help="Путь до шаблона КЭД. (опциональный)",
    )
    parser.add_argument(
        "--pps-doc",
        help="Путь до файла ППС (опциональный).",
    )

    args = parser.parse_args()

    build_ked(
        year=args.year,
        program_number=args.program_number,
        program_name=args.program_name,
        program_name_kk=args.program_name_kk,
        xlsx_doc=args.xlsx_doc,
        ru_doc=args.ru_doc,
        kz_doc=args.kz_doc,
        output_doc=args.output_doc,
        template_doc=args.template_doc,
        pps_doc=args.pps_doc,
    )
