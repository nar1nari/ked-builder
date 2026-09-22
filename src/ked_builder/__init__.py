from ked_builder.generate_ked import build_ked

def main() -> None:
    XLSX_DOC = "./documents/--УП-6В06150-ИТИИИ набор-2025 ОБНОВЛЕННЫЙ (1).xlsx"
    RU_DOC = "./documents/ОП_ИТиИИ[2025-2026] проверено 17.09.2025.docx"
    KZ_DOC = "./documents/ОП_ИТиИИ_КАЗ[2025-2026].docx"
    OUTPUT_DOC = "./documents/КЭД_ИТиИИ_2025-2026.docx"

    build_ked(
        year=2025,
        program_number="6В06150",
        program_name="Информационные технологии и искусственный интеллект",
        program_name_kk="Ақпараттық технологиялар және жасанды интеллект",
        xlsx_doc=XLSX_DOC,
        ru_doc=RU_DOC,
        kz_doc=KZ_DOC,
        output_doc=OUTPUT_DOC
    )
