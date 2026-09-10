from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import inspect
import logging

import converter_pdf


PDF = Path("ANALITICO_TAXA_0054041451-1.PDF")


def test_split_code_and_name_uses_code_pattern():
    assert converter_pdf.split_code_and_name(
        "0976.0663.000060-00   ADRIANA NASS"
    ) == ("0976.0663.000060-00", "ADRIANA NASS")
    assert converter_pdf.split_code_and_name(
        "0976.0663.000061-00   JOAO CARLOS DA SILVA"
    ) == ("0976.0663.000061-00", "JOAO CARLOS DA SILVA")


def test_field_text_preserves_source_whitespace():
    assert converter_pdf._field_text([(100.5, "  original\tvalue  ")], 100.5, 1.0) == (
        "  original\tvalue  "
    )


def test_record_operation_preserves_existing_span_separation():
    layout = converter_pdf.DEFAULT_LAYOUT
    record, _ = converter_pdf._record_from_row(
        1,
        10.0,
        [
            (layout.field_x["operacao"], "  A\t"),
            (layout.field_x["lancamento"], "  MENSALIDADE  "),
        ],
        ("", ""),
        layout,
    )

    assert record.op_lancamento == "  A\t  MENSALIDADE  "
    assert converter_pdf._compose_operation("A", "MENSALIDADE") == "A MENSALIDADE"


def test_reference_pdf_reconstructs_all_records_and_preserves_fields():
    result = converter_pdf.extract_records(PDF)

    assert result.page_count == 10
    assert len(result.records) == 185
    assert len(result.titular_totals) == 137
    assert result.records[0].codigo_beneficiario == "0976.0663.000060-00" + " " * 32
    assert result.records[0].beneficiario == "ADRIANA NASS" + " " * 39
    assert result.records[0].quantidade == "   1,00000"
    assert result.records[0].valor_unit == "           484,75"
    assert result.records[0].valor_bonif == "           -77,56"
    assert result.records[0].valor_total == "           407,19"

    special = next(record for record in result.records if record.valor_unit.strip() == "-169,90")
    assert special.quantidade == "   0,58065"
    assert special.valor_bonif == "             0,00"
    assert special.valor_total == "           -98,65"
    assert result.page_record_counts == [20, 24, 22, 21, 22, 20, 20, 24, 12, 0]


def test_validation_reports_no_total_divergence():
    result = converter_pdf.extract_records(PDF)
    assert converter_pdf.validate_records(result) == result.observations
    assert not any("diverge" in message for message in result.observations)


def test_cli_writes_ten_column_workbook(tmp_path):
    output = tmp_path / "resultado.xlsx"
    exit_code = converter_pdf.run([str(PDF), "-o", str(output)])
    assert exit_code == 0
    assert output.exists()

    with ZipFile(output) as archive:
        assert archive.testzip() is None
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows = sheet.findall(".//main:sheetData/main:row", namespace)
    assert len(rows) == 186
    assert sheet.find("main:autoFilter", namespace).attrib["ref"] == "A1:J186"
    assert sheet.find(".//main:pane", namespace).attrib["state"] == "frozen"


def test_cli_rejects_missing_input(tmp_path):
    assert converter_pdf.run([str(tmp_path / "missing.pdf")]) == 2


def test_legacy_cli_rejects_incompatible_layout(tmp_path):
    output = tmp_path / "nao_deve_existir.xlsx"

    assert converter_pdf.run(
        ["FLEXIV_1166046_1_082026.pdf", "-o", str(output)]
    ) == 1
    assert not output.exists()


def test_legacy_cli_returns_error_code_for_invalid_arguments():
    assert converter_pdf.run([]) == 2


def test_cli_logs_pages_lines_and_observations(tmp_path, caplog):
    output = tmp_path / "resultado.xlsx"
    from cli import run

    with caplog.at_level(logging.INFO, logger="rateiosrh.cli"):
        assert run(
            [
                "bradesco_dental",
                "FLEXIV_1166046_1_082026.pdf",
                "-o",
                str(output),
            ]
        ) == 0

    assert "7 páginas" in caplog.text
    assert "279 linhas" in caplog.text
    assert "0 observações" in caplog.text


def test_legacy_facade_exports_keep_their_public_signatures():
    assert callable(converter_pdf.extract_records)
    assert callable(converter_pdf.validate_records)
    assert callable(converter_pdf.write_workbook)
    assert callable(converter_pdf.run)
    assert list(inspect.signature(converter_pdf.extract_records).parameters) == [
        "pdf_path",
        "layout",
    ]
    assert list(inspect.signature(converter_pdf.validate_records).parameters) == [
        "result",
    ]
    assert list(inspect.signature(converter_pdf.write_workbook).parameters) == [
        "result",
        "output_path",
    ]
    assert list(inspect.signature(converter_pdf.run).parameters) == ["argv"]


def test_cli_writes_bradesco_dental_workbook(tmp_path):
    output = tmp_path / "resultado.xlsx"
    from cli import run

    assert run(
        [
            "bradesco_dental",
            "FLEXIV_1166046_1_082026.pdf",
            "-o",
            str(output),
        ]
    ) == 0
    assert output.exists()


def test_cli_rejects_incompatible_selected_parser(tmp_path):
    output = tmp_path / "nao_deve_existir.xlsx"
    from cli import run

    assert run(
        [
            "bradesco_dental",
            "ANALITICO_TAXA_0054041451-1.PDF",
            "-o",
            str(output),
        ]
    ) != 0
    assert not output.exists()
