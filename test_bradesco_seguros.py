from pathlib import Path
import zlib

import pandas as pd
import pytest

from rateiosrh.core.pdf_reader import PdfDocument
from rateiosrh.parsers.bradesco_seguros import BradescoSegurosParser
from rateiosrh.registry import list_parsers


PDF = Path("bradescoseguros.pdf")
REFERENCE_XLSX = Path("rateamento-bradesco_seguros.xlsx")
EXPECTED_COLUMNS = [
    "Matrícula",
    "Certificado",
    "Segurado",
    "Tipo Segurado",
    "Data Nascto.",
    "Data Inclusão",
    "Início Vigência",
    "Fim Vigência",
    "Prazo",
    "Capital",
    "Prêmio",
]


def _reference_as_text() -> pd.DataFrame:
    return pd.read_excel(REFERENCE_XLSX, dtype="string").fillna("")


def test_bradesco_seguros_parser_has_exact_schema_and_row_count():
    frame = BradescoSegurosParser().parse(PDF)

    assert frame.shape == (185, 11)
    assert frame.columns.tolist() == EXPECTED_COLUMNS
    assert all(dtype == "string" for dtype in frame.dtypes.astype(str))


def test_bradesco_seguros_preserves_text_and_explicit_name_continuations():
    frame = BradescoSegurosParser().parse(PDF)

    assert frame.iloc[0].tolist() == [
        "11920429999",
        "353",
        "ABIGAIL SCHMIDT",
        "P",
        "19/05/2007",
        "01/07/2026",
        "01/04/2026",
        "",
        "",
        "203.168,00",
        "26,71",
    ]
    assert frame.loc[frame["Certificado"].eq("131"), "Segurado"].item() == (
        "DANIELE PAULA DOS SANTOS CARDOSO"
    )
    assert frame.loc[frame["Matrícula"].eq("04008240004"), "Matrícula"].item() == (
        "04008240004"
    )
    assert frame["Fim Vigência"].eq("").all()
    assert frame["Prazo"].eq("").all()
    assert frame["Matrícula"].str.len().eq(11).all()


def test_bradesco_seguros_layout_is_specific_and_registry_is_automatic():
    parser = BradescoSegurosParser()

    assert parser.parser_id == "bradesco_seguros"
    assert parser.display_name == "Converter Bradesco Seguros"
    assert parser.validate_layout(PDF)
    assert not parser.validate_layout(Path("ANALITICO_TAXA_0054041451-1.PDF"))
    assert not parser.validate_layout(Path("FLEXIV_1166046_1_082026.pdf"))
    assert {item["id"] for item in list_parsers()} >= {
        "unimed",
        "bradesco_dental",
        "bradesco_seguros",
    }


def test_bradesco_seguros_validation_is_non_mutating():
    parser = BradescoSegurosParser()
    frame = parser.parse(PDF)
    before = frame.copy(deep=True)

    assert parser.validate_dataframe(frame) == []
    pd.testing.assert_frame_equal(frame, before)


def test_inline_image_removal_accepts_pdf_whitespace_variants():
    parser = BradescoSegurosParser()
    image = zlib.compress(b"image payload")
    stream = (
        b"q BI \r\n/W 1\r\n/H 1\r\n/F /FlateDecode\r\nID\r\n"
        + image
        + b"\r\n EI\t\r\nBT\n(visible)Tj\nET"
    )

    cleaned = parser._remove_inline_images(stream)

    assert image not in cleaned
    assert b"image payload" not in cleaned
    assert b"(visible)Tj" in cleaned


def test_layout_rejects_report_with_identifiable_incompatible_page(monkeypatch):
    parser = BradescoSegurosParser()
    valid_rows = parser._page_rows(PdfDocument(PDF.read_bytes()).page_streams()[0])
    invalid_rows = [(545.0, [(0.0, "Relação dos Segurados Vigentes")])]

    class FakeDocument:
        def page_streams(self):
            return [b"valid", b"invalid"]

    monkeypatch.setattr(
        BradescoSegurosParser,
        "_page_rows",
        classmethod(
            lambda cls, stream: valid_rows if stream == b"valid" else invalid_rows
        ),
    )

    assert not parser._document_matches(FakeDocument())


def test_bradesco_seguros_matches_the_checked_excel_literal_values():
    frame = BradescoSegurosParser().parse(PDF)

    pd.testing.assert_frame_equal(frame, _reference_as_text(), check_dtype=True)
