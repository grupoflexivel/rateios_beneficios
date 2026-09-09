from pathlib import Path

import pandas as pd

from rateiosrh.parsers.bradesco_dental import BradescoDentalParser


PDF = Path("FLEXIV_1166046_1_082026.pdf")
REFERENCE_XLSX = Path("FLEXIV_1166046_1_082026_referencia.xlsx")
EXPECTED_COLUMNS = [
    "Certif.",
    "Nome Beneficiário",
    "Subfatura Nº N=Nova",
    "Data Nascimento",
    "Sexo",
    "Est. Civil",
    "Paren.",
    "Plano",
    "Data Início",
    "Mov.",
    "Mês/ano",
    "Valor",
    "Part. Benef",
]


def _reference_as_text() -> pd.DataFrame:
    return pd.read_excel(REFERENCE_XLSX, dtype="string").fillna("")


def test_bradesco_dental_parser_has_exact_schema_and_row_count():
    frame = BradescoDentalParser().parse(PDF)

    assert frame.shape == (279, 13)
    assert frame.columns.tolist() == EXPECTED_COLUMNS
    assert all(dtype == "string" for dtype in frame.dtypes.astype(str))


def test_bradesco_dental_matches_the_checked_excel_literal_values():
    frame = BradescoDentalParser().parse(PDF)

    pd.testing.assert_frame_equal(frame, _reference_as_text(), check_dtype=True)


def test_bradesco_dental_preserves_continuations_and_trailing_negative_sign():
    frame = BradescoDentalParser().parse(PDF)

    assert frame["Certif."].eq("").sum() == 77
    continuation = frame[frame["Certif."].eq("")].iloc[0]
    assert continuation["Mov."] == "CM"
    assert continuation["Valor"] == "15,00-"
    assert frame["Subfatura Nº N=Nova"].eq("").all()


def test_bradesco_dental_rejects_unimed_pdf():
    assert not BradescoDentalParser().validate_layout(
        Path("ANALITICO_TAXA_0054041451-1.PDF")
    )


def test_bradesco_dental_dataframe_validation_is_structural_and_non_mutating():
    parser = BradescoDentalParser()
    frame = parser.parse(PDF)
    before = frame.copy(deep=True)

    assert parser.validate_dataframe(frame) == []
    pd.testing.assert_frame_equal(frame, before)


def test_bradesco_dental_validation_reports_non_textual_movement_without_raising():
    parser = BradescoDentalParser()
    frame = pd.DataFrame(
        [
            [
                "0000001/00",
                "NOME",
                "",
                "01/01/2000",
                "MAS",
                "SOLT",
                "",
                "TNDA",
                "01/01/2026",
                7,
                "08/2026",
                "15,00",
                "0,00",
            ]
        ],
        columns=parser.columns,
    )

    errors = parser.validate_dataframe(frame)

    assert "Todos os campos devem ser texto." in errors
    assert "Movimentos fora do conjunto esperado: 7" in errors
