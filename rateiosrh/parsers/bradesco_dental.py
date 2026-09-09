"""Parser for the detailed Bradesco Dental invoice layout."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

import pandas as pd

from rateiosrh.core.exceptions import LayoutMismatchError, PdfStructureError
from rateiosrh.core.pdf_reader import PdfDocument, group_rows
from rateiosrh.parsers.base import BaseParser
from rateiosrh.services.validation import ValidationService


Span = Tuple[float, str]
Row = Tuple[float, List[Span]]


class BradescoDentalParser(BaseParser):
    """Extract one DataFrame row for each detailed line in the invoice."""

    parser_id = "bradesco_dental"
    display_name = "Converter Bradesco Dental"
    columns = [
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

    # The intervals follow the fixed x positions documented for this layout.
    _FIELD_RANGES = (
        (35.0, 70.0),
        (70.0, 230.0),
        (230.0, 260.0),
        (260.0, 295.0),
        (295.0, 315.0),
        (315.0, 338.0),
        (338.0, 360.0),
        (360.0, 385.0),
        (385.0, 420.0),
        (420.0, 440.0),
        (440.0, 485.0),
        (485.0, 545.0),
        (545.0, 570.0),
    )
    _DETAIL_Y_RANGE = (4.70, 748.70)
    _DOCUMENT_MARKER = "BRADESCO DENTAL - FATURA TECNICA"
    _DETAIL_HEADERS = (
        "Certif.",
        "Nome Beneficiário",
        "Mês/Ano",
        "Valor",
        "Part. Benef.",
    )
    _EXPECTED_MOVEMENTS = {"", "CM", "CR", "IR", "IM"}
    _REQUIRED_ALL_ROWS = ("Plano", "Data Início", "Mês/ano", "Valor", "Part. Benef")
    _REQUIRED_COMPLETE_ROWS = (
        "Certif.",
        "Nome Beneficiário",
        "Data Nascimento",
        "Sexo",
        "Est. Civil",
        "Plano",
        "Data Início",
        "Mês/ano",
        "Valor",
        "Part. Benef",
    )
    _CONTINUATION_LEADING_FIELDS = (
        "Certif.",
        "Nome Beneficiário",
        "Subfatura Nº N=Nova",
        "Data Nascimento",
        "Sexo",
        "Est. Civil",
        "Paren.",
    )

    @classmethod
    def _page_text(cls, rows: Iterable[Row]) -> str:
        return " ".join(text for _, spans in rows for _, text in spans)

    @classmethod
    def _is_detail_page(cls, rows: Sequence[Row]) -> bool:
        page_text = cls._page_text(rows)
        return all(header in page_text for header in cls._DETAIL_HEADERS)

    @classmethod
    def _document_matches(cls, document: PdfDocument) -> bool:
        has_document_marker = False
        has_detail_page = False
        for stream in document.page_streams():
            rows = group_rows(stream)
            page_text = cls._page_text(rows)
            has_document_marker = has_document_marker or cls._DOCUMENT_MARKER in page_text
            has_detail_page = has_detail_page or cls._is_detail_page(rows)
        return has_document_marker and has_detail_page

    def validate_layout(self, pdf_path: Path) -> bool:
        try:
            document = PdfDocument(Path(pdf_path).read_bytes())
            return self._document_matches(document)
        except (OSError, PdfStructureError):
            return False

    @staticmethod
    def _field_text(spans: Sequence[Span], x_range: Tuple[float, float]) -> str:
        lower, upper = x_range
        values = [text for x, text in spans if lower <= x < upper]
        return "".join(values)

    @classmethod
    def _row_values(cls, spans: Sequence[Span]) -> List[str]:
        return [cls._field_text(spans, x_range) for x_range in cls._FIELD_RANGES]

    @classmethod
    def _detail_rows(cls, document: PdfDocument) -> List[List[str]]:
        records: List[List[str]] = []
        for stream in document.page_streams():
            rows = group_rows(stream)
            if not cls._is_detail_page(rows):
                continue
            for y, spans in rows:
                if not cls._DETAIL_Y_RANGE[0] <= y <= cls._DETAIL_Y_RANGE[1]:
                    continue
                values = cls._row_values(spans)
                if any(values):
                    records.append(values)
        return records

    def parse(self, pdf_path: Path) -> pd.DataFrame:
        path = Path(pdf_path)
        document = PdfDocument(path.read_bytes())
        if not self._document_matches(document):
            raise LayoutMismatchError(self.parser_id, path)

        frame = pd.DataFrame(self._detail_rows(document), columns=self.columns)
        return frame.astype("string")

    def validate_dataframe(self, df: pd.DataFrame) -> List[str]:
        errors = ValidationService.validate_schema(df, self.columns)
        if errors:
            return errors

        if not all(pd.api.types.is_string_dtype(df[column]) for column in self.columns):
            errors.append("Todos os campos devem ser texto.")

        invalid_movements = sorted(
            {
                str(value)
                for value in df["Mov."]
                if not isinstance(value, str) or value not in self._EXPECTED_MOVEMENTS
            }
        )
        if invalid_movements:
            errors.append("Movimentos fora do conjunto esperado: " + ", ".join(invalid_movements))

        for index, row in df.iterrows():
            is_continuation = row["Certif."] == ""
            required_fields = (
                self._CONTINUATION_LEADING_FIELDS
                if is_continuation
                else self._REQUIRED_COMPLETE_ROWS
            )
            if is_continuation:
                if any(row[field] != "" for field in self._CONTINUATION_LEADING_FIELDS):
                    errors.append(f"Continuidade inválida na linha {index + 1}.")
            elif any(row[field] == "" for field in required_fields):
                errors.append(f"Registro completo incompleto na linha {index + 1}.")

            if any(row[field] == "" for field in self._REQUIRED_ALL_ROWS):
                errors.append(f"Campos de lançamento ausentes na linha {index + 1}.")

        return errors
