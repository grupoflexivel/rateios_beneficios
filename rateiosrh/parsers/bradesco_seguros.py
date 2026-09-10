"""Parser for the Bradesco Vida e Previdência insured-people report."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import pandas as pd

from rateiosrh.core.exceptions import LayoutMismatchError, PdfStructureError
from rateiosrh.core.pdf_reader import PdfDocument, group_rows, remove_inline_images
from rateiosrh.parsers.base import BaseParser
from rateiosrh.services.validation import ValidationService


Span = Tuple[float, str]
Row = Tuple[float, List[Span]]


class BradescoSegurosParser(BaseParser):
    """Extract one logical insured record per PDF registration."""

    parser_id = "bradesco_seguros"
    display_name = "Converter Bradesco Seguros"
    columns = [
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

    _FIELD_RANGES = (
        ("Matrícula", 30.0, 70.0),
        ("Certificado", 70.0, 130.0),
        ("Segurado", 130.0, 300.0),
        ("Tipo Segurado", 300.0, 350.0),
        ("Data Nascto.", 350.0, 425.0),
        ("Data Inclusão", 425.0, 490.0),
        ("Início Vigência", 490.0, 560.0),
        ("Fim Vigência", 560.0, 650.0),
        ("Prazo", 650.0, 690.0),
        ("Capital", 690.0, 780.0),
        ("Prêmio", 780.0, 820.0),
    )
    _DOCUMENT_MARKERS = (
        "Relação dos Segurados Vigentes",
        "686 - BRADESCO VIDA E PREVIDENCIA",
    )
    _HEADER_MARKERS = tuple(
        column for column, _, _ in _FIELD_RANGES if column not in {"Segurado", "Tipo Segurado"}
    ) + ("Segurado", "Tipo Segurado")
    _DATA_Y_RANGE = (20.0, 375.0)

    def __init__(self) -> None:
        self._observations: List[str] = []

    @staticmethod
    def _remove_inline_images(stream: bytes) -> bytes:
        """Keep a parser-local alias for the common inline-image handling."""

        return remove_inline_images(stream)

    @classmethod
    def _page_rows(cls, stream: bytes) -> List[Row]:
        return group_rows(stream)

    @staticmethod
    def _page_text(rows: Iterable[Row]) -> str:
        return " ".join(text for _, spans in rows for _, text in spans)

    @classmethod
    def _is_detail_page(cls, rows: Sequence[Row]) -> bool:
        page_text = cls._page_text(rows)
        return all(marker in page_text for marker in cls._DOCUMENT_MARKERS) and all(
            marker in page_text for marker in cls._HEADER_MARKERS
        )

    @classmethod
    def _document_matches(cls, document: PdfDocument) -> bool:
        matched_page = False
        for stream in document.page_streams():
            rows = cls._page_rows(stream)
            page_text = cls._page_text(rows)
            identifiable = any(marker in page_text for marker in cls._DOCUMENT_MARKERS)
            if identifiable and not cls._is_detail_page(rows):
                return False
            matched_page = matched_page or cls._is_detail_page(rows)
        return matched_page

    def validate_layout(self, pdf_path: Path) -> bool:
        try:
            document = PdfDocument(Path(pdf_path).read_bytes())
            return self._document_matches(document)
        except (OSError, PdfStructureError):
            return False

    @classmethod
    def _field_text(cls, spans: Sequence[Span], lower: float, upper: float) -> str:
        return "".join(text for x, text in spans if lower <= x < upper).strip()

    @classmethod
    def _new_record(cls, spans: Sequence[Span]) -> Dict[str, str]:
        return {
            column: cls._field_text(spans, lower, upper)
            for column, lower, upper in cls._FIELD_RANGES
        }

    @classmethod
    def _extract_records(
        cls, document: PdfDocument
    ) -> Tuple[List[List[str]], int, int]:
        records: List[Dict[str, str]] = []
        name_continuations = 0
        matrícula_fragments = 0

        for stream in document.page_streams():
            rows = cls._page_rows(stream)
            if not cls._is_detail_page(rows):
                continue
            current: Dict[str, str] | None = None
            for y, spans in rows:
                if not cls._DATA_Y_RANGE[0] < y < cls._DATA_Y_RANGE[1]:
                    continue
                cleaned_spans = [(x, text.strip()) for x, text in spans if text.strip()]
                matrícula = next(
                    (
                        text
                        for x, text in cleaned_spans
                        if x < 60.0 and text.isdigit() and len(text) >= 8
                    ),
                    None,
                )
                if matrícula is not None:
                    current = cls._new_record(cleaned_spans)
                    records.append(current)
                    continue
                if current is None:
                    continue

                matrícula_fragment = next(
                    (
                        text
                        for x, text in cleaned_spans
                        if x < 60.0 and text.isdigit() and len(text) < 8
                    ),
                    None,
                )
                name_fragment = cls._field_text(cleaned_spans, 130.0, 300.0)
                if matrícula_fragment:
                    current["Matrícula"] += matrícula_fragment
                    matrícula_fragments += 1
                if name_fragment:
                    current["Segurado"] = (
                        f"{current['Segurado']} {name_fragment}".strip()
                    )
                    name_continuations += 1

        return (
            [[record[column] for column in cls.columns] for record in records],
            name_continuations,
            matrícula_fragments,
        )

    def parse(self, pdf_path: Path) -> pd.DataFrame:
        path = Path(pdf_path)
        document = PdfDocument(path.read_bytes())
        if not self._document_matches(document):
            raise LayoutMismatchError(self.parser_id, path)

        values, name_continuations, matrícula_fragments = self._extract_records(document)
        self._observations = [
            f"{name_continuations} continuidades de nome recompostas explicitamente.",
            f"{matrícula_fragments} fragmentos de matrícula recompostos explicitamente.",
        ]
        return pd.DataFrame(values, columns=self.columns).astype("string")

    def validate_dataframe(self, df: pd.DataFrame) -> List[str]:
        errors = ValidationService.validate_schema(df, self.columns)
        if errors:
            return errors
        if not all(pd.api.types.is_string_dtype(df[column]) for column in self.columns):
            errors.append("Todos os campos devem ser texto.")
        if df.empty:
            errors.append("Nenhum registro Bradesco Seguros foi extraído.")

        for index, row in df.iterrows():
            if not row["Matrícula"].isdigit() or len(row["Matrícula"]) != 11:
                errors.append(f"Matrícula inválida na linha {index + 1}.")
            if row["Certificado"] == "" or row["Segurado"] == "":
                errors.append(f"Registro incompleto na linha {index + 1}.")
            if row["Tipo Segurado"] != "P":
                errors.append(f"Tipo de segurado inválido na linha {index + 1}.")
        return errors

    def get_observations(self) -> List[str]:
        return list(self._observations)
