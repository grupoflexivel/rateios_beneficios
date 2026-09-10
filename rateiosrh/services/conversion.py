"""Application service that orchestrates parser selection and conversion."""

from dataclasses import dataclass
from pathlib import Path
from typing import List

import pandas as pd

from rateiosrh.core.exceptions import (
    DataFrameValidationError,
    LayoutMismatchError,
)
from rateiosrh.core.pdf_reader import PdfDocument
from rateiosrh.registry import get_parser


@dataclass
class ConversionOutput:
    """Result of a successful parser conversion."""

    dataframe: pd.DataFrame
    parser_id: str
    observations: List[str]
    page_count: int = 0


class ConversionService:
    """Resolve a parser, validate the layout, parse, and validate its output."""

    def convert(self, parser_id: str, pdf_path: Path) -> ConversionOutput:
        path = Path(pdf_path)
        parser = get_parser(parser_id)
        if not parser.validate_layout(path):
            raise LayoutMismatchError(parser_id, path)

        frame = parser.parse(path)
        errors = parser.validate_dataframe(frame)
        if errors:
            raise DataFrameValidationError(errors)
        page_count = len(PdfDocument(path.read_bytes()).page_streams())

        return ConversionOutput(
            dataframe=frame,
            parser_id=parser.parser_id,
            observations=parser.get_observations(),
            page_count=page_count,
        )
