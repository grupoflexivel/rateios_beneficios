"""Exceptions shared by the PDF conversion layers."""

from pathlib import Path
from typing import Iterable, Union


class PdfStructureError(RuntimeError):
    """Raised when a PDF has no usable objects, pages, contents, or streams."""


class LayoutMismatchError(RuntimeError):
    """Raised when a PDF is incompatible with the selected parser."""

    def __init__(self, parser_id: str, pdf_path: Union[Path, str]):
        self.parser_id = parser_id
        self.pdf_path = Path(pdf_path)
        super().__init__(
            f"PDF incompatível com o parser '{parser_id}': {self.pdf_path}"
        )


class UnknownParserError(RuntimeError):
    """Raised when a parser identifier is not registered."""

    def __init__(self, parser_id: str):
        self.parser_id = parser_id
        super().__init__(f"Parser desconhecido: {parser_id}")


class DataFrameValidationError(RuntimeError):
    """Raised when a parsed DataFrame violates its declared contract."""

    def __init__(self, errors: Iterable[str]):
        self.errors = list(errors)
        super().__init__(" ".join(self.errors))
