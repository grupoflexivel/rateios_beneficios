"""Abstract contract shared by layout-specific PDF parsers."""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import ClassVar, List

import pandas as pd

from rateiosrh.services.validation import ValidationService


class BaseParser(ABC):
    """Define the common interface for a PDF layout parser."""

    parser_id: ClassVar[str]
    display_name: ClassVar[str]
    columns: ClassVar[List[str]]

    @abstractmethod
    def validate_layout(self, pdf_path: Path) -> bool:
        """Return whether ``pdf_path`` matches this parser's layout."""

    @abstractmethod
    def parse(self, pdf_path: Path) -> pd.DataFrame:
        """Parse ``pdf_path`` into a DataFrame with the declared columns."""

    def validate_dataframe(self, df: pd.DataFrame) -> List[str]:
        """Return schema errors without changing ``df``."""

        return ValidationService.validate_schema(df, self.columns)

    def get_columns(self) -> List[str]:
        """Return a copy of the parser's declared column order."""

        return list(self.columns)
