"""Validation shared by parser and conversion services."""

from typing import List, Sequence

import pandas as pd


class ValidationService:
    """Validate only the structural schema of a parsed DataFrame."""

    @staticmethod
    def validate_schema(
        df: pd.DataFrame, expected_columns: Sequence[str]
    ) -> List[str]:
        errors: List[str] = []
        if list(df.columns) != list(expected_columns):
            errors.append("Colunas ausentes ou fora de ordem.")
        return errors
