#!/usr/bin/env python3
"""Legacy UNIMED converter facade.

The parsing implementation lives in :mod:`rateiosrh.parsers.unimed`.  This
module keeps the historical CLI and public functions available to callers.
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd

from rateiosrh.core.exceptions import (
    DataFrameValidationError,
    LayoutMismatchError,
    PdfStructureError,
)
from rateiosrh.core.pdf_reader import PdfDocument, TextSpan, clean, group_rows
from rateiosrh.parsers.unimed import (
    DEFAULT_LAYOUT,
    HEADERS,
    ExtractionResult,
    LayoutConfig,
    Record,
    TitularTotal,
    UnimedParser,
    _compose_operation,
    _field_text,
    _parse_decimal,
    _record_from_row,
    _title_parts,
    extract_records,
    split_code_and_name,
    validate_records,
)
from rateiosrh.services.conversion import ConversionService
from rateiosrh.services.excel_exporter import ExcelExporter


LOGGER = logging.getLogger("converter_pdf")


def write_workbook(result: ExtractionResult, output_path: Path) -> Path:
    """Write a legacy extraction result through the generic Excel exporter."""

    frame = pd.DataFrame(
        [
            [
                record.codigo_beneficiario,
                record.beneficiario,
                record.op_lancamento,
                record.filial,
                record.centro_custo,
                record.quantidade,
                record.valor_unit,
                record.valor_bonif,
                record.valor_total,
                record.mensagem,
            ]
            for record in result.records
        ],
        columns=HEADERS,
    ).astype("string")
    return ExcelExporter.export(frame, output_path, validate_records(result))


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Converte um relatório analítico PDF do layout de referência para Excel."
    )
    parser.add_argument("pdf", type=Path, help="arquivo PDF de entrada")
    parser.add_argument("-o", "--output", type=Path, help="arquivo XLSX de saída")
    parser.add_argument("-v", "--verbose", action="store_true", help="habilita logs detalhados")
    return parser


def run(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_argument_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return int(exc.code)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if not args.pdf.is_file():
        LOGGER.error("Arquivo PDF não encontrado: %s", args.pdf)
        return 2
    output = args.output or args.pdf.with_suffix(".xlsx")
    try:
        conversion = ConversionService().convert("unimed", args.pdf)
        if conversion.observations:
            LOGGER.warning(
                "Validações com observações: %d", len(conversion.observations)
            )
            for message in conversion.observations:
                LOGGER.warning(message)
        ExcelExporter.export(conversion.dataframe, output, conversion.observations)
        LOGGER.info("Excel gerado: %s", output)
        return 0
    except (
        DataFrameValidationError,
        FileNotFoundError,
        LayoutMismatchError,
        PdfStructureError,
        OSError,
        ValueError,
    ) as exc:
        LOGGER.error("Falha na conversão: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(run())
