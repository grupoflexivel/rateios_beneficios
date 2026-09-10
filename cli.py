#!/usr/bin/env python3
"""Command-line entry point for all registered PDF layout converters."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Optional, Sequence

from rateiosrh.core.exceptions import (
    DataFrameValidationError,
    LayoutMismatchError,
    PdfStructureError,
    UnknownParserError,
)
from rateiosrh.services.conversion import ConversionService
from rateiosrh.services.excel_exporter import ExcelExporter


LOGGER = logging.getLogger("rateiosrh.cli")


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Converte um relatório PDF para Excel usando o layout selecionado."
    )
    parser.add_argument("converter_id", help="identificador do conversor registrado")
    parser.add_argument("pdf", type=Path, help="arquivo PDF de entrada")
    parser.add_argument("-o", "--output", type=Path, help="arquivo XLSX de saída")
    parser.add_argument("-v", "--verbose", action="store_true", help="habilita logs detalhados")
    return parser


def run(argv: Optional[Sequence[str]] = None) -> int:
    argument_parser = build_argument_parser()
    try:
        args = argument_parser.parse_args(argv)
    except SystemExit as exc:
        return int(exc.code)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if not args.pdf.is_file():
        LOGGER.error("Arquivo PDF não encontrado: %s", args.pdf)
        return 2

    output_path = args.output or args.pdf.with_suffix(".xlsx")
    try:
        result = ConversionService().convert(args.converter_id, args.pdf)
        LOGGER.info(
            "Conversor %s: %d páginas, %d linhas, %d observações.",
            result.parser_id,
            result.page_count,
            len(result.dataframe),
            len(result.observations),
        )
        for message in result.observations:
            LOGGER.warning(message)
        ExcelExporter.export(result.dataframe, output_path, result.observations)
        LOGGER.info("Excel gerado: %s", output_path)
        return 0
    except UnknownParserError as exc:
        LOGGER.error("Falha na conversão: %s", exc)
        return 2
    except (DataFrameValidationError, LayoutMismatchError, FileNotFoundError,
            PdfStructureError, OSError, ValueError) as exc:
        LOGGER.error("Falha na conversão: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(run())
