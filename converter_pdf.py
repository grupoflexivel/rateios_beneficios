#!/usr/bin/env python3
"""Convert analytical health-benefit PDFs with the reference layout to XLSX.

The reference PDF is generated from a native text layer.  Its text is drawn
using fixed page coordinates, so this converter intentionally reads the text
streams and reconstructs table rows from coordinates instead of using OCR or
whitespace splitting.
"""

from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from rateiosrh.core.exceptions import PdfStructureError
from rateiosrh.core.pdf_reader import PdfDocument, TextSpan, clean, group_rows


LOGGER = logging.getLogger("converter_pdf")


@dataclass(frozen=True)
class LayoutConfig:
    """Coordinates and patterns for the known report layout."""

    page_type: bytes = b"/Type/Page"
    filial_value: str = "663"
    coordinate_tolerance: float = 1.0
    row_precision: int = 2
    field_x: Dict[str, float] = field(
        default_factory=lambda: {
            "codigo": 20.25,
            "beneficiario": 100.50,
            "operacao": 197.25,
            "lancamento": 204.75,
            "filial": 278.25,
            "centrocusto": 318.75,
            "quantidade": 350.25,
            "valor_unit": 366.00,
            "valor_bonif": 408.00,
            "valor_total": 450.00,
            "mensagem": 516.00,
        }
    )
    beneficiary_code_pattern: re.Pattern[str] = re.compile(
        r"(?P<code>\d{4}\.\d{4}\.\d{6}-\d{2})"
    )
    titular_pattern: re.Pattern[str] = re.compile(
        r"^(?P<name>.+?)\s+-\s+(?P<code>\d{4}\.\d{4}\.\d{6})$"
    )
    total_pattern: re.Pattern[str] = re.compile(
        r"TOTAL TITULAR\s*:\s*(?P<value>[+-]?[\d.,]+)"
    )


DEFAULT_LAYOUT = LayoutConfig()


@dataclass(frozen=True)
class TitularTotal:
    page: int
    titular_code: str
    titular_name: str
    value: str


@dataclass(frozen=True)
class Record:
    codigo_beneficiario: str
    beneficiario: str
    op_lancamento: str
    filial: str
    centro_custo: str
    quantidade: str
    valor_unit: str
    valor_bonif: str
    valor_total: str
    mensagem: str
    page: int
    y: float
    titular_code: str
    titular_name: str
    source_beneficiary_repeated: bool = True


@dataclass
class ExtractionResult:
    pdf_path: Path
    page_count: int
    records: List[Record]
    titular_totals: List[TitularTotal]
    observations: List[str]
    ignored_lines: int
    page_record_counts: List[int]


def _field_text(items: Sequence[Tuple[float, str]], target: float, tolerance: float) -> str:
    values = [
        text
        for x, text in items
        if abs(x - target) <= tolerance and clean(text)
    ]
    return values[-1] if values else ""


def _compose_operation(operation: str, launch: str) -> str:
    """Combine layout columns without changing whitespace supplied by either span."""

    if not operation:
        return launch
    if not launch:
        return operation
    if operation[-1].isspace() or launch[0].isspace():
        return operation + launch
    return operation + " " + launch


def split_code_and_name(text: str) -> Tuple[str, str]:
    """Split a beneficiary field using its structural identifier pattern."""

    normalized = clean(text)
    match = DEFAULT_LAYOUT.beneficiary_code_pattern.match(normalized)
    if not match:
        return "", normalized
    code = match.group("code")
    name = clean(normalized[match.end() :])
    return code, name


def _parse_decimal(value: str) -> Optional[Decimal]:
    if not value:
        return None
    try:
        return Decimal(value.replace(".", "").replace(",", "."))
    except InvalidOperation:
        return None


def _title_parts(text: str, layout: LayoutConfig) -> Optional[Tuple[str, str]]:
    match = layout.titular_pattern.match(text)
    if not match:
        return None
    return match.group("code"), clean(match.group("name"))


def _record_from_row(
    page: int,
    y: float,
    items: Sequence[Tuple[float, str]],
    current_titular: Tuple[str, str],
    layout: LayoutConfig,
) -> Tuple[Record, Optional[str]]:
    code = _field_text(items, layout.field_x["codigo"], layout.coordinate_tolerance)
    name = _field_text(items, layout.field_x["beneficiario"], layout.coordinate_tolerance)
    operation = _field_text(items, layout.field_x["operacao"], layout.coordinate_tolerance)
    launch = _field_text(items, layout.field_x["lancamento"], layout.coordinate_tolerance)
    filial = _field_text(items, layout.field_x["filial"], layout.coordinate_tolerance)
    centro_custo = _field_text(items, layout.field_x["centrocusto"], layout.coordinate_tolerance)
    quantity = _field_text(items, layout.field_x["quantidade"], layout.coordinate_tolerance)
    unit_value = _field_text(items, layout.field_x["valor_unit"], layout.coordinate_tolerance)
    bonus_value = _field_text(items, layout.field_x["valor_bonif"], layout.coordinate_tolerance)
    total_value = _field_text(items, layout.field_x["valor_total"], layout.coordinate_tolerance)
    message = _field_text(items, layout.field_x["mensagem"], layout.coordinate_tolerance)

    source_beneficiary_repeated = bool(code or name)
    observation = None
    if not source_beneficiary_repeated:
        code, name = current_titular
        observation = (
            f"Página {page}, y={y:.2f}: o beneficiário não foi repetido nesta linha; "
            f"associação mantida com o titular corrente ({name} - {code})."
        )

    return (
        Record(
            codigo_beneficiario=code,
            beneficiario=name,
            op_lancamento=_compose_operation(operation, launch),
            filial=filial,
            centro_custo=centro_custo,
            quantidade=quantity,
            valor_unit=unit_value,
            valor_bonif=bonus_value,
            valor_total=total_value,
            mensagem=message,
            page=page,
            y=y,
            titular_code=current_titular[0],
            titular_name=current_titular[1],
            source_beneficiary_repeated=source_beneficiary_repeated,
        ),
        observation,
    )


def extract_records(pdf_path: Path, layout: LayoutConfig = DEFAULT_LAYOUT) -> ExtractionResult:
    """Extract every launch from every page of a compatible PDF."""

    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"Arquivo PDF não encontrado: {pdf_path}")
    document = PdfDocument(pdf_path.read_bytes())
    streams = document.page_streams()
    records: List[Record] = []
    titular_totals: List[TitularTotal] = []
    observations: List[str] = []
    ignored_lines = 0
    page_record_counts: List[int] = []
    current_titular = ("", "")
    current_group_records: List[Record] = []

    for page_number, stream in enumerate(streams, start=1):
        page_start = len(records)
        for y, items in group_rows(stream, layout.row_precision):
            line = clean("".join(text for _, text in items))
            if not line:
                continue

            title = _title_parts(line, layout)
            if title:
                current_titular = title
                current_group_records = []
                continue

            total_match = layout.total_pattern.search(line)
            if total_match:
                value = total_match.group("value")
                titular_totals.append(
                    TitularTotal(page_number, current_titular[0], current_titular[1], value)
                )
                expected = _parse_decimal(value)
                actual = sum(
                    (_parse_decimal(record.valor_total) or Decimal("0"))
                    for record in current_group_records
                )
                if expected is not None and actual != expected:
                    observations.append(
                        f"Página {page_number}, y={y:.2f}: TOTAL TITULAR {value} diverge da soma "
                        f"dos lançamentos ({actual})."
                    )
                current_group_records = []
                continue

            filial = _field_text(items, layout.field_x["filial"], layout.coordinate_tolerance)
            if filial != layout.filial_value:
                ignored_lines += 1
                continue
            if not current_titular[0]:
                observations.append(
                    f"Página {page_number}, y={y:.2f}: lançamento encontrado sem titular corrente."
                )
            record, observation = _record_from_row(
                page_number, y, items, current_titular, layout
            )
            records.append(record)
            current_group_records.append(record)
            if observation:
                observations.append(observation)
        page_record_counts.append(len(records) - page_start)

    if current_group_records:
        observations.append("Há lançamentos sem TOTAL TITULAR posterior no documento.")

    result = ExtractionResult(
        pdf_path=pdf_path,
        page_count=len(streams),
        records=records,
        titular_totals=titular_totals,
        observations=observations,
        ignored_lines=ignored_lines,
        page_record_counts=page_record_counts,
    )
    LOGGER.info(
        "PDF %s: %d páginas, %d lançamentos, %d totais de titular, %d linhas ignoradas.",
        pdf_path,
        result.page_count,
        len(result.records),
        len(result.titular_totals),
        result.ignored_lines,
    )
    for page, count in enumerate(result.page_record_counts, start=1):
        LOGGER.debug("Página %d: %d lançamentos.", page, count)
    return result


def validate_records(result: ExtractionResult) -> List[str]:
    """Return validation observations without changing extracted values."""

    observations = list(result.observations)
    for index, record in enumerate(result.records, start=1):
        required = {
            "código": record.codigo_beneficiario,
            "beneficiário": record.beneficiario,
            "operação/lançamento": record.op_lancamento,
            "filial": record.filial,
            "centrocusto": record.centro_custo,
            "quantidade": record.quantidade,
            "valor unitário": record.valor_unit,
            "valor bonificação": record.valor_bonif,
            "valor total": record.valor_total,
        }
        for label, value in required.items():
            if not value:
                observations.append(f"Lançamento {index}: campo obrigatório vazio: {label}.")
    return observations


HEADERS = [
    "CÓDIGO BENEFICIÁRIO",
    "BENEFICIÁRIO",
    "OP. LANÇAMENTO",
    "FILIAL",
    "CENTROCUSTO",
    "QUANT.",
    "VALOR UNIT.",
    "VALOR BONIF.",
    "VALOR TOTAL",
    "MENSAGEM",
]


def _style_sheet(sheet, widths: Sequence[int]) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(widths))}{sheet.max_row}"
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows():
        for cell in row:
            cell.number_format = "@"


def write_workbook(result: ExtractionResult, output_path: Path) -> Path:
    """Write the main launch table and a separate validation-observation sheet."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lançamentos"
    sheet.append(HEADERS)
    for record in result.records:
        sheet.append(
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
        )
    _style_sheet(sheet, [25, 38, 24, 10, 14, 13, 15, 16, 15, 28])

    observations = workbook.create_sheet("Observações")
    observations.append(["TIPO", "DETALHE"])
    for message in validate_records(result):
        observations.append(["VALIDAÇÃO", message])
    observations.column_dimensions["A"].width = 16
    observations.column_dimensions["B"].width = 120
    observations.freeze_panes = "A2"
    observations.auto_filter.ref = f"A1:B{observations.max_row}"
    for cell in observations[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
    for row in observations.iter_rows():
        for cell in row:
            cell.number_format = "@"

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    LOGGER.info("Excel gerado: %s", output_path)
    return output_path


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
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if not args.pdf.is_file():
        LOGGER.error("Arquivo PDF não encontrado: %s", args.pdf)
        return 2
    output = args.output or args.pdf.with_suffix(".xlsx")
    try:
        result = extract_records(args.pdf)
        observations = validate_records(result)
        if observations:
            LOGGER.warning("Validações com observações: %d", len(observations))
            for message in observations:
                LOGGER.warning(message)
        unique_beneficiaries = {
            record.codigo_beneficiario
            for record in result.records
            if record.codigo_beneficiario and record.source_beneficiary_repeated
        }
        LOGGER.info("Beneficiários/dependentes identificados: %d", len(unique_beneficiaries))
        write_workbook(result, output)
        return 0
    except (FileNotFoundError, PdfStructureError, OSError, ValueError) as exc:
        LOGGER.error("Falha na conversão: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(run())
