"""Parser for the validated UNIMED analytical tax report layout."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd

from rateiosrh.core.exceptions import PdfStructureError
from rateiosrh.core.pdf_reader import PdfDocument, clean, group_rows
from rateiosrh.parsers.base import BaseParser


LOGGER = logging.getLogger("rateiosrh.parsers.unimed")
LAYOUT_MARKER = "ANALÍTICO DE TAXA - FATURA NRO"
LAYOUT_HEADER_TOKENS = ("BENEFICIÁRIO", "LANÇAMENTO", "FILIAL", "CUSTOMENSAGEM")


@dataclass(frozen=True)
class LayoutConfig:
    """Coordinates and patterns for the known UNIMED report layout."""

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


_REQUIRED_FIELDS = (
    ("codigo_beneficiario", "CÓDIGO BENEFICIÁRIO", "código"),
    ("beneficiario", "BENEFICIÁRIO", "beneficiário"),
    ("op_lancamento", "OP. LANÇAMENTO", "operação/lançamento"),
    ("filial", "FILIAL", "filial"),
    ("centro_custo", "CENTROCUSTO", "centrocusto"),
    ("quantidade", "QUANT.", "quantidade"),
    ("valor_unit", "VALOR UNIT.", "valor unitário"),
    ("valor_bonif", "VALOR BONIF.", "valor bonificação"),
    ("valor_total", "VALOR TOTAL", "valor total"),
)


def _is_empty_required(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value
    return bool(pd.isna(value))


def _required_field_errors(index: int, value_getter) -> List[str]:
    errors: List[str] = []
    for attribute, _, label in _REQUIRED_FIELDS:
        if _is_empty_required(value_getter(attribute)):
            errors.append(f"Lançamento {index}: campo obrigatório vazio: {label}.")
    return errors


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
        observations.extend(
            _required_field_errors(index, lambda attribute: getattr(record, attribute))
        )
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


def _dataframe_value(value: str) -> str:
    """Remove only fixed-width padding before exposing the public table."""

    return value.strip()


class UnimedParser(BaseParser):
    """Parse the existing UNIMED layout without changing its extraction rules."""

    parser_id = "unimed"
    display_name = "Converter UNIMED"
    columns = HEADERS

    def __init__(self) -> None:
        self._observations: List[str] = []

    def get_observations(self) -> List[str]:
        """Expose non-fatal observations generated during the last parse."""

        return list(self._observations)

    def validate_dataframe(self, df: pd.DataFrame) -> List[str]:
        """Validate schema and the same required fields used by the legacy API."""

        errors = super().validate_dataframe(df)
        if errors:
            return errors

        column_indexes = {column: index for index, column in enumerate(self.columns)}
        for row_number, row in enumerate(df.itertuples(index=False, name=None), start=1):
            values = {
                attribute: row[column_indexes[column]]
                for attribute, column, _ in _REQUIRED_FIELDS
            }
            errors.extend(
                _required_field_errors(row_number, lambda attribute: values[attribute])
            )
        return errors

    def validate_layout(self, pdf_path: Path) -> bool:
        """Check the document marker and the UNIMED table header structure."""

        pdf_path = Path(pdf_path)
        if not pdf_path.is_file():
            return False
        try:
            streams = PdfDocument(pdf_path.read_bytes()).page_streams()
        except (OSError, PdfStructureError):
            return False

        marker_found = False
        header_found = False
        for stream in streams:
            for _, items in group_rows(stream, DEFAULT_LAYOUT.row_precision):
                line = clean("".join(text for _, text in items))
                marker_found = marker_found or LAYOUT_MARKER in line
                header_found = header_found or all(
                    token in line for token in LAYOUT_HEADER_TOKENS
                )
                if marker_found and header_found:
                    return True
        return False

    def parse(self, pdf_path: Path) -> pd.DataFrame:
        """Return the ten public UNIMED fields as textual columns."""

        result = extract_records(pdf_path)
        self._observations = list(result.observations)
        rows = [
            [
                _dataframe_value(record.codigo_beneficiario),
                _dataframe_value(record.beneficiario),
                _dataframe_value(record.op_lancamento),
                _dataframe_value(record.filial),
                _dataframe_value(record.centro_custo),
                _dataframe_value(record.quantidade),
                _dataframe_value(record.valor_unit),
                _dataframe_value(record.valor_bonif),
                _dataframe_value(record.valor_total),
                _dataframe_value(record.mensagem),
            ]
            for record in result.records
        ]
        return pd.DataFrame(rows, columns=self.columns).astype("string")
