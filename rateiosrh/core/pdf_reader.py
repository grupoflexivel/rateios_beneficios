"""Small, layout-agnostic reader for text streams in simple PDF documents."""

from __future__ import annotations

import re
import zlib
from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple

from .exceptions import PdfStructureError


NUMBER_TOKEN = re.compile(
    rb"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$"
)
INLINE_IMAGE_RE = re.compile(rb"(?<!\S)BI(?=\s)")
INLINE_IMAGE_ID_RE = re.compile(rb"(?<!\S)ID(?=\s)")
INLINE_IMAGE_END_RE = re.compile(rb"(?<!\S)EI(?=\s)")
PDF_WHITESPACE = b" \t\r\n\f\0"


@dataclass(frozen=True)
class TextSpan:
    """Text drawn at a position in a PDF content stream."""

    y: float
    x: float
    text: str


class PdfDocument:
    """Index PDF objects and expose page content streams without layout rules."""

    def __init__(self, data: bytes):
        self.data = data
        self.objects = self._index_objects(data)

    @staticmethod
    def _index_objects(data: bytes) -> Dict[int, Tuple[int, int]]:
        matches = list(re.finditer(rb"(?m)^(\d+)\s+(\d+)\s+obj\s*", data))
        objects: Dict[int, Tuple[int, int]] = {}
        for index, match in enumerate(matches):
            end = matches[index + 1].start() if index + 1 < len(matches) else len(data)
            objects[int(match.group(1))] = (match.end(), end)
        if not objects:
            raise PdfStructureError("Nenhum objeto PDF foi encontrado.")
        return objects

    def object_body(self, object_number: int) -> bytes:
        try:
            start, end = self.objects[object_number]
        except KeyError as exc:
            raise PdfStructureError(f"Objeto PDF ausente: {object_number} 0 R") from exc
        body = self.data[start:end]
        endobj = body.rfind(b"endobj")
        return body[:endobj] if endobj >= 0 else body

    def stream(self, object_number: int) -> bytes:
        body = self.object_body(object_number)
        marker = body.find(b"stream")
        if marker < 0:
            raise PdfStructureError(f"Objeto {object_number} não contém stream.")
        stream_start = marker + len(b"stream")
        while body[stream_start : stream_start + 1] in b"\r\n":
            stream_start += 1

        dictionary = body[:marker]
        indirect_length = re.search(rb"/Length\s+(\d+)\s+\d+\s+R", dictionary)
        if indirect_length:
            length_body = self.object_body(int(indirect_length.group(1)))
            length_match = re.search(rb"\b(\d+)\b", length_body)
            if not length_match:
                raise PdfStructureError(f"Comprimento inválido no objeto {object_number}.")
            length = int(length_match.group(1))
        else:
            direct_length = re.search(rb"/Length\s+(\d+)\b", dictionary)
            if not direct_length:
                raise PdfStructureError(f"Comprimento ausente no stream {object_number}.")
            length = int(direct_length.group(1))

        payload = body[stream_start : stream_start + length]
        if b"/FlateDecode" in dictionary:
            try:
                return zlib.decompress(payload)
            except zlib.error as exc:
                raise PdfStructureError(f"Falha ao descomprimir stream {object_number}.") from exc
        return payload

    def page_streams(self) -> List[bytes]:
        pages_object = None
        for object_number in self.objects:
            body = self.object_body(object_number)
            if re.search(rb"/Type\s*/Pages", body) and b"/Kids" in body:
                pages_object = body
                break
        if pages_object is None:
            raise PdfStructureError("Objeto /Pages não encontrado.")

        kids_match = re.search(rb"/Kids\s*\[(.*?)\]", pages_object, re.S)
        if not kids_match:
            raise PdfStructureError("Lista de páginas ausente.")
        page_numbers = [
            int(value) for value in re.findall(rb"(\d+)\s+\d+\s+R", kids_match.group(1))
        ]
        if not page_numbers:
            raise PdfStructureError("Lista de páginas vazia.")

        streams: List[bytes] = []
        for page_number in page_numbers:
            body = self.object_body(page_number)
            if not re.search(rb"/Type\s*/Page\b", body):
                raise PdfStructureError(f"Objeto {page_number} não é uma página PDF.")
            contents_match = re.search(rb"/Contents\s+(\d+)\s+\d+\s+R", body)
            if not contents_match:
                raise PdfStructureError(f"Conteúdo ausente na página {page_number}.")
            streams.append(self.stream(int(contents_match.group(1))))
        if not streams:
            raise PdfStructureError("Nenhum stream de página foi encontrado.")
        return streams


def remove_inline_images(stream: bytes) -> bytes:
    """Remove inline image payloads before tokenizing PDF text operators."""

    chunks: List[bytes] = []
    cursor = 0
    while True:
        image_match = INLINE_IMAGE_RE.search(stream, cursor)
        if image_match is None:
            chunks.append(stream[cursor:])
            break

        image_start = image_match.start()
        id_match = INLINE_IMAGE_ID_RE.search(stream, image_match.end())
        if id_match is None:
            raise PdfStructureError("Imagem inline sem operador ID no stream PDF.")

        data_start = id_match.end()
        while stream[data_start : data_start + 1] in PDF_WHITESPACE:
            data_start += 1
        dictionary = stream[image_match.end() : id_match.start()]
        image_end = None

        for end_match in INLINE_IMAGE_END_RE.finditer(stream, data_start):
            candidate = end_match.start()
            payload = stream[data_start:candidate].rstrip(PDF_WHITESPACE)
            if b"/FlateDecode" not in dictionary:
                image_end = candidate
                break
            try:
                decompressor = zlib.decompressobj()
                decompressor.decompress(payload)
            except zlib.error:
                continue
            if decompressor.eof and not decompressor.unused_data:
                image_end = candidate
                break

        if image_end is None:
            raise PdfStructureError("Imagem inline sem terminador no stream PDF.")

        chunks.append(stream[cursor:image_start])
        cursor = image_end + len(b"EI")
    return b"".join(chunks)


def _tokens(data: bytes) -> Iterable[Tuple[str, object]]:
    index = 0
    while index < len(data):
        while index < len(data) and data[index] in b" \t\r\n\f\0":
            index += 1
        if index >= len(data):
            return
        if data[index] == ord("%"):
            while index < len(data) and data[index] not in b"\r\n":
                index += 1
            continue
        if data[index] == ord("("):
            index += 1
            depth = 1
            value = bytearray()
            while index < len(data) and depth:
                char = data[index]
                if char == ord("\\"):
                    if index + 1 >= len(data):
                        break
                    escaped = data[index + 1]
                    escapes = {ord("n"): 10, ord("r"): 13, ord("t"): 9, ord("b"): 8, ord("f"): 12}
                    if escaped in escapes:
                        value.append(escapes[escaped])
                        index += 2
                    elif escaped in b"()\\":
                        value.append(escaped)
                        index += 2
                    elif 48 <= escaped <= 55:
                        end = index + 1
                        while end < len(data) and end < index + 4 and 48 <= data[end] <= 55:
                            end += 1
                        value.append(int(data[index + 1 : end], 8))
                        index = end
                    else:
                        value.append(escaped)
                        index += 2
                elif char == ord("("):
                    depth += 1
                    value.append(char)
                    index += 1
                elif char == ord(")"):
                    depth -= 1
                    index += 1
                    if depth:
                        value.append(char)
                else:
                    value.append(char)
                    index += 1
            yield "str", bytes(value)
            continue
        if data[index] in b"[]":
            yield chr(data[index]), chr(data[index])
            index += 1
            continue
        end = index
        while end < len(data) and data[end] not in b" \t\r\n\f\0[]()":
            end += 1
        yield "tok", data[index:end]
        index = end


def _text_spans(stream: bytes) -> List[TextSpan]:
    operands: List[Tuple[str, object]] = []
    spans: List[TextSpan] = []
    x = 0.0
    y = 0.0
    leading = 0.0

    for kind, value in _tokens(remove_inline_images(stream)):
        if kind in ("str", "[", "]") or (kind == "tok" and NUMBER_TOKEN.match(value)):  # type: ignore[arg-type]
            operands.append((kind, value))
            continue

        operator = value.decode("latin1")  # type: ignore[union-attr]
        if operator == "Tm" and len(operands) >= 6:
            x = float(operands[-2][1])
            y = float(operands[-1][1])
        elif operator == "Td" and len(operands) >= 2:
            x += float(operands[-2][1])
            y += float(operands[-1][1])
        elif operator == "TL" and operands:
            leading = float(operands[-1][1])
        elif operator == "T*":
            y -= leading
        elif operator in ("Tj", "TJ"):
            strings = [item for item_kind, item in operands if item_kind == "str"]
            if strings:
                text = b"".join(strings).decode("cp1252", "replace")
                spans.append(TextSpan(y, x, text))
        operands = []
    return spans


def group_rows(
    stream: bytes, row_precision: int = 2
) -> List[Tuple[float, List[Tuple[float, str]]]]:
    """Group text spans by y coordinate while retaining original text values."""

    rows: Dict[float, List[Tuple[float, str]]] = {}
    for span in _text_spans(stream):
        row_key = round(span.y, row_precision)
        rows.setdefault(row_key, []).append((span.x, span.text))
    return [(y, sorted(items)) for y, items in sorted(rows.items(), reverse=True)]


def clean(value: str) -> str:
    """Normalize whitespace only when rebuilding text from positioned spans."""

    return re.sub(r"\s+", " ", value).strip()
