from pathlib import Path

import pytest

from rateiosrh.core.pdf_reader import PdfDocument, group_rows


def test_application_packages_are_importable():
    import rateiosrh.core
    import rateiosrh.parsers
    import rateiosrh.services
    import rateiosrh.web

    assert rateiosrh.core is not None


def test_reader_finds_ten_pages_in_unimed_reference():
    document = PdfDocument(Path("ANALITICO_TAXA_0054041451-1.PDF").read_bytes())
    assert len(document.page_streams()) == 10


def test_reader_finds_seven_pages_in_bradesco_dental_reference():
    document = PdfDocument(Path("FLEXIV_1166046_1_082026.pdf").read_bytes())
    assert len(document.page_streams()) == 7


@pytest.mark.parametrize(
    ("pages_type", "page_type"),
    [
        (b"/Type /Pages", b"/Type /Page"),
        (b"/Type/Pages", b"/Type/Page"),
        (b"/Type /Pages", b"/Type/Page"),
        (b"/Type/Pages", b"/Type /Page"),
    ],
)
def test_reader_accepts_page_type_spacing_variants(pages_type, page_type):
    data = (
        b"1 0 obj\n<< "
        + pages_type
        + b" /Kids [2 0 R] >>\nendobj\n"
        + b"2 0 obj\n<< "
        + page_type
        + b" /Contents 3 0 R >>\nendobj\n"
        + b"3 0 obj\n<< /Length 5 >>\nstream\nBT ET\nendstream\nendobj\n"
    )

    assert PdfDocument(data).page_streams() == [b"BT ET"]


def test_group_rows_uses_requested_precision_without_pre_rounding_spans():
    stream = (
        b"1 0 0 1 0 10.001 Tm (first) Tj "
        b"1 0 0 1 0 10.004 Tm (second) Tj"
    )

    assert [y for y, _ in group_rows(stream, 3)] == [10.004, 10.001]
