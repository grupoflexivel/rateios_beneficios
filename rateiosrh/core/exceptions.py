"""Exceptions raised by shared PDF infrastructure."""


class PdfStructureError(RuntimeError):
    """Raised when a PDF has no usable objects, pages, contents, or streams."""
