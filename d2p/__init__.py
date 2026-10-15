"""d2p -- a self-contained DOCX -> PDF conversion toolkit for Linux.

The package is intentionally dependency-light:

* ``python-docx``  reads the WordprocessingML package
* ``reportlab``    lays out and writes the PDF
* ``fontTools``    (optional) inspects system fonts for glyph coverage
* ``Pillow``       (optional) reads intrinsic image sizes
* ``PyMuPDF``      (optional, tests only) verifies the produced PDFs

Public entry point is :func:`d2p.convert.convert_file`.
"""

__version__ = "0.1.0"

from .convert import (ConversionError, ConversionResult, available_engines, convert_file,
                      convert_many)  # noqa: E402  (import after __version__)

__all__ = [
    "__version__",
    "convert_file",
    "convert_many",
    "available_engines",
    "ConversionResult",
    "ConversionError",
]
