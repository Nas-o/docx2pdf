"""High level conversion API (used by the CLI, usable as a library).

    from d2p import convert_file

    result = convert_file("report.docx", "report.pdf")
    if result.ok:
        print(result.pages, "pages via", result.engine)
    else:
        print(result.error)
"""

from __future__ import annotations

import os
from typing import Iterable, List, Optional, Sequence

from .engines import (DEFAULT_TIMEOUT, ConversionError, ConversionResult, LibreOfficeEngine,
                      engines_status, select_engine)
from .fonts import get_font_book

__all__ = ["convert_file", "convert_many", "ConversionResult", "ConversionError",
           "available_engines"]


def convert_file(source: str, target: Optional[str] = None, *, engine: str = "auto",
                 font_dirs: Optional[Sequence[str]] = None,
                 timeout: float = DEFAULT_TIMEOUT, overwrite: bool = True) -> ConversionResult:
    """Convert *source* to PDF and return a :class:`ConversionResult`.

    ``target`` defaults to the source path with a ``.pdf`` suffix.  Errors are
    reported in the result (``ok=False``, ``error=...``) instead of raising, so
    batches can continue; ``ConversionError`` is only raised for bad arguments.
    """
    if not source:
        raise ConversionError("source 不能为空 / source must not be empty")
    target = target or os.path.splitext(source)[0] + ".pdf"
    result = ConversionResult(source=source, target=target)
    if not overwrite and os.path.exists(target):
        result.ok = True
        result.skipped = True
        result.engine = "-"
        return result
    try:
        engine_instance = select_engine(engine, os.path.splitext(source)[1], font_dirs)
        result.engine = engine_instance.name
        converted = engine_instance.convert(source, target, timeout=timeout)
        result.ok = True
        result.pages = converted.pages
        result.warnings = converted.warnings
    except ConversionError as exc:
        result.error = str(exc)
    except Exception as exc:  # pragma: no cover - defensive
        result.error = f"{type(exc).__name__}: {exc}"
    return result


def convert_many(sources: Iterable[str], outdir: Optional[str] = None, *,
                 engine: str = "auto", font_dirs: Optional[Sequence[str]] = None,
                 timeout: float = DEFAULT_TIMEOUT, overwrite: bool = True,
                 progress=None) -> List[ConversionResult]:
    """Convert several files serially (the CLI adds real parallelism)."""
    sources = list(sources)
    results: List[ConversionResult] = []
    for index, source in enumerate(sources, start=1):
        target = None
        if outdir:
            os.makedirs(outdir, exist_ok=True)
            target = os.path.join(outdir, os.path.splitext(os.path.basename(source))[0] + ".pdf")
        result = convert_file(source, target, engine=engine, font_dirs=font_dirs,
                             timeout=timeout, overwrite=overwrite)
        results.append(result)
        if progress is not None:
            progress(index, len(sources), result)
    return results


def available_engines(font_dirs: Optional[Sequence[str]] = None) -> List[str]:
    """Names of engines usable on this machine (``builtin`` / ``libreoffice``)."""
    return [status["name"] for status in engines_status(font_dirs) if status["available"]]


def warm_up(font_dirs: Optional[Sequence[str]] = None) -> None:
    """Pre-build the font index (useful before converting many documents)."""
    get_font_book(font_dirs)
