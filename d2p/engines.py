"""Conversion engines.

Two engines are provided:

``libreoffice``
    Drives a headless LibreOffice (``soffice``) installation.  Highest fidelity
    and the only engine that can read legacy ``.doc``, ``.odt`` and ``.rtf``.
    Used automatically when a LibreOffice binary can be found.

``builtin``
    Pure Python (python-docx + ReportLab).  No office suite required, so it
    works on a bare machine, and is usually several times faster.  Reads
    ``.docx`` / ``.docm`` only.
"""

from __future__ import annotations

import glob
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from . import __version__

log = logging.getLogger("d2p.engines")

BUILTIN_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}
LIBREOFFICE_EXTENSIONS = BUILTIN_EXTENSIONS | {
    ".doc", ".dot", ".odt", ".ott", ".rtf", ".fodt", ".wps", ".txt", ".xml", ".uot",
}

DEFAULT_TIMEOUT = 180.0


class ConversionError(RuntimeError):
    """Raised when a document cannot be converted."""


@dataclass
class ConversionResult:
    source: str
    target: str = ""
    engine: str = ""
    ok: bool = False
    skipped: bool = False
    error: str = ""
    seconds: float = 0.0
    pages: int = 0
    warnings: List[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Built-in engine
# --------------------------------------------------------------------------- #


class BuiltinEngine:
    name = "builtin"
    extensions = BUILTIN_EXTENSIONS

    def __init__(self, font_dirs: Optional[Sequence[str]] = None) -> None:
        self.font_dirs = list(font_dirs) if font_dirs else None

    def available(self) -> bool:
        try:
            import docx  # noqa: F401
            import reportlab  # noqa: F401
        except Exception:
            return False
        return True

    def unavailable_reason(self) -> str:
        missing = []
        for module, package in (("docx", "python-docx"), ("reportlab", "reportlab")):
            try:
                __import__(module)
            except Exception:
                missing.append(package)
        if missing:
            return ("缺少 Python 依赖: " + ", ".join(missing) +
                    " —— 请运行: python3 -m pip install -r requirements.txt")
        return ""

    def supports(self, extension: str) -> bool:
        return extension.lower() in self.extensions

    def convert(self, source: str, target: str, timeout: float = DEFAULT_TIMEOUT,
                meta: Optional[Dict[str, Any]] = None) -> ConversionResult:
        from .docxmodel import read_docx
        from .renderer import PdfRenderer

        os.makedirs(os.path.dirname(os.path.abspath(target)) or ".", exist_ok=True)
        model = read_docx(source)
        renderer = PdfRenderer(font_dirs=self.font_dirs)
        stats = renderer.render(model, target, title=os.path.splitext(os.path.basename(source))[0])
        result = ConversionResult(source=source, target=target, engine=self.name, ok=True)
        result.pages = stats.pages
        result.warnings = list(model.warnings)
        return result


# --------------------------------------------------------------------------- #
# LibreOffice engine
# --------------------------------------------------------------------------- #


def find_soffice() -> Optional[str]:
    """Locate a usable ``soffice`` binary."""
    env = os.environ.get("D2P_SOFFICE")
    if env:
        if os.path.isfile(env) and os.access(env, os.X_OK):
            return env
        found = shutil.which(env)
        if found:
            return found
    for name in ("soffice", "libreoffice", "soffice.bin"):
        found = shutil.which(name)
        if found:
            return found
    tool_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    home = os.path.expanduser("~")
    patterns = [
        os.path.join(tool_dir, ".tools", "libreoffice*", "program", "soffice"),
        os.path.join(tool_dir, ".tools", "**", "program", "soffice"),
        "/usr/lib/libreoffice/program/soffice",
        "/usr/lib64/libreoffice/program/soffice",
        "/opt/libreoffice*/program/soffice",
        "/opt/libreoffice*/program/soffice.bin",
        os.path.join(home, "libreoffice*/program/soffice"),
        os.path.join(home, "squashfs-root", "AppRun"),
        "/snap/bin/libreoffice",
        # flatpak: system and per-user installs
        "/var/lib/flatpak/exports/bin/org.libreoffice.LibreOffice",
        os.path.join(home, ".local/share/flatpak/exports/bin/org.libreoffice.LibreOffice"),
        "/app/bin/soffice",
    ]
    for pattern in patterns:
        for candidate in sorted(glob.glob(pattern, recursive=True)):
            if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
                return candidate
    return None


class LibreOfficeEngine:
    name = "libreoffice"
    extensions = LIBREOFFICE_EXTENSIONS

    def __init__(self, binary: Optional[str] = None) -> None:
        self.binary = binary or find_soffice()

    def available(self) -> bool:
        return bool(self.binary) and os.path.exists(self.binary)

    def unavailable_reason(self) -> str:
        return ("未找到 LibreOffice/soffice（可选组件）。安装后可用 --engine libreoffice 获得"
                "最高保真度并支持 .doc/.odt/.rtf；不装也能用内置引擎转换 .docx，"
                "安装方式见 README 的「安装 LibreOffice（可选）」一节。"
                " / LibreOffice not found (optional): see README to install it, "
                "or keep using the built-in engine for .docx")

    def supports(self, extension: str) -> bool:
        return extension.lower() in self.extensions

    def version(self) -> str:
        if not self.binary:
            return ""
        try:
            out = subprocess.run([self.binary, "--version"], capture_output=True, text=True,
                                 timeout=60)
            return (out.stdout or out.stderr).strip().splitlines()[0] if (out.stdout or out.stderr) else ""
        except Exception:
            return ""

    def convert(self, source: str, target: str, timeout: float = DEFAULT_TIMEOUT,
                meta: Optional[Dict[str, Any]] = None) -> ConversionResult:
        if not self.available():
            raise ConversionError(self.unavailable_reason())
        source = os.path.abspath(source)
        target = os.path.abspath(target)
        os.makedirs(os.path.dirname(target) or ".", exist_ok=True)

        with tempfile.TemporaryDirectory(prefix="d2p-lo-") as tmp:
            outdir = os.path.join(tmp, "out")
            profile = os.path.join(tmp, "profile")
            os.makedirs(outdir, exist_ok=True)
            os.makedirs(profile, exist_ok=True)
            profile_uri = Path(profile).resolve().as_uri()
            args = [
                self.binary,
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--norestore",
                "--invisible",
                "--nolockcheck",
                "--nodefault",
                "--nofirststartwizard",
                "--nologo",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", outdir,
                source,
            ]
            env = dict(os.environ)
            env.setdefault("SAL_USE_VCLPLUGIN", "svp")
            env["HOME"] = profile
            env.pop("DISPLAY", None)
            try:
                proc = subprocess.run(args, capture_output=True, text=True, timeout=timeout,
                                      env=env, cwd=os.path.dirname(source) or None)
            except subprocess.TimeoutExpired:
                raise ConversionError(
                    f"LibreOffice 转换超时（{timeout:.0f}s）/ conversion timed out") from None
            produced = glob.glob(os.path.join(outdir, "*.pdf"))
            if not produced:
                detail = (proc.stderr or proc.stdout or "").strip()
                raise ConversionError(
                    "LibreOffice 未生成 PDF" + (f": {detail.splitlines()[-1]}" if detail else ""))
            shutil.move(produced[0], target)

        result = ConversionResult(source=source, target=target, engine=self.name, ok=True)
        result.pages = count_pages(target)
        return result


# --------------------------------------------------------------------------- #
# Selection helpers
# --------------------------------------------------------------------------- #


def count_pages(path: str) -> int:
    """Best-effort PDF page count with no mandatory dependency.

    Tries, in order: PyMuPDF (optional, AGPL), the ``pdfinfo`` utility from
    poppler, then a lightweight scan of the raw page objects.  Returns ``0``
    when the count cannot be determined -- it is only used for reporting.
    """
    try:
        import pymupdf

        with pymupdf.open(path) as pdf:
            return int(pdf.page_count)
    except Exception:
        pass
    pdfinfo = shutil.which("pdfinfo")
    if pdfinfo:
        try:
            proc = subprocess.run([pdfinfo, path], capture_output=True, text=True, timeout=60)
            for line in proc.stdout.splitlines():
                if line.lower().startswith("pages:"):
                    return int(line.split(":", 1)[1].strip())
        except Exception:
            pass
    try:
        import re

        if os.path.getsize(path) > 200 * 1024 * 1024:
            return 0
        with open(path, "rb") as handle:
            data = handle.read()
        return len(re.findall(rb"/Type\s*/Page[^s]", data))
    except Exception:
        return 0


def get_engine(name: str, font_dirs: Optional[Sequence[str]] = None):
    """Instantiate an engine by name (``builtin``/``libreoffice``)."""
    if name == "builtin":
        return BuiltinEngine(font_dirs)
    if name in ("libreoffice", "lo", "soffice"):
        return LibreOfficeEngine()
    raise ValueError(f"未知引擎 / unknown engine: {name}")


def select_engine(name: str, extension: str, font_dirs: Optional[Sequence[str]] = None):
    """Pick the best engine for *extension* honouring the user's ``--engine``."""
    builtin = BuiltinEngine(font_dirs)
    office = LibreOfficeEngine()

    if name == "auto":
        if office.available():
            return office
        if builtin.supports(extension):
            return builtin
        raise ConversionError(
            f"内置引擎不支持 {extension} 文件，且未安装 LibreOffice "
            f"/ built-in engine cannot read {extension}: install LibreOffice "
            f"(see README) or save the file as .docx first")

    engine = get_engine(name, font_dirs)
    if not engine.available():
        reason = engine.unavailable_reason()
        raise ConversionError(reason)
    if not engine.supports(extension):
        raise ConversionError(
            f"{engine.name} 引擎不支持 {extension} 文件 "
            f"/ engine {engine.name} cannot read {extension} files")
    return engine


def engines_status(font_dirs: Optional[Sequence[str]] = None) -> List[Dict[str, Any]]:
    """Environment report used by ``--doctor`` / ``--engines``."""
    builtin = BuiltinEngine(font_dirs)
    office = LibreOfficeEngine()
    status = [
        {
            "name": "builtin",
            "available": builtin.available(),
            "detail": ("python-docx + reportlab" if builtin.available()
                       else builtin.unavailable_reason()),
            "extensions": sorted(builtin.extensions),
        },
        {
            "name": "libreoffice",
            "available": office.available(),
            "detail": (office.version() or office.binary) if office.available()
                      else office.unavailable_reason(),
            "extensions": sorted(office.extensions),
        },
    ]
    return status


def doctor(font_dirs: Optional[Sequence[str]] = None) -> str:
    """Human readable environment diagnostics."""
    from .fonts import get_font_book

    lines = [f"d2p {__version__} 环境自检 / environment check", "-" * 52,
             f"Python      : {sys.version.split()[0]} ({sys.executable})"]
    for status in engines_status(font_dirs):
        mark = "OK " if status["available"] else "-- "
        lines.append(f"[{mark}] engine {status['name']}: {status['detail']}")
    try:
        book = get_font_book(font_dirs)
        lines.append("-" * 52)
        lines.append(book.explain())
        missing = book.missing_chars
        if missing:
            lines.append(f"缺少字形的字符 / characters without glyphs: {len(missing)}")
    except Exception as exc:
        lines.append(f"字体检测失败 / font detection failed: {exc}")
    return "\n".join(lines)
