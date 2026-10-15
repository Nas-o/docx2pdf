"""System font discovery, registration and script aware font selection.

The built-in renderer needs to draw arbitrary Word documents, which in practice
means a mix of Latin text, CJK text, punctuation and symbols.  Word asks for
fonts by name (``Arial``, ``宋体`` ...) but a Linux box rarely has those files.
This module therefore:

1. discovers usable TrueType fonts on the system (``/usr/share/fonts`` and
   friends, or directories given by ``$D2P_FONT_DIRS`` / ``--font-dir``),
2. registers them with ReportLab under stable, internal names,
3. maps Word font names onto a small set of *families*
   (``latin-sans``, ``latin-serif``, ``latin-mono``, ``cjk-sans``,
   ``cjk-serif``, ``cjk-kai``),
4. resolves, per character, the family that can actually draw it -- so a
   Chinese glyph inside an otherwise Arial paragraph automatically switches to
   a CJK face and vice versa.

Everything here degrades gracefully: if ``fontTools`` is missing we fall back to
a codepoint heuristic and if no TrueType font at all is found we use the PDF
base-14 fonts plus ReportLab's built-in Adobe CJK CID font.
"""

from __future__ import annotations

import logging
import os
import threading
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont, TTFError

log = logging.getLogger("d2p.fonts")

# --------------------------------------------------------------------------- #
# Candidate fonts
# --------------------------------------------------------------------------- #

Variant = Tuple[bool, bool]  # (bold, italic)

#: Font file name candidates per family and variant, best first.  Anything
#: ReportLab cannot load (CFF/PostScript outlines, e.g. Ubuntu's Noto CJK
#: ``.ttc`` builds) is skipped automatically and the next candidate is tried.
_FAMILY_SPECS: Dict[str, Dict[Variant, List[str]]] = {
    "latin-sans": {
        (False, False): [
            "DejaVuSans.ttf", "LiberationSans-Regular.ttf", "NotoSans-Regular.ttf",
            "FreeSans.ttf", "Carlito-Regular.ttf", "Arimo-Regular.ttf",
        ],
        (True, False): [
            "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf", "NotoSans-Bold.ttf",
            "FreeSansBold.ttf", "Carlito-Bold.ttf", "Arimo-Bold.ttf",
        ],
        (False, True): [
            "DejaVuSans-Oblique.ttf", "LiberationSans-Italic.ttf", "NotoSans-Italic.ttf",
            "FreeSansOblique.ttf", "Carlito-Italic.ttf", "Arimo-Italic.ttf",
        ],
        (True, True): [
            "DejaVuSans-BoldOblique.ttf", "LiberationSans-BoldItalic.ttf",
            "NotoSans-BoldItalic.ttf", "FreeSansBoldOblique.ttf",
            "Carlito-BoldItalic.ttf", "Arimo-BoldItalic.ttf",
        ],
    },
    "latin-serif": {
        (False, False): [
            "DejaVuSerif.ttf", "LiberationSerif-Regular.ttf", "NotoSerif-Regular.ttf",
            "FreeSerif.ttf", "Tinos-Regular.ttf",
        ],
        (True, False): [
            "DejaVuSerif-Bold.ttf", "LiberationSerif-Bold.ttf", "NotoSerif-Bold.ttf",
            "FreeSerifBold.ttf", "Tinos-Bold.ttf",
        ],
        (False, True): [
            "DejaVuSerif-Italic.ttf", "LiberationSerif-Italic.ttf", "NotoSerif-Italic.ttf",
            "FreeSerifItalic.ttf", "Tinos-Italic.ttf",
        ],
        (True, True): [
            "DejaVuSerif-BoldItalic.ttf", "LiberationSerif-BoldItalic.ttf",
            "NotoSerif-BoldItalic.ttf", "FreeSerifBoldItalic.ttf", "Tinos-BoldItalic.ttf",
        ],
    },
    "latin-mono": {
        (False, False): [
            "DejaVuSansMono.ttf", "LiberationMono-Regular.ttf", "NotoSansMono-Regular.ttf",
            "FreeMono.ttf", "Cousine-Regular.ttf",
        ],
        (True, False): [
            "DejaVuSansMono-Bold.ttf", "LiberationMono-Bold.ttf",
            "NotoSansMono-Bold.ttf", "FreeMonoBold.ttf", "Cousine-Bold.ttf",
        ],
        (False, True): [
            "DejaVuSansMono-Oblique.ttf", "LiberationMono-Italic.ttf",
            "NotoSansMono-Italic.ttf", "FreeMonoOblique.ttf", "Cousine-Italic.ttf",
        ],
        (True, True): [
            "DejaVuSansMono-BoldOblique.ttf", "LiberationMono-BoldItalic.ttf",
            "NotoSansMono-BoldItalic.ttf", "FreeMonoBoldOblique.ttf",
            "Cousine-BoldItalic.ttf",
        ],
    },
    # CJK faces usually ship in a single weight; bold/italic then reuse the
    # regular face (Word's bold is not synthesised -- see README limitations).
    "cjk-sans": {
        (False, False): [
            "DroidSansFallbackFull.ttf", "DroidSansFallback.ttf", "wqy-zenhei.ttc",
            "wqy-microhei.ttc", "NotoSansCJKsc-Regular.otf", "NotoSansCJK-Regular.ttc",
            "SourceHanSansSC-Regular.otf", "msyh.ttf", "simhei.ttf", "Arial Unicode.ttf",
            "unifont.ttf",
        ],
        (True, False): [
            "wqy-zenhei.ttc", "NotoSansCJKsc-Bold.otf", "NotoSansCJK-Bold.ttc",
            "SourceHanSansSC-Bold.otf", "msyhbd.ttf", "DroidSansFallbackFull.ttf",
        ],
        (False, True): [
            "NotoSansCJKsc-Regular.otf", "NotoSansCJK-Regular.ttc",
            "DroidSansFallbackFull.ttf", "wqy-zenhei.ttc",
        ],
        (True, True): [
            "NotoSansCJKsc-Bold.otf", "NotoSansCJK-Bold.ttc", "wqy-zenhei.ttc",
            "DroidSansFallbackFull.ttf",
        ],
    },
    "cjk-serif": {
        (False, False): [
            "NotoSerifCJKsc-Regular.otf", "NotoSerifCJK-Regular.ttc", "uming.ttc",
            "uming.ttf", "SourceHanSerifSC-Regular.otf", "simsun.ttc", "simsun.ttf",
        ],
        (True, False): [
            "NotoSerifCJKsc-Bold.otf", "NotoSerifCJK-Bold.ttc", "SourceHanSerifSC-Bold.otf",
            "uming.ttc", "simsun.ttc",
        ],
        (False, True): [
            "NotoSerifCJKsc-Regular.otf", "NotoSerifCJK-Regular.ttc", "uming.ttc",
        ],
        (True, True): [
            "NotoSerifCJKsc-Bold.otf", "NotoSerifCJK-Bold.ttc", "uming.ttc",
        ],
    },
    "cjk-kai": {
        (False, False): ["ukai.ttc", "ukai.ttf", "STKaiti.ttf", "simkai.ttf", "Kaiti.ttc"],
        (True, False): ["ukai.ttc", "STKaiti.ttf", "simkai.ttf"],
        (False, True): ["ukai.ttc", "STKaiti.ttf"],
        (True, True): ["ukai.ttc", "STKaiti.ttf"],
    },
}

#: When a family/variant is unavailable, try these in order.
_FALLBACK_CHAIN: Dict[str, List[str]] = {
    "latin-sans": ["latin-serif", "latin-mono"],
    "latin-serif": ["latin-sans", "latin-mono"],
    "latin-mono": ["latin-sans", "latin-serif"],
    "cjk-sans": ["cjk-serif", "cjk-kai", "latin-sans"],
    "cjk-serif": ["cjk-sans", "cjk-kai", "latin-serif"],
    "cjk-kai": ["cjk-serif", "cjk-sans", "latin-sans"],
}

#: Per character search order, relative to the family the document asked for.
_CHAR_SEARCH_ORDER: Dict[str, List[str]] = {
    "latin-sans": ["latin-sans", "latin-serif", "latin-mono", "cjk-sans", "cjk-serif", "cjk-kai"],
    "latin-serif": ["latin-serif", "latin-sans", "latin-mono", "cjk-serif", "cjk-sans", "cjk-kai"],
    "latin-mono": ["latin-mono", "latin-sans", "latin-serif", "cjk-sans", "cjk-serif", "cjk-kai"],
    "cjk-sans": ["cjk-sans", "cjk-serif", "cjk-kai", "latin-sans", "latin-serif", "latin-mono"],
    "cjk-serif": ["cjk-serif", "cjk-sans", "cjk-kai", "latin-serif", "latin-sans", "latin-mono"],
    "cjk-kai": ["cjk-kai", "cjk-serif", "cjk-sans", "latin-sans", "latin-serif", "latin-mono"],
}

_ALL_FAMILIES = tuple(_FAMILY_SPECS)

#: Word font name -> internal family.
_WORD_FONT_MAP: Dict[str, str] = {
    # Latin
    "times new roman": "latin-serif", "times": "latin-serif", "georgia": "latin-serif",
    "garamond": "latin-serif", "cambria": "latin-serif", "book antiqua": "latin-serif",
    "palatino linotype": "latin-serif", "constantia": "latin-serif", "serif": "latin-serif",
    "arial": "latin-sans", "helvetica": "latin-sans", "verdana": "latin-sans",
    "tahoma": "latin-sans", "calibri": "latin-sans", "segoe ui": "latin-sans",
    "carlito": "latin-sans", "lato": "latin-sans", "open sans": "latin-sans",
    "roboto": "latin-sans", "trebuchet ms": "latin-sans", "franklin gothic": "latin-sans",
    "sans-serif": "latin-sans", "comic sans ms": "latin-sans", "candara": "latin-sans",
    "courier new": "latin-mono", "courier": "latin-mono", "consolas": "latin-mono",
    "monaco": "latin-mono", "menlo": "latin-mono", "lucida console": "latin-mono",
    "dejavu sans mono": "latin-mono", "sf mono": "latin-mono", "monospace": "latin-mono",
    # Simplified Chinese
    "宋体": "cjk-serif", "simsun": "cjk-serif", "nsimsun": "cjk-serif",
    "新宋体": "cjk-serif", "songti sc": "cjk-serif", "songti": "cjk-serif",
    "stsong": "cjk-serif", "华文中宋": "cjk-serif", "仿宋": "cjk-serif",
    "fangsong": "cjk-serif", "fang song": "cjk-serif", "fangsong_gb2312": "cjk-serif",
    "微软雅黑": "cjk-sans", "microsoft yahei": "cjk-sans", "msyh": "cjk-sans",
    "黑体": "cjk-sans", "simhei": "cjk-sans", "heiti sc": "cjk-sans",
    "等线": "cjk-sans", "dengxian": "cjk-sans", "苹方": "cjk-sans",
    "pingfang sc": "cjk-sans", "思源黑体": "cjk-sans", "source han sans sc": "cjk-sans",
    "noto sans cjk sc": "cjk-sans", "华文细黑": "cjk-sans", "stxihei": "cjk-sans",
    "幼圆": "cjk-sans", "youyuan": "cjk-sans", "雅黑": "cjk-sans",
    "楷体": "cjk-kai", "kaiti": "cjk-kai", "kaiti sc": "cjk-kai",
    "华文楷体": "cjk-kai", "stkaiti": "cjk-kai", "楷体_gb2312": "cjk-kai",
    "隶书": "cjk-kai", "lisu": "cjk-kai",
    # Traditional Chinese / Japanese / Korean common names
    "新細明體": "cjk-serif", "mingliu": "cjk-serif", "細明體": "cjk-serif",
    "蘋方-繁": "cjk-sans", "ms jhenghei": "cjk-sans",
    "ms mincho": "cjk-serif", "yumincho": "cjk-serif", "メイリオ": "cjk-sans",
    "meiryo": "cjk-sans", "yu gothic": "cjk-sans", "malgun gothic": "cjk-sans",
    "batang": "cjk-serif", "gulim": "cjk-sans",
}

#: Font families with TrueType *symbol* glyphs for bullets.
_BULLET_HINTS = ("symbol", "wingdings", "webdings")


def _normalise(name: str) -> str:
    return " ".join(name.replace("_", " ").split()).strip().lower()


def is_wide_char(ch: str) -> bool:
    """Rough test for East Asian / wide codepoints (CJK, kana, Hangul, ...)."""
    cp = ord(ch)
    for low, high in ((0x2E80, 0x303F), (0x3040, 0x30FF), (0x3100, 0x4DBF),
                      (0x4E00, 0x9FFF), (0xA000, 0xA4CF), (0xAC00, 0xD7AF),
                      (0xF900, 0xFAFF), (0xFE30, 0xFE4F), (0xFF00, 0xFFEF),
                      (0x20000, 0x2FA1F), (0x1F300, 0x1FAFF)):
        if low <= cp <= high:
            return True
    return False


# --------------------------------------------------------------------------- #
# Discovery helpers
# --------------------------------------------------------------------------- #


def default_font_dirs() -> List[str]:
    """Font search path.

    ``$D2P_FONT_DIRS`` (``:`` separated) *replaces* the built-in list, which is
    handy in containers or when only specific fonts should be embedded.
    """
    from_env: List[str] = []
    for chunk in os.environ.get("D2P_FONT_DIRS", "").replace(";", os.pathsep).split(os.pathsep):
        chunk = chunk.strip()
        if chunk:
            from_env.append(os.path.expanduser(chunk))
    if from_env:
        return _existing_dirs(from_env)
    home = os.path.expanduser("~")
    dirs = [
        os.path.join(home, ".fonts"),
        os.path.join(home, ".local", "share", "fonts"),
        "/usr/local/share/fonts",
        "/usr/share/fonts",
        "/opt/fonts",
        "/Library/Fonts",
        "/System/Library/Fonts",
        "C:\\Windows\\Fonts",
    ]
    return _existing_dirs(dirs)


def _existing_dirs(dirs: Iterable[str]) -> List[str]:
    seen, out = set(), []
    for directory in dirs:
        if directory and directory not in seen and os.path.isdir(directory):
            seen.add(directory)
            out.append(directory)
    return out


def _index_font_files(roots: Sequence[str], max_depth: int = 6) -> Dict[str, str]:
    """Map ``filename.lower()`` -> absolute path for every font file found."""
    index: Dict[str, str] = {}
    exts = (".ttf", ".ttc", ".otf", ".otc")
    for root in roots:
        root = os.path.abspath(root)
        if not os.path.isdir(root):
            continue
        base_depth = root.rstrip(os.sep).count(os.sep)
        for dirpath, dirnames, filenames in os.walk(root):
            if dirpath.count(os.sep) - base_depth >= max_depth:
                dirnames[:] = []
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for fn in filenames:
                if fn.lower().endswith(exts):
                    index.setdefault(fn.lower(), os.path.join(dirpath, fn))
    return index


def _parse_spec(spec: str) -> Tuple[str, int]:
    """``"uming.ttc#1"`` -> ``("uming.ttc", 1)``."""
    if "#" in spec:
        name, _, idx = spec.partition("#")
        try:
            return name, int(idx)
        except ValueError:
            return name, 0
    return spec, 0


def _coverage_of(path: str, index: int) -> Optional[Set[int]]:
    """Set of Unicode codepoints covered by a font, or ``None`` if unknown."""
    try:
        from fontTools.ttLib import TTCollection, TTFont as FTFont
    except Exception:  # pragma: no cover - fontTools is optional
        return None
    try:
        if path.lower().endswith((".ttc", ".otc")):
            coll = TTCollection(path, lazy=True)
            try:
                font = coll.fonts[index]
            except IndexError:
                font = coll.fonts[0]
            try:
                cmap = font.getBestCmap() or {}
                return set(cmap.keys())
            finally:
                coll.close()
        font = FTFont(path, lazy=True, fontNumber=index)
        try:
            cmap = font.getBestCmap() or {}
            return set(cmap.keys())
        finally:
            font.close()
    except Exception as exc:  # pragma: no cover - malformed font
        log.debug("coverage lookup failed for %s: %s", path, exc)
        return None


# --------------------------------------------------------------------------- #
# Families
# --------------------------------------------------------------------------- #


@dataclass
class Face:
    path: str
    index: int
    registered_name: str
    coverage: Optional[Set[int]] = None


@dataclass
class Family:
    key: str
    faces: Dict[Variant, Face] = field(default_factory=dict)
    #: union of all face coverages, ``None`` when nothing was measurable
    coverage: Optional[Set[int]] = None
    name: str = ""

    def available(self) -> bool:
        return bool(self.faces)


class FontBook:
    """Discovers, registers and resolves fonts."""

    def __init__(self, font_dirs: Optional[Sequence[str]] = None) -> None:
        # Extra directories are *prepended* so user supplied fonts win, but the
        # system directories are still searched ($D2P_FONT_DIRS replaces them).
        self.requested_dirs = list(font_dirs) if font_dirs else None
        candidates: List[str] = []
        for extra in self.requested_dirs or []:
            candidates.append(os.path.expanduser(extra))
        candidates.extend(default_font_dirs())
        seen: Set[str] = set()
        self.font_dirs: List[str] = []
        for directory in candidates:
            if directory and directory not in seen and os.path.isdir(directory):
                seen.add(directory)
                self.font_dirs.append(directory)
        self._index = _index_font_files(self.font_dirs)
        self.families: Dict[str, Family] = {}
        self.warnings: List[str] = []
        self._slot_cache: Dict[Tuple[str, int], Tuple[str, bool]] = {}
        self._missing: Set[str] = set()
        self._helvetica = "Helvetica"
        self._cjk_fallback: Optional[str] = None
        self._build()

    # -- construction ------------------------------------------------------ #

    def _build(self) -> None:
        for key in _ALL_FAMILIES:
            fam = Family(key=key, name=key)
            self.families[key] = fam
            specs = _FAMILY_SPECS[key]
            for variant, candidates in specs.items():
                extra = os.environ.get("D2P_FONT_" + key.replace("-", "_").upper())
                if extra and variant == (False, False):
                    candidates = [extra] + list(candidates)
                face = self._load_first(candidates, key, variant)
                if face is not None:
                    fam.faces.setdefault(variant, face)
            # Fill holes: any missing variant reuses the closest available face.
            if fam.faces:
                for variant in ((False, False), (True, False), (False, True), (True, True)):
                    if variant not in fam.faces:
                        donor = fam.faces.get((False, False)) or next(iter(fam.faces.values()))
                        fam.faces[variant] = donor
                cov: Optional[Set[int]] = set()
                seen_faces: Dict[int, Face] = {}
                for face in fam.faces.values():
                    seen_faces[id(face)] = face
                for face in seen_faces.values():
                    if face.coverage is None:
                        cov = None
                        break
                    cov |= face.coverage
                fam.coverage = cov
            else:
                self.warnings.append(f"no usable font found for family {key!r}")
        self._ensure_cjk_fallback_family()

    def _ensure_cjk_fallback_family(self) -> None:
        """Last resort when the machine has no CJK TrueType font at all.

        ReportLab ships an Adobe CID font (``STSong-Light``) which viewers
        substitute with a local CJK face, so Chinese text still renders instead
        of disappearing.
        """
        if any(self.families[key].available() for key in ("cjk-sans", "cjk-serif", "cjk-kai")):
            return
        self._ensure_base14()
        if not self._cjk_fallback:
            return
        face = Face(path="<built-in Adobe CJK CID font>", index=0,
                    registered_name=self._cjk_fallback, coverage=None)
        family = Family(key="cjk-fallback", name="cjk-fallback", coverage=None)
        for variant in ((False, False), (True, False), (False, True), (True, True)):
            family.faces[variant] = face
        self.families["cjk-fallback"] = family
        # The three generic "no usable font" notes are implied by the summary.
        self.warnings = [w for w in self.warnings
                         if not any(key in w for key in ("'cjk-sans'", "'cjk-serif'", "'cjk-kai'"))]
        self.warnings.append(
            "未找到 CJK 字体，使用内置 Adobe CID 字体 STSong-Light（显示由阅读器替换）"
            " / no CJK TrueType font found; using built-in STSong-Light")

    def _load_first(self, candidates: Iterable[str], key: str, variant: Variant) -> Optional[Face]:
        suffix = ("b" if variant[0] else "") + ("i" if variant[1] else "")
        base_name = "d2p-" + key + ("-" + suffix if suffix else "")
        for spec in candidates:
            filename, index = _parse_spec(spec)
            if os.path.isabs(filename):
                path = filename if os.path.exists(filename) else None
            else:
                path = self._index.get(filename.lower())
            if not path:
                continue
            try:
                reg = TTFont(base_name, path, subfontIndex=index)
            except (TTFError, OSError, ValueError) as exc:
                log.debug("cannot embed %s: %s", path, exc)
                continue
            try:
                pdfmetrics.registerFont(reg)
            except Exception as exc:  # already registered / broken tables
                log.debug("registerFont failed for %s: %s", path, exc)
                continue
            return Face(path=path, index=index, registered_name=base_name,
                        coverage=_coverage_of(path, index))
        return None

    # -- registration helpers --------------------------------------------- #

    def _ensure_base14(self) -> None:
        """Guarantee a usable fallback even on a font-less system."""
        if self._cjk_fallback:
            return
        try:
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont

            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
            self._cjk_fallback = "STSong-Light"
        except Exception as exc:  # pragma: no cover
            log.debug("CID font unavailable: %s", exc)
            self._cjk_fallback = "Helvetica"

    def _base14_for(self, key: str) -> str:
        if key.startswith("latin-serif"):
            return "Times-Roman"
        if key.startswith("latin-mono"):
            return "Courier"
        return "Helvetica"

    def resolve(self, key: str, bold: bool = False, italic: bool = False) -> str:
        """Registered ReportLab font name for *key* with the requested style."""
        variant = (bool(bold), bool(italic))
        seen = []
        for family_key in [key] + _FALLBACK_CHAIN.get(key, []):
            if family_key in seen:
                continue
            seen.append(family_key)
            fam = self.families.get(family_key)
            if fam and fam.available():
                face = fam.faces.get(variant)
                if face is None:
                    face = fam.faces.get((False, False)) or next(iter(fam.faces.values()))
                return face.registered_name
        self._ensure_base14()
        if key.startswith("cjk"):
            return self._cjk_fallback or "Helvetica"
        return self._base14_for(key)

    def resolves_family(self, key: str) -> Optional[str]:
        """The family key that :meth:`resolve` would actually use."""
        for family_key in [key] + _FALLBACK_CHAIN.get(key, []):
            fam = self.families.get(family_key)
            if fam and fam.available():
                return family_key
        return None

    # -- per character font selection -------------------------------------- #

    def family_for_word_font(self, name: Optional[str]) -> str:
        """Map a Word ``w:rFonts`` name onto an internal family."""
        if not name:
            return "latin-sans"
        key = _normalise(name)
        if key in _WORD_FONT_MAP:
            return _WORD_FONT_MAP[key]
        for word_name, family in _WORD_FONT_MAP.items():
            if word_name and word_name in key:
                return family
        if any(h in key for h in _BULLET_HINTS):
            return "symbol"
        # Unknown name: guess by script of the name itself.
        if any(ord(ch) > 0x2E80 for ch in name):
            return "cjk-sans"
        return "latin-sans"

    def family_for_char(self, ch: str, requested: str) -> str:
        """Best family able to draw *ch*, preferring *requested*."""
        cp = ord(ch)
        cache_key = (requested, cp)
        hit = self._slot_cache.get(cache_key)
        if hit is not None:
            return hit[0]
        order = _CHAR_SEARCH_ORDER.get(requested)
        if order is None:
            order = _CHAR_SEARCH_ORDER["latin-sans"]
        chosen = None
        for family_key in order:
            fam = self.families.get(family_key)
            if fam and fam.coverage and cp in fam.coverage:
                chosen = family_key
                break
        if chosen is None:
            # No cmap information, or genuinely uncovered: use the requested
            # family so at least the metrics stay close.
            fam = self.families.get(requested)
            if fam and fam.coverage is None and fam.available():
                chosen = requested
            elif ("cjk-fallback" in self.families
                  and (is_wide_char(ch) or requested.startswith("cjk"))):
                chosen = "cjk-fallback"
            else:
                chosen = self.resolves_family(requested) or requested
                self._missing.add(ch)
        self._slot_cache[cache_key] = (chosen, True)
        return chosen

    def split(self, text: str, requested: str, bold: bool = False,
              italic: bool = False) -> List[Tuple[str, str]]:
        """Split *text* into ``(chunk, font_name)`` pairs of uniform font."""
        out: List[Tuple[str, str]] = []
        if not text:
            return out
        current_family: Optional[str] = None
        buf: List[str] = []

        def flush() -> None:
            if buf:
                out.append(("".join(buf), self.resolve(current_family or requested, bold, italic)))
                buf.clear()

        for ch in text:
            fam = self.family_for_char(ch, requested)
            if fam != current_family:
                flush()
                current_family = fam
            buf.append(ch)
        flush()
        return out

    @property
    def missing_chars(self) -> Set[str]:
        return set(self._missing)

    def explain(self) -> str:
        lines = ["字体目录 / font directories:"]
        lines += [f"  - {d}" for d in self.font_dirs]
        lines.append("已注册字族 / registered families:")
        for key in _ALL_FAMILIES:
            fam = self.families[key]
            if not fam.available():
                lines.append(f"  {key:<12} (unavailable)")
                continue
            reg = fam.faces[(False, False)]
            bold = fam.faces.get((True, False))
            extra = f" glyphs={len(fam.coverage)}" if fam.coverage else " glyphs=?"
            b = f", bold={os.path.basename(bold.path)}" if bold and bold.path != reg.path else ""
            lines.append(f"  {key:<12} {os.path.basename(reg.path)}#{reg.index}{extra}{b}")
        for w in self.warnings:
            lines.append(f"  ! {w}")
        return "\n".join(lines)


_BOOK: Optional[FontBook] = None
_BOOK_LOCK = threading.Lock()


def get_font_book(font_dirs: Optional[Sequence[str]] = None) -> FontBook:
    """Process-wide :class:`FontBook` (constructed on first use)."""
    global _BOOK
    with _BOOK_LOCK:
        if _BOOK is None or (
                font_dirs is not None and list(font_dirs) != (_BOOK.requested_dirs or [])):
            _BOOK = FontBook(font_dirs)
        return _BOOK


def reset_font_book() -> None:
    global _BOOK
    with _BOOK_LOCK:
        _BOOK = None
