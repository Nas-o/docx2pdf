"""Read a ``.docx`` package into a small, renderer friendly layout model.

``python-docx`` gives object access to the most common constructs but loses the
document order of paragraphs and tables, and does not expose ``w:rFonts``'s
East-Asian slot, fields, numbering definitions or per-cell borders.  This module
therefore walks the WordprocessingML tree directly and produces plain dataclasses
(:class:`Paragraph`, :class:`Table`, :class:`ImageBlock`, :class:`PageBreak`)
that :mod:`d2p.renderer` turns into PDF flowables.

Scope: body text, headings, lists (bullet + numbered, incl. Chinese numeral
formats), hyperlinks, inline/floating images, tables (borders, shading, merged
cells, nested tables), page breaks, sections, headers/footers and ``PAGE`` /
``NUMPAGES`` fields.  Footnotes, comments, equations and text boxes are not
rendered; they are reported as warnings.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field, replace
from typing import Any, Dict, Iterable, List, Optional, Tuple

from docx import Document as _open_docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.table import Table as _DocxTable
from docx.text.paragraph import Paragraph as _DocxParagraph

log = logging.getLogger("d2p.docx")

TWIPS = 1.0 / 20.0          # twips -> pt
EMU = 1.0 / 12700.0         # English Metric Unit -> pt
HALF_PT = 0.5               # w:sz is in half-points

RGB = Tuple[float, float, float]
_BLACK: RGB = (0.0, 0.0, 0.0)

_CJK_RANGES: Tuple[Tuple[int, int], ...] = (
    (0x2E80, 0x303F), (0x3040, 0x30FF), (0x3100, 0x312F), (0x3190, 0x31FF),
    (0x3200, 0x4DBF), (0x4E00, 0x9FFF), (0xA000, 0xA4CF), (0xAC00, 0xD7AF),
    (0xF900, 0xFAFF), (0xFE30, 0xFE4F), (0xFF00, 0xFFEF),
    (0x20000, 0x2FA1F), (0x1F300, 0x1FAFF),
)


def is_cjk_char(ch: str) -> bool:
    cp = ord(ch)
    for lo, hi in _CJK_RANGES:
        if lo <= cp <= hi:
            return True
    return False


def has_cjk(text: str) -> bool:
    return any(is_cjk_char(c) for c in text)


# --------------------------------------------------------------------------- #
# Model
# --------------------------------------------------------------------------- #


@dataclass
class Run:
    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: bool = False
    strike: bool = False
    size: Optional[float] = None          # pt
    color: Optional[RGB] = None
    highlight: Optional[RGB] = None
    family: Optional[str] = None          # family for Latin characters
    family_ea: Optional[str] = None       # family for CJK characters
    vert: Optional[str] = None            # 'super' | 'sub'
    link: Optional[str] = None
    is_marker: bool = False               # list bullet/number


@dataclass
class Paragraph:
    runs: List[Run] = field(default_factory=list)
    style_name: str = ""
    align: str = "left"
    size: float = 11.0
    leading: Optional[float] = None
    space_before: float = 0.0
    space_after: float = 0.0
    left_indent: float = 0.0
    right_indent: float = 0.0
    first_line_indent: float = 0.0
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    family: Optional[str] = None
    family_ea: Optional[str] = None
    color: Optional[RGB] = None
    shading: Optional[RGB] = None
    marker: Optional[str] = None          # rendered list marker, e.g. "1." / "•"
    marker_family: Optional[str] = None
    marker_size: Optional[float] = None
    keep_with_next: bool = False

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def is_empty(self) -> bool:
        return not self.text.strip() and self.marker is None


@dataclass
class ImageBlock:
    data: bytes
    width: float
    height: float
    name: str = ""
    align: str = "left"


@dataclass
class Border:
    style: str = "none"        # none | single | double | dashed | dotted
    size: float = 0.5          # pt
    color: RGB = _BLACK


@dataclass
class Cell:
    blocks: List[Any] = field(default_factory=list)
    grid_span: int = 1
    row_span: int = 1
    spanned: bool = False      # placeholder covered by a merge
    shading: Optional[RGB] = None
    v_align: str = "top"
    borders: Dict[str, Border] = field(default_factory=dict)


@dataclass
class Table:
    rows: List[List[Cell]] = field(default_factory=list)
    col_widths: List[float] = field(default_factory=list)
    header_rows: int = 0
    align: str = "left"
    cell_margins: Tuple[float, float, float, float] = (0.0, 5.4, 0.0, 5.4)
    row_heights: List[Optional[float]] = field(default_factory=list)


@dataclass
class PageBreak:
    pass


@dataclass
class Section:
    index: int = 0
    page_width: float = 595.28      # A4
    page_height: float = 841.89
    margin_left: float = 72.0
    margin_right: float = 72.0
    margin_top: float = 72.0
    margin_bottom: float = 72.0
    header_distance: float = 36.0
    footer_distance: float = 36.0
    header: List[Any] = field(default_factory=list)
    footer: List[Any] = field(default_factory=list)
    #: Word section break before this section starts a new page
    starts_new_page: bool = False
    #: blocks that belong to this section
    blocks: List[Any] = field(default_factory=list)

    @property
    def content_width(self) -> float:
        return self.page_width - self.margin_left - self.margin_right


@dataclass
class DocModel:
    sections: List[Section] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def section(self) -> Section:
        """First section; kept for convenience and single-section documents."""
        return self.sections[0] if self.sections else Section()

    @property
    def blocks(self) -> List[Any]:
        """All body blocks in document order (all sections flattened)."""
        out: List[Any] = []
        for section in self.sections:
            out.extend(section.blocks)
        return out

    def add_warning(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)


# --------------------------------------------------------------------------- #
# Low level XML helpers
# --------------------------------------------------------------------------- #


def _qpath(path: str) -> str:
    """``"w:rPr/w:rFonts"`` -> ``"{ns}rPr/{ns}rFonts"`` for lxml's ``find``."""
    if path.startswith("."):
        return path
    return "/".join(qn(part) if ":" in part else part for part in path.split("/"))


def _find(el, path: str):
    if el is None:
        return None
    return el.find(_qpath(path))


def _findall(el, path: str) -> List[Any]:
    if el is None:
        return []
    return list(el.findall(_qpath(path)))


def _attr(el, name: str) -> Optional[str]:
    if el is None:
        return None
    value = el.get(qn(name))
    if value is None:
        value = el.get(name)
    return value


def _toggle(el) -> Optional[bool]:
    """Interpret an OOXML boolean element (``w:b`` etc.)."""
    if el is None:
        return None
    val = _attr(el, "w:val")
    if val is None:
        return True
    return val.strip().lower() not in ("0", "false", "off", "no")


def parse_hex_color(value: Optional[str]) -> Optional[RGB]:
    if not value:
        return None
    value = value.strip().lstrip("#")
    if value.lower() in ("auto", "none"):
        return None
    if len(value) == 3:
        value = "".join(c * 2 for c in value)
    if len(value) != 6:
        return None
    try:
        r = int(value[0:2], 16) / 255.0
        g = int(value[2:4], 16) / 255.0
        b = int(value[4:6], 16) / 255.0
    except ValueError:
        return None
    return (r, g, b)


_HIGHLIGHT_COLORS: Dict[str, RGB] = {
    "black": (0, 0, 0), "blue": (0, 0, 1), "cyan": (0, 1, 1), "green": (0, 0.5, 0),
    "magenta": (1, 0, 1), "red": (1, 0, 0), "yellow": (1, 1, 0), "white": (1, 1, 1),
    "darkblue": (0, 0, 0.55), "darkcyan": (0, 0.55, 0.55), "darkgreen": (0, 0.39, 0),
    "darkmagenta": (0.55, 0, 0.55), "darkred": (0.55, 0, 0), "darkyellow": (0.5, 0.5, 0),
    "darkgray": (0.66, 0.66, 0.66), "lightgray": (0.83, 0.83, 0.83),
    "darkgrey": (0.66, 0.66, 0.66), "lightgrey": (0.83, 0.83, 0.83),
}


def _length(el, path: str) -> Optional[float]:
    """Read a ``w:...`` twips value at *path* and return points."""
    node = _find(el, path)
    if node is None:
        return None
    raw = _attr(node, "w:w")
    if raw is None:
        return None
    try:
        return float(raw) * TWIPS
    except ValueError:
        return None


# --------------------------------------------------------------------------- #
# Paragraph / run property resolution
# --------------------------------------------------------------------------- #

_RPR_TAGS = ("w:rFonts", "w:b", "w:i", "w:u", "w:strike", "w:dstrike", "w:sz",
             "w:szCs", "w:color", "w:vertAlign", "w:highlight", "w:bCs", "w:iCs")


class _RunProps:
    """Mutable bag of resolved run properties."""

    __slots__ = ("bold", "italic", "underline", "strike", "size", "color",
                 "highlight", "vert", "font_ascii", "font_hansi", "font_ea",
                 "font_cs")

    def __init__(self) -> None:
        self.bold: Optional[bool] = None
        self.italic: Optional[bool] = None
        self.underline: bool = False
        self.strike: bool = False
        self.size: Optional[float] = None
        self.color: Optional[RGB] = None
        self.highlight: Optional[RGB] = None
        self.vert: Optional[str] = None
        self.font_ascii: Optional[str] = None
        self.font_hansi: Optional[str] = None
        self.font_ea: Optional[str] = None
        self.font_cs: Optional[str] = None

    def apply(self, rpr) -> None:
        if rpr is None:
            return
        bold = _toggle(rpr.find(qn("w:b")))
        if bold is None:
            bold = _toggle(rpr.find(qn("w:bCs")))
        if bold is not None:
            self.bold = bold
        italic = _toggle(rpr.find(qn("w:i")))
        if italic is None:
            italic = _toggle(rpr.find(qn("w:iCs")))
        if italic is not None:
            self.italic = italic
        u = rpr.find(qn("w:u"))
        if u is not None:
            self.underline = (_attr(u, "w:val") or "single").lower() not in ("none", "0", "false")
        strike = rpr.find(qn("w:strike"))
        if strike is None:
            strike = rpr.find(qn("w:dstrike"))
        st = _toggle(strike)
        if st is not None:
            self.strike = st
        sz = rpr.find(qn("w:sz"))
        if sz is None:
            sz = rpr.find(qn("w:szCs"))
        val = _attr(sz, "w:val") if sz is not None else None
        if val:
            try:
                self.size = float(val) * HALF_PT
            except ValueError:
                pass
        color = rpr.find(qn("w:color"))
        if color is not None:
            rgb = parse_hex_color(_attr(color, "w:val"))
            if rgb is not None:
                self.color = rgb
            elif (_attr(color, "w:val") or "").lower() == "auto":
                self.color = None
        hl = rpr.find(qn("w:highlight"))
        if hl is not None:
            name = (_attr(hl, "w:val") or "").lower()
            if name not in ("", "none"):
                self.highlight = _HIGHLIGHT_COLORS.get(name)
        va = rpr.find(qn("w:vertAlign"))
        if va is not None:
            value = (_attr(va, "w:val") or "").lower()
            if value in ("superscript", "subscript"):
                self.vert = value
            elif value == "baseline":
                self.vert = None
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None:
            for key, attr in (("font_ascii", "w:ascii"), ("font_hansi", "w:hAnsi"),
                              ("font_ea", "w:eastAsia"), ("font_cs", "w:cs")):
                value = _attr(fonts, attr)
                if value:
                    setattr(self, key, value)


def _style_chain(style) -> List[Any]:
    chain, seen = [], set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        chain.append(style)
        try:
            style = style.base_style
        except Exception:
            break
    chain.reverse()          # base first, most derived last
    return chain


# --------------------------------------------------------------------------- #
# List numbering
# --------------------------------------------------------------------------- #

_BULLET_ALIASES: Dict[str, str] = {
    "\uf0b7": "•", "\uf0a7": "▪", "\uf0a8": "▪", "\uf0d8": "➢", "\uf0e0": "➢",
    "\uf075": "◆", "\uf06e": "■", "\uf0fc": "✔", "\uf0a1": "", "\uf0a2": "",
    "\uf0b6": "•", "\uf0d7": "▪", "\uf0fe": "■", "\uf0a4": "●", "\uf0a5": "●",
    "\u25cf": "●", "\u25cb": "○", "\u25a0": "■", "\u25aa": "▪", "\u2023": "‣",
    "\u2043": "⁃", "\u00b7": "·", "\u25c6": "◆", "\u25b6": "▶", "\u27a2": "➢",
}

_NUM_FMT_ALIASES: Dict[str, str] = {
    "decimal": "decimal", "decimalzero": "decimalZero", "lowerletter": "lowerLetter",
    "upperletter": "upperLetter", "lowerroman": "lowerRoman", "upperroman": "upperRoman",
    "chinesecounting": "chineseCounting", "chinesecountingthousand": "chineseCounting",
    "ideographdigital": "chineseCounting", "japanesecounting": "chineseCounting",
    "japaneselegal": "chineseCounting", "taiwanesecounting": "chineseCounting",
    "koreancounting": "chineseCounting", "arabicabjad": "decimal",
    "arabicalpha": "decimal", "hebrew1": "decimal", "thailetters": "decimal",
    "none": "none", "bullet": "bullet",
}

_CN_DIGITS = "零一二三四五六七八九"


def _to_chinese(number: int) -> str:
    if number <= 0:
        return str(number)
    if number < 10:
        return _CN_DIGITS[number]
    if number < 20:
        return "十" + (_CN_DIGITS[number - 10] if number > 10 else "")
    units = ["", "十", "百", "千"]
    out = ""
    digits = list(map(int, str(number)))
    length = len(digits)
    zero_pending = False
    for idx, digit in enumerate(digits):
        unit_index = length - idx - 1
        if digit == 0:
            zero_pending = True
            continue
        if zero_pending and out:
            out += "零"
        zero_pending = False
        if unit_index >= len(units):
            out += _CN_DIGITS[digit]
        else:
            out += _CN_DIGITS[digit] + units[unit_index]
    return out


def _to_roman(number: int, upper: bool) -> str:
    if number <= 0 or number > 3999:
        return str(number)
    table = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
             (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
             (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for value, letters in table:
        while number >= value:
            out += letters
            number -= value
    return out.upper() if upper else out


def _to_letters(number: int, upper: bool) -> str:
    if number <= 0:
        return str(number)
    out = ""
    while number > 0:
        number, rem = divmod(number - 1, 26)
        out = chr(ord("a") + rem) + out
    return out.upper() if upper else out


def format_counter(number: int, fmt: str) -> str:
    if fmt == "decimalZero":
        return f"{number:02d}"
    if fmt == "lowerLetter":
        return _to_letters(number, False)
    if fmt == "upperLetter":
        return _to_letters(number, True)
    if fmt == "lowerRoman":
        return _to_roman(number, False)
    if fmt == "upperRoman":
        return _to_roman(number, True)
    if fmt == "chineseCounting":
        return _to_chinese(number)
    return str(number)


# --------------------------------------------------------------------------- #
# Reader
# --------------------------------------------------------------------------- #


class DocxReader:
    def __init__(self, path: str) -> None:
        self.path = path
        self._temp_dir: Optional[str] = None
        self.doc, self._temp_dir = _open_word_document(path)
        self.model = DocModel()
        self._numbering: Dict[int, int] = {}          # numId -> abstractNumId
        self._abstract: Dict[int, Dict[int, Dict[str, Any]]] = {}
        self._counters: Dict[Tuple[int, int], int] = {}
        self._style_counters: Dict[str, int] = {}
        self._header_ids: set = set()
        self._docx_sections: List[Any] = []
        self._geometries: List[Section] = []
        self._section_cursor = 0
        self._pending_blocks: List[Any] = []
        self._last_header: List[Any] = []
        self._last_footer: List[Any] = []
        self._parse_numbering()
        self._parse_sections()

    # -- setup ------------------------------------------------------------- #

    def _parse_numbering(self) -> None:
        try:
            part = self.doc.part.numbering_part
        except Exception:
            return
        element = getattr(part, "element", None)
        if element is None:
            return
        for abstract in _findall(element, "w:abstractNum"):
            try:
                abstract_id = int(_attr(abstract, "w:abstractNumId") or -1)
            except ValueError:
                continue
            levels: Dict[int, Dict[str, Any]] = {}
            for lvl in _findall(abstract, "w:lvl"):
                try:
                    ilvl = int(_attr(lvl, "w:ilvl") or 0)
                except ValueError:
                    ilvl = 0
                fmt_raw = (_attr(_find(lvl, "w:numFmt"), "w:val") or "decimal").lower()
                start_raw = _attr(_find(lvl, "w:start"), "w:val")
                try:
                    start = int(start_raw) if start_raw else 1
                except ValueError:
                    start = 1
                levels[ilvl] = {
                    "fmt": _NUM_FMT_ALIASES.get(fmt_raw, "decimal"),
                    "text": _attr(_find(lvl, "w:lvlText"), "w:val") or "",
                    "start": start,
                    "font": _attr(_find(lvl, "w:rPr/w:rFonts"), "w:ascii"),
                    "indent_left": _length(_find(lvl, "w:pPr"), "w:ind"),
                    "suff": (_attr(_find(lvl, "w:suff"), "w:val") or "tab").lower(),
                }
                ind = _find(lvl, "w:pPr/w:ind")
                if ind is not None:
                    left = _attr(ind, "w:left") or _attr(ind, "w:start")
                    hanging = _attr(ind, "w:hanging")
                    if left:
                        try:
                            levels[ilvl]["indent_left"] = float(left) * TWIPS
                        except ValueError:
                            pass
                    if hanging:
                        try:
                            levels[ilvl]["hanging"] = float(hanging) * TWIPS
                        except ValueError:
                            pass
            self._abstract[abstract_id] = levels
        for num in _findall(element, "w:num"):
            try:
                num_id = int(_attr(num, "w:numId") or -1)
                abs_id = int(_attr(_find(num, "w:abstractNumId"), "w:val") or -1)
            except (TypeError, ValueError):
                continue
            override = _find(num, "w:lvlOverride")
            self._numbering[num_id] = abs_id
            if override is not None and abs_id in self._abstract:
                try:
                    ilvl = int(_attr(override, "w:ilvl") or 0)
                except ValueError:
                    ilvl = 0
                start_el = _find(override, "w:startOverride")
                if start_el is not None:
                    try:
                        self._abstract[abs_id].setdefault(ilvl, {})["start"] = int(
                            _attr(start_el, "w:val") or 1)
                    except ValueError:
                        pass

    def _parse_sections(self) -> None:
        """Read page geometry for every section up front.

        Content width is needed while parsing (tables/images are sized against
        it) but a section's ``sectPr`` only appears at the *end* of that
        section, so the geometry is collected before walking the body.
        """
        try:
            self._docx_sections = list(self.doc.sections)
        except Exception:
            self._docx_sections = []
        if not self._docx_sections:
            self.model.add_warning("文档没有节信息，使用 A4 默认页面设置 / no section, using A4 defaults")
            self._geometries = [Section()]
            return
        self._geometries = []
        for index, section in enumerate(self._docx_sections):
            target = Section(index=index)
            try:
                target.page_width = float(section.page_width.pt)
                target.page_height = float(section.page_height.pt)
                target.margin_left = float(section.left_margin.pt)
                target.margin_right = float(section.right_margin.pt)
                target.margin_top = float(section.top_margin.pt)
                target.margin_bottom = float(section.bottom_margin.pt)
                target.header_distance = float(section.header_distance.pt)
                target.footer_distance = float(section.footer_distance.pt)
            except Exception as exc:
                self.model.add_warning(f"页面设置读取失败 / cannot read page setup: {exc}")
            if index:
                try:
                    start_type = self._docx_sections[index - 1].start_type
                    target.starts_new_page = (start_type is None or
                                              str(start_type).split()[0] != "CONTINUOUS")
                except Exception:
                    target.starts_new_page = True
            self._geometries.append(target)

    def _current_section(self) -> Section:
        index = min(self._section_cursor, len(self._geometries) - 1)
        return self._geometries[max(index, 0)]

    def _close_section(self) -> None:
        """Finish the section currently being read and start the next one."""
        if self._geometries and len(self.model.sections) >= len(self._geometries):
            # Malformed package with more section breaks than sections: keep the
            # remaining content in the last section instead of dropping it.
            if self._pending_blocks:
                self.model.sections[-1].blocks.extend(self._pending_blocks)
                self._pending_blocks = []
            return
        index = min(self._section_cursor, len(self._geometries) - 1)
        geometry = self._current_section()
        section = replace(geometry, blocks=self._pending_blocks)
        if index < len(self._docx_sections):
            docx_section = self._docx_sections[index]
            try:
                section.header = self._read_hf(docx_section.header,
                                               self._last_header)
                section.footer = self._read_hf(docx_section.footer,
                                               self._last_footer)
            except Exception as exc:
                self.model.add_warning(f"页眉/页脚读取失败 / cannot read header/footer: {exc}")
            try:
                if docx_section.different_first_page_header_footer:
                    self.model.add_warning(
                        "首页页眉页脚不同，内置引擎未单独处理 "
                        "/ different first-page header/footer ignored")
            except Exception:
                pass
        self._last_header = section.header
        self._last_footer = section.footer
        self.model.sections.append(section)
        self._pending_blocks = []
        self._section_cursor = min(index + 1, max(len(self._geometries) - 1, 0))

    def _read_hf(self, container, inherited: Optional[List[Any]]) -> List[Any]:
        """Return header/footer blocks; inherits from the previous section."""
        if container is None:
            return list(inherited or [])
        try:
            if container.is_linked_to_previous:
                return list(inherited or [])
        except Exception:
            pass
        element = getattr(container, "_element", None)
        if element is None:
            return list(inherited or [])
        if id(element) in self._header_ids:
            return list(inherited or [])
        self._header_ids.add(id(element))
        blocks: List[Any] = []
        for child in self._iter_container(element, container):
            blocks.append(child)
        # Headers/footers are single-line in practice: drop trailing empties.
        while blocks and isinstance(blocks[-1], Paragraph) and blocks[-1].is_empty():
            blocks.pop()
        return blocks

    # -- block iteration --------------------------------------------------- #

    def _iter_container(self, element, parent) -> Iterable[Any]:
        for child in element.iterchildren():
            tag = child.tag
            if tag == qn("w:p"):
                yield from self._read_paragraph(_DocxParagraph(child, parent))
            elif tag == qn("w:tbl"):
                yield self._read_table(_DocxTable(child, parent))
            elif tag == qn("w:sdt"):
                content = _find(child, "w:sdtContent")
                if content is not None:
                    yield from self._iter_container(content, parent)

    def _walk_body(self, element, parent) -> Iterable[Tuple[str, Any]]:
        """Yield ``('blocks', [...])`` and ``('section_end', None)`` events."""
        for child in element.iterchildren():
            tag = child.tag
            if tag == qn("w:p"):
                yield ("blocks", list(self._read_paragraph(_DocxParagraph(child, parent))))
                if _find(child, "w:pPr/w:sectPr") is not None:
                    yield ("section_end", None)
            elif tag == qn("w:tbl"):
                yield ("blocks", [self._read_table(_DocxTable(child, parent))])
            elif tag == qn("w:sdt"):
                content = _find(child, "w:sdtContent")
                if content is not None:
                    yield from self._walk_body(content, parent)

    def read(self) -> DocModel:
        body = self.doc.element.body
        for kind, payload in self._walk_body(body, self.doc):
            if kind == "blocks":
                self._pending_blocks.extend(payload)
            else:
                self._close_section()
        # The body-level w:sectPr describes the final section.
        self._close_section()
        return self.model

    # -- paragraphs -------------------------------------------------------- #

    def _read_paragraph(self, paragraph) -> Iterable[Any]:
        el = paragraph._p
        ppr = el.find(qn("w:pPr"))

        blocks: List[Any] = []
        props = self._paragraph_props(paragraph, ppr)
        images: List[ImageBlock] = []

        # page break before
        if ppr is not None and _toggle(ppr.find(qn("w:pageBreakBefore"))) is True:
            blocks.append(PageBreak())

        segments = self._paragraph_segments(paragraph, props, images)
        for seg, page_break in segments:
            if page_break and (blocks or not seg.is_empty()):
                if blocks and not isinstance(blocks[-1], PageBreak):
                    blocks.append(PageBreak())
                elif not blocks:
                    blocks.append(PageBreak())
            if not seg.is_empty() or seg.runs or seg.marker:
                blocks.append(seg)
        for image in images:
            blocks.append(image)
        return blocks

    def _paragraph_props(self, paragraph, ppr) -> Paragraph:
        """Resolve paragraph level formatting into a :class:`Paragraph`."""
        para = Paragraph()
        try:
            para.style_name = paragraph.style.name or ""
        except Exception:
            para.style_name = ""
        style_chain = []
        try:
            style_chain = _style_chain(paragraph.style)
        except Exception:
            pass

        # --- defaults from document + styles + paragraph mark --------------
        base = _RunProps()
        base.apply(self._doc_default_rpr())
        for style in style_chain:
            base.apply(_find(getattr(style, "element", None), "w:rPr"))
        base.apply(_find(ppr, "w:rPr"))
        para.bold = base.bold
        para.italic = base.italic
        para.color = base.color
        para.family = self._family_of(base.font_ascii or base.font_hansi)
        para.family_ea = self._family_of(base.font_ea) or para.family

        size: Optional[float] = base.size
        leading: Optional[float] = None
        spacing_before: Optional[float] = None
        spacing_after: Optional[float] = None
        indent_left: Optional[float] = None
        indent_right: Optional[float] = None
        indent_first: Optional[float] = None
        align: Optional[str] = None
        keep_next: Optional[bool] = None

        def merge_pf(pf) -> None:
            nonlocal size, leading, spacing_before, spacing_after, indent_left
            nonlocal indent_right, indent_first, align, keep_next
            if pf is None:
                return
            try:
                if pf.space_before is not None:
                    spacing_before = float(pf.space_before.pt)
            except Exception:
                pass
            try:
                if pf.space_after is not None:
                    spacing_after = float(pf.space_after.pt)
            except Exception:
                pass
            try:
                if pf.left_indent is not None:
                    indent_left = float(pf.left_indent.pt)
            except Exception:
                pass
            try:
                if pf.right_indent is not None:
                    indent_right = float(pf.right_indent.pt)
            except Exception:
                pass
            try:
                if pf.first_line_indent is not None:
                    indent_first = float(pf.first_line_indent.pt)
            except Exception:
                pass
            try:
                if pf.alignment is not None:
                    align = _ALIGN_MAP.get(int(pf.alignment), align)
            except Exception:
                pass
            try:
                if pf.keep_with_next is not None:
                    keep_next = bool(pf.keep_with_next)
            except Exception:
                pass
            # line spacing
            try:
                rule = pf.line_spacing_rule
                value = pf.line_spacing
            except Exception:
                return
            if value is None:
                return
            if isinstance(value, float) or rule in (WD_LINE_SPACING.SINGLE,
                                                    WD_LINE_SPACING.ONE_POINT_FIVE,
                                                    WD_LINE_SPACING.DOUBLE,
                                                    WD_LINE_SPACING.MULTIPLE):
                multiple = float(value)
                if rule in (WD_LINE_SPACING.SINGLE,):
                    multiple = 1.0
                elif rule in (WD_LINE_SPACING.ONE_POINT_FIVE,):
                    multiple = 1.5
                elif rule in (WD_LINE_SPACING.DOUBLE,):
                    multiple = 2.0
                base_size = size or 11.0
                leading = max(base_size * 1.2 * multiple, base_size + 1.0)
            else:
                try:
                    exact = float(value.pt)
                except Exception:
                    return
                if rule == WD_LINE_SPACING.AT_LEAST:
                    leading = max(exact, (size or 11.0) * 1.15)
                else:
                    leading = exact

        for style in style_chain:
            try:
                merge_pf(style.paragraph_format)
            except Exception:
                pass
        try:
            merge_pf(paragraph.paragraph_format)
        except Exception:
            pass

        # --- list numbering ------------------------------------------------
        marker, marker_family, marker_size, list_indent, list_hanging = self._list_marker(para, ppr)

        if size is None:
            size = self._heading_size(para.style_name, base.size)
        para.size = size or 11.0
        para.leading = leading
        para.space_before = spacing_before if spacing_before is not None else 0.0
        para.space_after = spacing_after if spacing_after is not None else 0.0
        para.align = align or "left"
        para.keep_with_next = bool(keep_next)
        para.left_indent = indent_left if indent_left is not None else 0.0
        para.right_indent = indent_right if indent_right is not None else 0.0
        para.first_line_indent = indent_first if indent_first is not None else 0.0
        if list_indent is not None:
            base_indent = para.left_indent
            if list_indent < 0:
                list_indent = 0.0
            para.left_indent = max(base_indent, list_indent)
        if list_hanging:
            para.first_line_indent = -float(list_hanging)
        para.marker = marker
        para.marker_family = marker_family
        para.marker_size = marker_size

        # paragraph shading
        shd = _find(ppr, "w:shd")
        if shd is not None:
            fill = _attr(shd, "w:fill")
            if fill and fill.lower() not in ("auto", "none"):
                para.shading = parse_hex_color(fill)
        # character style / run-level overrides are applied per run
        return para

    def _doc_default_rpr(self):
        try:
            styles_el = self.doc.styles.element
        except Exception:
            return None
        dd = _find(styles_el, "w:docDefaults")
        if dd is None:
            return None
        rpr_default = _find(dd, "w:rPrDefault")
        if rpr_default is None:
            return None
        return _find(rpr_default, "w:rPr")

    def _family_of(self, word_font: Optional[str]) -> Optional[str]:
        if not word_font:
            return None
        return _FONTBOOK_FAMILY_CACHE.get(word_font) or self._lookup_family(word_font)

    def _lookup_family(self, word_font: str) -> str:
        from .fonts import get_font_book

        family = get_font_book().family_for_word_font(word_font)
        _FONTBOOK_FAMILY_CACHE[word_font] = family
        return family

    def _heading_size(self, style_name: str, base_size: Optional[float]) -> float:
        """Fallback heading sizes when the style carries no explicit size."""
        base = base_size or 11.0
        name = (style_name or "").lower()
        if not name:
            return base
        if "title" in name or "标题" == name:
            return base * 2.0
        if "subtitle" in name or "副标题" in name:
            return base * 1.3
        for index, factor in ((1, 1.45), (2, 1.27), (3, 1.13), (4, 1.06), (5, 1.0)):
            if f"heading {index}" in name or f"heading{index}" in name or f"标题 {index}" in name \
                    or f"标题{index}" in name or name.endswith(f" {index}"):
                return round(base * factor, 1)
        return base

    # -- runs -------------------------------------------------------------- #

    def _paragraph_segments(self, paragraph, props: Paragraph,
                            images: List[ImageBlock]) -> List[Tuple[Paragraph, bool]]:
        """Split a paragraph's runs into segments delimited by page breaks."""
        segments: List[Tuple[Paragraph, bool]] = []
        current = self._clone_shell(props)
        pending_break = False

        for child in paragraph._p.iterchildren():
            tag = child.tag
            if tag == qn("w:r"):
                for event, value in self._read_run_children(child, props, paragraph):
                    if event == "text":
                        current.runs.append(value)
                    elif event == "image":
                        images.append(value)
                    elif event == "pagebreak":
                        segments.append((current, pending_break))
                        pending_break = True
                        current = self._clone_shell(props)
                    elif event == "warning":
                        self.model.add_warning(value)
            elif tag == qn("w:hyperlink"):
                link = self._hyperlink_target(child, paragraph)
                run_props = self._run_props_for_hyperlink(child, props)
                for r in _findall(child, "w:r"):
                    for event, value in self._read_run_children(r, props, paragraph,
                                                               force_link=link,
                                                               inherited=run_props):
                        if event == "text":
                            current.runs.append(value)
                        elif event == "image":
                            images.append(value)
                        elif event == "pagebreak":
                            segments.append((current, pending_break))
                            pending_break = True
                            current = self._clone_shell(props)
            elif tag == qn("w:fldSimple"):
                instr = (_attr(child, "w:instr") or "").strip().upper()
                token = None
                if instr.startswith("PAGE"):
                    token = "{{PAGE}}"
                elif instr.startswith("NUMPAGES") or instr.startswith("SECTIONPAGES"):
                    token = "{{NUMPAGES}}"
                if token:
                    run = Run(text=token, family=props.family, family_ea=props.family_ea,
                              size=props.size, bold=props.bold, italic=props.italic,
                              color=props.color)
                    current.runs.append(run)
                else:
                    for r in _findall(child, "w:r"):
                        for event, value in self._read_run_children(r, props, paragraph):
                            if event == "text":
                                current.runs.append(value)
                            elif event == "image":
                                images.append(value)
            elif tag == qn("w:sdt"):
                content = _find(child, "w:sdtContent")
                if content is not None:
                    text = "".join((t.text or "") for t in content.iter(qn("w:t")))
                    if text:
                        current.runs.append(Run(text=text, family=props.family,
                                                family_ea=props.family_ea, size=props.size,
                                                bold=props.bold, italic=props.italic))
            elif tag == qn("w:bookmarkStart") or tag == qn("w:bookmarkEnd"):
                continue
            elif tag == qn("w:commentReference") or tag == qn("w:footnoteReference"):
                self.model.add_warning("忽略批注/脚注引用 / comments and footnotes are not rendered")

        segments.append((current, pending_break))
        return segments

    def _clone_shell(self, props: Paragraph) -> Paragraph:
        return Paragraph(
            runs=[],
            style_name=props.style_name,
            align=props.align,
            size=props.size,
            leading=props.leading,
            space_before=props.space_before,
            space_after=props.space_after,
            left_indent=props.left_indent,
            right_indent=props.right_indent,
            first_line_indent=props.first_line_indent,
            bold=props.bold,
            italic=props.italic,
            family=props.family,
            family_ea=props.family_ea,
            color=props.color,
            shading=props.shading,
            marker=props.marker,
            marker_family=props.marker_family,
            marker_size=props.marker_size,
            keep_with_next=props.keep_with_next,
        )

    def _hyperlink_target(self, element, paragraph) -> Optional[str]:
        rid = _attr(element, "r:id")
        if not rid:
            return None
        try:
            rel = paragraph.part.rels[rid]
            return rel.target_ref
        except Exception:
            return None

    def _run_props_for_hyperlink(self, element, paragraph_props: Paragraph) -> _RunProps:
        rpr = _find(element, "w:rPr")
        props = _RunProps()
        props.apply(self._doc_default_rpr())
        props.apply(rpr)
        return props

    def _read_run_children(self, run_el, props: Paragraph, paragraph,
                           force_link: Optional[str] = None,
                           inherited: Optional[_RunProps] = None
                           ) -> Iterable[Tuple[str, Any]]:
        rpr = run_el.find(qn("w:rPr"))
        resolved = _RunProps()
        resolved.apply(self._doc_default_rpr())
        if inherited is not None:
            for slot in _RunProps.__slots__:
                value = getattr(inherited, slot)
                if value is not None and value is not False:
                    setattr(resolved, slot, value)
        # paragraph style chain
        try:
            for style in _style_chain(paragraph.style):
                resolved.apply(_find(getattr(style, "element", None), "w:rPr"))
        except Exception:
            pass
        # paragraph mark properties
        resolved.apply(_find(paragraph._p.find(qn("w:pPr")), "w:rPr"))
        # character style
        if rpr is not None:
            rstyle = _attr(_find(rpr, "w:rStyle"), "w:val")
            if rstyle:
                try:
                    style = paragraph.part.get_style(rstyle, "character")
                    for s in _style_chain(style):
                        resolved.apply(_find(getattr(s, "element", None), "w:rPr"))
                except Exception:
                    pass
        resolved.apply(rpr)

        bold = resolved.bold if resolved.bold is not None else props.bold
        italic = resolved.italic if resolved.italic is not None else props.italic
        size = resolved.size or props.size
        color = resolved.color if resolved.color is not None else props.color
        family = self._family_of(resolved.font_ascii or resolved.font_hansi) or props.family
        family_ea = self._family_of(resolved.font_ea) or props.family_ea or family
        if _is_symbol_family(resolved.font_ascii):
            # Wingdings/Symbol bullets & glyphs: keep the requested family but
            # the renderer will remap private-use bullet codepoints.
            family = "latin-sans"

        def make_run(text: str, **kw) -> Run:
            return Run(text=text, bold=bold, italic=italic, size=size, color=color,
                       family=family, family_ea=family_ea,
                       underline=resolved.underline, strike=resolved.strike,
                       highlight=resolved.highlight, vert=resolved.vert,
                       link=force_link, **kw)

        for child in run_el.iterchildren():
            tag = child.tag
            if tag == qn("w:t"):
                text = child.text or ""
                if text:
                    yield ("text", make_run(text))
            elif tag == qn("w:br"):
                br_type = (_attr(child, "w:type") or "text").lower()
                if br_type == "page":
                    yield ("pagebreak", None)
                elif br_type == "column":
                    yield ("text", make_run("\n"))
                else:
                    yield ("text", make_run("\n"))
            elif tag == qn("w:tab"):
                yield ("text", make_run("\t"))
            elif tag in (qn("w:noBreakHyphen"),):
                yield ("text", make_run("\u2011"))
            elif tag == qn("w:softHyphen"):
                yield ("text", make_run("\u00ad"))
            elif tag == qn("w:sym"):
                char = _attr(child, "w:char")
                if char:
                    try:
                        yield ("text", make_run(chr(int(char, 16))))
                    except ValueError:
                        pass
            elif tag == qn("w:drawing"):
                found = list(self._images_from_drawing(child, paragraph))
                if not found:
                    text = "".join((t.text or "") for t in child.iter(qn("w:t")))
                    if text:
                        yield ("text", make_run(text))
                    else:
                        yield ("warning", "忽略不受支持的图形对象（图表/形状/EMF）"
                                          " / unsupported drawing skipped")
                for image in found:
                    yield ("image", image)
            elif tag == qn("w:pict"):
                found = list(self._images_from_vml(child, paragraph))
                if not found:
                    text = "".join((t.text or "") for t in child.iter(qn("w:t")))
                    if text:
                        yield ("text", make_run(text))
                for image in found:
                    yield ("image", image)
            elif tag == qn("w:object"):
                text = "".join((t.text or "") for t in child.iter(qn("w:t")))
                if text:
                    yield ("text", make_run(text))
            elif tag == qn("w:fldChar"):
                # complex fields: keep the cached result text, drop the code
                continue
            elif tag == qn("w:instrText"):
                continue
            elif tag == qn("w:rPr"):
                continue

    # -- images ------------------------------------------------------------ #

    def _image_part(self, rid: Optional[str], paragraph) -> Optional[Tuple[bytes, str]]:
        if not rid:
            return None
        try:
            part = paragraph.part.related_parts[rid]
        except Exception:
            return None
        blob = getattr(part, "blob", None)
        if not blob:
            return None
        name = os.path.basename(str(getattr(part, "partname", "") or ""))
        return blob, name

    def _images_from_drawing(self, drawing, paragraph) -> Iterable[ImageBlock]:
        for blip in drawing.iter(qn("a:blip")):
            rid = _attr(blip, "r:embed") or _attr(blip, "r:link")
            got = self._image_part(rid, paragraph)
            if not got:
                continue
            blob, name = got
            width = height = 0.0
            for parent in drawing.iter():
                if parent.tag in (qn("wp:extent"), qn("a:ext")):
                    try:
                        width = float(_attr(parent, "cx") or 0) * EMU
                        height = float(_attr(parent, "cy") or 0) * EMU
                    except (TypeError, ValueError):
                        width = height = 0.0
                    if width and height:
                        break
            image = self._make_image(blob, name, width, height)
            if image is not None:
                yield image

    def _images_from_vml(self, pict, paragraph) -> Iterable[ImageBlock]:
        for node in pict.iter(qn("v:imagedata")):
            rid = _attr(node, "r:id")
            got = self._image_part(rid, paragraph)
            if not got:
                continue
            blob, name = got
            width = height = 0.0
            shape = node.getparent()
            style = _attr(shape, "style") if shape is not None else None
            if style:
                for chunk in style.split(";"):
                    key, _, value = chunk.partition(":")
                    key, value = key.strip().lower(), value.strip().lower()
                    try:
                        if key == "width":
                            width = _css_length_to_pt(value)
                        elif key == "height":
                            height = _css_length_to_pt(value)
                    except ValueError:
                        pass
            image = self._make_image(blob, name, width, height)
            if image is not None:
                yield image

    def _make_image(self, blob: bytes, name: str, width: float, height: float) -> Optional[ImageBlock]:
        ext = os.path.splitext(name)[1].lower().lstrip(".")
        if ext in ("emf", "wmf", "svg", "eps"):
            self.model.add_warning(f"跳过矢量图 {name or ext}（内置引擎不支持）"
                                   f" / vector image {name or ext} skipped")
            return None
        if not width or not height:
            size = _image_intrinsic_size(blob)
            if size:
                width, height = size
            else:
                width, height = 300.0, 200.0
        # Cap to the printable width, preserving aspect ratio.
        available = max(self._current_section().content_width - 2.0, 72.0)
        if width > available:
            height *= available / width
            width = available
        return ImageBlock(data=blob, width=width, height=height, name=name)

    # -- lists ------------------------------------------------------------- #

    def _list_marker(self, para: Paragraph, ppr
                     ) -> Tuple[Optional[str], Optional[str], Optional[float],
                                Optional[float], Optional[float]]:
        num_pr = _find(ppr, "w:numPr")
        num_id = ilvl = None
        if num_pr is not None:
            raw_num = _attr(_find(num_pr, "w:numId"), "w:val")
            raw_ilvl = _attr(_find(num_pr, "w:ilvl"), "w:val")
            if raw_ilvl is not None:
                try:
                    ilvl = int(raw_ilvl)
                except ValueError:
                    ilvl = None
            if raw_num is not None:
                try:
                    num_id = int(raw_num)
                except ValueError:
                    num_id = None
                if num_id == 0:      # numbering removed
                    return None, None, None, None, None
        style_name = (para.style_name or "").lower()
        style_based = None
        if num_id is None:
            if "list bullet" in style_name or "列表段落" in style_name or "项目符号" in style_name:
                style_based = "bullet"
            elif "list number" in style_name or "列表编号" in style_name or "编号" in style_name:
                style_based = "decimal"
        if num_id is None and style_based is None:
            return None, None, None, None, None

        if num_id is None:
            key = f"style::{para.style_name}"
            count = self._style_counters.get(key, 0) + 1
            self._style_counters[key] = count
            if style_based == "bullet":
                return "•", "latin-sans", para.size, None, None
            return f"{count}.", "latin-sans", para.size, None, None

        ilvl = ilvl or 0
        abstract_id = self._numbering.get(num_id)
        levels = self._abstract.get(abstract_id, {}) if abstract_id is not None else {}
        level = levels.get(ilvl) or {}
        fmt = level.get("fmt", "decimal")
        if fmt == "none":
            return None, None, None, None, None

        # counters: increment this level, reset deeper ones
        start = int(level.get("start", 1))
        key = (num_id, ilvl)
        if key not in self._counters:
            self._counters[key] = start - 1
        self._counters[key] += 1
        for (nid, lvl) in list(self._counters):
            if nid == num_id and lvl > ilvl:
                del self._counters[(nid, lvl)]

        if fmt == "bullet":
            raw = level.get("text") or "•"
            text = _BULLET_ALIASES.get(raw, raw if raw.strip() else "•")
            if not text:
                return None, None, None, None, None
            family = "latin-sans"
            font_name = (level.get("font") or "")
            if _is_symbol_family(font_name):
                family = "latin-sans"
            return text, family, para.size, level.get("indent_left"), level.get("hanging")

        template = level.get("text") or f"%{ilvl + 1}."
        text = _render_number_template(template, num_id, ilvl, self._counters, fmt)
        if level.get("suff") in ("nothing", "space", "tab"):
            if level.get("suff") == "space":
                text += " "
            elif level.get("suff") in ("nothing", "tab"):
                text += "\t"
        return text, "latin-sans", para.size, level.get("indent_left"), level.get("hanging")

    # -- tables ------------------------------------------------------------ #

    def _read_table(self, table) -> Table:
        el = table._tbl
        result = Table()
        grid_cols = [_attr(c, "w:w") for c in _findall(el, "w:tblGrid/w:gridCol")]
        widths: List[float] = []
        for raw in grid_cols:
            try:
                widths.append(float(raw) * TWIPS)
            except (TypeError, ValueError):
                widths.append(0.0)

        rows_el = _findall(el, "w:tr")
        n_cols = len(widths)
        if n_cols == 0:
            for row_el in rows_el:
                n_cols = max(n_cols, sum(
                    max(1, int(_attr(_find(_find(tc, "w:tcPr"), "w:gridSpan"), "w:val") or 1))
                    for tc in _findall(row_el, "w:tc")))
        if n_cols == 0:
            n_cols = 1

        tbl_pr = _find(el, "w:tblPr")
        tbl_borders = self._read_borders(_find(tbl_pr, "w:tblBorders"))
        style_borders = self._style_borders(table)
        for name, border in style_borders.items():
            tbl_borders.setdefault(name, border)

        cell_margins = (0.0, 5.4, 0.0, 5.4)
        mar = _find(tbl_pr, "w:tblCellMar")
        if mar is not None:
            cell_margins = (
                (_length(mar, "w:top") or 0.0), (_length(mar, "w:left") or 5.4),
                (_length(mar, "w:bottom") or 0.0), (_length(mar, "w:right") or 5.4),
            )
        result.cell_margins = cell_margins

        align_el = _find(tbl_pr, "w:jc")
        if align_el is not None:
            result.align = {"center": "center", "right": "right", "end": "right",
                            "left": "left", "start": "left"}.get(
                (_attr(align_el, "w:val") or "left").lower(), "left")

        vmerge_state: Dict[int, Cell] = {}
        for row_index, row_el in enumerate(rows_el):
            row_cells: List[Cell] = []
            grid_pos = 0
            tr_pr = _find(row_el, "w:trPr")
            if _toggle(_find(tr_pr, "w:tblHeader")) is True:
                result.header_rows = row_index + 1
            row_height = None
            height_el = _find(tr_pr, "w:trHeight")
            if height_el is not None:
                try:
                    value = float(_attr(height_el, "w:val") or 0) * TWIPS
                    if (_attr(height_el, "w:hRule") or "auto").lower() == "exact":
                        row_height = value
                except ValueError:
                    row_height = None
            result.row_heights.append(row_height)

            for tc in _findall(row_el, "w:tc"):
                tc_pr = _find(tc, "w:tcPr")
                span_el = _find(tc_pr, "w:gridSpan")
                try:
                    span = max(1, int(_attr(span_el, "w:val") or 1))
                except ValueError:
                    span = 1
                vmerge_el = _find(tc_pr, "w:vMerge")
                vmerge = (_attr(vmerge_el, "w:val") or "continue").lower() if vmerge_el is not None else None

                cell = Cell(grid_span=span)
                shd = _find(tc_pr, "w:shd")
                if shd is not None:
                    fill = _attr(shd, "w:fill")
                    if fill and fill.lower() not in ("auto", "none"):
                        cell.shading = parse_hex_color(fill)
                valign = _find(tc_pr, "w:vAlign")
                if valign is not None:
                    cell.v_align = {"center": "middle", "bottom": "bottom",
                                    "top": "top"}.get(
                        (_attr(valign, "w:val") or "top").lower(), "top")
                cell.borders = self._cell_borders(tc_pr, tbl_borders, n_cols, grid_pos, span,
                                                  row_index, len(rows_el) - 1)

                if vmerge == "continue":
                    cell.spanned = True
                    start = vmerge_state.get(grid_pos)
                    if start is not None:
                        start.row_span += 1
                else:
                    blocks: List[Any] = []
                    for block in self._iter_container(tc, table):
                        blocks.append(block)
                    if not blocks:
                        blocks = [Paragraph()]
                    cell.blocks = blocks
                    if vmerge == "restart":
                        vmerge_state[grid_pos] = cell

                row_cells.append(cell)
                for extra in range(1, span):
                    row_cells.append(Cell(spanned=True, grid_span=0))
                    if vmerge == "restart":
                        vmerge_state[grid_pos + extra] = cell
                grid_pos += span

            while len(row_cells) < n_cols:
                row_cells.append(Cell(spanned=True))
            result.rows.append(row_cells[:n_cols])

        # column widths --------------------------------------------------
        available = max(self._current_section().content_width, 72.0)
        table_width_el = _find(tbl_pr, "w:tblW")
        table_pct = None
        if table_width_el is not None:
            kind = (_attr(table_width_el, "w:type") or "auto").lower()
            raw = _attr(table_width_el, "w:w")
            if kind == "pct" and raw:
                try:
                    table_pct = float(raw) / 50.0 / 100.0
                except ValueError:
                    table_pct = None
            elif kind == "dxa" and raw:
                try:
                    total = float(raw) * TWIPS
                    if total > 0:
                        table_pct = min(total / available, 1.0)
                except ValueError:
                    table_pct = None
        if len(widths) < n_cols:
            widths += [0.0] * (n_cols - len(widths))
        widths = widths[:n_cols]
        total = sum(widths)
        if total <= 1.0:
            widths = [available / n_cols] * n_cols
        else:
            if table_pct:
                target = available * table_pct
                widths = [w * target / total for w in widths]
            elif total > available:
                widths = [w * available / total for w in widths]
        result.col_widths = widths
        return result

    def _read_borders(self, borders_el) -> Dict[str, Border]:
        out: Dict[str, Border] = {}
        if borders_el is None:
            return out
        for child in borders_el.iterchildren():
            name = child.tag.split("}")[-1]
            out[name] = _border_from_element(child)
        return out

    def _style_borders(self, table) -> Dict[str, Border]:
        """Use the table style name as a hint (e.g. ``Table Grid``)."""
        try:
            name = (table.style.name or "").lower()
        except Exception:
            return {}
        if "grid" in name or "网格" in name or "all line" in name or "全部" in name:
            border = Border(style="single", size=0.5, color=_BLACK)
            return {key: border for key in ("top", "left", "bottom", "right",
                                            "insideH", "insideV")}
        return {}

    def _cell_borders(self, tc_pr, tbl_borders: Dict[str, Border], n_cols: int,
                      col: int, span: int, row: int, last_row: int) -> Dict[str, Border]:
        overrides = self._read_borders(_find(tc_pr, "w:tcBorders"))

        def pick(direct: Optional[Border], table_key: str) -> Border:
            if direct is not None:
                return direct
            return tbl_borders.get(table_key, Border(style="none", size=0.0))

        top = pick(overrides.get("top"), "top" if row == 0 else "insideH")
        bottom = pick(overrides.get("bottom"), "bottom" if row == last_row else "insideH")
        left = pick(overrides.get("left") or overrides.get("start"),
                    "left" if col == 0 else "insideV")
        right = pick(overrides.get("right") or overrides.get("end"),
                     "right" if col + span >= n_cols else "insideV")
        return {"top": top, "bottom": bottom, "left": left, "right": right}


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #

_ALIGN_MAP = {
    int(WD_ALIGN_PARAGRAPH.LEFT): "left",
    int(WD_ALIGN_PARAGRAPH.CENTER): "center",
    int(WD_ALIGN_PARAGRAPH.RIGHT): "right",
    int(WD_ALIGN_PARAGRAPH.JUSTIFY): "justify",
    int(WD_ALIGN_PARAGRAPH.DISTRIBUTE): "justify",
}

_FONTBOOK_FAMILY_CACHE: Dict[str, str] = {}


def _is_symbol_family(name: Optional[str]) -> bool:
    if not name:
        return False
    lowered = name.lower()
    return any(hint in lowered for hint in ("symbol", "wingdings", "webdings"))


def _border_from_element(el) -> Border:
    style = (_attr(el, "w:val") or "none").lower()
    if style in ("nil", "none"):
        return Border(style="none", size=0.0)
    raw_size = _attr(el, "w:sz")
    try:
        size = float(raw_size) / 8.0 if raw_size else 0.5
    except ValueError:
        size = 0.5
    color = parse_hex_color(_attr(el, "w:color")) or _BLACK
    if style in ("dashed", "dashsmallgap", "dotdash", "dashdotstroked"):
        style = "dashed"
    elif style in ("dotted", "dotdotdash"):
        style = "dotted"
    elif style in ("double", "doublewave", "thickthinlargegap"):
        style = "double"
    else:
        style = "single"
    return Border(style=style, size=max(size, 0.25), color=color)


def _render_number_template(template: str, num_id: int, ilvl: int,
                            counters: Dict[Tuple[int, int], int], fmt: str) -> str:
    out = template
    for level in range(0, ilvl + 1):
        token = f"%{level + 1}"
        if token in out:
            value = counters.get((num_id, level), 1)
            out = out.replace(token, format_counter(value, fmt if level == ilvl else "decimal"))
    # levels from other lists referenced with %N are left as-is; drop leftovers
    for level in range(ilvl + 1, 9):
        out = out.replace(f"%{level + 1}", "")
    return out


def _css_length_to_pt(value: str) -> float:
    value = value.strip().lower()
    try:
        if value.endswith("pt"):
            return float(value[:-2])
        if value.endswith("px"):
            return float(value[:-2]) * 0.75
        if value.endswith("in"):
            return float(value[:-2]) * 72.0
        if value.endswith("cm"):
            return float(value[:-2]) * 28.3465
        if value.endswith("mm"):
            return float(value[:-2]) * 2.83465
        if value.endswith("emu"):
            return float(value[:-3]) * EMU
        return float(value)
    except ValueError:
        return 0.0


def _image_intrinsic_size(blob: bytes) -> Optional[Tuple[float, float]]:
    try:
        import io

        from PIL import Image as PILImage

        with PILImage.open(io.BytesIO(blob)) as img:
            dpi = img.info.get("dpi") or (96, 96)
            try:
                dpi_x, dpi_y = float(dpi[0]), float(dpi[1])
            except (TypeError, ValueError, IndexError):
                dpi_x = dpi_y = 96.0
            dpi_x = dpi_x or 96.0
            dpi_y = dpi_y or 96.0
            return (img.width / dpi_x * 72.0, img.height / dpi_y * 72.0)
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# Package opening
# --------------------------------------------------------------------------- #

#: ``python-docx`` only accepts the plain document content type, so macro
#: enabled documents (``.docm``/``.dotm``) and templates (``.dotx``) are opened
#: through a rewritten copy of the package.
_MAIN_CONTENT_TYPES = {
    "application/vnd.ms-word.document.macroEnabled.main+xml":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
}


def _open_word_document(path: str):
    """Open *path*, working around python-docx's content-type restriction.

    Returns ``(document, temp_dir_or_None)``; the caller owns the temp dir.
    """
    try:
        return _open_docx(path), None
    except ValueError as exc:
        if "content type" not in str(exc).lower():
            raise
        patched = _rewrite_main_content_type(path)
        if patched is None:
            raise
        temp_dir, patched_path = patched
        try:
            return _open_docx(patched_path), temp_dir
        except Exception:
            import shutil

            shutil.rmtree(temp_dir, ignore_errors=True)
            raise


def _rewrite_main_content_type(path: str):
    """Copy the package with a standard main-part content type."""
    import shutil
    import tempfile
    import zipfile

    temp_dir = tempfile.mkdtemp(prefix="d2p-pkg-")
    target = os.path.join(temp_dir, os.path.basename(path) or "document.docx")
    try:
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as sink:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    text = data.decode("utf-8", "replace")
                    for old, new in _MAIN_CONTENT_TYPES.items():
                        text = text.replace(old, new)
                    data = text.encode("utf-8")
                sink.writestr(item, data)
        return temp_dir, target
    except Exception as exc:
        log.debug("cannot rewrite content types of %s: %s", path, exc)
        shutil.rmtree(temp_dir, ignore_errors=True)
        return None


def read_docx(path: str) -> DocModel:
    """Parse *path* into a :class:`DocModel`."""
    reader = DocxReader(path)
    try:
        model = reader.read()
    finally:
        if reader._temp_dir:
            import shutil

            shutil.rmtree(reader._temp_dir, ignore_errors=True)
            reader._temp_dir = None
    missing = None
    try:
        from .fonts import get_font_book
        book = get_font_book()
        missing = book.missing_chars
        for warning in book.warnings:
            model.add_warning(warning)
    except Exception:
        missing = None
    if missing:
        sample = "".join(sorted(missing)[:20])
        model.add_warning(f"部分字符缺少可用字体，可能显示为空白: {sample!r} "
                          f"/ {len(missing)} characters have no available glyph")
    return model
