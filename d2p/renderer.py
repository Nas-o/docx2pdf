"""Turn a :class:`~d2p.docxmodel.DocModel` into a PDF using ReportLab.

The renderer is deliberately conservative: it maps Word's paragraph/run/table
formatting onto ReportLab flowables (``Paragraph``, ``Table``, ``Image``) so
that pagination, line breaking, justification and table splitting stay
battle-tested instead of being re-implemented.  Fonts come from
:mod:`d2p.fonts`, which also splits mixed Latin/CJK text across faces.
"""

from __future__ import annotations

import io
import logging
import os
from dataclasses import dataclass
from functools import partial
from typing import Any, Dict, List, Optional, Sequence, Tuple

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfgen import canvas as _canvas
from reportlab.platypus import (BaseDocTemplate, Flowable, Frame, Image, NextPageTemplate,
                               PageBreak as RLPageBreak, PageTemplate,
                               Paragraph as RLParagraph, Spacer, Table as RLTable, TableStyle)

from . import __version__
from .docxmodel import (Cell, DocModel, ImageBlock, PageBreak, Paragraph, Run, Section,
                        Table, has_cjk, is_cjk_char)
from .engines import count_pages
from .fonts import FontBook, get_font_book

log = logging.getLogger("d2p.renderer")

_ALIGN = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT, "justify": TA_JUSTIFY}


def _css(color) -> str:
    from reportlab.lib.colors import Color

    c = Color(*color) if not isinstance(color, Color) else color
    return "#%02x%02x%02x" % (int(c.red * 255 + 0.5), int(c.green * 255 + 0.5),
                              int(c.blue * 255 + 0.5))


def _rl_color(color) -> colors.Color:
    if isinstance(color, colors.Color):
        return color
    return colors.Color(*color)


def _escape(text: str) -> str:
    out = []
    for ch in text:
        if ch == "&":
            out.append("&amp;")
        elif ch == "<":
            out.append("&lt;")
        elif ch == ">":
            out.append("&gt;")
        elif ch == "\u00a0":
            out.append("&nbsp;")
        elif ch in "\x0b\x0c":
            out.append("<br/>")
        elif ch == "\n":
            out.append("<br/>")
        elif ch == "\r":
            continue
        elif ch == "\u00ad":       # soft hyphen: invisible
            continue
        elif ch == "\t":
            out.append("&nbsp;&nbsp;&nbsp;&nbsp;")
        elif ord(ch) < 32:
            continue
        else:
            out.append(ch)
    return "".join(out)


def _escape_attr(text: str) -> str:
    return (text.replace("&", "&amp;").replace('"', "&quot;")
            .replace("<", "&lt;").replace(">", "&gt;"))


# --------------------------------------------------------------------------- #
# Canvas with deferred page furniture (needed for NUMPAGES)
# --------------------------------------------------------------------------- #


class _DeferredCanvas(_canvas.Canvas):
    def __init__(self, *args, drawer=None, **kwargs):
        self._page_states: List[Dict[str, Any]] = []
        super().__init__(*args, **kwargs)
        self._drawer = drawer

    def showPage(self):  # noqa: N802 - ReportLab API
        self._page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._page_states)
        for state in self._page_states:
            self.__dict__.update(state)
            if self._drawer is not None:
                try:
                    self._drawer(self, total)
                except Exception as exc:  # never fail a build over a header
                    log.warning("header/footer drawing failed: %s", exc)
            _canvas.Canvas.showPage(self)
        self._page_states = []
        _canvas.Canvas.save(self)


# --------------------------------------------------------------------------- #
# Renderer
# --------------------------------------------------------------------------- #


@dataclass
class RenderStats:
    pages: int = 0
    paragraphs: int = 0
    tables: int = 0
    images: int = 0


class PdfRenderer:
    def __init__(self, book: Optional[FontBook] = None, font_dirs: Optional[Sequence[str]] = None,
                 quiet: bool = False) -> None:
        self.book = book or get_font_book(font_dirs)
        self.stats = RenderStats()
        self.quiet = quiet

    # -- public ------------------------------------------------------------ #

    def render(self, model: DocModel, out_path: str, title: Optional[str] = None) -> RenderStats:
        sections = list(model.sections) or [model.section]
        if not any(section.blocks for section in sections):
            sections[0].blocks = []

        # One page template per Word section (portrait/landscape mixes, different
        # margins and per-section headers/footers).
        self._page_section: Dict[int, int] = {}
        story: List[Flowable] = []
        templates: List[PageTemplate] = []
        for index, section in enumerate(sections):
            template_id = f"d2p-sec{index}"
            frame = Frame(
                section.margin_left,
                section.margin_bottom,
                max(section.content_width, 36.0),
                max(section.page_height - section.margin_top - section.margin_bottom, 36.0),
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                id=f"body{index}",
            )
            templates.append(PageTemplate(
                id=template_id,
                frames=[frame],
                pagesize=(section.page_width, section.page_height),
                onPage=self._page_recorder(index),
            ))
            if index:
                story.append(NextPageTemplate(template_id))
                if section.starts_new_page:
                    story.append(RLPageBreak())
            story.extend(self._build_story(section.blocks, section, None))
        if not story:
            story = [Spacer(1, 1)]

        first = sections[0]
        drawer = self._make_drawer(sections)
        doc = BaseDocTemplate(
            out_path,
            pagesize=(first.page_width, first.page_height),
            leftMargin=first.margin_left,
            rightMargin=first.margin_right,
            topMargin=first.margin_top,
            bottomMargin=first.margin_bottom,
            title=title or os.path.basename(out_path),
            author="d2p",
            creator=f"d2p {__version__} (built-in engine)",
            producer=f"d2p {__version__}",
        )
        doc.addPageTemplates(templates)
        doc.build(story, canvasmaker=partial(_DeferredCanvas, drawer=drawer))

        # ReportLab tracks the page count itself, so no PDF library is needed
        # at runtime (PyMuPDF is AGPL and only used as an optional extra).
        pages = int(getattr(doc, "page", 0) or 0)
        if pages <= 0:
            pages = count_pages(out_path)
        self.stats.pages = pages
        return self.stats

    def _page_recorder(self, section_index: int):
        """Remember which Word section each PDF page belongs to."""

        def record(canv, doc):
            self._page_section[canv.getPageNumber()] = section_index

        return record

    # -- story construction ------------------------------------------------ #

    def _build_story(self, blocks: Sequence[Any], section: Section,
                     tokens: Optional[Dict[str, str]]) -> List[Flowable]:
        story: List[Flowable] = []
        for block in blocks:
            flowables = self._block_to_flowables(block, section, tokens)
            story.extend(flowables)
        return story

    def _block_to_flowables(self, block: Any, section: Section,
                            tokens: Optional[Dict[str, str]]) -> List[Flowable]:
        if isinstance(block, Paragraph):
            return self._paragraph_flowables(block, section, tokens)
        if isinstance(block, Table):
            self.stats.tables += 1
            return [self._table_flowable(block, section, tokens)]
        if isinstance(block, ImageBlock):
            self.stats.images += 1
            return [self._image_flowable(block, section)]
        if isinstance(block, PageBreak):
            return [RLPageBreak()]
        return []

    # -- paragraphs -------------------------------------------------------- #

    def _paragraph_flowables(self, para: Paragraph, section: Section,
                             tokens: Optional[Dict[str, str]]) -> List[Flowable]:
        size = para.size or 11.0
        leading = para.leading or max(size * 1.2, size + 1.5)
        align = _ALIGN.get(para.align, TA_LEFT)
        base_family = para.family or "latin-sans"
        base_font = self.book.resolve(base_family, bool(para.bold), bool(para.italic))

        marker = para.marker
        bullet_indent = None
        bullet_font = self.book.resolve(para.marker_family or "latin-sans")
        if marker:
            marker = marker.replace("\t", "").rstrip()
            hanging = abs(para.first_line_indent) or min(para.left_indent or 18.0, 18.0)
            if para.first_line_indent >= 0:
                para.first_line_indent = -hanging
            bullet_indent = max(para.left_indent - hanging, 0.0)
            # The marker is drawn in a single font: pick one that has the glyphs
            # (Chinese list numbers such as "一、" need a CJK face).
            first = next((c for c in marker if not c.isspace()), "")
            if first:
                requested = (para.marker_family or "latin-sans")
                if is_cjk_char(first):
                    requested = "cjk-sans" if not requested.startswith("cjk") else requested
                bullet_font = self.book.resolve(
                    self.book.family_for_char(first, requested), bool(para.bold))

        style = ParagraphStyle(
            name="d2p-body",
            fontName=base_font,
            fontSize=size,
            leading=leading,
            alignment=align,
            leftIndent=para.left_indent,
            rightIndent=para.right_indent,
            firstLineIndent=para.first_line_indent,
            spaceBefore=para.space_before,
            spaceAfter=para.space_after,
            backColor=_rl_color(para.shading) if para.shading else None,
            textColor=_rl_color(para.color) if para.color else colors.black,
            keepWithNext=1 if para.keep_with_next else 0,
            bulletFontName=bullet_font,
            bulletFontSize=para.marker_size or size,
            bulletIndent=bullet_indent if bullet_indent is not None else 0,
            wordWrap="CJK" if has_cjk(para.text) else None,
            splitLongWords=1,
            allowWidows=1,
            allowOrphans=1,
        )

        if para.is_empty() and not marker:
            out: List[Flowable] = []
            if para.space_before:
                out.append(Spacer(1, para.space_before))
            out.append(Spacer(1, leading))
            if para.space_after:
                out.append(Spacer(1, para.space_after))
            return out

        markup = self._paragraph_markup(para, tokens)
        if not markup.strip() and not marker:
            return [Spacer(1, leading)]
        self.stats.paragraphs += 1
        try:
            return [RLParagraph(markup, style, bulletText=marker)]
        except Exception as exc:
            log.warning("paragraph layout failed (%s); falling back to plain text", exc)
            plain = _escape(para.text)
            return [RLParagraph(plain, style, bulletText=marker)]

    def _paragraph_markup(self, para: Paragraph, tokens: Optional[Dict[str, str]]) -> str:
        chunks: List[str] = []
        for run in para.runs:
            if not run.text:
                continue
            chunks.append(self._run_markup(run, para, tokens))
        return "".join(chunks)

    def _run_markup(self, run: Run, para: Paragraph, tokens: Optional[Dict[str, str]]) -> str:
        text = run.text
        if tokens:
            for key, value in tokens.items():
                if key in text:
                    text = text.replace(key, value)
        bold = run.bold if run.bold is not None else bool(para.bold)
        italic = run.italic if run.italic is not None else bool(para.italic)
        size = run.size or para.size or 11.0
        family_latin = run.family or para.family or "latin-sans"
        family_ea = run.family_ea or para.family_ea or family_latin

        opens: List[str] = []
        closes: List[str] = []
        if run.link:
            opens.append(f'<a href="{_escape_attr(run.link)}" color="#0b5cad">')
            closes.append("</a>")
        if run.strike:
            opens.append("<strike>")
            closes.append("</strike>")
        if run.underline:
            opens.append("<u>")
            closes.append("</u>")
        if run.vert == "super":
            opens.append("<super>")
            closes.append("</super>")
        elif run.vert == "sub":
            opens.append("<sub>")
            closes.append("</sub>")
        open_str = "".join(opens)
        close_str = "".join(reversed(closes))

        segments = self._split_text(text, family_latin, family_ea, bold, italic)
        out: List[str] = []
        for chunk, font_name in segments:
            attrs = [f'name="{font_name}"']
            if abs(size - (para.size or 11.0)) > 0.05:
                attrs.append(f'size="{size:g}"')
            if run.color:
                attrs.append(f'color="{_css(run.color)}"')
            if run.highlight:
                attrs.append(f'backColor="{_css(run.highlight)}"')
            out.append(f"<font {' '.join(attrs)}>{_escape(chunk)}</font>")
        return open_str + "".join(out) + close_str

    def _split_text(self, text: str, family_latin: str, family_ea: str,
                    bold: bool, italic: bool) -> List[Tuple[str, str]]:
        """Split into ``(chunk, font)`` pairs, picking the CJK face per char."""
        out: List[Tuple[str, str]] = []
        buf: List[str] = []
        current: Optional[str] = None

        def flush() -> None:
            if buf:
                out.append(("".join(buf), current or self.book.resolve(family_latin, bold, italic)))
                buf.clear()

        for ch in text:
            requested = family_ea if is_cjk_char(ch) else family_latin
            font = self.book.resolve(self.book.family_for_char(ch, requested), bold, italic)
            if font != current:
                flush()
                current = font
            buf.append(ch)
        flush()
        return out

    # -- tables ------------------------------------------------------------ #

    def _table_flowable(self, table: Table, section: Section,
                        tokens: Optional[Dict[str, str]]) -> Flowable:
        data: List[List[Any]] = []
        col_of: Dict[int, int] = {}
        for r_index, row in enumerate(table.rows):
            data_row: List[Any] = []
            for c_index, cell in enumerate(row):
                if cell.spanned:
                    data_row.append("")
                    continue
                content = self._cell_flowables(cell, section, tokens)
                data_row.append(content)
            data.append(data_row)

        col_widths = list(table.col_widths) or None
        row_heights = list(table.row_heights) or None
        if row_heights is not None and all(h is None for h in row_heights):
            row_heights = None

        rl_table = RLTable(
            data,
            colWidths=col_widths,
            rowHeights=row_heights,
            repeatRows=table.header_rows,
            hAlign={"left": "LEFT", "center": "CENTER", "right": "RIGHT"}.get(table.align, "LEFT"),
            splitByRow=1,
        )
        rl_table.setStyle(TableStyle(self._table_commands(table)))
        return rl_table

    def _cell_flowables(self, cell: Cell, section: Section,
                        tokens: Optional[Dict[str, str]]) -> List[Flowable]:
        flowables: List[Flowable] = []
        for block in cell.blocks:
            flowables.extend(self._block_to_flowables(block, section, tokens))
        if not flowables:
            style = ParagraphStyle(name="d2p-cell", fontName=self.book.resolve("latin-sans"),
                                   fontSize=10, leading=12)
            flowables = [RLParagraph("", style)]
        return flowables

    def _table_commands(self, table: Table) -> List[Tuple[Any, ...]]:
        commands: List[Tuple[Any, ...]] = []
        top, left, bottom, right = table.cell_margins
        commands.append(("TOPPADDING", (0, 0), (-1, -1), top))
        commands.append(("BOTTOMPADDING", (0, 0), (-1, -1), bottom))
        commands.append(("LEFTPADDING", (0, 0), (-1, -1), left))
        commands.append(("RIGHTPADDING", (0, 0), (-1, -1), right))
        commands.append(("VALIGN", (0, 0), (-1, -1), "TOP"))

        for r_index, row in enumerate(table.rows):
            for c_index, cell in enumerate(row):
                if cell.spanned and cell.grid_span == 0:
                    continue
                row_end = min(r_index + max(cell.row_span, 1) - 1, len(table.rows) - 1)
                if cell.grid_span > 1 or cell.row_span > 1:
                    commands.append(("SPAN", (c_index, r_index),
                                     (c_index + max(cell.grid_span, 1) - 1, row_end)))
                if cell.shading:
                    commands.append(("BACKGROUND", (c_index, r_index), (c_index + max(cell.grid_span, 1) - 1, row_end),
                                     _rl_color(cell.shading)))
                if cell.v_align and cell.v_align != "top":
                    commands.append(("VALIGN", (c_index, r_index),
                                     (c_index + max(cell.grid_span, 1) - 1, row_end),
                                     cell.v_align.upper()))
                for edge, command, start, stop in (
                    ("top", "LINEABOVE", (c_index, r_index), (c_index + max(cell.grid_span, 1) - 1, r_index)),
                    ("bottom", "LINEBELOW", (c_index, row_end), (c_index + max(cell.grid_span, 1) - 1, row_end)),
                    ("left", "LINEBEFORE", (c_index, r_index), (c_index, row_end)),
                    ("right", "LINEAFTER", (c_index + max(cell.grid_span, 1) - 1, r_index),
                     (c_index + max(cell.grid_span, 1) - 1, row_end)),
                ):
                    border = cell.borders.get(edge)
                    if border is None or border.style == "none" or border.size <= 0:
                        continue
                    dash = None
                    if border.style == "dashed":
                        dash = (2.0, 2.0)
                    elif border.style == "dotted":
                        dash = (0.6, 1.8)
                    commands.append((command, start, stop, border.size,
                                     _rl_color(border.color), None, dash))
        return commands

    # -- images ------------------------------------------------------------ #

    def _image_flowable(self, block: ImageBlock, section: Section) -> Flowable:
        width, height = block.width, block.height
        max_width = max(section.content_width - 1.0, 36.0)
        max_height = max(section.page_height - section.margin_top - section.margin_bottom - 12.0, 36.0)
        if width > max_width:
            height *= max_width / width
            width = max_width
        if height > max_height:
            width *= max_height / height
            height = max_height
        image = Image(io.BytesIO(block.data), width=width, height=height)
        image.hAlign = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}.get(block.align, "LEFT")
        return image

    # -- header / footer --------------------------------------------------- #

    def _make_drawer(self, sections: Sequence[Section]):
        furniture = [(list(section.header), list(section.footer)) for section in sections]
        if not any(header or footer for header, footer in furniture):
            return None
        page_section = self._page_section

        def drawer(canv, total_pages):
            page_number = canv.getPageNumber()
            index = page_section.get(page_number, 0)
            index = index if 0 <= index < len(sections) else 0
            section = sections[index]
            header, footer = furniture[index]
            tokens = {"{{PAGE}}": str(page_number), "{{NUMPAGES}}": str(total_pages)}
            width = max(section.content_width, 36.0)
            canv.saveState()
            try:
                if header:
                    flowables = self._build_story(header, section, tokens)
                    height = _measure(flowables, width)
                    available = max(section.margin_top - section.header_distance, 12.0)
                    height = min(height, available)
                    y = section.page_height - section.header_distance - height
                    frame = Frame(section.margin_left, y, width, height,
                                  leftPadding=0, rightPadding=0, topPadding=0,
                                  bottomPadding=0, showBoundary=0)
                    frame.addFromList(flowables, canv)
                if footer:
                    flowables = self._build_story(footer, section, tokens)
                    height = _measure(flowables, width)
                    available = max(section.margin_bottom - section.footer_distance, 12.0)
                    height = min(height, available)
                    frame = Frame(section.margin_left, section.footer_distance, width, height,
                                  leftPadding=0, rightPadding=0, topPadding=0,
                                  bottomPadding=0, showBoundary=0)
                    frame.addFromList(flowables, canv)
            finally:
                canv.restoreState()

        return drawer


def _measure(flowables: Sequence[Flowable], width: float) -> float:
    total = 0.0
    for flowable in flowables:
        try:
            _, height = flowable.wrap(width, 10_000)
        except Exception:
            height = 0.0
        total += height
    return total


# --------------------------------------------------------------------------- #
# Convenience
# --------------------------------------------------------------------------- #


def render_model(model: DocModel, out_path: str, book: Optional[FontBook] = None,
                 font_dirs: Optional[Sequence[str]] = None,
                 title: Optional[str] = None) -> RenderStats:
    renderer = PdfRenderer(book=book, font_dirs=font_dirs)
    return renderer.render(model, out_path, title=title)
