#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx>=1.0.2,<2", "typer>=0.12,<1"]
# ///
r"""Render an author-written XML layout file into a native, editable `.pptx`.

The author states the layout directly in a small XML document (boxes, text, figures,
columns, with explicit geometry/style); this renderer maps each declared element
onto exactly one real `python-pptx` shape — real text boxes, figures as image
objects, editable in Google Slides. What-you-write-is-what-you-get: every output
shape traces back to a declared element, no silent drops, no surprise splits.

Low abstraction by design. The vocabulary is the native PowerPoint object model
(slide -> shapes -> text frame -> paragraphs -> runs; fonts, sizes, colours,
positions, images), exposed as XML elements/attributes that map ~1:1 onto
python-pptx calls. It is a thin transparent layer, not a DSL — the author can
reach any setting a person clicking around PowerPoint could. The author *states*
geometry; the renderer never *computes* layout (the one offset it derives —
column children's x — comes straight from the column widths/gap the author
declared, see `_layout_columns`).

The format is documented in `LAYOUT_FORMAT.md` (this dir); the deck "look"
(border, slide numbers, text-box outlines, default sizes) lives in
`slides_look.py`.

Run with uv (dependencies are declared inline, PEP 723):
`uv run layout_to_pptx.py <layout.xml> [--out deck.pptx]`, or `bash build.sh <layout.xml>`.
"""

from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

try:
    import typer
    from pptx import Presentation
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.presentation import Presentation as PresentationType
    from pptx.slide import Slide
    from pptx.text.text import TextFrame, _Paragraph
    from pptx.util import Emu, Inches, Length
except ImportError as exc:  # a bare `python3` without the deps installed
    sys.exit(
        f"[layout] missing dependency: {exc.name}. Run via `uv run {Path(__file__).name}` "
        "(deps are declared inline) or `pip install python-pptx typer`."
    )

import slides_look as look  # noqa: E402 — sibling module, after the dependency guard

# Inline elements allowed inside a paragraph's text — each maps to a styled run.
# Anything else inside a <li>/<p> is a hard error (no silent drops).
_INLINE_TAGS = {"b", "i", "code", "a", "br", "run", "muted"}

# Alignment / anchor names accepted in the layout file -> python-pptx enums.
_ALIGN = {
    "left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
    "centre": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
_ANCHOR = {
    "top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM,
}


class LayoutError(Exception):
    """A malformed layout file. Raised eagerly so the author sees the problem
    rather than a silently-wrong deck."""


@dataclass
class Box:
    """A shape's explicit geometry, in EMU. None on any side means "fall back to
    the element's named-layout default" (resolved per element type)."""

    x: Length | None = None
    y: Length | None = None
    w: Length | None = None
    h: Length | None = None

    @classmethod
    def from_el(cls, el: ET.Element) -> Box:
        """Read x/y/w/h attributes (inches, floats) off an element."""
        def emu(name: str) -> Length | None:
            v = el.get(name)
            return Inches(float(v)) if v is not None else None
        return cls(emu("x"), emu("y"), emu("w"), emu("h"))

    def filled(self, default: tuple[float, float, float, float]) -> _FullBox:
        """Resolve missing sides from an inches default 4-tuple (x, y, w, h)."""
        dx, dy, dw, dh = (Inches(v) for v in default)
        return _FullBox(self.x or dx, self.y or dy, self.w or dw, self.h or dh)


@dataclass
class _FullBox:
    """A fully-resolved geometry (no None) ready to hand to python-pptx."""

    x: Length
    y: Length
    w: Length
    h: Length


@dataclass
class RenderReport:
    """What the render produced, for the fidelity check: a flat list of
    (slide_index, element_tag, shape_name) the renderer actually emitted, so the
    checker can assert declared elements -> shapes with zero drops."""

    shapes: list[tuple[int, str, str]] = field(default_factory=list)

    def add(self, slide_idx: int, tag: str, name: str) -> None:
        self.shapes.append((slide_idx, tag, name))


# --- text / runs -----------------------------------------------------------


@dataclass
class _RunStyle:
    """The accumulated inline style for a text fragment — the union of every
    enclosing inline tag, so `<i><b>x</b></i>` makes a bold+italic run. Carried
    down the recursion; each inline element layers its own attribute on."""

    size: int
    color: str
    bold: bool = False
    italic: bool = False
    mono: bool = False
    href: str | None = None

    def layered(self, el: ET.Element) -> _RunStyle:
        """Return a copy with this inline element's tag + attributes applied."""
        s = _RunStyle(self.size, self.color, self.bold, self.italic, self.mono,
                      self.href)
        tag = el.tag
        if tag == "b":
            s.bold = True
        elif tag == "i":
            s.italic = True
        elif tag == "code":
            s.mono = True
        elif tag == "a":
            s.href = el.get("href")
            s.color = el.get("color", "1A73E8")
        elif tag == "muted":
            s.italic = True
            s.color = el.get("color", "707070")
            s.size = int(el.get("size", max(9, self.size - 4)))
        elif tag == "run":
            pass  # all style comes from explicit attributes below
        # Any inline element may override colour/size/bold/italic explicitly.
        if (c := el.get("color")) is not None and tag != "muted":
            s.color = c
        if (sz := el.get("size")) is not None and tag != "muted":
            s.size = int(sz)
        if (b := el.get("bold")) is not None:
            s.bold = b.lower() in ("1", "true", "yes")
        if (i := el.get("italic")) is not None:
            s.italic = i.lower() in ("1", "true", "yes")
        return s


def _add_styled_run(para: _Paragraph, text: str, style: _RunStyle) -> None:
    """Append one run carrying `text` styled by `style`. Empty text is skipped
    (no zero-width runs)."""
    if not text:
        return
    run = para.add_run()
    run.text = text
    font = run.font
    font.name = look.CARD_BODY_MONO if style.mono else look.FONT_FAMILY
    font.size = look.pt(style.size)
    font.bold = style.bold
    font.italic = style.italic
    font.color.rgb = look.hex_to_rgb(style.color or look.BODY_COLOR_HEX)
    if style.href is not None:
        run.hyperlink.address = style.href   # a real, editable hyperlink object


@dataclass
class _InlineState:
    r"""Whitespace-collapsing state threaded through one paragraph's inline walk.
    XML carries the author's source indentation as text (a `<a>` on its own line
    becomes a "\n      " text node); we normalise it like HTML inline content so
    output is predictable however the author wraps the XML: newline-bearing
    whitespace collapses to a single space, and the paragraph's leading/trailing
    whitespace is dropped. `at_start` suppresses the leading space; `pending`
    holds a deferred space so a trailing one is dropped if nothing follows."""

    at_start: bool = True
    # A trailing space held back to drop it if the paragraph ends; carries the
    # style it was deferred under, so the inter-word space stays in the
    # *preceding* run's style (a space between "x " and a <b> isn't bold).
    pending_style: _RunStyle | None = None


# A run of whitespace that contains a newline = source-formatting indentation.
_NL_WS = re.compile(r"\s*\n\s*")


def _feed_text(para: _Paragraph, text: str, style: _RunStyle,
               st: _InlineState) -> None:
    """Emit `text` as styled runs, applying inline whitespace collapsing. Plain
    inter-word spaces (no newline) are author-significant and preserved; only
    newline-bearing whitespace is treated as formatting."""
    text = _NL_WS.sub(" ", text)          # collapse source indentation to a space
    if not text:
        return
    if st.at_start:
        text = text.lstrip(" ")           # no leading space on the paragraph
        if not text:
            return
    elif st.pending_style is not None:    # a deferred space, now followed by text
        _add_styled_run(para, " ", st.pending_style)
    # Defer a single trailing space: drop it if the paragraph ends here.
    st.pending_style = style if text.endswith(" ") else None
    body = text[:-1] if st.pending_style is not None else text
    if body:
        _add_styled_run(para, body, style)
        st.at_start = False


@dataclass
class _ParaSink:
    r"""The current output paragraph for one inline walk, plus how to break a line.
    A `<br/>` anywhere in the walk — even nested inside `<b>`/`<run>` — swaps
    `.para` for a fresh continuation paragraph, so every enclosing frame's tail
    text flows onto the new line. `tf=None` keeps the legacy single-paragraph soft
    break (used only where no text frame is wired in)."""

    para: _Paragraph
    tf: TextFrame | None = None

    def line_break(self) -> None:
        if self.tf is None:
            self.para.add_line_break()
        else:
            self.para = _continuation_paragraph(self.tf, self.para)


def _continuation_paragraph(tf: TextFrame, prev: _Paragraph) -> _Paragraph:
    r"""Start a new paragraph continuing `prev` — a `<br/>` line break that SURVIVES
    the Google Slides import. python-pptx's `add_line_break` emits a single-paragraph
    `<a:br/>`, which the Slides import silently drops (collapsing every `<br/>` line
    onto one); a real `<a:p>` break does not. The continuation clones prev's
    paragraph properties (alignment, line spacing, hanging indent) so it reads as the
    same block, suppresses any bullet glyph so a `<br/>` inside a `<li>` doesn't
    repeat the bullet, and tightens the seam — the line we leave drops its
    space-after and the new line its space-before — so only the block's outer spacing
    brackets the whole run of lines."""
    import copy

    from pptx.oxml.ns import qn

    para = tf.add_paragraph()
    if (src_pPr := prev._p.find(qn("a:pPr"))) is not None:
        # Clone prev's props FIRST (carries its original space-after onto this, the
        # new last line), THEN zero prev's space-after below so the gap lands only
        # after the final line — never between the lines of one block.
        if (own := para._p.find(qn("a:pPr"))) is not None:
            para._p.remove(own)
        para._p.insert(0, copy.deepcopy(src_pPr))
        pPr = para._p.find(qn("a:pPr"))
        for tag in ("a:buChar", "a:buAutoNum", "a:buNone"):
            for existing in pPr.findall(qn(tag)):
                pPr.remove(existing)
        pPr.append(_oxml('<a:buNone xmlns:a="%s"/>' % look._A))
        pPr.set("indent", "0")               # no glyph, so no hanging gutter
    para.space_before = look.pt(0)
    prev.space_after = look.pt(0)
    return para


def _emit_inline(sink: _ParaSink, el: ET.Element, style: _RunStyle,
                 st: _InlineState) -> None:
    r"""Recursively turn an element's mixed text + inline children into styled
    runs, layering each enclosing inline tag's style (so nesting composes:
    `<i>… <b>x</b> …</i>` keeps the italic across the bold span and its tail).
    `<br>` starts a fresh continuation paragraph (an import-robust line break, see
    `_continuation_paragraph`); any non-inline child tag is a hard error — nothing
    is silently dropped."""
    if el.text:
        _feed_text(sink.para, el.text, style, st)
    for child in el:
        if child.tag not in _INLINE_TAGS:
            raise LayoutError(
                f"<{child.tag}> is not a valid inline element inside "
                f"<{el.tag}>; allowed: {sorted(_INLINE_TAGS)}"
            )
        if child.tag == "br":
            sink.line_break()
            st.at_start, st.pending_style = True, None  # fresh line, no lead space
        else:
            _emit_inline(sink, child, style.layered(child), st)
        # Tail continues the (possibly new) paragraph in the *enclosing* style.
        if child.tail:
            _feed_text(sink.para, child.tail, style, st)


def _emit_runs(
    para: _Paragraph, el: ET.Element, *, size: int, base_color: str | None = None,
    mono: bool = False, tf: TextFrame | None = None,
) -> None:
    """Entry point: render one paragraph element's text + inline children to
    styled runs. Mirrors how a person types a line with mid-sentence
    formatting — each span is its own run, nesting composes. `mono=True` seeds
    the root style as monospace (a <turn> body default), so plain text renders
    in Consolas while inline tags still compose on top. Pass `tf` (the paragraph's
    text frame) so a `<br/>` can start a real continuation paragraph that survives
    the Google Slides import."""
    _emit_inline(
        _ParaSink(para, tf), el,
        _RunStyle(size=size, color=base_color or look.BODY_COLOR_HEX, mono=mono),
        _InlineState(),
    )


def _style_frame(tf: TextFrame, el: ET.Element) -> None:
    """Apply text-frame-level options: wrap, autosize off (we honour the box),
    vertical anchor, and internal margins if stated."""
    tf.word_wrap = True
    if (anchor := el.get("anchor")) is not None:
        if anchor not in _ANCHOR:
            raise LayoutError(f"anchor={anchor!r} not in {sorted(_ANCHOR)}")
        tf.vertical_anchor = _ANCHOR[anchor]
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        if (v := el.get(side.replace("_", "-"))) is not None:
            setattr(tf, side, Inches(float(v)))


def _para_align(para: _Paragraph, el: ET.Element) -> None:
    if (a := el.get("align")) is not None:
        if a not in _ALIGN:
            raise LayoutError(f"align={a!r} not in {sorted(_ALIGN)}")
        para.alignment = _ALIGN[a]


def _set_line_spacing(
    para: _Paragraph, el: ET.Element, parent: ET.Element | None = None
) -> None:
    r"""Write an explicit line-spacing on the paragraph. Without it the paragraph
    carries no `<a:lnSpc>` and inherits the theme's default spacing — which the
    LibreOffice PDF path renders at ~100% (fine) but Google Slides / PowerPoint
    render tight enough that a *wrapped* large title's second line overlaps the
    first. Forcing 100% fixes the collapse without changing the single-line look.

    Two forms of `line-spacing=` (on the element, or a `<p>`, falling back to the
    parent box's value, then 1.0):

    - a **plain float** (`"1.15"`) → a *multiple* of the line height (`spcPct`).
      Convenient, but Google Slides scales the percentage off the font's own line
      metrics, so a big bold title can still crowd at 100%.
    - a **point value** (`"40pt"`) → an *exact* baseline-to-baseline distance
      (`spcPts`). Google Slides honours it literally regardless of font metrics,
      so it's the reliable knob for a wrapped large title (pick ~1.2–1.35× the
      font size, e.g. `40pt` for a 32pt title)."""
    v = el.get("line-spacing")
    if v is None and parent is not None:
        v = parent.get("line-spacing")
    if v is None:
        para.line_spacing = 1.0
    elif v.strip().lower().endswith("pt"):
        para.line_spacing = look.pt(float(v.strip()[:-2]))   # exact points -> spcPts
    else:
        para.line_spacing = float(v)                          # multiple -> spcPct


# --- shape emitters (one per declared element) -----------------------------


def _add_textbox(slide: Slide, box: _FullBox):  # noqa: ANN201 — pptx Shape
    """A real text box at an explicit box. Used for <title>, <text>, <bullets>."""
    return slide.shapes.add_textbox(box.x, box.y, box.w, box.h)


def _render_title(slide: Slide, el: ET.Element, box: Box) -> str:
    """A title text box: left-aligned, bold, down-sized by default (the deck
    look), but every attribute is overridable. Inline <muted> carries the
    right-aligned owner tag from the existing decks."""
    full = box.filled(look.DEFAULT_TITLE_BOX_IN)
    shape = _add_textbox(slide, full)
    shape.name = "Title"
    tf = shape.text_frame
    _style_frame(tf, el)
    size = int(el.get("size", look.TITLE_SZ // 100))   # default 20pt
    para = tf.paragraphs[0]
    para.alignment = _ALIGN.get(el.get("align", "left"), PP_ALIGN.LEFT)
    # Title line spacing: an explicit line-spacing= wins; otherwise default to EXACT
    # point spacing at look.TITLE_LINE_SPACING_FACTOR x the font size (spcPts), which
    # Google Slides honours literally so a wrapped 2-line title never collides (a flat
    # 100% multiple does — see _set_line_spacing).
    if el.get("line-spacing") is not None:
        _set_line_spacing(para, el)
    else:
        para.line_spacing = look.pt(size * look.TITLE_LINE_SPACING_FACTOR)
    # Title plain text defaults to bold (the deck look); a bold="false" override
    # or any inline <b>/<i>/<muted> still composes on top via the style stack.
    bold_default = el.get("bold", "true").lower() in ("1", "true", "yes")
    root = _RunStyle(size=size, color=el.get("color", look.TITLE_COLOR_HEX),
                     bold=bold_default)
    _emit_inline(_ParaSink(para, tf), el, root, _InlineState())
    return shape.name


def _render_text(slide: Slide, el: ET.Element, box: Box) -> str:
    """A free text box — a single paragraph (or multiple <p> children)."""
    full = box.filled(look.DEFAULT_BODY_BOX_IN)
    shape = _add_textbox(slide, full)
    shape.name = el.get("name", "Text")
    tf = shape.text_frame
    _style_frame(tf, el)
    size = int(el.get("size", look.BODY_SZ_PT))
    color = el.get("color")
    paras = el.findall("p")
    if paras:   # explicit paragraphs
        for idx, p in enumerate(paras):
            para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            _para_align(para, p)
            _set_line_spacing(para, p, el)
            _emit_runs(para, p, size=int(p.get("size", size)), base_color=color,
                       tf=tf)
    else:       # the element's own text/inline children are one paragraph
        para = tf.paragraphs[0]
        _para_align(para, el)
        _set_line_spacing(para, el)
        _emit_runs(para, el, size=size, base_color=color, tf=tf)
    return shape.name


def _render_caption(slide: Slide, el: ET.Element, box: Box) -> str:
    """A small, muted, italic contextual note — the house caption look as one
    element (figure captions, side annotations beside a plot), so authors don't
    restate size/colour/italic each time. A text box otherwise: `<p>` paragraphs
    and the full inline vocabulary compose on top, and size/color/italic/align
    are all overridable (e.g. `italic="false"`, or `<code italic="false">` to
    keep an identifier upright). Emits exactly one shape, like <text>."""
    full = box.filled(look.DEFAULT_BODY_BOX_IN)
    shape = _add_textbox(slide, full)
    shape.name = el.get("name", "Caption")
    tf = shape.text_frame
    _style_frame(tf, el)
    size = int(el.get("size", look.CAPTION_SZ_PT))
    color = el.get("color", look.CAPTION_COLOR_HEX)
    italic = el.get("italic", str(look.CAPTION_ITALIC)).lower() in ("1", "true", "yes")
    # Seed italic into the root run style so the caption reads muted+italic by
    # default; inline <b>/<code>/<run> still layer on top via the style stack.
    for idx, p in enumerate(el.findall("p") or [el]):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        _para_align(para, p)
        _emit_inline(
            _ParaSink(para, tf), p,
            _RunStyle(size=int(p.get("size", size)), color=color, italic=italic),
            _InlineState(),
        )
    return shape.name


def _render_source(slide: Slide, el: ET.Element, box: Box) -> str:
    """A small, muted provenance link from a data slide back to where its data
    lives (a report, a notebook, a run directory). State href= (a full URL), or
    run= (a path) joined onto base= / look.SOURCE_BASE_URL, and the element renders
    a short corner link so every data slide points back the same way. label= (or
    inner text) and geometry are overridable. The scheme lives in slides_look.py
    (SOURCE_*). Emits exactly one shape."""
    href = el.get("href")
    if href is None:
        run = el.get("run")
        base = el.get("base", look.SOURCE_BASE_URL).rstrip("/")
        if run is None or not base:
            raise LayoutError(
                '<source> needs href=, or run= with base= (or SOURCE_BASE_URL set)'
            )
        href = f"{base}/{run.strip('/')}/{look.SOURCE_SUMMARY_FILE}"
    label = (el.text or "").strip() or el.get("label", look.SOURCE_LABEL)
    full = box.filled(look.SOURCE_BOX_IN)
    shape = _add_textbox(slide, full)
    shape.name = el.get("name", "Source")
    tf = shape.text_frame
    _style_frame(tf, el)
    para = tf.paragraphs[0]
    para.alignment = _ALIGN.get(el.get("align", "left"), PP_ALIGN.LEFT)
    _add_styled_run(para, label, _RunStyle(
        size=int(el.get("size", look.SOURCE_SZ_PT)),
        color=el.get("color", look.SOURCE_COLOR_HEX), italic=True, href=href,
    ))
    return shape.name


def _render_bullets(slide: Slide, el: ET.Element, box: Box) -> str:
    """A bulleted list. Each <li> is one paragraph; level=N indents it. A real
    bullet glyph + hanging indent is set per paragraph so spacing survives in
    Slides. The list-wide bullet style (default char, or `bullet="arabic"` for
    auto-numbering) sits on <bullets>; a <li> can override its own."""
    full = box.filled(look.DEFAULT_BODY_BOX_IN)
    shape = _add_textbox(slide, full)
    shape.name = el.get("name", "Bullets")
    tf = shape.text_frame
    _style_frame(tf, el)
    size = int(el.get("size", look.BODY_SZ_PT))
    color = el.get("color")
    list_bullet = el.get("bullet")          # list-wide default bullet style
    list_align = el.get("align")            # list-wide default alignment
    indent_in = float(el.get("indent", "0.3"))   # hanging indent per level (in)
    space_pt = float(el.get("space-after", "6"))  # gap between items (pt)
    items = el.findall("li")
    if not items:
        raise LayoutError("<bullets> must contain at least one <li>")
    for idx, li in enumerate(items):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        level = int(li.get("level", "0"))
        para.level = level
        # Per-li align wins; otherwise the list-wide default (if any).
        if li.get("align") is None and list_align is not None:
            li.set("align", list_align)
        _para_align(para, li)
        li_size = int(li.get("size", size))
        para.space_after = look.pt(space_pt)
        _set_bullet(para, li.get("bullet", list_bullet), level=level,
                    indent_in=indent_in, size=li_size)
        _emit_runs(para, li, size=li_size, base_color=li.get("color", color),
                   tf=tf)
    return shape.name


# Named auto-number schemes the layout file accepts -> OOXML buAutoNum types.
_AUTONUM = {
    "arabic": "arabicPeriod", "arabic-period": "arabicPeriod",
    "arabic-paren": "arabicParenR", "alpha": "alphaLcPeriod",
    "alpha-upper": "alphaUcPeriod", "roman": "romanLcPeriod",
}


def _set_bullet(
    para: _Paragraph, glyph: str | None, *, level: int, indent_in: float,
    size: int,
) -> None:
    r"""Set a paragraph's bullet + the hanging indent that gives it breathing
    room (a plain text box otherwise renders the glyph flush against the text).
    `bullet="none"` suppresses the glyph; `bullet="arabic"` (etc.) auto-numbers;
    any other value is taken as a literal bullet character (default •). marL is
    the left margin per level; the negative `indent` pulls the bullet back into
    the gutter so wrapped lines align under the text, not the bullet."""
    from pptx.oxml.ns import qn
    pPr = para._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buNone", "a:buAutoNum", "a:buFont"):
        for existing in pPr.findall(qn(tag)):
            pPr.remove(existing)
    # Hanging indent: marL grows per level; the gutter is ~the glyph width.
    gutter = max(int(Inches(0.28)), int(look.pt(size) * 1.6))
    marL = int(Inches(indent_in)) * (level + 1)
    pPr.set("marL", str(marL))
    pPr.set("indent", str(-gutter))
    if glyph == "none":
        pPr.append(_oxml('<a:buNone xmlns:a="%s"/>' % look._A))
        pPr.set("indent", "0")              # nothing in the gutter
        return
    if glyph in _AUTONUM:
        pPr.append(_oxml(
            '<a:buAutoNum xmlns:a="%s" type="%s"/>' % (look._A, _AUTONUM[glyph])
        ))
        return
    char = glyph or "•"
    pPr.append(_oxml('<a:buChar xmlns:a="%s" char="%s"/>' % (look._A, char)))


def _oxml(xml: str):  # noqa: ANN201 — returns an lxml element
    from pptx.oxml import parse_xml
    return parse_xml(xml)


def _render_image(
    slide: Slide, el: ET.Element, box: Box, *, base_dir: Path
) -> str:
    """A figure as a real image object. Geometry is honoured exactly; with only
    a width (or only a box width) the height follows the image aspect ratio so
    figures aren't distorted (fit="contain", the default). fit="stretch" fills
    the box. A relative src resolves against the layout file's directory."""
    src = el.get("src")
    if src is None:
        raise LayoutError("<image> requires a src attribute")
    path = (base_dir / src) if not Path(src).is_absolute() else Path(src)
    if not path.is_file():
        raise LayoutError(f"<image> src not found: {path}")
    full = Box.from_el(el)
    fit = el.get("fit", "contain")
    if fit == "stretch":
        # Both dimensions stated (or defaulted) -> fill exactly.
        resolved = full.filled(look.DEFAULT_BODY_BOX_IN)
        pic = slide.shapes.add_picture(
            str(path), resolved.x, resolved.y, resolved.w, resolved.h
        )
    elif fit == "contain":
        # State width and/or x,y; height derives from the native aspect ratio.
        x = full.x if full.x is not None else Inches(look.DEFAULT_BODY_BOX_IN[0])
        y = full.y if full.y is not None else Inches(look.DEFAULT_BODY_BOX_IN[1])
        if full.w is not None:
            pic = slide.shapes.add_picture(str(path), x, y, width=full.w)
        elif full.h is not None:
            pic = slide.shapes.add_picture(str(path), x, y, height=full.h)
        else:                                   # native size
            pic = slide.shapes.add_picture(str(path), x, y)
    else:
        raise LayoutError(f'fit={fit!r} must be "contain" or "stretch"')
    pic.name = el.get("name", f"Image:{Path(src).name}")
    return pic.name


def _render_shape(
    slide: Slide, el: ET.Element, box: Box, *, base_dir: Path
) -> str:
    """A drawn shape (rect / rounded-rect / oval / line) with optional text —
    the low-abstraction escape hatch for native chrome the named elements don't
    cover (filled call-out boxes, dividers). prst names are OOXML preset
    geometries; fill/line/text are all explicit."""
    from pptx.enum.shapes import MSO_SHAPE
    full = box.filled(look.DEFAULT_BODY_BOX_IN)
    prst = el.get("prst", "rect").upper().replace("-", "_")
    try:
        autoshape = MSO_SHAPE[prst]
    except KeyError as exc:
        raise LayoutError(f"unknown shape prst={el.get('prst')!r}") from exc
    shape = slide.shapes.add_shape(autoshape, full.x, full.y, full.w, full.h)
    shape.name = el.get("name", f"Shape:{prst}")
    fill = el.get("fill")
    if fill == "none":
        shape.fill.background()
    elif fill is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = look.hex_to_rgb(fill)
    line = el.get("line")
    if line == "none":
        shape.line.fill.background()
    elif line is not None:
        shape.line.color.rgb = look.hex_to_rgb(line)
        if (lw := el.get("line-width")) is not None:
            shape.line.width = look.pt(float(lw))
    # Corner radius for rounded-rectangles: radius= (inches) maps to the shape's
    # first adjustment (a fraction of the shorter side). python-pptx defaults to a
    # fat ~0.167 radius that clips top-anchored text into the corner, so we default
    # to a small absolute radius (look.SHAPE_RADIUS_IN) — square-ish card. Only
    # rounded-rectangles get the default; other adjustable shapes (arrows) keep
    # their adjustment, which is geometry, not a corner. radius= overrides.
    want_radius = el.get("radius")
    if want_radius is None and prst == "ROUNDED_RECTANGLE":
        want_radius = look.SHAPE_RADIUS_IN
    if want_radius is not None and len(shape.adjustments):
        from pptx.util import Inches
        shorter = min(full.w, full.h)
        adj = (Inches(float(want_radius)) / shorter) if shorter else 0.0
        shape.adjustments[0] = max(0.0, min(0.5, adj))
    # Optional text inside the shape (same inline vocabulary as <text>). Multiple
    # <p> children become separate paragraphs; an inline <br/> also splits into a
    # real continuation paragraph (see _continuation_paragraph), so both survive
    # the Google Slides import.
    if (el.text and el.text.strip()) or len(el):
        tf = shape.text_frame
        _style_frame(tf, el)
        size = int(el.get("size", look.BODY_SZ_PT))
        color = el.get("color")
        paras = el.findall("p")
        if paras:
            # Autoshapes default to centred text; the shape's own align= seeds each
            # <p> that doesn't state its own (a per-p align still wins).
            shape_align = el.get("align")
            for idx, p in enumerate(paras):
                para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                if p.get("align") is None and shape_align is not None:
                    p.set("align", shape_align)
                _para_align(para, p)
                _emit_runs(para, p, size=int(p.get("size", size)),
                           base_color=p.get("color", color), tf=tf)
        else:
            para = tf.paragraphs[0]
            _para_align(para, el)
            _emit_runs(para, el, size=size, base_color=color, tf=tf)
    return shape.name


def _render_turn(slide: Slide, el: ET.Element, box: Box) -> str:
    """A transcript card: ONE rounded-rect carrying a small top-left context
    LABEL (auto-generated from actor[+tag], or label= override) above a mono
    BODY (the message). Part of the card family (look.CARD_*); the actor selects
    the fill/border/accent so the card reads per-speaker. Emits exactly one
    shape — label + body are paragraphs inside its text frame."""
    from pptx.enum.shapes import MSO_SHAPE

    actor = el.get("actor")
    if actor is None:
        raise LayoutError("<turn> requires an actor attribute")
    try:
        style = look.card_style(actor)
    except ValueError as exc:                # translate look's ValueError
        raise LayoutError(str(exc)) from exc

    full = box.filled(look.DEFAULT_BODY_BOX_IN)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, full.x, full.y, full.w, full.h
    )
    shape.name = el.get("name", f"Turn:{actor}")
    shape.fill.solid()
    shape.fill.fore_color.rgb = look.hex_to_rgb(style.fill)
    shape.line.color.rgb = look.hex_to_rgb(style.accent)   # accent reads as the border
    shape.line.width = look.pt(style.border_pt)
    # python-pptx's default rounded-rect adjustment is a fraction of the shorter side, so a
    # tall card gets a very fat corner. `look.CARD_RADIUS_IN` is the family's absolute radius
    # (`<shape>` already resolves it this way); `radius=` overrides per card.
    radius = float(el.get("radius", look.CARD_RADIUS_IN))
    if len(shape.adjustments):
        shorter = min(full.w, full.h)
        shape.adjustments[0] = max(0.0, min(0.5, Inches(radius) / shorter if shorter else 0.0))

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP                     # label hugs the top
    pl, pt_, pr, pb = look.CARD_PAD_IN
    tf.margin_left, tf.margin_top = Inches(pl), Inches(pt_)
    tf.margin_right, tf.margin_bottom = Inches(pr), Inches(pb)

    # Label string: explicit label= wins; else actor[ · TAG].
    if (label := el.get("label")) is None:
        tag = el.get("tag")
        label = actor + (f" · {tag.upper()}" if tag else "")
    label_para = tf.paragraphs[0]
    label_para.alignment = PP_ALIGN.LEFT
    label_para.space_after = look.pt(look.CARD_LABEL_GAP_PT)
    _add_styled_run(
        label_para, label,
        _RunStyle(size=look.CARD_LABEL_SZ, color=style.accent, bold=True),
    )

    # Body: the element's inline content, mono by default. Multiple <p> children
    # become separate body paragraphs; an inline <br/> likewise splits into a real
    # continuation paragraph, so both survive the Google Slides import. Bare inline
    # content (no <p>) stays one paragraph.
    mono = el.get("mono", "true").lower() in ("1", "true", "yes")
    for src in el.findall("p") or [el]:
        body = tf.add_paragraph()
        body.alignment = PP_ALIGN.LEFT
        body.line_spacing = look.CARD_LINE_SPACING
        _emit_runs(body, src, size=look.CARD_BODY_SZ,
                   base_color=el.get("color"), mono=mono, tf=tf)
    return shape.name


_RULES = ("none", "header", "rows")


def _column_widths(value: str | None, n: int, total: float) -> list[float]:
    """`cols=` on a <table>: a comma list of column widths, each absolute inches
    or a `%` of the table's `w`. Omitted, the width splits evenly."""
    if value is None:
        return [total / n] * n
    parts = [p.strip() for p in value.split(",") if p.strip()]
    if len(parts) != n:
        raise LayoutError(f"<table> cols= has {len(parts)} entries, needs {n}")
    return [total * float(p[:-1]) / 100.0 if p.endswith("%") else float(p)
            for p in parts]


def _column_aligns(value: str | None, n: int) -> list[str]:
    """`align=` on a <table>: one alignment for every column, or a comma list of
    one per column. A `<td align=…>` still overrides its own cell."""
    parts = [p.strip() for p in (value or "left").split(",")]
    if len(parts) == 1:
        parts *= n
    if len(parts) != n:
        raise LayoutError(f"<table> align= has {len(parts)} entries, needs {n}")
    for p in parts:
        if p not in _ALIGN:
            raise LayoutError(f"align={p!r} not in {sorted(_ALIGN)}")
    return parts


def _cell_text(cell, el: ET.Element, *, size: int, color: str, bold: bool,  # noqa: ANN001
               align: str) -> None:
    """Fill one cell's text frame — `<p>` paragraphs and the full inline
    vocabulary, exactly as in a <text> box."""
    tf = cell.text_frame
    tf.word_wrap = True
    for idx, p in enumerate(el.findall("p") or [el]):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        para.alignment = _ALIGN[p.get("align", align)]
        _set_line_spacing(para, p, el)
        _emit_inline(
            _ParaSink(para, tf), p,
            _RunStyle(size=int(p.get("size", size)), color=color, bold=bold),
            _InlineState(),
        )


def _render_table(slide: Slide, el: ET.Element, box: Box) -> str:
    """A native pptx table — ONE graphic frame, editable as a real table in
    Google Slides (drag a column, add a row). `<tr>`/`<td>` are containers, so
    the one-element-one-shape rule still holds.

    Row contrast is stated, not inherited: the Office theme's blue banded style
    is stripped (look.unstyle_table) and the emitter writes every fill and border
    itself. Two knobs, composable — `band=` zebra-fills alternate body rows,
    `rules=` draws horizontal lines (`none` / `header` / `rows`)."""
    rows = el.findall("tr")
    if not rows:
        raise LayoutError("<table> must contain at least one <tr>")
    widths_per_row = {len(tr.findall("td")) for tr in rows}
    if len(widths_per_row) != 1:
        raise LayoutError(f"<table> rows are ragged: {sorted(widths_per_row)} cells")
    n_cols = widths_per_row.pop()

    full = box.filled(look.DEFAULT_BODY_BOX_IN)
    col_w = _column_widths(el.get("cols"), n_cols, full.w.inches)
    default_h = el.get("row-height", str(look.TABLE_ROW_H_IN))
    row_h = [float(tr.get("h", default_h)) for tr in rows]
    aligns = _column_aligns(el.get("align"), n_cols)

    frame = slide.shapes.add_table(
        len(rows), n_cols, full.x, full.y,
        Inches(sum(col_w)), Inches(sum(row_h)),
    )
    frame.name = el.get("name", "Table")
    table = frame.table
    look.unstyle_table(table)
    for i, w in enumerate(col_w):
        table.columns[i].width = Inches(w)
    for i, h in enumerate(row_h):
        table.rows[i].height = Inches(h)

    header = el.get("header", "true").lower() in ("1", "true", "yes")
    band = el.get("band", look.TABLE_BAND_HEX)
    rules = el.get("rules", "header")
    if rules not in _RULES:
        raise LayoutError(f"rules={rules!r} not in {list(_RULES)}")
    rule_hex = el.get("rule-color", look.TABLE_RULE_HEX)
    size = int(el.get("size", look.TABLE_SZ_PT))
    color = el.get("color", look.BODY_COLOR_HEX)
    pad_l, pad_t, pad_r, pad_b = look.TABLE_PAD_IN
    table_valign = el.get("valign", "middle")

    for r, tr in enumerate(rows):
        is_header = header and r == 0
        body_idx = r - (1 if header else 0)          # 0-based index among body rows
        zebra = band != "none" and not is_header and body_idx % 2 == 1
        row_fill = tr.get("fill") or (
            look.TABLE_HEADER_FILL if is_header else (band if zebra else "none")
        )
        # header rule under row 0; rules="rows" also rules every body row but the last
        wants_rule = (rules == "header" and is_header) or (
            rules == "rows" and r < len(rows) - 1)
        for c, td in enumerate(tr.findall("td")):
            cell = table.cell(r, c)
            cell.vertical_anchor = _ANCHOR[td.get("valign", table_valign)]
            cell.margin_left, cell.margin_right = Inches(pad_l), Inches(pad_r)
            cell.margin_top, cell.margin_bottom = Inches(pad_t), Inches(pad_b)
            fill = td.get("fill", row_fill)
            if fill == "none":
                cell.fill.background()
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = look.hex_to_rgb(fill)
            bold_default = is_header and look.TABLE_HEADER_BOLD
            bold = td.get("bold", str(bold_default)).lower() in ("1", "true", "yes")
            _cell_text(cell, td, size=int(td.get("size", size)),
                       color=td.get("color", color), bold=bold,
                       align=td.get("align", aligns[c]))
            if wants_rule:
                weight = look.TABLE_HEADER_RULE_PT if is_header else look.TABLE_RULE_PT
                look.cell_border(cell, "bottom", rule_hex, weight)
    return frame.name


# Dispatch: declared element tag -> emitter. Each returns the shape name(s) it
# created. Kept as a table so adding an element is one entry + one function.
_EMITTERS = {
    "title": _render_title,
    "text": _render_text,
    "caption": _render_caption,
    "source": _render_source,
    "bullets": _render_bullets,
    "turn": _render_turn,
    "table": _render_table,
}


def _layout_columns(
    slide: Slide, el: ET.Element, report: RenderReport, slide_idx: int,
    *, base_dir: Path,
) -> None:
    r"""Lay out `<columns>` children into vertical bands. The author states each
    column's width (% of the content band, or absolute inches) and an optional
    gap; the renderer derives only each band's *x offset* from those declared
    widths — it never reflows content. A child shape with its own x is left
    where the author put it; otherwise it inherits its column's band as its box
    default."""
    band_x = float(el.get("x", look.DEFAULT_BODY_BOX_IN[0]))
    band_y = float(el.get("y", look.DEFAULT_BODY_BOX_IN[1]))
    band_w = float(el.get("w", look.DEFAULT_BODY_BOX_IN[2]))
    band_h = float(el.get("h", look.DEFAULT_BODY_BOX_IN[3]))
    gap = float(el.get("gap", "0.3"))
    cols = el.findall("column")
    if not cols:
        raise LayoutError("<columns> must contain at least one <column>")
    # Resolve each column's width in inches (percent of band, or absolute).
    widths: list[float] = []
    for col in cols:
        w = col.get("w")
        if w is None:
            raise LayoutError("each <column> needs a w (e.g. w=\"62%\" or w=\"4\")")
        if w.endswith("%"):
            widths.append(band_w * float(w[:-1]) / 100.0)
        else:
            widths.append(float(w))
    cx = band_x
    for col, cw in zip(cols, widths):
        col_default = (cx, band_y, cw, band_h)
        for child in col:
            _emit_child(slide, child, report, slide_idx,
                        base_dir=base_dir, box_default=col_default)
        cx += cw + gap


def _emit_child(
    slide: Slide, el: ET.Element, report: RenderReport, slide_idx: int,
    *, base_dir: Path, box_default: tuple[float, float, float, float] | None = None,
) -> None:
    """Emit one declared slide-child element to a shape and record it. A
    box_default (from an enclosing column) seeds geometry the element leaves
    implicit; an element's own x/y/w/h still win."""
    tag = el.tag
    box = Box.from_el(el)
    # An enclosing column seeds any side the element didn't state.
    if box_default is not None:
        dx, dy, dw, dh = (Inches(v) for v in box_default)
        box = Box(box.x or dx, box.y or dy, box.w or dw, box.h or dh)
    if tag == "image":
        name = _render_image(slide, el, box, base_dir=base_dir)
    elif tag == "shape":
        name = _render_shape(slide, el, box, base_dir=base_dir)
    elif tag == "columns":
        _layout_columns(slide, el, report, slide_idx, base_dir=base_dir)
        return
    elif tag in _EMITTERS:
        name = _EMITTERS[tag](slide, el, box)
    else:
        raise LayoutError(
            f"<{tag}> is not a valid slide element; allowed: "
            f"{sorted({*_EMITTERS, 'image', 'shape', 'columns', 'column'})}"
        )
    # `outline="none"` opts a content box out of the faint editing outline: set a
    # transparent line now so _apply_chrome sees a line present and leaves it be.
    if el.get("outline") == "none":
        last = slide.shapes[-1]
        if last.has_text_frame:
            look.suppress_outline(last)
    report.add(slide_idx, tag, name)


def _render_slide(
    prs: PresentationType, sl_el: ET.Element, report: RenderReport, idx: int,
    *, base_dir: Path,
) -> None:
    """Render one `<slide>` onto a blank layout. We always use the blank layout
    (index 6 in the default template) and place every shape explicitly — no
    inherited placeholders, so geometry is exactly what's declared."""
    blank = prs.slide_layouts[6]    # the default template's blank layout
    slide = prs.slides.add_slide(blank)
    if (bg := sl_el.get("background")) is not None:
        # Solid slide background (rare, but it's a normal-PowerPoint knob).
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = look.hex_to_rgb(bg)
    for child in sl_el:
        _emit_child(slide, child, report, idx, base_dir=base_dir)


def render(layout_path: Path, out_path: Path) -> RenderReport:
    """Parse the XML layout file and render it to `out_path`, returning a report
    of every shape emitted (for the fidelity check). Raises LayoutError on any
    malformed input — fail loud, never drop."""
    try:
        tree = ET.parse(layout_path)
    except ET.ParseError as exc:
        raise LayoutError(f"{layout_path}: malformed XML — {exc}") from exc
    root = tree.getroot()
    if root.tag != "deck":
        raise LayoutError(f"root element must be <deck>, got <{root.tag}>")
    # Deck-wide font/colour defaults override the module knobs for this render.
    if (f := root.get("font")) is not None:
        look.FONT_FAMILY = f
    prs = Presentation()    # default template; we resize to our 16:9 canvas
    prs.slide_width = Emu(look.SLIDE_W)
    prs.slide_height = Emu(look.SLIDE_H)
    report = RenderReport()
    slides = root.findall("slide")
    if not slides:
        raise LayoutError("<deck> has no <slide> children")
    base_dir = layout_path.resolve().parent
    for idx, sl_el in enumerate(slides):
        _render_slide(prs, sl_el, report, idx, base_dir=base_dir)
    _apply_chrome(prs)
    _set_core_properties(prs, root, layout_path)
    prs.save(str(out_path))
    return report


def _set_core_properties(prs: PresentationType, root: ET.Element, layout_path: Path) -> None:
    """Replace python-pptx's template metadata (its author's name, 2013 dates, a
    'generated using python-pptx' comment) with this deck's: <deck title= author=>
    when stated, else the layout file's stem and no author."""
    now = datetime.now(timezone.utc)
    cp = prs.core_properties
    cp.title = root.get("title", layout_path.stem)
    cp.author = root.get("author", "")
    cp.last_modified_by = "research-slides-skill"
    cp.comments = ""
    cp.subject = ""
    cp.created = now
    cp.modified = now
    cp.revision = 1


def _apply_chrome(prs: PresentationType) -> None:
    """Add the deck chrome (border on master, slide numbers, faint text-box
    outlines) after content is placed — one source of truth in slides_look.py."""
    for master in prs.slide_masters:
        look.add_border(master)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.name != "Title":
                look.outline_textbox(shape)
        look.add_slide_number(slide)


def main(
    layout: Path = typer.Argument(..., help="the .xml layout file to render"),
    out: Path = typer.Option(
        None, "--out", "-o",
        help="output .pptx (default: <layout>.pptx beside the layout)",
    ),
) -> None:
    """Render an XML layout file to a native, editable .pptx."""
    if not layout.is_file():
        typer.echo(f"[layout] not found: {layout}", err=True)
        raise typer.Exit(1)
    out_path = out or layout.with_suffix(".pptx")
    try:
        report = render(layout, out_path)
    except LayoutError as exc:
        typer.echo(f"[layout] {exc}", err=True)
        raise typer.Exit(1) from exc
    n_slides = len({s for s, _, _ in report.shapes})
    typer.echo(
        f"[layout] wrote {out_path} — {len(report.shapes)} shapes "
        f"across {n_slides} slides"
    )


if __name__ == "__main__":
    typer.run(main)
