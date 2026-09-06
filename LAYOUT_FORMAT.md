# The slide layout format (`slides.xml`)

The native, editable `.pptx` for Google Slides is built from a **layout file** —
a small XML document the agent writes (and a person can hand-edit), one per deck (e.g. `slides.xml` in the
deck's directory). [`layout_to_pptx.py`](layout_to_pptx.py) maps **each declared
element onto exactly one real `python-pptx` shape**: real text boxes, figures as
image objects, all editable downstream.

**What-you-write-is-what-you-get.** You *state* geometry and style; the renderer
never *computes* layout. Every output shape traces back to a declared element —
no silent drops, no surprise slide splits (the
[fidelity check](check_fidelity.py) asserts this). The vocabulary is the native
PowerPoint object model, mapped ~1:1 onto `python-pptx` — a thin transparent
layer, not a DSL. You can reach any setting a person clicking around PowerPoint
could.

## Why XML (not HTML)

XML via the stdlib `xml.etree.ElementTree` — dependency-free, and the nested
element/attribute model maps directly onto the object model (slide → shapes →
text frame → paragraphs → runs). It is strict (every tag closes; no void-element
or unclosed-tag ambiguity), so a malformed deck is a parse error, not a guess.
HTML's `html.parser` would be lower-level and lossier for this attribute-heavy
nested structure.

## Document shape

```xml
<?xml version="1.0" encoding="utf-8"?>
<deck font="Calibri" title="Q3 review" author="Platform team">  <!-- font, and the file's document properties; all optional -->
  <slide>                              <!-- one slide; a blank PowerPoint canvas -->
    <title>Heading</title>             <!-- a text box; see elements below -->
    <bullets x="0.7" y="1.4" w="8.6" h="3.4" size="15">
      <li>A point with <b>bold</b> and a <a href="…">link</a>.</li>
      <li level="1">A sub-point (indented).</li>
    </bullets>
    <image src="figures/x.png" x="1.0" y="1.5" w="7"/>
  </slide>
</deck>
```

## Geometry

Every shape takes `x` `y` `w` `h` in **inches** (floats), measured from the top-left
of the 10×5.625in 16:9 canvas. Omit them and the element falls back to a sensible
named-layout default (a title box up top, a body box below). State them and the
shape lands exactly there. There is no auto-flow — two shapes with the same box
overlap, by design.

## Elements (each → one shape)

| Element | Shape | Key attributes |
| --- | --- | --- |
| `<title>` | text box (bold, left, 20pt by default) | `x y w h size color align bold outline` |
| `<text>` | text box; one paragraph, or several `<p>` children | `x y w h size color align anchor outline` |
| `<caption>` | small muted italic note (figure caption / side annotation) | `x y w h size color align anchor italic` |
| `<source>` | small muted provenance link back to the slide's data | `href` (or `run` + `base`) `label size color align x y w h` |
| `<bullets>` | bulleted text box; each `<li>` is a paragraph | `x y w h size color align bullet indent space-after` |
| `<image>` | picture (figure as an image object) | `src x y w h fit` |
| `<shape>` | autoshape (rect/oval/line…) with optional text | `prst x y w h fill line line-width radius` + text attrs |
| `<turn>` | transcript card (1 shape: top-left label + mono body) | `actor`(req) `tag label mono accent-stripe x y w h` |
| `<table>` | native pptx table (1 graphic frame) | `x y w cols align size color row-height header band rules rule-color valign` |
| `<columns>`/`<column>` | container only (emits no shape) | `<columns> x y w h gap`; `<column> w` |

- **`<li>` nesting** — `level="N"` indents (N≥0); each level adds a hanging indent
  so wrapped lines align under the text.
- **`<bullets bullet=…>`** — `•` (default) or any literal char; `none` to suppress;
  `arabic` / `arabic-paren` / `alpha` / `roman` to auto-number. A `<li>` can
  override its own `bullet`.
- **`<caption>`** — a `<text>` box pre-styled as the house caption: small (12pt),
  muted grey, italic. Use for figure captions and side notes; don't wrap a plot in
  a heavy filled `<shape>` card. `<p>` paragraphs + the full inline vocab compose on
  top; everything is overridable (`size`, `color`, `italic="false"`, and
  `<code italic="false">` to keep an identifier upright). The look lives in
  [`slides_look.py`](slides_look.py) (`CAPTION_*`).
- **`<source>`** — a small, muted provenance link from a data slide back to where
  its data lives (a report, notebook, or run directory). State `href=` (a full URL),
  or `run="<path>"` to be joined onto `base=` / `SOURCE_BASE_URL`; it renders a short
  `data ↗` link (bottom-left by default). `label=` (or inner text) and geometry are
  overridable. The scheme lives in [`slides_look.py`](slides_look.py) (`SOURCE_*`).
  **Add one to every slide that shows data** (see SKILL.md).
- **`<image fit=…>`** — `contain` (default) keeps the aspect ratio inside the box
  (state just `w`, height follows the figure); `stretch` fills the box exactly.
  A relative `src` resolves against the layout file's directory.
- **`radius="<inches>"`** — corner radius of a rounded-rectangle `<shape>`.
  Rounded-rects default to a small radius (`look.SHAPE_RADIUS_IN`, ~0.05") so text
  doesn't clip into a fat corner; set `radius` to override. Other autoshapes (arrows)
  ignore it — their adjustment is geometry, not a corner.
- **`outline="none"`** — opt a text box out of the faint editing-bounds outline
  (use it on title-slide boxes for a clean look).
- **`line-spacing=`** (on `<title>`/`<text>`, or a `<p>`) — a **plain float**
  (`"1.15"`) is a *multiple* of the line height (`spcPct`); a **point value**
  (`"40pt"`) is an *exact* baseline-to-baseline distance (`spcPts`). Google Slides
  scales `spcPct` off the font's own metrics, so a big bold title crowds at 100% —
  so **`<title>` defaults to exact points at `look.TITLE_LINE_SPACING_FACTOR`×
  the font size** (1.25 → 40pt on a 32pt title) and a wrapped 2-line title never
  collides; `<text>` defaults to 1.0. State `line-spacing=` to override either.
- **`<slide background="RRGGBB">`** — a solid slide fill (rare).
- **`<turn>`** — a transcript dialogue card; emits **one** rounded-rect with a
  small top-left context **label** above a monospace **body**. `actor` (required:
  `user`/`agent`/`main`/`leaf`/`result`) selects the per-speaker fill/border/accent
  from the card family in [`slides_look.py`](slides_look.py). The label is
  auto-generated as `actor[ · TAG]` (`tag="draft"` → `agent · DRAFT`); `label=`
  overrides it wholesale (e.g. `label="Agent A · leaf"`). The body is the element's
  inline content (full inline vocab + `<br/>` for multiple lines) and renders mono
  by default — `mono="false"` opts back to the deck sans for prose-heavy turns.
  The card look (`CARD_*` in `slides_look.py`) is shared, so transcript cards and
  any image-side code/diff cards you render read as one object.

## Tables

`<table>` emits **one** native pptx table — a real table in Google Slides (drag a
column, add a row), not a grid of text boxes. `<tr>`/`<td>` are containers, so the
one-element-one-shape rule (and the fidelity check) still holds.

```xml
<table x="0.4" y="1.0" w="9.2" cols="2.6,1.8,1.6,3.2"
       align="left,center,center,center" size="13" row-height="0.55" rules="rows">
  <tr><td>environment</td><td>replicas</td><td>autoscale</td><td>database</td></tr>
  <tr><td>staging</td><td color="555555">2</td><td color="A23A3A">off</td><td>snapshot of production</td></tr>
  <tr><td>production</td><td color="555555">6–24</td><td color="1A3A6B">on</td><td>primary + two replicas</td></tr>
</table>
```

- **Row contrast is stated, not inherited.** The renderer strips the Office theme's
  blue banded table style and writes every fill and border itself, so the look is the
  `look.TABLE_*` knobs. Two composable ways to separate rows:
  - **`band="RRGGBB"`** — zebra fill on alternate body rows (default
    `look.TABLE_BAND_HEX`); `band="none"` off.
  - **`rules="none|header|rows"`** — horizontal lines. `header` (the default) rules
    under the header only; `rows` rules under every row but the last. `rule-color=`
    overrides the line colour.
- **`cols=`** — comma list of column widths, absolute inches or `%` of `w`; omit to
  split `w` evenly. **`row-height=`** (inches) applies to every row; a `<tr h=…>`
  overrides its own. The frame's height is the sum, so you don't state `h`.
- **`align=`** — one alignment for all columns, or a comma list of one per column.
  A `<td align=…>` still wins for its own cell.
- **`header="false"`** — no header row (row 0 is body: not bold, banded normally).
- **`<td>`** takes `align color size bold fill valign` and the full inline vocabulary
  (`<p>` paragraphs too), exactly like a `<text>` box. `fill` also works on a `<tr>`
  to override the band for one row.
- Ragged rows are a hard error, and there is no cell merging — a table here is a
  plain grid. Reach for `<shape>` cards when you need a merged or irregular one.

## Inline runs (rich text inside any paragraph)

Mix text with inline tags; each formatting span becomes its own run, and **nesting
composes** (`<i>see <b>this</b> now</i>` is italic throughout, bold on `this`):

| Tag | Effect |
| --- | --- |
| `<b>` | bold |
| `<i>` | italic |
| `<code>` | monospace (Consolas) |
| `<a href="…">` | a real, clickable hyperlink |
| `<muted>` | down-sized grey italic (the right-aligned owner tag) |
| `<run …>` | an explicitly-styled run (`bold italic color size font`) |
| `<br/>` | a line break — rendered as a real paragraph break, so it survives the Google Slides import |

Any inline tag also accepts `color` / `size` / `bold` / `italic` to override.
**Whitespace is normalised like HTML inline content**: source indentation (a
newline + spaces, e.g. an `<a>` on its own line) collapses away; single
inter-word spaces are preserved. So you can wrap the XML for readability without
leaking indentation into the slide.

## Columns

`<columns>` lays its `<column>` children into vertical bands. You state each
column's `w` (a `%` of the band, or absolute inches) and an optional `gap`; the
renderer derives **only each band's x offset** from those declared widths — it
never reflows content. A child shape with its own `x` stays put; otherwise it
inherits its column's band as its geometry default. (This is the one derived
number in the whole format, and it comes straight from widths you declared.)

## Build & check

```bash
SKILL=<path to this skill dir>
bash $SKILL/build.sh decks/my-deck          # decks/my-deck/slides.xml -> decks/my-deck/my-deck.pptx
uv run $SKILL/layout_to_pptx.py slides.xml --out deck.pptx   # or the renderer directly

# fidelity: element-presence (the gate) + PNG render (the eyeball check)
uv run $SKILL/check_fidelity.py slides.xml --pptx deck.pptx --png-dir /tmp/v   # then look at /tmp/v/slide-*.png
```

The look knobs (border, slide numbers, default sizes/colours, text-box outline)
live in [`slides_look.py`](slides_look.py) — one source of truth.
