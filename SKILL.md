---
name: research-slides-skill
license: MIT
compatibility: "Needs uv (the scripts declare their dependencies inline) or Python >=3.10 with python-pptx, typer, pillow. Optional: LibreOffice (soffice) and poppler for the PNG check and PDF export; the Carlito font for a faithful Calibri render on Linux."
description: "Build, edit, and export slide decks as native, editable .pptx (opens cleanly in Google Slides and PowerPoint) from an XML layout file you write — author slides.xml, render it with python-pptx, verify every declared element became a shape, eyeball the PNG render, export to PDF. Use when asked to make/update slides, a deck, a presentation, or a .pptx/PowerPoint."
---

# research-slides-skill

Built for turning **AI research results into slides** — eval numbers, plots, verbatim
transcripts — where the deck is regenerated from fresh data each week and must stay
editable afterwards. The **toolchain** lives in this directory: `build.sh`,
`layout_to_pptx.py` (the renderer), `check_fidelity.py` (the gate), `export_pdf.py`
(PDF export), `slides_look.py` (the look), plus `LAYOUT_FORMAT.md` (the full `.xml`
vocabulary) and `example-slides.xml` (the gallery source; `example-slides.pptx`/`.pdf`
are its built form, one slide per element). Deck **data** lives wherever you keep it —
one directory per deck holding `slides.xml`, a `figures/` folder, and the rendered
`.pptx`.

## Model — one path, one source

**The native, editable `.pptx`** is built from an **XML layout file** (`slides.xml`)
that you, the agent, write, via `layout_to_pptx.py`. You *state* the layout — each declared
element becomes exactly one real pptx shape — by setting:

- boxes, text, figures, columns, tables
- sizes and positions
- fonts and colours

**What-you-write-is-what-you-get**: no silent drops, no surprise slide splits.

## Setup & build

Scripts declare their dependencies inline (PEP 723), so `uv` is the only
prerequisite and the user's project venv is never touched; LibreOffice (`soffice`)
and poppler (`pdftoppm`) are optional, for the PNG render and PDF export.
`${CLAUDE_SKILL_DIR}` is this directory.

    bash ${CLAUDE_SKILL_DIR}/build.sh decks/<deck>            # decks/<deck>/slides.xml -> decks/<deck>/<deck>.pptx
    bash ${CLAUDE_SKILL_DIR}/build.sh path/to/slides.xml out.pptx

## Working loop

1. Read the previous deck's `slides.xml` (or `example-slides.xml`) and copy its shape.
2. Author or edit `slides.xml` — scaffold first (title, dividers), fill content later.
3. Build, then run `check_fidelity.py … --png-dir /tmp/v`; act on its text-fit
   warnings and **look at the PNGs**.
4. Fix overflow, wrapped titles, and collisions; rebuild; repeat.
5. Report "N slides, M/M shapes, fidelity PASS" and the `.pptx` path.

## Authoring the `.pptx` (write `slides.xml`)

Vocabulary (see `LAYOUT_FORMAT.md` for every attribute):

- `<deck>` wraps `<slide>`s; each slide is a blank canvas you place shapes on.
- Shapes — one declared element → one shape: `<title>`, `<text>` (`<p>`
  paragraphs), `<caption>` (small muted italic figure/side note — don't box a
  plot in a heavy `<shape>` card), `<source>` (provenance link, see below),
  `<bullets>` (`<li>`, `level=` to indent, `bullet=` `•`/`none`/`arabic`),
  `<image src=… x=… y=… w=…>`, `<shape prst=…>` (autoshapes), `<turn>`
  (transcript card), `<table>` (`<tr>`/`<td>`, a real editable pptx table),
  `<columns>`/`<column w="62%">` for side-by-side bands.
- Geometry is explicit: `x y w h` in **inches** on a 10 × 5.625" canvas; omit to
  take a sensible default. You state positions — the renderer never computes layout.
- Rich runs in any paragraph, composable: `<b> <i> <code> <a href> <muted>
  <run> <br/>`.
- `outline="none"` drops a box's faint editing outline (use on title slides).

Source whitespace/indentation is collapsed like HTML, so wrap the XML for
readability freely.

## Provenance — every data slide links back to its data

Whenever a slide shows data (a plot, a rate, any number from a run), add a
`<source href="…"/>` so a click goes from slide back to the report, notebook, or
run directory it came from. It renders a small muted `data ↗` link (bottom-left by
default). Set `SOURCE_BASE_URL` in `slides_look.py` to use the shorter
`<source run="<path>"/>` form.

## Pick the rendering path

A slide is a flat bag of absolutely-placed shapes — no layout engine, no flow, no
syntax highlighting (`<table>` is the one grid primitive). Choose by
**editability**, **relational structure**, and **text-flow needs**:

- **Native shapes** (`<text>`/`<bullets>`/`<shape>`/`<table>`/short `<turn>`) — the
  slide's *argument*: editable, crisp at any zoom. **Simple graphics too**: rounded
  boxes + **block-arrow autoshapes** (`prst="down-arrow"`, …) + text = a flow
  editable in Slides, no image. Block arrows only (no thin connectors), no rotation.
- **Diagram → image** — anything **complex or relational** (multi-branch flows,
  architecture, timelines, charts): let graphviz/matplotlib place it, embed the PNG.
- **HTML → screenshot** — content needing a real **flow engine + typography** (long
  code, diffs, highlighting, wrapping tables); embed the screenshot as one image.

Tie-breakers: edited live → native · a few boxes + arrows → native · many/branching
edges → diagram · needs wrap/mono-alignment/highlighting → HTML ·
**>~8–12 hand-placed shapes for a faithful render → image, always**.

## Word budget — ≤50 facing words per slide

A slide carries **≤50 facing words**: the prose the audience reads as the slide's
argument — title, subtitle, box labels, captions. Over budget → cut, split, or move
to speaker notes. **Data isn't counted**: a verbatim artifact shown *for inspection*
(a prompt, a transcript excerpt, code, raw output) is data — keep it in full and count
only the framing around it.

## Slide rhetoric — disclose gradually, lead with data

- **Data over narrative.** A slide's payload is an *artifact* — a diagram, a worked
  example, a stat, a plot — not prose. Narrative is the ≤50-word frame around it.
- **Gradual disclosure.** Open a topic with a **lead-in** slide (one concrete
  example), then reveal **one** takeaway per slide. Don't front-load everything.
- **Caption-free results slides** — title + figure + `<source>` is usually enough;
  size the image to fit (aspect × available height), not `w="9"` by reflex.

## Recurring patterns — worked examples

`example-slides.xml` is a gallery of verified slide bodies, one per element or
pattern, each naming the elements it uses — copy a `<slide>` and adapt: **title
slide** · **section divider** · **`<text>` + inline runs** · **`<bullets>`** (levels,
numbering) · **native flow** (boxes + `down-arrow` autoshapes) · **tier ladder** ·
**`<image>` results slide** (+ `<caption>` + `<source>`) · **data-led intro** ·
**`<columns>` compare** · **`<table>`** · **`<turn>` cards**. The built
`example-slides.pptx` / `.pdf` sit beside it. Rebuild after editing:

    bash ${CLAUDE_SKILL_DIR}/build.sh ${CLAUDE_SKILL_DIR}/example-slides.xml

## Figures

Render to PNG in the deck's `figures/`, size at the source, then
`<image src="figures/x.png" w="…"/>` (`fit="contain"` keeps aspect ratio):

- Graphviz: `dot -Tpng -Gdpi=200 figures/x.dot -o figures/x.png`.
- matplotlib scripts live beside the deck (e.g. `render/`); **read the numbers from
  the source data**, never hardcode them, so the figure regenerates with the data.
- Oversize chart fonts — charts sit in narrow columns/strips.
- **Author at ~16:9**; a near-square figure hits the height limit before it fills the
  width, and a very wide one leaves dead space below.
- Pick the encoding that foregrounds the signal (e.g. a dot plot when most series
  pin to one extreme and grouped bars become a wall).

## Verify fidelity — never trust the build blind

    uv run ${CLAUDE_SKILL_DIR}/check_fidelity.py slides.xml --pptx <deck>.pptx --png-dir /tmp/v

- **Pass `--pptx`** — `build.sh` names the file after the deck dir, so the check's
  default (`slides.pptx`) won't exist.
- **Element presence** asserts every declared element maps to a shape (no drops);
  non-zero exit on failure, so it works as a build gate.
- **Text fit** warns when a box's text needs more height than it has (a wrapped
  title spilling into the subtitle) — a glyph-width heuristic, so treat it as a
  prompt to shorten, shrink, or give headroom; `--strict` makes it fail the build.
  It also warns when a deck font isn't installed here (see Gotchas).
- **PNG render** (LibreOffice → poppler) is for a visual look. Read the PNGs.

## Export to PDF (one slide per page)

For a shareable read-only copy of the deck or a section (pages are 1-indexed and
match slide order):

    uv run ${CLAUDE_SKILL_DIR}/export_pdf.py <deck>.pptx                       # whole deck -> <deck>.pdf
    uv run ${CLAUDE_SKILL_DIR}/export_pdf.py <deck>.pptx --pages 4-10 --out section.pdf
    uv run ${CLAUDE_SKILL_DIR}/export_pdf.py <deck>.pptx --max-mb 1            # rasterize if over the cap

LibreOffice PDFs are heavy (font embedding, ~1 MB/page); `--max-mb` rasterizes at
150 DPI (still legible) when the file must be small. Rendered exports are
regenerable — hand over the path rather than committing them.

## The look — `slides_look.py`

The chrome (border, slide numbers, text-box outlines) + default sizes/colours/font
live in `slides_look.py` — one source of truth. Tweak the look there.

## Gotchas

- **No `--` (double hyphen) inside `<!-- -->` comments** — illegal XML; the build
  fails with "malformed XML". Write an em dash instead.
- **16:9 geometry** — the slide is **10 × 5.625"** (not 7.5): centre against 5.625,
  expect long `<title>`s to wrap into the body (shorten, drop the size, or leave
  headroom).
- **Google Slides import** — `<br/>` is emitted as a real paragraph break and titles
  get exact-point line spacing because Slides drops `<a:br/>` and crowds wrapped
  titles; `Consolas` isn't a Slides font, so mono text uses `Courier New`.
- **Font substitution in the local render** — PowerPoint and Google Slides have
  Calibri; a Linux box usually doesn't, and LibreOffice falls back to a wider font,
  so the PNG/PDF wraps earlier than the real viewer will. `check_fidelity.py` warns;
  install Carlito (metric-compatible; `fonts-crosextra-carlito`) to render faithfully.
- **`spare-parts.xml`** — keep cut-but-maybe-return slides in a sibling file; the
  deck only builds from `slides.xml`. Preview it with an explicit out path so it
  doesn't clobber the deck `.pptx`.
- **Tables are plain grids** — no cell merging; use `<shape>` cards for irregular
  layouts. Ragged rows are a hard error.
- **Viewing a `.pptx` locally** — VS Code's Office Viewer extension
  (`cweijan.vscode-office`) renders it in the editor; otherwise use the PNGs from
  `check_fidelity.py` or upload to Google Slides.
