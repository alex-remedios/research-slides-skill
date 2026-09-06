# Contributing

Everything hangs off one rule: **each declared XML element becomes exactly one
pptx shape**, and the author states all geometry. Adding an element means adding an
emitter to `_EMITTERS` in `layout_to_pptx.py`, its tag to `_SHAPE_TAGS` in
`check_fidelity.py`, its look knobs to `slides_look.py`, a row in `LAYOUT_FORMAT.md`,
and a gallery slide in `example-slides.xml` — then rebuild `example-slides.pptx` /
`.pdf` (both are committed) and run `uv run ruff check`, `uv run ty check`, and
`claude plugin validate . --strict`. Keep the renderer transparent: no layout engine,
no silent drops — malformed input is a `LayoutError`.

## How the pieces fit

```mermaid
flowchart LR
  X["slides.xml<br/>author-written layout"] --> R["layout_to_pptx.py<br/>one element → one shape"]
  L["slides_look.py<br/>chrome + defaults + card styles"] --> R
  R --> P["deck.pptx<br/>native, editable"]
  X --> C["check_fidelity.py"]
  P --> C
  C --> G["element presence (gate)<br/>text fit + fonts (warn)"]
  C -->|soffice → pdftoppm| N["slide-N.png"]
  P --> E["export_pdf.py"] --> D["deck.pdf<br/>whole or page range"]
  B["build.sh"] -.-> R
  S["SKILL.md"] -. tells Claude the loop .-> X
```

## The authoring loop Claude runs

```mermaid
sequenceDiagram
  actor U as User
  participant A as Claude following SKILL.md
  participant B as build.sh and layout_to_pptx.py
  participant F as check_fidelity.py
  participant E as export_pdf.py
  U->>A: make a deck, or add a results slide
  A->>A: read the last deck's slides.xml or example-slides.xml
  A->>A: write or edit slides.xml
  loop until PASS and PNGs look right
    A->>B: build
    B-->>A: deck.pptx, N shapes across M slides
    A->>F: check --pptx deck.pptx --png-dir /tmp/v
    F-->>A: presence PASS or FAIL, text-fit warnings, PNGs
    A->>A: look at the PNGs and fix overflow and collisions
  end
  opt share a read-only copy
    A->>E: export --pages a-b
    E-->>A: section.pdf
  end
  A-->>U: deck.pptx path, slide list, shape count
```
