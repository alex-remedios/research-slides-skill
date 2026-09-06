#!/usr/bin/env bash
# Render an XML layout file to a native, editable .pptx (python-pptx), via
# layout_to_pptx.py — every shape traces back to a declared element.
#
# Pass a deck directory (uses slides.xml inside) or the layout file directly,
# and optionally an output path:
#   bash build.sh decks/2026-01-14-review            # -> decks/2026-01-14-review/2026-01-14-review.pptx
#   bash build.sh decks/2026-01-14-review/slides.xml out.pptx
#
# Dependencies are declared inline in the scripts (PEP 723), so `uv run` fetches
# them; without uv, `pip install python-pptx typer pydantic` and it falls back to
# python3.
set -euo pipefail

usage() { echo "usage: $0 <deck-dir | layout.xml> [out.pptx]" >&2; exit 2; }

[ $# -ge 1 ] && [ $# -le 2 ] || usage
input=$1
# A deck dir resolves to its layout source.
[ -d "$input" ] && input="${input%/}/slides.xml"
[ -f "$input" ] || { echo "[slides] not found: $input" >&2; exit 1; }
# A slides.xml is named after its deck directory so a shared/downloaded file is
# self-identifying (2026-01-14-review.pptx, not slides.pptx); any other layout
# file keeps its own stem (example-slides.xml -> example-slides.pptx).
deck_dir=$(cd -- "$(dirname -- "$input")" && pwd -P)
stem=$(basename "${input%.xml}")
[ "$stem" = slides ] && stem=$(basename "$deck_dir")
output=${2:-$deck_dir/$stem.pptx}
here=$(cd -- "$(dirname -- "$0")" && pwd -P)

if command -v uv >/dev/null 2>&1; then
  uv run --quiet "$here/layout_to_pptx.py" "$input" --out "$output"
else
  python3 -c "import pptx, typer" 2>/dev/null \
    || { echo "[slides] no uv, and python3 lacks python-pptx/typer: pip install python-pptx typer" >&2; exit 1; }
  python3 "$here/layout_to_pptx.py" "$input" --out "$output"
fi

echo "[slides] wrote $output"
