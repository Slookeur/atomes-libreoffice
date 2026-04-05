#!/usr/bin/env bash
# build.sh — Package the atomes LibreOffice extension into an .oxt file

function convert_svg
{
  SVG="${SCRIPT_DIR}/icons/atomes.svg"
  if [ -f "$SVG" ]; then
      echo "→ Converting icon …"
      if command -v rsvg-convert &>/dev/null; then
          rsvg-convert -w 26 -h 26 "$SVG" -o "${SCRIPT_DIR}/icons/atomes_26.png"
          rsvg-convert -w 16 -h 16 "$SVG" -o "${SCRIPT_DIR}/icons/atomes_16.png"
      elif command -v inkscape &>/dev/null; then
          inkscape --export-type=png --export-width=26 --export-height=26 \
                   --export-filename="${SCRIPT_DIR}/icons/atomes_26.png" "$SVG" 2>/dev/null
          inkscape --export-type=png --export-width=16 --export-height=16 \
                   --export-filename="${SCRIPT_DIR}/icons/atomes_16.png" "$SVG" 2>/dev/null
      else
          echo "  ⚠ rsvg-convert/inkscape not found — PNG icons skipped"
      fi
  fi
}

rm -f *.oxt

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT="${SCRIPT_DIR}/atomes_extension.oxt"
echo "=== atomes LibreOffice Extension — Build ==="

# convert_svg

# 2. Package as .oxt
echo "→ Building ${OUT} …"
rm -f "$OUT"
cd "$SCRIPT_DIR"
zip -r "$OUT" META-INF/ description.xml Addons.xcu icons/ python/ pkg-description/ 
echo ""
echo "✓ Extension built: ${OUT}"
echo "Install: Tools → Extension Manager → Add …"
echo "Install using command line: unopkg add atomes_extension.oxt"
echo "Remove using command line : unopkg remove atomes_extension.oxt"
