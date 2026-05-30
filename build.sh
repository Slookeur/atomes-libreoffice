#!/usr/bin/env bash
# build.sh — Package the atomes LibreOffice extension into an .oxt file

rm -f *.oxt

set -euo pipefail
OUT="atomes_extension.oxt"
echo "=== atomes LibreOffice Extension — Build ==="

# 2. Package as .oxt
echo "→ Building ${OUT} …"
rm -f "$OUT"
zip -r "$OUT" META-INF/ description.xml Addons.xcu Jobs.xcu icons/ python/ pkg-description/ 
echo ""
echo "✓ Extension built: ${OUT}"
echo "Install: Tools → Extension Manager → Add …"
echo "Install using command line: unopkg add atomes_extension.oxt"
echo "Remove using command line : unopkg remove atomes_extension.oxt"
