#!/usr/bin/env bash
# build_package.sh
# ─────────────────────────────────────────────────────────────────────────
# Creates a distributable ZIP of the GHG QA Tool for Windows deployment.
#
# What it produces:
#   dist/ghg_qa_tool_v<VERSION>.zip
#
# The ZIP contains everything an analyst needs:
#   ghg_qa_tool/
#     run_app.bat         ← double-click to start
#     setup_windows.bat   ← run once on first install
#     README.txt          ← instructions for analysts
#     requirements.txt    ← for IT / manual installs
#     app/                ← tool source code
#       streamlit_app.py
#       column_matcher.py
#       diff_engine.py
#       annotation_store.py
#       export.py
#
# Usage:
#   bash build_package.sh
#   bash build_package.sh 1.1        # specify version
#
# To include embedded WinPython (zero-dependency install):
#   1. Download WinPython portable from https://winpython.github.io/
#      (choose "dot" version — smallest, no IDE bundled)
#   2. Extract the inner python-X.X.X.amd64 folder alongside this script
#      and rename it to  python/
#   3. Run:  bash build_package.sh --embed
#      This adds the python/ folder to the ZIP (~80MB total)
# ─────────────────────────────────────────────────────────────────────────

set -euo pipefail

VERSION="${1:-1.0}"
EMBED=false
[[ "${1:-}" == "--embed" || "${2:-}" == "--embed" ]] && EMBED=true

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"
DIST_DIR="$SCRIPT_DIR/dist"
PACKAGE_NAME="ghg_qa_tool_v${VERSION}"
STAGING_DIR="$DIST_DIR/staging/$PACKAGE_NAME"
ZIP_OUT="$DIST_DIR/${PACKAGE_NAME}.zip"

echo "========================================================"
echo "  GHG QA Tool — Build Package v${VERSION}"
$EMBED && echo "  Mode: EMBEDDED (includes WinPython)" || echo "  Mode: STANDARD (requires IT Python)"
echo "========================================================"
echo

# ── Validate source ────────────────────────────────────────────────────
REQUIRED_FILES=(
    "app/streamlit_app.py"
    "app/column_matcher.py"
    "app/diff_engine.py"
    "app/annotation_store.py"
    "app/export.py"
    "run_app.bat"
    "setup_windows.bat"
    "README.txt"
    "ANALYST_GUIDE.md"
    "requirements.txt"
)

echo "Checking source files..."
for f in "${REQUIRED_FILES[@]}"; do
    if [[ ! -f "$SCRIPT_DIR/$f" ]]; then
        echo "  ❌ MISSING: $f"
        echo "  Run the build from the project root directory."
        exit 1
    fi
    echo "  ✓ $f"
done

if $EMBED; then
    if [[ ! -f "$SCRIPT_DIR/python/python.exe" ]]; then
        echo
        echo "  ❌ Embedded mode requested but python/python.exe not found."
        echo "  Download WinPython from https://winpython.github.io/"
        echo "  Extract the inner python-X.X.X.amd64/ folder and rename it to python/"
        exit 1
    fi
    echo "  ✓ python/python.exe (embedded)"
fi

# ── Build staging directory ────────────────────────────────────────────
echo
echo "Building staging directory..."
rm -rf "$STAGING_DIR"
mkdir -p "$STAGING_DIR/app"

# Top-level files
cp "$SCRIPT_DIR/run_app.bat"        "$STAGING_DIR/"
cp "$SCRIPT_DIR/setup_windows.bat"  "$STAGING_DIR/"
cp "$SCRIPT_DIR/README.txt"         "$STAGING_DIR/"
cp "$SCRIPT_DIR/ANALYST_GUIDE.md"   "$STAGING_DIR/"
cp "$SCRIPT_DIR/requirements.txt"   "$STAGING_DIR/"

# App source
cp "$SCRIPT_DIR/app/streamlit_app.py"    "$STAGING_DIR/app/"
cp "$SCRIPT_DIR/app/column_matcher.py"   "$STAGING_DIR/app/"
cp "$SCRIPT_DIR/app/diff_engine.py"      "$STAGING_DIR/app/"
cp "$SCRIPT_DIR/app/annotation_store.py" "$STAGING_DIR/app/"
cp "$SCRIPT_DIR/app/export.py"           "$STAGING_DIR/app/"

# Embedded Python (optional)
if $EMBED; then
    echo "  Copying embedded Python runtime (this may take a moment)..."
    cp -r "$SCRIPT_DIR/python" "$STAGING_DIR/python"
fi

# ── Create ZIP ─────────────────────────────────────────────────────────
mkdir -p "$DIST_DIR"
rm -f "$ZIP_OUT"

echo
echo "Creating ZIP..."
(cd "$DIST_DIR/staging" && zip -r "$ZIP_OUT" "$PACKAGE_NAME" -x "*.pyc" -x "*/__pycache__/*")

# ── Summary ────────────────────────────────────────────────────────────
ZIP_SIZE=$(du -sh "$ZIP_OUT" | cut -f1)
FILE_COUNT=$(unzip -l "$ZIP_OUT" | tail -1 | awk '{print $2}')

echo
echo "========================================================"
echo "  ✅ Build complete"
echo "  Output:     $ZIP_OUT"
echo "  Size:       $ZIP_SIZE"
echo "  Files:      $FILE_COUNT"
echo
echo "  To deploy:"
echo "    1. Send $PACKAGE_NAME.zip to the analyst"
echo "    2. They extract it anywhere (Desktop, C:\Tools, etc.)"
echo "    3. They run setup_windows.bat  (once)"
echo "    4. They run run_app.bat        (every time)"
echo "========================================================"

# Clean up staging
rm -rf "$DIST_DIR/staging"
