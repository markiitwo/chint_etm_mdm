#!/usr/bin/env bash
# Builds the portable Linux version: dist/CHINT ETM MDM and dist/CHINT_ETM_MDM_portable_linux.tar.gz
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

APP_NAME="CHINT ETM MDM"
VENV_DIR="$ROOT/.venv"
VENV_PYTHON="$VENV_DIR/bin/python"
DIST_DIR="$ROOT/dist"
APP_DIR="$DIST_DIR/$APP_NAME"
ARCHIVE_PATH="$DIST_DIR/CHINT_ETM_MDM_portable_linux.tar.gz"

if [ -x "$VENV_PYTHON" ]; then
    echo "Local venv already exists: $VENV_DIR"
else
    echo "Creating local Python venv in .venv..."
    PYTHON_BIN="$(command -v python3 || command -v python || true)"
    if [ -z "$PYTHON_BIN" ]; then
        echo "Python was not found. Install Python 3.10+ to build the portable app." >&2
        exit 1
    fi
    "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

echo "Updating build tools..."
"$VENV_PYTHON" -m pip install --upgrade pip
"$VENV_PYTHON" -m pip install -r requirements.txt pyinstaller

echo "Building portable app..."
"$VENV_PYTHON" -m PyInstaller \
    --noconfirm \
    --clean \
    --windowed \
    --onedir \
    --name "$APP_NAME" \
    --paths src \
    --add-data "src/chint_etm_mdm/rules/attribute_mappings.json:chint_etm_mdm/rules" \
    --add-data "src/chint_etm_mdm/partners/*.json:chint_etm_mdm/partners" \
    run_app.py

if [ ! -x "$APP_DIR/$APP_NAME" ]; then
    echo "Build finished without the expected program: $APP_DIR/$APP_NAME" >&2
    exit 1
fi

echo "Creating portable archive..."
rm -f "$ARCHIVE_PATH"
tar -czf "$ARCHIVE_PATH" -C "$DIST_DIR" "$APP_NAME"

echo ""
echo "Done."
echo "App folder: $APP_DIR"
echo "Portable archive: $ARCHIVE_PATH"
