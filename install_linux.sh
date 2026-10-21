#!/usr/bin/env bash
# Installs CHINT ETM MDM for everyday use on Linux and adds a menu/desktop icon.
# The icon runs start_linux.sh, which pulls the latest version before starting,
# so the program updates itself. Run once: bash install_linux.sh
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_NAME="CHINT ETM MDM"

PYTHON_BIN="$(command -v python3 || command -v python || true)"
if [ -z "$PYTHON_BIN" ]; then
    echo "Python 3.10+ не найден. Установите python3 и python3-venv." >&2
    exit 1
fi
if [ ! -x "$ROOT/.venv/bin/python" ]; then
    "$PYTHON_BIN" -m venv "$ROOT/.venv"
fi
"$ROOT/.venv/bin/python" -m pip install --quiet --upgrade pip
"$ROOT/.venv/bin/python" -m pip install --quiet -r "$ROOT/requirements.txt"
sha256sum "$ROOT/requirements.txt" | cut -d' ' -f1 > "$ROOT/.venv/requirements.sha256"
chmod +x "$ROOT/start_linux.sh"

ENTRY="[Desktop Entry]
Type=Application
Name=$APP_NAME
Comment=Заполнение файлов для ЭТМ и РАЭК из базы CHINT
Exec=\"$ROOT/start_linux.sh\"
Path=$ROOT
Terminal=false
Categories=Office;"
APPS_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
mkdir -p "$APPS_DIR"
printf '%s\n' "$ENTRY" > "$APPS_DIR/chint-etm-mdm.desktop"
DESKTOP_DIR="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/Desktop")"
if [ -d "$DESKTOP_DIR" ]; then
    printf '%s\n' "$ENTRY" > "$DESKTOP_DIR/chint-etm-mdm.desktop"
    chmod +x "$DESKTOP_DIR/chint-etm-mdm.desktop"
    gio set "$DESKTOP_DIR/chint-etm-mdm.desktop" metadata::trusted true 2>/dev/null || true
fi
echo "Готово. Значок «$APP_NAME» есть в меню приложений и на рабочем столе."
