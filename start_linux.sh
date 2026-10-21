#!/usr/bin/env bash
# Starts CHINT ETM MDM after pulling the latest version from GitHub.
# Without internet or on any update error it starts the version already here.
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT" || exit 1
LOG="$ROOT/.venv/update.log"

if git fetch --quiet origin main >"$LOG" 2>&1 && git checkout --quiet -B main origin/main >>"$LOG" 2>&1; then
    WANT="$(sha256sum requirements.txt | cut -d' ' -f1)"
    if [ "$WANT" != "$(cat .venv/requirements.sha256 2>/dev/null)" ]; then
        .venv/bin/python -m pip install --quiet -r requirements.txt >>"$LOG" 2>&1 && echo "$WANT" > .venv/requirements.sha256
    fi
fi

# Move onto the ready-made builds from GitHub releases: they ask before updating
# and can roll back. The installed build takes over the desktop and menu icons.
STATE="${XDG_DATA_HOME:-$HOME/.local/share}/chint-etm-mdm/state.json"
if [ ! -f "$STATE" ]; then
    notify-send "CHINT ETM MDM" "Устанавливаю новую версию программы, это займёт пару минут" 2>/dev/null || true
fi
if EXE="$(PYTHONPATH="$ROOT/src" .venv/bin/python -m chint_etm_mdm.updater --install-latest 2>>"$LOG")" && [ -x "$EXE" ]; then
    exec "$EXE" "$@"
fi
PYTHONPATH="$ROOT/src" exec .venv/bin/python run_app.py "$@"
