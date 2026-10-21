#!/usr/bin/env bash
# Starts CHINT ETM MDM after pulling the latest version from GitHub.
# Without internet or on any update error it starts the version already here.
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT" || exit 1
LOG="$ROOT/.venv/update.log"

if git -C "$ROOT" pull --ff-only --quiet origin main >"$LOG" 2>&1; then
    WANT="$(sha256sum requirements.txt | cut -d' ' -f1)"
    if [ "$WANT" != "$(cat .venv/requirements.sha256 2>/dev/null)" ]; then
        .venv/bin/python -m pip install --quiet -r requirements.txt >>"$LOG" 2>&1 && echo "$WANT" > .venv/requirements.sha256
    fi
fi
PYTHONPATH="$ROOT/src" exec .venv/bin/python run_app.py "$@"
