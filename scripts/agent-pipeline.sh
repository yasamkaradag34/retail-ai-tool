#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
REPO_DIR=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)

if [ -f "$REPO_DIR/.env.agent" ]; then
  set -a
  # shellcheck disable=SC1091
  . "$REPO_DIR/.env.agent"
  set +a
fi

cd "$REPO_DIR"
exec python3 -m automation.codex_clickup.pipeline --repo "$REPO_DIR" "$@"
