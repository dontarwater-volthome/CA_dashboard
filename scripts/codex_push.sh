#!/usr/bin/env bash
set -euo pipefail
msg="${1:-Codex update}"

# Always operate on deploy
git checkout deploy

# If there are local changes, stash them temporarily so pull/rebase can run
STASHED=0
if ! git diff --quiet || ! git diff --cached --quiet; then
  git stash push -u -m "codex_push_autostash"
  STASHED=1
fi

# Update branch first
git pull --rebase origin deploy

# Restore changes if we stashed them
if [ "$STASHED" -eq 1 ]; then
  git stash pop || true
fi

# Commit & push
git add -A
git commit -m "$msg" || { echo "Nothing to commit."; exit 0; }
git push origin deploy
