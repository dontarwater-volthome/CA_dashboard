#!/usr/bin/env bash
set -euo pipefail
msg="${1:-Codex update}"

git checkout deploy
git pull --rebase origin deploy

git add -A
git commit -m "$msg" || { echo "Nothing to commit."; exit 0; }

git push origin deploy
