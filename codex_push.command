#!/usr/bin/env bash
set -euo pipefail

MSG=$(osascript -e 'display dialog "Commit message:" default answer "StreamDeck deploy" buttons {"Cancel","OK"} default button "OK"' \
      -e 'text returned of result')

cd /Users/furbo33/CA_dashboard
./scripts/codex_push.sh "$MSG"
