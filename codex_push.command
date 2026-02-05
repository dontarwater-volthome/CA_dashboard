#!/usr/bin/env bash
set -euo pipefail

cd /Users/furbo33/CA_dashboard
./scripts/codex_push.sh "StreamDeck push $(date '+%Y-%m-%d %H:%M:%S')"
