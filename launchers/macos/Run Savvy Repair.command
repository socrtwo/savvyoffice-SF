#!/usr/bin/env bash
# Savvy Repair for Microsoft Office — macOS launcher
set -e
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
WEB="$DIR/web"

if command -v python3 >/dev/null 2>&1; then
  (sleep 1 && open "http://localhost:8765/") &
  cd "$WEB"
  exec python3 -m http.server 8765
fi

echo "Python 3 not found. Opening the app file directly instead…"
open "$WEB/index.html"
