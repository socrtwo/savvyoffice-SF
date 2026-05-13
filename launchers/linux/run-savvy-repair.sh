#!/usr/bin/env bash
# Savvy Repair for Microsoft Office — Linux launcher
set -e
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
WEB="$DIR/web"
PORT="${PORT:-8765}"
URL="http://localhost:$PORT/"

open_browser() {
  if command -v xdg-open >/dev/null 2>&1; then xdg-open "$URL" >/dev/null 2>&1 || true
  elif command -v sensible-browser >/dev/null 2>&1; then sensible-browser "$URL" >/dev/null 2>&1 || true
  elif command -v firefox >/dev/null 2>&1; then firefox "$URL" >/dev/null 2>&1 || true
  elif command -v google-chrome >/dev/null 2>&1; then google-chrome "$URL" >/dev/null 2>&1 || true
  elif command -v chromium >/dev/null 2>&1; then chromium "$URL" >/dev/null 2>&1 || true
  fi
}

if command -v python3 >/dev/null 2>&1; then
  (sleep 1 && open_browser) &
  cd "$WEB"
  exec python3 -m http.server "$PORT"
elif command -v python >/dev/null 2>&1; then
  (sleep 1 && open_browser) &
  cd "$WEB"
  exec python -m SimpleHTTPServer "$PORT"
fi

echo "Python not found. Open this file in your browser:"
echo "  $WEB/index.html"
exit 1
