#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
echo "Starting Sales Statistics Analyzer..."
echo "Open http://localhost:8787 in your browser."
python3 -m http.server 8787

