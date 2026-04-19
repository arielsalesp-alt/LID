@echo off
setlocal
cd /d "%~dp0"
echo Starting Sales Statistics Analyzer...
echo Open http://localhost:8787 in your browser if it does not open automatically.
start "" "http://localhost:8787"
python -m http.server 8787
