@echo off
setlocal
rem Savvy Repair for Microsoft Office — Windows launcher
rem Starts a local static server out of .\web and opens the default browser.
set DIR=%~dp0
set WEB=%DIR%web

where python >nul 2>nul
if %errorlevel%==0 (
  start "" http://localhost:8765/
  pushd "%WEB%"
  python -m http.server 8765
  popd
  goto :eof
)

where py >nul 2>nul
if %errorlevel%==0 (
  start "" http://localhost:8765/
  pushd "%WEB%"
  py -3 -m http.server 8765
  popd
  goto :eof
)

echo Python was not found on PATH.
echo You can still use the app: just open "%WEB%\index.html" in your browser.
pause
