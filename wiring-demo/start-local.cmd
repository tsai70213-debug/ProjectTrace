@echo off
setlocal
set "WIRING_NODE=node.exe"
where node.exe >nul 2>nul
if errorlevel 1 (
  if not exist "%ProgramFiles%\nodejs\node.exe" goto node_missing
  set "WIRING_NODE=%ProgramFiles%\nodejs\node.exe"
)

"%WIRING_NODE%" "%~dp0server.cjs" --start
if errorlevel 1 goto failed
start "" "http://localhost:8764/"
if errorlevel 1 goto browser_failed
exit /b 0

:node_missing
echo Node.js was not found.
goto failed

:browser_failed
echo Open http://localhost:8764/ in your browser.
:failed
pause
exit /b 1
