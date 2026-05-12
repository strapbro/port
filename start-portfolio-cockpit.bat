@echo off
setlocal EnableExtensions
cd /d "%~dp0"

where node.exe >nul 2>nul
if errorlevel 1 goto :run_setup

where npm.cmd >nul 2>nul
if errorlevel 1 goto :run_setup

node.exe -e "const v=process.versions.node.split('.').map(Number); process.exit(((v[0]===20&&v[1]>=19)||v[0]>=22)?0:1)"
if errorlevel 1 goto :run_setup

if not exist node_modules goto :run_setup
goto :start_app

:run_setup
call "%~dp0setup-portfolio-cockpit.bat" --no-start
if errorlevel 1 exit /b %errorlevel%

:start_app

start "" "http://127.0.0.1:3002"
call npm.cmd run dev -- --host 127.0.0.1 --port 3002
if errorlevel 1 (
  echo.
  echo The dev server did not start. Try running setup-portfolio-cockpit.bat again.
  pause
  exit /b 1
)
