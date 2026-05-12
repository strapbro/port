@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "APP_NAME=AI Buildout Portfolio Recomp Cockpit"
set "NODE_DOWNLOAD_URL=https://nodejs.org/en/download"
set "START_AFTER=prompt"

if /I "%~1"=="--no-start" set "START_AFTER=no"
if /I "%~1"=="/no-start" set "START_AFTER=no"
if /I "%~1"=="--start" set "START_AFTER=yes"
if /I "%~1"=="/start" set "START_AFTER=yes"

echo.
echo === %APP_NAME% first-time setup ===
echo.

where node.exe >nul 2>nul
if errorlevel 1 goto :missing_node

where npm.cmd >nul 2>nul
if errorlevel 1 goto :missing_npm

for /f "delims=" %%v in ('node.exe -p "process.versions.node"') do set "NODE_VERSION=%%v"
node.exe -e "const v=process.versions.node.split('.').map(Number); const ok=(v[0]===20&&v[1]>=19)||v[0]>=22; if(!ok){console.error('Node '+process.versions.node+' is too old. This app requires Node 20.19+ or 22.12+.'); process.exit(1)}"
if errorlevel 1 goto :old_node

for /f "delims=" %%v in ('npm.cmd --version') do set "NPM_VERSION=%%v"

echo Found Node %NODE_VERSION%
echo Found npm %NPM_VERSION%
echo.

if exist node_modules goto :install_existing
if exist package-lock.json goto :install_clean
goto :install_no_lock

:install_existing
echo Existing node_modules folder found.
echo Refreshing dependencies without deleting files that may be in use...
call npm.cmd install
set "INSTALL_EXIT=%ERRORLEVEL%"
goto :install_done

:install_clean
echo Installing exact dependencies from package-lock.json...
call npm.cmd ci
set "INSTALL_EXIT=%ERRORLEVEL%"
goto :install_done

:install_no_lock
echo Installing dependencies...
call npm.cmd install
set "INSTALL_EXIT=%ERRORLEVEL%"
goto :install_done

:install_done
if not "%INSTALL_EXIT%"=="0" goto :npm_failed

echo.
echo Verifying the app builds...
call npm.cmd run build
if errorlevel 1 goto :build_failed

echo.
echo Setup complete.
echo.

if /I "%START_AFTER%"=="yes" goto :start_app
if /I "%START_AFTER%"=="no" goto :done_no_pause

choice /C YN /N /M "Start the app now? [Y/N] "
if errorlevel 2 goto :done
if errorlevel 1 goto :start_app

:start_app
echo.
echo Starting dev server at http://127.0.0.1:3002
start "" "http://127.0.0.1:3002"
call npm.cmd run dev -- --host 127.0.0.1 --port 3002
if errorlevel 1 goto :start_failed
goto :eof

:done
echo.
echo Run start-portfolio-cockpit.bat whenever you want to open it.
pause
goto :eof

:done_no_pause
echo Run start-portfolio-cockpit.bat whenever you want to open it.
exit /b 0

:missing_node
echo ERROR: Node.js is not installed or is not available on PATH.
echo Install the current LTS version from:
echo %NODE_DOWNLOAD_URL%
echo.
echo After installing Node.js, close this window and run this setup file again.
pause
exit /b 1

:missing_npm
echo ERROR: npm.cmd was not found.
echo npm normally installs with Node.js. Reinstall Node.js from:
echo %NODE_DOWNLOAD_URL%
echo.
pause
exit /b 1

:old_node
echo.
echo ERROR: Node %NODE_VERSION% is too old for this app.
echo Install Node.js 20.19 or newer, or Node.js 22.12 or newer:
echo %NODE_DOWNLOAD_URL%
echo.
pause
exit /b 1

:npm_failed
echo.
echo ERROR: Dependency install failed.
echo Check the npm error above, then run this setup file again.
pause
exit /b 1

:build_failed
echo.
echo ERROR: The app installed, but the build failed.
echo Check the build error above before starting the app.
pause
exit /b 1

:start_failed
echo.
echo ERROR: The dev server did not start.
pause
exit /b 1
