@echo off
setlocal
cd /d "%~dp0"

if not exist node_modules (
  echo Installing dependencies...
  call npm.cmd install
)

start "" "http://127.0.0.1:3002"
call npm.cmd run dev -- --host 127.0.0.1 --port 3002
