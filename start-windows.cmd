@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>&1
if errorlevel 1 (
  echo Please install Node.js 22 or later: https://nodejs.org/
  pause
  exit /b 1
)
if not exist node_modules\yauzl\package.json (
  call npm ci --omit=dev
  if errorlevel 1 (
    pause
    exit /b 1
  )
)
node product-demo\server.mjs --open-browser
if errorlevel 1 pause
