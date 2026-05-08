@echo off
setlocal EnableDelayedExpansion

REM ---- Find a real Python 3.10+ ----
set "PYCMD="

REM Prefer the 'py' launcher (it finds the real install, not the WindowsApps stub).
where py >nul 2>&1
if not errorlevel 1 (
  py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
  if not errorlevel 1 set "PYCMD=py -3"
)

REM Fall back to plain `python` if py isn't there.
if "!PYCMD!"=="" (
  where python >nul 2>&1
  if not errorlevel 1 (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
    if not errorlevel 1 set "PYCMD=python"
  )
)

if "!PYCMD!"=="" (
  echo.
  echo Could not find Python 3.10 or newer.
  echo Install it from https://python.org and tick "Add Python to PATH",
  echo then re-run this script.
  echo.
  pause
  exit /b 1
)

echo Using: !PYCMD!
!PYCMD! --version

REM ---- Warn (don't block) on Python 3.14+, where some ML wheels lag behind ----
!PYCMD! -c "import sys; sys.exit(0 if sys.version_info < (3,14) else 1)" >nul 2>&1
if errorlevel 1 (
  echo.
  echo NOTE: Python 3.14+ may not have wheels for every optional package yet.
  echo       The app will still run with whatever installs successfully.
  echo       For best OCR support consider Python 3.11-3.13.
  echo.
)

REM ---- Create venv if missing ----
if not exist .venv (
  echo Creating virtual environment...
  !PYCMD! -m venv .venv
  if errorlevel 1 (
    echo Failed to create virtual environment. Aborting.
    pause
    exit /b 1
  )
)

call .venv\Scripts\activate.bat

REM ---- Core deps: must install ----
echo Installing core dependencies...
pip install -q -r requirements.txt
if errorlevel 1 (
  echo.
  echo Core dependency install failed. The app cannot start.
  pause
  exit /b 1
)

REM ---- Optional extras: install each line independently, ignore failures ----
echo Installing optional extras (some may skip - that's OK)...
for /F "usebackq tokens=* eol=#" %%P in ("requirements-extras.txt") do (
  set "PKG=%%P"
  if not "!PKG!"=="" (
    pip install -q "%%P" >nul 2>&1
    if not errorlevel 1 (
      echo   [ ok ] %%P
    ) else (
      echo   [skip] %%P
    )
  )
)

echo.
echo Open http://127.0.0.1:8000 in your browser.
echo Press Ctrl+C to stop.
echo.
uvicorn app.main:app --host 127.0.0.1 --port 8000
