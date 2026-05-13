@echo off
setlocal EnableDelayedExpansion

REM ============================================================
REM  boss-pdf launcher
REM  Prefers a paddle-compatible Python (3.13 / 3.12 / 3.11 / 3.10)
REM  so PaddleOCR — the primary OCR engine — installs cleanly.
REM ============================================================

set "PYCMD="

REM 1) Look for paddle-compatible Pythons via the py launcher, in order
for %%V in (3.13 3.12 3.11 3.10) do (
  if "!PYCMD!"=="" (
    py -%%V --version >nul 2>&1
    if not errorlevel 1 set "PYCMD=py -%%V"
  )
)

REM 2) Fall back to any py -3 (might be 3.14 — we'll warn loudly below)
if "!PYCMD!"=="" (
  where py >nul 2>&1
  if not errorlevel 1 (
    py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
    if not errorlevel 1 set "PYCMD=py -3"
  )
)

REM 3) Last resort: plain `python` on PATH (NOT the WindowsApps stub)
if "!PYCMD!"=="" (
  where python >nul 2>&1
  if not errorlevel 1 (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
    if not errorlevel 1 set "PYCMD=python"
  )
)

if "!PYCMD!"=="" (
  echo.
  echo ============================================================
  echo   Python 3.10-3.13 not found.
  echo   Install Python 3.13 from https://python.org for full
  echo   PaddleOCR support, then re-run this script.
  echo ============================================================
  echo.
  pause
  exit /b 1
)

echo Using: !PYCMD!
!PYCMD! --version

REM Loudly warn if Python is too new for paddlepaddle
!PYCMD! -c "import sys; sys.exit(0 if sys.version_info >= (3,14) else 1)" >nul 2>&1
if not errorlevel 1 (
  echo.
  echo ============================================================
  echo  WARNING: Python 3.14+ has NO paddlepaddle wheels yet.
  echo  PaddleOCR (the primary OCR engine) WILL NOT INSTALL.
  echo  Scanned PDFs will fall back to RapidOCR / Tesseract or fail.
  echo  Install Python 3.13 from https://python.org for full OCR.
  echo ============================================================
  echo.
)

REM ---- venv ----
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

REM ---- pip itself ----
python -m pip install -q --upgrade pip >nul 2>&1

REM ---- core deps (must succeed) ----
echo Installing core dependencies...
pip install -q -r requirements.txt
if errorlevel 1 (
  echo Core install failed. The app cannot start.
  pause
  exit /b 1
)

REM ---- extras (each line independent; skip on failure) ----
echo.
echo Installing OCR / table extras (each independent; some may skip)...
for /F "usebackq tokens=* eol=#" %%P in ("requirements-extras.txt") do (
  set "PKG=%%P"
  if not "!PKG!"=="" (
    pip install -q "%%P" >nul 2>&1
    if not errorlevel 1 (
      echo   [ ok ] %%P
    ) else (
      echo   [skip] %%P -- install failed but not fatal
    )
  )
)

REM ---- show what OCR engines are usable ----
echo.
echo Active OCR engines:
python -c "from app import converter as c; print('  paddleocr:', 'YES (v' + str(c._PADDLE_API_VERSION) + ')' if c._HAS_PADDLE else 'no'); print('  rapidocr: ', 'YES' if c._HAS_RAPIDOCR else 'no'); print('  tesseract:', 'YES' if c._HAS_TESSERACT else 'no'); print('  PyMuPDF:  ', 'YES' if c._HAS_FITZ else 'no')" 2>nul

echo.
echo Open http://127.0.0.1:8000 in your browser.
echo Press Ctrl+C to stop.
echo.
uvicorn app.main:app --host 127.0.0.1 --port 8000

REM Keep window open if uvicorn exits for any reason (so you can read the error).
echo.
echo Server stopped. Press any key to close this window.
pause >nul
