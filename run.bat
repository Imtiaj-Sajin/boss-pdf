@echo off
setlocal

if not exist .venv (
  echo Creating virtual environment...
  python -m venv .venv
)

call .venv\Scripts\activate.bat

echo Installing dependencies...
pip install -q -r requirements.txt

echo.
echo Open http://127.0.0.1:8000 in your browser.
echo Press Ctrl+C to stop.
echo.
uvicorn app.main:app --host 127.0.0.1 --port 8000
