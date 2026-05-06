#!/usr/bin/env bash
set -e

if [ ! -d ".venv" ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
fi

source .venv/bin/activate

echo "Installing dependencies..."
pip install -q -r requirements.txt

echo
echo "Open http://127.0.0.1:8000 in your browser."
echo "Press Ctrl+C to stop."
echo
uvicorn app.main:app --host 127.0.0.1 --port 8000
