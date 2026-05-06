# boss-pdf

PDF → Excel converter. Drop in a PDF, get back a clean `.xlsx` with one sheet per page.
Runs locally. Files never leave your machine.

## What it does

For each page it tries multiple table-extraction strategies and picks whichever produced the best result:

1. **pdfplumber** — best for native PDFs with text layers
2. **Camelot lattice** — for tables with visible borders
3. **Camelot stream** — for tables without borders
4. **Tesseract OCR** — fallback for scanned PDFs

The output spreadsheet has:
- One sheet per PDF page (multiple tables on a page are stacked)
- Numeric, currency and percentage cells coerced to real numbers
- Header rows styled and frozen
- Auto-fit column widths

## Setup

### 1. Python

Python 3.10+ recommended.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

### 2. System tools (for the best results)

These are optional but recommended — without them you lose Camelot and OCR fallback.

**Windows**
- Ghostscript: https://ghostscript.com/releases/gsdnld.html (needed by Camelot)
- Poppler: https://github.com/oschwartz10612/poppler-windows/releases (needed by `pdf2image` for OCR)
  - Add the `bin` folder to your `PATH`
- Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki
  - Add the install folder to your `PATH`

**macOS**
```bash
brew install ghostscript poppler tesseract
```

**Linux (Debian/Ubuntu)**
```bash
sudo apt install ghostscript poppler-utils tesseract-ocr
```

If any of these are missing, the corresponding extractor is skipped — pdfplumber will still work for native PDFs.

## Run

**Windows:**
```bash
run.bat
```

**macOS/Linux:**
```bash
./run.sh
```

Or directly:
```bash
uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Then open http://127.0.0.1:8000 in your browser.

## API

Single endpoint:

```
POST /api/convert
Content-Type: multipart/form-data
file: <your.pdf>
```

Returns the `.xlsx` as an attachment, or a JSON `{detail: "..."}` error.

## Limits

- 50 MB per upload
- Conversion quality depends on the PDF. Native (text-layer) PDFs with structured tables are near-perfect. Heavily scanned, multi-column, or visually-complex pages are best-effort.

## Project layout

```
app/
  main.py        # FastAPI app + /api/convert endpoint
  converter.py   # extraction pipeline (pdfplumber / camelot / OCR) + xlsx rendering
web/
  index.html     # drag-drop UI
  style.css
  app.js
requirements.txt
run.bat / run.sh
```
