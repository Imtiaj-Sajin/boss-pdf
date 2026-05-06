"""PDF -> Excel conversion engine.

Strategy: try several extraction methods per page, score each, pick the best.
Falls back to OCR for scanned PDFs.
"""
from __future__ import annotations

import io
import logging
import os
import tempfile
from dataclasses import dataclass
from typing import Iterable

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Optional deps — imported lazily so the app still works if a method is unavailable.
try:
    import camelot  # type: ignore
    _HAS_CAMELOT = True
except Exception as e:
    logger.warning("camelot unavailable: %s", e)
    _HAS_CAMELOT = False

try:
    import pytesseract  # type: ignore
    from pdf2image import convert_from_path  # type: ignore
    from PIL import Image  # noqa: F401
    _HAS_OCR = True
except Exception as e:
    logger.warning("OCR stack unavailable: %s", e)
    _HAS_OCR = False


Table = list[list[str]]  # 2D grid of strings


@dataclass
class PageResult:
    page_num: int  # 1-indexed
    tables: list[Table]
    method: str  # which extractor produced the result
    score: float


def _clean_cell(v) -> str:
    if v is None:
        return ""
    s = str(v).replace(" ", " ").strip()
    # collapse runs of internal whitespace but preserve newlines as-is for multi-line cells
    return "\n".join(" ".join(line.split()) for line in s.splitlines())


def _normalize_table(rows: Iterable[Iterable]) -> Table:
    out: Table = []
    for r in rows:
        out.append([_clean_cell(c) for c in r])
    # pad rows to equal width
    if not out:
        return out
    width = max(len(r) for r in out)
    for r in out:
        if len(r) < width:
            r.extend([""] * (width - len(r)))
    # drop fully-empty trailing rows
    while out and all(c == "" for c in out[-1]):
        out.pop()
    return out


def _score_table(t: Table) -> float:
    """Higher = better. Rewards cells with content, penalizes empty grids."""
    if not t or not t[0]:
        return 0.0
    rows = len(t)
    cols = len(t[0])
    cells = rows * cols
    if cells == 0:
        return 0.0
    filled = sum(1 for row in t for c in row if c)
    fill_ratio = filled / cells
    # reward bigger tables that are also dense
    return filled * (0.5 + 0.5 * fill_ratio)


def _extract_pdfplumber(pdf_path: str) -> list[PageResult]:
    results: list[PageResult] = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            tables: list[Table] = []
            try:
                # try strict first (visible lines), then loose (no lines)
                for settings in (
                    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                    {"vertical_strategy": "text", "horizontal_strategy": "text",
                     "intersection_tolerance": 5},
                ):
                    raw = page.extract_tables(table_settings=settings) or []
                    for r in raw:
                        norm = _normalize_table(r)
                        if norm:
                            tables.append(norm)
                    if tables:
                        break
            except Exception as e:
                logger.warning("pdfplumber page %d failed: %s", i, e)
            score = sum(_score_table(t) for t in tables)
            results.append(PageResult(page_num=i, tables=tables, method="pdfplumber", score=score))
    return results


def _camelot_for_flavor(pdf_path: str, flavor: str, page_count: int) -> dict[int, list[Table]]:
    """Return {page_num: [tables...]} for the given camelot flavor."""
    out: dict[int, list[Table]] = {}
    if not _HAS_CAMELOT:
        return out
    try:
        tlist = camelot.read_pdf(pdf_path, pages=f"1-{page_count}", flavor=flavor,
                                 suppress_stdout=True)
    except Exception as e:
        logger.warning("camelot %s failed: %s", flavor, e)
        return out
    for t in tlist:
        try:
            page = int(t.page)
            df = t.df
            rows = df.values.tolist()
            norm = _normalize_table(rows)
            if norm:
                out.setdefault(page, []).append(norm)
        except Exception as e:
            logger.warning("camelot %s table parse failed: %s", flavor, e)
    return out


def _extract_ocr_page(pdf_path: str, page_num: int) -> Table | None:
    if not _HAS_OCR:
        return None
    try:
        images = convert_from_path(pdf_path, dpi=300, first_page=page_num, last_page=page_num)
        if not images:
            return None
        img = images[0]
        # Use TSV output to recover column structure from word boxes
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        # Group words by line, sort by x-position
        lines: dict[tuple, list[tuple]] = {}
        for i, txt in enumerate(data["text"]):
            if not txt or not txt.strip():
                continue
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            lines.setdefault(key, []).append((data["left"][i], txt.strip()))
        # Build rows
        rows: list[list[str]] = []
        for key in sorted(lines.keys()):
            words = sorted(lines[key], key=lambda x: x[0])
            rows.append([w for _, w in words])
        return _normalize_table(rows) if rows else None
    except Exception as e:
        logger.warning("OCR page %d failed: %s", page_num, e)
        return None


def convert_pdf_to_workbook(pdf_path: str) -> Workbook:
    """Convert a PDF to an openpyxl Workbook. Returns the workbook (not saved).

    For each page, picks the best of: pdfplumber, camelot-lattice, camelot-stream.
    Falls back to OCR if no text-based extractor returns anything.
    """
    # Page 1: pdfplumber pass (also gives us the page count & whether text exists)
    pp_results = _extract_pdfplumber(pdf_path)
    page_count = len(pp_results)

    # Camelot passes (whole-document)
    lattice = _camelot_for_flavor(pdf_path, "lattice", page_count)
    stream = _camelot_for_flavor(pdf_path, "stream", page_count)

    # Detect "looks like a scanned PDF" — no text on any page from pdfplumber
    pp_total_score = sum(r.score for r in pp_results)
    pp_total_tables = sum(len(r.tables) for r in pp_results)
    cm_total = sum(len(v) for v in lattice.values()) + sum(len(v) for v in stream.values())
    do_ocr = (pp_total_tables == 0 and cm_total == 0) and _HAS_OCR

    chosen_per_page: list[PageResult] = []
    for r in pp_results:
        candidates: list[tuple[float, str, list[Table]]] = []
        candidates.append((r.score, "pdfplumber", r.tables))
        if r.page_num in lattice:
            ts = lattice[r.page_num]
            candidates.append((sum(_score_table(t) for t in ts), "camelot-lattice", ts))
        if r.page_num in stream:
            ts = stream[r.page_num]
            candidates.append((sum(_score_table(t) for t in ts), "camelot-stream", ts))

        if do_ocr:
            ocr_table = _extract_ocr_page(pdf_path, r.page_num)
            if ocr_table:
                candidates.append((_score_table(ocr_table), "ocr", [ocr_table]))

        score, method, tables = max(candidates, key=lambda c: c[0]) if candidates else (0.0, "none", [])
        chosen_per_page.append(
            PageResult(page_num=r.page_num, tables=tables, method=method, score=score)
        )

    return _build_workbook(chosen_per_page)


# ---------- Excel rendering ----------

_HEADER_FILL = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFFFF")
_THIN = Side(border_style="thin", color="FFB7B7B7")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_WRAP = Alignment(vertical="center", wrap_text=True)


def _looks_like_header(row: list[str]) -> bool:
    if not row:
        return False
    non_empty = [c for c in row if c]
    if len(non_empty) < max(2, len(row) // 2):
        return False
    # heuristic: header rows tend to be short text, not numeric
    numeric = sum(1 for c in non_empty if _is_numeric(c))
    return numeric <= len(non_empty) // 3


def _is_numeric(s: str) -> bool:
    s = s.replace(",", "").replace("$", "").replace("%", "").strip()
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def _autosize(ws) -> None:
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            v = cell.value
            if v is None:
                continue
            for line in str(v).splitlines():
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 60)


def _write_table(ws, table: Table, start_row: int = 1) -> int:
    """Write a table starting at start_row. Returns the next free row."""
    if not table:
        return start_row
    has_header = _looks_like_header(table[0])
    for r_off, row in enumerate(table):
        excel_row = start_row + r_off
        for c_off, val in enumerate(row, start=1):
            cell = ws.cell(row=excel_row, column=c_off)
            if _is_numeric(val):
                try:
                    cell.value = float(val.replace(",", "").replace("$", "").replace("%", ""))
                    if "%" in val:
                        cell.number_format = "0.00%"
                        cell.value = cell.value / 100
                    elif "$" in val:
                        cell.number_format = "$#,##0.00"
                except ValueError:
                    cell.value = val
            else:
                cell.value = val
            cell.border = _BORDER
            cell.alignment = _WRAP
            if has_header and r_off == 0:
                cell.fill = _HEADER_FILL
                cell.font = _HEADER_FONT
    return start_row + len(table) + 1  # blank spacer row


def _build_workbook(pages: list[PageResult]) -> Workbook:
    wb = Workbook()
    # remove default sheet — we'll add our own
    default = wb.active
    wb.remove(default)

    any_content = False
    for page in pages:
        if not page.tables:
            continue
        any_content = True
        # one sheet per page; if multiple tables on a page, stack them with spacer rows
        sheet_name = f"Page {page.page_num}"
        ws = wb.create_sheet(title=sheet_name[:31])
        next_row = 1
        for t in page.tables:
            next_row = _write_table(ws, t, start_row=next_row)
        _autosize(ws)
        ws.freeze_panes = "A2"

    if not any_content:
        # produce an empty sheet with a friendly note rather than a broken file
        ws = wb.create_sheet(title="No tables found")
        ws["A1"] = "No tables were detected in this PDF."
        ws["A2"] = "If the PDF is scanned, ensure Tesseract & Poppler are installed."

    return wb


def convert_pdf_bytes_to_xlsx_bytes(pdf_bytes: bytes) -> bytes:
    """Convenience: bytes-in, bytes-out. Writes the PDF to a tempfile (camelot needs a path)."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp_path = f.name
    try:
        wb = convert_pdf_to_workbook(tmp_path)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
