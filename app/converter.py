"""PDF -> Excel conversion engine.

Strategy: try several extraction methods per page, score each, pick the best.
The pdfplumber path also reads char-level styling (bold, color) from the PDF
so the output xlsx visually resembles the source. Falls back to OCR for
scanned PDFs.
"""
from __future__ import annotations

import io
import logging
import os
import re
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from typing import Iterable, Optional

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


# ---------- Rich data model ----------

@dataclass
class CellStyle:
    bold: bool = False
    italic: bool = False
    color_hex: Optional[str] = None  # 6-digit hex, no alpha, e.g. "FF0000"


@dataclass
class RichCell:
    text: str = ""
    style: CellStyle = field(default_factory=CellStyle)


RichTable = list[list[RichCell]]


@dataclass
class PageResult:
    page_num: int  # 1-indexed
    tables: list[RichTable]
    title_lines: list[str]
    method: str
    score: float


# ---------- Char-level helpers (pdfplumber) ----------

def _chars_in_bbox(chars: list[dict], bbox: tuple[float, float, float, float]) -> list[dict]:
    x0, top, x1, bottom = bbox
    out = []
    for c in chars:
        cx = (c["x0"] + c["x1"]) / 2
        cy = (c["top"] + c["bottom"]) / 2
        if x0 <= cx <= x1 and top <= cy <= bottom:
            out.append(c)
    return out


def _is_bold(chars: list[dict]) -> bool:
    if not chars:
        return False
    bold_kw = ("bold", "black", "heavy", "semibold", "demibold")
    n_bold = sum(1 for c in chars
                 if any(k in (c.get("fontname") or "").lower() for k in bold_kw))
    return n_bold > len(chars) / 2


def _color_to_hex(color) -> Optional[str]:
    """Convert a pdfplumber non_stroking_color value to '#RRGGBB' (no '#')."""
    if color is None:
        return None
    if isinstance(color, (int, float)):
        v = max(0, min(255, int(round(float(color) * 255))))
        return f"{v:02X}{v:02X}{v:02X}"
    if isinstance(color, (list, tuple)):
        nums = [float(x) for x in color if isinstance(x, (int, float))]
        if len(nums) == 1:
            v = max(0, min(255, int(round(nums[0] * 255))))
            return f"{v:02X}{v:02X}{v:02X}"
        if len(nums) == 3:
            r, g, b = (max(0, min(255, int(round(x * 255)))) for x in nums)
            return f"{r:02X}{g:02X}{b:02X}"
        if len(nums) == 4:
            c, m, y, k = nums
            r = (1 - c) * (1 - k)
            g = (1 - m) * (1 - k)
            b = (1 - y) * (1 - k)
            r, g, b = (max(0, min(255, int(round(v * 255)))) for v in (r, g, b))
            return f"{r:02X}{g:02X}{b:02X}"
    return None


def _dominant_color(chars: list[dict]) -> Optional[str]:
    if not chars:
        return None
    counts: Counter = Counter()
    for c in chars:
        col = c.get("non_stroking_color")
        hexv = _color_to_hex(col)
        if hexv and hexv != "000000":
            counts[hexv] += 1
    return counts.most_common(1)[0][0] if counts else None


# ---------- Text/cell normalization ----------

def _clean_cell_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = str(s).replace(" ", " ").strip()
    return "\n".join(" ".join(line.split()) for line in s.splitlines())


def _normalize_rich_table(rows: list[list[RichCell]]) -> RichTable:
    if not rows:
        return rows
    width = max(len(r) for r in rows)
    for r in rows:
        if len(r) < width:
            r.extend([RichCell() for _ in range(width - len(r))])
    while rows and all(c.text == "" for c in rows[-1]):
        rows.pop()
    while rows and all(c.text == "" for c in rows[0]):
        rows.pop(0)
    return rows


def _normalize_plain_table(rows: Iterable[Iterable]) -> RichTable:
    return _normalize_rich_table(
        [[RichCell(text=_clean_cell_text(c)) for c in r] for r in rows]
    )


def _score_table(t: RichTable) -> float:
    if not t or not t[0]:
        return 0.0
    cells = sum(len(r) for r in t)
    if cells == 0:
        return 0.0
    filled = sum(1 for row in t for c in row if c.text)
    return filled * (0.5 + 0.5 * (filled / cells))


# ---------- Extraction: pdfplumber with rich style ----------

_TABLE_SETTINGS_OPTIONS = (
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    {"vertical_strategy": "lines", "horizontal_strategy": "text",
     "intersection_tolerance": 5},
    {"vertical_strategy": "text", "horizontal_strategy": "text",
     "intersection_tolerance": 5},
)


def _extract_rich_table_from_pdfplumber(page, table) -> RichTable:
    """Pull a pdfplumber Table into a RichTable with style info from chars."""
    chars = page.chars
    rows: list[list[RichCell]] = []
    for row in table.rows:
        rich_row: list[RichCell] = []
        for cell_bbox in row.cells:
            if cell_bbox is None:
                rich_row.append(RichCell())
                continue
            try:
                sub = page.within_bbox(cell_bbox)
                text = _clean_cell_text(sub.extract_text(x_tolerance=2, y_tolerance=2))
            except Exception:
                text = ""
            cell_chars = _chars_in_bbox(chars, cell_bbox)
            style = CellStyle(
                bold=_is_bold(cell_chars),
                color_hex=_dominant_color(cell_chars),
            )
            rich_row.append(RichCell(text=text, style=style))
        rows.append(rich_row)
    return _normalize_rich_table(rows)


def _extract_page_title_lines(page, table_bboxes: list[tuple]) -> list[str]:
    """Lines of text on the page above the first table, keep their original order."""
    if not table_bboxes:
        return []
    top_of_tables = min(b[1] for b in table_bboxes)
    if top_of_tables <= 5:
        return []
    try:
        crop = page.within_bbox((0, 0, float(page.width), float(top_of_tables)))
        text = crop.extract_text(x_tolerance=2, y_tolerance=2) or ""
    except Exception:
        return []
    lines = [ln.strip() for ln in text.splitlines()]
    return [ln for ln in lines if ln]


def _extract_pdfplumber(pdf_path: str) -> list[PageResult]:
    results: list[PageResult] = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            tables: list[RichTable] = []
            table_bboxes: list[tuple] = []
            for settings in _TABLE_SETTINGS_OPTIONS:
                try:
                    found = page.find_tables(table_settings=settings) or []
                except Exception as e:
                    logger.warning("find_tables p%d %s: %s", i, settings, e)
                    found = []
                for t in found:
                    try:
                        rt = _extract_rich_table_from_pdfplumber(page, t)
                        if rt:
                            tables.append(rt)
                            table_bboxes.append(t.bbox)
                    except Exception as e:
                        logger.warning("rich extract p%d failed: %s", i, e)
                if tables:
                    break

            title_lines = _extract_page_title_lines(page, table_bboxes)
            score = sum(_score_table(t) for t in tables)
            results.append(PageResult(
                page_num=i, tables=tables, title_lines=title_lines,
                method="pdfplumber", score=score,
            ))
    return results


# ---------- Extraction: camelot (fallback, plain text) ----------

def _camelot_pages(pdf_path: str, flavor: str, page_count: int) -> dict[int, list[RichTable]]:
    out: dict[int, list[RichTable]] = {}
    if not _HAS_CAMELOT:
        return out
    try:
        tlist = camelot.read_pdf(pdf_path, pages=f"1-{page_count}",
                                 flavor=flavor, suppress_stdout=True)
    except Exception as e:
        logger.warning("camelot %s failed: %s", flavor, e)
        return out
    for t in tlist:
        try:
            page = int(t.page)
            rows = t.df.values.tolist()
            norm = _normalize_plain_table(rows)
            if norm:
                out.setdefault(page, []).append(norm)
        except Exception as e:
            logger.warning("camelot %s parse failed: %s", flavor, e)
    return out


# ---------- Extraction: OCR (fallback for scanned PDFs) ----------

def _extract_ocr_page(pdf_path: str, page_num: int) -> Optional[RichTable]:
    if not _HAS_OCR:
        return None
    try:
        images = convert_from_path(pdf_path, dpi=300, first_page=page_num, last_page=page_num)
        if not images:
            return None
        data = pytesseract.image_to_data(images[0], output_type=pytesseract.Output.DICT)
        lines: dict[tuple, list[tuple]] = {}
        for i, txt in enumerate(data["text"]):
            if not txt or not txt.strip():
                continue
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            lines.setdefault(key, []).append((data["left"][i], txt.strip()))
        rows: list[list[str]] = []
        for key in sorted(lines.keys()):
            words = sorted(lines[key], key=lambda x: x[0])
            rows.append([w for _, w in words])
        return _normalize_plain_table(rows) if rows else None
    except Exception as e:
        logger.warning("OCR page %d failed: %s", page_num, e)
        return None


# ---------- Pipeline ----------

def convert_pdf_to_workbook(pdf_path: str) -> Workbook:
    pp_results = _extract_pdfplumber(pdf_path)
    page_count = len(pp_results)
    lattice = _camelot_pages(pdf_path, "lattice", page_count)
    stream = _camelot_pages(pdf_path, "stream", page_count)

    pp_total_tables = sum(len(r.tables) for r in pp_results)
    cm_total = sum(len(v) for v in lattice.values()) + sum(len(v) for v in stream.values())
    do_ocr = (pp_total_tables == 0 and cm_total == 0) and _HAS_OCR

    chosen: list[PageResult] = []
    for r in pp_results:
        candidates: list[tuple[float, str, list[RichTable], list[str]]] = []
        candidates.append((r.score, "pdfplumber", r.tables, r.title_lines))
        if r.page_num in lattice:
            ts = lattice[r.page_num]
            candidates.append((sum(_score_table(t) for t in ts), "camelot-lattice", ts, []))
        if r.page_num in stream:
            ts = stream[r.page_num]
            candidates.append((sum(_score_table(t) for t in ts), "camelot-stream", ts, []))
        if do_ocr:
            ocr_t = _extract_ocr_page(pdf_path, r.page_num)
            if ocr_t:
                candidates.append((_score_table(ocr_t), "ocr", [ocr_t], []))
        score, method, tables, titles = max(candidates, key=lambda c: c[0]) \
            if candidates else (0.0, "none", [], [])
        chosen.append(PageResult(
            page_num=r.page_num, tables=tables, title_lines=titles,
            method=method, score=score,
        ))

    return _build_workbook(chosen)


# ---------- Excel rendering ----------

_THIN = Side(border_style="thin", color="FFB7B7B7")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_FILL = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
_SECTION_FILL = PatternFill(start_color="FFE8EEF7", end_color="FFE8EEF7", fill_type="solid")
_TOTAL_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")

_NUM_RE = re.compile(r"^\(?-?\$?\s*-?[\d,]+(\.\d+)?\)?\s*%?$")


def _looks_like_header_row(row: list[RichCell]) -> bool:
    non_empty = [c for c in row if c.text]
    if len(non_empty) < max(2, len(row) // 2):
        return False
    numeric = sum(1 for c in non_empty if _is_numeric(c.text))
    return numeric <= len(non_empty) // 3


def _is_numeric(s: str) -> bool:
    if not s:
        return False
    s2 = s.strip().replace(",", "").replace("$", "").replace("%", "")
    s2 = s2.replace("(", "-").replace(")", "")
    if not s2 or s2 in ("-", ".", "-."):
        return False
    try:
        float(s2)
        return True
    except ValueError:
        return False


def _parse_number(s: str) -> float:
    s2 = s.strip().replace(",", "").replace("$", "").replace("%", "")
    s2 = s2.replace("(", "-").replace(")", "")
    return float(s2)


def _is_section_row(text_row: list[str]) -> bool:
    """Row where only the first column has text (section/header divider)."""
    if not text_row or not text_row[0].strip():
        return False
    return all(not t.strip() for t in text_row[1:])


def _is_total_row(text_row: list[str]) -> bool:
    if not text_row or not text_row[0]:
        return False
    head = text_row[0].strip().lower()
    return (
        head.startswith("total ")
        or head.startswith("sub total")
        or head.startswith("subtotal")
        or head.startswith("net ")
        or head == "total"
        or head.startswith("grand total")
    )


def _autosize(ws, n_cols: int) -> None:
    for col_idx in range(1, n_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[col_letter]:
            v = cell.value
            if v is None:
                continue
            for line in str(v).splitlines():
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 60)


def _write_title_rows(ws, title_lines: list[str], n_cols: int, start_row: int) -> int:
    if not title_lines:
        return start_row
    row = start_row
    for i, line in enumerate(title_lines):
        cell = ws.cell(row=row, column=1, value=line)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        size = 14 if i == 0 else (12 if i == 1 else 10)
        cell.font = Font(bold=(i <= 1), size=size, color="FF1F2937")
        if n_cols > 1:
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=n_cols)
        row += 1
    return row + 1  # spacer


def _apply_value_and_format(cell, raw: str) -> None:
    if _is_numeric(raw):
        try:
            num = _parse_number(raw)
        except ValueError:
            cell.value = raw
            return
        if "%" in raw:
            cell.value = num / 100
            cell.number_format = "0.00%;[Red]-0.00%"
        elif "$" in raw:
            cell.value = num
            cell.number_format = '_("$"* #,##0.00_);[Red]_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        else:
            cell.value = num
            cell.number_format = "#,##0.00;[Red]-#,##0.00"
    else:
        cell.value = raw


def _write_rich_table(ws, table: RichTable, start_row: int, n_cols: int) -> int:
    if not table:
        return start_row

    is_header_row = [False] * len(table)
    is_header_row[0] = _looks_like_header_row(table[0])

    for r_off, row in enumerate(table):
        excel_row = start_row + r_off
        text_row = [c.text for c in row]
        section = _is_section_row(text_row)
        total = _is_total_row(text_row)
        header = is_header_row[r_off]

        for c_off, rich in enumerate(row, start=1):
            cell = ws.cell(row=excel_row, column=c_off)
            _apply_value_and_format(cell, rich.text)
            cell.border = _BORDER
            cell.alignment = Alignment(
                vertical="center",
                horizontal=("right" if _is_numeric(rich.text) else "left"),
                wrap_text=True,
            )

            font_kwargs = {}
            if rich.style.bold:
                font_kwargs["bold"] = True
            if rich.style.color_hex and rich.style.color_hex != "000000":
                font_kwargs["color"] = "FF" + rich.style.color_hex

            if header:
                font_kwargs["bold"] = True
                font_kwargs["color"] = "FFFFFFFF"
                cell.fill = _HEADER_FILL
                cell.alignment = Alignment(
                    vertical="center", horizontal="center", wrap_text=True
                )
            elif section:
                font_kwargs["bold"] = True
                font_kwargs["size"] = 11
                cell.fill = _SECTION_FILL
            elif total:
                font_kwargs["bold"] = True
                cell.fill = _TOTAL_FILL

            if font_kwargs:
                cell.font = Font(**font_kwargs)

        if section and n_cols > 1:
            try:
                ws.merge_cells(start_row=excel_row, start_column=1,
                               end_row=excel_row, end_column=n_cols)
            except Exception:
                pass  # already merged in some pathological case

    return start_row + len(table) + 1  # spacer row


def _build_workbook(pages: list[PageResult]) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    any_content = False
    for page in pages:
        if not page.tables:
            continue
        any_content = True
        ws = wb.create_sheet(title=f"Page {page.page_num}"[:31])

        n_cols = max((len(t[0]) for t in page.tables if t), default=1)

        next_row = 1
        if page.title_lines:
            next_row = _write_title_rows(ws, page.title_lines, n_cols, next_row)

        for t in page.tables:
            next_row = _write_rich_table(ws, t, start_row=next_row, n_cols=n_cols)

        # Freeze panes below the title block (or at row 2 if no title)
        freeze_row = (1 + len(page.title_lines) + 1 + 1) if page.title_lines else 2
        ws.freeze_panes = f"A{freeze_row}"

        _autosize(ws, n_cols)

    if not any_content:
        ws = wb.create_sheet(title="No tables found")
        ws["A1"] = "No tables were detected in this PDF."
        ws["A2"] = "If the PDF is scanned, ensure Tesseract & Poppler are installed."

    return wb


def convert_pdf_bytes_to_xlsx_bytes(pdf_bytes: bytes) -> bytes:
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
