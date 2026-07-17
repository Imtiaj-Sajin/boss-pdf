"""Table Scrape — see boss-pdf's own grid, fix it, re-extract.

This does NOT invent its own table detector. It surfaces the exact grid that
boss-pdf's core converter already derives (pdfplumber `find_tables` with the
same settings ladder), lets the user correct it — typically to split a column
boss-pdf accidentally merged — and then feeds those lines straight back into
pdfplumber via the `explicit` strategies. Extraction and workbook building then
run through the core converter, so the output is exactly "Make it Excel" plus
your correction.

Lines are normalized fractions of the page ([0,1], top-left origin) so the
browser can render at any scale and still line up, and so one set of columns
travels across pages of differing size for free.

Columns are structural: they sit at the same x on every page, so the user's
calibration travels. Rows are just text lines and drift page to page, so by
default each page keeps boss-pdf's own row detection.
"""
from __future__ import annotations

import io
import logging
import os
import tempfile

import pdfplumber

from .converter import (
    _CAMELOT_SCORE_GATE,
    _HAS_CAMELOT,
    _TABLE_SETTINGS_OPTIONS,
    _build_workbook,
    _extract_rich_table_from_pdfplumber,
    _extract_page_title_lines,
    _normalize_plain_table,
    _score_table,
    PageResult,
)

logger = logging.getLogger(__name__)


# ---------- camelot: the other engine that may win a page ----------

def _camelot_tables(pdf_path: str, page_num: int, flavor: str):
    """Raw camelot Table objects for one page (we need their geometry, not just
    the dataframe, so we can draw the lines camelot actually used)."""
    if not _HAS_CAMELOT:
        return []
    try:
        import camelot
        return list(camelot.read_pdf(pdf_path, pages=str(page_num),
                                     flavor=flavor, suppress_stdout=True))
    except Exception as e:
        logger.warning("camelot %s p%d failed: %s", flavor, page_num, e)
        return []


def _camelot_grid(t, page_height: float) -> tuple[list[float], list[float]]:
    """Camelot's own cut lines -> (x_edges, y_edges) in pdfplumber space.

    Camelot reports PDF coordinates with the origin at the BOTTOM-left and y
    growing upward; pdfplumber (and our normalized space) put the origin at the
    TOP-left. So y flips: top = page_height - y.
    """
    cols = list(getattr(t, "cols", []) or [])
    rows = list(getattr(t, "rows", []) or [])
    if not cols or not rows:
        return [], []
    xs = [c[0] for c in cols] + [cols[-1][1]]
    ys_bottom = [r[0] for r in rows] + [rows[-1][1]]
    ys = sorted(page_height - y for y in ys_bottom)
    return sorted(xs), ys


def _camelot_rich(t):
    try:
        return _normalize_plain_table(t.df.values.tolist())
    except Exception:
        return []


def _is_real_grid(xs: list[float], ys: list[float]) -> bool:
    """A real table has at least 2 columns AND 2 rows of cells (>=3 cut lines
    on each axis). A lone rule — a 'Totals' bar, an underline, a page border —
    yields a 1-column or 1-row "table", which is not a table."""
    return len(xs) >= 3 and len(ys) >= 3


def _winning_tables(page):
    """Walk boss-pdf's ladder IN ORDER; take the first strategy giving a real grid.

    Order matters and must be respected: `lines` comes first because when a
    table is genuinely BORDERED, those ruled lines ARE its structure. Deriving
    columns from text spacing instead would ignore the borders the document
    already tells us about, and tends to over-segment. So a bordered table wins
    on its own rules — we never let a text-derived grid outvote real borders.

    Core's bug is that it takes the first strategy returning *any* table, so a
    page carrying a stray rule (that 'Totals' bar) hands lines/lines a
    degenerate 1x2 "table" which wins by default and buries the real one. The
    fix is not to score-maximise across strategies (that would let text/text
    beat genuine borders) — it's to skip *degenerate* results and fall through.
    """
    fallback = ([], None)
    for settings in _TABLE_SETTINGS_OPTIONS:
        try:
            found = page.find_tables(table_settings=settings) or []
        except Exception as e:
            logger.warning("find_tables %s failed: %s", settings, e)
            continue
        if not found:
            continue
        xs, ys = _grid_of(_biggest(found))
        if _is_real_grid(xs, ys):
            return found, settings          # real grid — borders respected first
        if not fallback[0]:
            fallback = (found, settings)    # keep as last resort
    return fallback


def _grid_of(table) -> tuple[list[float], list[float]]:
    """The table's own cut lines, in PDF points: (x_edges, y_edges)."""
    cells = [c for c in table.cells if c]
    if not cells:
        return [], []
    xs = sorted({round(c[0], 2) for c in cells} | {round(c[2], 2) for c in cells})
    ys = sorted({round(c[1], 2) for c in cells} | {round(c[3], 2) for c in cells})
    return xs, ys


def _biggest(tables):
    """The table covering the most area — the real one when stray bits appear."""
    return max(tables, key=lambda t: (t.bbox[2] - t.bbox[0]) * (t.bbox[3] - t.bbox[1]))


def _clean_row_cuts(page, table) -> list[float]:
    """One cut BETWEEN text rows, instead of one above and one below each.

    pdfplumber's text strategy puts a cut at both the top and the bottom of
    every text line. The strip left over between two lines therefore becomes its
    own band with nothing in it: 46 real rows come back as 91 (46 text + 45
    phantom empties). Collapsing each gap to a single cut at its midpoint gives
    exactly one row per text line — no empties to filter out afterwards, and the
    lines you see are the lines you get.
    """
    chars = page.chars
    bands = []
    for row in table.rows:
        top, bottom = float(row.bbox[1]), float(row.bbox[3])
        has_text = any(top <= (c["top"] + c["bottom"]) / 2 < bottom for c in chars)
        bands.append((top, bottom, has_text))
    text_bands = [(t, b) for t, b, has in bands if has]
    if not text_bands:
        return []
    cuts = [text_bands[0][0]]
    for i in range(len(text_bands) - 1):
        cuts.append((text_bands[i][1] + text_bands[i + 1][0]) / 2.0)   # midpoint of the gap
    cuts.append(text_bands[-1][1])
    return cuts


def _best_grid(pdf_path: str, page, page_num: int) -> tuple[str, float, list[float], list[float]]:
    """Boss-pdf's real per-page best-of, but returning the WINNER'S GRID.

    Core picks the highest-scoring of pdfplumber / camelot-lattice /
    camelot-stream. Whichever engine wins is the one that "knows the lines" for
    that page — so that's the grid we must draw. (On a rent roll's totals page
    pdfplumber scores ~2 and camelot-stream wins; showing pdfplumber's grid
    there would show nothing useful.)
    Returns (method, score, x_edges, y_edges) in pdfplumber points.
    """
    h = float(page.height)
    best = ("none", 0.0, [], [])

    tables, _settings = _winning_tables(page)
    if tables:
        t = _biggest(tables)
        try:
            score = _score_table(_extract_rich_table_from_pdfplumber(page, t))
        except Exception:
            score = 0.0
        xs, _raw_ys = _grid_of(t)
        ys = _clean_row_cuts(page, t) or _raw_ys   # one cut between rows, no phantoms
        if score > best[1]:
            best = ("pdfplumber", score, xs, ys)

    # Only reach for camelot when pdfplumber came up weak — the same gate core
    # uses. Camelot drives Ghostscript and is by far the slowest engine, so
    # running it on every page (as this used to) makes detecting a whole
    # document cost N Ghostscript round-trips for nothing.
    if best[1] >= _CAMELOT_SCORE_GATE:
        return best

    for flavor in ("lattice", "stream"):
        for t in _camelot_tables(pdf_path, page_num, flavor):
            score = _score_table(_camelot_rich(t))
            if score <= best[1]:
                continue
            xs, ys = _camelot_grid(t, h)
            if xs and ys:
                best = (f"camelot-{flavor}", score, xs, ys)
    return best


def detect(pdf_bytes: bytes, page_num: int) -> dict:
    """Return the grid of whichever engine boss-pdf uses for this page."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp = f.name
    try:
        with pdfplumber.open(tmp) as pdf:
            n = len(pdf.pages)
            page_num = max(1, min(page_num, n))
            page = pdf.pages[page_num - 1]
            w, h = float(page.width), float(page.height)
            method, score, xs, ys = _best_grid(tmp, page, page_num)
            return {
                "page_count": n,
                "page_width": w,
                "page_height": h,
                "method": method,
                "score": round(score, 1),
                "columns": [x / w for x in xs],
                "rows": [y / h for y in ys],
                "n_cols": max(0, len(xs) - 1),
                "n_rows": max(0, len(ys) - 1),
                "found": bool(xs and ys),
            }
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def _camelot_page_result(pdf_path: str, page_num: int, columns_pt: list[float],
                         chars: int, drop_empty: bool = True) -> PageResult | None:
    """Re-run camelot-stream with the user's calibrated column separators.
    Stream takes explicit column x-positions, which is exactly our calibration."""
    if not _HAS_CAMELOT or len(columns_pt) < 2:
        return None
    inner = [str(round(x, 2)) for x in columns_pt[1:-1]]   # separators, not outer edges
    if not inner:
        return None
    try:
        import camelot
        tabs = list(camelot.read_pdf(pdf_path, pages=str(page_num), flavor="stream",
                                     columns=[",".join(inner)], suppress_stdout=True))
    except Exception as e:
        logger.warning("camelot stream w/ columns p%d failed: %s", page_num, e)
        return None
    rich = []
    for t in tabs:
        rt = _camelot_rich(t)
        if drop_empty:
            rt = _drop_empty_rows(rt)
        if rt:
            rich.append(rt)
    if not rich:
        return None
    return PageResult(page_num=page_num, tables=rich, title_lines=[],
                      method="tablescrape-camelot",
                      score=sum(_score_table(t) for t in rich), char_count=chars)


def _drop_empty_rows(table):
    """Rows where every cell is blank. The grid spans the whole table bbox, so
    the gaps between text lines come back as fully empty rows; core's builder
    only trims them at the top/bottom, not in the middle."""
    return [row for row in table if any((c.text or "").strip() for c in row)]


def _page_result(page, page_num: int, columns: list[float],
                 rows: list[float] | None = None,
                 drop_empty: bool = True) -> PageResult:
    """Re-run boss-pdf's extraction on one page using the user's column lines.

    `rows` is this page's OWN row lines, and only when the user actually edited
    them. Rows drift page to page, so one page's cuts must never be forced onto
    another — pages the user didn't touch keep boss-pdf's per-page detection,
    which measures exactly to that page's true line count.
    """
    w, h = float(page.width), float(page.height)

    # Learn the horizontal strategy boss-pdf would have picked for this page, so
    # untouched pages behave exactly as they do in "Make it Excel".
    _tables, winning = _winning_tables(page)
    settings = dict(winning or {"horizontal_strategy": "text",
                                "intersection_tolerance": 5})
    settings["vertical_strategy"] = "explicit"
    settings["explicit_vertical_lines"] = sorted(c * w for c in columns)

    # Rows: your edits for THIS page if you made any, else this page's own
    # cleaned cuts (one between text rows). Either way they go in explicitly,
    # so the export can't drift from the lines you were shown.
    if rows:
        ys_pt = sorted(r * h for r in rows)
    else:
        ys_pt = _clean_row_cuts(page, _biggest(_tables)) if _tables else []
    if ys_pt:
        settings["horizontal_strategy"] = "explicit"
        settings["explicit_horizontal_lines"] = ys_pt

    try:
        found = page.find_tables(table_settings=settings) or []
    except Exception as e:
        logger.warning("explicit find_tables p%d failed: %s", page_num, e)
        found = []

    rich, bboxes = [], []
    for t in found:
        try:
            rt = _extract_rich_table_from_pdfplumber(page, t)
        except Exception as e:
            logger.warning("rich extract p%d failed: %s", page_num, e)
            continue
        if drop_empty:
            rt = _drop_empty_rows(rt)
        if rt:
            rich.append(rt)
            bboxes.append(t.bbox)

    try:
        titles = _extract_page_title_lines(page, bboxes)
    except Exception:
        titles = []
    try:
        chars = len(page.chars)
    except Exception:
        chars = 0

    return PageResult(page_num=page_num, tables=rich, title_lines=titles,
                      method="tablescrape", score=sum(_score_table(t) for t in rich),
                      char_count=chars)


def scrape_all(pdf_bytes: bytes, columns: list[float],
               rows_by_page: dict[int, list[float]] | None = None,
               drop_empty: bool = True) -> list[PageResult]:
    """Apply the calibrated columns to every page and extract via the core."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp = f.name
    out: list[PageResult] = []
    try:
        with pdfplumber.open(tmp) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                pr = _page_result(page, i, columns,
                                  (rows_by_page or {}).get(i), drop_empty)
                # Mirror the same gate detect() uses: when pdfplumber is weak on
                # this page, let camelot try with the SAME calibrated columns and
                # keep whichever scores better. Without this the export could use
                # a different engine than the one whose grid you were shown, so
                # the drawn lines wouldn't match the result.
                if pr.score < _CAMELOT_SCORE_GATE:
                    alt = _camelot_page_result(
                        tmp, i, sorted(c * float(page.width) for c in columns),
                        pr.char_count, drop_empty)
                    if alt is not None and alt.score > pr.score:
                        pr = alt
                out.append(pr)
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass
    return out


def preview_rows(pages: list[PageResult], limit: int = 400) -> list[list[str]]:
    """Flat page-prefixed rows for the on-screen preview."""
    out: list[list[str]] = []
    for p in pages:
        for t in p.tables:
            for row in t:
                out.append([str(p.page_num)] + [c.text for c in row])
                if len(out) >= limit:
                    return out
    return out


def build_xlsx(pages: list[PageResult]) -> tuple[bytes, int]:
    """Build the workbook with the CORE builder — same per-page sheets, 'All
    pages' sheet, styling and number formats as 'Make it Excel'."""
    wb = _build_workbook(pages)
    buf = io.BytesIO()
    wb.save(buf)
    total = sum(len(t) for p in pages for t in p.tables)
    return buf.getvalue(), total
