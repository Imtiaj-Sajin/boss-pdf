"""Regression tests for aging-column header detection.

Written after a real report failed with `list indices must be integers or
slices, not NoneType`: _header_geometry assumed every header token existed and
indexed find()'s result unconditionally, so one unexpected spelling produced an
opaque TypeError instead of a usable message.

Synthetic header rows only — no customer PDFs required. Stdlib unittest so it
runs in the app venv without adding a test dependency:

    .venv\\Scripts\\python.exe test_header_geometry.py
"""
import unittest

from ar_to_excel import BUCKET_KEYS, _header_geometry


def row(*labels_at):
    """Build a fake extracted-word line: (text, x_centre) -> word dicts."""
    return [{"text": t, "x0": cx - 4, "x1": cx + 4} for t, cx in labels_at]


WELL_FORMED = row(
    ("Original", 218), ("Amount", 247), ("Open", 287), ("Amount", 311),
    ("Current", 388),
    ("1", 455), ("-", 465), ("30", 476),
    ("31", 527), ("-", 537), ("60", 548),
    ("61", 599), ("-", 609), ("90", 620),
    ("91", 671), ("-", 681), ("120", 692),
    ("Over", 729), ("120", 741),
    ("Remark", 758),
)

MERGED_LABELS = row(
    ("Current", 388), ("1-30", 465), ("31-60", 537),
    ("61-90", 609), ("91-120", 681), ("Over120", 735), ("Remark", 758),
)

# The real failing file: every multi-digit number lost its leading digits, so
# the header reads "Current 1 - 0 1 - 0 1 - 0 1 - 0 Over 0 Remark".
DIGITS_MANGLED = row(
    ("Current", 387.9),
    ("1", 460.4), ("-", 465.5), ("0", 470.6),
    ("1", 532.4), ("-", 537.5), ("0", 542.6),
    ("1", 604.4), ("-", 609.5), ("0", 614.6),
    ("1", 676.4), ("-", 681.5), ("0", 686.6),
    ("Over", 728.7), ("0", 740.6),
    ("Remark", 757.6),
)

VARIANTS = [("split digits", WELL_FORMED),
            ("merged labels", MERGED_LABELS),
            ("mangled digits", DIGITS_MANGLED)]


class HeaderGeometry(unittest.TestCase):

    def test_all_six_columns_are_located(self):
        for label, line in VARIANTS:
            with self.subTest(label):
                centers, current_x, remark_x = _header_geometry(line)
                self.assertEqual(set(centers), set(BUCKET_KEYS))
                self.assertEqual(current_x, centers["Current"])
                self.assertGreater(remark_x, centers["Over120"])

    def test_columns_are_ordered_left_to_right(self):
        for label, line in VARIANTS:
            with self.subTest(label):
                centers, _, _ = _header_geometry(line)
                xs = [centers[k] for k in BUCKET_KEYS]
                self.assertEqual(xs, sorted(xs), "columns must run left to right")
                gaps = [b - a for a, b in zip(xs, xs[1:])]
                # Snapping to the nearest centre is only meaningful if the
                # centres are actually apart.
                self.assertGreater(min(gaps), 20, f"columns too close: {gaps}")

    def test_mangled_header_lands_near_the_real_columns(self):
        """The positional fallback must put columns within a few px of where a
        well-formed header puts them — otherwise amounts snap to the wrong
        bucket, which is worse than failing outright."""
        good, _, _ = _header_geometry(WELL_FORMED)
        bad, _, _ = _header_geometry(DIGITS_MANGLED)
        for k in BUCKET_KEYS:
            self.assertLess(abs(good[k] - bad[k]), 12, (k, good[k], bad[k]))

    def test_unrecognisable_header_is_actionable_not_a_typeerror(self):
        line = row(("Current", 388), ("Remark", 758))
        with self.assertRaises(ValueError) as cm:
            _header_geometry(line)
        msg = str(cm.exception)
        self.assertIn("1-30", msg)
        self.assertIn("Over120", msg)
        self.assertIn("header row was read as", msg)

    def test_missing_current_is_reported_not_crashed(self):
        line = row(("Something", 388), ("Else", 470), ("Remark", 758))
        with self.assertRaises(ValueError):
            _header_geometry(line)


if __name__ == "__main__":
    unittest.main(verbosity=2)
