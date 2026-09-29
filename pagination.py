"""
pagination.py

Stage 3.5 of the Hymnal Edition 2 Typst pilot.

Decides which hymns need a forced column break, based on a completed
compile's own <hymn-debug> / <hymn-measured> metadata. Validated rule
(2026-09-23): flag a hymn if it starts in the last column of a page AND
its own measured height would carry it past the bottom of that column.
Confirmed against all 451 pilot hymns — flags exactly {128, 276, 441},
no false positives, no misses.

These constants mirror template.typ's page geometry. If page size,
margins, or columns change in template.typ, update them here too.

Also decides which category-index headings need a forced page break
(first_category_break_violation). The tags <hymn-debug>, <hymn-measured>
and <category-debug> are defined in template.typ; do not remove or rename them.
"""

import json
import subprocess

COLUMN_X_THRESHOLD = 210.0   # pt — x below this = column 1, at/above = column 2
COLUMN_BOTTOM_Y = 581.1      # pt — writable column bottom (A5 height minus margins)


def _pt(value):
    if isinstance(value, str) and value.endswith("pt"):
        return float(value[:-2])
    return float(value)


def run_query(main_typ_path, selector):
    result = subprocess.run(
        ["typst", "query", str(main_typ_path), selector, "--pretty"],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        raise RuntimeError(f"typst query failed for {selector}:\n{result.stderr}")
    return json.loads(result.stdout)


def compute_force_breaks(main_typ_path):
    items = run_query(main_typ_path, "selector(<hymn-debug>).or(<hymn-measured>)")
    positions = {}
    heights = {}
    for item in items:
        v = item["value"]
        if item["label"] == "<hymn-debug>":
            positions[v["hymn"]] = v
        elif item["label"] == "<hymn-measured>":
            heights[v["hymn"]] = v["measured-height"]

    flagged = []
    for hymn_no, pos in positions.items():
        if hymn_no not in heights:
            continue
        x = _pt(pos["x"])
        y = _pt(pos["y"])
        h = _pt(heights[hymn_no])
        in_last_column = x >= COLUMN_X_THRESHOLD
        overflows = (y + h) > COLUMN_BOTTOM_Y
        if in_last_column and overflows:
            flagged.append(hymn_no)

    return sorted(flagged)

CATEGORY_BREAK_FRACTION = 0.7   # heading starting below this fraction of usable height -> new page
PAGE_HEIGHT = 595.28             # pt, A5
PAGE_MARGIN_TB = 14.17           # pt, 0.5 cm — mirrors template.typ


def first_category_break_violation(main_typ_path, already_flagged):
    """Top-most category heading that starts past the threshold and is not
    already flagged, or None. One at a time, top-to-bottom, so a flag can
    never go stale (same reasoning as the hymn loop)."""
    usable = PAGE_HEIGHT - 2 * PAGE_MARGIN_TB
    break_y = PAGE_MARGIN_TB + usable * CATEGORY_BREAK_FRACTION
    for item in run_query(main_typ_path, "<category-debug>"):
        v = item["value"]
        if v["category"] in already_flagged:
            continue
        if _pt(v["y"]) > break_y:
            return v["category"]
    return None
