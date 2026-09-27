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

        capture_output=True, text=True,

    )

    if result.returncode != 0:

        raise RuntimeError(f"typst query failed for {selector}:\n{result.stderr}")

    return json.loads(result.stdout)





def compute_force_breaks(main_typ_path):

    positions = {item["value"]["hymn"]: item["value"] for item in run_query(main_typ_path, "<hymn-debug>")}

    heights = {item["value"]["hymn"]: item["value"]["measured-height"] for item in run_query(main_typ_path, "<hymn-measured>")}



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