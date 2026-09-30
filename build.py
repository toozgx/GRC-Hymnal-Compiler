"""
build.py

Full pilot build: workbook -> hymnal_data.json -> PDF, via the local Typst
compiler (`typst` must be on PATH).

Usage:
    python build.py "path/to/workbook.xlsx"
    python build.py "https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit"

If given a Google Sheets URL (sharing set to "Anyone with the link can
view"), the Sheet is downloaded as .xlsx into a "snapshots" folder next to
this script, with a timestamp in the file name, and the build runs from that
copy. Keep the snapshot with the PDF it produced.

Expects template.typ and main.typ in the same folder as this script.
Writes hymnal_data.json and hymnal.pdf into that same folder.
"""

import sys
import re
import json
import subprocess
import shutil
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

from transform import load_workbook, build_hymn_data, build_indexes, load_title_page

from pagination import compute_force_breaks, first_category_break_violation


def fetch_workbook(source, here):
    """Returns a local .xlsx path. Local paths pass through unchanged;
    Google Sheets URLs are downloaded to snapshots/ first."""
    match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", source)
    if not match:
        return source  # not a Sheets URL - treat as a local file path

    sheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    snapshots = here / "snapshots"
    snapshots.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = snapshots / f"workbook_{stamp}.xlsx"

    print(f"Downloading Google Sheet ({sheet_id}) ...")
    try:
        request = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=60) as response:
            data = response.read()
    except urllib.error.URLError as err:
        print(f"Download FAILED: {err}")
        print("Check the internet connection and that the Sheet is shared as "
              "'Anyone with the link' -> Viewer.")
        sys.exit(1)

    # A real .xlsx is a zip file and starts with "PK". Anything else is
    # almost certainly a Google login page (Sheet not shared publicly).
    if not data.startswith(b"PK"):
        print("Download did not return a spreadsheet file.")
        print("Most likely the Sheet is not shared as 'Anyone with the link' -> Viewer.")
        sys.exit(1)

    target.write_bytes(data)
    print(f"Saved snapshot: {target}")
    return str(target)


def write_data(path, hymn_data, indexes, title_page, category_breaks):
    """Single place that writes hymnal_data.json, so no write can forget
    the category_breaks list."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "hymns": hymn_data,
                "indexes": indexes,
                "title_page": title_page,
                "category_breaks": category_breaks,
            },
            f, ensure_ascii=False, indent=2,
        )


def main():
    if len(sys.argv) != 2:
        print("Usage: python3 build.py <path-to-workbook.xlsx | Google Sheets URL>")
        sys.exit(1)

    if shutil.which("typst") is None:
        print("Typst was not found on PATH. Install it or add it to PATH, then run again.")
        sys.exit(1)

    here = Path(__file__).resolve().parent
    workbook_path = fetch_workbook(sys.argv[1], here)

    check = subprocess.run([sys.executable, str(here / "validate_workbook.py"), workbook_path])
    if check.returncode != 0:
        print("Workbook validation found ERRORS - build stopped. Fix them and run again.")
        sys.exit(1)
    data_path = here / "hymnal_data.json"
    main_typ = here / "main.typ"
    output_pdf = here / "hymnal.pdf"

    hymns, sections, contributors, categories = load_workbook(workbook_path)
    title_page = load_title_page(workbook_path)
    hymn_data = build_hymn_data(hymns, sections, contributors)
    flag_total = sum(len(h["flags"]) for h in hymn_data)
    print(f"Attribution flags: {flag_total}")
    for h in hymn_data:
        for f in h["flags"]:
            print(f"  FLAG #{h['printed_no']} {h['title']}: {f}")
    indexes, index_warnings = build_indexes(hymns, hymn_data, contributors, categories)

    for warning in index_warnings:
        print(f"INDEX WARNING: {warning}")

    category_breaks = []
    write_data(data_path, hymn_data, indexes, title_page, category_breaks)
    print(f"[1] Wrote {len(hymn_data)} active hymn(s) to {data_path.name}")

    print(f"[2] Compiling {main_typ.name} (pass 1, measurement) ...")
    result = subprocess.run(
        ["typst", "compile", str(main_typ), str(output_pdf)],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        print("Typst compilation FAILED (pass 1):")
        print(result.stderr or result.stdout)
        sys.exit(1)

    print("[3] Resolving category-index page breaks (one at a time, top to bottom) ...")
    for iteration in range(1, 101):
        try:
            violation = first_category_break_violation(main_typ, category_breaks)
        except RuntimeError as err:
            print(f"Typst query FAILED (category breaks, iteration {iteration}):")
            print(err)
            sys.exit(1)
        if violation is None:
            print(f"       Converged after {iteration} iteration(s). Breaks before: {category_breaks}")
            break

        category_breaks.append(violation)
        print(f"       iteration {iteration}: adding break before category '{violation}'")

        write_data(data_path, hymn_data, indexes, title_page, category_breaks)

    else:
        print("WARNING: category breaks did not converge within 100 iterations.")
        print(f"         Final list: {category_breaks} — inspect manually before trusting output.")

    print("[4] Resolving forced column breaks (one at a time, top to bottom) ...")
    flagged = set()
    MAX_ITERATIONS = 100
    for iteration in range(1, MAX_ITERATIONS + 1):
        try:
            violations = sorted(set(compute_force_breaks(main_typ)) - flagged)
        except RuntimeError as err:
            print(f"Typst query FAILED (column breaks, iteration {iteration}):")
            print(err)
            sys.exit(1)
        if not violations:
            print(f"       Converged after {iteration} iteration(s). Flagged: {sorted(flagged)}")
            break

        first = violations[0]
        flagged.add(first)
        print(f"       iteration {iteration}: adding break before hymn {first}"
              f"  (total flagged = {sorted(flagged)})")

        for hymn in hymn_data:
            hymn["force_break_before"] = hymn["printed_no"] in flagged

        write_data(data_path, hymn_data, indexes, title_page, category_breaks)

    else:
        print(f"WARNING: did not converge within {MAX_ITERATIONS} iterations.")
        print(f"         Final flag set: {sorted(flagged)} — inspect manually before trusting output.")

    print(f"[5] Compiling {main_typ.name} -> {output_pdf.name} ...")
    result = subprocess.run(
        ["typst", "compile", str(main_typ), str(output_pdf)],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        print("Typst compilation FAILED (final):")
        print(result.stderr or result.stdout)
        sys.exit(1)
    if result.stderr:
        print("Typst warnings:")
        print(result.stderr)

    print(f"Done: {output_pdf}")

if __name__ == "__main__":
    main()
