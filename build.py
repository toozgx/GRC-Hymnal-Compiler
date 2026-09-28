"""
build.py

Full pilot build: workbook -> hymnal_data.json -> PDF, via the local Typst
compiler (`typst` must be on PATH).

Usage:
    python build.py "path/to/workbook.xlsx"

Expects template.typ and main.typ in the same folder as this script.
Writes hymnal_data.json and hymnal.pdf into that same folder.
"""

import sys
import json
import subprocess
from pathlib import Path

from transform import load_workbook, build_hymn_data, build_indexes, load_title_page

from pagination import compute_force_breaks, first_category_break_violation


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
        print("Usage: python3 build.py <path-to-workbook.xlsx>")
        sys.exit(1)

    workbook_path = sys.argv[1]
    here = Path(__file__).resolve().parent
    data_path = here / "hymnal_data.json"
    main_typ = here / "main.typ"
    output_pdf = here / "hymnal.pdf"

    hymns, sections, contributors, categories = load_workbook(workbook_path)
    title_page = load_title_page(workbook_path)
    hymn_data = build_hymn_data(hymns, sections, contributors)
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
        violation = first_category_break_violation(main_typ, category_breaks)
        if violation is None:
            print(f"       Converged after {iteration} iteration(s). Breaks before: {category_breaks}")
            break

        category_breaks.append(violation)
        print(f"       iteration {iteration}: adding break before category '{violation}'")

        write_data(data_path, hymn_data, indexes, title_page, category_breaks)

        result = subprocess.run(
            ["typst", "compile", str(main_typ), str(output_pdf)],
            capture_output=True, text=True, encoding="utf-8",
        )
        if result.returncode != 0:
            print(f"Typst compilation FAILED (category breaks, iteration {iteration}):")
            print(result.stderr or result.stdout)
            sys.exit(1)
    else:
        print("WARNING: category breaks did not converge within 100 iterations.")
        print(f"         Final list: {category_breaks} — inspect manually before trusting output.")

    print("[4] Resolving forced column breaks (one at a time, top to bottom) ...")
    flagged = set()
    MAX_ITERATIONS = 100
    for iteration in range(1, MAX_ITERATIONS + 1):
        violations = sorted(set(compute_force_breaks(main_typ)) - flagged)
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

        result = subprocess.run(
            ["typst", "compile", str(main_typ), str(output_pdf)],
            capture_output=True, text=True, encoding="utf-8",
        )
        if result.returncode != 0:
            print(f"Typst compilation FAILED (iteration {iteration}):")
            print(result.stderr or result.stdout)
            sys.exit(1)
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
