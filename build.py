"""
build.py

Full pilot build: workbook -> hymnal_data.json -> PDF, via the local Typst
compiler (`typst` must be on PATH).

Usage:
    python3 build.py "path/to/workbook.xlsx"

Expects template.typ and main.typ in the same folder as this script.
Writes hymnal_data.json and hymnal_pilot.pdf into that same folder.
"""

import sys
import json
import subprocess
from pathlib import Path

from transform import load_workbook, build_hymn_data, build_indexes, load_title_page

from pagination import compute_force_breaks


def main():
    if len(sys.argv) != 2:
        print("Usage: python3 build.py <path-to-workbook.xlsx>")
        sys.exit(1)

    workbook_path = sys.argv[1]
    here = Path(__file__).resolve().parent
    data_path = here / "hymnal_data.json"
    main_typ = here / "main.typ"
    output_pdf = here / "hymnal_pilot.pdf"

    hymns, sections, contributors, categories = load_workbook(workbook_path)
    title_page = load_title_page(workbook_path)
    hymn_data = build_hymn_data(hymns, sections, contributors)
    for hymn in hymn_data:
        if hymn["id"] == "MOSWAN":
            for section in hymn["sections"]:
                print("SECTION:", section["type"])
                print(repr(section["text"]))
    indexes, index_warnings = build_indexes(hymns, hymn_data, contributors, categories)

    for warning in index_warnings:
        print(f"INDEX WARNING: {warning}")

    with open(data_path, "w", encoding="utf-8") as f:
        json.dump({"hymns": hymn_data, "indexes": indexes, "title_page": title_page}, f, ensure_ascii=False, indent=2)
    print(f"[1/4] Wrote {len(hymn_data)} active hymn(s) to {data_path.name}")

    print(f"[2/4] Compiling {main_typ.name} (pass 1, measurement) ...")
    result = subprocess.run(
        ["typst", "compile", str(main_typ), str(output_pdf)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print("Typst compilation FAILED (pass 1):")
        print(result.stderr or result.stdout)
        sys.exit(1)

    print("[3/*] Resolving forced column breaks (iterative) ...")
    flagged = set()
    MAX_ITERATIONS = 10
    for iteration in range(1, MAX_ITERATIONS + 1):
        new_flagged = set(compute_force_breaks(main_typ))
        added = new_flagged - flagged
        flagged |= new_flagged

        print(f"       iteration {iteration}: total flagged = {sorted(flagged)}"
              + (f"  (new: {sorted(added)})" if added else "  (no change)"))

        if not added:
            print(f"       Converged after {iteration} iteration(s).")
            break

        for hymn in hymn_data:
            hymn["force_break_before"] = hymn["printed_no"] in flagged

        with open(data_path, "w", encoding="utf-8") as f:
            json.dump({"hymns": hymn_data, "indexes": indexes}, f, ensure_ascii=False, indent=2)

        result = subprocess.run(
            ["typst", "compile", str(main_typ), str(output_pdf)],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(f"Typst compilation FAILED (iteration {iteration}):")
            print(result.stderr or result.stdout)
            sys.exit(1)
    else:
        print(f"WARNING: did not converge within {MAX_ITERATIONS} iterations.")
        print(f"         Final flag set: {sorted(flagged)} — inspect manually before trusting output.")

    print(f"[final] Compiling {main_typ.name} -> {output_pdf.name} ...")
    result = subprocess.run(
        ["typst", "compile", str(main_typ), str(output_pdf)],
        capture_output=True, text=True,
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
