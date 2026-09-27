# GRC-Hymnal-Compiler

Python → Typst publishing workflow for Hymns We Sing, 2nd Edition.

## Workflow

Workbook
→ Python validation/transformation
→ semantic Typst
→ Typst compiler
→ PDF

## Requirements

- Python
- openpyxl
- Typst

## Build

Run:

    python src/build.py "path/to/workbook.xlsx"
