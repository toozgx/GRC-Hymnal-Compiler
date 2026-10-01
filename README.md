# GRC-Hymnal-Compiler

Python → Typst publishing workflow for *Hymns We Sing*, 2nd Edition.

    Workbook → Python validation/transformation → Typst compiler → PDF

The workbook is the editorial source of truth. Python validates it, assigns
hymn numbers, generates attributions and indexes, and drives Typst, which
handles all layout and pagination.

## Quick start (Windows 10/11)

1. Download or clone this repository.
2. Open the `HymnalCompiler\` subfolder. Confirm its `fonts\` folder contains the PT Sans `.ttf` files
   (Regular, Bold, Italic, Bold Italic).
3. Double-click **`Setup.cmd`**. It downloads Typst and a portable Python
   into `tools\` and installs openpyxl. This needs an internet connection
   and is done once per machine.

## Build

The build accepts either a Google Sheets link or a local Excel workbook.

### Option A: Google Sheets — Recommended

Google Sheets is recommended for shared editing. Each build downloads a fresh copy of the sheet, so the published PDF is always based on a timestamped snapshot.

1. Set the sheet's sharing to **Anyone with the link → Viewer**.
2. Run `Build Hymnal.cmd`.
3. When prompted, enter the Google Sheets link:

   `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit`

   Copy this from the Google Sheet's shared link.

The workbook is downloaded to the `src\snapshots` folder with a timestamp, and the build runs from that snapshot.

**Optional: keep the snapshot together with the PDF it produced for better version control.**

### Option B: Local workbook

A local copy of the template can be edited in Excel. This is useful when the workbook does not need to be shared through Google Sheets.

1. Copy `Hymnal Master Data Template.xlsx` and rename the copy, for example:

   `Hymnal Master Data.xlsx`

   It is recommended to keep the workbook in the same folder as `Build Hymnal.cmd`.

2. Fill in the workbook. **Do not rename the tabs or change the column headings.** Delete the sample rows or replace them with real content.

3. Save and close the workbook.

4. Run `Build Hymnal.cmd`.

5. When prompted, enter the workbook filename:

   `Hymnal Master Data.xlsx`

   If the workbook is stored elsewhere, enter its full path, including the filename and extension. For example:

   `C:\Users\Downloads\GRC-Hymnal-Compiler-main\Hymnal Master Data.xlsx`

To rebuild after making changes, save and close the workbook, then run `Build Hymnal.cmd` again.

### What the build does

1. Validates the workbook. Errors stop the build; warnings are shown
   for review.
2. Generates `hymnal_data.json`.
3. Compiles the PDF with Typst, resolving page and column breaks
   automatically.

When the build succeeds, the PDF opens automatically. Nothing needs to be
installed system-wide, and no administrator rights are required.

## Google Sheets input

The Sheet must be shared as **Anyone with the link → Viewer**. The build
downloads a timestamped copy into `src\snapshots\`. Keep the snapshot with the
PDF it produced, so any PDF can be reproduced.

## Folder layout

    HymnalCompiler\
      Hymnal Master Data Template.xlsx (data template)
      Setup.cmd (one-time setup - downloads tools)
      Build Hymnal.cmd (the everyday entry point)
      src\           build.py, transform.py, pagination.py,
                     validate_workbook.py, workbook_schema.py,
                     main.typ, template.typ
        snapshots\      workbook copies from Google Sheets (not committed)
      fonts\         PT Sans (committed to the repository)
      tools\         Typst + portable Python (created by Setup.cmd, not committed)
      tests\         test_workbook_validation.py (debug tool)
                             

The PDF is written to `HymnalCompiler\hymnal.pdf`. Generated data is written
to `HymnalCompiler\src\hymnal_data.json`.

## What is and is not in the repository

Committed: `src\`, `fonts\`, `Setup.cmd`, `Build Hymnal.cmd`, this README.

Not committed (see `.gitignore`): `tools\`, `snapshots\`, and generated
output. They are recreated by `Setup.cmd` or by a build.

## Typst version

Pagination is calibrated against a specific Typst version and the bundled
fonts. A different Typst release can shift line or column breaks.

- `Setup.cmd` currently pins Typst to **0.15.1**. Leaving `TYPST_VERSION`
  blank selects the newest release instead.
- Once the layout is settled, run `tools\typst\typst.exe --version`, then
  set `TYPST_VERSION` at the top of `Setup.cmd` to that number so every
  machine builds identically.
- To change versions: delete `tools\typst\`, edit `TYPST_VERSION`, run
  `Setup.cmd` again, and re-check the pagination.

## Troubleshooting

**Validation errors stop the build.** Read the ERRORS list, fix the
workbook, and rebuild. Warnings do not stop the build but should be reviewed.

**"ModuleNotFoundError" or "ImportError".** Run `Setup.cmd` again. If it
persists, open `tools\python\python*._pth` and confirm it lists
`..\..\src` and an uncommented `import site`.

**Setup fails while downloading.** Check the internet connection and whether
antivirus or a firewall is blocking PowerShell downloads. Re-running
`Setup.cmd` skips steps that already finished.

**Text appears in the wrong font, or pagination looks different.** The PT
Sans files are missing from `fonts\`, or Typst is reading a different
installed copy of the font.

**Windows SmartScreen warns about `Build Hymnal.cmd` or `typst.exe`.** Choose
"More info" → "Run anyway" if the files came from this repository or from
Typst's official releases.

## Manual build (developers)

From the `HymnalCompiler\` folder, with Python, openpyxl and Typst installed
and on `PATH` (set `TYPST_FONT_PATHS` to the bundled `fonts\` folder):

    python src/build.py "path/to/workbook.xlsx"

To run the workbook validation tests from that same folder:

    python -m unittest discover -s tests -v

## Requirements

- Windows 10/11, 64-bit
- Internet access for `Setup.cmd` (and for Google Sheets input)

## Development and AI Assistance Disclosure

This project was developed using a human-directed, AI-assisted workflow.
Generative AI (OpenAI ChatGPT & Anthropic Claude) was used extensively as a programming and
technical-development assistant, including for software architecture
discussion, Python and Typst development, debugging, code review,
data-transformation logic, pagination analysis, and documentation.

The project maintainer remains responsible for the project requirements,
editorial decisions, hymnal source data, attribution decisions, validation,
testing, and final acceptance of the implementation. AI-generated suggestions
were reviewed, modified, tested, or rejected as appropriate.

AI assistance does not constitute authorship or editorial authority over the
hymnal's content.
