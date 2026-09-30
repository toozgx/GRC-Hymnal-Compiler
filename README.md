# GRC-Hymnal-Compiler

Python → Typst publishing workflow for *Hymns We Sing*, 2nd Edition.

    Workbook → Python validation/transformation → Typst compiler → PDF

The workbook is the editorial source of truth. Python validates it, assigns
hymn numbers, generates attributions and indexes, and drives Typst, which
handles all layout and pagination.

## Quick start (Windows 10/11)

1. Download or clone this repository.
2. Confirm the `fonts\` folder contains the PT Sans `.ttf` files
   (Regular, Bold, Italic, Bold Italic).
3. Double-click **`Setup.cmd`**. It downloads Typst and a portable Python
   into `tools\` and installs openpyxl. This needs an internet connection
   and is done once per machine.

## Build

Two ways to work with the workbook:

- **Google Sheets (recommended).** Easier for shared editing, and the
  build downloads a fresh copy each time.
- **Local workbook.** A locally saved copy of the template, edited in
  Excel or a similar program.

The build accepts either a Google Sheets link or a local file path.

### Option A: Google Sheets

1. Set the sheet's sharing to "Anyone with the link" → Viewer.
2. Run: `Build Hymnal.cmd`

- input: `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit` (copy from Google Sheet shared link)

The sheet is downloaded into the `snapshots` folder with a timestamp,
and the build runs from that copy. Keep the snapshot together with the
PDF it produced.

### Option B: Local workbook

1. Copy `Hymnal Master Data Template.xlsx` and rename the copy
   (for example `Hymnal Master Data.xlsx`).
2. Fill in the sheets. Do not rename the sheets or change the column
   headings. Delete the sample rows, or replace them with real content.
3. Save and close the file.
4. Run: `Build Hymnal.cmd`:

- input: file path to Hymnal Master Data.xlsx (e.g. `C:\User\Downloads\GRC-Hymnal-Compiler-main\Hymnal Master Data.xlsx`)

   If the file is stored elsewhere, input the full path along with filename and extension (.xlsx).

To rebuild after making changes, save the workbook and run the same
command again.

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
downloads a timestamped copy into `snapshots\`. Keep the snapshot with the
PDF it produced, so any PDF can be reproduced.

## Folder layout

    HymnalCompiler\
      Setup.cmd            one-time setup (downloads tools\)
      Build Hymnal.cmd     the everyday entry point
      src\                 build.py, transform.py, pagination.py,
                           validate_workbook.py, main.typ, template.typ
      fonts\               PT Sans (committed to the repository)
      tools\               Typst + portable Python (created by Setup.cmd,
                           NOT committed)
      snapshots\           workbook copies from Google Sheets (not committed)

Build output (`hymnal.pdf`, `hymnal_data.json`) is written to `src\`.

## What is and is not in the repository

Committed: `src\`, `fonts\`, `Setup.cmd`, `Build Hymnal.cmd`, this README.

Not committed (see `.gitignore`): `tools\`, `snapshots\`, and generated
output. They are recreated by `Setup.cmd` or by a build.

## Typst version

Pagination is calibrated against a specific Typst version and the bundled
fonts. A different Typst release can shift line or column breaks.

- `Setup.cmd` installs the newest Typst by default.
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

With Python, openpyxl and Typst installed and on `PATH`:

    python src/build.py "path/to/workbook.xlsx"

## Requirements

- Windows 10/11, 64-bit
- Internet access for `Setup.cmd` (and for Google Sheets input)
