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
4. Double-click **`Build Hymnal.cmd`**, then paste a workbook path or a
   Google Sheets URL when prompted. Alternatively, drag a workbook `.xlsx`
   file onto `Build Hymnal.cmd`.

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
