# pdf2sheet

Convert PDFs organised in a multi-level folder hierarchy into a single formatted Excel workbook. Each folder combination becomes a sheet containing section headers and scaled images of the PDF pages.

## How it works

```
<BASE_DIR>/
  <LEVEL-1>/
    <LEVEL-2>/        ← any depth supported
      F1.pdf
      Length.pdf
```

For every folder combination the script:

1. Rasterises each configured PDF to PNG.
2. Creates an Excel sheet with a styled header row and the image embedded below it.

## Prerequisites

### System — Poppler

`pdf2image` requires the Poppler PDF rendering library.

| OS | Command |
|---|---|
| Ubuntu / Debian | `sudo apt install poppler-utils` |
| macOS (Homebrew) | `brew install poppler` |
| Windows | Download from [poppler releases](https://github.com/oschwartz10612/poppler-windows/releases) and add `bin/` to `PATH` |

### uv

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

## Installation

```bash
git clone <repo-url>
cd pdf2sheet
uv sync
```

## Usage

```bash
# Run with all defaults
uv run pdf2sheet

# Equivalently
uv run python generate_excel.py
```

Every option can be overridden from the command line — no need to edit the script for one-off runs.

```
uv run pdf2sheet --help
```

```
usage: generate_excel.py [-h]

paths:
  --base-dir PATH        Root folder containing the input PDFs
  --output-dir PATH      Folder for rasterised PNGs
  --excel-output PATH    Output Excel file path

folder hierarchy:
  --level A,B,C          Comma-separated folder names for one hierarchy level.
                         Repeat the flag once per level.

PDF sections:
  --section STEM:LABEL:ROW
                         PDF section as 'stem:label:header_row'.
                         Repeat for each section.

conversion:
  --dpi INT              Rasterisation DPI
  --page INT             PDF page index to extract, 0-based

layout:
  --img-max-width PX     Scale down images wider than this
  --header-span-col COL  Last column letter of merged header rows
  --num-columns INT      Number of columns to apply width formatting
  --column-width INT     Excel column width in character units

header style:
  --header-bg-color HEX  Background colour without '#'
  --header-font-size INT Font size in points
```

### Examples

```bash
# Custom paths
uv run pdf2sheet --base-dir ./data --excel-output ./report.xlsx

# Two-level hierarchy
uv run pdf2sheet --level "G1,G2,GF" --level "TS,WP,WS"

# Three-level hierarchy
uv run pdf2sheet --level "2024,2025" --level "G1,G2" --level "TS,WP"

# Single-level hierarchy
uv run pdf2sheet --level "GroupA,GroupB,GroupC"

# Custom PDF sections (stem:label:header_row)
uv run pdf2sheet --section "chart:Chart:1" --section "table:Table:30"

# Mix of overrides
uv run pdf2sheet \
  --base-dir ./experiments \
  --level "control,treatment" \
  --level "pre,post" \
  --section "results:Results:1" \
  --dpi 50 \
  --excel-output ./experiments/report.xlsx
```

## Configuration reference

All options have defaults defined as constants at the top of `generate_excel.py`. Edit them to change the permanent defaults; use CLI flags for one-off overrides.

### Paths

| Constant | CLI flag | Default |
|---|---|---|
| `BASE_DIR` | `--base-dir` | `./data` |
| `OUTPUT_DIR` | `--output-dir` | `./png_output` |
| `EXCEL_OUTPUT` | `--excel-output` | `./output.xlsx` |

### Folder hierarchy

| Constant | CLI flag | Description |
|---|---|---|
| `LEVELS` | `--level` (repeatable) | List of lists of folder names. Each `--level "A,B"` adds one depth level. |

### PDF sections

| Constant | CLI flag | Format |
|---|---|---|
| `PDF_SECTIONS` | `--section` (repeatable) | `stem:label:header_row` — e.g. `F1:F1:1` |

### Conversion

| Constant | CLI flag | Default | Description |
|---|---|---|---|
| `PDF_DPI` | `--dpi` | `50` | Rasterisation resolution |
| `PDF_PAGE` | `--page` | `0` | Page index to extract (0 = first) |

### Layout

| Constant | CLI flag | Default | Description |
|---|---|---|---|
| `IMG_MAX_WIDTH` | `--img-max-width` | `1500` | Max image width in pixels before scaling |
| `HEADER_SPAN_COL` | `--header-span-col` | `P` | Last column of merged header rows |
| `NUM_COLUMNS` | `--num-columns` | `16` | Columns to apply width formatting |
| `COLUMN_WIDTH` | `--column-width` | `15` | Excel column width (character units) |

### Header style

| Constant | CLI flag | Default |
|---|---|---|
| `HEADER_BG_COLOR` | `--header-bg-color` | `EFEFEF` |
| `HEADER_FONT_SIZE` | `--header-font-size` | `12` |
| `HEADER_FONT_BOLD` | *(edit constant)* | `True` |

## Project structure

```
.
├── generate_excel.py   # main script — constants at the top, CLI flags override them
├── pyproject.toml      # project metadata and dependencies
├── README.md
├── .gitignore
├── png_output/         # auto-created; rasterised PNGs (git-ignored)
└── output.xlsx         # generated workbook (git-ignored)
```

## License

MIT
