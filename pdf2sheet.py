import argparse
from dataclasses import dataclass
from itertools import product
from pathlib import Path

from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

# ── Defaults (used when no CLI argument is provided) ──────────────────────────

BASE_DIR     = Path(__file__).parent / "data"
OUTPUT_DIR   = Path(__file__).parent / "png_output"
EXCEL_OUTPUT = Path(__file__).parent / "output.xlsx"

# Folder hierarchy fallback — used only when --level is given explicitly
# or when auto-discovery finds nothing in BASE_DIR.
# Leave empty ([]) to rely entirely on auto-discovery.
LEVELS = []

# PDF sections: (pdf_filename_stem, excel_label, header_row)
PDF_SECTIONS = [
    ("F1",     "F1",     1),
    ("Length", "Length", 23),
]

PDF_DPI          = 50
PDF_PAGE         = 0
IMG_MAX_WIDTH    = 1500
HEADER_SPAN_COL  = "P"
NUM_COLUMNS      = 16
COLUMN_WIDTH     = 15
HEADER_BG_COLOR  = "EFEFEF"
HEADER_FONT_BOLD = True
HEADER_FONT_SIZE = 12

SHEET_NAME_REPLACEMENTS = {"[": "(", "]": ")"}


# ── Config dataclass ──────────────────────────────────────────────────────────

@dataclass
class Config:
    base_dir:                Path
    output_dir:              Path
    excel_output:            Path
    levels:                  list   # list[list[str]]
    pdf_sections:            list   # list[tuple[str, str, int]]
    pdf_dpi:                 int
    pdf_page:                int
    img_max_width:           int
    header_span_col:         str
    num_columns:             int
    column_width:            int
    header_bg_color:         str
    header_font_bold:        bool
    header_font_size:        int
    sheet_name_replacements: dict


# ── CLI ───────────────────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert PDFs in a multi-level folder hierarchy into a formatted Excel workbook.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  # run with all defaults (hierarchy auto-discovered from BASE_DIR)
  uv run pdf2sheet

  # custom paths
  uv run pdf2sheet --base-dir ./data --excel-output ./report.xlsx

  # two-level hierarchy (overrides auto-discovery)
  uv run pdf2sheet --level "G1,G2,GF" --level "TS,WP,WS"

  # three-level hierarchy
  uv run pdf2sheet --level "2024,2025" --level "G1,G2" --level "TS,WP"

  # custom PDF sections (stem:label:header_row)
  uv run pdf2sheet --section "chart:Chart:1" --section "table:Table:30"
        """,
    )

    g = parser.add_argument_group("paths")
    g.add_argument("--base-dir",     type=Path, default=BASE_DIR,
                   metavar="PATH", help=f"Root folder containing the input PDFs (default: {BASE_DIR})")
    g.add_argument("--output-dir",   type=Path, default=OUTPUT_DIR,
                   metavar="PATH", help=f"Folder for rasterised PNGs (default: {OUTPUT_DIR})")
    g.add_argument("--excel-output", type=Path, default=EXCEL_OUTPUT,
                   metavar="PATH", help=f"Output Excel file path (default: {EXCEL_OUTPUT})")

    g = parser.add_argument_group("folder hierarchy")
    g.add_argument(
        "--level", action="append", metavar="A,B,C",
        help="Comma-separated folder names for one hierarchy level. "
             "Repeat the flag once per level. "
             "By default the hierarchy is discovered automatically from BASE_DIR.",
    )

    g = parser.add_argument_group("PDF sections")
    g.add_argument(
        "--section", action="append", metavar="STEM:LABEL:ROW",
        help="PDF section as 'stem:label:header_row'. Repeat for each section. "
             f"Default: {[f'{s}:{l}:{r}' for s, l, r in PDF_SECTIONS]}",
    )

    g = parser.add_argument_group("conversion")
    g.add_argument("--dpi",  type=int, default=PDF_DPI,
                   help=f"Rasterisation DPI (default: {PDF_DPI})")
    g.add_argument("--page", type=int, default=PDF_PAGE,
                   help=f"PDF page index to extract, 0-based (default: {PDF_PAGE})")

    g = parser.add_argument_group("layout")
    g.add_argument("--img-max-width",   type=int, default=IMG_MAX_WIDTH, metavar="PX",
                   help=f"Scale down images wider than this in pixels (default: {IMG_MAX_WIDTH})")
    g.add_argument("--header-span-col", default=HEADER_SPAN_COL, metavar="COL",
                   help=f"Last column letter of merged header rows (default: {HEADER_SPAN_COL})")
    g.add_argument("--num-columns",     type=int, default=NUM_COLUMNS,
                   help=f"Number of columns to apply width formatting (default: {NUM_COLUMNS})")
    g.add_argument("--column-width",    type=int, default=COLUMN_WIDTH,
                   help=f"Excel column width in character units (default: {COLUMN_WIDTH})")

    g = parser.add_argument_group("header style")
    g.add_argument("--header-bg-color",  default=HEADER_BG_COLOR, metavar="HEX",
                   help=f"Header background colour without '#' (default: {HEADER_BG_COLOR})")
    g.add_argument("--header-font-size", type=int, default=HEADER_FONT_SIZE,
                   help=f"Header font size in points (default: {HEADER_FONT_SIZE})")

    return parser


def discover_levels(base_dir: Path, pdf_stems: list) -> list:
    """
    Scan base_dir and infer the folder hierarchy by locating every directory
    that contains at least one of the configured PDFs.

    Returns a list of lists, one per depth level, each containing the unique
    folder names found at that depth (sorted alphabetically).

    Example — given:
        data/[G1]/[TS]/F1.pdf
        data/[G1]/[WP]/F1.pdf
        data/[G2]/[TS]/F1.pdf
    Returns:
        [["[G1]", "[G2]"], ["[TS]", "[WP]"]]
    """
    leaf_dirs: set[Path] = set()
    for stem in pdf_stems:
        for pdf_path in base_dir.rglob(f"{stem}.pdf"):
            leaf_dirs.add(pdf_path.parent)

    if not leaf_dirs:
        return []

    # Relative path components for every leaf directory
    rel_parts = [p.relative_to(base_dir).parts for p in leaf_dirs]
    depth = max(len(p) for p in rel_parts)

    levels = []
    for i in range(depth):
        # dict.fromkeys preserves insertion order while deduplicating;
        # sorting rel_parts first gives a deterministic alphabetical order.
        names = dict.fromkeys(
            parts[i] for parts in sorted(rel_parts) if i < len(parts)
        )
        levels.append(list(names))

    return levels


def _parse_config() -> Config:
    args = _build_parser().parse_args()

    if args.level:
        levels = [item.split(",") for item in args.level]
    else:
        levels = discover_levels(
            args.base_dir,
            [stem for stem, _label, _row in PDF_SECTIONS],
        )
        if not levels:
            levels = LEVELS  # last-resort fallback to the hardcoded constant

    sections = PDF_SECTIONS
    if args.section:
        sections = []
        for item in args.section:
            parts = item.split(":")
            if len(parts) != 3:
                raise SystemExit(f"--section must be 'stem:label:row', got: {item!r}")
            stem, label, row = parts
            sections.append((stem, label, int(row)))

    return Config(
        base_dir=args.base_dir,
        output_dir=args.output_dir,
        excel_output=args.excel_output,
        levels=levels,
        pdf_sections=sections,
        pdf_dpi=args.dpi,
        pdf_page=args.page,
        img_max_width=args.img_max_width,
        header_span_col=args.header_span_col,
        num_columns=args.num_columns,
        column_width=args.column_width,
        header_bg_color=args.header_bg_color,
        header_font_bold=HEADER_FONT_BOLD,
        header_font_size=args.header_font_size,
        sheet_name_replacements=SHEET_NAME_REPLACEMENTS,
    )


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sanitise_sheet_name(name: str, replacements: dict) -> str:
    for old, new in replacements.items():
        name = name.replace(old, new)
    return name


def _apply_header_row(ws, row: int, label: str, cfg: Config):
    """Merge, style, and label a single header row."""
    last_col_idx = column_index_from_string(cfg.header_span_col)
    fill  = PatternFill(start_color=cfg.header_bg_color, end_color=cfg.header_bg_color, fill_type="solid")
    font  = Font(bold=cfg.header_font_bold, size=cfg.header_font_size)
    align = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"A{row}:{cfg.header_span_col}{row}")
    cell = ws.cell(row=row, column=1)
    cell.value, cell.fill, cell.font, cell.alignment = label, fill, font, align
    for col in range(1, last_col_idx + 1):
        ws.cell(row=row, column=col).fill = fill


def _insert_image(ws, png_path: str, anchor: str, max_width: int):
    """Load, scale if necessary, and insert a PNG at the given cell anchor."""
    img = XlImage(png_path)
    if img.width > max_width:
        ratio      = max_width / img.width
        img.width  = int(img.width  * ratio)
        img.height = int(img.height * ratio)
    ws.add_image(img, anchor)


# ── Core logic ────────────────────────────────────────────────────────────────

def convert_pdfs_to_png(cfg: Config) -> dict:
    """Rasterise configured PDFs for every folder combination."""
    cfg.output_dir.mkdir(exist_ok=True)
    png_map = {}

    for combo in product(*cfg.levels):
        folder = cfg.base_dir.joinpath(*combo)
        if not folder.exists():
            print(f"  Skipping {folder} (not found)")
            continue

        key = "".join(combo)
        png_map[key] = {}

        for pdf_stem, _label, _row in cfg.pdf_sections:
            pdf_path = folder / f"{pdf_stem}.pdf"
            if not pdf_path.exists():
                print(f"  Skipping {pdf_path} (not found)")
                continue

            png_filename = "_".join(combo) + f"_{pdf_stem}.png"
            png_path     = cfg.output_dir / png_filename

            print(f"  Converting {pdf_path.name} -> {png_filename}")
            pages = convert_from_path(str(pdf_path), dpi=cfg.pdf_dpi)
            pages[cfg.pdf_page].save(str(png_path), "PNG")
            png_map[key][pdf_stem] = str(png_path)

    return png_map


def create_excel(png_map: dict, cfg: Config):
    """Create the Excel workbook with one sheet per folder combination."""
    wb = Workbook()
    wb.remove(wb.active)

    for combo in product(*cfg.levels):
        key = "".join(combo)
        if key not in png_map:
            continue

        sheet_name = _sanitise_sheet_name(key, cfg.sheet_name_replacements)
        ws = wb.create_sheet(title=sheet_name)

        for pdf_stem, label, header_row in cfg.pdf_sections:
            _apply_header_row(ws, header_row, label, cfg)
            if pdf_stem in png_map[key]:
                _insert_image(ws, png_map[key][pdf_stem], f"A{header_row + 1}", cfg.img_max_width)

        for col in range(1, cfg.num_columns + 1):
            ws.column_dimensions[get_column_letter(col)].width = cfg.column_width

        print(f"  Created sheet: {sheet_name}")

    wb.save(str(cfg.excel_output))
    print(f"\nExcel saved to: {cfg.excel_output}")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    cfg = _parse_config()

    print("Step 1: Converting PDFs to PNG...")
    png_map = convert_pdfs_to_png(cfg)

    print("\nStep 2: Creating Excel workbook...")
    create_excel(png_map, cfg)

    print("\nDone!")


if __name__ == "__main__":
    main()
