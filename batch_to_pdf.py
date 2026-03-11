#!/usr/bin/env python3
"""
batch_to_pdf.py

Convert CSV and TXT files to PDFs. Options:
 - choose input folder via CLI or a pop-up folder chooser
 - convert each file to its own PDF, or combine all into a single PDF
 - each page shows the source filename in the header
"""

import argparse
import os
from pathlib import Path
import sys
import pandas as pd

# reportlab imports
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
    Preformatted,
    Flowable,
)
from reportlab.lib.styles import ParagraphStyle

# tkinter for folder selection pop-up
try:
    import tkinter as tk
    from tkinter import filedialog
except Exception:
    tk = None
    filedialog = None


PAGE_SIZE = A4
PAGE_WIDTH, PAGE_HEIGHT = PAGE_SIZE
LEFT_MARGIN = RIGHT_MARGIN = top_margin = bottom_margin = 0.5 * inch
CONTENT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN

styles = getSampleStyleSheet()
filename_style = ParagraphStyle(
    "filename_style",
    parent=styles["Heading2"],
    alignment=TA_CENTER,
    spaceAfter=10,
)

text_style = ParagraphStyle(
    "text_style",
    parent=styles["Normal"],
    fontName="Courier",
    fontSize=9,
    leading=11,
    alignment=TA_LEFT,
)

# Flowable that sets the filename into the canvas so the onPage callback can draw it
class SetCurrentFilename(Flowable):
    def __init__(self, filename):
        super().__init__()
        self.filename = filename

    def wrap(self, availWidth, availHeight):
        # zero height; will not consume vertical space
        return (0, 0)

    def draw(self):
        # store on the canvas object for onPage to read
        self.canv._current_filename = self.filename


def header_footer(canvas, doc):
    # Draw filename (set by SetCurrentFilename) and page number in header/footer
    canvas.saveState()
    fn = getattr(canvas, "_current_filename", "")
    # header: filename centered
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawCentredString(PAGE_WIDTH / 2.0, PAGE_HEIGHT - 0.35 * inch, fn)
    # footer: page number right
    canvas.setFont("Helvetica", 8)
    page_text = f"Page {doc.page}"
    canvas.drawRightString(PAGE_WIDTH - RIGHT_MARGIN, 0.35 * inch, page_text)
    canvas.restoreState()


def csv_to_table_flowables(csv_path):
    """Return list of flowables for a csv file (includes SetCurrentFilename at start)."""
    try:
        df = pd.read_csv(csv_path, dtype=str, keep_default_na=False, na_filter=False)
    except Exception as e:
        # if read_csv fails, treat as text file fallback
        return txt_to_text_flowables(csv_path, fallback_reason=f"CSV read error: {e}")

    data = [list(df.columns)]
    # ensure all cells are strings and safe
    for _, row in df.iterrows():
        data.append([str(x) for x in row.tolist()])

    # compute col widths (distribute across available width)
    ncols = max(1, len(data[0]))
    min_col_width = 50  # px
    col_width = max(min_col_width, CONTENT_WIDTH / ncols)

    col_widths = [col_width] * ncols

    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle(
        [
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#dddddd")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f3f3f3")),
            ("FONT", (0, 0), (-1, -1), "Helvetica", 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]
    )
    tbl.setStyle(tbl_style)

    return [SetCurrentFilename(os.path.basename(csv_path)), Spacer(1, 6), tbl, PageBreak()]


def txt_to_text_flowables(txt_path, fallback_reason=None):
    """Return flowables for a txt file (preformatted)."""
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception:
        with open(txt_path, "r", encoding="latin-1", errors="ignore") as f:
            content = f.read()

    if fallback_reason:
        content = f"(Note: {fallback_reason})\n\n" + content

    # Preformatted to preserve whitespace and wrapping
    pre = Preformatted(content, text_style)
    return [SetCurrentFilename(os.path.basename(txt_path)), Spacer(1, 6), pre, PageBreak()]


def build_pdf_for_files(file_paths, output_path, combined=False, landscape_mode=False):
    """Create PDF(s) from list of file_paths.

    - If combined=True: create one PDF at output_path containing all files.
    - If combined=False: create separate PDFs per file in same dir as output_path or in file dir.
    """
    if combined:
        doc = SimpleDocTemplate(
            output_path,
            pagesize=landscape(PAGE_SIZE) if landscape_mode else PAGE_SIZE,
            leftMargin=LEFT_MARGIN,
            rightMargin=RIGHT_MARGIN,
            topMargin=top_margin,
            bottomMargin=bottom_margin,
        )

        story = []
        for p in file_paths:
            if p.suffix.lower() == ".csv":
                story.extend(csv_to_table_flowables(p))
            elif p.suffix.lower() == ".txt":
                story.extend(txt_to_text_flowables(p))
            else:
                # skip unknown
                continue

        # If last flowable is PageBreak, remove it to avoid blank last page
        if story and isinstance(story[-1], PageBreak):
            story = story[:-1]

        doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
        print(f"Combined PDF written to: {output_path}")
    else:
        # Make separate PDFs for each file
        for p in file_paths:
            out_name = p.with_suffix(".pdf")
            doc = SimpleDocTemplate(
                out_name,
                pagesize=landscape(PAGE_SIZE) if landscape_mode else PAGE_SIZE,
                leftMargin=LEFT_MARGIN,
                rightMargin=RIGHT_MARGIN,
                topMargin=top_margin,
                bottomMargin=bottom_margin,
            )
            if p.suffix.lower() == ".csv":
                story = csv_to_table_flowables(p)
            elif p.suffix.lower() == ".txt":
                story = txt_to_text_flowables(p)
            else:
                continue

            if story and isinstance(story[-1], PageBreak):
                story = story[:-1]
            doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
            print(f"Wrote: {out_name}")


def find_files_in_folder(folder: Path, patterns=("*.csv", "*.txt"), recursive=False):
    files = []
    for pat in patterns:
        if recursive:
            files.extend(folder.rglob(pat))
        else:
            files.extend(folder.glob(pat))
    # sort for stable order
    files = sorted([f for f in files if f.is_file()])
    return files


def choose_folder_with_dialog():
    if filedialog is None:
        print("tkinter not available; please supply a --path argument.", file=sys.stderr)
        sys.exit(1)
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Select folder containing CSV / TXT files")
    root.update()
    root.destroy()
    return folder


def main():
    parser = argparse.ArgumentParser(
        description="Convert CSV and TXT files in a folder to PDFs. Option to combine files."
    )
    parser.add_argument(
        "--path",
        help="Path to folder containing files. If omitted, a folder-picker popup will appear (requires tkinter).",
        default=None,
    )
    parser.add_argument(
        "--files",
        nargs="+",
        help="Optional explicit list of files to convert (overrides --path). Provide full paths or relative.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search subfolders recursively when using --path.",
    )
    parser.add_argument(
        "--combine",
        action="store_true",
        help="Combine all selected files into a single PDF. If not set, creates one PDF per file.",
    )
    parser.add_argument(
        "--output",
        help="Output file path when --combine is set (default: combined.pdf in current dir).",
        default="combined.pdf",
    )
    parser.add_argument(
        "--landscape",
        action="store_true",
        help="Generate PDFs in landscape orientation.",
    )

    args = parser.parse_args()

    if args.files:
        p_list = [Path(f) for f in args.files]
        files = [p for p in p_list if p.exists() and p.suffix.lower() in (".csv", ".txt")]
        if not files:
            print("No valid CSV or TXT files found in --files list.", file=sys.stderr)
            sys.exit(1)
    else:
        if args.path:
            folder = Path(args.path)
            if not folder.exists():
                print(f"Path does not exist: {folder}", file=sys.stderr)
                sys.exit(1)
        else:
            chosen = choose_folder_with_dialog()
            if not chosen:
                print("No folder selected. Exiting.", file=sys.stderr)
                sys.exit(0)
            folder = Path(chosen)

        files = find_files_in_folder(folder, recursive=args.recursive)
        if not files:
            print(f"No .csv or .txt files found in: {folder}", file=sys.stderr)
            sys.exit(1)

    if args.combine:
        out = Path(args.output)
        # ensure extension
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")
        build_pdf_for_files(files, str(out), combined=True, landscape_mode=args.landscape)
    else:
        build_pdf_for_files(files, None, combined=False, landscape_mode=args.landscape)


if __name__ == "__main__":
    main()
