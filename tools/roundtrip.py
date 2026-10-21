"""End-to-end check: PDF -> xlsx -> (edit one cell) -> PDF, with PNG renders.

    python tools/roundtrip.py samples/clean_statement_scan.pdf out_dir [--scan]

--scan first rasterises the PDF into an image-only PDF to exercise the OCR path.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import fitz  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

from backend.excel_writer import LAYOUT_SHEET, write_workbook  # noqa: E402
from backend.extract import extract_pages  # noqa: E402
from backend.pdf_writer import excel_to_pdf  # noqa: E402


def make_scan(pdf_bytes: bytes, dpi: int = 200) -> bytes:
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    for p in src:
        pix = p.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
        page = out.new_page(width=p.rect.width, height=p.rect.height)
        page.insert_image(page.rect, stream=pix.tobytes("jpeg"))
    return out.tobytes()


def render(pdf_bytes: bytes, path: Path, page: int = 0) -> None:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    doc[page].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(str(path))


def main() -> None:
    src = Path(sys.argv[1])
    out = Path(sys.argv[2])
    out.mkdir(parents=True, exist_ok=True)
    pdf_bytes = src.read_bytes()
    if "--scan" in sys.argv:
        pdf_bytes = make_scan(pdf_bytes)
        (out / "scan_input.pdf").write_bytes(pdf_bytes)

    pages = extract_pages(pdf_bytes)
    for p in pages:
        print(f"page {p.index}: {p.source}, {len(p.words)} words, {len(p.rules)} rules, {len(p.pictures)} pictures")

    xlsx = write_workbook(pages, pdf_bytes, src.name)
    (out / "statement.xlsx").write_bytes(xlsx)

    wb = load_workbook(io.BytesIO(xlsx))
    for name in wb.sheetnames:
        if name.startswith("Page"):
            ws = wb[name]
            print(f"\n== {name}: {ws.max_row} rows x {ws.max_column} cols, {len(ws.merged_cells.ranges)} merges")
            for row in ws.iter_rows():
                vals = [(c.coordinate, c.value) for c in row if c.value is not None]
                if vals:
                    print("  " + " | ".join(f"{k}={v!r}" for k, v in vals))

    # formulas: list them and check they evaluate to the statement's own totals
    from backend.formulas import evaluate, format_amount, parse_amount
    lay = wb[LAYOUT_SHEET]
    for row in lay.iter_rows(min_row=2, values_only=True):
        sheet, ref, text = row[1], row[2], str(row[-1])
        c = wb[sheet][ref]
        if isinstance(c.value, str) and c.value.startswith("="):
            val = evaluate(c.value, wb[sheet], wb[sheet])
            like = parse_amount(text)
            shown = format_amount(val, like) if (val is not None and like) else val
            flag = "OK" if shown == text else f"MISMATCH (statement says {text!r})"
            print(f"formula {sheet}!{ref}: {c.value} -> {shown}  {flag}")

    # unchanged round-trip must be byte-identical
    same_pdf, n = excel_to_pdf(xlsx)
    assert n == 0 and same_pdf == pdf_bytes, f"unchanged workbook must return the original PDF ({n} changes)"
    print("\nunchanged round-trip: identical bytes OK")

    # edit: change the first numeric-looking cell and the first long text cell on Page 1
    lay = wb[LAYOUT_SHEET]
    edits = 0
    for row in list(lay.iter_rows(min_row=2, values_only=True)):
        sheet, ref, text = row[1], row[2], str(row[-1])
        if edits == 0 and text.replace(",", "").replace(".", "").isdigit():
            wb[sheet][ref].value = "9,999.99"; edits += 1; print(f"edit {sheet}!{ref}: {text!r} -> '9,999.99'")
        elif edits == 1 and len(text) > 12 and text.isupper():
            wb[sheet][ref].value = "EDITED NAME HERE"; edits += 1; print(f"edit {sheet}!{ref}: {text!r} -> 'EDITED NAME HERE'")
        if edits == 2:
            break
    buf = io.BytesIO()
    wb.save(buf)
    (out / "statement_edited.xlsx").write_bytes(buf.getvalue())
    new_pdf, n = excel_to_pdf(buf.getvalue())
    print(f"edited round-trip: {n} change(s) applied")
    (out / "statement_edited.pdf").write_bytes(new_pdf)
    render(pdf_bytes, out / "original_p1.png")
    render(new_pdf, out / "edited_p1.png")
    print("wrote", out)


if __name__ == "__main__":
    main()
