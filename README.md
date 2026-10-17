# Statement PDF ⇄ Excel

Convert a bank-statement PDF (scanned or digital) into an Excel workbook that
mirrors the page layout, edit values in Excel, and export a PDF that is
identical to the original except for the edited values.

## How it works

1. **Extract** – each page is rasterised at 200 dpi. If the PDF has a text
   layer, words and font sizes come straight from it; otherwise Tesseract OCR
   provides words with bounding boxes. Horizontal rules and non-text regions
   (logos) are detected from the raster.
2. **Layout** – words are grouped into text runs, runs into rows (y-bands)
   and columns (x-boundaries). Column widths and row heights are proportional
   to the page, so the sheet lines up with the statement.
3. **Excel** – one visible sheet per page (`Page 1`, …) with fonts, alignment,
   merged cells, borders and the logo image. Two hidden sheets hold the
   coordinates of every text run (`_layout`) and the original PDF (`_source`).
4. **Export** – the original PDF is reused untouched. For every cell whose
   text differs from the original, the old glyphs are painted over with the
   sampled background colour and the new text is drawn at the same baseline,
   size and alignment. No edits ⇒ output is byte-identical to the input.

## Setup (Windows)

```powershell
winget install --id UB-Mannheim.TesseractOCR -e   # OCR engine (scanned PDFs)
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\run.ps1                                          # http://127.0.0.1:8000
```

## API

| Method | Path                 | Body (multipart) | Returns |
|--------|----------------------|------------------|---------|
| POST   | `/api/pdf-to-excel`  | `file` = .pdf    | .xlsx   |
| POST   | `/api/excel-to-pdf`  | `file` = .xlsx   | .pdf    |
| GET    | `/api/health`        |                  | `{ok, tesseract}` |

## Project layout

```
backend/
  app.py           FastAPI routes
  extract.py       rasterise, text layer / OCR, rules & logos
  layout.py        words -> segments -> rows/columns grid
  excel_writer.py  grid -> .xlsx (+ hidden layout/source sheets)
  pdf_writer.py    edited .xlsx -> PDF overlay
  static/index.html
samples/           reference statements
tools/roundtrip.py local end-to-end test
```

## Formulas (statement arithmetic)

When the table header contains WITHDRAWALS / DEPOSITS / BALANCE (or
DEBIT / CREDIT), `backend/formulas.py` makes the sheet live:

* amount cells become **numbers** with a number format that reproduces the
  statement text exactly (`#,##0.00`, `#,##0.00"Cr"`, …);
* the **Page Total** row gets `=SUM(first:last)` per amount column and the
  page-total balance points at the last balance;
* every **balance** is `=previous balance - withdrawal + deposit`; the first
  one uses the (unprinted) opening balance as a literal. A balance formula is
  only written where the statement's own arithmetic confirms it.

Edit any amount and the totals/balances follow; on export every cell whose
value differs from the statement is redrawn. Workbooks that were never
recalculated by Excel are evaluated by the built-in mini evaluator.

## Notes / limits

* Cells are stored as **text** (`@` number format) so Excel never reformats
  `6,186.00` or `74,559.86Cr`. If you retype a value as a number, the export
  re-applies the original thousands/decimals formatting.
* Replacement text is drawn in Helvetica (regular/bold), the closest built-in
  match to the Arial-style fonts used on most statements.
* Do not rename, delete or reorder the hidden `_layout` / `_source` sheets.
