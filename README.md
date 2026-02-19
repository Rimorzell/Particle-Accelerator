# Particle-Accelerator

Bulk extractor for SKU / Order Code + specifications from large nested PDF collections.

## What this does
- Recursively scans a root folder for all `.pdf` files.
- Uses multiple text extraction engines (tries `pypdf`, then `pdfplumber`, then `PyMuPDF`) for better robustness.
- Finds SKU/order codes from common labels (`SKU`, `Order Code`, `Model No`, etc.) and fallback patterns.
- Collects specification key/value pairs near each SKU and from specification sections.
- Exports an Excel workbook with:
  - `sku_specs` sheet: one row per extracted SKU/spec block.
  - `summary` sheet: PDF count, extracted row count, and unique SKU count.

## Install
```bash
pip install pandas openpyxl pypdf pdfplumber pymupdf
```

> You can still run with only `pandas openpyxl pypdf`, but `pdfplumber` and `pymupdf` improve extraction coverage.

## Usage
From the repository root:

```bash
python3 pdf_sku_spec_extractor.py /path/to/your/pdf/root -o sku_spec_extraction.xlsx
```

If omitted, input defaults to current folder:

```bash
python3 pdf_sku_spec_extractor.py -o sku_spec_extraction.xlsx
```

## Notes for messy PDFs
- If a PDF is image-only, you may need OCR before extraction.
- For highly inconsistent files, results are still best-effort; review the Excel and spot-check by PDF path/page.
- You can run this multiple times and compare outputs as you refine regex patterns for your catalog naming style.
