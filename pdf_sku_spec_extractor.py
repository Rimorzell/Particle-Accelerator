#!/usr/bin/env python3
"""
Bulk PDF SKU / Order Code + specification extractor.

Designed for mixed-quality catalogs with inconsistent layouts.

Strategy:
1) Walk recursively and find every PDF.
2) Extract page text with multiple engines (PyPDF2 -> pdfplumber -> PyMuPDF).
3) Detect SKU/Order Code candidates from labels and stand-alone patterns.
4) Collect specification-like key/value lines from nearby and global context.
5) Export a normalized Excel workbook.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

SKU_LABEL_RE = re.compile(
    r"\b(?:SKU|Order\s*Code|Ordering\s*Code|Model\s*(?:No\.?|Number)|Part\s*(?:No\.?|Number)|Item\s*Code)\b\s*[:#\-]?\s*([A-Z0-9][A-Z0-9._\-/]{2,})",
    re.IGNORECASE,
)

SKU_STANDALONE_RE = re.compile(
    r"\b([A-Z]{2,}[A-Z0-9]*[-_/][A-Z0-9][A-Z0-9._\-/]{1,}|[A-Z]{2,}\d{2,}[A-Z0-9._\-/]*)\b"
)

SPEC_SECTION_RE = re.compile(r"\b(?:Specifications?|Technical\s*Data|Specs?)\b", re.IGNORECASE)
KEY_VALUE_RE = re.compile(
    r"^\s*([A-Za-z][A-Za-z0-9 /().%+\-]{1,40})\s*[:\-]\s*(.{1,220})\s*$"
)

NOISE_SPEC_KEYS = {
    "page",
    "copyright",
    "www",
    "http",
    "email",
    "fax",
    "phone",
}


@dataclass
class PageText:
    pdf_path: str
    page_num: int
    text: str


@dataclass
class ExtractionRow:
    pdf_path: str
    page_num: int
    sku: str
    specs: str
    source: str


def discover_pdfs(root: Path) -> List[Path]:
    return sorted({p for p in root.rglob("*.pdf") if p.is_file()} | {p for p in root.rglob("*.PDF") if p.is_file()})


def _extract_with_pypdf(pdf_path: Path) -> List[PageText]:
    pages: List[PageText] = []
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(pdf_path))
        for idx, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            pages.append(PageText(str(pdf_path), idx, text))
    except Exception:
        return []
    return pages


def _extract_with_pdfplumber(pdf_path: Path) -> List[PageText]:
    pages: List[PageText] = []
    try:
        import pdfplumber

        with pdfplumber.open(str(pdf_path)) as pdf:
            for idx, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                pages.append(PageText(str(pdf_path), idx, text))
    except Exception:
        return []
    return pages


def _extract_with_pymupdf(pdf_path: Path) -> List[PageText]:
    pages: List[PageText] = []
    try:
        import fitz  # PyMuPDF

        doc = fitz.open(str(pdf_path))
        try:
            for idx, page in enumerate(doc, start=1):
                text = page.get_text("text") or ""
                pages.append(PageText(str(pdf_path), idx, text))
        finally:
            doc.close()
    except Exception:
        return []
    return pages


def extract_pdf_text(pdf_path: Path) -> List[PageText]:
    for extractor in (_extract_with_pypdf, _extract_with_pdfplumber, _extract_with_pymupdf):
        pages = extractor(pdf_path)
        if pages and any(p.text.strip() for p in pages):
            return pages
    return []


def clean_sku(raw: str) -> str:
    sku = raw.strip().strip(".,;:)]}")
    sku = re.sub(r"\s+", "", sku)
    return sku.upper()


def find_skus(lines: Sequence[str]) -> List[Tuple[str, int, str]]:
    found: List[Tuple[str, int, str]] = []
    seen = set()

    for i, line in enumerate(lines):
        for m in SKU_LABEL_RE.finditer(line):
            sku = clean_sku(m.group(1))
            if len(sku) >= 4 and sku not in seen:
                seen.add(sku)
                found.append((sku, i, "label"))

    if not found:
        for i, line in enumerate(lines):
            if len(line) > 220:
                continue
            for m in SKU_STANDALONE_RE.finditer(line):
                sku = clean_sku(m.group(1))
                if len(sku) >= 5 and sku not in seen:
                    seen.add(sku)
                    found.append((sku, i, "pattern"))

    return found


def looks_like_spec_key(key: str) -> bool:
    k = key.strip().lower()
    if any(noise in k for noise in NOISE_SPEC_KEYS):
        return False
    if len(k) < 2:
        return False
    return True


def collect_spec_lines(lines: Sequence[str], start_idx: int, window: int = 40) -> Dict[str, str]:
    start = max(0, start_idx - 6)
    end = min(len(lines), start_idx + window)
    specs: Dict[str, str] = {}

    in_spec_section = False
    for idx in range(start, end):
        line = lines[idx].strip()
        if not line:
            continue

        if SPEC_SECTION_RE.search(line):
            in_spec_section = True
            continue

        m = KEY_VALUE_RE.match(line)
        if m:
            key = re.sub(r"\s+", " ", m.group(1).strip())
            val = re.sub(r"\s+", " ", m.group(2).strip())
            if looks_like_spec_key(key) and val:
                specs[key] = val
            continue

        if in_spec_section and 3 <= len(line) <= 120:
            parts = re.split(r"\s{2,}|\t", line)
            if len(parts) == 2 and all(parts):
                key, val = parts
                key = key.strip()
                val = val.strip()
                if looks_like_spec_key(key):
                    specs[key] = val

    return specs


def summarize_specs(specs: Dict[str, str], max_items: int = 30) -> str:
    if not specs:
        return ""
    items = list(specs.items())[:max_items]
    return " | ".join(f"{k}: {v}" for k, v in items)


def parse_page(page_text: PageText) -> List[ExtractionRow]:
    lines = [ln.strip() for ln in page_text.text.splitlines()]
    lines = [ln for ln in lines if ln]
    if not lines:
        return []

    sku_hits = find_skus(lines)
    rows: List[ExtractionRow] = []

    if not sku_hits:
        specs = collect_spec_lines(lines, start_idx=0, window=min(len(lines), 120))
        if specs:
            rows.append(
                ExtractionRow(
                    pdf_path=page_text.pdf_path,
                    page_num=page_text.page_num,
                    sku="",
                    specs=summarize_specs(specs),
                    source="spec_only",
                )
            )
        return rows

    for sku, line_idx, source in sku_hits:
        specs = collect_spec_lines(lines, start_idx=line_idx)
        rows.append(
            ExtractionRow(
                pdf_path=page_text.pdf_path,
                page_num=page_text.page_num,
                sku=sku,
                specs=summarize_specs(specs),
                source=source,
            )
        )

    return rows


def dedupe_rows(rows: Iterable[ExtractionRow]) -> List[ExtractionRow]:
    out: List[ExtractionRow] = []
    seen = set()
    for row in rows:
        key = (row.pdf_path, row.sku, row.specs)
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def run(root: Path, output: Path) -> Tuple[int, int]:
    try:
        import pandas as pd
    except Exception as exc:  # pragma: no cover
        raise SystemExit(
            "Missing dependency: pandas. Install with `pip install pandas openpyxl pypdf`."
        ) from exc

    pdfs = discover_pdfs(root)
    all_rows: List[ExtractionRow] = []

    for i, pdf in enumerate(pdfs, start=1):
        print(f"[{i}/{len(pdfs)}] Parsing {pdf}")
        pages = extract_pdf_text(pdf)
        for page in pages:
            all_rows.extend(parse_page(page))

    all_rows = dedupe_rows(all_rows)

    data = [
        {
            "pdf_path": r.pdf_path,
            "page_num": r.page_num,
            "sku_or_order_code": r.sku,
            "specifications": r.specs,
            "extraction_source": r.source,
        }
        for r in all_rows
    ]

    df = pd.DataFrame(data)
    if not df.empty:
        df = df.sort_values(["pdf_path", "page_num", "sku_or_order_code"]).reset_index(drop=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="sku_specs")

        summary = pd.DataFrame(
            {
                "metric": ["pdf_count", "rows_extracted", "unique_skus"],
                "value": [
                    len(pdfs),
                    len(df),
                    int(df["sku_or_order_code"].replace("", pd.NA).dropna().nunique())
                    if not df.empty
                    else 0,
                ],
            }
        )
        summary.to_excel(writer, index=False, sheet_name="summary")

    return len(pdfs), len(df)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract SKU/Order Code and specifications from nested PDF folders into Excel."
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=".",
        help="Root directory to scan recursively for PDFs (default: current directory)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="sku_spec_extraction.xlsx",
        help="Output Excel path",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    root = Path(args.input_dir).resolve()
    output = Path(args.output).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Input directory does not exist or is not a directory: {root}")

    pdf_count, row_count = run(root, output)
    print(f"Done. PDFs scanned: {pdf_count}. Rows extracted: {row_count}. Output: {output}")


if __name__ == "__main__":
    main()
