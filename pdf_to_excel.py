"""
PDF -> Excel parser for the Italian financial ledger layout:
DATA / N.RIF. / N.Prot. / DESCRIZIONE / CONTO / DARE / AVERE

Built per the functional spec in `Prompts`:
  - pdfplumber only (extract_text() to locate the header, extract_words()
    for coordinate-based column assignment)
  - columns are decided by each word's x-position relative to the header's
    column x-coordinates, NOT by regex token order
  - amounts use a comma as a DECIMAL POINT (e.g. "61,00" -> 61.00), never
    split into two fields
  - DARE vs AVERE is decided purely by which column-x-range the amount
    word falls into
  - CONTO = everything (numeric identifiers + description words) sitting
    in the CONTO column's x-range
  - N.Prot. carries forward the last non-empty value when missing; DATA,
    N.RIF., DESCRIZIONE, CONTO are left blank when missing
  - dev scope: page 1 only for now (see `process_all_pages` flag), but
    structured so extending to every page is a one-line change

NOTE ON THE SPEC: the "Amount Rule" section says to "replace the comma
with 0", but its own worked example converts "61,00" -> "61.00" (comma
replaced with a decimal point). This implementation follows the worked
example. Flag this if that's not what you intended.
"""

import re
import sys
import traceback
from pathlib import Path

import pdfplumber
import pandas as pd


# -------------------- CONFIG --------------------

EXPECTED_HEADER_NORM = "DATANRIFNPROTDESCRIZIONECONTODAREAVERE"

# Internal column keys, in left-to-right order as they appear in the header
COLUMN_LABELS = ["DATA", "NRIF", "NPROT", "DESCRIZIONE", "CONTO", "DARE", "AVERE"]

# Output Excel columns, in required order
HEADERS = ["DATA", "N.RIF.", "N.Prot.", "DESCRIZIONE", "CONTO", "DARE", "AVERE"]

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
NRIF_RE = re.compile(r"^[A-Za-z]+$")
NPROT_RE = re.compile(r"^\d+$")

Y_TOLERANCE = 3    # px tolerance for grouping words into the same visual line
COL_TOLERANCE = 5  # px tolerance when binning a word's x0 into a column


# -------------------- HELPERS --------------------

def normalize_header(text):
    return re.sub(r"[^A-Z0-9]", "", text.upper())


def should_skip_row(line_text):
    n = normalize_header(line_text)
    return n.startswith("RIPORTI") or n.startswith("TOTALIPROGRESSIVI")


def italian_amount_to_float(raw):
    """Comma is a decimal point here, not a field separator.
    '61,00' -> 61.00 ; '1.234,56' -> 1234.56
    """
    if not raw:
        return None
    val = raw.strip().replace(".", "").replace(",", ".")
    try:
        return float(val)
    except ValueError:
        return None


# -------------------- LINE GROUPING --------------------

def group_words_into_lines(words, y_tolerance=Y_TOLERANCE):
    """Cluster pdfplumber words into visual lines using their 'top' coordinate."""
    lines = []
    buffer = []
    ref_top = None

    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if ref_top is None or abs(w["top"] - ref_top) <= y_tolerance:
            buffer.append(w)
            ref_top = w["top"] if ref_top is None else ref_top
        else:
            lines.append(sorted(buffer, key=lambda w: w["x0"]))
            buffer = [w]
            ref_top = w["top"]

    if buffer:
        lines.append(sorted(buffer, key=lambda w: w["x0"]))

    return lines


# -------------------- HEADER / COLUMN DETECTION --------------------

def detect_header_columns(line_words):
    """
    Try to match this line's words against the expected header, consuming
    characters label-by-label. Handles labels split across multiple words
    by the PDF extractor (e.g. 'D A R E' as four separate word objects).
    Returns {label: x0_start} on a full match, else None.
    """
    col_starts = {}
    label_idx = 0
    remaining = COLUMN_LABELS[0]
    current_start_x = None

    for w in line_words:
        norm = normalize_header(w["text"])
        if not norm:
            continue
        if current_start_x is None:
            current_start_x = w["x0"]

        while norm:
            if label_idx >= len(COLUMN_LABELS):
                return None  # extra text beyond the 7 expected labels

            take = min(len(norm), len(remaining))
            chunk, norm = norm[:take], norm[take:]

            if not remaining.startswith(chunk):
                return None  # doesn't match the expected header at all

            remaining = remaining[len(chunk):]

            if remaining == "":
                col_starts[COLUMN_LABELS[label_idx]] = current_start_x
                label_idx += 1
                current_start_x = None
                if label_idx < len(COLUMN_LABELS):
                    remaining = COLUMN_LABELS[label_idx]

    return col_starts if label_idx == len(COLUMN_LABELS) else None


def build_column_bins(col_starts):
    """Return (labels, x_starts) sorted left-to-right for word binning."""
    ordered = sorted(col_starts.items(), key=lambda kv: kv[1])
    labels = [label for label, _ in ordered]
    starts = [x for _, x in ordered]
    return labels, starts


def assign_column(x0, labels, starts, tolerance=COL_TOLERANCE):
    idx = 0
    for i, start in enumerate(starts):
        if x0 >= start - tolerance:
            idx = i
    return labels[idx]


# -------------------- TRANSACTION ROW PARSER --------------------

def parse_transaction_line(line_words, labels, starts, last_nprot):
    """Bin every word by x-position into its column, then validate each field."""
    bins = {label: [] for label in labels}

    for w in line_words:
        col = assign_column(w["x0"], labels, starts)
        bins[col].append(w["text"])

    raw = {label: " ".join(bins[label]).strip() for label in labels}

    row = {
        "DATA": "",
        "N.RIF.": "",
        "N.Prot.": last_nprot,
        "DESCRIZIONE": raw.get("DESCRIZIONE", ""),
        "CONTO": raw.get("CONTO", ""),
        "DARE": "",
        "AVERE": "",
    }

    # DATA
    data_val = raw.get("DATA", "")
    if DATE_RE.match(data_val):
        row["DATA"] = data_val

    # N.RIF.
    nrif_val = raw.get("NRIF", "")
    if NRIF_RE.match(nrif_val):
        row["N.RIF."] = nrif_val

    # N.Prot. -> numeric, else carry forward the last known value
    nprot_val = raw.get("NPROT", "")
    if NPROT_RE.match(nprot_val):
        row["N.Prot."] = nprot_val
        last_nprot = nprot_val
    else:
        row["N.Prot."] = last_nprot

    # DARE / AVERE -> column position decides which field; comma is a decimal point
    dare_amount = italian_amount_to_float(raw.get("DARE", ""))
    avere_amount = italian_amount_to_float(raw.get("AVERE", ""))

    row["DARE"] = round(dare_amount, 2) if dare_amount is not None else ""
    row["AVERE"] = round(avere_amount, 2) if avere_amount is not None else ""

    return row, last_nprot


# -------------------- PAGE PROCESSING --------------------

def process_page(page, last_nprot, labels, starts):
    """
    Parse one page. `labels`/`starts` carry the column layout forward from
    a previous page (None on the very first page, until the header is found).
    Returns (records, last_nprot, labels, starts).
    """
    words = page.extract_words()
    lines = group_words_into_lines(words)

    records = []
    inside_transactions = labels is not None

    for line_words in lines:
        if not line_words:
            continue

        if not inside_transactions:
            col_starts = detect_header_columns(line_words)
            if col_starts:
                labels, starts = build_column_bins(col_starts)
                inside_transactions = True
            continue

        # Header repeating on this page (e.g. per-page banner) -> skip it
        if detect_header_columns(line_words):
            continue

        line_text = " ".join(w["text"] for w in line_words)
        if should_skip_row(line_text):
            continue

        row, last_nprot = parse_transaction_line(line_words, labels, starts, last_nprot)

        if any(row[col] != "" for col in HEADERS):
            records.append(row)

    return records, last_nprot, labels, starts


# -------------------- MAIN DRIVER --------------------

def convert(pdf_path, output_path, process_all_pages=False):
    """
    process_all_pages=False matches the current dev scope (page 1 only).
    Flip to True once the page-1 output is validated.
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("Input file must be a PDF")

    records = []
    last_nprot = ""
    labels, starts = None, None

    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages if process_all_pages else pdf.pages[:1]
        print(f"\nProcessing {len(pages)} page(s) of {len(pdf.pages)} total")

        for page_num, page in enumerate(pages, start=1):
            try:
                page_records, last_nprot, labels, starts = process_page(
                    page, last_nprot, labels, starts
                )
                records.extend(page_records)
            except Exception:
                print(f"\nPAGE PROCESSING ERROR (page {page_num})")
                traceback.print_exc()

    if labels is None:
        print("\nWARNING: transaction header was never found — 0 rows extracted.")

    df = pd.DataFrame(records, columns=HEADERS)
    df.to_excel(output_path, index=False)

    print(f"\nTotal rows extracted: {len(df)}")
    print(f"Excel saved to: {output_path}")

    return df


# -------------------- ENTRY POINT --------------------

if __name__ == "__main__":
    try:
        if len(sys.argv) < 2:
            print("Usage: python giornale_to_excel.py input.pdf [output.xlsx]")
            sys.exit(1)

        pdf = sys.argv[1]
        output = sys.argv[2] if len(sys.argv) > 2 else str(Path(pdf).with_suffix(".xlsx"))

        convert(pdf, output, process_all_pages=False)

    except Exception:
        print("\nPROGRAM FAILED")
        traceback.print_exc()