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
NRIF_RE = re.compile(r"^[A-Za-z]$")  # a single alphabetic character only (A, B, C...)
NPROT_RE = re.compile(r"^\d+$")

Y_TOLERANCE = 3    # px tolerance for grouping words into the same visual line


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
    """Return (labels, boundaries) sorted left-to-right for word binning.
    `boundaries` are the midpoints between each pair of adjacent column
    starts — a word is assigned to a column based on which side of these
    midpoints it falls on, rather than the raw column start. This is far
    more forgiving of real-world text not being perfectly left-aligned to
    the header label above it (e.g. a DESCRIZIONE word sitting slightly
    left of DESCRIZIONE's own header start would otherwise get misassigned
    to N.Prot. under a naive "nearest preceding start" rule).
    """
    ordered = sorted(col_starts.items(), key=lambda kv: kv[1])
    labels = [label for label, _ in ordered]
    starts = [x for _, x in ordered]
    boundaries = [(starts[i] + starts[i + 1]) / 2 for i in range(len(starts) - 1)]
    return labels, boundaries


def assign_column(x0, labels, boundaries):
    idx = 0
    for i, boundary in enumerate(boundaries):
        if x0 >= boundary:
            idx = i + 1
        else:
            break
    return labels[idx]


# -------------------- TRANSACTION ROW PARSER (RIGHT TO LEFT) --------------------

def parse_transaction_line(line_words, labels, boundaries, last_nprot):
    """
    Parse a transaction row RIGHT TO LEFT, popping tokens off the end of the
    line one column at a time, in this order:
        1. AMOUNT (DARE or AVERE, whichever column the last word sits in)
        2. CONTO       - keep popping while still inside the CONTO column
        3. DESCRIZIONE / N.Prot. / N.RIF. / DATA zone - keep popping and
           route each word by its column; whatever isn't N.Prot./N.RIF./DATA
           falls into DESCRIZIONE

    Column boundaries come from the header's x-coordinates (see
    detect_header_columns). This still anchors on position rather than pure
    token content, because a DESCRIZIONE word can itself look numeric or
    date-like (e.g. this spec's own example description "100567 Ts.bsst.n.
    6724/8-2025 yep") and would otherwise be misread as N.Prot. or DATA.
    """
    remaining = sorted(line_words, key=lambda w: w["x0"])

    def col_at(x0):
        return assign_column(x0, labels, boundaries)

    row = {
        "DATA": "",
        "N.RIF.": "",
        "N.Prot.": last_nprot,
        "DESCRIZIONE": "",
        "CONTO": "",
        "DARE": "",
        "AVERE": "",
    }

    # STEP 1: AMOUNT at the extreme right
    if remaining:
        last = remaining[-1]
        col = col_at(last["x0"])
        if col in ("DARE", "AVERE"):
            amount = italian_amount_to_float(last["text"])
            if amount is not None:
                row[col] = round(amount, 2)
                remaining.pop()

    # STEP 2: CONTO block - keep collecting while still in the CONTO column.
    # NOTE: the amount was already popped in STEP 1, so if a word's x0 still
    # falls in the DARE/AVERE zone at this point, it's virtually certain to
    # be CONTO text overflowing rightward (e.g. a long account description),
    # not a second amount -- so DARE/AVERE also counts as "still CONTO" here.
    conto_words = []
    while remaining and col_at(remaining[-1]["x0"]) in ("CONTO", "DARE", "AVERE"):
        conto_words.insert(0, remaining.pop()["text"])
    row["CONTO"] = " ".join(conto_words)

    # STEP 3: everything left of CONTO - keep popping right-to-left, routing
    # each word by column; N.Prot./N.RIF./DATA are pulled out, everything
    # else accumulates into DESCRIZIONE
    desc_words, nprot_words, nrif_words, data_words = [], [], [], []

    while remaining:
        w = remaining.pop()
        col = col_at(w["x0"])
        if col == "NPROT":
            nprot_words.insert(0, w["text"])
        elif col == "NRIF":
            nrif_words.insert(0, w["text"])
        elif col == "DATA":
            data_words.insert(0, w["text"])
        else:
            desc_words.insert(0, w["text"])

    row["DESCRIZIONE"] = " ".join(desc_words).strip()

    # STEP 4: N.Prot. (numeric) - carry forward previous value if missing
    nprot_val = " ".join(nprot_words).strip()
    if NPROT_RE.match(nprot_val):
        row["N.Prot."] = nprot_val
        last_nprot = nprot_val
    else:
        row["N.Prot."] = last_nprot

    # STEP 5: N.RIF. (single alphabetic character)
    nrif_val = " ".join(nrif_words).strip()
    if NRIF_RE.match(nrif_val):
        row["N.RIF."] = nrif_val

    # STEP 6: DATA (dd/mm/yyyy)
    data_val = " ".join(data_words).strip()
    if DATE_RE.match(data_val):
        row["DATA"] = data_val

    return row, last_nprot


# -------------------- PAGE PROCESSING --------------------

def process_page(page, last_nprot, labels, boundaries):
    """
    Parse one page. `labels`/`boundaries` carry the column layout forward from
    a previous page (None on the very first page, until the header is found).
    Returns (records, last_nprot, labels, boundaries).
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
                labels, boundaries = build_column_bins(col_starts)
                inside_transactions = True
            continue

        # Header repeating on this page (e.g. per-page banner) -> skip it
        if detect_header_columns(line_words):
            continue

        line_text = " ".join(w["text"] for w in line_words)
        if should_skip_row(line_text):
            continue

        row, last_nprot = parse_transaction_line(line_words, labels, boundaries, last_nprot)

        if any(row[col] != "" for col in HEADERS):
            records.append(row)

    return records, last_nprot, labels, boundaries


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
    labels, boundaries = None, None

    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages if process_all_pages else pdf.pages[:1]
        print(f"\nProcessing {len(pages)} page(s) of {len(pdf.pages)} total")

        for page_num, page in enumerate(pages, start=1):
            try:
                page_records, last_nprot, labels, boundaries = process_page(
                    page, last_nprot, labels, boundaries
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
