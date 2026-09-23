#!/usr/bin/env python3
"""
copy_docs_by_number.py

Reads an Excel file with a "Document Number" column and a "Status" column.
For each document number, finds the matching PDF in a source folder
(matching the number even if it's embedded in a filename with a prefix
like "Sid-", "Cx-", extra letters, etc.), copies it to a destination
folder, and updates the Status column to "Copied".

Run it and just answer the prompts - no need to edit this file.
"""

import os
import re
import shutil
import sys

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required. Install it with: pip install openpyxl")


def ask_path(prompt, must_exist=True, is_dir=None):
    while True:
        raw = input(prompt).strip().strip('"').strip("'")
        if not raw:
            print("  Please enter a path.")
            continue
        path = os.path.expanduser(raw)
        if must_exist and not os.path.exists(path):
            print(f"  Not found: {path}")
            continue
        if is_dir is True and not os.path.isdir(path):
            print(f"  Not a folder: {path}")
            continue
        if is_dir is False and not os.path.isfile(path):
            print(f"  Not a file: {path}")
            continue
        return path


def normalize_token(s):
    """Strip everything but letters/digits, uppercase, into one string."""
    return re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()


def build_pdf_index(source_folder):
    """Map filename -> its normalized token, for every PDF in source_folder."""
    index = []
    for fname in os.listdir(source_folder):
        if fname.lower().endswith(".pdf"):
            full = os.path.join(source_folder, fname)
            index.append((fname, normalize_token(os.path.splitext(fname)[0]), full))
    return index


def token_matches(doc_token, fname_token):
    """
    True if doc_token occurs in fname_token, allowing it to be glued
    directly to other text/letters on either side (e.g. "VCUSGSG6262"
    inside "TAXINGIVEVCUSGSG6262"), but NOT if that would silently
    extend a digit run - i.e. doc_token starting/ending in a digit must
    not have another digit immediately outside the match (so "27383829292"
    will not match inside "273838292920" or "2738382929212").
    """
    start = 0
    while True:
        idx = fname_token.find(doc_token, start)
        if idx == -1:
            return False
        end = idx + len(doc_token)
        before_digit_clash = doc_token[0].isdigit() and idx > 0 and fname_token[idx - 1].isdigit()
        after_digit_clash = doc_token[-1].isdigit() and end < len(fname_token) and fname_token[end].isdigit()
        if not before_digit_clash and not after_digit_clash:
            return True
        start = idx + 1


def find_match(doc_number, pdf_index):
    """
    Match if the document number's normalized token (separators like
    "/", "-", "_" stripped out) occurs inside the filename's normalized
    token - allowing it to be glued to surrounding text with no
    separator - while still refusing to match a shorter/longer digit
    run (see token_matches). Returns the matching (fname, full_path) or
    None. Also returns 'ambiguous' if more than one file matches.
    """
    doc_token = normalize_token(doc_number)
    if not doc_token:
        return None, "empty"

    matches = [
        (fname, full)
        for fname, fname_token, full in pdf_index
        if token_matches(doc_token, fname_token)
    ]

    if len(matches) == 0:
        return None, "not_found"
    if len(matches) > 1:
        return matches, "ambiguous"
    return matches[0], "ok"


def find_column(header_row, keywords):
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        name = str(cell).strip().lower()
        if any(k in name for k in keywords):
            return idx
    return None


def main():
    print("=== Copy PDFs by Document Number ===\n")

    excel_path = ask_path("Path to the Excel file (.xlsx): ", must_exist=True, is_dir=False)
    source_folder = ask_path("Path to the SOURCE folder (where the PDFs currently are): ", must_exist=True, is_dir=True)
    dest_folder = ask_path("Path to the DESTINATION folder (where matched PDFs should be copied): ", must_exist=False)
    os.makedirs(dest_folder, exist_ok=True)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active  # uses the first/active sheet

    header_row = [c.value for c in ws[1]]
    doc_col = find_column(header_row, ["document number", "doc number", "document no", "docket"])
    status_col = find_column(header_row, ["status"])

    if doc_col is None or status_col is None:
        print("\nCould not auto-detect columns. Here are the headers found:")
        for i, h in enumerate(header_row):
            print(f"  {i+1}. {h}")
        doc_col = int(input("Enter the column NUMBER for Document Number: ")) - 1
        status_col = int(input("Enter the column NUMBER for Status: ")) - 1
    else:
        print(f"\nUsing column '{header_row[doc_col]}' for document number, "
              f"'{header_row[status_col]}' for status.\n")

    pdf_index = build_pdf_index(source_folder)
    print(f"Found {len(pdf_index)} PDF(s) in source folder.\n")

    copied, skipped_done, not_found, ambiguous, empty = 0, 0, 0, 0, 0

    for row in ws.iter_rows(min_row=2):
        doc_cell = row[doc_col]
        status_cell = row[status_col]
        doc_number = doc_cell.value

        if doc_number is None or str(doc_number).strip() == "":
            empty += 1
            continue

        current_status = str(status_cell.value).strip().lower() if status_cell.value else ""
        if current_status == "copied":
            skipped_done += 1
            continue

        result, kind = find_match(doc_number, pdf_index)

        if kind == "ok":
            fname, full_path = result
            dest_path = os.path.join(dest_folder, fname)
            shutil.copy2(full_path, dest_path)
            status_cell.value = "Copied"
            copied += 1
            print(f"  [OK] {doc_number} -> {fname}")
        elif kind == "ambiguous":
            names = ", ".join(m[0] for m in result)
            status_cell.value = "Not found"
            ambiguous += 1
            print(f"  [AMBIGUOUS -> marked Not found] {doc_number} matched multiple files: {names}")
        elif kind == "not_found":
            status_cell.value = "Not found"
            not_found += 1
            print(f"  [MISSING] {doc_number} - no matching PDF found")
        elif kind == "empty":
            empty += 1

    wb.save(excel_path)

    print("\n=== Done ===")
    print(f"Copied:            {copied}")
    print(f"Already 'Copied':  {skipped_done}")
    print(f"Not found:         {not_found}")
    print(f"Ambiguous:         {ambiguous}")
    print(f"Empty rows:        {empty}")
    print(f"\nExcel updated in place: {excel_path}")


if __name__ == "__main__":
    main()
