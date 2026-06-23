"""
excel_grouper.py
================
Excel entry point for the journal-entry grouping algorithm. Each worksheet
in the input workbook is treated as one independent journal — same columns
as the CSV format (Journal, Code, Date de facturation, Communication,
Partenaire, Débit, Crédit), just one sheet per journal instead of one file
per journal.

All algorithm logic lives in journal_grouper_core.py — this file only
handles reading workbook sheets into row dicts and writing the result back
out as a new workbook (one output sheet per input sheet). See algorithm.txt
for the full step-by-step specification, and journal_grouper.py for the
equivalent single-journal CSV entry point (same core logic, different file
format).

Usage:
    python excel_grouper.py input.xlsx output.xlsx
"""

import sys
import datetime

import openpyxl

import journal_grouper_core as core


def cell_to_text(value):
    """Converts a raw Excel cell value into the same kind of stripped text
    a CSV would have given us — handles numbers stored as numbers (not
    text) and dates stored as native date/datetime objects."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, datetime.date):
        # Matches the ddmmyy text convention used in the CSV journals
        # (e.g. "020126" for 2026-01-02).
        return value.strftime("%d%m%y")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value == int(value) else str(value)
    return str(value).strip()


def read_sheet(ws, sheet_name: str) -> tuple[list[dict], list[str]]:
    rows_iter = ws.iter_rows(values_only=True)
    header_raw = next(rows_iter, None)
    if header_raw is None:
        return [], []

    # Sheets are often padded with empty trailing columns (e.g. out to Z)
    # left over from formatting — trim those off rather than carrying a
    # pile of blank-named columns into the output.
    last_real = max(
        (i for i, h in enumerate(header_raw) if h is not None), default=-1
    )
    header_raw = header_raw[:last_real + 1]
    ncols = len(header_raw)
    fieldnames = [cell_to_text(h) for h in header_raw]

    # If the sheet has no Journal column, the sheet name *is* the journal —
    # add the column ourselves rather than treating every row as missing it.
    has_journal_col = core.COL_JOURNAL in fieldnames
    if not has_journal_col:
        fieldnames = [core.COL_JOURNAL] + fieldnames

    rows = []
    idx = 0
    for raw_row in rows_iter:
        raw_row = raw_row[:ncols]
        if all(v is None for v in raw_row):
            continue
        row = {}
        if not has_journal_col:
            row[core.COL_JOURNAL] = sheet_name
        data_fields = fieldnames[1:] if not has_journal_col else fieldnames
        for col, value in zip(data_fields, raw_row):
            if col == core.COL_DEBIT or col == core.COL_CREDIT:
                row[col] = value
            else:
                row[col] = cell_to_text(value)
        row["_idx"]        = idx
        row["_debit_val"]  = core.parse_amount(row.get(core.COL_DEBIT))
        row["_credit_val"] = core.parse_amount(row.get(core.COL_CREDIT))
        rows.append(row)
        idx += 1
    return rows, fieldnames


def write_sheet(ws_out, confirmed_groups, flagged_rows, fieldnames, resolution_meta):
    ncols = len(fieldnames) + 2  # + CONFIDENCE_SCORE, FLAG_REASON
    output_fields = None
    for event, payload in core.generate_output_rows(
        confirmed_groups, flagged_rows, fieldnames, resolution_meta
    ):
        if event == "fields":
            output_fields = payload
            ws_out.append(output_fields)
        elif event == "row":
            ws_out.append([payload.get(k, "") for k in output_fields])
        elif event == "blank":
            ws_out.append([None] * ncols)
        elif event == "header":
            ws_out.append([payload] + [None] * (ncols - 1))


def run(input_path, output_path):
    print(f"Reading: {input_path}")
    wb_in = openpyxl.load_workbook(input_path, data_only=True)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    total_confirmed = 0
    total_flagged   = 0
    any_errors      = False

    for sheet_name in wb_in.sheetnames:
        ws = wb_in[sheet_name]
        rows, fieldnames = read_sheet(ws, sheet_name)

        print(f"\n--- Sheet: {sheet_name} ---")
        if not rows:
            print(f"[{sheet_name}] Empty sheet, skipping")
            wb_out.create_sheet(title=sheet_name)
            continue

        confirmed, flagged, resolution_meta, errors = core.process_and_report(
            rows, label=sheet_name
        )

        total_confirmed += len(confirmed)
        total_flagged   += len(flagged)
        any_errors       = any_errors or bool(errors)

        ws_out = wb_out.create_sheet(title=sheet_name)
        write_sheet(ws_out, confirmed, flagged, fieldnames, resolution_meta)

    print(f"\n=== Summary across {len(wb_in.sheetnames)} sheet(s) ===")
    print(f"  Total confirmed entries : {total_confirmed}")
    print(f"  Total flagged rows      : {total_flagged}")
    print(f"  Verification           : {'FAILED — see above' if any_errors else 'all sheets passed'}")

    wb_out.save(output_path)
    print(f"\nOutput written to: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python excel_grouper.py <input.xlsx> <output.xlsx>")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
