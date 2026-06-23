"""
journal_grouper.py
==================
CSV entry point for the journal-entry grouping algorithm. Groups accounting
journal entry lines from a single-journal CSV file into balanced entries.

All algorithm logic lives in journal_grouper_core.py — this file only
handles reading a CSV into row dicts and writing the result back out as
CSV. See algorithm.txt for the full step-by-step specification, and
excel_grouper.py for the equivalent multi-sheet Excel entry point (same
core logic, different file format).

Usage:
    python journal_grouper.py input.csv output.csv
"""

import csv
import sys

import journal_grouper_core as core


CSV_DELIMITER = ","


def parse_csv(filepath: str) -> tuple[list[dict], list[str]]:
    rows = []
    with open(filepath, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
        fieldnames = [h.strip() for h in (reader.fieldnames or [])]
        for i, row in enumerate(reader):
            cleaned = {k.strip(): (v.strip() if v else "") for k, v in row.items()}
            cleaned["_idx"]        = i
            cleaned["_debit_val"]  = core.parse_amount(cleaned.get(core.COL_DEBIT,  ""))
            cleaned["_credit_val"] = core.parse_amount(cleaned.get(core.COL_CREDIT, ""))
            rows.append(cleaned)
    return rows, fieldnames


def write_output(filepath, confirmed_groups, flagged_rows, fieldnames, resolution_meta):
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = None
        for event, payload in core.generate_output_rows(
            confirmed_groups, flagged_rows, fieldnames, resolution_meta
        ):
            if event == "fields":
                writer = csv.DictWriter(
                    f, fieldnames=payload,
                    delimiter=CSV_DELIMITER, extrasaction="ignore",
                )
                writer.writeheader()
            elif event == "row":
                writer.writerow(payload)
            elif event == "blank":
                writer.writerow({k: "" for k in writer.fieldnames})
            elif event == "header":
                sep = {k: "" for k in writer.fieldnames}
                sep[core.COL_JOURNAL] = payload
                writer.writerow(sep)


def run(input_path, output_path):
    print(f"Reading: {input_path}")

    rows, fieldnames = parse_csv(input_path)

    confirmed, flagged, resolution_meta, errors = core.process_and_report(rows)

    write_output(output_path, confirmed, flagged, fieldnames, resolution_meta)
    print(f"Output written to: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python journal_grouper.py <input.csv> <output.csv>")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
