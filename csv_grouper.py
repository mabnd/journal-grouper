"""
csv_grouper.py
==============
CSV entry point for the journal-entry grouping algorithm. Groups accounting
journal entry lines from a single-journal CSV file into balanced entries.

All algorithm logic lives in journal_grouper_core.py — this file only
handles reading a CSV into row dicts and writing the result back out as
CSV. See algorithm.txt for the full step-by-step specification, and
excel_grouper.py for the equivalent multi-sheet Excel entry point (same
core logic, different file format).

Usage:
    python csv_grouper.py input.csv [clients.csv]

The output is written next to the input as <input>_processed.csv. If a
clients list is given (and exists), partner values not found in it are
written to <input>_missing_clients.csv.
"""

import csv
import os
import sys

import journal_grouper_core as core
import config


CSV_DELIMITER      = config.CSV_DELIMITER
CLIENT_NAME_COLUMN = config.CLIENT_NAME_COLUMN


def parse_csv(filepath: str) -> tuple[list[dict], list[str]]:
    rows = []
    with open(filepath, newline="", encoding=config.CSV_ENCODING) as f:
        reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
        fieldnames_raw = [h.strip() for h in (reader.fieldnames or [])]
        fieldnames = core.normalize_fieldnames(fieldnames_raw)
        field_map = dict(zip(fieldnames_raw, fieldnames))
        for i, row in enumerate(reader):
            cleaned = {field_map.get(k.strip(), k.strip()): (v.strip() if v else "") for k, v in row.items()}
            cleaned[core.COL_DATE]  = core.normalize_date(cleaned.get(core.COL_DATE, ""))
            cleaned["_idx"]        = i
            cleaned["_debit_val"]  = core.parse_amount(cleaned.get(core.COL_DEBIT,  ""))
            cleaned["_credit_val"] = core.parse_amount(cleaned.get(core.COL_CREDIT, ""))
            rows.append(cleaned)
    return rows, fieldnames


def derive_path(input_path: str, suffix: str) -> str:
    root, ext = os.path.splitext(input_path)
    return f"{root}{suffix}{ext}"


def load_clients(filepath: str) -> set[str]:
    with open(filepath, newline="", encoding=config.CSV_ENCODING) as f:
        reader = csv.DictReader(f, delimiter=CSV_DELIMITER)
        return {
            row.get(CLIENT_NAME_COLUMN, "").strip()
            for row in reader
            if row.get(CLIENT_NAME_COLUMN, "").strip()
        }


def write_missing_clients(filepath: str, issues: dict) -> None:
    cols = [
        config.MISSING_CLIENT_COL_PARTNER,
        config.MISSING_CLIENT_COL_OCCURRENCES,
        config.MISSING_CLIENT_COL_STATUS,
        config.MISSING_CLIENT_COL_SUGGESTION,
    ]
    with open(filepath, "w", newline="", encoding=config.CSV_ENCODING) as f:
        writer = csv.DictWriter(f, fieldnames=cols, delimiter=CSV_DELIMITER)
        writer.writeheader()
        for name, (count, status, best_match) in sorted(issues.items()):
            writer.writerow({
                config.MISSING_CLIENT_COL_PARTNER:     name,
                config.MISSING_CLIENT_COL_OCCURRENCES: count,
                config.MISSING_CLIENT_COL_STATUS:      config.MISSING_CLIENT_STATUS_TYPO if status == core.REASON_PARTNER_TYPO else config.MISSING_CLIENT_STATUS_UNKNOWN,
                config.MISSING_CLIENT_COL_SUGGESTION:  best_match,
            })


def write_output(filepath, confirmed_groups, flagged_rows, fieldnames, resolution_meta):
    with open(filepath, "w", newline="", encoding=config.CSV_ENCODING) as f:
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
                sep[writer.fieldnames[0]] = payload
                writer.writerow(sep)


def run(input_path, clients_path=None):
    print(f"Reading: {input_path}")

    rows, fieldnames = parse_csv(input_path)

    missing = core.validate_required_columns(fieldnames)
    if missing:
        print(f"ABORTED — missing required column(s): {', '.join(missing)}")
        sys.exit(1)

    confirmed, flagged, resolution_meta, errors = core.process_and_report(rows)

    output_path = derive_path(input_path, "_processed")
    write_output(output_path, confirmed, flagged, fieldnames, resolution_meta)
    print(f"Output written to: {output_path}")

    if clients_path:
        if not os.path.exists(clients_path):
            print(f"Clients list not found, skipping partner check: {clients_path}")
        else:
            known_names = load_clients(clients_path)
            issues = core.find_partner_issues(rows, known_names)
            if issues:
                missing_path = derive_path(input_path, "_missing_clients")
                write_missing_clients(missing_path, issues)
                typos   = sum(1 for _, status, _ in issues.values() if status == core.REASON_PARTNER_TYPO)
                unknown = len(issues) - typos
                print(f"{len(issues)} partner issue(s) ({typos} possible typo(s), {unknown} unknown) — written to: {missing_path}")
            else:
                print("All partners found in clients list")


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        print("Usage: python csv_grouper.py <input.csv> [clients.csv]")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2] if len(sys.argv) == 3 else None)
