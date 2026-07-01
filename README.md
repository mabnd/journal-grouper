# Journal Entry Grouper

Takes a flat export of bookkeeping lines and reconstructs the original **journal entries** — grouping lines so that debits equal credits within each entry.

Works with CSV files and Excel workbooks.

**Input:**

| Date     | Communication | Debit   | Credit  |
|----------|---------------|---------|---------|
| 01/01/26 | Invoice #1042 | 500.00  |         |
| 01/01/26 | Salary Jan    | 3000.00 |         |
| 01/01/26 | Salary Jan    |         | 3000.00 |
| 01/01/26 | Invoice #1042 |         | 500.00  |
| 02/01/26 | Rent Q1       | 1200.00 |         |
| 02/01/26 | VAT           |  240.00 |         |
| 02/01/26 | Rent Q1       |         | 1440.00 |
| 03/01/26 | Misc expense  |  750.00 |         |

**Output:**

| Date     | Communication | Debit   | Credit  | CONFIDENCE_SCORE | FLAG_REASON                               |
|----------|---------------|---------|---------|------------------|-------------------------------------------|
| 01/01/26 | Invoice #1042 | 500.00  |         | 75               |                                           |
|          | Invoice #1042 |         | 500.00  |                  |                                           |
|          |               |         |         |                  |                                           |
| 01/01/26 | Salary Jan    | 3000.00 |         | 100              |                                           |
|          | Salary Jan    |         | 3000.00 |                  |                                           |
|          |               |         |         |                  |                                           |
| 02/01/26 | Rent Q1       | 1200.00 |         | 100              |                                           |
|          | VAT           |  240.00 |         |                  |                                           |
|          | Rent Q1       |         | 1440.00 |                  |                                           |
|          |               |         |         |                  |                                           |
| 03/01/26 | Misc expense  |  750.00 |         | 0                | No matching credit found — check manually |

## Output

A processed file (same name, `_processed` suffix) with:

- **Confirmed entries** — grouped lines, each with a confidence score showing how certain the grouping is
- **Flagged rows** — lines that couldn't be grouped automatically, with a reason for manual review

## Configuration

Edit [`config.py`](config.py) — the only file you need to touch:

| Section | What to set |
|---|---|
| **Input columns** | Your file's column names |
| **Output columns** | How to rename columns in the output |
| **Messages** | Text shown to reviewers — translate as needed |
| **Numeric settings** | How aggressively lines are grouped (defaults are fine) |

## Usage

```
pip install -r requirements.txt

python csv_grouper.py   input.csv   [clients.csv]   # clients.csv is optional
python excel_grouper.py input.xlsx  [clients.csv]   # clients.csv is optional
```

See [algorithm.md](algorithm.md) for a full description of how it works.
