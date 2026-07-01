# Journal Entry Grouper

Accounting systems often export journals as a flat list of lines with no explicit grouping — every debit and credit is on its own row, with no marker showing where one journal entry ends and the next begins. This tool reconstructs those groupings automatically, using the fundamental double-entry rule that debits must equal credits within each entry.

It takes a flat list of bookkeeping lines and reassembles them into **journal entries** — groups where total debits = total credits.

Input is a CSV file or an Excel workbook. Each sheet in a multi-sheet workbook is treated as one independent journal. See [algorithm.txt](algorithm.txt) for a full description of how it works.

---

## Output

A processed file written next to the input (`_processed` suffix), containing:

- **Confirmed entries** — lines grouped together, each with a confidence score (0–100) reflecting how much guesswork was needed. Higher is better.
- **Flagged rows** — lines the algorithm couldn't resolve with confidence, each with a plain-language reason, for manual review.

Every input line is guaranteed to appear in the output exactly once, and debit/credit totals are verified to match before anything is written.

**Optional partner check:** if a clients list is provided, partner values not found in it are reported separately — as a second CSV file or an extra sheet — classified as probable typos or unknown clients.

> **Known limitation:** entries that balance across multiple dates (e.g. a running-total petty-cash style) cannot be detected — every line in such an entry will be flagged rather than guessed at.

---

## Configuration

Edit **`config.py`** — the only file you need to touch. It has four sections:

| Section | What to set |
|---|---|
| **Input column names** | Map each algorithm role to your file's column headers. Required: `date`, `communication`, `debit`, `credit`. Optional: `journal`, `code`, `partner`. |
| **Output column names** | Rename any output column — core or extra passthrough. Anything not listed keeps its original name. |
| **User-facing strings** | Flag reasons and missing-client report text. Translate or rewrite to match your language. |
| **Algorithm tuning** | Thresholds controlling grouping aggressiveness. Defaults work well for standard accounting exports. |

---

## Usage

```
pip install -r requirements.txt

python csv_grouper.py   input.csv   [clients.csv]
python excel_grouper.py input.xlsx  [clients.csv]
```

Column names are matched **case-insensitively**. If a required column is missing, a CSV run stops immediately; an Excel sheet is skipped with a note while the rest of the workbook still processes normally.
