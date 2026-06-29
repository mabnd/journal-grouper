What the algorithm does
========================

You give it one accounting journal's worth of lines — either as a CSV file, or
as one sheet in a multi-sheet Excel workbook where each sheet is its own
journal. Either way it's a flat list of debit/credit lines with no markers
showing which lines belong together. The algorithm figures out, on its own,
which lines form a complete "journal entry" (a transaction where total debits
= total credits), and returns a cleaned-up file where those lines are visibly
grouped together. A multi-sheet workbook is processed one sheet at a time —
each journal is resolved entirely on its own, with one matching output sheet
per input sheet.


How it does it
========================

1. Clean the numbers.
   Convert French-formatted amounts ("1,135,000") into real numbers, and set
   aside any line with no debit or credit at all — it carries no information.

2. Bucket by date.
   Every line in one journal entry shares the same date, so grouping only ever
   happens within a single day's lines (with one deliberate exception, below).

3. Sweep through the file in its original order.
   This is the main grouping mechanism. Starting from the first line of a date,
   it keeps adding the next line in file order to a "current group." Once that
   group's debits equal its credits, it checks whether the next line's
   description is nearly identical — if so, it keeps growing the group
   (probably still the same entry); if not, it closes the group and starts a
   new one. The order lines appear in the file is the backbone of this whole
   step — nothing here is reordered.

4. Catch coincidental overlaps.
   Occasionally two separate entries balance together by sheer coincidence and
   get treated as one group in step 3. It checks for this by clustering lines
   with similar descriptions, and splits the group apart only if each cluster
   can independently balance on its own.

5. Rescue entries trapped inside a bigger mess.
   If a group never balances — usually because one of its lines ended up
   elsewhere — it first tries to pull out any complete, self-balancing entries
   hiding inside by their description, then searches mathematically for any
   leftover combination of lines whose amounts cancel out to zero.

6. Search the rest of the file for what's missing.
   If a group is still short, it looks everywhere else in the file for 1, 2,
   or 3 lines whose amounts, added together, exactly cover the gap (not each
   line individually — their combined total). If one candidate combination's
   descriptions clearly match better than any other, it's accepted; if several
   look equally plausible, it's left for a human to decide rather than guessed
   at.

7. One last attempt, then a safety net.
   Anything still unresolved gets a final try at splitting it into balanced
   pieces. Whatever's left after that is flagged for review, and a final check
   makes sure no line was silently lost along the way.

8. Verify, then write the result.
   Before producing any output, it confirms every single input line is
   accounted for exactly once and that the totals still match. Then it writes
   the grouped entries with a confidence score, plus a separate section for
   anything it couldn't resolve.


What it prioritizes
========================

- File order first, for grouping.
  The original sequence of lines is the primary signal used to propose where
  one entry ends and the next begins — it's not just "considered," it's how
  the first pass works.

- Balance over text, after that.
  Whether debits equal credits is the only fact it fully trusts beyond order.
  Descriptions only ever assist — confirming a group should stay open,
  distinguishing two coincidentally-balanced entries, or breaking a tie
  between several balanced candidates.

- Certainty over completeness.
  It refuses to guess when multiple explanations are equally plausible. It
  would rather flag something for a human than silently group it wrong.

- Every line matters, no exceptions.
  Even lines that look like duplicates are treated as real, deliberate
  entries — every one must end up somewhere in the output, exactly once.

- Visibility over silence.
  Anything uncertain is clearly marked with a reason and a low confidence
  score, rather than buried inside a possibly-wrong grouping.


What it returns
========================

A single output file — a CSV, or an Excel workbook with one output sheet per
input sheet — with two parts in each:

- Confirmed entries — each journal entry's lines grouped together (journal
  and date shown only once, at the top of each group), followed by a
  confidence score (0-100) reflecting how much guesswork was needed to
  assemble it. Entries that were trivially closed by the file-order sweep
  score highest; entries that needed clustering, extraction, or stray-line
  hunting score progressively lower.

- A flagged section — every line the algorithm couldn't resolve with
  confidence, each with a plain-language reason and a score of 0, listed in
  the same order they appeared in the original file so patterns are easy
  to spot.

And underlying both: a guarantee, checked on every run for every journal,
that the number of lines and the total debit/credit amounts going in exactly
match what comes out.

The output columns are renamed and reordered to a fixed layout (matching
an accounting import template) rather than just mirroring whatever the
source file called its own columns, and dates are always shown as
dd/mm/yyyy no matter what format the source used.

Optionally, you can also give it a clients list to check the Partenaire
values against — see "How to run it" below.

One thing worth knowing: some journals close out many small expense lines
recorded on different days with a single summary line much later (a
"running total since last close" style), instead of balancing within one
day. The algorithm can't see that pattern — it sorts by date before it ever
checks balance — so every line in an entry built that way gets flagged
rather than guessed at. This was found in a petty-cash journal and is
treated as a known limitation rather than something the algorithm tries to
work around automatically.


How to run it
========================

There are two scripts, one per input format. Both produce the same kind of
output and use the exact same grouping logic underneath.

For a single journal in a CSV file:

    python csv_grouper.py input.csv [clients.csv]

For a workbook with multiple journals, one per sheet:

    python excel_grouper.py input.xlsx [clients.csv]

The output is written next to the input automatically (input.csv produces
input_processed.csv, and likewise for Excel) — no need to name it yourself.
If a sheet in the Excel workbook has no "Journal" column, the sheet's own
name is used as the journal name automatically.

The clients.csv argument is optional. If given, partner names that don't
appear in it are reported separately — as a second file for CSV input, or
as an extra sheet for Excel input. The report tells apart names that look
like a typo of a known client (with a suggested match to verify) from
names with no close match at all.

Both scripts expect the Date de facturation, Communication, Partenaire,
Débit, and Crédit columns to be present (under those exact names) — these
are what the grouping logic actually reads. If a CSV is missing one, the
run stops before writing anything, so you don't get a wrong result by
mistake. If an Excel sheet is missing one, only that sheet is skipped (with
a clear note in its place) — the rest of the workbook still processes
normally.
