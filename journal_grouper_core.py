"""
journal_grouper_core.py
========================
Format-agnostic core of the journal-entry grouping algorithm. See
algorithm.txt for the full step-by-step specification.

This module knows nothing about CSV or Excel — it operates purely on lists
of row dicts. Each row dict must carry:
  - "_idx"        : a unique integer index, stable across the whole run
                    for the journal/sheet this row belongs to
  - "_debit_val"  : the debit amount as a float (0.0 if none)
  - "_credit_val" : the credit amount as a float (0.0 if none)
  - the text fields named by COL_JOURNAL, COL_CODE, COL_DATE, COL_COMM,
    COL_PARTNER, COL_DEBIT, COL_CREDIT

Both journal_grouper.py (CSV) and excel_grouper.py (Excel) parse their own
input format into this shape, call process_and_report() here, and then
format the result (generate_output_rows()) into their own output format.
This keeps the algorithm itself in exactly one place.

Algorithm steps:
  1.  Parse and clean the input (amount parsing lives here; reading the
      file itself is format-specific and lives in each entry point)
  2.  Group by date
  3.  Propose initial groups using a balance-first sweep
  4.  Refine balanced groups with mixed communications
  5a. Extract balanced sub-groups from unbalanced groups
  5b. Hunt for stray lines (up to 3)
  6.  Final split attempt on remaining unbalanced groups
  7.  Final conflict check
  8a. Verification
  8b. Compute confidence scores (0-100) for each confirmed entry
  8c. Assign review priority to flagged rows
  8d. Generate output rows (format-agnostic event stream)

Every input line — including exact textual duplicates of another line —
is a legitimate transaction and must appear in the output exactly once.
Duplicates are never removed or specially flagged; they are grouped and
balanced like any other line.
"""

from itertools import combinations
from difflib import SequenceMatcher
from collections import Counter


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

COL_JOURNAL = "Journal"
COL_CODE    = "Code"
COL_DATE    = "Date de facturation"
COL_COMM    = "Communication"
COL_PARTNER = "Partenaire"
COL_DEBIT   = "Débit"
COL_CREDIT  = "Crédit"

# Columns the algorithm actually depends on. Missing any of these doesn't
# raise an error on its own — every lookup falls back to "" or 0.0 — but it
# silently changes behavior (e.g. a missing Communication column makes every
# blank "match" every other blank, merging unrelated entries with full
# confidence). COL_JOURNAL is deliberately excluded: a missing Journal column
# is a supported case, handled by falling back to the sheet name. COL_CODE
# is excluded too since it's pure passthrough, never read by the algorithm.
REQUIRED_COLUMNS = [COL_DATE, COL_COMM, COL_PARTNER, COL_DEBIT, COL_CREDIT]


def validate_required_columns(fieldnames) -> list[str]:
    """Returns the subset of REQUIRED_COLUMNS missing from `fieldnames`,
    preserving REQUIRED_COLUMNS order. Empty list means all present."""
    present = set(fieldnames)
    return [col for col in REQUIRED_COLUMNS if col not in present]


# A balanced group is only kept open if the next line's communication
# is at or above this similarity — otherwise it is closed and a new group starts.
FUZZY_MERGE_THRESHOLD = 0.95

# Stray-line hunting: top candidate must beat second by this margin to win.
STRAY_CONFIDENCE_MARGIN = 0.20

# Rounding tolerance for balance checks
BALANCE_TOLERANCE = 0.01

# Step 5a Pass 2 subset-sum DP guard — bounds the number of distinct partial
# sums tracked, not the line count, so it stays safe even on large groups.
MAX_SUBSET_SUM_STATES = 200_000

# Step 6 recursive split guard (spec: capped at 20 lines)
MAX_SPLIT_SIZE = 20

# Recursive split depth guard
MAX_SPLIT_DEPTH = 10

# Flag reasons
REASON_ZERO_AMOUNT     = "Zero or missing amount on both debit and credit"
REASON_INCOMPLETE      = "Incomplete entry — no matching line found to balance"
REASON_AMBIGUOUS_STRAY = "Ambiguous stray — multiple candidates found, none clearly better"
REASON_AMBIGUOUS_SPLIT = "Ambiguous split — multiple entries on same date cannot be cleanly separated"
REASON_ORPHAN          = "Orphaned line — not assigned to any group"

# Priority levels
PRIORITY_HIGH   = "HIGH   — Entry incomplete, no candidate found anywhere in file"
PRIORITY_MEDIUM = "MEDIUM — Ambiguous stray, multiple candidates equally plausible"
PRIORITY_LOW    = "LOW    — Data quality issue (zero amount)"
PRIORITY_INFO   = "INFO   — Orphaned line, may be related to another flagged entry"

PRIORITY_MAP = {
    REASON_INCOMPLETE:      PRIORITY_HIGH,
    REASON_AMBIGUOUS_STRAY: PRIORITY_MEDIUM,
    REASON_AMBIGUOUS_SPLIT: PRIORITY_MEDIUM,
    REASON_ZERO_AMOUNT:     PRIORITY_LOW,
    REASON_ORPHAN:          PRIORITY_INFO,
}


# ---------------------------------------------------------------------------
# Step 1 — Parse and clean (amount parsing only; reading the file is
# the responsibility of each format-specific entry point)
# ---------------------------------------------------------------------------

def parse_amount(value) -> float:
    """Accepts either a native number (already a number cell, e.g. from
    Excel) or French-formatted text (e.g. from CSV) and returns a float."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str) or not value.strip():
        return 0.0
    raw = value.strip()
    negative = raw.startswith("-")
    if negative:
        raw = raw[1:].strip()
    raw = raw.replace("\xa0", "").replace(" ", "")
    has_dot   = "." in raw
    has_comma = "," in raw
    if has_dot and has_comma:
        if raw.rindex(",") > raw.rindex("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif has_comma:
        parts = raw.split(",")
        if all(len(p) == 3 for p in parts[1:]):
            raw = raw.replace(",", "")
        else:
            raw = raw.replace(",", ".")
    try:
        result = float(raw)
        return -result if negative else result
    except ValueError:
        return 0.0


def flag_zero_amount_lines(rows):
    clean, flagged = [], []
    for row in rows:
        if row["_debit_val"] == 0.0 and row["_credit_val"] == 0.0:
            flagged.append({**row, "_flag_reason": REASON_ZERO_AMOUNT})
        else:
            clean.append(row)
    return clean, flagged


# ---------------------------------------------------------------------------
# Step 2 — Group by date
# ---------------------------------------------------------------------------

def group_by_date(rows):
    buckets = {}
    for row in rows:
        date = row.get(COL_DATE, "")
        if isinstance(date, str):
            date = date.strip()
        buckets.setdefault(date, []).append(row)
    return buckets


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def comm_similarity(a: str, b: str) -> float:
    if not a and not b: return 1.0
    if not a or not b:  return 0.0
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def is_balanced(rows: list[dict]) -> bool:
    d = sum(r["_debit_val"]  for r in rows)
    c = sum(r["_credit_val"] for r in rows)
    return abs(d - c) <= BALANCE_TOLERANCE


def missing_amount(rows: list[dict]) -> float:
    return sum(r["_debit_val"] for r in rows) - sum(r["_credit_val"] for r in rows)


def representative_comm(rows: list[dict]) -> str:
    """Most common communication among `rows`. Ties are broken by first
    occurrence in `rows`, not by set iteration order — set iteration order
    depends on string hash randomization, which varies between Python
    processes and would otherwise make the result nondeterministic."""
    comms = [r.get(COL_COMM, "") for r in rows]
    if not comms:
        return ""
    counts = Counter(comms)
    max_count = max(counts.values())
    for c in comms:
        if counts[c] == max_count:
            return c
    return comms[0]


def build_comm_clusters(rows: list[dict]) -> list[list[dict]]:
    """Single-pass clustering: a line joins the first existing cluster whose
    founding communication matches it above FUZZY_MERGE_THRESHOLD, otherwise
    it starts a new cluster. Used by Step 4 and Step 5a Pass 1."""
    clusters: list[dict] = []
    for r in rows:
        comm = r.get(COL_COMM, "")
        for cluster in clusters:
            if comm_similarity(cluster["rep"], comm) >= FUZZY_MERGE_THRESHOLD:
                cluster["rows"].append(r)
                break
        else:
            clusters.append({"rep": comm, "rows": [r]})
    return [c["rows"] for c in clusters]


# ---------------------------------------------------------------------------
# Step 3 — Propose groups (balance-first sweep)
# ---------------------------------------------------------------------------

def propose_groups(date_rows: list[dict]) -> list[list[dict]]:
    if not date_rows:
        return []
    groups: list[list[dict]] = []
    current: list[dict] = [date_rows[0]]
    for row in date_rows[1:]:
        if is_balanced(current):
            rep   = representative_comm(current)
            new_c = row.get(COL_COMM, "")
            sim   = comm_similarity(rep, new_c)
            if sim < FUZZY_MERGE_THRESHOLD:
                groups.append(current)
                current = [row]
                continue
        current.append(row)
    groups.append(current)
    return groups


# ---------------------------------------------------------------------------
# Step 5a — Extract balanced sub-groups from an unbalanced group
# ---------------------------------------------------------------------------

def _extract_pass1_communication_guided(rows: list[dict]) -> tuple[list[list[dict]], list[dict]]:
    """Pass 1 — cluster by communication, extract any cluster that balances
    independently, and repeat on what's left until no more can be extracted."""
    extracted: list[list[dict]] = []
    remaining = list(rows)
    changed = True
    while changed and len(remaining) >= 1:
        changed = False
        clusters = build_comm_clusters(remaining)
        if len(clusters) <= 1:
            break
        balancing = [c for c in clusters if len(c) < len(remaining) and is_balanced(c)]
        if not balancing:
            break
        balancing_idxs = {r["_idx"] for c in balancing for r in c}
        extracted.extend(balancing)
        remaining = [r for r in remaining if r["_idx"] not in balancing_idxs]
        changed = True
    return extracted, remaining


def find_one_balanced_subset(rows: list[dict], max_states: int = MAX_SUBSET_SUM_STATES):
    """Subset-sum search via dynamic programming: returns the smallest
    subset (size >= 2) of `rows` whose net amounts sum to zero, or None.

    This is a drop-in replacement for brute-force combinatorial enumeration.
    Enumerating combinations is exponential in the number of lines, which is
    why the spec caps that search at a small line count. Subset-sum DP is
    instead bounded by the number of distinct achievable partial sums, which
    stays small for realistic accounting amounts (lots of repeated/related
    values) even when the line count is large — so the search can run
    safely over much bigger groups without an arbitrary line-count cap.
    """
    cents = [round((r["_credit_val"] - r["_debit_val"]) * 100) for r in rows]
    achievable: dict[int, tuple] = {0: ()}  # sum -> combo of indices; 0:() is a fixed anchor, never overwritten
    best_zero = None
    for i, amt in enumerate(cents):
        if amt == 0:
            continue
        for s, combo in list(achievable.items()):
            ns = s + amt
            new_combo = combo + (i,)
            if ns == 0:
                if best_zero is None or len(new_combo) < len(best_zero):
                    best_zero = new_combo
                continue
            if ns not in achievable:
                achievable[ns] = new_combo
        if len(achievable) > max_states:
            break
    if best_zero:
        return [rows[i] for i in best_zero]
    return None


def _extract_pass2_subset_sum(rows: list[dict]) -> tuple[list[list[dict]], list[dict]]:
    """Pass 2 — repeatedly extract the smallest balanced subset found via
    subset-sum DP until none remain."""
    extracted: list[list[dict]] = []
    remaining = list(rows)
    while len(remaining) >= 2:
        subset = find_one_balanced_subset(remaining)
        if not subset:
            break
        extracted.append(subset)
        subset_idxs = {r["_idx"] for r in subset}
        remaining = [r for r in remaining if r["_idx"] not in subset_idxs]
    return extracted, remaining


def extract_balanced_subgroups(rows: list[dict]) -> tuple[list[list[dict]], list[dict]]:
    pass1_extracted, remaining = _extract_pass1_communication_guided(rows)
    pass2_extracted, remaining = _extract_pass2_subset_sum(remaining)
    return pass1_extracted + pass2_extracted, remaining


# ---------------------------------------------------------------------------
# Step 5b — Hunt for stray lines (up to 3, via an amount index)
# ---------------------------------------------------------------------------

def build_amount_index(pool):
    """Maps each rounded net amount (credit - debit) to the lines carrying
    it, so a single stray lookup is O(1) and pair/triple lookups only need
    to probe the index instead of scanning the whole pool."""
    index: dict[float, list[dict]] = {}
    for r in pool:
        net = round(r["_credit_val"] - r["_debit_val"], 2)
        index.setdefault(net, []).append(r)
    return index


def find_stray(group, pool):
    needed = missing_amount(group)
    rep    = representative_comm(group)

    if abs(needed) <= BALANCE_TOLERANCE:
        return "not_found", None

    needed_r = round(needed, 2)
    index    = build_amount_index(pool)

    # Candidates are pre-filtered by amount before any combination loop —
    # a line whose individual net amount exceeds the missing balance can
    # never be part of a valid pair/triple.
    candidates = [
        r for r in pool
        if abs(r["_credit_val"] - r["_debit_val"]) <= abs(needed_r) + BALANCE_TOLERANCE
    ]

    matching_combos: list[list[dict]] = []

    # Size 1 — direct O(1) lookup.
    matching_combos.extend([r] for r in index.get(needed_r, []))

    # Size 2 — for each plausible candidate, look up whether the remainder exists.
    if not matching_combos:
        seen = set()
        for r1 in candidates:
            amt1      = r1["_credit_val"] - r1["_debit_val"]
            remainder = round(needed_r - amt1, 2)
            for r2 in index.get(remainder, []):
                if r2["_idx"] == r1["_idx"]:
                    continue
                key = frozenset((r1["_idx"], r2["_idx"]))
                if key in seen:
                    continue
                seen.add(key)
                matching_combos.append([r1, r2])

    # Size 3 — same idea, one level deeper.
    if not matching_combos:
        seen = set()
        for i, r1 in enumerate(candidates):
            amt1 = r1["_credit_val"] - r1["_debit_val"]
            for r2 in candidates[i + 1:]:
                amt2      = r2["_credit_val"] - r2["_debit_val"]
                remainder = round(needed_r - amt1 - amt2, 2)
                for r3 in index.get(remainder, []):
                    if r3["_idx"] in (r1["_idx"], r2["_idx"]):
                        continue
                    key = frozenset((r1["_idx"], r2["_idx"], r3["_idx"]))
                    if key in seen:
                        continue
                    seen.add(key)
                    matching_combos.append([r1, r2, r3])

    if not matching_combos:
        return "not_found", None

    if len(matching_combos) == 1:
        return "balanced", matching_combos[0]

    def avg_sim(combo):
        sims = [comm_similarity(rep, r.get(COL_COMM, "")) for r in combo]
        return sum(sims) / len(sims) if sims else 0.0

    scored       = sorted(matching_combos, key=avg_sim, reverse=True)
    top_score    = avg_sim(scored[0])
    second_score = avg_sim(scored[1])

    if top_score - second_score > STRAY_CONFIDENCE_MARGIN:
        return "balanced", scored[0]

    return "ambiguous", None


# ---------------------------------------------------------------------------
# Step 6 — Recursively split into as many balancing sub-entries as possible
# ---------------------------------------------------------------------------

def try_split(rows, depth=0, memo=None):
    if memo is None:
        memo = {}

    n = len(rows)
    if n < 2 or n > MAX_SPLIT_SIZE or depth >= MAX_SPLIT_DEPTH:
        return None

    key = frozenset(r["_idx"] for r in rows)
    if key in memo:
        return memo[key]

    best       = None
    best_count = 0

    for size in range(1, n):
        for combo in combinations(range(n), size):
            a = [rows[i] for i in combo]
            b = [rows[i] for i in range(n) if i not in combo]

            if not (is_balanced(a) and is_balanced(b)):
                continue

            sub_a  = try_split(a, depth + 1, memo) or [a]
            sub_b  = try_split(b, depth + 1, memo) or [b]
            result = sub_a + sub_b
            count  = len(result)

            if count > best_count:
                best_count = count
                best = result

    memo[key] = best
    return best


# ---------------------------------------------------------------------------
# Core resolution — Steps 3 to 6 for one date bucket
# ---------------------------------------------------------------------------

def resolve_date_bucket(date_rows, global_pool):
    confirmed:      list[list[dict]] = []
    flagged:        list[dict]       = []
    used_from_pool: list[dict]       = []

    # Resolution metadata keyed by id(group list):
    # (resolution_method, stray_count, stray_crossdate)
    resolution_meta: dict[int, tuple] = {}

    def confirm(group, resolution="sweep", stray_count=0, stray_crossdate=False):
        confirmed.append(group)
        resolution_meta[id(group)] = (resolution, stray_count, stray_crossdate)

    proposed          = propose_groups(date_rows)
    unbalanced_groups = [g for g in proposed if not is_balanced(g)]

    # Step 4 — refine balanced groups with mixed communications
    for g in proposed:
        if not is_balanced(g):
            continue

        if len(g) <= 3:
            confirm(g, "sweep")
            continue

        rep = representative_comm(g)
        all_similar = all(
            comm_similarity(rep, r.get(COL_COMM, "")) >= FUZZY_MERGE_THRESHOLD
            for r in g
        )

        if all_similar:
            confirm(g, "sweep")
            continue

        clusters = build_comm_clusters(g)
        balancing_clusters = [c for c in clusters if is_balanced(c)]

        if len(balancing_clusters) >= 2:
            balancing_idxs = {r["_idx"] for c in balancing_clusters for r in c}
            for c in balancing_clusters:
                confirm(c, "split")
            remainder = [r for r in g if r["_idx"] not in balancing_idxs]
            if remainder:
                confirm(remainder, "split")
        else:
            confirm(g, "sweep")

    local_lines = [r for g in unbalanced_groups for r in g]
    pool        = local_lines + global_pool
    bucket_idxs = {r["_idx"] for r in date_rows}
    used_idxs: set[int] = set()

    for group in unbalanced_groups:
        group = [r for r in group if r["_idx"] not in used_idxs]
        if not group:
            continue
        if is_balanced(group):
            confirm(group, "sweep")
            continue

        # Step 5a — extract trapped balanced sub-entries
        extracted, remaining = extract_balanced_subgroups(group)
        if extracted:
            for sub in extracted:
                confirm(sub, "extracted")
                for r in sub:
                    used_idxs.add(r["_idx"])
            group = remaining

        if not group:
            continue

        if is_balanced(group):
            confirm(group, "sweep")
            continue

        # Step 5b — hunt for stray lines
        available = [
            r for r in pool
            if r["_idx"] not in used_idxs and r not in group
        ]
        status, strays = find_stray(group, available)

        if status == "balanced" and strays:
            merged     = group + strays
            crossdate  = any(r["_idx"] not in bucket_idxs for r in strays)
            if is_balanced(merged):
                confirm(merged, "stray",
                        stray_count=len(strays),
                        stray_crossdate=crossdate)
                for s in strays:
                    used_idxs.add(s["_idx"])
                    if s in global_pool:
                        used_from_pool.append(s)
                continue

        if status == "ambiguous":
            for row in group:
                flagged.append({**row, "_flag_reason": REASON_AMBIGUOUS_STRAY})
            continue

        # Step 6 — final recursive split attempt
        split = try_split(group)
        if split:
            for sub in split:
                confirm(sub, "split")
        else:
            for row in group:
                flagged.append({**row, "_flag_reason": REASON_INCOMPLETE})

    return confirmed, flagged, used_from_pool, resolution_meta


# ---------------------------------------------------------------------------
# Optional step — Partner existence check
#
# Independent of grouping/balancing: just reports which non-blank COL_PARTNER
# values in the journal don't appear in a client list supplied by the caller.
# Reading that client list is format-specific and lives in each entry point,
# same as journal/Excel parsing — this only takes the resulting set of names.
#
# Unmatched names are further split into two buckets: "typo" (close enough to
# one known name, by the same fuzzy comparison used for communications, that
# it's probably that name misspelled) and "unknown" (no close match at all).
# This is a suggestion for a human to verify, never an automatic correction.
# ---------------------------------------------------------------------------

# How close an unmatched partner name must be to a known client name to be
# reported as a probable typo instead of a flat unknown.
PARTNER_TYPO_THRESHOLD = 0.85

REASON_PARTNER_TYPO    = "typo"
REASON_PARTNER_UNKNOWN = "unknown"


def normalize_partner_name(name) -> str:
    return name.strip().casefold() if isinstance(name, str) else ""


def find_partner_issues(rows, known_names: set) -> dict:
    """Returns {partner name (as written in the journal): (line count,
    status, best_match)} for every non-blank COL_PARTNER value whose
    normalized form isn't exactly in `known_names`. status is
    REASON_PARTNER_TYPO with best_match set to the closest known name when
    similarity is at or above PARTNER_TYPO_THRESHOLD, otherwise
    REASON_PARTNER_UNKNOWN with best_match empty."""
    known_norms = {normalize_partner_name(n) for n in known_names}

    counts: dict[str, int] = {}
    for r in rows:
        partner = r.get(COL_PARTNER, "")
        if not isinstance(partner, str) or not partner.strip():
            continue
        if normalize_partner_name(partner) in known_norms:
            continue
        name = partner.strip()
        counts[name] = counts.get(name, 0) + 1

    issues: dict[str, tuple] = {}
    for name, count in counts.items():
        best_match, best_score = "", 0.0
        for known_name in known_names:
            score = comm_similarity(name, known_name)
            if score > best_score:
                best_score, best_match = score, known_name
        if best_score >= PARTNER_TYPO_THRESHOLD:
            issues[name] = (count, REASON_PARTNER_TYPO, best_match)
        else:
            issues[name] = (count, REASON_PARTNER_UNKNOWN, "")
    return issues


# ---------------------------------------------------------------------------
# Step 7 — Final conflict check
# ---------------------------------------------------------------------------

def final_conflict_check(confirmed_groups, flagged, all_rows):
    assigned = {row["_idx"] for group in confirmed_groups for row in group}
    assigned |= {row["_idx"] for row in flagged}
    return [
        {**row, "_flag_reason": REASON_ORPHAN}
        for row in all_rows
        if row["_idx"] not in assigned
    ]


# ---------------------------------------------------------------------------
# Step 8a — Verification
# ---------------------------------------------------------------------------

def verify(original_rows, confirmed_groups, flagged_rows):
    errors: list[str] = []

    confirmed_idxs = [row["_idx"] for group in confirmed_groups for row in group]
    flagged_idxs   = [row["_idx"] for row in flagged_rows]
    all_output_idxs = confirmed_idxs + flagged_idxs
    original_idxs   = [row["_idx"] for row in original_rows]

    # Check 1 — Line count
    if len(all_output_idxs) != len(original_idxs):
        errors.append(
            f"LINE COUNT MISMATCH: input has {len(original_idxs)} lines "
            f"but output has {len(all_output_idxs)} lines "
            f"({len(confirmed_idxs)} confirmed + {len(flagged_idxs)} flagged)"
        )

    # Check 2 — No line appears twice
    seen: dict[int, int] = {}
    for idx in all_output_idxs:
        seen[idx] = seen.get(idx, 0) + 1
    duplicated = [idx for idx, count in seen.items() if count > 1]
    if duplicated:
        errors.append(
            f"DUPLICATE LINES IN OUTPUT: {len(duplicated)} line(s) appear "
            f"more than once — indices: {duplicated[:10]}"
            f"{'...' if len(duplicated) > 10 else ''}"
        )

    # Check 2b — No line from input is missing
    missing = set(original_idxs) - set(all_output_idxs)
    if missing:
        errors.append(
            f"MISSING LINES IN OUTPUT: {len(missing)} input line(s) do not "
            f"appear in output — indices: {sorted(missing)[:10]}"
            f"{'...' if len(missing) > 10 else ''}"
        )

    # Check 3 — Amount reconciliation
    orig_debit  = sum(r["_debit_val"]  for r in original_rows)
    orig_credit = sum(r["_credit_val"] for r in original_rows)
    out_rows    = [row for group in confirmed_groups for row in group] + flagged_rows
    out_debit   = sum(r["_debit_val"]  for r in out_rows)
    out_credit  = sum(r["_credit_val"] for r in out_rows)

    if abs(orig_debit - out_debit) > BALANCE_TOLERANCE:
        errors.append(
            f"DEBIT MISMATCH: original={orig_debit:.2f}, "
            f"output={out_debit:.2f}, diff={orig_debit - out_debit:.2f}"
        )
    if abs(orig_credit - out_credit) > BALANCE_TOLERANCE:
        errors.append(
            f"CREDIT MISMATCH: original={orig_credit:.2f}, "
            f"output={out_credit:.2f}, diff={orig_credit - out_credit:.2f}"
        )

    return errors


# ---------------------------------------------------------------------------
# Step 8b — Compute confidence score (0–100) for each confirmed group
# ---------------------------------------------------------------------------

def compute_confidence(group, resolution, stray_count, stray_crossdate):
    """
    Start at 100 and deduct for each factor that reduces certainty:

    Resolution method:
      sweep    → 0 deduction  (balanced cleanly from the file order)
      split    → -10          (required splitting a balanced group)
      extracted→ -15          (rescued from inside a larger unbalanced group)
      stray    → -10 base, then additional penalties below

    Stray penalties (only when resolution == "stray"):
      base -10 covers the first stray line; -10 for each additional stray
      stray from a different date → -15

    Communication consistency:
      mixed communications across lines → -10

    Line proximity:
      non-consecutive indices in original file → -10
    """
    score = 100

    if resolution == "split":
        score -= 10
    elif resolution == "extracted":
        score -= 15
    elif resolution == "stray":
        score -= 10
        score -= (stray_count - 1) * 10
        if stray_crossdate:
            score -= 15

    # Communication consistency
    comms = [r.get(COL_COMM, "") for r in group]
    rep   = representative_comm(group)
    all_similar = all(
        comm_similarity(rep, c) >= FUZZY_MERGE_THRESHOLD for c in comms
    )
    if not all_similar:
        score -= 10

    # Line proximity
    idxs         = sorted(r["_idx"] for r in group)
    consecutive  = all(idxs[i+1] - idxs[i] == 1 for i in range(len(idxs) - 1))
    if not consecutive:
        score -= 10

    return max(0, score)


# ---------------------------------------------------------------------------
# Step 8c — Assign review priority to flagged rows
# ---------------------------------------------------------------------------

def assign_priorities(flagged_rows):
    for row in flagged_rows:
        reason = row.get("_flag_reason", "")
        row["_priority"] = PRIORITY_MAP.get(reason, PRIORITY_INFO)

    priority_order = {
        PRIORITY_HIGH:   0,
        PRIORITY_MEDIUM: 1,
        PRIORITY_LOW:    2,
        PRIORITY_INFO:   3,
    }
    flagged_rows.sort(key=lambda r: (
        priority_order.get(r["_priority"], 99),
        r.get(COL_DATE, "") if isinstance(r.get(COL_DATE, ""), str) else str(r.get(COL_DATE, "")),
        r["_idx"],
    ))
    return flagged_rows


# ---------------------------------------------------------------------------
# Orchestration — runs steps 1 (zero-flagging) through 8c for one
# journal's worth of rows, and prints the same progress/verification
# messages regardless of the input format.
# ---------------------------------------------------------------------------

def process_and_report(rows, label=None, log=print):
    """Runs the full pipeline (Steps 1 zero-flagging through 8c) for one
    journal's rows. `rows` must already carry _idx/_debit_val/_credit_val.
    Returns (confirmed_groups, flagged_rows, resolution_meta, errors)."""
    prefix = f"[{label}] " if label else ""

    original_rows = list(rows)
    log(f"{prefix}Loaded {len(original_rows)} lines")

    rows, zero_flagged = flag_zero_amount_lines(rows)
    if zero_flagged:
        log(f"{prefix}Flagged {len(zero_flagged)} zero-amount line(s)")

    all_flagged = list(zero_flagged)

    date_buckets = group_by_date(rows)
    log(f"{prefix}Found {len(date_buckets)} distinct date(s)")

    all_confirmed  = []
    all_resolution = {}
    global_pool    = list(rows)

    for date, date_rows in date_buckets.items():
        # date_rows is a fixed snapshot taken before any bucket was
        # processed. Some of its own rows may since have been stolen by an
        # earlier-processed date bucket's cross-date stray search — drop
        # those here, or this bucket would try to resolve (and possibly
        # flag) a line that's already confirmed elsewhere.
        global_pool_idxs = {r["_idx"] for r in global_pool}
        date_rows        = [r for r in date_rows if r["_idx"] in global_pool_idxs]
        if not date_rows:
            continue

        bucket_idxs     = {r["_idx"] for r in date_rows}
        cross_date_pool = [r for r in global_pool if r["_idx"] not in bucket_idxs]

        confirmed, flagged, used, resolution_meta = resolve_date_bucket(
            date_rows, cross_date_pool
        )
        all_confirmed.extend(confirmed)
        all_flagged.extend(flagged)
        all_resolution.update(resolution_meta)

        # Once this bucket's own rows are confirmed or flagged, they are
        # spent — whether resolved locally or given up on, they must never
        # be reconsidered as a stray candidate by a later-processed date
        # bucket (which would double-count them: once here, once there).
        used_idxs   = {r["_idx"] for r in used} | bucket_idxs
        global_pool = [r for r in global_pool if r["_idx"] not in used_idxs]

    orphans = final_conflict_check(all_confirmed, all_flagged, rows)
    if orphans:
        log(f"{prefix}Flagged {len(orphans)} orphaned line(s)")
    all_flagged.extend(orphans)

    log(f"{prefix}Confirmed entries : {len(all_confirmed)}")
    log(f"{prefix}Flagged rows      : {len(all_flagged)}")

    log(f"{prefix}Running verification...")
    errors = verify(original_rows, all_confirmed, all_flagged)
    if errors:
        log(f"{prefix}VERIFICATION FAILED:")
        for e in errors:
            log(f"{prefix}  ✗ {e}")
    else:
        log(f"{prefix}  ✓ Line count     : all input lines accounted for")
        log(f"{prefix}  ✓ No duplicates  : no line appears more than once")
        log(f"{prefix}  ✓ Debit totals   : input and output match")
        log(f"{prefix}  ✓ Credit totals  : input and output match")

    all_flagged = assign_priorities(all_flagged)

    return all_confirmed, all_flagged, all_resolution, errors


# ---------------------------------------------------------------------------
# Step 8d — Generate output rows as a format-agnostic event stream.
#
# Each entry point (CSV writer, Excel writer, ...) consumes this generator
# and renders the events into its own format. Events are tuples:
#   ("fields", output_fields)        - the output column order, yielded once
#   ("row", out_row_dict)            - one data row, keyed by output_fields
#   ("blank", None)                  - a blank separator row
#   ("header", text)                 - a section header (e.g. flagged banner)
# ---------------------------------------------------------------------------

def generate_output_rows(confirmed_groups, flagged_rows, fieldnames, resolution_meta):
    output_fields = fieldnames + ["CONFIDENCE_SCORE", "FLAG_REASON"]
    yield ("fields", output_fields)

    confirmed_sorted = sorted(confirmed_groups, key=lambda g: g[0]["_idx"])

    for group in confirmed_sorted:
        meta       = resolution_meta.get(id(group), ("sweep", 0, False))
        resolution, stray_count, stray_crossdate = meta
        score      = compute_confidence(group, resolution, stray_count, stray_crossdate)

        for i, row in enumerate(group):
            out_row = {k: row.get(k, "") for k in output_fields}
            out_row["FLAG_REASON"] = ""
            if i == 0:
                out_row["CONFIDENCE_SCORE"] = score
            else:
                out_row["CONFIDENCE_SCORE"] = ""
                out_row[COL_JOURNAL] = ""
                out_row[COL_DATE]    = ""
            yield ("row", out_row)

        yield ("blank", None)

    if flagged_rows:
        yield ("header", "=== FLAGGED FOR MANUAL REVIEW ===")
        yield ("blank", None)

        for row in flagged_rows:
            out_row = {k: row.get(k, "") for k in output_fields}
            out_row["CONFIDENCE_SCORE"] = 0
            out_row["FLAG_REASON"]      = row.get("_flag_reason", "")
            yield ("row", out_row)
