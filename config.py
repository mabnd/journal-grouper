"""
config.py
=========
All user-configurable settings for the journal-entry grouper. This is the
only file you need to edit to adapt the tool to a different organization's
data format, language, or target import system.
"""

# ---------------------------------------------------------------------------
# Input column roles
# Map each role to the column header as it appears in your input file.
# The four required roles must be present in every processed file/sheet;
# the optional ones are handled gracefully if absent.
# ---------------------------------------------------------------------------

INPUT_COLUMNS = {
    "date":          "Date de facturation",   # required: the transaction date
    "communication": "Communication",          # required: narrative/reference used for matching
    "debit":         "Débit",                  # required
    "credit":        "Crédit",                 # required
    "journal":       "Journal",                # optional: falls back to sheet name if absent
    "code":          "Code",                   # optional: account code, passthrough
    "partner":       "Partenaire",             # optional: only needed for client-matching check
}

# ---------------------------------------------------------------------------
# Output column names
# Controls what each column is called in the output file.
# Key: the input column name (from INPUT_COLUMNS values, or any extra column
#      present in your source file that you want to rename).
# Value: the desired name in the output.
# Columns not listed here keep their original name unchanged.
# ---------------------------------------------------------------------------

OUTPUT_COLUMN_NAMES = {
    "Journal":             "Journal",
    "Date de facturation": "Date",
    "Communication":       "Ecritures comptables / Libellés",
    "Débit":               "Ecritures comptables /Debit",
    "Crédit":              "Ecritures comptables /Credit",
    "Code":                "Ecritures comptables / Compte",
    "Partenaire":          "Ecritures comptables / partenaire",
}

# Auto-generated reference column inserted after the date column in the output.
REFERENCE_COLUMN_NAME = "Référence"
REFERENCE_VALUE       = "IMPORT MOUVEMENT"

# ---------------------------------------------------------------------------
# Client matching (optional partner check)
# ---------------------------------------------------------------------------

CLIENT_NAME_COLUMN = "Nom complet"   # column header in your clients CSV file

# ---------------------------------------------------------------------------
# Missing-client report
# Column names and status text for the partner-issue output.
# ---------------------------------------------------------------------------

MISSING_CLIENT_SHEET_NAME      = "Clients manquants"       # Excel sheet title
MISSING_CLIENT_COL_SHEET       = "Feuille"                 # which sheet the partner came from (Excel only)
MISSING_CLIENT_COL_PARTNER     = "Partenaire"              # the unmatched partner name
MISSING_CLIENT_COL_OCCURRENCES = "Occurrences"             # how many times it appeared
MISSING_CLIENT_COL_STATUS      = "Statut"                  # typo or unknown
MISSING_CLIENT_COL_SUGGESTION  = "Client suggéré"          # closest known client name
MISSING_CLIENT_STATUS_TYPO     = "Probable faute de frappe — vérifiez avec le client suggéré"
MISSING_CLIENT_STATUS_UNKNOWN  = "Client inconnu — à créer ou à corriger"

# ---------------------------------------------------------------------------
# Flag reasons
# Text that appears verbatim in the FLAG_REASON output column for flagged rows.
# ---------------------------------------------------------------------------

REASON_ZERO_AMOUNT     = "Montant manquant — cette ligne n'a ni débit ni crédit. Ajoutez le montant manquant dans le fichier source."
REASON_INCOMPLETE      = "Ligne sans correspondance — aucune autre ligne n'a pu être trouvée pour que les débits et crédits s'équilibrent. Vérifiez si une ligne est manquante ou si un montant est incorrect."
REASON_AMBIGUOUS_STRAY = "Correspondance incertaine — plusieurs lignes pourraient s'associer à celle-ci, mais aucune ne s'impose clairement. Choisissez manuellement la bonne ligne à associer."
REASON_AMBIGUOUS_SPLIT = "Regroupement impossible — les lignes à cette date peuvent être combinées de plusieurs façons différentes sans qu'une solution soit évidente. Vérifiez manuellement quelles lignes vont ensemble."
REASON_ORPHAN          = "Ligne orpheline — cette ligne n'a pu être rattachée à aucune écriture. Vérifiez si elle appartient à une écriture existante ou si une ligne de contrepartie est manquante."

# ---------------------------------------------------------------------------
# Algorithm tuning knobs
# These rarely need changing unless your data has unusual characteristics.
# ---------------------------------------------------------------------------

# Similarity threshold (0–1) for keeping a balanced group open or clustering
# communications. Higher = stricter matching, more groups; lower = looser.
FUZZY_MERGE_THRESHOLD   = 0.95

# Stray-line hunting: the top candidate's similarity score must exceed the
# second candidate's by at least this margin to be accepted automatically;
# otherwise the group is flagged for manual review.
STRAY_CONFIDENCE_MARGIN = 0.20

# Floating-point rounding tolerance for all balance checks.
BALANCE_TOLERANCE       = 0.01

# Subset-sum DP state limit — bounds memory usage, not line count.
MAX_SUBSET_SUM_STATES   = 200_000

# Recursive split size and depth guards.
MAX_SPLIT_SIZE          = 20
MAX_SPLIT_DEPTH         = 10

# Similarity threshold for classifying an unmatched partner name as a probable
# typo vs. a completely unknown client.
PARTNER_TYPO_THRESHOLD  = 0.85

# ---------------------------------------------------------------------------
# I/O settings
# ---------------------------------------------------------------------------

CSV_DELIMITER = ","
CSV_ENCODING  = "utf-8-sig"   # BOM-aware; handles files exported from Excel on Windows

# How native Excel date cells are converted to text before date parsing.
# Must produce a format that normalize_date() in journal_grouper_core.py accepts.
# Default "ddmmyy" produces e.g. "020126" for 2026-01-02.
DATE_FORMAT = "%d%m%y"
