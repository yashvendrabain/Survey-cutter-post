"""Runtime constants and feature thresholds for the Survey Insight Engine.

These are configuration values only — no computation logic, parsing, or
business rules belong in this module.
"""

APP_NAME = "Survey Insight Engine"
VERSION = "0.1.0-day1"

# Strings treated as missing values during decode
MISSING_VALUE_TOKENS = frozenset({
    "", "NA", "N/A", "n/a", "Don't know", "DK",
    "-1", "99", "999", "9999"
})

# Datamap structure (verified against the real sample file)
DATAMAP_SHEET_NAME = "Sheet1"
DATAMAP_INDEX_SHEET = "Index"   # ignored by parser

# Patterns observed in the real datamap
QUESTION_HEADER_PATTERN = r"^\[?([A-Za-z][A-Za-z0-9_]*)\]?:\s*(.+)$"
VALUES_LINE_PATTERN = r"^Values:\s*(-?\d+)\s*-\s*(-?\d+)$"
OPEN_NUMERIC_LINE = "Open numeric response"
OPEN_TEXT_LINE = "Open text response"

# Sub-column pattern for multi-select and grid:
#   group 1 = parent question id (e.g. "Q53")
#   group 2 = sub-index (e.g. "1")
#   group 3 = optional "oe" suffix marking open-text follow-up
SUB_COLUMN_PATTERN = r"^[A-Za-z][A-Za-z0-9_]*$"

# Default allocation parameters (numeric_allocation questions)
DEFAULT_ALLOCATION_TARGET = 100.0
ALLOCATION_TOLERANCE = 2.0

# Quality thresholds
LOW_SAMPLE_THRESHOLD = 30
HIGH_MISSING_THRESHOLD = 0.20

# File constraints
MAX_UPLOAD_SIZE_MB = 200
ACCEPTED_RAWDATA_EXTENSIONS = (".csv", ".xlsx")
ACCEPTED_DATAMAP_EXTENSIONS = (".xlsx",)
  # CSV cannot represent the multi-sheet structure we need.

# Memory safety valve for the Excel exporter. When the decoded dataframe has
# more rows than this threshold, the exporter switches to static-values mode:
# _RawData is collapsed to a single placeholder row, and every formula cell
# (COUNTIFS / helper masks / wrapped lookups) is written as its pre-computed
# cached value instead of a live formula. The workbook still shows correct
# numbers but loses live filter interactivity. Configure as needed.
RAW_DATA_SHEET_ROW_LIMIT = 50000
CROSS_TAB_MAX_GROUPS = 12

# ---------- AI insight layer (Stage A) ----------
PORTKEY_BASE_URL = "https://portkey.bain.dev/v1"
PORTKEY_DEFAULT_MODEL = "@personal-openai/gpt-4o-mini"
PORTKEY_PREMIUM_MODEL = "@personal-openai/gpt-4o"
AI_INSIGHT_TEMPERATURE = 0.1
AI_INSIGHT_MAX_TOKENS = 350
AI_INSIGHT_TIMEOUT_SECONDS = 30

# Read API key from environment ONLY. Never hardcode.
# If the key is missing, the layer falls back to template output.
