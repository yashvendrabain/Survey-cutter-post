"""Configuration constants for the Survey Analysis Engine."""

APP_NAME = "Survey Analysis Engine"
VERSION = "Stage 2"

DATAMAP_SHEET_NAME = "Sheet1"
QUESTION_HEADER_PATTERN = r"^\[?([A-Za-z][A-Za-z0-9_]*)\]?:\s*(.+)$"
VALUES_LINE_PATTERN = r"^Values:\s*(-?\d+)\s*-\s*(-?\d+)$"
OPEN_NUMERIC_LINE = "Open numeric response"
OPEN_TEXT_LINE = "Open text response"
SUB_COLUMN_PATTERN = r"^[A-Za-z][A-Za-z0-9_]*$"

MISSING_VALUE_TOKENS = {"", "NA", "N/A", "NULL", "null", "None", "nan"}
HIGH_MISSING_THRESHOLD = 0.5

DEFAULT_ALLOCATION_TARGET = 100.0
ALLOCATION_TOLERANCE = 2.0
LOW_SAMPLE_THRESHOLD = 30

# ---------- AI insight layer (Stage A) ----------
PORTKEY_BASE_URL = "https://portkey.bain.dev/v1"
PORTKEY_DEFAULT_MODEL = "@personal-openai/gpt-4o-mini"
PORTKEY_PREMIUM_MODEL = "@personal-openai/gpt-4o"
AI_INSIGHT_TEMPERATURE = 0.1
AI_INSIGHT_MAX_TOKENS = 350
AI_INSIGHT_TIMEOUT_SECONDS = 30

# Read API key from environment ONLY. Never hardcode.
# If the key is missing, the layer falls back to template output.
