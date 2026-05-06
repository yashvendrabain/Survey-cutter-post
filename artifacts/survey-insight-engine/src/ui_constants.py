"""UI display constants for the Survey Analysis Engine."""

APP_TITLE = "Survey Analysis Engine"
APP_TAGLINE = (
    "Upload a survey, explore the data, "
    "export consultant-ready workbooks."
)

# Section labels — keep these consistent across the app
SECTION_UPLOAD = "1. Upload your survey"
SECTION_GLOBAL_FILTER = "2. Apply a global filter (optional)"
SECTION_RESULTS = "3. Explore single cuts"
SECTION_CROSS_CUTS = "4. Build cross cuts"
SECTION_DOWNLOADS = "5. Download workbooks"

# Tooltips — written for a consultant who hasn't seen the
# tool before. Keep them under 200 chars.
TOOLTIP_GLOBAL_FILTER = (
    "Restrict every analysis below to a specific subset of "
    "respondents. For example, set Region = APAC to view the "
    "entire survey through an APAC-only lens."
)
TOOLTIP_PER_QUESTION_FILTER = (
    "Filter just this single question. Add multiple filters to "
    "narrow down further (e.g. Region = APAC AND Industry = "
    "Tech AND Role = CTO)."
)
TOOLTIP_BREAKDOWN = (
    "Choose 'All values (breakdown)' on a filter to see the "
    "question split across that dimension automatically."
)
TOOLTIP_CROSS_CUT_SUGGESTIONS = (
    "The tool inspects your survey and suggests cross cuts "
    "likely to be useful. Tick the box to see them."
)
TOOLTIP_THREE_DOWNLOADS = (
    "Three workbooks are produced: a full single-cut workbook, "
    "a cross-cut-only workbook with what you've selected, and "
    "a filtered workbook with any per-question filters you've "
    "applied."
)

# Status text for the run pipeline
PIPELINE_STAGES = (
    "Reading data map…",
    "Decoding raw data…",
    "Classifying questions…",
    "Computing single cuts…",
    "Building Excel workbook…",
)

# Empty state messages
EMPTY_NO_RESULTS = (
    "No results yet. Upload your data files and click "
    "'Run analysis' to begin."
)
EMPTY_NO_CROSS_CUTS = (
    "No cross cuts yet. Build one using the dropdowns above, "
    "or tick the 'Show suggested cross cuts' checkbox."
)
EMPTY_NO_FILTERS_APPLIED = (
    "No filters applied. Use the per-question filter panel to "
    "slice an individual question, or set a global filter to "
    "restrict every analysis."
)

# Status badges
STATUS_GLOBAL_FILTER_ACTIVE = "🔵 Global filter active"
STATUS_GLOBAL_FILTER_INACTIVE = "⚪ No global filter"
