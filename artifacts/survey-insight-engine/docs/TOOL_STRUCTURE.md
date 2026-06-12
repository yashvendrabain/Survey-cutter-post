# Survey Insight Engine — Tool Structure

A Streamlit application that ingests a raw survey data file plus a paired data map,
parses and classifies every question, computes statistics with a full audit trail,
and exports a live-filterable Excel workbook. Built in Python inside a pnpm monorepo
and served as a stateful web artifact.

- **Entry point:** `app.py`
- **Served at:** `/` (port `21049`), workflow `artifacts/survey-insight-engine: web`
- **Runtime:** Streamlit 1.58, pandas 2.2, openpyxl/xlsxwriter, plotly, scipy, openai
- **Artifact kind:** `web` (stateful — runs as a long-lived websocket app)

---

## Top-level layout

```
survey-insight-engine/
├── app.py                       # Streamlit entry point — UI, session state, page flow
├── config.py                    # Tunables: missing-value tokens, thresholds, limits
├── requirements.txt             # Python dependencies (pinned)
├── README.md                    # Project overview
├── advanced_segmentation_ui.py  # (root copy) advanced segmentation dashboard
├── winner_scoring.py            # (root copy) top-group vs rest cohort scoring
├── .streamlit/config.toml       # Streamlit server/theme config
├── .replit-artifact/artifact.toml  # Artifact + service/routing definition
│
├── src/                         # Core engine (parsing → classify → compute → export)
├── tests/                       # 64 pytest modules + fixtures
├── docs/                        # PRD.md, this file
├── sample_data/                 # Example survey workbooks + PPTX template
├── public/                      # opengraph.jpg
├── outputs/                     # Scratch run artifacts (logs, pickles, generated xlsx)
│
├── diagnose_q1_labels.py        # Diagnostic: trace option labels through the pipeline
├── io_diag.py                   # Diagnostic: exercise the real io load path
└── verify_5s5t.py               # Verification helper script
```

---

## `src/` — core engine

The pipeline runs in stages: **load → parse data map → decode raw data →
classify questions → compute → export**.

### Input / loading
| Module | Responsibility |
|---|---|
| `io.py` | App-facing loader (`load_survey_inputs`) — accepts uploaded file(s), routes to the parser + decoder, returns the data map, decoded dataframe, and a load report. |
| `datamap_parser.py` | Parses the data-map sheet into `ParsedQuestion`/`DataMap` (question ids, value ranges, options, sub-columns, numeric-label metadata). |
| `raw_decoder.py` | Loads the raw-data sheet, strips/normalises values, coerces to numeric, decodes option columns. Picks the correct sheet by matching expected columns. |

### Classification
| Module | Responsibility |
|---|---|
| `question_classifier.py` | Assigns each question a `QuestionType` (single/multi-select, grid, NPS, rank, numeric, open text) and builds the option/grid label maps the calculators read. Includes `reconcile_multiselect_value_subtypes`. |
| `survey_type_detector.py` | Detects overall survey shape / type to drive defaults. |

### Adapters (`src/adapters/`)
Pluggable handlers for non-standard data-map layouts, selected via a registry.
| Module | Layout it handles |
|---|---|
| `registry.py` / `base.py` | Adapter lookup + shared base class. |
| `label_pattern_subcolumn.py` | Sub-column labels carried as patterns. |
| `grid_categorical_row.py` | Categorical grid rows. |
| `grid_rated_double_colon.py` | Rated grids using `::` notation. |
| `compact_two_column.py` | Compact two-column data maps. |
| `six_column_combined.py` | Six-column combined layout. |
| `bcn_multicolumn.py` | BCN-style multi-column exports. |
| `wizard_configured.py` | Layout described via the in-app wizard. |

### Computation (`src/single_cut/`)
| Module | Responsibility |
|---|---|
| `engine.py` | Orchestrates single-cut computation across all questions. |
| `_single_select.py` | Single-select frequencies. |
| `_multi_select.py` | Multi-select (binary option) counts/percentages. |
| `_numeric.py` | Numeric summaries. |
| `_grid.py`, `grid_rated.py`, `grid_binary_pivot.py` | Grid question variants. |
| `nps.py` | Net Promoter Score buckets/score. |
| `rank_order.py` | Rank-order metrics (weighted avg, sum of ranks, position counts). |
| `_conditional.py` | Conditional/derived computations. |

### Cross-cuts, filters & segmentation
| Module | Responsibility |
|---|---|
| `cross_cut_engine.py` | Computes a question broken down by another (cross-tabs). |
| `cross_cut_suggestions.py` | Scores/suggests useful cross-cuts for an outcome. |
| `global_filter.py` | Workbook-wide respondent filter. |
| `filter_options.py` | Builds selectable filter specs from question options. |
| `filtered_single_cut.py` | Single cut recomputed under an active filter. |
| `outcome_segmentation.py` | Outcome-based segmentation result model + math. |
| `winner_scoring.py` (`src/`) | "Top group vs Rest" cohort scoring used by the advanced view. |
| `advanced_segmentation_ui.py` (`src/`) | Advanced segmentation dashboard rendering. |
| `hypothesis_validator.py` | Hypothesis-check computations. |

### Output / export
| Module | Responsibility |
|---|---|
| `excel_exporter.py` | The large export engine — builds the live-filterable workbook (raw data, options, filters, per-theme sheets, NPS, run summary, calc log, warnings, embedded inputs). |
| `thinkcell_table_formatter.py` | Formats tables into Think-cell-compatible payloads. |
| `chart_recommender.py` | Picks an appropriate chart type per question. |
| `chart_renderer.py` | Renders recommendations as Plotly figures (Bain styling). |
| `bain_palette.py` | Brand colour palette + series colours. |
| `ppttc_generator.py` | PowerPoint / Think-cell generation. |
| `word_survey_parser.py` | Parses Word-format survey questionnaires. |

### AI & assistant
| Module | Responsibility |
|---|---|
| `ai_insights.py` | Generates narrative insights, short labels, table/outlier commentary. |
| `assistant_bot.py` / `chat_panel.py` | In-app assistant + chat UI panel. |
| `product_tour.py` | Guided product tour. |

### Support / shared
| Module | Responsibility |
|---|---|
| `models.py` | Dataclasses/enums for the whole pipeline (`SurveySchema`, `QuestionType`, all `*Result` types, reports). |
| `calc_primitives.py` | Low-level calculation helpers. |
| `calculation_log.py` | Audit trail of every computed value. |
| `memory_profiler.py` | Optional memory profiling around export steps. |
| `ui_constants.py` | UI strings, section labels, tooltips, pipeline stages. |
| `ui/wizard.py` | Multi-step data-map configuration wizard. |

---

## Data flow

```
uploaded file(s)
   │
   ▼  src/io.load_survey_inputs
parse data map ──────► src/datamap_parser  (+ src/adapters/* for odd layouts)
   │
   ▼
decode raw data ─────► src/raw_decoder
   │
   ▼
classify questions ──► src/question_classifier  → SurveySchema
   │
   ▼
compute ─────────────► src/single_cut/*  (+ cross_cut_engine, outcome_segmentation,
   │                                        winner_scoring, hypothesis_validator)
   ▼
render in app  ──────► chart_recommender → chart_renderer (Plotly)
   │
   ▼
export ──────────────► src/excel_exporter  → live-filterable .xlsx
                       (+ ppttc_generator for PPTX/Think-cell)
```

Every computed value is recorded in `calculation_log.py` so outputs are auditable.

---

## Tests

`tests/` holds **64** pytest modules covering parsers, adapters, classification,
each calculator (single/multi-select, grid, NPS, rank, numeric), cross-cuts,
filters, segmentation, hypothesis checks, AI insights, chart rendering, and the
Excel exporter (with dedicated suites for grid-rated, rank-order, hypothesis,
input-embed, manual-cohort, and winners-vs-laggards formulas). Shared fixtures
live in `tests/fixtures/` and `tests/conftest.py`.

---

## Configuration & serving

- **`config.py`** — missing-value tokens, high-missing threshold, raw-data row
  limits, cross-tab max groups.
- **`.streamlit/config.toml`** — server flags + theme (suppresses the usage-stats line).
- **`.replit-artifact/artifact.toml`** — declares the `web` service on port `21049`
  routed at `/`, with identical dev/production Streamlit run commands.

## Diagnostics (root)

- **`diagnose_q1_labels.py`** — traces a question's option labels through parse →
  adapter → classification → live compute to localise label loss.
- **`io_diag.py`** — runs the real `src.io.load_survey_inputs` path (what the app
  actually uses) and dumps the resulting option maps and selections.
- **`verify_5s5t.py`** — verification helper.
