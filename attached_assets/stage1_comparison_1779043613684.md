# Stage 1 Streaming Raw Data Comparison

## Status

The `_RawData` sheet now routes to the streaming writer for exports above 2,000 rows and below the 50,000-row static safety threshold. The streaming path preserves the full `_RawData` sheet, helper columns, formula cache patching, and named ranges used by theme-sheet `COUNTIFS` formulas.

## BCN Profile

The requested BCN source files were not present in this workspace, so the BCN end-to-end export could not be run locally. A workspace scan found generated test workbooks only, not `BCN_LTB_raw_data...xlsx`, `Rawdata_sample.xlsx`, or `Datamap_sample.xlsx`.

| Metric | Before | After |
| --- | ---: | ---: |
| `build_raw_data_sheet` duration | 100+ seconds from prior profiler note | BCN not runnable in this workspace |
| `build_raw_data_sheet` memory | Not available locally | BCN not runnable in this workspace |
| Total export duration | Not available locally | BCN not runnable in this workspace |
| Peak RSS | Not available locally | `n/a` here because `psutil` is not installed in the bundled runtime |

## Local Streaming Regression Profile

Synthetic workbook: 3,000 respondents, fully streamed `_RawData`.

| Metric | Value |
| --- | ---: |
| Export duration | 5.915 seconds |
| Workbook file size | 199,187 bytes |
| Reloaded `_RawData` rows | 3,001 including header |
| `build_raw_data_sheet` tracemalloc peak | 10.2 MiB |
| `patch_formula_caches` tracemalloc peak | 33.1 MiB |

Profiler report used: `outputs/excel_exporter_test_streaming_raw_data_named_ranges_cover_all_rows_b65900b00617408f8d75178913ac786e.memory_report.txt`.

## Verification

- `respondent_id_data`, `Q_STREAM_data`, and helper named ranges point to rows 2 through 3001.
- Reloading the saved workbook confirms all streamed rows are physically present.
- Full test suite passed: 384 tests.
