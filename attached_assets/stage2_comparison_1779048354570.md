# Stage 2 Streaming Formula Cache Patcher Comparison

## Status

`_write_formula_caches()` now patches the `.xlsx` package with a stream-and-pipe flow:

- Copies zip entries one at a time into a temporary workbook package.
- Patches worksheet XML in chunks instead of building a full ElementTree DOM.
- Spools calcChain references to a temporary file and streams `xl/calcChain.xml` into the output zip.
- Drops each sheet's formula-cache entries after that sheet is patched.
- Uses `os.replace()` for the final swap, with a streamed overwrite fallback for local synced Windows folders that deny atomic replacement.

## BCN Profile

The BCN input workbook files were not present in this workspace, so I could not rerun the BCN end-to-end export locally. The baseline below is from the task context.

| Metric | Stage 1 Baseline | Stage 2 Local Result |
| --- | ---: | ---: |
| `patch_formula_caches` RSS delta | 2.76 GiB | BCN not runnable here |
| `patch_formula_caches` tracemalloc | Not provided | BCN not runnable here |
| `write_calc_chain` RSS delta | Not provided | BCN not runnable here |
| `write_calc_chain` tracemalloc | Not provided | BCN not runnable here |
| Total export duration | 429 seconds | BCN not runnable here |
| Total peak RSS | 4.35 GiB | BCN not runnable here |
| Workbook file size | Not provided | BCN not runnable here |

## Local Streaming Regression Profile

Synthetic workbook: 3,000 respondents, streamed `_RawData`.

| Metric | Stage 1 Local | Stage 2 Local |
| --- | ---: | ---: |
| `patch_formula_caches` tracemalloc peak | 33.1 MiB | 12.0 MiB |
| `write_calc_chain` tracemalloc peak | 21.2 MiB | 4.1 MiB |
| Workbook file size | 199,187 bytes | 199,364 bytes |

Latest memory report:

```text
step | rss_start | rss_end | rss_delta | tracemalloc_peak
--- | ---: | ---: | ---: | ---:
load_or_receive_decoded_df | n/a | n/a | n/a | 1.8 MiB
generate_short_labels | n/a | n/a | n/a | 1.9 MiB
build_raw_data_sheet | n/a | n/a | n/a | 11.1 MiB
build_options_sheet | n/a | n/a | n/a | 5.4 MiB
build_filters_sheet | n/a | n/a | n/a | 5.4 MiB
build_helper_columns | n/a | n/a | n/a | 5.4 MiB
build_run_summary_sheet | n/a | n/a | n/a | 5.4 MiB
build_question_metadata_sheet | n/a | n/a | n/a | 5.4 MiB
build_single_cut_index_sheet | n/a | n/a | n/a | 5.5 MiB
build_theme_sheets | n/a | n/a | n/a | 5.5 MiB
build_calculation_log_sheet | n/a | n/a | n/a | 5.6 MiB
build_filter_log_sheet | n/a | n/a | n/a | 5.6 MiB
build_warnings_sheet | n/a | n/a | n/a | 5.6 MiB
save_workbook | n/a | n/a | n/a | 6.4 MiB
patch_formula_caches | n/a | n/a | n/a | 12.0 MiB
write_calc_chain | n/a | n/a | n/a | 4.1 MiB
```

Profiler report used: `outputs/excel_exporter_test_streaming_formula_cache_patch_peak_under_100_mib_0c42bbc43936499fb4e657decc1386f8.memory_report.txt`.

## Verification

- Exporter test suite passed: 102 tests.
- Full test suite passed: 385 tests.
- New regression guard verifies `patch_formula_caches` stays below 100 MiB tracemalloc on a 3,000-row streamed export.
