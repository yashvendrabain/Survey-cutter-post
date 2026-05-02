"""Survey Insight Engine — Streamlit shell (Day 1).

This file provides the application chrome and file upload entry points.
No parsing, decoding, classification, or computation is performed here today.
"""

import streamlit as st

from config import (
    ACCEPTED_DATAMAP_EXTENSIONS,
    ACCEPTED_RAWDATA_EXTENSIONS,
    APP_NAME,
    MAX_UPLOAD_SIZE_MB,
    VERSION,
)


def _strip_dot(extensions):
    """Convert ('.csv', '.xlsx') -> ['csv', 'xlsx'] for st.file_uploader."""
    return [ext.lstrip(".") for ext in extensions]


def _format_size_kb(num_bytes: int) -> str:
    return f"{num_bytes / 1024:.1f}"


def _init_session_state() -> None:
    defaults = {
        "data_map": None,
        "survey_schema": None,
        "decoded_df": None,
        "quality_report": None,
        "single_cut_results": None,
        "cross_cut_results": None,
        "audit_log": None,
        "run_timestamp": None,
        "raw_data_file": None,
        "datamap_file": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def main() -> None:
    st.set_page_config(layout="wide", page_title=APP_NAME)

    _init_session_state()

    st.title(APP_NAME)
    st.caption(f"Version {VERSION}")

    with st.sidebar:
        st.header("Settings")
        st.caption("Configuration options will appear here as features are added")

        st.header("About")
        st.write(
            "Ingests a raw survey data file and a paired data map, "
            "then produces classified, audited analytics."
        )

    left, right = st.columns(2)

    with left:
        st.subheader("Raw data")
        raw_upload = st.file_uploader(
            "Upload raw data file",
            type=_strip_dot(ACCEPTED_RAWDATA_EXTENSIONS),
            key="raw_data_uploader",
            help=f"Maximum upload size: {MAX_UPLOAD_SIZE_MB} MB.",
        )
        if raw_upload is not None:
            st.session_state["raw_data_file"] = raw_upload

    with right:
        st.subheader("Data map")
        datamap_upload = st.file_uploader(
            "Upload data map file",
            type=_strip_dot(ACCEPTED_DATAMAP_EXTENSIONS),
            key="datamap_uploader",
            help="Data map must be an .xlsx file.",
        )
        if datamap_upload is not None:
            st.session_state["datamap_file"] = datamap_upload

    raw_file = st.session_state.get("raw_data_file")
    if raw_file is not None:
        st.text(f"Raw data: {raw_file.name} ({_format_size_kb(raw_file.size)} KB)")
    else:
        st.text("Raw data: not uploaded")

    datamap_file = st.session_state.get("datamap_file")
    if datamap_file is not None:
        st.text(
            f"Data map: {datamap_file.name} ({_format_size_kb(datamap_file.size)} KB)"
        )
    else:
        st.text("Data map: not uploaded")

    with st.expander("Expected file format"):
        st.write(
            "The data map should be an xlsx file with a 'Sheet1' tab containing "
            "question blocks separated by blank rows. Each block has a header line, "
            "a Values: or 'Open response' line, and one row per option."
        )
        st.write(
            "The raw data should be a csv or xlsx with one row per respondent, "
            "where column names match the question IDs declared in the data map."
        )


if __name__ == "__main__":
    main()
