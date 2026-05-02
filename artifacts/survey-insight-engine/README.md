# Survey Insight Engine

The Survey Insight Engine ingests a raw survey data file and a paired data map, parses them, classifies questions, and computes statistics with a full audit trail. It is designed to turn opaque survey exports into auditable, well-typed analytics outputs without hand-rolled spreadsheet work.

## Status

Day 1 — foundation only. The project scaffold, dependencies, and a Streamlit shell with file upload widgets are in place. No parsing, decoding, classification, or computation logic has been implemented yet.

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```
