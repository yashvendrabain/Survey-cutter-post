"""One-shot patcher: wires the example-driven wizard (app.py) AND the single-select
raw_id decode fix (src/raw_decoder.py). Self-verifying; backs up before writing.

Run from the project root:  python patch_survey_wizard.py
"""

import re, ast, shutil, sys, os

APP = "app.py"
DECODER = os.path.join("src", "raw_decoder.py")

# ---------------------------------------------------------------- app.py
NEW_STEP3 = '''def _render_format_wizard_step_3(config: dict[str, Any]) -> None:
    app = _require_streamlit()
    from src.adapters.wizard_configured import infer_question_id_regex

    example = app.text_input(
        "Paste ONE example of a question ID, exactly as it appears in the data "
        "map / codebook sheet (its left column, e.g. q1, 1q, q-1, Q15, Item_3). "
        "The wizard infers the pattern from it.",
        value=str(config.get("question_id_example") or ""),
        placeholder="q1",
        key="wizard_question_id_example",
    )
    config["question_id_example"] = example.strip()
    if config["question_id_example"]:
        try:
            config["question_id_pattern"] = infer_question_id_regex(
                config["question_id_example"]
            ).pattern
            app.caption(f"Inferred question-ID pattern: {config['question_id_pattern']}")
        except ValueError:
            config["question_id_pattern"] = ""
            app.warning(
                "That example has no number in it. Include the question number, "
                "e.g. 'q1' or 'Item_3'."
            )
    else:
        config["question_id_pattern"] = ""'''

NEW_STEP4 = """def _render_format_wizard_step_4(config: dict[str, Any]) -> None:
    app = _require_streamlit()
    from src.adapters.wizard_configured import infer_multi_select_spec

    example = app.text_input(
        "Paste ONE example of a multi-select / grid option COLUMN, exactly as it "
        "appears in the raw data sheet header (e.g. 6q1a, Q6r1, q6_1, or "
        "'Q6: Field sales'). Leave blank only if this survey has no multi-part "
        "questions.",
        value=str(config.get("multi_select_example") or ""),
        placeholder="6q1a",
        key="wizard_multi_select_example",
    )
    config["multi_select_example"] = example.strip()
    config["sub_column_separator"] = "none"
    if config["multi_select_example"]:
        try:
            mode, _multi, _single = infer_multi_select_spec(
                config["multi_select_example"]
            )
            app.caption(f"Inferred multi-select mode: {mode}")
        except ValueError:
            app.warning(
                "That example has no number in it. Include the question number, "
                "e.g. '6q1a' or 'Q6: Field sales'."
            )"""

OLD_PREVIEW = """        matched_columns.update(
            question["canonical_id"]
            for question in questions
            if question["canonical_id"] in set(str(column) for column in raw_df.columns)
        )"""
NEW_PREVIEW = """        raw_column_set = {str(column) for column in raw_df.columns}
        matched_columns.update(
            str(question.get("raw_id") or question["canonical_id"])
            for question in questions
            if str(question.get("raw_id") or question["canonical_id"]) in raw_column_set
        )"""


def _sub_func(text, name, next_name, replacement):
    pat = re.compile(
        r"(?sm)^def " + re.escape(name) + r"\(config: dict\[str, Any\]\) -> None:.*?"
        r"(?=^def " + re.escape(next_name) + r"\b)"
    )
    new, n = pat.subn(replacement.rstrip("\n") + "\n\n\n", text)
    if n != 1:
        raise SystemExit(f"ABORT: expected 1 match for {name}, got {n}")
    return new


def patch_app():
    src = open(APP, "r").read()
    if "wizard_question_id_example" in src:
        print("app.py already patched; skipping.")
        return
    src = _sub_func(
        src, "_render_format_wizard_step_3", "_render_format_wizard_step_4", NEW_STEP3
    )
    src = _sub_func(
        src, "_render_format_wizard_step_4", "_render_format_wizard_step_5", NEW_STEP4
    )
    if src.count(OLD_PREVIEW) != 1:
        raise SystemExit(
            f"ABORT: preview block found {src.count(OLD_PREVIEW)} times (want 1)"
        )
    src = src.replace(OLD_PREVIEW, NEW_PREVIEW, 1)
    ast.parse(src)
    shutil.copy(APP, APP + ".bak")
    open(APP, "w").write(src)
    print("app.py PATCHED (backup app.py.bak)")


# ---------------------------------------------------------------- raw_decoder.py
HELPER = '''def _effective_raw_id(question: ParsedQuestion) -> str:
    """Raw-data column id for a question with no sub-columns.

    Defaults to canonical_id, but the wizard sets raw_id to the actual raw
    column when codebook id and raw column id differ (e.g. "q1" vs "1q").
    """
    return str(question.get("raw_id") or question["canonical_id"])


'''


def patch_decoder():
    src = open(DECODER, "r").read()
    if "_effective_raw_id" in src:
        print("raw_decoder.py already patched; skipping.")
        return
    anchor = "def _question_expected_columns("
    if src.count(anchor) != 1:
        raise SystemExit(f"ABORT: anchor {anchor!r} found {src.count(anchor)} times")
    src = src.replace(anchor, HELPER + anchor, 1)

    old1 = '    return (question["canonical_id"],)'
    new1 = "    return (_effective_raw_id(question),)"
    if src.count(old1) != 1:
        raise SystemExit(f"ABORT: single-select return found {src.count(old1)} times")
    src = src.replace(old1, new1, 1)

    old2 = '        expected.add(question["canonical_id"])'
    new2 = "        expected.add(_effective_raw_id(question))"
    if src.count(old2) != 1:
        raise SystemExit(f"ABORT: expected.add line found {src.count(old2)} times")
    src = src.replace(old2, new2, 1)

    ast.parse(src)
    shutil.copy(DECODER, DECODER + ".bak")
    open(DECODER, "w").write(src)
    print("src/raw_decoder.py PATCHED (backup src/raw_decoder.py.bak)")


if __name__ == "__main__":
    patch_app()
    patch_decoder()
    print(
        "\nDONE. Restart Streamlit, then in the wizard type:  step 3 -> q1   step 4 -> 6q1a"
    )
