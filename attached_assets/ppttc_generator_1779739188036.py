"""think-cell .ppttc generator.

Produces a .ppttc command file that tells think-cell to load the Bain survey
template and populate the chart matching a question's recommended ChartType.
The numbers come verbatim from ThinkCellTablePayload; this module only adapts
the table shape to think-cell's JSON automation format.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
import io
import json
import re
import zipfile

from src.chart_recommender import ChartRecommendation, ChartType
from src.thinkcell_table_formatter import ThinkCellTablePayload


CHART_ANCHOR_MAP: dict[ChartType, str | None] = {
    ChartType.COLUMN_STACKED: "Chart_stacked_column",
    ChartType.COLUMN_CLUSTERED: "Chart_clustered_column",
    ChartType.BAR_STACKED: "Chart_stacked_horizontal",
    ChartType.LINE: "Chart_line",
    ChartType.COMBO: "Chart_clustered_column",
    ChartType.HEATMAP_TABLE: None,
}


@dataclass(frozen=True, slots=True)
class PpttcArtifact:
    """Result of .ppttc generation."""

    filename: str
    content_bytes: bytes
    chart_anchor: str
    warnings: list[str] = field(default_factory=list)


def resolve_chart_anchor(
    recommendation: ChartRecommendation,
    payload: ThinkCellTablePayload,
) -> tuple[str | None, list[str]]:
    """Resolve the template anchor for a chart recommendation."""

    warnings: list[str] = []
    chart_type = recommendation.chart_type
    if chart_type is ChartType.HEATMAP_TABLE:
        return None, []
    if chart_type is ChartType.BAR_CLUSTERED:
        if _count_series(payload) <= 1:
            return "Chart_horizontal_bar", warnings
        return "Chart_clustered_horizontal", warnings
    if chart_type is ChartType.COMBO:
        warnings.append(
            "COMBO chart type not natively in template; using "
            "Chart_clustered_column as fallback."
        )
    return CHART_ANCHOR_MAP.get(chart_type), warnings


def _count_series(payload: ThinkCellTablePayload) -> int:
    """Return the number of data series represented by a payload."""

    if payload.chart_type is ChartType.BAR_STACKED:
        return len(payload.rows)
    return max(1, len(payload.headers))


def build_ppttc(
    recommendation: ChartRecommendation,
    payload: ThinkCellTablePayload,
    template_path: Path,
    question_id: str,
    question_label: str | None = None,
) -> PpttcArtifact | None:
    """Build a .ppttc artifact for a single question chart."""

    anchor, warnings = resolve_chart_anchor(recommendation, payload)
    if anchor is None:
        return None

    if not template_path.exists():
        raise FileNotFoundError(
            f"think-cell template not found at {template_path}. "
            "Copy bain_survey_template.pptx into sample_data/ first."
        )

    rows = _payload_to_tc_rows(payload, recommendation)
    ppttc_dict = _assemble_ppttc(template_path, anchor, rows)
    content_bytes = json.dumps(ppttc_dict, indent=2).encode("utf-8")

    safe_label = _safe_filename_segment(question_label or question_id)
    filename = f"{_safe_filename_segment(question_id)}_{safe_label}.ppttc"

    return PpttcArtifact(
        filename=filename,
        content_bytes=content_bytes,
        chart_anchor=anchor,
        warnings=warnings,
    )


def _payload_to_tc_rows(
    payload: ThinkCellTablePayload,
    recommendation: ChartRecommendation,
) -> list[list[Any]]:
    """Convert a Round 1 table payload to think-cell rows-of-rows."""

    if not payload.headers:
        raise ValueError("payload.headers must be non-empty")
    for row in payload.rows:
        if len(row) != len(payload.headers) + 1:
            raise ValueError(
                "each payload row must contain a label plus one value per header"
            )
    return [[None] + list(payload.headers)] + [list(row) for row in payload.rows]


def _assemble_ppttc(
    template_path: Path,
    chart_anchor: str,
    rows: list[list[Any]],
) -> list[dict[str, Any]]:
    """Build the .ppttc JSON shape think-cell accepts.

    Format is a top-level list with sibling `template` and `data` entries
    rather than a nested object. The template entry uses only the filename so
    the .ppttc is portable when bundled beside the template.
    """

    return [
        {
            "template": template_path.name,
        },
        {
            "data": [
                {
                    "name": chart_anchor,
                    "table": [[_transform_cell(cell) for cell in row] for row in rows],
                }
            ],
        }
    ]


def _transform_cell(value: Any) -> dict[str, Any]:
    """Convert a Python value to a think-cell cell dict.

    Blank cells must be encoded as empty objects. Some think-cell chart types
    reject JSON null in the table payload.
    """

    if value is None:
        return {}
    if isinstance(value, bool):
        return {"string": str(value)}
    if isinstance(value, (int, float)):
        return {"number": value}
    return {"string": str(value)}


def build_ppttc_bundle(
    artifact: PpttcArtifact,
    template_path: Path,
) -> tuple[bytes, str]:
    """Bundle a .ppttc and its template into a single zip for download."""

    if not template_path.exists():
        raise FileNotFoundError(f"think-cell template not found at {template_path}.")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(artifact.filename, artifact.content_bytes)
        archive.write(template_path, arcname=template_path.name)

    zip_filename = artifact.filename.replace(".ppttc", "_thinkcell.zip")
    return buffer.getvalue(), zip_filename


def _safe_filename_segment(text: str) -> str:
    """Strip filesystem-unsafe characters and truncate to 40 characters."""

    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text).strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("._-")
    return (cleaned or "chart")[:40]
