"""Locked Bain color palette and chart color sequencing helpers."""

from __future__ import annotations

from types import MappingProxyType


BAIN_PALETTE = MappingProxyType(
    {
        "RED": "#CC0000",
        "BLACK": "#000000",
        "WHITE": "#FFFFFF",
        "GRAPHITE_1": "#333333",
        "GRAPHITE_2": "#5C5C5C",
        "GRAPHITE_3": "#858585",
        "GRAPHITE_4": "#B4B4B4",
        "GRAPHITE_5": "#D6D6D6",
        "SKY_1": "#2D475A",
        "SKY_2": "#46647B",
        "SKY_3": "#7891AA",
        "SKY_4": "#A3BCD3",
        "SKY_5": "#DCE5EA",
        "FOREST_1": "#104C3E",
        "FOREST_2": "#507867",
        "FOREST_3": "#83AC9A",
        "FOREST_4": "#BBCABA",
        "FOREST_5": "#DCE2D6",
        "SUNSET_1": "#AB8933",
        "SUNSET_2": "#C6AA3D",
        "SUNSET_3": "#E9CD49",
        "SUNSET_4": "#F2DE8A",
        "SUNSET_5": "#FAEEC3",
    }
)


_NO_HERO_SEQUENCES: dict[int, list[str]] = {
    1: [BAIN_PALETTE["GRAPHITE_1"]],
    2: [BAIN_PALETTE["GRAPHITE_1"], BAIN_PALETTE["GRAPHITE_3"]],
    3: [
        BAIN_PALETTE["GRAPHITE_1"],
        BAIN_PALETTE["GRAPHITE_3"],
        BAIN_PALETTE["SKY_2"],
    ],
    4: [
        BAIN_PALETTE["GRAPHITE_1"],
        BAIN_PALETTE["GRAPHITE_3"],
        BAIN_PALETTE["SKY_2"],
        BAIN_PALETTE["FOREST_2"],
    ],
    5: [
        BAIN_PALETTE["GRAPHITE_1"],
        BAIN_PALETTE["GRAPHITE_3"],
        BAIN_PALETTE["SKY_2"],
        BAIN_PALETTE["FOREST_2"],
        BAIN_PALETTE["SUNSET_2"],
    ],
    6: [
        BAIN_PALETTE["GRAPHITE_1"],
        BAIN_PALETTE["GRAPHITE_3"],
        BAIN_PALETTE["GRAPHITE_4"],
        BAIN_PALETTE["SKY_2"],
        BAIN_PALETTE["FOREST_2"],
        BAIN_PALETTE["SUNSET_2"],
    ],
}


_LONG_SERIES_SEQUENCE = [
    BAIN_PALETTE["GRAPHITE_1"],
    BAIN_PALETTE["GRAPHITE_2"],
    BAIN_PALETTE["GRAPHITE_3"],
    BAIN_PALETTE["GRAPHITE_4"],
    BAIN_PALETTE["GRAPHITE_5"],
    BAIN_PALETTE["SKY_2"],
    BAIN_PALETTE["SKY_3"],
    BAIN_PALETTE["SKY_4"],
    BAIN_PALETTE["FOREST_2"],
    BAIN_PALETTE["FOREST_3"],
    BAIN_PALETTE["FOREST_4"],
    BAIN_PALETTE["SUNSET_2"],
    BAIN_PALETTE["SUNSET_3"],
    BAIN_PALETTE["SUNSET_4"],
]


_HERO_RECEDING_SEQUENCE = [
    BAIN_PALETTE["GRAPHITE_3"],
    BAIN_PALETTE["GRAPHITE_4"],
    BAIN_PALETTE["GRAPHITE_5"],
    BAIN_PALETTE["GRAPHITE_2"],
]


def get_series_palette(n_series: int, hero_index: int | None = None) -> list[str]:
    """Return n_series color hex codes using the locked Bain rules."""

    if n_series <= 0:
        return []
    if hero_index is not None:
        colors: list[str] = []
        receding_index = 0
        for index in range(n_series):
            if index == hero_index:
                colors.append(get_hero_color())
            else:
                colors.append(
                    _HERO_RECEDING_SEQUENCE[
                        receding_index % len(_HERO_RECEDING_SEQUENCE)
                    ]
                )
                receding_index += 1
        return colors
    if n_series in _NO_HERO_SEQUENCES:
        return list(_NO_HERO_SEQUENCES[n_series])
    colors = []
    for index in range(n_series):
        colors.append(_LONG_SERIES_SEQUENCE[index % len(_LONG_SERIES_SEQUENCE)])
    return colors


def get_hero_color() -> str:
    """Always returns Bain Red."""

    return BAIN_PALETTE["RED"]
