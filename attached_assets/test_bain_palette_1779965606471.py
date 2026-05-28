"""Regression tests for the locked Bain chart palette."""

from __future__ import annotations

import unittest

from src.bain_palette import BAIN_PALETTE, get_hero_color, get_series_palette


class TestBainPalette(unittest.TestCase):
    def test_one_series_returns_graphite_1(self) -> None:
        self.assertEqual(get_series_palette(1), [BAIN_PALETTE["GRAPHITE_1"]])

    def test_five_series_returns_documented_sequence(self) -> None:
        self.assertEqual(
            get_series_palette(5),
            [
                BAIN_PALETTE["GRAPHITE_1"],
                BAIN_PALETTE["GRAPHITE_3"],
                BAIN_PALETTE["SKY_2"],
                BAIN_PALETTE["FOREST_2"],
                BAIN_PALETTE["SUNSET_2"],
            ],
        )

    def test_seven_series_are_distinct_and_in_palette(self) -> None:
        colors = get_series_palette(7)
        self.assertEqual(len(colors), 7)
        self.assertEqual(len(set(colors)), 7)
        self.assertTrue(set(colors).issubset(set(BAIN_PALETTE.values())))

    def test_hero_at_first_position_uses_red_and_receding_graphites(self) -> None:
        self.assertEqual(
            get_series_palette(3, hero_index=0),
            [get_hero_color(), BAIN_PALETTE["GRAPHITE_3"], BAIN_PALETTE["GRAPHITE_4"]],
        )

    def test_hero_at_middle_position_uses_red_and_receding_graphites(self) -> None:
        self.assertEqual(
            get_series_palette(3, hero_index=1),
            [BAIN_PALETTE["GRAPHITE_3"], get_hero_color(), BAIN_PALETTE["GRAPHITE_4"]],
        )

    def test_get_hero_color_returns_bain_red(self) -> None:
        self.assertEqual(get_hero_color(), "#CC0000")

    def test_palette_values_match_documented_hex_values(self) -> None:
        self.assertEqual(
            dict(BAIN_PALETTE),
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
            },
        )

    def test_sand_family_is_not_present(self) -> None:
        for key in ("SAND_1", "SAND_2", "sand_1", "sand_2"):
            with self.assertRaises(KeyError):
                _ = BAIN_PALETTE[key]


if __name__ == "__main__":
    unittest.main()
