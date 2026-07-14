import unittest

import coverage_studio as studio


class TrendAxisFormattingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        studio._load_heavy_modules()

    def test_adds_thousands_separators_to_unscaled_values(self) -> None:
        self.assertEqual(studio.format_trend_axis_tick(123_456.78), "123,456.78")
        self.assertEqual(studio.format_trend_axis_tick(-12_500), "-12,500")

    def test_normalizes_small_zero_ticks(self) -> None:
        self.assertEqual(studio.format_trend_axis_tick(-0.0), "0")

    def test_explains_scientific_notation_for_millions(self) -> None:
        fig, axis = studio.plt.subplots()
        try:
            axis.plot([0, 1, 2], [0, 1_000_000, 2_000_000])
            axis.yaxis.set_major_formatter(studio.build_trend_axis_formatter(2))
            fig.canvas.draw()

            self.assertEqual(axis.yaxis.get_offset_text().get_text(), "1e6 (millones)")
        finally:
            studio.plt.close(fig)

    def test_each_language_has_a_clear_magnitude_label(self) -> None:
        self.assertEqual(studio.trend_axis_scale_text(6, 1), "1e6 (milhões)")
        self.assertEqual(studio.trend_axis_scale_text(6, 2), "1e6 (millones)")
        self.assertEqual(studio.trend_axis_scale_text(9, 3), "1e9 (billions)")

    def test_uses_engineering_scales_only_for_millions_and_above(self) -> None:
        self.assertEqual(studio.trend_axis_magnitude_exponent([999_999]), 0)
        self.assertEqual(studio.trend_axis_magnitude_exponent([1_000_000]), 6)
        self.assertEqual(studio.trend_axis_magnitude_exponent([25_000_000]), 6)
        self.assertEqual(studio.trend_axis_magnitude_exponent([2_000_000_000]), 9)


if __name__ == "__main__":
    unittest.main()
