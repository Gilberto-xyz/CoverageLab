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

    def test_compacts_thousands_and_identifies_the_axis(self) -> None:
        fig, axis = studio.plt.subplots()
        try:
            values = [0, 100_000, 200_000]
            exponent = studio.trend_axis_magnitude_exponent(values)
            axis.plot([0, 1, 2], values)
            axis.set_ylim(0, 200_000)
            axis.yaxis.set_major_formatter(studio.build_trend_axis_formatter(2, exponent))
            fig.canvas.draw()

            labels = [tick.get_text() for tick in axis.get_yticklabels()]
            self.assertIn("200K", labels)
            self.assertEqual(studio.trend_axis_title("Sell-in", exponent, 2), "Sell-in (K = miles)")
        finally:
            studio.plt.close(fig)

    def test_places_the_million_scale_in_the_axis_title(self) -> None:
        exponent = studio.trend_axis_magnitude_exponent([2_000_000])

        self.assertEqual(exponent, 6)
        self.assertEqual(
            studio.trend_axis_title("Compras WP", exponent, 2),
            "Compras WP (M = millones)",
        )

    def test_each_language_has_a_clear_magnitude_label(self) -> None:
        self.assertEqual(studio.trend_axis_scale_text(6, 1), "M (milhões)")
        self.assertEqual(studio.trend_axis_scale_text(6, 2), "M (millones)")
        self.assertEqual(studio.trend_axis_scale_text(9, 3), "B (billions)")

    def test_abbreviates_every_supported_magnitude(self) -> None:
        self.assertEqual(studio.trend_axis_magnitude_abbreviation(3), "K")
        self.assertEqual(studio.trend_axis_magnitude_abbreviation(6), "M")
        self.assertEqual(studio.trend_axis_magnitude_abbreviation(9), "B")
        self.assertEqual(studio.trend_axis_magnitude_abbreviation(12), "T")

    def test_adds_the_million_abbreviation_to_tick_labels(self) -> None:
        fig, axis = studio.plt.subplots()
        try:
            axis.set_ylim(0, 2_000_000)
            axis.yaxis.set_major_formatter(studio.build_trend_axis_formatter(2, 6))
            fig.canvas.draw()

            labels = [tick.get_text() for tick in axis.get_yticklabels()]
            self.assertIn("2M", labels)
        finally:
            studio.plt.close(fig)

    def test_uses_compact_engineering_scales_from_thousands(self) -> None:
        self.assertEqual(studio.trend_axis_magnitude_exponent([999]), 0)
        self.assertEqual(studio.trend_axis_magnitude_exponent([1_000]), 3)
        self.assertEqual(studio.trend_axis_magnitude_exponent([999_999]), 3)
        self.assertEqual(studio.trend_axis_magnitude_exponent([1_000_000]), 6)
        self.assertEqual(studio.trend_axis_magnitude_exponent([25_000_000]), 6)
        self.assertEqual(studio.trend_axis_magnitude_exponent([2_000_000_000]), 9)

    def test_shortens_only_the_axis_label(self) -> None:
        self.assertEqual(studio.short_visible_sell_out_axis_label(1), "Compras WP")
        self.assertEqual(studio.short_visible_sell_out_axis_label(2), "Compras WP")
        self.assertEqual(studio.short_visible_sell_out_axis_label(3), "WP Purchases")

    def test_monthly_grid_is_soft_and_behind_the_series(self) -> None:
        fig, axis = studio.plt.subplots()
        try:
            axis.set_xticks([0, 1, 2])
            studio.apply_trend_grid_style(axis, "monthly")
            fig.canvas.draw()

            monthly_lines = axis.get_xgridlines()
            self.assertTrue(monthly_lines)
            self.assertTrue(all(line.get_visible() for line in monthly_lines))
            self.assertTrue(all(line.get_color() == "#AEB4BA" for line in monthly_lines))
            self.assertTrue(all(line.get_linewidth() == 0.45 for line in monthly_lines))
            self.assertTrue(all(line.get_alpha() == 0.12 for line in monthly_lines))
            self.assertTrue(axis.get_axisbelow())
        finally:
            studio.plt.close(fig)

    def test_quarterly_grid_does_not_add_monthly_guides(self) -> None:
        fig, axis = studio.plt.subplots()
        try:
            axis.set_xticks([0, 1, 2])
            studio.apply_trend_grid_style(axis, "quarterly")
            fig.canvas.draw()

            self.assertTrue(all(not line.get_visible() for line in axis.get_xgridlines()))
        finally:
            studio.plt.close(fig)


if __name__ == "__main__":
    unittest.main()
