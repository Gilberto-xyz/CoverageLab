"""El título y el estado de la animación deben permanecer en filas separadas."""
import unittest
from unittest import mock

import pipeline_explorer as explorer


class TrendAnimationHeaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.core = explorer.engine()

    def test_header_fits_all_animation_phases_and_layouts(self):
        core = self.core
        dates = core.pd.date_range("2023-09-01", periods=36, freq="MS")
        frame = core.pd.DataFrame({
            core.COL_DATA: dates.strftime("%m-%y"),
            core.COL_SELL_IN: [28_000_000 + (i % 7) * 1_000_000 for i in range(36)],
            core.COL_SELL_OUT: [18_000_000 + (i % 5) * 1_000_000 for i in range(36)],
        })
        for language, figsize, legend_y, dual in (
            (1, (6.4, 5.6), -0.22, False),  # Columna Natura Brasil.
            (2, (13, 5.2), -0.25, False),
            (3, (13.8, 5.2), -0.20, True),
        ):
            with self.subTest(language=language, dual=dual):
                titles, statuses, bounds = [], [], []

                def inspect_frame(figure):
                    figure.canvas.draw()
                    renderer = figure.canvas.get_renderer()
                    title = figure._suptitle
                    status = next(text for text in figure.axes[0].texts
                                  if text.get_transform() == figure.axes[0].transAxes)
                    title_box = title.get_window_extent(renderer)
                    status_box = status.get_window_extent(renderer)
                    self.assertGreaterEqual(title_box.y0 - status_box.y1,
                                            6 * figure.dpi / 72 - 0.1)
                    for box in (title_box, status_box):
                        self.assertGreaterEqual(box.x0, 0)
                        self.assertGreaterEqual(box.y0, 0)
                        self.assertLessEqual(box.x1, figure.bbox.width)
                        self.assertLessEqual(box.y1, figure.bbox.height)
                    self.assertEqual(figure.axes[0].get_title(), "")
                    self.assertEqual(list(figure.axes[0].lines[0].get_ydata()),
                                     frame[core.COL_SELL_IN].tolist())
                    titles.append(title.get_text())
                    statuses.append(status.get_text())
                    bounds.append(figure.axes[0].get_position().bounds)
                    return core.Image.new("RGB", (8, 8), "white")

                with mock.patch.object(core, "_figure_to_rgb_frame", side_effect=inspect_frame):
                    result = core._build_animated_trend_gif(
                        marca_clean="MULT Natura", pipeline=2, df_plot=frame,
                        lang_idx=language,
                        labels_dict=core.build_labels(language, "Natura", "08-26"),
                        doble_eje=dual, figsize=figsize, legend_y=legend_y,
                    )
                self.assertTrue(result.getvalue().startswith(b"GIF"))
                self.assertEqual(len(titles), core.TREND_ANIMATION_TRANSITION_FRAMES + 1)
                self.assertTrue(all("MULT Natura" in title for title in titles))
                self.assertIn(core.trend_animation_phase_text(language, 2, "original"), statuses)
                self.assertIn(core.trend_animation_phase_text(language, 2, "moving"), statuses)
                self.assertIn(core.trend_animation_phase_text(language, 2, "aligned"), statuses)
                self.assertTrue(all(bound == bounds[0] for bound in bounds))


if __name__ == "__main__":
    unittest.main()
