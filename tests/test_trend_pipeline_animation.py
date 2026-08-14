import io
import unittest

import coverage_studio as studio


class TrendPipelineAnimationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        studio._load_heavy_modules()

    @staticmethod
    def _trend_df():
        return studio.pd.DataFrame(
            {
                studio.COL_DATA: [
                    "01-25", "02-25", "03-25", "04-25", "05-25", "06-25",
                    "07-25", "08-25", "09-25", "10-25", "11-25", "12-25",
                ],
                studio.COL_SELL_IN: [10, 14, 20, 13, 11, 18, 12, 16, 22, 15, 12, 10],
                studio.COL_SELL_OUT: [8, 9, 10, 12, 15, 11, 9, 14, 10, 13, 17, 12],
            }
        )

    def test_pipeline_trend_is_embedded_as_a_single_play_animated_gif(self) -> None:
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            2,
            self._trend_df(),
            2,
            {},
        )

        picture = slide.shapes[-1]
        with studio.Image.open(io.BytesIO(picture.image.blob)) as image:
            self.assertEqual(image.format, "GIF")
            self.assertEqual(
                image.n_frames,
                studio.TREND_ANIMATION_TRANSITION_FRAMES + 1,
            )
            self.assertNotIn("loop", image.info)
            self.assertEqual(image.info.get("duration"), studio.TREND_ANIMATION_INITIAL_DURATION_MS)
            self.assertGreaterEqual(image.width, 2_000)
            self.assertGreaterEqual(image.height, 800)
            expected_extent = (0, 0, image.width, image.height)
            for frame_idx in range(image.n_frames):
                image.seek(frame_idx)
                self.assertEqual(image.dispose_extent, expected_extent)

    def test_dual_axis_animation_fills_the_requested_chart_bounds(self) -> None:
        presentation = studio.Presentation()
        presentation.slide_width = studio.Inches(13.333333)
        presentation.slide_height = studio.Inches(7.5)
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        margin = studio.Inches(0.4)
        top = studio.Inches(studio.AUTO_TREND_CHART_TOP_INCHES)
        width = presentation.slide_width - (2 * margin)
        height = studio.Inches(studio.AUTO_TREND_CHART_HEIGHT_INCHES)

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            2,
            self._trend_df(),
            2,
            {},
            doble_eje=True,
            box_left=margin,
            box_top=top,
            box_width=width,
            box_height=height,
            figsize=(13.8, 5.8),
            legend_y=studio.AUTO_TREND_LEGEND_Y,
        )

        picture = slide.shapes[-1]
        self.assertAlmostEqual(picture.left, margin, delta=studio.Inches(0.02))
        self.assertAlmostEqual(picture.top, top, delta=studio.Inches(0.02))
        self.assertAlmostEqual(picture.width, width, delta=studio.Inches(0.02))
        self.assertAlmostEqual(picture.height, height, delta=studio.Inches(0.02))
        with studio.Image.open(io.BytesIO(picture.image.blob)) as image:
            self.assertEqual(image.format, "GIF")
            self.assertGreaterEqual(image.width, 2_000)
            self.assertGreaterEqual(image.height, 900)
            expected_extent = (0, 0, image.width, image.height)
            for frame_idx in range(image.n_frames):
                image.seek(frame_idx)
                self.assertEqual(image.dispose_extent, expected_extent)

    def test_static_dual_axis_uses_the_same_compact_chart_bounds(self) -> None:
        presentation = studio.Presentation()
        presentation.slide_width = studio.Inches(13.333333)
        presentation.slide_height = studio.Inches(7.5)
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        margin = studio.Inches(0.35)
        top = studio.Inches(studio.AUTO_TREND_CHART_TOP_INCHES)
        width = presentation.slide_width - (2 * margin)
        height = studio.Inches(studio.AUTO_TREND_CHART_HEIGHT_INCHES)

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            2,
            self._trend_df(),
            2,
            {},
            doble_eje=True,
            box_left=margin,
            box_top=top,
            box_width=width,
            box_height=height,
            legend_y=studio.AUTO_TREND_LEGEND_Y,
            animate_pipeline=False,
        )

        picture = slide.shapes[-1]
        self.assertAlmostEqual(picture.left, margin, delta=studio.Inches(0.02))
        self.assertAlmostEqual(picture.top, top, delta=studio.Inches(0.02))
        self.assertAlmostEqual(picture.width, width, delta=studio.Inches(0.02))
        self.assertAlmostEqual(picture.height, height, delta=studio.Inches(0.02))
        with studio.Image.open(io.BytesIO(picture.image.blob)) as image:
            self.assertEqual(image.format, "PNG")
            self.assertGreaterEqual(image.width, 2_000)
            self.assertGreaterEqual(image.height, 800)

    def test_animation_can_be_disabled_for_static_exports(self) -> None:
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            2,
            self._trend_df(),
            2,
            {},
            animate_pipeline=False,
        )

        picture = slide.shapes[-1]
        with studio.Image.open(io.BytesIO(picture.image.blob)) as image:
            self.assertEqual(image.format, "PNG")
            self.assertEqual(getattr(image, "n_frames", 1), 1)

    def test_animation_copy_is_localized_and_avoids_pipeline_jargon(self) -> None:
        self.assertEqual(
            studio.trend_animation_phase_text(2, 2, "moving"),
            "Mover ventas 2 meses →",
        )
        self.assertNotIn(
            "pipeline",
            studio.trend_animation_phase_text(2, 2, "aligned").lower(),
        )

    def test_animation_month_quantity_uses_localized_singular_and_plural(self) -> None:
        expected = {
            1: ("1 mês", "2 meses"),
            2: ("1 mes", "2 meses"),
            3: ("1 month", "2 months"),
        }
        for language, (singular, plural) in expected.items():
            with self.subTest(language=language, quantity=1):
                self.assertEqual(
                    studio.trend_month_quantity_text(language, 1),
                    singular,
                )
                self.assertIn(
                    singular,
                    studio.trend_animation_phase_text(language, 1, "aligned"),
                )
            with self.subTest(language=language, quantity=2):
                self.assertEqual(
                    studio.trend_month_quantity_text(language, 2),
                    plural,
                )


if __name__ == "__main__":
    unittest.main()
