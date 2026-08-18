import unittest
import io

import coverage_studio as studio


class AutoTrendChartHeightTests(unittest.TestCase):
    def test_scenario_three_enables_the_default_pipeline_animation(self) -> None:
        options = studio.ExecutionOptions.from_scenario("3")

        self.assertIsNotNone(options)
        self.assertTrue(options.auto_mode)
        self.assertEqual(options.trend_axis, "simple")
        self.assertEqual(options.trend_granularity, "monthly")
        self.assertTrue(options.animate_trend_pipeline)

    def test_scenario_three_default_config_renders_an_animated_single_axis_chart(self) -> None:
        studio._load_heavy_modules()
        options = studio.ExecutionOptions.from_scenario("3")
        presentation = studio.Presentation()
        presentation.slide_width = studio.Inches(13.333333)
        presentation.slide_height = studio.Inches(7.5)
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        trend_df = studio.pd.DataFrame(
            {
                studio.COL_DATA: [
                    "01-25", "02-25", "03-25", "04-25", "05-25", "06-25",
                    "07-25", "08-25", "09-25", "10-25", "11-25", "12-25",
                ],
                studio.COL_SELL_IN: [10, 14, 20, 13, 11, 18, 12, 16, 22, 15, 12, 10],
                studio.COL_SELL_OUT: [8, 9, 10, 12, 15, 11, 9, 14, 10, 13, 17, 12],
            }
        )
        margin = studio.Inches(0.35)

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            1,
            trend_df,
            2,
            {},
            doble_eje=(options.trend_axis == "doble"),
            granularity=options.trend_granularity,
            box_left=margin,
            box_top=studio.Inches(studio.AUTO_TREND_CHART_TOP_INCHES),
            box_width=presentation.slide_width - (2 * margin),
            box_height=studio.Inches(options.trend_chart_height_inches),
            figsize=(13.8, 5.8),
            legend_y=studio.AUTO_TREND_LEGEND_Y,
            animate_pipeline=options.animate_trend_pipeline,
        )

        picture = slide.shapes[-1]
        with studio.Image.open(io.BytesIO(picture.image.blob)) as image:
            self.assertEqual(image.format, "GIF")
            self.assertEqual(
                image.n_frames,
                studio.TREND_ANIMATION_TRANSITION_FRAMES + 1,
            )
        self.assertEqual(picture.top, studio.Inches(studio.AUTO_TREND_CHART_TOP_INCHES))
        self.assertEqual(picture.height, studio.Inches(options.trend_chart_height_inches))

    def test_auto_templates_three_and_four_use_taller_trend_chart(self) -> None:
        for scenario in ("3", "4"):
            with self.subTest(scenario=scenario):
                options = studio.ExecutionOptions.from_scenario(scenario)

                self.assertIsNotNone(options)
                self.assertEqual(
                    options.trend_chart_height_inches,
                    studio.AUTO_TREND_CHART_HEIGHT_INCHES,
                )

    def test_other_presets_keep_default_trend_chart_height(self) -> None:
        for scenario in ("5", "6", "7"):
            with self.subTest(scenario=scenario):
                options = studio.ExecutionOptions.from_scenario(scenario)

                self.assertIsNotNone(options)
                self.assertEqual(
                    options.trend_chart_height_inches,
                    studio.DEFAULT_TREND_CHART_HEIGHT_INCHES,
                )

    def test_auto_legend_is_closer_to_the_chart_content(self) -> None:
        self.assertGreater(studio.AUTO_TREND_LEGEND_Y, studio.DEFAULT_TREND_LEGEND_Y)
        layout = studio._trend_subplot_layout(studio.AUTO_TREND_LEGEND_Y, True)
        self.assertGreater(layout["top"], studio.TREND_DEFAULT_TOP_MARGIN)
        self.assertLess(layout["bottom"], studio.TREND_DEFAULT_BOTTOM_MARGIN)
        legend_lower_edge = layout["bottom"] + (
            studio.AUTO_TREND_LEGEND_Y * (layout["top"] - layout["bottom"])
        )
        self.assertGreater(legend_lower_edge, 0.01)
        self.assertLess(legend_lower_edge, 0.025)

    def test_trend_chart_uses_requested_picture_height(self) -> None:
        studio._load_heavy_modules()
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        trend_df = studio.pd.DataFrame(
            {
                studio.COL_DATA: ["01-26", "02-26", "03-26", "04-26"],
                studio.COL_SELL_IN: [10.0, 11.0, 12.0, 13.0],
                studio.COL_SELL_OUT: [8.0, 9.0, 9.5, 10.0],
            }
        )
        requested_height = studio.Inches(studio.AUTO_TREND_CHART_HEIGHT_INCHES)

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            1,
            trend_df,
            2,
            {},
            picture_height=requested_height,
        )

        picture = slide.shapes[-1]
        self.assertEqual(picture.height, requested_height)

    def test_tall_trend_chart_fits_inside_the_slide(self) -> None:
        studio._load_heavy_modules()
        presentation = studio.Presentation()
        presentation.slide_width = studio.Inches(13.333333)
        presentation.slide_height = studio.Inches(7.5)
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        trend_df = studio.pd.DataFrame(
            {
                studio.COL_DATA: ["01-26", "02-26", "03-26", "04-26"],
                studio.COL_SELL_IN: [10.0, 11.0, 12.0, 13.0],
                studio.COL_SELL_OUT: [8.0, 9.0, 9.5, 10.0],
            }
        )
        margin = studio.Inches(0.35)
        chart_top = studio.Inches(studio.AUTO_TREND_CHART_TOP_INCHES)

        studio.generar_grafico_tendencia(
            slide,
            "Marca",
            1,
            trend_df,
            2,
            {},
            box_left=margin,
            box_top=chart_top,
            box_width=presentation.slide_width - (2 * margin),
            box_height=studio.Inches(studio.AUTO_TREND_CHART_HEIGHT_INCHES),
            figsize=(13.8, 5.8),
            legend_y=studio.AUTO_TREND_LEGEND_Y,
        )

        picture = slide.shapes[-1]
        self.assertAlmostEqual(
            picture.height,
            studio.Inches(studio.AUTO_TREND_CHART_HEIGHT_INCHES),
            delta=1,
        )
        self.assertEqual(picture.top, chart_top)
        self.assertGreaterEqual(picture.left, margin)
        self.assertLessEqual(picture.left + picture.width, presentation.slide_width - margin)
        self.assertLessEqual(
            picture.top + picture.height,
            presentation.slide_height - studio.Inches(0.6),
        )


if __name__ == "__main__":
    unittest.main()
