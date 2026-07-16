import unittest

import coverage_studio as studio


class TrendVariationTableHighlightTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        studio._load_heavy_modules()

    def test_softly_highlights_only_type_and_period_for_twelve_month_comparisons(self) -> None:
        variations = studio.pd.DataFrame(
            {
                "Tipo": ["Anual", "Semestral", "Semestral", "Trimestral", "Trimestral"],
                "Periodo": [
                    "MAT May-26 x MAT May-25",
                    "SEM May-26 x SEM Nov-25",
                    "SEM May-26 x SEM May-25",
                    "TRI May-26 x TRI Feb-26",
                    "TRI May-26 x TRI May-25",
                ],
                "WP by Numerator": [0.10, 0.08, 0.06, -0.04, -0.02],
                "Cliente P0": [0.09, 0.07, 0.05, -0.03, -0.01],
                "_CompareLagMonths": [12, 6, 12, 3, 12],
            }
        )
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        builder = studio.SlideBuilder(
            presentation,
            2,
            {},
            "Cobertura Absoluta",
            "Absoluta",
            "05-26",
            "Cliente",
            "México",
            "Categoría",
            "simple",
        )

        builder._add_editable_variations_table(
            slide,
            variations,
            left=studio.Inches(0.3),
            top=studio.Inches(0.3),
            width=studio.Inches(6.2),
            max_height=studio.Inches(1.15),
        )

        table = next(shape.table for shape in slide.shapes if shape.has_table)
        highlighted_rows = {1, 3, 5}
        for row_idx in range(1, len(table.rows)):
            type_and_period_fills = {
                str(table.cell(row_idx, col_idx).fill.fore_color.rgb)
                for col_idx in (0, 1)
            }
            value_fills = {
                str(table.cell(row_idx, col_idx).fill.fore_color.rgb)
                for col_idx in range(2, len(table.columns))
            }
            if row_idx in highlighted_rows:
                self.assertEqual(type_and_period_fills, {"E6FEF8"})
            else:
                self.assertEqual(type_and_period_fills, {"FFFFFF"})
            self.assertNotIn("E6FEF8", value_fills)


if __name__ == "__main__":
    unittest.main()
