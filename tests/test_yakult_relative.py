import unittest
import os
from unittest.mock import patch
import tempfile
import shutil
from pathlib import Path
from types import SimpleNamespace

import coverage_studio as studio


class YakultRelativeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        studio._load_heavy_modules()

    def test_scenario_selects_relative_without_rounding(self):
        for alias in ('8', 'Yakult - BR- Relativa', 'YAKULT_BR_RELATIVA'):
            options = studio.ExecutionOptions.from_scenario(alias)
            self.assertEqual(options.coverage_type, 'relativa')
            self.assertEqual(options.coverage_slide_variant, 'yakult')
            self.assertFalse(options.round_coverage)
            self.assertFalse(options.include_english)

    def test_mode_eight_prompts_for_trend_axis(self):
        for answer, expected in [('1', 'simple'), ('2', 'doble')]:
            with patch.dict(os.environ, {}, clear=True), \
                 patch.object(studio, 'tipo_cobertura', return_value='8'), \
                 patch.object(studio, 'clear_and_print_summary'), \
                 patch('builtins.print'), patch('builtins.input', return_value=answer) as prompt:
                options = studio.CoverageStudioUltraApp().gather_interactive_options()
                prompt.assert_called_once()
                self.assertEqual(options.trend_axis, expected)
                self.assertEqual(options.coverage_slide_variant, 'yakult')
                self.assertEqual(options.coverage_type, 'relativa')

    def test_mode_eight_accepts_axis_in_automatic_runs(self):
        for axis in ('simple', 'doble'):
            with patch.dict(os.environ, {'AUTO_FILE': 'example.xlsx',
                                         'AUTO_COV_TYPE': '8', 'AUTO_EJE': axis}, clear=True):
                self.assertEqual(studio.ExecutionOptions.from_environment().trend_axis, axis)

    def test_header_uses_reference_month_and_mat_mean(self):
        ppt = studio.Presentation()
        ppt.slide_width = studio.Inches(13.333)
        builder = studio.SlideBuilder(
            ppt, 1, {}, 'Cobertura Relativa', 'relativa', '07-26',
            'Yakult', 'Brasil', 'Lacteos', 'simple', coverage_slide_variant='yakult',
        )
        assets = SimpleNamespace(
            coverage_series=studio.pd.Series(
                [40.7, 44.3], index=studio.pd.to_datetime(['2025-07-01', '2026-07-01']),
            ),
            variation_table=studio.pd.DataFrame(),
            penet_mat_actual=2.3583333333,
        )
        slide = ppt.slides.add_slide(ppt.slide_layouts[6])
        builder._add_yakult_coverage_header(slide, assets)
        text = '\n'.join(shape.text for shape in slide.shapes if shape.has_text_frame)
        self.assertIn('Mensal 2026', text)
        self.assertIn('2,4', text)
        table = next(shape.table for shape in slide.shapes if shape.has_table)
        self.assertEqual([table.cell(2, i).text for i in range(3)], ['40,7', '44,3', '3,6'])
        self.assertIn('25', table.cell(1, 0).text)
        self.assertIn('26', table.cell(1, 1).text)
        assets.coverage_series.iloc[0] = float('nan')
        slide_missing = ppt.slides.add_slide(ppt.slide_layouts[6])
        builder._add_yakult_coverage_header(slide_missing, assets)
        table_missing = next(shape.table for shape in slide_missing.shapes if shape.has_table)
        self.assertEqual(table_missing.cell(2, 2).text, '-')

    def test_relative_coverage_adjusts_for_population(self):
        dates = studio.pd.date_range('2025-01-01', periods=24, freq='MS')
        frame = studio.pd.DataFrame({studio.COL_DATA: dates, studio.COL_SELL_IN: 100,
                                     studio.COL_SELL_OUT: 30})
        result = studio.compute_coverage_dataframe(frame, 'Brasil', 'relativa', False)
        expected = round(30 / (studio.get_population_coverage_percent('Brasil') / 100), 1)
        self.assertEqual(result['P1'].iloc[-1], expected)

    def test_embedded_template_is_reused_only_for_yakult(self):
        root = Path(studio.__file__).parent
        with tempfile.TemporaryDirectory() as directory:
            shutil.copyfile(root / 'Modelo_PPT.pptx', Path(directory) / 'Modelo_PPT.pptx')
            for language in ('ES', 'PT', 'EN'):
                for variant in ('classic', 'complemented', 'pg', 'yakult'):
                    with self.subTest(language=language, variant=variant):
                        ppt, _ = studio.copy_and_prune_template(directory, language, variant)
                        self.assertEqual(len(ppt.slides), 7)
                        self.assertNotIn(studio.YAKULT_TEMPLATE_SLIDE_NAME, [s.name for s in ppt.slides])
                        slide = studio.add_coverage_slide(ppt, variant)
                        houses = [s for s in slide.shapes if s.name == 'Yakult.Penetration.House']
                        self.assertEqual(len(houses), int(variant == 'yakult'))
                        if houses:
                            self.assertEqual(houses[0].image.size, (605, 605))
                            second = studio.add_coverage_slide(ppt, variant)
                            self.assertEqual(second.shapes[0].image.blob, houses[0].image.blob)
                        output = Path(directory) / 'result.pptx'
                        ppt.save(output)
                        reopened = studio.Presentation(output)
                        self.assertNotIn(studio.YAKULT_TEMPLATE_SLIDE_NAME, [s.name for s in reopened.slides])


if __name__ == '__main__':
    unittest.main()
