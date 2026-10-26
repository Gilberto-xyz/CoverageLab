"""Spanish display text survives capture, metadata lookup and Excel output."""
import tempfile
import unicodedata
import unittest
from contextlib import closing
from pathlib import Path

from openpyxl import load_workbook

import archivos_studio as nomenclature
from pipeline_explorer import engine
from studio_inputs import InputStudio
from studio_5w1h_vendor import report_engine, template_engine


class SpanishCatalogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.core = engine()

    def test_processing_catalogs_display_correct_spanish(self):
        categories = self.core.load_categories()
        for code, expected in {
            'DIAP': 'Pañales-Pañales Desechables',
            'CWIP': 'Paños Húmedos', 'FTIS': 'Pañuelos Faciales',
            'DRUB': 'Destapacañerías', 'CRTO': 'Pañitos + Papel Higiénico',
            'TOIL': 'Jabón de Tocador', 'SUNP': 'Protección Solar',
            'CLPC': 'Lápices de Colores', 'COFF': 'Café-Consolidado de Café',
        }.items():
            with self.subTest(code=code):
                self.assertEqual(self.core.CATEGORY_MAP[code], expected)
                self.assertEqual(categories.loc[code, 'cat'], expected)
                self.assertEqual(report_engine.lookup_category_info(code)[1], expected)

    def test_capture_search_accepts_accents_case_and_decomposed_unicode(self):
        for catalog in (nomenclature, template_engine):
            for query, code in [('pañales', 'DIAP'), ('PAÑALES', 'DIAP'),
                                ('panales', 'DIAP'), ('jabón', 'TOIL'),
                                ('jabon', 'TOIL'), ('lápices', 'CLPC'),
                                ('papel higienico', 'TOIP')]:
                for text in (query, unicodedata.normalize('NFD', query)):
                    with self.subTest(catalog=catalog.__name__, query=text):
                        matches = catalog.buscar_categorias(text)
                        self.assertIn(code, [item['cod'] for item in matches])
                        if code == 'DIAP':
                            self.assertTrue(all('Pañales' in item['descripcion'] for item in matches))

    def test_country_names_keep_accents_and_population_coverage(self):
        for code, name in [('52', 'México'), ('51', 'Perú'), ('67', 'Panamá'),
                           ('69', 'República Dominicana')]:
            for catalog in (nomenclature, template_engine):
                for query in (name, catalog.strip_accents(name), unicodedata.normalize('NFD', name)):
                    with self.subTest(catalog=catalog.__name__, query=query):
                        self.assertEqual(catalog.obtener_codigo_pais(query), (code, name))
                        self.assertEqual(catalog.obtener_nombre_pais(query), name)
                        self.assertIsNotNone(catalog.pop_coverage.get(name))
            self.assertEqual(self.core.COUNTRY_MAP[code], name)
            self.assertEqual(report_engine.lookup_country_name(int(code)), name)
            self.assertEqual(self.core.get_population_coverage_percent(name),
                             self.core.get_population_coverage_percent(nomenclature.strip_accents(name)))

    def test_subcategory_lookup_returns_accented_canonical_text(self):
        for query, expected in [('Liquido', 'Líquido'), ('Gato Humedo', 'Gato Húmedo'),
                                ('Bebe', 'Bebé'), ('Topico - Unguento', 'Tópico - Ungüento')]:
            self.assertEqual(self.core.extract_sheet_subcategory(f'P0_Marca ({query})'), expected)

    def test_input_and_excel_keep_spanish_characters(self):
        output = Path(__file__).resolve().parents[1] / 'outputs'
        output.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=output) as directory:
            studio = InputStudio(Path(directory))
            book = studio.create({'country_code': '51', 'category_code': 'DIAP',
                                  'manufacturer': 'Compañía ÁÉÍÓÚ ü', 'brands': ['Niñez'], 'unit': 'VOL1_P'})
            self.assertEqual(book['file'], '51_DIAP_Compañía ÁÉÍÓÚ ü.xlsx')
            self.assertEqual(book['sheet'], 'P0_Niñez')
            self.assertEqual(book['metadata']['country'], 'Perú')
            self.assertEqual(book['metadata']['category'], 'Pañales_Pañales Desechables')
            with closing(load_workbook(Path(directory) / book['file'])) as workbook:
                self.assertEqual(workbook['P0_Niñez']['H2'].value, 'SELL-IN: Niñez')


if __name__ == '__main__':
    unittest.main()
