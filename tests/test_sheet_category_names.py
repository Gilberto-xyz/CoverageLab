import unittest
from datetime import datetime
from io import StringIO

import coverage_studio as studio


class SheetCategoryNameTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        studio._load_heavy_modules()
        cls.categories = studio.load_categories()

    def test_extracts_pipeline_category_and_brand_from_explicit_name(self) -> None:
        identity = studio.parse_sheet_name_identity("P1_BISC_Clorox")

        self.assertEqual(identity.pipeline, 1)
        self.assertEqual(identity.category_code, "BISC")
        self.assertEqual(identity.brand_name, "Clorox")

    def test_log_label_shows_resolved_category_for_natura_sheets(self) -> None:
        expected = {
            "P1_crpc_Natura": ("CRPC", "Cross Category (Personal Care)", "Natura [CRPC - Cross Category (Personal Care)]"),
            "P1_FRAG_Natura": ("FRAG", "Fragancias", "Natura [FRAG - Fragancias]"),
            "P1_BDCR_Natura": ("BDCR", "Cremas Corporales", "Natura [BDCR - Cremas Corporales]"),
            "P1_MAKE_Natura": ("MAKE", "Maquillaje", "Natura [MAKE - Maquillaje]"),
            "P1_FCCR_Natura": ("FCCR", "Cremas Faciales", "Natura [FCCR - Cremas Faciales]"),
        }

        for sheet_name, (category_code, category_name, expected_label) in expected.items():
            with self.subTest(sheet_name=sheet_name):
                self.assertEqual(
                    studio.format_sheet_log_label(sheet_name, category_code, category_name),
                    expected_label,
                )

    def test_log_label_preserves_legacy_scenario_without_category(self) -> None:
        self.assertEqual(
            studio.format_sheet_log_label("P1_Embasa_original"),
            "Embasa_original",
        )

    def test_recognizes_mole_as_an_explicit_category(self) -> None:
        identity = studio.parse_sheet_name_identity("P1_MOLE_Doña Maria")
        hints = studio._extract_sheet_metadata_hints(
            studio.pd.DataFrame([["Table"]]),
            "P1_MOLE_Doña Maria",
        )
        metadata = studio.resolve_sheet_bank_metadata(
            category_code="CROS",
            fabricante="Herdez",
            marca_nombre_limpio=identity.brand_name,
            section_title=identity.brand_name,
            categories_df=self.categories,
            default_pais_nombre="Mexico",
            default_cesta_nombre="Diversos",
            default_categoria_nombre="Cross Category",
            default_categoria_nombre_corto="Cross Category",
            sheet_metadata_hints=hints,
        )

        self.assertEqual(identity.pipeline, 1)
        self.assertEqual(identity.category_code, "MOLE")
        self.assertEqual(identity.brand_name, "Doña Maria")
        self.assertEqual(metadata.categoria_codigo, "MOLE")
        self.assertEqual(metadata.categoria_nombre, "Mole")

    def test_preserves_brand_scenario_after_explicit_category(self) -> None:
        identity = studio.parse_sheet_name_identity("P3_BISC_Clorox_original")

        self.assertEqual(identity.pipeline, 3)
        self.assertEqual(identity.category_code, "BISC")
        self.assertEqual(identity.brand_name, "Clorox_original")

    def test_preserves_legacy_scenario_names_when_first_token_is_not_a_category(self) -> None:
        expected = {
            "P1_Embasa_original": "Embasa_original",
            "P1_Embasa_rvol1": "Embasa_rvol1",
            "P1_Embasa_ajustado": "Embasa_ajustado",
        }

        for sheet_name, expected_brand in expected.items():
            with self.subTest(sheet_name=sheet_name):
                identity = studio.parse_sheet_name_identity(sheet_name)
                self.assertEqual(identity.pipeline, 1)
                self.assertEqual(identity.category_code, "")
                self.assertEqual(identity.brand_name, expected_brand)

    def test_does_not_consume_a_category_code_without_a_brand(self) -> None:
        identity = studio.parse_sheet_name_identity("P1_BISC")

        self.assertEqual(identity.category_code, "")
        self.assertEqual(identity.brand_name, "BISC")

    def test_explicit_sheet_category_wins_for_cross_category_metadata(self) -> None:
        hints = studio._extract_sheet_metadata_hints(
            studio.pd.DataFrame([["Table"]]),
            "P1_BISC_Clorox",
        )
        metadata = studio.resolve_sheet_bank_metadata(
            category_code="CROS",
            fabricante="Herdez",
            marca_nombre_limpio="Clorox",
            section_title="Clorox",
            categories_df=self.categories,
            default_pais_nombre="Mexico",
            default_cesta_nombre="Diversos",
            default_categoria_nombre="Cross Category",
            default_categoria_nombre_corto="Cross Category",
            sheet_metadata_hints=hints,
        )

        self.assertEqual(hints.explicit_sheet_category_code, "BISC")
        self.assertEqual(metadata.categoria_codigo, "BISC")
        self.assertEqual(metadata.categoria_nombre, "Galletas")

    def test_explicit_natura_sheet_category_overrides_crpc_file_category(self) -> None:
        expected = {
            "P1_crpc_Natura": ("CRPC", "Cross Category (Personal Care)"),
            "P1_FRAG_Natura": ("FRAG", "Fragancias"),
            "P1_BDCR_Natura": ("BDCR", "Cremas Corporales"),
            "P1_MAKE_Natura": ("MAKE", "Maquillaje-Cosmeticos"),
            "P1_FCCR_Natura": ("FCCR", "Cremas Faciales"),
        }

        for sheet_name, (expected_code, expected_name) in expected.items():
            with self.subTest(sheet_name=sheet_name):
                identity = studio.parse_sheet_name_identity(sheet_name)
                hints = studio._extract_sheet_metadata_hints(
                    studio.pd.DataFrame([["Table"]]),
                    sheet_name,
                )
                metadata = studio.resolve_sheet_bank_metadata(
                    category_code="CRPC",
                    fabricante="Natura",
                    marca_nombre_limpio=identity.brand_name,
                    section_title=identity.brand_name,
                    categories_df=self.categories,
                    default_pais_nombre="Mexico",
                    default_cesta_nombre="Cuidado Personal",
                    default_categoria_nombre="Cross Category (Personal Care)",
                    default_categoria_nombre_corto="Cross Category (Personal Care)",
                    sheet_metadata_hints=hints,
                )

                self.assertEqual(metadata.categoria_codigo, expected_code)
                self.assertEqual(metadata.categoria_nombre, expected_name)

    def test_summary_columns_include_category_and_manual_brand(self) -> None:
        columns, _, _ = studio.build_summary_columns(
            lang_index=2,
            fabricante="Herdez",
            ref_dt=datetime(2026, 5, 1),
            summary_extra_months=[],
            summary_extra_months_mode="recent",
            include_category=True,
        )

        self.assertEqual(columns[:3], ["Categoría", "Fabricante/Marca", "Pipeline"])

    def test_summary_columns_omit_category_without_an_explicit_sheet_code(self) -> None:
        columns, _, _ = studio.build_summary_columns(
            lang_index=2,
            fabricante="Herdez",
            ref_dt=datetime(2026, 5, 1),
            summary_extra_months=[],
            summary_extra_months_mode="recent",
        )

        self.assertNotIn("Categoría", columns)
        self.assertEqual(columns[:2], ["Fabricante/Marca", "Pipeline"])

    def test_summary_row_uses_resolved_category_and_clean_brand(self) -> None:
        ref_dt = datetime(2026, 5, 1)
        labels = studio.build_labels(2, "Herdez", "05-26", include_summary_category=True)
        coverage_series = studio.pd.Series(
            [30.0, 35.0],
            index=studio.pd.to_datetime(["2025-05-01", "2026-05-01"]),
        )
        variations = studio.pd.DataFrame(
            [{"Tipo": "Anual", "Cliente P1": 0.10, "WP by Numerator": 0.08}]
        )

        summary_row, bank_row, *_ = studio.build_summary_and_bank_rows(
            pipeline=1,
            marca_nombre_limpio="Clorox",
            subcategoria_nombre="",
            coverage_series=coverage_series,
            df_variations=variations,
            averages={"Penet_MAT_Actual": 42.0},
            labels=labels,
            lang_index=2,
            fabricante="Herdez",
            pais_nombre="Mexico",
            categoria_nombre="Galletas",
            cesta_nombre="Alimentos",
            coverage_reason="Prueba",
            measure_unit="Unidades",
            coverage_type="Absoluta",
            ref_month_year="05-26",
            round_coverage=False,
            summary_extra_months=[],
            summary_extra_months_mode="recent",
            include_summary_category=True,
            categoria_codigo="BISC",
        )

        self.assertEqual(summary_row["Categoría"], "Galletas")
        self.assertEqual(summary_row["Fabricante/Marca"], "Clorox")
        self.assertEqual(summary_row["Pipeline"], 1)
        self.assertEqual(bank_row["Codigo Categoria"], "BISC")
        self.assertEqual(bank_row["Categoria"], "Galletas")
        self.assertEqual(bank_row["Cesta"], "Alimentos")
        self.assertEqual(bank_row["Fabricante/Marca"], "Clorox")

        no_category_row, _, *_ = studio.build_summary_and_bank_rows(
            pipeline=1,
            marca_nombre_limpio="Embasa_original",
            subcategoria_nombre="",
            coverage_series=coverage_series,
            df_variations=variations,
            averages={"Penet_MAT_Actual": 42.0},
            labels=studio.build_labels(2, "Herdez", "05-26"),
            lang_index=2,
            fabricante="Herdez",
            pais_nombre="Mexico",
            categoria_nombre="Cross Category",
            cesta_nombre="Diversos",
            coverage_reason="Prueba",
            measure_unit="Unidades",
            coverage_type="Absoluta",
            ref_month_year="05-26",
            round_coverage=False,
            summary_extra_months=[],
            summary_extra_months_mode="recent",
        )

        self.assertNotIn("Categoría", no_category_row)
        self.assertEqual(no_category_row["Fabricante/Marca"], "Embasa_original")
        self.assertEqual(no_category_row["Pipeline"], 1)

    def test_low_penetration_key_distinguishes_categories_for_same_brand(self) -> None:
        crpc_key = studio.build_low_penetration_key(
            "Natura",
            "Cross Category (Personal Care)",
            include_category=True,
        )
        frag_key = studio.build_low_penetration_key(
            "Natura",
            "Fragancias",
            include_category=True,
        )

        self.assertNotEqual(crpc_key, frag_key)
        self.assertEqual(studio.build_low_penetration_key("Natura"), "natura")

    def test_summary_highlights_only_low_category_for_same_brand(self) -> None:
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        builder = studio.SlideBuilder(
            presentation,
            lang_index=2,
            labels=studio.build_labels(
                2,
                "Natura",
                "05-26",
                include_summary_category=True,
            ),
            coverage_label="Cobertura Absoluta",
            coverage_type="Absoluta",
            ref_month_year="05-26",
            manufacturer_name="Natura",
            country_name="Mexico",
            category_name_display="Cross Category (Personal Care)",
            tipo_eje_tend="simple",
            include_summary_category=True,
        )
        summary_df = studio.pd.DataFrame(
            [
                {
                    "Categoría": "Cross Category (Personal Care)",
                    "Fabricante/Marca": "Natura",
                    "Pipeline": 1,
                },
                {
                    "Categoría": "Fragancias",
                    "Fabricante/Marca": "Natura",
                    "Pipeline": 1,
                },
            ]
        )
        low_keys = [
            studio.build_low_penetration_key(
                "Natura",
                "Fragancias",
                include_category=True,
            )
        ]

        builder._add_editable_summary_table(
            slide,
            summary_df,
            left=studio.Inches(0.5),
            top=studio.Inches(1.0),
            width=studio.Inches(9.0),
            max_height=studio.Inches(4.6),
            low_penetration_brands=low_keys,
        )

        table = next(shape.table for shape in slide.shapes if shape.has_table)
        soft_red = studio.RGBColor(255, 235, 235)
        self.assertNotEqual(table.cell(1, 0).fill.fore_color.rgb, soft_red)
        self.assertEqual(table.cell(2, 0).fill.fore_color.rgb, soft_red)

    def test_category_summary_uses_low_penetration_line_wording(self) -> None:
        labels = studio.build_labels(
            2,
            "Natura",
            "05-26",
            include_summary_category=True,
        )

        self.assertEqual(
            labels[(2, "LowPenSummaryLinePlural")].format(n=5),
            "El estudio contiene 5 líneas de baja penetración (<200 buyers). "
            "Resultados para uso interno",
        )

    def test_buyers_log_table_identifies_category_brand_value_and_status(self) -> None:
        output = StringIO()
        output_console = studio.Console(
            file=output,
            color_system=None,
            force_terminal=False,
            width=140,
        )

        studio.report_buyers_threshold_table(
            [
                ("CRPC · Cross Category (Personal Care)", "Natura", 158),
                ("FRAG · Fragancias", "Natura", 250),
            ],
            output_console=output_console,
        )

        rendered = output.getvalue()
        self.assertIn("Validación de compradores promedio", rendered)
        self.assertIn("CRPC · Cross Category (Personal Care)", rendered)
        self.assertIn("FRAG · Fragancias", rendered)
        self.assertIn("Natura", rendered)
        self.assertIn("158", rendered)
        self.assertIn("250", rendered)
        self.assertIn("PRECAUCIÓN", rendered)
        self.assertIn("OK", rendered)
        self.assertIn("1 de 2 líneas por debajo de 200", rendered)

    def test_pg_summary_table_renders_category_and_manual_brand(self) -> None:
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        builder = studio.SlideBuilder(
            presentation,
            lang_index=2,
            labels=studio.build_labels(2, "Herdez", "05-26"),
            coverage_label="Cobertura Absoluta",
            coverage_type="Absoluta",
            ref_month_year="05-26",
            manufacturer_name="Herdez",
            country_name="Mexico",
            category_name_display="Cross Category",
            tipo_eje_tend="simple",
            coverage_slide_variant="pg",
            include_summary_category=True,
        )
        bank_df = studio.pd.DataFrame(
            [
                {
                    "Categoria": "Galletas",
                    "Fabricante": "Herdez",
                    "Fabricante/Marca": "Clorox",
                    "Pipeline": 1,
                    "%VAR Cliente": 10.0,
                    "% VAR WP by Numerator": 8.0,
                    "Cobertura Año Mov Anterior": 30.0,
                    "Cobertura Año Mov Actual": 35.0,
                }
            ]
        )

        builder._add_pg_summary_table(
            slide,
            bank_df,
            left=studio.Inches(0.5),
            top=studio.Inches(1.0),
            width=studio.Inches(9.0),
            max_height=studio.Inches(4.6),
        )

        table = next(shape.table for shape in slide.shapes if shape.has_table)
        self.assertEqual(len(table.columns), 12)
        self.assertEqual(table.cell(0, 0).text, "Category")
        self.assertEqual(table.cell(0, 1).text, "Manufacturer/Brand")
        self.assertEqual(table.cell(2, 0).text, "Galletas")
        self.assertEqual(table.cell(2, 1).text, "Clorox")

    def test_pg_summary_omits_category_without_an_explicit_sheet_code(self) -> None:
        presentation = studio.Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        builder = studio.SlideBuilder(
            presentation,
            lang_index=2,
            labels=studio.build_labels(2, "Herdez", "05-26"),
            coverage_label="Cobertura Absoluta",
            coverage_type="Absoluta",
            ref_month_year="05-26",
            manufacturer_name="Herdez",
            country_name="Mexico",
            category_name_display="Cross Category",
            tipo_eje_tend="simple",
            coverage_slide_variant="pg",
        )
        bank_df = studio.pd.DataFrame(
            [
                {
                    "Categoria": "Cross Category",
                    "Fabricante": "Herdez",
                    "Fabricante/Marca": "Embasa_original",
                    "Pipeline": 1,
                    "%VAR Cliente": 10.0,
                    "% VAR WP by Numerator": 8.0,
                    "Cobertura Año Mov Anterior": 30.0,
                    "Cobertura Año Mov Actual": 35.0,
                }
            ]
        )

        builder._add_pg_summary_table(
            slide,
            bank_df,
            left=studio.Inches(0.5),
            top=studio.Inches(1.0),
            width=studio.Inches(9.0),
            max_height=studio.Inches(4.6),
        )

        table = next(shape.table for shape in slide.shapes if shape.has_table)
        self.assertEqual(len(table.columns), 11)
        self.assertEqual(table.cell(0, 0).text, "Manufacturer/Brand")
        self.assertEqual(table.cell(2, 0).text, "Embasa_original")


if __name__ == "__main__":
    unittest.main()
