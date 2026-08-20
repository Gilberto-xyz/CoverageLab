import unittest

import coverage_studio as studio


class PresentationSectionGroupingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        studio._load_heavy_modules()

    def test_natura_sections_include_explicit_sheet_category(self) -> None:
        expected = {
            "P1_CRPC_Natura": (
                "CRPC",
                "Cross Category (Personal Care)",
                "Natura — Personal Care",
            ),
            "P1_FRAG_Natura": ("FRAG", "Fragancias", "Natura — Fragancias"),
            "P1_BDCR_Natura": ("BDCR", "Cremas Corporales", "Natura — Cremas Corporales"),
            "P1_MAKE_Natura": ("MAKE", "Maquillaje-Cosmeticos", "Natura — Maquillaje"),
            "P1_FCCR_Natura": ("FCCR", "Cremas Faciales", "Natura — Cremas Faciales"),
        }

        for sheet_name, (category_code, category_name, expected_title) in expected.items():
            with self.subTest(sheet_name=sheet_name):
                section_title, current_title = studio.build_section_title_for_sheet(
                    sheet_name,
                    None,
                    category_code=category_code,
                    category_name=category_name,
                )
                self.assertEqual(section_title, expected_title)
                self.assertEqual(current_title, expected_title)

    def test_pipeline_slide_titles_use_category_label(self) -> None:
        label = studio.build_presentation_brand_category_label(
            "P1_FRAG_Natura",
            category_code="FRAG",
            category_name="Fragancias",
        )

        self.assertEqual(
            studio.build_pipeline_presentation_title(label, 6),
            "Natura — Fragancias | P6",
        )
        self.assertEqual(
            studio.build_pipeline_presentation_title(
                label,
                6,
                slide_kind="evolution",
                lang_index=2,
            ),
            "Natura — Fragancias | P6 | Evolución y variación",
        )

    def test_legacy_section_does_not_inherit_file_category(self) -> None:
        section_title, _ = studio.build_section_title_for_sheet(
            "P1_Embasa_original",
            None,
            category_code="CROS",
            category_name="Cross Category",
        )

        self.assertEqual(section_title, "Embasa original")

    def test_each_unilever_sheet_gets_its_own_visible_section_title(self) -> None:
        sheet_names = [
            "P5_T.UL Sabonetes",
            "P5_T.UL Sabonetes Barra",
            "P6_T.UL Sabonetes Líquido",
            "P2_T.UL FabClean",
            "P2_T.UL FabClean Roupa Pó",
            "P2_T.UL FabClean Roupa Líquido",
        ]
        expected_titles = [
            "T.UL Sabonetes",
            "T.UL Sabonetes Barra",
            "T.UL Sabonetes Líquido",
            "T.UL FabClean",
            "T.UL FabClean Roupa Pó",
            "T.UL FabClean Roupa Líquido",
        ]

        current_title = None
        actual_titles = []
        for sheet_name in sheet_names:
            section_title, current_title = studio.build_section_title_for_sheet(
                sheet_name,
                current_title,
            )
            actual_titles.append(section_title)

        self.assertEqual(actual_titles, expected_titles)

    def test_each_sheet_registers_only_its_own_three_slides(self) -> None:
        section_slide_map = {}
        section_titles = [
            "T.UL Sabonetes",
            "T.UL Sabonetes Barra",
            "T.UL Sabonetes Líquido",
            "T.UL FabClean",
            "T.UL FabClean Roupa Pó",
            "T.UL FabClean Roupa Líquido",
        ]

        for group_index, section_title in enumerate(section_titles):
            studio.register_section_slide_range(
                section_slide_map,
                section_title,
                start_idx=7 + (group_index * 3),
                count=3,
            )

        self.assertEqual(list(section_slide_map), section_titles)
        for group_index, section_title in enumerate(section_titles):
            expected_start = 7 + (group_index * 3)
            self.assertEqual(
                section_slide_map[section_title],
                [expected_start, expected_start + 1, expected_start + 2],
            )


if __name__ == "__main__":
    unittest.main()
