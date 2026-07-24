import unittest

import coverage_studio as studio


class PresentationSectionGroupingTests(unittest.TestCase):
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
