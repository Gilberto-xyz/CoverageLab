from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scorecards_studio import (
    _resolve_output_target,
    _scorecard_criteria_to_generate,
    _select_criterio,
)


def _pipeline_data():
    return {
        "Marca": pd.DataFrame(
            {
                "date": ["05-26"],
            }
        )
    }


class ScorecardOutputNameTests(unittest.TestCase):
    def test_output_name_uses_unilever_suffix_for_unilever_criterion(self):
        _, output_name = _resolve_output_target(
            "55_MULT_Unilever.xlsx",
            ["Marca"],
            _pipeline_data(),
            "Si",
        )

        self.assertTrue(output_name.endswith("_unilever.xlsx"))

    def test_output_name_uses_custom_suffix_for_non_unilever_criterion(self):
        _, output_name = _resolve_output_target(
            "55_MULT_Unilever.xlsx",
            ["Marca"],
            _pipeline_data(),
            "No",
        )

        self.assertTrue(output_name.endswith("_personalizado.xlsx"))

    def test_option_three_generates_both_scorecard_criteria(self):
        self.assertEqual(
            _scorecard_criteria_to_generate("Ambos (Unilever y Personalizado)"),
            ["Si", "No"],
        )

    @patch("builtins.input", return_value="3")
    def test_third_menu_option_selects_both_formats(self, _mock_input):
        self.assertEqual(
            _select_criterio(),
            "Ambos (Unilever y Personalizado)",
        )


if __name__ == "__main__":
    unittest.main()
