from pathlib import Path
import sys

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scorecards_studio import _resolve_output_target


def _pipeline_data():
    return {
        "Marca": pd.DataFrame(
            {
                "date": ["05-26"],
            }
        )
    }


def test_output_name_uses_unilever_suffix_for_unilever_criterion():
    _, output_name = _resolve_output_target(
        "55_MULT_Unilever.xlsx",
        ["Marca"],
        _pipeline_data(),
        "Si",
    )

    assert output_name.endswith("_unilever.xlsx")


def test_output_name_uses_custom_suffix_for_non_unilever_criterion():
    _, output_name = _resolve_output_target(
        "55_MULT_Unilever.xlsx",
        ["Marca"],
        _pipeline_data(),
        "No",
    )

    assert output_name.endswith("_personalizado.xlsx")
