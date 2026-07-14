from pathlib import Path
import sys

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scorecards_studio import _build_scorecards, _load_source_data


def _valid_raw_sheet():
    rows = [["Fecha", "Sell out", "Penetración", "Compra media", "Compra ocasión", "Frecuencia", "Buyers", "Sell in"]]
    for date in pd.date_range("2023-01-01", periods=24, freq="MS"):
        rows.append([date, 100, 20, 5, 2, 3, 240, 100])
    return pd.DataFrame(rows)


def test_load_source_data_skips_invalid_sheet_and_keeps_valid_ones(tmp_path):
    source_file = tmp_path / "55_MULT_Unilever.xlsx"
    with pd.ExcelWriter(source_file) as writer:
        _valid_raw_sheet().to_excel(writer, sheet_name="Marca válida", header=False, index=False)
        pd.DataFrame([["Sin", "formato"]]).to_excel(writer, sheet_name="Hoja inválida", header=False, index=False)

    _, sheet_names, total_pipeline, pipeline_by_brand, skipped = _load_source_data(source_file)

    assert sheet_names == ["Marca válida"]
    assert not total_pipeline.empty
    assert list(pipeline_by_brand) == ["Marca válida"]
    assert skipped[0]["sheet"] == "Hoja inválida"
    assert skipped[0]["stage"] == "la carga"


def test_build_scorecards_skips_failed_sheet_without_losing_previous_results():
    dates = pd.date_range("2024-01-01", periods=13, freq="MS").strftime("%m-%y")
    valid = pd.DataFrame(
        {
            "date": dates,
            "penetracion": [20.0] * 13,
            "buyers": [240.0] * 13,
            **{f"Pipeline {pipeline}": [90.0] * 13 for pipeline in range(7)},
        }
    )
    invalid = valid.drop(columns=["Pipeline 0"])

    scorecards, skipped = _build_scorecards(
        pais="Brasil",
        criterio="Si",
        sheet_names=["Marca válida", "P0_Hoja inválida"],
        pipeline_by_brand={"Marca válida": valid, "P0_Hoja inválida": invalid},
    )

    assert len(scorecards) == 7
    assert {entry["marca"] for entry in scorecards} == {"Marca válida"}
    assert skipped[0]["sheet"] == "P0_Hoja inválida"
    assert skipped[0]["stage"] == "el cálculo"
