import unittest

import coverage_studio as studio


studio._load_heavy_modules()


def candidate(
    pipeline: int,
    correlation: float,
    variation_gap: float,
    *,
    trend_match: bool = True,
) -> studio.OptimalPipelineCandidate:
    return studio.OptimalPipelineCandidate(
        pipeline=pipeline,
        current_correlation=correlation,
        current_variation=0.10,
        wp_current_variation=0.10,
        variation_distance_points=variation_gap,
        current_trend_match=trend_match,
        previous_year_correlation=float("nan"),
        two_year_correlation=float("nan"),
        previous_year_variation=float("nan"),
        wp_previous_year_variation=float("nan"),
        historical_trend_match=False,
        recent_shipment_outlier=False,
        forced_by_sheet=False,
    )


class AutoPipelineComparisonTests(unittest.TestCase):
    def test_auto_correlation_uses_the_highest_current_mat_correlation(self) -> None:
        candidates = (
            candidate(1, 0.80, 0.10),
            candidate(4, 0.40, 0.01),
            candidate(6, 0.92, 2.00, trend_match=False),
        )

        selection = studio.select_correlation_pipeline(candidates, forced_pipeline=1)

        self.assertEqual(selection.pipeline, 6)
        self.assertIn("máxima correlación MAT", selection.reason)

    def test_high_correlation_sacrifice_is_reported_for_balanced_override(self) -> None:
        candidates = (
            candidate(4, 0.387, 0.05),
            candidate(6, 0.925, 1.94),
        )
        comparison = studio.AutoPipelineComparison(
            correlation=studio.select_correlation_pipeline(candidates),
            balanced=studio.OptimalPipelineSelection(
                4,
                "ajuste casi exacto de variación anual",
                candidates,
            ),
        )

        diagnostics = studio.build_auto_pipeline_comparison_diagnostics(comparison)

        self.assertEqual(diagnostics["Pipeline AUTO Correlación"], 6)
        self.assertEqual(diagnostics["Pipeline AUTO Balanceado"], 4)
        self.assertEqual(diagnostics["Conflicto AUTO Correlación vs Balanceado"], "Alto")
        self.assertEqual(diagnostics["Revisión requerida"], "SI")
        self.assertEqual(
            diagnostics["Tipo de decisión AUTO Balanceado"],
            "Override por alineación de variación",
        )

    def test_same_pipeline_is_reported_without_conflict(self) -> None:
        candidates = (
            candidate(1, 0.70, 1.00),
            candidate(2, 0.90, 0.20),
        )
        correlation = studio.select_correlation_pipeline(candidates)
        comparison = studio.AutoPipelineComparison(
            correlation=correlation,
            balanced=studio.OptimalPipelineSelection(
                2,
                "correlación MAT de Año Actual positiva y variación anual alineada",
                candidates,
            ),
        )

        diagnostics = studio.build_auto_pipeline_comparison_diagnostics(comparison)

        self.assertEqual(diagnostics["Conflicto AUTO Correlación vs Balanceado"], "Sin conflicto")
        self.assertEqual(diagnostics["Revisión requerida"], "NO")

    def test_report_columns_include_both_auto_modes(self) -> None:
        columns = studio.build_pipeline_report_columns("05-26")

        self.assertIn("Pipeline AUTO Correlación", columns)
        self.assertIn("Pipeline AUTO Balanceado", columns)
        self.assertIn("Revisión requerida", columns)


if __name__ == "__main__":
    unittest.main()
