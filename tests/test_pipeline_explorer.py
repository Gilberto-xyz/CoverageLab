"""Integration checks for preview, source identity and manual export selection."""
import contextlib
import io
import json
import math
from pathlib import Path
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request

import pipeline_explorer as explorer


class PipelineExplorerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.core = explorer.engine()

    def setUp(self):
        (explorer.ROOT / "outputs").mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(prefix="explorer_test_", dir=explorer.ROOT / "outputs")
        self.root = Path(self.temp.name)
        self.source = self.root / "52_CARB_Explorador.xlsx"
        self.make_workbook()
        shutil.copy2(explorer.ROOT / "Modelo_PPT.pptx", self.root / "Modelo_PPT.pptx")
        self.state = explorer.ExplorerState(self.root)
        self.options = explorer.options_from_request({"animate_trend_pipeline": False})
        self.original_cwd = Path.cwd()

    def tearDown(self):
        import os
        os.chdir(self.original_cwd)
        self.temp.cleanup()

    def make_workbook(self, months=36):
        from openpyxl import Workbook
        workbook = Workbook()
        workbook.remove(workbook.active)
        for name in ("P1_Cola", "P4_Limon"):
            sheet = workbook.create_sheet(name)
            sheet.append(["Table", "Weighted R_VOL1", "PENET", "VO1_BUY", "VO1_DAY", "FREQ", "BUYERS", "SELL-IN"])
            for index in range(months):
                date = self.core.pd.Timestamp(2023, 1, 1) + self.core.pd.DateOffset(months=index)
                volume = 1200 + index * 17 + 180 * math.sin(index * .7)
                sheet.append([date.to_pydatetime(), volume, 20, 6, 2, 3, 900, volume * 1.7 + math.cos(index) * 75])
        workbook.create_sheet("Sin datos").append(["vacío"])
        workbook.save(self.source)

    def analyze(self):
        with contextlib.redirect_stdout(io.StringIO()):
            return self.state.analyze({"file": self.source.name, "animate_trend_pipeline": False})

    def test_preview_shift_and_coverage_use_the_same_months(self):
        analysis = self.analyze()
        sheet = analysis["sheets"][0]
        self.assertEqual(sheet["default_pipeline"], 1)
        self.assertEqual(len(sheet["pipelines"]), 7)
        self.assertEqual(sheet["pipelines"][3]["sell_in"][:3], [None] * 3)
        self.assertEqual(sheet["pipelines"][3]["sell_in"][3:], sheet["sell_in"][:-3])
        expected = round(sum(sheet["sell_out"][-12:]) / sum(sheet["sell_in"][-15:-3]) * 100, 1)
        self.assertEqual(sheet["pipelines"][3]["coverage"], expected)
        self.assertIsNotNone(analysis["sheets"][-1]["error"])
        json.dumps(analysis, allow_nan=False)

    def test_short_history_disables_unavailable_pipelines(self):
        self.make_workbook(months=13)
        analysis = self.analyze()
        pipelines = analysis["sheets"][0]["pipelines"]
        self.assertTrue(pipelines[1]["available"])
        self.assertFalse(pipelines[2]["available"])
        self.assertIsNone(pipelines[0]["correlation"])
        with self.assertRaises(ValueError):
            self.state.reviewed_selection({"analysis_id": analysis["id"], "selections": {"P1_Cola": 6}})

    def test_manual_p0_has_diagnostics_without_changing_auto_candidates(self):
        analysis = self.analyze()
        with self.core.pd.ExcelFile(self.source) as workbook:
            df, _ = self.core.load_and_preprocess_sheet(workbook, "P1_Cola")
        variations = self.core.compute_variations_dataframe(df)
        automatic = self.core.build_optimal_pipeline_candidates(df, variations)
        reviewed = self.core.build_optimal_pipeline_candidates(df, variations, include_reference=True)
        self.assertEqual([candidate.pipeline for candidate in automatic], list(range(1, 7)))
        diagnostics = self.core.build_pipeline_selection_diagnostics(0, reviewed, "Elegido manualmente")
        self.assertAlmostEqual(diagnostics['Correlación seleccionada'], analysis['sheets'][0]['pipelines'][0]['correlation'], places=3)

    def test_invalid_month_sequences_are_blocked(self):
        core = self.core
        dates = core.pd.date_range("2023-01-01", periods=14, freq="MS")
        for values in (list(dates[:3]) + list(dates[2:]), list(reversed(dates)), list(dates[:5]) + list(dates[6:])):
            self.assertIsNotNone(explorer.monthly_problem(core.pd.DataFrame({core.COL_DATA: values}), core))
        self.assertIsNone(explorer.monthly_problem(core.pd.DataFrame({core.COL_DATA: dates}), core))

    def test_selections_reject_invalid_values_and_changed_source(self):
        analysis = self.analyze()
        for selection in ({}, {"Otra hoja": 1}, {"P1_Cola": True}, {"P1_Cola": 7}, {"Sin datos": 0}):
            with self.assertRaises(ValueError):
                self.state.reviewed_selection({"analysis_id": analysis["id"], "selections": selection})
        with self.source.open("ab") as source:
            source.write(b"changed")
        with self.assertRaisesRegex(ValueError, "cambió"):
            self.state.reviewed_selection({"analysis_id": analysis["id"], "selections": {"P1_Cola": 3}})

    def test_saved_choices_survive_a_new_session(self):
        analysis = self.analyze()
        self.state.save_selection({"analysis_id": analysis["id"], "selections": {"P1_Cola": 3}})
        self.state = explorer.ExplorerState(self.root)
        restored = self.analyze()
        self.assertEqual(restored["saved_selection"]["selections"], {"P1_Cola": 3})
        with self.assertRaises(ValueError):
            self.state.source("../52_CARB_Explorador.xlsx")

    def test_equal_bytes_with_different_names_keep_their_own_metadata(self):
        first = self.analyze()
        alternate = self.root / "51_CARB_OtroFabricante.xlsx"
        shutil.copy2(self.source, alternate)
        self.state.save_selection({"analysis_id": first["id"], "selections": {"P1_Cola": 3}})
        with contextlib.redirect_stdout(io.StringIO()):
            second = self.state.analyze({"file": alternate.name, "animate_trend_pipeline": False})
        self.assertEqual(first["fingerprint"], second["fingerprint"])
        self.assertNotEqual(first["selection_key"], second["selection_key"])
        self.assertEqual(second["country"], "Peru")
        self.assertEqual(second["manufacturer"], "OtroFabricante")
        self.assertIsNone(second["saved_selection"])

    def test_http_api_requires_local_session_and_origin(self):
        server = explorer.ThreadingHTTPServer(("127.0.0.1", 0), explorer.make_handler(self.state))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        base = f"http://127.0.0.1:{server.server_port}"
        try:
            with urllib.request.urlopen(base + "/") as response:
                self.assertIn(self.state.token, response.read().decode())
            for headers in ({}, {"X-Explorer-Token": self.state.token, "Origin": "https://example.com"},
                            {"X-Explorer-Token": self.state.token, "Host": "example.com"}):
                request = urllib.request.Request(base + "/api/files", headers=headers)
                with self.assertRaises(urllib.error.HTTPError) as error:
                    urllib.request.urlopen(request)
                self.assertEqual(error.exception.code, 403)
            request = urllib.request.Request(base + "/api/files", headers={"X-Explorer-Token": self.state.token})
            with urllib.request.urlopen(request) as response:
                self.assertEqual(json.load(response)["files"], [self.source.name])
        finally:
            server.shutdown()
            server.server_close()
            thread.join()

    def test_real_export_respects_override_and_excludes_other_sheets(self):
        from openpyxl import load_workbook
        self.source.rename(self.root / "55_MULT_Unilever.xlsx")
        self.source = self.root / "55_MULT_Unilever.xlsx"
        source = load_workbook(self.source)
        source["P1_Cola"].title = "P1_BISC_T.UL Galletas"
        source["P4_Limon"].title = "P4_MarcaDemo"
        source.save(self.source)
        source.close()
        analysis = self.analyze()
        selected = analysis["sheets"][1]
        self.assertEqual(selected["metadata"]["categoria_codigo"], "BISC")
        options = self.options
        options.pipeline_overrides = {"P4_MarcaDemo": 3}
        options.sheet_metadata_overrides = {"P4_MarcaDemo": self.core.SheetBankMetadata(**selected["metadata"])}
        app = self.core.CoverageStudioUltraApp()
        app.root_dir = str(self.root)
        with contextlib.redirect_stdout(io.StringIO()):
            outputs = app.process_file(self.source.name, options, 1, 1)
        self.assertIsNotNone(outputs)
        self.assertTrue(all(Path(path).is_file() for path in outputs.values()))
        bank = load_workbook(outputs["bank"], data_only=True)
        ws = bank.active
        headers = {cell.value: cell.column for cell in ws[1]}
        self.assertEqual(ws.cell(2, headers["Pipeline"]).value, 3)
        self.assertEqual(ws.cell(2, headers["Fabricante/Marca"]).value, "MarcaDemo")
        self.assertEqual(ws.cell(2, headers["Categoria"]).value, "Galletas")
        self.assertEqual(ws.cell(2, headers["Cobertura Año Mov Actual"]).value, selected["pipelines"][3]["coverage"])
        bank.close()
        template = load_workbook(outputs["template"])
        self.assertIn("P4_MarcaDemo", template.sheetnames)
        self.assertNotIn("P1_BISC_T.UL Galletas", template.sheetnames)
        self.assertNotIn("Sin datos", template.sheetnames)
        ws = template["P4_MarcaDemo"]
        headers = {cell.value: cell.column_letter for cell in ws[1] if cell.value is not None}
        coverage_chart = ws._charts[0]
        self.assertIn(f"${headers['P3']}$", coverage_chart.series[0].val.numRef.f)
        template.close()
        ppt = self.core.Presentation(outputs["presentation"])
        texts = " ".join(shape.text for slide in ppt.slides for shape in slide.shapes if shape.has_text_frame)
        self.assertIn("MarcaDemo | P3", texts)
        self.assertNotIn("T.UL Galletas |", texts)
        report = load_workbook(outputs["pipeline_report"], data_only=True)
        ws = report.active
        headers = {cell.value: cell.column for cell in ws[1]}
        self.assertEqual(ws.cell(2, headers["Pipeline"]).value, 3)
        self.assertIn("explorador", " ".join(str(cell.value) for cell in ws[2]))
        report.close()


if __name__ == "__main__":
    unittest.main()
