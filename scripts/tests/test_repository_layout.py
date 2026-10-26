from pathlib import Path
import os
import tempfile
import unittest
from unittest import mock

import pipeline_explorer as explorer
from project_paths import external_research_directory, state_directory, data_directory, coverage_template


class RepositoryLayoutTests(unittest.TestCase):
    def test_first_run_creates_the_data_directory_before_export(self):
        core = explorer.engine()
        previous = Path.cwd()
        with tempfile.TemporaryDirectory(dir=explorer.ROOT/'outputs') as directory:
            data = Path(directory)/'recursos'
            self.assertFalse(data.exists())
            try:
                with mock.patch.object(core, 'data_directory', return_value=data):
                    app = core.CoverageStudioUltraApp()
                self.assertTrue(data.is_dir())
                self.assertEqual(Path(app.root_dir), data)
                self.assertEqual(Path.cwd(), data)
            finally:
                os.chdir(previous)

    def test_packaged_assets_and_launcher_resolve_after_relocation(self):
        root=Path(__file__).resolve().parents[1]
        self.assertEqual(explorer.ROOT,root)
        self.assertTrue((explorer.ASSETS/'index.html').is_file())
        self.assertTrue((root/'Modelo_PPT.pptx').is_file())
        self.assertTrue((root/'studio_5w1h_vendor'/'Modelo_5W1H.pptx').is_file())
        if root.name=='scripts':
            data=root.parent/'recursos'
            self.assertEqual(data_directory(),data)
            self.assertEqual(explorer.DATA_ROOT,data)
            self.assertEqual(state_directory(data),data/'.estado')
            self.assertEqual(coverage_template(data),root/'Modelo_PPT.pptx')
            self.assertTrue((root.parent/'Abrir_Explorador.cmd').is_file())
            self.assertEqual(external_research_directory(data),root.parent.parent/'5W1H')
            launcher=(root.parent/'Abrir_Explorador.cmd').read_text()
            self.assertIn('%~dp0scripts',launcher)

    def test_old_and_new_saved_selections_keep_their_priority(self):
        with tempfile.TemporaryDirectory(dir=explorer.ROOT/'outputs') as directory:
            root=Path(directory)
            state=explorer.ExplorerState(root)
            previous=root/'outputs'/'pipeline_explorer'/'selections'/'version.json'
            previous.parent.mkdir(parents=True);previous.write_text('{}')
            self.assertEqual(state.saved_selection_path('version'),previous)
            current=state_directory(root)/'selections'/'version.json'
            current.parent.mkdir(parents=True);current.write_text('{}')
            self.assertEqual(state.saved_selection_path('version'),current)
            self.assertEqual(external_research_directory(root),root.parent/'5W1H')
            local_template=root/'Modelo_PPT.pptx'
            local_template.touch()
            self.assertEqual(coverage_template(root),local_template)


if __name__=='__main__':unittest.main()
