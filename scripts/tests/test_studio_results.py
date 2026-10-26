import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pipeline_explorer as explorer
from studio_results import catalog, open_item


class ResultsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=explorer.ROOT / 'outputs')
        self.root = Path(self.temp.name)
        self.state = explorer.ExplorerState(self.root)
        self.coverage = self.root / 'Output' / 'Peru-Clorox'
        self.research = self.root / 'Output' / 'Brasil-Natura_5W1H'
        for folder in (self.coverage, self.research):
            folder.mkdir(parents=True)
            (folder / 'Informe.pptx').write_bytes(b'ppt')
            (folder / 'Template.xlsx').write_bytes(b'xlsx')
            (folder / '_sistema').mkdir()
            (folder / '_sistema' / 'privado.xlsx').write_bytes(b'auxiliar')
            (folder / 'job.json').write_text('{}')
            (folder / '~$Informe.pptx').touch()

    def tearDown(self):
        self.temp.cleanup()

    def test_catalog_includes_old_deliveries_from_both_modules_and_only_deliverables(self):
        result = catalog(self.state)
        self.assertEqual(Path(result['directory']), self.root / 'Output')
        self.assertEqual({entry['module'] for entry in result['deliveries']}, {'coverage', 'research'})
        self.assertEqual(len(result['deliveries']), 2)
        self.assertTrue(all(entry['status'] == 'available' for entry in result['deliveries']))
        self.assertTrue(all({out['name'] for out in entry['outputs']} == {'Informe.pptx', 'Template.xlsx'}
                            for entry in result['deliveries']))

    def test_registered_delivery_has_verified_status_and_incomplete_jobs_are_visible(self):
        self.state.jobs['ok'] = {'id': 'ok', 'directory': str(self.coverage), 'status': 'complete',
                                 'file': '51_BLEA_Clorox.xlsx', 'title': 'Clorox',
                                 'started_at': '2026-10-20T10:00:00-06:00', 'outputs': []}
        self.state.workflow.research_jobs['failed'] = {'id': 'failed', 'status': 'failed',
              'file': 'local/Natura.xlsx', 'title': 'Natura', 'started_at': '2026-10-21T10:00:00-06:00',
              'directory': None, 'error': 'No fue posible exportar', 'outputs': []}
        entries = catalog(self.state)['deliveries']
        self.assertEqual(next(entry for entry in entries if entry['id'] == 'Peru-Clorox')['status'], 'complete')
        failed = next(entry for entry in entries if entry['id'] == 'job:failed')
        self.assertEqual(failed['outputs'], [])
        with self.assertRaises(ValueError):open_item(self.state, {'id': failed['id'], 'output': 'folder'})

    def test_early_deliveries_nested_in_a_run_folder_are_also_accessible(self):
        nested = self.root / 'Output' / '20261012_run' / 'Peru-Clorox'
        nested.mkdir(parents=True)
        (nested / 'Template_Peru.xlsx').write_bytes(b'xlsx')
        entry = next(entry for entry in catalog(self.state)['deliveries'] if entry['id'].startswith('20261012_run/'))
        self.assertEqual(entry['outputs'][0]['label'], 'Excel de cobertura')
        with patch('studio_results.sys.platform', 'win32'), patch('studio_results.os.startfile', create=True) as launch:
            open_item(self.state, {'id': entry['id'], 'output': 'Template_Peru.xlsx'})
            launch.assert_called_once_with(str(nested / 'Template_Peru.xlsx'))

    def test_open_uses_only_catalog_files_and_never_a_client_path(self):
        for output in ('../../README.md', '_sistema/privado.xlsx', 'job.json', str(self.root / 'other.xlsx')):
            with self.assertRaises(ValueError):open_item(self.state, {'id': self.coverage.name, 'output': output})
        with self.assertRaises(ValueError):open_item(self.state, {'id': '../', 'output': 'folder'})
        with patch('studio_results.sys.platform', 'win32'), patch('studio_results.os.startfile', create=True) as launch:
            open_item(self.state, {'id': self.coverage.name, 'output': 'Informe.pptx'})
            launch.assert_called_once_with(str(self.coverage / 'Informe.pptx'))
        with patch('studio_results.open_folder', return_value={'message': 'Abierta'}) as launch:
            open_item(self.state, {'id': self.research.name, 'output': 'folder'})
            launch.assert_called_once_with(self.research)

    def test_results_http_endpoint_keeps_session_protection(self):
        server = explorer.ExplorerHTTPServer(('127.0.0.1', 0), explorer.make_handler(self.state))
        thread = threading.Thread(target=server.serve_forever, daemon=True);thread.start()
        url = f'http://127.0.0.1:{server.server_port}/api/results'
        try:
            with self.assertRaises(HTTPError) as failure:urlopen(url)
            self.assertEqual(failure.exception.code, 403)
            failure.exception.close()
            with urlopen(Request(url, headers={'X-Explorer-Token': self.state.token})) as response:
                self.assertEqual(len(json.load(response)['deliveries']), 2)
            with urlopen(url.replace('/api/results', '/results.js')) as response:
                self.assertEqual(response.status, 200)
                self.assertIn(b'/api/results', response.read())
        finally:
            server.shutdown();server.server_close();thread.join()


if __name__ == '__main__':unittest.main()
