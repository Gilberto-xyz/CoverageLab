"""Organización de primera ejecución sin perder libros ni entregas anteriores."""
from contextlib import redirect_stdout
import io
import json
import os
from pathlib import Path
import shutil
import socket
import subprocess
import sys
import tempfile
import unittest
from unittest import mock

import preparar_proyecto as preparation


class ProjectPreparationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1]/'outputs')
        self.base = Path(self.temp.name)
        self.project = self.base / 'Proyecto con espacios'
        self.project.mkdir()
        (self.project / 'scripts').mkdir()
        (self.project / 'README.md').write_text('Mi proyecto', encoding='utf-8')

    def tearDown(self): self.temp.cleanup()

    def write(self, path, content=b'datos originales'):
        target = self.project / path
        target.parent.mkdir(parents=True, exist_ok=True)
        if isinstance(content, dict): target.write_text(json.dumps(content, ensure_ascii=False), encoding='utf-8')
        else: target.write_bytes(content)
        return target

    def run_migration(self): return preparation.prepare_project(self.project)

    def test_moves_inputs_and_loose_outputs_without_changing_their_bytes(self):
        files = {'56_SNAC_PepsiCo año.xlsx': 'recursos/56_SNAC_PepsiCo año.xlsx',
                 '51_DIAP_Softys_5w1h.xlsx': 'recursos/inputs/5w1h/51_DIAP_Softys_5w1h.xlsx',
                 '51_DIAP_Softys_5W1H - copia.xlsx': 'recursos/inputs/5w1h/51_DIAP_Softys_5W1H - copia.xlsx',
                 'inputs/5w1h/55_MULT_Natura.xlsx': 'recursos/inputs/5w1h/55_MULT_Natura.xlsx',
                 'inputs/drafts/registro.json': 'recursos/inputs/drafts/registro.json',
                 'Template_Anterior.xlsx': 'recursos/Output/Archivos_anteriores/Template_Anterior.xlsx',
                 'Template_Anterior_5w1h.xlsx': 'recursos/Output/Archivos_anteriores/Template_Anterior_5w1h.xlsx',
                 'Anterior.pptx': 'recursos/Output/Archivos_anteriores/Anterior.pptx',
                 '.estado/selections/seleccion.json': 'recursos/.estado/selections/seleccion.json'}
        for index, name in enumerate(files): self.write(name, b'{}' if name.endswith('.json') else f'contenido {index}'.encode())
        unchanged = ['scripts/codigo.py', 'Modelo_PPT.pptx', 'Plantilla_Entrada_5W1H.xlsx', '~$56_SNAC_PepsiCo año.xlsx', 'nota.txt']
        for name in unchanged: self.write(name)
        outside = self.base / '5W1H'; outside.mkdir(); (outside/'original.xlsx').write_bytes(b'exterior')
        report = self.run_migration()
        self.assertTrue(report['complete'])
        for index, (old, new) in enumerate(files.items()):
            self.assertFalse((self.project/old).exists())
            self.assertEqual((self.project/new).read_bytes(), b'{}' if old.endswith('.json') else f'contenido {index}'.encode())
        for name in unchanged: self.assertEqual((self.project/name).read_bytes(), b'datos originales')
        self.assertEqual((outside/'original.xlsx').read_bytes(), b'exterior')

    def test_coverage_and_research_results_keep_working_after_path_changes(self):
        import pipeline_explorer as explorer
        old = self.project/'Output'/'Entrega'
        file = self.write('Output/Entrega/Informe.pptx', b'ppt real')
        self.write('Output/Entrega/job.json', {'id':'coverage', 'status':'complete', 'started_at':'2026-10-26T10:00:00-06:00',
            'file':'56_SNAC_PepsiCo.xlsx', 'directory':str(old), 'outputs':[{'label':'presentation','name':file.name,'path':str(file)}]})
        self.write('Output/Entrega/selection.json', {'output_root':str(self.project/'Output')})
        research = self.project/'Output'/'Entrega5W1H'
        ppt = self.write('Output/Entrega5W1H/5W1H.pptx', b'ppt investigacion')
        self.write('Output/Entrega5W1H/_sistema/research-job.json', {'id':'research','kind':'research','status':'complete',
            'started_at':'2026-10-26T11:00:00-06:00','file':'local/51_DIAP_Softys_5w1h.xlsx','fingerprint':'f'*64,
            'directory':str(research),'outputs':[{'label':'PPT','name':ppt.name,'path':str(ppt)}]})
        self.assertTrue(self.run_migration()['complete'])
        state = explorer.ExplorerState(self.project/'recursos')
        self.assertEqual(state.output_path(state.job('coverage'), 'presentation').read_bytes(), b'ppt real')
        self.assertIn('research', state.workflow.research_jobs)
        current = self.project/'recursos/Output/Entrega'
        self.assertEqual(json.loads((current/'_sistema/selection.json').read_text())['output_root'], str(self.project/'recursos/Output'))
        self.assertTrue(list((self.project/'recursos/.estado/respaldo_organizacion').glob('*.json')))

    def test_legacy_pipeline_runs_and_selection_files_keep_their_structure(self):
        import pipeline_explorer as explorer
        run = '20261026_100000_abc'
        output = self.write(f'outputs/pipeline_explorer/{run}/Marca/Template.xlsx')
        self.write(f'outputs/pipeline_explorer/{run}/result.json', {'outputs':{'template':str(output)}})
        self.write(f'outputs/pipeline_explorer/{run}/selection.json', {'file':'56_SNAC_PepsiCo.xlsx','fingerprint':'f'*64})
        self.write('outputs/pipeline_explorer/selections/version.json', {'selections':{'Marca':2}})
        self.assertTrue(self.run_migration()['complete'])
        state = explorer.ExplorerState(self.project/'recursos')
        self.assertIn(run, state.jobs)
        self.assertTrue(state.saved_selection_path('version').is_file())

    def test_detects_old_script_delivery_folders_and_ignores_active_staging(self):
        folder = 'Mexico-Marca_Cobertura Absoluta'
        for file in ['Template_Marca.xlsx','Banco_Marca.xlsx','Marca.pptx']: self.write(f'{folder}/{file}')
        running = self.write('Output/.runs/en_proceso/entrada.xlsx')
        self.assertTrue(self.run_migration()['complete'])
        self.assertFalse((self.project/folder).exists())
        self.assertTrue((self.project/'recursos/Output'/folder/'Marca.pptx').is_file())
        self.assertEqual(running.read_bytes(), b'datos originales')

    def test_conflicts_preserve_both_versions_and_retry_on_the_next_launch(self):
        original = self.write('56_SNAC_PepsiCo.xlsx', b'primera')
        target = self.write('recursos/56_SNAC_PepsiCo.xlsx', b'segunda')
        source_dir = self.write('Output/Entrega/Informe.pptx', b'viejo')
        destination = self.write('recursos/Output/Entrega/Informe.pptx', b'nuevo')
        report = self.run_migration()
        self.assertFalse(report['complete']); self.assertEqual(len(report['pending']), 2)
        self.assertEqual(original.read_bytes(), b'primera'); self.assertEqual(target.read_bytes(), b'segunda')
        self.assertEqual(source_dir.read_bytes(), b'viejo'); self.assertEqual(destination.read_bytes(), b'nuevo')
        target.rename(target.with_name('Otro.xlsx'))
        destination.parent.rename(destination.parent.with_name('Otra_entrega'))
        self.assertTrue(self.run_migration()['complete'])
        self.assertEqual(target.read_bytes(), b'primera')

    def test_locked_files_are_left_in_place_and_retried(self):
        source = self.write('56_SNAC_PepsiCo.xlsx')
        with mock.patch.object(preparation.os, 'rename', side_effect=PermissionError('Excel abierto')):
            report = self.run_migration()
        self.assertFalse(report['complete']); self.assertTrue(source.exists())
        self.assertIn('Excel abierto', report['pending'][0]['reason'])
        self.assertTrue(self.run_migration()['complete'])

    def test_atomic_journal_write_retries_a_temporary_windows_lock(self):
        target = self.project/'registro.json'
        replace = preparation.os.replace
        calls = []
        def busy_once(source, destination):
            calls.append(destination)
            if len(calls) == 1: raise PermissionError('Registro temporalmente ocupado')
            return replace(source, destination)
        with mock.patch.object(preparation.os, 'replace', side_effect=busy_once), mock.patch.object(preparation.time, 'sleep'):
            preparation._atomic_json(target, {'completo':True})
        self.assertEqual(json.loads(target.read_text()), {'completo':True})
        self.assertEqual(len(calls), 2)

    def test_an_interrupted_move_recovers_its_manifest_on_the_next_launch(self):
        source = self.project/'Output/Entrega'
        destination = self.project/'recursos/Output/Entrega'
        self.write('Output/Entrega/job.json', {'directory':str(source)})
        destination.parent.mkdir(parents=True)
        os.rename(source, destination)
        self.write('recursos/.estado/'+preparation.REPORT, {'version':1,'complete':False,'moves':[
            {'source':'Output/Entrega','target':'recursos/Output/Entrega','status':'planned'}]})
        report = self.run_migration()
        self.assertTrue(report['complete'])
        self.assertEqual(report['moves'][0]['status'], 'moved')
        self.assertEqual(json.loads((destination/'job.json').read_text())['directory'], str(destination))

    def test_locked_metadata_is_preserved_and_repaired_without_repeating_the_move(self):
        original = self.project/'Output/Entrega'
        self.write('Output/Entrega/job.json', {'directory':str(original)})
        write = preparation._atomic_json
        def locked(path, value):
            if path.name == 'job.json': raise PermissionError('Archivo ocupado')
            return write(path, value)
        with mock.patch.object(preparation, '_atomic_json', side_effect=locked):
            report = self.run_migration()
        target = self.project/'recursos/Output/Entrega/job.json'
        self.assertFalse(report['complete'])
        self.assertEqual(json.loads(target.read_text())['directory'], str(original))
        self.assertTrue(self.run_migration()['complete'])
        self.assertEqual(json.loads(target.read_text())['directory'], str(target.parent))

    def test_repeated_launch_is_idempotent_and_finds_new_legacy_files(self):
        self.write('56_SNAC_PepsiCo.xlsx')
        first = self.run_migration()
        marker = self.project/'recursos/.estado'/preparation.REPORT
        contents = marker.read_bytes()
        self.assertEqual(self.run_migration()['moved_now'], 0)
        self.assertEqual(marker.read_bytes(), contents)
        self.write('51_DIAP_Softys.xlsx')
        self.assertEqual(self.run_migration()['moved_now'], 1)
        self.assertGreater(len(self.run_migration()['moves']), len(first['moves']))

    def test_relocation_rebases_a_delivery_copied_from_another_computer(self):
        self.write('Output/Entrega/PPT.pptx')
        self.write('Output/Entrega/job.json', {'directory':r'C:\Anterior\CoverageLab\Output\Entrega',
            'outputs':[{'path':r'C:\Anterior\CoverageLab\Output\Entrega\PPT.pptx'}]})
        self.assertTrue(self.run_migration()['complete'])
        job = json.loads((self.project/'recursos/Output/Entrega/job.json').read_text())
        self.assertEqual(job['directory'], str(self.project/'recursos/Output/Entrega'))
        self.assertEqual(job['outputs'][0]['path'], str(self.project/'recursos/Output/Entrega/PPT.pptx'))

    def test_existing_reparse_points_and_invalid_data_directory_are_not_followed(self):
        source = self.write('56_SNAC_PepsiCo.xlsx')
        original = preparation._linked
        with mock.patch.object(preparation, '_linked', side_effect=lambda path: path == source or original(path)):
            report = self.run_migration()
        self.assertFalse(report['complete']); self.assertTrue(source.exists())
        with mock.patch.object(preparation, '_linked', side_effect=lambda path: path == self.project/'recursos' or original(path)):
            with self.assertRaises(ValueError): self.run_migration()

    def test_active_server_prevents_migration_and_launcher_forwards_arguments(self):
        with socket.socket() as server:
            server.bind(('127.0.0.1',0)); server.listen()
            with redirect_stdout(io.StringIO()), mock.patch.object(preparation, 'prepare_project') as prepare:
                self.assertEqual(preparation.main(['--port',str(server.getsockname()[1])]), 1)
                prepare.assert_not_called()
        if os.name != 'nt': return
        code = Path(preparation.__file__).resolve()
        shutil.copy2(code, self.project/'scripts'/code.name)
        launcher = code.parent.parent/'Abrir_Explorador.cmd'
        shutil.copy2(launcher, self.project/launcher.name)
        self.write('scripts/pipeline_explorer.py', b'import pathlib,sys,json\npathlib.Path("args.json").write_text(json.dumps(sys.argv[1:]))\n')
        self.write('56_SNAC_PepsiCo año.xlsx', b'datos a preservar')
        result = subprocess.run(['cmd','/d','/c',str(self.project/launcher.name),'--port','0','--no-browser'],
            stdin=subprocess.DEVNULL, capture_output=True, timeout=25, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, result.stdout+result.stderr)
        self.assertEqual(json.loads((self.project/'scripts/args.json').read_text()), ['--port','0','--no-browser'])
        self.assertEqual((self.project/'recursos/56_SNAC_PepsiCo año.xlsx').read_bytes(), b'datos a preservar')


if __name__ == '__main__': unittest.main()
