from datetime import datetime, timezone
from pathlib import Path
import json
import shutil
import subprocess
import tempfile
import unittest

from export_progress import ExportProgressReporter, historical_estimate, snapshot


ROOT=Path(__file__).resolve().parents[1]


class ExportProgressTests(unittest.TestCase):
    def test_real_units_drive_progress_and_eta_without_reaching_100_early(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'outputs') as directory:
            clock=[1000.0]
            started=datetime.fromtimestamp(clock[0],timezone.utc).isoformat()
            path=Path(directory)/'progress.json'
            reporter=ExportProgressReporter(path,started,clock=lambda:clock[0])
            reporter.update('starting')
            first=json.loads(path.read_text())
            self.assertIsNone(snapshot(first,'running',started,now=1000)['remaining_seconds'])
            clock[0]=1005;reporter.update('excel',1,4,'Hoja A')
            self.assertEqual(json.loads(path.read_text())['percent'],16.2)
            reporter.update('presentation',0,20)
            clock[0]=1015;reporter.update('presentation',2,20)
            raw=json.loads(path.read_text())
            estimate=snapshot(raw,'running',started,now=1015)
            self.assertGreater(estimate['remaining_seconds'],90)
            self.assertLess(snapshot(raw,'running',started,now=1020)['remaining_seconds'],estimate['remaining_seconds'])
            self.assertIsNone(snapshot(raw,'running',started,now=2000)['remaining_seconds'])
            before=raw['percent'];reporter.update('presentation',1,20)
            self.assertGreaterEqual(json.loads(path.read_text())['percent'],before)
            reporter.update('packaging',1,1)
            raw=json.loads(path.read_text())
            self.assertLess(raw['percent'],100)
            finished=snapshot(raw,'complete',started,now=1025)
            self.assertEqual(finished['percent'],100)
            self.assertEqual(finished['remaining_seconds'],0)
            self.assertEqual(snapshot(finished,'complete',started,now=5000)['elapsed_seconds'],25)

    def test_snapshot_sanitizes_invalid_data_and_failure_never_claims_completion(self):
        started='1970-01-01T00:00:00+00:00'
        raw={'percent':float('inf'),'completed':float('nan'),'total':float('inf'),'estimated_end_at':200}
        value=snapshot(raw,'failed',started,now=100)
        self.assertLess(value['percent'],100)
        self.assertIsNone(value['remaining_seconds'])
        json.dumps(value,allow_nan=False)
        self.assertEqual(snapshot(None,'running','invalid',now=100)['elapsed_seconds'],0)

    def test_history_uses_only_completed_measured_runs(self):
        self.assertIsNone(historical_estimate([{'status':'complete'}],4))
        jobs=[{'status':'complete','elapsed_seconds':20,'work_units':2},
              {'status':'complete','elapsed_seconds':60,'work_units':3},
              {'status':'failed','elapsed_seconds':1000,'work_units':1}]
        self.assertEqual(historical_estimate(jobs,4),60)

    def test_browser_views_show_remaining_time_preparation_failure_and_completion(self):
        node=shutil.which('node')
        if not node:self.skipTest('Node no está disponible')
        script=r'''
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const context={window:{}};vm.createContext(context);
vm.runInContext(fs.readFileSync('explorer_web/generation.js','utf8'),context);
const api=context.window.GenerationProgress;
let job={status:'running',progressReceivedAt:10000,progress:{stage:'presentation',label:'Diapositivas PPT',
  completed:12,total:80,percent:44,elapsed_seconds:20,remaining_seconds:65}};
let view=api.view(job,null,15000);
assert.equal(view.primary,'PPT 12/80 · 44%');assert.equal(view.secondary,'≈ 1 min restantes');
assert.ok(view.detail.includes('25 s'));
assert.equal(api.view(job,null,80000).secondary,'Recalculando estimación…');
job.progress.remaining_seconds=null;assert.equal(api.view(job,null,15000).secondary,'Estimando tiempo…');
job.status='failed';view=api.view(job,null,15000);assert.equal(view.mode,'failed');assert.ok(view.percent<100);
job.status='complete';job.completedObservedAt=15000;
assert.equal(api.view(job,null,16000).percent,100);assert.equal(api.view(job,null,20000),null);
const prep=api.view(null,{startedAt:1000,completed:3,total:12},4000);
assert.equal(prep.primary,'Gráficas 3/12');assert.equal(prep.percent,1.25);assert.ok(prep.secondary.includes('preparación'));
assert.equal(api.duration(3661),'1 h 1 min');assert.equal(api.view(null,null,0),null);
'''
        subprocess.run([node,'-e',script],cwd=ROOT,check=True,capture_output=True,text=True)


if __name__=='__main__':unittest.main()
