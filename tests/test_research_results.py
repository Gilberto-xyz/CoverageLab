"""Resultados 5W1H: archivo correcto, última entrega y acciones del panel."""
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
from unittest import mock

import pipeline_explorer as explorer


class ResearchResultsTests(unittest.TestCase):
    def test_latest_delivery_uses_dates_instead_of_restore_order(self):
        with tempfile.TemporaryDirectory(dir=explorer.ROOT/'outputs') as temporary:
            state=explorer.ExplorerState(Path(temporary))
            workflow=state.workflow
            self.assertIsNone(workflow.research_job(file_id='local/51_DIAP_Softys.xlsx'))
            directory=Path(temporary)/'Output'/'Entrega'
            directory.mkdir(parents=True)
            ppt=directory/'Informe.pptx';ppt.write_bytes(b'existing')
            base={'file':'local/51_DIAP_Softys.xlsx','status':'complete','directory':str(directory),
                  'fingerprint':'f'*64,'outputs':[{'label':'PPT','path':str(ppt)}]}
            workflow.research_jobs={'new':dict(base,id='new',started_at='2026-10-14T13:00:00-06:00'),
                                    'old':dict(base,id='old',started_at='2026-10-14T18:00:00+00:00')}
            latest=workflow.research_job(file_id=base['file'])
            self.assertEqual(latest['id'],'new')
            self.assertEqual(latest['available_outputs'],['PPT'])
            with mock.patch('studio_workflow.open_folder') as opening:
                workflow.open_research_output({'id':'new','file':base['file'],'output':'folder'})
                opening.assert_called_once_with(directory.resolve())
            ppt.unlink()
            self.assertEqual(workflow.research_job('new')['available_outputs'],[])
            with self.assertRaises(ValueError):
                workflow.open_research_output({'id':'new','file':base['file'],'output':'PPT'})

    def test_web_result_states_collapse_and_opening_follow_selected_file(self):
        node=shutil.which('node')
        if not node:self.skipTest('Node no está disponible')
        script=r'''
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
class Element{
 constructor(tag,attrs={},text=''){this.tag=tag;this.attrs=attrs;Object.assign(this,attrs);this.textContent=text;this.children=[];this.events={};}
 append(...children){this.children.push(...children);} replaceChildren(...children){this.children=children;}
 addEventListener(name,callback){this.events[name]=callback;}
}
const $=(...args)=>new Element(...args),messages=[],calls=[],stored=new Map();
const state={file:'local/A.xlsx',job:null,opening:null,resultsCollapsed:false};
const nodes={results:$('div'),generate:$('button',{},'Generar informe 5W1H')};
const context={state,nodes,$,message:text=>messages.push(text),updateActions(){},
 localStorage:{setItem:(k,v)=>stored.set(k,v)},window:{CoverageApp:{resultButton:(button,label,kind)=>{button.textContent=label;button.kind=kind;}}}};
vm.createContext(context);
const source=fs.readFileSync('explorer_web/research.js','utf8'),start=source.indexOf('  function renderJob() {'),end=source.indexOf('  async function pollJob(',start);
vm.runInContext(source.slice(start,end)+';this.renderJob=renderJob;',context);
const all=(root=nodes.results)=>[root,...root.children.flatMap(child=>all(child))];
const button=label=>all().find(node=>node.tag==='button'&&node.textContent===label);
async function run(){
 context.renderJob();assert(button('Abrir Excel').disabled);assert(button('Abrir carpeta').disabled);
 state.job={id:'run',file:state.file,status:'running',outputs:[]};context.renderJob();assert(button('Abrir PPT').disabled);
 state.job={...state.job,status:'complete',title:'Entrega A',available_outputs:['Excel','PPT','Registro'],directory:'Output/A'};
 context.renderJob();assert(!button('Abrir Excel').disabled);assert(!button('Abrir PPT').disabled);
 assert(button('Selección y origen').disabled);assert.equal(button('Abrir carpeta').kind,'folder');
 await button('Contraer resultados').events.click();
 assert.equal(stored.get('coveragelab-research-results-collapsed'),'true');
 assert(all().find(node=>node.id==='research-results-content').hidden);
 assert(all().includes(nodes.generate),'Generar sigue disponible al contraer');
 await button('Mostrar resultados').events.click();
 let complete;state.bridge={api:async(path,payload)=>{calls.push({path,payload});return new Promise(resolve=>complete=resolve);}};
 const opening=button('Abrir PPT').events.click();
 assert(button('Abrir Excel').disabled);assert(button('Abriendo…').disabled);
 await button('Abrir carpeta').events.click();assert.equal(calls.length,1);
 assert.equal(calls[0].path,'/api/research/open-output');assert.equal(calls[0].payload.file,'local/A.xlsx');
 state.file='local/B.xlsx';state.job=null;complete({message:'Respuesta de A'});await opening;
 assert(!messages.includes('Respuesta de A'),'No mostrar respuesta de otra entrada');
 assert(button('Abrir PPT').disabled);assert.equal(state.opening,null);
}
run().catch(error=>{console.error(error);process.exitCode=1;});
'''
        result=subprocess.run([node,'-e',script],cwd=explorer.ROOT,capture_output=True,text=True,encoding='utf-8')
        self.assertEqual(result.returncode,0,result.stdout+result.stderr)


if __name__=='__main__':unittest.main()
