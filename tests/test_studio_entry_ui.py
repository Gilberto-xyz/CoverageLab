"""Behavioral checks for entry editing, identity and the optional Excel handoff."""
from __future__ import annotations

from pathlib import Path
import shutil
import subprocess
import unittest


ROOT = Path(__file__).resolve().parents[1]
BOOTSTRAP = r'''
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const source=fs.readFileSync('explorer_web/studio.js','utf8');
function extract(startMarker,endMarker) {
  const start=source.indexOf(startMarker),end=source.indexOf(endMarker,start);
  assert(start>=0 && end>start,`No se encontró el código real: ${startMarker}`);
  return source.slice(start,end);
}
const plain=value=>JSON.parse(JSON.stringify(value));
const elements=new Map(),cells=new Map(),messages=[],callbacks=new Map();
function classList() {
  const values=new Set();
  return {add(...names){names.forEach(name=>values.add(name));},remove(...names){names.forEach(name=>values.delete(name));},
    contains(name){return values.has(name);},toggle(name,enabled){if(enabled===undefined)enabled=!values.has(name);if(enabled)values.add(name);else values.delete(name);}};
}
function element() {
  return {value:'',textContent:'',dataset:{},classList:classList(),disabled:false,hidden:false,options:[],
    replaceChildren(...values){this.options=values;},append(...values){this.options.push(...values);},
    addEventListener(type,handler){callbacks.set(this.id+':'+type,handler);},
    focus(){context.document.activeElement=this;},setAttribute(){}};
}
function byId(id) {
  if(!elements.has(id)){const value=element();value.id=id;elements.set(id,value);}
  return elements.get(id);
}
const grid=byId('input-grid');
grid.querySelector=selector=>{
  const match=/data-row="(\d+)"\]\[data-column="(\d+)"/.exec(selector);
  return match ? cells.get(`${match[1]}:${match[2]}`) || null : null;
};
grid.querySelectorAll=()=>[...cells.values()];
function addCell(row,column) {
  const td=element();td.dataset={row:String(row),column:String(column)};
  cells.set(`${row}:${column}`,td);return td;
}
const editor={kind:'coverage',file:'52_CARB_Actual.xlsx',book:{file:'52_CARB_Actual.xlsx',name:'52_CARB_Actual.xlsx',
  sheet:'P0_Actual',row:3,column:1,fingerprint:'original',validation:{ready:true},
  values:[['2024-01-01T00:00:00',123.25,45292]],
  display_values:[['ene 2024',123.25,'ene 2024']],display_types:[['date','number','date']]},
  pending:new Map(),originals:new Map(),undo:[],selected:null,anchor:null,busy:false};
for(let column=1;column<=3;column++) {
  addCell(3,column);editor.originals.set(JSON.stringify(['P0_Actual',3,column]),editor.book.values[0][column-1]);
}
const context={editor,byId,valueTarget:null,committingValue:false,Option:class{constructor(text,value){this.text=text;this.value=value;}},
  document:{activeElement:null,createRange(){return {selectNodeContents(){},collapse(){}};}},
  getSelection(){return {removeAllRanges(){},addRange(){}};},
  controls(){},message(text,error){messages.push({text,error});},console,
  window:{CoverageApp:{inputPending(){}}}};
vm.createContext(context);
function editingFunctions() {
  vm.runInContext(extract('  function letters(column)', '  function parseTSV(text)')+
    extract('  function beginEdit(', '  function renderGrid()')+
    extract('  function selectedRange()', '  function undo()')+
    extract('  function commitValue()', "  byId('input-cell-value').addEventListener('keydown'"),context);
}
'''


class StudioEntryUITests(unittest.TestCase):
    def run_node(self, script):
        node = shutil.which("node")
        if not node:
            self.skipTest("Node no está disponible para verificar el comportamiento web.")
        result = subprocess.run([node, "-e", BOOTSTRAP + script], cwd=ROOT,
                                capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_month_display_never_replaces_original_when_f2_is_committed_unchanged(self):
        self.run_node(r'''
editingFunctions();
for(const column of [1,2,3]) {
  const td=cells.get(`3:${column}`),raw=editor.book.values[0][column-1];
  context.paintCell(td);
  assert.equal(td.textContent,String(editor.book.display_values[0][column-1]));
  context.selection(3,column);
  assert.equal(byId('input-cell-value').value,String(raw),'La barra debe mostrar el valor original');
  context.beginEdit(td);
  assert.equal(td.textContent,String(raw),'F2 debe editar ISO/serial original, no el mes visual');
  context.commit(td);
  assert.equal(editor.pending.size,0,'F2 y salir sin modificar no debe alterar fechas ni números');
  assert.equal(editor.undo.length,0);
  assert.equal(td.textContent,String(editor.book.display_values[0][column-1]));
}
assert.match(cells.get('3:1').title,/2024-01-01/);
assert.equal(cells.get('3:3').title,'Valor original: 45292');
''')

    def test_copy_preserves_raw_dates_and_formula_bar_edits_its_original_target(self):
        self.run_node(r'''
editingFunctions();
context.selection(3,1);context.selection(3,3,true);
let copied=null,prevented=false;
callbacks.get('input-grid:copy')({preventDefault(){prevented=true;},clipboardData:{setData(type,value){assert.equal(type,'text/plain');copied=value;}}});
assert(prevented);assert.equal(copied,'2024-01-01T00:00:00\t123.25\t45292');
context.selection(3,1);
context.valueTarget={sheet:'P0_Actual',row:3,column:1};
byId('input-cell-value').value='2024-02-01';context.document.activeElement=byId('input-cell-value');
context.selection(3,2);
assert.equal(editor.pending.get(JSON.stringify(['P0_Actual',3,1])).value,'2024-02-01',
  'Cambiar de celda debe guardar la barra en su coordenada anterior');
assert.equal(context.currentValue('P0_Actual',3,2),123.25);
assert.equal(byId('input-cell-value').value,'123.25');
assert.equal(cells.get('3:1').textContent,'2024-02-01','El borrador debe mostrar su valor antes de guardar');
''')

    def test_cancel_edit_keeps_pending_raw_date_and_repaints_original_after_undo(self):
        self.run_node(r'''
editingFunctions();
vm.runInContext(extract('  function undo()', '  function validation()'),context);
const td=cells.get('3:1');context.applyBlock(3,1,[['2024-02-01']]);
context.beginEdit(td,'ene 2030');context.commit(td,true);
assert.equal(context.currentValue('P0_Actual',3,1),'2024-02-01');
assert.equal(td.textContent,'2024-02-01');
context.undo();assert.equal(editor.pending.size,0);
assert.equal(context.currentValue('P0_Actual',3,1),'2024-01-01T00:00:00');
assert.equal(td.textContent,'ene 2024');
''')

    def test_new_entry_uses_editor_identity_or_explicit_caller_without_leaking_other_project(self):
        self.run_node(r'''
editor.catalog={countries:[{code:'51',name:'Peru'},{code:'52',name:'Mexico'}],
  categories:[{id:'cloro',code:'BLEA'},{id:'gaseosas',code:'CARB'}]};
editor.book.metadata={country_code:'52',category_code:'CARB',manufacturer:'Investigación actual'};
context.inputFilters={country:'51',category:'BLEA',manufacturer:'Filtro actual'};
let opened=0;context.createDialog={showModal(){opened++;}};
context.mayChangeFile=()=>true;context.initialize=async()=>{};
context.categoryOptions=()=>{byId('create-category').value='cloro';};
context.createKind=async()=>{};
context.window.CoverageApp.context=()=>({loaded:true,file:{name:'51_BLEA_Otro.xlsx',
  country_code:'51',category_code:'BLEA',manufacturer:'Otro proyecto'},brandPipelines:{'Otra marca':4}});
vm.runInContext(extract('  async function showCreate(', "  byId('input-create').addEventListener"),context);
async function run() {
  await context.showCreate({},'research');
  assert.equal(byId('create-country').value,'52');assert.equal(byId('create-category').value,'gaseosas');
  assert.equal(byId('create-manufacturer').value,'Investigación actual');
  assert.equal(byId('create-brands').value,'','Las marcas de otra cobertura no pertenecen a esta entrada');
  editor.book=null;await context.showCreate({},'coverage');
  assert.equal(byId('create-country').value,'51');assert.equal(byId('create-category').value,'cloro');
  assert.equal(byId('create-manufacturer').value,'Filtro actual');
  await context.showCreate({file:{name:'51_BLEA_Otro.xlsx',country_code:'51',category_code:'BLEA',manufacturer:'Caller'}},'research');
  assert.equal(byId('create-manufacturer').value,'Caller');assert.equal(byId('create-brands').value,'Otra marca');
  editor.catalog.categories.push({id:'otra-descripción',code:'CARB'});
  await context.showCreate({file:{country_code:'52',category_code:'CARB',category_id:'otra-descripción',manufacturer:'Categoría explícita'}},'coverage');
  assert.equal(byId('create-category').value,'otra-descripción','El id explícito distingue categorías que comparten código');
  const before=opened;context.mayChangeFile=()=>false;await context.showCreate({},'coverage');assert.equal(opened,before);
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

    def test_identity_filters_refuse_to_clear_pending_data_and_clear_saved_nonmatching_book(self):
        self.run_node(r'''
const metadata={country_code:'52',category_code:'CARB',manufacturer:'Actual'};
editor.book.metadata=metadata;context.inputFilters={country:'52',category:'CARB',manufacturer:'Actual'};
context.commitValue=()=>{};
let changedModule=null,cleared=0,filled=0;
context.switchModule=next=>changedModule=next;
context.renderFilters=()=>{for(const [id,key] of [['input-country','country'],['input-category','category'],['input-manufacturer','manufacturer']])byId(id).value=context.inputFilters[key];};
context.clearBook=()=>{cleared++;editor.book=null;editor.file=null;};context.fillFiles=()=>filled++;
vm.runInContext(extract('  function matchesFilters(', '  function renderFilters()')+
  extract('  function mayChangeFile()', '  async function openEditor(')+
  extract("  for(const [id,key] of [['input-country','country'],['input-category','category'],['input-manufacturer','manufacturer']])byId(id).addEventListener", "  byId('input-clear-filters').addEventListener"),context);
const draftId=JSON.stringify(['P0_Actual',3,2]);editor.pending.set(draftId,{value:'999'});
byId('input-country').value='51';callbacks.get('input-country:change')();
assert.equal(byId('input-country').value,'52');assert.equal(context.inputFilters.country,'52');
assert.equal(editor.pending.get(draftId).value,'999');assert.equal(cleared,0);
assert.equal(editor.book.metadata.manufacturer,'Actual');assert.equal(changedModule,'inputs');
editor.pending.clear();byId('input-country').value='51';callbacks.get('input-country:change')();
assert.equal(context.inputFilters.country,'51');assert.equal(context.inputFilters.category,'');
assert.equal(context.inputFilters.manufacturer,'');assert.equal(cleared,1);assert.equal(filled,1);
''')

    def test_excel_actions_block_drafts_and_reload_invalidates_analysis_after_saved_data(self):
        self.run_node(r'''
let dirtyCommitted=0;context.commitValue=()=>dirtyCommitted++;
context.busy=value=>{editor.busy=value;};context.cleanEditor=()=>{editor.pending.clear();editor.originals.clear();};
context.refreshFiles=async()=>{};context.loadBook=async()=>{};
let invalidated=null,researchEdited=null;
context.window.CoverageApp.sourceChanged=async file=>invalidated=file;
context.window.StudioResearch={edited:async file=>researchEdited=file};
const calls=[];let response;
context.api=async(path,payload)=>{calls.push({path,payload});return response;};
byId('input-row').value=3;byId('input-column').value=1;
vm.runInContext(extract('  async function saved(', '  async function save()')+
  extract("  byId('input-open-excel').addEventListener", "  byId('input-discard').addEventListener"),context);
async function run() {
  editor.pending.set('draft',{value:999});
  await callbacks.get('input-open-excel:click')();await callbacks.get('input-reload-excel:click')();
  assert.equal(calls.length,0,'Ninguna acción nativa debe reemplazar datos pendientes');
  assert.equal(dirtyCommitted,2);assert.equal(editor.pending.size,1);
  editor.pending.clear();response={file:editor.file,kind:'coverage',sheet:'P0_Actual',name:editor.file,sheets:[{name:'P0_Actual'}],fingerprint:'excel-version'};
  await callbacks.get('input-reload-excel:click')();
  assert.equal(calls[0].path,'/api/studio/reload-excel');assert.equal(calls[0].payload.file,'52_CARB_Actual.xlsx');
  assert.equal(editor.book.fingerprint,'excel-version');assert.equal(invalidated,'52_CARB_Actual.xlsx');
  editor.kind='research';editor.file='external/52_CARB_Original.xlsx';editor.book={...editor.book,file:editor.file,sheet:'5_Actual_2'};
  response={...response,file:'local/52_CARB_Original.xlsx',kind:'research',sheet:'5_Actual_2',
    sheets:[{name:'5_Actual_2'}],copied_from:'external/52_CARB_Original.xlsx'};
  await callbacks.get('input-open-excel:click')();
  assert.equal(calls[1].path,'/api/studio/open-excel');assert.equal(calls[1].payload.fingerprint,'excel-version');
  assert.equal(editor.file,'local/52_CARB_Original.xlsx');assert.equal(researchEdited,editor.file);
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

    def test_excel_controls_require_installation_and_saved_entry(self):
        self.run_node(r'''
vm.runInContext(extract('  function controls()', '  function letters(column)'),context);
editor.desktop={available:false,message:'Excel no instalado'};context.controls();
assert(byId('input-open-excel').disabled);assert.equal(byId('input-open-excel').title,'Excel no instalado');
assert(!byId('input-reload-excel').disabled);
editor.desktop.available=true;context.controls();assert(!byId('input-open-excel').disabled);
editor.pending.set('draft',{value:123});context.controls();
assert(byId('input-open-excel').disabled);assert(byId('input-reload-excel').disabled);assert(byId('input-next').disabled);
editor.pending.clear();editor.busy=true;context.controls();assert(byId('input-open-excel').disabled);
editor.busy=false;editor.book=null;context.controls();assert(byId('input-open-excel').disabled);assert(byId('input-reload-excel').disabled);
''')


if __name__ == "__main__":
    unittest.main()
