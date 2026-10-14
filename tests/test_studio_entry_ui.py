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
  controls(){},commitValue(){},message(text,error){messages.push({text,error});},console,
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

    def test_capture_methods_show_one_editor_and_switch_without_discarding_values(self):
        self.run_node(r'''
context.nativePanel=element();context.entryMode='web';context.modeChosen=false;
const preferences=[];context.localStorage={setItem(key,value){preferences.push({key,value});}};
vm.runInContext(extract('  function setEntryMode(', "  byId('input-mode-excel').addEventListener"),context);
const id=JSON.stringify(['P0_Actual',3,1]);
editor.pending.set(id,{sheet:'P0_Actual',row:3,column:1,value:'Fecha por corregir'});
context.setEntryMode('excel',true);
assert.equal(context.entryMode,'excel');assert(byId('input-web-editor').hidden);assert(byId('input-web-actions').hidden);
assert(!context.nativePanel.hidden,'El método Excel debe mostrar la guía de apertura y actualización');
assert.equal(editor.pending.get(id).value,'Fecha por corregir','Cambiar método no puede descartar el borrador');
assert.equal(preferences[0].key,'coverage-entry-mode');assert.equal(preferences[0].value,'excel');
context.setEntryMode('web');
assert.equal(context.entryMode,'web');assert(!byId('input-web-editor').hidden);assert(!byId('input-web-actions').hidden);
assert(context.nativePanel.hidden,'La guía Excel no debe repetir el espacio de edición web');
assert.equal(editor.pending.size,1);assert.equal(preferences.length,1,'Restaurar el método no debe sobreescribir una preferencia explícita');
''')

    def test_checkpoint_preserves_invalid_multisheet_values_and_failed_storage_keeps_entry_open(self):
        self.run_node(r'''
editor.book.kind='coverage';byId('input-number-locale').value='decimal-comma';
let committed=0,initialized=0,switched=null;
context.commitValue=()=>committed++;
context.busy=value=>editor.busy=value;
context.initialize=async()=>initialized++;
context.switchModule=value=>switched=value;
context.cleanEditor=()=>{editor.pending.clear();editor.book=null;};
context.refreshFiles=async()=>{};context.loadBook=async()=>{};
const entries=[
  {sheet:'P0_Actual',row:3,column:1,value:'¿Contamos con datos mensuales?'},
  {sheet:'P0_Actual',row:3,column:2,value:0},
  {sheet:'P0_Actual',row:3,column:3,value:null},
  {sheet:'P0_Otra',row:80,column:1,value:'2024-02-01T00:00:00'},
  {sheet:'P0_Otra',row:80,column:2,value:'20%'},
];
entries.forEach(entry=>editor.pending.set(JSON.stringify([entry.sheet,entry.row,entry.column]),entry));
const calls=[];context.api=async(path,payload)=>{calls.push({path,payload:plain(payload)});return {draft:plain(payload)};};
vm.runInContext(extract('  function compactEdits()', '  async function saved(')+
  extract('  async function openEditor(', '  function compactEdits()'),context);
async function run() {
  await context.checkpointDraft();
  assert.equal(calls.length,1);assert.equal(calls[0].path,'/api/studio/draft');
  assert.equal(calls[0].payload.fingerprint,'original');assert.equal(calls[0].payload.number_locale,'decimal-comma');
  assert.deepEqual(calls[0].payload.edits,[
    {sheet:'P0_Actual',row:3,column:1,values:[['¿Contamos con datos mensuales?',0,null]]},
    {sheet:'P0_Otra',row:80,column:1,values:[['2024-02-01T00:00:00','20%']]},
  ]);
  assert.equal(editor.pending.size,5,'Conservar un borrador no debe confundirlo con un guardado Excel válido');
  assert.equal(editor.originals.get(JSON.stringify(['P0_Actual',3,1])),'2024-01-01T00:00:00');
  context.api=async(path,payload)=>{calls.push({path,payload:plain(payload)});throw new Error('Disco lleno');};
  const opened=await context.openEditor('52_CARB_Nueva.xlsx','coverage');
  assert.equal(opened,false);assert.equal(initialized,0,'No puede avanzar a otra entrada si conservar el borrador falló');
  assert.equal(editor.file,'52_CARB_Actual.xlsx');assert.equal(editor.pending.size,5);
  assert.equal(switched,'inputs');assert.equal(calls.length,2);assert.equal(calls[1].path,'/api/studio/draft');
  assert.match(messages.at(-1).text,/No se pudo conservar el borrador/);assert.match(messages.at(-1).text,/Disco lleno/);
  assert(committed>=2,'El valor en edición se captura antes de cada intento de transición');
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

    def test_create_waits_for_invalid_draft_checkpoint_and_does_not_require_an_excel_save(self):
        self.run_node(r'''
editor.book.kind='coverage';byId('input-number-locale').value='auto';
const id=JSON.stringify(['P0_Actual',3,1]);
editor.pending.set(id,{sheet:'P0_Actual',row:3,column:1,value:'Fecha por corregir'});
for(const [field,value] of Object.entries({'create-kind':'coverage','create-country':'51',
  'create-category':'cloro','create-manufacturer':'Nuevo proyecto','create-brands':'','create-unit':'R_VOL1'}))byId(field).value=value;
context.busy=value=>editor.busy=value;
let closed=0,created=null,checkpointDone=false;
context.createDialog={close(){closed++;}};context.switchModule=()=>{};
context.saved=async(result,selectNew)=>{created={result,selectNew};};
let finishCheckpoint;const calls=[];
context.api=async(path,payload)=>{
  calls.push({path,payload:plain(payload)});
  if(path==='/api/studio/draft')return new Promise(resolve=>finishCheckpoint=()=>{checkpointDone=true;resolve({draft:plain(payload)});});
  assert.equal(path,'/api/studio/create');assert(checkpointDone,'La creación no puede adelantarse a conservar la captura previa');
  return {file:'51_BLEA_Nuevo proyecto.xlsx',kind:'coverage',name:'51_BLEA_Nuevo proyecto.xlsx'};
};
vm.runInContext(extract('  function compactEdits()', '  async function saved(')+
  extract("  byId('input-create-form').addEventListener", "  byId('input-add-sheet').addEventListener"),context);
const flush=async()=>{for(let count=0;count<10;count++)await Promise.resolve();};
async function run() {
  let prevented=false;const creation=callbacks.get('input-create-form:submit')({preventDefault(){prevented=true;}});
  await flush();assert(prevented);assert.equal(calls.length,1);assert.equal(calls[0].path,'/api/studio/draft');
  assert(editor.busy);assert(byId('create-confirm').disabled);assert.equal(closed,0);assert.equal(created,null);
  assert.equal(editor.file,'52_CARB_Actual.xlsx');assert.equal(editor.pending.get(id).value,'Fecha por corregir');
  finishCheckpoint();await creation;
  assert.equal(calls[1].path,'/api/studio/create');assert.equal(calls[1].payload.country_code,'51');
  assert.equal(calls[1].payload.category_id,'cloro');assert.equal(calls[1].payload.manufacturer,'Nuevo proyecto');
  assert.deepEqual(calls[1].payload.brands,['Nuevo proyecto']);assert.equal(closed,1);assert(created.selectNew);
  assert(!calls.some(call=>call.path==='/api/studio/save'),'La captura inválida nunca debe escribirse al Excel para permitir crear otra entrada');
  assert(!editor.busy);assert(!byId('create-confirm').disabled);
  context.api=async(path,payload)=>{calls.push({path,payload:plain(payload)});throw new Error('Disco lleno');};
  await callbacks.get('input-create-form:submit')({preventDefault(){}});
  assert.equal(calls.length,3);assert.equal(calls[2].path,'/api/studio/draft');assert.equal(closed,1);
  assert.equal(editor.pending.get(id).value,'Fecha por corregir');assert.equal(editor.file,'52_CARB_Actual.xlsx');
  assert.match(messages.at(-1).text,/Disco lleno/);assert(!editor.busy);assert(!byId('create-confirm').disabled);
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

    def test_restore_recovers_raw_values_from_all_sheets_without_changing_excel_version(self):
        self.run_node(r'''
editor.book.kind='coverage';editor.draftChecked=null;
let fetched=0,mode=null;context.setEntryMode=value=>mode=value;
const draft={file:editor.file,kind:'coverage',fingerprint:'original',sheet:'P0_Actual',row:3,column:1,
  number_locale:'decimal-comma',edits:[
    {sheet:'P0_Actual',row:3,column:1,values:[['Fecha por corregir',0,null]]},
    {sheet:'P0_Otra',row:80,column:1,values:[['2024-02-01T00:00:00','20%']]},
  ]};
context.api=async(path)=>{fetched++;const params=new URL('http://localhost'+path).searchParams;
  assert.equal(params.get('file'),editor.file);assert.equal(params.get('kind'),'coverage');
  return {draft:plain(draft),source_changed:false};};
context.URLSearchParams=URLSearchParams;
vm.runInContext(extract('  async function restoreDraft(', '  async function discardStoredDraft(')+
  extract('  function key(', '  function currentValue('),context);
async function run() {
  const notice=await context.restoreDraft(editor.book);
  assert.match(notice,/5 celda/);assert.equal(editor.pending.size,5);assert.equal(mode,'web');
  assert.equal(editor.pending.get(JSON.stringify(['P0_Actual',3,1])).value,'Fecha por corregir');
  assert.equal(editor.pending.get(JSON.stringify(['P0_Actual',3,2])).value,0);
  assert.equal(editor.pending.get(JSON.stringify(['P0_Actual',3,3])).value,null);
  assert.equal(editor.pending.get(JSON.stringify(['P0_Otra',80,2])).value,'20%');
  assert.equal(byId('input-number-locale').value,'decimal-comma');assert(byId('input-draft-details').hidden);
  assert.equal(editor.book.fingerprint,'original');assert.equal(editor.book.values[0][0],'2024-01-01T00:00:00');
  assert.equal(editor.originals.get(JSON.stringify(['P0_Actual',3,2])),123.25);
  await context.restoreDraft({...editor.book,sheet:'P0_Otra',row:80});assert.equal(fetched,1,'Cambiar de hoja no debe duplicar ni reemplazar cambios pendientes');
  editor.pending.clear();await context.restoreDraft(editor.book);assert.equal(fetched,1,'La misma versión debe consultarse una sola vez');
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

    def test_changed_source_preserves_draft_for_recovery_without_overlaying_current_excel(self):
        self.run_node(r'''
editor.book.kind='coverage';editor.book.fingerprint='new-version';editor.draftChecked=null;
context.URLSearchParams=URLSearchParams;
const draft={file:editor.file,kind:'coverage',fingerprint:'old-version',number_locale:'auto',
  edits:[{sheet:'P0_Actual',row:3,column:1,values:[['Fecha por corregir',0,null]]}]};
context.api=async()=>({draft:plain(draft),source_changed:true});
let mode=null;context.setEntryMode=value=>mode=value;
vm.runInContext(extract('  async function restoreDraft(', '  async function discardStoredDraft(')+
  extract('  function key(', '  function currentValue(')+
  extract('  function controls()', '  function letters(column)'),context);
async function run() {
  const notice=await context.restoreDraft(editor.book);
  assert.match(notice,/no se aplicó/);assert.equal(editor.pending.size,0);
  assert(!byId('input-draft-details').hidden,'El borrador anterior debe seguir accesible para recuperación');
  assert.deepEqual(plain(editor.savedDraft),draft);assert.equal(editor.book.fingerprint,'new-version');
  assert.equal(editor.book.values[0][0],'2024-01-01T00:00:00');
  assert.equal(mode,'web','La recuperación debe mostrar sus acciones sin aplicar la captura de otra versión');
  context.controls();assert(byId('input-save').disabled);
  assert(!byId('input-discard').disabled,'Un borrador anterior debe poder descartarse aunque no se haya aplicado al editor');
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

    def test_delayed_draft_response_cannot_apply_to_another_entry(self):
        self.run_node(r'''
editor.book.kind='coverage';editor.draftChecked=null;editor.request=1;
context.URLSearchParams=URLSearchParams;context.setEntryMode=()=>{};
let finish;context.api=async()=>new Promise(resolve=>finish=resolve);
vm.runInContext(extract('  async function restoreDraft(', '  async function discardStoredDraft(')+
  extract('  function key(', '  function currentValue('),context);
async function run() {
  for(const requestChanged of [true,false]) {
    editor.kind='coverage';editor.file='52_CARB_Actual.xlsx';editor.draftChecked=null;editor.savedDraft=null;
    editor.book={file:editor.file,kind:'coverage',fingerprint:'original',sheet:'P0_Actual'};
    const prior=editor.book,restoring=context.restoreDraft(prior);
    // openEditor also changes identity while awaiting refreshFiles, before loadBook increments request.
    if(requestChanged)editor.request++;
    editor.file='local/52_CARB_Investigación.xlsx';editor.kind='research';
    editor.book={file:editor.file,kind:'research',fingerprint:'research-version',sheet:'5_Otra_2'};
    finish({source_changed:false,draft:{file:prior.file,kind:'coverage',fingerprint:prior.fingerprint,
      number_locale:'auto',edits:[{sheet:'P0_Actual',row:3,column:1,values:[['Dato de otra entrada']]}]}});
    await restoring;
    assert.equal(editor.pending.size,0,'La respuesta de A no puede inyectar celdas de cobertura sobre B de 5W1H');
    assert.equal(editor.file,'local/52_CARB_Investigación.xlsx');assert.equal(editor.savedDraft ?? null,null);
    assert.equal(editor.draftChecked ?? null,null,'Una respuesta obsoleta no puede marcar la entrada actual como revisada');
  }
}
run().catch(error=>{console.error(error);process.exitCode=1;});
''')

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
let opened=0,closed=0;context.createDialog={showModal(){opened++;},close(){closed++;}};
context.mayChangeFile=()=>true;context.initialize=async()=>{};
context.categoryOptions=()=>{byId('create-category').value='cloro';};
context.createKind=async()=>{};
context.window.CoverageApp.context=()=>({loaded:true,file:{name:'51_BLEA_Otro.xlsx',
  country_code:'51',category_code:'BLEA',manufacturer:'Otro proyecto'},brandPipelines:{'Otra marca':4}});
vm.runInContext(extract('  async function showCreate(', "  byId('input-create').addEventListener"),context);
const cancel=source.match(/  byId\('create-cancel'\)\.addEventListener[^\n]+/);
assert(cancel,'No se encontró la acción real Cancelar');vm.runInContext(cancel[0],context);
async function run() {
  await context.showCreate({},'research');
  assert.equal(byId('create-country').value,'52');assert.equal(byId('create-category').value,'gaseosas');
  assert.equal(byId('create-manufacturer').value,'Investigación actual');
  assert.equal(byId('create-brands').value,'','Las marcas de otra cobertura no pertenecen a esta entrada');
  editor.book=null;await context.showCreate({},'coverage');
  assert.equal(byId('create-manufacturer').value,'','Sin archivo actual no se debe heredar otro proyecto cargado');
  await context.showCreate({file:{name:'51_BLEA_Otro.xlsx',country_code:'51',category_code:'BLEA',manufacturer:'Caller'}},'research');
  assert.equal(byId('create-manufacturer').value,'Caller');assert.equal(byId('create-brands').value,'Otra marca');
  editor.catalog.categories.push({id:'otra-descripción',code:'CARB'});
  await context.showCreate({file:{country_code:'52',category_code:'CARB',category_id:'otra-descripción',manufacturer:'Categoría explícita'}},'coverage');
  assert.equal(byId('create-category').value,'otra-descripción','El id explícito distingue categorías que comparten código');
  const before=opened;context.mayChangeFile=()=>false;
  const draftId=JSON.stringify(['P0_Actual',3,1]);
  editor.pending.set(draftId,{sheet:'P0_Actual',row:3,column:1,value:'¿Contamos con datos mensuales?'});
  await context.showCreate({},'coverage');
  assert.equal(opened,before+1,'Una fecha inválida pendiente no debe impedir abrir Nueva entrada');
  assert.equal(editor.pending.get(draftId).value,'¿Contamos con datos mensuales?',
    'Abrir el formulario conserva el valor bruto del borrador pendiente');
  assert.equal(editor.file,'52_CARB_Actual.xlsx');
  callbacks.get('create-cancel:click')();assert.equal(closed,1);
  assert.equal(editor.pending.get(draftId).value,'¿Contamos con datos mensuales?',
    'Cancelar la creación debe regresar al borrador sin eliminarlo');
}
run().catch(error=>{console.error(error);process.exitCode=1;});
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
