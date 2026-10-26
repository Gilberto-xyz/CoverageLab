'use strict';
(() => {
  const byId = id => document.getElementById(id);
  const node = (tag, text, className) => {
    const value = document.createElement(tag);
    if (text != null) value.textContent = text;
    if (className) value.className = className;
    return value;
  };
  const token = document.querySelector('meta[name="explorer-token"]').content;
  const regionalNumberLocale = new Intl.NumberFormat().formatToParts(1.1).find(part=>part.type==='decimal')?.value===',' ? 'decimal-comma' : 'decimal-dot';
  async function api(path, body) {
    const response = await fetch(path, {method:body ? 'POST' : 'GET',
      headers:{'X-Explorer-Token':token, ...(body ? {'Content-Type':'application/json'} : {})},
      ...(body ? {body:JSON.stringify(body)} : {})});
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'No se pudo completar la operación.');
    return result;
  }
  const editor = {kind:'coverage', file:null, book:null, catalog:null, files:null, researchCatalog:null, desktop:null,
    pending:new Map(), originals:new Map(), undo:[], selected:null, anchor:null, busy:false, request:0,draftChecked:null,savedDraft:null};
  let module = 'pipelines';
  let valueTarget=null,committingValue=false;
  let entryMode='web',modeChosen=false;
  try{const preference=localStorage.getItem('coverage-entry-mode');if(['excel','web'].includes(preference)){entryMode=preference;modeChosen=true;}}catch{}
  let pipelineHash = location.hash.startsWith('#excel-') ? location.hash : '#pipelines';

  byId('input-module').innerHTML = `
    <div class="studio-heading"><div><h1>Archivos de entrada</h1><p>Crea o elige el archivo y selecciona dónde prefieres llenarlo.</p></div><div class="studio-actions"><button id="input-create" class="button secondary">Nueva entrada</button><label class="button secondary" for="input-upload">Importar Excel</label><input id="input-upload" type="file" accept=".xlsx" hidden><button id="input-open-excel" class="button secondary" disabled>Abrir en Excel</button><button id="input-reload-excel" class="button secondary" disabled>Actualizar desde Excel</button><button id="input-next" class="button primary" disabled>Explorar coberturas →</button></div></div>
    <div id="input-status" class="studio-message" role="status" aria-live="polite" hidden></div>
    <section class="input-context studio-panel" aria-label="Datos del archivo seleccionado"><div id="input-file-summary" class="input-file-summary" aria-live="polite">Selecciona un archivo o crea una entrada con Nueva entrada.</div></section>
    <div class="input-capture-method"><div class="input-mode-buttons" role="group" aria-label="Método de captura"><button id="input-mode-excel" class="button secondary" aria-pressed="false">Excel de escritorio</button><button id="input-mode-web" class="button secondary" aria-pressed="true">Editar en la web</button></div><p id="input-mode-hint" class="studio-note"></p></div>
    <details id="input-draft-details" class="studio-panel padded" hidden><summary>Borrador anterior conservado</summary><p class="studio-note">El archivo cambió desde ese borrador. Puedes recuperar sus valores sin sobrescribir la versión actual.</p><pre id="input-draft-content"></pre></details>
    <div class="studio-layout"><aside class="studio-sidebar" aria-label="Archivo de entrada"><section class="studio-panel padded"><h2>Archivo y hojas</h2><label for="input-kind">Etapa de los datos</label><select id="input-kind"><option value="coverage">Cobertura · series mensuales</option><option value="research">5W1H · segmentos</option></select><label for="input-file">Archivo</label><select id="input-file"><option value="">Selecciona un archivo…</option></select><button id="input-refresh" class="text-button">Actualizar archivos</button><label for="input-sheet">Hoja</label><select id="input-sheet" disabled></select><p id="input-origin" class="studio-note"></p><button id="input-add-sheet" class="button secondary" disabled>Añadir o duplicar hoja</button></section><section class="studio-panel padded"><h2>Edición</h2><p class="studio-note">Selecciona una celda y pega un rango de Excel. Doble clic o F2 edita; Enter y Tab recorren la tabla. Mayús + clic selecciona un rango.</p><div class="studio-actions"><button id="input-paste" class="button secondary" disabled>Pegar bloque</button><button id="input-undo" class="button secondary" disabled>Deshacer</button><button id="input-discard" class="button secondary" disabled>Descartar cambios</button><button id="input-save" class="button primary" disabled>Guardar datos</button></div><p id="input-dirty" class="studio-note">Sin cambios pendientes.</p></section></aside>
    <div class="studio-panel"><div class="studio-toolbar"><label for="input-address" class="sr-only">Celda seleccionada</label><input id="input-address" class="cell-address" value="A1" aria-label="Celda seleccionada"><label for="input-cell-value" class="sr-only">Valor de la celda</label><input id="input-cell-value" class="cell-value" aria-label="Valor de la celda" placeholder="Selecciona una celda para editar su valor" disabled></div><div class="studio-toolbar"><label for="input-row">Desde fila</label><input id="input-row" type="number" min="1" max="20000" value="1"><button id="input-prev-row" class="button secondary">← Filas</button><button id="input-next-row" class="button secondary">Filas →</button><label for="input-column">Desde columna</label><input id="input-column" type="number" min="1" max="256" value="1"><button id="input-prev-column" class="button secondary">← Columnas</button><button id="input-next-column" class="button secondary">Columnas →</button><button id="input-go" class="button secondary">Ver rango</button><span id="input-range" class="studio-note"></span></div><div class="studio-toolbar"><label for="input-number-locale">Formato numérico de la fuente</label><select id="input-number-locale"><option value="auto">Automático (como Excel)</option><option value="decimal-dot">Punto decimal · 1,234.56</option><option value="decimal-comma">Coma decimal · 1.234,56</option></select><span class="studio-note">Detecta el formato del bloque y usa tu configuración regional. Cambia esta opción solo si la fuente utiliza otro formato.</span></div><div class="studio-grid-scroll"><table id="input-grid" class="studio-grid" aria-label="Celdas editables del Excel"><thead></thead><tbody></tbody></table><div id="input-empty" class="research-empty">Selecciona o crea un archivo para comenzar.</div></div><div id="input-validation" class="studio-validation"></div></div></div>`;

  const createDialog = node('dialog');
  createDialog.id='input-create-dialog';
  createDialog.innerHTML=`<h2>Nueva entrada</h2><form id="input-create-form"><div class="studio-form"><div><label for="create-kind">Formato</label><select id="create-kind"><option value="coverage">Cobertura mensual</option><option value="research">Investigación 5W1H</option></select></div><div><label for="create-country">País</label><select id="create-country" required></select></div><div class="full-field"><label for="create-category-search">Buscar categoría</label><input id="create-category-search" type="search" placeholder="Nombre, código o cesta"><label for="create-category">Categoría</label><select id="create-category" required></select></div><div class="full-field"><label for="create-manufacturer">Fabricante</label><input id="create-manufacturer" maxlength="100" required></div><div class="full-field"><label for="create-brands">Marcas · una por línea</label><textarea id="create-brands" placeholder="Si lo dejas vacío, se usará el fabricante"></textarea></div><div id="create-unit-field"><label id="create-unit-label" for="create-unit">Unidad</label><input id="create-unit" value="R_VOL1" required></div><div id="create-start-field" hidden><label for="create-start">Primer mes</label><input id="create-start" type="month"></div><div id="create-months-field" hidden><label for="create-months">Meses de historia</label><input id="create-months" type="number" min="1" max="120" value="36"></div><div id="create-segments-field" class="full-field" hidden><div class="create-segments-heading"><span id="create-segments-label">Hojas 5W1H · por marca</span><span id="create-segments-count" class="studio-note" role="status" aria-live="polite"></span></div><div class="create-segments-actions" role="group" aria-label="Seleccionar hojas 5W1H"><button id="create-segments-all" class="button secondary" type="button">Seleccionar todas</button><button id="create-segments-none" class="button secondary" type="button">Deseleccionar todas</button></div><div id="create-segments" class="research-elements" role="group" aria-labelledby="create-segments-label"></div></div></div><p id="create-note" class="studio-note">Se crearán encabezados y hojas sin observaciones. Los datos se completan en el editor.</p><div id="create-error" class="studio-message error" role="alert" hidden></div><div class="dialog-actions"><button id="create-cancel" class="button secondary" type="button">Cancelar</button><button id="create-confirm" class="button primary" type="submit">Crear entrada</button></div></form>`;
  document.body.append(createDialog);
  const createFields=createDialog.querySelector('.studio-form'),createLayout=node('div',null,'create-layout');
  createFields.classList.add('create-source-fields');createFields.before(createLayout);
  createLayout.append(createFields,byId('create-segments-field'));
  const pasteDialog = node('dialog',null,'paste-dialog');
  pasteDialog.id='input-paste-dialog';
  pasteDialog.innerHTML=`<h2>Pegar un bloque de Excel</h2><p id="paste-target"></p><textarea id="input-paste-text" aria-label="Datos separados por tabulaciones" placeholder="Pega aquí las celdas copiadas de Excel"></textarea><p class="studio-note">Se conservan las columnas y filas del bloque. Las celdas vacías quedan vacías; las fórmulas existentes están protegidas.</p><div id="paste-error" class="studio-message error" role="alert" hidden></div><div class="dialog-actions"><button id="paste-cancel" class="button secondary">Cancelar</button><button id="paste-apply" class="button primary">Aplicar al editor</button></div>`;
  document.body.append(pasteDialog);
  const sheetDialog=node('dialog');
  sheetDialog.id='input-sheet-dialog';
  sheetDialog.innerHTML=`<h2>Añadir una hoja</h2><div class="studio-form"><div class="full-field"><label for="new-sheet-name">Nombre</label><input id="new-sheet-name" maxlength="31" placeholder="Marca o nombre del segmento"></div><div class="full-field"><label for="new-sheet-source">Estructura de origen</label><select id="new-sheet-source"><option value="">Nueva hoja de cobertura vacía</option></select></div></div><div id="new-sheet-error" class="studio-message error" role="alert" hidden></div><div class="dialog-actions"><button id="new-sheet-cancel" class="button secondary">Cancelar</button><button id="new-sheet-confirm" class="button primary">Añadir hoja</button></div>`;
  document.body.append(sheetDialog);
  byId('input-grid').closest('.studio-panel').id='input-web-editor';
  byId('input-save').closest('section').id='input-web-actions';
  const nativePanel=node('section',null,'studio-panel padded input-native-panel');nativePanel.id='input-native-panel';
  nativePanel.append(node('h2','Captura en Excel'),node('p','1. Abre el archivo en Excel.\n2. Pega la información y guarda en Excel.\n3. Vuelve aquí y actualiza los datos para continuar.','studio-note'));
  const nativeActions=node('div',null,'studio-actions');nativeActions.append(byId('input-open-excel'),byId('input-reload-excel'));nativePanel.append(nativeActions);
  byId('input-web-editor').parentElement.append(nativePanel);
  function setEntryMode(mode,persist=false){
    entryMode=mode==='excel' ? 'excel' : 'web';byId('input-module').dataset.editorMode=entryMode;
    byId('input-web-editor').hidden=byId('input-web-actions').hidden=entryMode!=='web';nativePanel.hidden=entryMode!=='excel';
    for(const [id,value] of [['input-mode-excel','excel'],['input-mode-web','web']])byId(id).setAttribute('aria-pressed',String(entryMode===value));
    byId('input-mode-hint').textContent=entryMode==='excel' ? editor.pending.size ? 'Hay un borrador pendiente. Retómalo en Editar en la web o crea otra entrada; primero guarda o descarta para abrir este archivo en Excel.' : 'Usa la interfaz de Excel para pegar y guardar; vuelve a esta página para analizar.' : 'Pega y edita valores aquí. Puedes crear otra entrada conservando este borrador.';
    if(persist){modeChosen=true;try{localStorage.setItem('coverage-entry-mode',entryMode);}catch{}}
  }
  byId('input-mode-excel').addEventListener('click',()=>setEntryMode('excel',true));
  byId('input-mode-web').addEventListener('click',()=>setEntryMode('web',true));
  byId('input-draft-details').addEventListener('toggle',()=>{if(byId('input-draft-details').open)byId('input-draft-content').textContent=JSON.stringify(editor.savedDraft,null,2);});
  setEntryMode(entryMode);

  function message(text, error=false, id='input-status') {
    const element=byId(id);element.textContent=text || '';element.hidden=!text;element.classList.toggle('error',error);
  }
  function busy(value) {editor.busy=value;controls();}
  function controls() {
    const loaded=!!editor.book, dirty=editor.pending.size>0;
    window.CoverageApp.inputPending(editor.kind==='coverage' ? editor.file : null, dirty && editor.kind==='coverage');
    ['input-kind','input-file','input-sheet','input-refresh','input-create','input-go','input-prev-row','input-next-row','input-prev-column','input-next-column','input-upload','input-number-locale'].forEach(id=>byId(id).disabled=editor.busy || (id==='input-sheet' && !loaded));
    byId('input-open-excel').disabled=editor.busy || !loaded || !editor.desktop?.available;
    const numberToolbar=byId('input-number-locale').parentElement;if(numberToolbar)numberToolbar.hidden=editor.kind==='research';
    byId('input-open-excel').title=editor.desktop?.message || 'Comprobando Excel de escritorio…';
    byId('input-reload-excel').disabled=editor.busy || !loaded || dirty;
    byId('input-cell-value').disabled=editor.busy || !loaded || (editor.selected && cell(editor.selected.row,editor.selected.column)?.classList.contains('readonly'));
    byId('input-save').disabled=editor.busy || !dirty;
    byId('input-discard').disabled=editor.busy || (!dirty && byId('input-draft-details').hidden);
    byId('input-undo').disabled=editor.busy || !editor.undo.length;
    byId('input-paste').disabled=editor.busy || !loaded;
    byId('input-add-sheet').disabled=editor.busy || !loaded || dirty;
    const coverageReady=editor.book?.validation?.ready || editor.book?.validation?.sheets?.some(sheet=>sheet.ready);
    byId('input-next').disabled=editor.busy || !loaded || dirty || (editor.kind==='coverage' && !coverageReady);
    byId('input-next').textContent=editor.kind==='research' ? 'Investigar 5W1H →' : 'Explorar coberturas →';
    byId('input-dirty').textContent=dirty ? `${editor.pending.size.toLocaleString('es-MX')} celda(s) en borrador. Guarda para analizar o crea otra entrada para retomarlo después.` : 'Sin cambios pendientes.';
    byId('input-save').textContent=editor.busy ? 'Procesando…' : 'Guardar datos';
  }
  function letters(column) {let result='';for(let n=column;n>0;n=Math.floor((n-1)/26))result=String.fromCharCode(65+(n-1)%26)+result;return result;}
  function address(row,column) {return `${letters(column)}${row}`;}
  function key(sheet,row,column) {return JSON.stringify([sheet,row,column]);}
  function currentValue(sheet,row,column) {
    const id=key(sheet,row,column);
    return editor.pending.has(id) ? editor.pending.get(id).value : editor.originals.get(id) ?? null;
  }
  function cell(row,column) {return byId('input-grid').querySelector(`td[data-row="${row}"][data-column="${column}"]`);}
  function selection(row,column,extend=false,focus=false) {
    if (!editor.book) return;
    if(!committingValue && document.activeElement===byId('input-cell-value'))commitValue();
    const target=cell(row,column);if(!target)return;
    editor.selected={row,column};if(!extend || !editor.anchor)editor.anchor={row,column};
    if(!committingValue && valueTarget && (valueTarget.row!==row || valueTarget.column!==column || valueTarget.sheet!==editor.book.sheet))valueTarget=null;
    byId('input-grid').querySelectorAll('.range-selected,.cell-selected').forEach(td=>td.classList.remove('range-selected','cell-selected'));
    const firstRow=Math.min(row,editor.anchor.row),lastRow=Math.max(row,editor.anchor.row),firstCol=Math.min(column,editor.anchor.column),lastCol=Math.max(column,editor.anchor.column);
    byId('input-grid').querySelectorAll('td').forEach(td=>{const r=+td.dataset.row,c=+td.dataset.column;if(r>=firstRow && r<=lastRow && c>=firstCol && c<=lastCol)td.classList.add('range-selected');});
    target.classList.add('cell-selected');byId('input-address').value=address(row,column);
    const value=currentValue(editor.book.sheet,row,column);byId('input-cell-value').value=value==null ? '' : String(value);
    byId('input-cell-value').disabled=editor.busy || target.classList.contains('readonly');
    if(focus)target.focus();
  }
  function paintCell(td) {
    const row=+td.dataset.row,column=+td.dataset.column,id=key(editor.book.sheet,row,column);
    const value=currentValue(editor.book.sheet,row,column),r=row-editor.book.row,c=column-editor.book.column;
    const display=editor.pending.has(id) ? value : editor.book.display_values?.[r]?.[c] ?? value;
    td.textContent=display==null ? '' : String(display);
    td.title=display!==value && value!=null ? `Valor original: ${String(value)}` : '';
    td.classList.toggle('date-cell',editor.book.display_types?.[r]?.[c]==='date');
    td.classList.toggle('invalid-date',editor.book.display_types?.[r]?.[c]==='invalid-date');
    if(td.classList.contains('invalid-date'))td.title='Fecha no reconocida. Revisa el valor original antes de analizar.';
    td.classList.toggle('pending',editor.pending.has(id));
  }
  function applyBlock(row,column,values) {
    if (!editor.book || editor.busy) throw new Error('Selecciona un archivo listo para editar.');
    const width=Math.max(...values.map(line=>line.length));
    if(!values.length || !width || values.length*width>50000 || row<1 || row+values.length-1>20000 || column<1 || column+width-1>256)throw new Error('El bloque excede los límites del editor (50 000 celdas, 20 000 filas y 256 columnas).');
    values=values.map(line=>Array.from({length:width},(_,c)=>line[c] ?? null));
    values.forEach((line,r)=>line.forEach((value,c)=>{
      const td=cell(row+r,column+c);
      if(td?.classList.contains('readonly'))throw new Error(`${address(row+r,column+c)} está protegida o contiene una fórmula. El bloque no se aplicó.`);
      if(typeof value==='string' && value.trimStart().startsWith('='))throw new Error('Pega valores calculados, no fórmulas. El bloque no se aplicó.');
    }));
    const changes=[];
    values.forEach((line,r)=>line.forEach((value,c)=>{
      const id=key(editor.book.sheet,row+r,column+c);changes.push({id,previous:editor.pending.get(id)});
      const normalized=value==='' ? null : value;
      if(editor.originals.has(id) && (normalized===editor.originals.get(id) || (normalized==null && editor.originals.get(id)==null)))editor.pending.delete(id);
      else editor.pending.set(id,{sheet:editor.book.sheet,row:row+r,column:column+c,value:normalized});
      const td=cell(row+r,column+c);if(td)paintCell(td);
    }));
    editor.undo.push(changes);if(editor.undo.length>20)editor.undo.shift();
    selection(row,column,false,false);controls();
    message(`${values.length} fila(s) × ${width} columna(s) aplicadas al editor. Guarda para actualizar el archivo.`);
  }
  function parseTSV(text) {
    const rows=[];let row=[],field='',quoted=false;
    const input=text.replace(/\r\n/g,'\n').replace(/\r/g,'\n');
    for(let i=0;i<input.length;i++){
      const ch=input[i];
      if(ch==='"' && quoted && input[i+1]==='"'){field+='"';i++;}
      else if(ch==='"' && (quoted || !field)){quoted=!quoted;}
      else if(!quoted && ch==='\t'){row.push(field);field='';}
      else if(!quoted && ch==='\n'){row.push(field);rows.push(row);row=[];field='';}
      else field+=ch;
    }
    if(quoted)throw new Error('El bloque tiene comillas sin cerrar. Vuelve a copiar el rango completo.');
    if(field || row.length || !rows.length){row.push(field);rows.push(row);}
    return rows;
  }
  function beginEdit(td,replacement) {
    if(editor.busy || td.classList.contains('readonly'))return;
    td.contentEditable='true';td.dataset.editing='true';
    const value=currentValue(editor.book.sheet,+td.dataset.row,+td.dataset.column);
    td.textContent=replacement!=null ? replacement : value==null ? '' : String(value);
    td.focus();const range=document.createRange();range.selectNodeContents(td);range.collapse(false);const caret=getSelection();caret.removeAllRanges();caret.addRange(range);
  }
  function commit(td,cancel=false) {
    if(td.dataset.editing!=='true')return;
    const value=td.textContent;delete td.dataset.editing;td.contentEditable='false';
    const original=currentValue(editor.book.sheet,+td.dataset.row,+td.dataset.column);
    if(cancel || value===(original==null ? '' : String(original)))paintCell(td);else try{applyBlock(+td.dataset.row,+td.dataset.column,[[value]]);}catch(error){paintCell(td);message(error.message,true);}
  }
  function renderGrid() {
    const book=editor.book, head=byId('input-grid').tHead,body=byId('input-grid').tBodies[0];head.replaceChildren();body.replaceChildren();
    if(!book)return;
    const count=editor.kind==='coverage' ? Math.max(1,Math.min(32,8-book.column+1)) : Math.min(32,Math.max(book.values[0]?.length || 1,(book.sheets.find(s=>s.name===book.sheet)?.columns || 1)-book.column+3));
    const header=node('tr');header.append(node('th','#'));for(let c=0;c<count;c++)header.append(node('th',letters(book.column+c)));head.append(header);
    for(let r=0;r<60;r++){
      const absoluteRow=book.row+r;if(absoluteRow>20000)break;
      const tr=node('tr');tr.append(node('th',String(absoluteRow)));
      for(let c=0;c<count;c++){
        const absoluteColumn=book.column+c;if(absoluteColumn>256)break;
        const id=key(book.sheet,absoluteRow,absoluteColumn);
        editor.originals.set(id,book.values[r]?.[c] ?? null);
        const td=node('td');td.dataset.row=absoluteRow;td.dataset.column=absoluteColumn;td.tabIndex=0;
        td.setAttribute('aria-label',`${book.sheet}, ${address(absoluteRow,absoluteColumn)}`);
        if(book.readonly[r]?.[c])td.classList.add('readonly');
        const fill=book.fills?.[r]?.[c];if(typeof fill==='string' && /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(fill) && fill!=='00000000')td.style.setProperty('--source-fill',`#${fill.slice(-6)}`);
        td.addEventListener('pointerdown',event=>{if(td.dataset.editing!=='true')selection(absoluteRow,absoluteColumn,event.shiftKey,false);});
        td.addEventListener('focus',()=>{if(!td.classList.contains('cell-selected'))selection(absoluteRow,absoluteColumn,false,false);});
        td.addEventListener('dblclick',()=>beginEdit(td));
        td.addEventListener('blur',()=>commit(td));
        td.addEventListener('paste',event=>{event.preventDefault();if(td.dataset.editing==='true'){delete td.dataset.editing;td.contentEditable='false';}try{applyBlock(absoluteRow,absoluteColumn,parseTSV(event.clipboardData.getData('text/plain')));}catch(error){message(error.message,true);}});
        td.addEventListener('keydown',event=>{
          const editing=td.dataset.editing==='true';
          if(event.key==='Escape' && editing){event.preventDefault();commit(td,true);return;}
          if(['Tab','Enter'].includes(event.key)){
            event.preventDefault();commit(td);const direction=event.shiftKey ? -1 : 1;
            selection(absoluteRow+(event.key==='Enter' ? direction : 0),absoluteColumn+(event.key==='Tab' ? direction : 0),false,true);return;
          }
          if(editing)return;
          if(event.key==='F2'){event.preventDefault();beginEdit(td);return;}
          const arrows={ArrowLeft:[0,-1],ArrowRight:[0,1],ArrowUp:[-1,0],ArrowDown:[1,0]};
          if(arrows[event.key]){event.preventDefault();selection(absoluteRow+arrows[event.key][0],absoluteColumn+arrows[event.key][1],event.shiftKey,true);return;}
          if(['Delete','Backspace'].includes(event.key)){event.preventDefault();const range=selectedRange();try{applyBlock(range.row,range.column,Array.from({length:range.height},()=>Array(range.width).fill(null)));}catch(error){message(error.message,true);}return;}
          if((event.ctrlKey || event.metaKey) && event.key.toLowerCase()==='z'){event.preventDefault();undo();return;}
          if(event.key.length===1 && !event.ctrlKey && !event.metaKey && !event.altKey){event.preventDefault();beginEdit(td,event.key);}
        });
        tr.append(td);paintCell(td);
      }
      body.append(tr);
    }
    byId('input-empty').hidden=true;
    byId('input-range').textContent=`${address(book.row,book.column)}–${address(Math.min(20000,book.row+59),book.column+count-1)}`;
    const start=editor.selected && cell(editor.selected.row,editor.selected.column) ? editor.selected : {row:book.row,column:book.column};
    selection(start.row,start.column,false,false);
    controls();
  }
  function selectedRange() {
    const selected=editor.selected || {row:editor.book.row,column:editor.book.column},anchor=editor.anchor || selected;
    return {row:Math.min(selected.row,anchor.row),column:Math.min(selected.column,anchor.column),height:Math.abs(selected.row-anchor.row)+1,width:Math.abs(selected.column-anchor.column)+1};
  }
  byId('input-grid').addEventListener('copy',event=>{
    if(!editor.book || byId('input-grid').querySelector('[data-editing="true"]'))return;
    const range=selectedRange();event.preventDefault();
    const text=Array.from({length:range.height},(_,r)=>Array.from({length:range.width},(_,c)=>{
      const value=currentValue(editor.book.sheet,range.row+r,range.column+c);const string=value==null ? '' : String(value);return /[\t\n"]/.test(string) ? `"${string.replace(/"/g,'""')}"` : string;
    }).join('\t')).join('\n');event.clipboardData.setData('text/plain',text);
  });
  function undo() {
    const last=editor.undo.pop();if(!last)return;
    last.forEach(({id,previous})=>{if(previous===undefined)editor.pending.delete(id);else editor.pending.set(id,previous);});
    byId('input-grid').querySelectorAll('td').forEach(paintCell);if(editor.selected)selection(editor.selected.row,editor.selected.column);controls();
  }
  function validation() {
    const target=byId('input-validation');target.replaceChildren();const data=editor.book?.validation;if(!data)return;
    target.append(node('strong',data.ready ? 'Entradas listas para explorar.' : data.message || 'Revisa las entradas antes de continuar.'));
    const errors=data.errors || [],warnings=data.warnings || [];
    if(errors.length){const list=node('ul');errors.slice(0,12).forEach(error=>list.append(node('li',typeof error==='string' ? error : JSON.stringify(error))));target.append(list);if(errors.length>12)target.append(node('p',`Y ${errors.length-12} incidencias más.`));}
    if(warnings.length){const detail=node('details');detail.append(node('summary',`${data.warning_count || warnings.length} aviso(s)`));const list=node('ul');warnings.slice(0,30).forEach(warning=>list.append(node('li',warning)));detail.append(list);target.append(detail);}
  }
  function fillFiles() {
    const select=byId('input-file');select.replaceChildren(new Option('Selecciona un archivo…',''));
    const files=editor.files?.[editor.kind] || [];
    files.forEach(file=>{
      const match=/^(\d+)_([^_]+)_(.+)\.xlsx$/i.exec(file.name);
      const country=match && editor.catalog?.countries.find(c=>c.code===match[1]);
      const category=match && editor.catalog?.categories.find(c=>c.code.toUpperCase()===match[2].toUpperCase());
      const meta=file.metadata;
      const label=meta?.recognized ? `${meta.country} · ${meta.category} · ${meta.manufacturer}` : country && category ? `${country.name} · ${category.name.split('_')[0]} · ${match[3].replace(/_5w1h$/i,'')}` : file.name;
      select.append(new Option(`${label}${file.origin==='external' ? ' · original 5W1H' : ''}`,file.file));
    });
    if(editor.file && [...select.options].some(option=>option.value===editor.file))select.value=editor.file;
    summary();
  }
  function summary() {
    const target=byId('input-file-summary');target.replaceChildren();
    const meta=editor.book?.metadata;
    if(editor.book){
      target.append(node('strong',meta?.recognized ? `${meta.country} · ${meta.country_code} / ${meta.category} · ${meta.category_code} / ${meta.manufacturer}` : editor.book.name));
      target.append(node('span',`${editor.kind==='research' ? 'Investigación 5W1H' : 'Cobertura mensual'} · ${editor.book.name}`,'studio-note'));
    }else{
      const count=(editor.files?.[editor.kind] || []).length;
      target.append(node('span',count ? `${count} archivo(s) disponibles. Elige uno en Archivo o usa Nueva entrada para crear otro.` : 'Crea una entrada para comenzar.'));
    }
  }
  function clearBook() {
    cleanEditor();editor.file=null;editor.book=null;byId('input-file').value='';byId('input-sheet').replaceChildren();
    byId('input-grid').tHead.replaceChildren();byId('input-grid').tBodies[0].replaceChildren();byId('input-empty').hidden=false;byId('input-origin').textContent='';
    byId('input-cell-value').value='';byId('input-validation').replaceChildren();controls();summary();
  }
  async function refreshFiles() {editor.files=await api('/api/studio/files');fillFiles();}
  async function initialize() {
    if(!editor.catalog){const [catalog,files,desktop]=await Promise.all([api('/api/studio/catalog'),api('/api/studio/files'),api('/api/studio/desktop').catch(()=>({available:false,message:'Excel de escritorio no está disponible.'}))]);editor.catalog=catalog;editor.files=files;editor.desktop=desktop;if(!modeChosen)setEntryMode(desktop.available ? 'excel' : 'web');fillFiles();controls();}
  }
  async function loadBook(file=editor.file,sheet=editor.book?.sheet,row=+byId('input-row').value,column=+byId('input-column').value) {
    if(!file)return;
    const request=++editor.request;busy(true);
    try{
      const params=new URLSearchParams({file,kind:editor.kind,row,limit:Math.min(60,20001-row),column,columns:Math.min(32,257-column)});if(sheet)params.set('sheet',sheet);
      let book=await api(`/api/studio/workbook?${params}`);
      if(!sheet && editor.kind==='research' && !/^[1-8][-_]/.test(book.sheet)){
        const firstDataSheet=book.sheets.find(entry=>/^[1-8][-_]/.test(entry.name));
        if(firstDataSheet){params.set('sheet',firstDataSheet.name);book=await api(`/api/studio/workbook?${params}`);}
      }
      if(request!==editor.request)return;
      if(editor.pending.size && editor.book?.fingerprint!==book.fingerprint)throw new Error('El archivo cambió fuera del editor. Conserva el bloque pendiente, descarta y vuelve a cargar la versión actual.');
      editor.file=book.file;editor.book=book;
      const draftMessage=await restoreDraft(book,request);if(request!==editor.request)return;
      summary();
      byId('input-sheet').replaceChildren(...book.sheets.map(sheet=>new Option(`${sheet.name} · ${sheet.rows} filas × ${sheet.columns} columnas`,sheet.name)));
      byId('input-sheet').value=book.sheet;byId('input-file').value=book.file;
      byId('input-origin').textContent=book.copy_on_save || book.origin==='external' ? 'Original de 5W1H. Al guardar se creará una copia editable dentro de CoverageLab.' : `Archivo local · ${book.name}`;
      byId('input-row').value=book.row;byId('input-column').value=book.column;
      renderGrid();validation();message(draftMessage);
    }catch(error){message(error.message,true);}
    finally{if(request===editor.request)busy(false);}
  }
  function cleanEditor() {editor.pending.clear();editor.originals.clear();editor.undo=[];editor.selected=null;editor.anchor=null;valueTarget=null;editor.draftChecked=null;editor.savedDraft=null;byId('input-draft-details').hidden=true;}
  async function openEditor(file,kind='coverage',sheet=null) {
    commitValue();busy(true);
    try{await checkpointDraft();await initialize();cleanEditor();editor.kind=kind;editor.file=file;editor.book=null;
      byId('input-kind').value=kind;byId('input-row').value=1;byId('input-column').value=1;
      await refreshFiles();switchModule('inputs');await loadBook(file,sheet,1,1);return true;
    }catch(error){switchModule('inputs');message(error.message,true);return false;}finally{busy(false);}
  }
  function compactEdits() {
    const entries=[...editor.pending.values()].sort((a,b)=>a.sheet.localeCompare(b.sheet)||a.row-b.row||a.column-b.column);
    const runs=[];
    entries.forEach(entry=>{const last=runs.at(-1);if(last && last.sheet===entry.sheet && last.row===entry.row && last.column+last.values[0].length===entry.column)last.values[0].push(entry.value);else runs.push({sheet:entry.sheet,row:entry.row,column:entry.column,values:[[entry.value]]});});
    const blocks=[],tails=new Map();
    runs.forEach(run=>{const id=JSON.stringify([run.sheet,run.column,run.values[0].length]),last=tails.get(id);if(last && last.row+last.values.length===run.row)last.values.push(run.values[0]);else{blocks.push(run);tails.set(id,run);}});
    return blocks;
  }
  async function checkpointDraft(){
    commitValue();if(!editor.book || !editor.pending.size)return;
    try{
      const result=await api('/api/studio/draft',{file:editor.file,kind:editor.kind,fingerprint:editor.book.fingerprint,sheet:editor.book.sheet,row:editor.book.row,column:editor.book.column,number_locale:byId('input-number-locale').value,edits:compactEdits()});
      editor.savedDraft=result.draft;
    }catch(error){throw new Error(`No se pudo conservar el borrador. La entrada actual sigue abierta: ${error.message}`);}
  }
  async function restoreDraft(book,request=editor.request){
    const identity=JSON.stringify([book.kind,book.file,book.fingerprint]);
    if(editor.pending.size || editor.draftChecked===identity)return '';
    const result=await api(`/api/studio/draft?${new URLSearchParams({file:book.file,kind:book.kind})}`);
    if(request!==editor.request || editor.book!==book || editor.file!==book.file || editor.kind!==book.kind)return '';
    editor.draftChecked=identity;editor.savedDraft=result.draft;byId('input-draft-details').hidden=true;
    if(!result.draft)return '';
    if(result.source_changed || result.draft.fingerprint!==book.fingerprint){byId('input-draft-details').hidden=false;setEntryMode('web');return 'El archivo cambió desde el borrador anterior. Ese borrador sigue conservado y no se aplicó sobre esta versión.';}
    for(const block of result.draft.edits)block.values.forEach((line,r)=>line.forEach((value,c)=>editor.pending.set(key(block.sheet,block.row+r,block.column+c),{sheet:block.sheet,row:block.row+r,column:block.column+c,value})));
    byId('input-number-locale').value=result.draft.number_locale;setEntryMode('web');
    return `Borrador recuperado: ${editor.pending.size} celda(s) pendientes. Puedes corregirlas, guardarlas o crear otra entrada para retomarlas después.`;
  }
  async function discardStoredDraft(file=editor.file,kind=editor.kind){
    if(file)await api('/api/studio/discard-draft',{file,kind});
  }
  async function saved(result,selectNew=false,refreshAnalysis=true) {
    const previousSheet=editor.book?.sheet;
    cleanEditor();editor.file=result.file;editor.book=result;editor.kind=result.kind;
    byId('input-kind').value=editor.kind;
    await refreshFiles();byId('input-file').value=result.file;
    const selectedSheet=!selectNew && result.sheets.some(entry=>entry.name===previousSheet) ? previousSheet : selectNew && result.kind==='research' && !/^[1-8][-_]/.test(result.sheet) ? result.sheets.find(entry=>/^[1-8][-_]/.test(entry.name))?.name || result.sheet : result.sheet;
    await loadBook(result.file,selectedSheet,selectNew ? 1 : +byId('input-row').value,selectNew ? 1 : +byId('input-column').value);
    if(refreshAnalysis) {
      if(editor.kind==='coverage')await window.CoverageApp.sourceChanged(result.file);
      else await window.StudioResearch?.edited(result.file);
    }
    message(result.copied_from ? `Copia guardada: ${result.name}. El original sigue intacto.` : `${result.saved_cells || 0} celda(s) guardadas. Los análisis se recalculan con esta versión.`);
  }
  async function save() {
    if(!editor.book || !editor.pending.size)return;
    busy(true);
    try{const result=await api('/api/studio/save',{file:editor.file,kind:editor.kind,fingerprint:editor.book.fingerprint,edits:compactEdits(),number_locale:byId('input-number-locale').value,number_locale_hint:regionalNumberLocale});await discardStoredDraft().catch(()=>{});await saved(result);}
    catch(error){message(error.message,true);}finally{busy(false);}
  }
  byId('input-save').addEventListener('click',save);
  byId('input-open-excel').addEventListener('click',async()=>{
    commitValue();if(!editor.book)return;
    busy(true);message('Abriendo el libro en Excel de escritorio…');
    try{
      if(editor.pending.size) {
        const written=await api('/api/studio/save',{file:editor.file,kind:editor.kind,fingerprint:editor.book.fingerprint,edits:compactEdits(),preserve_text:true});
        await discardStoredDraft().catch(()=>{});await saved(written,false,false);
      }
      const result=await api('/api/studio/open-excel',{file:editor.file,kind:editor.kind,fingerprint:editor.book.fingerprint,sheet:editor.book.sheet,row:editor.book.row,column:editor.book.column,limit:Math.min(60,20001-editor.book.row),columns:Math.min(32,257-editor.book.column)});
      if(result.copied_from)await saved(result);
      message(`${result.copied_from ? 'Se creó una copia local; el original sigue intacto. ' : ''}Excel está abierto. Pega y guarda allí; después pulsa Actualizar desde Excel para continuar aquí.`);
    }catch(error){message(error.message,true);}finally{busy(false);}
  });
  byId('input-reload-excel').addEventListener('click',async()=>{
    commitValue();if(!editor.book)return;
    if(editor.pending.size)return message('Guarda o descarta los cambios del editor web antes de actualizar desde Excel.',true);
    busy(true);message('Leyendo la versión guardada en Excel…');
    try{
      const result=await api('/api/studio/reload-excel',{file:editor.file,kind:editor.kind,sheet:editor.book.sheet,row:editor.book.row,column:editor.book.column,limit:Math.min(60,20001-editor.book.row),columns:Math.min(32,257-editor.book.column)});
      await saved(result);message('Datos guardados en Excel cargados. Revisa los períodos y continúa con el análisis.');
    }catch(error){message(error.message,true);}finally{busy(false);}
  });
  byId('input-discard').addEventListener('click',async()=>{busy(true);try{await discardStoredDraft();cleanEditor();await loadBook();message('Borrador descartado. Se cargó la versión actual del archivo.');}catch(error){message(error.message,true);}finally{busy(false);}});
  byId('input-undo').addEventListener('click',undo);
  byId('input-file').addEventListener('change',async()=>{const file=byId('input-file').value;commitValue();busy(true);try{await checkpointDraft();cleanEditor();editor.book=null;editor.file=file;byId('input-row').value=1;byId('input-column').value=1;if(file)await loadBook(file,null,1,1);else clearBook();}catch(error){byId('input-file').value=editor.file || '';message(error.message,true);}finally{busy(false);}});
  byId('input-kind').addEventListener('change',async()=>{const kind=byId('input-kind').value;commitValue();busy(true);try{await checkpointDraft();cleanEditor();editor.kind=kind;clearBook();fillFiles();validation();}catch(error){byId('input-kind').value=editor.kind;message(error.message,true);}finally{busy(false);}});
  byId('input-sheet').addEventListener('change',()=>{editor.selected=null;byId('input-row').value=1;byId('input-column').value=1;loadBook(editor.file,byId('input-sheet').value,1,1);});
  byId('input-refresh').addEventListener('click',()=>refreshFiles().catch(error=>message(error.message,true)));
  byId('input-go').addEventListener('click',()=>loadBook());
  [['input-prev-row','input-row',-60],['input-next-row','input-row',60],['input-prev-column','input-column',-32],['input-next-column','input-column',32]].forEach(([id,field,delta])=>byId(id).addEventListener('click',()=>{byId(field).value=Math.max(1,Math.min(field==='input-row' ? 20000 : 256,+byId(field).value+delta));editor.selected=null;loadBook();}));
  byId('input-address').addEventListener('keydown',event=>{
    if(event.key!=='Enter' || !editor.book)return;event.preventDefault();
    const match=/^([A-Z]{1,3})([1-9]\d*)$/.exec(byId('input-address').value.trim().toUpperCase());
    if(!match)return message('Escribe una referencia como A3 o AF20.',true);
    const column=[...match[1]].reduce((total,ch)=>total*26+ch.charCodeAt(0)-64,0),row=+match[2];
    if(row>20000 || column>256)return message('La referencia excede los límites del editor.',true);
    if(cell(row,column))selection(row,column,false,true);else{byId('input-row').value=row;byId('input-column').value=column;editor.selected={row,column};loadBook();}
  });
  function commitValue() {
    if(committingValue || !valueTarget || !editor.book || valueTarget.sheet!==editor.book.sheet || editor.busy)return;
    const value=byId('input-cell-value').value,original=currentValue(valueTarget.sheet,valueTarget.row,valueTarget.column);
    if(value===(original==null ? '' : String(original)))return;
    committingValue=true;
    try{applyBlock(valueTarget.row,valueTarget.column,[[value]]);}catch(error){message(error.message,true);}finally{committingValue=false;}
  }
  byId('input-cell-value').addEventListener('keydown',event=>{if(event.key!=='Enter')return;event.preventDefault();commitValue();});
  byId('input-cell-value').addEventListener('focus',()=>{valueTarget=editor.selected ? {...editor.selected,sheet:editor.book.sheet} : null;});
  byId('input-cell-value').addEventListener('change',commitValue);
  byId('input-paste').addEventListener('click',()=>{const selected=editor.selected || {row:editor.book.row,column:editor.book.column};byId('paste-target').textContent=`Se pegará desde ${editor.book.sheet}!${address(selected.row,selected.column)}.`;byId('input-paste-text').value='';message('',false,'paste-error');pasteDialog.showModal();});
  byId('paste-cancel').addEventListener('click',()=>pasteDialog.close());
  byId('paste-apply').addEventListener('click',()=>{try{if(!byId('input-paste-text').value)throw new Error('Pega primero las celdas de la fuente.');const selected=editor.selected || {row:editor.book.row,column:editor.book.column};applyBlock(selected.row,selected.column,parseTSV(byId('input-paste-text').value));pasteDialog.close();}catch(error){message(error.message,true,'paste-error');}});

  async function researchCatalog() {if(!editor.researchCatalog)editor.researchCatalog=await api('/api/research/catalog');return editor.researchCatalog;}
  function categoryOptions() {
    // Normaliza solo la búsqueda; las etiquetas conservan ñ, tildes y diéresis.
    const searchKey=value=>String(value).normalize('NFD').replace(/\p{M}/gu,'').toLocaleLowerCase();
    const query=searchKey(byId('create-category-search').value.trim()),selected=byId('create-category').value;
    const categories=editor.catalog.categories.filter(c=>searchKey(`${c.name} ${c.code} ${c.basket}`).includes(query));
    byId('create-category').replaceChildren(...categories.map(category=>new Option(`${category.name} · ${category.code} · ${category.basket}`,category.id)));
    if([...byId('create-category').options].some(option=>option.value===selected))byId('create-category').value=selected;
  }
  const segmentIconPaths={
    trend:['M3 19h18','m5 14 4-4 4 3 6-8'],
    tree:['M12 3v6','M5 14V9h14v5','M12 9v5','M3 15h4v5H3z','M10 15h4v5h-4z','M17 15h4v5h-4z'],
    size:['M4 8h16v10H4z','M8 8v4','M12 8v6','M16 8v4'],
    brand:['M4 4h10l6 6-10 10-6-6z','M9 8h.01'],
    flavor:['M12 3c-2 4-7 8-7 12a7 7 0 0 0 14 0c0-4-5-8-7-12Z','M9 15a3 3 0 0 0 3 3'],
    layers:['m12 3 9 5-9 5-9-5 9-5Z','m3 12 9 5 9-5','m3 16 9 5 9-5'],
    category:['M3 3h7v7H3z','M14 3h7v7h-7z','M3 14h7v7H3z','M14 14h7v7h-7z'],
    people:['M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2','M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8','M17 4a4 4 0 0 1 0 7','M22 21v-2a4 4 0 0 0-3-4'],
    map:['m3 6 6-3 6 3 6-3v15l-6 3-6-3-6 3V6Z','M9 3v15','M15 6v15'],
    channels:['M4 4h16v5H4z','M4 15h6v5H4z','M14 15h6v5h-6z','M12 9v3','M7 15v-3h10v3'],
    modern:['M5 21V3h14v18','M9 7h1','M14 7h1','M9 11h1','M14 11h1','M10 21v-6h4v6','M3 21h18'],
    store:['m3 9 2-6h14l2 6','M3 9a3 3 0 0 0 6 0 3 3 0 0 0 6 0 3 3 0 0 0 6 0','M5 12v9h14v-9','M9 21v-6h6v6'],
    players:['M3 20h18','M4 20v-8h5v8','M9 20V5h6v15','M15 20V9h5v11'],
    price:['M4 4h10l6 6-10 10-6-6z','M8 8h.01','m11 12 3-3','m11 15 5-5'],
    distribution:['M3 6h12v12H3z','M15 10h4l3 4v4h-7','M7 21a2 2 0 1 0 0-4 2 2 0 0 0 0 4','M18 21a2 2 0 1 0 0-4 2 2 0 0 0 0 4'],
    confidence:['M4 6v12','M20 6v12','M4 12h16','M12 8v8'],
  };
  function createSegmentIcon(kind) {
    const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
    Object.entries({viewBox:'0 0 24 24','aria-hidden':'true',focusable:'false'}).forEach(([key,value])=>svg.setAttribute(key,value));
    (segmentIconPaths[kind] || segmentIconPaths.category).forEach(d=>{
      const path=document.createElementNS('http://www.w3.org/2000/svg','path');path.setAttribute('d',d);svg.append(path);
    });
    return svg;
  }
  function renderCreateSegments(segments) {
    const target=byId('create-segments');target.replaceChildren();
    const defaults=new Set(['1','2','3-1','3-2','4','5-1','5-2','6','7','8']);
    const icons={'1':'trend','2':'tree','3-1':'size','3-2':'brand','3-3':'flavor','3-4':'layers','3-5':'category',
      '4':'people','5-1':'map','5-2':'channels','5-3':'modern','5-4':'store','6':'players','6-1':'price','7':'distribution','8':'confidence'};
    const groups=[
      {id:'context',title:'Evolución y compradores',description:'Cuándo, por qué y quiénes'},
      {id:'product',title:'Producto',description:'Qué se compra'},
      {id:'place',title:'Lugar de compra',description:'Regiones y canales'},
      {id:'market',title:'Mercado y precisión',description:'Competencia, distribución y confianza'},
    ];
    const groupFor=id=>['1','2','4'].includes(id) ? 'context' : id.startsWith('3-') ? 'product' : id.startsWith('5-') ? 'place' : 'market';
    groups.forEach(group=>{
      const items=segments.filter(segment=>groupFor(segment.id)===group.id);if(!items.length)return;
      const section=node('section',null,'create-segment-group');section.dataset.group=group.id;
      const heading=node('h3',group.title);heading.id=`create-segment-${group.id}`;
      section.setAttribute('aria-labelledby',heading.id);section.append(heading,node('p',group.description));
      items.forEach(segment=>{
        const label=node('label',null,'create-segment-option'),input=node('input');
        const fullLabel=segment.label || segment.name;
        input.type='checkbox';input.value=segment.id;input.dataset.segmentOrder=String(segments.indexOf(segment));input.checked=defaults.has(segment.id);input.setAttribute('aria-label',fullLabel);
        const icon=node('span',null,'create-segment-icon');icon.append(createSegmentIcon(icons[segment.id]));
        const name=fullLabel.includes(' · ') ? fullLabel.split(' · ').slice(1).join(' · ') : fullLabel;
        label.append(icon,node('span',name.charAt(0).toLocaleUpperCase()+name.slice(1),'create-segment-name'),input);
        section.append(label);
      });
      target.append(section);
    });
    updateCreateSegments();
  }
  function updateCreateSegments() {
    const inputs=[...byId('create-segments').querySelectorAll('input[type="checkbox"]')];
    const selected=inputs.filter(input=>input.checked).length;
    byId('create-segments-count').textContent=`${selected} de ${inputs.length} seleccionadas`;
    byId('create-segments-all').disabled=!inputs.length || selected===inputs.length;
    byId('create-segments-none').disabled=selected===0;
    const research=byId('create-kind').value==='research';
    const needsUnit=!research || inputs.some(input=>input.value==='2' && input.checked);
    byId('create-unit-field').hidden=!needsUnit;
    byId('create-unit').disabled=!needsUnit;
    byId('create-unit').required=needsUnit;
    byId('create-unit-label').textContent=research ? 'Unidad del árbol' : 'Unidad';
  }
  function selectCreateSegments(checked) {
    byId('create-segments').querySelectorAll('input[type="checkbox"]').forEach(input=>{input.checked=checked;});
    updateCreateSegments();
  }
  byId('create-segments-all').addEventListener('click',()=>selectCreateSegments(true));
  byId('create-segments-none').addEventListener('click',()=>selectCreateSegments(false));
  byId('create-segments').addEventListener('change',updateCreateSegments);
  async function createKind() {
    const research=byId('create-kind').value==='research';
    createDialog.classList.toggle('research-create',research);
    ['create-start-field','create-months-field','create-segments-field'].forEach(id=>byId(id).hidden=!research);
    byId('create-unit').value=research ? 'K' : 'R_VOL1';
    byId('create-unit').title=research ? 'Solo para Árbol de medidas: K=kg, L=litros, U=unidades, T=toneladas, R=rollos, M=metros, H=hojas' : 'Nombre de la unidad en la fuente, por ejemplo R_VOL1 o VOL1_P';
    updateCreateSegments();
    if(research){const catalog=await researchCatalog();renderCreateSegments(catalog.segments);}
  }
  async function showCreate(context={},kind=editor.kind) {
    commitValue();
    await initialize();byId('create-kind').value=kind;
    byId('create-country').replaceChildren(...editor.catalog.countries.map(country=>new Option(`${country.name} · ${country.code}`,country.code)));
    byId('create-category-search').value='';categoryOptions();
    const file=context.file || editor.book?.metadata;
    if(file){if(file.country_code)byId('create-country').value=file.country_code;const category=editor.catalog.categories.find(c=>c.id===file.category_id) || editor.catalog.categories.find(c=>c.code===file.category_code);if(category)byId('create-category').value=category.id;}
    byId('create-manufacturer').value=file?.manufacturer || '';
    const coverage=window.CoverageApp.context();
    const sameContext=context.file && coverage.loaded && context.file.name===coverage.file?.name;
    byId('create-brands').value=kind==='research' && sameContext ? Object.keys(coverage.brandPipelines).join('\n') : '';
    byId('create-start').value=`${new Date().getFullYear()-2}-01`;
    await createKind();message('',false,'create-error');createDialog.showModal();
  }
  byId('input-create').addEventListener('click',()=>showCreate().catch(error=>message(error.message,true)));
  byId('create-kind').addEventListener('change',()=>createKind().catch(error=>message(error.message,true,'create-error')));
  byId('create-category-search').addEventListener('input',categoryOptions);
  byId('create-cancel').addEventListener('click',()=>createDialog.close());
  byId('input-create-form').addEventListener('submit',async event=>{
    event.preventDefault();if(editor.busy)return;busy(true);byId('create-confirm').disabled=true;
    try{
      await checkpointDraft();
      const manufacturer=byId('create-manufacturer').value.trim();
      const spec={kind:byId('create-kind').value,country_code:byId('create-country').value,category_id:byId('create-category').value,manufacturer,
        brands:byId('create-brands').value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean),unit:byId('create-unit').value.trim()};
      if(!spec.brands.length)spec.brands=[manufacturer];
      if(spec.kind==='research')Object.assign(spec,{start_date:byId('create-start').value,months:+byId('create-months').value,segments:[...byId('create-segments').querySelectorAll('input:checked')].sort((a,b)=>Number(a.dataset.segmentOrder)-Number(b.dataset.segmentOrder)).map(input=>input.value)});
      const result=await api('/api/studio/create',spec);createDialog.close();switchModule('inputs');await saved(result,true);
      message(`Entrada creada: ${result.name}. Selecciona ${spec.kind==='coverage' ? 'A3' : 'la primera celda de datos'} y pega los datos reales.`);
    }catch(error){message(error.message,true,'create-error');}finally{busy(false);byId('create-confirm').disabled=false;}
  });
  byId('input-add-sheet').addEventListener('click',()=>{
    byId('new-sheet-name').value='';const options=editor.kind==='coverage' ? [new Option('Nueva hoja de cobertura vacía','')] : [];
    byId('new-sheet-source').replaceChildren(...options,...editor.book.sheets.map(sheet=>new Option(`Duplicar ${sheet.name}`,sheet.name)));
    message('',false,'new-sheet-error');sheetDialog.showModal();
  });
  byId('new-sheet-cancel').addEventListener('click',()=>sheetDialog.close());
  byId('new-sheet-confirm').addEventListener('click',async()=>{
    busy(true);byId('new-sheet-confirm').disabled=true;
    try{const result=await api('/api/studio/add-sheet',{file:editor.file,kind:editor.kind,fingerprint:editor.book.fingerprint,name:byId('new-sheet-name').value,source_sheet:byId('new-sheet-source').value || null});sheetDialog.close();await saved(result,true);message(`Hoja ${result.sheet} añadida.`);}
    catch(error){message(error.message,true,'new-sheet-error');}finally{busy(false);byId('new-sheet-confirm').disabled=false;}
  });
  byId('input-upload').addEventListener('change',async()=>{
    if(editor.busy)return;
    const upload=byId('input-upload').files[0];if(!upload)return;
    commitValue();
    busy(true);
    try{
      if(upload.size>10*1024*1024)throw new Error('El archivo excede el límite de 10 MiB.');
      const bytes=new Uint8Array(await upload.arrayBuffer());let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));
      await checkpointDraft();const result=await api('/api/studio/import',{kind:editor.kind,filename:upload.name,content:btoa(binary)});await saved(result,true);message(`Archivo importado: ${result.name}.`);
    }catch(error){message(error.message,true);}finally{busy(false);byId('input-upload').value='';}
  });
  byId('input-next').addEventListener('click',async()=>{
    if(editor.pending.size)return message('Guarda los cambios antes de continuar.',true);
    busy(true);
    try{if(editor.kind==='coverage'){await window.CoverageApp.openFile(editor.file);switchModule('pipelines');}else{switchModule('research');await window.StudioResearch.open(editor.file);}}
    catch(error){message(error.message,true);}finally{busy(false);}
  });
  async function activateInputs() {
    try{await initialize();if(!editor.book && !editor.file){const context=window.CoverageApp.context();if(context.file){editor.kind='coverage';editor.file=context.file.name;byId('input-kind').value='coverage';fillFiles();await loadBook(editor.file,null,1,1);}}}
    catch(error){message(error.message,true);}
  }
  function switchModule(next, updateURL=true) {
    if(!['inputs','pipelines','research'].includes(next))return;
    if(module==='pipelines' && location.hash.startsWith('#excel-'))pipelineHash=location.hash;
    module=next;document.body.dataset.module=next;
    byId('pipeline-layout').hidden=next!=='pipelines';byId('input-module').hidden=next!=='inputs';byId('research-module').hidden=next!=='research';
    document.querySelectorAll('.module-nav [data-module]').forEach(button=>{const active=button.dataset.module===next;button.classList.toggle('active',active);if(active)button.setAttribute('aria-current','page');else button.removeAttribute('aria-current');});
    if(updateURL)history.replaceState(null,'',next==='pipelines' ? pipelineHash : `#${next}`);
    window.CoverageApp.geometry();
    if(next==='pipelines')window.CoverageApp.redraw();
    if(next==='inputs')activateInputs();
    if(next==='research')window.StudioResearch?.activate(window.CoverageApp.context());
  }
  document.querySelectorAll('.module-nav [data-module]').forEach(button=>button.addEventListener('click',()=>switchModule(button.dataset.module)));
  byId('return-inputs').addEventListener('click',()=>switchModule('inputs'));
  window.addEventListener('beforeunload',event=>{if(editor.pending.size){event.preventDefault();event.returnValue='';}});
  window.StudioResearch?.init(byId('research-module'),{api,openEditor,createResearch:context=>showCreate(context,'research'),coverageContext:()=>window.CoverageApp.context(),hasPendingInput:file=>editor.kind==='research' && editor.file===file && editor.pending.size>0,openChart:(...args)=>window.CoverageApp.openChart(...args)});
  controls();switchModule(location.hash==='#inputs' ? 'inputs' : location.hash==='#research' ? 'research' : 'pipelines',false);
  window.StudioApp={openEditor,switchModule};
})();
