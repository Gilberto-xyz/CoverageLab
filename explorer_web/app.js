'use strict';
const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name="explorer-token"]').content;
const state = {analysis: null, active: null, selections: {}, included: {}, pptExtras: {}, busy: false, dirty: false, job: null, files: [], openingResults:false, openingOutput:null, resultRequest:0, pendingInput:null};
const outputNames = {template:'Abrir Excel', presentation:'Abrir PPT', bank:'Banco', pipeline_report:'Pipelines', selection:'Selección', log:'Registro'};
const fmt = (value, digits = 1, suffix = '') => value == null ? '—' : new Intl.NumberFormat('es-MX', {minimumFractionDigits: digits, maximumFractionDigits: digits}).format(value) + suffix;
const month = date => new Intl.DateTimeFormat('es-MX', {month: 'short', year: '2-digit'}).format(new Date(date + 'T12:00:00'));
const item = (tag, text, className) => { const element = document.createElement(tag); if (text != null) element.textContent = text; if (className) element.className = className; return element; };
const resultIconPaths = {
  template:['M5 3h10l4 4v14H5z','M15 3v5h4','M8 11h8v7H8z','M8 14h8','M11 11v7'],
  presentation:['M3 4h18v13H3z','M12 17v4','M8 21h8','M7 9h4v4H7z','M14 9h3','M14 12h3'],
  bank:['M4 6c0-2 16-2 16 0s-16 2-16 0','M4 6v12c0 2 16 2 16 0V6','M4 12c0 2 16 2 16 0'],
  pipeline_report:['M4 3h5v5H4z','M15 16h5v5h-5z','M4 16h5v5H4z','M6.5 8v8','M6.5 11h11v5'],
  selection:['M9 4H5v17h14V4h-4','M9 2h6v4H9z','m8 13 2 2 5-5'],
  log:['M5 3h10l4 4v14H5z','M15 3v5h4','M8 12h8','M8 16h8'],
  folder:['M3 7V4h6l2 3h10v13H3z','M3 10h18'],
};
function resultButton(button, label, kind) {
  const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
  Object.entries({viewBox:'0 0 24 24','aria-hidden':'true',focusable:'false',class:`output-icon output-icon-${kind}`}).forEach(([key,value])=>svg.setAttribute(key,value));
  (resultIconPaths[kind] || resultIconPaths.log).forEach(d=>{
    const path=document.createElementNS('http://www.w3.org/2000/svg','path');path.setAttribute('d',d);svg.append(path);
  });
  button.replaceChildren(svg,item('span',label));
}
const fold = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase();
const driverMetrics = {
  penetration: {title: 'Penetración mensual', unit: '%', digits: 1},
  buyers_series: {title: 'Compradores', unit: 'compradores', digits: 0},
  purchase_average: {title: 'Compra media', unit: 'por comprador', digits: 2},
  purchase_occasion: {title: 'Compra por ocasión', unit: 'por ocasión', digits: 2},
  frequency: {title: 'Frecuencia de compra', unit: 'ocasiones por comprador', digits: 2},
};

async function api(path, body) {
  const response = await fetch(path, {method: body ? 'POST' : 'GET', headers: {'X-Explorer-Token': token, ...(body ? {'Content-Type': 'application/json'} : {})}, ...(body ? {body: JSON.stringify(body)} : {})});
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || 'No se pudo completar la operación.');
  return result;
}
function status(message, error = false) {
  $('status').hidden = !message;
  $('status').textContent = message || '';
  $('status').classList.toggle('error', error);
  $('status').setAttribute('role', error ? 'alert' : 'status');
  syncChromeGeometry();
}
function syncChromeGeometry() {
  const root=document.documentElement;
  root.style.setProperty('--header-height',`${$('header-main').getBoundingClientRect().height}px`);
  root.style.setProperty('--results-height',`${$('result-dock').getBoundingClientRect().height}px`);
  root.style.setProperty('--top-height',`${$('app-top').getBoundingClientRect().height}px`);
}
function setHeaderCollapsed(collapsed, persist = true) {
  const generateContainer = $(collapsed ? 'header-actions' : 'result-heading');
  if ($('generate-files').parentElement !== generateContainer) generateContainer.append($('generate-files'));
  $('app-top').classList.toggle('header-collapsed', collapsed);
  $('result-dock').hidden = collapsed;
  $('toggle-header').setAttribute('aria-expanded', String(!collapsed));
  $('toggle-header').setAttribute('aria-label', collapsed ? 'Mostrar encabezado' : 'Contraer encabezado');
  $('header-toggle-label').textContent = collapsed ? 'Mostrar encabezado' : 'Contraer encabezado';
  $('toggle-header').title = collapsed ? 'Mostrar archivo actual y resultados' : 'Ocultar archivo actual y resultados para dar más espacio a los gráficos';
  if (persist) {
    try { localStorage.setItem('coveragelab-header-collapsed', String(collapsed)); } catch (_) { /* El control funciona sin almacenamiento local. */ }
  }
  syncChromeGeometry();
}
function requestOptions() {
  return {file: $('file').value, template: $('template').value, trend_axis: $('trend-axis').value,
    include_english: $('english').checked, round_coverage: $('round').checked, animate_trend_pipeline: $('animate').checked};
}
function selectedSheets() {
  return Object.fromEntries(Object.entries(state.selections).filter(([name]) => state.included[name]));
}
function payload() { return {analysis_id: state.analysis.id, selections: selectedSheets(),
  ppt_extras:Object.fromEntries(Object.keys(selectedSheets()).filter(name=>state.pptExtras[name]?.length).map(name=>[name,state.pptExtras[name]]))}; }
function draftKey() { return `coveragelab-selection-${state.analysis.selection_key}`; }
function saveDraft() {
  try { localStorage.setItem(draftKey(), JSON.stringify({selections: state.selections, included: state.included,ppt_extras:state.pptExtras})); } catch (_) { /* Disk save remains available. */ }
}
function controls() {
  const working = state.busy || state.job?.status === 'running';
  const pendingInput=state.pendingInput===$('file').value;
  ['file', 'template', 'trend-axis', 'english', 'round', 'animate', 'refresh-files', 'load', 'browse-files'].forEach(id => { $(id).disabled = working; });
  $('load').disabled=working || pendingInput;
  $('input-pending-notice').hidden=!pendingInput;
  const ready = state.analysis && !state.dirty && Object.keys(selectedSheets()).length && !working && !pendingInput;
  $('save').disabled = !ready;
  const refreshNeeded=!state.analysis || state.dirty;
  $('generate-files').disabled = working || pendingInput || !$('file').value || (!refreshNeeded && !Object.keys(selectedSheets()).length);
  $('generate-files').textContent=state.job?.status==='running' ? 'Generando…' : state.busy ? 'Procesando…' : 'Generar archivos';
  $('generate-files').title = pendingInput ? 'Guarda los cambios del editor de entradas antes de generar' : refreshNeeded ? 'Cargar los datos y revisar antes de crear una nueva entrega' : !Object.keys(selectedSheets()).length ? 'Incluye al menos una hoja para generar' : `Generar una nueva entrega con ${$('template').selectedOptions[0].textContent} y los pipelines elegidos`;
  const outputsReady=state.job?.file===$('file').value && state.job?.status==='complete';
  $('open-folder').disabled = !outputsReady || working || !!state.openingResults || !!state.openingOutput;
  resultButton($('open-folder'),state.openingResults ? 'Abriendo…' : 'Abrir carpeta','folder');
  $('open-folder').title = outputsReady ? `Abrir esta entrega: ${state.job.directory}` : 'Genera una entrega para habilitar sus resultados';
  document.querySelectorAll('[data-output]').forEach(button=>{
    const label=button.dataset.output, available=outputsReady && state.job.available_outputs.includes(label);
    button.disabled=!available || working || !!state.openingResults || !!state.openingOutput;
    resultButton(button,state.openingOutput===label ? 'Abriendo…' : outputNames[label],label);
    button.title=available ? `Abrir ${state.job.outputs.find(out=>out.label===label)?.name || (label==='selection' ? 'selection.json' : 'generation.log')}` : 'Se habilita cuando exista este archivo en una entrega completada';
  });
  $('job-log-details').hidden = !state.job || state.job.file!==$('file').value;
  $('select-all').disabled = working;
  $('select-none').disabled = working;
  document.querySelectorAll('.sheet-row input, .pipeline-button, .ppt-extra-input').forEach(control => { control.disabled = working || control.dataset.unavailable === 'true'; });
  const active = state.analysis?.sheets.find(sheet => sheet.name === state.active);
  $('use-recommended').disabled = working || !active?.pipelines?.[active.recommended]?.available;
}
async function refreshFiles() {
  let previous = $('file').value;
  if(!previous)try{previous=localStorage.getItem('coveragelab-source-file') || '';}catch(_){/* Sigue disponible sin almacenamiento local. */}
  const result = await api('/api/files');
  if (!Array.isArray(result.details)) throw new Error('El servidor local usa una versión anterior. Cierra su terminal y vuelve a abrir Abrir_Explorador.cmd para actualizar los datos.');
  state.files = result.details;
  $('file').replaceChildren();
  if (!result.files.length) $('file').append(new Option('No hay archivos Excel en el proyecto', ''));
  else state.files.forEach(file => $('file').append(new Option(file.error ? `${file.name} · Revisar nombre` : `${file.country} · ${file.short_category || file.category} · ${file.manufacturer}`, file.name)));
  if (result.files.includes(previous)) $('file').value = previous;
  if (!result.files.length) status('Coloca un .xlsx en la carpeta del proyecto y pulsa Actualizar archivos.');
  ['file-country', 'file-category'].forEach((id, index) => {
    const current = $(id).value;
    const key = index ? 'category' : 'country';
    $(id).replaceChildren(new Option(index ? 'Todas las categorías' : 'Todos los países', ''));
    [...new Set(state.files.map(file => file[key]).filter(Boolean))].sort((a,b) => a.localeCompare(b, 'es')).forEach(value => $(id).append(new Option(value, value)));
    if ([...$(id).options].some(option => option.value === current)) $(id).value = current;
  });
  renderFileDetail(); renderFileCards(); controls();
  await refreshJobForFile();
}
function fact(label, value) { const div = item('div', null, 'fact'); div.append(item('small', label), item('strong', value || '—')); return div; }
function paintExcelCell(cell, kind, value, sheet, monthIndex = null) {
  const colors=state.analysis?.table_colors;if(!colors)return;
  let style;
  if(kind==='correlation')style=CoverageDashboard.correlationStyle(value,sheet.template.correlations.flatMap(row=>row.pipelines),colors.correlation);
  if(kind==='variation' && value!=null && Number.isFinite(value))style=value>0 ? colors.positive : value<0 ? colors.negative : null;
  if(kind==='coverage') {
    if(monthIndex===sheet.dates.length-1)style={fill:colors.coverage_current,ink:'#172c3c'};
    else if(monthIndex===sheet.dates.length-13)style={fill:colors.coverage_previous,ink:'#172c3c'};
  }
  if(!style)return;
  cell.classList.add('excel-colored');cell.style.setProperty('--cell-bg',style.fill);cell.style.setProperty('--cell-ink',style.ink);
  if(kind==='correlation')cell.title=`Pearson r=${fmt(value,3)} · Color relativo: mínimo ${fmt(style.low,3)}, mediana ${fmt(style.mid,3)}, máximo ${fmt(style.high,3)}`;
}
function renderColorKey(root, kinds) {
  root.replaceChildren();const colors=state.analysis?.table_colors;root.hidden=!colors;if(!colors)return;
  const keys=[];
  if(kinds.includes('correlation'))keys.push(['Correlación menor',colors.correlation[0]],['Mediana',colors.correlation[1]],['Correlación mayor',colors.correlation[2]]);
  if(kinds.includes('variation'))keys.push(['Disminuye (<0)',colors.negative.fill],['Aumenta (>0)',colors.positive.fill]);
  if(kinds.includes('coverage'))keys.push(['Cobertura · corte actual',colors.coverage_current],['Cobertura · hace 12 meses',colors.coverage_previous]);
  keys.forEach(([label,color])=>{const chip=item('span',label);chip.style.setProperty('--key-color',color);root.append(chip);});
  if(kinds.includes('correlation')) {
    root.append(item('small','Escala relativa a toda la tabla de correlaciones, como en Excel.'));
    const sheet=state.analysis.sheets.find(entry=>entry.name===state.active);
    const values=sheet?.template?.correlations.flatMap(row=>row.pipelines).filter(value=>value!=null&&Number.isFinite(value)) || [];
    if(values.length && Math.max(...values)<0)root.append(item('small','Todos los coeficientes disponibles son negativos, incluso los de color verde.'));
  }
}
function renderFileDetail() {
  const file = state.files.find(entry => entry.name === $('file').value);
  const panel = $('file-detail'); panel.replaceChildren(); panel.hidden = !file;
  if (!file) return;
  try{localStorage.setItem('coveragelab-source-file',file.name);}catch(_){/* La selección funciona durante esta sesión. */}
  const identity = item('div', null, 'file-identity');
  identity.append(item('strong', file.manufacturer || 'Nombre por revisar', 'file-manufacturer'), item('div', file.name, 'file-original-name'));
  panel.append(identity);
  if (file.error) panel.append(item('p', file.error, 'file-warning'));
  else {
    const facts = item('div', null, 'file-facts');
    facts.append(fact('País', `${file.country} · ${file.country_code}`), fact('Categoría', `${file.short_category || file.category} · ${file.category_code}`), fact('Cesta', file.basket));
    panel.append(facts);
    if (file.descriptor) panel.append(item('p', file.descriptor, 'source-note'));
    if (file.per_sheet) panel.append(item('p', 'Categorías por hoja: el detalle se resuelve al cargar.', 'file-warning'));
    const extra=item('details',null,'file-extra');
    extra.append(item('summary','Detalles del archivo'),item('p',`${fmt(file.size / 1024, 0)} KB · Modificado ${new Intl.DateTimeFormat('es-MX', {dateStyle:'medium'}).format(new Date(file.modified))}`,'file-size'));
    panel.append(extra);
  }
  if (state.analysis && state.analysis.file !== file.name) panel.append(item('p', 'Archivo seleccionado. Pulsa Cargar y explorar para cambiar los datos visibles.', 'file-warning'));
}
function renderFileCards() {
  const search = fold($('file-search').value);
  const matches = state.files.filter(file => (!search || fold(Object.values(file).join(' ')).includes(search)) && (!$('file-country').value || file.country === $('file-country').value) && (!$('file-category').value || file.category === $('file-category').value));
  $('file-cards').replaceChildren();
  $('catalog-count').textContent = `${matches.length} de ${state.files.length} archivos`;
  matches.forEach(file => {
    const card = item('button', null, `file-card${file.name === $('file').value ? ' selected' : ''}`);
    card.setAttribute('aria-pressed', String(file.name === $('file').value));
    card.append(item('span', file.error ? 'Requiere corrección' : `${file.country} · ${file.country_code}`, 'file-country-tag'), item('strong', file.manufacturer || file.name), item('span', file.error || `${file.category} · ${file.category_code}`, 'file-category-name'), item('small', file.name));
    if (file.per_sheet) card.append(item('span', 'Varias categorías · detalle por hoja', 'table-tag'));
    card.addEventListener('click', () => { $('file').value = file.name; markDirty(); renderFileDetail(); renderFileCards(); $('file-dialog').close(); refreshJobForFile().catch(error=>status(error.message,true)); });
    $('file-cards').append(card);
  });
  if (!matches.length) $('file-cards').append(item('p', 'No hay archivos con estos filtros.', 'section-description'));
}
function markDirty() {
  if (state.analysis) { state.dirty = true; status('La configuración cambió. Cargar y explorar o Generar archivos actualizará los datos antes de la revisión final.'); }
  controls();
}
async function load() {
  if (!$('file').value) return status('Selecciona un archivo Excel.', true);
  state.busy = true; state.analysis = null; state.dirty = false;
  $('workspace').hidden = true; $('sheet-panel').hidden = true; $('empty').hidden = false;
  status('Cargando hojas y calculando P0–P6…'); controls();
  try {
    const analysis = await api('/api/analyze', requestOptions());
    if (analysis.sheets.some(sheet => !sheet.error && !sheet.dashboard)) throw new Error('Reinicia el explorador local para cargar el panel completo del Excel.');
    state.analysis = analysis; state.selections = {}; state.included = {}; state.pptExtras = {};
    $('window').value = 'all'; $('dashboard-window').value = 'all';
    $('template-view').value = 'all'; $('table-columns').value = 'all';
    renderFileDetail();
    let saved;
    try { saved = JSON.parse(localStorage.getItem(draftKey()) || 'null'); } catch (_) { saved = null; }
    if (!saved && analysis.saved_selection) saved = {selections: analysis.saved_selection.selections,ppt_extras:analysis.saved_selection.ppt_extras,
      included: Object.fromEntries(Object.keys(analysis.saved_selection.selections).map(name => [name, true]))};
    analysis.sheets.forEach(sheet => {
      if (sheet.error) { state.included[sheet.name] = false; return; }
      const previous = saved?.selections?.[sheet.name];
      state.selections[sheet.name] = Number.isInteger(previous) && sheet.pipelines[previous]?.available ? previous : sheet.default_pipeline;
      state.included[sheet.name] = saved ? !!saved.included?.[sheet.name] : true;
      const extras=saved?.ppt_extras?.[sheet.name];
      state.pptExtras[sheet.name] = Array.isArray(extras) ? [...new Set(extras.filter(key=>Object.hasOwn(driverMetrics,key)))] : [];
    });
    state.active = analysis.sheets.find(sheet => !sheet.error)?.name || analysis.sheets[0]?.name;
    $('empty').hidden = true; $('workspace').hidden = false; $('sheet-panel').hidden = false;
    $('search').value = '';
    const invalid = analysis.sheets.filter(sheet => sheet.error).length;
    status(invalid ? `${invalid} hoja(s) requieren corrección y quedan excluidas.` : 'Panel del Excel listo.' + (saved ? ' Selección recuperada.' : ''));
    render();
  } catch (error) { status(error.message, true); }
  finally { state.busy = false; controls(); }
}
function renderSheets() {
  const filter = $('search').value.toLocaleLowerCase();
  $('sheets').replaceChildren();
  let visible = 0;
  for (const sheet of state.analysis.sheets) {
    if (!`${sheet.name} ${sheet.brand} ${sheet.category || ''}`.toLocaleLowerCase().includes(filter)) continue;
    visible++;
    const row = item('div', null, `sheet-row${sheet.name === state.active ? ' active' : ''}${!state.included[sheet.name] ? ' excluded' : ''}${sheet.error ? ' error' : ''}`);
    const checkbox = item('input'); checkbox.type = 'checkbox'; checkbox.checked = !!state.included[sheet.name];
    checkbox.disabled = !!sheet.error; checkbox.dataset.unavailable = String(!!sheet.error);
    checkbox.setAttribute('aria-label', `Incluir ${sheet.name} en los entregables`);
    checkbox.addEventListener('change', () => { state.included[sheet.name] = checkbox.checked; saveDraft(); render(); });
    const button = item('button', null, 'sheet-button'); button.title = sheet.name;
    button.append(item('strong', sheet.brand), item('small', sheet.error ? 'Requiere corrección' : `${sheet.category} · ${sheet.months} meses`));
    button.addEventListener('click', () => { state.active = sheet.name; render(); });
    row.append(checkbox, button, item('span', sheet.error ? '!' : `P${state.selections[sheet.name]}`, 'sheet-pipeline'));
    $('sheets').append(row);
  }
  if (!visible) $('sheets').append(item('p', 'No hay hojas que coincidan.', 'source-note'));
  $('sheet-count').textContent = state.analysis.sheets.length;
}
function choose(pipeline) {
  const sheet = state.analysis.sheets.find(sheet => sheet.name === state.active);
  if (state.job?.status === 'running' || state.busy || !sheet.pipelines[pipeline]?.available) return;
  state.selections[state.active] = pipeline; saveDraft(); render();
  $('pipeline-buttons').querySelectorAll('button')[pipeline].focus({preventScroll: true});
}
function render() {
  if (!state.analysis) return;
  renderSheets();
  const sheet = state.analysis.sheets.find(sheet => sheet.name === state.active);
  const count = Object.keys(selectedSheets()).length;
  $('included-count').textContent = `${count} hojas incluidas`;
  $('export-summary').textContent = count ? `${count} hoja(s) con un pipeline elegido` : 'Incluye al menos una hoja para generar';
  $('context').textContent = `${sheet?.metadata?.pais_nombre || state.analysis.country} / ${state.analysis.manufacturer} / ${sheet?.category || state.analysis.category}`;
  $('sheet-facts').replaceChildren(fact('País del archivo', `${state.analysis.country} · ${state.analysis.country_code}`), fact('Categoría de la hoja', `${sheet?.metadata?.categoria_nombre || state.analysis.category} · ${sheet?.metadata?.categoria_codigo || state.analysis.category_code}`), fact('Cesta', sheet?.metadata?.cesta_nombre || state.analysis.basket), fact('Unidad de volumen', sheet?.unit), fact('Último mes', sheet?.dates ? month(sheet.dates.at(-1)) : '—'));
  $('brand-title').textContent = sheet?.brand || 'Sin hojas';
  $('sheet-subtitle').textContent = sheet ? `${sheet.name}${sheet.dates ? ` · ${month(sheet.dates[0])} — ${month(sheet.dates.at(-1))} · ${sheet.months} meses` : ''}` : '';
  $('warnings').hidden = !sheet?.warnings?.length;
  $('warnings').textContent = sheet?.warnings?.join('\n') || '';
  $('sheet-error').hidden = !sheet?.error;
  $('sheet-error').textContent = sheet?.error || '';
  $('sheet-content').hidden = !sheet || !!sheet.error;
  if (!sheet || sheet.error) { controls(); return; }
  const selected = state.selections[sheet.name];
  $('pipeline-buttons').replaceChildren();
  sheet.pipelines.forEach(p => {
    const button = item('button', `P${p.pipeline}`, `pipeline-button${p.pipeline === selected ? ' selected' : ''}${p.pipeline === sheet.recommended ? ' recommended' : ''}`);
    button.dataset.unavailable = String(!p.available); button.disabled = !p.available;
    button.setAttribute('aria-pressed', String(p.pipeline === selected));
    button.title = p.available ? `Elegir P${p.pipeline}: desplazar Sell-in ${p.pipeline} mes(es)` : 'Historia insuficiente para cobertura MAT';
    button.addEventListener('click', () => choose(p.pipeline)); $('pipeline-buttons').append(button);
  });
  $('chosen-label').textContent = `P${selected}`;
  $('recommendation-title').textContent = `AUTO Balanceado: P${sheet.recommended} · Mayor correlación: P${sheet.correlation_pipeline}`;
  $('recommendation-reason').textContent = sheet.reason;
  const comparison = $('compare').value;
  $('compare').replaceChildren(new Option('Sin comparación', ''));
  sheet.pipelines.forEach(p => { if (p.available && p.pipeline !== selected) $('compare').append(new Option(`Pipeline ${p.pipeline}`, String(p.pipeline))); });
  if ([...$('compare').options].some(option => option.value === comparison)) $('compare').value = comparison;
  const p = sheet.pipelines[selected];
  $('metric-corr').textContent = fmt(p.correlation, 3);
  $('metric-corr-note').textContent = p.correlation == null ? 'Sin base suficiente o sin variación' : 'Pearson entre las curvas MAT';
  $('metric-cov').textContent = fmt(p.coverage, 1, '%');
  $('metric-cov-note').textContent = `Estabilidad: ${fmt(p.stability, 1, ' pp')} vs. MAT anterior`;
  $('metric-var').textContent = fmt(p.client_variation, 1, '%');
  $('metric-var-note').textContent = `Worldpanel: ${fmt(p.wp_variation, 1, '%')}`;
  $('metric-gap').textContent = fmt(p.gap, 1, ' pp');
  $('pipeline-table').replaceChildren();
  sheet.pipelines.forEach(pipeline => {
    const row = item('tr', null, `${pipeline.pipeline === selected ? 'selected' : ''} ${!pipeline.available ? 'unavailable' : ''}`);
    row.setAttribute('aria-selected', String(pipeline.pipeline === selected));
    if (pipeline.available) { row.tabIndex = 0; row.addEventListener('click', () => choose(pipeline.pipeline)); row.addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); choose(pipeline.pipeline); } }); }
    const labels = [];
    if (pipeline.pipeline === sheet.recommended) labels.push('AUTO Balanceado');
    if (pipeline.pipeline === sheet.forced) labels.push('Prefijo de hoja');
    if (pipeline.outlier) labels.push('Pico reciente');
    if (!pipeline.available) labels.push('Historia insuficiente');
    [ `P${pipeline.pipeline}`, fmt(pipeline.correlation, 3), fmt(pipeline.coverage, 1, '%'), fmt(pipeline.client_variation, 1, '%'), fmt(pipeline.gap), fmt(pipeline.stability), labels.join(' · ') || (pipeline.pipeline === 0 ? 'Sin desplazamiento' : '—') ].forEach((value, index) => {
      const cell=item('td',value,index===6?'table-tag':'');
      if(index===1)paintExcelCell(cell,'correlation',pipeline.correlation,sheet);
      if(index===3)paintExcelCell(cell,'variation',pipeline.client_variation,sheet);
      if(index===5)paintExcelCell(cell,'variation',pipeline.stability,sheet);
      row.append(cell);
    });
    $('pipeline-table').append(row);
  });
  renderColorKey($('pipeline-color-key'),['correlation','variation']);
  drawChart(); controls();
}

function renderTemplate() {
  const sheet = state.analysis?.sheets.find(entry => entry.name === state.active);
  if (!sheet?.template) return;
  const selected = state.selections[sheet.name], pipeline = sheet.pipelines[selected];
  const template = sheet.template, averages = template.averages;
  $('template-reference').textContent = `Corte ${month(sheet.dates.at(-1))} · ${sheet.months} meses · Selecciona un indicador para verlo en el gráfico.`;
  $('template-pipeline').textContent = `Sell-in P${selected}`;
  const cards = [
    ['penetration', 'Penetración · último mes', sheet.penetration.at(-1), `MAT anterior: ${fmt(averages.Penet_MAT_Anterior, 1, '%')}`, '%', 1],
    ['buyers_series', 'Compradores · último mes', sheet.buyers_series.at(-1), `MAT anterior: ${fmt(averages.Buyers_MAT_Anterior, 0)}${averages.Buyers_MAT_Actual < 200 ? ' · Base baja' : ''}`, '', 0],
    ['purchase_average', 'Compra media · último mes', sheet.purchase_average.at(-1), `${sheet.unit} por comprador`, '', 2],
    ['purchase_occasion', 'Compra por ocasión · último mes', sheet.purchase_occasion.at(-1), `${sheet.unit} por ocasión`, '', 2],
    ['frequency', 'Frecuencia · último mes', sheet.frequency.at(-1), 'Ocasiones por comprador', '', 2],
  ];
  $('purchase-metrics').replaceChildren(...cards.map(([key, label, value, note, suffix, digits]) => {
    const card = item('button', null, 'purchase-card');
    card.append(item('span', label), item('strong', fmt(value, digits, suffix)), item('small', note));
    CoverageDashboard.kpiAverages(card,sheet.dashboard.kpis.find(kpi=>kpi.key===key));
    card.append(item('span', 'Ver tendencia ↗', 'metric-action'));
    card.addEventListener('click', () => { $('metric').value = key; drawChart(); $('chart-title').scrollIntoView({block:'start'}); }); return card;
  }));
  const inspect = $('inspect-month').value;
  $('inspect-month').replaceChildren(...sheet.dates.map(date => new Option(month(date), date)));
  $('inspect-month').value = sheet.dates.includes(inspect) ? inspect : sheet.dates.at(-1);
  const view = $('template-view').value, monthly = view === 'monthly' || view === 'all';
  renderColorKey($('template-color-key'),monthly ? ['coverage','variation'] : view==='correlations' ? ['correlation'] : view==='variations' ? ['variation'] : []);
  $('template-color-key').hidden=view==='averages';
  $('template-table').className = `template-${monthly ? 'monthly' : view}`;
  ['table-columns', 'table-columns-label', 'inspect-month', 'inspect-month-label', 'monthly-inspector'].forEach(id => { $(id).hidden = !monthly; });
  $('template-head').replaceChildren(); $('template-body').replaceChildren();
  const columns = [];
  let rows = [];
  if (monthly) {
    const allColumns = [
      ['Mes', i => month(sheet.dates[i]), 'date'],
      ['Worldpanel', i => fmt(sheet.sell_out[i]), 'volume'],
      ['Sell-in original P0', i => fmt(sheet.sell_in[i]), 'volume'],
      [`Sell-in P${selected}`, i => fmt(pipeline.sell_in[i]), 'volume'],
      ['MAT Worldpanel', i => fmt(sheet.wp_mat[i]), 'volume'],
      ['MAT Sell-in original P0', i => fmt(sheet.pipelines[0].sell_in_mat[i]), 'volume'],
      [`MAT Sell-in P${selected}`, i => fmt(pipeline.sell_in_mat[i]), 'volume'],
      [`Cobertura P${selected}`, i => fmt(pipeline.coverage_series[i], 1, '%'), 'volume'],
      ['YoY MAT Worldpanel', i => fmt(sheet.wp_yoy[i], 1, '%'), 'volume'],
      [`YoY MAT Sell-in P${selected}`, i => fmt(pipeline.yoy[i], 1, '%'), 'volume'],
      ['Penetración mensual', i => fmt(sheet.penetration[i], 1, '%'), 'shopping'],
      ['Compradores', i => fmt(sheet.buyers_series[i], 0), 'shopping'],
      ['Compra media', i => fmt(sheet.purchase_average[i], 2), 'shopping'],
      ['Compra por ocasión', i => fmt(sheet.purchase_occasion[i], 2), 'shopping'],
      ['Frecuencia', i => fmt(sheet.frequency[i], 2), 'shopping'],
      ...sheet.pipelines.filter(p => p.pipeline !== selected).map(p => [`Cobertura P${p.pipeline}`, i => fmt(p.coverage_series[i],1,'%'),'volume']),
      ['Año', i => sheet.dates[i].slice(0,4),'volume'],
      ['Trimestre', i => String(Math.ceil(Number(sheet.dates[i].slice(5,7))/3)),'volume'],
      ['Semestre', i => Number(sheet.dates[i].slice(5,7)) <= 6 ? '1' : '2','volume'],
    ];
    columns.push(...allColumns.filter(([, , group]) => group === 'date' || $('table-columns').value === 'all' || group === $('table-columns').value));
    const first = $('window').value === 'all' ? 0 : Math.max(0, sheet.dates.length - Number($('window').value));
    rows = sheet.dates.slice(first).map((date, offset) => ({cells: columns.map(([, get]) => get(first + offset)), date}));
    const i = sheet.dates.indexOf($('inspect-month').value);
    const source = i - selected;
    $('monthly-inspector').replaceChildren(fact('Mes observado', month(sheet.dates[i])), fact('Worldpanel', fmt(sheet.sell_out[i])), fact(`Sell-in P${selected}`, fmt(pipeline.sell_in[i])), fact('Mes de origen Sell-in', source >= 0 ? month(sheet.dates[source]) : 'Sin historia'), fact('Penetración', fmt(sheet.penetration[i], 1, '%')), fact('Compradores', fmt(sheet.buyers_series[i], 0)));
    $('template-note').textContent = `La tabla usa el selector de período y conserva los meses originales, incluso sin base MAT. Volumen: ${sheet.unit}. Haz clic en un mes para ver su detalle. Los hábitos de compra conservan el mes del Excel. Los YoY MAT siguen las columnas de evolución del template (denominador 1 si el MAT anterior es cero).`;
  } else if (view === 'variations') {
    columns.push(['Comparación'], ['Worldpanel'], ...sheet.pipelines.map(p => [`Sell-in P${p.pipeline}`]));
    rows = template.variations.map(row => ({cells: [row.label, fmt(row.worldpanel, 1, '%'), ...row.pipelines.map(value => fmt(value, 1, '%'))], numeric:[null,row.worldpanel,...row.pipelines], periods: [row.wp_period, ...row.pipeline_periods]}));
    $('template-note').textContent = 'Ventanas idénticas a las fórmulas del template: semestre y trimestre vs. período anterior o mismo período del año anterior. Pasa sobre un valor para ver las fechas exactas. — indica historia insuficiente, ceros en alguna ventana o denominador cero; el template también lo omite.';
  } else if (view === 'correlations') {
    columns.push(['Ventana MAT'], ...sheet.pipelines.map(p => [`P${p.pipeline}`]));
    rows = template.correlations.map(row => ({cells: [row.label, ...row.pipelines.map(value => fmt(value, 3))],numeric:[null,...row.pipelines]}));
    $('template-note').textContent = 'Correlación Pearson entre las series MAT, respetando el desplazamiento de Sell-in. — indica falta de historia o una serie sin variación. Se destaca el pipeline elegido.';
  } else {
    columns.push(['Indicador'], ['Últimos 12 meses'], ['12 meses anteriores'], ['Cambio']);
    const pen = averages.Penet_MAT_Actual, penPrev = averages.Penet_MAT_Anterior, buyers = averages.Buyers_MAT_Actual, buyersPrev = averages.Buyers_MAT_Anterior;
    rows = [{cells: ['Penetración media mensual', fmt(pen, 1, '%'), fmt(penPrev, 1, '%'), pen != null && penPrev != null ? fmt(pen - penPrev, 1, ' pp') : '—']}, {cells: ['Compradores promedio', fmt(buyers, 0), fmt(buyersPrev, 0), buyers != null && buyersPrev != null ? fmt(buyers - buyersPrev, 0) : '—']}, {cells: ['Frecuencia promedio mensual', fmt(averages.Freq_MAT_Actual, 2), '—', '—']}];
    $('template-note').textContent = 'Promedios mensuales de ventanas completas de 12 meses. La penetración media MAT es un promedio; no representa penetración anual de compradores únicos. Los compradores no se suman. Frecuencia promedio se calcula con el motor de CoverageLab.';
  }
  const header = item('tr'); columns.forEach(([label], index) => { const th = item('th', label); th.scope = 'col'; if ((view === 'variations' && index === selected + 2) || (view === 'correlations' && index === selected + 1)) th.className = 'template-selected'; header.append(th); }); $('template-head').append(header);
  rows.forEach(row => {
    const tr = item('tr');
    if (monthly) {
      tr.classList.toggle('selected', row.date === $('inspect-month').value); tr.tabIndex = 0;
      const inspectRow = () => { $('inspect-month').value = row.date; renderTemplate(); };
      tr.addEventListener('click', inspectRow); tr.addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); inspectRow(); } });
    }
    row.cells.forEach((value, index) => {
      const td = item('td', value);
      if ((view === 'variations' && index === selected + 2) || (view === 'correlations' && index === selected + 1)) td.className = 'template-selected';
      if(view==='variations' && index>0)paintExcelCell(td,'variation',row.numeric[index],sheet);
      if(view==='correlations' && index>0)paintExcelCell(td,'correlation',row.numeric[index],sheet);
      if(monthly) {
        const i=sheet.dates.indexOf(row.date),label=columns[index][0],cov=label.match(/^Cobertura P([0-6])$/);
        if(cov)paintExcelCell(td,'coverage',sheet.pipelines[Number(cov[1])].coverage_series[i],sheet,i);
        if(label==='YoY MAT Worldpanel')paintExcelCell(td,'variation',sheet.wp_yoy[i],sheet);
        if(label.startsWith('YoY MAT Sell-in'))paintExcelCell(td,'variation',pipeline.yoy[i],sheet);
      }
      const periods = row.periods?.[index - 1];
      if (periods) td.title = `${periods.current.map(month).join(' — ')} vs. ${periods.previous.map(month).join(' — ')}`;
      if (view === 'variations' && index === 0) {
        const wpPeriod = row.periods[0], chosenPeriod = row.periods[selected + 1];
        if (wpPeriod) td.append(item('small', `WP: ${wpPeriod.current.map(month).join(' — ')} / ${wpPeriod.previous.map(month).join(' — ')}`, 'period-detail'));
        if (chosenPeriod) td.append(item('small', `P${selected}: ${chosenPeriod.current.map(month).join(' — ')} / ${chosenPeriod.previous.map(month).join(' — ')}`, 'period-detail'));
      }
      tr.append(td);
    }); $('template-body').append(tr);
  });
  renderCompleteTables(sheet, selected, view === 'all');
}

function renderCompleteTables(sheet, selected, visible) {
  const root = $('template-extra-tables'); root.replaceChildren(); root.hidden = !visible;
  if (!visible) return;
  const avg = sheet.template.averages;
  const descriptions = [
    {key:'variations',title:'Variaciones por período · P0–P6',columns:['Comparación','Worldpanel',...sheet.pipelines.map(p=>`Sell-in P${p.pipeline}`)],chosen:selected+2,
      rows:sheet.template.variations.map(row=>({cells:[row.label,fmt(row.worldpanel,1,'%'),...row.pipelines.map(v=>fmt(v,1,'%'))],numeric:[null,row.worldpanel,...row.pipelines],periods:[row.wp_period,...row.pipeline_periods]})),
      note:'Las ventanas con ceros o historia insuficiente aparecen como —, igual que las fórmulas del Excel. Se muestran las fechas exactas de Worldpanel y del pipeline elegido.'},
    {key:'correlations',title:'Correlaciones MAT · P0–P6',columns:['Ventana',...sheet.pipelines.map(p=>`P${p.pipeline}`)],chosen:selected+1,
      rows:sheet.template.correlations.map(row=>({cells:[row.label,...row.pipelines.map(v=>fmt(v,3))],numeric:[null,...row.pipelines]})),
      note:'Últimos 12 meses, año anterior y ventana de 24 meses. Se conserva el desplazamiento de Sell-in y la regla de base insuficiente.'},
    {key:'averages',title:'Promedios MAT',columns:['Indicador','Últimos 12 meses','12 meses anteriores'],
      rows:[{cells:['Penetración media mensual',fmt(avg.Penet_MAT_Actual,1,'%'),fmt(avg.Penet_MAT_Anterior,1,'%')]},{cells:['Compradores promedio',fmt(avg.Buyers_MAT_Actual,0),fmt(avg.Buyers_MAT_Anterior,0)]},{cells:['Frecuencia promedio',fmt(avg.Freq_MAT_Actual,2),'—']}],
      note:'Promedios mensuales; los compradores no se suman y la penetración media MAT no equivale a penetración anual de compradores únicos.'},
    {key:'annual-mat',title:'Cortes anuales MAT · volúmenes, YoY y cobertura',columns:['Corte MAT','Worldpanel',`Sell-in P${selected}`,'YoY Worldpanel',`YoY Sell-in P${selected}`,`Cobertura P${selected}`],chosen:2,
      rows:sheet.dashboard.mat_periods.map(row=>({cells:[month(row.date),fmt(row.worldpanel),fmt(row.sell_in[selected]),fmt(row.worldpanel_yoy,1,'%'),fmt(row.sell_in_yoy[selected],1,'%'),fmt(row.coverage[selected],1,'%')],numeric:[null,row.worldpanel,row.sell_in[selected],row.worldpanel_yoy,row.sell_in_yoy[selected],row.coverage[selected]],date:row.date})),
      note:'Mismo mes de corte en todos los años. Cada volumen exige 12 meses completos; el YoY anual queda vacío si falta el año previo o su volumen es cero.'},
  ];
  descriptions.forEach(data=>{
    const section=item('section',null,'complete-table-section');section.dataset.detail=data.key;section.id=`excel-table-${data.key}`;
    const heading=item('div',null,'table-heading');heading.append(item('h3',data.title));section.append(heading);
    if(data.key!=='averages') { const key=item('div',null,'excel-color-key');renderColorKey(key,data.key==='correlations'?['correlation']:data.key==='variations'?['variation']:['coverage','variation']);section.append(key); }
    const wrapper=item('div',null,'table-scroll template-scroll'),table=item('table',null,`template-${data.key}`),thead=item('thead'),head=item('tr');
    data.columns.forEach((label,i)=>{const th=item('th',label);th.scope='col';if(i===data.chosen)th.className='template-selected';head.append(th);});thead.append(head);table.append(thead);
    const tbody=item('tbody');
    data.rows.forEach(row=>{
      const tr=item('tr');row.cells.forEach((value,i)=>{
        const td=item('td',value);if(i===data.chosen)td.className='template-selected';
        if(data.key==='correlations' && i>0)paintExcelCell(td,'correlation',row.numeric[i],sheet);
        if(data.key==='variations' && i>0)paintExcelCell(td,'variation',row.numeric[i],sheet);
        if(data.key==='annual-mat' && [3,4].includes(i))paintExcelCell(td,'variation',row.numeric[i],sheet);
        if(data.key==='annual-mat' && i===5)paintExcelCell(td,'coverage',row.numeric[i],sheet,sheet.dates.indexOf(row.date));
        const period=row.periods?.[i-1];if(period)td.title=`${period.current.map(month).join(' — ')} vs. ${period.previous.map(month).join(' — ')}`;
        if(i===0 && row.periods) [0,selected+1].forEach((p,index)=>{const dates=row.periods[p];if(dates)td.append(item('small',`${index ? `P${selected}` : 'WP'}: ${dates.current.map(month).join(' — ')} / ${dates.previous.map(month).join(' — ')}`,'period-detail'));});
        tr.append(td);
      });tbody.append(tr);
    });table.append(tbody);wrapper.append(table);section.append(wrapper,item('p',data.note,'table-note'));root.append(section);
  });
}

function dashboardOptions() { return {window:$('dashboard-window').value,labels:$('dashboard-labels').checked,scope:$('dashboard-label-scope').value,detail:$('dashboard-detail').checked}; }
function renderDashboard() {
  const sheet=state.analysis?.sheets.find(entry=>entry.name===state.active);
  if(!sheet?.dashboard)return;
  $('dashboard-window').value=$('window').value;
  const selected=state.selections[sheet.name];
  const count=CoverageDashboard.render($('dashboard-grid'),$('dashboard-nav'),sheet,selected,state.analysis.chart_colors,dashboardOptions(),(model,options)=>{
    $('chart-dialog-title').textContent=model.title;$('chart-dialog-description').textContent=`${sheet.brand} · P${selected} · ${model.description}`;
    $('chart-dialog').showModal();CoverageDashboard.plot($('chart-dialog-content'),model,options,true);
  },{selected:state.pptExtras[sheet.name] || [],disabled:state.busy || state.job?.status==='running',change:(key,checked)=>{
    const extras=new Set(state.pptExtras[sheet.name] || []);if(checked)extras.add(key);else extras.delete(key);
    state.pptExtras[sheet.name]=[...extras];saveDraft();
  }});
  $('dashboard-reference').textContent=`${count} gráficos visibles · ${sheet.brand} · P${selected} · Corte ${month(sheet.dates.at(-1))}. Los gráficos básicos se mantienen en el PPT. Marca «Incluir en PPT» en los KPI extra de esta hoja.`;
  [['Diagnóstico KPI','excel-diagnostics'],['Datos mensuales y P0–P6','template-table'],['Variaciones','excel-table-variations'],['Correlaciones','excel-table-correlations'],['Promedios MAT','excel-table-averages'],['Detalle anual MAT','excel-table-annual-mat']].forEach(([title,id])=>{
    const link=item('a',title);link.href=`#${id}`;link.className='dashboard-detail-link';
    link.addEventListener('click',()=>{if($('template-view').value!=='all'){$('template-view').value='all';renderTemplate();}});$('dashboard-nav').append(link);
  });
  $('dashboard-label-scope').disabled=!$('dashboard-labels').checked;
  renderDiagnostics();
}
function renderDiagnostics() {
  const sheet=state.analysis?.sheets.find(entry=>entry.name===state.active);if(!sheet?.dashboard)return;
  const change=$('kpi-change-threshold').valueAsNumber,range=$('kpi-range-threshold').valueAsNumber;
  const valid=Number.isFinite(change)&&change>=0&&change<=1000&&Number.isFinite(range)&&range>=0&&range<=1000;
  $('diagnostic-body').replaceChildren();
  sheet.dashboard.kpis.forEach(kpi=>{
    const percent=kpi.key==='penetration',digits=kpi.key==='buyers_series'?0:2;
    const review=kpi.average===0||Math.abs(kpi.change)>change/100||kpi.range>range/100;
    const tr=item('tr');
    [kpi.label,fmt(kpi.latest,digits,percent?'%':''),fmt(kpi.average,CoverageDashboard.kpiDigits(kpi,kpi.average),percent?'%':''),fmt(kpi.total_average,CoverageDashboard.kpiDigits(kpi,kpi.total_average),percent?'%':''),fmt(kpi.change*100,1,'%'),fmt(kpi.minimum,digits,percent?'%':''),fmt(kpi.maximum,digits,percent?'%':''),fmt(kpi.range*100,1,'%'),valid?(review?'REVISAR':'OK'):'Ajusta límites'].forEach((value,i)=>tr.append(item('td',value,i===8?`diagnostic-signal ${!valid||review?'review':'ok'}`:'')));
    $('diagnostic-body').append(tr);
  });
}

const ns = 'http://www.w3.org/2000/svg';
function svgElement(tag, attrs, text) {
  const node = document.createElementNS(ns, tag);
  Object.entries(attrs || {}).forEach(([key, value]) => node.setAttribute(key, String(value)));
  if (text != null) node.textContent = text;
  return node;
}
function drawChart() {
  const sheet = state.analysis?.sheets.find(sheet => sheet.name === state.active);
  if (!sheet?.pipelines) return;
  const selected = state.selections[sheet.name];
  const compare = $('compare').value;
  const metric = $('metric').value;
  const p = sheet.pipelines[selected];
  const plots = [];
  const colors = state.analysis.chart_colors;
  const coverageView = metric === 'coverage';
  const driver = driverMetrics[metric];
  const yoyView = metric === 'yoy';
  const percentView = coverageView || metric === 'penetration' || yoyView;
  const detailed = $('visual-detail').checked;
  const labels = $('data-labels').checked;
  const labelAll = labels && $('label-scope').value === 'all';
  $('label-scope').disabled = !labels;
  $('penetration-control').hidden = !coverageView;
  const dual = state.analysis.options.trend_axis === 'doble' && !coverageView && !driver && !yoyView;
  const dataKey = coverageView ? 'coverage_series' : metric === 'mat' ? 'sell_in_mat' : yoyView ? 'yoy' : 'sell_in';
  if (driver) plots.push({name: driver.title, values: sheet[metric], color: metric === 'penetration' ? colors.penetration : colors.worldpanel, axis: 'si', className: 'wp'});
  else {
  if (!coverageView) plots.push({name: 'Worldpanel', values: metric === 'mat' ? sheet.wp_mat : yoyView ? sheet.wp_yoy : sheet.sell_out, color: colors.worldpanel, axis: 'wp', className: 'wp'});
  if (coverageView && $('show-penetration').checked) plots.push({name: 'Penetración mensual', values: sheet.penetration, color: colors.penetration, axis: 'si', className: 'penetration', bar: true});
  plots.push({name: `${coverageView ? 'Cobertura' : 'Sell-in'} P${selected}`, values: p[dataKey], color: coverageView ? colors.coverage : colors.sellin, axis: 'si', className: coverageView ? 'coverage' : 'sellin', bar: coverageView});
  if (compare !== '') plots.push({name: `${coverageView ? 'Cobertura' : 'Sell-in'} P${compare}`, values: sheet.pipelines[Number(compare)][dataKey], color: colors.comparison, dash: '6 5', axis: 'si', className: 'compare'});
  if ($('original').checked && selected !== 0 && compare !== '0') plots.push({name: `${coverageView ? 'Cobertura' : 'Sell-in'} original (P0)`, values: sheet.pipelines[0][dataKey], color: colors.original, dash: '3 5', axis: 'si', className: 'original'});
  }
  const view=CoverageDashboard.chartData({dates:sheet.dates,series:plots},$('window').value);
  const dates=view.dates,n=dates.length,first=0;
  plots.splice(0,plots.length,...view.series);
  $('chart-period').textContent=n ? `${month(dates[0])} – ${month(dates.at(-1))} · Datos disponibles en todas las series visibles` : 'Sin período compartido con datos';
  $('chart-kpi-averages').replaceChildren();
  $('chart-kpi-averages').hidden=!driver;
  if(driver)CoverageDashboard.kpiAverages($('chart-kpi-averages'),sheet.dashboard.kpis.find(kpi=>kpi.key===metric));
  $('compare').hidden = !!driver; document.querySelector('label[for="compare"]').hidden = !!driver; document.querySelector('.original-control').hidden = !!driver;
  $('chart-title').textContent = driver ? driver.title : yoyView ? 'Variación interanual de volumen MAT' : coverageView ? ($('show-penetration').checked ? 'Cobertura MAT y penetración mensual' : 'Cobertura MAT por pipeline') : metric === 'mat' ? 'Tendencia de volumen MAT' : 'Tendencia de volumen mensual';
  $('chart-description').textContent = driver ? 'Serie mensual de Worldpanel · Conserva las fechas originales' : `P${selected}: ${selected === 0 ? 'sin desplazamiento' : `Sell-in desplazado ${selected} mes(es)`}${dual ? ' · Doble eje: Sell-in a la izquierda, Worldpanel a la derecha' : ' · Eje compartido'}${percentView ? ' (%)' : ''}`;
  $('unit-label').textContent = driver ? `${['purchase_average', 'purchase_occasion'].includes(metric) ? `${sheet.unit} ` : ''}${driver.unit}` : yoyView ? 'Variación MAT (%)' : coverageView ? `Cobertura ${state.analysis.options.coverage_type.toLowerCase()}${$('show-penetration').checked ? ' y penetración' : ''} (%)` : sheet.unit;
  $('original-label').textContent = coverageView ? 'Mostrar cobertura original (P0)' : 'Mostrar Sell-in original (P0)';
  $('original').disabled = selected === 0 || compare === '0';
  $('coverage-context').hidden = !coverageView;
  const reference = compare !== '' ? Number(compare) : 0;
  const baseline = sheet.pipelines[reference];
  const difference = (value, previous, digits, suffix) => value == null || previous == null ? '—' : `${value - previous > 0 ? '+' : ''}${fmt(value - previous, digits, suffix)}`;
  const pen = sheet.penetration.at(-1), penPrev = sheet.penetration.at(-2);
  $('coverage-context').textContent = `Penetración mensual: ${fmt(pen, 1, '%')} (${difference(pen, penPrev, 1, ' pp')} vs. mes anterior). ${selected === reference ? 'P0 conserva el Sell-in sin desplazamiento.' : `P${selected} vs. P${reference}: cobertura ${difference(p.coverage, baseline.coverage, 1, ' pp')}; correlación MAT ${difference(p.correlation, baseline.correlation, 3, '')}; cambio anual de cobertura ${fmt(p.stability, 1, ' pp')} (P${reference}: ${fmt(baseline.stability, 1, ' pp')}).`}`;
  $('chart-footnote').textContent = driver ? 'Estos indicadores conservan los valores mensuales del Excel y no se desplazan con el pipeline. Compra media y compra por ocasión usan la unidad de volumen de la hoja.' : yoyView ? 'YoY MAT compara la suma de 12 meses con la suma de los 12 meses anteriores. Las columnas de evolución del template usan 1 como denominador si el MAT anterior es cero.' : coverageView
    ? 'La penetración es mensual y conserva las fechas del Excel; no cambia con el pipeline. La cobertura usa volúmenes MAT. Revisa correlación, gap y estabilidad junto con estas curvas: una cobertura más alta por sí sola no indica un mejor pipeline.'
    : 'P1 desplaza Sell-in un mes hacia adelante; P6, seis meses. Los espacios sin datos se conservan. Correlación y variaciones se calculan sobre MAT, incluso al ver volumen mensual.';
  const graph = $('chart'); graph.replaceChildren();
  graph.append(svgElement('title', {}, `${$('chart-title').textContent}. ${$('chart-description').textContent}.`));
  if(!n) {
    graph.setAttribute('viewBox','0 0 640 230');graph.style.width='100%';
    graph.append(svgElement('text',{x:320,y:115,'text-anchor':'middle',fill:'#637888','font-size':14},'Sin base común en el período elegido'));
    graph.onpointermove=null;graph.onkeydown=null;$('chart-tooltip').textContent='Las bases insuficientes se muestran como —.';
    $('legend').replaceChildren();renderTemplate();renderDashboard();return;
  }
  const availableWidth = Math.max(640, graph.parentElement.clientWidth);
  const width = labelAll ? Math.max(availableWidth, (n - first) * 54) : availableWidth;
  graph.setAttribute('viewBox', `0 0 ${width} 340`);
  graph.style.width = width > availableWidth ? `${width}px` : '100%';
  const left = 68, right = width - (dual ? 77 : 50), top = 35, bottom = 274;
  const x = i => left + (i - first + .5) / Math.max(1, n - first) * (right - left);
  function scale(axis) {
    const values = plots.filter(plot => !dual || plot.axis === axis).flatMap(plot => plot.values.slice(first)).filter(v => v != null && Number.isFinite(v));
    return CoverageDashboard.axisScale(values,.1,.06);
  }
  const si = scale('si'), wp = dual ? scale('wp') : si;
  const y = (value, axis) => { const s = axis === 'wp' ? wp : si; return bottom - (value - s.min) / (s.max - s.min) * (bottom - top); };
  const compact = value => new Intl.NumberFormat('es-MX', {notation: 'compact', maximumFractionDigits: 1}).format(value);
  for (let tick = 0; tick <= 4; tick++) {
    const position = bottom - tick / 4 * (bottom - top);
    if (detailed) graph.append(svgElement('line', {class: 'chart-grid', x1: left, x2: right, y1: position, y2: position, stroke: '#e8eef2', 'stroke-width': 1}));
    graph.append(svgElement('text', {x: left - 12, y: position + 4, 'text-anchor': 'end', fill: '#637888', 'font-size': 12}, CoverageDashboard.axisTick(si.min + tick / 4 * (si.max - si.min),si.max-si.min)));
    if (dual) graph.append(svgElement('text', {x: right + 10, y: position + 4, fill: colors.worldpanel, 'font-size': 12}, CoverageDashboard.axisTick(wp.min + tick / 4 * (wp.max - wp.min),wp.max-wp.min)));
  }
  const step = Math.max(1, Math.ceil((n - first) / 10));
  for (let i = first; i < n; i += step) graph.append(svgElement('text', {x: x(i), y: bottom + 25, 'text-anchor': 'middle', fill: '#637888', 'font-size': 12}, month(dates[i])));
  const barPlots = plots.filter(plot => plot.bar);
  const barWidth = (right - left) / Math.max(1, n - first) * (barPlots.length > 1 ? .28 : .55);
  const pointX = (plot, i) => x(i) + (plot.bar && barPlots.length > 1 ? (barPlots.indexOf(plot) - .5) * barWidth * 1.1 : 0);
  const labelPoints = [];
  for (const plot of plots) {
    let path = '', segment = false;
    const segments = []; let points = [];
    const lastValid = plot.values.findLastIndex(value => value != null);
    for (let i = first; i < n; i++) {
      if (plot.values[i] == null) { segment = false; if (points.length) segments.push(points); points = []; continue; }
      const px = pointX(plot, i), py = y(plot.values[i], plot.axis);
      path += `${segment ? 'L' : 'M'}${px.toFixed(2)},${py.toFixed(2)} `; segment = true;
      points.push([px, py]);
      if (plot.bar) graph.append(svgElement('rect', {class: 'chart-bar', 'data-series': plot.className, x: px - barWidth / 2, y: Math.min(py, y(0, plot.axis)), width: barWidth, height: Math.abs(py - y(0, plot.axis)), fill: plot.color, stroke: detailed ? '#637888' : 'none', 'stroke-width': plot.className === 'coverage' ? 1.2 : .5, opacity: .88}));
      if (labels && (labelAll || i === lastValid)) labelPoints.push({plot, i, px, py});
    }
    if (points.length) segments.push(points);
    if (!plot.bar) {
      if (detailed && plot.className === 'wp') segments.forEach(segmentPoints => {
        const area = `M${segmentPoints[0][0]},${y(0, plot.axis)} ` + segmentPoints.map(([px, py]) => `L${px},${py}`).join(' ') + ` L${segmentPoints.at(-1)[0]},${y(0, plot.axis)} Z`;
        graph.append(svgElement('path', {class: 'chart-area', d: area, fill: colors.trend_fill}));
      });
      graph.append(svgElement('path', {class: 'chart-series', 'data-series': plot.className, d: path, fill: 'none', stroke: plot.color, 'stroke-width': plot.className === 'original' ? 1.8 : 2.6, 'stroke-linejoin': 'round', 'stroke-linecap': 'round', ...(plot.dash ? {'stroke-dasharray': plot.dash} : {})}));
      if (detailed) segments.flat().forEach(([px, py]) => graph.append(svgElement('circle', {class: 'chart-point', cx: px, cy: py, r: 2.6, fill: plot.color, stroke: '#fff', 'stroke-width': .8})));
    }
    const valid = plot.values.slice(first).filter(v => v != null);
    if (valid.length === 1) { const i = plot.values.findLastIndex(v => v != null); graph.append(svgElement('circle', {cx: x(i), cy: y(plot.values[i], plot.axis), r: 4, fill: plot.color})); }
  }
  const occupied = [];
  labelPoints.forEach(({plot, i, px, py}) => {
    const text = percentView ? fmt(plot.values[i], 1, '%') : driver && driver.digits === 2 ? fmt(plot.values[i], 2) : compact(plot.values[i]);
    const half = Math.max(20, text.length * 3.7), lx = Math.max(left + half, Math.min(right - half, px));
    let ly = Math.max(14, py - 12);
    // Move neighboring labels vertically, retaining a leader to their own data point.
    while (occupied.some(box => Math.abs(box.x - lx) < box.half + half + 3 && Math.abs(box.y - ly) < 19)) ly += 20;
    if (ly > bottom - 5) {
      ly = Math.min(bottom - 5, py + 19);
      while (occupied.some(box => Math.abs(box.x - lx) < box.half + half + 3 && Math.abs(box.y - ly) < 19)) ly -= 20;
    }
    occupied.push({x: lx, y: ly, half});
    const group = svgElement('g', {class: 'data-label', 'data-series': plot.className});
    group.append(svgElement('title', {}, `${plot.name} · ${month(dates[i])}: ${fmt(plot.values[i], driver?.digits ?? 1, percentView ? '%' : '')}`));
    group.append(svgElement('line', {x1: px, y1: py, x2: lx, y2: ly, stroke: plot.color, 'stroke-width': .8}));
    group.append(svgElement('rect', {x: lx - half, y: ly - 12, width: half * 2, height: 18, rx: 3, fill: '#fff', stroke: plot.color, 'stroke-width': .9}));
    group.append(svgElement('text', {x: lx, y: ly + 1, 'text-anchor': 'middle', fill: driver || ['coverage', 'penetration', 'sellin', 'original'].includes(plot.className) ? '#172c3c' : plot.color, 'font-size': 12, 'font-weight': 600}, text));
    graph.append(group);
  });
  if (!plots.some(plot => plot.values.slice(first).some(value => value != null))) graph.append(svgElement('text', {x: width / 2, y: 145, 'text-anchor': 'middle', fill: '#718797', 'font-size': 15}, 'Sin historia suficiente en el periodo visible'));
  const guide = svgElement('line', {x1: left, x2: left, y1: top, y2: bottom, stroke: '#aabac5', 'stroke-dasharray': '3 4', visibility: 'hidden'}); graph.append(guide);
  let hoverIndex = n - 1;
  function tooltip(i) {
    hoverIndex = i; guide.setAttribute('visibility', detailed ? 'visible' : 'hidden'); guide.setAttribute('x1', x(i)); guide.setAttribute('x2', x(i));
    $('chart-tooltip').textContent = `${month(dates[i])}   ·   ${plots.map(plot => `${plot.name}: ${fmt(plot.values[i], driver?.digits ?? 1, percentView ? '%' : '')}`).join('   ·   ')}`;
  }
  graph.onpointermove = event => { const box = graph.getBoundingClientRect(); const px = (event.clientX - box.left) / box.width * width; tooltip(Math.max(first, Math.min(n - 1, Math.floor(first + (px - left) / (right - left) * (n - first))))); };
  graph.onkeydown = event => { if (['ArrowLeft', 'ArrowRight'].includes(event.key)) { event.preventDefault(); tooltip(Math.max(first, Math.min(n - 1, hoverIndex + (event.key === 'ArrowRight' ? 1 : -1)))); } };
  tooltip(n - 1);
  $('legend').replaceChildren(...plots.map(plot => { const legend = item('span', plot.name, plot.className); legend.style.setProperty('--series-color', plot.color); return legend; }));
  renderTemplate();
  renderDashboard();
}

async function saveSelection() {
  state.busy = true; controls();
  try { const result = await api('/api/save', payload()); saveDraft(); status(result.message); }
  catch (error) { status(error.message, true); }
  finally { state.busy = false; controls(); }
}
function reviewExport() {
  $('dialog-file').textContent = `${state.analysis.file} · ${$('template').selectedOptions[0].textContent}`;
  $('dialog-selection').replaceChildren();
  Object.entries(selectedSheets()).forEach(([name, pipeline]) => { const row = item('div', null, 'dialog-row');
    const extras=(state.pptExtras[name] || []).map(key=>driverMetrics[key].title);
    row.append(item('span',name+(extras.length ? ` · Extra: ${extras.join(', ')}` : ' · Solo gráficos básicos')), item('strong', `P${pipeline}`)); $('dialog-selection').append(row); });
  $('export-dialog').showModal();
}
async function prepareExport() {
  if(state.busy || state.job?.status==='running')return;
  if(!state.analysis || state.dirty)await load();
  if(state.analysis && !state.dirty && Object.keys(selectedSheets()).length)reviewExport();
  else if(state.analysis)status('Incluye al menos una hoja para generar.',true);
}
async function exportFiles() {
  $('export-dialog').close(); state.busy = true; controls();
  try {
    await api('/api/save', payload());
    state.job = await api('/api/export', payload());
    state.resultRequest++;
    status('Generación iniciada con tu selección. El registro se actualiza aquí.');
    renderResults();
    pollJob();
  } catch (error) { status(error.message, true); }
  finally { state.busy = false; controls(); }
}
function renderResults() {
  const job=state.job?.file===$('file').value ? state.job : null;
  const coverage=job?.options?.coverage_type;
  $('job-title').textContent=!job ? 'Aún no hay una entrega para este archivo' : job.source_changed ? 'Entrega anterior · datos modificados' : job.status==='complete' ? (coverage ? `Cobertura ${coverage.toLocaleLowerCase()}` : 'Entrega lista') : job.status==='failed' ? 'La generación requiere revisión' : 'Generando entregables…';
  $('job-badge').hidden=!job;
  $('job-badge').textContent=job?.status==='complete' ? 'Completado' : job?.status==='failed' ? 'Error' : 'En curso';
  $('job-date').textContent=job?.started_at ? new Intl.DateTimeFormat('es-MX',{dateStyle:'medium',timeStyle:'short'}).format(new Date(job.started_at)) : '';
  $('result-hint').textContent=job?.source_changed ? 'Estos resultados corresponden a los datos anteriores. Genera una nueva entrega para actualizar.' : job?.status==='complete' ? 'Archivos de la última entrega de este Excel.' : 'Se habilitan al finalizar la generación.';
  $('job-delivery-title').textContent=job?.title || '';
  $('job-message').textContent=job ? `Ruta: ${job.directory}${job.error ? `\n${job.error}` : ''}` : '';
  $('job-log').textContent=(job?.log || '').replace(/\x1b\[[0-9;]*[a-zA-Z]/g,'');
  controls();syncChromeGeometry();
}
async function refreshJobForFile() {
  const request=++state.resultRequest, file=$('file').value;
  state.job=null;$('job-log-details').open=false;renderResults();
  if(!file)return;
  status('');
  let result;
  try{result=await api(`/api/jobs?file=${encodeURIComponent(file)}`);}
  catch(error){if(request===state.resultRequest && file===$('file').value)status(error.message,true);return;}
  if(request!==state.resultRequest || file!==$('file').value)return;
  state.job=result.latest;renderResults();
  if(state.job?.status==='running')pollJob();
}
async function openOutput(label) {
  const job=state.job;
  if(!job || job.file!==$('file').value || job.status!=='complete' || !job.available_outputs.includes(label) || state.openingOutput || state.openingResults)return;
  state.openingOutput=label;controls();status(`Abriendo ${outputNames[label].replace('Abrir ','')}…`);
  try { await api('/api/open-output',{id:job.id,file:job.file,output:label});status(`${outputNames[label].replace('Abrir ','')}: abierto en su aplicación.`); }
  catch(error){status(error.message,true);}
  finally{state.openingOutput=null;controls();}
}
async function openResults() {
  if(state.openingResults || state.openingOutput || state.job?.status!=='complete' || state.job.file!==$('file').value)return;
  state.openingResults=true;controls();
  status('Abriendo la carpeta de resultados…');
  try { const result=await api('/api/open-results',{scope:'job',id:state.job.id});status(result.message); }
  catch(error) { status(error.message,true); }
  finally { state.openingResults=false;controls(); }
}
async function pollJob() {
  const id=state.job?.id;if(!id)return;
  try {
    const job = await api(`/api/job?id=${encodeURIComponent(id)}`);
    if(state.job?.id!==id || job.file!==$('file').value)return;
    state.job=job;renderResults();
    if (job.status !== 'running') {
      status(job.error || 'Entrega lista para abrir.', !!job.error);
    }
    controls();
    syncChromeGeometry();
    if (job.status === 'running') setTimeout(pollJob, 2000);
  } catch (error) { if(state.job?.id!==id)return;status(`No se pudo consultar el avance: ${error.message}`, true); setTimeout(pollJob, 5000); }
}

$('load').addEventListener('click', load);
$('browse-files').addEventListener('click', () => { renderFileCards(); $('file-dialog').showModal(); });
$('close-files').addEventListener('click', () => $('file-dialog').close());
$('close-chart').addEventListener('click', () => $('chart-dialog').close());
let chartBackdropPress=false;
$('chart-dialog').addEventListener('pointerdown',event=>{
  const box=$('chart-dialog').getBoundingClientRect();
  chartBackdropPress=event.target===$('chart-dialog') && (event.clientX<box.left||event.clientX>box.right||event.clientY<box.top||event.clientY>box.bottom);
});
$('chart-dialog').addEventListener('click',event=>{
  const box=$('chart-dialog').getBoundingClientRect();
  if(chartBackdropPress && event.target===$('chart-dialog') && (event.clientX<box.left||event.clientX>box.right||event.clientY<box.top||event.clientY>box.bottom))$('chart-dialog').close();
  chartBackdropPress=false;
});
['dashboard-labels','dashboard-label-scope','dashboard-detail'].forEach(id=>$(id).addEventListener('change',renderDashboard));
$('dashboard-window').addEventListener('change',()=>{$('window').value=$('dashboard-window').value;drawChart();});
['kpi-change-threshold','kpi-range-threshold'].forEach(id=>$(id).addEventListener('input',renderDiagnostics));
$('file-search').addEventListener('input', renderFileCards);
['file-country', 'file-category'].forEach(id => $(id).addEventListener('change', renderFileCards));
$('file').addEventListener('change',()=>{renderFileDetail();refreshJobForFile().catch(error=>status(error.message,true));});
['template-view', 'table-columns', 'inspect-month'].forEach(id => $(id).addEventListener('change', renderTemplate));
let chartResizeFrame;
window.addEventListener('resize', () => { cancelAnimationFrame(chartResizeFrame); chartResizeFrame = requestAnimationFrame(drawChart); });
$('refresh-files').addEventListener('click', () => refreshFiles().then(markDirty).catch(error => status(error.message, true)));
['file', 'trend-axis', 'english', 'round', 'animate'].forEach(id => $(id).addEventListener('change', markDirty));
$('template').addEventListener('change', () => { $('english').checked = $('template').value === '5'; $('trend-axis').value = $('template').value === '4' ? 'doble' : 'simple'; markDirty(); });
$('search').addEventListener('input', () => { if (state.analysis) { renderSheets(); controls(); } });
$('select-all').addEventListener('click', () => { state.analysis.sheets.forEach(sheet => { state.included[sheet.name] = !sheet.error; }); saveDraft(); render(); });
$('select-none').addEventListener('click', () => { Object.keys(state.included).forEach(name => { state.included[name] = false; }); saveDraft(); render(); });
$('use-recommended').addEventListener('click', () => choose(state.analysis.sheets.find(sheet => sheet.name === state.active).recommended));
['compare', 'metric', 'window', 'original', 'data-labels', 'label-scope', 'visual-detail', 'show-penetration'].forEach(id => $(id).addEventListener('change', () => {
  try { localStorage.setItem('coveragelab-chart-view', JSON.stringify({labels: $('data-labels').checked, scope: $('label-scope').value, detail: $('visual-detail').checked, penetration: $('show-penetration').checked})); } catch (_) { /* View remains usable without local storage. */ }
  drawChart();
}));
try {
  const view = JSON.parse(localStorage.getItem('coveragelab-chart-view') || 'null');
  if (view) {
    if (typeof view.labels === 'boolean') $('data-labels').checked = view.labels;
    if (['last', 'all'].includes(view.scope)) $('label-scope').value = view.scope;
    if (typeof view.detail === 'boolean') $('visual-detail').checked = view.detail;
    if (typeof view.penetration === 'boolean') $('show-penetration').checked = view.penetration;
  }
} catch (_) { /* Default visual settings are retained. */ }
$('save').addEventListener('click', saveSelection);
$('generate-files').addEventListener('click', prepareExport);
$('open-folder').addEventListener('click',openResults);
document.querySelectorAll('[data-output]').forEach(button=>button.addEventListener('click',()=>openOutput(button.dataset.output)));
$('cancel-export').addEventListener('click', () => $('export-dialog').close());
$('confirm-export').addEventListener('click', exportFiles);
$('toggle-header').addEventListener('click', () => setHeaderCollapsed(!$('result-dock').hidden));
let headerCollapsed = false;
try { headerCollapsed = localStorage.getItem('coveragelab-header-collapsed') === 'true'; } catch (_) { /* El encabezado se muestra por defecto. */ }
setHeaderCollapsed(headerCollapsed, false);
const chromeObserver=new ResizeObserver(syncChromeGeometry);
chromeObserver.observe($('app-top'));
window.addEventListener('resize',syncChromeGeometry);
syncChromeGeometry();controls();
refreshFiles().catch(error => status(error.message, true));

window.CoverageApp = {
  resultButton,
  inputPending(name, pending) {state.pendingInput=pending ? name : null;controls();},
  context() {
    const file = state.files.find(entry => entry.name === $('file').value);
    const sheet = state.analysis?.sheets.find(entry => entry.name === state.active);
    return {file, activeSheet:sheet?.name, pipeline:sheet ? state.selections[sheet.name] : 0,
      selections:selectedSheets(), brandPipelines:Object.fromEntries((state.analysis?.sheets || []).filter(entry=>state.included[entry.name] && !entry.error).map(entry=>[entry.brand,state.selections[entry.name]])), loaded:!!state.analysis && !state.dirty};
  },
  async openFile(name) {
    if (state.busy || state.job?.status === 'running') throw new Error('Espera a que termine la operación de cobertura.');
    await refreshFiles();
    if (![...$('file').options].some(option => option.value === name)) throw new Error('El archivo no está disponible para cobertura.');
    $('file').value=name;renderFileDetail();await refreshJobForFile();
    await load();
    if (!state.analysis) throw new Error($('status').textContent || 'Revisa las entradas antes de explorar.');
  },
  async sourceChanged(name) {
    if (state.analysis?.file === name || $('file').value === name) {
      state.analysis=null;state.dirty=true;$('workspace').hidden=true;$('sheet-panel').hidden=true;$('empty').hidden=false;
      $('export-dialog').close();
      status('Entradas actualizadas. Carga de nuevo para recalcular los pipelines y resultados.');
    }
    await refreshFiles();
    controls();
  },
  openChart(title, description, content) {
    $('chart-dialog-title').textContent=title;$('chart-dialog-description').textContent=description;
    $('chart-dialog-content').replaceChildren(content);$('chart-dialog').showModal();
  },
  geometry:syncChromeGeometry,
  redraw:drawChart,
};
