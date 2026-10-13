'use strict';
const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name="explorer-token"]').content;
const state = {analysis: null, active: null, selections: {}, included: {}, busy: false, dirty: false, job: null};
const fmt = (value, digits = 1, suffix = '') => value == null ? '—' : new Intl.NumberFormat('es-MX', {minimumFractionDigits: digits, maximumFractionDigits: digits}).format(value) + suffix;
const month = date => new Intl.DateTimeFormat('es-MX', {month: 'short', year: '2-digit'}).format(new Date(date + 'T12:00:00'));
const item = (tag, text, className) => { const element = document.createElement(tag); if (text != null) element.textContent = text; if (className) element.className = className; return element; };

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
}
function requestOptions() {
  return {file: $('file').value, template: $('template').value, trend_axis: $('trend-axis').value,
    include_english: $('english').checked, round_coverage: $('round').checked, animate_trend_pipeline: $('animate').checked};
}
function selectedSheets() {
  return Object.fromEntries(Object.entries(state.selections).filter(([name]) => state.included[name]));
}
function payload() { return {analysis_id: state.analysis.id, selections: selectedSheets()}; }
function draftKey() { return `coveragelab-selection-${state.analysis.selection_key}`; }
function saveDraft() {
  try { localStorage.setItem(draftKey(), JSON.stringify({selections: state.selections, included: state.included})); } catch (_) { /* Disk save remains available. */ }
}
function controls() {
  const working = state.busy || state.job?.status === 'running';
  ['file', 'template', 'trend-axis', 'english', 'round', 'animate', 'refresh-files', 'load'].forEach(id => { $(id).disabled = working; });
  const ready = state.analysis && !state.dirty && Object.keys(selectedSheets()).length && !working;
  $('save').disabled = !ready;
  $('export').disabled = !ready;
  $('select-all').disabled = working;
  $('select-none').disabled = working;
  document.querySelectorAll('.sheet-row input, .pipeline-button').forEach(control => { control.disabled = working || control.dataset.unavailable === 'true'; });
  const active = state.analysis?.sheets.find(sheet => sheet.name === state.active);
  $('use-recommended').disabled = working || !active?.pipelines?.[active.recommended]?.available;
}
async function refreshFiles() {
  const previous = $('file').value;
  const result = await api('/api/files');
  $('file').replaceChildren();
  if (!result.files.length) $('file').append(new Option('No hay archivos Excel en el proyecto', ''));
  else result.files.forEach(name => $('file').append(new Option(name, name)));
  if (result.files.includes(previous)) $('file').value = previous;
  if (!result.files.length) status('Coloca un .xlsx en la carpeta del proyecto y pulsa Actualizar archivos.');
}
function markDirty() {
  if (state.analysis) { state.dirty = true; status('La configuración cambió. Pulsa «Cargar y explorar» para actualizar las métricas antes de generar.'); controls(); }
}
async function load() {
  if (!$('file').value) return status('Selecciona un archivo Excel.', true);
  state.busy = true; state.analysis = null; state.dirty = false;
  $('workspace').hidden = true; $('sheet-panel').hidden = true; $('empty').hidden = false;
  $('job-panel').hidden = true;
  status('Cargando hojas y calculando P0–P6…'); controls();
  try {
    const analysis = await api('/api/analyze', requestOptions());
    state.analysis = analysis; state.selections = {}; state.included = {};
    let saved;
    try { saved = JSON.parse(localStorage.getItem(draftKey()) || 'null'); } catch (_) { saved = null; }
    if (!saved && analysis.saved_selection) saved = {selections: analysis.saved_selection.selections,
      included: Object.fromEntries(Object.keys(analysis.saved_selection.selections).map(name => [name, true]))};
    analysis.sheets.forEach(sheet => {
      if (sheet.error) { state.included[sheet.name] = false; return; }
      const previous = saved?.selections?.[sheet.name];
      state.selections[sheet.name] = Number.isInteger(previous) && sheet.pipelines[previous]?.available ? previous : sheet.default_pipeline;
      state.included[sheet.name] = saved ? !!saved.included?.[sheet.name] : true;
    });
    state.active = analysis.sheets.find(sheet => !sheet.error)?.name || analysis.sheets[0]?.name;
    $('empty').hidden = true; $('workspace').hidden = false; $('sheet-panel').hidden = false;
    $('search').value = '';
    const invalid = analysis.sheets.filter(sheet => sheet.error).length;
    status(invalid ? `${invalid} hoja(s) requieren corrección y quedan excluidas. Puedes revisar sus motivos en la lista.` : saved ? 'Selección anterior recuperada. Revisa las curvas antes de generar.' : 'Archivo listo. Explora cada hoja y elige su pipeline.');
    $('step-file').classList.remove('active'); $('step-review').classList.add('active'); $('step-export').classList.remove('active');
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
  $('context').textContent = `${state.analysis.country} / ${state.analysis.manufacturer} / ${sheet?.category || state.analysis.category}`;
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
    [ `P${pipeline.pipeline}`, fmt(pipeline.correlation, 3), fmt(pipeline.coverage, 1, '%'), fmt(pipeline.client_variation, 1, '%'), fmt(pipeline.gap), fmt(pipeline.stability), labels.join(' · ') || (pipeline.pipeline === 0 ? 'Sin desplazamiento' : '—') ].forEach((value, index) => row.append(item('td', value, index === 6 ? 'table-tag' : '')));
    $('pipeline-table').append(row);
  });
  drawChart(); controls();
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
  const n = sheet.dates.length;
  const first = $('window').value === 'all' ? 0 : Math.max(0, n - Number($('window').value));
  const plots = [];
  const dual = state.analysis.options.trend_axis === 'doble' && metric !== 'coverage';
  const dataKey = metric === 'coverage' ? 'coverage_series' : metric === 'mat' ? 'sell_in_mat' : 'sell_in';
  if (metric !== 'coverage') plots.push({name: 'Worldpanel', values: metric === 'mat' ? sheet.wp_mat : sheet.sell_out, color: '#078c86', axis: 'wp', className: 'wp'});
  plots.push({name: `${metric === 'coverage' ? 'Cobertura' : 'Sell-in'} P${selected}`, values: p[dataKey], color: '#7361c5', axis: 'si', className: 'sellin'});
  if (compare !== '') plots.push({name: `${metric === 'coverage' ? 'Cobertura' : 'Sell-in'} P${compare}`, values: sheet.pipelines[Number(compare)][dataKey], color: '#d6953b', dash: '6 5', axis: 'si', className: 'compare'});
  if ($('original').checked && selected !== 0 && metric !== 'coverage') plots.unshift({name: 'Sell-in original', values: sheet.pipelines[0][dataKey], color: '#b3bdc7', dash: '3 5', axis: 'si', className: 'original'});
  $('chart-title').textContent = metric === 'coverage' ? 'Cobertura MAT por pipeline' : metric === 'mat' ? 'Tendencia de volumen MAT' : 'Tendencia de volumen mensual';
  $('chart-description').textContent = `P${selected}: ${selected === 0 ? 'sin desplazamiento' : `Sell-in desplazado ${selected} mes(es)`}${dual ? ' · Doble eje: Sell-in a la izquierda, Worldpanel a la derecha' : ' · Eje compartido'}`;
  $('unit-label').textContent = metric === 'coverage' ? `Cobertura ${state.analysis.options.coverage_type.toLowerCase()} (%)` : sheet.unit;
  $('original').disabled = metric === 'coverage';
  const graph = $('chart'); graph.replaceChildren();
  graph.append(svgElement('title', {}, `${$('chart-title').textContent}. ${$('chart-description').textContent}.`));
  const left = 68, right = dual ? 923 : 965, top = 24, bottom = 266;
  const x = i => left + (i - first) / Math.max(1, n - first - 1) * (right - left);
  function scale(axis) {
    const values = plots.filter(plot => !dual || plot.axis === axis).flatMap(plot => plot.values.slice(first)).filter(v => v != null && Number.isFinite(v));
    const minimum = Math.min(0, ...values), maximum = Math.max(1, ...values);
    const spread = maximum - minimum;
    return {min: minimum < 0 ? minimum - spread * .06 : 0, max: maximum + spread * .1};
  }
  const si = scale('si'), wp = dual ? scale('wp') : si;
  const y = (value, axis) => { const s = axis === 'wp' ? wp : si; return bottom - (value - s.min) / (s.max - s.min) * (bottom - top); };
  const compact = value => new Intl.NumberFormat('es-MX', {notation: 'compact', maximumFractionDigits: 1}).format(value);
  for (let tick = 0; tick <= 4; tick++) {
    const position = bottom - tick / 4 * (bottom - top);
    graph.append(svgElement('line', {x1: left, x2: right, y1: position, y2: position, stroke: '#e8eef2', 'stroke-width': 1}));
    graph.append(svgElement('text', {x: left - 12, y: position + 4, 'text-anchor': 'end', fill: '#91a0ac', 'font-size': 10}, compact(si.min + tick / 4 * (si.max - si.min))));
    if (dual) graph.append(svgElement('text', {x: right + 10, y: position + 4, fill: '#078c86', 'font-size': 10}, compact(wp.min + tick / 4 * (wp.max - wp.min))));
  }
  const step = Math.max(1, Math.ceil((n - first) / 10));
  for (let i = first; i < n; i += step) graph.append(svgElement('text', {x: x(i), y: bottom + 25, 'text-anchor': 'middle', fill: '#8c9dab', 'font-size': 10}, month(sheet.dates[i])));
  for (const plot of plots) {
    let path = '', segment = false;
    for (let i = first; i < n; i++) {
      if (plot.values[i] == null) { segment = false; continue; }
      path += `${segment ? 'L' : 'M'}${x(i).toFixed(2)},${y(plot.values[i], plot.axis).toFixed(2)} `; segment = true;
    }
    graph.append(svgElement('path', {d: path, fill: 'none', stroke: plot.color, 'stroke-width': plot.className === 'original' ? 1.6 : 2.6, 'stroke-linejoin': 'round', 'stroke-linecap': 'round', ...(plot.dash ? {'stroke-dasharray': plot.dash} : {})}));
    const valid = plot.values.slice(first).filter(v => v != null);
    if (valid.length === 1) { const i = plot.values.findLastIndex(v => v != null); graph.append(svgElement('circle', {cx: x(i), cy: y(plot.values[i], plot.axis), r: 4, fill: plot.color})); }
  }
  if (!plots.some(plot => plot.values.slice(first).some(value => value != null))) graph.append(svgElement('text', {x: 500, y: 145, 'text-anchor': 'middle', fill: '#718797', 'font-size': 15}, 'Sin historia suficiente en el periodo visible'));
  const guide = svgElement('line', {x1: left, x2: left, y1: top, y2: bottom, stroke: '#aabac5', 'stroke-dasharray': '3 4', visibility: 'hidden'}); graph.append(guide);
  let hoverIndex = n - 1;
  function tooltip(i) {
    hoverIndex = i; guide.setAttribute('visibility', 'visible'); guide.setAttribute('x1', x(i)); guide.setAttribute('x2', x(i));
    $('chart-tooltip').textContent = `${month(sheet.dates[i])}   ·   ${plots.map(plot => `${plot.name}: ${fmt(plot.values[i], 1, metric === 'coverage' ? '%' : '')}`).join('   ·   ')}`;
  }
  graph.onpointermove = event => { const box = graph.getBoundingClientRect(); const px = (event.clientX - box.left) / box.width * 1000; tooltip(Math.max(first, Math.min(n - 1, Math.round(first + (px - left) / (right - left) * (n - first - 1))))); };
  graph.onkeydown = event => { if (['ArrowLeft', 'ArrowRight'].includes(event.key)) { event.preventDefault(); tooltip(Math.max(first, Math.min(n - 1, hoverIndex + (event.key === 'ArrowRight' ? 1 : -1)))); } };
  tooltip(n - 1);
  $('legend').replaceChildren(...plots.map(plot => item('span', plot.name, plot.className)));
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
  Object.entries(selectedSheets()).forEach(([name, pipeline]) => { const row = item('div', null, 'dialog-row'); row.append(item('span', name), item('strong', `P${pipeline}`)); $('dialog-selection').append(row); });
  $('export-dialog').showModal();
}
async function exportFiles() {
  $('export-dialog').close(); state.busy = true; controls();
  try {
    await api('/api/save', payload());
    state.job = await api('/api/export', payload());
    status('Generación iniciada con tu selección. El registro se actualiza aquí.');
    $('job-panel').hidden = false;
    $('step-review').classList.remove('active'); $('step-export').classList.add('active');
    pollJob();
  } catch (error) { status(error.message, true); }
  finally { state.busy = false; controls(); }
}
function downloadLink(url, label) {
  const a = item('a', label); a.href = `${url}${url.includes('?') ? '&' : '?'}token=${encodeURIComponent(token)}`; a.setAttribute('referrerpolicy', 'no-referrer'); return a;
}
async function pollJob() {
  try {
    state.job = await api(`/api/job?id=${encodeURIComponent(state.job.id)}`);
    const job = state.job;
    $('job-title').textContent = job.status === 'complete' ? 'Entregables generados' : job.status === 'failed' ? 'La generación requiere revisión' : 'Generando entregables…';
    $('job-badge').textContent = job.status === 'complete' ? 'Completado' : job.status === 'failed' ? 'Error' : 'En curso';
    $('job-message').textContent = job.error || (job.status === 'complete' ? `Archivos guardados en ${job.directory}` : 'Puedes consultar el avance en el registro. Esta generación usa la selección que acabas de confirmar.');
    $('job-log').textContent = (job.log || '').replace(/\x1b\[[0-9;]*[a-zA-Z]/g, '');
    $('downloads').replaceChildren();
    const names = {template: '↓ Template Excel', presentation: '↓ Presentación PPT', bank: '↓ Banco de coberturas', pipeline_report: '↓ Reporte de pipelines'};
    job.outputs.forEach(output => $('downloads').append(downloadLink(output.url, names[output.label] || output.name)));
    $('downloads').append(downloadLink(`/api/download?job=${encodeURIComponent(job.id)}&output=selection`, '↓ Selección utilizada'));
    if (job.status !== 'running') {
      $('downloads').append(downloadLink(`/api/download?job=${encodeURIComponent(job.id)}&output=log`, '↓ Registro'));
      status(job.error || 'Generación completa. Los entregables y la configuración están disponibles para descargar.', !!job.error);
    }
    controls();
    if (job.status === 'running') setTimeout(pollJob, 2000);
  } catch (error) { status(`No se pudo consultar el avance: ${error.message}. La generación puede seguir en su carpeta local.`, true); setTimeout(pollJob, 5000); }
}

$('load').addEventListener('click', load);
$('refresh-files').addEventListener('click', () => refreshFiles().then(markDirty).catch(error => status(error.message, true)));
['file', 'trend-axis', 'english', 'round', 'animate'].forEach(id => $(id).addEventListener('change', markDirty));
$('template').addEventListener('change', () => { $('english').checked = $('template').value === '5'; $('trend-axis').value = $('template').value === '4' ? 'doble' : 'simple'; markDirty(); });
$('search').addEventListener('input', () => { if (state.analysis) { renderSheets(); controls(); } });
$('select-all').addEventListener('click', () => { state.analysis.sheets.forEach(sheet => { state.included[sheet.name] = !sheet.error; }); saveDraft(); render(); });
$('select-none').addEventListener('click', () => { Object.keys(state.included).forEach(name => { state.included[name] = false; }); saveDraft(); render(); });
$('use-recommended').addEventListener('click', () => choose(state.analysis.sheets.find(sheet => sheet.name === state.active).recommended));
['compare', 'metric', 'window', 'original'].forEach(id => $(id).addEventListener('change', drawChart));
$('save').addEventListener('click', saveSelection);
$('export').addEventListener('click', reviewExport);
$('cancel-export').addEventListener('click', () => $('export-dialog').close());
$('confirm-export').addEventListener('click', exportFiles);
controls();
refreshFiles().catch(error => status(error.message, true));
api('/api/jobs').then(result => {
  if (result.latest) {
    state.job = result.latest;
    $('job-panel').hidden = false;
    controls();
    pollJob();
  }
}).catch(error => status(error.message, true));
