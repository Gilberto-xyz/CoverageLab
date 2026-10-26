/* 5W1H stays in the same local studio; chart values come from the Python engine. */
(() => {
  'use strict';
  const $ = (tag, attrs = {}, text) => {
    const node = document.createElement(tag);
    for (const [key, value] of Object.entries(attrs)) {
      if (key === 'class') node.className = value;
      else if (key === 'checked' || key === 'disabled' || key === 'hidden') node[key] = Boolean(value);
      else node.setAttribute(key, String(value));
    }
    if (text !== undefined) node.textContent = String(text);
    return node;
  };
  const svgNode = (tag, attrs = {}, text) => {
    const node = document.createElementNS('http://www.w3.org/2000/svg', tag);
    Object.entries(attrs).forEach(([key, value]) => node.setAttribute(key, String(value)));
    if (text !== undefined) node.textContent = String(text);
    return node;
  };
  const normalize = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s*[-_]?\s*(?:copia|copy|5w1h)\s*$/g, '').replace(/[^a-z0-9]/g, '');
  const finite = value => typeof value === 'number' && Number.isFinite(value);
  const number = (value, unit = '') => {
    if (!finite(value)) return '—';
    if (unit === 'percent') return `${(value * 100).toLocaleString('es-MX', {maximumFractionDigits: 2})}%`;
    if (unit === '%') return `${value.toLocaleString('es-MX', {maximumFractionDigits: 2})}%`;
    if (unit === 'pp') return `${value.toLocaleString('es-MX', {maximumFractionDigits: 2})} pp`;
    return value.toLocaleString('es-MX', {maximumFractionDigits: Math.abs(value) < 1 ? 4 : 2});
  };
  const period = value => /^\d{4}-\d{2}$/.test(String(value)) ? new Date(`${value}-01T12:00:00`).toLocaleDateString('es-MX', {month: 'short', year: '2-digit'}) : String(value);
  const state = {container: null, bridge: null, catalog: null, file: null, inspection: null, analysis: null,
    job: null, busy: false, request: 0, poll: null, timer: null, manual: false, labels: true, detail: true,
    elements: null, contextKey: null, opening: null, resultsCollapsed: false};
  try {state.resultsCollapsed=localStorage.getItem('coveragelab-research-results-collapsed')==='true';} catch (_) {}
  const nodes = {};
  let generationTimer;
  function renderResearchProgress() {
    const job = state.job?.file === state.file ? state.job : null;
    const button = nodes.generate, view = window.GenerationProgress?.view(job, null);
    button.classList.toggle('is-generation-active', !!view);
    button.classList.toggle('generation-failed', view?.mode === 'failed');
    button.classList.toggle('generation-complete', view?.mode === 'complete');
    if (!view) {
      button.textContent = 'Generar informe 5W1H';
      button.style.removeProperty('--generation-progress');
      button.removeAttribute('aria-busy'); button.removeAttribute('aria-label'); button.title = '';
      if (generationTimer) {clearInterval(generationTimer); generationTimer = null;}
      return;
    }
    if (!button.querySelector('.generation-fill')) {
      button.replaceChildren($('span', {class:'generation-fill', role:'progressbar', 'aria-valuemin':0, 'aria-valuemax':100, 'aria-label':'Avance de exportación 5W1H'}),
        $('span', {class:'generation-main'}), $('small', {class:'generation-time'}));
    }
    button.style.setProperty('--generation-progress', `${view.percent}%`);
    const meter = button.querySelector('.generation-fill');
    meter.setAttribute('aria-valuenow', String(Math.floor(view.percent)));
    meter.setAttribute('aria-valuetext', `${view.primary}. ${view.secondary}`);
    button.querySelector('.generation-main').textContent = view.primary;
    button.querySelector('.generation-time').textContent = view.secondary;
    button.setAttribute('aria-busy', String(view.mode === 'running'));
    button.setAttribute('aria-label', `${view.primary}. ${view.secondary}`); button.title = view.detail;
    if (generationTimer && view.mode === 'failed') {clearInterval(generationTimer); generationTimer = null;}
    if (!generationTimer && view.mode !== 'failed') generationTimer = setInterval(renderResearchProgress, 1000);
  }
  const pendingInputs = () => Boolean(state.file && state.bridge?.hasPendingInput?.(state.file));

  function message(text, error = false) {
    nodes.message.textContent = text;
    nodes.message.className = `studio-message${error ? ' error' : ''}`;
    nodes.message.hidden = !text;
  }
  function updateActions() {
    nodes.analyze.disabled = !state.file || !nodes.sheet.value || state.busy;
    nodes.edit.disabled = !state.file || state.busy;
    nodes.edit.title = state.file ? 'Abre el archivo seleccionado en el editor de datos.' : 'Primero abre o selecciona un archivo de investigación.';
    nodes.create.disabled = state.busy;
    if (nodes.refresh) nodes.refresh.disabled = state.busy;
    nodes.generate.disabled = !state.analysis?.ready || state.analysis?.source_id !== state.file || state.busy || state.job?.status === 'running' || pendingInputs();
    nodes.file.disabled = state.busy;
    if (nodes.upload) nodes.upload.disabled = state.busy;
    if (nodes.importFile) nodes.importFile.disabled = state.busy;
    nodes.dimensions?.querySelectorAll('button').forEach(node => {node.disabled = state.busy;});
    nodes.sheet.disabled = state.busy || !state.inspection?.sheets.length;
    [nodes.mode, nodes.pipeline, nodes.start, nodes.end].forEach(node => {node.disabled = state.busy;});
    if (state.analysis?.origins?.length > 1 && state.analysis.origins.every(origin => origin.kind === 'Compras')) nodes.pipeline.disabled = true;
    nodes.elements.querySelectorAll('input').forEach(node => {node.disabled = state.busy;});
  }
  function clearAnalysis() {
    state.analysis = null;
    nodes.output.replaceChildren($('div', {class: 'research-empty'}, state.file ? 'Selecciona una hoja con datos o complétala con «Editar datos» para explorarla.' : 'Abre un Excel de tu PC, retoma un archivo disponible o crea una plantilla vacía para comenzar.'));
    updateActions();
  }
  function sourceDetails() {
    const file = state.catalog?.files.find(f => f.file === state.file || f.id === state.file);
    nodes.details.replaceChildren();
    if (!file) {
      nodes.details.append($('p', {class: 'studio-note'}, 'Al elegir un archivo verás aquí su país, categoría, fabricante y hojas disponibles.'));
      return;
    }
    const meta = file.metadata || file;
    nodes.details.append($('strong', {}, `${meta.country} · ${meta.category} · ${meta.manufacturer}`),
      $('p', {class: 'studio-note research-source-name', title: file.name}, file.name),
      $('p', {class: 'source-note'}, file.read_only ? 'Al editar se crea una copia local; el original se conserva.' : 'Entrada local. «Editar datos» permite completarla o corregirla.'));
    if (state.inspection?.requires_input) nodes.details.append($('p', {class: 'studio-note'}, 'Entrada vacía · pendiente de llenar.'));
  }
  async function catalog() {
    const response = await state.bridge.api('/api/research/catalog');
    state.catalog = response;
    nodes.file.replaceChildren($('option', {value: ''}, 'Selecciona un archivo ya disponible…'));
    for (const file of response.files) {
      const meta = file.metadata || file;
      nodes.file.append($('option', {value: file.file || file.id, title: file.name}, `${meta.country} · ${meta.category} · ${meta.manufacturer} · ${file.name}${file.read_only ? ' · origen' : ' · local'}`));
    }
    const previousMode = nodes.mode.value || '1';
    nodes.mode.replaceChildren(...response.modes.map(m => $('option', {value: m.id}, m.name)));
    nodes.mode.value = previousMode;
    if (state.file) nodes.file.value = state.file;
    return response;
  }
  function chooseMatch(context) {
    const current = context?.file?.metadata || context?.file || {};
    const matches = (state.catalog?.files || []).filter(f => {
      const meta = f.metadata || f;
      return String(meta.country_code) === String(current.country_code) && String(meta.category_code) === String(current.category_code)
        && normalize(meta.manufacturer) === normalize(current.manufacturer);
    });
    return matches.find(f => !f.read_only) || matches[0];
  }
  function sameProject(context) {
    const current = context?.file?.metadata || context?.file || {};
    const file = state.catalog?.files.find(f => (f.file || f.id) === state.file);
    const selected = file?.metadata || file || {};
    return String(current.country_code) === String(selected.country_code) && String(current.category_code) === String(selected.category_code)
      && normalize(current.manufacturer) === normalize(selected.manufacturer);
  }
  async function open(fileId) {
    const request = ++state.request;
    clearTimeout(state.poll);
    clearTimeout(state.timer);
    state.file = fileId;
    state.analysis = null;
    state.inspection = null;
    state.job = null;
    state.busy = false;
    state.elements = null;
    nodes.file.value = fileId;
    nodes.sheet.replaceChildren();
    renderDimensions();
    nodes.start.value = nodes.end.value = '';
    clearAnalysis();
    renderJob();
    if (!fileId) { sourceDetails(); message(''); return; }
    state.busy = true;
    updateActions();
    message('Leyendo hojas y origen de los datos…');
    try {
      const [info, job] = await Promise.all([state.bridge.api(`/api/research/inspect?file=${encodeURIComponent(fileId)}`),
        state.bridge.api(`/api/research/job?file=${encodeURIComponent(fileId)}`)]);
      if (request !== state.request) return;
      state.inspection = info;
      state.job = job;
      nodes.sheet.replaceChildren(...info.sheets.map(s => $('option', {value: s.name}, `${s.title || s.name} · ${s.brand || ''}${s.has_data ? '' : ' · vacío'}`)));
      const context = state.bridge.coverageContext();
      const activeBrand = typeof context?.activeSheet === 'string' ? context.activeSheet.replace(/^P\d+_/, '') : context?.activeSheet?.brand || context?.activeSheet?.name;
      const match = info.sheets.find(s => normalize(s.brand) === normalize(activeBrand));
      nodes.sheet.value = match?.name || info.sheets.find(s => s.has_data)?.name || info.sheets[0]?.name || '';
      const pipeline = match && context?.loaded && sameProject(context) ? context.pipeline : null;
      nodes.pipeline.value = Number.isInteger(pipeline) && pipeline >= 0 && pipeline <= 6 ? String(pipeline) : 'default';
      sourceDetails();
      renderDimensions();
      renderJob();
      message(info.sheets.length ? 'Archivo listo. Explora los datos o ábrelos en el editor.' : 'No se encontraron hojas con el formato 5W1H. Revisa la estructura en el editor.', !info.sheets.length);
      if (job?.status === 'running') pollJob(job.id, fileId);
    } catch (error) {
      if (request === state.request) message(error.message, true);
    } finally {
      if (request === state.request) {state.busy = false; updateActions();}
    }
    if (request === state.request && nodes.sheet.value) await analyze();
  }
  async function importResearch(file) {
    if (!file || state.busy) return;
    state.busy = true; updateActions();
    message('Abriendo el archivo seleccionado…');
    let imported = null;
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error('Selecciona un libro Excel .xlsx.');
      if (file.size > 10 * 1024 * 1024) throw new Error('El archivo excede el límite de 10 MiB.');
      const bytes = new Uint8Array(await file.arrayBuffer());
      let binary = '';
      for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
      imported = await state.bridge.api('/api/studio/import', {kind: 'research', filename: file.name, content: btoa(binary), reuse_existing: true});
      await catalog();
      state.manual = true;
    } catch (error) {message(error.message, true);}
    finally {state.busy = false; nodes.upload.value = ''; updateActions();}
    if (imported) {
      await open(imported.file);
      if (state.file === imported.file && state.inspection) {
        message(`${imported.reused_existing ? 'Se abrió la entrada local existente' : 'Copia de trabajo lista'}: ${imported.name}. ${state.inspection.sheets.length} hojas disponibles en «Salidas del archivo».${state.analysis && !state.analysis.ready ? ' Revisa los avisos de la hoja seleccionada.' : ''}`);
      }
    }
  }
  function renderDimensions() {
    if (!nodes.dimensions) return;
    const sheets = state.inspection?.sheets || [];
    nodes.sidebar.hidden = !sheets.length;
    nodes.controlsPanel.hidden = !sheets.length;
    nodes.layout.classList.toggle('research-without-sheets', !sheets.length);
    nodes.dimensions.replaceChildren($('h2', {}, 'Salidas del archivo'));
    nodes.dimensions.hidden = !sheets.length;
    if (!sheets.length) return;
    nodes.dimensions.append($('p', {class: 'studio-note'}, `${sheets.length} hojas · elige una dimensión para ver sus gráficos y tablas.`));
    for (const sheet of sheets) {
      const button = $('button', {type: 'button', class: 'research-dimension', 'aria-pressed': sheet.name === nodes.sheet.value, disabled: state.busy});
      button.append($('strong', {}, `${sheet.title} · ${sheet.brand}`), $('span', {}, `${sheet.name} · ${sheet.has_data ? 'Con datos' : 'Pendiente de llenar'}`));
      button.addEventListener('click', () => {
        if (state.busy) return;
        nodes.sheet.value = sheet.name; state.elements = null; renderDimensions(); changed();
      });
      nodes.dimensions.append(button);
    }
    const missing = state.inspection.unavailable_segments || [];
    if (missing.length) {
      const detail = $('details'); detail.append($('summary', {}, `${missing.length} dimensiones sin hoja en este archivo`));
      detail.append($('p', {class: 'studio-note'}, missing.map(item => item.title).join(' · ')),
        $('p', {class: 'studio-note'}, 'Para obtener estas salidas, agrega sus hojas y datos desde Archivos de entrada o elige otro libro.'));
      nodes.dimensions.append(detail);
    }
  }
  function changed(reanalyze = true) {
    if (state.analysis) {
      state.analysis = null;
      message('Filtros cambiados. Actualiza los gráficos para revisar esta selección.');
      updateActions();
    }
    if (reanalyze) {
      clearTimeout(state.timer);
      state.timer = setTimeout(analyze, 250);
    }
  }
  async function analyze() {
    if (!state.file || !nodes.sheet.value || state.busy) return;
    const request = ++state.request;
    const fileId = state.file;
    state.busy = true;
    updateActions();
    message('Calculando gráficos y métricas con el motor 5W1H…');
    const payload = {file: fileId, sheet: nodes.sheet.value, mode: nodes.mode.value,
      start: nodes.start.value || null, end: nodes.end.value || null, elements: state.elements,
      pipeline: nodes.pipeline.value === 'default' ? null : Number(nodes.pipeline.value)};
    try {
      const analysis = await state.bridge.api('/api/research/analyze', payload);
      if (request !== state.request || state.file !== fileId) return;
      state.analysis = analysis;
      renderAnalysis();
      message(analysis.ready ? 'Investigación lista. Los filtros afectan la vista; el informe conserva las hojas completas.' : analysis.warnings.join('\n') || 'Completa datos reales para analizar esta hoja.', !analysis.ready);
    } catch (error) {
      if (request === state.request) {state.analysis = null; message(error.message, true);}
    } finally {
      if (request === state.request) {state.busy = false; updateActions();}
    }
  }
  function renderElements() {
    const names = state.analysis?.filters.elements || [];
    nodes.elements.replaceChildren();
    nodes.elementDetail.hidden = !names.length;
    const chosen = new Set(state.elements === null ? names : state.elements);
    for (const name of names) {
      const checkbox = $('input', {type: 'checkbox', checked: chosen.has(name)});
      const label = $('label');
      label.append(checkbox, document.createTextNode(name));
      checkbox.addEventListener('change', () => {
        const picked = new Set(state.elements === null ? names : state.elements);
        if (checkbox.checked) picked.add(name); else picked.delete(name);
        state.elements = [...picked];
        changed();
      });
      nodes.elements.append(label);
    }
  }
  function chartContent(chart, expanded = false) {
    const wrap = $('div', {class: 'research-chart-content'});
    const W = 720, H = expanded ? 370 : 300, pad = {left: 62, right: 20, top: 28, bottom: 52};
    const dates = chart.dates || [], series = chart.series || [];
    const data = series.flatMap(s => s.values.filter(finite));
    if (!data.length || !dates.length) {wrap.append($('p', {}, 'No hay observaciones en este período.')); return wrap;}
    let min = Math.min(0, ...data), max = Math.max(0, ...data);
    if (chart.axis_range?.length === 2 && chart.axis_range.every(finite)) [min, max] = chart.axis_range;
    if (max === min) max = min + 1;
    const spread = max - min;
    max += spread * .12;
    if (min < 0) min -= spread * .08;
    const x = index => dates.length === 1 ? (pad.left + W - pad.right) / 2 : pad.left + index / (dates.length - 1) * (W - pad.left - pad.right);
    const y = value => pad.top + (max - value) / (max - min) * (H - pad.top - pad.bottom);
    const svg = svgNode('svg', {viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': chart.title});
    const unit = series[0]?.unit || '';
    if (state.detail) {
      for (let i = 0; i <= 4; i++) {
        const value = min + (max - min) * i / 4;
        svg.append(svgNode('line', {x1: pad.left, x2: W - pad.right, y1: y(value), y2: y(value), stroke: '#e0e9ee'}),
          svgNode('text', {x: pad.left - 8, y: y(value) + 4, 'text-anchor': 'end', fill: '#607d93', 'font-size': 10}, number(value, unit)));
      }
    }
    const indices = [...new Set(Array.from({length: Math.min(6, dates.length)}, (_, i) => Math.round(i * (dates.length - 1) / Math.max(1, Math.min(6, dates.length) - 1))))];
    const groupWidth = (W - pad.left - pad.right) / dates.length;
    indices.forEach(i => svg.append(svgNode('text', {x: chart.type === 'bar' ? pad.left + groupWidth * (i + .5) : x(i), y: H - 25, 'text-anchor': 'middle', fill: '#607d93', 'font-size': 10}, period(dates[i]).slice(0, 26))));
    const readout = $('p', {class: 'studio-note'}, 'Pasa el cursor para ver el detalle de cada período.');
    const readable = (s, i) => `${period(dates[i])} · ${s.name}: ${number(s.values[i], s.unit)}`;
    const barWidth = Math.min(42, groupWidth * .82 / Math.max(1, series.length));
    const barX = (i, j) => pad.left + groupWidth * (i + .5) - barWidth * series.length / 2 + j * barWidth;
    series.forEach((s, seriesIndex) => {
      const color = /^#[0-9a-f]{6}$/i.test(s.color || '') ? s.color : '#078c86';
      if (chart.type === 'bar') {
        s.values.forEach((value, i) => {
          if (!finite(value)) return;
          const rect = svgNode('rect', {x: barX(i, seriesIndex), y: Math.min(y(0), y(value)), width: Math.max(1, barWidth - 1), height: Math.max(1, Math.abs(y(value) - y(0))), fill: color, opacity: .88});
          rect.append(svgNode('title', {}, readable(s, i)));
          rect.addEventListener('pointerenter', () => {readout.textContent = readable(s, i);});
          svg.append(rect);
          if (state.labels && dates.length * series.length <= 24) svg.append(svgNode('text', {x: barX(i, seriesIndex) + barWidth / 2, y: y(value) - 6, fill: color, 'font-size': 10, 'text-anchor': 'middle'}, number(value, s.unit)));
        });
      } else {
        let path = '', preceding = false;
        s.values.forEach((value, i) => {
          if (!finite(value)) {preceding = false; return;}
          path += `${preceding ? 'L' : 'M'}${x(i).toFixed(2)} ${y(value).toFixed(2)} `;
          preceding = true;
        });
        svg.append(svgNode('path', {d: path, fill: 'none', stroke: color, 'stroke-width': 2, 'stroke-dasharray': s.dash ? '6 4' : ''}));
        const validIndices = s.values.map((value, i) => finite(value) ? i : -1).filter(i => i >= 0);
        const labelIndices = new Set([validIndices[0], validIndices.at(-1)]);
        if (validIndices.length > 15) labelIndices.add(validIndices[Math.floor(validIndices.length / 2)]);
        validIndices.forEach(i => {
          if (state.detail) {
            const point = svgNode('circle', {cx: x(i), cy: y(s.values[i]), r: expanded ? 3.5 : 2.5, fill: color});
            point.append(svgNode('title', {}, readable(s, i)));
            point.addEventListener('pointerenter', () => {readout.textContent = readable(s, i);});
            svg.append(point);
          }
          if (state.labels && series.length <= 8 && labelIndices.has(i)) svg.append(svgNode('text', {x: x(i), y: y(s.values[i]) - 7 - (seriesIndex % 2) * 5, fill: color, 'font-size': 10, 'text-anchor': i === validIndices.at(-1) ? 'end' : 'start'}, number(s.values[i], s.unit)));
        });
      }
    });
    if (chart.type !== 'bar') svg.addEventListener('pointermove', event => {
      const bounds = svg.getBoundingClientRect(), coordinate = (event.clientX - bounds.left) / bounds.width * W;
      const i = Math.max(0, Math.min(dates.length - 1, Math.round((coordinate - pad.left) / (W - pad.left - pad.right) * (dates.length - 1))));
      readout.textContent = `${period(dates[i])} · ${series.filter(s => finite(s.values[i])).map(s => `${s.name}: ${number(s.values[i], s.unit)}`).join(' | ')}`;
    });
    const legend = $('div', {class: 'legend'});
    (chart.comparison_legend || series).forEach(s => {
      const item = $('span', {class: `research-series${s.dash ? ' compare' : ''}`}, s.name);
      item.style.setProperty('--series-color', s.color || '#078c86');
      legend.append(item);
    });
    wrap.append(svg, readout, legend);
    return wrap;
  }
  function purchaseComparisonContent(model) {
    const wrap=$('div',{class:'research-chart-content'}), stats=$('div',{class:'research-stats'});
    [[model.labels[0],model.total.base,''],[model.labels[1],model.total.comparison,''],['Diferencia vs. base',model.total.percent,'%']].forEach(([label,value,unit])=>{
      const card=$('article');card.append($('small',{},`${label} · MAT completo`),$('strong',{},number(value,unit)));stats.append(card);
    });
    wrap.append(stats);
    const rows=model.rows, W=1100,H=Math.max(300,rows.length*34+95),top=55,left=210,mid=685,zero=900;
    const svg=svgNode('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'Volumen MAT y diferencia entre fuentes'});
    const max=Math.max(1,...rows.flatMap(row=>[row.base,row.comparison]).filter(finite));
    const maxDelta=Math.max(1,...rows.map(row=>Math.abs(row.percent)).filter(finite));
    svg.append(svgNode('text',{x:left,y:22,fill:'#172c3c','font-size':14,'font-weight':700},'Volumen MAT'),
      svgNode('text',{x:mid+80,y:22,fill:'#172c3c','font-size':14,'font-weight':700},'Diferencia vs. base (%)'),
      svgNode('line',{x1:zero,x2:zero,y1:38,y2:H-38,stroke:'#bdcdd6'}));
    rows.forEach((row,index)=>{
      const y=top+index*34;
      const name=row.label.replace(/^T\.\s*/i,'');
      svg.append(svgNode('text',{x:left-10,y:y+4,'text-anchor':'end',fill:'#172c3c','font-size':11},name.slice(0,33)));
      [[row.base,-8,'#2c3e50'],[row.comparison,3,'#078c86']].forEach(([value,offset,color])=>{
        if(!finite(value))return;
        const bar=svgNode('rect',{x:left,y:y+offset,width:value/max*(mid-left-20),height:9,fill:color});
        bar.append(svgNode('title',{},`${row.left_label || 'Sin equivalente'} ↔ ${row.right_label || 'Sin equivalente'} · ${number(value)}`));svg.append(bar);
      });
      if(finite(row.percent)){
        const width=Math.abs(row.percent)/maxDelta*125,color=row.percent>0?'#238354':row.percent<0?'#ba4141':'#8c9aa3';
        svg.append(svgNode('rect',{x:row.percent<0?zero-width:zero,y:y-8,width,height:17,fill:color}),
          svgNode('text',{x:row.percent<0?zero-width-5:zero+width+5,y:y+4,'text-anchor':row.percent<0?'end':'start',fill:color,'font-size':11},`${row.percent>0?'+':''}${number(row.percent,'%')}`));
      } else svg.append(svgNode('text',{x:zero,y:y+4,'text-anchor':'middle',fill:'#637888','font-size':10},'Sin base comparable'));
    });
    const legend=$('div',{class:'legend'});
    model.labels.forEach((label,index)=>{const item=$('span',{class:'research-series'},label);item.style.setProperty('--series-color',index?'#078c86':'#2c3e50');legend.append(item);});
    wrap.append(svg,legend,$('p',{class:'studio-note'},model.note+' Las equivalencias de nombres solo alinean colores y posiciones; se conservan las etiquetas y los datos de cada fuente.'));
    return wrap;
  }
  function confidenceContent(model) {
    const view = model.confidence_view;
    const wrap = $('div', {class:'confidence-content'});
    const scroll = $('div', {class:'table-scroll', tabindex:'0', 'aria-label':model.title});
    const table = $('table', {class:'confidence-table'});
    const columns = $('colgroup');
    [90,150,95,95,310,110,225,125].forEach(width => {const col=$('col');col.style.width=`${width/12}%`;columns.append(col);});
    table.append(columns);
    const head=$('thead'), group=$('tr'), labels=$('tr');
    group.append($('th',{rowspan:2,scope:'col',class:'confidence-buyers-heading'},'Compradores · prom. 12 meses'),$('th',{rowspan:2,scope:'col'},'Marca'),$('th',{colspan:4,scope:'colgroup'},'Comparación MAT (%)'),$('th',{colspan:2,scope:'colgroup'},'Detalle de muestra'));
    ['Fabricante','Worldpanel','Intervalo WP','Fabricante vs. intervalo','Nivel de penetración','Error muestral'].forEach(label=>labels.append($('th',{scope:'col'},label)));
    head.append(group,labels);table.append(head);
    const pct=(value,ratio=false)=>finite(value)?`${(ratio?value*100:value).toFixed(1)}%`:'—';
    const signed=value=>finite(value)?`${value>0?'+':value<0?'−':''}${Math.abs(value).toFixed(1)}%`:'—';
    const color=value=>finite(value)&&value<0?'#C0392B':finite(value)&&value>0?'#168345':'#183143';
    const axis=view.axis, x=value=>28+(value-axis.min)/(axis.max-axis.min)*274;
    const body=$('tbody');
    view.rows.forEach(row=>{
      const line=$('tr');line.append($('td',{class:'confidence-buyers'},finite(row.buyers)?row.buyers.toLocaleString('es-MX',{maximumFractionDigits:0}):'—'),$('th',{scope:'row'},row.brand));
      [row.manufacturer,row.worldpanel].forEach((value,i)=>{const cell=$('td',{class:i?'confidence-wp':'confidence-manufacturer'},signed(value));cell.style.color=color(value);line.append(cell);});
      const interval=$('td',{class:'confidence-interval'});
      const graph=svgNode('svg',{viewBox:'0 0 330 68',role:'img','aria-label':`${row.brand}: intervalo WP ${signed(row.lower)} a ${signed(row.upper)}, fabricante ${signed(row.manufacturer)}, Worldpanel ${signed(row.worldpanel)}`});
      graph.append(svgNode('line',{x1:28,x2:302,y1:22,y2:22,stroke:'#CBD5DF'}),svgNode('line',{x1:x(0),x2:x(0),y1:5,y2:43,stroke:'#7C8F9B','stroke-dasharray':'3 3'}));
      if(finite(row.lower)&&finite(row.upper)){
        graph.append(svgNode('rect',{x:x(row.lower),y:18,width:Math.max(0,x(row.upper)-x(row.lower)),height:8,fill:'#B8D0DF'}));
        [row.lower,row.upper].forEach(value=>graph.append(svgNode('line',{x1:x(value),x2:x(value),y1:15,y2:29,stroke:'#B8D0DF'})));
      }
      if(finite(row.worldpanel))graph.append(svgNode('circle',{cx:x(row.worldpanel),cy:22,r:4,fill:'#293F52'}));
      if(finite(row.manufacturer))graph.append(svgNode('path',{d:`M${x(row.manufacturer)} 16 l6 6 -6 6 -6 -6 Z`,fill:'#BF9900'}));
      const bounds=svgNode('text',{x:165,y:58,'text-anchor':'middle','font-size':12});
      bounds.append(svgNode('tspan',{fill:color(row.lower)},signed(row.lower)),svgNode('tspan',{fill:'#526878'},' a '),svgNode('tspan',{fill:color(row.upper)},signed(row.upper)));
      graph.append(bounds);interval.append(graph);line.append(interval);
      const position=$('td',{class:'confidence-position'},row.position);position.style.color=row.position==='Dentro'?'#236B58':'#79571C';line.append(position);
      const level=$('td',{class:'confidence-level'});level.append($('span',{},row.penetration[0]));if(row.penetration[1])level.append($('strong',{},row.penetration[1]));line.append(level);
      const style=row.error_style, error=$('td',{class:'confidence-error'}), signal=$('div',{class:`confidence-signal${style.extreme?' extreme':''}`});
      signal.style.setProperty('--error-ink',`#${style.ink}`);signal.style.setProperty('--error-accent',`#${style.accent}`);signal.style.setProperty('--error-background',`#${style.background}`);
      if(style.extreme){
        const warning=svgNode('svg',{viewBox:'0 0 16 16','aria-hidden':'true',class:'confidence-warning'});
        warning.append(svgNode('path',{d:'M8 1.5 L15 14 H1 Z',fill:'none',stroke:'currentColor','stroke-width':1.3}),svgNode('text',{x:8,y:11,'text-anchor':'middle','font-size':10,fill:'currentColor'},'!'));signal.append(warning);
      }
      signal.append($('strong',{},pct(row.error,true)));error.append(signal);line.append(error);body.append(line);
    });
    table.append(body);
    const foot=$('tfoot'), axisRow=$('tr');
    axisRow.append($('td',{colspan:4}));
    const axisCell=$('td'), graph=svgNode('svg',{viewBox:'0 0 330 24',class:'confidence-axis','aria-label':'Escala de variación porcentual'});
    axis.ticks.forEach(tick=>graph.append(svgNode('text',{x:x(tick),y:15,'text-anchor':'middle','font-size':10,fill:'#526878'},`${Number(tick.toPrecision(8))}%`)));
    axisCell.append(graph);axisRow.append(axisCell,$('td',{colspan:3}));foot.append(axisRow);table.append(foot);
    scroll.append(table);wrap.append(scroll);
    const legend=$('div',{class:'confidence-legend'});
    [['interval','Intervalo WP'],['wp','Worldpanel'],['manufacturer','Fabricante']].forEach(([kind,label])=>{const item=$('span');item.append($('i',{class:`confidence-key ${kind}`,'aria-hidden':'true'}),document.createTextNode(label));legend.append(item);});
    wrap.append(legend);return wrap;
  }
  function renderAnalysis() {
    const analysis = state.analysis;
    nodes.output.replaceChildren();
    renderElements();
    if (!analysis) return;
    const origins = analysis.origins || [];
    const base = origins.find(origin => origin.role === 'Compras'), comparison = origins.find(origin => origin.role === 'Ventas');
    const purchaseVersions = origins.length > 1 && origins.every(origin => origin.kind === 'Compras');
    const leftOrigin = purchaseVersions ? base : comparison, rightOrigin = purchaseVersions ? comparison : base;
    nodes.originHelp.textContent = comparison ? `Izquierda: ${leftOrigin.label} · Derecha: ${rightOrigin?.label || 'Base'}.${purchaseVersions ? ' Ambas compras se comparan sin desfase.' : ''}` : base ? `Base: ${base.label}.` : '';
    if (nodes.mode.value === '3' && base) nodes.originHelp.textContent = `Vista del bloque base: ${base.label}.`;
    nodes.originHelp.hidden = !origins.length;
    nodes.mode.options[0].textContent = purchaseVersions ? 'Comparar versiones juntas' : 'Compras y ventas juntas';
    nodes.mode.options[1].textContent = purchaseVersions ? 'Comparar versiones por separado' : 'Compras y ventas separadas';
    nodes.mode.options[2].textContent = base && base.label !== 'Compras' ? `Solo base · ${base.label}` : 'Solo bloque base';
    if (purchaseVersions) nodes.pipeline.value = 'default';
    nodes.pipeline.labels[0].textContent = purchaseVersions ? 'Desfase entre versiones' : 'Pipeline de ventas';
    nodes.pipeline.options[0].textContent = purchaseVersions ? 'Sin desfase · compras' : 'Pipeline del archivo';
    if (analysis.warnings.length) nodes.output.append($('div', {class: 'studio-message error'}, analysis.warnings.join('\n')));
    const heading = $('div', {class: 'studio-heading'});
    const titles = $('div');
    titles.append($('h2', {}, analysis.title), $('p', {}, `${analysis.sheet} · ${analysis.charts.length} gráficos · ${analysis.tables.length} tablas · motor original 5W1H`));
    heading.append(titles);
    nodes.output.append(heading);
    const navigation = $('nav', {class: 'research-output-nav', 'aria-label': 'Gráficos y tablas de esta hoja'});
    [...analysis.charts.map(chart => [chart.id, chart.title, 'chart']), ...analysis.tables.map(table => [table.id, table.title, 'table'])].forEach(([id, title, kind]) => {
      const link = $('a', {href: `#research-${kind}-${id}`, class: 'button secondary'}, title);
      link.addEventListener('click', event => {event.preventDefault(); document.getElementById(`research-${kind}-${id}`)?.scrollIntoView({block: 'start', behavior: 'smooth'});});
      navigation.append(link);
    });
    nodes.output.append(navigation);
    if (analysis.stats.length) {
      const stats = $('div', {class: 'research-stats'});
      analysis.stats.forEach(stat => {const card = $('article'); card.append($('span', {}, stat.label), $('strong', {}, number(stat.value, stat.unit))); stats.append(card);});
      nodes.output.append(stats);
    }
    const comparisonModel=analysis.purchase_comparison;
    if(analysis.ready && comparisonModel?.ready && comparisonModel.rows.length){
      const panel=$('section',{class:'studio-panel',id:'research-purchase-comparison',tabindex:'0',role:'button','aria-label':'Ampliar comparación MAT entre compras'});
      panel.append($('h3',{},`Comparación entre fuentes · MAT ${period(comparisonModel.start)}–${period(comparisonModel.end)}`),purchaseComparisonContent(comparisonModel));
      const expand=()=>state.bridge.openChart('Comparación MAT entre compras',analysis.title,purchaseComparisonContent(comparisonModel));
      panel.addEventListener('click',expand);panel.addEventListener('keydown',event=>{if(event.key==='Enter'){event.preventDefault();expand();}});
      nodes.output.append(panel);
    }
    const charts = $('div', {class: 'research-charts'});
    analysis.charts.forEach(chart => {
      const card = $('article', {id: `research-chart-${chart.id}`, class: 'research-chart', tabindex: '0', role: 'button', 'aria-label': `Ampliar ${chart.title}`});
      const dashedOrigins = [...new Set(chart.series.filter(series => series.dash).map(series => series.origin))];
      card.append($('h3', {}, chart.title), $('p', {}, `Haz clic para ampliar.${dashedOrigins.length ? ` Líneas discontinuas: ${dashedOrigins.join(', ')}.` : ''}`), chartContent(chart));
      const expand = () => state.bridge.openChart(chart.title, `${analysis.title} · ${analysis.sheet}`, chartContent(chart, true));
      card.addEventListener('click', expand);
      card.addEventListener('keydown', event => {if (event.key === 'Enter' || event.key === ' ') {event.preventDefault(); expand();}});
      charts.append(card);
    });
    nodes.output.append(charts);
    const tables = $('div', {class: 'research-tables'});
    analysis.tables.forEach(model => {
      const panel = $('section', {id: `research-table-${model.id}`, class: 'studio-panel'});
      panel.append($('h3', {}, model.title));
      if(model.id==='confidence'&&model.confidence_view){
        panel.classList.add('confidence-panel');
        const expand=$('button',{type:'button',class:'button secondary confidence-expand'},'Ampliar ↗');
        expand.addEventListener('click',()=>state.bridge.openChart(model.title,analysis.title,confidenceContent(model)));
        panel.append(expand,confidenceContent(model));tables.append(panel);return;
      }
      const scroll = $('div', {class: 'table-scroll', tabindex: '0', 'aria-label': model.title});
      const table = $('table');
      const head = $('thead'), tr = $('tr');
      (model.display_columns || model.columns).forEach(column => tr.append($('th', {}, column)));
      head.append(tr); table.append(head);
      const body = $('tbody');
      model.rows.forEach(row => {
        const line = $('tr');
        row.forEach((value, index) => {
          const cell = $('td', {}, finite(value) ? number(value, model.column_units?.[index]) : value === null ? '—' : typeof value === 'object' ? JSON.stringify(value) : value);
          if (typeof value === 'string' && /^[-+]?\d+(?:\.\d+)?%$/.test(value)) {
            const variation = Number(value.replace('%', ''));
            if (String(row[0]).toLowerCase().includes('var') || String(row[0]).toLowerCase().includes('aporte')) {
              cell.style.background = variation > 0 ? '#c6efce' : variation < 0 ? '#ffc7ce' : '#f5f7f8';
              cell.style.color = variation > 0 ? '#006100' : variation < 0 ? '#9c0006' : '';
            }
          }
          line.append(cell);
        });
        body.append(line);
      });
      table.append(body); scroll.append(table); panel.append(scroll); tables.append(panel);
    });
    nodes.output.append(tables);
    const provenance = $('details', {class: 'research-provenance'});
    provenance.append($('summary', {}, 'Origen, cálculo y alcance del informe'), $('p', {}, 'La web explora el período y los elementos seleccionados. El PPT/Excel se genera con las hojas completas; no inventa regiones, canales o NSE a partir de cobertura.'),
      $('pre', {}, JSON.stringify(analysis.provenance, null, 2)));
    nodes.output.append(provenance);
    if (!analysis.ready) nodes.output.append($('div', {class: 'research-empty'}, 'Usa «Editar datos» para completar este segmento.'));
  }
  function renderJob() {
    nodes.results.hidden = !state.file;
    nodes.results.replaceChildren();
    const job = state.job?.file===state.file ? state.job : null;
    const panel = $('section', {class: 'studio-panel research-results', 'aria-label':'Resultados del archivo 5W1H seleccionado'});
    const heading=$('div',{class:'result-heading'}), identity=$('div');
    const title=job?.status==='complete' ? `${job.source_changed?'Entrega anterior · ':''}${job.title}` : job?.status==='running' ? 'Generando informe 5W1H…' : job?.status==='failed' ? 'La generación requiere revisión' : 'Aún no hay una entrega para este archivo';
    identity.append($('div',{class:'eyebrow'},'RESULTADOS 5W1H'),$('h2',{title},title));
    const headingActions=$('div',{class:'studio-actions'});
    const toggle=$('button',{class:'button secondary',type:'button','aria-expanded':!state.resultsCollapsed,'aria-controls':'research-results-content'},state.resultsCollapsed?'Mostrar resultados':'Contraer resultados');
    toggle.addEventListener('click',()=>{
      state.resultsCollapsed=!state.resultsCollapsed;
      try {localStorage.setItem('coveragelab-research-results-collapsed',String(state.resultsCollapsed));} catch (_) {}
      renderJob();
    });
    headingActions.append(toggle,nodes.generate);heading.append(identity,headingActions);panel.append(heading);
    const content=$('div',{id:'research-results-content',hidden:state.resultsCollapsed});
    const summary=$('div',{class:'result-state'});
    if(job) {
      summary.append($('span',{class:'pill'},job.status==='complete'?'Completado':job.status==='running'?'En curso':'Error'));
      if(job.started_at)summary.append($('span',{},new Date(job.started_at).toLocaleString('es-MX',{dateStyle:'medium',timeStyle:'short'})));
    }
    summary.append($('span',{},job?.status==='failed' ? job.error : job?.source_changed ? 'Los datos cambiaron. Puedes abrir la entrega anterior o generar otra.' : job?.status==='complete' ? 'Archivos de la última entrega de este 5W1H.' : 'Los archivos se habilitan al finalizar la generación.'));
    content.append(summary);
    const actions = $('div', {class: 'output-buttons',role:'group','aria-label':'Abrir resultados 5W1H en sus aplicaciones'});
    [['Excel', 'Abrir Excel'], ['PPT', 'Abrir PPT'], ['Provenance', 'Selección y origen'], ['Registro', 'Registro'], ['folder', 'Abrir carpeta']].forEach(([output, label]) => {
      const button = $('button', {class: 'button result-file', type: 'button', disabled: !!state.opening || !job || job.status !== 'complete' || output !== 'folder' && !job.available_outputs?.includes(output)}, label);
      const iconKind={Excel:'template',PPT:'presentation',Provenance:'selection',Registro:'log',folder:'folder'}[output];
      window.CoverageApp?.resultButton?.(button,label,iconKind);
      button.addEventListener('click', async () => {
        if(state.opening || !job || state.file!==job.file || state.job?.id!==job.id)return;
        const file=job.file;state.opening=output;renderJob();message(`Abriendo ${label.replace('Abrir ','')}…`);
        try {const result = await state.bridge.api('/api/research/open-output', {id: job.id, file, output});if(state.file===file && state.job?.id===job.id)message(result.message);}
        catch (error) {if(state.file===file && state.job?.id===job.id)message(error.message, true);}
        finally {state.opening=null;renderJob();}
      });
      if(state.opening===output)window.CoverageApp?.resultButton?.(button,'Abriendo…',iconKind);
      actions.append(button);
    });
    content.append(actions);
    if(job?.directory || job?.log) {
      const details=$('details',{class:'job-log-details'});details.append($('summary',{},'Ruta y registro'));
      if(job.directory)details.append($('p',{class:'studio-note'},job.directory));
      if(job.log)details.append($('pre',{},job.log));content.append(details);
    }
    panel.append(content);
    nodes.results.append(panel);
    renderResearchProgress();
    updateActions();
  }
  async function pollJob(id, fileId) {
    try {
      const job = await state.bridge.api(`/api/research/job?id=${encodeURIComponent(id)}`);
      if (state.file !== fileId) return;
      job.progressReceivedAt = Date.now();
      if (job.status === 'complete' && state.job?.status === 'running') job.completedObservedAt = Date.now();
      state.job = job; renderJob();
      if (job.status === 'running') state.poll = setTimeout(() => pollJob(id, fileId), 1800);
      else message(job.status === 'complete' ? 'Informe 5W1H generado. Ya puedes abrir el Excel, el PPT y la carpeta de esta entrega.' : job.error, job.status !== 'complete');
    } catch (error) {if (state.file === fileId) message(error.message, true);}
  }
  function reviewExport() {
    if (pendingInputs()) {message('Guarda las entradas pendientes antes de generar el informe 5W1H.', true); return;}
    if (!state.analysis?.ready || state.busy || state.job?.status === 'running') return;
    const dialog = $('dialog', {class: 'research-review'});
    const title = $('div', {class: 'dialog-heading'});
    title.append($('div', {class: 'eyebrow'}, 'REVISIÓN FINAL 5W1H'), $('h2', {}, 'Generar informe completo'),
      $('p', {}, `${state.analysis.file.name} · ${nodes.mode.selectedOptions[0].textContent}`));
    dialog.append(title, $('p', {}, 'Los filtros de fechas y elementos aplican a la exploración web. El informe conserva los datos completos de cada hoja incluida y se guarda en una carpeta nueva de Output.'));
    const selection = $('div', {class: 'dialog-selection'});
    const sheetInputs = [];
    for (const sheet of state.inspection.sheets) {
      const checkbox = $('input', {type: 'checkbox', checked: sheet.has_data, disabled: !sheet.has_data});
      const label = $('label');
      label.append(checkbox, document.createTextNode(`${sheet.title} · ${sheet.brand} · ${sheet.name}${sheet.has_data ? '' : ' · pendiente de llenar'}`));
      selection.append(label); sheetInputs.push([checkbox, sheet.name]);
    }
    dialog.append(selection);
    const context = state.bridge.coverageContext();
    const brandOverrides = {};
    const fromBrands = Boolean(context?.brandPipelines);
    const selectionValues = context?.loaded && sameProject(context) ? context.brandPipelines || context.selections || {} : {};
    Object.entries(selectionValues).forEach(([sheetName, pipeline]) => {
      const normalizedBrand = fromBrands ? sheetName : sheetName.replace(/^P\d+_/, '');
      const matching = state.inspection.sheets.filter(s => normalize(s.brand) === normalize(normalizedBrand));
      if (matching.length && Number.isInteger(pipeline) && pipeline >= 0 && pipeline <= 6) matching.forEach(s => {brandOverrides[s.brand] = pipeline;});
    });
    const purchaseVersions = state.analysis.origins?.length > 1 && state.analysis.origins.every(origin => origin.kind === 'Compras');
    dialog.append($('p', {class: 'studio-note'}, purchaseVersions ? 'Las versiones Compras_ conservan sus nombres y se comparan sin desfase. Si incluyes otras hojas con Ventas_n, esas hojas conservan sus pipelines de ventas.' : Object.keys(brandOverrides).length ? `Pipelines de cobertura que coinciden con las marcas: ${Object.entries(brandOverrides).map(([brand, p]) => `${brand}: P${p}`).join(' · ')}. Las demás hojas conservan su pipeline de origen${nodes.pipeline.value !== 'default' ? ` o el manual P${nodes.pipeline.value}` : ''}.` : nodes.pipeline.value === 'default' ? 'Se conserva el pipeline indicado en cada encabezado Ventas_n.' : `Se usará P${nodes.pipeline.value} como pipeline manual de las hojas con ventas.`));
    const thresholdLabel = $('label', {class: 'studio-note'}, 'Agrupar tamaños pequeños (%) · opcional');
    const threshold = $('input', {type: 'number', min: '.01', max: '99.99', step: '.01', placeholder: 'Sin agrupación'});
    thresholdLabel.append(threshold); dialog.append(thresholdLabel);
    const feedback = $('p', {class: 'studio-note'}); dialog.append(feedback);
    const actions = $('div', {class: 'dialog-actions'});
    const cancel = $('button', {class: 'button secondary', type: 'button'}, 'Volver a explorar');
    const confirm = $('button', {class: 'button primary', type: 'button'}, 'Generar PPT y Excel');
    cancel.addEventListener('click', () => dialog.close());
    confirm.addEventListener('click', async () => {
      if (pendingInputs()) {feedback.textContent = 'Guarda las entradas pendientes antes de generar el informe 5W1H.'; return;}
      const sheets = sheetInputs.filter(([input]) => input.checked).map(([, name]) => name);
      if (!sheets.length) {feedback.textContent = 'Incluye al menos una hoja con datos.'; return;}
      if (threshold.value && (!finite(Number(threshold.value)) || Number(threshold.value) <= 0 || Number(threshold.value) >= 100)) {feedback.textContent = 'El umbral debe estar entre 0 y 100%.'; return;}
      confirm.disabled = true; cancel.disabled = true;
      try {
        const job = await state.bridge.api('/api/research/export', {analysis_id: state.analysis.id, sheets, pipeline_overrides: brandOverrides, size_threshold: threshold.value ? Number(threshold.value) : null});
        job.progressReceivedAt = Date.now();
        state.job = job; dialog.close(); renderJob(); pollJob(job.id, state.file);
      } catch (error) {feedback.textContent = error.message; confirm.disabled = false; cancel.disabled = false;}
    });
    actions.append(cancel, confirm); dialog.append(actions);
    dialog.addEventListener('click', event => {if (event.target === dialog) {const rect = dialog.getBoundingClientRect(); if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) dialog.close();}});
    dialog.addEventListener('close', () => dialog.remove());
    document.body.append(dialog); dialog.showModal();
  }
  function init(container, bridge) {
    state.container = container; state.bridge = bridge;
    const heading = $('div', {class: 'studio-heading'}), headingText = $('div');
    headingText.append($('div', {class: 'eyebrow'}, 'MOMENTO 3 · INVESTIGACIÓN'), $('h1', {}, 'Investigar con 5W1H'), $('p', {}, 'Elige un archivo, explora sus hojas y genera el informe de marcas, canales, regiones y compradores.'));
    nodes.create = $('button', {id: 'research-create', class: 'button secondary', type: 'button'}, 'Crear plantilla vacía');
    nodes.edit = $('button', {id: 'research-edit', class: 'button secondary', type: 'button', disabled: true}, 'Editar datos');
    nodes.importFile = $('button', {id: 'research-open-pc', class: 'button primary', type: 'button'}, 'Abrir Excel de mi PC');
    nodes.upload = $('input', {id: 'research-upload', type: 'file', accept: '.xlsx', hidden: true, 'aria-label': 'Elegir archivo de investigación desde cualquier carpeta'});
    window.CoverageApp?.resultButton?.(nodes.importFile, 'Abrir Excel de mi PC', 'folder');
    nodes.importFile.addEventListener('click', () => nodes.upload.click());
    nodes.upload.addEventListener('change', () => importResearch(nodes.upload.files[0]));
    nodes.generate = $('button', {id:'research-generate', class: 'button primary', type: 'button', disabled: true}, 'Generar informe 5W1H');
    nodes.create.addEventListener('click', () => bridge.createResearch(bridge.coverageContext()));
    nodes.edit.addEventListener('click', () => bridge.openEditor(state.file, 'research', nodes.sheet.value));
    nodes.generate.addEventListener('click', reviewExport);
    heading.append(headingText);
    nodes.message = $('div', {class: 'studio-message', role: 'status', hidden: true});
    const layout = $('div', {class: 'studio-layout research-layout'}), aside = $('aside', {class: 'studio-sidebar'}), source = $('section', {id: 'research-source', class: 'studio-panel research-source-panel', 'aria-labelledby': 'research-source-title'});
    nodes.layout = layout; nodes.sidebar = aside;
    source.append($('h2', {id: 'research-source-title'}, 'Archivo de investigación'));
    const choices = $('div', {class: 'research-source-choices'});
    const fromPC = $('div', {class: 'research-source-choice'});
    fromPC.append($('span', {class: 'research-source-label'}, 'Desde tu computadora'), nodes.importFile,
      $('p', {class: 'studio-note'}, 'Excel 5W1H (.xlsx) de cualquier carpeta. Se usa una copia local.'));
    const existing = $('div', {class: 'research-source-choice research-source-existing'});
    existing.append($('label', {for: 'research-file', class: 'research-source-label'}, 'Archivos ya disponibles en CoverageLab'));
    nodes.file = $('select', {id: 'research-file', 'aria-label': 'Archivo de investigación 5W1H'});
    nodes.details = $('div', {id: 'research-source-details', class: 'research-source-details', 'aria-live': 'polite'});
    nodes.file.addEventListener('change', () => {state.manual = true; open(nodes.file.value);});
    const refresh = $('button', {id: 'research-refresh', class: 'button secondary', type: 'button', title: 'Actualiza la lista y vuelve a leer el archivo seleccionado.'}, 'Actualizar lista');
    nodes.refresh = refresh;
    refresh.addEventListener('click', async () => {try {await catalog(); if (state.file) await open(state.file);} catch (error) {message(error.message, true);}});
    refresh.addEventListener('click', () => {state.manual = true;});
    const picker = $('div', {class: 'research-source-picker'}); picker.append(nodes.file, refresh); existing.append(picker);
    const newEntry = $('div', {class: 'research-source-choice'});
    newEntry.append($('span', {class: 'research-source-label'}, 'Para empezar desde cero'), nodes.create,
      $('p', {class: 'studio-note'}, 'Crea un archivo sin datos, listo para llenar.'));
    choices.append(fromPC, existing, newEntry);
    const selected = $('div', {class: 'research-source-selected'}); selected.append(nodes.details, nodes.edit);
    source.append(choices, nodes.upload, selected);
    nodes.dimensions = $('section', {class: 'studio-panel padded research-dimensions', hidden: true}); aside.append(nodes.dimensions);
    const main = $('div');
    const controlsPanel = $('section', {class: 'studio-panel padded'}), controls = $('div', {class: 'research-controls'});
    nodes.controlsPanel = controlsPanel;
    const field = (label, node, wide = false) => {const wrap = $('div', {class: wide ? 'wide-field' : ''}); wrap.append($('label', {for: node.id}, label), node); controls.append(wrap);};
    nodes.sheet = $('select', {id: 'research-sheet'}); field('Hoja / dimensión', nodes.sheet, true);
    nodes.mode = $('select', {id: 'research-mode'}); field('Origen de los gráficos', nodes.mode);
    nodes.pipeline = $('select', {id: 'research-pipeline'});
    nodes.pipeline.append($('option', {value: 'default'}, 'Pipeline del archivo'), ...Array.from({length: 7}, (_, p) => $('option', {value: p}, `P${p}`)));
    field('Pipeline de ventas', nodes.pipeline);
    nodes.start = $('input', {id: 'research-start', type: 'month'}); field('Desde · vista web', nodes.start);
    nodes.end = $('input', {id: 'research-end', type: 'month'}); field('Hasta · corte del comparativo', nodes.end);
    nodes.sheet.addEventListener('change', () => {state.elements = null; renderDimensions(); changed();});
    [nodes.mode, nodes.pipeline, nodes.start, nodes.end].forEach(node => node.addEventListener('change', () => changed()));
    controlsPanel.append(controls);
    nodes.originHelp = $('p', {class: 'studio-note', role: 'status', hidden: true}); controlsPanel.append(nodes.originHelp);
    const originGuide = $('details', {class: 'studio-note'});
    originGuide.append($('summary', {}, 'Comparar dos versiones de compras'), $('p', {}, 'En la misma hoja, conserva Compras como encabezado de fecha del primer bloque: se verá a la izquierda. En el segundo bloque, donde antes iban ventas, escribe Compras_Mosaico, Compras_ajuste, Compras_ajt o Compras_ seguido del texto que quieras: se verá a la derecha. Cada bloque lleva sus fechas y categorías; la columna vacía entre ellos es opcional. También se admiten nombres personalizados en ambos bloques. Los nombres se mantienen en la web, el PPT y el Excel, sin aplicar desfase de ventas. Con Ventas_n se conserva el comparativo anterior: ventas a la izquierda y compras a la derecha.'));
    controlsPanel.append(originGuide);
    const viewActions = $('div', {class: 'studio-actions'});
    nodes.analyze = $('button', {class: 'button primary', type: 'button', disabled: true}, 'Actualizar gráficos');
    nodes.analyze.addEventListener('click', analyze);
    const reset = $('button', {class: 'button secondary', type: 'button'}, 'Toda la historia');
    reset.addEventListener('click', () => {nodes.start.value = nodes.end.value = ''; state.elements = null; changed();});
    const labels = $('label'), labelCheck = $('input', {type: 'checkbox', checked: true});
    labels.append(labelCheck, document.createTextNode('Etiquetas')); labelCheck.addEventListener('change', () => {state.labels = labelCheck.checked; renderAnalysis();});
    const details = $('label'), detailCheck = $('input', {type: 'checkbox', checked: true});
    details.append(detailCheck, document.createTextNode('Detalle visual')); detailCheck.addEventListener('change', () => {state.detail = detailCheck.checked; renderAnalysis();});
    viewActions.append(nodes.analyze, reset, labels, details); controlsPanel.append(viewActions);
    nodes.elementDetail = $('details', {hidden: true});
    nodes.elementDetail.append($('summary', {}, 'Elementos visibles'));
    const elementActions = $('div', {class: 'studio-actions'}), all = $('button', {class: 'button subtle', type: 'button'}, 'Todos'), none = $('button', {class: 'button subtle', type: 'button'}, 'Ninguno');
    all.addEventListener('click', () => {state.elements = null; changed();}); none.addEventListener('click', () => {state.elements = []; changed();});
    elementActions.append(all, none); nodes.elements = $('div', {class: 'research-elements'}); nodes.elementDetail.append(elementActions, nodes.elements); controlsPanel.append(nodes.elementDetail);
    nodes.results = $('div',{class:'research-results-dock'}); nodes.output = $('div');
    main.append(controlsPanel, nodes.output); layout.append(aside, main);
    container.replaceChildren(nodes.results,heading,source,nodes.message,layout); renderJob(); sourceDetails(); renderDimensions(); clearAnalysis();
  }
  async function activate(context) {
    try {
      const contextValue = context || state.bridge.coverageContext();
      const meta = contextValue?.file?.metadata || contextValue?.file || {};
      const contextKey = `${meta.country_code}/${meta.category_code}/${meta.manufacturer}`;
      if (state.contextKey !== null && state.contextKey !== contextKey) state.manual = false;
      state.contextKey = contextKey;
      if (!state.catalog) await catalog();
      const matching = chooseMatch(contextValue);
      if (!state.manual && contextValue?.file && !matching) {
        await open(null);
        message('Elige un archivo disponible, abre un Excel de tu PC o crea una plantilla vacía para este contexto.');
      } else if (!state.manual && matching && (matching.file || matching.id) !== state.file) await open(matching.file || matching.id);
      else updateActions();
      if (pendingInputs()) message('Hay entradas pendientes de guardar. La vista muestra los datos guardados; guarda los cambios en el editor antes de generar el informe.', true);
    } catch (error) {message(error.message, true);}
  }
  async function edited(fileId) {await catalog(); state.manual = true; await open(fileId);}
  window.StudioResearch = {init, activate, edited, open: async fileId => {if (!state.catalog) await catalog(); state.manual = true; await open(fileId);}};
})();
