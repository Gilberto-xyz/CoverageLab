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
  const pendingInputs = () => Boolean(state.file && state.bridge?.hasPendingInput?.(state.file));

  function message(text, error = false) {
    nodes.message.textContent = text;
    nodes.message.className = `studio-message${error ? ' error' : ''}`;
    nodes.message.hidden = !text;
  }
  function updateActions() {
    nodes.analyze.disabled = !state.file || !nodes.sheet.value || state.busy;
    nodes.edit.disabled = !state.file || state.busy;
    nodes.create.disabled = state.busy;
    if (nodes.refresh) nodes.refresh.disabled = state.busy;
    nodes.generate.disabled = !state.analysis?.ready || state.analysis?.source_id !== state.file || state.busy || state.job?.status === 'running' || pendingInputs();
    nodes.file.disabled = state.busy;
    if (nodes.upload) nodes.upload.disabled = state.busy;
    if (nodes.importFile) nodes.importFile.disabled = state.busy;
    nodes.dimensions?.querySelectorAll('button').forEach(node => {node.disabled = state.busy;});
    nodes.sheet.disabled = state.busy || !state.inspection?.sheets.length;
    [nodes.mode, nodes.pipeline, nodes.start, nodes.end].forEach(node => {node.disabled = state.busy;});
    nodes.elements.querySelectorAll('input').forEach(node => {node.disabled = state.busy;});
  }
  function clearAnalysis() {
    state.analysis = null;
    nodes.output.replaceChildren($('div', {class: 'research-empty'}, 'Selecciona una hoja con datos o llénala en el editor para explorarla.'));
    updateActions();
  }
  function sourceDetails() {
    const file = state.catalog?.files.find(f => f.file === state.file || f.id === state.file);
    nodes.details.replaceChildren();
    if (!file) return;
    const meta = file.metadata || file;
    nodes.details.append($('strong', {}, `${meta.country} · ${meta.category}`), $('p', {class: 'studio-note'}, meta.manufacturer),
      $('p', {class: 'studio-note research-source-name', title: file.name}, file.name),
      $('p', {class: 'source-note'}, file.read_only ? 'Origen 5W1H. Al guardar cambios se creará una copia local para continuar aquí.' : 'Entrada local. Puedes editarla y analizarla sin abrir Excel.'));
    if (state.inspection?.requires_input) nodes.details.append($('p', {class: 'studio-note'}, 'Entrada vacía · pendiente de llenar.'));
  }
  async function catalog() {
    const response = await state.bridge.api('/api/research/catalog');
    state.catalog = response;
    nodes.file.replaceChildren($('option', {value: ''}, 'Selecciona un archivo 5W1H'));
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
    if (!fileId) { sourceDetails(); return; }
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
        message(`${imported.reused_existing ? 'Se abrió la entrada local existente' : 'Copia de trabajo guardada en inputs/5w1h'}: ${imported.name}. ${state.inspection.sheets.length} hojas disponibles en «Salidas del archivo».${state.analysis && !state.analysis.ready ? ' Revisa los avisos de la hoja seleccionada.' : ''}`);
      }
    }
  }
  function renderDimensions() {
    if (!nodes.dimensions) return;
    const sheets = state.inspection?.sheets || [];
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
    series.forEach(s => {
      const item = $('span', {class: `research-series${s.dash ? ' compare' : ''}`}, s.name);
      item.style.setProperty('--series-color', s.color || '#078c86');
      legend.append(item);
    });
    wrap.append(svg, readout, legend);
    return wrap;
  }
  function renderAnalysis() {
    const analysis = state.analysis;
    nodes.output.replaceChildren();
    renderElements();
    if (!analysis) return;
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
    const charts = $('div', {class: 'research-charts'});
    analysis.charts.forEach(chart => {
      const card = $('article', {id: `research-chart-${chart.id}`, class: 'research-chart', tabindex: '0', role: 'button', 'aria-label': `Ampliar ${chart.title}`});
      card.append($('h3', {}, chart.title), $('p', {}, 'Haz clic en este cuadro para ampliar. Las líneas discontinuas corresponden a ventas.'), chartContent(chart));
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
    if (!analysis.ready) nodes.output.append($('div', {class: 'research-empty'}, 'Usa Llenar / editar datos para completar este segmento.'));
  }
  function renderJob() {
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
    updateActions();
  }
  async function pollJob(id, fileId) {
    try {
      const job = await state.bridge.api(`/api/research/job?id=${encodeURIComponent(id)}`);
      if (state.file !== fileId) return;
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
    dialog.append($('p', {class: 'studio-note'}, Object.keys(brandOverrides).length ? `Pipelines de cobertura que coinciden con las marcas: ${Object.entries(brandOverrides).map(([brand, p]) => `${brand}: P${p}`).join(' · ')}. Las demás hojas conservan su pipeline de origen${nodes.pipeline.value !== 'default' ? ` o el manual P${nodes.pipeline.value}` : ''}.` : nodes.pipeline.value === 'default' ? 'Se conserva el pipeline indicado en cada encabezado Ventas_n.' : `Se usará P${nodes.pipeline.value} como pipeline manual de las hojas con ventas.`));
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
    headingText.append($('div', {class: 'eyebrow'}, 'MOMENTO 3 · INVESTIGACIÓN'), $('h1', {}, 'Investigar con 5W1H'), $('p', {}, 'Explora marcas, tamaños, canales, regiones, NSE y sus aportes. Los datos adicionales se llenan en el mismo estudio.'));
    const actions = $('div', {class: 'studio-actions'});
    nodes.create = $('button', {class: 'button secondary', type: 'button'}, 'Nueva entrada 5W1H');
    nodes.edit = $('button', {class: 'button secondary', type: 'button', disabled: true}, 'Llenar / editar datos');
    nodes.importFile = $('button', {class: 'button secondary', type: 'button'}, 'Elegir archivo Excel');
    nodes.upload = $('input', {type: 'file', accept: '.xlsx', hidden: true, 'aria-label': 'Elegir archivo de investigación desde cualquier carpeta'});
    nodes.importFile.addEventListener('click', () => nodes.upload.click());
    nodes.upload.addEventListener('change', () => importResearch(nodes.upload.files[0]));
    nodes.generate = $('button', {class: 'button primary', type: 'button', disabled: true}, 'Generar informe 5W1H');
    nodes.create.addEventListener('click', () => bridge.createResearch(bridge.coverageContext()));
    nodes.edit.addEventListener('click', () => bridge.openEditor(state.file, 'research', nodes.sheet.value));
    nodes.generate.addEventListener('click', reviewExport);
    actions.append(nodes.importFile, nodes.upload, nodes.create, nodes.edit); heading.append(headingText, actions);
    nodes.message = $('div', {class: 'studio-message', role: 'status', hidden: true});
    const layout = $('div', {class: 'studio-layout'}), aside = $('aside', {class: 'studio-sidebar'}), source = $('section', {class: 'studio-panel padded'});
    source.append($('h2', {}, 'Fuente 5W1H'), $('label', {for: 'research-file'}, 'Archivo de investigación'));
    nodes.file = $('select', {id: 'research-file', 'aria-label': 'Archivo de investigación 5W1H'});
    nodes.details = $('div');
    nodes.file.addEventListener('change', () => {state.manual = true; open(nodes.file.value);});
    source.append(nodes.file, nodes.details, $('p', {class: 'studio-note'}, 'La cobertura aporta contexto y pipelines. Las regiones, canales y perfiles requieren sus propias observaciones.'));
    const refresh = $('button', {class: 'button subtle', type: 'button'}, 'Actualizar archivos');
    nodes.refresh = refresh;
    refresh.addEventListener('click', async () => {try {await catalog(); if (state.file) await open(state.file);} catch (error) {message(error.message, true);}});
    refresh.addEventListener('click', () => {state.manual = true;});
    source.append(refresh, $('p', {class: 'studio-note'}, 'Elige un Excel de cualquier carpeta. Se conserva el original y se usa una copia local; si ya existe idéntico, se abre sin duplicarlo.')); aside.append(source);
    nodes.dimensions = $('section', {class: 'studio-panel padded research-dimensions', hidden: true}); aside.append(nodes.dimensions);
    const main = $('div');
    const controlsPanel = $('section', {class: 'studio-panel padded'}), controls = $('div', {class: 'research-controls'});
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
    container.replaceChildren(nodes.results,heading,nodes.message,layout); renderJob(); clearAnalysis();
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
        message('No hay una entrada 5W1H para este contexto. Crea una entrada o elige un archivo de investigación para continuar.');
      } else if (!state.manual && matching && (matching.file || matching.id) !== state.file) await open(matching.file || matching.id);
      else if (!state.manual && !state.file && !contextValue?.file && state.catalog.files.length) await open(state.catalog.files[0].file || state.catalog.files[0].id);
      else updateActions();
      if (pendingInputs()) message('Hay entradas pendientes de guardar. La vista muestra los datos guardados; guarda los cambios en el editor antes de generar el informe.', true);
    } catch (error) {message(error.message, true);}
  }
  async function edited(fileId) {await catalog(); state.manual = true; await open(fileId);}
  window.StudioResearch = {init, activate, edited, open: async fileId => {if (!state.catalog) await catalog(); state.manual = true; await open(fileId);}};
})();
