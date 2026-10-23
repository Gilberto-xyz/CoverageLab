/* One row per input sheet; the chosen pipeline is shared with the detail view. */
(() => {
  'use strict';
  const el = (tag, text, cls) => {const n = document.createElement(tag); if (text != null) n.textContent = text; if (cls) n.className = cls; return n;};
  const valid = sheet => !sheet.error && sheet.pipelines?.some(p => p.available);
  const eligible = sheet => Number.isFinite(sheet.buyers) && sheet.buyers >= 200;
  function preset(state, mode) {
    state.pptSlides ||= {};
    for (const sheet of state.analysis.sheets) {
      if (mode === 'summary' || mode === 'slides-eligible') {
        state.pptSlides[sheet.name] = mode === 'slides-eligible' && valid(sheet) && eligible(sheet) && !!state.included[sheet.name];
        continue;
      }
      const include = valid(sheet) && (mode === 'all' || eligible(sheet));
      state.included[sheet.name] = include;
      state.pptSlides[sheet.name] = include;
    }
  }
  function counts(state) {
    const included = state.analysis.sheets.filter(s => valid(s) && state.included[s.name]);
    return {summary: included.length, slides: included.filter(s => state.pptSlides?.[s.name] !== false).length,
      low: included.filter(s => Number.isFinite(s.buyers) && s.buyers < 200).length,
      unknown: included.filter(s => !Number.isFinite(s.buyers)).length};
  }
  function render(root, state, {fmt, month, changed, inspect, extraTitles = {}}, review = false) {
    const filter = root.querySelector('input[type="search"]')?.value || '';
    const wasOpen = root.querySelector('details')?.open ?? true;
    const previousScroll = root.querySelector('.coverage-summary-scroll');
    const position = [previousScroll?.scrollTop || 0, previousScroll?.scrollLeft || 0];
    root.replaceChildren();
    const details = el('details'); details.open = wasOpen;
    const title = el('summary', review ? 'Resumen y diapositivas a generar' : 'Resumen de todas las marcas');
    const count = counts(state);
    title.append(el('span', `${count.summary} en resumen · ${count.slides} con slides`, 'coverage-summary-count'));
    details.append(title);
    const toolbar = el('div', null, 'coverage-summary-tools');
    const search = el('input'); search.type = 'search'; search.placeholder = 'Buscar marca, hoja o categoría…'; search.value = filter;
    search.setAttribute('aria-label', 'Filtrar el resumen de coberturas');
    const presets = [['all', 'Incluir todas'], ['eligible', 'Solo ≥200 compradores'], ['slides-eligible', 'Slides solo ≥200'], ['summary', 'Solo resumen']];
    for (const [mode, label] of presets) {
      const button = el('button', label, 'button secondary'); button.type = 'button'; button.disabled = state.busy;
      button.addEventListener('click', () => {if (state.busy) return; preset(state, mode); changed();}); toolbar.append(button);
    }
    toolbar.append(search); details.append(toolbar);
    details.append(el('p', 'Resumen incluye la marca en la tabla del PPT, el Excel y el banco. Slides agrega sus diapositivas básicas y extras. «Solo ≥200» usa el promedio exacto de compradores de los últimos 12 meses, antes de redondearlo; sin dato quedan fuera de ese filtro.', 'table-note'));
    const scroll = el('div', null, 'table-scroll coverage-summary-scroll'), table = el('table');
    const head = el('thead'), header = el('tr');
    const labels = ['Resumen', 'Slides', 'Marca / hoja', 'Pipeline elegido', 'Corte MAT', 'Compradores · prom. 12 meses', 'Penetración media', 'VAR MAT Sell-in', 'VAR MAT Worldpanel', 'Cobertura actual', 'Cobertura anterior', 'Estabilidad (pp)', 'Correlación MAT'];
    labels.forEach(label => {const th = el('th', label); th.scope = 'col'; header.append(th);}); head.append(header); table.append(head);
    const body = el('tbody');
    for (const sheet of state.analysis.sheets) {
      const row = el('tr'); row.dataset.search = `${sheet.brand} ${sheet.name} ${sheet.category || ''}`.toLocaleLowerCase();
      row.className = !valid(sheet) ? 'coverage-invalid' : !state.included[sheet.name] ? 'coverage-omitted' : '';
      if (Number.isFinite(sheet.buyers) && sheet.buyers < 200) row.className += ' coverage-low-buyers';
      for (const kind of ['summary', 'slides']) {
        const cell = el('td'), check = el('input'); check.type = 'checkbox';
        check.checked = kind === 'summary' ? !!state.included[sheet.name] : !!state.included[sheet.name] && state.pptSlides?.[sheet.name] !== false;
        check.disabled = state.busy || !valid(sheet) || (kind === 'slides' && !state.included[sheet.name]);
        check.dataset.unavailable = String(!valid(sheet) || (kind === 'slides' && !state.included[sheet.name]));
        check.setAttribute('aria-label', `${kind === 'summary' ? 'Incluir en resumen' : 'Generar diapositivas de'} ${sheet.name}`);
        check.addEventListener('change', () => {
          if (state.busy) return;
          if (kind === 'summary') state.included[sheet.name] = check.checked;
          else {(state.pptSlides ||= {})[sheet.name] = check.checked;}
          changed();
        }); cell.append(check); row.append(cell);
      }
      const name = el('td'), label = el(review ? 'strong' : 'button', sheet.brand || sheet.name, review ? '' : 'coverage-summary-brand');
      if (!review) {label.type = 'button'; label.disabled = state.busy; label.title = `Analizar ${sheet.name}`; label.addEventListener('click', () => {if (!state.busy) inspect(sheet.name);});}
      name.append(label, el('small', `${sheet.category || ''} · ${sheet.name}`));
      if (review && state.included[sheet.name] && state.pptSlides?.[sheet.name] !== false) {
        const extras = state.pptExtras?.[sheet.name] || [];
        if (extras.length) name.append(el('small', `Extras: ${extras.map(key => extraTitles[key] || key).join(', ')}`));
      }
      if (sheet.error) name.append(el('small', sheet.error, 'coverage-summary-error'));
      else if (!Number.isFinite(sheet.buyers)) name.append(el('small', 'Sin dato de compradores', 'coverage-summary-error'));
      else if (!eligible(sheet)) name.append(el('small', 'Menos de 200 compradores · uso interno', 'coverage-summary-error'));
      row.append(name);
      const pipelineCell = el('td'), select = el('select'); select.setAttribute('aria-label', `Pipeline de ${sheet.name}`);
      for (const p of sheet.pipelines || []) {const option = el('option', `P${p.pipeline}`); option.value = String(p.pipeline); option.disabled = !p.available; select.append(option);}
      select.value = String(state.selections[sheet.name] ?? ''); select.disabled = state.busy || !valid(sheet);
      select.dataset.unavailable = String(!valid(sheet));
      select.addEventListener('change', () => {if (state.busy) return; state.selections[sheet.name] = Number(select.value); changed();}); pipelineCell.append(select); row.append(pipelineCell);
      const pipeline = sheet.pipelines?.[state.selections[sheet.name]], averages = sheet.template?.averages || {};
      const data = [sheet.dates?.length ? month(sheet.dates.at(-1)) : '—', fmt(sheet.buyers, 0), fmt(averages.Penet_MAT_Actual, 1, '%'),
        fmt(pipeline?.client_variation, 1, '%'), fmt(pipeline?.wp_variation, 1, '%'), fmt(pipeline?.coverage, 1, '%'),
        fmt(pipeline?.previous_coverage, 1, '%'), fmt(pipeline?.stability, 1), fmt(pipeline?.correlation, 3)];
      data.forEach((value, index) => {const cell = el('td', value); if ([3, 4, 7].includes(index)) {
        const raw = [pipeline?.client_variation, pipeline?.wp_variation, pipeline?.stability][[3, 4, 7].indexOf(index)];
        cell.className = Number.isFinite(raw) ? raw < 0 ? 'coverage-negative' : raw > 0 ? 'coverage-positive' : '' : '';
      } row.append(cell);}); body.append(row);
    }
    const applyFilter = () => {for (const row of body.children) row.hidden = !row.dataset.search.includes(search.value.toLocaleLowerCase());};
    search.addEventListener('input', applyFilter); applyFilter(); table.append(body); scroll.append(table); details.append(scroll);
    if (count.low || count.unknown) details.append(el('p', `${count.low} marca(s) incluidas con menos de 200 compradores${count.unknown ? ` · ${count.unknown} sin dato de compradores` : ''}. Puedes conservarlas solo en el resumen o excluirlas con el filtro.`, 'table-note coverage-summary-error'));
    root.append(details);
    scroll.scrollTop = position[0]; scroll.scrollLeft = position[1];
  }
  function setBusy(root, busy) {
    if (!root) return;
    root.inert = !!busy;
    root.querySelectorAll('button,input,select').forEach(control => {control.disabled = !!busy || control.dataset.unavailable === 'true';});
  }
  window.CoverageSummary = {render, counts, preset, eligible, setBusy};
})();
