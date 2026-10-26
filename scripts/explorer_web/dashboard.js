'use strict';
// Panel de lectura: trabaja exclusivamente con el análisis cargado, sin exportar.
window.CoverageDashboard = (() => {
  const ns = 'http://www.w3.org/2000/svg';
  const node = (tag, text, cls) => { const el = document.createElement(tag); if (text != null) el.textContent = text; if (cls) el.className = cls; return el; };
  const svg = (tag, attrs, text) => { const el = document.createElementNS(ns, tag); Object.entries(attrs).forEach(([k,v]) => el.setAttribute(k, v)); if (text != null) el.textContent = text; return el; };
  const fmt = (v, digits = 1, suffix = '') => v == null ? '—' : new Intl.NumberFormat('es-MX', {maximumFractionDigits: digits, minimumFractionDigits: digits}).format(v) + suffix;
  const compact = v => new Intl.NumberFormat('es-MX', {notation:'compact',maximumFractionDigits:1}).format(v);
  const month = date => new Intl.DateTimeFormat('es-MX', {month:'short',year:'2-digit'}).format(new Date(date + 'T12:00:00'));
  const valid = v => v != null && Number.isFinite(v);
  const signed = v => v == null ? '—' : (v > 0 ? '+' : '') + fmt(v, 1, '%');

  // Recorta extremos sin base común y conserva los huecos del calendario interior.
  function chartData(model, window = 'all') {
    const n=model.dates.length, start=model.annual || window==='all' ? 0 : Math.max(0,n-Number(window));
    const complete=model.dates.map((_,i)=>model.series.length>0 && model.series.every(s=>valid(s.values[i])));
    const first=complete.findIndex((value,i)=>i>=start && value);
    const end=first<0 ? 0 : complete.findLastIndex(value=>value)+1;
    const begin=first<0 ? 0 : first;
    return {...model,dates:model.dates.slice(begin,end),series:model.series.map(s=>({...s,
      values:s.values.slice(begin,end).map((value,i)=>complete[begin+i] ? value : null),
      ...(s.yoy ? {yoy:s.yoy.slice(begin,end)} : {})}))};
  }
  function kpiDigits(kpi, value) {
    return kpi.key==='buyers_series' ? 0 : ['purchase_average','purchase_occasion'].includes(kpi.key) && valid(value) && Math.abs(value)>0 && Math.abs(value)<.1 ? 4 : 2;
  }
  function kpiAverages(container, kpi,language='ES') {
    const lang=window.CoverageI18n?.get(language),fmtValue=lang?.number||fmt,dateLabel=lang?.month||month;
    const suffix=kpi.key==='penetration' ? '%' : '';
    const summary=node('dl',null,'kpi-averages');
    for(const [label,value,count,start] of [[lang?.average12||'Promedio 12 meses',kpi.average,kpi.recent_count,kpi.recent_start],[lang?.averageTotal||'Promedio total',kpi.total_average,kpi.total_count,kpi.history_start]]) {
      const entry=node('div');entry.append(node('dt',label),node('dd',fmtValue(value,kpiDigits(kpi,value),suffix)));
      entry.title=`${dateLabel(start)} – ${dateLabel(kpi.end)} · ${count}`;summary.append(entry);
    }
    container.append(summary);
    const note=language==='PT'?`Médias mensais · 12 meses: ${kpi.recent_count} dados · Total: ${kpi.total_count} dados até ${dateLabel(kpi.end)}`:language==='EN'?`Monthly averages · 12 months: ${kpi.recent_count} observations · Overall: ${kpi.total_count} observations through ${dateLabel(kpi.end)}`:`Promedios mensuales · 12 meses: ${kpi.recent_count} datos · Total: ${kpi.total_count} datos hasta ${dateLabel(kpi.end)}`;
    container.append(node('p',note,'kpi-average-note'));
  }

  function axisScale(values, topPadding = .18, bottomPadding = .08) {
    const numbers = values.filter(valid);
    const min = Math.min(0,...numbers), max = Math.max(0,...numbers);
    const span = max-min;
    return span ? {min:min < 0 ? min-span*bottomPadding : 0,max:max+span*topPadding} : {min:0,max:1};
  }
  function axisTick(value, span,locale='es-MX') {
    if(Math.abs(value)>=10000)return new Intl.NumberFormat(locale,{notation:'compact',maximumFractionDigits:1}).format(value);
    const step=Math.abs(span)/4;
    const digits=step ? Math.min(8,Math.max(0,1-Math.floor(Math.log10(step)))) : 1;
    return new Intl.NumberFormat(locale,{maximumFractionDigits:digits}).format(value);
  }
  function correlationStyle(value, samples, colors) {
    if(!valid(value))return null;
    const sorted=samples.filter(valid).sort((a,b)=>a-b);
    if(!sorted.length)return null;
    const low=sorted[0],high=sorted.at(-1),mid=(sorted[Math.floor((sorted.length-1)/2)]+sorted[Math.ceil((sorted.length-1)/2)])/2;
    const mix=(from,to,amount)=>'#'+[0,2,4].map(i=>{
      const a=parseInt(from.replace('#','').slice(i,i+2),16),b=parseInt(to.replace('#','').slice(i,i+2),16);
      return Math.round(a+(b-a)*Math.max(0,Math.min(1,amount))).toString(16).padStart(2,'0');
    }).join('').toUpperCase();
    const fill=low===high ? colors[1] : value<=low ? colors[0] : value>=high ? colors[2] : value<=mid ? mix(colors[0],colors[1],(value-low)/(mid-low)) : mix(colors[1],colors[2],(value-mid)/(high-mid));
    return {fill,ink:'#172c3c',low,mid,high};
  }

  function models(sheet, pipeline, colors) {
    const p = sheet.pipelines[pipeline];
    const wp = (values, kind = 'line') => ({name:'Worldpanel',values,color:colors.worldpanel,kind});
    const si = (values, kind = 'line') => ({name:`Sell-in P${pipeline}`,values,color:colors.sellin,kind});
    const base = (key,title,series,description,percent = false) => ({key,title,series,description,dates:sheet.dates,percent});
    const result = [
      base('coverage','Cobertura MAT y penetración', [{name:`Cobertura P${pipeline}`,values:p.coverage_series,color:colors.coverage,kind:'bar'}, {name:'Penetración mensual',values:sheet.penetration,color:colors.penetration,kind:'line'}], 'Cobertura MAT y penetración mensual en el mismo eje (%)',true),
      base('monthly','Tendencia mensual de volumen',[si(p.sell_in),wp(sheet.sell_out)],`${sheet.unit} · Sell-in desplazado ${pipeline} mes(es)`),
      base('mat','Tendencia de volumen MAT',[si(p.sell_in_mat),wp(sheet.wp_mat)],`${sheet.unit} · Sumas móviles de 12 meses`),
      base('yoy','Evolución y variación interanual MAT',[si(p.yoy),wp(sheet.wp_yoy)],'Variación de volúmenes MAT vs. los 12 meses anteriores',true),
    ];
    for (const kpi of sheet.dashboard.kpis) result.push({...base(kpi.key,kpi.label,[{name:kpi.label,values:sheet[kpi.key],color:kpi.color,kind:'line',digits:kpi.key === 'buyers_series' ? 0 : 2}], 'Serie mensual original · Mismo color del Excel', kpi.key === 'penetration'),kpi});
    const periods = sheet.dashboard.mat_periods;
    result.push({key:'annual_mat',title:'Volumen MAT, YoY y cobertura por año',dates:periods.map(row => row.date),annual:true,
      description:'Mismo mes de corte en cada año · Barras de volumen, etiquetas YoY y cobertura en eje derecho',
      series:[{...si(periods.map(row=>row.sell_in[pipeline]),'bar'),yoy:periods.map(row=>row.sell_in_yoy[pipeline])},
              {...wp(periods.map(row=>row.worldpanel),'bar'),yoy:periods.map(row=>row.worldpanel_yoy)},
              {name:`Cobertura P${pipeline}`,values:periods.map(row=>row.coverage[pipeline]),color:colors.comparison,kind:'line',axis:'right',percent:true}]});
    return result;
  }

  function plot(container, model, options, expanded = false) {
    container.replaceChildren();
    const lang=window.CoverageI18n?.get(options.language||'ES');
    const dateLabel=lang?.month||month,numberLabel=lang?.number||fmt,compactLabel=lang?.compact||compact,signedLabel=lang?.signed||signed;
    if(lang)model=lang.model(model);
    model=chartData(model,options.window);
    if(model.kpi)kpiAverages(container,model.kpi,options.language);
    const period=node('p',model.dates.length ? `${dateLabel(model.dates[0])} – ${dateLabel(model.dates.at(-1))} · ${lang?.data||'Datos disponibles en todas las series'}` : lang?.noPeriod||'Sin período compartido con datos','chart-period');container.append(period);
    const n = model.dates.length;
    const first = 0;
    const baseWidth = expanded ? 900 : 540;
    const width = options.labels && options.scope === 'all' ? Math.max(baseWidth,(n-first)*58) : baseWidth;
    const graph = svg('svg',{viewBox:`0 0 ${width} 300`,role:'img',tabindex:0,'aria-label':model.title});
    if (width > baseWidth) graph.style.width = `${width}px`;
    graph.append(svg('title',{},`${model.title}. ${model.description}`));
    const scroll = node('div',null,'dashboard-chart-scroll'); scroll.append(graph); container.append(scroll);
    const tooltip = node('p',null,'dashboard-tooltip'); container.append(tooltip);
    const legend = node('div',null,'dashboard-legend');
    model.series.forEach(series => { const text = node('span',series.name); text.style.setProperty('--series-color',series.color); legend.append(text); }); container.append(legend);
    if (!n) { graph.append(svg('text',{x:width/2,y:130,'text-anchor':'middle',fill:'#637888','font-size':14},lang?.noBase||'Sin base común en el período elegido')); tooltip.textContent=lang?.insufficient||'Las bases insuficientes se muestran como —.'; return; }
    const dual = model.series.some(s=>s.axis === 'right');
    const left = 59, right = width - (dual ? 69 : 25), top = 38, bottom = 248;
    const x = i => left + (i-first+.5) / Math.max(1,n-first)*(right-left);
    const scale = axis => {
      if(model.scales?.[axis])return model.scales[axis];
      const values = model.series.filter(s=>(s.axis || 'left')===axis).flatMap(s=>s.values.slice(first)).filter(valid);
      return axisScale(values);
    };
    const scales = {left:scale('left'),right:scale('right')};
    const y = (v,axis='left') => bottom - (v-scales[axis].min)/(scales[axis].max-scales[axis].min)*(bottom-top);
    for (let i=0;i<=4;i++) {
      const py=bottom-i/4*(bottom-top), value=scales.left.min+i/4*(scales.left.max-scales.left.min);
      if (options.detail) graph.append(svg('line',{x1:left,x2:right,y1:py,y2:py,stroke:'#e5edf1',class:'dashboard-gridline'}));
      graph.append(svg('text',{x:left-8,y:py+4,'text-anchor':'end','font-size':11,fill:'#637888',class:'dashboard-axis-tick'},axisTick(value,scales.left.max-scales.left.min,lang?.locale)+(model.percent ? '%' : '')));
      if (dual) graph.append(svg('text',{x:right+7,y:py+4,'font-size':11,fill:'#008c95'},axisTick(scales.right.min+i/4*(scales.right.max-scales.right.min),scales.right.max-scales.right.min,lang?.locale)+'%'));
    }
    const step = Math.max(1,Math.ceil((n-first)/7));
    for (let i=first;i<n;i+=step) graph.append(svg('text',{x:x(i),y:bottom+23,'text-anchor':'middle','font-size':11,fill:'#637888'},`${model.annual ? 'MAT ' : ''}${dateLabel(model.dates[i])}`));
    const barSeries=model.series.filter(s=>s.kind==='bar');
    const barWidth=Math.min(model.annual ? 54 : 22,(right-left)/Math.max(1,n-first)*.65/Math.max(1,barSeries.length));
    const marks=[];
    const labelIndexes=series => {
      const indexes=series.values.map((value,i)=>valid(value) && i>=first ? i : -1).filter(i=>i>=0);
      if (!options.labels || !indexes.length) return new Set();
      if (model.annual || options.scope==='all') return new Set(indexes);
      if (options.scope==='last') return new Set([indexes.at(-1)]);
      const cutoff=model.dates.at(-1).slice(5,7);
      return new Set([indexes[0],indexes.at(-1),...indexes.filter(i=>model.dates[i].slice(5,7)===cutoff)]);
    };
    model.series.forEach(series => {
      let d='',connected=false;
      const axis=series.axis || 'left', labels=labelIndexes(series);
      for (let i=first;i<n;i++) {
        const value=series.values[i];
        if (!valid(value)) { connected=false; continue; }
        const px=x(i)+(series.kind==='bar' && barSeries.length>1 ? (barSeries.indexOf(series)-(barSeries.length-1)/2)*barWidth*1.15 : 0), py=y(value,axis);
        if (series.kind==='bar') graph.append(svg('rect',{x:px-barWidth/2,y:Math.min(py,y(0,axis)),width:barWidth,height:Math.abs(py-y(0,axis)),fill:series.color,stroke:options.detail ? '#637888' : 'none','stroke-width':.5,class:'dashboard-bar','data-series':series.name}));
        else { d+=`${connected ? 'L' : 'M'}${px},${py} `; connected=true; if(options.detail) graph.append(svg('circle',{cx:px,cy:py,r:2.5,fill:series.color,class:'dashboard-point'})); }
        if (labels.has(i)) marks.push({series,i,px,py,value});
      }
      if(series.kind!=='bar') graph.append(svg('path',{d,fill:'none',stroke:series.color,'stroke-width':2.6,'stroke-linejoin':'round','stroke-linecap':'round',class:'dashboard-series','data-series':series.name}));
    });
    const boxes=[],yoyMarks=[];
    marks.forEach(({series,i,px,py,value})=>{
      const isPercent=series.percent || model.percent;
      const text=isPercent ? numberLabel(value,1,'%') : model.annual ? compactLabel(value) : series.digits===2 ? numberLabel(value,2) : compactLabel(value);
      const half=Math.max(17,text.length*3.4), lx=Math.max(left+half,Math.min(right-half,px));
      let ly=Math.max(16,py-11);
      while(boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18)) ly+=20;
      if(ly>bottom-5) { ly=Math.min(bottom-5,py+18); while(boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18)) ly-=20; }
      boxes.push({x:lx,y:ly,half});
      const group=svg('g',{class:'dashboard-data-label','data-series':series.name});
      group.append(svg('title',{},`${dateLabel(model.dates[i])} · ${series.name}: ${numberLabel(value,series.digits ?? 1,isPercent ? '%' : '')}`));
      group.append(svg('line',{x1:px,y1:py,x2:lx,y2:ly,stroke:series.color,'stroke-width':.7}));
      group.append(svg('rect',{x:lx-half,y:ly-12,width:half*2,height:18,rx:3,fill:'#fff',stroke:series.color,'stroke-width':.8}));
      group.append(svg('text',{x:lx,y:ly+1,'text-anchor':'middle','font-size':11,fill:'#172c3c'},text)); graph.append(group);
      if(model.annual && series.yoy && valid(series.yoy[i]))yoyMarks.push({px,py,value:series.yoy[i]});
    });
    // Reserva primero las etiquetas de volumen y cobertura, después ubica el YoY.
    yoyMarks.forEach(({px,py,value})=>{
      const text=signedLabel(value),half=Math.max(17,text.length*3.4),lx=Math.max(left+half,Math.min(right-half,px));
      let ly=Math.max(14,py-30);
      while(ly>14 && boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18))ly=Math.max(14,ly-19);
      if(boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18)) {
        ly=Math.min(bottom-5,py+36);
        while(ly<bottom-5 && boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18))ly=Math.min(bottom-5,ly+19);
      }
      boxes.push({x:lx,y:ly,half});
      graph.append(svg('text',{x:lx,y:ly,'text-anchor':'middle','font-size':11,'font-weight':600,fill:value>=0 ? '#1e8449' : '#8b0000',class:'dashboard-yoy-label'},text));
    });
    if(!model.series.some(s=>s.values.slice(first).some(valid))) graph.append(svg('text',{x:width/2,y:130,'text-anchor':'middle',fill:'#637888','font-size':13},lang?.noValues||'Sin base suficiente en el período visible'));
    const guide=svg('line',{x1:left,x2:left,y1:top,y2:bottom,stroke:'#8298a5','stroke-dasharray':'3 4',visibility:'hidden'}); graph.append(guide);
    let active=n-1;
    function inspect(i) {
      active=i; guide.setAttribute('visibility',options.detail ? 'visible' : 'hidden'); guide.setAttribute('x1',x(i));guide.setAttribute('x2',x(i));
      tooltip.textContent=`${dateLabel(model.dates[i])} · `+model.series.map(s=>`${s.name}: ${numberLabel(s.values[i],s.digits ?? 1,s.percent || model.percent ? '%' : '')}${model.annual && s.yoy ? ` · YoY ${signedLabel(s.yoy[i])}` : ''}`).join(' | ');
    }
    graph.onpointermove=event=>{const box=graph.getBoundingClientRect(), px=(event.clientX-box.left)/box.width*width;inspect(Math.max(first,Math.min(n-1,Math.floor(first+(px-left)/(right-left)*(n-first)))));};
    graph.onkeydown=event=>{if(['ArrowLeft','ArrowRight'].includes(event.key)){event.preventDefault();inspect(Math.max(first,Math.min(n-1,active+(event.key==='ArrowRight'?1:-1))));}};
    inspect(n-1);
  }

  function pptAvailable(sheet, pipeline, key, colors) {
    if (!sheet?.dashboard || !sheet.pipelines?.[pipeline]?.available) return false;
    if(key==='trend_correlations')return !!sheet.template?.correlations?.some(row=>row.pipelines.some(valid));
    const model=models(sheet,pipeline,colors).find(model=>model.key===key);
    return !!model && (model.kpi ? model.kpi.total_count>0 : !!model.annual && chartData(model).dates.length>0);
  }

  function correlationView(sheet) {
    const rows=(sheet.template?.correlations || []).filter(row=>row.pipelines.some(valid));
    const pipelines=Array.from({length:7},(_,p)=>p).filter(p=>rows.some(row=>valid(row.pipelines[p])));
    return {rows,pipelines,samples:rows.flatMap(row=>pipelines.map(p=>row.pipelines[p]))};
  }
  const correlationPercent=value=>valid(value) ? fmt(value*100,1,'%') : '—';

  function renderTrendCorrelations(root, sheet, pipeline, enabled,language='ES') {
    if(!root)return;
    root.replaceChildren();root.hidden=!enabled;if(!enabled)return;
    const {rows,pipelines,samples}=correlationView(sheet);
    const lang=window.CoverageI18n?.get(language);
    if(!rows.length){root.hidden=true;return;}
    root.append(node('strong',`${language==='PT'?'Correlações MAT (%)':language==='EN'?'MAT correlations (%)':'Correlaciones MAT (%)'} · ${lang?.chosen||'Elegido'} P${pipeline}`));
    const wrapper=node('div',null,'table-scroll'),table=node('table'),body=node('tbody'),labels=node('tr');
    labels.append(node('th','Ventana MAT'));
    for(const p of pipelines) {
      const th=node('th',`P${p}`);if(p===pipeline)th.className='trend-correlations-selected';labels.append(th);
    }
    body.append(labels);
    for(const row of rows) {
      const tr=node('tr');tr.append(node('th',lang?.correlationLabel(row.label)||row.label));
      pipelines.forEach(p=>{
        const value=row.pipelines[p];
        const td=node('td',lang?lang.number(value==null?null:value*100,1,'%'):correlationPercent(value)),style=correlationStyle(value,samples,['#F8696B','#FFEB84','#63BE7B']);
        if(style){td.style.backgroundColor=style.fill;td.style.color=style.ink;}
        if(p===pipeline)td.className='trend-correlations-selected';tr.append(td);
      });body.append(tr);
    }
    table.append(body);wrapper.append(table);root.append(wrapper);
    root.append(node('small',language==='PT'?'Coeficiente r em %. Apenas janelas e pipelines com dados. — indica uma correlação não calculável.':language==='EN'?'Correlation r as %. Only windows and pipelines with data are shown. — indicates an unavailable correlation.':'Coeficiente r en %. Solo se muestran ventanas y pipelines con datos. — indica una correlación no calculable.'));
  }

  function render(root, nav, sheet, pipeline, colors, options, open, ppt = {}) {
    root.replaceChildren();nav.replaceChildren();
    const figures=models(sheet,pipeline,colors).map(model=>window.CoverageI18n?.get(options.language).model(model)||model);
    figures.forEach((model,index)=>{
      const card=node('article',null,'dashboard-chart-card');card.id=`excel-chart-${model.key}`;card.dataset.chart=model.key;
      const heading=node('div',null,'dashboard-chart-heading'), title=node('h3',model.title);
      let pptActions;
      const extraKey=model.key==='monthly' ? 'trend_correlations' : model.key;
      const extraTitle=model.key==='monthly' ? 'Correlaciones MAT debajo de tendencia' : model.title;
      if(model.kpi || model.annual) {
        const available=pptAvailable(sheet,pipeline,extraKey,colors);
        const label=node('label',null,'ppt-chart-choice'), input=node('input');
        input.type='checkbox';input.className='ppt-extra-input';input.checked=!!ppt.selected?.includes(extraKey);input.disabled=!!ppt.disabled || !available;
        input.dataset.extra=extraKey;
        if(!available){input.dataset.unavailable='true';label.title='Este gráfico no tiene datos para el PPT';}
        input.setAttribute('aria-label',`Incluir ${extraTitle} en PPT`);
        if(available)label.title=model.annual ? 'Añadir la comparación anual con el pipeline elegido a la presentación de esta hoja' : 'Añadir la serie mensual completa de este KPI a la presentación de esta hoja';
        label.append(input,node('span','Incluir en PPT'));
        label.addEventListener('click',event=>event.stopPropagation());
        input.addEventListener('change',()=>ppt.change?.(extraKey,input.checked));
        pptActions=node('div',null,'ppt-chart-actions');pptActions.append(label);
        if(ppt.all) {
          const all=node('button','Incluir en todas','ppt-extra-all text-button');
          all.type='button';all.dataset.extra=extraKey;all.disabled=!!ppt.disabled;
          all.setAttribute('aria-label',`Incluir ${extraTitle} en todas las hojas incluidas`);
          all.title='Aplicar este gráfico a todas las hojas incluidas que tengan datos, usando el pipeline elegido en cada una';
          all.addEventListener('click',event=>{event.stopPropagation();ppt.all(extraKey);});
          pptActions.append(all);
        }
        if(ppt.none){const none=node('button','Quitar de todas','ppt-extra-none text-button');none.type='button';none.dataset.extra=extraKey;none.addEventListener('click',event=>{event.stopPropagation();ppt.none(extraKey);});pptActions.append(none);}
        pptActions.append(node('small',null,'ppt-extra-count'));
      }
      const expand=()=>open(model,options);
      const button=node('button','Ampliar ↗','text-button');button.setAttribute('aria-label',`Ampliar ${model.title}`);
      button.addEventListener('click',event=>{event.stopPropagation();expand();});
      card.tabIndex=0;card.setAttribute('aria-label',`${model.title}. Clic en el cuadro para ampliar.`);card.setAttribute('aria-haspopup','dialog');
      let pointerStart;
      card.addEventListener('pointerdown',event=>{pointerStart={x:event.clientX,y:event.clientY};});
      card.addEventListener('click',event=>{
        if(event.target.closest('button,input,label') || (pointerStart && Math.hypot(event.clientX-pointerStart.x,event.clientY-pointerStart.y)>8))return;
        expand();
      });
      card.addEventListener('keydown',event=>{
        if(event.target!==card)return;
        if(['Enter',' '].includes(event.key)){event.preventDefault();expand();}
      });
      heading.append(title,button);card.append(heading,node('p',model.description,'section-description'));
      if(pptActions)card.append(pptActions);
      const content=node('div',null,'dashboard-chart-content');card.append(content);root.append(card);plot(content,model,options);
      if(model.key==='monthly') {
        const correlations=node('section',null,'trend-correlations-preview');card.append(correlations);
        renderTrendCorrelations(correlations,sheet,pipeline,!!ppt.selected?.includes(extraKey),options.language);
      }
      const link=node('a',`${String(index+1).padStart(2,'0')} ${model.title}`);link.href=`#${card.id}`;nav.append(link);
    });
    return figures.length;
  }
  async function capturePpt(sheet, pipeline, colors, options, keys) {
    const lang=window.CoverageI18n.get(options.language||'ES'),fmt=lang.number,month=lang.month,signed=lang.signed;
    const images={};
    if(keys.includes('trend_correlations') && correlationView(sheet).rows.length) {
      const {rows,pipelines,samples}=correlationView(sheet),height=(rows.length+1)*60;
      const root=svg('svg',{xmlns:ns,width:1200,height,viewBox:`0 0 1200 ${height}`,'font-family':'Segoe UI, Arial, sans-serif'});
      for(let row=0;row<=rows.length;row++)for(let column=0;column<=pipelines.length;column++) {
        const x=column ? 264+(column-1)*(936/pipelines.length) : 0,w=column ? 936/pipelines.length : 264;
        const p=pipelines[column-1],value=row && column ? rows[row-1].pipelines[p] : null;
        const style=correlationStyle(value,samples,['#F8696B','#FFEB84','#63BE7B']);
        const chosen=column>0 && p===pipeline;
        const fill=!row && chosen ? '#008C95' : !row || !column ? '#ECF5F3' : style?.fill || '#F7FAFC';
        const text=!row ? (column ? `P${p}` : `r MAT (%) · P${pipeline}`) : !column ? lang.correlationLabel(rows[row-1].label) : fmt(value==null?null:value*100,1,'%');
        root.append(svg('rect',{x,y:row*60,width:w,height:60,fill,stroke:'#FFFFFF','stroke-width':1}));
        root.append(svg('text',{x:column ? x+w/2 : x+12,y:row*60+37,'text-anchor':column ? 'middle' : 'start',
          'font-size':column ? 24 : 20,'font-weight':!row || !column || chosen ? 650 : 400,
          fill:!row && chosen ? '#FFFFFF' : '#172C3C'},text));
      }
      const image=new Image();image.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(new XMLSerializer().serializeToString(root));
      await image.decode();
      const canvas=document.createElement('canvas');canvas.width=2400;canvas.height=height*2;
      canvas.getContext('2d').drawImage(image,0,0,2400,height*2);
      images.trend_correlations=canvas.toDataURL('image/png');
    }
    for(const source of models(sheet,pipeline,colors).filter(model=>keys.includes(model.key))) {
      const model=lang.model(source);
      const content=node('div');plot(content,model,options,true);
      const graph=content.querySelector('svg');
      const root=svg('svg',{xmlns:ns,width:1200,height:675,viewBox:'0 0 1200 675','font-family':'Segoe UI, Arial, sans-serif'});
      root.append(svg('rect',{width:1200,height:675,fill:'#fff'}));
      root.append(svg('text',{x:42,y:40,'font-size':23,'font-weight':650,fill:'#172c3c'},model.title));
      root.append(svg('text',{x:42,y:66,'font-size':13,fill:'#637888'},`${sheet.brand} · P${pipeline} · ${model.description}`));
      if(model.kpi) {
        for(const [i,label,value] of [[0,lang.average12,model.kpi.average],[1,lang.averageTotal,model.kpi.total_average]]) {
          root.append(svg('text',{x:42+i*310,y:99,'font-size':13,fill:'#637888'},label));
          root.append(svg('text',{x:42+i*310,y:127,'font-size':23,'font-weight':650,fill:'#172c3c'},fmt(value,kpiDigits(model.kpi,value),model.key==='penetration'?'%':'')));
        }
      }
      root.append(svg('text',{x:42,y:153,'font-size':12,fill:'#637888'},content.querySelector('.chart-period').textContent));
      // Mismo SVG que la vista ampliada, sin cursor ni controles interactivos.
      graph.querySelectorAll('line[stroke-dasharray]').forEach(line=>line.remove());
      graph.setAttribute('x','30');graph.setAttribute('y','164');graph.setAttribute('width','1140');graph.setAttribute('height','407');
      graph.removeAttribute('style');graph.removeAttribute('tabindex');root.append(graph);
      const legendWidth=model.series.reduce((sum,series)=>sum+series.name.length*7+42,0);
      let legendX=Math.max(42,(1200-legendWidth)/2);
      for(const series of model.series) {
        root.append(svg('line',{x1:legendX,x2:legendX+18,y1:598,y2:598,stroke:series.color,'stroke-width':3}));
        root.append(svg('text',{x:legendX+24,y:602,'font-size':13,fill:'#637888'},series.name));legendX+=series.name.length*7+42;
      }
      const visible=chartData(model,options.window),last=visible.dates.length-1;
      if(last>=0)root.append(svg('text',{x:42,y:642,'font-size':12,fill:'#637888'},`${month(visible.dates[last])} · `+
        visible.series.map(s=>`${s.name}: ${fmt(s.values[last],s.digits??1,s.percent||model.percent?'%':'')}${model.annual&&s.yoy?` · YoY ${signed(s.yoy[last])}`:''}`).join(' | ')));
      const img=new Image();img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(new XMLSerializer().serializeToString(root));
      await img.decode();
      const canvas=document.createElement('canvas');canvas.width=2400;canvas.height=1350;
      canvas.getContext('2d').drawImage(img,0,0,canvas.width,canvas.height);
      images[model.key]=canvas.toDataURL('image/png');
    }
    return images;
  }
  return {render,plot,capturePpt,axisScale,axisTick,correlationStyle,chartData,kpiAverages,kpiDigits,pptAvailable,renderTrendCorrelations,correlationView,correlationPercent,
    annualModel:(sheet,pipeline,colors)=>models(sheet,pipeline,colors).find(model=>model.key==='annual_mat')};
})();
