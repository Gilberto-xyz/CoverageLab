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
  function kpiAverages(container, kpi) {
    const suffix=kpi.key==='penetration' ? '%' : '';
    const summary=node('dl',null,'kpi-averages');
    for(const [label,value,count,start] of [['Promedio 12 meses',kpi.average,kpi.recent_count,kpi.recent_start],['Promedio total',kpi.total_average,kpi.total_count,kpi.history_start]]) {
      const entry=node('div');entry.append(node('dt',label),node('dd',fmt(value,kpiDigits(kpi,value),suffix)));
      entry.title=`${month(start)} – ${month(kpi.end)} · ${count} mes(es) con datos`;summary.append(entry);
    }
    container.append(summary);
    container.append(node('p',`Promedios mensuales · 12 meses: ${kpi.recent_count} datos · Total: ${kpi.total_count} datos hasta ${month(kpi.end)}`,'kpi-average-note'));
  }

  function axisScale(values, topPadding = .18, bottomPadding = .08) {
    const numbers = values.filter(valid);
    const min = Math.min(0,...numbers), max = Math.max(0,...numbers);
    const span = max-min;
    return span ? {min:min < 0 ? min-span*bottomPadding : 0,max:max+span*topPadding} : {min:0,max:1};
  }
  function axisTick(value, span) {
    if(Math.abs(value)>=10000)return compact(value);
    const step=Math.abs(span)/4;
    const digits=step ? Math.min(8,Math.max(0,1-Math.floor(Math.log10(step)))) : 1;
    return new Intl.NumberFormat('es-MX',{maximumFractionDigits:digits}).format(value);
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
    model=chartData(model,options.window);
    if(model.kpi)kpiAverages(container,model.kpi);
    const period=node('p',model.dates.length ? `${month(model.dates[0])} – ${month(model.dates.at(-1))} · Datos disponibles en todas las series` : 'Sin período compartido con datos','chart-period');container.append(period);
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
    if (!n) { graph.append(svg('text',{x:width/2,y:130,'text-anchor':'middle',fill:'#637888','font-size':14},'Sin base común en el período elegido')); tooltip.textContent='Las bases insuficientes se muestran como —.'; return; }
    const dual = model.series.some(s=>s.axis === 'right');
    const left = 59, right = width - (dual ? 69 : 25), top = 38, bottom = 248;
    const x = i => left + (i-first+.5) / Math.max(1,n-first)*(right-left);
    const scale = axis => {
      const values = model.series.filter(s=>(s.axis || 'left')===axis).flatMap(s=>s.values.slice(first)).filter(valid);
      return axisScale(values);
    };
    const scales = {left:scale('left'),right:scale('right')};
    const y = (v,axis='left') => bottom - (v-scales[axis].min)/(scales[axis].max-scales[axis].min)*(bottom-top);
    for (let i=0;i<=4;i++) {
      const py=bottom-i/4*(bottom-top), value=scales.left.min+i/4*(scales.left.max-scales.left.min);
      if (options.detail) graph.append(svg('line',{x1:left,x2:right,y1:py,y2:py,stroke:'#e5edf1',class:'dashboard-gridline'}));
      graph.append(svg('text',{x:left-8,y:py+4,'text-anchor':'end','font-size':11,fill:'#637888',class:'dashboard-axis-tick'},axisTick(value,scales.left.max-scales.left.min)+(model.percent ? '%' : '')));
      if (dual) graph.append(svg('text',{x:right+7,y:py+4,'font-size':11,fill:'#008c95'},axisTick(scales.right.min+i/4*(scales.right.max-scales.right.min),scales.right.max-scales.right.min)+'%'));
    }
    const step = Math.max(1,Math.ceil((n-first)/7));
    for (let i=first;i<n;i+=step) graph.append(svg('text',{x:x(i),y:bottom+23,'text-anchor':'middle','font-size':11,fill:'#637888'},`${model.annual ? 'MAT ' : ''}${month(model.dates[i])}`));
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
    const boxes=[];
    marks.forEach(({series,i,px,py,value})=>{
      const isPercent=series.percent || model.percent;
      const text=isPercent ? fmt(value,1,'%') : model.annual ? compact(value) : series.digits===2 ? fmt(value,2) : compact(value);
      const half=Math.max(17,text.length*3.4), lx=Math.max(left+half,Math.min(right-half,px));
      let ly=Math.max(16,py-11);
      while(boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18)) ly+=20;
      if(ly>bottom-5) { ly=Math.min(bottom-5,py+18); while(boxes.some(b=>Math.abs(b.x-lx)<b.half+half+3 && Math.abs(b.y-ly)<18)) ly-=20; }
      boxes.push({x:lx,y:ly,half});
      const group=svg('g',{class:'dashboard-data-label','data-series':series.name});
      group.append(svg('title',{},`${month(model.dates[i])} · ${series.name}: ${fmt(value,series.digits ?? 1,isPercent ? '%' : '')}`));
      group.append(svg('line',{x1:px,y1:py,x2:lx,y2:ly,stroke:series.color,'stroke-width':.7}));
      group.append(svg('rect',{x:lx-half,y:ly-12,width:half*2,height:18,rx:3,fill:'#fff',stroke:series.color,'stroke-width':.8}));
      group.append(svg('text',{x:lx,y:ly+1,'text-anchor':'middle','font-size':11,fill:'#172c3c'},text)); graph.append(group);
      if(model.annual && series.yoy && valid(series.yoy[i])) graph.append(svg('text',{x:px,y:Math.max(14,py-30),'text-anchor':'middle','font-size':11,'font-weight':600,fill:series.yoy[i]>=0 ? '#1e8449' : '#8b0000',class:'dashboard-yoy-label'},signed(series.yoy[i])));
    });
    if(!model.series.some(s=>s.values.slice(first).some(valid))) graph.append(svg('text',{x:width/2,y:130,'text-anchor':'middle',fill:'#637888','font-size':13},'Sin base suficiente en el período visible'));
    const guide=svg('line',{x1:left,x2:left,y1:top,y2:bottom,stroke:'#8298a5','stroke-dasharray':'3 4',visibility:'hidden'}); graph.append(guide);
    let active=n-1;
    function inspect(i) {
      active=i; guide.setAttribute('visibility',options.detail ? 'visible' : 'hidden'); guide.setAttribute('x1',x(i));guide.setAttribute('x2',x(i));
      tooltip.textContent=`${month(model.dates[i])} · `+model.series.map(s=>`${s.name}: ${fmt(s.values[i],s.digits ?? 1,s.percent || model.percent ? '%' : '')}${model.annual && s.yoy ? ` · YoY ${signed(s.yoy[i])}` : ''}`).join(' | ');
    }
    graph.onpointermove=event=>{const box=graph.getBoundingClientRect(), px=(event.clientX-box.left)/box.width*width;inspect(Math.max(first,Math.min(n-1,Math.floor(first+(px-left)/(right-left)*(n-first)))));};
    graph.onkeydown=event=>{if(['ArrowLeft','ArrowRight'].includes(event.key)){event.preventDefault();inspect(Math.max(first,Math.min(n-1,active+(event.key==='ArrowRight'?1:-1))));}};
    inspect(n-1);
  }

  function render(root, nav, sheet, pipeline, colors, options, open, ppt = {}) {
    root.replaceChildren();nav.replaceChildren();
    const figures=models(sheet,pipeline,colors);
    figures.forEach((model,index)=>{
      const card=node('article',null,'dashboard-chart-card');card.id=`excel-chart-${model.key}`;card.dataset.chart=model.key;
      const heading=node('div',null,'dashboard-chart-heading'), title=node('h3',model.title);
      if(model.kpi || model.annual) {
        const available=model.kpi ? model.kpi.total_count>0 : chartData(model).dates.length>0;
        const label=node('label',null,'ppt-chart-choice'), input=node('input');
        input.type='checkbox';input.className='ppt-extra-input';input.checked=!!ppt.selected?.includes(model.key);input.disabled=!!ppt.disabled || !available;
        if(!available){input.dataset.unavailable='true';label.title='Este gráfico no tiene datos para el PPT';}
        input.setAttribute('aria-label',`Incluir ${model.title} en PPT`);
        if(available)label.title=model.annual ? 'Añadir la comparación anual con el pipeline elegido a la presentación de esta hoja' : 'Añadir la serie mensual completa de este KPI a la presentación de esta hoja';
        label.append(input,node('span','Incluir en PPT'));
        label.addEventListener('click',event=>event.stopPropagation());
        input.addEventListener('change',()=>ppt.change?.(model.key,input.checked));
        heading.append(label);
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
      const content=node('div',null,'dashboard-chart-content');card.append(content);root.append(card);plot(content,model,options);
      const link=node('a',`${String(index+1).padStart(2,'0')} ${model.title}`);link.href=`#${card.id}`;nav.append(link);
    });
    return figures.length;
  }
  async function capturePpt(sheet, pipeline, colors, options, keys) {
    const images={};
    for(const model of models(sheet,pipeline,colors).filter(model=>keys.includes(model.key))) {
      const content=node('div');plot(content,model,options,true);
      const graph=content.querySelector('svg');
      const root=svg('svg',{xmlns:ns,width:1200,height:675,viewBox:'0 0 1200 675','font-family':'Segoe UI, Arial, sans-serif'});
      root.append(svg('rect',{width:1200,height:675,fill:'#fff'}));
      root.append(svg('text',{x:42,y:40,'font-size':23,'font-weight':650,fill:'#172c3c'},model.title));
      root.append(svg('text',{x:42,y:66,'font-size':13,fill:'#637888'},`${sheet.brand} · P${pipeline} · ${model.description}`));
      if(model.kpi) {
        for(const [i,label,value] of [[0,'Promedio 12 meses',model.kpi.average],[1,'Promedio total',model.kpi.total_average]]) {
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
  return {render,plot,capturePpt,axisScale,axisTick,correlationStyle,chartData,kpiAverages,kpiDigits};
})();
