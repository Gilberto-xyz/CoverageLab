'use strict';
// Comparativo especial. Comparte cálculos, marcas, etiquetas y colores con el panel.
window.MosaicComparison=(()=>{
  const ns='http://www.w3.org/2000/svg',valid=v=>v!=null&&Number.isFinite(v);
  const svg=(tag,attrs,text)=>{const el=document.createElementNS(ns,tag);for(const [k,v] of Object.entries(attrs))el.setAttribute(k,v);if(text!=null)el.textContent=text;return el;};
  const month=date=>new Intl.DateTimeFormat('es-MX',{month:'short',year:'2-digit'}).format(new Date(date+'T12:00:00'));
  const variationPercent=value=>valid(value)?new Intl.NumberFormat('es-MX',{minimumFractionDigits:1,maximumFractionDigits:1,signDisplay:'exceptZero'}).format(value)+'%':'—';
  function variationStyle(value,colors={}) {
    if(!valid(value)||value===0)return {fill:'#F7FAFC',ink:'#172C3C'};
    return value>0 ? colors.positive||{fill:'#C6EFCE',ink:'#006100'} : colors.negative||{fill:'#FFC7CE',ink:'#9C0006'};
  }
  function prepare(analysis,selections,pair) {
    const sheets=['original','mosaic'].map(key=>analysis.sheets.find(s=>s.name===pair[key]));
    if(sheets.some(sheet=>!sheet||!Number.isInteger(selections[sheet.name])))return null;
    const models=sheets.map(sheet=>CoverageDashboard.chartData(CoverageDashboard.annualModel(sheet,selections[sheet.name],analysis.chart_colors),'all'));
    const dates=models[0].dates.filter(date=>models[1].dates.includes(date));if(!dates.length)return null;
    models.forEach(model=>{
      const indexes=dates.map(date=>model.dates.indexOf(date));model.dates=dates;
      model.series=model.series.map(s=>({...s,values:indexes.map(i=>s.values[i]),...(s.yoy?{yoy:indexes.map(i=>s.yoy[i])}:{})}));
    });
    const scales=Object.fromEntries(['left','right'].map(axis=>[axis,CoverageDashboard.axisScale(models.flatMap(m=>m.series.filter(s=>(s.axis||'left')===axis).flatMap(s=>s.values)))]));
    models.forEach(model=>{model.scales=scales;});
    const allRows=sheets.flatMap(s=>s.template?.correlations||[]);
    const rows=[...new Set(allRows.filter(r=>r.pipelines.some(valid)).map(r=>r.label))];
    const pipelines=Array.from({length:7},(_,i)=>i).filter(p=>allRows.some(r=>rows.includes(r.label)&&valid(r.pipelines[p])));
    const samples=sheets.map(sheet=>(sheet.template?.correlations||[]).flatMap(row=>row.pipelines).filter(valid));
    const variations=sheets.map(sheet=>{
      const annual=(sheet.template?.variations||[]).find(row=>row.label==='MAT vs. MAT anterior');
      const cutoff=sheet.dashboard.mat_periods.find(row=>row.date===dates.at(-1));
      return annual && sheet.dates.at(-1)?.slice(0,10)===dates.at(-1) ? annual :
        {worldpanel:cutoff?.worldpanel_yoy??null,pipelines:cutoff?.sell_in_yoy||Array(7).fill(null)};
    });
    // Incluye los pipelines con variación aunque su correlación no sea calculable.
    for(let p=0;p<7;p++)if(!pipelines.includes(p)&&variations.some(row=>valid(row.pipelines[p])))pipelines.push(p);
    pipelines.sort((a,b)=>a-b);
    return {sheets,models,dates,rows,pipelines,samples,variations};
  }
  function eligible(analysis,selections) {
    return (analysis?.mosaic_pairs||[]).filter(pair=>prepare(analysis,selections,pair));
  }
  function render(analysis,selections,pair,options) {
    const data=prepare(analysis,selections,pair);if(!data)throw new Error('La pareja requiere ambas hojas y cortes MAT compartidos.');
    const lang=window.CoverageI18n.get(options.language||analysis.output_language||'ES'),month=lang.month,variationPercent=lang.signed;
    data.models=data.models.map(lang.model);
    const root=svg('svg',{xmlns:ns,width:1600,height:900,viewBox:'0 0 1600 900',role:'img','aria-label':`Comparativo Original/Mosaic: ${pair.label}`,'font-family':'Segoe UI, Arial, sans-serif'});
    root.append(svg('rect',{width:1600,height:900,fill:'#FFFFFF'}));
    root.append(svg('text',{x:44,y:43,'font-size':30,'font-weight':700,fill:'#172C3C'},`${pair.label.replace(/^[A-Z]{3,5}_/,'')} · Original / Mosaic`));
    root.append(svg('text',{x:44,y:75,'font-size':17,fill:'#637888'},`${lang.category(pair.category)} · ${lang.title('annual_mat')} · ${lang.unit(pair.unit)}`));
    root.append(svg('line',{x1:800,x2:800,y1:103,y2:854,stroke:'#DCE7EB'}));
    data.sheets.forEach((sheet,index)=>{
      const x=32+index*800,pipeline=selections[sheet.name],model=data.models[index];
      root.append(svg('text',{x:x+12,y:122,'font-size':24,'font-weight':650,fill:index?'#008C95':'#172C3C'},`${index?lang.mosaic:lang.original} · P${pipeline}`));
      root.append(svg('text',{x:x+12,y:146,'font-size':15,fill:'#637888'},`${month(data.dates[0])} – ${month(data.dates.at(-1))} · ${lang.axes}`));
      const container=document.createElement('div');CoverageDashboard.plot(container,model,{...options,language:lang.code,window:'all',scope:'all'},false);
      const graph=container.querySelector('svg');graph.querySelectorAll('line[stroke-dasharray]').forEach(line=>line.remove());
      graph.setAttribute('x',x);graph.setAttribute('y',156);graph.setAttribute('width',760);graph.setAttribute('height',423);root.append(graph);
      model.series.forEach((series,i)=>{
        const lx=x+70+i*235;
        root.append(svg('line',{x1:lx,x2:lx+26,y1:598,y2:598,stroke:series.color,'stroke-width':5}));
        root.append(svg('text',{x:lx+34,y:604,'font-size':17,fill:'#172C3C'},series.name));
      });
      root.append(svg('text',{x:x+12,y:651,'font-size':19,'font-weight':650,fill:'#172C3C'},lang.matrix));
      const rows=data.rows,ps=data.pipelines,rowHeight=Math.min(43,Math.floor(184/(rows.length+3))),labelWidth=182,cellWidth=(736-labelWidth)/Math.max(1,ps.length);
      const fontSize=rowHeight<38?16:19,labelFont=rowHeight<38?14:17,textY=Math.round(rowHeight*.65);
      for(let row=0;row<=rows.length;row++)for(let col=0;col<=ps.length;col++) {
        const p=ps[col-1],left=x+12+(col?labelWidth+(col-1)*cellWidth:0),w=col?cellWidth:labelWidth;
        const originalRow=sheet.template.correlations.find(r=>r.label===rows[row-1]);
        const value=row&&col ? originalRow?.pipelines[p] : null;
        const style=CoverageDashboard.correlationStyle(value,data.samples[index],['#F8696B','#FFEB84','#63BE7B']);
        const chosen=col>0&&p===pipeline;
        const fill=!row&&chosen?'#008C95':!row||!col?'#ECF5F3':style?.fill||'#F7FAFC';
        const label=lang.correlationLabel(rows[row-1],true);
        const text=!row?(col?`P${p}`:`${lang.chosen} P${pipeline}`):!col?label:lang.number(value==null?null:value*100,1,'%');
        const top=670+row*rowHeight;
        root.append(svg('rect',{x:left,y:top,width:w,height:rowHeight,fill,stroke:'#FFFFFF','stroke-width':1}));
        root.append(svg('text',{x:col?left+w/2:left+10,y:top+textY,'text-anchor':col?'middle':'start','font-size':col?fontSize:labelFont,'font-weight':!row||!col||chosen?650:400,fill:!row&&chosen?'#FFFFFF':'#172C3C'},text));
      }
      const variation=data.variations[index],variationTop=670+(rows.length+1)*rowHeight;
      for(let row=0;row<2;row++) {
        const top=variationTop+row*rowHeight,label=row?lang.varWP:lang.varSell;
        root.append(svg('rect',{x:x+12,y:top,width:labelWidth,height:rowHeight,fill:'#ECF5F3',stroke:'#FFFFFF'}));
        root.append(svg('text',{x:x+22,y:top+textY,'font-size':labelFont,'font-weight':650,fill:'#172C3C'},label));
        const values=row?[variation.worldpanel]:ps.length?ps.map(p=>variation.pipelines[p]):[null];
        values.forEach((value,col)=>{
          const width=row||!ps.length?736-labelWidth:cellWidth,left=x+12+labelWidth+col*cellWidth;
          const style=variationStyle(value,analysis.table_colors),chosen=!row&&ps[col]===pipeline;
          const group=svg('g',{'data-variation':row?'worldpanel':'sell-in',...(row?{}:{'data-pipeline':ps[col]})});
          group.append(svg('title',{},`${label}${row?'':` P${ps[col]}`} · MAT ${month(data.dates.at(-1))} ${lang.previous}: ${variationPercent(value)}`));
          group.append(svg('rect',{x:left,y:top,width,height:rowHeight,fill:style.fill,stroke:'#FFFFFF','stroke-width':1}));
          group.append(svg('text',{x:left+width/2,y:top+textY,'text-anchor':'middle','font-size':fontSize,'font-weight':chosen?700:500,fill:style.ink},variationPercent(value)));
          root.append(group);
        });
      }
    });
    root.append(svg('text',{x:44,y:882,'font-size':15,fill:'#637888'},lang.footer));
    return root;
  }
  async function capture(analysis,selections,pair,options) {
    const root=render(analysis,selections,pair,options),image=new Image();
    image.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(new XMLSerializer().serializeToString(root));await image.decode();
    const canvas=document.createElement('canvas');canvas.width=3200;canvas.height=1800;
    canvas.getContext('2d').drawImage(image,0,0,canvas.width,canvas.height);return canvas.toDataURL('image/png');
  }
  return {prepare,eligible,render,capture,variationPercent,variationStyle};
})();
