'use strict';
// El código de idioma lo resuelve el mismo motor que genera Excel/PPT.
window.CoverageI18n=(()=>{
  const titles={
    ES:['Cobertura MAT y penetración','Tendencia mensual de volumen','Tendencia de volumen MAT','Evolución y variación interanual MAT','Compra media','Compra por ocasión','Frecuencia','Penetración','Compradores','Volumen MAT, YoY y cobertura por año'],
    PT:['Cobertura MAT e penetração','Tendência mensal de volume','Tendência de volume MAT','Evolução e variação anual MAT','Compra média','Compra por ocasião','Frequência','Penetração','Compradores','Volume MAT, YoY e cobertura por ano'],
    EN:['MAT coverage and penetration','Monthly volume trend','MAT volume trend','MAT evolution and year-on-year change','Average purchase','Purchase per occasion','Frequency','Penetration','Buyers','MAT volume, YoY and coverage by year']};
  const keys=['coverage','monthly','mat','yoy','purchase_average','purchase_occasion','frequency','penetration','buyers_series','annual_mat'];
  const words={
    ES:{name:'Español',locale:'es-MX',coverage:'Cobertura',penetration:'Penetración mensual',average12:'Promedio 12 meses',averageTotal:'Promedio total',data:'Datos disponibles en todas las series',noPeriod:'Sin período compartido con datos',noBase:'Sin base común en el período elegido',insufficient:'Las bases insuficientes se muestran como —.',noValues:'Sin base suficiente en el período visible',original:'Original',mosaic:'Expansión Mosaic',axes:'Ejes de volumen y cobertura comunes',matrix:'Correlaciones y variaciones MAT (%)',chosen:'Elegido',varSell:'Var. MAT Sell-in',varWP:'Var. MAT Worldpanel',previous:'vs. MAT anterior',footer:'Correlación: escala de color propia. Variación MAT: vs. MAT anterior, verde aumenta y rojo disminuye. Cortes compartidos. — = sin base.',corr:['Año actual · 12 meses','Año anterior · 12 meses','Últimos 24 meses'],shortCorr:['Actual · 12m','Anterior · 12m','Últimos 24m'],kpiDescription:'Serie mensual original · Mismo color del Excel',annualDescription:'Mismo mes de corte en cada año · Barras de volumen, etiquetas YoY y cobertura en eje derecho',coverageDescription:'Cobertura MAT y penetración mensual en el mismo eje (%)',yoyDescription:'Variación de volúmenes MAT vs. los 12 meses anteriores',matDescription:'Sumas móviles de 12 meses',shift:'Sell-in desplazado {p} mes(es)'},
    PT:{name:'Português (Brasil)',locale:'pt-BR',coverage:'Cobertura',penetration:'Penetração mensal',average12:'Média dos últimos 12 meses',averageTotal:'Média total',data:'Dados disponíveis em todas as séries',noPeriod:'Sem período comum com dados',noBase:'Sem base comum no período selecionado',insufficient:'Bases insuficientes são exibidas como —.',noValues:'Sem base suficiente no período visível',original:'Original',mosaic:'Expansão Mosaic',axes:'Eixos de volume e cobertura comuns',matrix:'Correlações e variações MAT (%)',chosen:'Selecionado',varSell:'Var. MAT Sell-in',varWP:'Var. MAT Worldpanel',previous:'vs. MAT anterior',footer:'Correlação: escala de cores própria. Variação MAT: vs. MAT anterior, verde indica aumento e vermelho indica queda. Cortes comuns. — = sem base.',corr:['Ano atual · 12 meses','Ano anterior · 12 meses','Últimos 24 meses'],shortCorr:['Atual · 12m','Anterior · 12m','Últimos 24m'],kpiDescription:'Série mensal original · Mesmas cores do Excel',annualDescription:'Mesmo mês de corte em cada ano · Barras de volume, rótulos YoY e cobertura no eixo direito',coverageDescription:'Cobertura MAT e penetração mensal no mesmo eixo (%)',yoyDescription:'Variação dos volumes MAT vs. os 12 meses anteriores',matDescription:'Somas móveis de 12 meses',shift:'Sell-in deslocado {p} mês(es)'},
    EN:{name:'English',locale:'en-US',coverage:'Coverage',penetration:'Monthly penetration',average12:'Last 12 months average',averageTotal:'Overall average',data:'Data available in all series',noPeriod:'No shared period with data',noBase:'No common base in the selected period',insufficient:'Insufficient bases are shown as —.',noValues:'Insufficient data in the visible period',original:'Original',mosaic:'Mosaic expansion',axes:'Shared volume and coverage axes',matrix:'MAT correlations and changes (%)',chosen:'Selected',varSell:'MAT change Sell-in',varWP:'MAT change Worldpanel',previous:'vs. previous MAT',footer:'Correlation: separate color scales. MAT change: vs. previous MAT, green indicates growth and red indicates decline. Shared cutoffs. — = no base.',corr:['Current year · 12 months','Previous year · 12 months','Last 24 months'],shortCorr:['Current · 12m','Previous · 12m','Last 24m'],kpiDescription:'Original monthly series · Same colors as Excel',annualDescription:'Same cutoff month each year · Volume bars, YoY labels and coverage on the right axis',coverageDescription:'MAT coverage and monthly penetration on the same axis (%)',yoyDescription:'MAT volume change vs. the previous 12 months',matDescription:'Rolling 12-month sums',shift:'Sell-in shifted by {p} month(s)'}};
  const categories={
    'Cross Category (Personal Care)':['Categorias de cuidados pessoais','Personal care categories'],Consolidado:['Consolidado','Consolidated'],
    'Jabón de Tocador':['Sabonetes','Bar soap'],'Cuidado del Cabello':['Cuidados com o cabelo','Hair care'],Desodorantes:['Desodorantes','Deodorants'],
    'Cremas Faciales':['Cremes faciais','Face creams'],'Cremas Corporales':['Cremes corporais','Body creams'],Fragancias:['Fragrâncias','Fragrances'],Maquillaje:['Maquiagem','Makeup'],
    'Bebidas Gaseosas':['Refrigerantes','Carbonated soft drinks'],Pañales:['Fraldas','Diapers'],Cloro:['Água sanitária','Bleach'],Galletas:['Biscoitos','Biscuits']};
  function get(value='ES') {
    const code=Object.hasOwn(words,value)?value:'ES',w=words[code];
    const number=(v,digits=1,suffix='')=>v==null||!Number.isFinite(v)?'—':new Intl.NumberFormat(w.locale,{minimumFractionDigits:digits,maximumFractionDigits:digits}).format(v)+suffix;
    const signed=v=>v==null||!Number.isFinite(v)?'—':new Intl.NumberFormat(w.locale,{minimumFractionDigits:1,maximumFractionDigits:1,signDisplay:'exceptZero'}).format(v)+'%';
    const month=date=>new Intl.DateTimeFormat(w.locale,{month:'short',year:'2-digit'}).format(new Date(date.slice(0,10)+'T12:00:00'));
    const title=(key,fallback)=>titles[code][keys.indexOf(key)]||fallback;
    const unit=value=>code==='PT'?String(value).replace(/\bUNITS\b|\bUnits\b/g,'Unidades'):value;
    function model(source){
      let description=source.description;
      if(source.key==='annual_mat')description=w.annualDescription;
      else if(source.kpi)description=w.kpiDescription;
      else if(source.key==='coverage')description=w.coverageDescription;
      else if(source.key==='yoy')description=w.yoyDescription;
      else if(source.key==='mat')description=source.description.split(' · ')[0]+' · '+w.matDescription;
      else if(source.key==='monthly')description=source.description.replace(/Sell-in desplazado (\d+) mes\(es\)/,(_,p)=>w.shift.replace('{p}',p));
      return {...source,title:title(source.key,source.title),description:unit(description),series:source.series.map(s=>({...s,name:s.name.replace(/^Cobertura/,w.coverage).replace(/^Penetración mensual$/,w.penetration).replace(/^(Compra media|Compra por ocasión|Frecuencia|Penetración|Compradores)$/,()=>title(source.key,s.name))}))};
    }
    const correlationLabel=(label,short=false)=>{const index=words.ES.corr.indexOf(label);return index<0?label:(short?w.shortCorr:w.corr)[index];};
    const category=label=>code==='ES'?label:categories[label]?.[code==='PT'?0:1]||label;
    return {...w,code,number,signed,month,title,model,correlationLabel,category,unit,compact:v=>new Intl.NumberFormat(w.locale,{notation:'compact',maximumFractionDigits:1}).format(v)};
  }
  return {get};
})();
