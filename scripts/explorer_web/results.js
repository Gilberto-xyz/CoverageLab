(() => {
  'use strict';
  const token = document.querySelector('meta[name="explorer-token"]').content;
  const node = (tag, text, className) => {
    const item = document.createElement(tag);
    if (text) item.textContent = text;
    if (className) item.className = className;
    return item;
  };
  const button = (text, action, className = 'button secondary') => {
    const item = node('button', text, className); item.type = 'button'; item.addEventListener('click', action); return item;
  };
  async function api(path, body) {
    const response = await fetch(path, {method: body ? 'POST' : 'GET',
      headers: {'X-Explorer-Token': token, ...(body ? {'Content-Type':'application/json'} : {})},
      ...(body ? {body:JSON.stringify(body)} : {})});
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'No se pudieron consultar los resultados.');
    return data;
  }
  let entries = [], timer, pending = false;
  const launch = button('Resultados', () => {dialog.showModal(); refresh();}, 'button secondary results-launch');
  launch.title = 'Ver y abrir todas las entregas de Coberturas y 5W1H';
  document.getElementById('header-actions').prepend(launch);
  const dialog = node('dialog', '', 'results-dialog');dialog.setAttribute('aria-labelledby', 'results-title');
  const heading = node('div', '', 'results-heading'), identity = node('div');
  const title = node('h2', 'Resultados');title.id = 'results-title';
  identity.append(title, node('p', 'Coberturas y 5W1H · Abre los archivos directamente en Excel o PowerPoint.'));
  const close = button('Cerrar ×', () => dialog.close());heading.append(identity, close);
  const summary = node('div', '', 'results-summary');
  const toolbar = node('div', '', 'results-toolbar');
  const filter = node('select');filter.setAttribute('aria-label', 'Filtrar resultados por módulo');
  [['all','Todos los módulos'],['coverage','Coberturas'],['research','Investigación 5W1H']].forEach(([value,label]) => {
    const option = node('option',label);option.value=value;filter.append(option);
  });
  const search = node('input');search.type='search';search.placeholder='Buscar país, fabricante o archivo…';search.setAttribute('aria-label','Buscar resultados');
  const openRoot = button('Abrir carpeta de resultados', async () => {
    openRoot.disabled=true;
    try {const result=await api('/api/open-results',{scope:'all'});message(result.message);}
    catch(error){message(error.message,true);}finally{openRoot.disabled=false;}
  });
  toolbar.append(filter,search,button('Actualizar',refresh),openRoot);
  const feedback=node('p','','results-feedback');feedback.setAttribute('role','status');
  const list=node('div','','results-list'), path=node('p','','results-path');
  dialog.append(heading,summary,toolbar,feedback,list,path);document.body.append(dialog);
  function message(text,error=false){feedback.textContent=text;feedback.classList.toggle('error',error);}
  function date(value){return value ? new Date(value).toLocaleString('es-MX',{dateStyle:'medium',timeStyle:'short'}) : '';}
  function render(){
    summary.replaceChildren();list.replaceChildren();
    for(const [kind,label] of [['coverage','Coberturas'],['research','5W1H']]){
      const matching=entries.filter(entry=>entry.module===kind), latest=matching.find(entry=>['complete','available'].includes(entry.status));
      const card=button('',()=>{filter.value=kind;search.value='';render();},'results-summary-card');
      card.append(node('strong',label),node('span',latest ? `Última entrega: ${date(latest.date)}` : 'Sin entregas disponibles'),
        node('small',`${matching.filter(entry=>['complete','available'].includes(entry.status)).length} entrega(s) disponibles`));summary.append(card);
    }
    const query=search.value.trim().toLocaleLowerCase();
    const visible=entries.filter(entry=>(filter.value==='all'||entry.module===filter.value)&&
      `${entry.title} ${entry.source} ${entry.outputs.map(out=>out.name).join(' ')}`.toLocaleLowerCase().includes(query));
    if(!visible.length)list.append(node('p','No hay resultados para esta búsqueda. Los entregables aparecerán aquí al finalizar la generación.','research-empty'));
    const seen=new Set();
    for(const entry of visible){
      const card=node('article','','results-card'), top=node('div','','results-card-heading'), text=node('div');
      const label=entry.module==='research'?'Investigación 5W1H':'Coberturas';
      const latest=!seen.has(entry.module);seen.add(entry.module);
      text.append(node('div',`${label}${latest?' · Más reciente':''} · ${date(entry.date)}`,'eyebrow'),node('h3',entry.title));
      if(entry.source)text.append(node('p',`Entrada: ${entry.source}`));
      const status={complete:'Generado correctamente',available:'Archivos disponibles',running:'Generando…',queued:'En espera',failed:'No se completó'}[entry.status]||entry.status;
      const badge=node('span',status,`pill results-state ${entry.status}`);top.append(text,badge);card.append(top);
      if(entry.error)card.append(node('p',entry.error,'error'));
      if(entry.status==='running'&&entry.progress?.percent!=null)card.append(node('p',`${Math.round(entry.progress.percent)}% · ${entry.progress.stage||'En proceso'}`));
      const actions=node('div','','output-buttons');
      for(const output of [...entry.outputs,{name:'folder',label:'Abrir carpeta'}]){
        if(!['complete','available'].includes(entry.status))continue;
        const action=button(output.label,async()=>{
          if(pending)return;pending=true;action.disabled=true;
          message(`Abriendo ${output.name==='folder'?'carpeta':output.name}…`);
          try{const result=await api('/api/results/open',{id:entry.id,output:output.name});message(result.message);}
          catch(error){message(error.message,true);}finally{pending=false;action.disabled=false;}
        },'button result-file');
        action.title=output.name==='folder'?entry.directory:output.name;
        window.CoverageApp?.resultButton?.(action,output.label,output.name==='folder'?'folder':output.name.endsWith('.pptx')?'presentation':'template');
        actions.append(action);
      }
      card.append(actions);
      if(entry.directory){const details=node('details'), route=node('summary','Ver ubicación y nombres');details.append(route,node('p',entry.directory,'results-path'));
        entry.outputs.forEach(output=>details.append(node('div',output.name,'results-filename')));card.append(details);}
      list.append(card);
    }
  }
  async function refresh(announce=true){
    clearTimeout(timer);
    if(pending){if(dialog.open)timer=setTimeout(()=>refresh(false),5000);return;}
    if(announce)message('Consultando entregas…');
    try{const result=await api('/api/results');const changed=JSON.stringify(entries)!==JSON.stringify(result.deliveries);entries=result.deliveries;path.textContent=`Todo se guarda en: ${result.directory}`;if(changed||!summary.childElementCount)render();if(announce)message(`${entries.length} entrega(s) encontradas.`);}
    catch(error){message(error.message,true);}
    if(dialog.open)timer=setTimeout(()=>refresh(false),5000);
  }
  filter.addEventListener('change',render);search.addEventListener('input',render);
  dialog.addEventListener('close',()=>clearTimeout(timer));
  dialog.addEventListener('click',event=>{if(event.target===dialog){const box=dialog.getBoundingClientRect();if(event.clientX<box.left||event.clientX>box.right||event.clientY<box.top||event.clientY>box.bottom)dialog.close();}});
})();
