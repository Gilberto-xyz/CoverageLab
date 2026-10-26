'use strict';
window.GenerationProgress=(()=>{
  function duration(seconds) {
    seconds=Math.max(0,Math.ceil(Number(seconds)||0));
    if(seconds<60)return `${seconds} s`;
    const minutes=Math.floor(seconds/60),rest=seconds%60;
    if(minutes<60)return `${minutes} min${rest ? ` ${rest} s` : ''}`;
    return `${Math.floor(minutes/60)} h ${minutes%60} min`;
  }
  function view(job,preparing,now=Date.now()) {
    if(preparing) {
      const completed=preparing.completed || 0,total=preparing.total || 0;
      const remaining=completed>0 ? (now-preparing.startedAt)/1000/completed*(total-completed) : null;
      return {mode:'running',percent:total ? 5*completed/total : 0,
        primary:total ? `Gráficas ${completed}/${total}` : 'Preparando exportación…',
        secondary:remaining>0 ? `≈ ${duration(remaining)} de preparación` : 'Preparando archivos…',
        detail:'Preparando las imágenes que se incluirán en la PPT'};
    }
    if(!job)return null;
    const progress=job.progress || {},percent=Math.max(0,Math.min(99.5,Number(progress.percent)||5));
    if(job.status==='failed')return {mode:'failed',percent,primary:'Reintentar generación',secondary:'La exportación falló',detail:job.error || 'Consulta el registro'};
    if(job.status==='complete') {
      if(!job.completedObservedAt || now-job.completedObservedAt>=4000)return null;
      return {mode:'complete',percent:100,primary:'Completado · 100%',secondary:`Listo en ${duration(progress.elapsed_seconds)}`,detail:'Todos los entregables están disponibles'};
    }
    if(job.status!=='running')return null;
    const age=Math.max(0,(now-(job.progressReceivedAt || now))/1000);
    const remaining=Number.isFinite(progress.remaining_seconds) ? progress.remaining_seconds-age : null;
    const labels={starting:'Iniciando',excel:'Excel',presentation:'PPT',research:'5W1H',saving_ppt:'Guardando PPT',bank:'Banco',report:'Reporte',comparison:'Comparativo',packaging:'Verificando'};
    const count=['excel','presentation','research'].includes(progress.stage) && progress.total>1 ? ` ${Math.floor(progress.completed || 0)}/${Math.floor(progress.total)}` : '';
    return {mode:'running',percent,primary:`${labels[progress.stage] || 'Generando'}${count} · ${Math.floor(percent)}%`,
      secondary:remaining>0 ? `≈ ${duration(remaining)} restantes` : percent>=98 ? 'Finalizando entregables…' : Number.isFinite(progress.remaining_seconds) ? 'Recalculando estimación…' : 'Estimando tiempo…',
      detail:`${progress.label || 'Generando archivos'}${count} · Transcurrido: ${duration((progress.elapsed_seconds || 0)+age)}${progress.detail ? ` · ${progress.detail}` : ''}`};
  }
  return {view,duration};
})();
