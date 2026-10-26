"""Catálogo común de entregables actuales y anteriores, sin archivos auxiliares."""
from datetime import datetime
from pathlib import Path
import os
import subprocess
import sys

from studio_desktop import open_folder


LABELS = {'template': 'Excel de cobertura', 'presentation': 'Presentación PPT',
          'bank': 'Banco de coberturas', 'pipeline_report': 'Reporte de pipelines',
          'mosaic_comparison': 'Comparativo Mosaic', 'Excel': 'Excel 5W1H', 'PPT': 'Presentación 5W1H'}


def file_label(file):
    if file.suffix.lower() == '.pptx':
        return 'Presentación PPT'
    for prefix, label in (('template_', 'Excel de cobertura'), ('banco_', 'Banco de coberturas'),
                          ('reporte', 'Reporte de pipelines'), ('scorecard', 'Scorecard')):
        if file.name.lower().startswith(prefix):
            return label
    return 'Excel'


def catalog(explorer):
    root = explorer.output_root.resolve()
    with explorer.lock:
        jobs = [(dict(job), 'coverage') for job in explorer.jobs.values()]
    with explorer.workflow.lock:
        jobs += [(dict(job), 'research') for job in explorer.workflow.research_jobs.values()]
    registered = {str(Path(job['directory']).resolve()): (job, kind)
                  for job, kind in jobs if job.get('directory')}
    entries = []
    # Las primeras versiones colocaban la entrega dentro de una carpeta de ejecución.
    directories = []
    for folder in root.iterdir() if root.is_dir() else ():
        if folder.is_dir() and not folder.name.startswith(('.', '_')) and folder.resolve().parent == root:
            directories.append(folder)
            directories.extend(child for child in folder.iterdir() if child.is_dir() and
                               not child.name.startswith(('.', '_')) and child.resolve().parent == folder.resolve())
    for directory in directories:
        if not directory.resolve().is_relative_to(root):
            continue
        fallback = registered.get(str(directory.parent.resolve()), ({}, 'research' if '5w1h' in directory.name.lower() else 'coverage'))
        job, kind = registered.get(str(directory.resolve()), fallback)
        labels = {out.get('name'): LABELS.get(out.get('label')) for out in job.get('outputs', [])}
        outputs = []
        for file in sorted(directory.iterdir()):
            if file.name.startswith('~$') or file.suffix.lower() not in {'.xlsx', '.pptx'}:
                continue
            if file.is_file() and file.resolve().parent == directory.resolve():
                outputs.append({'name': file.name, 'label': labels.get(file.name) or file_label(file)})
        if not outputs:
            continue
        entries.append({'id': directory.relative_to(root).as_posix(), 'module': kind, 'title': job.get('title') or directory.name,
                        'source': job.get('file', ''), 'status': job.get('status', 'available'),
                        'date': job.get('finished_at') or job.get('started_at') or
                                datetime.fromtimestamp(directory.stat().st_mtime).astimezone().isoformat(),
                        'directory': str(directory.resolve()), 'outputs': outputs})
    # Un trabajo en curso o fallido también debe verse, aunque todavía no tenga archivos.
    for job, kind in jobs:
        if job.get('status') not in {'running', 'failed', 'queued'}:
            continue
        entries.append({'id': 'job:' + job['id'], 'module': kind, 'title': job.get('title') or job.get('file', ''),
                        'source': job.get('file', ''), 'status': job['status'], 'date': job.get('started_at', ''),
                        'error': job.get('error'), 'progress': job.get('progress', {}), 'outputs': [], 'directory': ''})
    entries.sort(key=lambda entry: datetime.fromisoformat(entry['date']).timestamp() if entry['date'] else 0, reverse=True)
    return {'directory': str(root), 'deliveries': entries}


def open_item(explorer, payload):
    entry = next((entry for entry in catalog(explorer)['deliveries'] if entry['id'] == payload.get('id')), None)
    if not entry or entry['status'] not in {'complete', 'available'}:
        raise ValueError('Selecciona una entrega con archivos disponibles.')
    directory = Path(entry['directory']).resolve()
    if payload.get('output') == 'folder':
        return open_folder(directory)
    name = payload.get('output')
    if not any(output['name'] == name for output in entry['outputs']):
        raise ValueError('El archivo no pertenece a esta entrega.')
    target = (directory / name).resolve()
    if target.parent != directory or not target.is_file():
        raise ValueError('El archivo ya no está disponible.')
    if sys.platform == 'win32':
        os.startfile(str(target))
    else:
        subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', str(target)],
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return {'message': f'Abierto: {target.name}'}
