"""Avance de exportación local: unidades reales y estimación revisable del tiempo."""
from datetime import datetime
from pathlib import Path
import json
import math
import statistics
import time


STAGES = {
    'starting': (5, 5, 'Iniciando motor'),
    'excel': (10, 25, 'Excel'),
    'presentation': (35, 54, 'Diapositivas PPT'),
    'saving_ppt': (89, 5, 'Guardando PPT'),
    'bank': (94, 2.5, 'Banco de coberturas'),
    'report': (96.5, 1, 'Reporte de pipelines'),
    'comparison': (97.5, 1, 'Comparativo Original/Mosaic'),
    'packaging': (98.5, 1, 'Verificando entregables'),
}

RESEARCH_STAGES = {
    'starting': (0, 5, 'Preparando informe 5W1H'),
    'research': (5, 80, 'Hojas 5W1H'),
    'saving_ppt': (85, 5, 'Guardando PPT'),
    'excel': (90, 8.5, 'Excel 5W1H'),
    'packaging': (98.5, 1, 'Verificando entregables'),
}


def number(value, fallback=None):
    try:
        value = float(value)
        return value if math.isfinite(value) else fallback
    except (TypeError, ValueError, OverflowError):
        return fallback


def elapsed_seconds(started_at, now=None):
    try:
        return max(0, (time.time() if now is None else now) - datetime.fromisoformat(started_at).timestamp())
    except (ValueError, TypeError, OverflowError):
        return 0


def historical_estimate(jobs, work_units):
    rates = [number(job.get('elapsed_seconds'), 0) / number(job.get('work_units'), 1)
             for job in jobs if job.get('status') == 'complete'
             and 0 < number(job.get('elapsed_seconds'), 0) < 86400
             and number(job.get('work_units'), 0) > 0]
    return statistics.median(rates[-12:]) * work_units if rates else None


def snapshot(raw, status, started_at, now=None):
    now = time.time() if now is None else now
    raw = raw if isinstance(raw, dict) else {}
    result = {'stage':str(raw.get('stage','starting'))[:50], 'label':str(raw.get('label','Iniciando motor'))[:100],
              'detail':str(raw.get('detail',''))[:250], 'completed':number(raw.get('completed'),0),
              'total':max(1,number(raw.get('total'),1)), 'updated_at':number(raw.get('updated_at'),now)}
    result['percent'] = 100 if status == 'complete' else min(99.5, max(0, number(raw.get('percent'), 5)))
    saved_elapsed=number(raw.get('elapsed_seconds')) if status!='running' else None
    result['elapsed_seconds'] = round(saved_elapsed if saved_elapsed is not None else elapsed_seconds(started_at, now), 1)
    end = number(raw.get('estimated_end_at'))
    result['remaining_seconds'] = math.ceil(end-now) if status == 'running' and end and end > now else None
    if status == 'complete':
        result.update(stage='complete', label='Entrega lista', remaining_seconds=0)
    return result


class ExportProgressReporter:
    def __init__(self, path, started_at, estimate_seconds=None, *, clock=time.time, stages=None):
        self.path = Path(path)
        self.started_at = started_at
        self.clock = clock
        self.stages = stages or STAGES
        self.stage = None
        self.stage_started = clock()
        self.percent = 5
        estimate_seconds=number(estimate_seconds,0)
        self.estimated_end = clock()+estimate_seconds if estimate_seconds > 0 else None

    def update(self, stage, completed=0, total=1, detail=''):
        now = self.clock()
        if stage not in self.stages:
            return
        if stage != self.stage:
            self.stage, self.stage_started = stage, now
        completed = max(0, number(completed, 0))
        total = max(1, number(total, 1))
        completed = min(completed, total)
        base, span, label = self.stages[stage]
        self.percent = max(self.percent, min(99.5, base+span*completed/total))
        # La velocidad de las diapositivas incluye los gráficos y animaciones.
        # Antes de observar dos unidades, usa solo historial o muestra «Estimando».
        if stage in {'presentation', 'research'} and completed >= 2 and now > self.stage_started:
            rate = (now-self.stage_started)/completed
            remaining = (total-completed)*rate + rate*total*(100-base-span)/span
            self.estimated_end = now+remaining
        raw = {'stage': stage, 'label': label, 'completed': completed, 'total': total,
               'detail': str(detail), 'percent': round(self.percent, 1), 'updated_at': now,
               'estimated_end_at': self.estimated_end}
        temporary = self.path.with_suffix('.tmp')
        try:
            temporary.write_text(json.dumps(raw, ensure_ascii=False, allow_nan=False), encoding='utf-8')
            temporary.replace(self.path)
        except OSError:
            pass  # Un bloqueo del archivo de avance no debe cancelar una entrega.
