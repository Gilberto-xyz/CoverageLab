"""Organiza los datos de instalaciones anteriores antes de iniciar la web.

Solo usa la biblioteca estándar: puede ejecutarse antes de instalar dependencias.
Los libros y las presentaciones se mueven íntegros. Las rutas de los JSON internos
se ajustan con una copia de respaldo; cada movimiento se registra antes de hacerlo.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager, suppress
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import socket
import stat
import tempfile
import time


REPORT = 'organizacion_inicial_v1.json'
AUXILIARIES = {'job.json', 'research-job.json', 'selection.json', 'result.json',
               '_result.json', 'request.json', 'investigacion_5w1h.json'}
CONTAINERS = {'Output', 'outputs', 'outputs/pipeline_explorer',
              'outputs/pipeline_explorer/selections', 'inputs', 'inputs/5w1h',
              'inputs/drafts', '.estado', '.estado/selections'}


def _linked(path):
    if not path.exists() and not path.is_symlink(): return False
    return path.is_symlink() or bool(getattr(path.lstat(), 'st_file_attributes', 0)
                                    & getattr(stat, 'FILE_ATTRIBUTE_REPARSE_POINT', 0x400))


def _checked(path, project):
    path = Path(path)
    if not path.resolve().is_relative_to(project) or path.resolve() == project:
        raise ValueError('La ruta debe permanecer dentro de la carpeta del proyecto.')
    for part in [path, *path.parents]:
        if part == project: break
        if _linked(part): raise ValueError('Se conserva un enlace o punto de unión sin moverlo.')
    return path


def _check_tree(path, project):
    _checked(path, project)
    if path.is_dir():
        for child in path.iterdir(): _check_tree(child, project)


def _atomic_json(path, value):
    temporary = None
    try:
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=path.parent,
                                         prefix='.organizacion-', delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.flush(); os.fsync(stream.fileno())
        for attempt in range(6):
            try:
                temporary.replace(path)
                break
            except PermissionError:
                if attempt == 5: raise
                time.sleep(.02 * 2 ** attempt)
    finally:
        if temporary and temporary.exists():
            with suppress(OSError): temporary.unlink()


@contextmanager
def _lock(path):
    with path.open('a+b') as stream:
        stream.seek(0, 2)
        if not stream.tell(): stream.write(b'0'); stream.flush()
        stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == 'nt': msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else: fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def _plan(project):
    data = project / 'recursos'
    units, containers = [], []

    def add(source, target):
        relative = source.relative_to(project).as_posix()
        if source.is_dir() and relative.casefold() in {name.casefold() for name in CONTAINERS} and not _linked(source):
            containers.append(source)
            for child in sorted(source.iterdir()):
                if child.name in {'.runs', '__pycache__'} or child.name.startswith('~$'): continue
                add(child, target / child.name)
        else:
            units.append((source, target))

    for name in ('Output', 'outputs', 'inputs', '.estado'):
        source = project / name
        if source.exists() or source.is_symlink(): add(source, data / name)
    for source in sorted(project.iterdir()):
        if source.name.startswith(('.', '~$')) or source.name.casefold() in {'scripts', 'recursos', 'output', 'outputs', 'inputs'}:
            continue
        lower = source.name.casefold()
        if source.is_file() and source.suffix.casefold() in {'.xlsx', '.pptx'}:
            if lower.startswith(('modelo', 'plantilla')): continue
            if source.suffix.casefold() == '.pptx' or lower.startswith(('template_', 'banco_', 'reporte', 'scorecard')):
                target = data / 'Output' / 'Archivos_anteriores' / source.name
            elif '5w1h' in source.stem.casefold():
                target = data / 'inputs' / '5w1h' / source.name
            else:
                target = data / source.name
            units.append((source, target))
        elif source.is_dir() and not _linked(source):
            names = [p.name.casefold() for p in source.iterdir() if p.is_file()]
            if any(n.startswith('template_') for n in names) and any(n.startswith('banco_') for n in names) and any(n.endswith('.pptx') for n in names):
                units.append((source, data / 'Output' / source.name))
    return units, containers


def _translated(value, mappings):
    if isinstance(value, dict): return {key: _translated(item, mappings) for key, item in value.items()}
    if isinstance(value, list): return [_translated(item, mappings) for item in value]
    if not isinstance(value, str): return value
    normalized = value.replace('\\', '/').rstrip('/')
    for original, target in mappings:
        prefix = original.replace('\\', '/').rstrip('/')
        if normalized.casefold() == prefix.casefold() or normalized.casefold().startswith(prefix.casefold() + '/'):
            suffix = normalized[len(prefix):].lstrip('/')
            result = target / suffix
            if result.resolve().is_relative_to(target.resolve()): return str(result)
    return value


def _repair(records, project, state, pending):
    moved, mappings = [], []
    for record in records:
        if record.get('status') not in {'planned', 'moved'}: continue
        source, target = project / record['source'], project / record['target']
        _checked(source, project); _checked(target, project)
        if not target.resolve().is_relative_to(project / 'recursos'): continue
        if record['status'] in {'planned', 'moved'} and not source.exists() and target.exists():
            record['status'] = 'moved'; moved.append(target); mappings.append((str(source), target))
    mappings.sort(key=lambda item: len(item[0]), reverse=True)
    for target in moved:
        files = [target] if target.is_file() else list(target.rglob('*.json'))
        for file in files:
            if file.suffix.casefold() != '.json' or file.name == REPORT: continue
            parts = file.relative_to(project / 'recursos').parts
            if file.name not in AUXILIARIES and not any(p in {'selections', 'drafts', '.estado'} for p in parts): continue
            try:
                _checked(file, project)
                raw = file.read_bytes()
                if len(raw) > 16 * 1024 * 1024: raise ValueError('El JSON auxiliar es demasiado grande para actualizar sus rutas.')
                value = json.loads(raw.decode('utf-8-sig'))
                local_maps = list(mappings)
                # A copied delivery may still mention the previous computer's path.
                if isinstance(value, dict) and isinstance(value.get('directory'), str):
                    old = value['directory'].replace('\\', '/').rstrip('/')
                    folder = file.parent.parent if file.parent.name == '_sistema' else file.parent
                    if old.rsplit('/', 1)[-1].casefold() == folder.name.casefold(): local_maps.insert(0, (old, folder))
                updated = _translated(value, local_maps)
                if isinstance(updated, dict):
                    for key in ('output_root', 'root', 'base_dir'):
                        old = updated.get(key)
                        if not isinstance(old, str): continue
                        for container in ('Output', 'outputs', 'inputs', '.estado'):
                            if old.replace('\\', '/').rstrip('/').casefold() == str(project / container).replace('\\', '/').casefold():
                                updated[key] = str(project / 'recursos' / container)
                if updated == value: continue
                backup_dir = _checked(state / 'respaldo_organizacion', project)
                backup_dir.mkdir(exist_ok=True)
                backup = backup_dir / (hashlib.sha256(str(file).encode() + raw).hexdigest() + '.json')
                if not backup.exists():
                    with backup.open('xb') as stream: stream.write(raw)
                _atomic_json(file, updated)
            except (OSError, ValueError, TypeError) as error:
                pending.append({'path': str(file.relative_to(project)), 'reason': f'Rutas auxiliares pendientes: {error}'})


def prepare_project(project):
    project = Path(project).resolve()
    if not project.is_dir(): raise ValueError('No se encontró la carpeta del proyecto.')
    data = _checked(project / 'recursos', project); data.mkdir(exist_ok=True)
    state = _checked(data / '.estado', project); state.mkdir(exist_ok=True)
    report_path = _checked(state / REPORT, project)
    with _lock(_checked(state / 'organizacion_inicial.lock', project)):
        report = json.loads(report_path.read_text(encoding='utf-8')) if report_path.exists() else {'version': 1, 'complete': False, 'moves': []}
        if report.get('version') != 1 or not isinstance(report.get('moves'), list):
            raise ValueError('El registro de organización no es válido; se conserva para revisión.')
        units, containers = _plan(project)
        if report.get('complete') and not units: return dict(report, moved_now=0)
        report['complete'] = False
        pending, moved_now = [], 0
        records = {item['source']: item for item in report['moves']}
        for source, target in units:
            relative = source.relative_to(project).as_posix()
            record = records.setdefault(relative, {'source': relative, 'target': target.relative_to(project).as_posix(), 'status': 'planned'})
            report['moves'] = list(records.values()); _atomic_json(report_path, report)
            try:
                _check_tree(source, project); _checked(target, project)
                if target.exists() or target.is_symlink(): raise FileExistsError('El destino ya existe; se conservan ambos archivos o carpetas.')
                target.parent.mkdir(parents=True, exist_ok=True)
                # On Windows rename refuses to replace an existing destination.
                os.rename(source, target)
                record['status'] = 'moved'; record.pop('error', None); moved_now += 1
            except (OSError, ValueError) as error:
                record['status'] = 'pending'; record['error'] = str(error)
                pending.append({'path': relative, 'reason': str(error)})
            _atomic_json(report_path, report)
        _repair(report['moves'], project, state, pending)
        for folder in sorted(containers, key=lambda p: len(p.parts), reverse=True):
            try:
                _checked(folder, project)
                if folder.is_dir() and not any(folder.iterdir()): folder.rmdir()
            except OSError: pass
        report.update(complete=not pending, pending=pending, updated_at=datetime.now(timezone.utc).isoformat())
        _atomic_json(report_path, report)
        return dict(report, moved_now=moved_now)


def main(argv=None):
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--port', type=int, default=8765)
    args, _ = parser.parse_known_args(argv)
    if not 0 <= args.port <= 65535:
        print('El puerto debe estar entre 0 y 65535.'); return 1
    if args.port:
        try:
            with socket.create_connection(('127.0.0.1', args.port), timeout=.4):
                print('CoverageLab ya está ejecutándose. Cierra el servidor anterior (Ctrl+C en su ventana) antes de volver a iniciar y organizar sus archivos.')
                return 1
        except OSError: pass
    try:
        project = Path(__file__).resolve().parent.parent
        result = prepare_project(project)
        if result['moved_now'] or result.get('pending'):
            print(f"Organización: {result['moved_now']} elementos movidos; {len(result.get('pending', []))} pendientes.")
            print(f"Registro: {project / 'recursos' / '.estado' / REPORT}")
        return 0
    except (OSError, ValueError, KeyError, TypeError) as error:
        print(f'No se pudo preparar la carpeta del proyecto: {error}')
        return 1


if __name__ == '__main__': raise SystemExit(main())
