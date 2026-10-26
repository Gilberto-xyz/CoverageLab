"""Archivos internos de una entrega, con lectura compatible con salidas anteriores."""
from pathlib import Path
import json


def system_dir(directory):
    directory = Path(directory).resolve()
    target = (directory / "_sistema").resolve()
    if target.parent != directory:
        raise ValueError("La carpeta del sistema debe estar dentro de la entrega.")
    target.mkdir(exist_ok=True)
    return target


def auxiliary(directory, name, *, write=False):
    directory = Path(directory).resolve()
    if Path(name).name != name:
        raise ValueError("Nombre de archivo auxiliar inválido.")
    target = system_dir(directory) / name if write else directory / "_sistema" / name
    if not write and not target.exists():
        target = directory / name
    resolved = target.resolve()
    if not resolved.is_relative_to(directory) or resolved == directory:
        raise ValueError("El archivo auxiliar está fuera de la entrega.")
    return resolved


def manifests(root, name):
    """Prefiere el manifiesto nuevo si una entrega conserva ambas ubicaciones."""
    found = {path.parent: path for path in Path(root).glob(f"*/{name}")}
    found.update({path.parent.parent: path for path in Path(root).glob(f"*/_sistema/{name}")})
    return sorted(found.values())


def manifest_directory(path):
    return path.parent.parent if path.parent.name == "_sistema" else path.parent


def organize_auxiliaries(directory, names, *, input_copy=False):
    """Mueve solo auxiliares reconocidos; nunca reemplaza un archivo existente."""
    directory = Path(directory).resolve()
    moved = {}
    for name in names:
        source = directory / name
        if not source.is_file():
            continue
        if source.resolve().parent != directory:
            raise ValueError("El archivo auxiliar está fuera de la entrega.")
        try:
            target = auxiliary(directory, name, write=True)
            if target.exists():
                continue
            source.replace(target)
        except OSError:
            # Una entrega anterior de solo lectura sigue siendo reconocible.
            continue
        moved[str(source)] = str(target)
    source = directory / "_input"
    if input_copy and source.is_dir():
        try:
            target = system_dir(directory) / "_input"
        except OSError:
            return moved
        if source.resolve().parent != directory or target.resolve().parent != system_dir(directory):
            raise ValueError("La copia de trabajo debe estar dentro de la entrega.")
        if not target.exists():
            try:
                source.rename(target)
            except OSError:
                return moved
            request_file = target / "request.json"
            if request_file.is_file() and request_file.resolve().parent == target:
                try:
                    request = json.loads(request_file.read_text(encoding="utf-8"))
                    request["source"] = str(target / Path(request["source"]).name)
                    request["result"] = str(auxiliary(directory, "_result.json"))
                    request_file.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
                except (OSError, ValueError, KeyError, TypeError):
                    pass
    return moved
