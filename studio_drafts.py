"""Borradores de captura: conservan valores pendientes sin escribir en Excel."""
from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path
import re
import tempfile
import threading

from studio_inputs import (MAX_CELL_TEXT, MAX_COLUMNS, MAX_EDIT_CELLS, MAX_ROWS,
                           MAX_SHEETS, file_fingerprint)


MAX_DRAFT_BYTES = 4 * 1024 * 1024


def _coordinate(value, name, maximum):
    if type(value) is not int or not 1 <= value <= maximum:
        raise ValueError(f"{name} debe ser un entero entre 1 y {maximum}.")
    return value


def _sheet(value):
    if not isinstance(value, str) or not value or len(value) > 31 or re.search(r"[\\/*?:\[\]\x00-\x1f]", value):
        raise ValueError("Nombre de hoja inválido en el borrador.")
    return value


class StudioDrafts:
    def __init__(self, inputs):
        self.inputs = inputs
        self.root = inputs.root
        self.directory = self.root / "inputs" / "drafts"
        self.lock = threading.RLock()

    def _source(self, file, kind):
        path, origin = self.inputs._path(file, kind)
        canonical = path.name if kind == "coverage" else f"{origin}/{path.name}"
        return path, canonical

    def _path(self, canonical, kind, *, create=False):
        # Validate before mkdir as well as before each read/write/delete: inputs can
        # be replaced with a link while the application is running.
        for directory in (self.root / "inputs", self.directory):
            if not directory.resolve().is_relative_to(self.root) or directory.is_symlink():
                raise ValueError("La carpeta de borradores debe estar dentro del proyecto, sin enlaces.")
            if directory.exists() and not directory.is_dir():
                raise ValueError("La carpeta de borradores no es un directorio.")
        if create:
            self.directory.mkdir(parents=True, exist_ok=True)
        key = hashlib.sha256(json.dumps([kind, canonical], ensure_ascii=False,
                                       separators=(",", ":")).encode("utf-8")).hexdigest()
        path = self.directory / f"{key}.json"
        if path.is_symlink() or not path.resolve().is_relative_to(self.directory.resolve()):
            raise ValueError("El borrador está fuera de la carpeta permitida.")
        return path

    @staticmethod
    def _document(payload, canonical, kind, *, updated_at=None):
        fingerprint = payload.get("fingerprint")
        if not isinstance(fingerprint, str) or not re.fullmatch(r"[a-f0-9]{64}", fingerprint):
            raise ValueError("La huella del archivo del borrador es inválida.")
        sheet = _sheet(payload.get("sheet"))
        row = _coordinate(payload.get("row", 1), "La fila", MAX_ROWS)
        column = _coordinate(payload.get("column", 1), "La columna", MAX_COLUMNS)
        number_locale = payload.get("number_locale", "auto")
        if not isinstance(number_locale, str) or number_locale not in {"auto", "decimal-dot", "decimal-comma"}:
            raise ValueError("Formato numérico del borrador desconocido.")
        edits = payload.get("edits")
        if not isinstance(edits, list) or len(edits) > MAX_EDIT_CELLS:
            raise ValueError("Lista de cambios del borrador inválida.")
        prepared, sheets, count = [], {sheet}, 0
        for edit in edits:
            if not isinstance(edit, dict):
                raise ValueError("Cambio del borrador inválido.")
            name = _sheet(edit.get("sheet"))
            sheets.add(name)
            if len(sheets) > MAX_SHEETS:
                raise ValueError(f"El borrador admite como máximo {MAX_SHEETS} hojas.")
            start_row = _coordinate(edit.get("row"), "La primera fila", MAX_ROWS)
            start_column = _coordinate(edit.get("column"), "La primera columna", MAX_COLUMNS)
            values = edit.get("values")
            if not isinstance(values, list) or not values or not isinstance(values[0], list) or not values[0]:
                raise ValueError("Cada cambio debe ser un bloque rectangular de celdas.")
            width = len(values[0])
            if start_row + len(values) - 1 > MAX_ROWS or start_column + width - 1 > MAX_COLUMNS:
                raise ValueError("El bloque del borrador excede las filas o columnas permitidas.")
            count += len(values) * width
            if count > MAX_EDIT_CELLS:
                raise ValueError(f"El borrador admite como máximo {MAX_EDIT_CELLS} celdas.")
            block = []
            for cells in values:
                if not isinstance(cells, list) or len(cells) != width:
                    raise ValueError("Cada cambio debe ser un bloque rectangular de celdas.")
                copied = []
                for value in cells:
                    if value is None:
                        pass
                    elif isinstance(value, str):
                        if len(value) > MAX_CELL_TEXT:
                            raise ValueError(f"Cada celda admite como máximo {MAX_CELL_TEXT} caracteres.")
                    elif type(value) in (int, float):
                        if isinstance(value, float) and not math.isfinite(value):
                            raise ValueError("Los números del borrador deben ser finitos.")
                    else:
                        raise ValueError("Las celdas del borrador solo admiten texto, números o valores vacíos.")
                    copied.append(value)
                block.append(copied)
            prepared.append({"sheet": name, "row": start_row, "column": start_column, "values": block})
        timestamp = updated_at or datetime.now(timezone.utc).isoformat()
        return {"file": canonical, "kind": kind, "fingerprint": fingerprint, "sheet": sheet,
                "row": row, "column": column, "number_locale": number_locale,
                "edits": prepared, "updated_at": timestamp}

    def get(self, file, kind="coverage"):
        with self.lock:
            source, canonical = self._source(file, kind)
            path = self._path(canonical, kind)
            if not path.exists():
                return {"draft": None, "source_changed": False}
            if not path.is_file() or path.stat().st_size > MAX_DRAFT_BYTES:
                raise ValueError("El archivo de borrador es inválido o demasiado grande.")
            stored = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(stored, dict) or stored.get("file") != canonical or stored.get("kind") != kind:
                raise ValueError("El borrador no pertenece al archivo seleccionado.")
            timestamp = stored.get("updated_at")
            if not isinstance(timestamp, str):
                raise ValueError("Fecha del borrador inválida.")
            datetime.fromisoformat(timestamp)
            draft = self._document(stored, canonical, kind, updated_at=timestamp)
            return {"draft": draft, "source_changed": file_fingerprint(source) != draft["fingerprint"]}

    def save(self, payload):
        if not isinstance(payload, dict):
            raise ValueError("Solicitud de borrador inválida.")
        with self.lock:
            kind = payload.get("kind", "coverage")
            source, canonical = self._source(payload.get("file"), kind)
            draft = self._document(payload, canonical, kind)
            data = json.dumps(draft, ensure_ascii=False, allow_nan=False,
                              separators=(",", ":")).encode("utf-8")
            if len(data) > MAX_DRAFT_BYTES:
                raise ValueError("El borrador es demasiado grande.")
            path = self._path(canonical, kind, create=True)
            temporary = None
            try:
                with tempfile.NamedTemporaryFile(dir=self.directory, prefix=".draft_", suffix=".tmp", delete=False) as stream:
                    temporary = Path(stream.name)
                    stream.write(data)
                    stream.flush()
                    os.fsync(stream.fileno())
                self._path(canonical, kind)
                os.replace(temporary, path)
            finally:
                if temporary is not None and temporary.exists():
                    temporary.unlink()
            return {"draft": draft, "source_changed": file_fingerprint(source) != draft["fingerprint"]}

    def discard(self, file, kind="coverage"):
        with self.lock:
            _, canonical = self._source(file, kind)
            path = self._path(canonical, kind)
            path.unlink(missing_ok=True)
            return {"draft": None, "source_changed": False}
