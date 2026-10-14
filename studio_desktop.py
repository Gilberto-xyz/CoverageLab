"""Excel de escritorio como editor opcional de las entradas registradas.

No automatiza ni cierra sesiones de Excel. El usuario abre un libro, lo guarda
en Excel y vuelve a leerlo en la web; los motores conservan el mismo contrato.
"""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import re
import subprocess
import sys

from studio_inputs import InputConflictError, MAX_FILE_BYTES, _lock, file_fingerprint


def _excel_path(value) -> Path | None:
    """El registro debe contener una ruta a EXCEL.EXE, nunca un comando."""
    if not isinstance(value, str) or not value.strip():
        return None
    value = value.strip()
    if value.startswith('"') and value.endswith('"'):
        value = value[1:-1]
    if '"' in value or "\x00" in value:
        return None
    try:
        raw = Path(value)
        if not raw.is_absolute():
            return None
        candidate = raw.resolve()
        if candidate.name.casefold() != "excel.exe" or not candidate.is_file():
            return None
        return candidate
    except (OSError, ValueError):
        return None


def find_excel_executable() -> Path | None:
    """Lee App Paths de Windows sin ejecutar búsquedas ni comandos del registro."""
    if sys.platform != "win32":
        return None
    try:
        import winreg
    except ImportError:
        return None
    key_name = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    views = dict.fromkeys((0, winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY))
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in views:
            try:
                with winreg.OpenKey(hive, key_name, 0, winreg.KEY_READ | view) as key:
                    value, kind = winreg.QueryValueEx(key, "")
                if kind not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ):
                    continue
                if kind == winreg.REG_EXPAND_SZ:
                    value = os.path.expandvars(value)
                candidate = _excel_path(value)
                if candidate:
                    return candidate
            except OSError:
                continue
    return None


def launch_excel(executable: Path, workbook: Path) -> None:
    validated = _excel_path(str(executable))
    if validated is None:
        raise ValueError("La instalación de Excel ya no está disponible. Usa el editor web o revisa Office.")
    if workbook.suffix.lower() != ".xlsx" or not workbook.is_file():
        raise ValueError("El libro de entrada ya no está disponible.")
    try:
        subprocess.Popen([str(validated), str(workbook)], shell=False,
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL)
    except OSError as exc:
        raise ValueError("No se pudo abrir Excel de escritorio. Revisa la instalación de Office.") from exc


class DesktopStudio:
    def __init__(self, inputs, *, finder=None, launcher=None):
        self.inputs = inputs
        self.finder = finder or find_excel_executable
        self.launcher = launcher or launch_excel

    def availability(self) -> dict:
        available = self.finder() is not None
        message = ("Excel de escritorio disponible. Guarda el libro en Excel y pulsa Actualizar desde Excel."
                   if available else "Excel de escritorio no está instalado o no está registrado en Windows. Puedes usar el editor web.")
        if not available and sys.platform != "win32":
            message = "La conexión con Excel de escritorio está disponible en Windows. Puedes usar el editor web."
        return {"available": available, "message": message}

    def _registered_path(self, payload) -> tuple[Path, str]:
        kind = payload.get("kind", "coverage")
        path, origin = self.inputs._path(payload.get("file"), kind)
        reference = path.name if kind == "coverage" else f"{origin}/{path.name}"
        if not any(item["file"] == reference for item in self.inputs.files()[kind]):
            raise ValueError("Selecciona un archivo de entrada del catálogo.")
        return path, origin

    def _workbook(self, payload, *, reference=None) -> dict:
        return self.inputs.workbook(reference or payload.get("file"),
                                    kind=payload.get("kind", "coverage"),
                                    sheet=payload.get("sheet") or None,
                                    **{key: payload.get(key, default) for key, default in
                                       (("row", 1), ("limit", 60), ("column", 1), ("columns", 32))})

    @staticmethod
    def _expected(payload) -> str:
        expected = payload.get("fingerprint")
        if not isinstance(expected, str) or not re.fullmatch(r"[0-9a-f]{64}", expected):
            raise ValueError("Vuelve a cargar la entrada antes de abrirla en Excel.")
        return expected

    def open(self, payload) -> dict:
        path, origin = self._registered_path(payload)
        expected = self._expected(payload)
        executable = self.finder()
        if executable is None:
            raise ValueError("Excel de escritorio no está disponible. Revisa Office en Windows o usa el editor web.")
        copied_path = None
        with _lock(path):
            book = self._workbook(payload)
            if book["fingerprint"] != expected:
                raise InputConflictError("El archivo cambió. Vuelve a cargar la entrada antes de abrir Excel.")
            if origin == "external":
                if path.stat().st_size > MAX_FILE_BYTES:
                    raise ValueError("El libro admite como máximo 10 MiB.")
                content = path.read_bytes()
                if hashlib.sha256(content).hexdigest() != expected or file_fingerprint(path) != expected:
                    raise InputConflictError("El archivo cambió al copiarlo. Vuelve a cargar la entrada.")
                copied = self.inputs.import_workbook({"kind": "research", "filename": path.name, "content": content})
                copied_path, _ = self.inputs._path(copied["file"], "research")
                try:
                    if file_fingerprint(path) != expected:
                        raise InputConflictError("El archivo cambió al copiarlo. Vuelve a cargar la entrada.")
                    book = self._workbook(payload, reference=copied["file"])
                except Exception:
                    # Solo retira la copia recién creada, dentro de entradas locales.
                    if copied_path.resolve().is_relative_to(self.inputs.research_local.resolve()):
                        copied_path.unlink(missing_ok=True)
                    raise
                book["copied_from"] = payload["file"]
                path = copied_path
            elif file_fingerprint(path) != expected:
                raise InputConflictError("El archivo cambió. Vuelve a cargar la entrada antes de abrir Excel.")
            self.launcher(Path(executable), path)
            book["desktop_opened"] = True
            book["message"] = "Libro abierto en Excel. Guarda allí y pulsa Actualizar desde Excel para continuar en la web."
            return book

    def reload(self, payload) -> dict:
        path, _ = self._registered_path(payload)
        with _lock(path):
            book = self._workbook(payload)
        book["message"] = "Entrada actualizada desde el archivo guardado en Excel."
        return book
