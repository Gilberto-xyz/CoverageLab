"""Apertura de carpetas y Excel de escritorio para los archivos registrados.

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
import threading
import time

from studio_inputs import InputConflictError, MAX_FILE_BYTES, _lock, file_fingerprint


_folder_lock = threading.Lock()


def _folder_window(shell, directory, gui):
    """Identifica Explorer por su ruta completa, no por el título de la ventana."""
    try:
        windows = list(shell.Windows())
    except Exception:
        return None
    for window in windows:
        try:
            if Path(window.FullName).name.casefold() != "explorer.exe":
                continue
            if Path(window.Document.Folder.Self.Path).resolve() != directory:
                continue
            handle = int(window.HWND)
            if gui.IsWindow(handle):
                return handle
        except Exception:
            # Una pestaña puede cerrarse mientras Windows enumera las ventanas.
            continue
    return None


def _foreground_folder(handle, gui, process, api, constants):
    """Activa solo la ventana solicitada y libera siempre las colas de entrada."""
    attached = []
    try:
        if gui.IsIconic(handle):
            gui.ShowWindow(handle, constants.SW_RESTORE)
        if gui.GetForegroundWindow() == handle:
            return True
        try:
            gui.SetForegroundWindow(handle)
        except Exception:
            pass
        if gui.GetForegroundWindow() == handle:
            return True
        current = api.GetCurrentThreadId()
        foreground = gui.GetForegroundWindow()
        threads = {process.GetWindowThreadProcessId(handle)[0]}
        if foreground:
            threads.add(process.GetWindowThreadProcessId(foreground)[0])
        # Crea la cola del hilo HTTP antes de enlazarla durante esta acción.
        gui.PeekMessage(None, 0, 0, constants.PM_NOREMOVE)
        for thread in threads - {current}:
            try:
                process.AttachThreadInput(current, thread, True)
                attached.append(thread)
            except Exception:
                pass
        gui.BringWindowToTop(handle)
        gui.SetForegroundWindow(handle)
        return gui.GetForegroundWindow() == handle
    except Exception:
        return False
    finally:
        for thread in reversed(attached):
            try:
                process.AttachThreadInput(current, thread, False)
            except Exception:
                pass


def open_folder(directory: Path) -> dict:
    """Abre una carpeta validada por el llamador y activa su ventana en Windows."""
    directory = Path(directory).resolve()
    if not directory.is_dir():
        raise ValueError("La carpeta de resultados ya no está disponible.")
    if sys.platform != "win32":
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(directory)],
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return {"message": "Carpeta de resultados abierta.", "directory": str(directory)}
    try:
        import pythoncom
        import win32api
        import win32con
        import win32gui
        import win32process
        import win32com.client
    except ImportError:
        os.startfile(str(directory))
        return {"message": "Carpeta abierta. Selecciona el icono del Explorador en la barra de tareas.",
                "directory": str(directory), "foreground": False}
    with _folder_lock:
        initialized = False
        shell = None
        try:
            try:
                pythoncom.CoInitialize()
                initialized = True
                shell = win32com.client.Dispatch("Shell.Application")
            except Exception:
                shell = None
            handle = _folder_window(shell, directory, win32gui) if shell else None
            if not handle:
                os.startfile(str(directory))
                deadline = time.monotonic() + 3
                while shell and not handle and time.monotonic() < deadline:
                    time.sleep(0.1)
                    handle = _folder_window(shell, directory, win32gui)
            focused = bool(handle and _foreground_folder(handle, win32gui, win32process, win32api, win32con))
        finally:
            shell = None
            if initialized:
                pythoncom.CoUninitialize()
    return {"message": ("Carpeta de resultados abierta en primer plano." if focused else
                        "Carpeta abierta. Selecciona el icono del Explorador en la barra de tareas."),
            "directory": str(directory), "foreground": focused}


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
