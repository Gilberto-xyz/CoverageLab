#!/usr/bin/env python3
"""Analiza y migra filtros/campos personalizados entre bases PowerView.

Por seguridad, el modo predeterminado solo analiza. La migracion se limita a
familias de campos de usuario detectadas en specs/, tables/ y usrfield/; no
copia archivos de datos, periodos, logs, respaldos ni resultados temporales.

Ejemplo con deteccion automatica en una carpeta contenedora:
    py migrar_filtros_power_view.py --raiz "C:\\Bases\\Cliente.gbl"

Ejemplo indicando cualquier par de bases:
    py migrar_filtros_power_view.py --origen "BASE.pw'anterior" --destino "BASE.pw"

Despues del analisis, presione ENTER para ejecutar la migracion.

Si se ejecuta sin rutas, abre dos ventanas para elegir las carpetas.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable


FILTER_DIRECTORIES = {"specs", "tables", "usrfield"}
USER_FIELD_DATA_SUFFIXES = {
    ".uak", ".ui0", ".ui1", ".ui2",
    ".rak", ".ri0", ".ri1", ".ri2",
    ".cak", ".ci0", ".ci1", ".ci2",
}
EXCLUDED_SUFFIXES = {".log", ".bak", ".backup", ".tmp", ".temp"}
PROTECTED_FAMILY_NAMES = {
    "period", "periods", "periodo", "periodos", "date", "dates", "fecha", "fechas"
}
OLD_NAME_RE = re.compile(
    r"(?i)(?:[' _.-]+(?:anterior|old|previous|prev|backup|respaldo))$"
)


class Style:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    MAGENTA = "\033[95m"
    CYAN = "\033[96m"
    WHITE = "\033[97m"
    PROMPT = "\033[30;103m"


COLOR_ENABLED = False


def configure_colors(disabled: bool = False) -> None:
    global COLOR_ENABLED
    COLOR_ENABLED = not disabled and sys.stdout.isatty() and "NO_COLOR" not in os.environ
    if not COLOR_ENABLED or os.name != "nt":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        output_handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(output_handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(output_handle, mode.value | 0x0004)
    except (AttributeError, OSError):
        COLOR_ENABLED = False


def styled(value: object, *styles: str) -> str:
    text = str(value)
    if not COLOR_ENABLED:
        return text
    return "".join(styles) + text + Style.RESET


def highlighted_number(value: int) -> str:
    return styled(f"{value:,}", Style.BOLD, Style.YELLOW)

# Referencias binarias/textuales a archivos base que deben existir en destino.
DEPENDENCY_RE = re.compile(
    rb"(?i)([a-z0-9_][a-z0-9_.-]{0,80}\.(?:it[0-9]|vl[0-9]|fl[0-9]|"
    rb"ml[0-9]|al[0-9]|ct[0-9]|wi[0-9]|ni[0-9]|nv[0-9]|si[0-9]|sv[0-9]))"
)


@dataclass(frozen=True)
class FileInfo:
    relative: Path
    full_path: Path
    size: int

    @property
    def key(self) -> str:
        return self.relative.as_posix().casefold()

    @property
    def family(self) -> str:
        return self.relative.stem.casefold()


@dataclass
class Comparison:
    source_only: list[FileInfo]
    destination_only: list[FileInfo]
    modified: list[tuple[FileInfo, FileInfo]]
    identical: list[tuple[FileInfo, FileInfo]]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def inventory(root: Path) -> dict[str, FileInfo]:
    result: dict[str, FileInfo] = {}
    for path in root.rglob("*"):
        if path.is_symlink() or not path.is_file():
            continue
        relative = path.relative_to(root)
        info = FileInfo(relative=relative, full_path=path, size=path.stat().st_size)
        if info.key in result:
            raise RuntimeError(
                f"Hay dos rutas que solo difieren por mayusculas/minusculas: {relative}"
            )
        result[info.key] = info
    return result


def compare(
    source_files: dict[str, FileInfo], destination_files: dict[str, FileInfo]
) -> Comparison:
    source_only: list[FileInfo] = []
    destination_only: list[FileInfo] = []
    modified: list[tuple[FileInfo, FileInfo]] = []
    identical: list[tuple[FileInfo, FileInfo]] = []

    for key, source in source_files.items():
        destination = destination_files.get(key)
        if destination is None:
            source_only.append(source)
            continue
        same = source.size == destination.size and sha256(source.full_path) == sha256(
            destination.full_path
        )
        (identical if same else modified).append((source, destination))

    for key, destination in destination_files.items():
        if key not in source_files:
            destination_only.append(destination)

    source_only.sort(key=lambda item: item.key)
    destination_only.sort(key=lambda item: item.key)
    modified.sort(key=lambda pair: pair[0].key)
    identical.sort(key=lambda pair: pair[0].key)
    return Comparison(source_only, destination_only, modified, identical)


def detect_custom_families(source_files: Iterable[FileInfo]) -> set[str]:
    families: set[str] = set()
    for info in source_files:
        parts = info.relative.parts
        if len(parts) < 2 or parts[0].casefold() != "usrfield":
            continue
        family = info.family
        if is_protected_family(family):
            continue
        if info.relative.suffix.casefold() in USER_FIELD_DATA_SUFFIXES:
            families.add(family)
    return families


def normalize_family(value: str) -> str:
    return Path(value.strip()).stem.casefold()


def is_protected_family(family: str) -> bool:
    normalized = family.casefold().lstrip("_.- ")
    return normalized in PROTECTED_FAMILY_NAMES


def is_filter_artifact(info: FileInfo, families: set[str]) -> bool:
    parts = info.relative.parts
    if len(parts) < 2 or parts[0].casefold() not in FILTER_DIRECTORIES:
        return False
    if info.family not in families:
        return False
    if info.relative.suffix.casefold() in EXCLUDED_SUFFIXES:
        return False
    return True


def find_dependencies(files: Iterable[FileInfo]) -> set[str]:
    dependencies: set[str] = set()
    for info in files:
        data = info.full_path.read_bytes()
        for match in DEPENDENCY_RE.finditer(data):
            dependencies.add(match.group(1).decode("ascii", errors="ignore").casefold())
    return dependencies


def validate_dependencies(
    selected: Iterable[FileInfo], destination_files: dict[str, FileInfo]
) -> list[str]:
    available_names = {
        item.relative.name.casefold() for item in destination_files.values()
    }
    available_names.update(item.relative.name.casefold() for item in selected)
    return sorted(find_dependencies(selected) - available_names)


def describe_list(
    title: str, files: Iterable[FileInfo], path_style: str = Style.GREEN
) -> None:
    items = list(files)
    print(
        "\n"
        + styled(title, Style.BOLD, Style.CYAN)
        + " ("
        + highlighted_number(len(items))
        + "):"
    )
    if not items:
        print(styled("  (ninguno)", Style.DIM))
        return
    for item in items:
        print(
            "  - "
            + styled(item.relative.as_posix(), path_style)
            + " ("
            + highlighted_number(item.size)
            + " bytes)"
        )


def choose_directory(title: str) -> Path:
    """Abre un selector de carpetas; usa consola si tkinter no esta disponible."""
    selected = ""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            selected = filedialog.askdirectory(title=title, mustexist=True)
        finally:
            root.destroy()
    except (ImportError, RuntimeError, OSError):
        if sys.stdin.isatty():
            selected = input(f"{title}\nRuta: ").strip().strip('"')
    if not selected:
        raise ValueError(
            "No se selecciono una carpeta. Use --origen y --destino para indicar rutas."
        )
    return Path(selected)


def discover_pair(root: Path) -> tuple[Path, Path]:
    """Detecta un par BASE'anterior -> BASE entre subcarpetas inmediatas."""
    root = root.expanduser().resolve()
    if not root.is_dir():
        raise FileNotFoundError(f"No existe la carpeta contenedora: {root}")
    directories = [path for path in root.iterdir() if path.is_dir()]
    by_name = {path.name.casefold(): path for path in directories}
    pairs: list[tuple[Path, Path]] = []
    for old in directories:
        new_name = OLD_NAME_RE.sub("", old.name)
        if new_name == old.name:
            continue
        new = by_name.get(new_name.casefold())
        if new is not None and new != old:
            pairs.append((old, new))
    if len(pairs) == 1:
        return pairs[0]
    if not pairs:
        raise ValueError(
            "No se encontro un par automatico. La base anterior debe llamarse como "
            "la nueva mas un sufijo: 'anterior, _anterior, -old o _respaldo."
        )
    options = "\n".join(f"  - {old.name} -> {new.name}" for old, new in pairs)
    raise ValueError(
        "Se encontraron varios pares; indique --origen y --destino:\n" + options
    )


def resolve_input_paths(args: argparse.Namespace) -> tuple[Path, Path]:
    if args.raiz is not None:
        if args.origen is not None or args.destino is not None:
            raise ValueError("Use --raiz o --origen/--destino, no ambos modos.")
        return discover_pair(args.raiz)
    source = args.origen
    destination = args.destino
    if source is None:
        source = choose_directory("Seleccione la base ANTERIOR que contiene los filtros")
    if destination is None:
        destination = choose_directory("Seleccione la base NUEVA de destino")
    return source, destination


def ensure_valid_roots(source: Path, destination: Path) -> tuple[Path, Path]:
    source = source.expanduser().resolve()
    destination = destination.expanduser().resolve()
    if not source.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de origen: {source}")
    if not destination.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de destino: {destination}")
    if source == destination:
        raise ValueError("El origen y el destino no pueden ser la misma carpeta.")
    if source in destination.parents or destination in source.parents:
        raise ValueError("El origen y el destino no pueden estar contenidos entre si.")
    for label, root in (("origen", source), ("destino", destination)):
        child_directories = {
            path.name.casefold() for path in root.iterdir() if path.is_dir()
        }
        present = FILTER_DIRECTORIES & child_directories
        if "usrfield" not in present or len(present) < 2:
            raise ValueError(
                f"La carpeta de {label} no parece una base PowerView compatible: {root}. "
                "Debe contener usrfield y al menos una de las carpetas specs/tables."
            )
    return source, destination


def atomic_copy(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.", suffix=".migrando", dir=destination.parent
    )
    os.close(file_descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def migrate(
    selected: list[FileInfo],
    destination: Path,
    destination_files: dict[str, FileInfo],
    source: Path,
    summary: dict[str, object],
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_root = (
        destination.parent
        / "_respaldos_migracion_power_view"
        / f"{destination.name}_{timestamp}"
    )
    backup_files_root = backup_root / "archivos_reemplazados"
    backup_root.mkdir(parents=True, exist_ok=False)

    required_bytes = sum(info.size for info in selected)
    if shutil.disk_usage(destination).free < required_bytes * 2 + 10_000_000:
        raise OSError("No hay espacio libre suficiente para migracion y respaldo.")

    created: list[Path] = []
    overwritten: list[tuple[Path, Path]] = []
    operations: list[dict[str, object]] = []

    try:
        for info in selected:
            target = destination / info.relative
            previous = destination_files.get(info.key)
            if previous is not None:
                backup = backup_files_root / info.relative
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(previous.full_path, backup)
                overwritten.append((target, backup))
                action = "reemplazado"
            else:
                created.append(target)
                action = "creado"

            expected_hash = sha256(info.full_path)
            atomic_copy(info.full_path, target)
            actual_hash = sha256(target)
            if actual_hash != expected_hash:
                raise OSError(f"Fallo la verificacion SHA-256 de {info.relative}")
            operations.append(
                {
                    "ruta": info.relative.as_posix(),
                    "accion": action,
                    "bytes": info.size,
                    "sha256": actual_hash,
                }
            )
    except Exception:
        for target in reversed(created):
            target.unlink(missing_ok=True)
        for target, backup in reversed(overwritten):
            atomic_copy(backup, target)
        raise

    manifest = {
        "fecha": datetime.now().astimezone().isoformat(),
        "origen": str(source),
        "destino": str(destination),
        "resumen_analisis": summary,
        "operaciones": operations,
        "rollback": {
            "eliminar_creados": [str(path) for path in created],
            "restaurar_reemplazados_desde": str(backup_files_root),
        },
    }
    manifest_path = backup_root / "manifiesto_migracion.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return manifest_path


def write_analysis_report(
    path: Path,
    source: Path,
    destination: Path,
    comparison: Comparison,
    families: set[str],
    selected: list[FileInfo],
    ignored_source_only: list[FileInfo],
    missing_dependencies: list[str],
) -> None:
    report = {
        "fecha": datetime.now().astimezone().isoformat(),
        "origen": str(source),
        "destino": str(destination),
        "conteos": {
            "solo_origen": len(comparison.source_only),
            "solo_destino": len(comparison.destination_only),
            "modificados": len(comparison.modified),
            "identicos": len(comparison.identical),
            "seleccionados": len(selected),
        },
        "familias_personalizadas": sorted(families),
        "seleccionados": [item.relative.as_posix() for item in selected],
        "solo_origen_ignorados": [
            item.relative.as_posix() for item in ignored_source_only
        ],
        "dependencias_faltantes": missing_dependencies,
        "modificados_no_reemplazados": [
            pair[0].relative.as_posix() for pair in comparison.modified
        ],
    }
    path = path.expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nReporte guardado en: {path}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Migra de forma segura filtros/campos personalizados entre bases PowerView. "
            "Analiza y despues solicita ENTER para copiar."
        )
    )
    parser.add_argument(
        "--origen",
        type=Path,
        help="Carpeta de la base anterior que contiene los filtros.",
    )
    parser.add_argument(
        "--destino",
        type=Path,
        help="Carpeta de la base nueva que recibira los filtros.",
    )
    parser.add_argument(
        "--raiz",
        type=Path,
        help=(
            "Carpeta que contiene ambas bases. Detecta automaticamente nombres como "
            "BASE.pw'anterior y BASE.pw."
        ),
    )
    parser.add_argument(
        "--actualizar-existentes",
        action="store_true",
        help=(
            "Tambien reemplaza artefactos modificados de las mismas familias de filtros. "
            "Usar solo si se sabe que son compatibles con la base nueva."
        ),
    )
    parser.add_argument(
        "--solo-analizar",
        action="store_true",
        help="Muestra el diagnostico y termina sin solicitar la migracion.",
    )
    parser.add_argument(
        "--sin-color",
        action="store_true",
        help="Desactiva los colores ANSI de la terminal.",
    )
    parser.add_argument(
        "--reporte",
        type=Path,
        help="Ruta opcional para guardar el analisis en JSON.",
    )
    parser.add_argument(
        "--incluir-familia",
        action="append",
        default=[],
        metavar="NOMBRE",
        help=(
            "Incluye una familia de filtro adicional si la deteccion automatica no la "
            "reconoce. Se puede repetir."
        ),
    )
    parser.add_argument(
        "--excluir-familia",
        action="append",
        default=[],
        metavar="NOMBRE",
        help="Excluye una familia detectada. Se puede repetir.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    configure_colors(args.sin_color)
    try:
        source_input, destination_input = resolve_input_paths(args)
        source, destination = ensure_valid_roots(source_input, destination_input)
        print(styled("Inventariando y comparando contenido (SHA-256)...", Style.BLUE))
        source_files = inventory(source)
        destination_files = inventory(destination)
        comparison = compare(source_files, destination_files)
        families = detect_custom_families(source_files.values())
        explicitly_included = {
            normalize_family(value) for value in args.incluir_familia if value.strip()
        }
        explicitly_excluded = {
            normalize_family(value) for value in args.excluir_familia if value.strip()
        }
        protected_requested = sorted(
            family for family in explicitly_included if is_protected_family(family)
        )
        if protected_requested:
            raise ValueError(
                "No se pueden incluir familias protegidas de fecha/periodo: "
                + ", ".join(protected_requested)
            )
        families.update(explicitly_included)
        families.difference_update(explicitly_excluded)

        selected = [
            item
            for item in comparison.source_only
            if is_filter_artifact(item, families)
        ]
        if args.actualizar_existentes:
            selected.extend(
                source_info
                for source_info, _ in comparison.modified
                if is_filter_artifact(source_info, families)
            )
            selected.sort(key=lambda item: item.key)

        selected_keys = {item.key for item in selected}
        ignored_source_only = [
            item for item in comparison.source_only if item.key not in selected_keys
        ]
        missing_dependencies = validate_dependencies(selected, destination_files)

        print("\n" + styled("ORIGEN :", Style.BOLD, Style.BLUE) + " " + styled(source, Style.CYAN))
        print(styled("DESTINO:", Style.BOLD, Style.MAGENTA) + " " + styled(destination, Style.MAGENTA))
        print("\n" + styled("Resumen de comparacion:", Style.BOLD, Style.WHITE))
        print("  Solo en origen : " + highlighted_number(len(comparison.source_only)))
        print("  Solo en destino: " + highlighted_number(len(comparison.destination_only)))
        print(
            "  Mismo nombre, contenido diferente: "
            + highlighted_number(len(comparison.modified))
        )
        print("  Identicos: " + highlighted_number(len(comparison.identical)))
        print(
            "  Familias personalizadas detectadas: "
            + styled(", ".join(sorted(families)) or "ninguna", Style.BOLD, Style.GREEN)
        )

        describe_list("Archivos seleccionados para migrar", selected)
        describe_list(
            "Exclusivos del origen ignorados por seguridad",
            ignored_source_only,
            Style.DIM,
        )

        if missing_dependencies:
            print("\n" + styled("ERROR: faltan dependencias en la base destino:", Style.BOLD, Style.RED))
            for dependency in missing_dependencies:
                print("  - " + styled(dependency, Style.RED))
        else:
            print(
                "\n"
                + styled(
                    "Dependencias correctas: las referencias base existen en destino.",
                    Style.GREEN,
                )
            )

        summary: dict[str, object] = {
            "solo_origen": len(comparison.source_only),
            "solo_destino": len(comparison.destination_only),
            "modificados": len(comparison.modified),
            "identicos": len(comparison.identical),
            "seleccionados": len(selected),
            "familias": sorted(families),
        }

        if args.reporte:
            write_analysis_report(
                args.reporte,
                source,
                destination,
                comparison,
                families,
                selected,
                ignored_source_only,
                missing_dependencies,
            )

        if args.solo_analizar:
            print("\n" + styled("SOLO ANALISIS: no se modifico ningun archivo.", Style.YELLOW))
            return 0

        if not selected:
            print("\n" + styled("No hay filtros nuevos que migrar.", Style.GREEN))
            return 0
        if missing_dependencies:
            print(
                "\n" + styled("Migracion cancelada por dependencias faltantes.", Style.RED),
                file=sys.stderr,
            )
            return 2

        print("\n" + styled("IMPORTANTE: cierre PowerView antes de continuar.", Style.BOLD, Style.YELLOW))
        prompt = styled(
            f" Presione ENTER para copiar {len(selected):,} archivos de filtro ",
            Style.BOLD,
            Style.PROMPT,
        )
        try:
            answer = input(prompt + " (cualquier texto cancela): ")
        except (EOFError, KeyboardInterrupt):
            print("\n" + styled("Operacion cancelada; no se modifico ningun archivo.", Style.YELLOW))
            return 0
        if answer.strip():
            print(styled("Operacion cancelada; no se modifico ningun archivo.", Style.YELLOW))
            return 0

        manifest_path = migrate(
            selected, destination, destination_files, source, summary
        )
        print(
            "\n"
            + styled("Migracion completada: ", Style.BOLD, Style.GREEN)
            + highlighted_number(len(selected))
            + styled(" archivos verificados.", Style.GREEN)
        )
        print(styled("Respaldo y manifiesto: ", Style.BLUE) + styled(manifest_path, Style.CYAN))
        return 0
    except (OSError, ValueError, RuntimeError) as error:
        print(styled(f"ERROR: {error}", Style.BOLD, Style.RED), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
