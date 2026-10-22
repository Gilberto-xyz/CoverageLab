# Vendored from 5W1H/generador_informe_5w1h.py (2026-10-13).
# Calculation/chart functions preserved; CLI entry changed to explicit isolated worker.
# -*- coding: utf-8 -*-
from __future__ import annotations
"""Genera informes 5W1H en PowerPoint a partir de datos Excel.

Incluye tablas y graficos para multiples segmentos y comparativos.
"""
# Bibliotecas necesarias
#---------------------------------------------------------------------------------------------------------------------
import warnings
from datetime import datetime as dt
import os
import sys
import io
import colorsys
from copy import copy
import csv
import math
import textwrap
import re
import unicodedata
import threading
import tempfile
import uuid
import zipfile
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from functools import lru_cache
from itertools import cycle
from dataclasses import dataclass
from decimal import Decimal, ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_DOWN, ROUND_HALF_EVEN, ROUND_HALF_UP, ROUND_UP, ROUND_05UP
from typing import Dict, Iterable, List, NamedTuple, Optional, Tuple
from pathlib import Path

try:
    from . import confidence_design
except ImportError:
    import confidence_design
from collections import OrderedDict
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Modulos pesados: se cargan en background para reducir el tiempo al primer prompt.
pd = None
np = None
matplotlib = None
plt = None
FuncFormatter = None
mcolors = None
mpatches = None
Rectangle = None
FontProperties = None
Presentation = None
Inches = None
Cm = None
Pt = None
RGBColor = None

_HEAVY_LOADER_THREAD: Optional[threading.Thread] = None
_HEAVY_MODULES_READY = threading.Event()
_HEAVY_MODULES_ERROR: Optional[Exception] = None
_HEAVY_LOAD_LOCK = threading.Lock()

def _load_heavy_modules() -> None:
    """Importa modulos pesados y los deja disponibles en globals."""
    global pd, np, matplotlib, plt, FuncFormatter, mcolors, mpatches, Rectangle, FontProperties
    global Presentation, Inches, Cm, Pt, RGBColor, _HEAVY_MODULES_ERROR
    if _HEAVY_MODULES_READY.is_set():
        return
    try:
        import pandas as _pd
        import numpy as _np
        import matplotlib as _matplotlib
        _matplotlib.use("Agg")
        from matplotlib import pyplot as _plt
        from matplotlib.ticker import FuncFormatter as _FuncFormatter
        from matplotlib import colors as _mcolors
        import matplotlib.patches as _mpatches
        from matplotlib.patches import Rectangle as _Rectangle
        from matplotlib.font_manager import FontProperties as _FontProperties
        from pptx import Presentation as _Presentation
        from pptx.util import Inches as _Inches, Cm as _Cm, Pt as _Pt
        from pptx.dml.color import RGBColor as _RGBColor

        pd = _pd
        np = _np
        matplotlib = _matplotlib
        plt = _plt
        FuncFormatter = _FuncFormatter
        mcolors = _mcolors
        mpatches = _mpatches
        Rectangle = _Rectangle
        FontProperties = _FontProperties
        Presentation = _Presentation
        Inches = _Inches
        Cm = _Cm
        Pt = _Pt
        RGBColor = _RGBColor
        _HEAVY_MODULES_ERROR = None
    except Exception as exc:
        _HEAVY_MODULES_ERROR = exc
    finally:
        _HEAVY_MODULES_READY.set()

def start_heavy_module_loader() -> None:
    """Inicia carga en segundo plano de modulos pesados."""
    global _HEAVY_LOADER_THREAD
    if _HEAVY_MODULES_READY.is_set():
        return
    with _HEAVY_LOAD_LOCK:
        if _HEAVY_MODULES_READY.is_set():
            return
        if _HEAVY_LOADER_THREAD is not None and _HEAVY_LOADER_THREAD.is_alive():
            return
        _HEAVY_LOADER_THREAD = threading.Thread(
            target=_load_heavy_modules,
            name="heavy-module-loader",
            daemon=True,
        )
        _HEAVY_LOADER_THREAD.start()

def ensure_heavy_modules_loaded() -> None:
    """Bloquea hasta completar imports pesados y valida error de carga."""
    if _HEAVY_MODULES_READY.is_set():
        if _HEAVY_MODULES_ERROR is not None:
            raise RuntimeError("No se pudieron cargar los modulos pesados.") from _HEAVY_MODULES_ERROR
        return
    if _HEAVY_LOADER_THREAD is not None and _HEAVY_LOADER_THREAD.is_alive():
        _HEAVY_MODULES_READY.wait()
    else:
        _load_heavy_modules()
    if _HEAVY_MODULES_ERROR is not None:
        raise RuntimeError("No se pudieron cargar los modulos pesados.") from _HEAVY_MODULES_ERROR

try:
    from rich.console import Console
    from rich.panel import Panel
except Exception:
    Console = None
    Panel = None

# Fuerza E/S en UTF-8 para soportar acentos y ñ en la terminal de Windows
def _configure_utf8_io():
    """Configura E/S para forzar UTF-8 en la consola."""
    try:
        os.environ.setdefault('PYTHONUTF8', '1')
        os.environ.setdefault('PYTHONIOENCODING', 'utf-8')
        for stream_name in ('stdout', 'stderr', 'stdin'):
            stream = getattr(sys, stream_name, None)
            if stream is not None and hasattr(stream, 'reconfigure'):
                stream.reconfigure(encoding='utf-8')
    except Exception:
        # Si la consola no soporta reconfigure, seguimos sin romper ejecucion
        pass


COLOR_BLUE = '\033[94m'
COLOR_YELLOW = '\033[93m'
COLOR_GREEN = '\033[92m'
COLOR_RED = '\033[91m'
COLOR_RESET = '\033[0m'
COLOR_QUESTION = '\033[38;5;37m'
LOADING_PREFIX_COLOR = '\033[96m'
LOADING_SPINNER_FRAMES = ('-', '\\', '|', '/')
LOADING_SPINNER = cycle(LOADING_SPINNER_FRAMES)
MIN_READABLE_LUMINANCE = 170.0
TERMINAL_BRAND_COLOR_SEQUENCE = (
    (31, 119, 180),   # Azul
    (214, 39, 40),    # Rojo
    (44, 160, 44),    # Verde
    (255, 127, 14),   # Naranja
    (148, 103, 189),  # Morado
    (23, 190, 207),   # Cian
    (140, 86, 75),    # Marron
    (227, 119, 194),  # Rosa
    (188, 189, 34),   # Oliva
    (63, 81, 181),    # Indigo
    (0, 173, 181),    # Turquesa
    (255, 64, 129),   # Magenta
    (57, 59, 121),    # Azul oscuro
    (82, 84, 163),    # Azul violeta
    (107, 110, 207),  # Lavanda intensa
    (99, 121, 57),    # Verde musgo
    (140, 162, 82),   # Verde oliva claro
    (140, 109, 49),   # Cafe
    (189, 158, 57),   # Mostaza
    (132, 60, 57),    # Rojo ladrillo
    (173, 73, 74),    # Coral oscuro
    (123, 65, 115),   # Purpura profundo
    (165, 81, 148),   # Fucsia oscuro
    (230, 85, 13),    # Naranja quemado
)
PREVIEW_COPY_SUFFIX_RE = re.compile(r'^(?P<brand>.+?)(?P<suffix>(?:[\s_-]+(?:copia|copy|preview))*)$', re.IGNORECASE)
PPTX_PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
PPTX_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PPTX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PPTX_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PPTX_SECTION_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
PPTX_SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
HEADER_COLOR_PRIMARY = '#286B72'
HEADER_COLOR_SECONDARY = '#3EBBC7'
HEADER_FONT_COLOR = '#FFFFFF'
HEADER_TOTAL_FILL = '#E8F0FF'
HEADER_FIRST_COL_FILL = '#F5F5F5'
TABLE_TEXT_PRIMARY = '#1F1F1F'
TABLE_POSITIVE_COLOR = '#27AE60'
TABLE_NEGATIVE_COLOR = '#C0392B'
TABLE_GRID_COLOR = '#1A1A1A'
APORTE_NEGATIVE_COLOR = "#F3948B"
APORTE_NEUTRAL_COLOR = "#FFCE7E"
VOLUME_BAR_START = '#1452B8'
VOLUME_BAR_END = '#8CB8FF'
VOLUME_BAR_ALPHA = 0.85
VOLUME_BAR_BASE_ALPHA = 0.18
SHARE_BAR_TRACK_COLOR = '#F3F4F6'
SHARE_BAR_TRACK_EDGE = '#D2D6DC'
SHARE_BAR_TOP_EDGE = '#667085'
SHARE_BAR_TOP_COUNT = 5
SHARE_BAR_HEADROOM = 1.10
SHARE_RANK_COLORS = {
    1: '#E3A008',  # Oro
    2: '#8B9098',  # Plata
    3: '#A85412',  # Bronce
    4: '#56616F',
    5: '#737B86',
}
SHARE_RANK_MOVEMENT_COLORS = {
    'up': '#1E8449',
    'down': '#C0392B',
    'same': '#7F8C8D',
}
SHARE_RANK_MOVEMENT_SYMBOLS = {
    'up': '↑',
    'down': '↓',
    'same': '-',
}
LINE_COLOR_CLIENT = '#2C3E50'
LINE_COLOR_NUMERATOR = "#D4AC0D"
BAR_COLOR_VARIATION_CLIENT = '#7F8C8D'
BAR_COLOR_VARIATION_NUMERATOR = '#F1C40F'
BAR_EDGE_COLOR = 'black'
BAR_LABEL_COLOR_POS = '#1E8449'
BAR_LABEL_COLOR_NEG = '#8B0000'
BAR_LABEL_COLOR_POS_ALT = '#27AE60'
BAR_LABEL_COLOR_NEG_ALT = '#C0392B'
REFERENCE_SERIES_COLOR = '#6E6E6E'
ANNOTATION_BOX_FACE = '#F2F2F2'
ANNOTATION_BOX_EDGE = 'black'
TREND_COLOR_PALETTE = {
    'trend_01': '#1F77B4',
    'trend_02': '#FF7F0E',
    'trend_03': '#2CA02C',
    'trend_04': '#D62728',
    'trend_05': '#9467BD',
    'trend_06': '#8C564B',
    'trend_07': '#E377C2',
    'trend_08': '#7F7F7F',
    'trend_09': '#BCBD22',
    'trend_10': '#17BECF',
    'trend_11': '#AEC7E8',
    'trend_12': '#FFBB78',
    'trend_13': '#98DF8A',
    'trend_14': '#FF9896',
    'trend_15': '#C5B0D5',
    'trend_16': '#C49C94',
    'trend_17': '#F7B6D2',
    'trend_18': '#C7C7C7',
    'trend_19': '#DBDB8D',
    'trend_20': '#9EDAE5',
    'trend_21': '#393B79',
    'trend_22': '#5254A3',
    'trend_23': '#6B6ECF',
    'trend_24': '#9C9EDE',
    'trend_25': '#637939',
    'trend_26': '#8CA252',
    'trend_27': '#B5CF6B',
    'trend_28': '#CEDB9C',
    'trend_29': '#8C6D31',
    'trend_30': '#BD9E39',
    'trend_31': '#E7BA52',
    'trend_32': '#E7CB94',
    'trend_33': '#843C39',
    'trend_34': '#AD494A',
    'trend_35': '#D6616B',
    'trend_36': '#E7969C',
    'trend_37': '#7B4173',
    'trend_38': '#A55194',
    'trend_39': '#CE6DBD',
    'trend_40': '#DE9ED6',
    'trend_41': '#3182BD',
    'trend_42': '#6BAED6',
    'trend_43': '#9ECAE1',
    'trend_44': '#C6DBEF',
    'trend_45': '#E6550D',
    'trend_46': '#FD8D3C',
    'trend_47': '#FDAE6B',
    'trend_48': '#FDD0A2',
    'trend_49': '#31A354',
    'trend_50': '#74C476',
    'trend_51': '#A1D99B',
    'trend_52': '#C7E9C0',
    'trend_53': '#756BB1',
    'trend_54': '#9E9AC8',
    'trend_55': '#BCBDDC',
    'trend_56': '#DADAEB',
    'trend_57': '#636363',
    'trend_58': '#969696',
    'trend_59': '#BDBDBD',
    'trend_60': '#D9D9D9',
    'trend_61': '#393E46',
    'trend_62': '#00ADB5',
    'trend_63': '#FF5722',
    'trend_64': '#795548',
    'trend_65': '#607D8B',
    'trend_66': '#8BC34A',
    'trend_67': '#CDDC39',
    'trend_68': '#FFC107',
    'trend_69': '#FF4081',
    'trend_70': '#3F51B5'
}
TREND_COLOR_SEQUENCE = list(TREND_COLOR_PALETTE.values())
# Paleta reservada para títulos de Competencia: usa los últimos colores para no interferir con los de marcas.
COMPETITION_TITLE_PALETTE = list(reversed(TREND_COLOR_SEQUENCE[-8:]))
# Paleta separada para títulos de Categoría (segmento 8). Usa un tramo distinto
# al inicio de la paleta para evitar colisión con colores de marcas.
CATEGORY_TITLE_PALETTE = TREND_COLOR_SEQUENCE[8:16] if len(TREND_COLOR_SEQUENCE) >= 16 else TREND_COLOR_SEQUENCE
TABLE_WRAP_WIDTH = 14
DISPLAY_TREND_REFERENCE_TEXT = False
TREND_SCALE_RULES = [
    {
        "threshold": 1_000_000_000_000,
        "divisor": 1_000_000_000_000,
        "suffix": "T",
        "reference": {
            'E': 'Valores expresados en billones',
            'P': 'Valores expressos em trilhões',
            'default': 'Values expressed in trillions'
        }
    },
    {
        "threshold": 1_000_000_000,
        "divisor": 1_000_000_000,
        "suffix": "B",
        "reference": {
            'E': 'Valores expresados en miles de millones',
            'P': 'Valores expressos em bilhões',
            'default': 'Values expressed in billions'
        }
    },
    {
        "threshold": 1_000_000,
        "divisor": 1_000_000,
        "suffix": "M",
        "reference": {
            'E': 'Valores expresados en millones',
            'P': 'Valores expressos em milhões',
            'default': 'Values expressed in millions'
        }
    },
    {
        "threshold": 1_000,
        "divisor": 1_000,
        "suffix": "K",
        "reference": {
            'E': 'Valores expresados en miles',
            'P': 'Valores expressos em milhares',
            'default': 'Values expressed in thousands'
        }
    },
    {
        "threshold": 0,
        "divisor": 1,
        "suffix": "",
        "reference": {
            'E': '',
            'P': '',
            'default': ''
        }
    },
]
MONTH_NAMES = {
    'P': ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
    'E': ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
}
def colorize(text: str, color: str = COLOR_BLUE) -> str:
    """Devuelve el texto envuelto con el codigo ANSI de color."""
    return f"{color}{text}{COLOR_RESET}"
def print_colored(text: str, color: str = COLOR_BLUE) -> None:
    """Imprime texto coloreado en la terminal."""
    print(colorize(text, color))

def print_loading_message(text: str, color: str = COLOR_QUESTION) -> None:
    """Imprime un prefijo animado y conserva el historial de mensajes."""
    frame = next(LOADING_SPINNER)
    prefix = f"{LOADING_PREFIX_COLOR}{frame}{COLOR_RESET}"
    print(f"{prefix} {colorize(text, color)}")

def print_loading_done(text: str = "Generacion completada", color: str = COLOR_GREEN) -> None:
    """Imprime una paloma de finalizacion para cerrar el loader."""
    checkmark = colorize("✓", COLOR_GREEN)
    print(f"{checkmark} {colorize(text, color)}")

def _format_elapsed_hms(seconds: float) -> str:
    """Convierte segundos a H:MM:SS.mmm para mensajes en terminal."""
    try:
        total = float(seconds)
    except Exception:
        return "-"
    total = max(0.0, total)
    h = int(total // 3600)
    m = int((total % 3600) // 60)
    s = total - (h * 3600) - (m * 60)
    return f"{h}:{m:02d}:{s:06.3f}"

def print_file_locked_error(path_str: str, *, elapsed_seconds: Optional[float] = None) -> None:
    """Muestra un mensaje estilo CoverageLab cuando el archivo esta en uso/bloqueado."""
    path_disp = str(path_str or "").strip() or "-"
    base = os.path.basename(path_disp) if path_disp not in {"-", ""} else "-"
    hora_actual = dt.now().strftime("%I:%M:%S %p")
    elapsed_line = ""
    if elapsed_seconds is not None:
        elapsed_line = f"\nTiempo total: {_format_elapsed_hms(elapsed_seconds)}"
    if Console is not None and Panel is not None:
        try:
            console = Console()
            msg = (
                "[bright_white]Proceso terminado con error[/bright_white]\n\n"
                f"[white]Archivo en uso: [bold]{base}[/bold][/white]\n"
                "[white]No se pudo reescribir porque esta abierto o bloqueado.[/white]\n"
                "[white]Cierra el archivo y vuelve a ejecutar.[/white]\n\n"
                f"[white]Hora de finalizacion: [bold]{hora_actual}[/bold][/white]"
                + (f"\n[white]Tiempo total: [bold]{_format_elapsed_hms(elapsed_seconds)}[/bold][/white]" if elapsed_seconds is not None else "")
                + f"\n\n[grey]{path_disp}[/grey]"
            )
            console.print()
            console.print(Panel.fit(msg, border_style="red", title="Coverages Latam"))
            console.print()
            return
        except Exception:
            pass
    print_colored("", COLOR_RED)
    print_colored("[Coverages Latam] Proceso terminado con error", COLOR_RED)
    print_colored(f"Archivo en uso: {base}", COLOR_RED)
    print_colored("No se pudo reescribir porque esta abierto o bloqueado.", COLOR_RED)
    print_colored("Cierra el archivo y vuelve a ejecutar.", COLOR_RED)
    print_colored(f"Hora de finalizacion: {hora_actual}{elapsed_line}", COLOR_RED)
    print_colored(path_disp, COLOR_RED)
    print_colored("", COLOR_RED)


def build_locked_retry_panel(
    path_str: str,
    *,
    action_label: str,
    elapsed_seconds: Optional[float] = None,
):
    """Construye el panel de pausa para archivo bloqueado."""
    path_disp = str(path_str or "").strip() or "-"
    base = os.path.basename(path_disp) if path_disp not in {"-", ""} else "-"
    hora_actual = dt.now().strftime("%I:%M:%S %p")
    msg = (
        "[bold yellow]Proceso en pausa por archivo en uso[/bold yellow]\n\n"
        f"[bold red]Archivo afectado:[/bold red] [white]{base}[/white]\n"
        f"[bold red]Accion pendiente:[/bold red] [white]{action_label}[/white]\n"
        "[red]No se puede continuar mientras el archivo siga abierto o bloqueado.[/red]\n"
        "[yellow]Cierra el archivo para habilitar el reintento.[/yellow]\n\n"
        f"[yellow]Hora actual:[/yellow] [white]{hora_actual}[/white]"
        + (f"\n[yellow]Tiempo acumulado:[/yellow] [white]{_format_elapsed_hms(elapsed_seconds)}[/white]" if elapsed_seconds is not None else "")
        + f"\n\n[grey70]{path_disp}[/grey70]"
    )
    return Panel.fit(
        msg,
        border_style="yellow",
        title="[bold yellow]Coverages Latam[/bold yellow]",
    )


def prompt_file_locked_retry(
    path_str: str,
    *,
    action_label: str,
    elapsed_seconds: Optional[float] = None,
) -> Optional[bool]:
    """Pausa el flujo para permitir cerrar el archivo y reintentar sin recomenzar."""
    path_disp = str(path_str or "").strip() or "-"
    base = os.path.basename(path_disp) if path_disp not in {"-", ""} else "-"
    if Console is not None and Panel is not None:
        try:
            console = Console()
            console.print()
            console.print(build_locked_retry_panel(path_str, action_label=action_label, elapsed_seconds=elapsed_seconds))
            console.print()
        except Exception:
            pass
    else:
        print_colored("", COLOR_YELLOW)
        print_colored("[Coverages Latam] Proceso en pausa por archivo en uso", COLOR_YELLOW)
        print_colored(f"Archivo afectado: {base}", COLOR_RED)
        print_colored(f"Accion pendiente: {action_label}", COLOR_RED)
        print_colored("No se puede continuar mientras el archivo siga abierto o bloqueado.", COLOR_RED)
        print_colored("Cierra el archivo para habilitar el reintento.", COLOR_YELLOW)
        if elapsed_seconds is not None:
            print_colored(f"Tiempo acumulado: {_format_elapsed_hms(elapsed_seconds)}", COLOR_YELLOW)
        print_colored(path_disp, COLOR_YELLOW)
        print_colored("", COLOR_YELLOW)
    try:
        choice = input(colorize("Reintentar [Enter] / Cancelar [Q]: ", COLOR_QUESTION)).strip().lower()
    except EOFError:
        return None
    if choice in {"q", "quit", "salir", "cancelar", "n", "no"}:
        return None
    return True


def save_presentation_with_retry(
    ppt,
    output_path: Path,
    lang: str,
    *,
    elapsed_seconds: Optional[float] = None,
) -> Optional[bool]:
    """Guarda la presentacion y escribe secciones, reintentando si el archivo esta bloqueado."""
    while True:
        if is_locked_for_write(output_path):
            retry_choice = prompt_file_locked_retry(
                str(output_path),
                action_label="guardar y reemplazar la presentacion",
                elapsed_seconds=elapsed_seconds,
            )
            if retry_choice is True:
                continue
            return None
        if output_path.exists():
            try:
                output_path.unlink()
            except PermissionError:
                retry_choice = prompt_file_locked_retry(
                    str(output_path),
                    action_label="reemplazar la presentacion existente",
                    elapsed_seconds=elapsed_seconds,
                )
                if retry_choice is True:
                    continue
                return None
        try:
            ppt.save(str(output_path))
            break
        except PermissionError:
            retry_choice = prompt_file_locked_retry(
                str(output_path),
                action_label="guardar la presentacion",
                elapsed_seconds=elapsed_seconds,
            )
            if retry_choice is True:
                continue
            return None

    while True:
        try:
            section_slide_map = build_section_slides_from_presentation(output_path, lang)
            apply_powerpoint_sections(output_path, section_slide_map)
            return True
        except PermissionError:
            retry_choice = prompt_file_locked_retry(
                str(output_path),
                action_label="escribir las secciones de PowerPoint",
                elapsed_seconds=elapsed_seconds,
            )
            if retry_choice is True:
                continue
            return None


EXCEL_TABLE_START_OFFSET_COLUMNS = 2
EXCEL_TABLE_GAP_COLUMNS = 1
EXCEL_CHART_HORIZONTAL_GAP_COLUMNS = 2
EXCEL_CHART_SECONDARY_SHIFT_COLUMNS = 1
EXCEL_CHART_WIDTH_SCALE = 4.0
EXCEL_CHART_HEIGHT_SCALE = 4.0
EXCEL_CHART_PAIR_ROW_OFFSET = 32
EXCEL_CHART_PAIR_OVERLAP_RATIO = 0.48
EXCEL_TABLE_HEADER_HEIGHT = 34.5
EXCEL_TABLE_BODY_HEIGHT = 15.75
EXCEL_TABLE_NUMBER_FORMAT = "0.00%"
EXCEL_TABLE_POSITIVE_FILL = "FFC6EFCE"
EXCEL_TABLE_POSITIVE_FONT = "FF006100"
EXCEL_TABLE_NEGATIVE_FILL = "FFFFC7CE"
EXCEL_TABLE_NEGATIVE_FONT = "FF9C0006"
EXCEL_TABLE_NEUTRAL_FONT = "FF000000"
EXCEL_TABLE_BORDER = Border(
    left=Side(style="medium", color="FF000000"),
    right=Side(style="medium", color="FF000000"),
    top=Side(style="medium", color="FF000000"),
    bottom=Side(style="medium", color="FF000000"),
)
EXCEL_NO_FILL = PatternFill(fill_type=None)


@lru_cache(maxsize=8)
def _excel_alignment_style(horizontal: str = "center") -> Alignment:
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=True)


@lru_cache(maxsize=64)
def _excel_font_style(color: str = EXCEL_TABLE_NEUTRAL_FONT, bold: bool = False) -> Font:
    return Font(name="Arial", bold=bool(bold), color=str(color))


@lru_cache(maxsize=512)
def _excel_solid_fill_style(argb_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=str(argb_color))


def _hex_to_excel_argb(color_value: Optional[str], fallback: str = "FFFFFFFF") -> str:
    if not color_value:
        return fallback
    color_text = str(color_value).strip()
    if not color_text:
        return fallback
    if color_text.startswith("#"):
        color_text = color_text[1:]
    if len(color_text) == 6:
        return f"FF{color_text.upper()}"
    if len(color_text) == 8:
        return color_text.upper()
    return fallback


def _rgb_float_to_excel_argb(rgb_tuple) -> str:
    rgb = tuple(max(0, min(255, int(round(channel * 255)))) for channel in rgb_tuple[:3])
    return f"FF{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _excel_fill_style(color_value: Optional[str]) -> PatternFill:
    if not color_value:
        return EXCEL_NO_FILL
    return _excel_solid_fill_style(_hex_to_excel_argb(color_value))


def _excel_fill_style_from_argb(argb_color: Optional[str]) -> PatternFill:
    if not argb_color:
        return EXCEL_NO_FILL
    return _excel_solid_fill_style(str(argb_color).upper())


def _relative_luminance(rgb_tuple) -> float:
    return 0.2126 * rgb_tuple[0] + 0.7152 * rgb_tuple[1] + 0.0722 * rgb_tuple[2]


def _excel_font_color_for_fill(color_value: Optional[str], fallback_dark: str = EXCEL_TABLE_NEUTRAL_FONT) -> str:
    try:
        rgb = mcolors.to_rgb(color_value)
    except Exception:
        return fallback_dark
    return "FFFFFFFF" if _relative_luminance(rgb) < 0.6 else fallback_dark


def _set_excel_fill(cell, color_value: Optional[str]) -> None:
    cell.fill = _excel_fill_style(color_value)


def _estimate_excel_chart_row_span(height_emu) -> int:
    try:
        height_in = emu_to_inches(height_emu) if height_emu is not None else emu_to_inches(Cm(10))
    except Exception:
        height_in = 4.0
    return max(20, int(math.ceil(height_in * 6 * EXCEL_CHART_HEIGHT_SCALE)) + 3)


def _estimate_excel_chart_col_span(width_emu) -> int:
    try:
        width_in = emu_to_inches(width_emu) if width_emu is not None else 7.5
    except Exception:
        width_in = 7.5
    return max(8, int(math.ceil(width_in * 1.55 * EXCEL_CHART_WIDTH_SCALE)))


def _resolve_trend_source_columns(block: dict, source_labels: list[str]) -> list[int]:
    available_entries = list(_build_excel_block_column_entries(block))
    resolved_cols: list[int] = []
    for label in source_labels:
        match = _best_semantic_reference_match(
            label,
            [entry_label for _, entry_label in available_entries],
            allow_family=True,
        )
        if match is None:
            continue
        selected_col_idx, _ = available_entries.pop(match.reference_index)
        resolved_cols.append(selected_col_idx)
    return resolved_cols


def _quoted_sheet_ref(ws, col_idx: int, row_start: int, row_end: int) -> str:
    return f"'{ws.title}'!${get_column_letter(col_idx)}${row_start}:${get_column_letter(col_idx)}${row_end}"


def _register_excel_column_width(sheet_runtime: dict, col_idx: int, target_width: float) -> None:
    pending_widths = sheet_runtime.setdefault("pending_widths", {})
    pending_widths[col_idx] = max(float(target_width), float(pending_widths.get(col_idx, 0.0) or 0.0))


def _apply_pending_excel_column_widths(ws, sheet_runtime: dict) -> None:
    pending_widths = sheet_runtime.get("pending_widths") or {}
    for col_idx, target_width in pending_widths.items():
        column_letter = get_column_letter(int(col_idx))
        current_width = ws.column_dimensions[column_letter].width or 0
        ws.column_dimensions[column_letter].width = max(current_width, float(target_width))


def _next_excel_helper_column(ws, sheet_runtime: dict) -> int:
    current_cursor = int(sheet_runtime.get("helper_col_cursor") or 0)
    if current_cursor <= ws.max_column:
        current_cursor = ws.max_column + 1
    sheet_runtime["helper_col_cursor"] = current_cursor + 1
    return current_cursor


def _write_chart_helper_column(
    ws,
    sheet_runtime: dict,
    row_start: int,
    row_end: int,
    *,
    values: Optional[list] = None,
    formulas: Optional[list[str]] = None,
    number_format: Optional[str] = None,
    cache_key: Optional[tuple] = None,
) -> str:
    if row_end < row_start:
        raise ValueError("Invalid helper column range")
    items = values if values is not None else formulas
    if items is None:
        raise ValueError("Helper column requires values or formulas")
    if len(items) != (row_end - row_start + 1):
        raise ValueError("Helper column length does not match row range")
    helper_ref_cache = sheet_runtime.setdefault("helper_ref_cache", {})
    if cache_key is not None:
        cached_ref = helper_ref_cache.get(cache_key)
        if cached_ref:
            return cached_ref
    helper_col = _next_excel_helper_column(ws, sheet_runtime)
    for offset, row_idx in enumerate(range(row_start, row_end + 1)):
        cell = ws.cell(row_idx, helper_col)
        cell.value = items[offset]
        if number_format:
            cell.number_format = number_format
    ws.column_dimensions[get_column_letter(helper_col)].hidden = True
    helper_ref = _quoted_sheet_ref(ws, helper_col, row_start, row_end)
    if cache_key is not None:
        helper_ref_cache[cache_key] = helper_ref
    return helper_ref


def _build_native_trend_chart(ws, chart_spec: dict, blocks: dict, sheet_runtime: dict) -> Optional[LineChart]:
    chart_title = chart_spec.get("title") or "Tendencia"
    chart = LineChart()
    chart.title = chart_title
    chart.style = 2
    chart_height_emu = chart_spec.get("height_emu")
    chart_width_emu = chart_spec.get("width_emu")
    chart.height = (emu_to_inches(chart_height_emu) if chart_height_emu is not None else 7.2) * EXCEL_CHART_HEIGHT_SCALE
    chart.width = (emu_to_inches(chart_width_emu) if chart_width_emu is not None else 15) * EXCEL_CHART_WIDTH_SCALE
    chart.layout = Layout(
        manualLayout=ManualLayout(
            layoutTarget="inner",
            xMode="factor",
            yMode="factor",
            wMode="factor",
            hMode="factor",
            x=0.02,
            y=0.02,
            w=0.96,
            h=0.72,
        )
    )
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.visible_cells_only = False
    chart.y_axis.majorGridlines = None
    chart.x_axis.title = ""
    chart.y_axis.title = ""
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.crosses = "min"
    chart.x_axis.majorTickMark = "out"
    chart.x_axis.minorTickMark = "none"
    chart.y_axis.axPos = "l"
    chart.y_axis.crosses = "min"
    chart.y_axis.majorTickMark = "out"
    chart.y_axis.minorTickMark = "none"
    prepared_series = []
    for series_spec in chart_spec.get("series", []):
        block = blocks.get(series_spec.get("block_key"))
        if block is None:
            continue
        source_cols = _resolve_trend_source_columns(block, series_spec.get("source_labels", []))
        if not source_cols:
            continue
        start_idx = max(0, int(series_spec.get("start_idx", chart_spec.get("start_idx", 0)) or 0))
        data_first_row = int(block["first_data_row"])
        data_last_row = int(block["last_data_row"])
        cat_first_row = data_first_row + start_idx
        val_first_row = data_first_row if series_spec.get("shift_from_start") and start_idx > 0 else data_first_row + start_idx
        val_last_row = data_last_row - start_idx if series_spec.get("shift_from_start") and start_idx > 0 else data_last_row
        if cat_first_row > data_last_row or val_first_row > val_last_row:
            continue
        row_count = min(data_last_row - cat_first_row + 1, val_last_row - val_first_row + 1)
        if row_count <= 0:
            continue
        cat_last_row = cat_first_row + row_count - 1
        val_last_row = val_first_row + row_count - 1
        raw_category_values = [
            ws.cell(row_idx, int(block["date_col"])).value
            for row_idx in range(cat_first_row, cat_last_row + 1)
        ]
        parsed_category_values = [_try_parse_reference_date(value) for value in raw_category_values]
        prepared_series.append(
            {
                "series_spec": series_spec,
                "cat_first_row": cat_first_row,
                "cat_last_row": cat_last_row,
                "val_first_row": val_first_row,
                "val_last_row": val_last_row,
                "source_cols": source_cols,
                "raw_category_values": raw_category_values,
                "parsed_category_values": parsed_category_values,
            }
        )
    if not prepared_series:
        return None
    use_date_axis = all(all(value is not None for value in item["parsed_category_values"]) for item in prepared_series)
    if use_date_axis:
        chart.x_axis = DateAxis(crossAx=chart.y_axis.axId)
        chart.y_axis.crossAx = chart.x_axis.axId
        chart.x_axis.delete = False
        chart.x_axis.axPos = "b"
        chart.x_axis.number_format = "mmm-yy"
        chart.x_axis.baseTimeUnit = "months"
        chart.x_axis.majorTimeUnit = "months"
        chart.x_axis.majorUnit = 1
        chart.x_axis.tickLblPos = "nextTo"
        chart.x_axis.crosses = "min"
        chart.x_axis.majorTickMark = "out"
        chart.x_axis.minorTickMark = "none"
        chart.x_axis.lblOffset = 100
        chart.x_axis.title = ""
        chart.y_axis.delete = False
    added_series_specs = []
    for prepared in prepared_series:
        series_spec = prepared["series_spec"]
        block = blocks.get(series_spec.get("block_key"))
        if block is None:
            continue
        if use_date_axis:
            category_values = [
                value.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                for value in prepared["parsed_category_values"]
            ]
            cache_key = ("category", "date", tuple(category_values), "mmm-yy")
            cat_ref = _write_chart_helper_column(
                ws,
                sheet_runtime,
                prepared["cat_first_row"],
                prepared["cat_last_row"],
                values=category_values,
                number_format="mmm-yy",
                cache_key=cache_key,
            )
            cat_source = AxDataSource(numRef=NumRef(f=cat_ref))
        else:
            category_labels = [_safe_excel_period_label(value) for value in prepared["raw_category_values"]]
            cache_key = ("category", "label", tuple(category_labels), None)
            cat_ref = _write_chart_helper_column(
                ws,
                sheet_runtime,
                prepared["cat_first_row"],
                prepared["cat_last_row"],
                values=category_labels,
                cache_key=cache_key,
            )
            cat_source = AxDataSource(strRef=StrRef(f=cat_ref))
        if len(prepared["source_cols"]) == 1:
            val_ref = _quoted_sheet_ref(
                ws,
                prepared["source_cols"][0],
                prepared["val_first_row"],
                prepared["val_last_row"],
            )
        else:
            formula_cache_key = (
                "sum",
                tuple(int(col_idx) for col_idx in prepared["source_cols"]),
                int(prepared["val_first_row"]),
                int(prepared["val_last_row"]),
            )
            cached_formula_ref = sheet_runtime.setdefault("helper_ref_cache", {}).get(formula_cache_key)
            if cached_formula_ref is None:
                helper_formulas = []
                for row_idx in range(prepared["val_first_row"], prepared["val_last_row"] + 1):
                    row_refs = [f"{get_column_letter(col_idx)}{row_idx}" for col_idx in prepared["source_cols"]]
                    helper_formulas.append("=" + "+".join(row_refs))
                val_ref = _write_chart_helper_column(
                    ws,
                    sheet_runtime,
                    prepared["val_first_row"],
                    prepared["val_last_row"],
                    formulas=helper_formulas,
                    cache_key=formula_cache_key,
                )
            else:
                val_ref = cached_formula_ref
        series_obj = Series(
            cat=cat_source,
            val=NumDataSource(numRef=NumRef(f=val_ref)),
        )
        series_obj.title = SeriesLabel(v=str(series_spec.get("label") or "Serie"))
        chart.ser.append(series_obj)
        added_series_specs.append(series_spec)
    for series_obj, series_spec in zip(chart.ser, added_series_specs):
        color_value = series_spec.get("color")
        if color_value:
            series_obj.graphicalProperties.line.solidFill = _hex_to_excel_argb(color_value)[2:]
            series_obj.graphicalProperties.solidFill = _hex_to_excel_argb(color_value)[2:]
            series_obj.marker.graphicalProperties.line.solidFill = _hex_to_excel_argb(color_value)[2:]
            series_obj.marker.graphicalProperties.solidFill = _hex_to_excel_argb(color_value)[2:]
        series_obj.marker.symbol = "circle"
        series_obj.marker.size = 6
        if series_spec.get("dash"):
            series_obj.graphicalProperties.line.dashStyle = "sysDash"
    return chart


def _add_native_trend_charts_to_sheet(ws, trend_charts: list[dict], blocks: dict, sheet_runtime: dict, start_row: int, start_col: int = 1) -> None:
    if not trend_charts:
        return
    current_row = max(1, start_row)
    for chart_spec in trend_charts:
        chart = _build_native_trend_chart(ws, chart_spec, blocks, sheet_runtime)
        if chart is None:
            continue
        anchor_row = int(chart_spec.get("anchor_row") or current_row)
        anchor_col = int(chart_spec.get("anchor_col") or start_col)
        anchor_cell = f"{get_column_letter(anchor_col)}{anchor_row}"
        ws.add_chart(chart, anchor_cell)
        if "anchor_row" not in chart_spec:
            current_row += _estimate_excel_chart_row_span(chart_spec.get("height_emu"))


def save_workbook_with_retry(workbook, output_path: Path) -> bool:
    """Guarda un workbook evitando archivos temporales parciales."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=str(output_path.parent)) as tmp_file:
            tmp_path = Path(tmp_file.name)
        workbook.save(tmp_path)
        if output_path.exists():
            output_path.unlink()
        os.replace(tmp_path, output_path)
        return True
    except PermissionError:
        return False
    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink(missing_ok=True)
            except PermissionError:
                pass


def _find_excel_header_row(ws) -> Optional[int]:
    compras_row = None
    table_row = None
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell_value = ws.cell(row_idx, col_idx).value
            if isinstance(cell_value, str) and 'compras' in _normalize_header_text(cell_value):
                compras_row = row_idx
                break
            if isinstance(cell_value, str) and 'table' in _normalize_header_text(cell_value) and table_row is None:
                table_row = row_idx
        if compras_row is not None:
            break
    if compras_row is not None:
        return compras_row
    return table_row


def _find_excel_last_data_row(ws, date_col: int, start_row: int) -> Optional[int]:
    last_row = None
    for row_idx in range(start_row, ws.max_row + 1):
        if _try_parse_reference_date(ws.cell(row_idx, date_col).value) is not None:
            last_row = row_idx
    return last_row


def _extract_excel_block_columns(ws, header_row: int, start_col: int, end_col: int) -> list[tuple[int, str]]:
    columns: list[tuple[int, str]] = []
    for col_idx in range(start_col, end_col + 1):
        header_value = ws.cell(header_row, col_idx).value
        if _is_separator_column(header_value):
            continue
        header_text = str(header_value).strip() if header_value is not None else ''
        if not header_text:
            continue
        columns.append((col_idx, header_text))
    return columns


def detect_excel_table_blocks(ws) -> dict[str, dict]:
    """Detecta los bloques fuente de Compras/Ventas directamente en la hoja."""
    header_row = _find_excel_header_row(ws)
    if header_row is None:
        return {}
    ventas_header_col = None
    fallback_separator_col = None
    for col_idx in range(2, ws.max_column + 1):
        header_value = ws.cell(header_row, col_idx).value
        normalized = _normalize_header_text(header_value) if header_value is not None else ''
        if 'ventas' in normalized:
            ventas_header_col = col_idx
            break
        if fallback_separator_col is None and _is_separator_column(header_value):
            fallback_separator_col = col_idx
    if ventas_header_col is None and fallback_separator_col is not None:
        for candidate_col in range(fallback_separator_col + 1, ws.max_column + 1):
            if not _is_separator_column(ws.cell(header_row, candidate_col).value):
                ventas_header_col = candidate_col
                break

    blocks: dict[str, dict] = {}
    first_data_row = header_row + 1
    if ventas_header_col is not None:
        compras_data_cols = _extract_excel_block_columns(ws, header_row, 2, max(1, ventas_header_col - 1))
        if compras_data_cols:
            last_row = _find_excel_last_data_row(ws, 1, first_data_row)
            if last_row is not None:
                blocks["compras"] = {
                    "display_tipo": "Compras",
                    "header_row": header_row,
                    "date_col": 1,
                    "data_cols": compras_data_cols,
                    "first_data_row": first_data_row,
                    "last_data_row": last_row,
                    "last_used_col": compras_data_cols[-1][0],
                }
        ventas_data_cols = _extract_excel_block_columns(ws, header_row, ventas_header_col + 1, ws.max_column)
        if ventas_data_cols:
            last_row = _find_excel_last_data_row(ws, ventas_header_col, first_data_row)
            if last_row is not None:
                blocks["ventas"] = {
                    "display_tipo": "Ventas",
                    "header_row": header_row,
                    "date_col": ventas_header_col,
                    "data_cols": ventas_data_cols,
                    "first_data_row": first_data_row,
                    "last_data_row": last_row,
                    "last_used_col": ventas_data_cols[-1][0],
                    "pipeline": _extract_pipeline(ws.cell(header_row, ventas_header_col).value),
                }
    else:
        compras_data_cols = _extract_excel_block_columns(ws, header_row, 2, ws.max_column)
        if compras_data_cols:
            last_row = _find_excel_last_data_row(ws, 1, first_data_row)
            if last_row is not None:
                blocks["compras"] = {
                    "display_tipo": "Compras",
                    "header_row": header_row,
                    "date_col": 1,
                    "data_cols": compras_data_cols,
                    "first_data_row": first_data_row,
                    "last_data_row": last_row,
                    "last_used_col": compras_data_cols[-1][0],
                }
    return blocks


def _build_excel_block_column_entries(block: dict) -> list[tuple[int, str]]:
    cached_entries = block.get("_column_entries")
    if cached_entries is not None:
        return cached_entries
    entries: list[tuple[int, str]] = []
    for col_idx, label in block.get("data_cols", []):
        label_text = str(label).strip()
        if not label_text:
            continue
        entries.append((col_idx, label_text))
    block["_column_entries"] = entries
    return entries


def _resolve_excel_source_columns(block: dict, apo_df) -> list[tuple[str, list[int]]]:
    available_entries = list(_build_excel_block_column_entries(block))
    grouped_source_labels = {}
    if hasattr(apo_df, "attrs"):
        grouped_source_labels = apo_df.attrs.get("grouped_source_labels", {}) or {}
    resolved: list[tuple[str, list[int]]] = []
    for column_name in list(apo_df.columns)[1:]:
        column_label = str(column_name).strip()
        if not column_label:
            continue
        if column_label.lower() == "total":
            resolved.append((column_label, []))
            continue
        source_cols: list[int] = []
        source_labels = grouped_source_labels.get(column_label) or [column_label]
        for source_label in source_labels:
            match = _best_semantic_reference_match(
                source_label,
                [entry_label for _, entry_label in available_entries],
                allow_family=True,
            )
            if match is None:
                continue
            matched_col_idx, _ = available_entries.pop(match.reference_index)
            source_cols.append(matched_col_idx)
        resolved.append((column_label, source_cols))
    return resolved


def _safe_excel_period_label(value) -> str:
    parsed = _try_parse_reference_date(value)
    if parsed is None:
        return str(value).strip() if value is not None else ""
    return parsed.strftime("%b-%y")


def _excel_range_sum_formula(col_start: int, col_end: int, row_start: int, row_end: int) -> str:
    return f"SUM({get_column_letter(col_start)}{row_start}:{get_column_letter(col_end)}{row_end})"


def _excel_column_sum_formula(col_idx: int, row_start: int, row_end: int) -> str:
    return f"SUM({get_column_letter(col_idx)}{row_start}:{get_column_letter(col_idx)}{row_end})"


def _excel_multi_column_sum_formula(col_indexes: list[int], row_start: int, row_end: int) -> str:
    valid_indexes = []
    for col_idx in col_indexes:
        try:
            normalized_idx = int(col_idx)
        except (TypeError, ValueError):
            continue
        if normalized_idx > 0:
            valid_indexes.append(normalized_idx)
    if not valid_indexes:
        return "0"
    if len(valid_indexes) == 1:
        return _excel_column_sum_formula(valid_indexes[0], row_start, row_end)
    return "+".join(
        _excel_column_sum_formula(col_idx, row_start, row_end)
        for col_idx in valid_indexes
    )


def _apply_excel_cell_style(
    cell,
    *,
    alignment: str = "center",
    font: Optional[Font] = None,
    fill: Optional[PatternFill] = None,
) -> None:
    cell.border = EXCEL_TABLE_BORDER
    cell.alignment = _excel_alignment_style(alignment)
    cell.font = font or _excel_font_style(EXCEL_TABLE_NEUTRAL_FONT, False)
    cell.fill = fill if fill is not None else EXCEL_NO_FILL


def _apply_excel_table_style(cell, *, is_header: bool = False, alignment: str = "center") -> None:
    _apply_excel_cell_style(
        cell,
        alignment=alignment,
        font=_excel_font_style(EXCEL_TABLE_NEUTRAL_FONT, is_header),
        fill=EXCEL_NO_FILL,
    )


def _apply_excel_metric_fill(cell, numeric_value: Optional[float]) -> None:
    if numeric_value is None or not np.isfinite(numeric_value):
        cell.fill = EXCEL_NO_FILL
        cell.font = _excel_font_style(EXCEL_TABLE_NEUTRAL_FONT, False)
        return
    if numeric_value > 0:
        cell.fill = _excel_fill_style_from_argb(EXCEL_TABLE_POSITIVE_FILL)
        cell.font = _excel_font_style(EXCEL_TABLE_POSITIVE_FONT, False)
        return
    if numeric_value < 0:
        cell.fill = _excel_fill_style_from_argb(EXCEL_TABLE_NEGATIVE_FILL)
        cell.font = _excel_font_style(EXCEL_TABLE_NEGATIVE_FONT, False)
        return
    cell.fill = EXCEL_NO_FILL
    cell.font = _excel_font_style(EXCEL_TABLE_NEUTRAL_FONT, False)


def _prepare_excel_formula_table_visuals(
    apo_df,
    resolved_columns: list[tuple[str, Optional[int]]],
    column_color_mapping: Optional[dict[str, str]] = None,
) -> dict:
    total_column_name = next(
        (col_name for col_name, _ in resolved_columns if str(col_name).strip().lower() == "total"),
        None
    )
    data_columns = [
        idx
        for idx, (col_name, _) in enumerate(resolved_columns, start=1)
        if total_column_name is None or str(col_name).strip() != str(total_column_name).strip()
    ]
    column_color_mapping = column_color_mapping or {}
    exact_color_mapping = {
        str(key).strip(): value
        for key, value in column_color_mapping.items()
        if key is not None and value
    }
    normalized_color_mapping = {
        str(key).strip().lower(): value
        for key, value in exact_color_mapping.items()
        if str(key).strip()
    }
    suffix_variants = ('.c', '.v', '_c', '_v', '-c', '-v')

    def resolve_column_color(column_name):
        if column_name is None:
            return None
        key = str(column_name).strip()
        if not key:
            return None
        color_value = exact_color_mapping.get(key)
        if color_value:
            return color_value
        lower_key = key.lower()
        color_value = normalized_color_mapping.get(lower_key)
        if color_value:
            return color_value
        for suffix in suffix_variants:
            if lower_key.endswith(suffix):
                base = key[:-len(suffix)].strip()
                if not base:
                    continue
                color_value = exact_color_mapping.get(base)
                if color_value:
                    return color_value
                color_value = normalized_color_mapping.get(base.lower())
                if color_value:
                    return color_value
        return None

    start_rgb = np.array(mcolors.to_rgb(VOLUME_BAR_START))
    end_rgb = np.array(mcolors.to_rgb(VOLUME_BAR_END))

    def share_colors(count: int):
        if count <= 0:
            return []
        if count == 1:
            return [mcolors.to_hex(np.clip(start_rgb, 0, 1))]
        colors = []
        for idx in range(count):
            ratio = idx / (count - 1)
            rgb = np.clip(start_rgb + (end_rgb - start_rgb) * ratio, 0, 1)
            colors.append(mcolors.to_hex(rgb))
        return colors

    def to_percentage(value):
        try:
            text_value = str(value).strip().replace('%', '').replace(',', '.')
            if not text_value:
                return None
            return max(0.0, float(text_value) / 100.0)
        except Exception:
            return None

    def to_signed_percentage(value):
        try:
            text_value = str(value).strip().replace('%', '').replace(',', '.')
            if not text_value:
                return None
            return float(text_value) / 100.0
        except Exception:
            return None

    fallback_column_colors = share_colors(len(data_columns))
    column_colors = {}
    for rel_idx, col_idx in enumerate(data_columns):
        col_name = resolved_columns[col_idx - 1][0]
        resolved = resolve_column_color(col_name)
        if resolved is None:
            resolved = fallback_column_colors[rel_idx] if rel_idx < len(fallback_column_colors) else VOLUME_BAR_END
        try:
            column_colors[col_idx] = mcolors.to_hex(resolved)
        except Exception:
            column_colors[col_idx] = fallback_column_colors[rel_idx] if rel_idx < len(fallback_column_colors) else VOLUME_BAR_END

    header_font = _excel_font_style(_hex_to_excel_argb(HEADER_FONT_COLOR), True)
    label_font = _excel_font_style(EXCEL_TABLE_NEUTRAL_FONT, True)
    neutral_font = _excel_font_style(EXCEL_TABLE_NEUTRAL_FONT, False)
    positive_font = _excel_font_style(_hex_to_excel_argb(TABLE_POSITIVE_COLOR), False)
    negative_font = _excel_font_style(_hex_to_excel_argb(TABLE_NEGATIVE_COLOR), False)
    white_font = _excel_font_style("FFFFFFFF", False)
    white_fill = _excel_fill_style("#FFFFFF")
    total_fill = _excel_fill_style(HEADER_TOTAL_FILL)

    header_styles: dict[int, tuple[PatternFill, Font]] = {}
    for rel_col_idx, (column_label, _) in enumerate(resolved_columns, start=1):
        if total_column_name is not None and str(column_label).strip().lower() == 'total':
            header_styles[rel_col_idx] = (_excel_fill_style(HEADER_COLOR_PRIMARY), header_font)
            continue
        if rel_col_idx in data_columns:
            header_color = resolve_column_color(column_label) or HEADER_COLOR_SECONDARY
            header_styles[rel_col_idx] = (
                _excel_fill_style(header_color),
                _excel_font_style(_excel_font_color_for_fill(header_color), True),
            )
        else:
            header_styles[rel_col_idx] = (_excel_fill_style(HEADER_COLOR_SECONDARY), header_font)

    body_styles: dict[tuple[int, int], tuple[PatternFill, Font]] = {}
    volume_row_indexes = [idx for idx, label in enumerate(apo_df.iloc[:, 0]) if str(label).strip().lower().startswith('vol')]
    aporte_row_index = next((idx for idx, label in enumerate(apo_df.iloc[:, 0]) if str(label).strip().lower() == 'aporte'), None)

    for row_offset, label in enumerate(apo_df.iloc[:, 0]):
        row_label = str(label).strip().lower()
        for rel_col_idx, (column_label, _) in enumerate(resolved_columns, start=1):
            is_total_col = total_column_name is not None and str(column_label).strip().lower() == 'total'
            fill_style = total_fill if is_total_col else white_fill
            if row_label.startswith('vol'):
                font_style = neutral_font
            else:
                numeric = _parse_percent_cell(apo_df.iloc[row_offset, rel_col_idx])
                if numeric is None:
                    font_style = neutral_font
                else:
                    font_style = positive_font if numeric >= 0 else negative_font
            body_styles[(row_offset, rel_col_idx)] = (fill_style, font_style)

    if aporte_row_index is not None:
        aporte_values = []
        for rel_col_idx in data_columns:
            numeric = to_signed_percentage(apo_df.iloc[aporte_row_index, rel_col_idx])
            if numeric is not None:
                aporte_values.append((rel_col_idx, numeric))
        if aporte_values:
            aporte_numbers = [value for _, value in aporte_values]
            min_val = min(aporte_numbers)
            max_val = max(aporte_numbers)
            cmap = None
            norm = None
            if not math.isclose(min_val, max_val, rel_tol=1e-9, abs_tol=1e-9):
                if min_val < 0 and max_val > 0:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_diverging_excel',
                        [APORTE_NEGATIVE_COLOR, APORTE_NEUTRAL_COLOR, TABLE_POSITIVE_COLOR]
                    )
                    norm = mcolors.TwoSlopeNorm(vmin=min_val, vcenter=0, vmax=max_val)
                elif max_val <= 0:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_negative_excel',
                        [APORTE_NEGATIVE_COLOR, APORTE_NEUTRAL_COLOR]
                    )
                    norm = mcolors.Normalize(vmin=min_val, vmax=0)
                else:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_positive_excel',
                        [APORTE_NEUTRAL_COLOR, TABLE_POSITIVE_COLOR]
                    )
                    norm = mcolors.Normalize(vmin=0, vmax=max_val)
            for rel_col_idx, numeric in aporte_values:
                if cmap is not None and norm is not None:
                    rgb = tuple(cmap(norm(numeric))[:3])
                else:
                    if math.isclose(numeric, 0.0, abs_tol=1e-9):
                        rgb = mcolors.to_rgb(APORTE_NEUTRAL_COLOR)
                    elif numeric > 0:
                        rgb = mcolors.to_rgb(TABLE_POSITIVE_COLOR)
                    else:
                        rgb = mcolors.to_rgb(APORTE_NEGATIVE_COLOR)
                fill_style = _excel_fill_style_from_argb(_rgb_float_to_excel_argb(rgb))
                font_style = neutral_font if _relative_luminance(rgb) > 0.5 else white_font
                body_styles[(aporte_row_index, rel_col_idx)] = (fill_style, font_style)

    white_rgb = np.array([1.0, 1.0, 1.0])
    for row_offset in volume_row_indexes:
        row_percentages = {}
        max_percent = 0.0
        for rel_col_idx in data_columns:
            percent = to_percentage(apo_df.iloc[row_offset, rel_col_idx])
            row_percentages[rel_col_idx] = percent
            if percent is not None and percent > max_percent:
                max_percent = percent
        for rel_col_idx in data_columns:
            percent = row_percentages.get(rel_col_idx)
            if percent is None:
                continue
            bar_color = column_colors.get(rel_col_idx, VOLUME_BAR_END)
            base_rgb = np.array(mcolors.to_rgb(bar_color))
            intensity = min(percent / max_percent, 1.0) if max_percent > 0 else 0.0
            blended_rgb = white_rgb * (1.0 - intensity) + base_rgb * intensity
            body_styles[(row_offset, rel_col_idx)] = (
                _excel_fill_style_from_argb(_rgb_float_to_excel_argb(blended_rgb)),
                neutral_font,
            )

    return {
        "header_styles": header_styles,
        "title_fill": _excel_fill_style(HEADER_COLOR_PRIMARY),
        "title_font": header_font,
        "label_fill": _excel_fill_style(HEADER_FIRST_COL_FILL),
        "label_font": label_font,
        "body_styles": body_styles,
    }


def _write_dynamic_formula_table_to_sheet(
    ws,
    block: dict,
    apo_df,
    start_row: int,
    start_col: int,
    sheet_runtime: dict,
    column_color_mapping: Optional[dict[str, str]] = None,
) -> int:
    """Inserta una tabla con fórmulas dinámicas al lado del bloque original."""
    if apo_df is None or apo_df.empty or apo_df.shape[1] <= 1 or not block.get("data_cols"):
        return 0
    try:
        pipeline = int(block.get("pipeline", 0))
    except (TypeError, ValueError):
        pipeline = 0
    last_data_row = int(block["last_data_row"])
    prev_start = last_data_row - 23 - pipeline
    prev_end = last_data_row - 12 - pipeline
    curr_start = last_data_row - 11 - pipeline
    curr_end = last_data_row - pipeline
    if (
        prev_start < block["first_data_row"]
        or prev_end < prev_start
        or curr_start < block["first_data_row"]
        or curr_end < curr_start
    ):
        return 0
    resolved_columns = _resolve_excel_source_columns(block, apo_df)
    if not resolved_columns:
        return 0
    visuals = _prepare_excel_formula_table_visuals(
        apo_df,
        resolved_columns,
        column_color_mapping=column_color_mapping,
    )
    first_source_col = block["data_cols"][0][0]
    last_source_col = block["data_cols"][-1][0]
    metric_labels = [
        f"Vol {_safe_excel_period_label(ws.cell(prev_end, block['date_col']).value)}",
        f"Vol {_safe_excel_period_label(ws.cell(curr_end, block['date_col']).value)}",
        "Var %",
        "Aporte",
    ]
    total_width = len(resolved_columns) + 1
    header_row = start_row
    body_start_row = start_row + 1
    ws.row_dimensions[header_row].height = EXCEL_TABLE_HEADER_HEIGHT
    for row_idx in range(body_start_row, body_start_row + 4):
        ws.row_dimensions[row_idx].height = EXCEL_TABLE_BODY_HEIGHT

    title_cell = ws.cell(header_row, start_col)
    title_cell.value = block.get("display_tipo", "")
    _apply_excel_cell_style(
        title_cell,
        alignment="center",
        font=visuals["title_font"],
        fill=visuals["title_fill"],
    )
    for header_offset, (column_label, _) in enumerate(resolved_columns, start=1):
        header_cell = ws.cell(header_row, start_col + header_offset)
        header_cell.value = column_label
        header_fill, header_font = visuals["header_styles"][header_offset]
        _apply_excel_cell_style(
            header_cell,
            alignment="center",
            font=header_font,
            fill=header_fill,
        )

    for row_offset, metric_label in enumerate(metric_labels):
        metric_cell = ws.cell(body_start_row + row_offset, start_col)
        metric_cell.value = metric_label
        _apply_excel_cell_style(
            metric_cell,
            alignment="left",
            font=visuals["label_font"],
            fill=visuals["label_fill"],
        )

    prev_total_formula = _excel_range_sum_formula(first_source_col, last_source_col, prev_start, prev_end)
    curr_total_formula = _excel_range_sum_formula(first_source_col, last_source_col, curr_start, curr_end)
    for col_offset, (column_label, source_cols) in enumerate(resolved_columns, start=1):
        target_col = start_col + col_offset
        cell_prev = ws.cell(body_start_row, target_col)
        cell_curr = ws.cell(body_start_row + 1, target_col)
        cell_var = ws.cell(body_start_row + 2, target_col)
        cell_aporte = ws.cell(body_start_row + 3, target_col)
        for row_offset, current_cell in enumerate((cell_prev, cell_curr, cell_var, cell_aporte)):
            fill_style, font_style = visuals["body_styles"][(row_offset, col_offset)]
            _apply_excel_cell_style(
                current_cell,
                alignment="center",
                font=font_style,
                fill=fill_style,
            )
            current_cell.number_format = EXCEL_TABLE_NUMBER_FORMAT
        if column_label.lower() == "total":
            prev_formula = f"=IFERROR({prev_total_formula}/{prev_total_formula},0)"
            curr_formula = f"=IFERROR({curr_total_formula}/{curr_total_formula},0)"
            var_formula = f"=IFERROR({curr_total_formula}/{prev_total_formula}-1,0)"
        elif not source_cols:
            prev_formula = "=0"
            curr_formula = "=0"
            var_formula = "=0"
        else:
            prev_source_formula = _excel_multi_column_sum_formula(source_cols, prev_start, prev_end)
            curr_source_formula = _excel_multi_column_sum_formula(source_cols, curr_start, curr_end)
            prev_formula = f"=IFERROR(({prev_source_formula})/{prev_total_formula},0)"
            curr_formula = f"=IFERROR(({curr_source_formula})/{curr_total_formula},0)"
            var_formula = f"=IFERROR(({curr_source_formula})/({prev_source_formula})-1,0)"
        aporte_formula = f"=IFERROR({get_column_letter(target_col)}{body_start_row + 2}*{get_column_letter(target_col)}{body_start_row},0)"
        cell_prev.value = prev_formula
        cell_curr.value = curr_formula
        cell_var.value = var_formula
        cell_aporte.value = aporte_formula

    for col_idx in range(start_col, start_col + total_width):
        if col_idx == start_col:
            _register_excel_column_width(sheet_runtime, col_idx, 13)
        else:
            header_text = str(ws.cell(header_row, col_idx).value or "")
            target_width = min(22, max(13, len(header_text) + 2))
            _register_excel_column_width(sheet_runtime, col_idx, target_width)
    return total_width


def generate_excel_output_with_tables(source_excel_path: Path, output_path: Path, sheet_exports: list[dict]) -> bool:
    """Copia el Excel fuente y agrega tablas dinámicas con fórmulas editables."""
    workbook = load_workbook(source_excel_path)
    sheet_runtime_cache: dict[str, dict] = {}
    try:
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
    except Exception:
        pass
    for sheet_export in sheet_exports:
        sheet_name = sheet_export.get("sheet_name")
        if not sheet_name or sheet_name not in workbook.sheetnames:
            continue
        ws = workbook[sheet_name]
        export_images = sheet_export.get("images", []) or []
        trend_charts = sheet_export.get("trend_charts", []) or []
        sheet_runtime = sheet_runtime_cache.get(sheet_name)
        if sheet_runtime is None:
            sheet_runtime = {
                "blocks": detect_excel_table_blocks(ws),
                "helper_ref_cache": {},
                "pending_widths": {},
            }
            sheet_runtime_cache[sheet_name] = sheet_runtime
        blocks = sheet_runtime["blocks"]
        table_anchor_row = None
        table_anchor_col = None
        table_anchor_map: dict[str, int] = {}
        export_tables = [
            table_entry
            for table_entry in sheet_export.get("tables", [])
            if isinstance(table_entry.get("apo"), pd.DataFrame)
            and not table_entry["apo"].empty
            and not table_entry["apo"].attrs.get("skip_table", False)
        ]
        last_block = blocks.get("ventas") or blocks.get("compras") if blocks else None
        if export_tables and last_block is not None:
            anchor_row = max(int(last_block["header_row"]) + 1, int(last_block["last_data_row"]) - 4)
            table_anchor_row = anchor_row
            anchor_col = int(last_block["last_used_col"]) + EXCEL_TABLE_START_OFFSET_COLUMNS
            table_anchor_col = anchor_col
            for idx, table_entry in enumerate(export_tables):
                display_tipo = str(table_entry.get("display_tipo", "")).strip().lower()
                if "venta" in display_tipo:
                    block = blocks.get("ventas")
                elif "compra" in display_tipo:
                    block = blocks.get("compras")
                else:
                    block = blocks.get("compras") or blocks.get("ventas")
                if block is None:
                    continue
                if idx > 0:
                    anchor_col += EXCEL_TABLE_GAP_COLUMNS
                anchor_key = ""
                if "venta" in display_tipo:
                    anchor_key = "ventas"
                elif "compra" in display_tipo:
                    anchor_key = "compras"
                if anchor_key and anchor_key not in table_anchor_map:
                    table_anchor_map[anchor_key] = anchor_col
                width_used = _write_dynamic_formula_table_to_sheet(
                    ws,
                    block,
                    table_entry["apo"],
                    anchor_row,
                    anchor_col,
                    sheet_runtime,
                    column_color_mapping=table_entry.get("column_color_mapping"),
                )
                if width_used:
                    anchor_col += width_used
        if trend_charts:
            if table_anchor_row is not None and last_block is not None:
                chart_start_row = max(int(last_block["header_row"]) + 6, int(table_anchor_row) - 16)
            elif last_block is not None:
                chart_start_row = max(int(last_block["last_data_row"]) + 3, 55)
            else:
                chart_start_row = max(ws.max_row + 3, 5)
            if table_anchor_col is not None:
                chart_start_col = int(table_anchor_col)
            elif last_block is not None:
                chart_start_col = int(last_block["last_used_col"]) + EXCEL_TABLE_START_OFFSET_COLUMNS
            else:
                chart_start_col = 1
            if len(trend_charts) == 1 and table_anchor_row is not None and last_block is not None:
                chart_start_row = max(int(last_block["header_row"]) + 4, int(table_anchor_row) - EXCEL_CHART_PAIR_ROW_OFFSET)
                trend_charts[0]["anchor_row"] = chart_start_row
                trend_charts[0]["anchor_col"] = int(table_anchor_map.get("compras", chart_start_col))
            if len(trend_charts) == 2:
                def _chart_primary_key(chart_spec: dict) -> str:
                    series_list = chart_spec.get("series") or []
                    if series_list and isinstance(series_list[0], dict):
                        block_key = str(series_list[0].get("block_key") or "").strip().lower()
                        if block_key in {"compras", "ventas"}:
                            return block_key
                    block_key = str(chart_spec.get("category_block_key") or "").strip().lower()
                    return block_key if block_key in {"compras", "ventas"} else ""

                chart_specs_by_key = {
                    _chart_primary_key(spec): spec
                    for spec in trend_charts
                    if _chart_primary_key(spec)
                }
                if table_anchor_row is not None and last_block is not None:
                    chart_start_row = max(int(last_block["header_row"]) + 4, int(table_anchor_row) - EXCEL_CHART_PAIR_ROW_OFFSET)
                compras_spec = chart_specs_by_key.get("compras")
                ventas_spec = chart_specs_by_key.get("ventas")
                if compras_spec is not None:
                    compras_col = table_anchor_map.get("compras", chart_start_col)
                    compras_spec["anchor_row"] = chart_start_row
                    compras_spec["anchor_col"] = compras_col
                if compras_spec is not None and ventas_spec is not None:
                    compras_col = int(compras_spec.get("anchor_col", chart_start_col))
                    compras_span = _estimate_excel_chart_col_span(compras_spec.get("width_emu"))
                    ventas_base_col = table_anchor_map.get("ventas", compras_col + compras_span)
                    ventas_col = max(
                        ventas_base_col + EXCEL_CHART_SECONDARY_SHIFT_COLUMNS,
                        compras_col + max(1, int(round(compras_span * EXCEL_CHART_PAIR_OVERLAP_RATIO))),
                    )
                    ventas_spec["anchor_row"] = chart_start_row
                    ventas_spec["anchor_col"] = ventas_col
            _add_native_trend_charts_to_sheet(
                ws,
                trend_charts,
                blocks,
                sheet_runtime,
                chart_start_row,
                start_col=chart_start_col,
            )
        if export_images:
            base_row = max(ws.max_row + 3, 5)
            current_row = base_row
            for image_entry in export_images:
                image_bytes = image_entry.get("image_bytes")
                if not image_bytes:
                    continue
                anchor_cell = str(image_entry.get("anchor_cell") or "").strip()
                anchor_col = max(1, int(image_entry.get("anchor_col", 1) or 1))
                anchor_row = max(current_row, int(image_entry.get("anchor_row", base_row) or base_row))
                image_stream = io.BytesIO(image_bytes)
                image_stream.seek(0)
                excel_image = OpenpyxlImage(image_stream)
                width_px = image_entry.get("width_px")
                height_px = image_entry.get("height_px")
                try:
                    if width_px:
                        excel_image.width = float(width_px)
                    if height_px:
                        excel_image.height = float(height_px)
                except Exception:
                    pass
                if anchor_cell:
                    ws.add_image(excel_image, anchor_cell)
                else:
                    ws.add_image(excel_image, f"{get_column_letter(anchor_col)}{anchor_row}")
                estimated_rows = max(18, int(math.ceil((float(height_px) if height_px else 420.0) / 20.0)))
                current_row = anchor_row + estimated_rows
        _apply_pending_excel_column_widths(ws, sheet_runtime)
    return save_workbook_with_retry(workbook, output_path)


def normalize_brand_key(brand: str) -> str:
    """Normaliza fabricante para mapearlo siempre al mismo color."""
    normalized = str(brand or "").strip()
    normalized = re.sub(
        r'(?:[._\-\s](?:k|m|g|kg|grs?|gr|l|lt|lts|ml))+$',
        '',
        normalized,
        flags=re.IGNORECASE,
    )
    normalized = re.sub(r'\s*-\s*share\s*100%\s*apilado\s*$', '', normalized, flags=re.IGNORECASE)
    normalized = re.sub(r'\s*-\s*share\s*100%\s*stacked\s*$', '', normalized, flags=re.IGNORECASE)
    normalized = normalized.replace('.', ' ').replace('_', ' ')
    normalized = unicodedata.normalize("NFKD", normalized)
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized

def parse_filename_brand(filename: str) -> Tuple[str, str, str, str]:
    """Separa nombre de archivo en prefijo, marca, sufijo y extension."""
    base, ext = os.path.splitext(filename)
    parts = base.split("_", 2)
    if len(parts) < 3:
        return "", "", "", ext
    prefix = f"{parts[0]}_{parts[1]}_"
    raw_brand_segment = parts[2].strip()
    if not raw_brand_segment:
        return prefix, "", "", ext
    match = PREVIEW_COPY_SUFFIX_RE.match(raw_brand_segment)
    if not match:
        return prefix, raw_brand_segment, "", ext
    brand = (match.group("brand") or "").strip()
    suffix = match.group("suffix") or ""
    return prefix, brand, suffix, ext


def parse_standard_sheet_name(sheet_name: str) -> Optional[dict[str, str]]:
    """Parsea hojas tipo '3_Marca_1 (ML)' sin depender de indices fijos."""
    if not isinstance(sheet_name, str):
        return None
    match = re.match(
        r'^(?P<segment>[1-8])[-_](?P<label>.+?)(?:_(?P<variant>\d+)(?:\s*\((?P<qualifier>[^)]*)\))?)?$',
        sheet_name.strip(),
        re.IGNORECASE,
    )
    if not match:
        return None
    label = (match.group('label') or '').replace('_', ' ').strip()
    if not label:
        return None
    return {
        'segment': (match.group('segment') or '').strip(),
        'label': label,
        'variant': (match.group('variant') or '').strip(),
        'qualifier': (match.group('qualifier') or '').strip(),
    }

def relative_luminance(rgb: Tuple[int, int, int]) -> float:
    """Calcula luminancia relativa para evitar colores oscuros."""
    r, g, b = rgb
    return (0.2126 * r) + (0.7152 * g) + (0.0722 * b)

def lift_color_to_min_luminance(
    rgb: Tuple[int, int, int],
    min_luminance: float = MIN_READABLE_LUMINANCE
) -> Tuple[int, int, int]:
    """Aclara color mezclandolo con blanco hasta luminancia minima."""
    lum = relative_luminance(rgb)
    if lum >= min_luminance:
        return rgb
    r, g, b = rgb
    mix = min(0.7, max(0.0, (min_luminance - lum) / 255.0))
    r = int(round(r + (255 - r) * mix))
    g = int(round(g + (255 - g) * mix))
    b = int(round(b + (255 - b) * mix))
    return (r, g, b)

def ansi_truecolor(text: str, rgb: Tuple[int, int, int]) -> str:
    """Envuelve texto con ANSI truecolor (24-bit)."""
    r, g, b = rgb
    return f"\033[38;2;{r};{g};{b}m{text}{COLOR_RESET}"

def brand_color_rgb(brand: str, color_cache: Dict[str, Tuple[int, int, int]]) -> Tuple[int, int, int]:
    """Asigna un color estable y distinguible por marca para listados en terminal."""
    key = normalize_brand_key(brand)
    if not key:
        return (255, 235, 59)
    cached = color_cache.get(key)
    if cached:
        return cached
    used_colors = set(color_cache.values())
    for candidate in TERMINAL_BRAND_COLOR_SEQUENCE:
        readable_candidate = lift_color_to_min_luminance(candidate)
        if readable_candidate not in used_colors:
            color_cache[key] = readable_candidate
            return readable_candidate
    fallback = TERMINAL_BRAND_COLOR_SEQUENCE[len(color_cache) % len(TERMINAL_BRAND_COLOR_SEQUENCE)]
    fallback = lift_color_to_min_luminance(fallback)
    color_cache[key] = fallback
    return fallback

def colorize_filename_brand(filename: str, color_cache: Dict[str, Tuple[int, int, int]]) -> str:
    """Colorea solo la marca dentro del nombre del archivo .xlsx."""
    prefix, brand, suffix, ext = parse_filename_brand(filename)
    if not brand:
        return filename
    rgb = brand_color_rgb(brand, color_cache)
    brand_colored = ansi_truecolor(brand, rgb)
    return f"{prefix}{brand_colored}{COLOR_BLUE}{suffix}{ext}"

def select_trend_scale_rule(max_value: float) -> dict:
    """Selecciona la regla de escala segun el valor maximo."""
    for rule in TREND_SCALE_RULES:
        if max_value >= rule["threshold"]:
            return rule
    return TREND_SCALE_RULES[-1]
def format_value_with_suffix(value: float, divisor: float, suffix: str) -> str:
    """Formatea un valor escalado con sufijo (K/M/B/T)."""
    if not np.isfinite(value):
        return ''
    scaled = value / divisor if divisor else value
    abs_scaled = abs(scaled)
    if abs_scaled >= 100:
        formatted = f"{scaled:.0f}"
    elif abs_scaled >= 10:
        formatted = f"{scaled:.1f}"
    else:
        formatted = f"{scaled:.2f}"
    if '.' in formatted:
        formatted = formatted.rstrip('0').rstrip('.')
    return f"{formatted}{suffix}"
def available_width(presentation, left=None, right=None):
    """Calcula el ancho disponible en la diapositiva."""
    if left is None:
        left = Inches(0)
    if right is None:
        right = Inches(0)
    return presentation.slide_width - int(left) - int(right)
def constrain_picture_width(picture, max_width):
    """Limita el ancho de una imagen preservando la relacion de aspecto."""
    if max_width is None:
        return
    max_width = int(max_width)
    if max_width <= 0:
        return
    current_width = int(picture.width)
    if current_width <= max_width:
        return
    current_height = int(picture.height)
    scale = current_width / max_width
    picture.width = max_width
    picture.height = max(int(round(current_height / scale)), 1)
EMU_PER_INCH = 914400
DEFAULT_LINE_CHART_RATIO = 3
TABLE_TARGET_HEIGHT_CM = 4.0
TABLE_SIDE_MARGIN_CM = 1.2
TABLE_PAIR_GAP_CM = 0.5
TABLE_HEADER_FONT_SIZE = 11
TABLE_BODY_FONT_SIZE = 11
TABLE_WRAP_WIDTH = 14
LINE_CHART_LEFT_MARGIN = 0.045
LINE_CHART_RIGHT_MARGIN = 0.99
LINE_CHART_BOTTOM_MARGIN = 0.1
LINE_CHART_TOP_MARGIN = 0.9
LINE_CHART_SINGLE_X_MARGIN = 0.16
LINE_CHART_MULTI_LEFT_MARGIN = 0.08
LINE_CHART_MULTI_RIGHT_MARGIN = 0.98
LINE_CHART_MULTI_BOTTOM_MARGIN = 0.12
LINE_CHART_MULTI_TOP_MARGIN = 0.85
LINE_CHART_MULTI_X_MARGIN = 0.16
DEFAULT_EXPORT_DPI = 110
TABLE_EXPORT_DPI = 220
TABLE_APORTE_EXPORT_DPI = 150
EXPORT_PAD_INCHES = 0.08
CHART_TOP_INCH = 0.72
def emu_to_inches(value: int) -> float:
    """Convierte unidades EMU a pulgadas."""
    return float(value) / EMU_PER_INCH
def wrap_table_text(value, max_width: int = TABLE_WRAP_WIDTH):
    """
    Devuelve el valor con saltos suaves para ajustar el texto en celdas.
    Mantiene los valores no string sin cambios.
    """
    if not isinstance(value, str):
        return value
    normalized = " ".join(value.split())
    if not normalized:
        return ""
    if len(normalized) <= max_width or " " not in normalized:
        return normalized
    return textwrap.fill(normalized, width=max_width, break_long_words=False)
def figure_to_stream(
    fig: plt.Figure,
    dpi: int = DEFAULT_EXPORT_DPI,
    bbox_inches="tight",
    pad_inches: float = EXPORT_PAD_INCHES,
    transparent: bool = True,
    close: bool = True,
) -> io.BytesIO:
    """
    Renderiza una figura en memoria y la cierra para liberar recursos.
    Replica el flujo rapido usado en coverage_studio (guardar en BytesIO y cerrar).
    """
    buf = io.BytesIO()
    savefig_kwargs = {
        "format": "png",
        "dpi": dpi,
        "bbox_inches": bbox_inches,
        "pad_inches": pad_inches,
        "transparent": transparent,
    }
    fig.savefig(buf, **savefig_kwargs)
    buf.seek(0)
    if close:
        try:
            plt.close(fig)
        except Exception:
            pass
    return buf

# ---------------- Segmento 8: tabla PPT (intervalos) ----------------

# Valores z: limites (2.0) y error muestral (1.96).
PPT8_Z_SCORE_INTERVAL = 2.0
PPT8_Z_SCORE_ERROR = 1.96


@dataclass
class Ppt8MonthlyBlock:
    brand: str
    rows: pd.DataFrame  # columnas: 0=date, 1=Weighted PENET, 2=BUYERS, 3=ventas_pipeline (opcional)
    pipeline: int = 0


@dataclass
class Ppt8AggBlock:
    brand: str
    block: pd.DataFrame  # columnas variables: label, MAT base, MAT actual
    label_col: int = 4
    base_col: int = 5
    current_col: int = 6


@dataclass
class Ppt8Row:
    marca: str
    var_fabricante: float
    mat_sep25: float
    lim_inf: float
    lim_sup: float
    compradores_prom: float
    pen12m: float
    nivel_pen: str
    error_muestral: float


def ppt8_normalize_label(label: str) -> str:
    """Normaliza etiquetas para emparejar bloques por marca."""
    normalized = unicodedata.normalize("NFKD", str(label))
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", normalized.lower())


def ppt8_is_header_label(label: str) -> bool:
    """Detecta etiquetas de encabezado en lugar de marcas."""
    raw = str(label).strip()
    if not raw:
        return True
    lower = raw.lower()
    if re.search(r"\btable\b", lower):
        return True
    if re.search(r"\bcateg", lower):
        return True
    if re.search(r"\bbrands\b", lower) and ("=" in lower or "total" in lower):
        return True
    compact = ppt8_normalize_label(label)
    return not compact


def ppt8_clean_dimension_label(label: str) -> str:
    """Limpia etiquetas de dimensiones tipo 'DIM = Total X\\Total DIM'."""
    raw = str(label).strip()
    if not raw:
        return ""
    cleaned = raw
    if "=" in cleaned:
        cleaned = cleaned.split("=", 1)[1].strip()
    if "\\" in cleaned:
        cleaned = cleaned.split("\\", 1)[0].strip()
    cleaned = re.sub(r"(?i)^total\s+", "", cleaned).strip()
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or raw


def _ppt8_normalize_text(value) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text.lower().strip()


def ppt8_is_monthly_header_row(row: pd.Series) -> bool:
    """Detecta encabezados mensuales del bloque izquierdo (PENET/BUYERS)."""
    if len(row) < 3:
        return False
    col0 = _ppt8_normalize_text(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
    col1 = _ppt8_normalize_text(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
    col2 = _ppt8_normalize_text(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
    has_penet = "penet" in col1
    has_buyers = ("buyer" in col2) or ("comprad" in col2)
    has_table = "table" in col0
    return has_penet and has_buyers and (has_table or bool(col1 and col2))


def ppt8_is_agg_header_row(row: pd.Series, label_col: int, base_col: int, current_col: int) -> bool:
    """Detecta encabezados MAT del bloque agregado derecho."""
    if len(row) <= max(label_col, base_col, current_col):
        return False
    col_label = _ppt8_normalize_text(row.iloc[label_col]) if pd.notna(row.iloc[label_col]) else ""
    col_base = _ppt8_normalize_text(row.iloc[base_col]) if pd.notna(row.iloc[base_col]) else ""
    col_curr = _ppt8_normalize_text(row.iloc[current_col]) if pd.notna(row.iloc[current_col]) else ""
    has_table = "table" in col_label
    has_mat = ("mat" in col_base) or ("mat" in col_curr)
    if not has_mat:
        has_mat = bool(re.search(r"y\d{4}", col_base)) and bool(re.search(r"y\d{4}", col_curr))
    return has_table and has_mat


def _ppt8_resolve_block_label(
    df: pd.DataFrame,
    header_idx: int,
    label_col: int,
    val_col_a: int,
    val_col_b: int,
) -> str:
    """Resuelve la etiqueta del bloque buscando fila de marca cercana al encabezado."""
    max_rows = len(df)
    # Escenario base: etiqueta inmediatamente debajo del encabezado.
    for row_idx in range(header_idx + 1, min(header_idx + 4, max_rows)):
        label_val = df.iloc[row_idx, label_col]
        if pd.isna(label_val):
            continue
        val_a = df.iloc[row_idx, val_col_a]
        val_b = df.iloc[row_idx, val_col_b]
        if pd.isna(val_a) and pd.isna(val_b):
            cleaned = ppt8_clean_dimension_label(label_val)
            if cleaned and not ppt8_is_header_label(cleaned):
                return cleaned
    # Escenario alterno: metadatos justo encima del encabezado (ej.: L_REGION).
    for row_idx in range(header_idx - 1, max(-1, header_idx - 5), -1):
        label_val = df.iloc[row_idx, label_col]
        if pd.isna(label_val):
            continue
        val_a = df.iloc[row_idx, val_col_a]
        val_b = df.iloc[row_idx, val_col_b]
        if pd.isna(val_a) and pd.isna(val_b):
            cleaned = ppt8_clean_dimension_label(label_val)
            if cleaned and not ppt8_is_header_label(cleaned):
                return cleaned
    fallback = str(df.iloc[header_idx, label_col]).strip()
    return ppt8_clean_dimension_label(fallback) or fallback or f"Bloque {header_idx + 1}"


def ppt8_block_has_numeric(block: pd.DataFrame, base_col: int = 5, current_col: int = 6) -> bool:
    """Verifica si el bloque tiene valores numericos en las columnas de MAT."""
    if block.empty:
        return False
    for col in (base_col, current_col):
        if block.shape[1] > col:
            numeric = pd.to_numeric(block.iloc[:, col], errors="coerce")
            if numeric.notna().any():
                return True
    return False


def ppt8_detect_agg_columns(df: pd.DataFrame) -> Tuple[int, int, int]:
    """Detecta columnas (label, MAT base, MAT actual) para formato viejo/nuevo."""
    if df is None or df.empty:
        return 4, 5, 6
    if df.shape[1] <= 6:
        return max(0, df.shape[1] - 3), max(0, df.shape[1] - 2), max(0, df.shape[1] - 1)
    last_start = max(0, df.shape[1] - 3)
    best_triplet = (4, 5, 6)
    best_score = -1
    for label_col in range(3, last_start + 1):
        base_col = label_col + 1
        current_col = label_col + 2
        labels = df.iloc[:, label_col].fillna("").astype(str)
        base_numeric = pd.to_numeric(df.iloc[:, base_col], errors="coerce").notna()
        curr_numeric = pd.to_numeric(df.iloc[:, current_col], errors="coerce").notna()
        numeric_mask = base_numeric | curr_numeric
        metric_like = labels.str.contains(
            r"weighted|r_vol|vol|penet|freq|hholds?|buyers?",
            case=False,
            regex=True,
            na=False,
        )
        score = int((metric_like & numeric_mask).sum())
        if score > best_score:
            best_score = score
            best_triplet = (label_col, base_col, current_col)
    if best_score < 1 and df.shape[1] > 6:
        return 4, 5, 6
    return best_triplet


def ppt8_filter_agg_blocks(
    monthly_blocks: Iterable[Ppt8MonthlyBlock],
    agg_blocks: Iterable[Ppt8AggBlock],
) -> Tuple[List[Ppt8AggBlock], List[str]]:
    """Omite bloques agregados sin datos o que son encabezados."""
    monthly_norms = {ppt8_normalize_label(block.brand) for block in monthly_blocks if block.brand}
    filtered: List[Ppt8AggBlock] = []
    dropped: List[str] = []
    for block in agg_blocks:
        brand_label = block.brand
        norm = ppt8_normalize_label(brand_label)
        header_like = ppt8_is_header_label(brand_label)
        has_numeric = ppt8_block_has_numeric(block.block, block.base_col, block.current_col)
        if header_like and not has_numeric:
            dropped.append(brand_label)
            continue
        if norm and norm not in monthly_norms and not has_numeric:
            dropped.append(brand_label)
            continue
        filtered.append(block)
    return filtered, dropped


def ppt8_parse_monthly_blocks(df: pd.DataFrame) -> List[Ppt8MonthlyBlock]:
    """Encuentra bloques de datos mensuales delimitados por marca en columnas 0-2."""
    if df.shape[1] < 3:
        return []
    header_blocks: List[Ppt8MonthlyBlock] = []
    header_rows = [idx for idx, row in df.iterrows() if ppt8_is_monthly_header_row(row)]
    if header_rows:
        for i, header_idx in enumerate(header_rows):
            end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df)
            data = df.loc[header_idx + 1 : end - 1]
            data = data[pd.to_numeric(data.iloc[:, 1], errors="coerce").notna()]
            if data.empty:
                continue
            brand_label = _ppt8_resolve_block_label(df, header_idx, 0, 1, 2)
            if not brand_label or ppt8_is_header_label(brand_label):
                continue
            pipeline = 0
            if df.shape[1] > 3:
                pipeline = _extract_pipeline(df.iloc[header_idx, 3])
            header_blocks.append(Ppt8MonthlyBlock(brand=brand_label, rows=data, pipeline=pipeline))

    first_col = df.iloc[:, 0]
    brand_rows = df[(first_col.notna()) & df.iloc[:, 1].isna() & df.iloc[:, 2].isna()].index.tolist()
    fallback_blocks: List[Ppt8MonthlyBlock] = []
    for i, start in enumerate(brand_rows):
        end = brand_rows[i + 1] if i + 1 < len(brand_rows) else len(df)
        data = df.loc[start + 1 : end - 1]
        brand_label = ppt8_clean_dimension_label(df.iloc[start, 0])
        if ppt8_is_header_label(brand_label):
            continue
        data = data[pd.to_numeric(data.iloc[:, 1], errors="coerce").notna()]
        if data.empty:
            continue
        pipeline = 0
        header_idx = start + 1
        if df.shape[1] > 3 and header_idx < len(df.index):
            pipeline = _extract_pipeline(df.iloc[header_idx, 3])
        fallback_blocks.append(Ppt8MonthlyBlock(brand=brand_label, rows=data, pipeline=pipeline))
    if header_blocks and len(header_blocks) >= len(fallback_blocks):
        return header_blocks
    return fallback_blocks


def ppt8_parse_agg_blocks(df: pd.DataFrame) -> List[Ppt8AggBlock]:
    """Encuentra bloques agregados por marca en columnas 4-6."""
    if df.shape[1] < 7:
        return []
    label_col, base_col, current_col = ppt8_detect_agg_columns(df)
    header_blocks: List[Ppt8AggBlock] = []
    header_rows = [
        idx for idx, row in df.iterrows()
        if ppt8_is_agg_header_row(row, label_col, base_col, current_col)
    ]
    if header_rows:
        for i, header_idx in enumerate(header_rows):
            end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df)
            block = df.loc[header_idx + 1 : end - 1]
            block = block[block.iloc[:, label_col].notna()]
            if block.empty:
                continue
            brand_label = _ppt8_resolve_block_label(df, header_idx, label_col, base_col, current_col)
            if not brand_label or ppt8_is_header_label(brand_label):
                continue
            header_blocks.append(
                Ppt8AggBlock(
                    brand=brand_label,
                    block=block,
                    label_col=label_col,
                    base_col=base_col,
                    current_col=current_col,
                )
            )

    label_series = df.iloc[:, label_col]
    agg_brand_rows = df[
        label_series.notna()
        & (~label_series.astype(str).str.contains("Weighted", case=False, na=False))
        & (~label_series.astype(str).str.contains("table", case=False, na=False))
        & (~label_series.astype(str).str.contains("categ", case=False, na=False))
        & df.iloc[:, base_col].isna()
        & df.iloc[:, current_col].isna()
    ].index.tolist()
    fallback_blocks: List[Ppt8AggBlock] = []
    for i, start in enumerate(agg_brand_rows):
        end = agg_brand_rows[i + 1] if i + 1 < len(agg_brand_rows) else len(df)
        block = df.loc[start + 1 : end - 1]
        block = block[block.iloc[:, label_col].notna()]
        if block.empty:
            continue
        brand_label = ppt8_clean_dimension_label(df.iloc[start, label_col])
        if ppt8_is_header_label(brand_label):
            continue
        fallback_blocks.append(
            Ppt8AggBlock(
                brand=brand_label,
                block=block,
                label_col=label_col,
                base_col=base_col,
                current_col=current_col,
            )
        )
    if header_blocks and len(header_blocks) >= len(fallback_blocks):
        return header_blocks
    return fallback_blocks


def ppt8_extract_metrics(
    block: pd.DataFrame,
    label_col: int = 4,
    base_col: int = 5,
    current_col: int = 6,
) -> Tuple[Dict[str, Tuple[float, float]], Dict[str, str]]:
    """Mapea etiqueta -> (MAT base, MAT actual) y registra la fuente detectada."""

    volume_metric_priority = {
        "RVOL1": 400,
        "RVOL2": 390,
        "VOL1PCO": 350,
        "VOL2PCO": 340,
        "VOL1P": 320,
        "VOL2P": 310,
        "VOL1": 280,
        "VOL2": 270,
        "VOLSU": 240,
        "UNITS": 200,
        "UNIDAD": 200,
        "UNIDADES": 200,
    }

    def canonical_metric(label: str) -> Tuple[Optional[str], Optional[str], int]:
        raw_label = str(label)
        if not raw_label or not raw_label.strip():
            return None, None, -1
        is_weighted = "WEIGHTED" in raw_label.upper()
        normalized = re.sub(r"(?i)\bweighted\b", "", raw_label).strip().upper()
        compact = re.sub(r"[^A-Z0-9]+", "", normalized)
        if not compact:
            return None, None, -1
        is_vertical = "VERT" in compact

        if is_vertical:
            return None, None, -1
        if compact in {
            "RVOL1", "RVOL2",
            "VOL1", "VOL2",
            "VOL1P", "VOL2P",
            "VOL1PCO", "VOL2PCO",
            "VOLSU", "UNITS", "UNIDAD", "UNIDADES"
        }:
            return "R_VOL1", compact, volume_metric_priority.get(compact, 0)
        if compact.startswith("PEN"):
            return "PENET", compact, 100
        if compact.startswith("FREQ") or compact in {"FRDAY", "FRDAYS"}:
            return "FREQ", compact, 100
        if compact.startswith("HH"):
            priority = 200 if not is_weighted else 100
            return "HHOLDS", compact, priority
        return normalized, compact, 0

    metrics: Dict[str, Tuple[float, float]] = {}
    metric_priority: Dict[str, int] = {}
    metric_source: Dict[str, str] = {}
    for _, row in block.iterrows():
        if len(row) <= max(label_col, base_col, current_col):
            continue
        canonical, source_compact, priority = canonical_metric(row.iloc[label_col])
        if not canonical:
            continue
        try:
            base_val = float(row.iloc[base_col])
            act_val = float(row.iloc[current_col])
        except (TypeError, ValueError):
            continue
        # Si solo uno de los dos valores es numerico, replica para no perder la medida
        if not np.isfinite(act_val) and np.isfinite(base_val):
            act_val = base_val
        if not np.isfinite(base_val) and np.isfinite(act_val):
            base_val = act_val
        previous_priority = metric_priority.get(canonical, -1)
        # Evita que metricas de menor prioridad (ej. VOL2_P) pisen R_VOL2/R_VOL1.
        if previous_priority > priority:
            continue
        metrics[canonical] = (base_val, act_val)
        metric_priority[canonical] = priority
        if source_compact:
            metric_source[canonical] = source_compact
    return metrics, metric_source


def ppt8_detect_volume_metric(agg_blocks: Iterable[Ppt8AggBlock]) -> Optional[str]:
    """Detecta la metrica de volumen usada en Segmento 8 para reportarla en terminal."""
    source_counts: Dict[str, int] = {}
    source_priority = {
        "RVOL1": 400,
        "RVOL2": 390,
        "VOL1PCO": 350,
        "VOL2PCO": 340,
        "VOL1P": 320,
        "VOL2P": 310,
        "VOL1": 280,
        "VOL2": 270,
        "VOLSU": 240,
        "UNITS": 200,
        "UNIDAD": 200,
        "UNIDADES": 200,
    }
    source_display = {
        "RVOL1": "R_VOL1",
        "RVOL2": "R_VOL2",
        "VOL1": "VOL1",
        "VOL2": "VOL2",
        "VOL1P": "VOL1_P",
        "VOL2P": "VOL2_P",
        "VOL1PCO": "VOL1_PCO",
        "VOL2PCO": "VOL2_PCO",
        "VOLSU": "VOLSU",
        "UNITS": "UNITS",
        "UNIDAD": "UNIDAD",
        "UNIDADES": "UNIDADES",
    }
    for block in agg_blocks:
        _, metric_sources = ppt8_extract_metrics(
            block.block,
            block.label_col,
            block.base_col,
            block.current_col,
        )
        source = metric_sources.get("R_VOL1")
        if source:
            source_counts[source] = source_counts.get(source, 0) + 1
    if not source_counts:
        return None
    selected_source = max(
        source_counts.keys(),
        key=lambda key: (source_counts.get(key, 0), source_priority.get(key, 0), key),
    )
    return source_display.get(selected_source, selected_source)


def ppt8_extract_metrics_values(
    block: pd.DataFrame,
    label_col: int = 4,
    base_col: int = 5,
    current_col: int = 6,
) -> Dict[str, Tuple[float, float]]:
    """Compatibilidad: devuelve solo valores de metricas agregadas."""
    metrics, _ = ppt8_extract_metrics(block, label_col, base_col, current_col)
    return metrics


def ppt8_find_mat_labels(df: pd.DataFrame) -> Tuple[str, str]:
    """Detecta las etiquetas de MAT (base y actual) en las columnas 5 y 6."""
    mat_base = "MAT Base"
    mat_current = "MAT Actual"
    for _, row in df.iterrows():
        val5 = row.iloc[5] if df.shape[1] > 5 else None
        val6 = row.iloc[6] if df.shape[1] > 6 else None
        if isinstance(val5, str) and "MAT" in val5.upper():
            mat_base = val5.strip()
        if isinstance(val6, str) and "MAT" in val6.upper():
            mat_current = val6.strip()
        if "MAT" in str(val5).upper() or "MAT" in str(val6).upper():
            if mat_base != "MAT Base" and mat_current != "MAT Actual":
                break
    return mat_base, mat_current


def ppt8_classify_penetration(p: float) -> str:
    """Clasifica la penetracion (0-1)."""
    if p <= 0:
        return "Fuera de rango"
    if 0.01 <= p <= 0.10:
        return "Marca Pequeña (1% a 10%)"
    if 0.10 < p <= 0.30:
        return "Marca Mediana (11% a 30%)"
    if 0.30 < p <= 0.50:
        return "Marca Grande (31% a 50%)"
    if 0.50 < p <= 0.99:
        return "Super Marca (51% a 99%)"
    return "Fuera de rango"


def compute_mat_variation_percent(values: pd.Series, pipeline_offset: int = 0) -> float:
    """Calcula variacion MAT % usando desplazamiento por pipeline."""
    if values is None:
        return np.nan
    try:
        p = int(pipeline_offset) if pipeline_offset is not None else 0
    except (TypeError, ValueError):
        p = 0
    p = max(0, p)
    series = pd.to_numeric(values, errors="coerce")
    total = len(series)
    if total < (24 + p):
        return np.nan
    current_end = total - p if p > 0 else total
    current_start = current_end - 12
    previous_start = current_start - 12
    if previous_start < 0 or current_start <= previous_start:
        return np.nan
    previous_sum = series.iloc[previous_start:current_start].sum(min_count=1)
    current_sum = series.iloc[current_start:current_end].sum(min_count=1)
    if not np.isfinite(previous_sum) or not np.isfinite(current_sum) or previous_sum == 0:
        return np.nan
    return ((current_sum / previous_sum) - 1.0) * 100.0


def ppt8_compute_rows(
    monthly_blocks: Iterable[Ppt8MonthlyBlock], agg_blocks: Iterable[Ppt8AggBlock]
) -> Tuple[List[Ppt8Row], List[str]]:
    """Combina bloques mensuales y agregados para construir filas del cuadro PPT."""
    rows: List[Ppt8Row] = []
    errors: List[str] = []
    required_metrics = {"R_VOL1", "PENET", "FREQ", "HHOLDS"}
    agg_lookup = {ppt8_normalize_label(block.brand): block for block in agg_blocks}

    for idx, mblock in enumerate(monthly_blocks):
        try:
            ablock = agg_lookup.get(ppt8_normalize_label(mblock.brand))
            if ablock is None:
                ablock = agg_blocks[idx] if idx < len(agg_blocks) else None
            if ablock is None:
                raise ValueError(f"{mblock.brand}: no se encontro bloque agregado correspondiente")
            monthly = mblock.rows
            if monthly.empty:
                raise ValueError(f"{mblock.brand}: sin datos mensuales validos")
            last12_pen = pd.to_numeric(monthly.iloc[:, 1], errors="coerce").tail(12)
            last12_buy = pd.to_numeric(monthly.iloc[:, 2], errors="coerce").tail(12)
            if last12_pen.dropna().empty or last12_buy.dropna().empty:
                raise ValueError(f"{mblock.brand}: faltan valores para promedio 12m")
            pen_avg = float(last12_pen.mean())
            buyers_avg = float(last12_buy.mean())

            metrics = ppt8_extract_metrics_values(
                ablock.block,
                ablock.label_col,
                ablock.base_col,
                ablock.current_col,
            )
            missing = [m for m in required_metrics if m not in metrics]
            if missing:
                raise ValueError(f"{mblock.brand}: faltan metricas {', '.join(missing)}")
            vol1, vol2 = metrics["R_VOL1"]
            pen1, pen2 = metrics["PENET"]
            freq1, freq2 = metrics["FREQ"]
            hh_tuple = metrics.get("HHOLDS", (np.nan, np.nan))
            hh = hh_tuple[1]
            if not np.isfinite(hh):
                hh = hh_tuple[0]

            for name, value in {
                "vol1": vol1,
                "vol2": vol2,
                "pen1": pen1,
                "pen2": pen2,
                "freq1": freq1,
                "freq2": freq2,
                "hh": hh,
            }.items():
                if value is None or not np.isfinite(value):
                    raise ValueError(f"{mblock.brand}: valor invalido en {name}")
            if hh <= 0:
                raise ValueError(f"{mblock.brand}: HHOLDS no es positivo")
            if vol1 == 0:
                raise ValueError(f"{mblock.brand}: MAT base es cero")

            dif = vol2 - vol1
            b = (pen1 + pen2) / 200
            w = (freq1 + freq2) / 2
            m_val = b * w
            if m_val <= 0:
                raise ValueError(f"{mblock.brand}: valor m invalido")
            se_rel = math.sqrt(2 / (m_val * hh))
            avg_vol = (vol1 + vol2) / 2
            int_dif = PPT8_Z_SCORE_INTERVAL * se_rel * avg_vol
            lim_dif_minus = dif - int_dif
            lim_dif_plus = dif + int_dif
            lim_cant2_minus = vol1 + lim_dif_minus
            lim_cant2_plus = vol1 + lim_dif_plus

            evol = (vol2 / vol1 - 1) * 100
            lim_evol_minus = (lim_cant2_minus / vol1 - 1) * 100
            lim_evol_plus = (lim_cant2_plus / vol1 - 1) * 100

            p_ratio = pen_avg / 100
            if p_ratio <= 0:
                err = np.nan
            else:
                err = PPT8_Z_SCORE_ERROR * math.sqrt((p_ratio * (1 - p_ratio)) / hh) / p_ratio
            nivel = ppt8_classify_penetration(p_ratio)
            var_fabricante = np.nan
            if monthly.shape[1] > 3:
                var_fabricante = compute_mat_variation_percent(monthly.iloc[:, 3], mblock.pipeline)

            rows.append(
                Ppt8Row(
                    marca=mblock.brand,
                    var_fabricante=var_fabricante,
                    mat_sep25=evol,
                    lim_inf=lim_evol_minus,
                    lim_sup=lim_evol_plus,
                    compradores_prom=buyers_avg,
                    pen12m=pen_avg,
                    nivel_pen=nivel,
                    error_muestral=err,
                )
            )
        except ValueError as exc:
            errors.append(str(exc))
    return rows, errors


def render_ppt8_table(
    rows: List[Ppt8Row],
    figure_id: int,
    brand_color_lookup: Dict[str, str],
    manufacturer_name: str,
) -> Tuple[io.BytesIO, Tuple[float, float]]:
    """Renderiza la tabla PPT del segmento 8 y devuelve un stream PNG."""
    COLOR_HEADER = "#1F4E78"
    COLOR_VAR_BG = "#C9CED6"
    COLOR_C = "#F2DCDB"
    COLOR_D = "#E2EFDA"
    COLOR_G = "#DCE6F1"
    COLOR_H = "#E2EFDA"
    COLOR_ERR_GREEN = "#A8D08D"
    COLOR_ERR_YELLOW = "#FFD966"
    COLOR_ERR_RED = "#C00000"

    def _interp_color(start_hex: str, end_hex: str, t: float) -> str:
        t = max(0.0, min(1.0, t))
        start_rgb = mcolors.to_rgb(start_hex)
        end_rgb = mcolors.to_rgb(end_hex)
        blended = tuple(s + (e - s) * t for s, e in zip(start_rgb, end_rgb))
        return mcolors.to_hex(blended)

    def _error_cell_color(err_ratio: float) -> str:
        """Color degradado según error muestral (%)."""
        if err_ratio is None or not np.isfinite(err_ratio):
            return COLOR_H
        val = err_ratio * 100.0
        if val <= 1.0:
            return COLOR_ERR_GREEN
        if val <= 5.0:
            t = (val - 1.0) / 4.0
            return _interp_color(COLOR_ERR_GREEN, COLOR_ERR_YELLOW, t)
        if val <= 51.0:
            t = (val - 5.0) / 46.0
            return _interp_color(COLOR_ERR_YELLOW, COLOR_ERR_RED, t)
        return COLOR_ERR_RED

    def _text_color_for_bg(hex_color: str) -> str:
        r, g, b = mcolors.to_rgb(hex_color)
        if 0.299 * r + 0.587 * g + 0.114 * b < 0.45:
            return "#FFFFFF"
        return "#1F1F1F"

    def _metric_text_color(value: float) -> str:
        if value is None or not np.isfinite(value):
            return "#1F1F1F"
        if value > 0:
            return TABLE_POSITIVE_COLOR
        if value < 0:
            return TABLE_NEGATIVE_COLOR
        return "#1F1F1F"

    fabricante_col_name = manufacturer_name.strip() if isinstance(manufacturer_name, str) else ""
    fabricante_col_name = fabricante_col_name or "Fabricante"
    header_bottom = [
        "Marca",
        f"% VAR {fabricante_col_name}",
        "% VAR MAT Actual (WP by Numerator)",
        "% VAR LIM INFERIOR",
        "% VAR LIM SUPERIOR",
        "Nivel de penetracion",
        "Error muestral",
    ]

    data = [
        [
            r.marca,
            f"{r.var_fabricante:.1f}" if np.isfinite(r.var_fabricante) else "-",
            f"{r.mat_sep25:.1f}" if np.isfinite(r.mat_sep25) else "-",
            f"{r.lim_inf:.1f}" if np.isfinite(r.lim_inf) else "-",
            f"{r.lim_sup:.1f}" if np.isfinite(r.lim_sup) else "-",
            r.nivel_pen,
            f"{r.error_muestral:.1%}" if np.isfinite(r.error_muestral) else "-",
        ]
        for r in rows
    ]

    fig_height = 1.4 + 0.55 * len(rows)
    fig, ax = plt.subplots(num=figure_id, figsize=(12, fig_height), dpi=TABLE_EXPORT_DPI)
    ax.axis("off")

    table = ax.table(
        cellText=[header_bottom] + data,
        colLabels=None,
        cellLoc="center",
        loc="center",
    )

    table.auto_set_font_size(False)
    table.set_fontsize(13)
    table.scale(1.0, 1.25)

    widths = [0.20, 0.15, 0.16, 0.13, 0.13, 0.15, 0.08]
    for i, w in enumerate(widths):
        table.auto_set_column_width(i)
        table._cells[(0, i)].set_width(w)
    for (r_idx, c_idx), cell in table._cells.items():
        cell.set_edgecolor("black")
        cell.set_linewidth(1.5 if r_idx == 0 else 1.2)
        cell.PAD = 0.08 if r_idx == 0 else 0.06

    for col in range(len(header_bottom)):
        cell = table[0, col]
        cell.set_facecolor(COLOR_HEADER)
        cell.set_text_props(color="white", weight="bold")

    start_body = 1
    for r_idx in range(start_body, len(data) + start_body):
        row_model = rows[r_idx - start_body]
        table[r_idx, 1].set_facecolor(COLOR_VAR_BG)
        table[r_idx, 1].set_text_props(color=_metric_text_color(row_model.var_fabricante), weight="bold")
        table[r_idx, 2].set_facecolor(COLOR_VAR_BG)
        table[r_idx, 2].set_text_props(color=_metric_text_color(row_model.mat_sep25), weight="bold")
        table[r_idx, 3].set_facecolor(COLOR_C)
        table[r_idx, 3].set_text_props(color=_metric_text_color(row_model.lim_inf), weight="bold")
        table[r_idx, 4].set_facecolor(COLOR_D)
        table[r_idx, 4].set_text_props(color=_metric_text_color(row_model.lim_sup), weight="bold")
        table[r_idx, 5].set_facecolor(COLOR_G)
        err_val = rows[r_idx - start_body].error_muestral
        err_color = _error_cell_color(err_val)
        table[r_idx, 6].set_facecolor(err_color)
        table[r_idx, 6].set_text_props(color=_text_color_for_bg(err_color))
        brand_label = rows[r_idx - start_body].marca
        brand_color = assign_brand_palette_color(brand_label, brand_color_lookup, TREND_COLOR_SEQUENCE)
        if brand_color:
            table[r_idx, 0].set_text_props(color=brand_color, weight="normal")

    plt.tight_layout()
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    table_bbox = table.get_tightbbox(renderer).transformed(fig.dpi_scale_trans.inverted())
    fig_size = fig.get_size_inches()
    bbox_width = float(table_bbox.width) if table_bbox.width else float(fig_size[0])
    bbox_height = float(table_bbox.height) if table_bbox.height else float(fig_size[1])
    return (
        figure_to_stream(
            fig,
            dpi=TABLE_EXPORT_DPI,
            bbox_inches=table_bbox,
            pad_inches=0.02,
            transparent=False,
        ),
        (bbox_width, bbox_height),
    )

def add_ppt8_table_to_slide(
    slide, rows: List[Ppt8Row], brand_color_lookup: Dict[str, str],
    manufacturer_name: str, left, top, width_emu, max_height_in: Optional[float] = None,
):
    """Editable unified comparison, including the shared soft double error signal."""
    colors = [assign_brand_palette_color(row.marca, brand_color_lookup, TREND_COLOR_SEQUENCE)
              or confidence_design.INK for row in rows]
    return confidence_design.add_table(slide, rows, colors, left, top, width_emu, max_height_in)

def _set_pandas_option_compat(option_name: str, value) -> None:
    """Configura opciones de pandas sin fallar cuando no existen o estan deprecadas."""
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            pd.set_option(option_name, value)
    except Exception:
        # Mantiene compatibilidad entre versiones con diferencias de opciones.
        pass

def _df_map_compat(df: pd.DataFrame, formatter) -> pd.DataFrame:
    """Aplica una funcion a cada celda con compatibilidad entre versiones de pandas."""
    if hasattr(df, "map"):
        return df.map(formatter)
    if hasattr(df, "applymap"):
        return df.applymap(formatter)
    return df.apply(lambda col: col.map(formatter))

_set_pandas_option_compat('future.no_silent_downcasting', True)
_set_pandas_option_compat('mode.chained_assignment', None)
# Segmento 1: utilidades para el grafico MAT (cuando).
# Funcion que prepara los datos para la creacion del grafico MAT
def df_mat(df,p):
    """Prepara el dataframe MAT con acumulados y variaciones."""
    v1 = pd.DataFrame([(df.iloc[i-12-p:i-p,1].sum()/df.iloc[i-24-p:i-12-p,1].sum()) - 1 if i >= 24 else np.nan for i in range(12, len(df)+1)],columns=['Var Sell-in'])
    v2 = pd.DataFrame([(df.iloc[i-12:i,2].sum()/df.iloc[i-24:i-12,2].sum()) - 1 if i >= 24 else np.nan for i in range(12, len(df)+1)],columns=['Var Sell-out'])
    ac1, ac2 = [pd.DataFrame(df[col].rolling(window=12).sum().dropna()) for col in df.columns[1:]]
    v1.index,v2.index=v1.index+11,v2.index+11
    v1 ,v2 = v1.dropna(), v2.dropna()
    mat =pd.DataFrame(df.iloc[:,0].copy())
    mat= mat.join(ac1.join(ac2)[-len(v1):].join([v1]).join([v2]))
    mat= mat[mat.notna().all(axis=1).idxmax():]
    return mat
#Grafico de MAT
def graf_mat(mat, c_fig, p):
    """Genera el grafico MAT y devuelve una imagen en memoria."""
    altura = 7  # Aumentamos aun mas la altura
    fig, ax = plt.subplots(num=c_fig, figsize=(15, altura), dpi=DEFAULT_EXPORT_DPI)
    fig.subplots_adjust(bottom=0.2, top=0.9)
    ran = [x.strftime('%m-%y') for x in mat.iloc[:, 0].copy()]
    ac1 = mat.iloc[:, 1].copy()
    ac2 = mat.iloc[:, 2].copy()
    v1 = mat.iloc[:, 3].copy()
    v2 = mat.iloc[:, 4].copy()
    l1 = ax.plot(ran, ac1, color=LINE_COLOR_CLIENT, linewidth=4.2, label='Acumulado Cliente')
    l2 = ax.plot(ran, ac2, color=LINE_COLOR_NUMERATOR, linewidth=4.2, label='Acumulado Numerator')
    ax2 = ax.twinx()
    b1 = ax2.bar(
        np.arange(len(ran)) - 0.3,
        v1.values,
        0.3,
        color=BAR_COLOR_VARIATION_CLIENT,
        edgecolor=BAR_EDGE_COLOR,
        label='Var % Cliente Pipeline: ' + str(p) + ' ' + labels[(lang, 'Var MAT')],
    )
    b2 = ax2.bar(
        np.arange(len(ran)) + 0.3,
        v2.values,
        0.3,
        color=BAR_COLOR_VARIATION_NUMERATOR,
        edgecolor=BAR_EDGE_COLOR,
        label='Var % Numerator ' + labels[(lang, 'Var MAT')],
    )
    ax.set_axisbelow(True)
    ax.grid(axis='y', linestyle='--', alpha=0.28, color='#D9D9D9', linewidth=0.9)
    ax.set_xticks(np.arange(len(ran)))
    ax.set_xticklabels(ran, rotation=30)
    if len(ran) > 12:
        for idx in range(len(ran) - 13, -1, -12):
            ax.axvline(
                x=idx,
                color="#B0B0B0",
                linestyle="--",
                linewidth=1.1,
                alpha=0.35,
                zorder=0,
            )
    ax2.tick_params(
        left=False,
        labelleft=False,
        top=False,
        labeltop=False,
        right=False,
        labelright=False,
        bottom=False,
        labelbottom=False,
    )
    bar_positions = np.arange(len(ran))
    label_specs = [
        (v1, -0.3, BAR_LABEL_COLOR_POS, BAR_LABEL_COLOR_NEG),
        (v2, 0.3, BAR_LABEL_COLOR_POS_ALT, BAR_LABEL_COLOR_NEG_ALT),
    ]
    for serie, x_offset, color_pos, color_neg in label_specs:
        for idx, y in enumerate(serie):
            if not np.isfinite(y):
                continue
            y_pos = y + 0.01 if y >= 0 else y - 0.01  # leve desplazamiento para dejar espacio
            va_align = 'bottom' if y >= 0 else 'top'
            ax2.text(
                bar_positions[idx] + x_offset,
                y_pos,
                f"{y*100:.1f}%",
                ha='center',
                va=va_align,
                fontsize=8,
                fontweight='bold',
                color=color_pos if y >= 0 else color_neg,
            )
    y2_min, y2_max = ax2.get_ylim()
    padding = max(abs(y2_min), abs(y2_max)) * 0.15 if np.isfinite(y2_min) and np.isfinite(y2_max) else 0
    ax2.set_ylim(y2_min - padding, y2_max + padding * 2)
    lns = l1 + l2 + [b1, b2]
    labs = [l.get_label() for l in l1 + l2] + [b1.get_label(), b2.get_label()]
    ax.legend(
        lns,
        labs,
        loc='upper center',
        bbox_to_anchor=(0.5, -0.15),
        borderaxespad=0.05,
        frameon=True,
        prop={'size': 12},
        ncol=2,
    )
    ax.set_title(labels[(lang, 'MAT')] + ' | ' + w[2:], size=18, pad=20)
    fig.tight_layout(rect=[0, 0.05, 1, 0.95])
    return figure_to_stream(fig, bbox_inches=None)
#Grafico de Lineas

def format_title_with_bold_suffix(base_title: str, suffix: str) -> str:
    """Devuelve un titulo con sufijo en negrita usando mathtext."""
    base_title = (base_title or '').strip()
    suffix = (suffix or '').strip()
    if not suffix:
        return base_title
    safe_suffix = suffix.replace('\\', '\\\\').replace('$', '\\$')
    spacer = ' ' if base_title else ''
    return f"{base_title}{spacer}$\\bf{{{safe_suffix}}}$"


MONTH_OVER_MONTH_MAX_PERIODS = 12
MONTH_OVER_MONTH_DENSE_SERIES_THRESHOLD = 8
MONTH_OVER_MONTH_DENSE_SERIES_LIMIT = 5
MONTH_OVER_MONTH_POSITIVE_COLOR = '#1E8449'
MONTH_OVER_MONTH_NEGATIVE_COLOR = '#C0392B'
MONTH_OVER_MONTH_NEUTRAL_COLOR = '#666666'


def _format_month_over_month_change(value: float) -> str:
    """Formatea una variación mensual porcentual con signo explícito."""
    if math.isclose(value, 0.0, abs_tol=0.05):
        return '0.0%'
    decimals = 0 if abs(value) >= 100 else 1
    return f"{value:+.{decimals}f}%"


def _build_month_over_month_annotations(
    series_points: dict,
    visible_period_count: int,
) -> list[dict]:
    """Construye etiquetas MoM entre puntos consecutivos para gráficos cortos."""
    if visible_period_count < 2 or visible_period_count > MONTH_OVER_MONTH_MAX_PERIODS:
        return []
    series_items = list(series_points.items())
    original_order = {column: order for order, (column, _) in enumerate(series_items)}
    if len(series_items) > MONTH_OVER_MONTH_DENSE_SERIES_THRESHOLD:
        ranked_series = []
        for column, raw_points in series_items:
            finite_values = [
                float(y_value)
                for _, y_value in raw_points
                if np.isfinite(y_value)
            ]
            if finite_values:
                ranked_series.append((column, abs(finite_values[-1])))
        ranked_series.sort(key=lambda item: item[1], reverse=True)
        selected_columns = {
            column for column, _ in ranked_series[:MONTH_OVER_MONTH_DENSE_SERIES_LIMIT]
        }
        series_items = [item for item in series_items if item[0] in selected_columns]
    annotations: list[dict] = []
    for column, raw_points in series_items:
        series_order = original_order[column]
        points = sorted(
            (
                (float(x_value), float(y_value))
                for x_value, y_value in raw_points
                if np.isfinite(x_value) and np.isfinite(y_value)
            ),
            key=lambda item: item[0],
        )
        for interval_index, (previous_point, current_point) in enumerate(zip(points, points[1:])):
            previous_x, previous_y = previous_point
            current_x, current_y = current_point
            if not math.isclose(current_x - previous_x, 1.0, rel_tol=0.0, abs_tol=1e-6):
                continue
            if math.isclose(previous_y, 0.0, rel_tol=0.0, abs_tol=1e-12):
                continue
            variation = ((current_y - previous_y) / abs(previous_y)) * 100.0
            if not np.isfinite(variation):
                continue
            annotations.append(
                {
                    'column': column,
                    'series_order': series_order,
                    'interval_index': interval_index,
                    'x': (previous_x + current_x) / 2.0,
                    'y': (previous_y + current_y) / 2.0,
                    'variation': variation,
                    'text': _format_month_over_month_change(variation),
                }
            )
    return annotations

def _trend_legend_columns(labels, width_inches, font_points, default_columns, figure):
    """Distribuye leyendas densas según el ancho real del texto, sin quitar series."""
    if len(labels) <= 12:
        return max(1, min(len(labels), default_columns))
    renderer = figure.canvas.get_renderer()
    font = FontProperties(size=font_points)
    label_widths = [max(renderer.get_text_width_height_descent(line, font, False)[0]
                        for line in str(label).split('\n')) for label in labels]
    point = figure.dpi / 72.0
    handle_width = (plt.rcParams['legend.handlelength'] + plt.rcParams['legend.handletextpad']) * font_points * point
    column_gap = plt.rcParams['legend.columnspacing'] * font_points * point
    border = 2 * plt.rcParams['legend.borderpad'] * font_points * point
    available = max(1, (width_inches - 0.08) * figure.dpi)
    for columns in range(min(12, len(labels)), 0, -1):
        count, extra = divmod(len(labels), columns)
        offset, widths = 0, []
        for index in range(columns):
            length = count + (index < extra)
            widths.append(max(label_widths[offset:offset + length]) + handle_width)
            offset += length
        if sum(widths) + column_gap * (columns - 1) + border <= available:
            return columns
    return 1


def _trend_legend_height_inches(ax, handles, labels, columns, font_points):
    if len(labels) <= 12:
        return font_points * 1.35 / 72.0 * math.ceil(len(labels) / columns)
    from matplotlib.legend import Legend
    legend = Legend(ax, handles, labels, ncols=columns, prop={'size': font_points})
    return legend.get_window_extent(ax.figure.canvas.get_renderer()).height / ax.figure.dpi + 0.02


def line_graf(
    df,
    p,
    title,
    c_fig,
    ven,
    width_emu=None,
    height_emu=None,
    multi_chart=None,
    share_lookup=None,
    y_axis_percent=False,
    top_value_annotations=0,
    color_collector=None,
    color_overrides=None,
    linestyle_overrides=None,
    show_title=True,
    include_origin_suffix=True,
    force_right_labels=False,
):
    """
    Renderiza la gráfica de tendencias y devuelve un buffer PNG listo para insertar en la slide.
    Cuando se generan múltiples gráficos en la misma diapositiva (multi_chart=True), se aplican
    márgenes más amplios para evitar recortes visuales.
    """
    if width_emu is not None:
        try:
            planned_width = emu_to_inches(width_emu)
        except Exception:
            planned_width = None
    else:
        planned_width = None
    if multi_chart is None:
        detected_multi = planned_width is not None and planned_width < 6.0
    else:
        detected_multi = bool(multi_chart)
    if detected_multi:
        chart_top_margin = LINE_CHART_MULTI_TOP_MARGIN
        chart_bottom_margin = LINE_CHART_MULTI_BOTTOM_MARGIN
        chart_left_margin = LINE_CHART_MULTI_LEFT_MARGIN
        chart_right_margin = LINE_CHART_MULTI_RIGHT_MARGIN
        chart_x_margin = LINE_CHART_MULTI_X_MARGIN
    else:
        chart_top_margin = LINE_CHART_TOP_MARGIN
        chart_bottom_margin = LINE_CHART_BOTTOM_MARGIN
        chart_left_margin = LINE_CHART_LEFT_MARGIN
        chart_right_margin = LINE_CHART_RIGHT_MARGIN
        chart_x_margin = LINE_CHART_SINGLE_X_MARGIN
    # Ajusta o tamanho da figura para combinar com o espaco reservado no slide
    width_inches = emu_to_inches(width_emu) if width_emu is not None else None
    height_inches = emu_to_inches(height_emu) if height_emu is not None else None
    if width_inches is None and height_inches is None:
        # Ancho dinámico según cantidad de datos (en pulgadas)
        n_points = len(df)
        ancho = max(15, n_points * 0.5)  # 0.5 puede ajustarse para más/menos espacio (pulgadas)
        # Altura fija de 10 cm -> convertir a pulgadas
        altura = emu_to_inches(Cm(10))
        figsize = (ancho, altura)
    else:
        if width_inches is None:
            width_inches = height_inches * DEFAULT_LINE_CHART_RATIO
        elif height_inches is None:
            height_inches = width_inches / DEFAULT_LINE_CHART_RATIO
        figsize = (width_inches, height_inches)
    fig, ax = plt.subplots(num=c_fig, figsize=figsize, dpi=DEFAULT_EXPORT_DPI)
    # Escalar elementos (fuentes, linewidth) según la altura para que el contenido se adapte
    # Se toma como referencia la altura original de 5 pulgadas usada previamente
    ref_height = 5.0
    actual_height = figsize[1]
    scale = actual_height / ref_height if ref_height else 1.0
    base_linewidth = 2.0 * max(0.6, scale)
    title_base_size = max(10, int(18 * scale))
    legend_base_size = max(8, int(10 * scale))
    xtick_size = max(6, int(8 * scale))
    xtick_font = FontProperties(family='DejaVu Sans', size=xtick_size, weight='regular')
    xtick_pad = max(6, int(8 * scale))
    marker_size = max(3.5, 4.5 * scale)
    share_lookup = share_lookup or {}
    aux = df.copy()
    if aux.shape[1] > 1:
        kept_columns = [aux.columns[0]]
        for col in aux.columns[1:]:
            series = pd.to_numeric(aux[col], errors='coerce')
            if series.replace(0, np.nan).dropna().empty:
                continue
            kept_columns.append(col)
        if len(kept_columns) == 1:
            aux = aux.iloc[:, :1]
        else:
            aux = aux.loc[:, kept_columns]
    if aux.empty or aux.shape[1] <= 1:
        return figure_to_stream(fig, bbox_inches=None)
    parsed_periods, _ = _parse_flexible_date_series(aux.iloc[:, 0])
    invalid_period_mask = parsed_periods.isna()
    if invalid_period_mask.any():
        invalid_count = int(invalid_period_mask.sum())
        chart_context = f" del grafico '{title}'" if title else " del grafico"
        print_colored(
            f"Se omitieron {invalid_count} fila(s) con periodo vacio o invalido{chart_context}.",
            COLOR_RED,
        )
        aux = aux.loc[~invalid_period_mask].copy().reset_index(drop=True)
        parsed_periods = parsed_periods.loc[~invalid_period_mask].reset_index(drop=True)
    if aux.empty:
        return figure_to_stream(fig, bbox_inches=None)
    aux.iloc[:, 0] = parsed_periods
    ran = [x.strftime('%m-%y') for x in parsed_periods]
    data_len = len(ran)
    if data_len == 0:
        return figure_to_stream(fig, bbox_inches=None)
    start_idx = max(0, min(p, data_len - 1))
    visible_period_count = max(0, data_len - start_idx)
    x_positions = np.arange(data_len)
    colunas = list(aux.columns[1:])
    def _is_sales_series(label: str) -> bool:
        normalized = str(label).strip().lower()
        if not normalized:
            return False
        if normalized.endswith(('.v', '_v', '-v')):
            return True
        return 'venta' in normalized or 'sell-in' in normalized or 'sellin' in normalized
    def _is_purchase_series(label: str) -> bool:
        normalized = str(label).strip().lower()
        if not normalized:
            return False
        if normalized.endswith(('.c', '_c', '-c')):
            return True
        return 'compra' in normalized or 'sell-out' in normalized or 'sellout' in normalized
    has_sales_markers = any(_is_sales_series(col) for col in colunas)
    has_purchase_markers = any(_is_purchase_series(col) for col in colunas)
    palette_values = TREND_COLOR_SEQUENCE or [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]
    color_mapping: dict[str, str] = {}
    color_key_lookup: dict[str, str] = {}
    chart_color_owners: dict[str, set[str]] = {}
    pending_reserved_labels: dict[str, Optional[str]] = {}
    pending_reserved_counts: dict[str, int] = {}
    preassigned_colors: dict[str, str] = {}
    def _register_color_keys(label: str, color_value: str, key_list: Optional[list[str]] = None) -> None:
        keys = key_list if key_list is not None else generate_color_lookup_keys(label)
        for key in keys:
            if key and key not in color_key_lookup:
                color_key_lookup[key] = color_value
    def _find_preferred_color(label: str, key_list: Optional[list[str]] = None) -> Optional[str]:
        existing_color = color_mapping.get(label)
        if existing_color:
            return existing_color
        # Un color explicito heredado de otro grafico debe prevalecer sobre
        # cualquier variante inferida por coincidencia semantica.
        planned_color = preassigned_colors.get(label)
        if planned_color:
            return planned_color
        keys = key_list if key_list is not None else generate_color_lookup_keys(label)
        for key in keys:
            reused_color = color_key_lookup.get(key)
            if reused_color:
                return reused_color
        return _lookup_semantic_mapping_value(label, color_mapping, allow_family=True)
    def _release_reserved_color(label: str) -> None:
        reserved_color = pending_reserved_labels.pop(label, None)
        if not reserved_color:
            return
        remaining = pending_reserved_counts.get(reserved_color, 0) - 1
        if remaining > 0:
            pending_reserved_counts[reserved_color] = remaining
        else:
            pending_reserved_counts.pop(reserved_color, None)
    def _assign_chart_color(label: str, color_value: str, key_list: Optional[list[str]] = None) -> str:
        color_mapping[label] = color_value
        _register_color_keys(label, color_value, key_list)
        owners = chart_color_owners.setdefault(color_value, set())
        owners.add(label)
        _release_reserved_color(label)
        return color_value
    def _color_is_available(color_value: str, label: str) -> bool:
        owners = chart_color_owners.get(color_value)
        return not owners or label in owners
    if color_overrides:
        for name, hex_color in color_overrides.items():
            color_mapping[name] = hex_color
            _register_color_keys(name, hex_color)
    current_label_set = {str(label).strip() for label in colunas if str(label).strip()}
    reference_labels_for_family = [
        label
        for label in color_mapping.keys()
        if str(label).strip() and str(label).strip() not in current_label_set
    ]
    if reference_labels_for_family:
        family_groups: dict[str, list[dict[str, object]]] = {}
        for original_index, label in enumerate(colunas):
            family_key = _extract_semantic_family_key(label, reference_labels_for_family)
            if not family_key:
                continue
            best_match = _best_semantic_reference_match(label, reference_labels_for_family, allow_family=True)
            anchor_color = None
            anchor_score = float('-inf')
            if best_match is not None:
                anchor_color = color_mapping.get(best_match.reference_label)
                anchor_score = best_match.score
            family_groups.setdefault(family_key, []).append(
                {
                    "label": label,
                    "order": original_index,
                    "anchor_color": anchor_color,
                    "anchor_score": anchor_score,
                }
            )
        for family_key, family_entries in family_groups.items():
            if not family_entries:
                continue
            anchored_entries = [
                entry for entry in family_entries
                if entry.get("anchor_color") and entry.get("anchor_score", float('-inf')) != float('-inf')
            ]
            if anchored_entries:
                anchored_entries.sort(key=lambda entry: (entry["anchor_score"], -entry["order"]), reverse=True)
                base_color = anchored_entries[0]["anchor_color"]
            else:
                base_color = _select_ambiguous_family_base_color(family_key)
            family_entries.sort(key=lambda entry: entry["order"])
            if len(family_entries) == 1 and anchored_entries:
                preassigned_colors[str(family_entries[0]["label"])] = str(base_color)
                continue
            family_palette = _build_family_variant_palette(str(base_color), len(family_entries))
            for idx_family, entry in enumerate(family_entries):
                label_key = str(entry["label"])
                preassigned_colors[label_key] = family_palette[idx_family] if idx_family < len(family_palette) else str(base_color)
    for label in colunas:
        keys = generate_color_lookup_keys(label)
        preferred_color = _find_preferred_color(label, keys)
        pending_reserved_labels[label] = preferred_color
        if preferred_color:
            pending_reserved_counts[preferred_color] = pending_reserved_counts.get(preferred_color, 0) + 1
    color_index = 0
    def _next_palette_color() -> str:
        nonlocal color_index
        if not palette_values:
            return '#1F77B4'
        palette_len = len(palette_values)
        attempts = 0
        while attempts < palette_len:
            candidate = palette_values[color_index % palette_len]
            color_index += 1
            attempts += 1
            if candidate not in chart_color_owners:
                if pending_reserved_counts.get(candidate, 0) > 0:
                    continue
                return candidate
        candidate = palette_values[color_index % palette_len]
        color_index += 1
        return candidate
    def _resolve_or_assign_color(label: str) -> str:
        keys = generate_color_lookup_keys(label)
        existing_color = color_mapping.get(label)
        if existing_color and _color_is_available(existing_color, label):
            return _assign_chart_color(label, existing_color, keys)
        planned_color = preassigned_colors.get(label)
        if planned_color and _color_is_available(planned_color, label):
            return _assign_chart_color(label, planned_color, keys)
        for key in keys:
            reused_color = color_key_lookup.get(key)
            if reused_color and _color_is_available(reused_color, label):
                return _assign_chart_color(label, reused_color, keys)
        semantic_color = _lookup_semantic_mapping_value(label, color_mapping, allow_family=True)
        if semantic_color and _color_is_available(semantic_color, label):
            return _assign_chart_color(label, semantic_color, keys)
        assigned_color = _next_palette_color()
        return _assign_chart_color(label, assigned_color, keys)
    lns = []
    legend_labels = []
    numeric_series_list = []
    plotted_columns = []
    series_points = {}
    for col in colunas:
        if ven > 1:
            estilo = '-' if '.v' in col.lower() else '--'
        else:
            estilo = '-'
        cor = _resolve_or_assign_color(col)
        numeric_series = pd.to_numeric(aux[col], errors='coerce')
        numeric_series_list.append(numeric_series)
        y = numeric_series.values
        if start_idx > 0:
            if has_sales_markers:
                apply_pipeline_shift = _is_sales_series(col)
            else:
                # Si no hay marcadores mixtos, asumimos grafico de ventas/sell-in.
                apply_pipeline_shift = not has_purchase_markers
        else:
            apply_pipeline_shift = False
        if apply_pipeline_shift:
            if data_len <= start_idx:
                continue
            x_slice = x_positions[start_idx:]
            y_slice = y[:data_len - start_idx]
        else:
            x_slice = x_positions[start_idx:]
            y_slice = y[start_idx:]
        valid_points = [(x_idx, y_val) for x_idx, y_val in zip(x_slice, y_slice) if pd.notna(y_val)]
        if not valid_points:
            continue
        series_points[col] = valid_points
        suffix_map = {".c": "Compras", ".v": "Ventas", "_c": "Compras", "_v": "Ventas", "-c": "Compras", "-v": "Ventas"}
        lower_col = col.lower()
        legend_label = col
        for suffix, suffix_text in suffix_map.items():
            if lower_col.endswith(suffix):
                base_label = col[:-len(suffix)].strip()
                if base_label:
                    legend_label = f"{base_label} ({suffix_text})" if include_origin_suffix else base_label
                break
        if linestyle_overrides and col in linestyle_overrides:
            estilo = linestyle_overrides[col]
        line, = ax.plot(
            x_slice,
            y_slice,
            color=cor,
            linewidth=base_linewidth,
            linestyle=estilo,
            label=col,
            marker='o',
            markersize=marker_size,
            markerfacecolor='white',
            markeredgewidth=max(0.7, base_linewidth / 2),
            markeredgecolor=cor,
        )
        lns.append(line)
        legend_labels.append(legend_label)
        plotted_columns.append(col)
    if data_len and start_idx < data_len:
        ax.set_xticks(x_positions[start_idx:])
        ax.set_xticklabels(
            ran[start_idx:],
            rotation=90,
            fontproperties=xtick_font,
            ha='center',
            va='center',
        )
        ax.tick_params(axis='x', pad=xtick_pad)
        ax.set_xlim(x_positions[start_idx], x_positions[-1])
        if data_len - start_idx > 12:
            for idx in range(data_len - 13, start_idx - 1, -12):
                ax.axvline(
                    x=idx,
                    color="#B0B0B0",
                    linestyle="--",
                    linewidth=1.1,
                    alpha=0.35,
                    zorder=0,
                )
    else:
        ax.set_xticks([])
        ax.set_xticklabels([])
    ax.set_axisbelow(True)
    ax.grid(axis='y', linestyle='--', alpha=0.28, color='#D9D9D9', linewidth=0.9)
    y_tick_size = max(8, int(10 * scale))
    ax.tick_params(axis='y', labelsize=y_tick_size)
    axis_percent = bool(y_axis_percent)
    max_abs_value = 0.0
    for series in numeric_series_list:
        numeric_values = series.to_numpy()
        if numeric_values.size == 0:
            continue
        finite_values = numeric_values[np.isfinite(numeric_values)]
        if finite_values.size == 0:
            continue
        series_max = float(np.max(np.abs(finite_values)))
        if series_max > max_abs_value:
            max_abs_value = series_max
    if axis_percent:
        def _trend_tick_formatter(value, _):
            return f"{value:.1f}%"
        ax.yaxis.set_major_formatter(FuncFormatter(_trend_tick_formatter))
    else:
        scale_rule = select_trend_scale_rule(max_abs_value)
        divisor = scale_rule["divisor"] or 1
        suffix = scale_rule["suffix"]
        def _trend_tick_formatter(value, _):
            return format_value_with_suffix(value, divisor, suffix)
        ax.yaxis.set_major_formatter(FuncFormatter(_trend_tick_formatter))
        if DISPLAY_TREND_REFERENCE_TEXT and suffix:
            reference_template = scale_rule.get("reference", {})
            reference_text = reference_template.get(lang, reference_template.get('default', '')).strip()
            if reference_text:
                reference_font_size = max(8, int(9 * scale))
                ax.text(
                    0.02,
                    0.95,
                    reference_text,
                    transform=ax.transAxes,
                    ha='left',
                    va='top',
                    fontsize=reference_font_size,
                    color='#4D4D4D',
                    alpha=0.85,
                )
    if axis_percent:
        def _format_axis_value(value):
            return f"{value:.1f}%"
    else:
        def _format_axis_value(value):
            return format_value_with_suffix(value, divisor, suffix)
    month_over_month_annotations = _build_month_over_month_annotations(
        series_points,
        visible_period_count,
    )
    ax.set_ylim(bottom=0)
    # Margen de los datos dentro del eje
    ax.margins(x=chart_x_margin, y=0.14 if month_over_month_annotations else 0.08)
    if month_over_month_annotations:
        dense_annotation_chart = detected_multi or len(plotted_columns) > 8
        if dense_annotation_chart:
            annotation_font_size = max(6.2, min(7.2, 8.0 * scale))
            minimum_label_gap_fraction = 0.055
        else:
            annotation_font_size = max(7.0, min(9.0, 9.0 * scale))
            minimum_label_gap_fraction = 0.045
        y_limits = ax.get_ylim()
        y_span = max(y_limits[1] - y_limits[0], 1.0)
        label_padding = y_span * 0.035
        minimum_label_gap = y_span * minimum_label_gap_fraction
        lower_label_limit = y_limits[0] + label_padding
        upper_label_limit = y_limits[1] - label_padding
        annotations_by_midpoint = {}
        for annotation in month_over_month_annotations:
            midpoint_y = annotation['y']
            if midpoint_y < y_limits[0] or midpoint_y > y_limits[1]:
                continue
            midpoint_key = round(annotation['x'], 6)
            annotations_by_midpoint.setdefault(midpoint_key, []).append(annotation)

        positioned_annotations = []
        for midpoint_annotations in annotations_by_midpoint.values():
            midpoint_annotations.sort(key=lambda item: item['y'])
            target_positions = [
                min(max(annotation['y'], lower_label_limit), upper_label_limit)
                for annotation in midpoint_annotations
            ]
            for index in range(1, len(target_positions)):
                target_positions[index] = max(
                    target_positions[index],
                    target_positions[index - 1] + minimum_label_gap,
                )
            if target_positions and target_positions[-1] > upper_label_limit:
                target_positions[-1] = upper_label_limit
                for index in range(len(target_positions) - 2, -1, -1):
                    target_positions[index] = min(
                        target_positions[index],
                        target_positions[index + 1] - minimum_label_gap,
                    )
            positioned_annotations.extend(zip(midpoint_annotations, target_positions))

        for annotation, target_y in positioned_annotations:
            midpoint_y = annotation['y']
            variation = annotation['variation']
            if variation > 0.05:
                variation_color = MONTH_OVER_MONTH_POSITIVE_COLOR
            elif variation < -0.05:
                variation_color = MONTH_OVER_MONTH_NEGATIVE_COLOR
            else:
                variation_color = MONTH_OVER_MONTH_NEUTRAL_COLOR
            connector = None
            if abs(target_y - midpoint_y) > y_span * 0.012:
                connector = dict(
                    arrowstyle='-',
                    color=variation_color,
                    linewidth=0.45,
                    alpha=0.55,
                )
            ax.annotate(
                annotation['text'],
                xy=(annotation['x'], midpoint_y),
                xytext=(annotation['x'], target_y),
                textcoords='data',
                ha='center',
                va='center',
                fontsize=annotation_font_size,
                fontweight='bold',
                color=variation_color,
                bbox=dict(
                    boxstyle='round,pad=0.16',
                    facecolor='white',
                    edgecolor=variation_color,
                    linewidth=0.45,
                    alpha=0.82,
                ),
                arrowprops=connector,
                zorder=7,
                clip_on=True,
            )

    # --- Parámetros de separación en coordenadas de figura ---
    # para gráficos "sencillos" vs "multi" (muchas series)
    base_gap = 0.075 if not detected_multi else 0.070         # espacio visual típico
    min_axes_legend_gap = 0.070 if not detected_multi else 0.080  # mínimo absoluto entre eje y leyenda
    margin_buffer = 0.020                                      # margen al borde inferior de la figura
    inner_plot_lift = 0.040 if not detected_multi else 0.032   # eleva el area del eje sin mover la leyenda
    needs_tight_export = False

    legend_bottom_margin = chart_bottom_margin
    legend_columns = 1
    legend_rows = 1
    legend_height_fraction = 0.0

    if lns:
        # Cálculo de columnas/filas en la leyenda
        max_columns = 3 if detected_multi else 4
        legend_columns = _trend_legend_columns(legend_labels, figsize[0] * (chart_right_margin - chart_left_margin), legend_base_size, max_columns, fig)
        legend_rows = max(1, math.ceil(len(legend_labels) / legend_columns))

        # Altura de la leyenda en pulgadas -> fracción de la figura
        legend_font_points = legend_base_size
        legend_line_height_points = legend_font_points * 1.35
        legend_height_inches = _trend_legend_height_inches(ax, lns, legend_labels, legend_columns, legend_font_points)
        figure_height_inches = fig.get_size_inches()[1]
        legend_height_fraction = (
            legend_height_inches / figure_height_inches if figure_height_inches else 0.0
        )

        # Espacio mínimo entre eje y leyenda
        legend_clearance = max(base_gap, min_axes_legend_gap)

        # Aseguramos espacio suficiente abajo para leyenda + gap + buffer
        legend_bottom_margin = max(
            chart_bottom_margin,
            legend_height_fraction + legend_clearance + margin_buffer,
        )
    else:
        legend_clearance = 0.0  # no hay leyenda, no hace falta espacio extra

    # --- Márgenes superior/inferior del eje ---
    effective_top_margin = chart_top_margin
    if not (show_title and title):
        # Si no hay título, no "regalemos" tanto espacio arriba
        effective_top_margin = min(0.97, chart_top_margin + 0.05)
    elif title:
        plt.title(title, size=title_base_size, pad=10)

    axis_bottom_margin = min(0.9, max(0.05, legend_bottom_margin + inner_plot_lift))
    if axis_bottom_margin >= effective_top_margin:
        axis_bottom_margin = max(0.05, effective_top_margin - 0.05)

    fig.subplots_adjust(
        top=effective_top_margin,
        bottom=axis_bottom_margin,
        left=chart_left_margin,
        right=chart_right_margin,
    )

    # --- Posición final de la leyenda ---
    if lns:
        axes_box = ax.get_position()  # caja del eje en coordenadas de figura

        # Queremos que la parte superior de la leyenda quede legend_clearance
        # por debajo del borde inferior del eje.
        legend_top = max(margin_buffer + legend_height_fraction, legend_bottom_margin - legend_clearance)

        legend = ax.legend(
            lns,
            legend_labels,
            loc="upper left",
            bbox_to_anchor=(axes_box.x0, legend_top, axes_box.width, 0.0),
            bbox_transform=fig.transFigure,
            borderaxespad=0.0,
            frameon=True,
            prop={"size": legend_base_size},
            ncol=legend_columns,
            mode="expand",
        )

        frame = legend.get_frame()
        frame.set_facecolor("white")
        frame.set_edgecolor("#D3D3D3")
        frame.set_alpha(0.85)

    if plotted_columns and (share_lookup or force_right_labels):
        legend_label_map = {col: label for col, label in zip(plotted_columns, legend_labels)}
        annotation_candidates = []
        share_lookup_lower = {str(key).lower(): value for key, value in share_lookup.items()} if share_lookup else {}
        for col, line in zip(plotted_columns, lns):
            points = series_points.get(col)
            if not points:
                continue
            last_x, last_y = points[-1]
            share_value = share_lookup.get(col) if share_lookup else None
            if share_value is None and share_lookup_lower:
                share_value = share_lookup_lower.get(str(col).lower())
            annotation_candidates.append({
                "column": col,
                "line": line,
                "legend": legend_label_map.get(col, col),
                "share": share_value,
                "last_x": last_x,
                "last_y": last_y,
                "color": line.get_color(),
            })
        total_series = len(annotation_candidates)
        if total_series:
            if total_series < 10:
                selected_series = annotation_candidates
            else:
                with_share = [info for info in annotation_candidates if info["share"] is not None]
                without_share = [info for info in annotation_candidates if info["share"] is None]
                with_share.sort(key=lambda info: info["share"], reverse=True)
                if force_right_labels:
                    without_share.sort(key=lambda info: info["last_y"], reverse=True)
                    selected_series = []
                    if with_share:
                        selected_series.extend(with_share[:3])
                    if without_share:
                        selected_series.extend(without_share[:2])
                    if len(selected_series) < 5:
                        remaining = [info for info in annotation_candidates if info not in selected_series]
                        remaining.sort(key=lambda info: info["last_y"], reverse=True)
                        selected_series.extend(remaining[: max(0, 5 - len(selected_series))])
                else:
                    selected_series = with_share[:5]
                    if len(selected_series) < 5:
                        remaining = [info for info in annotation_candidates if info not in selected_series]
                        selected_series.extend(remaining[: max(0, 5 - len(selected_series))])
            if selected_series:
                needs_tight_export = True
                x_limits = ax.get_xlim()
                y_limits = ax.get_ylim()
                y_range = y_limits[1] - y_limits[0] if y_limits[1] > y_limits[0] else 1.0
                min_gap = y_range * 0.04
                min_margin = y_range * 0.02
                x_range = x_limits[1] - x_limits[0] if x_limits[1] > x_limits[0] else 1.0
                x_offset = max(0.8, x_range * 0.06)
                arranged_series = sorted(selected_series, key=lambda info: info["last_y"], reverse=True)
                placed_positions = []
                y_max = y_limits[1] - min_margin
                y_min = y_limits[0] + min_margin
                for info in arranged_series:
                    y_pos = max(min(info["last_y"], y_max), y_min)
                    for prev_y in placed_positions:
                        if y_pos > prev_y - min_gap:
                            y_pos = prev_y - min_gap
                    y_pos = max(y_pos, y_min)
                    placed_positions.append(y_pos)
                    info["label_y"] = y_pos
                    info["label_x"] = info["last_x"] + x_offset
                max_label_x = max(info["label_x"] for info in arranged_series)
                if max_label_x > x_limits[1]:
                    ax.set_xlim(x_limits[0], max_label_x + x_offset * 1.15)
                size_factor = 0.7 if detected_multi else 0.8
                min_size = 6 if detected_multi else 7
                label_font_size = max(min_size, int(legend_base_size * size_factor))
                bbox_props = dict(facecolor='white', edgecolor='none', alpha=0.85, boxstyle='round,pad=0.2')
                for info in arranged_series:
                    text_value = info["legend"]
                    ax.text(
                        info["label_x"],
                        info["label_y"],
                        text_value,
                        ha='left',
                        va='center',
                        color=info["color"],
                        fontsize=label_font_size,
                        fontweight='bold',
                        bbox=bbox_props,
                        clip_on=False,
                    )
    if top_value_annotations and plotted_columns and lns:
        point_candidates = []
        line_lookup = {col: line for col, line in zip(plotted_columns, lns)}
        for col in plotted_columns:
            if str(col).strip().lower() == 'total':
                continue
            line = line_lookup.get(col)
            if line is None:
                continue
            for x_idx, y_val in series_points.get(col, []):
                if not np.isfinite(y_val):
                    continue
                point_candidates.append({
                    "column": col,
                    "line": line,
                    "x": x_idx,
                    "y": float(y_val),
                })
        if point_candidates:
            point_candidates.sort(key=lambda item: item["y"], reverse=True)
            y_limits = ax.get_ylim()
            y_range = y_limits[1] - y_limits[0] if y_limits[1] > y_limits[0] else 1.0
            label_offset = max(y_range * 0.03, 0.5 if axis_percent else y_range * 0.03)
            label_font_size = max(8, int(10 * scale))
            annotated = 0
            seen_points = set()
            for point in point_candidates:
                point_key = (point["x"], round(point["y"], 6))
                if point_key in seen_points:
                    continue
                seen_points.add(point_key)
                label_y = point["y"] + label_offset
                if label_y > y_limits[1]:
                    label_y = min(y_limits[1], point["y"] + label_offset * 0.6)
                text_value = _format_axis_value(point["y"])
                ax.annotate(
                    text_value,
                    xy=(point["x"], point["y"]),
                    xytext=(point["x"], label_y),
                    textcoords='data',
                    ha='center',
                    va='bottom',
                    fontsize=label_font_size,
                    color=point["line"].get_color(),
                    fontweight='bold',
                    bbox=dict(
                        boxstyle='round,pad=0.2',
                        facecolor='white',
                        edgecolor=point["line"].get_color(),
                        linewidth=0.6,
                        alpha=0.85
                    ),
                    arrowprops=dict(
                        arrowstyle='-',
                        color=point["line"].get_color(),
                        linewidth=max(0.6, base_linewidth * 0.5)
                    ),
                    clip_on=False,
                )
                annotated += 1
                if annotated >= top_value_annotations:
                    break
            if annotated > 0:
                needs_tight_export = True
    if isinstance(color_collector, dict):
        color_collector.clear()
        for col, line in zip(plotted_columns, lns):
            color_collector[col] = line.get_color()
    export_bbox = 'tight' if needs_tight_export else None
    export_pad = 0.02 if needs_tight_export else 0
    return figure_to_stream(fig, bbox_inches=export_bbox, pad_inches=export_pad)
#Normaliza etiquetas de periodo eliminando prefijos genericos
def normalize_period_label(label) -> str:
    """Normaliza etiquetas de periodo eliminando prefijos genericos."""
    if label is None:
        return ""
    text = str(label).strip()
    if not text:
        return text
    parts = text.split(None, 1)
    if parts and parts[0].lower().rstrip('.') in {'vol', 'volume', 'volumen'} and len(parts) > 1:
        return parts[1].strip()
    return text
#Grafico de barras apiladas 100% para share por periodo
def stacked_share_chart(period_label, share_values, color_mapping, c_fig, title=None):
    """
    Genera un gráfico 100% apilado a partir de la participación de un periodo puntual.
    """
    if not isinstance(share_values, (dict, OrderedDict)):
        share_values = {}
    palette_values = TREND_COLOR_SEQUENCE or [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]
    normalized_items = []
    total = 0.0
    for key, value in share_values.items():
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            try:
                numeric_value = float(str(value).replace('%', '').replace(',', '.')) / 100.0
            except Exception:
                numeric_value = 0.0
        if not np.isfinite(numeric_value) or numeric_value < 0:
            numeric_value = 0.0
        total += numeric_value
        normalized_items.append([key, numeric_value])
    if total <= 0:
        total = 1.0
    for item in normalized_items:
        item[1] = item[1] / total
    normalized_items = [item for item in normalized_items if item[1] > 0]
    segments = len(normalized_items)
    # Altura fija para mantener consistencia visual entre gráficos
    # con diferente cantidad de elementos.
    figure_height = 7.0
    figure_width = 4.6
    fig, ax = plt.subplots(num=c_fig, figsize=(figure_width, figure_height), dpi=DEFAULT_EXPORT_DPI)
    fig.patch.set_alpha(0.0)
    ax.set_facecolor('none')
    bottom = 0.0
    bar_width = 0.65
    label_infos = []
    for idx, (key, value) in enumerate(normalized_items):
        color = color_mapping.get(key)
        if color is None:
            color = palette_values[idx % len(palette_values)]
            color_mapping[key] = color
        ax.bar(0, value, width=bar_width, bottom=bottom, color=color, edgecolor='white')
        center_y = bottom + value / 2 if value > 0 else bottom
        share_label = f"{value * 100:.1f}%"
        label_infos.append(
            {
                "key": key,
                "text": f"{key} ({share_label})",
                "color": color,
                "center_y": center_y,
            }
        )
        bottom += value
    max_text_len = 0
    if label_infos:
        label_infos.sort(key=lambda info: info["center_y"])
        min_gap = max(0.035, min(0.12, 0.95 / max(segments, 1)))
        min_y = 0.04
        max_y = 0.96
        adjusted_positions = []
        for info in label_infos:
            proposed = np.clip(info["center_y"], min_y, max_y)
            if adjusted_positions and proposed - adjusted_positions[-1] < min_gap:
                proposed = min(max_y, adjusted_positions[-1] + min_gap)
            adjusted_positions.append(proposed)
        for idx in range(len(adjusted_positions) - 2, -1, -1):
            if adjusted_positions[idx + 1] - adjusted_positions[idx] < min_gap:
                adjusted_positions[idx] = max(min_y, adjusted_positions[idx + 1] - min_gap)
        for info, position in zip(label_infos, adjusted_positions):
            info["label_y"] = position
        max_text_len = max(len(info["text"]) for info in label_infos)
        label_offset = 1.1 + max(0, (max_text_len - 22) * 0.014)
        label_offset = min(label_offset, 1.8)
        label_x = bar_width / 2 + label_offset
    else:
        label_x = bar_width / 2 + 1.15
    connector_x = bar_width / 2 + 0.1
    label_font_size = 9 if segments <= 10 else (8 if segments <= 16 else 7)
    if label_infos:
        for info in label_infos:
            ax.plot(
                [connector_x, label_x - 0.1],
                [info["center_y"], info["label_y"]],
                color=info["color"],
                linewidth=1.2,
                alpha=0.65,
            )
            ax.text(
                label_x,
                info["label_y"],
                info["text"],
                ha='left',
                va='center',
                color=info["color"],
                fontsize=label_font_size,
                fontweight='bold',
                clip_on=False,
            )
    ax.set_ylim(0, 1)
    label_padding = max(1.0, min(3.2, 0.072 * max_text_len)) if label_infos else 1.4
    axis_right = label_x + label_padding if label_infos else (bar_width / 2 + label_padding)
    ax.set_xlim(-0.75, axis_right)
    ax.set_xticks([])
    ax.set_ylabel('')
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x * 100:.0f}%"))
    chart_title = period_label if title is None else str(title).strip()
    if chart_title:
        ax.set_title(chart_title, fontsize=15, pad=24, fontweight='bold')
    ax.set_yticks(np.linspace(0, 1, 5))
    ax.yaxis.tick_left()
    ax.grid(axis='y', linestyle='--', alpha=0.25)
    for spine in ax.spines.values():
        spine.set_visible(False)
    fig.subplots_adjust(left=0.06, right=0.98, top=0.985, bottom=0.025)
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    tight_bbox = fig.get_tightbbox(renderer)
    if not tight_bbox or tight_bbox.width <= 0 or tight_bbox.height <= 0:
        tight_bbox = None
    fig_size = fig.get_size_inches()
    fixed_width = float(fig_size[0])
    fixed_height = float(fig_size[1])
    return (
        figure_to_stream(fig, bbox_inches=tight_bbox, pad_inches=0.02, transparent=True),
        (fixed_width, fixed_height),
    )
# Segmento 2: helpers para arbol de medidas (por que).
def _limpiar_tabla_excel(df):
    """Detecta la fila de encabezados (MAT ...) tolerando filas adicionales al inicio."""
    header_idx = None
    for idx, row in df.iterrows():
        non_null = [str(v) for v in row.tolist() if pd.notna(v)]
        periodos = [val for val in non_null[1:] if str(val).startswith("MAT")]
        tiene_table = any("table" in str(val).lower() for val in non_null)
        if periodos or tiene_table:
            header_idx = idx
            break
    if header_idx is None:
        return None, None
    header_row = df.iloc[header_idx].tolist()
    periodos = [str(c).strip() for c in header_row[1:] if pd.notna(c)]
    if len(periodos) < 2:
        return None, None
    tabla = df.iloc[header_idx + 1:, : len(periodos) + 1].copy()
    tabla.columns = ["Metric"] + periodos
    tabla = tabla.dropna(how="all")
    tabla = tabla.dropna(subset=["Metric"])
    tabla["Metric"] = tabla["Metric"].astype(str).str.strip()
    for col in periodos:
        tabla[col] = pd.to_numeric(tabla[col], errors="coerce")
    return tabla.reset_index(drop=True), periodos
def _nombres_unidad(unidad):
    """Devuelve las etiquetas del arbol segun la unidad indicada."""
    unidad_lower = unidad.lower()
    nombre_vol = {
        "litros": ("LTs", "ML"),
        "units": ("Units", "Unit"),
        "kilos": ("KGs", "KG"),
        "toneladas": ("Tons", "Ton"),
        "rollos": ("Rollos", "Rollo"),
        "metros": ("MTs", "MT"),
        "hojas": ("Hojas", "Hoja"),
    }.get(unidad_lower, (unidad, unidad))
    vol_label, sub_label = nombre_vol
    precio_unit = vol_label[:-1] if vol_label.endswith("s") else vol_label
    return {
        "valor": "Valor $ 000s",
        "volumen": f"Volumen 000 {vol_label}",
        "precio": f"Precio Promedio $/{precio_unit}",
        "gasto": "Gasto Promedio $/Buyer",
        "compradores": "Compradores 000s",
        "volumen_prom": f"Volumen Promedio {sub_label}/Comprador",
        "penetracion": "% Penetracion",
        "hholds": "Total HHolds 000s",
        "frecuencia": "Frecuencia",
        "volumen_viaje": f"Volumen por Viaje {sub_label}/Viaje",
        "ticket": "Gasto por Ticket $/Viaje",
    }
def calcular_cambios(df, periodo_inicial, periodo_final, unidad='Units', penetracion_en_pp=False):
    """Calcula niveles y variaciones porcentuales entre dos periodos."""
    etiquetas = _nombres_unidad(unidad)
    def calculate_percentage_change(old, new):
        if old is None or new is None:
            return 0
        if not np.isfinite(old) or not np.isfinite(new) or old == 0:
            return 0
        return ((new - old) / old) * 100
    def calculate_percentage_point_change(old, new):
        if old is None or new is None:
            return 0
        if not np.isfinite(old) or not np.isfinite(new):
            return 0
        delta = new - old
        if abs(old) <= 1 and abs(new) <= 1:
            delta *= 100
        return delta
    factores = {
        'Units': 1,
        'Litros': 1,
        # La base viene en kilos: en kilos no se escala, en toneladas se divide por 1000.
        'Kilos': 1,
        # Para toneladas la base viene en kilos, se divide por 1000 en volumen
        # y se multiplica por 1000 en precios; factor < 1 corrige el inflado observado.
        'Toneladas': 0.001,
        'Rollos': 10,
        'Metros': 100,
        'Hojas': 1
    }
    factor = factores.get(unidad, 1)
    # Los promedios por comprador/viaje ya vienen en la unidad correcta
    # (sin escalar) para todas las unidades.
    factor_volumen_prom = 1
    factor_volumen_viaje = 1
    def _normalize_metric_label(label):
        normalized = unicodedata.normalize("NFKD", str(label))
        normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
        normalized = re.sub(r"\bweighted\b", "", normalized, flags=re.IGNORECASE)
        normalized = normalized.strip().upper()
        return re.sub(r"[^A-Z0-9]+", "", normalized)
    metric_index = df["Metric"].map(_normalize_metric_label).fillna("")
    metric_vertical_mask = metric_index.str.contains("VERT", na=False)
    def _metric_candidates(metrico):
        if isinstance(metrico, (list, tuple, set)):
            raw = list(metrico)
        else:
            raw = [metrico]
        return [_normalize_metric_label(m) for m in raw if m is not None]
    def leer_metric(metrico, columna):
        candidates = _metric_candidates(metrico)
        for candidate in candidates:
            mask = (metric_index == candidate) & (~metric_vertical_mask)
            if mask.any():
                return df.loc[mask, columna].values[0]
        for candidate in candidates:
            if not candidate:
                continue
            mask = metric_index.str.contains(candidate, na=False) & (~metric_vertical_mask)
            if mask.any():
                return df.loc[mask, columna].values[0]
        if isinstance(metrico, (list, tuple, set)):
            display_name = ", ".join(str(m) for m in metrico)
        else:
            display_name = str(metrico)
        raise KeyError(f"No se encontro la metrica '{display_name}' en los datos.")
    def build_entry(label, metrica, transform=lambda x, _: x, change_fn=calculate_percentage_change, change_suffix="%"):
        inicial = transform(leer_metric(metrica, periodo_inicial), periodo_inicial)
        final = transform(leer_metric(metrica, periodo_final), periodo_final)
        return {
            "valor_origen": inicial,
            "value": final,
            "change": change_fn(inicial, final),
            "change_suffix": change_suffix,
        }
    segment2_metric_candidates = {
        # Compatibilidad entre hojas legacy (R_VOL/VO1/PM1) y nuevas hojas en unidades (UNITS/VOPK/PMPK).
        "volumen": ["Weighted R_VOL1", "Weighted R_VOL2", "Weighted UNITS", "Weighted VOL1_P", "Weighted VOL2_P", "Weighted VOLSU"],
        "precio": ["Weighted PM1_LC", "Weighted PMSU_LC", "Weighted PM2_LC", "Weighted PMPK_LC"],
        "volumen_prom": ["Weighted VO1_BUY", "Weighted VOSU_BUY", "Weighted VO2_BUY", "Weighted VOPK_BUY"],
        "volumen_viaje": ["Weighted VO1_DAY", "Weighted VOSU_DAY", "Weighted VO2_DAY", "Weighted VOPK_DAY"],
    }
    metrics_calculated = {
        etiquetas["valor"]: build_entry(etiquetas["valor"], 'Weighted VAL_LC'),
        etiquetas["volumen"]: build_entry(
            etiquetas["volumen"],
            segment2_metric_candidates["volumen"],
            lambda v, _: v * factor
        ),
        etiquetas["precio"]: build_entry(
            etiquetas["precio"],
            segment2_metric_candidates["precio"],
            lambda v, _: v / factor
        ),
        etiquetas["gasto"]: build_entry(etiquetas["gasto"], 'Weighted VALC_BUY'),
        etiquetas["compradores"]: build_entry(etiquetas["compradores"], 'Weighted BUYERS'),
        etiquetas["volumen_prom"]: build_entry(
            etiquetas["volumen_prom"],
            segment2_metric_candidates["volumen_prom"],
            lambda v, _: v * factor_volumen_prom
        ),
        etiquetas["penetracion"]: build_entry(
            etiquetas["penetracion"],
            'Weighted PENET',
            change_fn=calculate_percentage_point_change if penetracion_en_pp else calculate_percentage_change,
            change_suffix=" pp" if penetracion_en_pp else "%"
        ),
        etiquetas["hholds"]: build_entry(etiquetas["hholds"], 'Weighted HHOLDS'),
        etiquetas["frecuencia"]: build_entry(etiquetas["frecuencia"], 'Weighted FREQ'),
        etiquetas["volumen_viaje"]: build_entry(
            etiquetas["volumen_viaje"],
            segment2_metric_candidates["volumen_viaje"],
            lambda v, _: v * factor_volumen_viaje
        ),
        etiquetas["ticket"]: build_entry(etiquetas["ticket"], 'Weighted VALC_DAY'),
    }
    return metrics_calculated
def _format_val(label_key, value):
    """Formato sensible a cada metrica para asemejarse al ejemplo de referencia."""
    reglas = {
        "Valor $ 000s": {"scale": 1 / 1000, "decimals": 0, "thousands": True},
        "Volumen 000": {"scale": 1, "decimals": 0, "thousands": True},
        "Precio Promedio": {"scale": 1, "decimals": 2, "thousands": False},
        "Gasto Promedio": {"scale": 1, "decimals": 2, "thousands": False},
        "Compradores 000s": {"scale": 1 / 1000, "decimals": 0, "thousands": True},
        # Mostrar con 1 decimal para que se vea 0.2 en lugar de 0 ó 170
        "Volumen Promedio": {"scale": 1, "decimals": 2, "thousands": False},
        "% Penetracion": {"scale": 1, "decimals": 1, "thousands": False},
        "Total HHolds 000s": {"scale": 1 / 1000, "decimals": 0, "thousands": True},
        "Frecuencia": {"scale": 1, "decimals": 1, "thousands": False},
        # Mostrar con 1 decimal (p.ej. 0.1) en lugar de redondear a entero
        "Volumen por Viaje": {"scale": 1, "decimals": 2, "thousands": False},
        "Gasto por Ticket": {"scale": 1, "decimals": 2, "thousands": False},
    }
    regla = next((v for k, v in reglas.items() if label_key.startswith(k)), {"scale": 1, "decimals": 2, "thousands": False})
    scaled = value * regla["scale"]
    if regla["thousands"]:
        fmt = f"{{:,.{regla['decimals']}f}}" if regla["decimals"] else "{:,.0f}"
    else:
        fmt = f"{{:.{regla['decimals']}f}}"
    return fmt.format(round(scaled, regla["decimals"]))
def graficar_arbol(metrics_calculated, volumen_unidad='Units', output_dir=None, hoja='', mostrar=False):
    """Dibuja el arbol de variables y retorna un stream PNG listo para insertar en la PPT."""
    etiquetas = _nombres_unidad(volumen_unidad)
    def wrap_label(text, max_chars=14):
        return textwrap.fill(text, width=max_chars, break_long_words=False, break_on_hyphens=False)
    level4_ids = {"penetracion", "hholds", "frecuencia1", "frecuencia2", "volumen_viaje", "ticket"}
    def card_size(node_id):
        if node_id in level4_ids:
            return 0.12, 0.11
        return 0.15, 0.12
    nodes = [
        {"id": "valor", "label": etiquetas["valor"], "metric": etiquetas["valor"], "pos": (0.50, 0.86), "color": "#8cc63e"},
        {"id": "volumen", "label": etiquetas["volumen"], "metric": etiquetas["volumen"], "pos": (0.33, 0.63), "color": "#8fd1c4"},
        {"id": "precio", "label": etiquetas["precio"], "metric": etiquetas["precio"], "pos": (0.70, 0.63), "color": "#bf5a9b"},
        {"id": "compradores", "label": etiquetas["compradores"], "metric": etiquetas["compradores"], "pos": (0.18, 0.46), "color": "#83a860"},
        {"id": "volumen_prom", "label": etiquetas["volumen_prom"], "metric": etiquetas["volumen_prom"], "pos": (0.50, 0.46), "color": "#6495c4"},
        {"id": "gasto", "label": etiquetas["gasto"], "metric": etiquetas["gasto"], "pos": (0.82, 0.46), "color": "#c671b7"},
        {"id": "penetracion", "label": etiquetas["penetracion"], "metric": etiquetas["penetracion"], "pos": (0.08, 0.29), "color": "#27a347"},
        {"id": "hholds", "label": etiquetas["hholds"], "metric": etiquetas["hholds"], "pos": (0.26, 0.29), "color": "#bdaa7e"},
        {"id": "frecuencia1", "label": etiquetas["frecuencia"], "metric": etiquetas["frecuencia"], "pos": (0.41, 0.29), "color": "#00aeef"},
        {"id": "volumen_viaje", "label": etiquetas["volumen_viaje"], "metric": etiquetas["volumen_viaje"], "pos": (0.58, 0.29), "color": "#0074c1"},
        {"id": "frecuencia2", "label": etiquetas["frecuencia"], "metric": etiquetas["frecuencia"], "pos": (0.74, 0.29), "color": "#00aeef"},
        {"id": "ticket", "label": etiquetas["ticket"], "metric": etiquetas["ticket"], "pos": (0.90, 0.29), "color": "#8a63ad"},
    ]
    edges = [
        ("valor", "volumen", "solid"),
        ("valor", "precio", "dashed"),
        ("valor", "gasto", "dashed"),
        ("volumen", "compradores", "solid"),
        ("volumen", "volumen_prom", "solid"),
        ("compradores", "penetracion", "solid"),
        ("compradores", "hholds", "solid"),
        ("volumen_prom", "frecuencia1", "solid"),
        ("volumen_prom", "volumen_viaje", "solid"),
        ("gasto", "frecuencia2", "solid"),
        ("gasto", "ticket", "solid"),
    ]
    pos_change_color = "#0bbdaf"
    neg_change_color = "#ef5350"
    neu_change_color = "#9e9e9e"
    def change_color(ch):
        return pos_change_color if ch > 3 else (neg_change_color if ch < -3 else neu_change_color)
    def draw_card(ax, center, node):
        metric_info = metrics_calculated[node["metric"]]
        value_txt = _format_val(node["label"], metric_info["value"])
        change_suffix = metric_info.get("change_suffix", "%")
        if str(change_suffix).strip().lower() == "pp":
            change_txt = f"{metric_info['change']:+.1f} pp"
        else:
            change_txt = f"{metric_info['change']:+.1f}%"
        label_wrapped = wrap_label(node["label"])
        label_lines = label_wrapped.split("\n")
        x, y = center
        width, height = card_size(node["id"])
        header_h = height * (0.36 + 0.08 * (len(label_lines) - 1))
        rounding = 0.02
        shadow = mpatches.FancyBboxPatch(
            (x - width / 2 + 0.008, y - height / 2 - 0.015),
            width, height,
            boxstyle=f"round,pad=0.01,rounding_size={rounding}",
            linewidth=0, facecolor="#9e9e9e", alpha=0.35, zorder=1)
        ax.add_patch(shadow)
        body = mpatches.FancyBboxPatch(
            (x - width / 2, y - height / 2),
            width, height,
            boxstyle=f"round,pad=0.01,rounding_size={rounding}",
            linewidth=1, edgecolor="#c1c1c1", facecolor="white", zorder=2)
        ax.add_patch(body)
        header = mpatches.FancyBboxPatch(
            (x - width / 2, y + height / 2 - header_h),
            width, header_h,
            boxstyle=f"round,pad=0.01,rounding_size={rounding}",
            linewidth=1, edgecolor=node["color"], facecolor=node["color"], zorder=3)
        ax.add_patch(header)
        ax.text(x, y + height / 2 - header_h / 2, label_wrapped, va="center", ha="center",
                fontsize=11, color="white", weight="bold", zorder=4)
        value_x = x - width * 0.22
        change_x = x + width * 0.22
        values_y = y - 0.035
        ax.text(value_x, values_y, value_txt, va="center", ha="center",
                fontsize=12, color="#222222", weight="bold", zorder=4)
        ax.text(change_x, values_y, change_txt, va="center", ha="center",
                fontsize=11, color=change_color(metric_info["change"]), weight="bold", zorder=4)
    def draw_edge(ax, from_id, to_id, style):
        x0, y0 = next(n for n in nodes if n["id"] == from_id)["pos"]
        x1, y1 = next(n for n in nodes if n["id"] == to_id)["pos"]
        dash_style = '--' if style == "dashed" else "solid"
        if from_id == "valor" and to_id == "gasto":
            w_valor, h_valor = card_size("valor")
            start_x = x0 + w_valor / 2
            start_y = y0 - h_valor * 0.15
            ax.annotate("",
                        xy=(x1, start_y), xytext=(start_x, start_y),
                        arrowprops=dict(arrowstyle='-', lw=1.2,
                                        linestyle=dash_style,
                                        color="#8c8c8c", shrinkA=0, shrinkB=0))
            ax.annotate("",
                        xy=(x1, y1 + 0.055), xytext=(x1, start_y),
                        arrowprops=dict(arrowstyle='-|>', lw=1.2,
                                        linestyle=dash_style,
                                        color="#8c8c8c", shrinkA=0, shrinkB=4))
            return
        _, h_from = card_size(from_id)
        split_offset = 0.68
        start_y = y0 - h_from * split_offset
        origin_offset = 0.60
        origin_y = y0 - h_from * origin_offset
        ax.annotate("",
                    xy=(x0, start_y), xytext=(x0, origin_y),
                    arrowprops=dict(arrowstyle='-', lw=1.2,
                                    linestyle=dash_style,
                                    color="#8c8c8c", shrinkA=0, shrinkB=0))
        ax.annotate("",
                    xy=(x1, y1 + 0.055), xytext=(x0, start_y),
                    arrowprops=dict(arrowstyle='-|>', lw=1.2,
                                    linestyle=dash_style,
                                    color="#8c8c8c", shrinkA=0, shrinkB=4))
    def draw_key(ax):
        ax.text(0.05, 0.88, "KEY", ha="left", va="center", fontsize=10, color="#565656", weight="bold")
        ax.add_patch(mpatches.Rectangle((0.09, 0.87), 0.02, 0.02, facecolor=pos_change_color, edgecolor="none"))
        ax.text(0.115, 0.88, "= > 3% Cambio", ha="left", va="center", fontsize=9, color=pos_change_color)
        ax.add_patch(mpatches.Rectangle((0.21, 0.87), 0.02, 0.02, facecolor=neg_change_color, edgecolor="none"))
        ax.text(0.235, 0.88, "= < -3% Cambio", ha="left", va="center", fontsize=9, color=neg_change_color)
    def draw_attribution_bar(ax, metrics):
        valor_info = metrics[etiquetas["valor"]]
        v0 = valor_info["valor_origen"]
        v1 = valor_info["value"]
        delta_v = v1 - v0
        if v0 <= 0 or v1 <= 0 or delta_v == 0:
            return
        drivers = [
            ("Frecuencia", etiquetas["frecuencia"], "#00b0f0"),
            ("% Penetracion", etiquetas["penetracion"], "#27a347"),
            ("Total HHolds 000s", etiquetas["hholds"], "#bdaa7e"),
            ("Volumen por Viaje", etiquetas["volumen_viaje"], "#1f78b4"),
            ("Precio Promedio", etiquetas["precio"], "#b4559c"),
        ]
        driver_logs = []
        for display, key, color in drivers:
            info = metrics[key]
            a, b = info["valor_origen"], info["value"]
            if a <= 0 or b <= 0:
                driver_logs.append((display, key, color, 0.0))
            else:
                driver_logs.append((display, key, color, math.log(b / a)))
        total_log = sum(item[3] for item in driver_logs)
        if total_log == 0:
            total_log = math.log(v1 / v0)
            if total_log == 0:
                return
        contribs = []
        for display, key, color, log_ratio in driver_logs:
            if log_ratio == 0:
                contribs.append((display, 0.0, 0.0, color))
                continue
            contrib_val = delta_v * (log_ratio / total_log)
            chart_pct = (contrib_val / abs(delta_v)) * 100
            contribs.append((display, chart_pct, contrib_val, color))
        total_abs = sum(abs(c[1]) for c in contribs)
        if total_abs == 0:
            return
        neg_abs = sum(abs(c[1]) for c in contribs if c[1] < 0)
        zero_x = neg_abs / total_abs
        bar_ax = fig.add_axes([0.065, 0.80, 0.36, 0.10])
        bar_ax.set_axis_off()
        bar_ax.set_xlim(0, 1)
        bar_ax.set_ylim(0, 1)
        bar_ax.plot([zero_x, zero_x], [0.30, 0.70], color="#5f5f5f", linewidth=1)
        x_left = zero_x
        x_right = zero_x
        for display, pct, _, color in contribs:
            width = abs(pct) / total_abs
            if pct < 0:
                x_left -= width
                bar_ax.add_patch(mpatches.Rectangle((x_left, 0.35), width, 0.30, facecolor=color, edgecolor="none"))
            else:
                bar_ax.add_patch(mpatches.Rectangle((x_right, 0.35), width, 0.30, facecolor=color, edgecolor="none"))
                x_right += width
        ax.text(0.06, 0.92, "Atribucion del cambio en el gasto", ha="left", va="center",
                fontsize=14, weight="bold", color="#333333")
        draw_key(ax)
    fig = plt.figure(figsize=(21, 9), dpi=DEFAULT_EXPORT_DPI)
    ax = plt.gca()
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    for p, c, style in edges:
        draw_edge(ax, p, c, style)
    for node in nodes:
        draw_card(ax, node["pos"], node)
    draw_attribution_bar(ax, metrics_calculated)
    plt.title(f"Arbol de Variables ({volumen_unidad})", fontsize=10, loc="right", color="#666")
    fig_size = tuple(fig.get_size_inches()) if fig is not None else None
    buf = figure_to_stream(fig, dpi=DEFAULT_EXPORT_DPI, bbox_inches='tight', pad_inches=0.02, close=False)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        suffix = f"_{hoja}" if hoja else ""
        image_path = os.path.join(output_dir, f"arbol_{volumen_unidad}{suffix}.png")
        with open(image_path, "wb") as handler:
            handler.write(buf.getvalue())
        buf.seek(0)
    if mostrar:
        plt.show()
    plt.close(fig)
    return buf, fig_size
def _unidad_desde_nombre_hoja(nombre_hoja):
    """
    Si el nombre cumple con el patron '2_NombreMarca_U' o '2_NombreMarca_U_pp',
    retorna la unidad solicitada. Ejemplo: '2_Coca Cola_L' -> 'Litros'.
    Si no coincide o la letra no es valida, devuelve None.
    """
    match = re.match(r"^\d+_(.+)_([A-Za-z])(?:_(?:p|pp|%))?$", nombre_hoja.strip(), flags=re.IGNORECASE)
    if not match:
        return None
    letra = match.group(2).upper()
    mapa = {
        "U": "Units",
        "L": "Litros",
        "K": "Kilos",
        "T": "Toneladas",
        "R": "Rollos",
        "M": "Metros",
        "H": "Hojas",
    }
    return mapa.get(letra)

def _modo_penetracion_desde_nombre_hoja(nombre_hoja):
    """Devuelve 'pp' cuando la hoja de Segmento 2 pide puntos porcentuales."""
    match = re.match(r"^\d+_.+_[A-Za-z]_(p|pp|%)$", nombre_hoja.strip(), flags=re.IGNORECASE)
    if not match:
        return "%"
    token = match.group(1).lower()
    return "pp" if token in {"p", "pp"} else "%"

def _etiqueta_segmento2_desde_nombre_hoja(nombre_hoja):
    """Quita prefijo, unidad y modo opcional para mostrar la marca en el titulo."""
    sheet_clean = str(nombre_hoja or "").strip()
    match = re.match(r"^\d+_(.+)_([A-Za-z])(?:_(?:p|pp|%))?$", sheet_clean, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    return sheet_clean[2:] if len(sheet_clean) > 2 else sheet_clean
# Segmentos 3-6: tabla de aportes y calculos de shares/variaciones.
def aporte(df,p,lang,tipo):
        """Construye la tabla de aporte y calcula shares/variaciones."""
        aux = df.copy()
        removed_non_numeric: list[str] = []
        grouped_label = ""
        grouped_columns_labels: list[str] = []
        if hasattr(df, "attrs"):
            grouped_label = str(df.attrs.get("size_grouped_label", "")).strip()
            grouped_columns_labels = [
                str(label).strip()
                for label in df.attrs.get("size_grouped_columns", [])
                if str(label).strip()
            ]
        if aux.empty:
            apo = pd.DataFrame(columns=[tipo])
            apo.attrs["removed_columns"] = []
            apo.attrs["share_period_values"] = []
            apo.attrs["share_mat_values"] = {}
            apo.attrs["grouped_source_labels"] = {}
            apo.attrs["skip_table"] = True
            apo.attrs["skip_table_reason"] = "empty_data"
            return apo
        first_col_name = aux.columns[0]
        first_col = aux.iloc[:, 0].copy()
        numeric_data = OrderedDict()
        for idx in range(1, aux.shape[1]):
            col_name = aux.columns[idx]
            series = aux.iloc[:, idx]
            if np.issubdtype(series.dtype, np.datetime64):
                removed_non_numeric.append(str(col_name))
                continue
            numeric_series = pd.to_numeric(series, errors='coerce')
            if numeric_series.dropna().empty:
                removed_non_numeric.append(str(col_name))
                continue
            numeric_data[col_name] = numeric_series
        if numeric_data:
            sanitized_aux = pd.concat(
                [first_col.to_frame(name=first_col_name), pd.DataFrame(numeric_data)],
                axis=1
            )
        else:
            sanitized_aux = first_col.to_frame(name=first_col_name)
        sanitized_aux.reset_index(drop=True, inplace=True)
        aux = sanitized_aux
        if aux.shape[1] <= 1:
            apo = pd.DataFrame(columns=[tipo])
            apo.attrs["removed_columns"] = list(dict.fromkeys(removed_non_numeric))
            apo.attrs["share_period_values"] = []
            apo.attrs["share_mat_values"] = {}
            apo.attrs["grouped_source_labels"] = {}
            apo.attrs["skip_table"] = True
            apo.attrs["skip_table_reason"] = "no_numeric_columns"
            return apo
        try:
            pipeline_offset = int(p) if p is not None else 0
        except (TypeError, ValueError):
            pipeline_offset = 0
        if pipeline_offset < 0:
            pipeline_offset = 0
        p = pipeline_offset
        required_rows = 24 + pipeline_offset
        if len(aux) < required_rows:
            apo = pd.DataFrame(columns=[tipo])
            apo.attrs["removed_columns"] = list(dict.fromkeys(removed_non_numeric))
            apo.attrs["share_period_values"] = []
            apo.attrs["share_mat_values"] = {}
            apo.attrs["grouped_source_labels"] = {}
            apo.attrs["skip_table"] = True
            apo.attrs["skip_table_reason"] = "insufficient_periods"
            return apo
        date_col_label = labels.get((lang, 'Data'))
        date_col_name = date_col_label if date_col_label in aux.columns else first_col_name
        aux_numeric = aux.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
        aux['Total'] = aux_numeric.sum(axis=1, skipna=True)
        apo=pd.DataFrame(columns=[tipo] + aux.columns[1:].tolist())
        #Vol ultimo MAT
        apo.loc[len(apo)] = [str('Vol' )+" "+aux.loc[len(aux)-1-12-p,date_col_name].strftime('%b-%y') ] + [aux.iloc[len(aux)-24-p:len(aux)-12-p, col].sum() / aux.iloc[len(aux)-24-p:len(aux)-12-p,aux.shape[1]-1].sum() for col in range(1,len(aux.columns) )]
        #Vol MAT actual
        apo.loc[len(apo)] = [str('Vol' )+" "+aux.loc[len(aux)-1-p,date_col_name].strftime('%b-%y') ] + [aux.iloc[len(aux)-12-p:len(aux)-p, col].sum() / aux.iloc[len(aux)-12-p:len(aux)-p,aux.shape[1]-1].values.sum() for col in range(1,len(aux.columns))]
        #Var
        apo.loc[len(apo)] = ["Var %"] +  [ aux.iloc[len(aux)-12-p:len(aux)-p, col].sum() / aux.iloc[len(aux)-24-p:len(aux)-12-p, col].sum()-1 for col in range(1,len(aux.columns))]
        #Aporte
        apo.loc[len(apo)] = ["Aporte"] + [apo.iloc[-1, col] * apo.iloc[0, col] for col in range(1,len(aux.columns))]
        # Remove categories with invalid metrics (NaN/Inf) before formatting
        def _is_invalid_metric(value):
            if pd.isna(value):
                return True
            try:
                numeric = float(value)
            except (TypeError, ValueError):
                return False
            return not np.isfinite(numeric)
        invalid_columns = []
        for col_idx in range(1, len(apo.columns)):
            col_name = apo.columns[col_idx]
            var_value = apo.iloc[2, col_idx]
            aporte_value = apo.iloc[3, col_idx]
            if _is_invalid_metric(var_value) or _is_invalid_metric(aporte_value):
                invalid_columns.append(col_name)
        if invalid_columns:
            apo.drop(columns=invalid_columns, inplace=True)
        removed_all = list(dict.fromkeys(removed_non_numeric + invalid_columns))
        apo.attrs["removed_columns"] = removed_all
        apo.attrs["grouped_source_labels"] = (
            {grouped_label: grouped_columns_labels}
            if grouped_label and grouped_columns_labels
            else {}
        )
        share_period_values = []
        max_share_rows = min(2, len(apo.index))
        for row_idx in range(max_share_rows):
            period_label = normalize_period_label(apo.iloc[row_idx, 0])
            period_shares = OrderedDict()
            for col_idx in range(1, len(apo.columns)):
                column_name = str(apo.columns[col_idx])
                normalized_name = column_name.strip().lower()
                if not normalized_name or normalized_name == 'total':
                    continue
                value = apo.iloc[row_idx, col_idx]
                numeric_value = None
                try:
                    numeric_value = float(value)
                except (TypeError, ValueError):
                    try:
                        numeric_value = float(str(value).replace('%', '').replace(',', '.')) / 100.0
                    except Exception:
                        numeric_value = None
                if numeric_value is None or not np.isfinite(numeric_value) or numeric_value < 0:
                    numeric_value = 0.0
                period_shares[column_name] = numeric_value
            share_period_values.append((period_label, period_shares))
        apo.attrs["share_period_values"] = share_period_values
        share_mat_values = {}
        if len(share_period_values) > 1:
            _, current_period_shares = share_period_values[1]
            share_mat_values = {
                column_name: value
                for column_name, value in current_period_shares.items()
                if np.isfinite(value)
            }
        apo.attrs["share_mat_values"] = share_mat_values
        # En pandas nuevos, asignar strings sobre columnas float via iloc produce error.
        # Forzamos estas columnas a object para mantener compatibilidad con 2.3.3 y 3.x.
        for column_name in apo.columns[1:]:
            apo[column_name] = apo[column_name].astype(object)
        #Formato del volumen
        apo.iloc[:2, 1:] = _df_map_compat(apo.iloc[:2, 1:], lambda x: f"{round(x * 100, 1)}%")
        #Formato de la variacion y el aporte
        apo.iloc[2:, 1:] = _df_map_compat(apo.iloc[2:, 1:], lambda x: f"{round(x * 100, 2)}%")
        return apo

def _parse_percent_cell(value) -> Optional[float]:
    """Convierte celdas tipo '12.3%' o numericas a float comparable."""
    if value is None:
        return None
    if isinstance(value, (int, float, np.number)):
        numeric = float(value)
        return numeric if np.isfinite(numeric) else None
    text_value = str(value).strip().replace('%', '').replace(',', '.')
    if not text_value:
        return None
    try:
        numeric = float(text_value)
    except (TypeError, ValueError):
        return None
    return numeric if np.isfinite(numeric) else None

SEMANTIC_LABEL_ORIGIN_TOKENS = {
    'c', 'v', 'sv',
    'compra', 'compras', 'purchase', 'purchases',
    'venta', 'ventas', 'sale', 'sales', 'venda', 'vendas',
    'sellin', 'sellout',
}
ALNUM_BOUNDARY_SPLIT_PATTERN = re.compile(r'(?<=\d)(?=[a-z])|(?<=[a-z])(?=\d)')

@dataclass(frozen=True)
class SemanticLabelSignature:
    raw: str
    canonical: str
    compact: str
    tokens: tuple[str, ...]
    alpha_tokens: tuple[str, ...]
    numeric_tokens: tuple[str, ...]

@dataclass(frozen=True)
class SemanticLabelMatch:
    reference_label: str
    reference_index: int
    score: float

AMBIGUOUS_FAMILY_BASE_PALETTE = ['#64748B', '#0F766E', '#9A3412', '#6D28D9', '#B45309']

def _prepare_semantic_label_text(value) -> str:
    normalized = normalize_color_text(value)
    if not normalized:
        return ''
    normalized = strip_color_suffixes(normalized)
    normalized = re.sub(r'\bsell[\s_-]*in\b', 'sellin', normalized)
    normalized = re.sub(r'\bsell[\s_-]*out\b', 'sellout', normalized)
    normalized = ALNUM_BOUNDARY_SPLIT_PATTERN.sub(' ', normalized)
    return re.sub(r'\s+', ' ', normalized).strip()

@lru_cache(maxsize=4096)
def build_semantic_label_signature(value) -> SemanticLabelSignature:
    raw = '' if value is None else str(value).strip()
    prepared = _prepare_semantic_label_text(raw)
    if not prepared:
        return SemanticLabelSignature(raw=raw, canonical='', compact='', tokens=(), alpha_tokens=(), numeric_tokens=())
    tokens = tuple(
        token
        for token in COLOR_TOKEN_SPLIT_PATTERN.split(prepared)
        if token and token not in SEMANTIC_LABEL_ORIGIN_TOKENS
    )
    alpha_tokens = tuple(token for token in tokens if any(ch.isalpha() for ch in token))
    numeric_tokens = tuple(token for token in tokens if token.isdigit())
    canonical = ' '.join(tokens)
    compact = ''.join(tokens)
    return SemanticLabelSignature(
        raw=raw,
        canonical=canonical,
        compact=compact,
        tokens=tokens,
        alpha_tokens=alpha_tokens,
        numeric_tokens=numeric_tokens,
    )

def _semantic_token_match_score(left_token: str, right_token: str) -> float:
    if not left_token or not right_token:
        return 0.0
    if left_token == right_token:
        return 1.0
    if left_token.isdigit() or right_token.isdigit():
        return 0.0
    ratio = SequenceMatcher(None, left_token, right_token).ratio()
    if ratio >= 0.92:
        return ratio
    shorter, longer = sorted((left_token, right_token), key=len)
    if len(shorter) >= 5 and longer.startswith(shorter) and (len(longer) - len(shorter)) <= 2:
        return 0.9
    return 0.0

def _pair_semantic_tokens(reference_sig: SemanticLabelSignature, candidate_sig: SemanticLabelSignature) -> list[tuple[int, int, float]]:
    matches: list[tuple[int, int, float]] = []
    matched_reference: set[int] = set()
    matched_candidate: set[int] = set()
    for exact_only in (True, False):
        for ref_idx, reference_token in enumerate(reference_sig.tokens):
            if ref_idx in matched_reference:
                continue
            best_candidate_idx = None
            best_score = 0.0
            for candidate_idx, candidate_token in enumerate(candidate_sig.tokens):
                if candidate_idx in matched_candidate:
                    continue
                score = 1.0 if reference_token == candidate_token else 0.0
                if not exact_only and score == 0.0:
                    score = _semantic_token_match_score(reference_token, candidate_token)
                if score > best_score:
                    best_score = score
                    best_candidate_idx = candidate_idx
            if best_candidate_idx is None or best_score <= 0.0:
                continue
            matches.append((ref_idx, best_candidate_idx, best_score))
            matched_reference.add(ref_idx)
            matched_candidate.add(best_candidate_idx)
    return matches

def _forward_semantic_label_match_score(reference_label, candidate_label, allow_family: bool = False) -> float:
    reference_sig = build_semantic_label_signature(reference_label)
    candidate_sig = build_semantic_label_signature(candidate_label)
    if not reference_sig.tokens or not candidate_sig.tokens:
        return float('-inf')
    if reference_sig.compact and reference_sig.compact == candidate_sig.compact:
        return 1000.0
    if (reference_sig.numeric_tokens or candidate_sig.numeric_tokens) and reference_sig.numeric_tokens != candidate_sig.numeric_tokens:
        return float('-inf')
    token_matches = _pair_semantic_tokens(reference_sig, candidate_sig)
    if not token_matches:
        return float('-inf')
    matched_count = len(token_matches)
    matched_alpha_count = sum(
        1
        for ref_idx, candidate_idx, _ in token_matches
        if any(ch.isalpha() for ch in reference_sig.tokens[ref_idx])
        and any(ch.isalpha() for ch in candidate_sig.tokens[candidate_idx])
    )
    if reference_sig.alpha_tokens and matched_alpha_count == 0:
        return float('-inf')
    coverage_reference = matched_count / len(reference_sig.tokens)
    coverage_candidate = matched_count / len(candidate_sig.tokens)
    alpha_reference = len(reference_sig.alpha_tokens)
    alpha_candidate = len(candidate_sig.alpha_tokens)
    alpha_coverage_reference = matched_alpha_count / alpha_reference if alpha_reference else 0.0
    alpha_coverage_candidate = matched_alpha_count / alpha_candidate if alpha_candidate else 0.0
    avg_similarity = sum(score for _, _, score in token_matches) / matched_count
    if coverage_reference == 1.0 and coverage_candidate == 1.0 and alpha_coverage_reference == 1.0 and alpha_coverage_candidate == 1.0:
        return 900.0 + avg_similarity * 50.0
    if not allow_family:
        return float('-inf')
    if coverage_reference == 1.0 and alpha_coverage_reference == 1.0 and coverage_candidate >= 0.6 and alpha_coverage_candidate >= 0.5:
        extra_candidate_tokens = max(0, len(candidate_sig.tokens) - matched_count)
        return 700.0 + avg_similarity * 40.0 - extra_candidate_tokens * 5.0
    return float('-inf')

def _best_semantic_reference_match(label, reference_labels: List[str], allow_family: bool = False) -> Optional[SemanticLabelMatch]:
    best_match: Optional[SemanticLabelMatch] = None
    second_best_score = float('-inf')
    for reference_index, reference_label in enumerate(reference_labels):
        score = _forward_semantic_label_match_score(reference_label, label, allow_family=allow_family)
        if score == float('-inf'):
            continue
        if best_match is None or score > best_match.score:
            second_best_score = best_match.score if best_match is not None else float('-inf')
            best_match = SemanticLabelMatch(
                reference_label=str(reference_label),
                reference_index=reference_index,
                score=score,
            )
        elif score > second_best_score:
            second_best_score = score
    if best_match is None:
        return None
    if second_best_score != float('-inf') and abs(best_match.score - second_best_score) < 1e-6 and best_match.score < 900.0:
        return None
    return best_match

def _reorder_labels_against_reference(labels: List[str], reference_labels: List[str], allow_family: bool = False) -> List[str]:
    if not labels or not reference_labels:
        return labels
    matched_items = []
    unmatched_items = []
    for original_index, label in enumerate(labels):
        match = _best_semantic_reference_match(label, reference_labels, allow_family=allow_family)
        if match is None:
            unmatched_items.append((original_index, label))
            continue
        matched_items.append((match.reference_index, original_index, label))
    if not matched_items:
        return labels
    matched_items.sort(key=lambda item: (item[0], item[1]))
    reordered = [label for _, _, label in matched_items]
    reordered.extend(label for _, label in unmatched_items)
    if reordered == labels:
        return labels
    return reordered

def _extract_semantic_family_key(label, reference_labels: List[str]) -> str:
    candidate_sig = build_semantic_label_signature(label)
    if not candidate_sig.tokens or not reference_labels:
        return ''
    family_tokens: list[str] = []
    for candidate_token in candidate_sig.tokens:
        matched = False
        for reference_label in reference_labels:
            reference_sig = build_semantic_label_signature(reference_label)
            for reference_token in reference_sig.tokens:
                if _semantic_token_match_score(candidate_token, reference_token) > 0.0:
                    matched = True
                    break
            if matched:
                break
        if matched:
            family_tokens.append(candidate_token)
    return ''.join(family_tokens)

def _build_family_variant_palette(base_color: str, count: int) -> list[str]:
    if count <= 0:
        return []
    try:
        r, g, b = mcolors.to_rgb(base_color)
    except Exception:
        return [base_color] * count
    hue, lightness, saturation = colorsys.rgb_to_hls(r, g, b)
    if count == 1:
        rgb_variant = colorsys.hls_to_rgb(hue, min(0.72, max(0.28, lightness)), max(0.2, saturation))
        return [mcolors.to_hex(rgb_variant)]
    offsets = np.linspace(0.12, -0.12, count)
    variants: list[str] = []
    for offset in offsets:
        variant_lightness = min(0.78, max(0.24, lightness + float(offset)))
        variant_saturation = min(0.95, max(0.18, saturation))
        rgb_variant = colorsys.hls_to_rgb(hue, variant_lightness, variant_saturation)
        variants.append(mcolors.to_hex(rgb_variant))
    return variants

def _select_ambiguous_family_base_color(family_key: str) -> str:
    if not family_key:
        return AMBIGUOUS_FAMILY_BASE_PALETTE[0]
    deterministic_uuid = uuid.uuid5(uuid.NAMESPACE_DNS, family_key)
    index = deterministic_uuid.int % len(AMBIGUOUS_FAMILY_BASE_PALETTE)
    return AMBIGUOUS_FAMILY_BASE_PALETTE[index]

def _normalize_sort_header(value) -> str:
    """Devuelve una clave canónica segura para deduplicar variaciones triviales del mismo label."""
    signature = build_semantic_label_signature(value)
    return signature.compact or signature.canonical

def _build_compras_sort_rank(apo_df: pd.DataFrame) -> list[str]:
    """Devuelve el ranking de encabezados segun Compras (de mayor a menor), preservando labels originales."""
    if apo_df is None or apo_df.empty or apo_df.shape[1] <= 1:
        return []
    total_col = next(
        (col for col in apo_df.columns[1:] if str(col).strip().lower() == 'total'),
        None
    )
    vol_rows = [
        idx for idx, label in enumerate(apo_df.iloc[:, 0])
        if str(label).strip().lower().startswith('vol')
    ]
    # Preferir Vol MAT actual (ultima fila Vol disponible).
    reference_row = vol_rows[-1] if vol_rows else 0
    rank_candidates = []
    for col_idx in range(1, len(apo_df.columns)):
        col_name = apo_df.columns[col_idx]
        if total_col is not None and col_name == total_col:
            continue
        norm_name = _normalize_sort_header(col_name)
        if not norm_name:
            continue
        score = _parse_percent_cell(apo_df.iloc[reference_row, col_idx])
        if score is None:
            score = float('-inf')
        rank_candidates.append((str(col_name), norm_name, score, col_idx))
    rank_candidates.sort(key=lambda item: (item[2], -item[3]), reverse=True)
    rank = []
    seen = set()
    for raw_label, norm_name, _, _ in rank_candidates:
        if norm_name in seen:
            continue
        seen.add(norm_name)
        rank.append(raw_label)
    return rank

def reorder_apo_entries_by_compras(apo_entries: List[dict]) -> None:
    """
    Reordena encabezados de tablas de aportes usando Compras como referencia.
    Solo aplica cuando hay coincidencia de encabezados normalizados.
    """
    if not apo_entries:
        return
    compras_entry = None
    for entry in apo_entries:
        display_tipo = str(entry.get("display_tipo", "")).strip().lower()
        apo_df = entry.get("apo")
        if display_tipo == 'compras' and isinstance(apo_df, pd.DataFrame):
            compras_entry = entry
            break
    if compras_entry is None:
        return
    compras_df = compras_entry.get("apo")
    if not isinstance(compras_df, pd.DataFrame):
        return
    compras_rank = _build_compras_sort_rank(compras_df)
    if not compras_rank:
        return
    for entry in apo_entries:
        apo_df = entry.get("apo")
        if not isinstance(apo_df, pd.DataFrame) or apo_df.empty or apo_df.shape[1] <= 1:
            continue
        first_col = apo_df.columns[0]
        total_cols = [col for col in apo_df.columns[1:] if str(col).strip().lower() == 'total']
        value_cols = [col for col in apo_df.columns[1:] if str(col).strip().lower() != 'total']
        reordered_value_cols = _reorder_labels_against_reference(
            [str(col) for col in value_cols],
            compras_rank,
            allow_family=True,
        )
        if reordered_value_cols == [str(col) for col in value_cols]:
            continue
        raw_to_original: dict[str, list] = {}
        for col in value_cols:
            raw_to_original.setdefault(str(col), []).append(col)
        reordered_cols = [first_col]
        for reordered_label in reordered_value_cols:
            original_candidates = raw_to_original.get(reordered_label) or []
            if not original_candidates:
                continue
            reordered_cols.append(original_candidates.pop(0))
        for original_candidates in raw_to_original.values():
            reordered_cols.extend(original_candidates)
        reordered_cols.extend(total_cols)
        if list(apo_df.columns) == reordered_cols:
            continue
        reordered_df = apo_df.loc[:, reordered_cols].copy()
        reordered_df.attrs = dict(getattr(apo_df, 'attrs', {}))
        entry["apo"] = reordered_df

def _build_share_rank_from_compras(share_chart_entries: List[dict]) -> list[str]:
    """Construye ranking de encabezados para share usando Compras (periodo mas reciente)."""
    compras_entry = None
    for entry in share_chart_entries:
        display_tipo = str(entry.get("display_tipo", "")).strip().lower()
        periods = entry.get("periods") or []
        if display_tipo == 'compras' and periods:
            compras_entry = entry
            break
    if compras_entry is None:
        return []
    periods = compras_entry.get("periods") or []
    latest_shares = periods[-1][1] if periods else {}
    if not isinstance(latest_shares, dict) or not latest_shares:
        return []
    rank_candidates = []
    for idx, (column_name, value) in enumerate(latest_shares.items()):
        norm_name = _normalize_sort_header(column_name)
        if not norm_name:
            continue
        score = _parse_percent_cell(value)
        if score is None:
            score = float('-inf')
        rank_candidates.append((str(column_name), norm_name, score, idx))
    rank_candidates.sort(key=lambda item: (item[2], -item[3]), reverse=True)
    rank: list[str] = []
    seen: set[str] = set()
    for raw_label, norm_name, _, _ in rank_candidates:
        if norm_name in seen:
            continue
        seen.add(norm_name)
        rank.append(raw_label)
    return rank

def _reorder_labels_by_rank(labels: List[str], compras_rank: List[str]) -> List[str]:
    """Reordena etiquetas manteniendo no-coincidentes al final en su orden original."""
    if not labels or not compras_rank:
        return labels
    return _reorder_labels_against_reference(labels, compras_rank, allow_family=True)

def _reorder_shares_by_rank(shares: dict, compras_rank: List[str]) -> OrderedDict:
    """Reordena un diccionario de shares usando el ranking de Compras."""
    if not isinstance(shares, dict):
        return OrderedDict()
    ordered_labels = _reorder_labels_by_rank(list(shares.keys()), compras_rank)
    if ordered_labels == list(shares.keys()):
        return OrderedDict(shares.items())
    reordered = OrderedDict()
    for label in ordered_labels:
        reordered[label] = shares[label]
    return reordered

def reorder_share_entries_by_compras(share_chart_entries: List[dict]) -> None:
    """Aplica a share 100% la misma regla de orden: coincidencias se ordenan por Compras."""
    if not share_chart_entries:
        return
    compras_rank = _build_share_rank_from_compras(share_chart_entries)
    if not compras_rank:
        return
    for entry in share_chart_entries:
        columns = entry.get("columns") or []
        if isinstance(columns, list) and columns:
            entry["columns"] = _reorder_labels_by_rank(columns, compras_rank)
        periods = entry.get("periods") or []
        reordered_periods = []
        for period_label, shares in periods:
            reordered_periods.append((period_label, _reorder_shares_by_rank(shares, compras_rank)))
        entry["periods"] = reordered_periods

def graf_apo(apo, c_fig, column_color_mapping=None):
    """Renderiza la tabla de aporte como imagen PNG usando dibujo manual para evitar el costo de matplotlib.table."""
    n_rows, n_cols = apo.shape
    export_dpi = TABLE_APORTE_EXPORT_DPI
    if n_rows == 0 or n_cols == 0:
        fig = plt.figure(num=c_fig, figsize=(2.5, 1.2), dpi=export_dpi)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis('off')
        return figure_to_stream(fig, dpi=export_dpi, bbox_inches=None, pad_inches=0)

    display_data = apo.copy()
    if display_data.shape[1] > 0:
        display_data.iloc[:, 0] = display_data.iloc[:, 0].map(wrap_table_text)
    display_col_labels = [wrap_table_text(str(label)) for label in apo.columns]

    header_main = HEADER_COLOR_PRIMARY
    header_secondary = HEADER_COLOR_SECONDARY
    total_fill = HEADER_TOTAL_FILL
    first_col_fill = HEADER_FIRST_COL_FILL
    positive_color = TABLE_POSITIVE_COLOR
    negative_color = TABLE_NEGATIVE_COLOR
    column_color_mapping = column_color_mapping or {}

    exact_color_mapping = {
        str(key).strip(): value
        for key, value in column_color_mapping.items()
        if key is not None and value
    }
    normalized_color_mapping = {
        key.lower(): value
        for key, value in exact_color_mapping.items()
        if key
    }
    suffix_variants = ('.c', '.v', '_c', '_v', '-c', '-v')

    def resolve_column_color(column_name):
        if column_name is None:
            return None
        key = str(column_name).strip()
        if not key:
            return None
        color_value = exact_color_mapping.get(key)
        if color_value:
            return color_value
        lower_key = key.lower()
        color_value = normalized_color_mapping.get(lower_key)
        if color_value:
            return color_value
        for suffix in suffix_variants:
            if lower_key.endswith(suffix):
                base = key[:-len(suffix)].strip()
                if not base:
                    continue
                color_value = exact_color_mapping.get(base)
                if color_value:
                    return color_value
                color_value = normalized_color_mapping.get(base.lower())
                if color_value:
                    return color_value
        return None

    def to_percentage(value):
        try:
            text_value = str(value).strip().replace('%', '').replace(',', '.')
            if not text_value:
                return None
            return max(0.0, float(text_value) / 100.0)
        except Exception:
            return None

    def to_signed_percentage(value):
        try:
            text_value = str(value).strip().replace('%', '').replace(',', '.')
            if not text_value:
                return None
            return float(text_value) / 100.0
        except Exception:
            return None

    def text_metrics(value) -> tuple[str, int, int]:
        if value is None or (isinstance(value, float) and math.isnan(value)):
            text = ""
        else:
            text = str(value)
        lines = text.splitlines() if text else [""]
        line_count = len(lines)
        max_len = max((len(line.strip()) for line in lines), default=0)
        return text, line_count, max_len

    total_column_name = next((col for col in apo.columns if str(col).strip().lower() == 'total'), None)
    data_columns = [
        idx
        for idx in range(1, n_cols)
        if total_column_name is None or apo.columns[idx] != total_column_name
    ]
    start_rgb = np.array(mcolors.to_rgb(VOLUME_BAR_START))
    end_rgb = np.array(mcolors.to_rgb(VOLUME_BAR_END))
    def share_colors(count: int):
        if count <= 0:
            return []
        if count == 1:
            return [mcolors.to_hex(np.clip(start_rgb, 0, 1))]
        colors = []
        for i in range(count):
            ratio = i / (count - 1)
            rgb = np.clip(start_rgb + (end_rgb - start_rgb) * ratio, 0, 1)
            colors.append(mcolors.to_hex(rgb))
        return colors

    fallback_column_colors = share_colors(len(data_columns))
    column_colors: dict[int, str] = {}
    for rel_idx, col_idx in enumerate(data_columns):
        resolved = resolve_column_color(apo.columns[col_idx])
        if resolved is None:
            resolved = fallback_column_colors[rel_idx] if rel_idx < len(fallback_column_colors) else VOLUME_BAR_END
        try:
            column_colors[col_idx] = mcolors.to_hex(resolved)
        except (ValueError, TypeError):
            column_colors[col_idx] = fallback_column_colors[rel_idx] if rel_idx < len(fallback_column_colors) else VOLUME_BAR_END

    volume_row_indexes = [idx for idx, label in enumerate(apo.iloc[:, 0]) if str(label).lower().startswith('vol')]
    aporte_row_index = next((idx for idx, label in enumerate(apo.iloc[:, 0]) if str(label).strip().lower() == 'aporte'), None)

    aporte_fill_colors: dict[tuple[int, int], object] = {}
    if aporte_row_index is not None:
        aporte_values = []
        for col_idx in data_columns:
            numeric = to_signed_percentage(apo.iloc[aporte_row_index, col_idx])
            if numeric is not None:
                aporte_values.append((col_idx, numeric))
        if aporte_values:
            aporte_numbers = [value for _, value in aporte_values]
            min_val = min(aporte_numbers)
            max_val = max(aporte_numbers)
            cmap = None
            norm = None
            if not math.isclose(min_val, max_val, rel_tol=1e-9, abs_tol=1e-9):
                if min_val < 0 and max_val > 0:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_diverging',
                        [APORTE_NEGATIVE_COLOR, APORTE_NEUTRAL_COLOR, TABLE_POSITIVE_COLOR]
                    )
                    norm = mcolors.TwoSlopeNorm(vmin=min_val, vcenter=0, vmax=max_val)
                elif max_val <= 0:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_negative',
                        [APORTE_NEGATIVE_COLOR, APORTE_NEUTRAL_COLOR]
                    )
                    norm = mcolors.Normalize(vmin=min_val, vmax=0)
                else:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_positive',
                        [APORTE_NEUTRAL_COLOR, TABLE_POSITIVE_COLOR]
                    )
                    norm = mcolors.Normalize(vmin=0, vmax=max_val)
            for col_idx, numeric in aporte_values:
                if cmap is not None and norm is not None:
                    rgba = cmap(norm(numeric))
                    aporte_fill_colors[(aporte_row_index, col_idx)] = tuple(rgba[:3])
                else:
                    if math.isclose(numeric, 0.0, abs_tol=1e-9):
                        aporte_fill_colors[(aporte_row_index, col_idx)] = mcolors.to_rgb(APORTE_NEUTRAL_COLOR)
                    elif numeric > 0:
                        aporte_fill_colors[(aporte_row_index, col_idx)] = mcolors.to_rgb(TABLE_POSITIVE_COLOR)
                    else:
                        aporte_fill_colors[(aporte_row_index, col_idx)] = mcolors.to_rgb(APORTE_NEGATIVE_COLOR)

    volume_bar_values: dict[tuple[int, int], float] = {}
    volume_row_maxima: dict[int, float] = {}
    for df_row_idx in volume_row_indexes:
        row_percentages: dict[int, Optional[float]] = {}
        max_percent = 0.0
        for col_idx in data_columns:
            percent = to_percentage(apo.iloc[df_row_idx, col_idx])
            row_percentages[col_idx] = percent
            if percent is not None and percent > max_percent:
                max_percent = percent
        volume_row_maxima[df_row_idx] = max_percent
        for col_idx in data_columns:
            percent = row_percentages.get(col_idx)
            if percent is None:
                continue
            volume_bar_values[(df_row_idx, col_idx)] = percent

    # Cada fila de share conserva su propio ranking 1-5 para poder comparar
    # visualmente el orden del MAT anterior contra el MAT actual.
    volume_share_all_ranks: dict[tuple[int, int], int] = {}
    volume_share_ranks: dict[tuple[int, int], int] = {}
    for df_row_idx in volume_row_indexes:
        ranked_row_shares = [
            (col_idx, volume_bar_values[(df_row_idx, col_idx)])
            for col_idx in data_columns
            if (df_row_idx, col_idx) in volume_bar_values
        ]
        ranked_row_shares.sort(key=lambda item: (-item[1], item[0]))
        for rank, (col_idx, _) in enumerate(ranked_row_shares, start=1):
            share_key = (df_row_idx, col_idx)
            volume_share_all_ranks[share_key] = rank
            if rank <= SHARE_BAR_TOP_COUNT:
                volume_share_ranks[share_key] = rank

    # El movimiento se muestra solo en el periodo actual y compara su posicion
    # contra la fila de share inmediatamente anterior.
    volume_share_movements: dict[tuple[int, int], str] = {}
    ordered_volume_rows = sorted(volume_row_indexes)
    if len(ordered_volume_rows) >= 2:
        previous_volume_row = ordered_volume_rows[-2]
        latest_volume_row = ordered_volume_rows[-1]
        for share_key, current_rank in volume_share_ranks.items():
            row_idx, col_idx = share_key
            if row_idx != latest_volume_row:
                continue
            previous_rank = volume_share_all_ranks.get((previous_volume_row, col_idx))
            if previous_rank is None or current_rank == previous_rank:
                movement = 'same'
            elif current_rank < previous_rank:
                movement = 'up'
            else:
                movement = 'down'
            volume_share_movements[share_key] = movement

    # El borde negro continua destacando solo el top 5 vigente. Como
    # ``graf_apo`` se invoca por separado para Compras y Ventas, cada bloque
    # conserva su propio ranking y nunca compite con el otro.
    top_share_cells: set[tuple[int, int]] = set()
    if volume_row_indexes:
        latest_volume_row = max(volume_row_indexes)
        top_share_cells = {
            share_key
            for share_key in volume_share_ranks
            if share_key[0] == latest_volume_row
        }

    header_metrics = [text_metrics(label) for label in display_col_labels]
    body_text_matrix: list[list[str]] = []
    body_line_counts: list[int] = []
    col_body_max_len = [0] * n_cols
    for row_idx in range(n_rows):
        row_values = []
        row_max_lines = 1
        for col_idx in range(n_cols):
            raw_value = display_data.iat[row_idx, col_idx] if col_idx == 0 else apo.iat[row_idx, col_idx]
            text_value, line_count, max_len = text_metrics(raw_value)
            row_values.append(text_value)
            row_max_lines = max(row_max_lines, line_count)
            if max_len > col_body_max_len[col_idx]:
                col_body_max_len[col_idx] = max_len
        body_text_matrix.append(row_values)
        body_line_counts.append(row_max_lines)

    base_row_height_in = 0.42
    header_line_count = max((line_count for _, line_count, _ in header_metrics), default=1)
    header_height_factor = min(2.0, 1.0 + 0.45 * (header_line_count - 1))
    header_height_in = base_row_height_in * header_height_factor
    row_heights_in = [
        base_row_height_in * min(1.45, 1.0 + 0.25 * (line_count - 1))
        for line_count in body_line_counts
    ]

    col_widths_in: list[float] = []
    for col_idx in range(n_cols):
        _, header_lines, header_max_len = header_metrics[col_idx]
        longest = max(header_max_len, col_body_max_len[col_idx])
        multiline_bonus = max(0, header_lines - 1) * 0.12
        if col_idx == 0:
            width_in = max(1.55, min(2.85, 1.0 + longest * 0.055 + multiline_bonus))
        elif total_column_name is not None and apo.columns[col_idx] == total_column_name:
            width_in = max(1.05, min(1.55, 0.82 + longest * 0.05 + multiline_bonus))
        else:
            width_in = max(1.35, min(2.35, 0.95 + longest * 0.07 + multiline_bonus))
        col_widths_in.append(width_in)

    fig_width = sum(col_widths_in) + 0.06
    fig_height = header_height_in + sum(row_heights_in) + 0.06
    fig = plt.figure(num=c_fig, figsize=(fig_width, fig_height), dpi=export_dpi)
    fig.patch.set_alpha(0.0)
    ax = fig.add_axes([0.004, 0.004, 0.992, 0.992])
    ax.set_xlim(0, fig_width)
    ax.set_ylim(fig_height, 0)
    ax.axis('off')
    ax.set_facecolor('none')

    x_edges = [0.03]
    for width_in in col_widths_in:
        x_edges.append(x_edges[-1] + width_in)
    current_y = 0.03

    def text_color_for_fill(fill_value, fallback=TABLE_TEXT_PRIMARY):
        try:
            rgb = mcolors.to_rgb(fill_value)
        except (ValueError, TypeError):
            return fallback
        luminance = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
        return '#FFFFFF' if luminance < 0.6 else TABLE_TEXT_PRIMARY

    def draw_cell(x0, y0, width_in, height_in, text_value, facecolor, text_color, *, bold=False, font_size=TABLE_BODY_FONT_SIZE):
        ax.add_patch(
            Rectangle(
                (x0, y0),
                width_in,
                height_in,
                facecolor=facecolor,
                edgecolor=TABLE_GRID_COLOR,
                linewidth=1.0,
                antialiased=False,
            )
        )
        ax.text(
            x0 + width_in / 2.0,
            y0 + height_in / 2.0,
            text_value,
            ha='center',
            va='center',
            fontsize=font_size,
            fontweight='bold' if bold else 'normal',
            color=text_color,
            family='DejaVu Sans',
        )

    def draw_share_cell(
        x0,
        y0,
        width_in,
        height_in,
        text_value,
        share_value,
        max_share_value,
        bar_color,
        *,
        highlight=False,
        share_rank=None,
        share_movement=None,
    ):
        """Dibuja una celda de share con barra proporcional y ranking visual."""
        ax.add_patch(
            Rectangle(
                (x0, y0),
                width_in,
                height_in,
                facecolor='white',
                edgecolor=TABLE_GRID_COLOR,
                linewidth=1.0,
                antialiased=False,
            )
        )
        horizontal_padding = width_in * 0.09
        track_x = x0 + horizontal_padding
        track_width = max(width_in - 2.0 * horizontal_padding, width_in * 0.10)
        track_height = height_in * 0.32
        track_y = y0 + height_in * 0.12
        track_radius = min(track_height * 0.24, track_width * 0.08)
        track_edge = SHARE_BAR_TOP_EDGE if highlight else SHARE_BAR_TRACK_EDGE
        track_linewidth = 1.2 if highlight else 0.8
        ax.add_patch(
            mpatches.FancyBboxPatch(
                (track_x, track_y),
                track_width,
                track_height,
                boxstyle=mpatches.BoxStyle(
                    "Round",
                    pad=0,
                    rounding_size=track_radius,
                ),
                facecolor=SHARE_BAR_TRACK_COLOR,
                edgecolor=track_edge,
                linewidth=track_linewidth,
                antialiased=True,
            )
        )
        if share_value is not None and share_value > 0 and max_share_value > 0:
            scale_max = max_share_value * SHARE_BAR_HEADROOM
            fill_ratio = min(max(share_value / scale_max, 0.0), 1.0)
            fill_width = track_width * fill_ratio
            fill_radius = min(track_height * 0.24, fill_width * 0.24)
            ax.add_patch(
                mpatches.FancyBboxPatch(
                    (track_x, track_y),
                    fill_width,
                    track_height,
                    boxstyle=mpatches.BoxStyle(
                        "Round",
                        pad=0,
                        rounding_size=fill_radius,
                    ),
                    facecolor=bar_color,
                    edgecolor='none',
                    linewidth=0,
                    antialiased=True,
                )
            )
        if highlight:
            ax.add_patch(
                mpatches.FancyBboxPatch(
                    (track_x, track_y),
                    track_width,
                    track_height,
                    boxstyle=mpatches.BoxStyle(
                        "Round",
                        pad=0,
                        rounding_size=track_radius,
                    ),
                    facecolor='none',
                    edgecolor=SHARE_BAR_TOP_EDGE,
                    linewidth=track_linewidth,
                    antialiased=True,
                )
            )
        if share_movement is not None:
            percentage_x = x0 + width_in * 0.40
        elif share_rank is not None:
            percentage_x = x0 + width_in * 0.45
        else:
            percentage_x = x0 + width_in * 0.50
        ax.text(
            percentage_x,
            y0 + height_in * 0.72,
            text_value,
            ha='center',
            va='center',
            fontsize=TABLE_BODY_FONT_SIZE,
            fontweight='bold',
            color=bar_color,
            family='DejaVu Sans',
        )
        if share_rank is not None:
            rank_x = 0.67 if share_movement is not None else 0.72
            ax.text(
                x0 + width_in * rank_x,
                y0 + height_in * 0.72,
                str(share_rank),
                ha='center',
                va='center',
                fontsize=max(TABLE_BODY_FONT_SIZE - 1, 8),
                fontweight='bold',
                color=SHARE_RANK_COLORS.get(share_rank, SHARE_RANK_COLORS[5]),
                family='DejaVu Sans',
            )
        if share_movement is not None:
            ax.text(
                x0 + width_in * 0.82,
                y0 + height_in * 0.72,
                SHARE_RANK_MOVEMENT_SYMBOLS.get(share_movement, '-'),
                ha='center',
                va='center',
                fontsize=max(TABLE_BODY_FONT_SIZE - 1, 8),
                fontweight='bold',
                color=SHARE_RANK_MOVEMENT_COLORS.get(
                    share_movement,
                    SHARE_RANK_MOVEMENT_COLORS['same'],
                ),
                family='DejaVu Sans',
            )

    for col_idx in range(n_cols):
        if col_idx == 0 or (total_column_name is not None and apo.columns[col_idx] == total_column_name):
            header_fill = header_main
            header_text_color = HEADER_FONT_COLOR
        else:
            header_fill = resolve_column_color(apo.columns[col_idx]) or header_secondary
            header_text_color = text_color_for_fill(header_fill, HEADER_FONT_COLOR)
        draw_cell(
            x_edges[col_idx],
            current_y,
            col_widths_in[col_idx],
            header_height_in,
            header_metrics[col_idx][0],
            header_fill,
            header_text_color,
            bold=True,
            font_size=TABLE_HEADER_FONT_SIZE,
        )
    current_y += header_height_in

    for row_idx in range(n_rows):
        row_height_in = row_heights_in[row_idx]
        row_label = str(apo.iloc[row_idx, 0]).strip().lower()
        for col_idx in range(n_cols):
            text_value = body_text_matrix[row_idx][col_idx]
            facecolor = 'white'
            text_color = TABLE_TEXT_PRIMARY
            is_bold = False

            if col_idx == 0:
                facecolor = first_col_fill
                text_color = TABLE_TEXT_PRIMARY
                is_bold = True
            else:
                if total_column_name is not None and apo.columns[col_idx] == total_column_name:
                    facecolor = total_fill
                if (row_idx, col_idx) in aporte_fill_colors:
                    facecolor = aporte_fill_colors[(row_idx, col_idx)]
                    text_color = TABLE_TEXT_PRIMARY if text_color_for_fill(facecolor) != '#FFFFFF' else '#FFFFFF'
                elif row_label.startswith('vol'):
                    text_color = TABLE_TEXT_PRIMARY
                else:
                    try:
                        numeric = float(str(apo.iat[row_idx, col_idx]).replace('%', '').replace(',', '.'))
                        text_color = positive_color if numeric >= 0 else negative_color
                    except Exception:
                        text_color = TABLE_TEXT_PRIMARY
                if row_label.startswith('var'):
                    is_bold = True

            share_key = (row_idx, col_idx)
            if col_idx > 0 and share_key in volume_bar_values:
                draw_share_cell(
                    x_edges[col_idx],
                    current_y,
                    col_widths_in[col_idx],
                    row_height_in,
                    text_value,
                    volume_bar_values[share_key],
                    volume_row_maxima.get(row_idx, 0.0),
                    column_colors.get(col_idx, VOLUME_BAR_END),
                    highlight=share_key in top_share_cells,
                    share_rank=volume_share_ranks.get(share_key),
                    share_movement=volume_share_movements.get(share_key),
                )
            else:
                draw_cell(
                    x_edges[col_idx],
                    current_y,
                    col_widths_in[col_idx],
                    row_height_in,
                    text_value,
                    facecolor,
                    text_color,
                    bold=is_bold,
                    font_size=TABLE_BODY_FONT_SIZE,
                )
        current_y += row_height_in

    return figure_to_stream(fig, dpi=export_dpi, bbox_inches=None, pad_inches=0)
def simplify_name_segment(value: str, max_len: int) -> str:
    """
    Reduce una cadena a un segmento seguro para archivos.
    Conserva alfanumericos, reemplaza otros con guiones y limita longitud.
    """
    if value is None:
        return 'NA'
    value_str = str(value).strip()
    if not value_str:
        return 'NA'
    cleaned = ''.join(ch if ch.isalnum() else '-' for ch in value_str)
    cleaned = cleaned.strip('-')
    if not cleaned:
        cleaned = 'NA'
    return cleaned[:max_len]
def simplify_folder_segment(value: str, max_len: int = 80) -> str:
    """
    Limpia un segmento para nombre de carpeta conservando espacios.
    Reemplaza caracteres invalidos de ruta por guion y limita longitud.
    """
    if value is None:
        return 'NA'
    value_str = str(value).strip()
    if not value_str:
        return 'NA'
    invalid_chars = '<>:"/\\|?*'
    cleaned = ''.join('-' if (ch in invalid_chars or ord(ch) < 32) else ch for ch in value_str)
    cleaned = ' '.join(cleaned.split()).strip(' .')
    if not cleaned:
        cleaned = 'NA'
    return cleaned[:max_len]
def is_locked_for_write(file_path: Path) -> bool:
    """Indica si un archivo existente parece bloqueado para escritura."""
    if not file_path.exists():
        return False
    try:
        with open(file_path, 'a+b'):
            pass
        return False
    except PermissionError:
        return True
    except OSError:
        return False
def select_excel_file(base_dir: Path) -> str:
    """Lista archivos Excel disponibles y devuelve el nombre seleccionado."""
    excel_files = sorted([p for p in base_dir.glob('*.xlsx') if not p.name.startswith('~$')])
    brand_color_cache: Dict[str, Tuple[int, int, int]] = {}
    def warn_if_excel_in_use(file_path: Path) -> None:
        """Avisa si el workbook parece estar abierto en Excel."""
        lock_path = file_path.with_name(f"~${file_path.name}")
        if lock_path.exists():
            print_colored(
                (
                    f"Advertencia: el archivo '{file_path.name}' parece estar en uso "
                    "(detectado archivo temporal de bloqueo). "
                    "Si tienes cambios sin guardar, no se veran reflejados hasta guardar/cerrar el Excel."
                ),
                COLOR_YELLOW
            )
    def metadata_for(file_path: Path) -> str:
        try:
            if file_path.stem.lower() == 'plantilla_entrada_5w1h':
                return 'Plantilla de entrada - usar como referencia'
            parts = file_path.stem.split('_')
            country_code = int(parts[0]) if parts else None
            category_code = parts[1].strip().upper() if len(parts) > 1 else None
            pais_name = lookup_country_name(country_code)
            cesta, categoria = lookup_category_info(category_code)
            return f"{pais_name} | {cesta} | {categoria}"
        except Exception:
            return 'Metadata no disponible'
    if not excel_files:
        raise FileNotFoundError(colorize(f'No se encontraron archivos .xlsx en {base_dir}', COLOR_RED))
    if len(excel_files) == 1:
        meta = metadata_for(excel_files[0])
        meta_color = COLOR_RED if meta == 'Metadata no disponible' else COLOR_YELLOW
        archivo_coloreado = colorize_filename_brand(excel_files[0].name, brand_color_cache)
        print_colored(f"Archivo encontrado: {archivo_coloreado} {colorize('[' + meta + ']', meta_color)}", COLOR_BLUE)
        warn_if_excel_in_use(excel_files[0])
        return excel_files[0].name
    print_colored('Archivos Excel disponibles:', COLOR_QUESTION)
    max_name_len = max(len(archivo.name) for archivo in excel_files)
    idx_width = len(str(len(excel_files)))
    for idx, archivo in enumerate(excel_files, 1):
        meta = metadata_for(archivo)
        meta_color = COLOR_RED if meta == 'Metadata no disponible' else COLOR_YELLOW
        archivo_coloreado = colorize_filename_brand(archivo.name, brand_color_cache)
        pad = " " * (max_name_len - len(archivo.name) + 3)
        prefix = colorize(f"  {idx:>{idx_width}}. ", COLOR_BLUE)
        print(f"{prefix}{archivo_coloreado}{pad}{colorize('[' + meta + ']', meta_color)}")
    prompt = colorize(f"Seleccione el numero del archivo (1-{len(excel_files)}). Enter para {excel_files[0].name}: ", COLOR_QUESTION)
    while True:
        choice = input(prompt).strip()
        if not choice:
            warn_if_excel_in_use(excel_files[0])
            return excel_files[0].name
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(excel_files):
                selected_file = excel_files[idx - 1]
                warn_if_excel_in_use(selected_file)
                return selected_file.name
        matching = [archivo.name for archivo in excel_files if archivo.name.lower() == choice.lower()]
        if matching:
            selected_file = base_dir / matching[0]
            warn_if_excel_in_use(selected_file)
            return matching[0]
        print_colored('Entrada invalida. Intente nuevamente.', COLOR_RED)
def configure_cover_slide(presentation, lang):
    """Configura el titulo de portada segun el idioma."""
    title_by_lang = {
        'P': 'Adicionais de Cobertura - 5W1H',
        'E': 'Adicionales de Cobertura - 5W1H'
    }
    cover_slide = presentation.slides[0]
    target_text = title_by_lang.get(lang, title_by_lang['E'])
    fallback_geometry = (Inches(0.42), Inches(2.0), Inches(10), Inches(0.8))
    geometry = None
    for shape in list(cover_slide.shapes):
        if not getattr(shape, 'has_text_frame', False):
            continue
        text_value = shape.text.strip()
        if 'Cobertura' in text_value and '5W1H' in text_value:
            geometry = (shape.left, shape.top, shape.width, shape.height)
            cover_slide.shapes._spTree.remove(shape._element)
    left, top, width, height = geometry if geometry else fallback_geometry
    text_box = cover_slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.clear()
    paragraph = tf.paragraphs[0]
    paragraph.text = target_text
    run = paragraph.runs[0]
    font = run.font
    font.name = 'Arial'
    font.size = Inches(0.45)
    font.bold = True
    font.color.rgb = RGBColor(255, 255, 255)
def plot_ven():
    """Solicita el modo de grafico para ventas y compras."""
    options = [
        ("1 - Plotear Ventas y Compras Juntas", "1"),
        ("2 - Plotear Ventas y Compras Separadas", "2"),
        ("3 - No hay W con Ventas en esa base", "3"),
    ]
    print_colored('\nOpciones de grafico disponibles:', COLOR_QUESTION)
    for idx, (texto, _) in enumerate(options, 1):
        print_colored(f"  {idx}. {texto}", COLOR_BLUE)
    prompt = colorize(f"Seleccione el numero de la opcion (1-{len(options)}). Enter para {options[0][1]}: ", COLOR_QUESTION)
    while True:
        choice = input(prompt).strip()
        if not choice:
            return options[0][1]
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(options):
                return options[idx - 1][1]
        for _, valor in options:
            if choice == valor:
                return valor
        print_colored('Entrada invalida. Intente nuevamente.', COLOR_RED)
#Codigos paises
COUNTRY_LOOKUP = {
    10: 'LatAm',
    54: 'Argentina',
    91: 'Bolivia',
    55: 'Brasil',
    12: 'CAM',
    56: 'Chile',
    57: 'Colombia',
    93: 'Ecuador',
    52: 'Mexico',
    51: 'Peru',
    66: 'Costa Rica',
    63: 'El Salvador',
    62: 'Guatemala',
    64: 'Honduras',
    65: 'Nicaragua',
    67: 'Panama',
    69: 'Republica Dominicana',
}

def lookup_country_name(code: Optional[int]) -> str:
    """Resuelve el pais desde codigo numerico."""
    if code is None:
        return 'Pais N/D'
    return COUNTRY_LOOKUP.get(int(code), 'Pais N/D')
#Codigos categorias
CATEG_CSV_DATA = """
cod,cest,cat
ALCB,Bebidas,Bebidas Alcoholicas
BEER,Bebidas,Cervezas
CARB,Bebidas,Bebidas Gaseosas
CWAT,Bebidas,Agua Gasificada
COCW,Bebidas,Agua de Coco
COFF,Bebidas,Cafe-Consolidado de Cafe
CRBE,Bebidas,Cross Category (Bebidas)
ENDR,Bebidas,Bebidas Energeticas
FLBE,Bebidas,Bebidas Saborizadas Sin Gas
GCOF,Bebidas,Cafe Tostado y Molido
HJUI,Bebidas,Jugos Caseros
ITEA,Bebidas,Te Helado
ICOF,Bebidas,Cafe Instantaneo-Cafe Sucedaneo
JUNE,Bebidas,Jugos y Nectares
VEJU,Bebidas,Zumos de Vegetales
WATE,Bebidas,Agua Natural
CSDW,Bebidas,Gaseosas + Aguas
MXCM,Bebidas,Mixta Cafe+Malta
MXDG,Bebidas,Mixta Dolce Gusto-Mixta Te Helado + Cafe + Modificadores
MXJM,Bebidas,Mixta Jugos y Leches
MXJS,Bebidas,Mixta Jugos Liquidos + Bebidas de Soja
MXTC,Bebidas,Mixta Te+Cafe
JUIC,Bebidas,Jugos Liquidos-Jugos Polvo
PWDJ,Bebidas,Refrescos en Polvo-Jugos - Bebidas Instantaneas En Polvo - Jugos Polvo
RFDR,Bebidas,Bebidas Refrescantes
RTDJ,Bebidas,Refrescos Liquidos-Jugos Liquidos
RTEA,Bebidas,Te Liquido - Listo para Tomar
SOYB,Bebidas,Bebidas de Soja
SPDR,Bebidas,Bebidas Isotonicas
TEAA,Bebidas,Te e Infusiones-Te-Infusion Hierbas
YERB,Bebidas,Yerba Mate
BUTT,Lacteos,Manteca
CHEE,Lacteos,Queso Fresco y para Untar
CMLK,Lacteos,Leche Condensada
CRCH,Lacteos,Queso Untable
DYOG,Lacteos,Yoghurt p-beber
EMLK,Lacteos,Leche Culinaria-Leche Evaporada
FRMM,Lacteos,Leche Fermentada
FMLK,Lacteos,Leche Liquida Saborizada-Leche Liquida Con Sabor
FRMK,Lacteos,Formulas Infantiles
LQDM,Lacteos,Leche Liquida
LLFM,Lacteos,Leche Larga Vida
MARG,Lacteos,Margarina
MCHE,Lacteos,Queso Fundido
MKCR,Lacteos,Crema de Leche
MXDI,Lacteos,Mixta Lacteos-Postre+Leches+Yogurt
MXMI,Lacteos,Mixta Leches
MXYD,Lacteos,Mixta Yoghurt+Postres
PTSS,Lacteos,Petit Suisse
PWDM,Lacteos,Leche en Polvo
SYOG,Lacteos,Yoghurt p-comer
MILK,Lacteos,Leche-Leche Liquida Blanca - Leche Liq. Natural
YOGH,Lacteos,Yoghurt
CLOT,Ropas y Calzados,Ropas
FOOT,Ropas y Calzados,Calzados
SOCK,Ropas y Calzados,Medias-Calcetines
AREP,Alimentos,Arepas
BCER,Alimentos,Cereales Infantiles
BABF,Alimentos,Nutricion Infantil-Colados y Picados
BEAN,Alimentos,Frijoles
BISC,Alimentos,Galletas
BOUI,Alimentos,Caldos-Caldos y Sazonadores
BREA,Alimentos,Pan
BRCR,Alimentos,Apanados-Empanizadores
BRDC,Alimentos,Empanados
CERE,Alimentos,Cereales-Cereales Desayuno-Avenas y Cereales
BURG,Alimentos,Hamburguesas
CCMX,Alimentos,Mezclas Listas para Tortas-Preparados Base Harina Trigo
CAKE,Alimentos,Queques-Ponques Industrializados
FISH,Alimentos,Conservas De Pescado
CFAV,Alimentos,Conservas de Frutas y Verduras
CRML,Alimentos,Dulce de Leche-Manjar
CMLC,Alimentos,Alfajores
CBAR,Alimentos,Barras de Cereal
CHCK,Alimentos,Pollo
CHOC,Alimentos,Chocolate
COCO,Alimentos,Chocolate de Taza-Achocolatados - Cocoas
COLS,Alimentos,Salsas Frias
COMP,Alimentos,Compotas
SPIC,Alimentos,Condimentos y Especias
CKCH,Alimentos,Chocolate de Mesa
COIL,Alimentos,Aceite-Aceites Comestibles
CSAU,Alimentos,Salsas Listas-Salsas Caseras Envasadas
CNML,Alimentos,"Grano, Harina y Masa de Maiz"
CNST,Alimentos,Fecula de Maiz
CNFL,Alimentos,Harina De Maiz
CAID,Alimentos,Ayudantes Culinarios
DESS,Alimentos,Postres Preparados
DHAM,Alimentos,Jamon Endiablado
DFNS,Alimentos,Semillas y Frutos Secos
EBRE,Alimentos,Pan de Pascua
EEGG,Alimentos,Huevos de Pascua
EGGS,Alimentos,Huevos
FLSS,Alimentos,Flash Cecinas
FLOU,Alimentos,Harinas
MEAT,Alimentos,Carne Fresca
FRDS,Alimentos,Platos Listos Congelados
FRFO,Alimentos,Alimentos Congelados
HAMS,Alimentos,Jamones
HCER,Alimentos,Cereales Calientes-Cereales Precocidos
HOTS,Alimentos,Salsas Picantes
ICEC,Alimentos,Helados
IBRE,Alimentos,Pan Industrializado
IMPO,Alimentos,Pure Instantaneo
INOO,Alimentos,Fideos Instantaneos
JAMS,Alimentos,Mermeladas
KETC,Alimentos,Ketchup
LJDR,Alimentos,Jugo de Limon Adereso
MALT,Alimentos,Maltas
SEAS,Alimentos,Adobos - Sazonadores
MAYO,Alimentos,Mayonesa
MEAT,Alimentos,Carnicos
MLKM,Alimentos,Modificadores de Leche-Saborizadores p-leche
MXCO,Alimentos,Mixta Cereales Infantiles+Avenas
MXBS,Alimentos,Mixta Caldos + Saborizantes
MXSB,Alimentos,Mixta Caldos + Sopas
MXCH,Alimentos,Mixta Cereales + Cereales Calientes
MXCC,Alimentos,Mixta Chocolate + Manjar
MXSN,Alimentos,"Galletas, snacks y mini tostadas"
COBT,Alimentos,Aceites + Mantecas
COCF,Alimentos,Aceites + Conservas De Pescado
CABB,Alimentos,Ayudantes Culinarios + Bolsa de Hornear
MXEC,Alimentos,Mixta Huevos de Pascua + Chocolates
MXDP,Alimentos,Mixta Platos Listos Congelados + Pasta
MXFR,Alimentos,Mixta Platos Congelados y Listos para Comer
MXFM,Alimentos,Mixta Alimentos Congelados + Margarina
MXMC,Alimentos,Mixta Modificadores + Cocoa
MXPS,Alimentos,Mixta Pastas
MXSO,Alimentos,Mixta Sopas+Cremas+Ramen
MXSP,Alimentos,Mixta Margarina + Mayonesa + Queso Crema
MXSW,Alimentos,Mixta Azucar+Endulzantes
MUST,Alimentos,Mostaza
NDCR,Alimentos,Sustitutos de Crema
NOOD,Alimentos,Fideos
NUGG,Alimentos,Nuggets
OAFL,Alimentos,Avena en hojuelas-liquidas
OLIV,Alimentos,Aceitunas
PANC,Alimentos,Tortilla
PANE,Alimentos,Paneton
PAST,Alimentos,Pastas
PSAU,Alimentos,Salsas para Pasta
PNOU,Alimentos,Turron de mani
PORK,Alimentos,Carne Porcina
PPMX,Alimentos,Postres en Polvo-Postres para Preparar - Horneables-Gelificables
PWSM,Alimentos,Leche de Soya en Polvo
PCCE,Alimentos,Cereales Precocidos
DOUG,Alimentos,Masas Frescas-Tapas Empanadas y Tarta
PPIZ,Alimentos,Pre-Pizzas
REFR,Alimentos,Meriendas listas
RICE,Alimentos,Arroz
RBIS,Alimentos,Galletas de Arroz
RTEB,Alimentos,Frijoles Procesados
RTEM,Alimentos,Pratos Prontos - Comidas Listas
SDRE,Alimentos,Aderezos para Ensalada
SALT,Alimentos,Sal
SLTC,Alimentos,Galletas Saladas-Galletas No Dulce
SARD,Alimentos,Sardina Envasada
SAUS,Alimentos,Cecinas
SCHN,Alimentos,Milanesas
SNAC,Alimentos,Snacks
SNOO,Alimentos,Fideos Sopa
SOUP,Alimentos,Sopas-Sopas Cremas
SOYS,Alimentos,Siyau
SPAG,Alimentos,Tallarines-Spaguetti
SPCH,Alimentos,Chocolate para Untar
SUGA,Alimentos,Azucar
SWCO,Alimentos,Galletas Dulces
SWSP,Alimentos,Untables Dulces
SWEE,Alimentos,Endulzantes
TOAS,Alimentos,Torradas - Tostadas
TOMA,Alimentos,Salsas de Tomate
TUNA,Alimentos,Atun Envasado
VMLK,Alimentos,Leche Vegetal
WFLO,Alimentos,Harinas de trigo
AIRC,Cuidado del Hogar,Ambientadores-Desodorante Ambiental
BARS,Cuidado del Hogar,Jabon en Barra-Jabon de lavar
BLEA,Cuidado del Hogar,Cloro-Lavandinas-Lejias-Blanqueadores
CBLK,Cuidado del Hogar,Pastillas para Inodoro
CGLO,Cuidado del Hogar,Guantes de latex
CLSP,Cuidado del Hogar,Esponjas de Limpieza-Esponjas y panos
CLTO,Cuidado del Hogar,Utensilios de Limpieza
FILT,Cuidado del Hogar,Filtros de Cafe
CRHC,Cuidado del Hogar,Cross Category (Limpiadores Domesticos)
CRLA,Cuidado del Hogar,Cross Category (Lavanderia)
CRPA,Cuidado del Hogar,Cross Category (Productos de Papel)
DISH,Cuidado del Hogar,Lavavajillas-Lavaplatos - Lavalozas mano
DPAC,Cuidado del Hogar,Empaques domesticos-Bolsas plasticas-Plastico Adherente-Papel encerado-Papel aluminio
DRUB,Cuidado del Hogar,Destapacanerias
FBRF,Cuidado del Hogar,Perfumantes para Ropa-Perfumes para Ropa
FWAX,Cuidado del Hogar,Cera p-pisos
FDEO,Cuidado del Hogar,Desodorante para Pies
FRNP,Cuidado del Hogar,Lustramuebles
GBBG,Cuidado del Hogar,Bolsas de Basura
GCLE,Cuidado del Hogar,Limpiadores verdes
CLEA,Cuidado del Hogar,Limpiadores-Limpiadores y Desinfectantes
INSE,Cuidado del Hogar,Insecticidas-Raticidas
KITT,Cuidado del Hogar,Toallas de papel-Papel Toalla - Toallas de Cocina - Rollos Absorbentes de Papel
LAUN,Cuidado del Hogar,Detergentes para ropa
LSTA,Cuidado del Hogar,Apresto
MXBC,Cuidado del Hogar,Mixta Pastillas para Inodoro + Limpiadores
MXHC,Cuidado del Hogar,Mixta Home Care-Cloro-Limpiadores-Ceras-Ambientadores
MXCB,Cuidado del Hogar,Mixta Limpiadores + Cloro
MXLB,Cuidado del Hogar,Mixta Detergentes + Cloro
MXLD,Cuidado del Hogar,Mixta Detergentes + Lavavajillas
CRTO,Cuidado del Hogar,Panitos + Papel Higienico
NAPK,Cuidado del Hogar,Servilletas
PLWF,Cuidado del Hogar,Film plastico e papel aluminio
SCOU,Cuidado del Hogar,Esponjas de Acero
SOFT,Cuidado del Hogar,Suavizantes de Ropa
STRM,Cuidado del Hogar,Quitamanchas-Desmanchadores
TOIP,Cuidado del Hogar,Papel Higienico
WIPE,Cuidado del Hogar,Panos de Limpieza
ANLG,OTC,Analgesicos-Painkillers
FSUP,OTC,Suplementos alimentares
GMED,OTC,Gastrointestinales-Efervescentes
VITA,OTC,Vitaminas y Calcio
nan,Otros,Categoria Desconocida
BATT,Otros,Pilas-Baterias
CGAS,Otros,Combustible Gas
PFHH,Otros,Panel Financiero de Hogares
PFIN,Otros,Panel Financiero de Hogares
INKC,Otros,Cartuchos de Tintas
PETF,Otros,Alimento para Mascota-Alim.p - perro - gato
TELE,Otros,Telecomunicaciones - Convergencia
TILL,Otros,Tickets - Till Rolls
TOBA,Otros,Tabaco - Cigarrillos
ADIP,Cuidado Personal,Incontinencia de Adultos
BSHM,Cuidado Personal,Shampoo Infantil
RAZO,Cuidado Personal,Maquinas de Afeitar
BDCR,Cuidado Personal,Cremas Corporales
CWIP,Cuidado Personal,Panos Humedos
COMB,Cuidado Personal,Cremas para Peinar
COND,Cuidado Personal,Acondicionador-Balsamo
CRHY,Cuidado Personal,Cross Category (Higiene)
CRPC,Cuidado Personal,Cross Category (Personal Care)
DEOD,Cuidado Personal,Desodorantes
DIAP,Cuidado Personal,Panales-Panales Desechables
FCCR,Cuidado Personal,Cremas Faciales
FTIS,Cuidado Personal,Panuelos Faciales
FEMI,Cuidado Personal,Proteccion Femenina-Toallas Femeninas
FRAG,Cuidado Personal,Fragancias
HAIR,Cuidado Personal,Cuidado del Cabello-Hair Care
HRCO,Cuidado Personal,Tintes para el Cabello-Tintes - Tintura - Tintes y Coloracion para el cabello
HREM,Cuidado Personal,Depilacion
HRST,Cuidado Personal,Alisadores para el Cabello
HSTY,Cuidado Personal,Fijadores para el Cabello-Modeladores-Gel-Fijadores para el cabello
HRTR,Cuidado Personal,Tratamientos para el Cabello
LINI,Cuidado Personal,leo Calcareo
MAKE,Cuidado Personal,Maquillaje-Cosmeticos
MEDS,Cuidado Personal,Jabon Medicinal
CRDT,Cuidado Personal,Panitos + Panales
MXMH,Cuidado Personal,Mixta Make Up+Tinturas
MOWA,Cuidado Personal,Enjuague Bucal-Refrescante Bucal
ORAL,Cuidado Personal,Cuidado Bucal
SPAD,Cuidado Personal,Protectores Femeninos
STOW,Cuidado Personal,Toallas Femininas
SHAM,Cuidado Personal,Shampoo
SHAV,Cuidado Personal,Afeitado-Crema afeitar-Locion de afeitar-Pord. Antes del afeitado
SKCR,Cuidado Personal,Cremas Faciales y Corporales-Cremas de Belleza - Cremas Cuerp y Faciales
SUNP,Cuidado Personal,Proteccion Solar
TALC,Cuidado Personal,Talcos-Talco para pies
TAMP,Cuidado Personal,Tampones Femeninos
TOIL,Cuidado Personal,Jabon de Tocador
TOOB,Cuidado Personal,Cepillos Dentales
TOOT,Cuidado Personal,Pastas Dentales
BAGS,Material Escolar,Morrales y MAletas Escoalres
CLPC,Material Escolar,Lapices de Colores
GRPC,Material Escolar,Lapices De Grafito
MRKR,Material Escolar,Marcadores
NTBK,Material Escolar,Cuadernos
SCHS,Material Escolar,tiles Escolares
CSTD,Diversos,Estudio de Categorias
CORP,Diversos,Corporativa
CROS,Diversos,Cross Category
CRBA,Diversos,Cross Category (Bebes)
CRBR,Diversos,"Cross Category (Desayuno)-Yogurt, Cereal, Pan y Queso"
CRDT,Diversos,Cross Category (Diet y Light)
CRDF,Diversos,Cross Category (Alimentos Secos)
CRFO,Diversos,Cross Category (Alimentos)
CRSA,Diversos,Cross Category (Salsas)-Mayonesas-Ketchup - Salsas Frias
CRSN,Diversos,Cross Category (Snacks)
DEMO,Diversos,Demo
FLSH,Diversos,Flash
HLVW,Diversos,Holistic View
COCP,Diversos,Mezcla para cafe instantaneo y crema no lactea
CRSN,Diversos,Mezclas nutricionales y suplementos
MULT,Diversos,Consolidado-Multicategory
PCHK,Diversos,Pantry Check
STCK,Diversos,Inventario
MIHC,Diversos,Leche y Cereales Calientes-Cereales Precocidos y Leche Liquida Blanca
FLWT,Alimentos,Agua Saborizada"""
def _build_category_lookup(csv_data: str) -> dict[str, tuple[str, str]]:
    """Construye lookup cod -> (cesta, categoria) preservando la primera ocurrencia."""
    lookup: dict[str, tuple[str, str]] = {}
    reader = csv.DictReader(io.StringIO(csv_data.lstrip()))
    for row in reader:
        raw_code = (row.get('cod') or '').strip().upper()
        if not raw_code or raw_code in lookup:
            continue
        cesta = (row.get('cest') or '').strip() or 'Cesta N/D'
        categoria = (row.get('cat') or '').strip() or 'Categoria N/D'
        lookup[raw_code] = (cesta, categoria)
    return lookup

CATEGORY_LOOKUP = _build_category_lookup(CATEG_CSV_DATA)

def lookup_category_info(code: Optional[str]) -> tuple[str, str]:
    """Devuelve (cesta, categoria) para el codigo indicado."""
    normalized = str(code).strip().upper() if code is not None else ''
    if not normalized:
        return 'Cesta N/D', 'Categoria N/D'
    return CATEGORY_LOOKUP.get(normalized, ('Cesta N/D', 'Categoria N/D'))
CLIENT_NAME_SUFFIX_PATTERN = re.compile(r'[\s_-]*5w1h$', re.IGNORECASE)
#obtém o país,categoria,cesta e fabricante para template e ppt
# Diccionario con las correspondencias de los números y su W respectivo
c_w = {
    ('P','1')   : '1W - Quando?',
    ('P','2')   : '2W - Por quê?',
    ('P','3-1') : '3W - O quê? Tamanhos',
    ('P','3-2') : '3W - O quê? Marcas',
    ('P','3-3') : '3W - O quê? Sabores',
    ('P','3-4') : '3W - O quê? Subsegmento',
    ('P','3-5') : '3W - O quê? Subcategoria',
    ('P','4')   : '4W - Quem? NSE (Nível Socioeconômico)',
    ('P','5-1') : '5W - Onde? Regiões',
    ('P','5-2') : '5W - Onde? Canais',
    ('P','5-3') : '5W - Onde? Canal moderno',
    ('P','5-4') : '5W - Onde? Canal tradicional',
    ('P','6')   : 'Players',
    ('P','6-1') : 'Players - Preço indexado',
    ('P','7-R') : '7W - Distribuição por regiões',
    ('P','7-NSE'): '7W - Distribuição por NSE',
    ('P','8')   : 'Intervalos de confiança',

    ('E','1')   : '1W - ¿Cuándo?',
    ('E','2')   : '2W - ¿Por qué?',
    ('E','3-1') : '3W - ¿Qué? Tamaños',
    ('E','3-2') : '3W - ¿Qué? Marcas',
    ('E','3-3') : '3W - ¿Qué? Sabores',
    ('E','3-4') : '3W - ¿Qué? Subsegmento',
    ('E','3-5') : '3W - ¿Qué? Subcategoría',
    ('E','4')   : '4W - ¿Quiénes? NSE (Nivel socioeconómico)',
    ('E','5-1') : '5W - ¿Dónde? Regiones',
    ('E','5-2') : '5W - ¿Dónde? Canales',
    ('E','5-3') : '5W - ¿Dónde? Canal moderno',
    ('E','5-4') : '5W - ¿Dónde? Canal tradicional',
    ('E','6')   : 'Players',
    ('E','6-1') : 'Players - Precio indexado',
    ('E','7-R') : '7W - Distribución por regiones',
    ('E','7-NSE'): '7W - Distribución por NSE',
    ('E','8')   : 'Intervalos de confianza'
}

# Etiquetas de los slides
labels = {
    ('P','Data')   : 'Data',
    ('P','MAT')    : 'Avaliação em ano móvel acumulado',
    ('P','Var MAT'): 'Variação em ano móvel',
    ('P','comp')   : 'Concorrência no mercado de: ',
    ('P','dist')   : 'Distribuição',

    ('E','Data')   : 'Fecha',
    ('E','MAT')    : 'Evaluación en año móvil acumulado',
    ('E','Var MAT'): 'Variación en año móvil',
    ('E','comp')   : 'Competencia en el mercado de: ',
    ('E','dist')   : 'Distribución'
}

class SeriesConfig(NamedTuple):
    data: pd.DataFrame
    raw_tipo: str
    display_tipo: str
    pipeline: int
COLOR_TOKEN_STOPWORDS = {
    'gr', 'gr.', 'grupo', 'grup', 'compras', 'compra', 'ventas', 'venta',
    'canal', 'canales', 'channel', 'channels', 'marca', 'marcas', 'brand',
    'brands'
}
COLOR_SUFFIXES = ('.c', '.v', '_c', '_v', '-c', '-v')
COLOR_TOKEN_SPLIT_PATTERN = re.compile(r'[^0-9a-z]+')
COLOR_UNIT_SUFFIX_PATTERN = re.compile(
    r'(?:[._\-\s]+(?:k|m|g|kg|grs?|gr|l|lt|lts|ml))+$',
    re.IGNORECASE
)
def normalize_color_text(value) -> str:
    """Normaliza texto para comparar etiquetas de color."""
    if value is None:
        return ''
    if not isinstance(value, str):
        value = str(value)
    normalized = unicodedata.normalize('NFKD', value)
    normalized = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
    return normalized.lower().strip()
def strip_color_suffixes(text: str) -> str:
    """Elimina sufijos de compras/ventas y unidades antes de buscar color."""
    base = text.strip()
    for suffix in COLOR_SUFFIXES:
        if base.endswith(suffix):
            base = base[:-len(suffix)]
            break
    # Elimina sufijos de unidad (p. ej. "_K", "-ML" o " K") para reutilizar el color de marca.
    base = COLOR_UNIT_SUFFIX_PATTERN.sub('', base)
    return base.strip()
def generate_color_lookup_keys(label: str) -> list[str]:
    """Genera claves canónicas seguras para reusar colores sin depender del texto literal."""
    signature = build_semantic_label_signature(label)
    if not signature.tokens:
        normalized = normalize_color_text(label)
        if not normalized:
            return []
        base = strip_color_suffixes(normalized)
        return [base] if base else []
    seen = set()
    keys: list[str] = []
    def _push(candidate: str):
        candidate = candidate.strip()
        if candidate and candidate not in seen:
            seen.add(candidate)
            keys.append(candidate)
    _push(signature.canonical)
    _push(signature.compact)
    if len(signature.tokens) > 1:
        _push('|'.join(signature.tokens))
        _push('|'.join(sorted(signature.tokens)))
    return keys

def _lookup_semantic_mapping_value(label, lookup_dict: dict[str, str], allow_family: bool = True) -> Optional[str]:
    if not lookup_dict or label is None:
        return None
    raw_label = str(label).strip()
    if not raw_label or raw_label.lower() == 'total':
        return None
    direct_color = lookup_dict.get(raw_label)
    if direct_color:
        return direct_color
    best_match = _best_semantic_reference_match(
        raw_label,
        [candidate for candidate in lookup_dict.keys() if candidate and str(candidate).strip().lower() != 'total'],
        allow_family=allow_family,
    )
    if best_match is None:
        return None
    return lookup_dict.get(best_match.reference_label)

def build_color_lookup_dict(color_mapping: dict[str, str]) -> dict[str, str]:
    """Construye un diccionario de busqueda de colores."""
    if not color_mapping:
        return {}
    lookup: dict[str, str] = {}
    for original_key, color_value in color_mapping.items():
        if color_value is None:
            continue
        normalized_key = str(original_key).strip()
        if not normalized_key or normalized_key.lower() == 'total':
            continue
        if normalized_key not in lookup:
            lookup[normalized_key] = color_value
    return lookup

def lookup_color_for_label(label, lookup_dict: dict[str, str]) -> Optional[str]:
    """Busca el color asociado a una etiqueta normalizada."""
    raw_key = str(label).strip() if label is not None else ''
    if not raw_key or raw_key.lower() == 'total':
        return None
    if not lookup_dict:
        return None
    direct_color = lookup_dict.get(raw_key)
    if direct_color:
        return direct_color
    for lookup_key in generate_color_lookup_keys(raw_key):
        color_value = lookup_dict.get(lookup_key)
        if color_value:
            return color_value
    return _lookup_semantic_mapping_value(raw_key, lookup_dict, allow_family=True)

def register_color_lookup(label, color_value: str, lookup_dict: dict[str, str], overwrite: bool = False) -> None:
    """Registra una etiqueta y su color en el lookup."""
    if label is None or not color_value:
        return
    raw_key = str(label).strip()
    if not raw_key or raw_key.lower() == 'total':
        return
    if overwrite or raw_key not in lookup_dict:
        lookup_dict[raw_key] = color_value
    for lookup_key in generate_color_lookup_keys(raw_key):
        if not lookup_key:
            continue
        if overwrite or lookup_key not in lookup_dict:
            lookup_dict[lookup_key] = color_value

def hex_to_rgb_color(color_value: str) -> Optional[RGBColor]:
    """Convierte un color HEX a RGBColor."""
    try:
        rgb_float = mcolors.to_rgb(color_value)
    except (ValueError, TypeError):
        return None
    rgb_int = tuple(int(round(val * 255)) for val in rgb_float)
    return RGBColor(*rgb_int)

def hex_to_ansi_color(color_value: str) -> Optional[str]:
    """
    Convierte un color HEX a un codigo ANSI de 24 bits para la terminal.
    """
    try:
        r, g, b = (int(round(val * 255)) for val in mcolors.to_rgb(color_value))
    except (ValueError, TypeError):
        return None
    return f"\033[38;2;{r};{g};{b}m"

def assign_brand_palette_color(label: str, lookup_dict: dict[str, str], palette_sequence: Optional[list[str]] = None) -> Optional[str]:
    """Asigna un color de paleta a una marca si no existe en el lookup."""
    if not label:
        return None
    color_value = lookup_color_for_label(label, lookup_dict)
    if color_value:
        return color_value
    palette = palette_sequence if palette_sequence else TREND_COLOR_SEQUENCE
    if not palette:
        return None
    used_colors = {value for value in lookup_dict.values() if value}
    for candidate in palette:
        if candidate not in used_colors:
            register_color_lookup(label, candidate, lookup_dict, overwrite=False)
            return candidate
    fallback = palette[len(used_colors) % len(palette)]
    register_color_lookup(label, fallback, lookup_dict, overwrite=False)
    return fallback

def set_title_with_brand_color(
    text_frame,
    prefix_text: str,
    brand_text: Optional[str],
    suffix_text: str,
    font_size_inches: float,
    brand_lookup: dict[str, str],
    palette_sequence: Optional[list[str]] = None
) -> None:
    """Escribe un titulo con el segmento de marca coloreado."""
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.text = ''
    base_size = Inches(font_size_inches)
    def _add_run(text: str, color_hex: Optional[str] = None):
        if not text:
            return
        run = paragraph.add_run()
        run.text = text
        font = run.font
        font.bold = True
        font.size = base_size
        if color_hex:
            rgb_color = hex_to_rgb_color(color_hex)
            if rgb_color:
                font.color.rgb = rgb_color
    _add_run(prefix_text)
    brand_segment = (brand_text or '').strip()
    if brand_segment:
        brand_color = assign_brand_palette_color(brand_segment, brand_lookup, palette_sequence)
        _add_run(brand_segment, brand_color)
    _add_run(suffix_text)

BRAND_TITLE_COLOR_LOOKUP: dict[str, str] = {}
# Lookup independiente para categorías (segmento 8).
CATEGORY_TITLE_COLOR_LOOKUP: dict[str, str] = {}

def colorize_brand_for_terminal(
    label: str,
    brand_lookup: Optional[dict[str, str]] = None,
    palette_sequence: Optional[list[str]] = None
) -> str:
    """
    Devuelve el nombre de la marca coloreado con su color asignado en la paleta.
    """
    if not label:
        return ''
    lookup = brand_lookup if brand_lookup is not None else BRAND_TITLE_COLOR_LOOKUP
    brand_color = assign_brand_palette_color(label, lookup, palette_sequence)
    ansi_color = hex_to_ansi_color(brand_color) if brand_color else None
    if ansi_color:
        return colorize(label, ansi_color)
    return label

def normalize_section_title(value: Optional[str]) -> str:
    """Limpia una etiqueta para usarla como nombre de seccion en PowerPoint."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        raw = strip_color_suffixes(raw)
    except Exception:
        pass
    raw = raw.replace('.', ' ').replace('_', ' ')
    raw = re.sub(r'\s*-\s*share\s*100%\s*apilado\s*$', '', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*-\s*share\s*100%\s*stacked\s*$', '', raw, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", raw).strip()

def extract_sheet_group_label(sheet_name: str, category_name: str) -> str:
    """Devuelve la etiqueta principal (marca/categoria) asociada a una hoja."""
    if not sheet_name:
        return ""
    sheet_clean = str(sheet_name).strip()
    if not sheet_clean:
        return ""
    distribution_match = DISTRIBUTION_SHEET_PATTERN.match(sheet_clean)
    first_char = sheet_clean[0]
    parsed_sheet = parse_standard_sheet_name(sheet_clean)
    brand_label = ''
    if distribution_match:
        brand_label = distribution_match.group(1).replace('.', ' ').strip()
    elif sheet_clean.startswith('6-1') or sheet_clean.startswith('6_1'):
        brand_label = category_name
    elif first_char == '6':
        brand_label = category_name
    elif parsed_sheet:
        brand_label = parsed_sheet['label']
    return normalize_section_title(brand_label)

def get_cover_section_title(lang: str) -> str:
    """Nombre de la seccion para la portada."""
    return 'Capa' if lang == 'P' else 'Portada'

def register_section_slides(
    section_slides: OrderedDict[str, list[int]],
    section_title: Optional[str],
    start_idx: int,
    end_idx: int,
) -> None:
    """Registra los indices de slides producidos para una seccion."""
    if end_idx <= start_idx:
        return
    title = normalize_section_title(section_title)
    if not title:
        return
    bucket = section_slides.setdefault(title, [])
    seen = set(bucket)
    for slide_idx in range(start_idx, end_idx):
        if slide_idx not in seen:
            bucket.append(slide_idx)
            seen.add(slide_idx)

def apply_powerpoint_sections(pptx_path: Path, section_slides: OrderedDict[str, list[int]]) -> None:
    """Inyecta la lista de secciones en presentation.xml del .pptx final."""
    if not pptx_path or not section_slides:
        return
    ET.register_namespace('a', PPTX_DRAWING_NS)
    ET.register_namespace('r', PPTX_REL_NS)
    ET.register_namespace('p', PPTX_PRESENTATION_NS)
    ET.register_namespace('p14', PPTX_SECTION_NS)
    with zipfile.ZipFile(pptx_path, 'r') as src_zip:
        presentation_xml = src_zip.read('ppt/presentation.xml')
    root = ET.fromstring(presentation_xml)
    sld_id_list = root.find(f'{{{PPTX_PRESENTATION_NS}}}sldIdLst')
    if sld_id_list is None:
        return
    slide_ids = [
        int(node.attrib['id'])
        for node in sld_id_list.findall(f'{{{PPTX_PRESENTATION_NS}}}sldId')
        if node.attrib.get('id')
    ]
    valid_sections: list[tuple[str, list[int]]] = []
    for title, slide_indexes in section_slides.items():
        slide_id_values: list[int] = []
        seen_slide_ids: set[int] = set()
        for slide_idx in slide_indexes:
            if 0 <= slide_idx < len(slide_ids):
                slide_id = slide_ids[slide_idx]
                if slide_id not in seen_slide_ids:
                    slide_id_values.append(slide_id)
                    seen_slide_ids.add(slide_id)
        if slide_id_values:
            valid_sections.append((title, slide_id_values))
    if not valid_sections:
        return
    ext_lst = root.find(f'{{{PPTX_PRESENTATION_NS}}}extLst')
    if ext_lst is None:
        ext_lst = ET.SubElement(root, f'{{{PPTX_PRESENTATION_NS}}}extLst')
    for ext in list(ext_lst.findall(f'{{{PPTX_PRESENTATION_NS}}}ext')):
        if ext.attrib.get('uri') == PPTX_SECTION_EXT_URI:
            ext_lst.remove(ext)
    sections_ext = ET.SubElement(
        ext_lst,
        f'{{{PPTX_PRESENTATION_NS}}}ext',
        {'uri': PPTX_SECTION_EXT_URI},
    )
    section_lst = ET.SubElement(sections_ext, f'{{{PPTX_SECTION_NS}}}sectionLst')
    for title, slide_id_values in valid_sections:
        section = ET.SubElement(
            section_lst,
            f'{{{PPTX_SECTION_NS}}}section',
            {
                'name': title,
                'id': '{' + str(uuid.uuid4()).upper() + '}',
            },
        )
        section_slide_list = ET.SubElement(section, f'{{{PPTX_SECTION_NS}}}sldIdLst')
        for slide_id in slide_id_values:
            ET.SubElement(
                section_slide_list,
                f'{{{PPTX_SECTION_NS}}}sldId',
                {'id': str(slide_id)},
            )
    updated_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    tmp_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx', dir=str(pptx_path.parent)) as tmp_file:
            tmp_path = Path(tmp_file.name)
        with zipfile.ZipFile(pptx_path, 'r') as read_zip, zipfile.ZipFile(tmp_path, 'w') as write_zip:
            for entry in read_zip.infolist():
                payload = updated_xml if entry.filename == 'ppt/presentation.xml' else read_zip.read(entry.filename)
                write_zip.writestr(entry, payload)
        os.replace(tmp_path, pptx_path)
    finally:
        if tmp_path is not None and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass

def _extract_top_slide_text_from_xml(slide_xml: bytes) -> str:
    """Lee el titulo del slide directamente desde su XML."""
    slide_root = ET.fromstring(slide_xml)
    text_candidates: list[tuple[int, int, str]] = []
    for shape in slide_root.findall(f'.//{{{PPTX_PRESENTATION_NS}}}sp'):
        texts = [node.text or '' for node in shape.findall(f'.//{{{PPTX_DRAWING_NS}}}t')]
        combined_text = ''.join(texts).strip()
        if not combined_text or normalize_section_title(combined_text).lower() == 'comentario':
            continue
        off = shape.find(f'./{{{PPTX_PRESENTATION_NS}}}spPr/{{{PPTX_DRAWING_NS}}}xfrm/{{{PPTX_DRAWING_NS}}}off')
        top = int(off.attrib.get('y', '0')) if off is not None else 0
        left = int(off.attrib.get('x', '0')) if off is not None else 0
        text_candidates.append((top, left, combined_text))
    if not text_candidates:
        return ""
    text_candidates.sort(key=lambda item: (item[0], item[1]))
    return text_candidates[0][2]

def _derive_section_title_from_slide_text(title_text: str, previous_title: str, lang: str) -> str:
    """Infere el nombre de seccion a partir del titulo visible del slide."""
    if title_text:
        if '|' in title_text:
            _, right_part = title_text.split('|', 1)
            parsed_title = normalize_section_title(right_part)
            if parsed_title:
                return parsed_title
        parsed_title = normalize_section_title(title_text)
        if parsed_title:
            return parsed_title
    if previous_title:
        return previous_title
    return 'Geral' if lang == 'P' else 'General'

def build_section_slides_from_presentation(pptx_path: Path, lang: str) -> OrderedDict[str, list[int]]:
    """Reconstruye el mapa de secciones leyendo los titulos del .pptx ya generado."""
    section_slides: OrderedDict[str, list[int]] = OrderedDict()
    with zipfile.ZipFile(pptx_path, 'r') as pptx_zip:
        presentation_root = ET.fromstring(pptx_zip.read('ppt/presentation.xml'))
        rels_root = ET.fromstring(pptx_zip.read('ppt/_rels/presentation.xml.rels'))
        slide_rel_targets: dict[str, str] = {}
        for rel in rels_root.findall(f'{{{PPTX_PACKAGE_REL_NS}}}Relationship'):
            rel_id = rel.attrib.get('Id')
            target = rel.attrib.get('Target')
            if rel_id and target:
                slide_rel_targets[rel_id] = target.lstrip('/')
        slide_order: list[str] = []
        sld_id_list = presentation_root.find(f'{{{PPTX_PRESENTATION_NS}}}sldIdLst')
        if sld_id_list is not None:
            for slide_node in sld_id_list.findall(f'{{{PPTX_PRESENTATION_NS}}}sldId'):
                rel_id = slide_node.attrib.get(f'{{{PPTX_REL_NS}}}id')
                target = slide_rel_targets.get(rel_id or '')
                if target:
                    slide_order.append(f'ppt/{target}')
        if not slide_order:
            return section_slides
        cover_title = get_cover_section_title(lang)
        register_section_slides(section_slides, cover_title, 0, 1)
        previous_title = cover_title
        for slide_idx, slide_name in enumerate(slide_order[1:], start=1):
            title_text = _extract_top_slide_text_from_xml(pptx_zip.read(slide_name))
            section_title = _derive_section_title_from_slide_text(title_text, previous_title, lang)
            register_section_slides(section_slides, section_title, slide_idx, slide_idx + 1)
            previous_title = section_title or previous_title
    return section_slides

def build_terminal_label(sheet_name: str, lang: str, category_name: str) -> Optional[str]:
    """
    Devuelve un label corto y legible (ej: '3W Marcas - X').
    """
    if not sheet_name:
        return None
    sheet_clean = sheet_name.strip()
    if not sheet_clean:
        return None
    distribution_match = DISTRIBUTION_SHEET_PATTERN.match(sheet_clean)
    first_char = sheet_clean[0]
    parsed_sheet = parse_standard_sheet_name(sheet_clean)
    variant = parsed_sheet['variant'] if parsed_sheet else ''
    brand_label = extract_sheet_group_label(sheet_clean, category_name)
    step_label: Optional[str] = None
    cw_key = None
    # Distribución (segmento 7) via patrón 7_*_* (ej.: R, NSE, WIFEAGE, FAMILY_SIZE)
    if distribution_match:
        dist_kind_raw = distribution_match.group(2)
        dist_kind = str(dist_kind_raw).upper().strip()
        pretty_dist = dist_kind.replace('_', ' ').replace('-', ' ')
        pretty_dist = ' '.join(word.capitalize() for word in pretty_dist.split())
        if dist_kind == 'R':
            cw_key = '7-R'
            step_label = '7W Distribución Regiones'
        elif dist_kind == 'NSE':
            cw_key = '7-NSE'
            step_label = '7W Distribución NSE'
        else:
            step_label = f'7W Distribución {pretty_dist or dist_kind}'
    elif sheet_clean.startswith('6-1') or sheet_clean.startswith('6_1'):
        cw_key = '6-1'
    elif first_char == '6':
        cw_key = '6'
    elif first_char == '8':
        cw_key = '8'
    elif first_char == '5':
        cw_key = f"5-{variant}" if variant in {'1', '2', '3', '4'} else '5'
    elif first_char == '4':
        cw_key = '4'
    elif first_char == '3':
        cw_key = f"3-{variant}" if variant in {'1', '2', '3', '4', '5'} else '3'
    elif first_char in {'2', '1'}:
        cw_key = first_char
    if cw_key:
        step_label = c_w.get((lang, cw_key), step_label if step_label else cw_key)
    if not step_label:
        return None
    # Para mantener el estilo corto en terminal, eliminamos el " - " del texto base
    if ' - ' in step_label:
        code, desc = step_label.split(' - ', 1)
        step_label = f"{code} {desc}"
    brand_terminal = colorize_brand_for_terminal(brand_label) if brand_label else ''
    detail = f" - {brand_terminal}" if brand_terminal else ''
    return f"{step_label}{detail}"

def build_terminal_progress_message(
    sheet_name: str,
    lang: str,
    category_name: str,
    suffix: str = "",
) -> Optional[str]:
    """
    Mensaje de progreso claro (ej: 'Generando 3W Marcas - X (hoja 3_Marca_P)').
    """
    label = build_terminal_label(sheet_name, lang, category_name)
    if not label:
        return None
    suffix_text = f" {suffix}" if suffix else ""
    return f"Generando {label} (hoja {sheet_name}){suffix_text}"

def build_terminal_done_message(sheet_name: str, lang: str, category_name: str) -> Optional[str]:
    """Devuelve el mensaje final de progreso en terminal."""
    label = build_terminal_label(sheet_name, lang, category_name)
    if not label:
        return None
    return f"Listo {label}"

def _extract_pipeline(col_name: str) -> int:
    """Extrae el numero de pipeline desde el encabezado."""
    if not isinstance(col_name, str):
        return 0
    text = str(col_name).strip()
    if not text:
        return 0
    max_pipeline = 12
    normalized = _normalize_header_text(text)
    explicit_match = re.search(r"(?:pipeline|pipe|p)\s*[:=_-]?\s*(\d{1,2})\b", normalized, re.IGNORECASE)
    if explicit_match:
        value = int(explicit_match.group(1))
        return value if 0 <= value <= max_pipeline else 0
    for token in re.split(r"[_\s\-]+", text):
        if token.isdigit():
            value = int(token)
            if 0 <= value <= max_pipeline:
                return value
    return 0
def _is_separator_column(col) -> bool:
    """Detecta columnas separadoras vacias o 'Unnamed'."""
    if col is None:
        return True
    if not isinstance(col, str):
        return False
    stripped = col.strip()
    return not stripped or stripped.lower().startswith('unnamed')
def _normalize_header_text(value: str) -> str:
    normalized = unicodedata.normalize('NFKD', str(value))
    normalized = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
    return normalized.lower().strip()
def _find_compras_header_row(df: pd.DataFrame, include_table: bool = True) -> Optional[int]:
    """Busca la fila con encabezado de Compras (y opcionalmente table)."""
    if df is None or df.empty:
        return None
    compras_idx = None
    table_idx = None
    for idx, row in df.iterrows():
        for value in row:
            if isinstance(value, str) and 'compras' in _normalize_header_text(value):
                compras_idx = idx
                break
            if include_table and isinstance(value, str) and 'table' in _normalize_header_text(value):
                if table_idx is None:
                    table_idx = idx
        if compras_idx is not None:
            break
    if compras_idx is not None:
        return compras_idx
    return table_idx


def _find_table_anchor_row(df: pd.DataFrame) -> Optional[int]:
    if df is None or df.empty or df.shape[1] == 0:
        return None
    first_col = df.iloc[:, 0]
    for idx, value in first_col.items():
        if isinstance(value, str) and 'table' in _normalize_header_text(value):
            return idx
    return None


def _find_first_date_row(
    df: pd.DataFrame,
    start_idx: int,
    date_format: str
) -> Optional[int]:
    if df is None or df.empty or df.shape[1] == 0:
        return None
    first_col = df.iloc[start_idx:, 0]
    parsed, non_empty = _parse_flexible_date_series(first_col)
    valid = parsed.notna() & non_empty
    if valid.any():
        return valid[valid].index[0]
    return None


def _find_header_row_near_date(df: pd.DataFrame, date_row_idx: int) -> Optional[int]:
    if df is None or df.empty:
        return None
    for idx in range(date_row_idx, max(date_row_idx - 6, -1), -1):
        row = df.iloc[idx]
        for value in row:
            if isinstance(value, str):
                text = _normalize_header_text(value)
                if 'compras' in text or 'ventas' in text:
                    return idx
    if date_row_idx > 0:
        return date_row_idx - 1
    return None


def _row_has_mat_headers(row: pd.Series) -> bool:
    for value in row:
        if isinstance(value, str) and 'mat' in _normalize_header_text(value):
            return True
    return False
# Segmentos 3-6: parsing de hojas con Compras/Ventas y series.
def parse_sheet_with_compras_header(
    excel_file: pd.ExcelFile,
    sheet_name: str,
    include_table: bool = True
) -> pd.DataFrame:
    """
    Carga la hoja buscando la fila con 'Compras' como encabezado para tolerar filas vacias o
    textos previos; si no se encuentra, usa el encabezado por defecto.
    """
    raw_df = excel_file.parse(sheet_name, header=None)
    header_row = _find_compras_header_row(raw_df, include_table=include_table)
    if header_row is None:
        return excel_file.parse(sheet_name)
    return excel_file.parse(sheet_name, header=header_row)


def parse_sheet_with_date_fallback(
    excel_file: pd.ExcelFile,
    sheet_name: str,
    date_format: str = '%b-%y  '
) -> pd.DataFrame:
    """
    Intenta leer la hoja con encabezado Compras/Ventas.
    Si falla, busca la primera fecha debajo de un ancla "table" y reintenta.
    """
    raw_df = excel_file.parse(sheet_name, header=None)
    try:
        base_df = parse_sheet_with_compras_header(excel_file, sheet_name, include_table=False)
        return ensure_date_column(base_df, sheet_name=sheet_name, date_format=date_format)
    except ValueError:
        table_idx = _find_table_anchor_row(raw_df)
        start_idx = table_idx + 1 if table_idx is not None else 0
        date_row_idx = _find_first_date_row(raw_df, start_idx, date_format)
        if date_row_idx is None:
            if table_idx is not None and _row_has_mat_headers(raw_df.iloc[table_idx]):
                raise ValueError(
                    f"La hoja {sheet_name} contiene un encabezado 'table' con periodos MAT, "
                    "pero no se encontraron fechas en la primera columna."
                )
            raise
        header_idx = _find_header_row_near_date(raw_df, date_row_idx)
        if header_idx is None:
            raise ValueError(
                f"No se encontro encabezado de Compras/Ventas alrededor de las fechas en la hoja {sheet_name}."
            )
        df = excel_file.parse(sheet_name, header=header_idx)
        return ensure_date_column(df, sheet_name=sheet_name, date_format=date_format)


def _parse_flexible_date_series(series: pd.Series) -> tuple[pd.Series, pd.Series]:
    """Convierte una serie de fechas usando el helper global soportando formatos mixtos."""
    if series is None:
        return pd.Series(dtype="datetime64[ns]"), pd.Series(dtype=bool)
    if pd.api.types.is_datetime64_any_dtype(series):
        parsed = pd.to_datetime(series, errors='coerce')
        non_empty = series.map(lambda value: not pd.isna(value)).astype(bool)
        return parsed, non_empty

    parsed_values = []
    non_empty = []
    for raw_value in series.tolist():
        value = raw_value
        if isinstance(value, str):
            value = value.strip()
        is_empty = value is None or value == "" or (isinstance(value, float) and math.isnan(value))
        if is_empty:
            parsed_values.append(pd.NaT)
            non_empty.append(False)
            continue
        parsed_value = _try_parse_reference_date(value)
        parsed_values.append(pd.Timestamp(parsed_value) if parsed_value is not None else pd.NaT)
        non_empty.append(True)

    return (
        pd.Series(parsed_values, index=series.index, dtype="datetime64[ns]"),
        pd.Series(non_empty, index=series.index, dtype=bool),
    )


def ensure_date_column(df: pd.DataFrame, sheet_name: Optional[str] = None, date_format: str = '%b-%y  ') -> pd.DataFrame:
    """
    Valida y convierte la primera columna a fecha (formato MMM-YY).
    Lanza un ValueError si no hay fechas legibles para evitar graficos con datos corruptos.
    """
    if df is None or df.empty or df.shape[1] == 0:
        return df
    first_col = df.iloc[:, 0]
    context = f" en la hoja {sheet_name}" if sheet_name else ""
    parsed, non_empty = _parse_flexible_date_series(first_col)
    if df.shape[1] > 1:
        payload_mask = df.iloc[:, 1:].apply(
            lambda column: column.map(
                lambda value: not (
                    value is None
                    or (isinstance(value, str) and not value.strip())
                    or pd.isna(value)
                )
            )
        ).any(axis=1)
    else:
        payload_mask = pd.Series(False, index=df.index, dtype=bool)
    invalid_mask = parsed.isna() & (non_empty | payload_mask)
    if invalid_mask.any() or parsed.notna().sum() == 0:
        invalid_rows = [str(int(idx) + 2) if isinstance(idx, (int, np.integer)) else str(idx) for idx in invalid_mask[invalid_mask].index[:5]]
        row_detail = f" Filas detectadas: {', '.join(invalid_rows)}." if invalid_rows else ""
        raise ValueError(
            f"No se pudieron leer fechas validas en la primera columna{context}. "
            f"Use el formato MMM-YY (ej.: Ene-24).{row_detail}"
        )
    empty_rows = parsed.isna() & ~payload_mask
    df_copy = df.copy()
    if empty_rows.any():
        df_copy = df_copy.loc[~empty_rows].copy()
        parsed = parsed.loc[~empty_rows]
    df_copy.iloc[:, 0] = parsed
    df_copy.reset_index(drop=True, inplace=True)
    return df_copy
# Segmento 7: distribuciones (R/NSE/otros cortes).
DISTRIBUTION_SHEET_PATTERN = re.compile(r'^7[-_](.+?)[-_](.+)$', re.IGNORECASE)
def _normalize_simple(text: str) -> str:
    """Normaliza texto a minusculas y sin tildes."""
    normalized = unicodedata.normalize('NFKD', str(text))
    normalized = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
    return normalized.lower().strip()
def _is_distribution_total_row(cat_value, row_values: pd.Series) -> bool:
    """
    Detecta filas de totales/resumen que no deben graficarse en distribuciones.
    Conserva filas como "Total Guatemala" o similares siempre que los valores
    no sean uniformes.
    """
    if not isinstance(cat_value, str):
        return False
    normalized = _normalize_simple(cat_value)
    if not normalized or 'total' not in normalized:
        return False
    numeric_vals = pd.to_numeric(row_values, errors='coerce').dropna()
    if numeric_vals.empty:
        return True
    unique_vals = pd.Series(np.round(numeric_vals, 3)).unique()
    if len(unique_vals) == 1:
        return True
    if np.allclose(numeric_vals, 100, atol=0.5):
        return True
    return False
def _find_distribution_header_indices(raw_df: pd.DataFrame) -> tuple[Optional[int], Optional[int]]:
    """
    Busca una fila de encabezados cuando no hay ancla "table/Compras".
    Heuristica: una fila con al menos dos celdas no vacías (col 1+) y la fila
    siguiente con al menos un número, asumiendo que es cabecera + datos.
    """
    if raw_df is None or raw_df.empty:
        return None, None
    max_scan = min(raw_df.shape[0] - 1, 12)
    for idx in range(max_scan):
        row_rest = raw_df.iloc[idx, 1:]
        non_empty = row_rest.dropna().astype(str).str.strip()
        if (non_empty != '').sum() < 2:
            continue
        next_row = raw_df.iloc[idx + 1, 1:] if idx + 1 < raw_df.shape[0] else None
        if next_row is None:
            continue
        numeric_count = pd.to_numeric(next_row, errors='coerce').count()
        if numeric_count >= 1:
            return idx, idx + 1
    return None, None
def parse_distribution_sheet(excel_file: pd.ExcelFile, sheet_name: str, client_name: str) -> tuple[pd.DataFrame, str]:
    """
    Extrae la distribucion (R/NSE) desde las hojas 7_*_*.
    Mantiene todas las series con datos y usa los encabezados de la fila 2 como nombres de marca.
    """
    raw_df = excel_file.parse(sheet_name, header=None)
    if raw_df.shape[0] < 3 or raw_df.shape[1] < 2:
        return pd.DataFrame(columns=['Categoria']), ''
    # Buscar la fila "table" (en la primera columna) para alinear la lectura, similar al segmento 2
    table_anchor_idx = None
    def _cell_has_anchor(value) -> bool:
        if value is None:
            return False
        text = str(value)
        if not text:
            return False
        normalized = unicodedata.normalize('NFKD', text)
        normalized = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
        normalized = normalized.lower()
        return 'table' in normalized or 'compras' in normalized
    for idx in range(raw_df.shape[0]):
        first_cell = raw_df.iat[idx, 0] if raw_df.shape[1] > 0 else None
        if _cell_has_anchor(first_cell):
            table_anchor_idx = idx
            break
    header_idx = None
    data_start_idx = None
    if table_anchor_idx is not None and table_anchor_idx + 1 < raw_df.shape[0]:
        header_idx = table_anchor_idx
        data_start_idx = table_anchor_idx + 1
    if header_idx is None:
        header_idx, data_start_idx = _find_distribution_header_indices(raw_df)
    if header_idx is None:
        header_idx, data_start_idx = 1, 2
    if data_start_idx is None or data_start_idx >= raw_df.shape[0]:
        return pd.DataFrame(columns=['Categoria']), ''
    header_row = raw_df.iloc[header_idx, 1:]
    data_block = raw_df.iloc[data_start_idx:, :]
    categories = data_block.iloc[:, 0].astype(str).str.strip()
    df = pd.DataFrame({'Categoria': categories})
    series_headers: list[str] = []
    for idx in range(1, data_block.shape[1]):
        serie = pd.to_numeric(data_block.iloc[:, idx], errors='coerce')
        if serie.dropna().empty:
            continue
        header_value = header_row.iloc[idx - 1] if (idx - 1) < len(header_row) else ''
        header_text = str(header_value).strip() if pd.notna(header_value) else ''
        if not header_text:
            fallback_name = None
            for candidate in data_block.iloc[:, idx]:
                if isinstance(candidate, str) and candidate.strip():
                    fallback_name = candidate.strip()
                    break
            header_text = fallback_name if fallback_name else f"Serie {idx}"
        df[header_text] = serie
        series_headers.append(header_text)
    df = df[df['Categoria'].str.len() > 0]
    df = df[~df['Categoria'].str.lower().eq('nan')]
    value_columns = [col for col in df.columns if col != 'Categoria']
    if not value_columns:
        return pd.DataFrame(columns=['Categoria']), ''
    df = df.dropna(subset=value_columns, how='all')
    df = df[~df.apply(lambda row: _is_distribution_total_row(row['Categoria'], row[value_columns]), axis=1)]
    def _is_total_label(label: str) -> bool:
        return isinstance(label, str) and 'total' in label.lower()
    total_columns = [col for col in value_columns if _is_total_label(col)]
    non_total_columns = [col for col in value_columns if col not in total_columns]
    ordered_columns = total_columns + non_total_columns
    df = df[['Categoria'] + ordered_columns]
    ordered_headers = [col for col in ordered_columns if col]
    return df, ', '.join(ordered_headers)
def plot_distribution_chart(
    df: pd.DataFrame,
    c_fig: int,
    color_lookup: Optional[dict[str, str]] = None,
    title: Optional[str] = None,  # se conserva la firma para no romper llamadas previas
    adaptive_label_size: bool = True,
    target_width_in: Optional[float] = None,
    max_height_in: Optional[float] = None
) -> tuple[io.BytesIO, tuple[float, float]]:
    """
    Crea un grafico de barras agrupadas con etiquetas numericas para distribuciones R/NSE.
    Permite escalar el tamano de las etiquetas de barra cuando hay muchas categorias/series.
    Devuelve el stream PNG y el tamano de la figura en pulgadas para insertarlo sin reescalar.
    """
    if df.empty or df.shape[1] <= 1:
        fig, ax = plt.subplots(num=c_fig)
        ax.axis('off')
        fig_size = fig.get_size_inches()
        return figure_to_stream(fig, bbox_inches=None, pad_inches=0), (float(fig_size[0]), float(fig_size[1]))
    categories = df['Categoria'].tolist()
    series_names = [col for col in df.columns if col != 'Categoria']
    filtered_series: list[tuple[str, pd.Series]] = []
    for serie_name in series_names:
        values_raw = pd.to_numeric(df[serie_name], errors='coerce')
        if values_raw.dropna().empty:
            continue
        values = values_raw.fillna(0)
        if values.replace(0, np.nan).dropna().empty:
            continue
        filtered_series.append((serie_name, values))
    if not categories or not filtered_series:
        fig, ax = plt.subplots(num=c_fig, figsize=(8, 4), dpi=DEFAULT_EXPORT_DPI)
        ax.axis('off')
        fig_size = fig.get_size_inches()
        return figure_to_stream(fig, bbox_inches=None, pad_inches=0), (float(fig_size[0]), float(fig_size[1]))
    # Descarta categorias cuyo valor sea ~100 para todas las series (evita barras "planas" de referencia)
    keep_mask = []
    tol_100 = 0.1  # tolerancia para considerar un valor como 100
    for cat_idx in range(len(categories)):
        vals_at_cat = []
        for _, serie_vals in filtered_series:
            try:
                val = float(serie_vals.iloc[cat_idx])
            except Exception:
                continue
            if np.isfinite(val):
                vals_at_cat.append(val)
        if vals_at_cat:
            max_v = max(vals_at_cat)
            min_v = min(vals_at_cat)
            if abs(max_v - 100.0) <= tol_100 and abs(min_v - 100.0) <= tol_100:
                keep_mask.append(False)
                continue
        else:
            keep_mask.append(False)
            continue
        keep_mask.append(True)
    if not any(keep_mask):
        fig, ax = plt.subplots(num=c_fig, figsize=(8, 4), dpi=DEFAULT_EXPORT_DPI)
        ax.axis('off')
        fig_size = fig.get_size_inches()
        return figure_to_stream(fig, bbox_inches=None, pad_inches=0), (float(fig_size[0]), float(fig_size[1]))
    keep_indices = [idx for idx, keep in enumerate(keep_mask) if keep]
    categories = [categories[idx] for idx in keep_indices]
    filtered_series = [
        (name, pd.Series([vals.iloc[idx] for idx in keep_indices])) for name, vals in filtered_series
    ]
    n_series = len(filtered_series)
    x = np.arange(len(categories))
    width = 0.8 / max(n_series, 1)
    base_fig_width = max(10.0, min(16.0, 1.3 * len(categories)))
    base_fig_height = max(6.5, min(12.0, 2.5 + 0.9 * len(categories)))
    fig_width = target_width_in if target_width_in else base_fig_width
    if fig_width <= 0:
        fig_width = base_fig_width
    fig_height = base_fig_height
    if max_height_in is not None and max_height_in > 0:
        fig_height = min(fig_height, max_height_in)
    default_label_font_size = 9
    label_font_size = default_label_font_size
    if adaptive_label_size:
        total_bars = max(len(categories) * n_series, 1)
        density = total_bars / max(fig_width, 1.0)  # barras por pulgada aproximada
        if density > 5.5:
            label_font_size = 6
        elif density > 4.0:
            label_font_size = 7
        elif density > 3.0:
            label_font_size = 8
        elif density < 1.5:
            label_font_size = min(11, default_label_font_size + 1)
    label_font_size = max(6, min(11, int(round(label_font_size))))
    fig, ax = plt.subplots(num=c_fig, figsize=(fig_width, fig_height), dpi=DEFAULT_EXPORT_DPI)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#F9FAFB')
    ax.set_axisbelow(True)
    palette_lookup = {}
    def _is_total_label(label: str) -> bool:
        return isinstance(label, str) and 'total' in label.lower()
    def _is_reference_total_series(label: str, serie_values: pd.Series, tol: float = 0.5) -> bool:
        """
        Considera serie de referencia solo si contiene 'total' y sus valores ~100.
        Evita pintar de negro series con 'total' en el nombre que no son totales reales.
        """
        if not _is_total_label(label):
            return False
        numeric_vals = pd.to_numeric(serie_values, errors='coerce')
        numeric_vals = numeric_vals[np.isfinite(numeric_vals)]
        if numeric_vals.empty:
            return False
        return np.allclose(numeric_vals, 100.0, atol=tol)
    if TREND_COLOR_PALETTE:
        palette_values = [
            TREND_COLOR_PALETTE[key]
            for key in sorted(
                TREND_COLOR_PALETTE.keys(),
                key=lambda k: int(k.split('_')[1]) if '_' in k and k.split('_')[1].isdigit() else k
            )
        ]
    else:
        palette_values = TREND_COLOR_SEQUENCE
    if not palette_values:
        palette_values = [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]
    palette_index = 0
    bars_by_series = []
    legend_handles = []
    legend_labels = []
    max_val = 0.0
    for idx, (serie_name, values) in enumerate(filtered_series):
        is_total_series = _is_reference_total_series(serie_name, values)
        if idx == 0 or is_total_series:
            color_val = REFERENCE_SERIES_COLOR
        else:
            color_val = palette_values[palette_index % len(palette_values)]
            palette_index += 1
        register_color_lookup(serie_name, color_val, palette_lookup, overwrite=True)
        offset = (idx - (n_series - 1) / 2) * width
        bars = ax.bar(
            x + offset,
            values,
            width,
            label=serie_name,
            color=color_val,
            edgecolor='none',
            alpha=0.9
        )
        bars_by_series.append((bars, values))
        legend_handles.append(bars)
        legend_labels.append(serie_name)
        serie_max = values.max() if not values.empty else 0
        if np.isfinite(serie_max):
            max_val = max(max_val, float(serie_max))
    ax.set_xticks(x)
    tick_font_size = 10
    ax.set_xticklabels(categories, rotation=30, ha='right', va='top', rotation_mode='anchor', fontweight='normal', fontsize=tick_font_size)
    ax.tick_params(axis='x', pad=2)
    ax.tick_params(axis='y', labelsize=tick_font_size, pad=2)
    ax.set_ylabel('%', fontweight='bold', fontsize=11, labelpad=10)
    ax.grid(axis='y', linestyle='--', alpha=0.45, color='#D9D9D9')
    ax.margins(x=0.02)
    for spine in ax.spines.values():
        spine.set_visible(False)
    limit = max_val * 1.15 if max_val > 0 else None
    if limit:
        ax.set_ylim(0, limit)
    ax.set_xlim(-0.5, len(categories) - 0.5)
    ref_idx = 0 if filtered_series else None
    highlight_map: dict[tuple[int, int], str] = {}
    if ref_idx is not None and len(filtered_series) > 1:
        for cat_idx in range(len(categories)):
            try:
                ref_val = float(bars_by_series[ref_idx][1].iloc[cat_idx])
            except Exception:
                ref_val = None
            if ref_val is None or not np.isfinite(ref_val):
                continue
            diffs = []
            for s_idx, (_, vals) in enumerate(bars_by_series):
                if s_idx == ref_idx:
                    continue
                try:
                    val = float(vals.iloc[cat_idx])
                except Exception:
                    continue
                if not np.isfinite(val):
                    continue
                diffs.append((s_idx, val - ref_val))
            if not diffs:
                continue
            max_entry = max(diffs, key=lambda t: t[1])
            min_entry = min(diffs, key=lambda t: t[1])
            if max_entry[1] > 0:
                highlight_map[(max_entry[0], cat_idx)] = BAR_LABEL_COLOR_POS_ALT
            if min_entry[1] < 0:
                highlight_map[(min_entry[0], cat_idx)] = BAR_LABEL_COLOR_NEG_ALT
    label_padding = max(max_val * 0.01, 0.05 * (label_font_size / default_label_font_size))
    max_label_y = 0.0
    for serie_idx, (bars, values) in enumerate(bars_by_series):
        for cat_idx, (bar, value) in enumerate(zip(bars, values)):
            label_val = f"{value:.1f}" if abs(value) >= 10 else f"{value:.2f}"
            color_val = highlight_map.get((serie_idx, cat_idx), TABLE_TEXT_PRIMARY)
            proposed_y = bar.get_height() + label_padding
            if proposed_y > max_label_y:
                max_label_y = proposed_y
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                proposed_y,
                label_val,
                ha='center',
                va='bottom',
                rotation=0,
                fontsize=label_font_size,
                fontweight='normal',
                color=color_val,
                clip_on=False
            )
    if max_label_y > 0:
        ylim_bottom, ylim_top = ax.get_ylim()
        extra_top = max(max_val * 0.03, label_padding)
        target_top = max(ylim_top, max_label_y + extra_top)
        ax.set_ylim(ylim_bottom, target_top)
    legend = ax.legend(
        legend_handles,
        legend_labels,
        frameon=True,
        loc='upper center',
        bbox_to_anchor=(0.5, 1.02),
        ncol=max(1, len(legend_labels)),
        borderaxespad=0.2,
        columnspacing=1.0,
        handlelength=1.4,
        handletextpad=0.4
    )
    if legend:
        legend.get_frame().set_facecolor('#FFFFFF')
        legend.get_frame().set_edgecolor('#CCCCCC')
        legend.get_frame().set_alpha(0.95)
    # Ajuste dinamico de margenes: conserva la ubicacion de la leyenda arriba
    # y expande el area util del eje para evitar que el contenido quede comprimido.
    fig.canvas.draw()
    fig_height_in = fig.get_size_inches()[1] or 1.0
    legend_height_frac = 0.0
    if legend:
        renderer = fig.canvas.get_renderer()
        if renderer:
            bbox = legend.get_window_extent(renderer=renderer)
            legend_height_frac = bbox.height / (fig_height_in * fig.dpi)
    # Margen inferior estimado segun el alto de las etiquetas del eje X (rotadas)
    xtick_labels = ax.get_xticklabels()
    bottom_margin = 0.1
    if xtick_labels:
        max_label_pts = max(label.get_size() for label in xtick_labels if label.get_text())
        estimated_bottom = (max_label_pts * 1.6) / (fig_height_in * 72.0)
        bottom_margin = max(0.06, min(0.16, estimated_bottom))
    # Margen superior reducido dinamicamente para dejar espacio a la leyenda
    top_margin = 0.93 - legend_height_frac * 1.2
    top_margin = max(0.84, min(0.95, top_margin))
    # Aseguramos consistencia con los bar charts de otras secciones (3-6)
    fig.subplots_adjust(left=0.06, right=0.98, top=top_margin, bottom=bottom_margin)
    try:
        # Afina el layout sin mover la leyenda
        fig.tight_layout(rect=[0.05, bottom_margin + 0.01, 0.97, top_margin - 0.01])
    except Exception:
        pass
    return figure_to_stream(fig, bbox_inches=None, pad_inches=0), (fig_width, fig_height)

def split_compras_ventas(df: pd.DataFrame, sheet_name: Optional[str] = None) -> tuple[list[pd.DataFrame], int]:
    """Divide la hoja en bloques de compras/ventas y detecta pipeline."""
    warning_context = f" en la hoja {sheet_name}" if sheet_name else ""
    normalized_columns = [
        str(col).strip().lower() if isinstance(col, str) else ''
        for col in df.columns
    ]
    has_compras_header = any('compra' in col for col in normalized_columns)
    has_ventas_header = any('venta' in col for col in normalized_columns)
    if not has_compras_header:
        print_colored(
            f"No se encontro el encabezado de \"Compras\"{warning_context}. Se usaran los datos disponibles.",
            COLOR_RED
        )
    ventas_idx = None
    pipeline_header = None
    for idx, col in enumerate(df.columns):
        if isinstance(col, str) and 'ventas' in col.lower():
            ventas_idx = idx
            pipeline_header = col
            break
    fallback_separator_idx = None
    for idx in range(len(df.columns) - 1):
        if _is_separator_column(df.columns[idx]):
            next_idx = idx + 1
            if next_idx < len(df.columns) and not _is_separator_column(df.columns[next_idx]):
                fallback_separator_idx = idx
                break
    suffix_hint = any(
        isinstance(col, str) and col.endswith('.1')
        for col in df.columns
    )
    fallback_used = False
    if ventas_idx is None and fallback_separator_idx is not None:
        candidate_idx = fallback_separator_idx + 1
        if candidate_idx < len(df.columns):
            ventas_idx = candidate_idx
            pipeline_header = df.columns[candidate_idx]
            fallback_used = True
    pipeline = _extract_pipeline(pipeline_header) if pipeline_header is not None else 0
    has_pipeline_digits = isinstance(pipeline_header, str) and any(ch.isdigit() for ch in pipeline_header)
    if ventas_idx is None:
        if (fallback_separator_idx is not None or suffix_hint) and not has_ventas_header:
            print_colored(
                f"No se encontro el encabezado de \"Ventas\"{warning_context}. No se pudo dividir la hoja.",
                COLOR_RED
            )
        return [df], 0
    if fallback_used and not has_ventas_header:
        print_colored(
            f"No se encontro el encabezado de \"Ventas\"{warning_context}. Se detecto el bloque por la columna en blanco y se graficara con pipeline = 0.",
            COLOR_RED
        )
        pipeline = 0
    if pipeline == 0 and pipeline_header is not None and not has_pipeline_digits:
        print_colored(
            f"No se definio un pipeline numerico en la columna '{pipeline_header}'{warning_context}. Se usara pipeline = 0.",
            COLOR_RED
        )
        pipeline = 0
    separator_idx = ventas_idx - 1
    has_separator = separator_idx >= 0 and _is_separator_column(df.columns[separator_idx])
    compras_end = separator_idx if has_separator else ventas_idx
    if compras_end <= 0:
        compras_end = ventas_idx
    compras = df.iloc[:, :compras_end].copy()
    ventas = df.iloc[:, ventas_idx:].copy()
    if not compras.empty:
        compras = compras.loc[:, ~pd.Series(compras.columns).apply(_is_separator_column).to_numpy()]
        if len(compras.columns) > 1:
            compras.columns = [compras.columns[0]] + [f"{col}.C" for col in compras.columns[1:]]
    if not ventas.empty:
        ventas = ventas.loc[:, ~pd.Series(ventas.columns).apply(_is_separator_column).to_numpy()]
        ventas_cols = list(ventas.columns)
        if ventas_cols:
            first = ventas_cols[0]
            base_name = first[:-2] if isinstance(first, str) and len(first) > 2 and first[-2] == '_' else first
            ventas_cols[0] = base_name
            normalized_ventas_cols = []
            for col in ventas_cols[1:]:
                # pandas agrega sufijos numericos (p. ej. ``.1``) solo cuando el
                # encabezado de Ventas coincide exactamente con uno de Compras.
                # No podemos depender de ese detalle: espacios finales o series
                # exclusivas de Ventas dejan el encabezado sin sufijo y el grafico
                # no las reconoce como sell-in para aplicarles el pipeline.
                base_col = re.sub(r'\.\d+$', '', str(col).strip()).strip()
                for origin_suffix in ('.v', '_v', '-v'):
                    if base_col.lower().endswith(origin_suffix):
                        base_col = base_col[:-len(origin_suffix)].strip()
                        break
                normalized_ventas_cols.append(f"{base_col}.V")
            ventas_cols[1:] = normalized_ventas_cols
            ventas.columns = ventas_cols
    # En hojas con pipeline es comun que el ultimo bloque de Ventas no traiga fecha ni valores.
    # Conservamos la fila para que el desplazamiento del pipeline siga alineado y heredamos el
    # periodo equivalente de Compras cuando este sea legible.
    aligned_periods = 0
    if not compras.empty and not ventas.empty and compras.shape[1] > 0 and ventas.shape[1] > 0:
        compras_periods, _ = _parse_flexible_date_series(compras.iloc[:, 0])
        aligned_length = min(len(compras), len(ventas))
        for row_pos in range(aligned_length):
            ventas_value = ventas.iat[row_pos, 0]
            ventas_missing = ventas_value is None or pd.isna(ventas_value) or (
                isinstance(ventas_value, str) and not ventas_value.strip()
            )
            compras_period = compras_periods.iloc[row_pos]
            if ventas_missing and pd.notna(compras_period):
                ventas.iat[row_pos, 0] = compras_period
                aligned_periods += 1
    if aligned_periods:
        print_colored(
            (
                f"Se completaron {aligned_periods} periodo(s) faltante(s) del bloque de Ventas"
                f"{warning_context} usando la fecha alineada de Compras; se conserva pipeline = {pipeline}."
            ),
            COLOR_YELLOW,
        )
    if ventas.shape[1] <= 1 or ventas.iloc[:, 1:].isna().all().all():
        return [compras if not compras.empty else df], 0
    df_list: list[pd.DataFrame] = []
    if not compras.empty:
        df_list.append(compras)
    df_list.append(ventas)
    return df_list, pipeline


def _detect_ventas_pipeline(columns: pd.Index) -> int:
    for col in columns:
        if isinstance(col, str) and 'ventas' in col.lower():
            return _extract_pipeline(col)
    return 0
def prepare_series_configs(df_list, lang, p_ventas, sheet_name: Optional[str] = None):
    """Prepara configuraciones de series para graficar."""
    configs = []
    for original_df in df_list:
        if original_df is None or original_df.empty or original_df.shape[1] == 0:
            context = f" en la hoja {sheet_name}" if sheet_name else ""
            print_colored(f"No se encontraron datos graficables{context}. Se omite esta seccion.", COLOR_RED)
            continue
        df_local = ensure_date_column(original_df.copy(), sheet_name=sheet_name)
        if df_local.shape[1] == 0:
            context = f" en la hoja {sheet_name}" if sheet_name else ""
            print_colored(f"No hay columnas disponibles{context}. Se omite esta seccion.", COLOR_RED)
            continue
        raw_tipo = str(df_local.columns[0])
        df_local.rename(columns={df_local.columns[0]: labels[(lang,'Data')]}, inplace=True)
        is_compras = 'compras' in raw_tipo.lower()
        if not is_compras and len(df_local.columns) > 1:
            is_compras = '.c' in str(df_local.columns[1]).lower()
        display_tipo = 'Compras' if is_compras else 'Ventas'
        pipeline = 0 if is_compras else p_ventas
        configs.append(SeriesConfig(df_local, raw_tipo, display_tipo, pipeline))
    return configs

def origin_display_order(display_tipo) -> int:
    """Orden visual de paneles: Ventas a la izquierda, Compras a la derecha."""
    return {'ventas': 0, 'compras': 1}.get(str(display_tipo).strip().lower(), 2)


def reorder_series_configs_by_compras(series_configs: list[SeriesConfig]) -> list[SeriesConfig]:
    """Alinea el orden de las series no-Compras usando Compras como referencia semántica."""
    if not series_configs:
        return series_configs
    compras_config = next(
        (
            cfg for cfg in series_configs
            if str(getattr(cfg, "display_tipo", "")).strip().lower() == "compras"
            and getattr(cfg, "data", None) is not None
            and getattr(cfg.data, "shape", (0, 0))[1] > 1
        ),
        None
    )
    if compras_config is None:
        return series_configs
    reference_labels = [
        str(col)
        for col in compras_config.data.columns[1:]
        if not _is_total_like_series(col)
    ]
    if not reference_labels:
        return series_configs
    reordered_configs: list[SeriesConfig] = []
    for cfg in series_configs:
        df_local = cfg.data
        if df_local is None or df_local.empty or df_local.shape[1] <= 1:
            reordered_configs.append(cfg)
            continue
        value_cols = [col for col in df_local.columns[1:] if not _is_total_like_series(col)]
        total_cols = [col for col in df_local.columns[1:] if _is_total_like_series(col)]
        reordered_value_labels = _reorder_labels_against_reference(
            [str(col) for col in value_cols],
            reference_labels,
            allow_family=True,
        )
        if reordered_value_labels == [str(col) for col in value_cols]:
            reordered_configs.append(cfg)
            continue
        raw_to_original: dict[str, list] = {}
        for col in value_cols:
            raw_to_original.setdefault(str(col), []).append(col)
        reordered_cols = [df_local.columns[0]]
        for reordered_label in reordered_value_labels:
            original_candidates = raw_to_original.get(reordered_label) or []
            if not original_candidates:
                continue
            reordered_cols.append(original_candidates.pop(0))
        for remaining_candidates in raw_to_original.values():
            reordered_cols.extend(remaining_candidates)
        reordered_cols.extend(total_cols)
        reordered_df = df_local.loc[:, reordered_cols].copy()
        reordered_df.attrs = dict(getattr(df_local, "attrs", {}))
        reordered_configs.append(
            SeriesConfig(reordered_df, cfg.raw_tipo, cfg.display_tipo, cfg.pipeline)
        )
    return reordered_configs
SIZE_GROUP_MAX_ELEMENTS = 10
SIZE_GROUP_WINDOW_MONTHS = 12
SIZE_GROUP_SUFFIXES = ('.C', '.V', '_C', '_V', '-C', '-V', '.c', '.v', '_c', '_v', '-c', '-v')
def is_segment3_size_sheet(sheet_name: str) -> bool:
    """Indica si la hoja corresponde al segmento 3-1 (Tamaños)."""
    parsed_sheet = parse_standard_sheet_name(str(sheet_name).strip())
    return bool(parsed_sheet) and parsed_sheet['segment'] == '3' and parsed_sheet['variant'] == '1'
def _is_total_like_series(label) -> bool:
    """Detecta encabezados de total ignorando sufijos de compras/ventas."""
    normalized = strip_color_suffixes(normalize_color_text(label))
    return normalized == 'total'
def _count_size_candidates_in_range(df: pd.DataFrame, start_col: int, end_col: int) -> int:
    """Cuenta columnas graficables de tamaños dentro de un rango de columnas."""
    if df is None or df.empty:
        return 0
    if start_col >= end_col:
        return 0
    upper_bound = min(end_col, df.shape[1])
    count = 0
    for idx in range(max(1, start_col), upper_bound):
        column_name = df.columns[idx]
        if _is_separator_column(column_name):
            continue
        if _is_total_like_series(column_name):
            continue
        numeric_series = pd.to_numeric(df.iloc[:, idx], errors='coerce')
        if numeric_series.dropna().empty:
            continue
        if numeric_series.replace(0, np.nan).dropna().empty:
            continue
        count += 1
    return count
def _find_ventas_header_index(columns: pd.Index) -> Optional[int]:
    """Devuelve el indice de la columna de Ventas si existe."""
    for idx, col in enumerate(columns):
        if idx == 0:
            continue
        if isinstance(col, str) and 'ventas' in col.lower():
            return idx
    return None
def quick_scan_segment3_size_max_elements(excel_file: pd.ExcelFile, sheet_names: list[str]) -> int:
    """Escaneo rapido para encontrar el maximo de elementos graficables en tamaños."""
    max_elements = 0
    for sheet_name in sheet_names:
        try:
            parsed_df = parse_sheet_with_compras_header(excel_file, sheet_name)
        except Exception:
            continue
        if parsed_df is None or parsed_df.empty or parsed_df.shape[1] <= 1:
            continue
        ventas_idx = _find_ventas_header_index(parsed_df.columns)
        if ventas_idx is None:
            sheet_count = _count_size_candidates_in_range(parsed_df, 1, parsed_df.shape[1])
        else:
            compras_count = _count_size_candidates_in_range(parsed_df, 1, ventas_idx)
            ventas_count = _count_size_candidates_in_range(parsed_df, ventas_idx + 1, parsed_df.shape[1])
            sheet_count = max(compras_count, ventas_count)
        if sheet_count > max_elements:
            max_elements = sheet_count
    return max_elements
def _parse_threshold_percent(raw_value: str) -> Optional[float]:
    """Parsea una entrada de porcentaje (admite coma y simbolo %)."""
    if raw_value is None:
        return None
    cleaned = str(raw_value).strip().replace('%', '').replace(',', '.')
    if not cleaned:
        return None
    try:
        numeric = float(cleaned)
    except (TypeError, ValueError):
        return None
    if not np.isfinite(numeric):
        return None
    return numeric
def format_threshold_percent(value: float) -> str:
    """Formatea porcentaje evitando decimales innecesarios."""
    rounded = round(float(value), 2)
    if float(rounded).is_integer():
        return str(int(rounded))
    return f"{rounded:.2f}".rstrip('0').rstrip('.')
def prompt_size_grouping_threshold(max_elements: int) -> Optional[float]:
    """Solicita umbral para agrupar elementos de tamaños cuando hay demasiadas series."""
    print_colored(
        f"Se detectaron {max_elements} elementos en Tamaños (segmento 3-1).",
        COLOR_YELLOW
    )
    prompt = (
        colorize(
            'Se detectaron demasiados elementos en el tamaño. '
            'Desea agrupar por umbral minimo? ',
            COLOR_QUESTION
        )
        + colorize('Ingrese porcentaje (ej: 2). Enter para no agrupar: ', COLOR_YELLOW)
    )
    while True:
        choice = input(prompt).strip()
        if not choice:
            return None
        threshold = _parse_threshold_percent(choice)
        if threshold is None or threshold <= 0:
            print_colored('Entrada invalida. Ingrese un porcentaje numerico mayor a 0.', COLOR_RED)
            continue
        print_colored(
            f"Se agruparan categorias por debajo de {format_threshold_percent(threshold)}%.",
            COLOR_YELLOW
        )
        return threshold
def _detect_series_suffix(df: pd.DataFrame, column_indexes: list[int]) -> str:
    """Detecta sufijo de serie (.C/.V) para mantener etiquetas consistentes."""
    for idx in column_indexes:
        if idx < 0 or idx >= df.shape[1]:
            continue
        text = str(df.columns[idx]).strip()
        for suffix in SIZE_GROUP_SUFFIXES:
            if text.endswith(suffix):
                return suffix
    return ''
def _build_others_label(df: pd.DataFrame, threshold_pct: float, column_indexes: list[int]) -> str:
    """Construye etiqueta de agrupacion por umbral y evita colisiones."""
    threshold_text = format_threshold_percent(threshold_pct)
    base_label = f"Agrupado menores a {threshold_text}%"
    suffix = _detect_series_suffix(df, column_indexes)
    candidate = base_label
    existing = {str(col).strip() for col in df.columns if str(col).strip()}
    if candidate not in existing:
        return candidate
    suffix_lower = str(suffix).lower()
    origin_label = ''
    if suffix_lower.endswith('c'):
        origin_label = 'Compras'
    elif suffix_lower.endswith('v'):
        origin_label = 'Ventas'
    if origin_label:
        candidate_with_origin = f"{base_label} ({origin_label})"
        if candidate_with_origin not in existing:
            return candidate_with_origin
    counter = 2
    while f"{candidate}_{counter}" in existing:
        counter += 1
    return f"{candidate}_{counter}"
def group_size_elements_by_threshold(
    df: pd.DataFrame,
    pipeline_offset: int,
    threshold_pct: float
) -> tuple[pd.DataFrame, bool]:
    """
    Agrupa columnas por debajo del umbral de share MAT actual.
    Cuando el usuario eligio un umbral, se aplica a cada bloque de Compras y
    Ventas aunque uno de ellos tenga diez series o menos.
    """
    if df is None or df.empty or df.shape[1] <= 2:
        return df, False
    try:
        pipeline = int(pipeline_offset) if pipeline_offset is not None else 0
    except (TypeError, ValueError):
        pipeline = 0
    pipeline = max(0, pipeline)
    end_idx = len(df) - pipeline if pipeline > 0 else len(df)
    start_idx = end_idx - SIZE_GROUP_WINDOW_MONTHS
    if start_idx < 0 or end_idx <= start_idx:
        return df, False
    numeric_by_index: dict[int, pd.Series] = {}
    candidate_indexes: list[int] = []
    total_indexes: list[int] = []
    for idx in range(1, df.shape[1]):
        column_name = df.columns[idx]
        if _is_separator_column(column_name):
            continue
        numeric_series = pd.to_numeric(df.iloc[:, idx], errors='coerce')
        if numeric_series.dropna().empty:
            continue
        numeric_by_index[idx] = numeric_series
        if _is_total_like_series(column_name):
            total_indexes.append(idx)
            continue
        if numeric_series.replace(0, np.nan).dropna().empty:
            continue
        candidate_indexes.append(idx)
    column_window_sums: dict[int, float] = {}
    total_window_sum = 0.0
    for idx in candidate_indexes:
        window_sum = float(numeric_by_index[idx].iloc[start_idx:end_idx].sum(skipna=True))
        if not np.isfinite(window_sum):
            window_sum = 0.0
        window_sum = max(window_sum, 0.0)
        column_window_sums[idx] = window_sum
        total_window_sum += window_sum
    if total_window_sum <= 0:
        return df, False
    keep_indexes: list[int] = []
    grouped_indexes: list[int] = []
    for idx in candidate_indexes:
        share_pct = (column_window_sums[idx] / total_window_sum) * 100.0
        if share_pct >= threshold_pct:
            keep_indexes.append(idx)
        else:
            grouped_indexes.append(idx)
    if not grouped_indexes:
        return df, False
    others_label = _build_others_label(df, threshold_pct, candidate_indexes)
    grouped_columns_labels = [
        str(df.columns[idx]).strip()
        for idx in grouped_indexes
        if str(df.columns[idx]).strip()
    ]
    grouped_sum_series = pd.DataFrame(
        {idx: numeric_by_index[idx] for idx in grouped_indexes}
    ).sum(axis=1, min_count=1)
    grouped_df = pd.DataFrame({df.columns[0]: df.iloc[:, 0]})
    keep_set = set(keep_indexes)
    grouped_set = set(grouped_indexes)
    total_set = set(total_indexes)
    inserted_group = False
    for idx in range(1, df.shape[1]):
        column_name = df.columns[idx]
        if _is_separator_column(column_name):
            continue
        if idx in keep_set:
            grouped_df[column_name] = numeric_by_index[idx]
            continue
        if idx in grouped_set:
            if not inserted_group:
                grouped_df[others_label] = grouped_sum_series
                inserted_group = True
            continue
        if idx in total_set:
            grouped_df[column_name] = numeric_by_index[idx]
            continue
        grouped_df[column_name] = df.iloc[:, idx]
    if not inserted_group:
        grouped_df[others_label] = grouped_sum_series
    grouped_df.attrs["size_grouped_columns"] = grouped_columns_labels
    grouped_df.attrs["size_grouped_label"] = others_label
    return grouped_df, True
def format_origin_label(display_tipo: Optional[str], lang: str) -> str:
    """
    Normaliza la etiqueta de origen de datos (Compras/Ventas) para textos como 'Corte a:'.
    """
    if not display_tipo:
        return ""
    base = str(display_tipo).strip()
    if not base:
        return ""
    normalized = base.lower()
    if normalized.startswith('comp') or 'compra' in normalized or normalized.endswith('.c'):
        return 'Compras'
    if normalized.startswith('vent') or 'venta' in normalized or normalized.endswith('.v') or 'sell-out' in normalized:
        return 'Vendas' if lang == 'P' else 'Ventas'
    if 'sell-in' in normalized:
        return 'Compras'
    return base


def _trend_block_key_from_display_tipo(display_tipo: Optional[str]) -> str:
    normalized = str(display_tipo or "").strip().lower()
    if 'venta' in normalized or 'venda' in normalized:
        return 'ventas'
    return 'compras'


def _build_excel_trend_series_specs(series_config: SeriesConfig, color_mapping: Optional[dict[str, str]], dashed: bool) -> list[dict]:
    specs: list[dict] = []
    grouped_label = ''
    grouped_columns: list[str] = []
    if hasattr(series_config.data, 'attrs'):
        grouped_label = str(series_config.data.attrs.get("size_grouped_label", "")).strip()
        grouped_columns = [
            str(label).strip()
            for label in series_config.data.attrs.get("size_grouped_columns", [])
            if str(label).strip()
        ]
    for col in list(series_config.data.columns)[1:]:
        label = str(col).strip()
        if not label:
            continue
        numeric_series = pd.to_numeric(series_config.data[col], errors='coerce')
        if numeric_series.replace(0, np.nan).dropna().empty:
            continue
        source_labels = grouped_columns if grouped_label and label == grouped_label and grouped_columns else [label]
        color_value = (color_mapping or {}).get(label)
        specs.append(
            {
                "label": label,
                "source_labels": source_labels,
                "block_key": _trend_block_key_from_display_tipo(series_config.display_tipo),
                "color": color_value,
                "dash": dashed,
                "shift_from_start": _trend_block_key_from_display_tipo(series_config.display_tipo) == 'ventas',
            }
        )
    return specs
def extract_players_base_key(sheet_name: str) -> Optional[str]:
    """Deriva la clave base para hojas de jugadores (segmento 6)."""
    if not isinstance(sheet_name, str):
        return None
    if not sheet_name or not sheet_name.startswith('6'):
        return None
    parts = [part for part in sheet_name.split('_') if part]
    if not parts:
        return None
    if parts[-1].isdigit():
        if len(parts) > 1:
            return '_'.join(parts[:-1])
        return parts[0]
    return '_'.join(parts)
# Precio indexado: identifica hojas y arma la tabla indexada con contexto Players.
def is_price_index_sheet(sheet_name: str) -> bool:
    """Indica si la hoja corresponde a precio indexado."""
    if not isinstance(sheet_name, str):
        return False
    if not sheet_name.startswith('6'):
        return False
    parts = [part for part in sheet_name.split('_') if part]
    if len(parts) < 3:
        return False
    last_part = parts[-1]
    return last_part == '1'
def prepare_price_index_dataframe(series_config: SeriesConfig, top10_columns: list[str]) -> tuple[pd.DataFrame, list[str], pd.DataFrame]:
    """Prepara dataframes para grafico y tabla de precio indexado."""
    df_price = series_config.data.copy()
    if df_price.empty or df_price.shape[1] <= 1:
        return pd.DataFrame(), [], pd.DataFrame()
    date_col = df_price.columns[0]
    data_cols = list(df_price.columns[1:])
    if not data_cols:
        return pd.DataFrame(), [], pd.DataFrame()
    total_col = next(
        (col for col in data_cols if 'total' in str(col).strip().lower()),
        data_cols[0]
    )
    df_price[total_col] = pd.to_numeric(df_price[total_col], errors='coerce')
    df_price = df_price[df_price[total_col].notna()]
    df_price = df_price[df_price[total_col] > 0]
    if df_price.empty:
        return pd.DataFrame(columns=[date_col]), [], pd.DataFrame(columns=[date_col])
    normalized_top10 = {str(col).strip().lower() for col in top10_columns if str(col).strip()}
    candidate_columns = []
    filtered_columns = []
    for col in data_cols:
        if col == total_col:
            continue
        numeric_series = pd.to_numeric(df_price[col], errors='coerce')
        if numeric_series.dropna().empty:
            continue
        if (numeric_series == 0).any():
            continue
        candidate_columns.append(col)
        col_lower = str(col).strip().lower()
        if not normalized_top10 or col_lower in normalized_top10:
            filtered_columns.append(col)
    if normalized_top10 and not filtered_columns:
        return pd.DataFrame(columns=[date_col]), [], pd.DataFrame(columns=[date_col])
    if not filtered_columns:
        filtered_columns = candidate_columns[:10]
    if not filtered_columns:
        return pd.DataFrame(columns=[date_col]), [], pd.DataFrame(columns=[date_col])
    total_numeric = pd.to_numeric(df_price[total_col], errors='coerce')
    ratio_df = pd.DataFrame({date_col: df_price[date_col]})
    ratio_df['Total'] = np.where(total_numeric.notna(), 100.0, np.nan)
    original_df = pd.DataFrame({date_col: df_price[date_col], 'Total': total_numeric})
    for col in filtered_columns:
        numeric_series = pd.to_numeric(df_price[col], errors='coerce')
        original_df[col] = numeric_series
        ratio_df[col] = numeric_series
    total_series = total_numeric.astype(float)
    with np.errstate(divide='ignore', invalid='ignore'):
        for col in filtered_columns:
            ratio_df[col] = ratio_df[col].astype(float) / total_series * 100.0
    ratio_df = ratio_df.replace([np.inf, -np.inf], np.nan)
    ratio_df.dropna(subset=filtered_columns, how='all', inplace=True)
    return ratio_df, ['Total'] + filtered_columns, original_df
def graf_price_index_table(
    columns: list[str],
    averages: dict[str, float],
    color_mapping: dict[str, str],
    metric_label: str,
    c_fig: int
) -> tuple[io.BytesIO, tuple[float, float]]:
    """Genera la tabla de indicadores de precio indexado."""
    column_count = max(len(columns), 1)
    fig_width = max(6.0, 1.25 * column_count)
    fig_height = 1.5
    fig, ax = plt.subplots(num=c_fig, figsize=(fig_width, fig_height), dpi=DEFAULT_EXPORT_DPI)
    ax.axis('off')
    ax.set_facecolor('white')
    header_labels = ['Indicador'] + [wrap_table_text(str(col)) for col in columns]
    display_columns = ['Indicador'] + columns
    row_values = [metric_label]
    for col in columns:
        value = averages.get(col)
        if value is None or not np.isfinite(value):
            row_values.append('-')
        else:
            row_values.append(f"{value:.2f}%")
    row_values[0] = wrap_table_text(row_values[0])
    table = ax.table(
        cellText=[row_values],
        colLabels=header_labels,
        cellLoc='center',
        loc='center'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.05, 1.4)
    default_header_color = HEADER_COLOR_PRIMARY
    header_cell = table[(0, 0)]
    header_cell.set_facecolor(default_header_color)
    header_cell.set_text_props(color='white', weight='bold')
    header_cell.set_height(header_cell.get_height() * 1.5)
    for idx, col in enumerate(columns, start=1):
        header_cell = table[(0, idx)]
        header_color = color_mapping.get(col, '#777777')
        header_cell.set_facecolor(header_color)
        header_cell.set_text_props(color='white', weight='bold')
        header_cell.set_height(header_cell.get_height() * 1.5)
    for idx in range(len(display_columns)):
        data_cell = table[(1, idx)]
        data_cell.set_height(data_cell.get_height() * 1.35)
        data_cell.set_facecolor('#FFFFFF')
        data_cell.set_edgecolor('#DDDDDD')
    fig.tight_layout(pad=0.2)
    fig_size = fig.get_size_inches()
    return figure_to_stream(fig, dpi=TABLE_EXPORT_DPI), fig_size
def build_price_index_slide(
    ppt: Presentation,
    lang: str,
    cat: str,
    apo_entries: list[dict],
    removed_headers: list[str],
    price_df: pd.DataFrame,
    price_original_df: pd.DataFrame,
    chart_pipeline: int,
    chart_share_lookup: dict,
    c_fig: int
) -> int:
    """Construye la diapositiva de precio indexado."""
    slide = ppt.slides.add_slide(ppt.slide_layouts[1])
    titulo_base = c_w.get((lang, '6-1'), 'Precio indexado')
    titulo_completo = f"{titulo_base} | {labels[(lang,'comp')]}{cat}"
    title_prefix = f"{titulo_base} | {labels[(lang,'comp')]}"
    title_box = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
    title_tf = title_box.text_frame
    competition_title_lookup: dict[str, str] = {}
    set_title_with_brand_color(
        title_tf,
        title_prefix,
        cat,
        '',
        0.33,
        competition_title_lookup,
        COMPETITION_TITLE_PALETTE
    )
    comment_box = slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
    comment_tf = comment_box.text_frame
    comment_tf.clear()
    comment_paragraph = comment_tf.paragraphs[0]
    comment_paragraph.text = "Comentario" if lang == 'P' else "Comentario"
    comment_paragraph.font.size = Inches(0.25)
    chart_share_lookup = chart_share_lookup or {}
    if not price_df.empty and price_df.shape[1] > 1:
        left_margin_chart = Inches(0.33)
        right_margin_chart = Inches(0.33)
        available_line_width = available_width(ppt, left_margin_chart, right_margin_chart)
        c_fig += 1
        chart_height = Cm(12)
        chart_colors = {}
        slide.shapes.add_picture(
            line_graf(
                price_df,
                chart_pipeline,
                titulo_completo,
                c_fig,
                1,
                width_emu=available_line_width,
                height_emu=chart_height,
                share_lookup=chart_share_lookup,
                y_axis_percent=True,
                top_value_annotations=2,
                color_collector=chart_colors,
                color_overrides={'Total': '#000000'},
                linestyle_overrides={'Total': '--'},
                show_title=False,
            ),
            left_margin_chart,
            Inches(CHART_TOP_INCH),
            width=available_line_width,
            height=chart_height
        )
        chart_colors.setdefault('Total', '#000000')
        if not price_original_df.empty and chart_colors:
            data_cols = [col for col in price_df.columns[1:] if col in chart_colors]
            if 'Total' in data_cols:
                data_cols = ['Total'] + [col for col in data_cols if col != 'Total']
            if data_cols:
                tail_df = price_original_df.iloc[-12:] if len(price_original_df) >= 12 else price_original_df
                averages_raw: dict[str, float] = {}
                for col in data_cols:
                    series = pd.to_numeric(tail_df[col], errors='coerce')
                    avg_value = series.mean() if not series.dropna().empty else np.nan
                    averages_raw[col] = avg_value
                total_key = next((col for col in data_cols if str(col).strip().lower() == 'total'), None)
                total_avg = averages_raw.get(total_key) if total_key else np.nan
                averages = {}
                if total_key and np.isfinite(total_avg) and not np.isclose(total_avg, 0.0):
                    for col, avg_value in averages_raw.items():
                        if avg_value is None or not np.isfinite(avg_value):
                            averages[col] = np.nan
                        elif col == total_key:
                            averages[col] = 100.0
                        else:
                            averages[col] = (avg_value / total_avg) * 100.0
                else:
                    for col, avg_value in averages_raw.items():
                        averages[col] = avg_value if avg_value is not None and np.isfinite(avg_value) else np.nan
                if any(np.isfinite(val) for val in averages.values()):
                    c_fig += 1
                    metric_label = "Preco indexado\nmedio 12 meses" if lang == 'P' else "Precio indexado 12m"
                    table_stream, fig_size = graf_price_index_table(
                        data_cols,
                        averages,
                        chart_colors,
                        metric_label,
                        c_fig
                    )
                    target_table_height = Cm(TABLE_TARGET_HEIGHT_CM)
                    bottom_margin = Cm(1.5)
                    slide_height_emu = int(ppt.slide_height)
                    chart_bottom_emu = int(Inches(CHART_TOP_INCH) + chart_height)
                    min_table_top = chart_bottom_emu + int(Cm(0.2))
                    table_top_emu = slide_height_emu - int(target_table_height) - int(bottom_margin)
                    if table_top_emu < min_table_top:
                        table_top_emu = min_table_top
                    pic_table = slide.shapes.add_picture(
                        table_stream,
                        left_margin_chart,
                        table_top_emu,
                        height=target_table_height
                    )
                    constrain_picture_width(pic_table, available_line_width)
    return c_fig
# Variable de control del numero de graficos
def _try_parse_reference_date(value):
    """Intenta parsear una fecha desde distintos formatos."""
    month_token_map = {
        'jan': 1, 'ene': 1,
        'feb': 2,
        'mar': 3,
        'apr': 4, 'abr': 4,
        'may': 5,
        'jun': 6,
        'jul': 7,
        'aug': 8, 'ago': 8,
        'sep': 9, 'set': 9,
        'oct': 10,
        'nov': 11,
        'dec': 12, 'dic': 12
    }
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        match = re.search(r'([A-Za-z]{3})[-/](\d{2,4})', value)
        if match:
            month_token, year_token = match.groups()
            month = month_token_map.get(month_token.lower())
            if month:
                year = int(year_token)
                if year < 100:
                    year += 2000 if year < 50 else 1900
                return dt(year, month, 1)
        # Intentar formatos cortos antes de delegar en pandas
        for fmt in ('%b-%y', '%b-%y  '):
            try:
                return dt.strptime(value, fmt)
            except ValueError:
                continue
    try:
        parsed = pd.to_datetime(value, errors='coerce')
    except Exception:
        return None
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime() if hasattr(parsed, 'to_pydatetime') else parsed


def run_report(source_path, output_root, mode="1", selected_sheets=None, size_threshold=None):
    """Original report pipeline with explicit inputs, intended for an isolated worker."""
    global land, cat, client, lang, base_dir, file, excel, ppt, W, plot
    ensure_heavy_modules_loaded()
    source_path = Path(source_path).resolve()
    base_dir, excel = source_path.parent, source_path.name
    file = pd.ExcelFile(str(source_path))
    W = [s for s in file.sheet_names if str(s).strip().lower() != "readme"]
    if selected_sheets is not None:
        if not selected_sheets or any(s not in W for s in selected_sheets):
            raise ValueError("Seleccion de hojas 5W1H invalida.")
        W = list(selected_sheets)
    if not W:
        raise ValueError("No hay hojas analiticas.")
    parts = source_path.stem.split("_", 2)
    land = lookup_country_name(int(parts[0]))
    cesta, cat = lookup_category_info(parts[1])
    client = CLIENT_NAME_SUFFIX_PATTERN.sub("", parts[2]).strip() or parts[2]
    lang = "P" if land == "Brasil" else "E"
    ppt = Presentation(str(Path(__file__).with_name("Modelo_5W1H.pptx")))
    configure_cover_slide(ppt, lang)
    BRAND_TITLE_COLOR_LOOKUP.clear()
    CATEGORY_TITLE_COLOR_LOOKUP.clear()
    c_fig=0
    plot=str(mode)
    size_grouping_threshold_pct = size_threshold
    segment3_size_sheets = [sheet_name for sheet_name in W if is_segment3_size_sheet(sheet_name)]
    if segment3_size_sheets:
        max_size_elements = quick_scan_segment3_size_max_elements(file, segment3_size_sheets)
        if max_size_elements > SIZE_GROUP_MAX_ELEMENTS:
            size_grouping_threshold_pct = size_threshold
    last_reference_source = None
    last_reference_origin = None
    last_tree_period_dt = None
    players_share_context = {}
    excel_sheet_exports = []
    #---------------------------------------------------------------------------------------------------------------------
    chart_generation_start = dt.now()
    for w in W:
        sheet_clean = str(w).strip()
        progress_label = build_terminal_label(w, lang, cat)
        if not progress_label:
            print_colored(
                (
                    f"Se omite la hoja '{w}': formato de nombre no reconocido. "
                    "Formatos esperados: 1_, 2_, 3_, 4_, 5_, 6_, 6-1_, 7_<segmento>_<corte>, 8_."
                ),
                COLOR_YELLOW
            )
            continue
        progress_message = f"Generando {progress_label} (hoja {w})"
        if progress_message and not sheet_clean.startswith('8'):
            print_loading_message(progress_message, COLOR_QUESTION)
        distribution_match = DISTRIBUTION_SHEET_PATTERN.match(w.strip())
        # Segmento 7: distribuciones (regiones/NSE/otros cortes) desde hojas 7_*_*.
        if distribution_match:
            dist_kind_raw = distribution_match.group(2)
            dist_kind = str(dist_kind_raw).upper().strip()
            category_segment = distribution_match.group(1).replace('.', ' ').strip()
            if dist_kind == 'R':
                dist_label = 'Regiones'
            elif dist_kind == 'NSE':
                dist_label = 'NSE'
            else:
                pretty_dist = dist_kind.replace('_', ' ').replace('-', ' ')
                dist_label = ' '.join(word.capitalize() for word in pretty_dist.split()) or dist_kind
            dist_df, dist_series = parse_distribution_sheet(file, w, client)
            if dist_df.empty:
                print_colored(f"No se encontraron datos para graficar en la hoja {w}.", COLOR_RED)
                continue
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            title_box = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
            title_tf = title_box.text_frame
            suffix_text = ''
            set_title_with_brand_color(
                title_tf,
                f"Distribucion {dist_label} | ",
                category_segment,
                suffix_text,
                0.35,
                BRAND_TITLE_COLOR_LOOKUP
            )
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available_width_emu = available_width(ppt, left_margin, right_margin)
            available_width_in = emu_to_inches(int(available_width_emu))
            slide_height_in = emu_to_inches(int(ppt.slide_height))
            comment_top_in = emu_to_inches(int(Inches(6.33)))
            max_height_in = min(
                slide_height_in - CHART_TOP_INCH - 0.3,
                comment_top_in - CHART_TOP_INCH - 0.2
            )
            if max_height_in <= 0:
                max_height_in = None
            c_fig += 1
            chart_stream, fig_size = plot_distribution_chart(
                dist_df,
                c_fig,
                BRAND_TITLE_COLOR_LOOKUP,
                target_width_in=available_width_in,
                max_height_in=max_height_in
            )
            fig_width_in, fig_height_in = fig_size
            target_height_in = fig_height_in
            if fig_width_in and fig_width_in > 0 and available_width_in:
                target_height_in = fig_height_in * (available_width_in / fig_width_in)
            if max_height_in is not None and target_height_in > max_height_in:
                target_height_in = max_height_in
            slide.shapes.add_picture(
                chart_stream,
                left_margin,
                Inches(CHART_TOP_INCH),
                width=available_width_emu,
                height=Inches(target_height_in)
            )
            comment_box = slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
            comment_tf = comment_box.text_frame
            comment_tf.clear()
            comment_para = comment_tf.paragraphs[0]
            comment_para.text = "Comentario"
            comment_para.font.size = Inches(0.25)
            continue
        # Segmento 8: intervalos y error muestral con bloques mensuales + agregados.
        if sheet_clean.startswith('8'):
            try:
                raw_df = file.parse(w, header=None)
            except Exception as exc:
                print_colored(f"No se pudo leer la hoja {w}: {exc}", COLOR_RED)
                continue
            monthly_blocks = ppt8_parse_monthly_blocks(raw_df)
            agg_blocks = ppt8_parse_agg_blocks(raw_df)
            if not monthly_blocks or not agg_blocks:
                print_colored(f"Segmento 8: no se detectaron bloques validos en la hoja {w}.", COLOR_RED)
                continue
            agg_blocks, agg_dropped = ppt8_filter_agg_blocks(monthly_blocks, agg_blocks)
            if agg_dropped:
                unique_labels: list[str] = []
                for label in agg_dropped:
                    cleaned = str(label).strip()
                    if cleaned and cleaned not in unique_labels:
                        unique_labels.append(cleaned)
                preview = ", ".join(unique_labels[:3])
                extra = len(unique_labels) - 3
                extra_note = f", +{extra} mas" if extra > 0 else ""
                detail = f": {preview}{extra_note}" if preview else ""
                print_colored(
                    f"Segmento 8: se omitieron {len(agg_dropped)} bloque(s) agregados sin datos{detail}.",
                    COLOR_YELLOW
                )
            if not agg_blocks:
                print_colored(f"Segmento 8: no se detectaron bloques agregados validos en la hoja {w}.", COLOR_RED)
                continue
            if len(monthly_blocks) != len(agg_blocks):
                print_colored(
                    f"Segmento 8: se usara el minimo comun de bloques en {w} (mensuales={len(monthly_blocks)}, agregados={len(agg_blocks)}).",
                    COLOR_YELLOW
                )
            volume_metric_used = ppt8_detect_volume_metric(agg_blocks) or "R_VOL1"
            progress_message = build_terminal_progress_message(
                w,
                lang,
                cat,
                suffix=f"- [{volume_metric_used}]",
            )
            if progress_message:
                print_loading_message(progress_message, COLOR_QUESTION)
            ppt_rows, ppt_errors = ppt8_compute_rows(monthly_blocks, agg_blocks)
            for err in ppt_errors:
                print_colored(err, COLOR_YELLOW)
            if not ppt_rows:
                print_colored(f"Segmento 8: no hay filas calculables en la hoja {w}.", COLOR_RED)
                continue
            left_margin = Inches(0.33)
            available_width_emu = available_width(ppt, left_margin, Inches(0.33))
            confidence_layout = confidence_design.slide_layout(emu_to_inches(int(ppt.slide_height)), CHART_TOP_INCH)
            max_height_in = confidence_layout['table_height']
            capacity = confidence_layout['capacity']
            ppt8_brand_color_lookup: dict[str, str] = {}
            pages = [ppt_rows[i:i + capacity] for i in range(0, len(ppt_rows), capacity)]
            for page_index, page_rows in enumerate(pages, 1):
                slide = ppt.slides.add_slide(ppt.slide_layouts[1])
                title_box = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
                title_prefix = c_w.get((lang, '8'), 'Intervalos') + ' | '
                brand_label = sheet_clean[2:].replace('_', ' ').strip() or cat
                suffix = f" ({page_index}/{len(pages)})" if len(pages) > 1 else ''
                set_title_with_brand_color(title_box.text_frame, title_prefix, brand_label, suffix,
                                          0.35, CATEGORY_TITLE_COLOR_LOOKUP, CATEGORY_TITLE_PALETTE)
                add_ppt8_table_to_slide(slide, page_rows, ppt8_brand_color_lookup, client,
                                       left=left_margin, top=Inches(CHART_TOP_INCH),
                                       width_emu=available_width_emu, max_height_in=max_height_in)
                comment_box = slide.shapes.add_textbox(Inches(11.07), Inches(confidence_layout['comment_top']), Inches(2), Inches(0.4))
                comment_box.name = 'Comentario · intervalos de confianza'
                comment_para = comment_box.text_frame.paragraphs[0]
                comment_para.text = "Comentario"
                comment_para.font.size = Inches(0.25)
                confidence_design.add_buyer_reference(slide, left_margin, Inches(confidence_layout['reference_top']), available_width_emu)
            last_reference_source = raw_df
            last_reference_origin = 'Segmento 8'
            continue
        # Precio indexado: usa el contexto de Players (segmento 6) para construir la tabla.
        if is_price_index_sheet(w):
            players_base_key = extract_players_base_key(w)
            context = players_share_context.get(players_base_key)
            if context is None:
                print_colored(f"No se encontro contexto de Players para la hoja {w}. Se omite Precio indexado.", COLOR_YELLOW)
                continue
            try:
                df_start = parse_sheet_with_compras_header(file, w)
                df_list, p_ventas = split_compras_ventas(df_start, sheet_name=w)
                series_configs = prepare_series_configs(df_list, lang, p_ventas, sheet_name=w)
            except Exception as exc:
                msg = str(exc)
                color = COLOR_YELLOW if "encabezado 'table' con periodos MAT" in msg else COLOR_RED
                print_colored(f"Se omite la hoja {w}: {msg}", color)
                continue
            price_series = None
            for cfg in series_configs:
                if cfg.display_tipo.lower() == 'compras':
                    price_series = cfg
                    break
            if price_series is None:
                price_series = series_configs[0] if series_configs else None
            if price_series is None:
                print_colored(f"No se detecto informacion de Compras en la hoja {w}.", COLOR_YELLOW)
                continue
            price_df, selected_columns, price_original_df = prepare_price_index_dataframe(price_series, context.get('compras_top10', []))
            share_lookup_context = context.get('share_lookup', {})
            chart_share_lookup = {}
            if share_lookup_context and selected_columns:
                share_lookup_lower = {
                    str(key).strip().lower(): value
                    for key, value in share_lookup_context.items()
                }
                for col in selected_columns:
                    share_val = share_lookup_lower.get(str(col).strip().lower())
                    if share_val is not None:
                        chart_share_lookup[col] = share_val
            c_fig = build_price_index_slide(
                ppt,
                lang,
                cat,
                context.get('apo_entries', []),
                context.get('removed_headers', []),
                price_df,
                price_original_df,
                price_series.pipeline,
                chart_share_lookup,
                c_fig
            )
            last_reference_source = price_series.data
            last_reference_origin = price_series.display_tipo
            continue
        # Segmento 2: arbol de medidas (por que) con comparacion MAT.
        if w.startswith('2_'):
            raw_tree_df = file.parse(w, header=None)
            sheet_name_clean = w.strip()
            tree_table, tree_periods = _limpiar_tabla_excel(raw_tree_df)
            if tree_table is None or not tree_periods:
                print_colored(f"No se encontro la tabla de variables en la hoja {w}.", COLOR_YELLOW)
                continue
            if len(tree_periods) < 2:
                print_colored(f"La hoja {w} no tiene al menos dos periodos MAT para comparar.", COLOR_YELLOW)
                continue
            unidad_detectada = _unidad_desde_nombre_hoja(w) or 'Units'
            penetracion_en_pp = _modo_penetracion_desde_nombre_hoja(w) == "pp"
            periodo_inicial, periodo_final = tree_periods[-2], tree_periods[-1]
            parsed_tree_dt = _try_parse_reference_date(periodo_final)
            if parsed_tree_dt is not None:
                last_tree_period_dt = parsed_tree_dt
            try:
                metrics_calculated = calcular_cambios(
                    tree_table,
                    periodo_inicial,
                    periodo_final,
                    unidad_detectada,
                    penetracion_en_pp=penetracion_en_pp
                )
            except KeyError as exc:
                print_colored(f"No se pudo calcular el arbol para {w}: {exc}", COLOR_RED)
                continue
            c_fig += 1
            tree_stream, fig_size = graficar_arbol(
                metrics_calculated,
                volumen_unidad=unidad_detectada,
                hoja=sheet_name_clean.replace(" ", "_"),
                output_dir=None
            )
            tree_image_bytes = tree_stream.getvalue()
            tree_stream.seek(0)
            if fig_size:
                fig_width_in, fig_height_in = fig_size
            else:
                fig_width_in, fig_height_in = (21, 9)
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            titulo_arbol = c_w.get((lang, '2'), '2W - Arbol de Medidas')
            display_target = _etiqueta_segmento2_desde_nombre_hoja(sheet_name_clean)
            title_box = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
            title_tf = title_box.text_frame
            set_title_with_brand_color(
                title_tf,
                f"{titulo_arbol} | ",
                display_target,
                '',
                0.35,
                BRAND_TITLE_COLOR_LOOKUP
            )
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            chart_top = Inches(CHART_TOP_INCH)
            available_width_emu = available_width(ppt, left_margin, right_margin)
            slide_width_in = emu_to_inches(ppt.slide_width)
            slide_height_in = emu_to_inches(ppt.slide_height)
            left_margin_in = emu_to_inches(left_margin)
            right_margin_in = emu_to_inches(right_margin)
            available_width_in = max(1.0, slide_width_in - left_margin_in - right_margin_in)
            aspect_ratio = fig_height_in / fig_width_in if fig_width_in else 0.52
            max_chart_height_in = max(3.0, slide_height_in - emu_to_inches(chart_top) - 0.4)
            target_width_in = available_width_in
            target_height_in = target_width_in * aspect_ratio if aspect_ratio else 5.5
            scale_factor = 1.12
            target_width_in *= scale_factor
            target_height_in *= scale_factor
            if target_width_in > available_width_in:
                target_width_in = available_width_in
                target_height_in = target_width_in * aspect_ratio if aspect_ratio else target_height_in
            if target_height_in > max_chart_height_in:
                target_height_in = max_chart_height_in
                target_width_in = target_height_in / aspect_ratio if aspect_ratio else target_width_in
            chart_shape = slide.shapes.add_picture(
                tree_stream,
                left_margin,
                chart_top,
                width=Inches(target_width_in),
                height=Inches(target_height_in)
            )
            # Caja de comentario en posicion fija aun si sobrepone el grafico
            comment_box = slide.shapes.add_textbox(Inches(0.33), Inches(5.8), Inches(10), Inches(0.5))
            comment_tf = comment_box.text_frame
            comment_tf.clear()
            comment_para = comment_tf.paragraphs[0]
            comment_para.text = "Comentario"
            comment_para.font.size = Inches(0.25)
            excel_sheet_exports.append(
                {
                    "sheet_name": w,
                    "tables": [],
                    "trend_charts": [],
                    "images": [
                        {
                            "image_bytes": tree_image_bytes,
                            "anchor_cell": "E1",
                            "width_px": max(960, int(round(fig_width_in * DEFAULT_EXPORT_DPI * 0.65))),
                            "height_px": max(360, int(round(fig_height_in * DEFAULT_EXPORT_DPI * 0.65))),
                        }
                    ],
                }
            )
            # No imprimimos mensaje de finalización para mantener la terminal concisa
            continue
        # Segmento 1: grafico de variaciones MAT (cuando).
        if w[0]=='1':
            try:
                sheet_df = parse_sheet_with_date_fallback(file, w)
            except ValueError as exc:
                print_colored(str(exc), COLOR_RED)
                continue
            #Cria o slide
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            #Define o titulo
            txTitle = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
            tf = txTitle.text_frame
            title_prefix = c_w[(lang,w[0])]+' '+ labels[(lang,'MAT')]+' | '
            title_brand = w[2:].strip()
            set_title_with_brand_color(tf, title_prefix, title_brand, '', 0.35, BRAND_TITLE_COLOR_LOOKUP)
            #Obtém pipeline das vendas
            p = _detect_ventas_pipeline(sheet_df.columns)
            #Cria a base
            mat=df_mat(sheet_df,p)
            #Elimina linhas com divisão com zero devido ao pipeline
            if not mat.empty and mat.shape[1] > 3:
                mat=mat[~np.isinf(mat.iloc[:, 3])]
            if mat.empty and p != 0:
                mat_fallback = df_mat(sheet_df, 0)
                if not mat_fallback.empty and mat_fallback.shape[1] > 3:
                    mat_fallback = mat_fallback[~np.isinf(mat_fallback.iloc[:, 3])]
                if not mat_fallback.empty:
                    print_colored(
                        f"Segmento 1 ({w}): pipeline {p} sin datos graficables; se usa pipeline 0.",
                        COLOR_YELLOW
                    )
                    p = 0
                    mat = mat_fallback
            if mat.empty:
                print_colored(
                    f"Segmento 1 ({w}): no hay datos suficientes para construir MAT.",
                    COLOR_RED
                )
                continue
            last_reference_source = mat
            last_reference_origin = None
            #Incrementa contador do gráfico
            c_fig+=1
            #Insere o gráfico do MAT
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available = available_width(ppt, left_margin, right_margin)
            pic=slide.shapes.add_picture(graf_mat(mat,c_fig,p), left_margin, Inches(CHART_TOP_INCH),width=available)
            #Insere caixa de texto para comentário do slide
            txTitle = slide.shapes.add_textbox(Inches(0.33), Inches(5.8), Inches(10), Inches(0.5))
            tf = txTitle.text_frame
            tf.clear()
            t = tf.paragraphs[0]
            t.text = "Comentario"
            t.font.size = Inches(0.28)
            #Limpa área de plotagem
            #mensaje de conclusion por cada slide
        # Segmentos 3-6: comparativos y competencia con grafico + tabla de aportes.
        else:
            #Carrega a base
            try:
                df_start=parse_sheet_with_compras_header(file, w)
                df_list, p_ventas = split_compras_ventas(df_start, sheet_name=w)
                series_configs = prepare_series_configs(df_list, lang, p_ventas, sheet_name=w)
            except Exception as exc:
                print_colored(f"Se omite la hoja {w}: {exc}", COLOR_RED)
                continue
            if (
                size_grouping_threshold_pct is not None
                and is_segment3_size_sheet(w)
                and series_configs
            ):
                grouped_series_configs: list[SeriesConfig] = []
                for serie in series_configs:
                    grouped_df, grouped_applied = group_size_elements_by_threshold(
                        serie.data,
                        serie.pipeline,
                        size_grouping_threshold_pct
                    )
                    if grouped_applied:
                        grouped_series_configs.append(
                            SeriesConfig(
                                grouped_df,
                                serie.raw_tipo,
                                serie.display_tipo,
                                serie.pipeline
                            )
                        )
                    else:
                        grouped_series_configs.append(serie)
                series_configs = grouped_series_configs
            series_configs = reorder_series_configs_by_compras(series_configs)
            last_reference_source = series_configs[0].data if series_configs else df_start
            last_reference_origin = series_configs[0].display_tipo if series_configs else None
            #Cria o slide
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            #Define o titulo
            txTitle = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
            tf = txTitle.text_frame
            title_prefix = ''
            title_brand_label = None
            title_font_inches = 0.35
            title_color_lookup = BRAND_TITLE_COLOR_LOOKUP
            title_palette_sequence = None
            parsed_sheet = parse_standard_sheet_name(w)
            if w[0] in ['3','5'] and parsed_sheet:
                step_variant = parsed_sheet['variant']
                step_key = f"{parsed_sheet['segment']}-{step_variant}" if step_variant else parsed_sheet['segment']
                title_prefix = c_w[(lang, step_key)]+' | '
                title_brand_label = parsed_sheet['label']
                titulo = f"{title_prefix}{title_brand_label}"
            elif w[0] == '4' and parsed_sheet:
                title_prefix = c_w[(lang,w[0])] + ' | '
                title_brand_label = parsed_sheet['label']
                titulo = f"{title_prefix}{title_brand_label}"
            elif w[0]=='6':
                title_prefix =  c_w[(lang,w[0])] + ' | ' + labels[(lang,'comp')]
                title_brand_label = cat
                titulo = f"{title_prefix}{title_brand_label}"
                title_font_inches = 0.33
                title_color_lookup = {}
                title_palette_sequence = COMPETITION_TITLE_PALETTE
            else:
                fallback_label = parsed_sheet['label'] if parsed_sheet else w[2:]
                title_prefix = c_w[(lang,w[0])]+' | '+ fallback_label
                titulo = title_prefix
            set_title_with_brand_color(tf, title_prefix, title_brand_label, '', title_font_inches, title_color_lookup, title_palette_sequence)
            #Insere caixa de texto para comentário do slide
            txTitle = slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
            tf = txTitle.text_frame
            tf.clear()
            t = tf.paragraphs[0]
            t.text ="Comentario"
            t.font.size= Inches(0.25)
            # Control para la tabla de aporte
            target_table_height = Cm(TABLE_TARGET_HEIGHT_CM)
            bottom_margin = Cm(1.5)
            left_margin = int(Cm(TABLE_SIDE_MARGIN_CM))
            right_margin = left_margin
            vertical_spacing = Cm(0.2)
            target_table_height_emu = int(target_table_height)
            bottom_margin_emu = int(bottom_margin)
            slide_width = int(ppt.slide_width)
            slide_height = int(ppt.slide_height)
            removed_headers = []
            series_share_lookup = {}
            stacked_share_sources = []
            should_collect_share = False
            apo_entries = []
            grouped_footer_headers = []
            grouped_footer_label = ''
            chart_color_mappings = {}
            excel_trend_charts = []
            chart_top_emu = int(Inches(CHART_TOP_INCH))
            chart_height_emu = int(Cm(10))
            min_table_top = chart_top_emu + chart_height_emu + int(vertical_spacing)
            if w:
                first_char = w[0]
                parsed_sheet = parse_standard_sheet_name(w)
                step_variant = parsed_sheet['variant'] if parsed_sheet else ''
                if first_char == '6':
                    should_collect_share = True
                elif first_char == '3' and step_variant in {'3', '4', '5'}:
                    should_collect_share = True
                elif first_char == '5' and step_variant in {'2', '3', '4'}:
                    should_collect_share = True
            for idx, serie in enumerate(series_configs):
                grouped_columns = serie.data.attrs.get("size_grouped_columns", []) if hasattr(serie.data, "attrs") else []
                if grouped_columns:
                    grouped_footer_headers.extend(
                        [str(header).strip() for header in grouped_columns if str(header).strip()]
                    )
                    if not grouped_footer_label:
                        grouped_label_candidate = str(serie.data.attrs.get("size_grouped_label", "")).strip()
                        if grouped_label_candidate:
                            grouped_footer_label = grouped_label_candidate
                apo = aporte(serie.data.copy(), serie.pipeline, lang, serie.raw_tipo)
                apo_entries.append({
                    "apo": apo,
                    "display_tipo": serie.display_tipo,
                })
                removed_headers.extend(apo.attrs.get("removed_columns", []))
                share_values = apo.attrs.get("share_mat_values", {})
                if share_values:
                    series_share_lookup.update(share_values)
                if should_collect_share:
                    share_periods = apo.attrs.get("share_period_values", [])
                    if share_periods:
                        stacked_share_sources.append(
                            {
                                "display_tipo": serie.display_tipo,
                                "share_periods": share_periods,
                                "columns": [
                                    str(col)
                                    for col in apo.columns[1:]
                                    if str(col).strip() and str(col).strip().lower() != 'total'
                                ],
                            }
                        )
            reorder_apo_entries_by_compras(apo_entries)
            # Cuadro de texto para columnas eliminadas y/o agrupadas por umbral
            unique_removed_headers = [header for header in dict.fromkeys(removed_headers) if header]
            unique_grouped_headers = [header for header in dict.fromkeys(grouped_footer_headers) if header]
            footer_messages = []
            if unique_removed_headers:
                footer_messages.append(
                    "Se eliminaron los encabezados sin información: " + ", ".join(unique_removed_headers)
                )
            if unique_grouped_headers:
                grouped_msg = "Se agruparon las siguientes lineas: " + ", ".join(unique_grouped_headers)
                if grouped_footer_label:
                    grouped_msg = f"{grouped_msg} -> {grouped_footer_label}"
                footer_messages.append(grouped_msg)
            if footer_messages:
                footer_left = Cm(4.6)
                footer_right_margin = Cm(1.2)
                footer_width = ppt.slide_width - footer_left - footer_right_margin
                if footer_width < Cm(2):
                    footer_width = Cm(2)
                estimated_lines = sum(max(1, math.ceil(len(message) / 95)) for message in footer_messages)
                footer_height_cm = min(2.6, max(0.8, 0.18 + 0.28 * estimated_lines))
                footer_height = Cm(footer_height_cm)
                footer_bottom_gap = Cm(0.25)
                footer_top = ppt.slide_height - footer_bottom_gap - footer_height
                footer_box = slide.shapes.add_textbox(footer_left, footer_top, footer_width, footer_height)
                footer_tf = footer_box.text_frame
                footer_tf.clear()
                footer_tf.word_wrap = True
                footer_tf.text = "\n".join(footer_messages)
                for paragraph in footer_tf.paragraphs:
                    for run in paragraph.runs:
                        footer_font = run.font
                        footer_font.name = 'Arial'
                        footer_font.size = Pt(10)
                        footer_font.color.rgb = RGBColor(120, 120, 120)
            compras_top10 = []
            if series_share_lookup:
                compras_items = []
                for col_name, share_value in series_share_lookup.items():
                    if col_name is None:
                        continue
                    col_str = str(col_name).strip()
                    if not col_str:
                        continue
                    if col_str.lower().find('total') != -1:
                        continue
                    if not col_str.lower().endswith('.c'):
                        continue
                    try:
                        numeric_share = float(share_value)
                    except (TypeError, ValueError):
                        continue
                    if not np.isfinite(numeric_share):
                        continue
                    compras_items.append((col_name, numeric_share))
                if compras_items:
                    compras_items.sort(key=lambda item: item[1], reverse=True)
                    compras_top10 = [item[0] for item in compras_items[:10]]
            players_base_key = extract_players_base_key(w)
            if players_base_key and apo_entries:
                players_share_context[players_base_key] = {
                    "apo_entries": apo_entries,
                    "removed_headers": unique_removed_headers,
                    "share_lookup": dict(series_share_lookup),
                    "titulo": titulo,
                    "compras_top10": compras_top10,
                }
            if plot=="1" and len(series_configs)>1:
                df_full = series_configs[0].data.copy()
                for extra in series_configs[1:]:
                    df_full = pd.concat([df_full, extra.data.iloc[:,1:]], axis=1)
                ventas_pipelines = []
                for cfg in series_configs:
                    if str(getattr(cfg, "display_tipo", "")).strip().lower() != "ventas":
                        continue
                    try:
                        current_pipeline = int(getattr(cfg, "pipeline", 0))
                    except (TypeError, ValueError):
                        continue
                    if current_pipeline >= 0:
                        ventas_pipelines.append(current_pipeline)
                if ventas_pipelines:
                    pipeline_combined = max(ventas_pipelines)
                else:
                    fallback_pipelines = []
                    for cfg in series_configs:
                        try:
                            fallback_pipeline = int(getattr(cfg, "pipeline", 0))
                        except (TypeError, ValueError):
                            continue
                        if fallback_pipeline >= 0:
                            fallback_pipelines.append(fallback_pipeline)
                    pipeline_combined = max(fallback_pipelines) if fallback_pipelines else 0
                c_fig+=1
                left_margin = Inches(0.33)
                right_margin = Inches(0.33)
                available_line_width = available_width(ppt, left_margin, right_margin)
                chart_colors = {}
                chart_stream = line_graf(
                    df_full,
                    pipeline_combined,
                    titulo,
                    c_fig,
                    len(series_configs),
                    width_emu=available_line_width,
                    height_emu=Cm(10),
                    share_lookup=series_share_lookup,
                    color_collector=chart_colors,
                    color_overrides=chart_color_mappings,
                    force_right_labels=True,
                    show_title=False,
                )
                if isinstance(chart_colors, dict):
                    chart_color_mappings.update(chart_colors)
                combined_series_specs = []
                for cfg in series_configs:
                    dashed = _trend_block_key_from_display_tipo(cfg.display_tipo) == 'compras'
                    combined_series_specs.extend(
                        _build_excel_trend_series_specs(cfg, chart_colors, dashed=dashed)
                    )
                excel_trend_charts.append(
                    {
                        "title": titulo,
                        "category_block_key": _trend_block_key_from_display_tipo(series_configs[0].display_tipo),
                        "start_idx": pipeline_combined,
                        "series": combined_series_specs,
                        "height_emu": Cm(10),
                        "width_emu": available_line_width,
                    }
                )
                pic=slide.shapes.add_picture(chart_stream, left_margin, Inches(CHART_TOP_INCH),width=available_line_width,height=Cm(10))
            elif plot=="2" and len(series_configs)>1:
                ven_param = max(len(series_configs)-1, 1)
                left_margin = Inches(0.33)
                right_margin = Inches(0.33)
                available = available_width(ppt, left_margin, right_margin)
                count = len(series_configs)
                gap = Inches(0.1) if count > 1 else Inches(0)
                total_gap = int(gap) * (count - 1)
                effective_width = available - total_gap
                if count:
                    if effective_width <= 0:
                        chart_width = available // count
                    else:
                        chart_width = effective_width // count
                else:
                    chart_width = available
                chart_width = max(chart_width, 1)
                panel_positions = {item.display_tipo: index for index, item in enumerate(
                    sorted(series_configs, key=lambda item: origin_display_order(item.display_tipo)))}
                # Se mantiene el orden de render para conservar la asignación de colores;
                # únicamente cambia la posición física de cada origen en la diapositiva.
                for idx, serie in enumerate(series_configs):
                    c_fig+=1
                    left_position = left_margin + panel_positions[serie.display_tipo] * (chart_width + int(gap))
                    chart_colors = {}
                    chart_title = format_title_with_bold_suffix(titulo, serie.display_tipo)
                    chart_stream = line_graf(
                        serie.data,
                        serie.pipeline,
                        chart_title,
                        c_fig,
                        ven_param,
                        width_emu=chart_width,
                        height_emu=Cm(10),
                        multi_chart=True,
                        share_lookup=series_share_lookup,
                        color_collector=chart_colors,
                        color_overrides=chart_color_mappings,
                        force_right_labels=True,
                        include_origin_suffix=False
                    )
                    if isinstance(chart_colors, dict):
                        chart_color_mappings.update(chart_colors)
                    block_key = _trend_block_key_from_display_tipo(serie.display_tipo)
                    plain_chart_title = f"{titulo} - {serie.display_tipo}"
                    excel_trend_charts.append(
                        {
                            "title": plain_chart_title,
                            "category_block_key": block_key,
                            "start_idx": serie.pipeline,
                            "series": _build_excel_trend_series_specs(
                                serie,
                                chart_colors,
                                dashed=block_key == 'compras',
                            ),
                            "height_emu": Cm(10),
                            "width_emu": chart_width,
                        }
                    )
                    pic=slide.shapes.add_picture(chart_stream, left_position, Inches(CHART_TOP_INCH),width=chart_width,height=Cm(10))
                    pic.name = f"5W1H · tendencia · {serie.display_tipo}"
            elif series_configs:
                c_fig+=1
                left_margin = Inches(0.33)
                right_margin = Inches(0.33)
                available_single_width = available_width(ppt, left_margin, right_margin)
                chart_colors = {}
                chart_stream = line_graf(
                    series_configs[0].data,
                    series_configs[0].pipeline,
                    titulo,
                    c_fig,
                    len(series_configs),
                    width_emu=available_single_width,
                    height_emu=Cm(10),
                    share_lookup=series_share_lookup,
                    color_collector=chart_colors,
                    color_overrides=chart_color_mappings,
                    show_title=False,
                )
                if isinstance(chart_colors, dict):
                    chart_color_mappings.update(chart_colors)
                block_key = _trend_block_key_from_display_tipo(series_configs[0].display_tipo)
                excel_trend_charts.append(
                    {
                        "title": titulo,
                        "category_block_key": block_key,
                        "start_idx": series_configs[0].pipeline,
                        "series": _build_excel_trend_series_specs(
                            series_configs[0],
                            chart_colors,
                            dashed=False,
                        ),
                        "height_emu": Cm(10),
                        "width_emu": available_single_width,
                    }
                )
                pic=slide.shapes.add_picture(chart_stream, left_margin, Inches(CHART_TOP_INCH),width=available_single_width,height=Cm(10))
            color_lookup_keys = build_color_lookup_dict(chart_color_mappings)
            table_entries = [
                entry for entry in apo_entries
                if not entry["apo"].attrs.get("skip_table")
                and not entry["apo"].empty
                and entry["apo"].shape[1] > 1
            ]
            table_entries.sort(key=lambda item: origin_display_order(item.get('display_tipo')))
            if table_entries:
                def column_color_for_table(column_name):
                    return lookup_color_for_label(column_name, color_lookup_keys)
                def build_table_color_mapping(apo_df):
                    mapping = {}
                    for column in apo_df.columns[1:]:
                        color_value = column_color_for_table(column)
                        if color_value:
                            mapping[column] = color_value
                    return mapping
                excel_sheet_exports.append(
                    {
                        "sheet_name": w,
                        "tables": [
                            {
                                "display_tipo": entry.get("display_tipo"),
                                "apo": entry.get("apo"),
                                "column_color_mapping": build_table_color_mapping(entry.get("apo")),
                            }
                            for entry in table_entries
                        ],
                        "trend_charts": excel_trend_charts,
                    }
                )
                table_top = slide_height - target_table_height_emu - bottom_margin_emu
                if table_top > min_table_top:
                    table_top = min_table_top
                if len(table_entries) == 2:
                    gap = Cm(TABLE_PAIR_GAP_CM)
                    half_slide_width = slide_width // 2
                    gap_half = int(gap) // 2
                    paired_shapes = []
                    for index, entry in enumerate(sorted(table_entries, key=lambda item: origin_display_order(item.get('display_tipo')))):
                        apo_df = entry["apo"]
                        if index == 1:
                            left_position = half_slide_width + gap_half
                            max_width = slide_width - left_position - right_margin
                        else:
                            left_position = left_margin
                            max_width = half_slide_width - left_position - gap_half
                        c_fig += 1
                        table_colors = build_table_color_mapping(apo_df)
                        pic = slide.shapes.add_picture(
                            graf_apo(apo_df, c_fig, column_color_mapping=table_colors),
                            left_position,
                            table_top,
                            height=target_table_height
                        )
                        pic.name = f"5W1H · aportes · {entry.get('display_tipo')}"
                        constrain_picture_width(pic, max_width)
                        paired_shapes.append((pic, max_width))
                    if paired_shapes:
                        min_height = min(shape.height for shape, _ in paired_shapes if shape is not None)
                        for shape, max_width in paired_shapes:
                            if shape is None or min_height <= 0:
                                continue
                            # Escalamos de forma proporcional para igualar alturas sin deformar
                            current_h = float(shape.height)
                            current_w = float(shape.width)
                            if current_h <= 0 or current_w <= 0:
                                continue
                            scale_h = min_height / current_h
                            # Evitar cualquier upscale involuntario
                            scale_h = min(scale_h, 1.0)
                            if max_width:
                                scale_w_limit = float(max_width) / current_w
                                if scale_w_limit <= 0:
                                    scale_w_limit = scale_h
                                scale_h = min(scale_h, scale_w_limit)
                            new_h = int(current_h * scale_h)
                            new_w = int(current_w * scale_h)
                            shape.height = new_h
                            shape.width = new_w
                else:
                    for entry in table_entries:
                        apo_df = entry["apo"]
                        c_fig += 1
                        table_colors = build_table_color_mapping(apo_df)
                        top_position = slide_height - target_table_height_emu - bottom_margin_emu
                        if top_position > min_table_top:
                            top_position = min_table_top
                        left_position = left_margin
                        max_width = slide_width - left_position - right_margin
                        pic = slide.shapes.add_picture(
                            graf_apo(apo_df, c_fig, column_color_mapping=table_colors),
                            left_position,
                            top_position,
                            height=target_table_height
                        )
                        constrain_picture_width(pic, max_width)
            elif excel_trend_charts:
                excel_sheet_exports.append(
                    {
                        "sheet_name": w,
                        "tables": [],
                        "trend_charts": excel_trend_charts,
                    }
                )
            render_share_items = []
            if should_collect_share and stacked_share_sources:
                preferred_order = ['ventas', 'compras']
                ordered_sources = []
                consumed_indexes = set()
                for desired in preferred_order:
                    for idx_source, entry in enumerate(stacked_share_sources):
                        if idx_source in consumed_indexes:
                            continue
                        display_tipo = (entry.get("display_tipo") or "").strip().lower()
                        if display_tipo == desired:
                            ordered_sources.append(entry)
                            consumed_indexes.add(idx_source)
                            break
                for idx_source, entry in enumerate(stacked_share_sources):
                    if idx_source not in consumed_indexes:
                        ordered_sources.append(entry)
                        consumed_indexes.add(idx_source)
                def _sanitize_share_periods(entry_dict):
                    valid_periods = []
                    for period_label, shares in entry_dict.get("share_periods", []):
                        if not isinstance(shares, dict):
                            continue
                        sanitized = {}
                        has_positive = False
                        for column_name, value in shares.items():
                            if column_name is None:
                                continue
                            try:
                                numeric_value = float(value)
                            except (TypeError, ValueError):
                                continue
                            if not np.isfinite(numeric_value):
                                continue
                            if numeric_value < 0:
                                numeric_value = 0.0
                            sanitized[column_name] = numeric_value
                            if numeric_value > 0:
                                has_positive = True
                        if sanitized and has_positive:
                            valid_periods.append((period_label, sanitized))
                    if len(valid_periods) > 2:
                        valid_periods = valid_periods[-2:]
                    return valid_periods
                share_chart_entries = []
                for entry in ordered_sources:
                    sanitized_periods = _sanitize_share_periods(entry)
                    if sanitized_periods:
                        share_chart_entries.append(
                            {
                                "display_tipo": entry.get("display_tipo") or "",
                                "columns": entry.get("columns") or [],
                                "periods": sanitized_periods,
                            }
                        )
                if share_chart_entries:
                    reorder_share_entries_by_compras(share_chart_entries)
                    max_sections = 2
                    share_chart_entries = share_chart_entries[:max_sections]
                    palette_values = TREND_COLOR_SEQUENCE or [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]
                    share_color_lookup = dict(color_lookup_keys)
                    used_color_values = {value for value in share_color_lookup.values() if value}
                    palette_index = [0]
                    palette_length = len(palette_values)
                    def _next_palette_color():
                        if not palette_values:
                            return '#888888'
                        attempts = 0
                        while attempts < max(1, palette_length):
                            candidate = palette_values[palette_index[0] % palette_length]
                            palette_index[0] += 1
                            attempts += 1
                            if candidate not in used_color_values:
                                used_color_values.add(candidate)
                                return candidate
                        candidate = palette_values[palette_index[0] % palette_length]
                        palette_index[0] += 1
                        return candidate
                    label_color_cache: dict[str, str] = {}
                    def _color_for_share_label(label):
                        if label is None:
                            return _next_palette_color()
                        normalized_label = str(label).strip()
                        if not normalized_label or normalized_label.lower() == 'total':
                            return _next_palette_color()
                        cached_color = label_color_cache.get(normalized_label)
                        if cached_color:
                            return cached_color
                        color_value = lookup_color_for_label(normalized_label, share_color_lookup)
                        if color_value:
                            label_color_cache[normalized_label] = color_value
                            register_color_lookup(normalized_label, color_value, share_color_lookup, overwrite=False)
                            return color_value
                        color_value = _next_palette_color()
                        label_color_cache[normalized_label] = color_value
                        register_color_lookup(normalized_label, color_value, share_color_lookup, overwrite=False)
                        return color_value
                    if len(share_chart_entries) == 1:
                        entry_info = share_chart_entries[0]
                        column_candidates = [col for col in entry_info["columns"] if col]
                        for _, shares in entry_info["periods"]:
                            for column_name in shares.keys():
                                if column_name and column_name not in column_candidates:
                                    column_candidates.append(column_name)
                        color_mapping = {}
                        for column_name in column_candidates:
                            color_mapping[column_name] = _color_for_share_label(column_name)
                        for idx_period, (period_label, shares) in enumerate(entry_info["periods"]):
                            previous_shares = entry_info["periods"][idx_period - 1][1] if idx_period > 0 else None
                            render_share_items.append(
                                {
                                    "display_tipo": entry_info["display_tipo"],
                                    "period_label": period_label,
                                    "shares": shares,
                                    "previous_shares": previous_shares,
                                    "color_mapping": color_mapping,
                                    "show_delta": previous_shares is not None,
                                }
                            )
                    else:
                        for entry_info in share_chart_entries:
                            column_candidates = [col for col in entry_info["columns"] if col]
                            for _, shares in entry_info["periods"]:
                                for column_name in shares.keys():
                                    if column_name and column_name not in column_candidates:
                                        column_candidates.append(column_name)
                            color_mapping = {}
                            for column_name in column_candidates:
                                color_mapping[column_name] = _color_for_share_label(column_name)
                            latest_label, latest_shares = entry_info["periods"][-1]
                            previous_shares = entry_info["periods"][-2][1] if len(entry_info["periods"]) > 1 else None
                            render_share_items.append(
                                {
                                    "display_tipo": entry_info["display_tipo"],
                                    "period_label": latest_label,
                                    "shares": latest_shares,
                                    "previous_shares": previous_shares,
                                    "color_mapping": color_mapping,
                                    "show_delta": False,
                                }
                            )
            if render_share_items:
                delta_header = 'Crescimento vs ano anterior' if lang == 'P' else 'Crecimiento vs año anterior'
                share_slide = ppt.slides.add_slide(ppt.slide_layouts[1])
                share_title_box = share_slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
                share_tf = share_title_box.text_frame
                share_suffix = " - Share 100% Apilado"
                set_title_with_brand_color(
                    share_tf,
                    title_prefix,
                    title_brand_label,
                    share_suffix,
                    0.33,
                    title_color_lookup,
                    title_palette_sequence
                )
                comment_box = share_slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
                comment_tf = comment_box.text_frame
                comment_tf.clear()
                comment_paragraph = comment_tf.paragraphs[0]
                comment_paragraph.text = "Comentario"
                comment_paragraph.font.size = Inches(0.25)
                chart_top = Inches(0.72)
                multiple_sections = len(render_share_items) > 1
                left_boundary_in = 1.70 if multiple_sections else 2.00
                right_margin_in = 0.35
                horizontal_gap_in = 0.08 if multiple_sections else 0.30
                layout_shift_right_in = 0.35 if multiple_sections else 0.25
                caption_gap = Inches(0.12)
                caption_height = Inches(0.35)
                reference_label = 'Corte em' if lang == 'P' else 'Corte a'
                delta_box_width_in = 2.6
                delta_box_gap_in = 0.2
                delta_shift_right_in = 0.18
                delta_needed = any(
                    item.get("show_delta") and item.get("previous_shares") is not None
                    for item in render_share_items
                )
                reserved_delta_in = (delta_box_width_in + delta_box_gap_in) if delta_needed else 0.0
                slide_height_in = emu_to_inches(ppt.slide_height)
                slide_width_in = emu_to_inches(ppt.slide_width)
                chart_top_in = emu_to_inches(chart_top)
                caption_gap_in = emu_to_inches(caption_gap)
                caption_height_in = emu_to_inches(caption_height)
                available_height_in = max(3.0, slide_height_in - chart_top_in - caption_gap_in - caption_height_in - 0.2)
                chart_area_left_in = left_boundary_in
                chart_area_right_in = slide_width_in - right_margin_in - reserved_delta_in
                if chart_area_right_in <= chart_area_left_in:
                    chart_area_right_in = slide_width_in - right_margin_in
                chart_area_width_in = max(3.5, chart_area_right_in - chart_area_left_in)

                valid_items = []
                for item in render_share_items:
                    shares = item["shares"]
                    total_share = sum(max(value, 0) for value in shares.values())
                    if total_share <= 0:
                        continue
                    valid_items.append(item)
                if not valid_items:
                    continue

                # Cuando hay 2 paneles (tipico Compras/Ventas), reducir el corrimiento
                # hacia la derecha para que el conjunto quede mas centrado.
                if len(valid_items) == 2:
                    layout_shift_right_effective_in = -0.18
                else:
                    layout_shift_right_effective_in = layout_shift_right_in

                prepared_charts = []
                slot_width_in = (
                    chart_area_width_in - horizontal_gap_in * max(0, len(valid_items) - 1)
                ) / max(1, len(valid_items))
                slot_width_in = max(2.4, slot_width_in)
                max_chart_width_in = max(2.4, slot_width_in)

                for item in valid_items:
                    shares = item["shares"]
                    c_fig += 1
                    chart_stream, fig_size = stacked_share_chart(
                        item["period_label"],
                        shares,
                        item["color_mapping"],
                        c_fig,
                        title=""
                    )
                    fig_width_in, fig_height_in = fig_size
                    if fig_height_in <= 0:
                        continue
                    aspect_ratio = fig_width_in / fig_height_in if fig_height_in else 1.0
                    prepared_charts.append(
                        {
                            "item": item,
                            "stream": chart_stream,
                            "aspect_ratio": max(0.1, aspect_ratio),
                        }
                    )
                if not prepared_charts:
                    continue

                # Fija una altura uniforme para todos los gráficos del slide.
                desired_height_in = min(available_height_in, 6.2)
                uniform_height_in = desired_height_in
                for entry in prepared_charts:
                    aspect_ratio = entry["aspect_ratio"]
                    max_height_by_width = max_chart_width_in / aspect_ratio if aspect_ratio > 0 else desired_height_in
                    uniform_height_in = min(uniform_height_in, max_height_by_width)
                uniform_height_in = max(3.2, uniform_height_in)

                for entry in prepared_charts:
                    aspect_ratio = entry["aspect_ratio"]
                    entry["target_height_in"] = uniform_height_in
                    entry["target_width_in"] = min(max_chart_width_in, uniform_height_in * aspect_ratio)

                total_charts_width_in = (
                    sum(entry["target_width_in"] for entry in prepared_charts)
                    + horizontal_gap_in * max(0, len(prepared_charts) - 1)
                )
                if len(prepared_charts) == 2 and delta_needed:
                    # En Compras + Compras actual con caja de crecimiento, centra
                    # el bloque completo (2 gráficos + gap + caja delta).
                    total_group_width_in = total_charts_width_in + delta_box_gap_in + delta_shift_right_in + delta_box_width_in
                    centered_start_in = (slide_width_in - total_group_width_in) / 2.0
                    max_start_left_in = max(0.2, slide_width_in - total_group_width_in - 0.2)
                    start_left_in = min(max(0.2, centered_start_in), max_start_left_in)
                elif len(prepared_charts) == 2 and not delta_needed:
                    # En Compras/Ventas, centra con márgenes simétricos respecto al slide completo.
                    start_left_in = max(0.2, (slide_width_in - total_charts_width_in) / 2.0)
                else:
                    start_left_in = chart_area_left_in + max(0.0, (chart_area_width_in - total_charts_width_in) / 2.0)
                    max_start_left_in = max(chart_area_left_in, chart_area_right_in - total_charts_width_in)
                    start_left_in = min(start_left_in + layout_shift_right_effective_in, max_start_left_in)
                    start_left_in = max(chart_area_left_in, start_left_in)
                current_left_in = start_left_in

                for idx_chart, entry in enumerate(prepared_charts):
                    item = entry["item"]
                    shares = item["shares"]
                    target_height = Inches(entry["target_height_in"])
                    target_width = Inches(entry["target_width_in"])
                    chart_shape = share_slide.shapes.add_picture(
                        entry["stream"],
                        Inches(current_left_in),
                        chart_top,
                        width=target_width,
                        height=target_height
                    )
                    chart_shape.name = f"5W1H · participación · {item.get('display_tipo')}"
                    display_tipo_share = _normalize_simple(item.get("display_tipo") or "")
                    if ("venta" in display_tipo_share) or ("venda" in display_tipo_share):
                        border_rgb = hex_to_rgb_color(LINE_COLOR_NUMERATOR) or RGBColor(212, 172, 13)
                        chart_shape.line.fill.solid()
                        chart_shape.line.fill.fore_color.rgb = border_rgb
                        chart_shape.line.color.rgb = border_rgb
                        chart_shape.line.width = Pt(2.25)
                    caption_left = chart_shape.left
                    caption_width = target_width
                    caption_top = chart_top + target_height + caption_gap
                    if idx_chart == 0:
                        # Fondo blanco para mejorar legibilidad de "Corte a ..."
                        # sobre logos del template.
                        caption_bg_left = Inches(0.28)
                        # Lo extendemos hacia arriba para cubrir mas area del logo,
                        # y hasta abajo en la zona de caption.
                        caption_bg_top = chart_top + target_height - Inches(0.30)
                        caption_bg_bottom = caption_top + caption_height + Inches(0.02)
                        caption_bg_height = max(Inches(0.58), caption_bg_bottom - caption_bg_top)
                        caption_bg_right = min(caption_left + caption_width, caption_left + Inches(3.05))
                        caption_bg_width = max(Inches(2.60), caption_bg_right - caption_bg_left)
                        caption_bg = share_slide.shapes.add_textbox(
                            caption_bg_left,
                            caption_bg_top,
                            caption_bg_width,
                            caption_bg_height
                        )
                        caption_bg.text_frame.clear()
                        caption_bg.fill.solid()
                        caption_bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
                        caption_bg.line.fill.background()
                        # Enviar al fondo para que quede detras de los graficos/textos.
                        sp_tree = share_slide.shapes._spTree
                        sp_tree.remove(caption_bg._element)
                        sp_tree.insert(2, caption_bg._element)
                    caption_box = share_slide.shapes.add_textbox(caption_left, caption_top, caption_width, caption_height)
                    caption_tf = caption_box.text_frame
                    caption_tf.clear()
                    caption_paragraph = caption_tf.paragraphs[0]
                    origin_label = format_origin_label(item.get("display_tipo"), lang)
                    origin_suffix = f" - {origin_label}" if origin_label else ""
                    caption_paragraph.text = f"{reference_label}: {item['period_label']}{origin_suffix}"
                    caption_paragraph.font.size = Pt(11)
                    caption_paragraph.font.bold = True
                    caption_paragraph.font.color.rgb = RGBColor(80, 80, 80)
                    should_show_delta = (
                        item["show_delta"]
                        and item["previous_shares"] is not None
                        and idx_chart == len(prepared_charts) - 1
                    )
                    if should_show_delta:
                        proposed_left_in = current_left_in + entry["target_width_in"] + delta_box_gap_in + delta_shift_right_in
                        max_left_in = max(0.2, slide_width_in - delta_box_width_in - 0.2)
                        delta_box_left_in = min(proposed_left_in, max_left_in)
                        min_gap_left_in = current_left_in + entry["target_width_in"] + 0.05
                        if delta_box_left_in < min_gap_left_in:
                            delta_box_left_in = min(min_gap_left_in, max_left_in)
                        delta_box_left = Inches(delta_box_left_in)
                        # Conserva la posicion del cuadro y extiende su borde derecho
                        # exactamente hasta el limite derecho del slide.
                        delta_box_width = ppt.slide_width - delta_box_left
                        delta_box = share_slide.shapes.add_textbox(
                            delta_box_left,
                            chart_top,
                            delta_box_width,
                            chart_shape.height + caption_gap + caption_height
                        )
                        delta_tf = delta_box.text_frame
                        delta_tf.clear()
                        delta_tf.word_wrap = True
                        header_para = delta_tf.paragraphs[0]
                        header_para.text = delta_header
                        header_para.font.bold = True
                        header_para.font.size = Pt(12)
                        header_para.font.color.rgb = RGBColor(70, 70, 70)
                        growth_entries = []
                        previous_shares = item["previous_shares"] or {}
                        for brand, current_value in shares.items():
                            prev_value = previous_shares.get(brand, 0.0)
                            if current_value is None or not np.isfinite(current_value):
                                current_value = 0.0
                            if prev_value is None or not np.isfinite(prev_value):
                                prev_value = 0.0
                            delta_value = (current_value - prev_value) * 100
                            growth_entries.append((brand, delta_value, current_value))
                        growth_entries.sort(key=lambda val: val[1], reverse=True)
                        for brand, delta_value, _ in growth_entries:
                            para = delta_tf.add_paragraph()
                            para.level = 1
                            para.text = f"{brand}: {delta_value:+.1f} pp"
                            para.font.size = Pt(11)
                            color_hex = item["color_mapping"].get(brand)
                            if color_hex:
                                rgb_tuple = mcolors.to_rgb(color_hex)
                                color_rgb = tuple(int(round(c * 255)) for c in rgb_tuple)
                                para.font.color.rgb = RGBColor(*color_rgb)
                            else:
                                para.font.color.rgb = RGBColor(90, 90, 90)
                    current_left_in += entry["target_width_in"] + horizontal_gap_in
            # Mensaje final omitido para evitar ruido en consola
    chart_generation_end = dt.now()
    # Referencia de la base
    ref_source = last_reference_source if last_reference_source is not None else parse_sheet_with_compras_header(file, W[0])
    last_dt = None
    if ref_source is not None and not ref_source.empty:
        for val in reversed(ref_source.iloc[:, 0].tolist()):
            last_dt = _try_parse_reference_date(val)
            if last_dt is not None:
                break
    if last_dt is None and last_tree_period_dt is not None:
        last_dt = last_tree_period_dt

    if last_dt is None:
        ref = 'NA'
        ref_display = 'NA'
    else:
        ref = last_dt.strftime('%m-%y')
        month_list = MONTH_NAMES.get(lang, MONTH_NAMES['E'])
        month_name = month_list[last_dt.month - 1]
        ref_display = f"{month_name}-{last_dt.year}"
    slide=ppt.slides[0]
    #Fabricante
    txBox = slide.shapes.add_textbox(Inches(0.42), Inches(3.77), Inches(10), Inches(0.5))
    tf = txBox.text_frame
    tf.clear()
    t = tf.paragraphs[0]
    t.text = "Numerator vs " + client
    run = t.runs[0]
    font = run.font
    font.name, font.size, font.bold = 'Arial', Inches(0.44), True
    font.color.rgb = RGBColor(255, 255, 255)
    #Referencia
    txBox = slide.shapes.add_textbox(Inches(0.42), Inches(5.98), Inches(10), Inches(0.5))
    tf = txBox.text_frame
    tf.clear()
    t = tf.paragraphs[0]
    reference_label = 'Corte em' if lang == 'P' else 'Corte a'
    display_ref = ref_display if ref_display != 'NA' else ref
    origin_label = format_origin_label(last_reference_origin, lang)
    reference_text = f"{reference_label} {display_ref}"
    if origin_label:
        reference_text = f"{reference_text} - {origin_label}"
    t.text = reference_text
    run = t.runs[0]
    font = run.font
    font.name, font.size, font.bold = 'Arial', Inches(0.32), True
    font.color.rgb = RGBColor(255, 255, 255)
    output_filename = "-".join([
        simplify_name_segment(land, 30),
        simplify_name_segment(cat, 40),
        simplify_name_segment(client, 40),
        'informe',
        '5W1H'
    ]) + '.pptx'
    output_foldername = "-".join([
        simplify_folder_segment(land, 30),
        simplify_folder_segment(cat, 80),
        simplify_folder_segment(client, 40),
        simplify_folder_segment(ref, 10)
    ]) + '_5W1H'
    output_folder = Path(output_root)
    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / output_filename
    excel_output_filename = "-".join([
        simplify_name_segment(land, 30),
        simplify_name_segment(cat, 40),
        simplify_name_segment(client, 40),
        'informe',
        '5W1H'
    ]) + '.xlsx'
    excel_output_path = output_folder / excel_output_filename
    total_elapsed_seconds = (chart_generation_end - chart_generation_start).total_seconds()
    try:
        saved_ok = save_presentation_with_retry(
            ppt,
            output_path,
            lang,
            elapsed_seconds=total_elapsed_seconds,
        )
    except Exception as exc:
        print_colored(f"No se pudieron escribir las secciones de PowerPoint: {exc}", COLOR_RED)
        raise RuntimeError("No se pudo guardar el informe 5W1H.")
    if saved_ok is None:
        print_colored("Guardado cancelado por el usuario.", COLOR_YELLOW)
        raise RuntimeError("No se pudo guardar el informe 5W1H.")
    if saved_ok is False:
        print_file_locked_error(str(output_path), elapsed_seconds=total_elapsed_seconds)
        raise RuntimeError("No se pudo guardar el informe 5W1H.")
    excel_saved_ok = generate_excel_output_with_tables(
        base_dir / excel,
        excel_output_path,
        excel_sheet_exports,
    )
    if not excel_saved_ok:
        print_file_locked_error(str(excel_output_path), elapsed_seconds=total_elapsed_seconds)
        raise RuntimeError("No se pudo guardar el informe 5W1H.")
    print_loading_done("Generacion de slides completada")
    print_colored(f'Presentacion guardada en: {output_path}', COLOR_GREEN)
    print_colored(f'Excel guardado en: {excel_output_path}', COLOR_GREEN)
    chart_elapsed = (chart_generation_end - chart_generation_start).total_seconds()
    print_colored(
        f'Tiempo de generacion de graficos : {int(chart_elapsed // 60)} min {chart_elapsed % 60:.3f} s'
        if chart_elapsed >= 60
        else f'Tiempo de generacion de graficos : {chart_elapsed:.3f} s',
        COLOR_BLUE
    )

    file.close()
    plt.close("all")
    return {"Excel": str(excel_output_path), "PPT": str(output_path), "title": output_foldername}

if __name__ == "__main__":
    import json
    request = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    result = run_report(request["source"], request["destination"], request.get("mode", "1"), request.get("sheets"), request.get("size_threshold"))
    Path(request["result"]).write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
