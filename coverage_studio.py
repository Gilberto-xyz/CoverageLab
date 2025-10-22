"""Coverage Studio Ultra

Refactor mantenible de coverage_studio.py. Conserva la logica original pero
organiza la generacion de reportes en componentes reutilizables
"""

from __future__ import annotations

import io
import os
import re
import shutil
import sys
import threading
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Callable
from calendar import month_abbr

import colorama
from colorama import Fore, Style
from rich.console import Console
from rich.panel import Panel

colorama.init(autoreset=True)
console = Console()

CATEGORIES_CSV_DATA = """cod,cest,cat
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
CNML,Alimentos,Grano- Harina y Masa de Maiz
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
MXSN,Alimentos,Galletas - snacks y mini tostadas
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
LINI,Cuidado Personal,Oleo Calcareo
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
SCHS,Material Escolar,Utiles Escolares
CSTD,Diversos,Estudio de Categorias
CORP,Diversos,Corporativa
CROS,Diversos,Cross Category
CRBA,Diversos,Cross Category (Bebes)
CRBR,Diversos,Cross Category (Desayuno)-Yogurt - Cereal - Pan y Queso
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
FLWT,Alimentos,Agua Saborizada
"""
COUNTRY_MAP = {
    "10": "LatAm",
    "54": "Argentina",
    "91": "Bolivia",
    "55": "Brasil",
    "12": "CAM",
    "56": "Chile",
    "57": "Colombia",
    "93": "Ecuador",
    "52": "Mexico",
    "51": "Peru",
    "69": "Republica Dominicana",
    "62": "Guatemala",
    "63": "El Salvador",
    "64": "Honduras",
    "65": "Nicaragua",
    "66": "Costa Rica",
    "67": "Panamá",
}
CATEGORY_MAP: dict[str, str] = {}
for _line in CATEGORIES_CSV_DATA.splitlines()[1:]:
    _parts = _line.split(',')
    if len(_parts) >= 3:
        CATEGORY_MAP[_parts[0]] = _parts[2]

PPT_LAYOUT_INDEX = 1
DEFAULT_POP_COVERAGE = "100%"
EXCEL_TEMP_FILENAME = "file_temp_coverage.xlsx"

COL_DATA = "Data"
COL_SELL_IN = "Sell_in"
COL_SELL_OUT = "Sell_out"
COL_PENET = "Penet"
COL_COMPRA_MEDIA = "Compra_Media"
COL_COMPRA_OCA = "Compra_por_Oca"
COL_FREQ = "Freq"
COL_BUYERS = "Buyers"
COL_SELL_IN_SIM = "Sell_in_sim"
COL_ACUM_SELL_OUT = "Acum_Sell_out"
COL_ACUM_SELL_IN = "Acum_Sell_in"
COL_ANO = "Ano"
COL_TRI = "Tri"
COL_SEM = "Sem"

COLOR_KANTAR_LINE = "#2C3E50"
COLOR_SELLIN_LINE = "#D4AC0D"
COLOR_SELLOUT_LINE = "#1F618D"
COLOR_TENDENCIA_FILL = "#EBF5FB"
COLOR_COVERAGE_BAR = "#3498DB"
COLOR_PENET_LINE = "#E74C3C"
COLOR_KANTAR_BAR_VAR = '#7F8C8D'
COLOR_SELLIN_BAR_VAR = '#F1C40F'
COLOR_KANTAR_EDGE_VAR = '#2C3E50'
COLOR_SELLIN_EDGE_VAR = '#B7950B'
COLOR_COBERTURA_BAR = '#D9D9D9'
COLOR_PENETRACION_BAR = '#FFC000'
COLOR_POS_LABEL = '#1E8449'
COLOR_NEG_LABEL = '#8B0000'
COLOR_POS_LABEL_ALT = '#27AE60'
COLOR_NEG_LABEL_ALT = '#C0392B'
COLOR_SELLIN_TREND_LINE = "#D4AC0D"
COLOR_SELLOUT_TREND_LINE = "#2C3E50"

def _load_heavy_modules() -> None:
    """Carga en segundo plano las bibliotecas pesadas y datos estaticos."""
    try:
        global pd, np, dfi, plt, warnings, matplotlib, dt, timedelta, pearsonr
        global Presentation, Inches, get_column_letter, tqdm, mtick, MonthLocator
        global DateFormatter, matplotlib_style, Progress, BarColumn, TextColumn
        global TimeElapsedColumn, TimeRemainingColumn, SpinnerColumn, Image, ImageOps
        global RGBColor, Pt, MSO_SHAPE, pais, pop_coverage

        import dataframe_image as dfi
        import pandas as pd
        import numpy as np
        import warnings
        import matplotlib

        matplotlib.use("Agg")
        from matplotlib import pyplot as plt
        from datetime import datetime as dt, timedelta
        from scipy.stats import pearsonr
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.dml.color import RGBColor
        from openpyxl.utils import get_column_letter
        from openpyxl import load_workbook
        from openpyxl.formatting.rule import ColorScaleRule
        from tqdm import tqdm
        import matplotlib.ticker as mtick
        from matplotlib.dates import MonthLocator, DateFormatter
        import matplotlib.style as matplotlib_style
        from rich.progress import (
            Progress,
            BarColumn,
            TextColumn,
            TimeElapsedColumn,
            TimeRemainingColumn,
            SpinnerColumn,
        )
        from PIL import Image, ImageOps

        pd.set_option('future.no_silent_downcasting', True)
        pd.set_option('mode.chained_assignment', None)
        warnings.filterwarnings('ignore')

        _codes = sorted((int(k), v) for k, v in COUNTRY_MAP.items())
        pais = pd.DataFrame({"cod": [c for c, _ in _codes], "pais": [v for _, v in _codes]})

        pop_coverage = {
            "Argentina": "90%",
            "Bolivia": "60%",
            "Brasil": "82%",
            "Chile": "78%",
            "Colombia": "65%",
            "Ecuador": "55%",
            "Mexico": "64%",
            "Peru": "66%",
            "CAM": "74%",
            "Costa Rica": "94%",
            "El Salvador": "85%",
            "Guatemala": "69%",
            "Honduras": "65%",
            "Nicaragua": "57%",
            "Panama": "92%",
            "Republica Dominicana": "63.29%",
        }
    finally:
        LOADER_READY.set()


LOADER_READY = threading.Event()

def wait_for_heavy_modules() -> None:
    """Bloquea hasta que los módulos pesados hayan terminado de cargarse."""
    if not LOADER_READY.is_set():
        _loader_thread.join()

_loader_thread = threading.Thread(target=_load_heavy_modules)
_loader_thread.start()

SELECTIONS: Dict[str, str] = {}
ROUND_COVERAGE = False

def quick_file_metadata(filename: str) -> str:
    """Obtiene metadatos básicos del nombre de archivo."""
    base = os.path.splitext(filename)[0]
    parts = base.split('_')
    if len(parts) < 2:
        return ""
    country = COUNTRY_MAP.get(parts[0], "Desconocido")
    category = CATEGORY_MAP.get(parts[1], "Categoria desconocida")
    return f"{country} - {category}"

# --- Datos Estaticos cargados en _load_heavy_modules

# --- Función para cargar categorías ---
def load_categories():
    """Carga el catálogo de categorías desde el string embebido."""
    try:
        categories_file = io.StringIO(CATEGORIES_CSV_DATA)
        df = pd.read_csv(categories_file, dtype={'cod': str}).set_index('cod')
        if os.environ.get('SHOW_CAT_MSG', '1') == '1' and df.index.duplicated().any():
            duplicates = df.index[df.index.duplicated()].unique().tolist()
            print(
                f"{Fore.YELLOW}Advertencia: Se encontraron códigos de categoría duplicados en los datos embebidos: {duplicates}. Se usará la última entrada encontrada para cada código."
            )
        if os.environ.get('SHOW_CAT_MSG', '1') == '1':
            print(Fore.GREEN + "Datos de categorías cargados correctamente desde el script.")
        return df
    except Exception as e:
        print(f"{Fore.RED}{Style.BRIGHT}Error Crítico al cargar datos de categorías desde el string embebido: {e}")
        exit()

# --- Variables Globales y Funciones de Utilidad ---
# --- Variables Globales y Funciones de Utilidad ---
# Nota: SELECTIONS y ROUND_COVERAGE se declaran al inicio del modulo.

def _round_half_up_series(series):
    """Redondea una serie numérica al entero más cercano con umbral .5 (ROUND_HALF_UP).
    Devuelve float con .0 para mantener NaN compatibles.
    """
    # Requiere numpy/pandas cargados; se usa después de esperar la carga pesada
    arr = series.to_numpy(dtype=float)
    # Usar isfinite para evitar afectar NaN/inf
    mask = np.isfinite(arr)
    arr[mask] = np.floor(arr[mask] + 0.5)
    return pd.Series(arr, index=series.index, name=series.name)

def round_coverage_flag():
    """Pregunta/lee si se debe redondear la cobertura (sin decimales, .5 hacia arriba)."""
    env_val = os.environ.get("AUTO_ROUND_COV")
    if env_val is not None:
        env_val_norm = str(env_val).strip().lower()
        do_round = env_val_norm in {"1", "true", "yes", "y", "si", "sí"}
    else:
        print(Fore.CYAN + "\n¿Desea redondear la cobertura (sin decimales, umbral .5)?")
        print(Fore.WHITE + "1 - Sí")
        print(Fore.WHITE + "2 - No")
        opciones = {'1': True, '2': False}
        eleccion = input(Fore.GREEN + "Elija 1 o 2: ")
        do_round = opciones.get(eleccion, False)
    SELECTIONS['Redondeo Cobertura'] = 'Sí' if do_round else 'No'
    clear_and_print_summary()
    return do_round

def clear_and_print_summary():
    """Limpia la terminal y muestra un resumen de las selecciones del usuario."""
    os.system('cls' if os.name == 'nt' else 'clear') # Compatible con Windows y Linux/Mac
    print(Fore.CYAN + Style.BRIGHT + "Resumen de opciones seleccionadas:")
    if 'Excel' in SELECTIONS:
        print(Fore.BLUE + "Archivo Excel: " + Fore.YELLOW + f"{SELECTIONS['Excel']}")
    if 'Cobertura' in SELECTIONS:
        print(Fore.BLUE + "Tipo de cobertura: " + Fore.YELLOW + f"{SELECTIONS['Cobertura']}")
    if 'Razón' in SELECTIONS:
        print(Fore.BLUE + "Razón de Cobertura: " + Fore.YELLOW + f"{SELECTIONS['Razón']}")
    if 'Eje tendencia' in SELECTIONS:
        print(Fore.BLUE + "Tipo de gráfico (tendencia): " + Fore.YELLOW + f"{SELECTIONS['Eje tendencia']}")
    if 'Idioma PPT' in SELECTIONS:
        print(Fore.BLUE + "Idioma PPT: " + Fore.YELLOW + f"{SELECTIONS['Idioma PPT']}")
    elif 'Inglés' in SELECTIONS:
        print(Fore.BLUE + "Idioma PPT: " + Fore.YELLOW + ("ENGLISH" if SELECTIONS['Inglés'] == 'Sí' else ("PORTUGUES" if SELECTIONS.get('Pais') == 'Brasil' else "ESPAÑOL")))
    if 'Redondeo Cobertura' in SELECTIONS:
        print(Fore.BLUE + "Redondeo de Cobertura: " + Fore.YELLOW + f"{SELECTIONS['Redondeo Cobertura']}")
    print("\n" + "-"*50 + "\n")

def print_file_header(idx: int, total: int, filename: str) -> None:
    """Muestra un encabezado visual para la ejecución de un archivo."""
    console.rule(f"[bold cyan]Procesando archivo {idx}/{total}: {filename}")

# --- Función para mostrar resumen de archivos generados ---
def print_file_summary(ruta_excel: str, ruta_ppt: str, ruta_banco: str) -> None:
    """Muestra un resumen con las rutas generadas para el archivo."""
    console.print("\n[blue]Resumen de archivos generados:[/blue]")
    if ruta_excel:
        console.print(f"[cyan]Excel:[/] [grey]{ruta_excel}")
    if ruta_ppt:
        console.print(f"[cyan]Presentación:[/] [grey]{ruta_ppt}")
    if ruta_banco:
        console.print(f"[cyan]Banco:[/] [grey]{ruta_banco}")
    # Mostrar panel de proceso completado con hora actual
    hora_actual = datetime.now().strftime("%H:%M:%S")
    mensaje = (
        "[bright_white]Proceso completado[/bright_white]\n\n"
        f"[white]Hora de finalización: [bold]{hora_actual}[/bold][/white]"
    )
    console.print()
    console.print(Panel.fit(mensaje, border_style="cyan", title="Coverages Latam"))
    console.print()



def calc_var1(df, coluna, p):
    """
    Calcula variaciones vs período anterior (Y-1) en Python.

    Args:
        df (pd.DataFrame): DataFrame con los datos.
        coluna (str): Nombre de la columna a calcular (e.g., COL_SELL_OUT).
        p (int): Pipeline (shift para Sell_in).

    Returns:
        list: Lista con variaciones [Anual, Semestral, Trimestral].
              Retorna NaN para cálculos imposibles (datos insuficientes).
    """
    n_rows = len(df)
    variations = []

    # Anual (12 vs 12 meses)
    if n_rows >= 24 + p:
        current_sum = df[coluna][n_rows-12-p : n_rows-p].sum() if p != 0 else df[coluna][-12:].sum()
        previous_sum = df[coluna][n_rows-24-p : n_rows-12-p].sum() if p!= 0 else df[coluna][-24:-12].sum()
        variations.append((current_sum / previous_sum) - 1 if previous_sum else np.nan)
    else:
        variations.append(np.nan)

    # Semestral (6 vs 6 meses)
    if n_rows >= 12 + p:
        current_sum = df[coluna][n_rows-6-p : n_rows-p].sum() if p != 0 else df[coluna][-6:].sum()
        previous_sum = df[coluna][n_rows-12-p : n_rows-6-p].sum() if p!= 0 else df[coluna][-12:-6].sum()
        variations.append((current_sum / previous_sum) - 1 if previous_sum else np.nan)
    else:
        variations.append(np.nan)

    # Trimestral (3 vs 3 meses)
    if n_rows >= 6 + p:
        current_sum = df[coluna][n_rows-3-p : n_rows-p].sum() if p != 0 else df[coluna][-3:].sum()
        previous_sum = df[coluna][n_rows-6-p : n_rows-3-p].sum() if p!= 0 else df[coluna][-6:-3].sum()
        variations.append((current_sum / previous_sum) - 1 if previous_sum else np.nan)
    else:
        variations.append(np.nan)

    return variations


def calc_var2(df, coluna, p):
    """
    Calcula variaciones vs período retrasado (Y-2) en Python.

    Args:
        df (pd.DataFrame): DataFrame con los datos.
        coluna (str): Nombre de la columna a calcular (e.g., COL_SELL_OUT).
        p (int): Pipeline (shift para Sell_in).

    Returns:
        list: Lista con variaciones [Anual, Semestral, Trimestral].
              Retorna NaN para cálculos imposibles (datos insuficientes).
    """
    n_rows = len(df)
    variations = []

    # Anual (12 meses actuales vs 12 meses de hace 2 años)
    if n_rows >= 36 + p:
        current_sum = df[coluna][n_rows-12-p : n_rows-p].sum() if p != 0 else df[coluna][-12:].sum()
        previous_sum = df[coluna][n_rows-36-p : n_rows-24-p].sum() if p!= 0 else df[coluna][-36:-24].sum()
        variations.append((current_sum / previous_sum) - 1 if previous_sum else np.nan)
    else:
        variations.append(np.nan)

    # Semestral (6 meses actuales vs 6 meses de hace 2 años) - CORREGIDO
    if n_rows >= 30 + p: # Necesitamos 6 actuales + 24 para ir 2 años atrás
        current_sum = df[coluna][n_rows-6-p : n_rows-p].sum() if p != 0 else df[coluna][-6:].sum()
        previous_sum = df[coluna][n_rows-30-p : n_rows-24-p].sum() if p!= 0 else df[coluna][-30:-24].sum()
        variations.append((current_sum / previous_sum) - 1 if previous_sum else np.nan)
    else:
        variations.append(np.nan)

    # Trimestral (3 meses actuales vs 3 meses de hace 2 años) - CORREGIDO
    if n_rows >= 27 + p: # Necesitamos 3 actuales + 24 para ir 2 años atrás
        current_sum = df[coluna][n_rows-3-p : n_rows-p].sum() if p != 0 else df[coluna][-3:].sum()
        previous_sum = df[coluna][n_rows-27-p : n_rows-24-p].sum() if p!= 0 else df[coluna][-27:-24].sum()
        variations.append((current_sum / previous_sum) - 1 if previous_sum else np.nan)
    else:
        variations.append(np.nan)

    return variations


def escalona(df_to_scale):
    """
    Desplaza los valores de cada columna hacia abajo, rellenando con NaN al principio.
    Se utiliza para alinear datos en fórmulas de Excel para cálculos de cobertura.

    Args:
        df_to_scale (pd.DataFrame): DataFrame cuyas columnas serán escalonadas.
    """
    for col in df_to_scale.columns:
        col_idx = df_to_scale.columns.get_loc(col)
        values = list(df_to_scale[col].values)
        # Invierte, trunca desde el inicio según índice, rellena, invierte de nuevo
        scaled_values = (values[::-1][col_idx:] + [np.nan]*col_idx)[::-1]
        df_to_scale[col] = scaled_values

def razao_cov():
    """Devuelve la razón de cobertura elegida o obtenida de las variables de entorno."""
    if os.environ.get("AUTO_RAZON"):
        razon_seleccionada = os.environ["AUTO_RAZON"]
    else:
        print(Fore.CYAN + "\nPregunta: ¿Cuál es la razón de la cobertura?")
        print(Fore.WHITE + "Opciones:")
        print(Fore.WHITE + "1 - Actualización periódica por contrato")
        print(Fore.WHITE + "2 - Conocer nivel de cobertura o pipeline")
        print(Fore.WHITE + "3 - Tendencias Contrarias")
        print(Fore.WHITE + "4 - Renovación de contrato")
        print(Fore.WHITE + "5 - Otras")

        razones = {
            '1': "Actualización periódica por contrato",
            '2': "Conocer nivel de cobertura o pipeline",
            '3': "Tendencias Contrarias",
            '4': "Renovación de contrato",
            '5': "Otras"
        }
        eleccion = input(Fore.GREEN + "Elija el número de la opción (1-5): ")
        razon_seleccionada = razones.get(eleccion, "Otras")  # Default a 'Otras'
    SELECTIONS['Razón'] = razon_seleccionada
    clear_and_print_summary()
    return razon_seleccionada

def tipo_cobertura():
    """Obtiene el tipo de cobertura interactivo o desde las variables de entorno."""
    if os.environ.get("AUTO_COV_TYPE"):
        tipo_seleccionado = os.environ["AUTO_COV_TYPE"]
    else:
        print(Fore.CYAN + "\nPregunta: ¿Qué tipo de cobertura se calculará?")
        print(Fore.WHITE + "Opciones:")
        print(Fore.WHITE + "1 - Cobertura Absoluta")
        print(Fore.WHITE + "2 - Cobertura Relativa")
        print(Fore.WHITE + "3 - AUTO (usar configuración predeterminada)")
        tipos = {'1': "Absoluta", '2': "relativa", '3': "AUTO"}
        eleccion = input(Fore.GREEN + "Elija 1, 2 o 3: ")
        tipo_seleccionado = tipos.get(eleccion, "Absoluta")  # Default a 'Absoluta'
    SELECTIONS['Cobertura'] = tipo_seleccionado
    clear_and_print_summary()
    return tipo_seleccionado

def tipo_eje_tendencia():
    """Elige tipo de gráfico de tendencia de forma interactiva o vía variables de entorno."""
    if os.environ.get("AUTO_EJE"):
        tipo_eje = os.environ["AUTO_EJE"]
    else:
        print(Fore.CYAN + "\n¿Desea el gráfico de tendencia con doble eje?")
        print(Fore.WHITE + "1 - Un solo eje (Sell-in y WP by Numerator juntos)")
        print(Fore.WHITE + "2 - Doble eje (WP by Numerator en eje secundario)")
        opciones = {'1': "simple", '2': "doble"}
        eleccion = input(Fore.GREEN + "Elija 1 o 2: ")
        tipo_eje = opciones.get(eleccion, "simple")
    SELECTIONS['Eje tendencia'] = tipo_eje
    clear_and_print_summary()
    return tipo_eje

def include_english_flag() -> bool:
    """Determina si se deben generar salidas en inglés.

    Usa AUTO_ENGLISH cuando está disponible; de lo contrario, solicita al usuario su preferencia.
    """
    env_val = os.environ.get("AUTO_ENGLISH")
    if env_val is not None:
        include_en = str(env_val).strip().lower() in {"1", "true", "yes", "y", "si", "sí"}
    else:
        print(Fore.CYAN + "\n¿Desea generar la presentación en inglés?")
        print(Fore.WHITE + "1 - Sí (usar bloque ENGLISH de la plantilla)")
        print(Fore.WHITE + "2 - No (usar idioma por país)")
        opciones = {"1": True, "2": False, "si": True, "no": False, "s": True, "n": False}
        while True:
            eleccion = input(Fore.GREEN + "Elija 1 o 2: ").strip().lower()
            if eleccion in opciones:
                include_en = opciones[eleccion]
                break
            if eleccion in {"", "\n"}:
                include_en = False
                break
            print(Fore.RED + "Entrada inválida. Intente nuevamente.")
    SELECTIONS['Inglés'] = 'Sí' if include_en else 'No'
    clear_and_print_summary()
    return include_en

def load_and_preprocess_sheet(excel_file_obj, sheet_name):
    """
    Carga una hoja del archivo Excel, la preprocesa (renombra, limpia, fechas)
    y devuelve el DataFrame procesado y la unidad de medida.

    Args:
        excel_file_obj (pd.ExcelFile): Objeto ExcelFile abierto.
        sheet_name (str): Nombre de la hoja a procesar.

    Returns:
        tuple: (pd.DataFrame, str) - El DataFrame procesado y la unidad de medida.
               Retorna (None, None) si hay un error al cargar o procesar.
    """
    try:
        df_sheet = excel_file_obj.parse(sheet_name)
        # Validar estructura mínima esperada (al menos 2 filas, 8 columnas)
        rows, cols = df_sheet.shape
        if rows < 2 or cols < 8:
            if cols == 7:
                # Caso específico: 7 columnas → probablemente falta Sell-in del cliente
                print(
                    f"{Fore.RED}{Style.BRIGHT}Error:{Style.RESET_ALL} "
                    f"La hoja '{sheet_name}' no cumple la estructura mínima "
                    f"({rows} filas, {cols} columnas)."
                )
                print(
                    f"{Fore.LIGHTMAGENTA_EX}{Style.BRIGHT}Sugerencia:{Style.RESET_ALL} "
                    f"Probablemente falta la columna de Sell-in del cliente."
                )
                # (Opcional) Ayuda de depuración:
                # print(f"{Fore.LIGHTMAGENTA_EX}Columnas detectadas: {list(df_sheet.columns)}{Style.RESET_ALL}")
            else:
                # Otros casos (<8 columnas o <2 filas)
                print(
                    f"{Fore.RED}{Style.BRIGHT}Error:{Style.RESET_ALL} "
                    f"La hoja '{sheet_name}' tiene una estructura inesperada "
                    f"({rows} filas, {cols} columnas). Se omitirá."
                )
            return None, None


        # === Validación temprana adicional: abortar si la columna 8 no tiene datos ===
        # Si existen los 8 encabezados pero no hay datos debajo del encabezado de la columna 8,
        # se omite la hoja para evitar que el programa se rompa más adelante.
        try:
            _col8 = df_sheet.iloc[1:, 7]  # índice 0-based: 7 es la 8ª columna
            _col8_empty = _col8.isna().all() or (_col8.astype(str).str.strip() == '').all()
        except Exception:
            _col8_empty = True  # si por alguna razón falla, tratamos como vacío

        if _col8_empty:
            print(f"{Fore.RED}Advertencia: La hoja '{sheet_name}' se omitirá porque la columna 8 (Sell-in) no tiene datos debajo del encabezado.")
            return None, None
        # === Fin validación adicional ===


        # Obtiene la 'unidad' o 'medida' de la primera fila, columna 2 (índice 1)
        measure = str(df_sheet.iat[0, 1]).replace('Weighted', '').strip()

        # Renombra las columnas al formato estándar
        df_sheet.columns = [COL_DATA, COL_SELL_OUT, COL_PENET, COL_COMPRA_MEDIA, COL_COMPRA_OCA, COL_FREQ, COL_BUYERS, COL_SELL_IN] + list(df_sheet.columns[8:]) # Mantiene columnas extra si existen
        df_sheet = df_sheet.loc[:, [COL_DATA, COL_SELL_IN, COL_SELL_OUT, COL_COMPRA_MEDIA, COL_COMPRA_OCA, COL_FREQ, COL_PENET, COL_BUYERS]] # Reordena y selecciona

        # Elimina la primera fila (encabezados repetidos) y resetea el índice
        df_sheet = df_sheet.iloc[1:].reset_index(drop=True)

        # Convierte la columna "Data" a tipo datetime
        # Maneja posibles errores de formato o valores nulos
        original_dates = df_sheet[COL_DATA].copy() # Guardar original por si falla
        try:
            # Intenta convertir primero todos los que sean strings
            is_string = df_sheet[COL_DATA].apply(lambda x: isinstance(x, str))
            if is_string.any():
                 # Intenta formato específico primero, maneja errores individuales
                 df_sheet.loc[is_string, COL_DATA] = df_sheet.loc[is_string, COL_DATA].apply(
                     lambda x: dt.strptime(x, '%b-%y  ') if isinstance(x, str) and re.match(r'\w{3}-\d{2}\s{2}', x) else x
                 )
            # Convierte el resto (o los ya convertidos) a datetime
            df_sheet[COL_DATA] = pd.to_datetime(df_sheet[COL_DATA], errors='coerce')
        except Exception as e:
             print(f"{Fore.YELLOW}Advertencia: Problema al convertir fechas en hoja '{sheet_name}'. Error: {e}. Se usará la columna original si es posible.")
             df_sheet[COL_DATA] = pd.to_datetime(original_dates, errors='coerce') # Reintentar con la original

        # Eliminar filas donde la fecha no se pudo convertir (NaT)
        initial_rows = len(df_sheet)
        df_sheet.dropna(subset=[COL_DATA], inplace=True)
        if len(df_sheet) < initial_rows:
            print(f"{Fore.YELLOW}Advertencia: Se eliminaron {initial_rows - len(df_sheet)} filas de la hoja '{sheet_name}' por fechas inválidas.")

        if df_sheet.empty:
            print(f"{Fore.red}Advertencia: La hoja '{sheet_name}' está vacía o no contiene fechas válidas después del preprocesamiento. Se omitirá.")
            return None, None

        # Asegurar tipos numéricos (intentar convertir, rellenar NaN con 0 si falla)
        numeric_cols = [COL_SELL_IN, COL_SELL_OUT, COL_COMPRA_MEDIA, COL_COMPRA_OCA, COL_FREQ, COL_PENET, COL_BUYERS]
        for col in numeric_cols:
            df_sheet[col] = pd.to_numeric(df_sheet[col], errors='coerce').fillna(0)

        # Añade columnas de Año, Trimestre, Semestre
        df_sheet[COL_ANO] = df_sheet[COL_DATA].dt.year
        df_sheet[COL_TRI] = df_sheet[COL_DATA].dt.quarter
        df_sheet[COL_SEM] = (df_sheet[COL_DATA].dt.month - 1) // 6 + 1
        df_sheet[COL_DATA] = df_sheet[COL_DATA].dt.date # Convertir a solo fecha al final

        return df_sheet, measure

    except Exception as e:
        print(f"{Fore.RED}Error crítico al cargar o preprocesar la hoja '{sheet_name}': {e}")
        return None, None


# --- Funciones de Generación de Gráficos ---

def generar_grafico_evolucion_mensual(df_graf, pipeline_meses=0, lang_idx=2):
    """
    Genera un gráfico de evolución mensual de WP by Numerator vs Sell-in con variación interanual.

    Args:
        df_graf (pd.DataFrame): DataFrame con datos mensuales (col 'Data' debe ser datetime).
        pipeline_meses (int): Número de meses de pipeline para desplazar Sell-in.

    Returns:
        matplotlib.figure.Figure: Figura de matplotlib con el gráfico, o None si no hay datos.
    """
    if df_graf is None or df_graf.empty or len(df_graf) < 24: # Necesita al menos 24 meses para var YOY
        print(f"{Fore.YELLOW}Advertencia: No se puede generar gráfico de evolución mensual. Datos insuficientes (se requieren >= 24 meses).")
        return None

    # Usar contexto de estilo para evitar afectar otros gráficos
    with matplotlib.style.context('seaborn-v0_8-whitegrid'):
        df_plot = df_graf.copy()
        df_plot[COL_DATA] = pd.to_datetime(df_plot[COL_DATA]) # Asegurar datetime

        # Si hay pipeline, desplazar Sell-in y guardar original si es necesario
        if pipeline_meses > 0:
            # df_plot["Sell_in_original"] = df_plot[COL_SELL_IN].copy() # Descomentar si se necesita el original
            df_plot[COL_SELL_IN] = df_plot[COL_SELL_IN].shift(pipeline_meses)

        # Calcular sumas móviles y variaciones interanuales
        df_plot["Kantar_12m"] = df_plot[COL_SELL_OUT].rolling(12).sum()
        df_plot["Sellin_12m"] = df_plot[COL_SELL_IN].rolling(12).sum()
        df_plot["Kantar_yoy"] = ((df_plot["Kantar_12m"] / df_plot["Kantar_12m"].shift(12)) - 1) * 100
        df_plot["Sellin_yoy"] = ((df_plot["Sellin_12m"] / df_plot["Sellin_12m"].shift(12)) - 1) * 100

        # Filtrar NaNs resultantes de rolling/shift
        df_plot = df_plot.dropna(subset=["Kantar_yoy", "Sellin_yoy"]).copy()

        if df_plot.empty:
            print(f"{Fore.YELLOW}Advertencia: No quedan datos para el gráfico de evolución después de calcular YOY.")
            return None

        # Crear figura y ejes con márgenes personalizados
        fig = plt.figure(figsize=(16.5, 8), dpi=100) # Ajustar tamaño si es necesario
        left_margin, right_margin, bottom_margin, top_margin = 0.08, 0.92, 0.18, 0.90
        ax1 = fig.add_axes([left_margin, bottom_margin, right_margin-left_margin, top_margin-bottom_margin])
        ax2 = ax1.twinx()

        # Eje primario (Líneas)
        sellin_label = (
            f"{COL_SELL_IN} (Mensual)" if lang_idx != 3 else f"{COL_SELL_IN} (Monthly)"
        ) + (f" - P:{pipeline_meses}" if pipeline_meses > 0 else "")
        ax1.plot(
            df_plot[COL_DATA], df_plot[COL_SELL_OUT],
            color=COLOR_KANTAR_LINE, marker="o", linewidth=2, markersize=5,
            label=f"{COL_SELL_OUT} (Mensual)" if lang_idx != 3 else f"{COL_SELL_OUT} (Monthly)"
        )
        ax1.plot(df_plot[COL_DATA], df_plot[COL_SELL_IN], color=COLOR_SELLIN_LINE, marker="o", linewidth=2, markersize=5, label=sellin_label)
        ax1.set_ylabel("Volumen Mensual" if lang_idx != 3 else "Monthly Volume", fontsize=11, labelpad=15)
        ax1.tick_params(axis='y', labelsize=9)
        ax1.set_ylim(bottom=0)
        ax1.grid(axis='y', linestyle='--', alpha=0.4)

        # Eje secundario (Barras de Variación YOY)
        width = 8
        offset = 4
        ax2.bar(df_plot[COL_DATA] - pd.DateOffset(days=offset), df_plot["Kantar_yoy"], width=width, color=COLOR_KANTAR_BAR_VAR, edgecolor=COLOR_KANTAR_EDGE_VAR, alpha=0.7, label="% Var Worldpanel by Numerator")
        ax2.bar(df_plot[COL_DATA] + pd.DateOffset(days=offset), df_plot["Sellin_yoy"], width=width, color=COLOR_SELLIN_BAR_VAR, edgecolor=COLOR_SELLIN_EDGE_VAR, alpha=0.7, label="% Var Sell-in")
        ax2.set_ylabel("Variación Interanual (%)" if lang_idx != 3 else "Year-over-Year Change (%)", fontsize=11, labelpad=15)
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
        ax2.tick_params(axis='y', labelsize=9)
        ax2.axhline(y=0, color='gray', linestyle='-', alpha=0.5, linewidth=0.8)

        # Etiquetas en barras
        for _, row in df_plot.iterrows():
            for col_yoy, x_offset, color_pos, color_neg in [("Kantar_yoy", -offset, COLOR_POS_LABEL, COLOR_NEG_LABEL),
                                                             ("Sellin_yoy", offset, COLOR_POS_LABEL_ALT, COLOR_NEG_LABEL_ALT)]:
                if not pd.isna(row[col_yoy]):
                    valor = row[col_yoy]
                    pos_vert = valor + 1 if valor >= 0 else valor - 1
                    va_align = "bottom" if valor >= 0 else "top"
                    color_etiq = color_pos if valor >= 0 else color_neg
                    ax2.text(row[COL_DATA] + pd.Timedelta(days=x_offset), pos_vert, f"{valor:.1f}%",
                             ha="center", va=va_align, fontsize=7, fontweight='bold', color=color_etiq) # Añadir etiquetas, con un decimal y color según el signo

        # Ajustar límites eje Y secundario para dar espacio a etiquetas
        y2_min, y2_max = ax2.get_ylim()
        padding = max(abs(y2_min), abs(y2_max)) * 0.15 # 15% padding
        ax2.set_ylim(y2_min - padding, y2_max + padding*2) # Más espacio arriba

        # Formato Eje X (Fechas) con extensión de un mes antes y después
        fechas_validas = df_plot[COL_DATA]
        fecha_min = fechas_validas.min() - pd.DateOffset(months=1)
        fecha_max = fechas_validas.max() + pd.DateOffset(months=1)
        ax1.set_xlim([fecha_min, fecha_max])
        ax1.xaxis.set_major_locator(MonthLocator(interval=1)) # Ajustar intervalo dinámicamente
        ax1.xaxis.set_major_formatter(DateFormatter('%b-%y'))
        ax1.tick_params(axis='x', rotation=45, labelsize=8)

        # Título y Leyenda
        # titulo = "Evolución Mensual y Variación " + (f" (Pipeline: {pipeline_meses})" if pipeline_meses > 0 else "")
        # fig.suptitle(titulo, fontsize=16, fontweight='bold', y=top_margin + 0.05) # Título de la figura
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left", bbox_to_anchor=(0.01, 0.98), fontsize=9, frameon=True, framealpha=0.8)

        # No usar tight_layout con add_axes, márgenes manuales ya aplicados
        # fig.tight_layout(rect=[0, 0, 1, 0.95]) # Ajustar rect si el título se solapa
  
        return fig

def generar_grafico_cobertura(slide, marca_clean, pipeline, df_cov_pipe, df_pen_pipe, lang_idx, coverage_label, labels_dict):
    """Genera el gráfico de barras de Cobertura vs Penetración y lo añade al slide."""
    cov_series = df_cov_pipe if isinstance(df_cov_pipe, pd.Series) else pd.Series(df_cov_pipe)
    pen_series = df_pen_pipe if isinstance(df_pen_pipe, pd.Series) else pd.Series(df_pen_pipe)
    cov_series = cov_series.rename('coverage')
    pen_series = pen_series.rename('penetracion')
    combined = pd.concat([cov_series, pen_series], axis=1, join='inner')
    combined = combined.dropna(subset=['coverage', 'penetracion'])
    if combined.empty:
        print(f"{Fore.YELLOW}Advertencia: No hay datos suficientes para el gráfico de cobertura/penetración (Marca: {marca_clean}, P:{pipeline}).")
        return
    cov_data = combined['coverage'].to_numpy(dtype=float)
    pen_data = combined['penetracion'].to_numpy(dtype=float)
    x_labels = [idx.strftime('%m-%y') if hasattr(idx, 'strftime') else str(idx) for idx in combined.index]
    x_pos = np.arange(len(x_labels))
    fig_cov, ax_cov = plt.subplots(figsize=(12, 4.25), dpi=100)
    bar_width = 0.35
    offset = bar_width / 2
    rects2 = ax_cov.bar(
        x_pos - offset / 1.2,
        pen_data,
        bar_width,
        label=labels_dict.get((lang_idx, 'Graf cob Penet Men'), 'Penetración Mensual'),
        color=COLOR_PENETRACION_BAR,
        edgecolor='black',
        zorder=1,
    )
    rects1 = ax_cov.bar(
        x_pos + offset,
        cov_data,
        bar_width,
        label=coverage_label,
        color=COLOR_COBERTURA_BAR,
        edgecolor='black',
        linewidth=2,
        zorder=2,
        alpha=0.85,
    )
    for rect_group in (rects2, rects1):
        for i, rect in enumerate(rect_group):
            height = rect.get_height()
            if height > 0.1:
                bbox_props = dict(facecolor='#F2F2F2', edgecolor='black', boxstyle='round,pad=0.3')
                if rect_group is rects1:
                    if i % 12 == (len(rect_group) % 12 - 1):
                        bbox_props['facecolor'] = '#A6A6A6'
                        bbox_props['edgecolor'] = 'black'
                    label_txt = f"{int(np.floor(height + 0.5))}" if globals().get('ROUND_COVERAGE', False) else f"{height:.1f}"
                else:
                    bbox_props['facecolor'] = '#FDEAD9'
                    label_txt = f"{height:.1f}"
                ax_cov.annotate(
                    label_txt,
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center',
                    va='bottom',
                    fontsize=8,
                    bbox=bbox_props,
                )
    ax_cov.set_ylabel(
        f"{coverage_label} | {labels_dict.get((lang_idx, 'Graf cob Penet Men'), 'Penetración Mensual')}",
        fontsize=9,
    )
    title_key = 'Titulo Cob'
    default_title = 'Cobertura Año Móvil' if lang_idx != 3 else 'MOVING YEAR COVERAGE'
    ax_cov.set_title(
        f"{labels_dict.get((lang_idx, title_key), default_title)} | {marca_clean} Pipeline {pipeline}",
        size=16,
    )
    ax_cov.set_xticks(x_pos)
    ax_cov.set_xticklabels(x_labels, rotation=30, ha='right', fontsize=9)
    ax_cov.legend(
        loc='lower center',
        bbox_to_anchor=(0.5, -0.30),
        frameon=False,
        prop={'size': 11},
        ncol=2,
    )
    ax_cov.grid(axis='y', linestyle='--', alpha=0.6)
    ax_cov.set_axisbelow(True)
    ax_cov.spines['top'].set_visible(False)
    ax_cov.spines['right'].set_visible(False)
    ax_cov.spines['left'].set_visible(False)
    max_val = max(np.nanmax(cov_data) if cov_data.size else 0, np.nanmax(pen_data) if pen_data.size else 0)
    ax_cov.set_ylim(bottom=0, top=max_val * 1.15 if max_val else 1)
    ax_cov.margins(x=0)
    plt.tight_layout()
    img_stream = io.BytesIO()
    fig_cov.savefig(img_stream, format='png', bbox_inches='tight', pad_inches=0.1, transparent=True)
    img_stream.seek(0)
    img_pil = Image.open(img_stream)
    bordered = ImageOps.expand(img_pil, border=1, fill='black')
    img_stream_bordered = io.BytesIO()
    bordered.save(img_stream_bordered, format='PNG')
    img_stream_bordered.seek(0)
    slide.shapes.add_picture(img_stream_bordered, Inches(0.5), Inches(2.0), height=Inches(4.2))
    plt.close(fig_cov)

def generar_grafico_tendencia(slide, marca_clean, pipeline, df_plot, lang_idx, labels_dict, doble_eje=False):
    """
    Genera el gráfico de líneas de Tendencia (Sell-in vs Sell-out) y lo añade al slide.
    Si doble_eje=True, WP by Numerator (Sell-out) va en eje secundario.
    """
    if df_plot is None or df_plot.empty or pipeline >= len(df_plot):
         print(f"{Fore.YELLOW}Advertencia: Datos insuficientes para gráfico de Tendencia (Marca: {marca_clean}, P:{pipeline}).")
         return

    fig_trend, ax_trend = plt.subplots(figsize=(13, 5), dpi=100)

    sell_out_data = df_plot[COL_SELL_OUT].iloc[pipeline:].values
    sell_in_data = df_plot[COL_SELL_IN].iloc[:len(df_plot)-pipeline].values
    x_labels = df_plot[COL_DATA].iloc[pipeline:].values

    if len(sell_out_data) != len(sell_in_data):
         print(f"{Fore.RED}Error: Discrepancia de longitud en datos de tendencia para {marca_clean} P:{pipeline}.")
         plt.close(fig_trend)
         return

    if doble_eje:
        ax2 = ax_trend.twinx()
        lns1 = ax_trend.plot(x_labels, sell_in_data, color=COLOR_SELLIN_TREND_LINE, linewidth=4, label=f'{COL_SELL_IN} (P:{pipeline})')
        lns2 = ax2.plot(x_labels, sell_out_data, color=COLOR_SELLOUT_TREND_LINE, linewidth=4, label=COL_SELL_OUT)
        ax_trend.set_ylabel(f'{COL_SELL_IN}', color=COLOR_SELLIN_TREND_LINE, fontsize=11)
        ax2.set_ylabel(f'{COL_SELL_OUT}', color=COLOR_SELLOUT_TREND_LINE, fontsize=11)
        # --- CORRECCIÓN: Configurar ambos ejes para empezar desde 0 ---
        ax_trend.set_ylim(bottom=0)
        ax2.set_ylim(bottom=0)
        lns = lns1 + lns2
        labs = [l.get_label() for l in lns]
        ax2.legend(lns, labs, loc='lower center', bbox_to_anchor=(0.5, -0.28), frameon=False, prop={'size': 11}, ncol=2)
    else:
        lns1 = ax_trend.plot(x_labels, sell_in_data, color=COLOR_SELLIN_TREND_LINE, linewidth=4, label=f'{COL_SELL_IN} (P:{pipeline})')
        lns2 = ax_trend.plot(x_labels, sell_out_data, color=COLOR_SELLOUT_TREND_LINE, linewidth=4, label=COL_SELL_OUT)
        ax_trend.set_ylabel(f'{COL_SELL_IN} / {COL_SELL_OUT}', color='black', fontsize=11)
        ax_trend.set_ylim(bottom=0)
        lns = lns1 + lns2
        labs = [l.get_label() for l in lns]
        ax_trend.legend(lns, labs, loc='lower center', bbox_to_anchor=(0.5, -0.28), frameon=False, prop={'size': 11}, ncol=2)

    ax_trend.tick_params(axis='x', rotation=30, labelsize=9)
    for label in ax_trend.get_xticklabels():
        label.set_ha('right')
    ax_trend.grid(axis='y', linestyle='--', alpha=0.6)
    ax_trend.spines['top'].set_visible(False)
    ax_trend.spines['right'].set_visible(False)
    ax_trend.set_title(f"{labels_dict.get((lang_idx, 'Titulo Vol'), 'Tendencia en Volumen')} | {marca_clean} P:{pipeline}", size=17)

    plt.tight_layout()
    img_stream = io.BytesIO()
    fig_trend.savefig(img_stream, format='png', bbox_inches='tight', pad_inches=0.1, transparent=True)
    img_stream.seek(0)
    img_pil = Image.open(img_stream)
    bordered = ImageOps.expand(img_pil, border=2, fill='black')
    img_stream_bordered = io.BytesIO()
    bordered.save(img_stream_bordered, format='PNG')
    img_stream_bordered.seek(0)
    slide.shapes.add_picture(img_stream_bordered, Inches(0.5), Inches(1.8), height=Inches(4.5))
    plt.close(fig_trend)
    

# --- Configuración y estructuras de alto nivel --------------------------------

@dataclass
class ExecutionOptions:
    coverage_type: str
    coverage_reason: str
    trend_axis: str
    include_english: bool
    round_coverage: bool
    auto_mode: bool = False

    @classmethod
    def from_environment(cls) -> Optional["ExecutionOptions"]:
        """Crea las opciones cuando se usa la ejecución en modo automático."""
        auto_file = os.environ.get("AUTO_FILE")
        if not auto_file:
            return None
        coverage_type = os.environ.get("AUTO_COV_TYPE", "Absoluta")
        auto_mode = coverage_type.strip().lower() == "auto"
        if auto_mode:
            coverage_type = "Absoluta"
            coverage_reason = "Actualización periódica por contrato"
            trend_axis = "simple"
            include_english = False
            round_cov = False
        else:
            coverage_reason = os.environ.get("AUTO_RAZON", "Otras")
            trend_axis = os.environ.get("AUTO_EJE", "simple")
            include_english = str(os.environ.get("AUTO_ENGLISH", "0")).strip().lower() in {"1", "true", "yes", "y", "si", "sí"}
            round_cov = str(os.environ.get("AUTO_ROUND_COV", "0")).strip().lower() in {"1", "true", "yes", "y", "si", "sí"}
        return cls(
            coverage_type=coverage_type,
            coverage_reason=coverage_reason,
            trend_axis=trend_axis,
            include_english=include_english,
            round_coverage=round_cov,
            auto_mode=auto_mode,
        )


@dataclass
class PipelineAssets:
    """Recursos calculados para generar las diapositivas de un pipeline."""

    pipeline: int
    marca: str
    coverage_series: "pd.Series"
    penetration_series: "pd.Series"
    variation_table: "pd.DataFrame"
    trend_plot_df: "pd.DataFrame"
    variations_detail: Optional["pd.DataFrame"]
    evolution_figure: Optional["plt.Figure"]


    summary_rows: List[Dict[str, str]] = field(default_factory=list)
    bank_rows: List[Dict[str, object]] = field(default_factory=list)
    lang_index: int = 2

    def as_summary_df(self, labels: Dict[Tuple[int, str], List[str]]) -> "pd.DataFrame":
        return pd.DataFrame(self.summary_rows, columns=labels[(self.lang_index, "Summary")])

    summary_rows: List[Dict[str, str]] = field(default_factory=list)
    bank_rows: List[Dict[str, object]] = field(default_factory=list)

    def as_summary_df(self, labels: Dict[Tuple[int, str], List[str]]) -> "pd.DataFrame":
        return pd.DataFrame(self.summary_rows, columns=labels[(self._lang_index, "Summary")])

    def configure(self, lang_index: int) -> None:
        self._lang_index = lang_index


# --- Pequeñas utilidades -------------------------------------------------------

def compute_coverage_label(coverage_type: str, include_english: bool) -> str:
    """Devuelve el texto de cobertura a mostrar en nombres de archivo y títulos."""
    ctype = coverage_type.strip().lower()
    if ctype == "auto":
        ctype = "absoluta"
    if include_english:
        return "MOVING YEAR COVERAGE" if ctype == "absoluta" else "MOVING YEAR COVERAGE RELATIVE"
    return "Cobertura Absoluta" if ctype == "absoluta" else "Cobertura Relativa"



def determine_language(include_english: bool, pais_nombre: str) -> Tuple[str, int]:
    """Determina el código de idioma y el índice numérico usado por la lógica heredada."""
    if include_english:
        return "EN", 3
    pais_norm = (pais_nombre or "").strip().lower()
    if pais_norm in {"brasil", "brazil"}:
        return "PT", 1
    return "ES", 2


def build_labels(lang_index: int, fabricante: str, ref_month_year: str) -> Dict[Tuple[int, str], List[str] | str]:
    """Reproduce el diccionario de etiquetas usado por el script original."""
    ref_dt = dt.strptime(ref_month_year, "%m-%y")
    previous_dt = ref_dt - timedelta(days=365)
    return {
        (1, "S1"): " ",
        (1, "Summary"): [
            "Fabricante/Marca",
            "Pipeline",
            "Penetração Média Mensal",
            fabricante,
            "Worldpanel by Numerator",
            f"Cobertura {previous_dt.strftime('%b-%y')}",
            f"Cobertura {ref_dt.strftime('%b-%y')}",
            "Estabilidade",
        ],
        (1, "Graf cob Penet Men"): "Penetração Mensal",
        (1, "Titulo Cob"): "Cobertura em Ano Móvel",
        (1, "Var"): "com",
        (1, "Titulo Vol"): "Tendência em Volumen",
        (2, "S1"): " ",
        (2, "Summary"): [
            "Fabricante/Marca",
            "Pipeline",
            "Penetración Media Mensual",
            f"%VAR {fabricante}",
            "% VAR Worldpanel by Numerator",
            f"Cobertura {previous_dt.strftime('%b-%y')}",
            f"Cobertura {ref_dt.strftime('%b-%y')}",
            "Estabilidad",
        ],
        (2, "Graf cob Penet Men"): "Penetración Mensual",
        (2, "Titulo Cob"): "Cobertura en Año Móvil",
        (2, "Var"): "con",
        (2, "Titulo Vol"): "Tendencia en Volumen",
        (3, "S1"): " ",
        (3, "Summary"): [
            "Manufacturer/Brand",
            "Pipeline",
            "Monthly Avg Penetration",
            f"%VAR {fabricante}",
            "% VAR Worldpanel by Numerator",
            f"Coverage {previous_dt.strftime('%b-%y')}",
            f"Coverage {ref_dt.strftime('%b-%y')}",
            "Stability",
        ],
        (3, "Graf cob Penet Men"): "PENETRATION BY PERIOD",
        (3, "Titulo Cob"): "MOVING YEAR COVERAGE",
        (3, "Var"): "with",
        (3, "Titulo Vol"): "TREND IN VOLUME",
    }

def dataframe_to_bordered_stream(
    df: "pd.DataFrame",
    hide_index: bool = True,
    dpi: int = 220,
    styler_fn: Optional[Callable] = None,
) -> io.BytesIO:
    """Convierte un DataFrame en imagen PNG con borde negro.

    Permite aplicar personalizaciones adicionales sobre el Styler mediante ``styler_fn``.
    """
    styler = df.style.set_table_styles(
        [
            {"selector": "*", "props": [("font-size", "10pt"), ("font-family", "Calibri"), ("color", "black"), ("border-style", "solid"), ("border-width", "1px"), ("text-align", "center")]},
            {"selector": "th", "props": [("background-color", "#D9E1F2"), ("font-weight", "bold"), ("padding", "3px 5px")]},
            {"selector": "td", "props": [("padding", "2px 4px")]},
        ]
    )
    if hide_index:
        styler = styler.hide(axis="index")
    if styler_fn is not None:
        styler = styler_fn(styler)
    buffer = io.BytesIO()
    dfi.export(styler, buffer, table_conversion="matplotlib", dpi=dpi)
    buffer.seek(0)
    img = Image.open(buffer)
    bordered = ImageOps.expand(img, border=2, fill="black")
    final_stream = io.BytesIO()
    bordered.save(final_stream, format="PNG")
    final_stream.seek(0)
    return final_stream


def ensure_title_frame(slide: "Presentation"):
    """Garantiza que el slide tenga un cuadro de título y devuelve su text_frame."""
    placeholder = slide.shapes.title
    if placeholder is not None:
        return placeholder.text_frame
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
    return textbox.text_frame


class SlideBuilder:
    """Encapsula la lógica de creación de slides para mantener el código ordenado."""

    def __init__(
        self,
        presentation: "Presentation",
        lang_index: int,
        labels: Dict[Tuple[int, str], List[str] | str],
        coverage_label: str,
        tipo_eje_tend: str,
    ) -> None:
        self.ppt = presentation
        self.lang_index = lang_index
        self.labels = labels
        self.coverage_label = coverage_label
        self.tipo_eje_tend = tipo_eje_tend

    # --- Portada -----------------------------------------------------------------
    def configure_cover(self, pais_nombre: str, fabricante: str, categoria_nombre: str, ref_month_year: str, chosen_lang: str) -> None:
        cover_slide = self.ppt.slides[0]
        line1 = f"{pais_nombre} | {fabricante}"
        try:
            ref_dt = dt.strptime(ref_month_year, "%m-%y")
            meses_es = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
            meses_pt = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            meses_en = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
            if chosen_lang == "PT":
                month_name = meses_pt[ref_dt.month].capitalize()
                line2 = f"{categoria_nombre} - Corte em {month_name} {ref_dt.year}"
            elif chosen_lang == "EN":
                month_name = meses_en[ref_dt.month]
                line2 = f"{categoria_nombre} - As of {month_name} {ref_dt.year}"
            else:
                month_name = meses_es[ref_dt.month].capitalize()
                line2 = f"{categoria_nombre} - Corte a {month_name} {ref_dt.year}"
        except Exception:
            if chosen_lang == "PT":
                line2 = f"{categoria_nombre} - Corte em {ref_month_year}"
            elif chosen_lang == "EN":
                line2 = f"{categoria_nombre} - As of {ref_month_year}"
            else:
                line2 = f"{categoria_nombre} - Corte a {ref_month_year}"
        textbox = cover_slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(9), Inches(2.5))
        text_frame = textbox.text_frame
        text_frame.clear()
        p1 = text_frame.add_paragraph()
        p1.text = line1
        p1.font.size = Pt(44)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(255, 255, 255)
        p1.alignment = 1
        p2 = text_frame.add_paragraph()
        p2.text = line2
        p2.font.size = Pt(36)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(255, 255, 255)
        p2.alignment = 1

    # --- Pipelines ---------------------------------------------------------------
    def add_pipeline_slides(
        self,
        assets: PipelineAssets,
        marca_nombre_limpio: str,
        lang_index: int,
        coverage_label: str,
        progress: Optional["Progress"] = None,
        task_id: Optional[int] = None,
    ) -> int:
        slides_created = 0
        slide_cov = self.ppt.slides.add_slide(self.ppt.slide_layouts[PPT_LAYOUT_INDEX])
        tx_title_cov = ensure_title_frame(slide_cov)
        p_cov = tx_title_cov.paragraphs[0]
        p_cov.text = f"{marca_nombre_limpio} - Pipeline {assets.pipeline}"
        p_cov.font.bold = True
        p_cov.font.size = Pt(24)
        generar_grafico_cobertura(
            slide_cov,
            marca_nombre_limpio,
            assets.pipeline,
            assets.coverage_series,
            assets.penetration_series,
            lang_index,
            coverage_label,
            self.labels,
        )
        try:
            table_stream = dataframe_to_bordered_stream(assets.variation_table, hide_index=True, dpi=200)
            slide_cov.shapes.add_picture(table_stream, Inches(0.5), Inches(1.1), height=Inches(0.6))
        except Exception as exc:
            print(f"{Fore.YELLOW}Advertencia: No se pudo generar la tabla de variación MAT para {marca_nombre_limpio} P{assets.pipeline}. Error: {exc}")
        slides_created += 1
        if progress and task_id is not None:
            progress.update(task_id, advance=1)
        slide_trend = self.ppt.slides.add_slide(self.ppt.slide_layouts[PPT_LAYOUT_INDEX])
        tx_title_trend = ensure_title_frame(slide_trend)
        p_trend = tx_title_trend.paragraphs[0]
        p_trend.text = f"{marca_nombre_limpio} - Pipeline {assets.pipeline}"
        p_trend.font.bold = True
        p_trend.font.size = Pt(24)
        generar_grafico_tendencia(
            slide_trend,
            marca_nombre_limpio,
            assets.pipeline,
            assets.trend_plot_df,
            lang_index,
            self.labels,
            doble_eje=(self.tipo_eje_tend == "doble"),
        )
        if assets.variations_detail is not None and not assets.variations_detail.empty:
            value_columns = [col for col in assets.variations_detail.columns if col not in {'Tipo', 'Periodo'}]

            def _variation_styler(styler):
                formatter = {}
                for col in value_columns:
                    formatter[col] = lambda v, _col=col: "-" if (pd.isna(v) or (isinstance(v, str) and str(v).strip() == "-")) else f"{v*100:.1f}%"
                if formatter:
                    styler = styler.format(formatter)

                    def _colorize(val):
                        if isinstance(val, str):
                            try:
                                numeric_val = val.replace('%', '').replace(',', '.')
                                val = float(numeric_val) / 100 if '%' in val else float(numeric_val)
                            except ValueError:
                                return ""
                        if pd.isna(val):
                            return ""
                        if val > 0:
                            return "background-color: #C6EFCE; color: #006100"
                        if val < 0:
                            return "background-color: #FFC7CE; color: #9C0006"
                        return ""

                    styler = styler.applymap(_colorize, subset=value_columns)
                text_columns = [col for col in ('Tipo', 'Periodo') if col in assets.variations_detail.columns]
                if text_columns:
                    styler = styler.set_properties(subset=text_columns, **{"text-align": "left"})
                return styler

            table_stream = dataframe_to_bordered_stream(
                assets.variations_detail,
                hide_index=True,
                dpi=200,
                styler_fn=_variation_styler,
            )
            table_stream.seek(0)
            scale_factor = .6
            try:
                with Image.open(table_stream) as img_preview:
                    base_width_in = img_preview.width / 200
                    base_height_in = img_preview.height / 200
            except Exception:
                base_width_in = 3.4
                base_height_in = base_width_in * 0.75
            finally:
                table_stream.seek(0)

            table_width = Inches(base_width_in * scale_factor)
            table_height = Inches(base_height_in * scale_factor)
            right_margin = Inches(0.3)
            left_pos = self.ppt.slide_width - table_width - right_margin
            if left_pos < Inches(0.1):
                left_pos = Inches(0.1)
            top_pos = Inches(0.4)
            slide_trend.shapes.add_picture(table_stream, left_pos, top_pos, width=table_width, height=table_height)
        slides_created += 1
        if progress and task_id is not None:
            progress.update(task_id, advance=1)
        if assets.evolution_figure is not None:
            slide_evol = self.ppt.slides.add_slide(self.ppt.slide_layouts[PPT_LAYOUT_INDEX])
            tx_title_evol = ensure_title_frame(slide_evol)
            p_evol = tx_title_evol.paragraphs[0]
            if lang_index == 3:
                p_evol.text = f"{marca_nombre_limpio} - Pipeline {assets.pipeline} - Monthly Evolution and YoY Variation"
            elif lang_index == 1:
                p_evol.text = f"{marca_nombre_limpio} - Pipeline {assets.pipeline} - Evolução Mensal e Variação"
            else:
                p_evol.text = f"{marca_nombre_limpio} - Pipeline {assets.pipeline} - Evolución Mensual y Variación"
            p_evol.font.bold = True
            p_evol.font.size = Pt(24)
            buffer = io.BytesIO()
            assets.evolution_figure.savefig(buffer, format="png", dpi=240, bbox_inches="tight", pad_inches=0.08, transparent=True)
            plt.close(assets.evolution_figure)
            buffer.seek(0)
            left = Inches(0.1)
            usable_w = self.ppt.slide_width - 2 * left
            slide_evol.shapes.add_picture(buffer, left, Inches(1.0), width=usable_w)
            slides_created += 1
            if progress and task_id is not None:
                progress.update(task_id, advance=1)
        return slides_created

    # --- Resumen -----------------------------------------------------------------
    def add_summary_slide(
        self,
        df_summary: "pd.DataFrame",
        pais_nombre: str,
        categoria_nombre: str,
    ) -> None:
        slide_summary = self.ppt.slides.add_slide(self.ppt.slide_layouts[PPT_LAYOUT_INDEX])
        title_frame = ensure_title_frame(slide_summary)
        p = title_frame.paragraphs[0]
        p.text = f"Summary - {pais_nombre} {categoria_nombre} - {self.coverage_label}"
        p.font.bold = True
        p.font.size = Pt(26)
        tx_s1 = slide_summary.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.5))
        s1_frame = tx_s1.text_frame
        s1_frame.text = self.labels.get((self.lang_index, "S1"), "")
        comentarios_box = slide_summary.shapes.add_textbox(Inches(0.5), Inches(6.0), Inches(8.5), Inches(0.7))
        comentarios_frame = comentarios_box.text_frame
        comentarios_frame.word_wrap = True
        comentarios_frame.auto_size = True
        comentarios_frame.text = "Comentarios:"
        if df_summary.empty:
            print(f"{Fore.YELLOW}Advertencia: No hay datos para generar la tabla de resumen en el PPT.")
            return
        try:
            summary_stream = dataframe_to_bordered_stream(df_summary, hide_index=True, dpi=250)
            left = Inches(0.5)
            top = Inches(1.0)
            usable_w = self.ppt.slide_width - 2 * left
            final_left = int((self.ppt.slide_width - usable_w) // 2)
            slide_summary.shapes.add_picture(summary_stream, final_left, top, width=usable_w)
        except Exception as exc:
            print(f"{Fore.YELLOW}Advertencia: No se pudo generar la tabla resumen en el PPT. Error: {exc}")

    # --- Post-procesamiento -------------------------------------------------------
    def insert_thanks_text(self, chosen_lang: str) -> None:
        thanks_map = {"ES": "Gracias", "PT": "Obrigado(a)", "EN": "Thank you"}
        thanks_txt = thanks_map.get(chosen_lang, "Gracias")
        if len(self.ppt.slides) <= 6:
            return
        slide7 = self.ppt.slides[6]
        tb = slide7.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(9), Inches(2.5))
        tf7 = tb.text_frame
        tf7.clear()
        p = tf7.add_paragraph()
        p.text = thanks_txt
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = 1

    def reorder_summary_and_credit(self) -> None:
        if len(self.ppt.slides) > 1:
            summary_slide_xml = self.ppt.slides._sldIdLst[-1]
            insert_idx = 7 if len(self.ppt.slides) > 7 else len(self.ppt.slides) - 1
            self.ppt.slides._sldIdLst.insert(insert_idx, summary_slide_xml)
        if len(self.ppt.slides) > 7:
            credit_slide_xml = self.ppt.slides._sldIdLst[6]
            self.ppt.slides._sldIdLst.append(credit_slide_xml)


def parse_file_metadata(excel_file_name: str, categories_df: "pd.DataFrame") -> Tuple[str, str, str, str, str]:
    """Obtiene país, cesta y categoría a partir del nombre del archivo."""
    parts = os.path.splitext(excel_file_name)[0].split('_')
    if len(parts) < 3:
        raise ValueError("El nombre de archivo no contiene suficientes partes (país_categoria_fabricante)")
    country_code_str, category_code, fabricante = parts[:3]
    try:
        country_code = int(country_code_str)
    except ValueError as exc:
        raise ValueError(f"El código de país '{country_code_str}' no es numérico") from exc
    try:
        pais_nombre = str(pais.loc[pais.cod == country_code, 'pais'].iloc[0]).strip()
    except Exception as exc:
        raise ValueError(f"No se encontró el país para el código {country_code_str}") from exc
    if category_code not in categories_df.index:
        raise ValueError(f"El código de categoría '{category_code}' no está en el catálogo")
    cesta_nombre = categories_df.loc[category_code, 'cest']
    categoria_nombre = categories_df.loc[category_code, 'cat']
    try:
        dash_split = re.split(r"\s*[-‑–—−‒]\s*", str(categoria_nombre), maxsplit=1)
        categoria_corta = dash_split[0].strip() if dash_split else str(categoria_nombre).strip()
        if not categoria_corta:
            categoria_corta = str(categoria_nombre).strip()
    except Exception:
        categoria_corta = str(categoria_nombre).strip()
    return pais_nombre, cesta_nombre, categoria_nombre, categoria_corta, fabricante


def ensure_output_folder(root_dir: str, nombre_base_archivo: str) -> str:
    carpeta_salida = os.path.join(root_dir, nombre_base_archivo)
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida, exist_ok=True)
    return carpeta_salida



def copy_and_prune_template(root_dir: str, chosen_lang: str) -> Tuple["Presentation", str]:
    """Copia la plantilla base, elimina slides según idioma y devuelve la presentación lista."""
    run_id = os.environ.get('RUN_ID') or datetime.now().strftime('%Y%m%d_%H%M%S')
    tmp_dir = os.path.join(root_dir, 'tmp')
    os.makedirs(tmp_dir, exist_ok=True)
    src_template_path = os.path.join(root_dir, 'Modelo_PPT.pptx')
    if not os.path.exists(src_template_path):
        raise FileNotFoundError(f"No se encontró la plantilla base: {src_template_path}")
    tmp_ppt_name = f"Modelo_PPT_{run_id}_{chosen_lang}.pptx"
    tmp_ppt_path = os.path.join(tmp_dir, tmp_ppt_name)
    shutil.copyfile(src_template_path, tmp_ppt_path)
    ppt = Presentation(tmp_ppt_path)
    keep_indices_by_lang = {
        'ES': {0, 1, 2, 3, 4, 5, 16},
        'PT': {0, 6, 7, 8, 9, 10, 16},
        'EN': {0, 11, 12, 13, 14, 15, 16},
    }
    keep_set = keep_indices_by_lang.get(chosen_lang, keep_indices_by_lang['ES'])
    total_initial = len(ppt.slides)
    delete_list = sorted([i for i in range(total_initial) if i not in keep_set], reverse=True)
    for di in delete_list:
        _delete_slide(ppt, di)
    ppt.save(tmp_ppt_path)
    return Presentation(tmp_ppt_path), tmp_ppt_path


def _delete_slide(pres_obj: "Presentation", idx: int) -> None:
    """Elimina un slide usando la API protegida de python-pptx."""
    sldIdLst = pres_obj.slides._sldIdLst  # type: ignore[attr-defined]
    sldId = sldIdLst[idx]
    rId = sldId.rId
    pres_obj.part.drop_rel(rId)
    sldIdLst.remove(sldId)


def generate_excel_template(
    root_dir: str,
    excel_file_obj: 'pd.ExcelFile',
    marcas: Sequence[str],
    pais_nombre: str,
    categoria_nombre: str,
    categoria_nombre_corto: str,
    fabricante: str,
    coverage_label: str,
    coverage_type: str,
    coverage_reason: str
) -> Tuple[str, str, str, str]:
    """Genera el archivo Excel temporal y devuelve datos clave."""
    print(Fore.CYAN + "\nGenerando archivo Excel temporal...")
    excel_temp_path = os.path.join(root_dir, EXCEL_TEMP_FILENAME)
    try:
        with pd.ExcelWriter(excel_temp_path) as writer:
            # Recorrer cada hoja (marca) del archivo
            for marca_sheet_name in tqdm(marcas, desc="Procesando Hojas Excel"):

                # 1.1) Carga y preprocesa la hoja usando la función refactorizada
                df_marca, measure_unit = load_and_preprocess_sheet(excel_file_obj, marca_sheet_name)

                # Si la carga falló, continuar con la siguiente hoja
                if df_marca is None:
                    continue

                # Guardar número original de filas de datos para fórmulas Excel
                original_data_rows = len(df_marca)
                if original_data_rows < 12:
                    print(f"{Fore.YELLOW}Advertencia: Hoja '{marca_sheet_name}' tiene < 12 meses de datos ({original_data_rows}). Algunos cálculos de Excel pueden fallar o dar NaN.")
                    # Continuar de todos modos, pero con precaución

                # Actualizar fecha de referencia global (usará la de la última hoja procesada con éxito)
                ref_month_year = df_marca[COL_DATA].iloc[-1].strftime('%m-%y')

                # --- 1.5) Creación de columnas con fórmulas Excel ---
                df_excel = df_marca.copy() # Trabajar sobre una copia para Excel
                # Hacer los índices basados en 1 y añadir offset de header (fila 1)
                excel_row_offset = 2

                # Sell_in_sim (Ejemplo - ajustable manualmente en Excel si se necesita)
                # La fórmula asume que Sell_in está en la columna B
                df_excel[COL_SELL_IN_SIM] = [f"=B{r}" for r in range(excel_row_offset, original_data_rows + excel_row_offset)] + [np.nan] * (len(df_excel) - original_data_rows)

                # Acumulados (MAT - Moving Annual Total) - comienzan desde la fila 12 de datos
                # Las fórmulas asumen Sell_out en C y Sell_in_sim en L
                for i in range(11, original_data_rows):
                    row_excel = i + excel_row_offset
                    df_excel.loc[i, COL_ACUM_SELL_OUT] = f"=SUM(C{row_excel - 11}:C{row_excel})"
                    df_excel.loc[i, COL_ACUM_SELL_IN] = f"=SUM(L{row_excel - 11}:L{row_excel})" # Usa Sell_in_sim (L)

                # --- 1.6) Cálculo de coberturas (pipeline 0 a 6) en Excel ---
                pop_value_str = pop_coverage.get(pais_nombre, DEFAULT_POP_COVERAGE)
                cov_formulas_list = []
                max_rows_excel = original_data_rows + excel_row_offset -1 # Última fila con datos en Excel

                for r_idx in range(original_data_rows): # Iterar sobre índices de datos (0 a N-1)
                    excel_current_row = r_idx + excel_row_offset
                    row_formulas = {}
                    if r_idx >= 11: # Cobertura solo se calcula desde el mes 12
                        for p in range(7): # Pipelines P0 a P6
                            # Fila de Excel para el numerador (Acum_Sell_in) - con pipeline
                            num_row_excel = excel_current_row + p
                            # Fila de Excel para el denominador (Acum_Sell_out) - sin pipeline
                            den_row_excel = excel_current_row

                            # Verificar que las filas referenciadas existan
                            if num_row_excel <= max_rows_excel and den_row_excel <= max_rows_excel:
                                # La fórmula asume Acum_Sell_in en N y Acum_Sell_out en M
                                #anterior  m{den_row_excel}/n{num_row_excel}*100
                                base_formula = f"M{num_row_excel}/N{den_row_excel}*100"
                                if coverage_type == "relativa":
                                    # CORRECCIÓN: Cambiar formato de porcentaje y usar NA()
                                    pop_value_decimal = float(pop_value_str.replace("%", "")) / 100
                                    formula = f"=IFERROR(({base_formula})/{pop_value_decimal},NA())"
                                else:
                                    formula = f"=IFERROR({base_formula},NA())"
                                row_formulas[f'P{p}'] = formula
                            else:
                                 row_formulas[f'P{p}'] = np.nan # O "" o NA()
                    else:
                        # Rellenar con NaN para las primeras 11 filas
                        for p in range(7):
                             row_formulas[f'P{p}'] = np.nan

                    cov_formulas_list.append(row_formulas)

                df_cov_excel = pd.DataFrame(cov_formulas_list, index=df_excel.index[:original_data_rows])

                # Escalonar las columnas de cobertura
                df_cov_excel_scaled = df_cov_excel.copy()
                escalona(df_cov_excel_scaled) # Escalonar la copia




    # -------------------------------------------------------
                # 1.7 & 1.8) Cálculo de variaciones (Y-1 e Y-2) en Excel
                # -------------------------------------------------------

                # ► VARIABLES EXTRA que tu código “heredado” sigue ocupando
                n_data          = original_data_rows                            # filas con datos
                last_row_excel  = n_data + excel_row_offset - 1                 # última fila real en Excel

                # ---------- Y-1 -------------------------------------------------
                var = pd.DataFrame([
                    ['Anual',      "MAT " + df_excel.loc[original_data_rows-1, COL_DATA].strftime('%b-%y') +
                                " x MAT " + df_excel.loc[original_data_rows-1-12, COL_DATA].strftime('%b-%y')],
                    ['Semestral',  "SEM " + df_excel.loc[original_data_rows-1, COL_DATA].strftime('%b-%y') +
                                " x SEM " + df_excel.loc[original_data_rows-1-6,  COL_DATA].strftime('%b-%y')],
                    ['Trimestral', "TRI " + df_excel.loc[original_data_rows-1, COL_DATA].strftime('%b-%y') +
                                " x TRI " + df_excel.loc[original_data_rows-1-3,  COL_DATA].strftime('%b-%y')]
                ], columns=['Tipo', 'Periodo'])

                # Variaciones WP by Numerator
                var['WP by Numerator'] = [
                    f"=SUM(C{original_data_rows+excel_row_offset-i - 2}:C{original_data_rows+excel_row_offset-1})/"
                    f"SUM(C{original_data_rows+excel_row_offset-2*j -2 }:C{original_data_rows+excel_row_offset-j - 2})-1"
                    for i, j in zip([10, 4, 1], [11, 5, 2])
                ]

                # Variaciones Cliente
                for p in range(7):
                    var[f'Cliente P{p}'] = [
                        f"=SUM(L{original_data_rows+excel_row_offset-i-p -2}:L{original_data_rows+excel_row_offset-p -1})/"
                        f"SUM(L{original_data_rows+excel_row_offset-2*j-p -2}:L{original_data_rows+excel_row_offset-j-p -2})-1"
                        for i, j in zip([10, 4, 1], [11, 5, 2])
                    ]

                # ---------- Y-2 -------------------------------------------------
                # Ventanas: MAT=12, SEM=6, TRI=3  (todas comparadas contra el mismo tamaño W, 24 meses antes)
                periods = [
                    ('Anual',      12,  24),   # (nombre, meses_ventana, lag_meses)
                    ('Semestral',   6,  24),
                    ('Trimestral',  3,  24),
                ]

                # Texto de periodo (tu formato actual)
                aux = pd.DataFrame([
                    ['Anual',      "MAT " + df_excel.loc[original_data_rows-1,     COL_DATA].strftime('%b-%y') +
                                " x MAT " + df_excel.loc[original_data_rows-1-24, COL_DATA].strftime('%b-%y')],
                    ['Semestral',  "SEM " + df_excel.loc[original_data_rows-1,     COL_DATA].strftime('%b-%y') +
                                " x SEM " + df_excel.loc[original_data_rows-1-24, COL_DATA].strftime('%b-%y')],
                    ['Trimestral', "TRI " + df_excel.loc[original_data_rows-1,     COL_DATA].strftime('%b-%y') +
                                " x TRI " + df_excel.loc[original_data_rows-1-24, COL_DATA].strftime('%b-%y')]
                ], columns=['Tipo', 'Periodo'])

                def rango_excel(end_row: int, meses: int) -> tuple[int, int]:
                    """Devuelve (inicio, fin) inclusivo para una ventana de 'meses' que termina en 'end_row'."""
                    return end_row - (meses - 1), end_row

                def formula_yoy_excel(col: str, end_row: int, meses: int, lag_meses: int) -> str:
                    """
                    = SUM( col[num_ini:num_fin] ) / SUM( col[den_ini:den_fin] ) - 1
                    donde el denominador termina en end_row - lag_meses y tiene el mismo tamaño 'meses'.
                    """
                    # Numerador: ventana actual (tamaño 'meses') que termina en end_row
                    num_ini, num_fin = rango_excel(end_row, meses)
                    # Denominador: misma ventana 'meses', pero que termina 'lag_meses' antes
                    den_fin = end_row - lag_meses
                    den_ini, den_fin = rango_excel(den_fin, meses)
                    return f"=SUM({col}{num_ini}:{col}{num_fin})/SUM({col}{den_ini}:{col}{den_fin})-1"

                # Reglas de suficiencia de datos por ventana para Y-2:
                #  - MAT (12): requiere >= 12 + 24 = 36 meses
                #  - SEM (6):  requiere >= 6  + 24 = 30 meses
                #  - TRI (3):  requiere >= 3  + 24 = 27 meses

                # ► WP by Numerator (columna C)
                wp_y2_formulas = []
                for _, meses, lag in periods:
                    required = meses + lag
                    if n_data >= required:
                        wp_y2_formulas.append(formula_yoy_excel("C", last_row_excel, meses, lag))
                    else:
                        wp_y2_formulas.append("-")
                aux['WP by Numerator'] = wp_y2_formulas

                # ► Clientes P0..P6 (columna L), ajustando el fin por 'p'
                for p in range(7):
                    end_row_p = last_row_excel - p
                    cli_y2 = []
                    for _, meses, lag in periods:
                        required = meses + lag
                        # Suficiencia: descontamos 'p' del total disponible para ese cliente
                        if (n_data - p) >= required:
                            cli_y2.append(formula_yoy_excel("L", end_row_p, meses, lag))
                        else:
                            cli_y2.append("-")
                    aux[f'Cliente P{p}'] = cli_y2

                # Limpiar variaciones sin sentido
                if 42 - original_data_rows >= 0:
                    for i in range(abs(42 - original_data_rows)):
                        aux.loc[0, f'Cliente P{6 - i}'] = np.nan


                # ---------- Unir Y-1 y Y-2 --------------------------------------
                df_variations_excel = pd.concat([var, aux], ignore_index=True)



                # --- 1.9) Cálculo de correlaciones en Excel (MAT) ---
                # Se genera un diccionario con fórmulas de correlación para cada pipeline (P0 a P6)
                # Se construyen fórmulas Excel que calculan la correlación Pearson entre dos rangos de 12 filas:
                #   uno en la columna M y otro en la columna N, considerando el desplazamiento (pipeline).
                # Los índices son base 1 y se garantiza que cada rango tenga exactamente 12 filas; de lo contrario, se asigna '-'.
            
                # ---------- Correlaciones: 12m, 2 años antes (12m terminando hace 24m), 2 años (ventana 24m) ----------

                def _build_correl_row(label: str, window: int, end_offset: int = 0) -> dict:
                    """
                    Genera una fila de correlaciones entre M y N para:
                    - window: tamaño de ventana (12 o 24)
                    - end_offset: 0 = ventana termina en last_row_excel (reciente)
                                    24 = ventana termina 24 meses antes (para '2 años antes')
                    N se alinea con M desplazando p filas hacia arriba (n_start = m_start - p).
                    Si no hay suficientes datos para esa p y esa ventana, devuelve '-'.
                    """
                    row = {'Correlacion': label}

                    # ¿Hay datos suficientes para esta ventana y desplazamiento?
                    if n_data >= window + end_offset:
                        # Ventana base en M
                        row_ini = last_row_excel - end_offset - (window - 1)
                        row_fin = last_row_excel - end_offset

                        # Respetar que la fila 1 es encabezado
                        m_start = max(row_ini, 2)
                        m_end   = max(row_fin, 2)

                        for p in range(0, 7):  # P0..P6
                            n_start = max(row_ini - p, 2)
                            n_end   = max(row_fin - p, 2)

                            # Ambas ventanas deben tener exactamente 'window' filas
                            if (m_end - m_start + 1 == window) and (n_end - n_start + 1 == window):
                                # Usa coma ',' en argumentos; función en inglés 'CORREL' como en tu flujo actual
                                row[f'P{p}'] = f"=CORREL(M{m_start}:M{m_end},N{n_start}:N{n_end})"
                            else:
                                row[f'P{p}'] = '-'
                    else:
                        for p in range(0, 7):
                            row[f'P{p}'] = '-'

                    return row

                # Construye las 3 filas en el orden solicitado
                rows = [
                    _build_correl_row('Año Actual', 12, end_offset=0),                   # últimos 12 meses
                    _build_correl_row('1 año antes', 12, end_offset=12),           # 12 meses que terminaron hace 12 meses (Año anterior)
                    _build_correl_row('2 años (ventana de 24 meses)', 24, 0),       # últimos 24 meses
                ]

                # Ordenar columnas: Correlacion, P0..P6 (incluye P6)
                cols = ['Correlacion'] + [f'P{i}' for i in range(7)]
                df_correlations_excel = pd.DataFrame(rows)[cols]




                # --- 1.10) Promedio de Penetración y Buyers (MAT) en Excel ---
                avg_formulas = []
                # MAT Actual
                if n_data >= 12:
                     start_avg_curr = last_row_excel - 11
                     end_avg_curr = last_row_excel
                     # Asume Penet en G, Buyers en H
                     avg_formulas.append({'Media': 'Penet MAT Actual', 'Valor': f"=AVERAGE(G{start_avg_curr}:G{end_avg_curr})"})
                     avg_formulas.append({'Media': 'Buyers MAT Actual', 'Valor': f"=AVERAGE(H{start_avg_curr}:H{end_avg_curr})"})
                else:
                     avg_formulas.append({'Media': 'Penet MAT Actual', 'Valor': f"=AVERAGE(G{excel_row_offset}:G{last_row_excel})"}) # Promedio de lo disponible
                     avg_formulas.append({'Media': 'Buyers MAT Actual', 'Valor': f"=AVERAGE(H{excel_row_offset}:H{last_row_excel})"})

                # MAT Anterior
                if n_data >= 24:
                     start_avg_prev = last_row_excel - 23
                     end_avg_prev = last_row_excel - 12
                     avg_formulas.append({'Media': 'Penet MAT Anterior', 'Valor': f"=AVERAGE(G{start_avg_prev}:G{end_avg_prev})"})
                     avg_formulas.append({'Media': 'Buyers MAT Anterior', 'Valor': f"=AVERAGE(H{start_avg_prev}:H{end_avg_prev})"})
                else:
                     avg_formulas.append({'Media': 'Penet MAT Anterior', 'Valor': np.nan}) # O NA()
                     avg_formulas.append({'Media': 'Buyers MAT Anterior', 'Valor': np.nan})

                df_averages_excel = pd.DataFrame(avg_formulas)


                # --- 1.11) Calcular Estabilidad en Excel ---
                # Diferencia entre última cobertura y cobertura de hace 12 meses
                estab_data = {"Estabilidad": "Estabilidad"}
                # Asume Cobertura P0-P6 en columnas O a U (después de escalonar)
                coverage_start_col_letter = 'O'
                coverage_start_col_idx = 15 # Col O es la 15

                last_data_row_idx = original_data_rows -1 # Índice base 0

                for p in range(7):
                     col_letter = get_column_letter(coverage_start_col_idx + p)
                     row_last_cov = last_row_excel - p
                     row_prev_cov = row_last_cov - 12

                     # Verificar si las filas son válidas y si hay suficientes datos
                     if row_last_cov >= excel_row_offset and row_prev_cov >= excel_row_offset and (original_data_rows >= 23+p):
                         # CORRECCIÓN: Usar IFERROR y NA()
                         formula = f"=IFERROR({col_letter}{row_last_cov}-{col_letter}{row_prev_cov},NA())"
                         estab_data[f'P{p}'] = formula
                     else:
                         estab_data[f'P{p}'] = np.nan
            
                # Crear DataFrame para estabilidad
                df_stability_excel = pd.DataFrame([estab_data])

                # --- 1.12) Ensamblar DataFrame final para Excel ---
                # Unir datos originales con coberturas escalonadas
                df_excel_final = pd.concat([df_excel, df_cov_excel_scaled], axis=1)

                # Crear la sección de resumen (Variaciones, Promedios, Correlación, Estabilidad)
                # Añadir filas vacías y reorganizar
                df_variations_excel['spacer1'] = np.nan
                # df_averages_excel['spacer2'] = np.nan
                df_correlations_excel['spacer3'] = np.nan

                # Aplanar las tablas de resumen para concatenarlas horizontalmente
                summary_part1 = df_variations_excel.T.reset_index().T # Variaciones
                summary_part2 = df_averages_excel.T.reset_index().T   # Promedios
                summary_part3 = df_correlations_excel.T.reset_index().T # Correlaciones
                summary_part4 = df_stability_excel.T.reset_index().T  # Estabilidad

                # Crear un DataFrame vacío con el número correcto de columnas para alinear
                max_cols = df_excel_final.shape[1]
                summary_placeholder = pd.DataFrame(np.nan, index=range(max(len(summary_part1), len(summary_part2), len(summary_part3), len(summary_part4))), columns=df_excel_final.columns)

                # Rellenar el placeholder (esto requiere manejo cuidadoso de índices y columnas)
                # Simplificación: Crear el df_excel_summary_part como antes y concatenar al final
                df_excel_summary_part = pd.concat([df_variations_excel.reset_index(drop=True),
                                                  df_averages_excel.reset_index(drop=True),
                                                  df_correlations_excel.reset_index(drop=True),
                                                  df_stability_excel.reset_index(drop=True)], axis=1)

                # Añadir fila vacía de separación
                df_excel_final.loc[len(df_excel_final)] = [np.nan] * len(df_excel_final.columns)

                # Añadir nombres de columnas del resumen como cabecera
                summary_header = pd.DataFrame([df_excel_summary_part.columns], columns=df_excel_summary_part.columns)
                df_excel_summary_part_with_header = pd.concat([summary_header, df_excel_summary_part], ignore_index=True)

                # Ajustar columnas del resumen para que coincidan con el df principal y concatenar
                # --- INICIO CAMBIO ---
                # Si el número de columnas no coincide, agrega columnas vacías
                n_main_cols = df_excel_final.shape[1]
                n_summary_cols = df_excel_summary_part_with_header.shape[1]
                if n_summary_cols < n_main_cols:
                    # Agrega columnas vacías al resumen
                    for i in range(n_summary_cols, n_main_cols):
                        df_excel_summary_part_with_header[f'empty_{i}'] = np.nan
                elif n_summary_cols > n_main_cols:
                    # Si el resumen tiene más columnas, recórtalas
                    df_excel_summary_part_with_header = df_excel_summary_part_with_header.iloc[:, :n_main_cols]
                # Ahora reasigna los nombres de columnas
                df_excel_summary_part_with_header.columns = df_excel_final.columns
                # --- FIN CAMBIO ---

                df_excel_final = pd.concat([df_excel_final, df_excel_summary_part_with_header], ignore_index=True)

                # --- 1.13) Exportar a la hoja de Excel ---
                df_excel_final.to_excel(writer, sheet_name=marca_sheet_name, index=False)

        print(Fore.GREEN + f"Archivo Excel temporal '{EXCEL_TEMP_FILENAME}' generado.")

        # Aplicar formato de color y porcentaje a la sección de Correlaciones (como en excel_color.py)
        try:
            def apply_correlation_formatting(xlsx_path: str) -> None:
                from openpyxl import load_workbook as _load_wb
                from openpyxl.formatting.rule import ColorScaleRule as _ColorScaleRule
                from openpyxl.utils import get_column_letter as _get_col_letter
                wb = _load_wb(xlsx_path)
                for ws in wb.worksheets:
                    found = False
                    for row in ws.iter_rows(values_only=False):
                        for cell in row:
                            if isinstance(cell.value, str) and str(cell.value).strip().lower() == 'correlacion':
                                header_row = cell.row
                                header_col = cell.column
                                start_col = header_col + 1  # P0
                                end_col = start_col + 6     # P6

                                # Detectar largo dinámico: desde la fila siguiente hasta que la fila esté vacía en P0..P6
                                r = header_row + 1
                                last_row = r - 1
                                while True:
                                    vals = [ws.cell(row=r, column=c).value for c in range(start_col, end_col + 1)]
                                    if all(v is None for v in vals):
                                        break
                                    last_row = r
                                    r += 1

                                if last_row >= header_row + 1:
                                    # Formato de porcentaje 0.0% en P0..P6
                                    for rr in range(header_row + 1, last_row + 1):
                                        for cc in range(start_col, end_col + 1):
                                            ws.cell(row=rr, column=cc).number_format = '0.0%'

                                    # Regla de escala de color 3-colores (rojo-amarillo-verde)
                                    rng = f"{_get_col_letter(start_col)}{header_row + 1}:{_get_col_letter(end_col)}{last_row}"
                                    color_scale = _ColorScaleRule(
                                        start_type='min', start_color='F8696B',
                                        mid_type='percentile', mid_value=50, mid_color='FFEB84',
                                        end_type='max', end_color='63BE7B'
                                    )
                                    ws.conditional_formatting.add(rng, color_scale)
                                found = True
                                break
                        if found:
                            break
                wb.save(xlsx_path)

            apply_correlation_formatting(excel_temp_path)
            print(Fore.GREEN + "Formato de correlaciones aplicado (colores y porcentaje).")

            # Variaciones: formato porcentaje y reglas de color rojo(<0)/verde(>0)
            def apply_variations_formatting(xlsx_path: str) -> None:
                from openpyxl import load_workbook as _load_wb2
                from openpyxl.utils import get_column_letter as _col_letter
                from openpyxl.formatting.rule import Rule as _Rule
                from openpyxl.styles import PatternFill as _PatternFill, Font as _Font
                from openpyxl.styles.differential import DifferentialStyle as _Diff
                wb2 = _load_wb2(xlsx_path)
                for ws in wb2.worksheets:
                    header_row = None
                    wp_col = None
                    # Buscar el encabezado 'WP by Numerator'
                    for row in ws.iter_rows(values_only=False):
                        for cell in row:
                            if isinstance(cell.value, str) and str(cell.value).strip().lower() == 'wp by numerator':
                                header_row = cell.row
                                wp_col = cell.column
                                break
                        if header_row:
                            break
                    if not header_row:
                        continue
                    # Detectar columnas de Cliente P0..P6 consecutivas hacia la derecha
                    end_col = wp_col
                    p = 0
                    while True:
                        header_cell = ws.cell(row=header_row, column=wp_col + 1 + p)
                        val = header_cell.value
                        if isinstance(val, str) and val.strip().lower() == f'cliente p{p}':
                            end_col = wp_col + 1 + p
                            p += 1
                            if p > 20:  # seguridad
                                break
                        else:
                            break
                    # Si no se detectaron clientes, por defecto tomar WP + 7 clientes
                    if end_col == wp_col:
                        end_col = wp_col + 7
                    # Determinar rango de filas con datos (hasta que todas las columnas estén vacías)
                    r = header_row + 1
                    last_row = r - 1
                    while True:
                        vals = [ws.cell(row=r, column=c).value for c in range(wp_col, end_col + 1)]
                        if all(v is None for v in vals):
                            break
                        last_row = r
                        r += 1
                    if last_row < header_row + 1:
                        continue
                    # Aplicar formato porcentaje
                    for rr in range(header_row + 1, last_row + 1):
                        for cc in range(wp_col, end_col + 1):
                            ws.cell(row=rr, column=cc).number_format = '0.0%'
                    data_range = f"{_col_letter(wp_col)}{header_row + 1}:{_col_letter(end_col)}{last_row}"
                    # Regla < 0%: relleno rojo claro (#FFC7CE), texto rojo oscuro (#9C0006)
                    dxf_red = _Diff(
                        fill=_PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
                        font=_Font(color='9C0006')
                    )
                    rule_red = _Rule(type='cellIs', operator='lessThan', formula=['0'], dxf=dxf_red)
                    ws.conditional_formatting.add(data_range, rule_red)

                    # Regla > 0%: relleno verde claro (#C6EFCE), texto verde oscuro (#006100)
                    dxf_green = _Diff(
                        fill=_PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
                        font=_Font(color='006100')
                    )
                    rule_green = _Rule(type='cellIs', operator='greaterThan', formula=['0'], dxf=dxf_green)
                    ws.conditional_formatting.add(data_range, rule_green)

                wb2.save(xlsx_path)
            apply_variations_formatting(excel_temp_path)
            print(Fore.GREEN + "Formato de variaciones aplicado (0.0% + rojo/verde).")
        except Exception as e:
            print(Fore.YELLOW + f"No se pudo aplicar el formato de correlaciones: {e}")

    except Exception as e:
        print(f"{Fore.RED}{Style.BRIGHT}Error crítico durante la generación del archivo Excel: {e}")
        if os.path.exists(excel_temp_path):
             os.remove(excel_temp_path) # Limpiar si falla
        exit()

    # --- 1.14) Renombrar y mover archivo Excel final ---
    if not ref_month_year:
         print(f"{Fore.RED}{Style.BRIGHT}No se pudo determinar la fecha de referencia. No se puede renombrar el archivo Excel.")
         if os.path.exists(excel_temp_path):
             os.remove(excel_temp_path)
         exit()

    nombre_base_archivo = f"{pais_nombre}-{categoria_nombre_corto}-{fabricante}-{ref_month_year}_{coverage_label}"
    carpeta_salida = os.path.join(root_dir, nombre_base_archivo) # Carpeta con el mismo nombre base

    if not os.path.exists(carpeta_salida):
        try:
            os.makedirs(carpeta_salida)
            print(Fore.BLUE + "Carpeta de salida creada")
        except OSError as e:
            print(f"{Fore.RED}Error al crear carpeta de salida '{carpeta_salida}': {e}")
            if os.path.exists(excel_temp_path): os.remove(excel_temp_path)
            exit()
    else:
        print(Fore.YELLOW + "Carpeta de salida ya existe, no se creara de nuevo")

    nombre_template_final = f"{nombre_base_archivo}.xlsx"
    ruta_template_final = os.path.join(carpeta_salida, nombre_template_final)

    try:
        if os.path.exists(ruta_template_final):
            print(Fore.YELLOW + f"Archivo Excel ya existe. Se sobrescribirá.")
            os.remove(ruta_template_final)
        os.rename(excel_temp_path, ruta_template_final)
        print(Fore.GREEN + "Archivo Excel final guardado")
    except Exception as e:
        print(f"{Fore.RED}Error al mover/renombrar archivo Excel final: {e}")
        if os.path.exists(excel_temp_path): os.remove(excel_temp_path) # Limpiar temporal si falla el renombrado
        exit()


    return ref_month_year, carpeta_salida, nombre_base_archivo, ruta_template_final

def compute_coverage_dataframe(
    df_marca: "pd.DataFrame",
    pais_nombre: str,
    coverage_type: str,
    round_coverage: bool,
) -> "pd.DataFrame":
    """Calcula la cobertura rolling de 12 meses para cada pipeline."""
    acum_sell_out_py = df_marca[COL_SELL_OUT].rolling(window=12, min_periods=12).sum()
    acum_sell_out_py.index = df_marca[COL_DATA]
    df_coverage = pd.DataFrame(index=acum_sell_out_py.index)
    for p in range(7):
        sell_in_shifted = df_marca[COL_SELL_IN].shift(p)
        acum_sell_in_shifted = sell_in_shifted.rolling(window=12, min_periods=12).sum()
        acum_sell_in_shifted.index = df_marca[COL_DATA]
        coverage_p = (acum_sell_out_py / acum_sell_in_shifted) * 100
        df_coverage[f'P{p}'] = coverage_p
    pop_val_num = float(pop_coverage.get(pais_nombre, DEFAULT_POP_COVERAGE).replace('%', '')) / 100.0
    if coverage_type.lower() == "relativa" and pop_val_num > 0:
        df_coverage = df_coverage / pop_val_num
    if round_coverage:
        df_coverage = df_coverage.apply(_round_half_up_series)
    else:
        df_coverage = df_coverage.round(1)
    return df_coverage


def compute_variations_dataframe(df_marca: "pd.DataFrame") -> "pd.DataFrame":
    period_types = ["Anual", "Semestral", "Trimestral"]
    df_variations = pd.DataFrame(columns=['Tipo', 'Periodo', 'WP by Numerator'] + [f'Cliente P{p}' for p in range(7)])
    kantar_vars_y1 = calc_var1(df_marca, COL_SELL_OUT, 0)
    cliente_vars_y1 = {p: calc_var1(df_marca, COL_SELL_IN, p) for p in range(7)}
    for i, p_type in enumerate(period_types):
        row = {'Tipo': p_type, 'Periodo': f'{p_type} vs Y-1', 'WP by Numerator': kantar_vars_y1[i]}
        for p in range(7):
            row[f'Cliente P{p}'] = cliente_vars_y1[p][i]
        df_variations.loc[len(df_variations)] = row
    kantar_vars_y2 = calc_var2(df_marca, COL_SELL_OUT, 0)
    cliente_vars_y2 = {p: calc_var2(df_marca, COL_SELL_IN, p) for p in range(7)}
    for i, p_type in enumerate(period_types):
        row = {'Tipo': p_type, 'Periodo': f'{p_type} vs Y-2', 'WP by Numerator': kantar_vars_y2[i]}
        for p in range(7):
            row[f'Cliente P{p}'] = cliente_vars_y2[p][i]
        df_variations.loc[len(df_variations)] = row
    return df_variations


def compute_averages(df_marca: "pd.DataFrame") -> Dict[str, float]:
    averages = {}
    n_data = len(df_marca)
    if n_data >= 12:
        averages['Penet_MAT_Actual'] = df_marca[COL_PENET].iloc[-12:].mean()
        averages['Buyers_MAT_Actual'] = df_marca[COL_BUYERS].iloc[-12:].mean()
    else:
        averages['Penet_MAT_Actual'] = df_marca[COL_PENET].mean()
        averages['Buyers_MAT_Actual'] = df_marca[COL_BUYERS].mean()
    if n_data >= 24:
        averages['Penet_MAT_Anterior'] = df_marca[COL_PENET].iloc[-24:-12].mean()
        averages['Buyers_MAT_Anterior'] = df_marca[COL_BUYERS].iloc[-24:-12].mean()
    else:
        averages['Penet_MAT_Anterior'] = np.nan
        averages['Buyers_MAT_Anterior'] = np.nan
    return averages


def compute_trend_plot_df(df_marca: "pd.DataFrame") -> "pd.DataFrame":
    df_trend_plot = df_marca[[COL_DATA, COL_SELL_IN, COL_SELL_OUT]].copy()
    df_trend_plot[COL_DATA] = df_trend_plot[COL_DATA].apply(lambda x: x.strftime('%m-%y'))
    return df_trend_plot


def build_variation_table(
    fabricante: str,
    labels: Dict[Tuple[int, str], List[str] | str],
    lang_index: int,
    pipeline: int,
    ref_month_year: str,
    var_cliente_mat: Optional[float],
    var_kantar_mat: Optional[float],
) -> "pd.DataFrame":
    label_var = labels[(lang_index, 'Var')]
    data = {
        " ": [f"VAR % MAT ({ref_month_year})"],
        f"{fabricante} {label_var} Pipeline {pipeline}": [f"{var_cliente_mat*100:.1f}%" if pd.notna(var_cliente_mat) else "-"],
        "Worldpanel by Numerator": [f"{var_kantar_mat*100:.1f}%" if pd.notna(var_kantar_mat) else "-"],
    }
    return pd.DataFrame(data)

def build_variations_detail_table(
    df_variations: "pd.DataFrame",
    pipeline: int,
    df_marca: "pd.DataFrame",
) -> "pd.DataFrame":
    """Construye la tabla de variaciones utilizada en el slide de tendencia."""
    if df_variations is None or df_variations.empty:
        return pd.DataFrame()
    filtered = df_variations[df_variations['Periodo'].astype(str).str.contains('Y-1', na=False)].copy()
    if filtered.empty:
        return pd.DataFrame()

    base_columns = [col for col in ['Tipo', 'Periodo', 'WP by Numerator', 'Cliente P0'] if col in filtered.columns]
    if not base_columns:
        return pd.DataFrame()
    detail_df = filtered[base_columns].copy()

    pipeline_col = f'Cliente P{pipeline}'
    if pipeline_col in df_variations.columns:
        detail_df[f'Cliente Pipeline (P{pipeline})'] = df_variations.loc[detail_df.index, pipeline_col].values

    def _format_month(dt: "pd.Timestamp") -> str:
        if pd.isna(dt):
            return "-"
        dt = pd.to_datetime(dt)
        return f"{month_abbr[dt.month]}-{dt.year % 100:02d}"

    if df_marca is not None and not df_marca.empty:
        try:
            current_dt = pd.to_datetime(df_marca[COL_DATA].iloc[-1])
            period_specs = {
                'Anual': ('MAT', 12),
                'Semestral': ('SEM', 6),
                'Trimestral': ('TRI', 3),
            }
            formatted_periods: List[str] = []
            for _, row in detail_df.iterrows():
                tipo = row.get('Tipo')
                label, offset = period_specs.get(tipo, ("", None))
                if offset is None or pd.isna(current_dt):
                    formatted_periods.append(row.get('Periodo', ''))
                    continue
                previous_dt = current_dt - pd.DateOffset(months=offset)
                formatted_periods.append(f"{label} {_format_month(current_dt)} x {label} {_format_month(previous_dt)}")
            detail_df['Periodo'] = formatted_periods
        except Exception:
            pass

    detail_df.reset_index(drop=True, inplace=True)
    return detail_df


def build_evolution_figure(df_marca: "pd.DataFrame", pipeline: int, lang_index: int) -> Optional["plt.Figure"]:
    if len(df_marca) < 24:
        return None
    df_evol = df_marca[[COL_DATA, COL_SELL_IN, COL_SELL_OUT]].copy()
    df_evol[COL_DATA] = pd.to_datetime(df_evol[COL_DATA])
    return generar_grafico_evolucion_mensual(df_evol, pipeline, lang_index)



def build_summary_and_bank_rows(
    pipeline: int,
    marca_nombre_limpio: str,
    coverage_series: "pd.Series",
    df_variations: "pd.DataFrame",
    averages: Dict[str, float],
    labels: Dict[Tuple[int, str], List[str] | str],
    lang_index: int,
    fabricante: str,
    pais_nombre: str,
    categoria_nombre: str,
    cesta_nombre: str,
    coverage_reason: str,
    measure_unit: str,
    coverage_type: str,
    ref_month_year: str,
    round_coverage: bool,
) -> Tuple[Dict[str, str], Dict[str, object], float, float, float, str]:
    coverage_series = coverage_series.dropna()
    if not coverage_series.empty:
        coverage_actual = coverage_series.iloc[-1]
        coverage_anterior = coverage_series.iloc[-13] if len(coverage_series) >= 13 else np.nan
    else:
        coverage_actual = np.nan
        coverage_anterior = np.nan
    var_cliente_anual_y1 = df_variations.loc[df_variations['Tipo'] == 'Anual', f'Cliente P{pipeline}'].iloc[0]
    var_kantar_anual_y1 = df_variations.loc[df_variations['Tipo'] == 'Anual', 'WP by Numerator'].iloc[0]
    tendencia_alineada = "NO"
    if pd.notna(var_cliente_anual_y1) and pd.notna(var_kantar_anual_y1):
        if (var_cliente_anual_y1 * var_kantar_anual_y1) > 0:
            tendencia_alineada = "SI"
        elif var_cliente_anual_y1 == 0 and var_kantar_anual_y1 == 0:
            tendencia_alineada = "SI"
    if round_coverage:
        cov_actual_val = int(np.floor(coverage_actual + 0.5)) if pd.notna(coverage_actual) else 0
        cov_anterior_val = int(np.floor(coverage_anterior + 0.5)) if pd.notna(coverage_anterior) else 0
        estabilidad = cov_actual_val - cov_anterior_val
    else:
        cov_actual_val = round(coverage_actual, 1) if pd.notna(coverage_actual) else 0
        cov_anterior_val = round(coverage_anterior, 1) if pd.notna(coverage_anterior) else 0
        estabilidad = round(cov_actual_val - cov_anterior_val, 1)
    summary_row = {
        labels[(lang_index, 'Summary')][0]: marca_nombre_limpio,
        labels[(lang_index, 'Summary')][1]: pipeline,
        labels[(lang_index, 'Summary')][2]: f"{averages.get('Penet_MAT_Actual', 0):.1f}%",
        labels[(lang_index, 'Summary')][3]: f"{var_cliente_anual_y1*100:.1f}%" if pd.notna(var_cliente_anual_y1) else "0.0%",
        labels[(lang_index, 'Summary')][4]: f"{var_kantar_anual_y1*100:.1f}%" if pd.notna(var_kantar_anual_y1) else "0.0%",
        labels[(lang_index, 'Summary')][5]: (str(cov_anterior_val) if round_coverage else (f"{coverage_anterior:.1f}" if pd.notna(coverage_anterior) else "0.0")),
        labels[(lang_index, 'Summary')][6]: (str(cov_actual_val) if round_coverage else (f"{coverage_actual:.1f}" if pd.notna(coverage_actual) else "0.0")),
        labels[(lang_index, 'Summary')][7]: (str(estabilidad) if round_coverage else (f"{estabilidad:.1f}" if pd.notna(estabilidad) else "0.0")),
    }
    banco_row = {
        'Periodo': dt.strptime(ref_month_year, '%m-%y').date(),
        'Fabricante': fabricante,
        'Categoria': categoria_nombre,
        'Fabricante/Marca': marca_nombre_limpio,
        'Cesta': cesta_nombre,
        'Panel': 'PNC',
        'Unidad': measure_unit,
        'Razon': coverage_reason,
        'Pais': pais_nombre,
        'Ampliacion': 'SI',
        'Penet Media Ano Mov Atual': round(averages.get('Penet_MAT_Actual', 0), 1),
        'Penet Media Ano Mov Anterior': round(averages.get('Penet_MAT_Anterior', 0), 1),
        'Raw Buyers Media Ano Mov Atual': round(averages.get('Buyers_MAT_Actual', 0), 1),
        'Pipeline': pipeline,
        'Cobertura Año Mov Actual': cov_actual_val,
        'Cobertura Año Mov Anterior': cov_anterior_val,
        '%VAR Cliente': round(var_cliente_anual_y1 * 100, 1) if pd.notna(var_cliente_anual_y1) else 0,
        '% VAR WP by Numerator': round(var_kantar_anual_y1 * 100, 1) if pd.notna(var_kantar_anual_y1) else 0,
        'Misma Tendencia': tendencia_alineada,
        'Estabilidad': estabilidad,
    }
    return summary_row, banco_row, cov_actual_val, cov_anterior_val, estabilidad, tendencia_alineada


COVERAGE_BANK_COLUMNS = [
    'Periodo', 'Fabricante', 'Categoria', 'Fabricante/Marca', 'Cesta', 'Panel', 'Unidad',
    'Razon', 'Pais', 'Ampliacion', 'Penet Media Ano Mov Atual', 'Penet Media Ano Mov Anterior',
    'Raw Buyers Media Ano Mov Atual', 'Pipeline', 'Cobertura Año Mov Actual',
    'Cobertura Año Mov Anterior', '%VAR Cliente', '% VAR WP by Numerator', 'Misma Tendencia', 'Estabilidad'
]

def generate_presentation_and_bank(
    root_dir: str,
    excel_file_obj: "pd.ExcelFile",
    marcas: Sequence[str],
    pais_nombre: str,
    categoria_nombre: str,
    categoria_nombre_corto: str,
    fabricante: str,
    cesta_nombre: str,
    coverage_label: str,
    coverage_type: str,
    coverage_reason: str,
    ref_month_year: str,
    carpeta_salida: str,
    nombre_base_archivo: str,
    include_english: bool,
    trend_axis: str,
    round_coverage: bool,
) -> Tuple[str, "pd.DataFrame", "pd.DataFrame"]:
    chosen_lang, lang_index = determine_language(include_english, pais_nombre)
    ppt, tmp_ppt_path = copy_and_prune_template(root_dir, chosen_lang)
    labels = build_labels(lang_index, fabricante, ref_month_year)
    builder = SlideBuilder(ppt, lang_index, labels, coverage_label, trend_axis)
    builder.configure_cover(pais_nombre, fabricante, categoria_nombre, ref_month_year, chosen_lang)

    summary_rows: List[Dict[str, str]] = []
    bank_rows: List[Dict[str, object]] = []

    total_slides_to_generate = 0
    for marca_sheet_name in marcas:
        df_marca_ppt, _ = load_and_preprocess_sheet(excel_file_obj, marca_sheet_name)
        if df_marca_ppt is None:
            continue
        match = re.match(r"(?i)^p([0-6])_", marca_sheet_name)
        pipelines_to_run = [int(match.group(1))] if match else list(range(7))
        n_slides_marca = len(pipelines_to_run) * (2 + (1 if len(df_marca_ppt) >= 24 else 0))
        total_slides_to_generate += n_slides_marca

    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TextColumn("{task.completed}/{task.total}"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        transient=True,
    )

    with progress:
        task_id = progress.add_task("Creando Diapositivas PPT", total=total_slides_to_generate + 1)
        for marca_sheet_name in marcas:
            df_marca_ppt, measure_unit = load_and_preprocess_sheet(excel_file_obj, marca_sheet_name)
            if df_marca_ppt is None:
                continue
            marca_nombre_limpio = re.sub(r"(?i)^p[0-6]_", "", marca_sheet_name)
            match = re.match(r"(?i)^p([0-6])_", marca_sheet_name)
            pipelines_to_run = [int(match.group(1))] if match else list(range(7))
            df_coverage = compute_coverage_dataframe(df_marca_ppt, pais_nombre, coverage_type, round_coverage)
            df_variations = compute_variations_dataframe(df_marca_ppt)
            averages = compute_averages(df_marca_ppt)
            df_trend_plot = compute_trend_plot_df(df_marca_ppt)
            for pipeline in pipelines_to_run:
                coverage_series = df_coverage[f'P{pipeline}']
                var_cliente_mat = df_variations.loc[df_variations['Tipo'] == 'Anual', f'Cliente P{pipeline}'].iloc[0]
                var_kantar_mat = df_variations.loc[df_variations['Tipo'] == 'Anual', 'WP by Numerator'].iloc[0]
                variation_table = build_variation_table(
                    fabricante,
                    labels,
                    lang_index,
                    pipeline,
                    ref_month_year,
                    var_cliente_mat,
                    var_kantar_mat,
                )
                variations_detail = build_variations_detail_table(df_variations, pipeline, df_marca_ppt)
                evolution_figure = build_evolution_figure(df_marca_ppt, pipeline, lang_index)
                assets = PipelineAssets(
                    pipeline=pipeline,
                    marca=marca_nombre_limpio,
                    coverage_series=coverage_series,
                    penetration_series=df_marca_ppt.set_index(COL_DATA)[COL_PENET].loc[coverage_series.dropna().index],
                    variation_table=variation_table,
                    trend_plot_df=df_trend_plot,
                    variations_detail=variations_detail,
                    evolution_figure=evolution_figure,
                )
                builder.add_pipeline_slides(
                    assets,
                    marca_nombre_limpio=marca_nombre_limpio,
                    lang_index=lang_index,
                    coverage_label=builder.coverage_label,
                    progress=progress,
                    task_id=task_id,
                )
                summary_row, bank_row, _, _, _, _ = build_summary_and_bank_rows(
                    pipeline=pipeline,
                    marca_nombre_limpio=marca_nombre_limpio,
                    coverage_series=coverage_series,
                    df_variations=df_variations,
                    averages=averages,
                    labels=labels,
                    lang_index=lang_index,
                    fabricante=fabricante,
                    pais_nombre=pais_nombre,
                    categoria_nombre=categoria_nombre,
                    cesta_nombre=cesta_nombre,
                    coverage_reason=coverage_reason,
                    measure_unit=measure_unit,
                    coverage_type=coverage_type,
                    ref_month_year=ref_month_year,
                    round_coverage=round_coverage,
                )
                summary_rows.append(summary_row)
                bank_rows.append(bank_row)
        progress.update(task_id, advance=1)

    df_summary = pd.DataFrame(summary_rows)
    if not df_summary.empty:
        df_summary = df_summary[labels[(lang_index, 'Summary')]]
    df_bank = pd.DataFrame(bank_rows, columns=COVERAGE_BANK_COLUMNS)

    builder.add_summary_slide(df_summary, pais_nombre, categoria_nombre)
    builder.insert_thanks_text(chosen_lang)
    builder.reorder_summary_and_credit()

    ruta_ppt_final = os.path.join(carpeta_salida, f"{nombre_base_archivo}.pptx")
    ppt.save(ruta_ppt_final)

    return ruta_ppt_final, df_summary, df_bank


def save_coverage_bank(
    df_bank: "pd.DataFrame",
    carpeta_salida: str,
    nombre_base_archivo: str,
    fabricante: str,
    categoria_nombre: str,
    categoria_nombre_corto: str,
    pais_nombre: str,
    ref_month_year: str,
    coverage_label: str,
) -> str:
    df_bank = df_bank.copy()
    try:
        mes_ejecucion_dt = datetime.now().date().replace(day=1)
        if 'Mes_Ejecucion' not in df_bank.columns:
            df_bank.insert(0, 'Mes_Ejecucion', mes_ejecucion_dt)
        else:
            df_bank['Mes_Ejecucion'] = mes_ejecucion_dt
    except Exception as exc:
        print(f"{Fore.YELLOW}Advertencia: No se pudo agregar la columna 'Mes_Ejecucion': {exc}")
    categoria_para_banco = categoria_nombre_corto or categoria_nombre
    nombre_banco_final = f"Banco_{fabricante}_{categoria_para_banco}_{pais_nombre}_{ref_month_year}_{coverage_label}.xlsx"
    ruta_banco_final = os.path.join(carpeta_salida, nombre_banco_final)
    df_bank.to_excel(ruta_banco_final, index=False)
    try:
        from openpyxl import load_workbook as _wb_load
        wb_bank = _wb_load(ruta_banco_final)
        for ws in wb_bank.worksheets:
            header_map = {}
            for cell in ws[1]:
                if cell.value is not None:
                    header_map[str(cell.value).strip().lower()] = cell.column
            for header_name in ['periodo', 'mes_ejecucion']:
                col_idx = header_map.get(header_name)
                if col_idx is None:
                    continue
                for r in range(2, ws.max_row + 1):
                    c = ws.cell(row=r, column=col_idx)
                    c.number_format = 'mmm-yy'
        wb_bank.save(ruta_banco_final)
    except Exception as exc:
        print(f"{Fore.YELLOW}Advertencia: No se pudo aplicar formato mmm-yy en Banco: {exc}")
    print(Fore.MAGENTA + "-> Banco de coberturas guardado")
    return ruta_banco_final


def cleanup_temp_dir(root_dir: str) -> None:
    tmp_dir = os.path.join(root_dir, 'tmp')
    if os.path.isdir(tmp_dir):
        shutil.rmtree(tmp_dir)
        print(Fore.BLUE + "Carpeta temporal ./tmp eliminada")


class CoverageStudioUltraApp:
    def __init__(self) -> None:
        self.root_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.root_dir)
        self.categories: Optional["pd.DataFrame"] = None

    def list_excel_files(self) -> List[str]:
        return [f for f in os.listdir(self.root_dir) if f.endswith('.xlsx') and not f.startswith('~$') and f != EXCEL_TEMP_FILENAME]

    def ensure_categories_loaded(self) -> None:
        if self.categories is None:
            wait_for_heavy_modules()
            self.categories = load_categories()

    def select_files(self, excel_list: Sequence[str]) -> List[str]:
        print(Fore.CYAN + "Archivos Excel (.xlsx) encontrados:")
        for i, archivo in enumerate(excel_list, start=1):
            meta = quick_file_metadata(archivo)
            if meta:
                print(Fore.BLUE + f"{i}. {archivo} " + Fore.YELLOW + f"| {meta}")
            else:
                print(Fore.BLUE + f"{i}. {archivo}")
        while True:
            opcion = input(
                Fore.WHITE
                + f"Seleccione el número de archivo a procesar (1-{len(excel_list)}).\n"
                + "Puede separar varios con comas o escribir 'all': "
            )
            opcion = opcion.strip().lower()
            if opcion in {"all", "todos", "*"}:
                selected_indices = list(range(1, len(excel_list) + 1))
            else:
                try:
                    selected_indices = [int(x) for x in opcion.split(',') if x]
                except ValueError:
                    print(Fore.RED + Style.BRIGHT + "Entrada inválida. Ingrese números separados por coma o 'all'.")
                    continue
                if not all(1 <= idx <= len(excel_list) for idx in selected_indices):
                    print(Fore.RED + "Uno o más números están fuera de rango. Intente nuevamente.")
                    continue
            selected_files = [excel_list[idx - 1] for idx in selected_indices]
            SELECTIONS['Excel'] = ", ".join(selected_files)
            clear_and_print_summary()
            return selected_files

    def gather_interactive_options(self) -> ExecutionOptions:
        coverage_type = tipo_cobertura()
        auto_mode = str(coverage_type).strip().lower() == "auto"
        if auto_mode:
            coverage_type_value = "Absoluta"
            coverage_reason = "Actualización periódica por contrato"
            trend_axis = "simple"
            include_english = False
            round_cov = False
            SELECTIONS['Razón'] = coverage_reason
            SELECTIONS['Eje tendencia'] = trend_axis
            SELECTIONS['Idioma PPT'] = 'ESPAÑOL'
            SELECTIONS['Inglés'] = 'No'
            SELECTIONS['Redondeo Cobertura'] = 'No'
            clear_and_print_summary()
        else:
            coverage_type_value = coverage_type
            coverage_reason = razao_cov()
            trend_axis = tipo_eje_tendencia()
            include_english = include_english_flag()
            round_cov = round_coverage_flag()
        return ExecutionOptions(
            coverage_type=coverage_type_value,
            coverage_reason=coverage_reason,
            trend_axis=trend_axis,
            include_english=include_english,
            round_coverage=round_cov,
            auto_mode=auto_mode,
        )


    def process_file(self, excel_file_name: str, options: ExecutionOptions, idx: int, total: int) -> None:
        global ROUND_COVERAGE
        ROUND_COVERAGE = options.round_coverage
        self.ensure_categories_loaded()
        print_file_header(idx, total, excel_file_name)
        excel_file_path = os.path.join(self.root_dir, excel_file_name)
        try:
            excel_file_obj = pd.ExcelFile(excel_file_path)
            marcas = excel_file_obj.sheet_names
        except FileNotFoundError:
            print(f"{Fore.RED}{Style.BRIGHT}Error: No se encontró el archivo seleccionado: {excel_file_path}")
            return
        except Exception as exc:
            print(f"{Fore.RED}{Style.BRIGHT}Error al abrir el archivo Excel '{excel_file_name}': {exc}")
            return
        try:
            pais_nombre, cesta_nombre, categoria_nombre, categoria_nombre_corto, fabricante = parse_file_metadata(excel_file_name, self.categories)
        except ValueError as exc:
            print(f"{Fore.RED}{Style.BRIGHT}{exc}")
            return
        SELECTIONS['Pais'] = pais_nombre
        coverage_label = compute_coverage_label(options.coverage_type, options.include_english)
        ref_month_year, carpeta_salida, nombre_base_archivo, ruta_template_final = generate_excel_template(
            self.root_dir,
            excel_file_obj,
            marcas,
            pais_nombre,
            categoria_nombre,
            categoria_nombre_corto,
            fabricante,
            coverage_label,
            options.coverage_type,
            options.coverage_reason,
        )
        ruta_ppt_final, df_summary, df_bank = generate_presentation_and_bank(
            root_dir=self.root_dir,
            excel_file_obj=excel_file_obj,
            marcas=marcas,
            pais_nombre=pais_nombre,
            categoria_nombre=categoria_nombre,
            categoria_nombre_corto=categoria_nombre_corto,
            fabricante=fabricante,
            cesta_nombre=cesta_nombre,
            coverage_label=coverage_label,
            coverage_type=options.coverage_type,
            coverage_reason=options.coverage_reason,
            ref_month_year=ref_month_year,
            carpeta_salida=carpeta_salida,
            nombre_base_archivo=nombre_base_archivo,
            include_english=options.include_english,
            trend_axis=options.trend_axis,
            round_coverage=options.round_coverage,
        )
        ruta_banco_final = save_coverage_bank(
            df_bank=df_bank,
            carpeta_salida=carpeta_salida,
            nombre_base_archivo=nombre_base_archivo,
            fabricante=fabricante,
            categoria_nombre=categoria_nombre,
            categoria_nombre_corto=categoria_nombre_corto,
            pais_nombre=pais_nombre,
            ref_month_year=ref_month_year,
            coverage_label=coverage_label,
        )
        print_file_summary(ruta_template_final, ruta_ppt_final, ruta_banco_final)

    def run(self) -> None:
        excel_list = self.list_excel_files()
        if not excel_list:
            print(f"{Fore.RED}{Style.BRIGHT}Error: No se encontraron archivos .xlsx en la carpeta: {self.root_dir}")
            return
        env_options = ExecutionOptions.from_environment()
        if env_options:
            excel_file_name = os.environ['AUTO_FILE']
            idx = int(os.environ.get('AUTO_INDEX', '1'))
            total = int(os.environ.get('AUTO_TOTAL', '1'))
            self.process_file(excel_file_name, env_options, idx, total)
            cleanup_temp_dir(self.root_dir)
            return
        selected_files = self.select_files(excel_list)
        options = self.gather_interactive_options()
        total = len(selected_files)
        for idx, excel_file_name in enumerate(selected_files, start=1):
            self.process_file(excel_file_name, options, idx, total)
        cleanup_temp_dir(self.root_dir)



def main() -> None:
    app = CoverageStudioUltraApp()
    app.run()


if __name__ == "__main__":
    main()
