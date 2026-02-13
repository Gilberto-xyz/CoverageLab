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
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Callable, Set
from calendar import month_abbr

import colorama
from colorama import Fore, Style
from rich.console import Console
from rich.panel import Panel

colorama.init(autoreset=True)
console = Console()

BRAND_EXCEPTION_REASONS: Dict[str, Set[str]] = {}

EXCEPTION_STYLES: Dict[str, Dict[str, str]] = {
    "zero_dash": {
        "brand_color": Fore.YELLOW,
        "message": "contiene 0s en los en algunos meses, graficando con exepcion",
        "summary_tag": "0/-",
    },
    "negative": {
        "brand_color": Fore.YELLOW,
        "message": "contiene valores negativos, graficando con exepcion",
        "summary_tag": "neg",
    },
}

SUMMARY_EXTRA_MONTHS_ENV_KEYS: Tuple[str, ...] = ("AUTO_EXTEA", "AUTO_EXTRA_MONTHS")
SUMMARY_EXTRA_MONTHS_MODE_ENV_KEYS: Tuple[str, ...] = ("AUTO_EXTEA_MODE", "AUTO_EXTRA_MONTHS_MODE")
VARIATIONS_BOX_STYLE_ENV_KEYS: Tuple[str, ...] = ("AUTO_VAR_BOX_STYLE", "AUTO_VAR_STYLE")
COVERAGE_SLIDE_VARIANT_ENV_KEYS: Tuple[str, ...] = ("AUTO_COV_SLIDE", "AUTO_COV_SLIDE_STYLE")
EVOLUTION_SLIDE_VARIANT_ENV_KEYS: Tuple[str, ...] = ("AUTO_EVO_SLIDE", "AUTO_EVO_SLIDE_STYLE")
MONTH_TOKEN_TO_NUMBER: Dict[str, int] = {
    "ene": 1, "enero": 1, "jan": 1, "janeiro": 1, "january": 1,
    "feb": 2, "febrero": 2, "fev": 2, "fevereiro": 2, "february": 2,
    "mar": 3, "marzo": 3, "marco": 3, "march": 3,
    "abr": 4, "abril": 4, "apr": 4, "april": 4,
    "may": 5, "mayo": 5, "maio": 5,
    "jun": 6, "junio": 6, "junho": 6, "june": 6,
    "jul": 7, "julio": 7, "julho": 7, "july": 7,
    "ago": 8, "agosto": 8, "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "set": 9, "septiembre": 9, "setiembre": 9, "setembro": 9, "september": 9,
    "oct": 10, "octubre": 10, "out": 10, "outubro": 10, "october": 10,
    "nov": 11, "noviembre": 11, "novembro": 11, "november": 11,
    "dic": 12, "diciembre": 12, "dez": 12, "dezembro": 12, "dec": 12, "december": 12,
}

def normalize_variations_box_style(raw_value: Optional[str]) -> str:
    """Normaliza el estilo del cuadro de variaciones (classic | pretty)."""
    val = (raw_value or "").strip().lower()
    if not val:
        return "classic"
    if val in {"pretty", "bonito", "nuevo", "nice", "card", "cards", "2"}:
        return "pretty"
    if val in {"classic", "clasico", "clásico", "tabla", "1"}:
        return "classic"
    return "classic"


def normalize_coverage_slide_variant(raw_value: Optional[str]) -> str:
    """Normaliza el modo del slide de cobertura (classic | complemented)."""
    val = (raw_value or "").strip().lower()
    if not val:
        return "classic"
    if val in {"complemented", "complementado", "complement", "penetracion", "penetración", "penetration", "2"}:
        return "complemented"
    if val in {"classic", "clasico", "clásico", "variacion", "variación", "var", "1"}:
        return "classic"
    return "classic"

def normalize_evolution_slide_variant(raw_value: Optional[str]) -> str:
    """Normaliza el modo del slide de Evolucion mensual y variacion (classic | simple)."""
    val = (raw_value or "").strip().lower()
    if not val:
        return "classic"
    if val in {"simple", "basico", "basica", "basic", "1"}:
        return "simple"
    if val in {"classic", "clasico", "clásico", "avanzado", "advanced", "2"}:
        return "classic"
    return "classic"

def _register_brand_exception(marca_label: Optional[str], reason: str) -> None:
    normalized = (marca_label or "N/D").strip() or "N/D"
    reason_set = BRAND_EXCEPTION_REASONS.setdefault(normalized, set())
    if reason in reason_set:
        return
    reason_set.add(reason)
    style = EXCEPTION_STYLES.get(reason, EXCEPTION_STYLES["zero_dash"])
    # Mensaje en rojo con el nombre de la marca en amarillo al inicio
    print(f"{Fore.RED}{Fore.YELLOW}{normalized}{Fore.RED} {style['message']}")


def notify_zero_months_exception(marca_label: Optional[str]) -> None:
    _register_brand_exception(marca_label, "zero_dash")


def notify_negative_values_exception(marca_label: Optional[str]) -> None:
    _register_brand_exception(marca_label, "negative")


def notify_buyers_threshold(marca_label: Optional[str], buyers_value: Optional[float], threshold: float = 200) -> None:
    if buyers_value is None:
        return
    try:
        if pd.isna(buyers_value):
            return
        buyers_num = float(buyers_value)
    except Exception:
        return
    normalized = (marca_label or "N/D").strip() or "N/D"
    buyers_display = f"{buyers_num:.0f}"
    if buyers_num < threshold:
        print(Fore.RED + f"{normalized} cuenta con {buyers_display} compradores promedio, tener precaución")
    else:
        print(Fore.GREEN + f"{normalized} si cuenta con al menos {int(threshold)} compradores")


def report_zero_months_exceptions() -> None:
    if not BRAND_EXCEPTION_REASONS:
        return
    print(f"{Fore.RED}Marcas con excepción detectada:")
    for marca in sorted(BRAND_EXCEPTION_REASONS):
        tags = "/".join(
            sorted(
                {
                    EXCEPTION_STYLES.get(reason, {}).get("summary_tag", reason)
                    for reason in BRAND_EXCEPTION_REASONS[marca]
                }
            )
        )
        print(f"{Fore.RED}- {Fore.YELLOW}{marca}{Fore.RED} [{tags}]")
    BRAND_EXCEPTION_REASONS.clear()


def detect_brand_data_issues(df_marca: "pd.DataFrame", window: int = 12) -> Set[str]:
    issues: Set[str] = set()
    if df_marca is None or df_marca.empty:
        return issues
    cols_to_check = [COL_SELL_IN, COL_SELL_OUT]
    tail_df = df_marca.tail(window) if window > 0 else df_marca
    for col in cols_to_check:
        series = tail_df[col]
        str_series = series.astype(str).str.strip()
        if str_series.eq("-").any():
            issues.add("zero_dash")
        numeric = pd.to_numeric(series, errors="coerce")
        if (numeric == 0).any():
            issues.add("zero_dash")
        if (numeric < 0).any():
            issues.add("negative")
    return issues

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
COL_EVO_KANTAR_YOY = "% VAR WP by Numerator"
COL_EVO_SELLIN_YOY = "% VAR Sell-in (Cliente)"

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

def _normalize_month_token(token: str) -> str:
    token = (token or "").strip().lower()
    if not token:
        return ""
    normalized = unicodedata.normalize("NFKD", token)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))

def parse_summary_extra_months(raw_value: Optional[str]) -> List[int]:
    """Convierte una entrada como '8,ago,12' en meses [8, 12]."""
    if raw_value is None:
        return []
    tokens = [t for t in re.split(r"[,\s;/|]+", str(raw_value).strip()) if t]
    if not tokens:
        return []
    months: List[int] = []
    invalid: List[str] = []
    for token in tokens:
        normalized = _normalize_month_token(token)
        if normalized.isdigit():
            month_num = int(normalized)
        else:
            month_num = MONTH_TOKEN_TO_NUMBER.get(normalized, 0)
        if 1 <= month_num <= 12:
            if month_num not in months:
                months.append(month_num)
        else:
            invalid.append(token)
    if invalid:
        raise ValueError(f"Mes(es) inválido(s): {', '.join(invalid)}")
    return months

def parse_summary_extra_months_mode(raw_value: Optional[str]) -> str:
    if raw_value is None:
        return "recent"
    normalized = str(raw_value).strip().lower()
    recent_values = {"recent", "actual", "current", "solo", "ultimo", "último", "ultimo_mes", "single", "one", "1"}
    both_values = {"both", "ambos", "dos", "doble", "dual", "2", "two", "all"}
    if normalized in recent_values:
        return "recent"
    if normalized in both_values:
        return "both"
    raise ValueError(f"Modo de meses extra inválido: {raw_value}")

def format_summary_extra_months(months: Sequence[int]) -> str:
    if not months:
        return "Ninguno"
    return ", ".join(month_abbr[m].capitalize() for m in months if 1 <= m <= 12)

def get_summary_extra_months_from_env() -> List[int]:
    for key in SUMMARY_EXTRA_MONTHS_ENV_KEYS:
        raw = os.environ.get(key)
        if raw is None:
            continue
        try:
            return parse_summary_extra_months(raw)
        except ValueError as exc:
            print(Fore.YELLOW + f"Advertencia: {exc}. Se ignora {key}.")
            return []
    return []

def get_summary_extra_months_mode_from_env() -> Optional[str]:
    for key in SUMMARY_EXTRA_MONTHS_MODE_ENV_KEYS:
        raw = os.environ.get(key)
        if raw is None:
            continue
        try:
            return parse_summary_extra_months_mode(raw)
        except ValueError as exc:
            print(Fore.YELLOW + f"Advertencia: {exc}. Se ignora {key}.")
            return None
    return None

def summary_extra_months_option() -> List[int]:
    """Obtiene meses extra a mostrar en la tabla summary de cobertura."""
    for key in SUMMARY_EXTRA_MONTHS_ENV_KEYS:
        raw = os.environ.get(key)
        if raw is not None:
            months = get_summary_extra_months_from_env()
            SELECTIONS['Meses extra summary'] = format_summary_extra_months(months)
            clear_and_print_summary()
            return months

    print(Fore.CYAN + "\n¿Desea agregar meses extra al summary de cobertura?")
    print(Fore.WHITE + "Ingrese mes(es) (1-12 o nombre, separados por coma). Ej: 8,ago,nov")
    print(Fore.WHITE + "Presione ENTER para continuar sin meses extra.")
    while True:
        raw = input(Fore.GREEN + "Mes(es) extra: ").strip()
        if not raw:
            months = []
            break
        try:
            months = parse_summary_extra_months(raw)
            break
        except ValueError as exc:
            print(Fore.RED + str(exc) + ". Intente nuevamente.")
    SELECTIONS['Meses extra summary'] = format_summary_extra_months(months)
    clear_and_print_summary()
    return months

def summary_extra_months_mode_option(has_extra_months: bool) -> str:
    env_mode = get_summary_extra_months_mode_from_env()
    if env_mode:
        # Evitar confusión: si no hay meses extra, el modo no aplica y no se muestra.
        if has_extra_months:
            SELECTIONS['Modo meses extra summary'] = "Mes más reciente" if env_mode == "recent" else "Año actual y anterior"
            clear_and_print_summary()
        return env_mode
    if not has_extra_months:
        return "recent"
    print(Fore.CYAN + "\n¿Modo de meses extra en summary?")
    print(Fore.WHITE + "1 - Solo mes más reciente (año actual)")
    print(Fore.WHITE + "2 - Dos meses (año actual y año anterior)")
    opciones = {"1": "recent", "2": "both"}
    eleccion = input(Fore.GREEN + "Elija 1 o 2: ").strip()
    modo = opciones.get(eleccion, "recent")
    SELECTIONS['Modo meses extra summary'] = "Mes más reciente" if modo == "recent" else "Año actual y anterior"
    clear_and_print_summary()
    return modo

def clear_and_print_summary():
    """Limpia la terminal y muestra un resumen de las selecciones del usuario."""
    os.system('cls' if os.name == 'nt' else 'clear') # Compatible con Windows y Linux/Mac
    print(Fore.CYAN + Style.BRIGHT + "Resumen de opciones seleccionadas:")

    displayed: Set[str] = set()

    def _get(key: str) -> Optional[object]:
        return SELECTIONS.get(key)

    def _as_text(val: object) -> str:
        if val is None:
            return "-"
        txt = str(val).strip()
        if not txt:
            return "-"
        # Evitar mojibake en consolas Windows (codepages) para el resumen.
        try:
            txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
        except Exception:
            pass
        return txt if txt else "-"

    def _line(label: str, key: str, value: object) -> None:
        displayed.add(key)
        print(Fore.BLUE + f"{label}: " + Fore.YELLOW + _as_text(value))

    # --- Archivo / contexto ---
    if _get("Excel") is not None:
        _line("Archivo Excel", "Excel", _get("Excel"))
    if _get("Pais") is not None:
        _line("Pais (detectado)", "Pais", _get("Pais"))
    elif _get("Excel") is not None:
        # El pais se infiere del nombre del archivo al momento de procesarlo.
        _line("Pais (detectado)", "Pais", "Pendiente (se detecta al procesar)")

    # --- Cobertura ---
    cov = _as_text(_get("Cobertura"))
    if cov != "-":
        cov_disp = cov
        if cov.strip().lower() == "auto":
            cov_disp = "AUTO (usa configuracion predeterminada)"
        _line("Tipo de cobertura", "Cobertura", cov_disp)
    if _get("Razón") is not None:
        _line("Razon de cobertura", "Razón", _get("Razón"))
    if _get("Redondeo Cobertura") is not None:
        round_val = str(_get("Redondeo Cobertura")).strip().lower()
        round_disp = "Si (sin decimales)" if round_val in {"si", "sí", "yes", "y", "true", "1"} else "No (1 decimal)"
        _line("Redondeo de cobertura", "Redondeo Cobertura", round_disp)

    # --- Slides ---
    if _get("Slide Cobertura") is not None:
        slide_mode = str(_get("Slide Cobertura")).strip().lower()
        if "complement" in slide_mode:
            slide_disp = "Complementado (Penetracion MAT + Cobertura puntual + Estabilidad)"
        else:
            slide_disp = "Clasico (tabla VAR % MAT)"
        _line("Slide de cobertura", "Slide Cobertura", slide_disp)
    if _get("Slide Evolucion") is not None:
        evo_mode = str(_get("Slide Evolucion")).strip().lower()
        evo_disp = "Simple (lineas de variacion)" if "simple" in evo_mode else "Clasico/avanzado (volumen + barras)"
        _line("Slide evolucion mensual", "Slide Evolucion", evo_disp)
    if _get("Estilo variaciones") is not None:
        var_style = str(_get("Estilo variaciones")).strip().lower()
        var_disp = "Bonito (tarjetas)" if "bonit" in var_style else "Clasico (tabla)"
        _line("Cuadro de variaciones (tendencia)", "Estilo variaciones", var_disp)

    # --- Tendencia ---
    if _get("Eje tendencia") is not None:
        eje = str(_get("Eje tendencia")).strip().lower()
        eje_disp = "Simple (un eje)" if eje == "simple" else ("Doble (2 ejes)" if eje == "doble" else eje)
        _line("Grafico de tendencia", "Eje tendencia", eje_disp)

    # --- Idioma ---
    # Mostrar de forma consistente y evitando depender de que el país esté disponible (a veces se define después).
    include_en = str(_get("Inglés") or "").strip().lower() in {"sí", "si", "yes", "y", "true", "1"}
    pais_norm = str(_get("Pais") or "").strip().lower()
    if include_en:
        idioma_disp = "EN (forzado)"
    elif pais_norm in {"brasil", "brazil"}:
        idioma_disp = "PT (por pais)"
    elif pais_norm:
        idioma_disp = "ES (por pais)"
    elif _get("Idioma PPT") is not None:
        # Compatibilidad con el texto legado si existiera.
        idioma_disp = _as_text(_get("Idioma PPT"))
    else:
        idioma_disp = "Auto (por pais)"
    _line("Idioma PPT", "Idioma PPT", idioma_disp)
    displayed.add("Inglés")  # se muestra en Idioma PPT (aunque no exista aún)

    # --- Summary extra months ---
    meses_extra_val = _get("Meses extra summary")
    if meses_extra_val is not None:
        _line("Meses extra (summary)", "Meses extra summary", meses_extra_val)
    modo_val = _get("Modo meses extra summary")
    if modo_val is not None:
        meses_txt = str(meses_extra_val or "").strip().lower()
        no_aplica = (not meses_txt) or (meses_txt in {"ninguno", "-"})
        modo_disp = f"{_as_text(modo_val)}{' (no aplica: no hay meses extra)' if no_aplica else ''}"
        _line("Modo meses extra (summary)", "Modo meses extra summary", modo_disp)

    # Mostrar cualquier otro valor no incluido para evitar "desaparecen opciones".
    remaining = [k for k in SELECTIONS.keys() if k not in displayed]
    if remaining:
        print(Fore.CYAN + "Otros:")
        for k in sorted(remaining):
            print(Fore.BLUE + f"- {k}: " + Fore.YELLOW + _as_text(SELECTIONS.get(k)))

    print("\n" + "-"*50 + "\n")

def print_file_header(idx: int, total: int, filename: str) -> None:
    """Muestra un encabezado visual para la ejecución de un archivo."""
    console.rule(f"[bold cyan]Procesando archivo {idx}/{total}: {filename}")

# --- Función para mostrar resumen de archivos generados ---
def _format_path_for_summary(path_str: str, *, base_dir: Optional[str] = None, max_len: int = 90) -> str:
    """
    Formatea rutas para mostrarlas en consola sin confundir con paths largos:
    - Preferir ruta relativa (a base_dir o al cwd) cuando sea posible.
    - Si sigue siendo muy larga, elidir el medio (mantener inicio y el final).
    """
    if not path_str:
        return ""

    norm = os.path.normpath(str(path_str))
    try:
        abs_path = os.path.abspath(norm)
    except Exception:
        abs_path = norm

    display = abs_path

    def _try_relpath(target: str, base: str) -> Optional[str]:
        try:
            rel = os.path.relpath(target, base)
            # Solo usar relpath si no se va "hacia arriba" (..)
            if not rel.startswith(".."):
                return rel
        except Exception:
            return None
        return None

    if base_dir:
        base_abs = os.path.abspath(os.path.normpath(base_dir))
        rel = _try_relpath(abs_path, base_abs)
        if rel:
            display = rel
    else:
        rel = _try_relpath(abs_path, os.getcwd())
        if rel:
            display = rel

    display = os.path.normpath(display)
    if len(display) <= max_len:
        return display

    # Elidir el medio manteniendo el final (más útil para ubicar el archivo).
    parts = display.split(os.sep)
    if len(parts) <= 2:
        return "..." + display[-(max_len - 3):]

    tail_parts = parts[-3:] if len(parts) >= 3 else parts[-2:]
    tail = os.sep.join(tail_parts)
    head = parts[0]
    candidate = head + os.sep + "..." + os.sep + tail
    if len(candidate) <= max_len:
        return candidate

    head_short = (head[:20] + "...") if len(head) > 23 else (head + "...")
    candidate = head_short + os.sep + "..." + os.sep + tail
    if len(candidate) <= max_len:
        return candidate

    # Fallback: solo el final
    candidate = "..." + os.sep + tail
    if len(candidate) <= max_len:
        return candidate
    return "..." + tail[-(max_len - 3):]

def _format_elapsed(seconds: float) -> str:
    try:
        total = int(round(float(seconds)))
    except Exception:
        return "-"
    if total < 0:
        total = 0
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h}:{m:02d}:{s:02d}"


def print_file_locked_error(path_str: str, *, elapsed_seconds: Optional[float] = None) -> None:
    """Muestra un panel rojo cuando no se puede reescribir un archivo por estar en uso (Windows/Excel/PPT abierto)."""
    path_disp = str(path_str or "").strip() or "-"
    try:
        base = os.path.basename(path_disp) if path_disp not in {"-", ""} else "-"
    except Exception:
        base = path_disp

    hora_actual = datetime.now().strftime("%I:%M:%S %p")
    elapsed_line = ""
    if elapsed_seconds is not None:
        elapsed_line = f"\n[white]Tiempo total: [bold]{_format_elapsed(elapsed_seconds)}[/bold][/white]"

    msg = (
        "[bright_white]Proceso terminado con error[/bright_white]\n\n"
        f"[white]Archivo en uso: [bold]{base}[/bold][/white]\n"
        "[white]No se pudo reescribir porque esta abierto o bloqueado.[/white]\n"
        "[white]Cierra el archivo y vuelve a ejecutar.[/white]\n\n"
        f"[white]Hora de finalizacion: [bold]{hora_actual}[/bold][/white]"
        f"{elapsed_line}\n\n"
        f"[grey]{path_disp}[/grey]"
    )
    console.print()
    console.print(Panel.fit(msg, border_style="red", title="Coverages Latam"))
    console.print()


def print_file_summary(ruta_excel: str, ruta_ppt: str, ruta_banco: str, *, elapsed_seconds: Optional[float] = None) -> None:
    """Muestra un resumen con las rutas generadas para el archivo."""
    console.print("\n[blue]Resumen de archivos generados:[/blue]")

    items: List[Tuple[str, str]] = [
        ("Excel", ruta_excel),
        ("Presentación", ruta_ppt),
        ("Banco", ruta_banco),
    ]
    present = [(label, p) for label, p in items if p]

    common_dir = ""
    if present:
        try:
            parents = [os.path.dirname(os.path.abspath(p)) for _, p in present]
            common_dir = os.path.commonpath(parents)
        except Exception:
            common_dir = ""

    if common_dir:
        console.print(f"[cyan]Carpeta:[/] [grey]{_format_path_for_summary(common_dir)}[/grey]")

    for label, p in present:
        filename = os.path.basename(p)
        parent = os.path.dirname(os.path.abspath(p))
        same_parent = False
        if common_dir:
            try:
                same_parent = os.path.normcase(parent) == os.path.normcase(os.path.abspath(common_dir))
            except Exception:
                same_parent = False

        if same_parent:
            console.print(f"[cyan]{label}:[/] [white]{filename}[/white]")
        else:
            parent_disp = _format_path_for_summary(parent, base_dir=common_dir or None)
            console.print(f"[cyan]{label}:[/] [white]{filename}[/white] [grey]({parent_disp})[/grey]")

    # Mostrar panel de proceso completado con hora actual
    hora_actual = datetime.now().strftime("%I:%M:%S %p")
    elapsed_line = ""
    if elapsed_seconds is not None:
        elapsed_line = f"\n[white]Tiempo total: [bold]{_format_elapsed(elapsed_seconds)}[/bold][/white]"
    mensaje = (
        "[bright_white]Proceso completado[/bright_white]\n\n"
        f"[white]Hora de finalizacion: [bold]{hora_actual}[/bold][/white]"
        f"{elapsed_line}"
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

def variations_box_style_option() -> str:
    """Elige el estilo del cuadro de variaciones (clásico o bonito)."""
    raw_env = next((os.environ.get(k) for k in VARIATIONS_BOX_STYLE_ENV_KEYS if os.environ.get(k) is not None), None)
    if raw_env is not None:
        style = normalize_variations_box_style(raw_env)
    else:
        print(Fore.CYAN + "\n¿Estilo del cuadro de variaciones (en slide de Tendencia)?")
        print(Fore.WHITE + "1 - Clásico (tabla actual)")
        print(Fore.WHITE + "2 - Bonito (tarjetas)")
        opciones = {"1": "classic", "2": "pretty", "clasico": "classic", "clásico": "classic", "bonito": "pretty"}
        eleccion = input(Fore.GREEN + "Elija 1 o 2: ").strip().lower()
        style = opciones.get(eleccion, "classic")
    SELECTIONS["Estilo variaciones"] = "Bonito" if style == "pretty" else "Clasico"
    clear_and_print_summary()
    return style

def coverage_slide_variant_option() -> str:
    """Elige el modo del slide de Cobertura (clásico o complementado)."""
    raw_env = next((os.environ.get(k) for k in COVERAGE_SLIDE_VARIANT_ENV_KEYS if os.environ.get(k) is not None), None)
    if raw_env is not None:
        variant = normalize_coverage_slide_variant(raw_env)
    else:
        print(Fore.CYAN + "\n¿Modo del slide de Cobertura?")
        print(Fore.WHITE + "1 - Clásico (tabla VAR % MAT)")
        print(Fore.WHITE + "2 - Complementado (Penetración MAT + Cobertura puntual + Estabilidad)")
        opciones = {
            "1": "classic",
            "2": "complemented",
            "clasico": "classic",
            "clásico": "classic",
            "complementado": "complemented",
            "complemented": "complemented",
        }
        eleccion = input(Fore.GREEN + "Elija 1 o 2: ").strip().lower()
        variant = opciones.get(eleccion, "classic")
    SELECTIONS["Slide Cobertura"] = "Complementado" if variant == "complemented" else "Clasico"
    clear_and_print_summary()
    return variant

def evolution_slide_variant_option() -> str:
    """Elige el modo del slide de Evolucion mensual y variacion (simple o clasico/avanzado)."""
    raw_env = next((os.environ.get(k) for k in EVOLUTION_SLIDE_VARIANT_ENV_KEYS if os.environ.get(k) is not None), None)
    if raw_env is not None:
        variant = normalize_evolution_slide_variant(raw_env)
    else:
        print(Fore.CYAN + "\n¿Modo del slide 'Evolucion mensual y variacion'?")
        print(Fore.WHITE + "1 - Simple (solo variacion: lineas, sin volumen mensual)")
        print(Fore.WHITE + "2 - Clasico/avanzado (volumen mensual + barras de variacion)")
        opciones = {
            "1": "simple",
            "2": "classic",
            "simple": "simple",
            "clasico": "classic",
            "clásico": "classic",
            "avanzado": "classic",
            "advanced": "classic",
        }
        eleccion = input(Fore.GREEN + "Elija 1 o 2: ").strip().lower()
        variant = opciones.get(eleccion, "classic")
    SELECTIONS["Slide Evolucion"] = "Simple" if variant == "simple" else "Clasico/Avanzado"
    clear_and_print_summary()
    return variant

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
        raw_sheet = excel_file_obj.parse(sheet_name, header=None)

        # Detectar inicio real de la tabla buscando "table" en la primera columna
        start_idx = 0
        meta_header_text = None
        try:
            first_col = raw_sheet.iloc[:, 0].astype(str)
            table_mask = first_col.str.contains(r"\btable\b", flags=re.IGNORECASE, na=False)
            if table_mask.any():
                start_idx = table_mask[table_mask].index[0]
                meta_header_text = raw_sheet.iloc[start_idx, 0]
        except Exception:
            start_idx = 0
            meta_header_text = None

        df_sheet = raw_sheet.iloc[start_idx:, :].reset_index(drop=True)

        # Validar estructura mínima esperada (al menos 2 filas, 8 columnas)
        rows, cols = df_sheet.shape
        if rows < 2 or cols < 8:
            if cols == 7:
                # Caso específico: 7 columnas – probablemente falta Sell-in del cliente
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
            _col8 = df_sheet.iloc[1:, 7]  # Índice 0-based: 7 es la 8ª columna
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
        df_sheet.columns = [COL_DATA, COL_SELL_OUT, COL_PENET, COL_COMPRA_MEDIA, COL_COMPRA_OCA, COL_FREQ, COL_BUYERS, COL_SELL_IN] + list(df_sheet.columns[8:])  # Mantiene columnas extra si existen
        df_sheet = df_sheet.loc[:, [COL_DATA, COL_SELL_IN, COL_SELL_OUT, COL_COMPRA_MEDIA, COL_COMPRA_OCA, COL_FREQ, COL_PENET, COL_BUYERS]]  # Reordena y selecciona

        # Elimina la primera fila (encabezados repetidos) y resetea el índice
        df_sheet = df_sheet.iloc[1:].reset_index(drop=True)

        # Convierte la columna "Data" a tipo datetime
        # Maneja posibles errores de formato o valores nulos
        original_dates = df_sheet[COL_DATA].copy()  # Guardar original por si falla
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
            df_sheet[COL_DATA] = pd.to_datetime(original_dates, errors='coerce')  # Reintentar con la original

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
        df_sheet[COL_DATA] = df_sheet[COL_DATA].dt.date  # Convertir a solo fecha al final

        return df_sheet, measure

    except Exception as e:
        print(f"{Fore.RED}Error crítico al cargar o preprocesar la hoja '{sheet_name}': {e}")
        return None, None

# --- Funciones de Generación de Gráficos ---

def generar_grafico_evolucion_mensual(
    df_graf,
    pipeline_meses: int = 0,
    lang_idx: int = 2,
    marca_nombre: Optional[str] = None,
    variant: str = "classic",
):
    """
    Genera un gráfico de evolución mensual de WP by Numerator vs Sell-in con variación interanual.

    Args:
        df_graf (pd.DataFrame): DataFrame con datos mensuales (col 'Data' debe ser datetime).
        pipeline_meses (int): Número de meses de pipeline para desplazar Sell-in.
        lang_idx (int): Identificador de idioma (impacta etiquetas).
        marca_nombre (str, opcional): Nombre de la marca para mensajes de advertencia.

    Returns:
        matplotlib.figure.Figure: Figura de matplotlib con el gráfico, o None si no hay datos.
    """
    if df_graf is None or df_graf.empty or len(df_graf) < 24: # Necesita al menos 24 meses para var YOY
        print(f"{Fore.YELLOW}Advertencia: No se puede generar gráfico de evolución mensual. Datos insuficientes (se requieren >= 24 meses).")
        return None

    variant_norm = normalize_evolution_slide_variant(variant)

    # Usar contexto de estilo para evitar afectar otros gráficos
    with matplotlib.style.context('seaborn-v0_8-whitegrid'):
        df_plot = df_graf.copy()
        df_plot[COL_DATA] = pd.to_datetime(df_plot[COL_DATA]) # Asegurar datetime
        marca_label = (marca_nombre or "N/D").strip() or "N/D"
        needs_exception_warning = False

        # Detectar '-' en valores numéricos y asegurar tipo float
        for col in (COL_SELL_IN, COL_SELL_OUT):
            col_as_str = df_plot[col].astype(str).str.strip()
            dash_mask = col_as_str.eq("-")
            if dash_mask.any():
                needs_exception_warning = True
                df_plot.loc[dash_mask, col] = 0
            df_plot[col] = pd.to_numeric(df_plot[col], errors='coerce').fillna(0)

        # Si hay pipeline, desplazar Sell-in y guardar original si es necesario
        if pipeline_meses > 0:
            # df_plot["Sell_in_original"] = df_plot[COL_SELL_IN].copy() # Descomentar si se necesita el original
            df_plot[COL_SELL_IN] = df_plot[COL_SELL_IN].shift(pipeline_meses)

        # Calcular sumas móviles y variaciones interanuales
        df_plot["Kantar_12m"] = df_plot[COL_SELL_OUT].rolling(12).sum()
        df_plot["Sellin_12m"] = df_plot[COL_SELL_IN].rolling(12).sum()
        kantar_prev = df_plot["Kantar_12m"].shift(12)
        sellin_prev = df_plot["Sellin_12m"].shift(12)
        zero_prev_kantar = kantar_prev == 0
        zero_prev_sellin = sellin_prev == 0
        if zero_prev_kantar.any() or zero_prev_sellin.any():
            needs_exception_warning = True
        safe_kantar_prev = kantar_prev.where(~zero_prev_kantar, 1)
        safe_sellin_prev = sellin_prev.where(~zero_prev_sellin, 1)
        df_plot["Kantar_yoy"] = ((df_plot["Kantar_12m"] / safe_kantar_prev) - 1) * 100
        df_plot["Sellin_yoy"] = ((df_plot["Sellin_12m"] / safe_sellin_prev) - 1) * 100

        # Filtrar NaNs resultantes de rolling/shift
        df_plot = df_plot.dropna(subset=["Kantar_yoy", "Sellin_yoy"]).copy()

        if df_plot.empty:
            print(f"{Fore.YELLOW}Advertencia: No quedan datos para el gráfico de evolución después de calcular YOY.")
            return None

        if needs_exception_warning:
            notify_zero_months_exception(marca_label)

        # Crear figura y ejes con márgenes personalizados
        fig = plt.figure(figsize=(16.5, 8), dpi=100) # Ajustar tamaño si es necesario
        left_margin, right_margin, bottom_margin, top_margin = 0.08, 0.92, 0.18, 0.90
        ax1 = fig.add_axes([left_margin, bottom_margin, right_margin-left_margin, top_margin-bottom_margin])
        ax2 = None
        if variant_norm == "classic":
            ax2 = ax1.twinx()

        var_title = "Variacion Interanual (%)" if lang_idx != 3 else "Year-over-Year Change (%)"
        def _tint_color(color_str: str, mix_with_white: float = 0.78) -> str:
            """Devuelve una version mas clara del color (mezclado con blanco)."""
            try:
                from matplotlib.colors import to_rgb, to_hex
                r, g, b = to_rgb(color_str)
                m = float(mix_with_white)
                if m < 0:
                    m = 0.0
                if m > 1:
                    m = 1.0
                tinted = (r + (1 - r) * m, g + (1 - g) * m, b + (1 - b) * m)
                return to_hex(tinted)
            except Exception:
                return "#E7E6E6"

        def _font_color(col_yoy: str, value: float) -> str:
            if value > 0:
                return COLOR_POS_LABEL if col_yoy == "Kantar_yoy" else COLOR_POS_LABEL_ALT
            if value < 0:
                return COLOR_NEG_LABEL if col_yoy == "Kantar_yoy" else COLOR_NEG_LABEL_ALT
            return "#333333"
        if variant_norm == "simple":
            # Simple: solo variacion (lineas), sin volumen mensual.
            ax1.plot(
                df_plot[COL_DATA],
                df_plot["Kantar_yoy"],
                color=COLOR_KANTAR_LINE,
                marker="o",
                linewidth=2.5,
                markersize=5,
                label="% Var Worldpanel by Numerator",
            )
            ax1.plot(
                df_plot[COL_DATA],
                df_plot["Sellin_yoy"],
                color=COLOR_SELLIN_LINE,
                marker="o",
                linewidth=2.5,
                markersize=5,
                label="% Var Sell-in" + (f" - P:{pipeline_meses}" if pipeline_meses > 0 else ""),
            )
            ax1.set_ylabel(var_title, fontsize=11, labelpad=15)
            ax1.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
            ax1.tick_params(axis='y', labelsize=9)
            ax1.axhline(y=0, color='gray', linestyle='-', alpha=0.5, linewidth=0.8)
            ax1.grid(axis='y', linestyle='--', alpha=0.4)

            offset = 4
            for _, row in df_plot.iterrows():
                for col_yoy, x_offset in [("Kantar_yoy", -offset), ("Sellin_yoy", offset)]:
                    if pd.isna(row[col_yoy]):
                        continue
                    valor = float(row[col_yoy])
                    pos_vert = valor + 1 if valor >= 0 else valor - 1
                    va_align = "bottom" if valor >= 0 else "top"
                    line_color = COLOR_KANTAR_LINE if col_yoy == "Kantar_yoy" else COLOR_SELLIN_LINE
                    bg = _tint_color(line_color, mix_with_white=0.78)
                    ax1.text(
                        row[COL_DATA] + pd.Timedelta(days=x_offset),
                        pos_vert,
                        f"{valor:.1f}%",
                        ha="center",
                        va=va_align,
                        fontsize=7,
                        fontweight="bold",
                        color=_font_color(col_yoy, valor),
                        bbox=dict(boxstyle="round,pad=0.18", facecolor=bg, edgecolor=line_color, linewidth=0.8),
                    )

            # Ajustar limites Y para dar aire a las cajas
            y_min, y_max = ax1.get_ylim()
            pad = max(abs(y_min), abs(y_max)) * 0.18
            ax1.set_ylim(y_min - pad, y_max + pad)
        else:
            # Clasico/avanzado: volumen mensual (lineas) + variacion (barras)
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

            width = 8
            offset = 4
            assert ax2 is not None
            ax2.bar(df_plot[COL_DATA] - pd.DateOffset(days=offset), df_plot["Kantar_yoy"], width=width, color=COLOR_KANTAR_BAR_VAR, edgecolor=COLOR_KANTAR_EDGE_VAR, alpha=0.7, label="% Var Worldpanel by Numerator")
            ax2.bar(df_plot[COL_DATA] + pd.DateOffset(days=offset), df_plot["Sellin_yoy"], width=width, color=COLOR_SELLIN_BAR_VAR, edgecolor=COLOR_SELLIN_EDGE_VAR, alpha=0.7, label="% Var Sell-in")
            ax2.set_ylabel(var_title, fontsize=11, labelpad=15)
            ax2.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
            ax2.tick_params(axis='y', labelsize=9)
            ax2.axhline(y=0, color='gray', linestyle='-', alpha=0.5, linewidth=0.8)

            # Etiquetas en barras con cuadro de fondo segun signo
            for _, row in df_plot.iterrows():
                for col_yoy, x_offset in [("Kantar_yoy", -offset), ("Sellin_yoy", offset)]:
                    if pd.isna(row[col_yoy]):
                        continue
                    valor = float(row[col_yoy])
                    pos_vert = valor + 1 if valor >= 0 else valor - 1
                    va_align = "bottom" if valor >= 0 else "top"
                    bg = "#C6EFCE" if valor > 0 else ("#FFC7CE" if valor < 0 else "#E7E6E6")
                    ax2.text(
                        row[COL_DATA] + pd.Timedelta(days=x_offset),
                        pos_vert,
                        f"{valor:.1f}%",
                        ha="center",
                        va=va_align,
                        fontsize=7,
                        fontweight="bold",
                        color=_font_color(col_yoy, valor),
                        bbox=dict(boxstyle="round,pad=0.18", facecolor=bg, edgecolor="black", linewidth=0.6),
                    )

            y2_min, y2_max = ax2.get_ylim()
            padding = max(abs(y2_min), abs(y2_max)) * 0.15
            ax2.set_ylim(y2_min - padding, y2_max + padding*2)

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
        if variant_norm == "classic":
            lines1, labels1 = ax1.get_legend_handles_labels()
            assert ax2 is not None
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left", bbox_to_anchor=(0.01, 0.98), fontsize=9, frameon=True, framealpha=0.8)
        else:
            ax1.legend(loc="upper left", bbox_to_anchor=(0.01, 0.98), fontsize=9, frameon=True, framealpha=0.8)

        # No usar tight_layout con add_axes, márgenes manuales ya aplicados
        # fig.tight_layout(rect=[0, 0, 1, 0.95]) # Ajustar rect si el título se solapa
  
        return fig

def generar_grafico_cobertura(slide, marca_clean, pipeline, df_cov_pipe, df_pen_pipe, lang_idx, coverage_label, labels_dict):
    """Genera el gráfico de barras de Cobertura vs Penetración y lo añade al slide."""
    cov_series = df_cov_pipe if isinstance(df_cov_pipe, pd.Series) else pd.Series(df_cov_pipe)
    pen_series = df_pen_pipe if isinstance(df_pen_pipe, pd.Series) else pd.Series(df_pen_pipe)
    cov_series = cov_series.rename('coverage')
    pen_series = pen_series.rename('penetracion')
    cov_series = pd.to_numeric(cov_series, errors='coerce')
    pen_series = pd.to_numeric(pen_series, errors='coerce')
    combined = pd.concat([cov_series, pen_series], axis=1, join='inner')
    combined = combined.replace([np.inf, -np.inf], np.nan)
    combined = combined.dropna(subset=['coverage', 'penetracion'])
    if combined.empty:
        print(f"{Fore.YELLOW}Advertencia: No hay datos suficientes para el gráfico de cobertura/penetración (Marca: {marca_clean}, P:{pipeline}).")
        return
    cov_data = combined['coverage'].to_numpy(dtype=float)
    pen_data = combined['penetracion'].to_numpy(dtype=float)
    cov_data = np.where(np.isfinite(cov_data), cov_data, np.nan)
    pen_data = np.where(np.isfinite(pen_data), pen_data, np.nan)
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
                    # Resaltar el último mes y cada 12 meses hacia atrás (evita el caso len%12==0).
                    if i % 12 == ((len(rect_group) - 1) % 12):
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

def generar_grafico_tendencia(
    slide,
    marca_clean,
    pipeline,
    df_plot,
    lang_idx,
    labels_dict,
    doble_eje: bool = False,
    box_left=None,
    box_top=None,
    box_width=None,
    box_height=None,
    figsize: Tuple[float, float] = (13, 5),
    legend_y: float = -0.28,
):
    """
    Genera el gráfico de líneas de Tendencia (Sell-in vs Sell-out) y lo añade al slide.
    Si doble_eje=True, WP by Numerator (Sell-out) va en eje secundario.
    """
    if df_plot is None or df_plot.empty or pipeline >= len(df_plot):
         print(f"{Fore.YELLOW}Advertencia: Datos insuficientes para gráfico de Tendencia (Marca: {marca_clean}, P:{pipeline}).")
         return

    fig_trend, ax_trend = plt.subplots(figsize=figsize, dpi=100)

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
        ax2.legend(lns, labs, loc='lower center', bbox_to_anchor=(0.5, legend_y), frameon=False, prop={'size': 11}, ncol=2)
    else:
        lns1 = ax_trend.plot(x_labels, sell_in_data, color=COLOR_SELLIN_TREND_LINE, linewidth=4, label=f'{COL_SELL_IN} (P:{pipeline})')
        lns2 = ax_trend.plot(x_labels, sell_out_data, color=COLOR_SELLOUT_TREND_LINE, linewidth=4, label=COL_SELL_OUT)
        ax_trend.set_ylabel(f'{COL_SELL_IN} / {COL_SELL_OUT}', color='black', fontsize=11)
        ax_trend.set_ylim(bottom=0)
        lns = lns1 + lns2
        labs = [l.get_label() for l in lns]
        ax_trend.legend(lns, labs, loc='lower center', bbox_to_anchor=(0.5, legend_y), frameon=False, prop={'size': 11}, ncol=2)

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
    # Sin contorno: el slide ya maneja el layout, y el borde negro se ve pesado.
    img_stream_bordered = io.BytesIO()
    img_pil.save(img_stream_bordered, format='PNG')
    img_stream_bordered.seek(0)
    # Ubicación por defecto (layout clásico)
    if box_left is None:
        box_left = Inches(0.5)
    if box_top is None:
        box_top = Inches(1.8)

    # Si no se pasa un "box", mantenemos el comportamiento anterior (solo altura).
    if box_width is None and box_height is None:
        slide.shapes.add_picture(img_stream_bordered, box_left, box_top, height=Inches(4.5))
        plt.close(fig_trend)
        return

    # Fit dentro del rectángulo manteniendo aspect ratio.
    if box_width is None:
        box_width = Inches(100)  # sin limite real, se acota por altura
    if box_height is None:
        box_height = Inches(100)  # sin limite real, se acota por ancho

    img_stream_bordered.seek(0)
    try:
        with Image.open(img_stream_bordered) as _img:
            px_w, px_h = _img.size
    except Exception:
        px_w, px_h = (1, 1)
    finally:
        img_stream_bordered.seek(0)

    aspect = (px_w / px_h) if px_h else 1.0
    box_w_in = float(box_width) / 914400.0
    box_h_in = float(box_height) / 914400.0
    placed_w_in = box_w_in
    placed_h_in = placed_w_in / aspect if aspect else box_h_in
    if placed_h_in > box_h_in:
        placed_h_in = box_h_in
        placed_w_in = placed_h_in * aspect

    placed_w = Inches(max(0.1, placed_w_in))
    placed_h = Inches(max(0.1, placed_h_in))
    left = box_left + int((box_width - placed_w) / 2)
    top = box_top + int((box_height - placed_h) / 2)
    slide.shapes.add_picture(img_stream_bordered, left, top, width=placed_w, height=placed_h)
    plt.close(fig_trend)
    

# --- Configuración y estructuras de alto nivel --------------------------------

@dataclass
class ExecutionOptions:
    coverage_type: str
    coverage_reason: str
    trend_axis: str
    include_english: bool
    round_coverage: bool
    variations_box_style: str = "classic"
    coverage_slide_variant: str = "classic"
    evolution_slide_variant: str = "classic"
    summary_extra_months: List[int] = field(default_factory=list)
    summary_extra_months_mode: str = "recent"
    auto_mode: bool = False

    @classmethod
    def from_environment(cls) -> Optional["ExecutionOptions"]:
        """Crea las opciones cuando se usa la ejecución en modo automático."""
        auto_file = os.environ.get("AUTO_FILE")
        if not auto_file:
            return None
        coverage_type = os.environ.get("AUTO_COV_TYPE", "Absoluta")
        variations_box_style = normalize_variations_box_style(
            next((os.environ.get(k) for k in VARIATIONS_BOX_STYLE_ENV_KEYS if os.environ.get(k) is not None), None)
        )
        coverage_slide_variant = normalize_coverage_slide_variant(
            next((os.environ.get(k) for k in COVERAGE_SLIDE_VARIANT_ENV_KEYS if os.environ.get(k) is not None), None)
        )
        evolution_slide_variant = normalize_evolution_slide_variant(
            next((os.environ.get(k) for k in EVOLUTION_SLIDE_VARIANT_ENV_KEYS if os.environ.get(k) is not None), None)
        )
        summary_extra_months = get_summary_extra_months_from_env()
        summary_extra_months_mode = get_summary_extra_months_mode_from_env() or "recent"
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
            variations_box_style=variations_box_style,
            include_english=include_english,
            round_coverage=round_cov,
            coverage_slide_variant=coverage_slide_variant,
            evolution_slide_variant=evolution_slide_variant,
            summary_extra_months=summary_extra_months,
            summary_extra_months_mode=summary_extra_months_mode,
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
    buyers_mat_actual: Optional[float] = None
    penet_mat_actual: Optional[float] = None
    penet_mat_anterior: Optional[float] = None


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


def build_summary_coverage_periods(
    ref_dt: datetime,
    summary_extra_months: Sequence[int],
    extra_months_mode: str,
) -> Tuple[List[datetime], List[datetime], datetime, datetime]:
    """Arma los periodos de cobertura ordenados e identifica los extras."""
    months_to_compare: List[int] = []
    for month_num in summary_extra_months:
        if 1 <= int(month_num) <= 12 and int(month_num) != ref_dt.month and int(month_num) not in months_to_compare:
            months_to_compare.append(int(month_num))

    base_prev = datetime(ref_dt.year - 1, ref_dt.month, 1)
    base_curr = datetime(ref_dt.year, ref_dt.month, 1)

    extras_prev = [datetime(ref_dt.year - 1, month_num, 1) for month_num in months_to_compare] if extra_months_mode == "both" else []
    extras_curr = [datetime(ref_dt.year, month_num, 1) for month_num in months_to_compare]

    if extra_months_mode == "both":
        ordered_periods = extras_prev + [base_prev] + extras_curr + [base_curr]
        extra_periods = extras_prev + extras_curr
    else:
        ordered_periods = [base_prev] + extras_curr + [base_curr]
        extra_periods = extras_curr

    return ordered_periods, extra_periods, base_prev, base_curr

def build_summary_columns(
    lang_index: int,
    fabricante: str,
    ref_dt: datetime,
    summary_extra_months: Sequence[int],
    summary_extra_months_mode: str,
) -> Tuple[List[str], List[datetime], List[str]]:
    coverage_periods, extra_periods, _, _ = build_summary_coverage_periods(
        ref_dt,
        summary_extra_months,
        summary_extra_months_mode,
    )
    summary_base_columns: Dict[int, List[str]] = {
        1: [
            "Fabricante/Marca",
            "Pipeline",
            "Penetração Média Mensal",
            fabricante,
            "Worldpanel by Numerator",
        ],
        2: [
            "Fabricante/Marca",
            "Pipeline",
            "Penetración Media Mensual",
            f"%VAR {fabricante}",
            "% VAR Worldpanel by Numerator",
        ],
        3: [
            "Manufacturer/Brand",
            "Pipeline",
            "Monthly Avg Penetration",
            f"%VAR {fabricante}",
            "% VAR Worldpanel by Numerator",
        ],
    }
    coverage_prefix = "Coverage" if lang_index == 3 else "Cobertura"
    stability_label = {1: "Estabilidade", 2: "Estabilidad", 3: "Stability"}[lang_index]
    summary_columns = list(summary_base_columns[lang_index])
    for period_dt in coverage_periods:
        summary_columns.append(f"{coverage_prefix} {period_dt.strftime('%b-%y')}")
    summary_columns.append(stability_label)
    extra_columns = [f"{coverage_prefix} {period_dt.strftime('%b-%y')}" for period_dt in extra_periods]
    return summary_columns, coverage_periods, extra_columns

def build_labels(
    lang_index: int,
    fabricante: str,
    ref_month_year: str,
    summary_extra_months: Optional[Sequence[int]] = None,
    summary_extra_months_mode: str = "recent",
) -> Dict[Tuple[int, str], List[str] | str]:
    """Reproduce el diccionario de etiquetas usado por el script original."""
    ref_dt = dt.strptime(ref_month_year, "%m-%y")
    extra_months = list(summary_extra_months or [])
    summary_pt, _, extra_cols_pt = build_summary_columns(1, fabricante, ref_dt, extra_months, summary_extra_months_mode)
    summary_es, _, extra_cols_es = build_summary_columns(2, fabricante, ref_dt, extra_months, summary_extra_months_mode)
    summary_en, _, extra_cols_en = build_summary_columns(3, fabricante, ref_dt, extra_months, summary_extra_months_mode)

    return {
        (1, "S1"): " ",
        (1, "Summary"): summary_pt,
        (1, "SummaryExtraCoverageCols"): extra_cols_pt,
        (1, "Graf cob Penet Men"): "Penetração Mensal",
        (1, "Titulo Cob"): "Cobertura em Ano Móvel",
        (1, "Var"): "com",
        (1, "Titulo Vol"): "Tendência em Volumen",
        (2, "S1"): " ",
        (2, "Summary"): summary_es,
        (2, "SummaryExtraCoverageCols"): extra_cols_es,
        (2, "Graf cob Penet Men"): "Penetración Mensual",
        (2, "Titulo Cob"): "Cobertura en Año Móvil",
        (2, "Var"): "con",
        (2, "Titulo Vol"): "Tendencia en Volumen",
        (3, "S1"): " ",
        (3, "Summary"): summary_en,
        (3, "SummaryExtraCoverageCols"): extra_cols_en,
        (3, "Graf cob Penet Men"): "PENETRATION BY PERIOD",
        (3, "Titulo Cob"): "MOVING YEAR COVERAGE",
        (3, "Var"): "with",
        (3, "Titulo Vol"): "TREND IN VOLUME",
        (1, "LowPenFooter"): "Marca de baixa penetração (<200 compradores) - Resultados para uso interno",
        (1, "LowPenFooterPlural"): "Marcas de baixa penetração (<200 compradores) - Resultados para uso interno",
        (1, "LowPenSummarySingular"): "O estudo contém 1 marca de baixa penetração (<200 buyers). Resultados para uso interno",
        (1, "LowPenSummaryPlural"): "O estudo contém {n} marcas de baixa penetração (<200 buyers). Resultados para uso interno",
        (2, "LowPenFooter"): "Marca de baja penetración (<200 compradores) - Resultados para uso interno",
        (2, "LowPenFooterPlural"): "Marcas de baja penetración (<200 compradores) - Resultados para uso interno",
        (2, "LowPenSummarySingular"): "El estudio contiene 1 marca de baja penetración (<200 buyers). Resultados para uso interno",
        (2, "LowPenSummaryPlural"): "El estudio contiene {n} marcas de baja penetración (<200 buyers). Resultados para uso interno",
        (3, "LowPenFooter"): "Low penetration brand (<200 buyers) - For internal use only",
        (3, "LowPenFooterPlural"): "Low penetration brands (<200 buyers) - For internal use only",
        (3, "LowPenSummarySingular"): "This study contains 1 low penetration brand (<200 buyers). For internal use only",
        (3, "LowPenSummaryPlural"): "This study contains {n} low penetration brands (<200 buyers). For internal use only",
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
        coverage_type: str,
        ref_month_year: str,
        tipo_eje_tend: str,
        variations_box_style: str = "classic",
        coverage_slide_variant: str = "classic",
    ) -> None:
        self.ppt = presentation
        self.lang_index = lang_index
        self.labels = labels
        self.coverage_label = coverage_label
        self.coverage_type = coverage_type
        self.ref_month_year = ref_month_year
        self.tipo_eje_tend = tipo_eje_tend
        self.variations_box_style = normalize_variations_box_style(variations_box_style)
        self.coverage_slide_variant = normalize_coverage_slide_variant(coverage_slide_variant)

    def _add_picture_fit(
        self,
        slide,
        img_stream: io.BytesIO,
        *,
        left,
        top,
        width,
        height,
        halign: str = "center",  # left|center|right
        valign: str = "center",  # top|center|bottom
    ) -> None:
        """Inserta una imagen en un rectángulo (fit) manteniendo aspect ratio.

        Nota: por defecto centra, pero para el header de cobertura se usa anclaje
        a izquierda/derecha para alinear con el gráfico de coberturas.
        """
        img_stream.seek(0)
        try:
            with Image.open(img_stream) as _img:
                px_w, px_h = _img.size
        except Exception:
            px_w, px_h = (1, 1)
        finally:
            img_stream.seek(0)

        aspect = (px_w / px_h) if px_h else 1.0
        box_w_in = float(width) / 914400.0
        box_h_in = float(height) / 914400.0
        placed_w_in = box_w_in
        placed_h_in = placed_w_in / aspect if aspect else box_h_in
        if placed_h_in > box_h_in:
            placed_h_in = box_h_in
            placed_w_in = placed_h_in * aspect

        placed_w = Inches(max(0.1, placed_w_in))
        placed_h = Inches(max(0.1, placed_h_in))

        # Horizontal alignment inside the box
        _h = (halign or "center").strip().lower()
        if _h == "left":
            left2 = left
        elif _h == "right":
            left2 = left + int(width - placed_w)
        else:
            left2 = left + int((width - placed_w) / 2)

        # Vertical alignment inside the box
        _v = (valign or "center").strip().lower()
        if _v == "top":
            top2 = top
        elif _v == "bottom":
            top2 = top + int(height - placed_h)
        else:
            top2 = top + int((height - placed_h) / 2)
        slide.shapes.add_picture(img_stream, left2, top2, width=placed_w, height=placed_h)

    def _month_abbr(self, month: int) -> str:
        # Abreviaciones locales en mayúsculas (para el cuadro "bonito").
        es = ["", "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
        pt = ["", "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
        en = ["", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        table = en if self.lang_index == 3 else (pt if self.lang_index == 1 else es)
        if 1 <= int(month) <= 12:
            return table[int(month)]
        return "-"

    def _coverage_metric_title(self) -> str:
        ctype = (self.coverage_type or "").strip().lower()
        if ctype == "auto":
            ctype = "absoluta"
        if self.lang_index == 3:
            return "Absolute Coverage" if ctype == "absoluta" else "Relative Coverage"
        return "Cobertura Absoluta" if ctype == "absoluta" else "Cobertura Relativa"

    def _pen_table_headers(self) -> Tuple[str, str]:
        if self.lang_index == 1:
            return "Ano", "Penetração\nMédia Mensal"
        if self.lang_index == 3:
            return "Year", "Monthly Avg\nPenetration"
        return "Año", "Penetración\nMedia Mensual"

    def _stability_label(self) -> str:
        return {1: "Estabilidade", 2: "Estabilidad", 3: "Stability"}[self.lang_index]

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> "RGBColor":
        raw = str(hex_color or "").strip().lstrip("#")
        if len(raw) != 6:
            return RGBColor(0, 0, 0)
        try:
            return RGBColor(int(raw[0:2], 16), int(raw[2:4], 16), int(raw[4:6], 16))
        except Exception:
            return RGBColor(0, 0, 0)

    def _set_table_cell_text(
        self,
        cell,
        text: object,
        *,
        fill_color: Optional["RGBColor"] = None,
        font_color: Optional["RGBColor"] = None,
        font_size: int = 12,
        bold: bool = True,
        align: int = 2,
        word_wrap: bool = True,
    ) -> None:
        if fill_color is not None:
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill_color

        tf = cell.text_frame
        tf.clear()
        tf.word_wrap = bool(word_wrap)
        tf.margin_left = Pt(2)
        tf.margin_right = Pt(2)
        tf.margin_top = Pt(2)
        tf.margin_bottom = Pt(2)

        p = tf.paragraphs[0]
        p.text = "" if text is None else str(text)
        p.alignment = align
        p.font.bold = bool(bold)
        p.font.size = Pt(font_size)
        p.font.color.rgb = font_color if font_color is not None else RGBColor(0, 0, 0)

    @staticmethod
    def _normalize_summary_table_value(value: object) -> str:
        if value is None:
            return "-"
        try:
            if "pd" in globals() and pd.isna(value):
                return "-"
        except Exception:
            pass
        txt = str(value).strip()
        return txt if txt else "-"

    @staticmethod
    def _parse_summary_percent_value(value: object) -> Optional[float]:
        txt = SlideBuilder._normalize_summary_table_value(value)
        if txt in {"-", ""}:
            return None
        txt = txt.replace("%", "").replace(",", ".").strip()
        try:
            return float(txt)
        except Exception:
            return None

    @staticmethod
    def _normalize_brand_key(value: object) -> str:
        txt = SlideBuilder._normalize_summary_table_value(value)
        if txt == "-":
            return ""
        txt = unicodedata.normalize("NFD", txt)
        txt = "".join(ch for ch in txt if unicodedata.category(ch) != "Mn")
        txt = re.sub(r"\s+", " ", txt).strip().lower()
        return txt

    @classmethod
    def _summary_row_fails_robustness(cls, row_values: Sequence[object], cols: int) -> bool:
        # Robustez: misma tendencia entre %VAR cliente y %VAR WP by Numerator.
        if cols < 5:
            return False
        var_cliente = cls._parse_summary_percent_value(row_values[3])
        var_wp = cls._parse_summary_percent_value(row_values[4])
        if var_cliente is None or var_wp is None:
            return False
        return not ((var_cliente * var_wp) > 0 or (var_cliente == 0 and var_wp == 0))

    def _add_editable_summary_table(
        self,
        slide,
        df_summary: "pd.DataFrame",
        *,
        left,
        top,
        width,
        max_height,
        low_penetration_brands: Optional[Sequence[str]] = None,
    ) -> None:
        if df_summary is None or df_summary.empty:
            return

        rows = int(len(df_summary.index)) + 1
        cols = int(len(df_summary.columns))
        if rows <= 1 or cols <= 0:
            return

        max_height = int(max_height)
        if max_height <= 0:
            max_height = int(Inches(4.8))

        header_h = int(Inches(0.34))
        body_rows = max(rows - 1, 0)
        preferred_body_h = int(Inches(0.25))
        min_body_h = int(Inches(0.17))

        if body_rows > 0:
            needed_h = header_h + (body_rows * preferred_body_h)
            if needed_h <= max_height:
                body_h = preferred_body_h
            else:
                body_h = max(min_body_h, int((max_height - header_h) / body_rows))
                needed_h = header_h + (body_rows * body_h)
        else:
            body_h = 0
            needed_h = header_h

        table_shape = slide.shapes.add_table(rows, cols, left, top, width, needed_h)
        table = table_shape.table

        col_weights: List[int] = []
        for col_name in df_summary.columns:
            sample_values = df_summary[col_name].head(15).tolist()
            max_cell_len = max((len(self._normalize_summary_table_value(v)) for v in sample_values), default=0)
            header_len = len(str(col_name))
            weight = max(8, min(40, max(max_cell_len, header_len)))
            col_weights.append(weight)
        weight_sum = sum(col_weights) if col_weights else cols

        width_assigned = 0
        for idx, weight in enumerate(col_weights):
            if idx == cols - 1:
                col_w = int(width - width_assigned)
            else:
                col_w = int(width * (float(weight) / float(weight_sum)))
                width_assigned += col_w
            table.columns[idx].width = max(col_w, int(width * 0.04))

        table.rows[0].height = header_h
        for r in range(1, rows):
            table.rows[r].height = body_h

        header_bg = RGBColor(217, 225, 242)
        stripe_bg = RGBColor(245, 247, 251)
        white_bg = RGBColor(255, 255, 255)
        soft_red_bg = RGBColor(255, 235, 235)
        black = RGBColor(0, 0, 0)

        if rows <= 9:
            body_font_size = 10
        elif rows <= 14:
            body_font_size = 9
        else:
            body_font_size = 8
        low_penetration_keys: Set[str] = set()
        for brand in (low_penetration_brands or []):
            key = self._normalize_brand_key(brand)
            if key:
                low_penetration_keys.add(key)

        for c, col_name in enumerate(df_summary.columns):
            self._set_table_cell_text(
                table.cell(0, c),
                str(col_name),
                fill_color=header_bg,
                font_color=black,
                font_size=10,
                bold=True,
                align=2,
                word_wrap=True,
            )

        for r, row_values in enumerate(df_summary.itertuples(index=False), start=1):
            brand_key = self._normalize_brand_key(row_values[0]) if cols > 0 else ""
            if low_penetration_keys:
                row_fails_robustness = brand_key in low_penetration_keys
            else:
                row_fails_robustness = self._summary_row_fails_robustness(row_values, cols)
            for c in range(cols):
                val = self._normalize_summary_table_value(row_values[c])
                align = 1 if c == 0 else 2
                fill = soft_red_bg if row_fails_robustness else (stripe_bg if r % 2 == 0 else white_bg)
                self._set_table_cell_text(
                    table.cell(r, c),
                    val,
                    fill_color=fill,
                    font_color=black,
                    font_size=body_font_size,
                    bold=False,
                    align=align,
                    word_wrap=False,
                )

    def _add_penetration_header_table_shape(
        self,
        slide,
        *,
        left,
        top,
        width,
        height,
        year_header: str,
        pen_header: str,
        rows: Sequence[Tuple[str, str]],
    ) -> None:
        table_shape = slide.shapes.add_table(3, 2, left, top, width, height)
        table_shape.height = height
        table = table_shape.table
        table.columns[0].width = int(width * 0.40)
        table.columns[1].width = width - table.columns[0].width
        table.rows[0].height = int(height * 0.40)
        table.rows[1].height = int((height - table.rows[0].height) / 2)
        table.rows[2].height = height - table.rows[0].height - table.rows[1].height

        header_bg = RGBColor(0, 0, 0)
        body_bg = self._hex_to_rgb("#D9D9D9")
        white = RGBColor(255, 255, 255)
        black = RGBColor(0, 0, 0)

        row_values: List[Tuple[str, str]] = list(rows[:2])
        while len(row_values) < 2:
            row_values.append(("-", "-"))

        self._set_table_cell_text(table.cell(0, 0), year_header, fill_color=header_bg, font_color=white, font_size=10, align=2)
        self._set_table_cell_text(
            table.cell(0, 1),
            pen_header,
            fill_color=header_bg,
            font_color=white,
            font_size=10,
            align=2,
            word_wrap=False,
        )

        self._set_table_cell_text(table.cell(1, 0), row_values[0][0], fill_color=body_bg, font_color=black, align=1)
        self._set_table_cell_text(table.cell(1, 1), row_values[0][1], fill_color=body_bg, font_color=black, align=3)
        self._set_table_cell_text(table.cell(2, 0), row_values[1][0], fill_color=body_bg, font_color=black, align=1)
        self._set_table_cell_text(table.cell(2, 1), row_values[1][1], fill_color=body_bg, font_color=black, align=3)

    def _add_coverage_stability_header_table_shape(
        self,
        slide,
        *,
        left,
        top,
        width,
        height,
        cov_title: str,
        prev_label: str,
        curr_label: str,
        stability_label: str,
        cov_prev_txt: str,
        cov_curr_txt: str,
        stability_txt: str,
    ) -> None:
        table_shape = slide.shapes.add_table(3, 3, left, top, width, height)
        table_shape.height = height
        table = table_shape.table
        table.columns[0].width = int(width * 0.34)
        table.columns[1].width = int(width * 0.34)
        table.columns[2].width = width - table.columns[0].width - table.columns[1].width
        table.rows[0].height = int(height * 0.34)
        table.rows[1].height = int(height * 0.28)
        table.rows[2].height = height - table.rows[0].height - table.rows[1].height

        header_bg = self._hex_to_rgb("#355D6C")
        white = RGBColor(255, 255, 255)
        black = RGBColor(0, 0, 0)
        white_bg = RGBColor(255, 255, 255)

        table.cell(0, 0).merge(table.cell(0, 1))
        table.cell(0, 2).merge(table.cell(1, 2))

        self._set_table_cell_text(table.cell(0, 0), cov_title, fill_color=header_bg, font_color=white, font_size=11, align=2)
        self._set_table_cell_text(table.cell(0, 2), stability_label, fill_color=header_bg, font_color=white, font_size=11, align=2)
        self._set_table_cell_text(table.cell(1, 0), prev_label, fill_color=header_bg, font_color=white, font_size=11, align=2)
        self._set_table_cell_text(table.cell(1, 1), curr_label, fill_color=header_bg, font_color=white, font_size=11, align=2)

        self._set_table_cell_text(table.cell(2, 0), cov_prev_txt, fill_color=white_bg, font_color=black, align=2)
        self._set_table_cell_text(table.cell(2, 1), cov_curr_txt, fill_color=white_bg, font_color=black, align=2)
        self._set_table_cell_text(table.cell(2, 2), stability_txt, fill_color=white_bg, font_color=black, align=2)

    def _add_cov_slide_header_boxes(self, slide, assets: PipelineAssets) -> None:
        """Header del slide de Cobertura en modo 'complemented'."""
        try:
            ref_dt = dt.strptime(self.ref_month_year, "%m-%y")
        except Exception:
            # Fallback: usar el último periodo disponible del índice si hay fechas.
            idx = pd.to_datetime(getattr(assets.coverage_series, "index", []), errors="coerce")
            idx = idx[~idx.isna()]
            if len(idx) == 0:
                return
            ref_dt = idx.max().to_pydatetime()

        prev_dt = ref_dt - pd.DateOffset(months=12)
        curr_label = f"{self._month_abbr(ref_dt.month)}-{ref_dt.year % 100:02d}"
        prev_label = f"{self._month_abbr(prev_dt.month)}-{prev_dt.year % 100:02d}"

        # --- Penetración (MAT actual vs anterior) ---
        year_col, pen_col = self._pen_table_headers()
        mat_curr = f"MAT {curr_label}"
        mat_prev = f"MAT {prev_label}"
        pen_curr = assets.penet_mat_actual
        pen_prev = assets.penet_mat_anterior
        pen_curr_txt = f"{float(pen_curr):.1f}" if (pen_curr is not None and pd.notna(pen_curr)) else "-"
        pen_prev_txt = f"{float(pen_prev):.1f}" if (pen_prev is not None and pd.notna(pen_prev)) else "-"

        # --- Cobertura puntual + estabilidad ---
        cov_title = self._coverage_metric_title()
        stability_label = self._stability_label()
        cov_prev = _coverage_value_for_year_month(assets.coverage_series, int(prev_dt.year), int(prev_dt.month))
        cov_curr = _coverage_value_for_year_month(assets.coverage_series, int(ref_dt.year), int(ref_dt.month))

        def _fmt_cov(v: float) -> str:
            if v is None or pd.isna(v):
                return "-"
            return str(int(np.floor(float(v) + 0.5))) if globals().get("ROUND_COVERAGE", False) else f"{float(v):.1f}"

        stability_txt = "-"
        if cov_prev is not None and cov_curr is not None and pd.notna(cov_prev) and pd.notna(cov_curr):
            if globals().get("ROUND_COVERAGE", False):
                stability_txt = str(int(np.floor(float(cov_curr) + 0.5)) - int(np.floor(float(cov_prev) + 0.5)))
            else:
                stability_txt = f"{(float(cov_curr) - float(cov_prev)):.1f}"

        # Construye tablas nativas de PowerPoint para que el texto sea editable.

        # --- Layout superior: dos cajas en una banda encima del gráfico ---
        # Alinear con el inicio/fin del gráfico de coberturas (que arranca en x=0.5in).
        top = Inches(0.95)
        chart_left = Inches(0.5)
        chart_right = self.ppt.slide_width - Inches(0.5)
        shared_h = Inches(0.90)
        # Cuadros mas angostos, manteniendo posiciones originales:
        # penetracion a la izquierda y cobertura a la derecha.
        total_w = chart_right - chart_left
        left_w = int(total_w * 0.27)
        right_w = int(total_w * 0.37)
        right_left = chart_right - right_w
        pen_top = top - Inches(0.03)

        self._add_penetration_header_table_shape(
            slide,
            left=chart_left,
            top=pen_top,
            width=left_w,
            height=shared_h,
            year_header=year_col,
            pen_header=pen_col,
            rows=[(mat_curr, pen_curr_txt), (mat_prev, pen_prev_txt)],
        )
        self._add_coverage_stability_header_table_shape(
            slide,
            left=right_left,
            top=top,
            width=right_w,
            height=shared_h,
            cov_title=cov_title,
            prev_label=prev_label,
            curr_label=curr_label,
            stability_label=stability_label,
            cov_prev_txt=_fmt_cov(cov_prev),
            cov_curr_txt=_fmt_cov(cov_curr),
            stability_txt=stability_txt,
        )

    @staticmethod
    def _date_minus_months(year: int, month: int, delta: int) -> Tuple[int, int]:
        total = year * 12 + (month - 1) - int(delta)
        y2 = total // 12
        m2 = (total % 12) + 1
        return int(y2), int(m2)

    @staticmethod
    def _safe_float(val: object) -> Optional[float]:
        try:
            if val is None or (isinstance(val, str) and val.strip() == "-"):
                return None
            if "pd" in globals() and pd.isna(val):
                return None
            return float(val)
        except Exception:
            return None

    def _fmt_pct(self, val: object) -> str:
        f = self._safe_float(val)
        if f is None:
            return "-"
        return f"{f * 100:.1f}%"

    def _tipo_label(self, tipo: str) -> str:
        t = (tipo or "").strip().lower()
        if t.startswith("an"):
            return "ANO" if self.lang_index != 3 else "YEAR"
        if t.startswith("sem"):
            return "SEMESTRE" if self.lang_index != 3 else "SEMESTER"
        if t.startswith("tri"):
            return "TRIMESTRE" if self.lang_index != 3 else "QUARTER"
        return (tipo or "").strip().upper()

    def _add_footer_text(self, slide: "Presentation", msg: str) -> None:
        if not msg:
            return
        # Centrar el texto en el espacio "util" a la derecha del logo del template.
        logo_clear = Inches(2.00)
        right = Inches(0.35)
        left = logo_clear
        width = self.ppt.slide_width - left - right
        height = Inches(0.35)
        top = self.ppt.slide_height - height - Inches(0.10)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = Pt(2)
        tf.margin_right = Pt(2)
        p = tf.paragraphs[0]
        p.text = str(msg)
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 0, 0)
        p.alignment = 1

    def _add_low_penetration_footer(self, slide: "Presentation", buyers_value: Optional[float], threshold: float = 200) -> None:
        """Agrega un aviso al pie del slide cuando buyers promedio < threshold."""
        if buyers_value is None:
            return
        try:
            if "pd" in globals() and pd.isna(buyers_value):
                return
            buyers_num = float(buyers_value)
        except Exception:
            return
        if buyers_num >= float(threshold):
            return

        msg = self.labels.get((self.lang_index, "LowPenFooter")) or "Low penetration brand (<200 buyers) - For internal use only"
        self._add_footer_text(slide, msg)

    def _add_variations_box_pretty(
        self,
        slide: "Presentation",
        variations_detail: "pd.DataFrame",
        pipeline: int,
        trend_plot_df: "pd.DataFrame",
        container_left=None,
        container_top=None,
        container_width=None,
        container_height=None,
    ) -> None:
        """Renderiza el cuadro de variaciones en estilo 'bonito' (shapes, no imagen)."""
        if variations_detail is None or variations_detail.empty:
            return

        wp_col = "WP by Numerator" if "WP by Numerator" in variations_detail.columns else None
        # Sem pipeline (P0) siempre se intenta mostrar cuando existe.
        p0_col = "Cliente P0" if "Cliente P0" in variations_detail.columns else ("Cliente Pipeline (P0)" if "Cliente Pipeline (P0)" in variations_detail.columns else None)
        px_col = f"Cliente Pipeline (P{pipeline})" if f"Cliente Pipeline (P{pipeline})" in variations_detail.columns else (f"Cliente P{pipeline}" if f"Cliente P{pipeline}" in variations_detail.columns else None)
        if wp_col is None and p0_col is None and px_col is None:
            return

        show_pipeline_group = int(pipeline) > 0 and px_col is not None

        # Ubicación: si se pasa un contenedor (layout lado derecho), se usa; si no, fallback.
        if container_left is None:
            container_left = Inches(6.8)
        if container_top is None:
            container_top = Inches(1.15)
        if container_width is None:
            container_width = Inches(6.0)
        if container_height is None:
            container_height = Inches(5.8)

        # Row heights dentro del contenedor.
        row_gap = Inches(0.12)
        row_h = int((container_height - (2 * row_gap)) / 3)
        if row_h <= 0:
            row_h = Inches(0.50)

        # Intentar derivar el mes base del gráfico (mm-yy) para construir periodos por pipeline.
        base_year = None
        base_month = None
        if trend_plot_df is not None and not trend_plot_df.empty and COL_DATA in trend_plot_df.columns:
            last_token = str(trend_plot_df[COL_DATA].iloc[-1]).strip()
            try:
                mm_s, yy_s = last_token.split("-")
                base_month = int(mm_s)
                base_year = 2000 + int(yy_s)
            except Exception:
                base_year = None
                base_month = None

        # Colores (aprox. al ejemplo)
        green_border = RGBColor(126, 201, 67)  # #7EC943
        sellin_fill = RGBColor(126, 201, 67)
        kantar_fill = RGBColor(58, 58, 58)     # #3A3A3A
        white = RGBColor(255, 255, 255)
        black = RGBColor(0, 0, 0)
        grey = RGBColor(120, 120, 120)
        red = RGBColor(208, 2, 27)

        # Columnas (se escalan para llenar el ancho del contenedor).
        def _emu_to_in(v) -> float:
            return float(v) / 914400.0

        base_cols = {
            "tipo": 1.20,
            "var0": 1.10,
            "wp": 1.10,
            "sell0": 1.20,
        }
        if show_pipeline_group:
            base_cols.update({"varp": 1.10, "sellp": 1.25})
        base_total = sum(base_cols.values())
        scale = _emu_to_in(container_width) / base_total if base_total else 1.0

        col_tipo_w = Inches(base_cols["tipo"] * scale)
        col_var0_w = Inches(base_cols["var0"] * scale)
        col_kantar_w = Inches(base_cols["wp"] * scale)
        col_sell0_w = Inches(base_cols["sell0"] * scale)
        col_varp_w = Inches(base_cols.get("varp", 0.0) * scale)
        col_sellp_w = Inches(base_cols.get("sellp", 0.0) * scale)

        left = container_left
        top = container_top
        total_w = container_width

        def _add_row_box(y):
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, y, total_w, row_h)
            box.fill.solid()
            box.fill.fore_color.rgb = white
            box.line.color.rgb = green_border
            box.line.width = Pt(1.5)

        def _add_tipo_text(x, y, text):
            tb = slide.shapes.add_textbox(x, y, col_tipo_w, row_h)
            tf = tb.text_frame
            tf.clear()
            tf.word_wrap = False
            tf.margin_left = Pt(2)
            tf.margin_right = Pt(2)
            p = tf.paragraphs[0]
            p.text = text
            p.font.bold = True
            p.font.size = Pt(12)
            p.font.color.rgb = black
            p.alignment = 1

        def _add_var_text(x, y, w, period_text: str):
            tb = slide.shapes.add_textbox(x, y, w, row_h)
            tf = tb.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.margin_left = Pt(2)
            tf.margin_right = Pt(2)
            tf.margin_top = Pt(2)
            tf.margin_bottom = Pt(2)

            p1 = tf.paragraphs[0]
            # Mantener texto base, ya que el diseño es un "badge" visual.
            p1.text = "VAR %\nMOVEL" if self.lang_index == 1 else ("YOY %\nCHANGE" if self.lang_index == 3 else "VAR %\nMOVIL")
            p1.font.size = Pt(8)
            p1.font.bold = True
            p1.font.color.rgb = grey
            p1.alignment = 1

            p2 = tf.add_paragraph()
            p2.alignment = 1
            # period_text esperado: "JUN-25 vs JUN-24" (o similar)
            parts = [p.strip() for p in str(period_text or "").split("vs")]
            left_txt = parts[0].strip()
            right_txt = parts[1].strip() if len(parts) > 1 else ""
            r1 = p2.add_run()
            r1.text = f"{left_txt} " if left_txt else ""
            r1.font.size = Pt(8)
            r1.font.color.rgb = black
            rvs = p2.add_run()
            rvs.text = "vs"
            rvs.font.size = Pt(8)
            rvs.font.bold = True
            rvs.font.color.rgb = red
            r2 = p2.add_run()
            r2.text = f" {right_txt}" if right_txt else ""
            r2.font.size = Pt(8)
            r2.font.color.rgb = black

        def _add_value_card(x, y, w, fill_rgb, title, value):
            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, row_h)
            card.fill.solid()
            card.fill.fore_color.rgb = fill_rgb
            card.line.color.rgb = fill_rgb
            card.line.width = Pt(0.5)
            tf = card.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.margin_left = Pt(4)
            tf.margin_right = Pt(4)
            tf.margin_top = Pt(4)
            tf.margin_bottom = Pt(4)

            p1 = tf.paragraphs[0]
            p1.text = title
            # "WP by Numerator" es más largo que SELL-IN; bajamos un poco el tamaño.
            p1.font.size = Pt(7) if len(str(title)) >= 11 else Pt(9)
            p1.font.bold = True
            p1.font.color.rgb = white
            p1.alignment = 1

            p2 = tf.add_paragraph()
            p2.text = value
            p2.font.size = Pt(22)
            p2.font.bold = True
            p2.font.color.rgb = white
            p2.alignment = 1

        def _period_label(end_year: Optional[int], end_month: Optional[int], offset: int) -> str:
            if end_year is None or end_month is None:
                return "-"
            prev_y, prev_m = self._date_minus_months(int(end_year), int(end_month), int(offset))
            m1 = f"{self._month_abbr(int(end_month))}-{int(end_year) % 100:02d}"
            m2 = f"{self._month_abbr(int(prev_m))}-{int(prev_y) % 100:02d}"
            return f"{m1} vs {m2}"

        # Encabezados de grupo (Sem pipeline / Pipeline p)
        # Se alinean sobre las tarjetas numéricas (no sobre el bloque completo) para que queden centrados.
        header_h = Inches(0.16)
        header_y = top - Inches(0.36)
        # Permite un pequeño offset negativo para subir un poco más sin mover el cuadro completo.
        if header_y < Inches(-0.06):
            header_y = Inches(-0.06)

        def _add_group_header(x: int, w: int, text: str) -> None:
            tb = slide.shapes.add_textbox(x, header_y, w, header_h)
            tf = tb.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = text
            p.font.size = Pt(10)
            p.font.color.rgb = grey
            p.alignment = 1

        # Sem pipeline: encima del SELL-IN (verde) del sem pipeline (P0), no sobre el bloque completo.
        x_sem = left + col_tipo_w + col_var0_w + col_kantar_w
        w_sem = col_sell0_w
        _add_group_header(x_sem, w_sem, "Sem pipeline" if self.lang_index != 3 else "No pipeline")

        if show_pipeline_group:
            # Pipeline p: encima del SELL-IN del pipeline (no incluye el badge de periodo).
            x_pip = left + col_tipo_w + col_var0_w + col_kantar_w + col_sell0_w + col_varp_w
            w_pip = col_sellp_w
            _add_group_header(x_pip, w_pip, f"Pipeline {int(pipeline)}")

        tipo_order = ["Anual", "Semestral", "Trimestral"]
        for idx, tipo in enumerate(tipo_order):
            y = top + (idx * (row_h + row_gap))
            _add_row_box(y)

            row = variations_detail[variations_detail["Tipo"].astype(str).str.lower().str.startswith(tipo[:3].lower())]
            wp_val = row[wp_col].iloc[0] if (wp_col and not row.empty and wp_col in row.columns) else None
            p0_val = row[p0_col].iloc[0] if (p0_col and not row.empty and p0_col in row.columns) else None
            px_val = row[px_col].iloc[0] if (px_col and not row.empty and px_col in row.columns) else None

            offsets = {"anual": 12, "semestral": 6, "trimestral": 3}
            tkey = (tipo or "").strip().lower()
            if tkey.startswith("an"):
                offset = offsets["anual"]
            elif tkey.startswith("sem"):
                offset = offsets["semestral"]
            elif tkey.startswith("tri"):
                offset = offsets["trimestral"]
            else:
                offset = 12

            # Periodo sem pipeline (p=0) y pipeline p (p=pipeline), usando el mes base del gráfico.
            sem_end_y, sem_end_m = (base_year, base_month)
            if base_year is not None and base_month is not None:
                pip_end_y, pip_end_m = self._date_minus_months(int(base_year), int(base_month), int(pipeline))
            else:
                pip_end_y, pip_end_m = (None, None)
            sem_period = _period_label(sem_end_y, sem_end_m, offset)
            pip_period = _period_label(pip_end_y, pip_end_m, offset)

            x = left
            _add_tipo_text(x, y, self._tipo_label(tipo))
            x += col_tipo_w
            _add_var_text(x, y, col_var0_w, sem_period)
            x += col_var0_w
            if wp_col is not None:
                _add_value_card(x, y, col_kantar_w, kantar_fill, "WP by Numerator", self._fmt_pct(wp_val))
            x += col_kantar_w
            if p0_col is not None:
                _add_value_card(x, y, col_sell0_w, sellin_fill, "SELL-IN", self._fmt_pct(p0_val))
            x += col_sell0_w
            if show_pipeline_group:
                _add_var_text(x, y, col_varp_w, pip_period)
                x += col_varp_w
                _add_value_card(x, y, col_sellp_w, sellin_fill, "SELL-IN", self._fmt_pct(px_val))

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
        if self.coverage_slide_variant == "complemented":
            try:
                self._add_cov_slide_header_boxes(slide_cov, assets)
            except Exception as exc:
                print(f"{Fore.YELLOW}Advertencia: No se pudo generar el header complementado (penetración/cobertura) para {marca_nombre_limpio} P{assets.pipeline}. Error: {exc}")
                try:
                    table_stream = dataframe_to_bordered_stream(assets.variation_table, hide_index=True, dpi=200)
                    slide_cov.shapes.add_picture(table_stream, Inches(0.5), Inches(1.1), height=Inches(0.6))
                except Exception as exc2:
                    print(f"{Fore.YELLOW}Advertencia: Tampoco se pudo generar la tabla VAR % MAT (fallback) para {marca_nombre_limpio} P{assets.pipeline}. Error: {exc2}")
        else:
            try:
                table_stream = dataframe_to_bordered_stream(assets.variation_table, hide_index=True, dpi=200)
                slide_cov.shapes.add_picture(table_stream, Inches(0.5), Inches(1.1), height=Inches(0.6))
            except Exception as exc:
                print(f"{Fore.YELLOW}Advertencia: No se pudo generar la tabla de variación MAT para {marca_nombre_limpio} P{assets.pipeline}. Error: {exc}")
        self._add_low_penetration_footer(slide_cov, getattr(assets, "buyers_mat_actual", None))
        slides_created += 1
        if progress and task_id is not None:
            progress.update(task_id, advance=1)
        slide_trend = self.ppt.slides.add_slide(self.ppt.slide_layouts[PPT_LAYOUT_INDEX])
        tx_title_trend = ensure_title_frame(slide_trend)
        p_trend = tx_title_trend.paragraphs[0]
        p_trend.text = f"{marca_nombre_limpio} - Pipeline {assets.pipeline}"
        p_trend.font.bold = True
        p_trend.font.size = Pt(24)
        has_variations = assets.variations_detail is not None and not assets.variations_detail.empty
        if self.variations_box_style == "pretty" and has_variations:
            # Layout: gráfico a la izquierda + cuadro bonito a la derecha, mismo alto.
            content_top = Inches(1.15)
            content_bottom = Inches(0.55)
            content_h = self.ppt.slide_height - content_top - content_bottom
            margin_l = Inches(0.35)
            margin_r = Inches(0.25)
            divider_w = Inches(0.06)
            avail_w = self.ppt.slide_width - margin_l - margin_r - divider_w
            left_w = int(avail_w / 2)
            right_w = avail_w - left_w
            chart_left = margin_l
            chart_top = content_top
            var_left = margin_l + left_w + divider_w
            var_top = content_top

            # Divisor vertical (como referencia)
            divider_x = margin_l + left_w
            divider = slide_trend.shapes.add_shape(MSO_SHAPE.RECTANGLE, divider_x, content_top, divider_w, content_h)
            divider.fill.solid()
            # Azul verdoso (segun referencia)
            divider.fill.fore_color.rgb = RGBColor(0, 229, 176)  # #00E5B0
            divider.line.fill.background()

            generar_grafico_tendencia(
                slide_trend,
                marca_nombre_limpio,
                assets.pipeline,
                assets.trend_plot_df,
                lang_index,
                self.labels,
                doble_eje=(self.tipo_eje_tend == "doble"),
                box_left=chart_left,
                box_top=chart_top,
                box_width=left_w,
                box_height=content_h,
                # Imagen más "alta" para que se aproveche mejor la columna izquierda sin estirar.
                figsize=(9.0, 7.0),
                legend_y=-0.22,
            )
        else:
            generar_grafico_tendencia(
                slide_trend,
                marca_nombre_limpio,
                assets.pipeline,
                assets.trend_plot_df,
                lang_index,
                self.labels,
                doble_eje=(self.tipo_eje_tend == "doble"),
            )
        if has_variations:
            if self.variations_box_style == "pretty":
                # Usa el mismo contenedor del lado derecho que el layout del gráfico.
                content_top = Inches(1.15)
                content_bottom = Inches(0.55)
                content_h = self.ppt.slide_height - content_top - content_bottom
                margin_l = Inches(0.35)
                margin_r = Inches(0.25)
                divider_w = Inches(0.06)
                avail_w = self.ppt.slide_width - margin_l - margin_r - divider_w
                left_w = int(avail_w / 2)
                right_w = avail_w - left_w
                var_left = margin_l + left_w + divider_w
                # Reducir altura a la mitad y centrar verticalmente en el área de contenido.
                var_h = int(content_h * 0.5)
                var_top = content_top + int((content_h - var_h) / 2)
                self._add_variations_box_pretty(
                    slide_trend,
                    assets.variations_detail,
                    assets.pipeline,
                    assets.trend_plot_df,
                    container_left=var_left,
                    container_top=var_top,
                    container_width=right_w,
                    container_height=var_h,
                )
            else:
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
        self._add_low_penetration_footer(slide_trend, getattr(assets, "buyers_mat_actual", None))
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
            self._add_low_penetration_footer(slide_evol, getattr(assets, "buyers_mat_actual", None))
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
        low_penetration_brands: Optional[Sequence[str]] = None,
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
        low_penetration_brands = list(low_penetration_brands or [])
        try:
            left = Inches(0.5)
            top = Inches(1.0)
            usable_w = self.ppt.slide_width - 2 * left
            max_h = Inches(4.8)
            self._add_editable_summary_table(
                slide_summary,
                df_summary,
                left=left,
                top=top,
                width=usable_w,
                max_height=max_h,
                low_penetration_brands=low_penetration_brands,
            )
        except Exception as exc:
            print(f"{Fore.YELLOW}Advertencia: No se pudo generar la tabla resumen en el PPT. Error: {exc}")
        if low_penetration_brands:
            unique_brands = sorted({str(b).strip() for b in low_penetration_brands if str(b).strip()})
            n = len(unique_brands)
            key = "LowPenSummaryPlural" if n > 1 else "LowPenSummarySingular"
            tpl = self.labels.get((self.lang_index, key))
            if not tpl:
                tpl = self.labels.get((self.lang_index, "LowPenSummaryPlural" if n > 1 else "LowPenSummarySingular"))
            if not tpl:
                tpl = "This study contains {n} low penetration brand(s) (<200 buyers). For internal use only"
            msg = str(tpl).format(n=n)
            self._add_footer_text(slide_summary, msg)

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


def _build_sheet_header_index(ws: "object") -> Dict[str, int]:
    """Devuelve un índice {header: columna} a partir de la fila 1."""
    header_index: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        raw_value = ws.cell(row=1, column=col).value
        if raw_value is None:
            continue
        header = str(raw_value).strip()
        if header and header not in header_index:
            header_index[header] = col
    return header_index


def _find_last_data_row(ws: "object", data_col: int, start_row: int = 2) -> int:
    """Encuentra la última fila contigua con datos en la columna de fecha."""
    row = start_row
    last_valid = start_row - 1
    while row <= ws.max_row:
        value = ws.cell(row=row, column=data_col).value
        if value is None or str(value).strip() == "":
            break
        last_valid = row
        row += 1
    return last_valid


def _find_first_nonempty_row(ws: "object", col: int, start_row: int, end_row: int) -> Optional[int]:
    """Devuelve la primera fila con valor no vacío en un rango vertical."""
    for row in range(start_row, end_row + 1):
        value = ws.cell(row=row, column=col).value
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return row
    return None


def _excel_lang_code(include_english: bool, pais_nombre: str) -> str:
    if include_english:
        return "EN"
    return "PT" if (pais_nombre or "").strip().lower() in {"brasil", "brazil"} else "ES"


def _parse_pipeline_from_sheet_name(sheet_name: str) -> int:
    match = re.match(r"(?i)^p([0-6])_", str(sheet_name or "").strip())
    return int(match.group(1)) if match else 0


def _clean_brand_name_from_sheet(sheet_name: str) -> str:
    cleaned = re.sub(r"(?i)^p[0-6]_", "", str(sheet_name or "")).strip()
    return cleaned or str(sheet_name or "N/D")


def _safe_hex(color_value: str) -> str:
    return str(color_value or "").strip().replace("#", "")[:6] or "000000"


def _set_line_series_color(series_obj: "object", color_value: str, width: int = 28575) -> None:
    color_hex = _safe_hex(color_value)
    try:
        series_obj.graphicalProperties.line.solidFill = color_hex
        series_obj.graphicalProperties.line.width = width
    except Exception:
        pass
    try:
        series_obj.graphicalProperties.solidFill = color_hex
    except Exception:
        pass


def _set_bar_series_color(series_obj: "object", color_value: str, line_color: str = "000000") -> None:
    fill_hex = _safe_hex(color_value)
    line_hex = _safe_hex(line_color)
    try:
        series_obj.graphicalProperties.solidFill = fill_hex
    except Exception:
        pass
    try:
        series_obj.graphicalProperties.line.solidFill = line_hex
    except Exception:
        pass


def add_native_excel_charts(
    xlsx_path: str,
    *,
    coverage_label: str,
    trend_axis: str,
    evolution_slide_variant: str,
    include_english: bool,
    pais_nombre: str,
) -> None:
    """
    Inserta graficos nativos de Excel (editables) en cada hoja de marca,
    replicando los 3 graficos principales del flujo PPT:
    - Cobertura vs penetracion mensual (barras).
    - Tendencia de volumen Sell-in vs Sell-out (lineas).
    - Evolucion mensual y variacion interanual (simple o clasico).
    """
    from openpyxl import load_workbook as _load_wb_chart
    from openpyxl.chart import (
        BarChart as _BarChart,
        LineChart as _LineChart,
        Reference as _Reference,
    )
    from openpyxl.chart.label import DataLabelList as _DataLabelList
    from openpyxl.chart.series import SeriesLabel as _SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties as _GraphicalProperties
    from openpyxl.utils import get_column_letter as _get_col_letter

    lang_code = _excel_lang_code(include_english, pais_nombre)
    trend_axis_norm = str(trend_axis or "").strip().lower()
    evolution_variant_norm = normalize_evolution_slide_variant(evolution_slide_variant)
    chart_titles = {
        "ES": {
            "coverage_title": "Cobertura en Ano Movil",
            "penetration_label": "Penetracion Mensual",
            "trend_title": "Tendencia en Volumen",
            "evolution_title": "Evolucion Mensual y Variacion",
            "evolution_var_axis": "Variacion Interanual",
            "evolution_monthly_axis": "Volumen Mensual",
        },
        "PT": {
            "coverage_title": "Cobertura em Ano Movel",
            "penetration_label": "Penetracao Mensal",
            "trend_title": "Tendencia em Volumen",
            "evolution_title": "Evolucao Mensal e Variacao",
            "evolution_var_axis": "Variacao Interanual",
            "evolution_monthly_axis": "Volumen Mensual",
        },
        "EN": {
            "coverage_title": "MOVING YEAR COVERAGE",
            "penetration_label": "PENETRATION BY PERIOD",
            "trend_title": "TREND IN VOLUME",
            "evolution_title": "Monthly Evolution and YoY Variation",
            "evolution_var_axis": "YoY Variation",
            "evolution_monthly_axis": "Monthly Volume",
        },
    }[lang_code]
    chart_scale = 1.2
    chart_anchor_col = "AA"

    def _tint_hex_color(color_value: str, mix_with_white: float = 0.78) -> str:
        """Aclara un color HEX mezclándolo con blanco."""
        hex_color = _safe_hex(color_value)
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
        except Exception:
            return "E7E6E6"
        m = max(0.0, min(1.0, float(mix_with_white)))
        r2 = int(round(r + (255 - r) * m))
        g2 = int(round(g + (255 - g) * m))
        b2 = int(round(b + (255 - b) * m))
        return f"{r2:02X}{g2:02X}{b2:02X}"

    def _apply_variation_labels(series_obj: "object", line_color: str) -> None:
        """Muestra valor puntual con fondo difuminado del color de línea y color por signo."""
        dlabels = _DataLabelList()
        dlabels.showVal = True
        dlabels.showSerName = False
        dlabels.showCatName = False
        dlabels.showLegendKey = False
        dlabels.showPercent = False
        dlabels.separator = " "
        # Color de fuente por signo (Excel evalúa el formato en tiempo de cálculo).
        dlabels.numFmt = "[Green]0.0%;[Red]-0.0%;0.0%"
        series_obj.dLbls = dlabels
        # Fondo difuminado con color de la línea de la serie.
        try:
            dlabels.spPr = _GraphicalProperties(solidFill=_tint_hex_color(line_color, mix_with_white=0.78))
            if getattr(dlabels.spPr, "line", None) is not None:
                dlabels.spPr.line.solidFill = _safe_hex(line_color)
        except Exception:
            pass

    wb = _load_wb_chart(xlsx_path)
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    for ws in wb.worksheets:
        headers = _build_sheet_header_index(ws)
        pipeline = _parse_pipeline_from_sheet_name(ws.title)
        cov_header = f"P{pipeline}"
        required_headers = [COL_DATA, COL_SELL_OUT, COL_SELL_IN_SIM, COL_PENET, cov_header]
        if any(header not in headers for header in required_headers):
            continue

        data_col = headers[COL_DATA]
        sell_out_col = headers[COL_SELL_OUT]
        sell_in_sim_col = headers[COL_SELL_IN_SIM]
        pen_col = headers[COL_PENET]
        cov_col = headers[cov_header]
        evo_kantar_col = headers.get(COL_EVO_KANTAR_YOY)
        evo_sellin_col = headers.get(COL_EVO_SELLIN_YOY)

        last_data_row = _find_last_data_row(ws, data_col=data_col, start_row=2)
        if last_data_row < 3:
            continue
        n_data_rows = last_data_row - 1
        if n_data_rows < 12:
            continue

        # Evita duplicados sin borrar graficos ajenos a esta rutina.
        def _chart_anchor_cell(ch: "object") -> str:
            try:
                anchor = ch.anchor
                if isinstance(anchor, str):
                    return anchor.upper()
                if hasattr(anchor, "_from"):
                    return f"{_get_col_letter(anchor._from.col + 1)}{anchor._from.row + 1}"
            except Exception:
                return ""
            return ""

        if hasattr(ws, "_charts"):
            target_anchors = {
                f"{chart_anchor_col}2",
                f"{chart_anchor_col}22",
                f"{chart_anchor_col}42",
                # Limpieza de anclas anteriores para evitar duplicados al regenerar.
                "W2",
                "W22",
                "W42",
            }
            ws._charts = [c for c in ws._charts if _chart_anchor_cell(c) not in target_anchors]  # type: ignore[attr-defined]

        brand_name = _clean_brand_name_from_sheet(ws.title)
        trend_start = 2 + pipeline
        trend_end = last_data_row
        sell_in_start = 2
        sell_in_end = last_data_row - pipeline

        # 1) Cobertura vs Penetracion (rangos directos de columnas originales).
        cov_start = _find_first_nonempty_row(ws, cov_col, start_row=2, end_row=last_data_row)
        if cov_start is not None and cov_start <= last_data_row:
            coverage_chart = _BarChart()
            coverage_chart.type = "col"
            coverage_chart.grouping = "clustered"
            coverage_chart.overlap = 0
            coverage_chart.gapWidth = 85
            coverage_chart.height = 7.1 * chart_scale
            coverage_chart.width = 16.2 * chart_scale
            coverage_chart.title = f"{chart_titles['coverage_title']} | {brand_name} Pipeline {pipeline}"
            coverage_chart.y_axis.title = f"{coverage_label} | {chart_titles['penetration_label']}"
            coverage_chart.y_axis.scaling.min = 0
            coverage_chart.y_axis.numFmt = "0.0"
            coverage_chart.x_axis.number_format = "yyyy-mm"
            coverage_chart.x_axis.numFmt = "yyyy-mm"
            coverage_chart.x_axis.tickLblPos = "low"
            coverage_chart.x_axis.tickLblSkip = 1
            coverage_chart.x_axis.tickMarkSkip = 1
            coverage_chart.x_axis.delete = False
            coverage_chart.legend.position = "b"
            coverage_chart.legend.overlay = False

            coverage_chart.add_data(
                _Reference(ws, min_col=pen_col, min_row=cov_start, max_row=last_data_row),
                titles_from_data=False,
            )
            coverage_chart.series[-1].title = _SeriesLabel(v=chart_titles["penetration_label"])
            coverage_chart.add_data(
                _Reference(ws, min_col=cov_col, min_row=cov_start, max_row=last_data_row),
                titles_from_data=False,
            )
            coverage_chart.series[-1].title = _SeriesLabel(v=coverage_label)
            coverage_chart.set_categories(
                _Reference(ws, min_col=data_col, min_row=cov_start, max_row=last_data_row)
            )
            coverage_chart.dataLabels = _DataLabelList()
            coverage_chart.dataLabels.showVal = True
            coverage_chart.dataLabels.showSerName = False
            coverage_chart.dataLabels.showCatName = False
            coverage_chart.dataLabels.showLegendKey = False
            coverage_chart.dataLabels.showPercent = False
            coverage_chart.dataLabels.numFmt = "0.0"
            coverage_chart.dataLabels.separator = " "
            _set_bar_series_color(coverage_chart.series[0], COLOR_PENETRACION_BAR)
            _set_bar_series_color(coverage_chart.series[1], COLOR_COBERTURA_BAR)
            ws.add_chart(coverage_chart, f"{chart_anchor_col}2")

        # 2) Tendencia (rangos directos para evitar graficos vacios por formulas auxiliares).
        if trend_start <= trend_end and sell_in_start <= sell_in_end:
            trend_categories = _Reference(
                ws,
                min_col=data_col,
                min_row=trend_start,
                max_row=trend_end,
            )
            trend_chart = _LineChart()
            trend_chart.style = 2
            trend_chart.height = 7.1 * chart_scale
            trend_chart.width = 16.2 * chart_scale
            trend_chart.title = f"{chart_titles['trend_title']} | {brand_name} P:{pipeline}"
            trend_chart.x_axis.number_format = "yyyy-mm"
            trend_chart.x_axis.numFmt = "yyyy-mm"
            trend_chart.x_axis.tickLblPos = "low"
            trend_chart.x_axis.tickLblSkip = 1
            trend_chart.x_axis.tickMarkSkip = 1
            trend_chart.x_axis.delete = False
            trend_chart.legend.position = "b"
            trend_chart.legend.overlay = False
            trend_chart.y_axis.scaling.min = 0

            if trend_axis_norm == "doble":
                trend_chart.y_axis.title = COL_SELL_IN
                trend_chart.add_data(
                    _Reference(ws, min_col=sell_in_sim_col, min_row=sell_in_start, max_row=sell_in_end),
                    titles_from_data=False,
                )
                trend_chart.series[-1].title = _SeriesLabel(v=f"{COL_SELL_IN} (P:{pipeline})")
                trend_chart.set_categories(trend_categories)

                trend_chart2 = _LineChart()
                trend_chart2.y_axis.axId = 200
                trend_chart2.y_axis.crosses = "max"
                trend_chart2.y_axis.title = COL_SELL_OUT
                trend_chart2.add_data(
                    _Reference(ws, min_col=sell_out_col, min_row=trend_start, max_row=trend_end),
                    titles_from_data=False,
                )
                trend_chart2.series[-1].title = _SeriesLabel(v=COL_SELL_OUT)
                trend_chart += trend_chart2
            else:
                trend_chart.y_axis.title = f"{COL_SELL_IN} / {COL_SELL_OUT}"
                trend_chart.add_data(
                    _Reference(ws, min_col=sell_in_sim_col, min_row=sell_in_start, max_row=sell_in_end),
                    titles_from_data=False,
                )
                trend_chart.series[-1].title = _SeriesLabel(v=f"{COL_SELL_IN} (P:{pipeline})")
                trend_chart.add_data(
                    _Reference(ws, min_col=sell_out_col, min_row=trend_start, max_row=trend_end),
                    titles_from_data=False,
                )
                trend_chart.series[-1].title = _SeriesLabel(v=COL_SELL_OUT)
                trend_chart.set_categories(trend_categories)

            if len(trend_chart.series) >= 1:
                _set_line_series_color(trend_chart.series[0], COLOR_SELLIN_TREND_LINE)
            if len(trend_chart.series) >= 2:
                _set_line_series_color(trend_chart.series[1], COLOR_SELLOUT_TREND_LINE)
            ws.add_chart(trend_chart, f"{chart_anchor_col}22")

        # 3) Evolucion mensual y variacion interanual (nutrida por columnas V/W del Excel).
        if (
            n_data_rows >= 24
            and trend_start <= trend_end
            and sell_in_start <= sell_in_end
            and evo_kantar_col is not None
            and evo_sellin_col is not None
        ):
            evo_start_k = _find_first_nonempty_row(ws, evo_kantar_col, start_row=2, end_row=last_data_row)
            evo_start_s = _find_first_nonempty_row(ws, evo_sellin_col, start_row=2, end_row=last_data_row)
            evo_start_s_shifted = (evo_start_s + pipeline) if evo_start_s is not None else None
            evo_candidates = [r for r in (evo_start_k, evo_start_s_shifted) if r is not None]
            if not evo_candidates:
                continue
            evo_start = max(evo_candidates)
            if evo_start > last_data_row:
                continue

            sellin_var_start = max(2, evo_start - pipeline)
            sellin_var_end = max(2, last_data_row - pipeline)
            if sellin_var_start > sellin_var_end:
                continue

            evo_categories = _Reference(
                ws,
                min_col=data_col,
                min_row=evo_start,
                max_row=last_data_row,
            )

            if evolution_variant_norm == "simple":
                evol_chart = _LineChart()
                evol_chart.style = 13
                evol_chart.height = 7.5 * chart_scale
                evol_chart.width = 16.2 * chart_scale
                evol_chart.title = f"{chart_titles['evolution_title']} | {brand_name} P:{pipeline}"
                evol_chart.y_axis.title = chart_titles["evolution_var_axis"]
                evol_chart.y_axis.numFmt = "0.0%"
                evol_chart.x_axis.number_format = "yyyy-mm"
                evol_chart.x_axis.numFmt = "yyyy-mm"
                evol_chart.x_axis.tickLblPos = "low"
                evol_chart.x_axis.tickLblSkip = 1
                evol_chart.x_axis.tickMarkSkip = 1
                evol_chart.x_axis.delete = False
                evol_chart.legend.position = "b"
                evol_chart.legend.overlay = False
                evol_chart.add_data(
                    _Reference(
                        ws,
                        min_col=evo_kantar_col,
                        min_row=evo_start,
                        max_row=last_data_row,
                    ),
                    titles_from_data=False,
                )
                evol_chart.series[-1].title = _SeriesLabel(v=COL_EVO_KANTAR_YOY)
                evol_chart.add_data(
                    _Reference(
                        ws,
                        min_col=evo_sellin_col,
                        min_row=sellin_var_start,
                        max_row=sellin_var_end,
                    ),
                    titles_from_data=False,
                )
                evol_chart.series[-1].title = _SeriesLabel(v=COL_EVO_SELLIN_YOY)
                evol_chart.set_categories(evo_categories)
                if len(evol_chart.series) >= 1:
                    _set_line_series_color(evol_chart.series[0], COLOR_KANTAR_LINE)
                    _apply_variation_labels(evol_chart.series[0], COLOR_KANTAR_LINE)
                if len(evol_chart.series) >= 2:
                    _set_line_series_color(evol_chart.series[1], COLOR_SELLIN_LINE)
                    _apply_variation_labels(evol_chart.series[1], COLOR_SELLIN_LINE)
                ws.add_chart(evol_chart, f"{chart_anchor_col}42")
            else:
                # Modo clasico: solo tendencia de variaciones (sin volumen), con etiquetas puntuales.
                evol_var_chart = _LineChart()
                evol_var_chart.style = 2
                evol_var_chart.height = 7.5 * chart_scale
                evol_var_chart.width = 16.2 * chart_scale
                evol_var_chart.title = f"{chart_titles['evolution_title']} | {brand_name} P:{pipeline}"
                evol_var_chart.y_axis.title = chart_titles["evolution_var_axis"]
                evol_var_chart.y_axis.numFmt = "0.0%"
                evol_var_chart.x_axis.number_format = "yyyy-mm"
                evol_var_chart.x_axis.numFmt = "yyyy-mm"
                evol_var_chart.x_axis.tickLblPos = "low"
                evol_var_chart.x_axis.tickLblSkip = 1
                evol_var_chart.x_axis.tickMarkSkip = 1
                evol_var_chart.x_axis.delete = False
                evol_var_chart.legend.position = "b"
                evol_var_chart.legend.overlay = False
                evol_var_chart.add_data(
                    _Reference(
                        ws,
                        min_col=evo_kantar_col,
                        min_row=evo_start,
                        max_row=last_data_row,
                    ),
                    titles_from_data=False,
                )
                evol_var_chart.series[-1].title = _SeriesLabel(v=COL_EVO_KANTAR_YOY)
                evol_var_chart.add_data(
                    _Reference(
                        ws,
                        min_col=evo_sellin_col,
                        min_row=sellin_var_start,
                        max_row=sellin_var_end,
                    ),
                    titles_from_data=False,
                )
                evol_var_chart.series[-1].title = _SeriesLabel(v=COL_EVO_SELLIN_YOY)
                evol_var_chart.set_categories(evo_categories)
                if len(evol_var_chart.series) >= 1:
                    _set_line_series_color(evol_var_chart.series[0], COLOR_KANTAR_BAR_VAR)
                    _apply_variation_labels(evol_var_chart.series[0], COLOR_KANTAR_BAR_VAR)
                if len(evol_var_chart.series) >= 2:
                    _set_line_series_color(evol_var_chart.series[1], COLOR_SELLIN_BAR_VAR)
                    _apply_variation_labels(evol_var_chart.series[1], COLOR_SELLIN_BAR_VAR)
                ws.add_chart(evol_var_chart, f"{chart_anchor_col}42")

    wb.save(xlsx_path)


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
    coverage_reason: str,
    trend_axis: str,
    evolution_slide_variant: str,
    include_english: bool,
) -> Tuple[str, str, str, str]:
    """Genera el archivo Excel temporal y devuelve datos clave."""
    try:
        console.print("\n[bold cyan]Generando archivo Excel temporal...[/bold cyan]")
    except Exception:
        print(Fore.CYAN + "\nGenerando archivo Excel temporal...")
    excel_temp_path = os.path.join(root_dir, EXCEL_TEMP_FILENAME)
    try:
        with pd.ExcelWriter(excel_temp_path) as writer:
            # Recorrer cada hoja (marca) del archivo
            total_sheets = len(marcas) if hasattr(marcas, "__len__") else 0
            status = console.status("Procesando hojas Excel...", spinner="line")
            status.start()
            try:
                for idx_sheet, marca_sheet_name in enumerate(marcas, start=1):
                    status.update(f"Procesando hoja {idx_sheet}/{total_sheets}: {marca_sheet_name}")

                    # 1.1) Carga y preprocesa la hoja usando la función refactorizada
                    df_marca, measure_unit = load_and_preprocess_sheet(excel_file_obj, marca_sheet_name)

                    # Si la carga falló, continuar con la siguiente hoja
                    if df_marca is None:
                        continue

                    # Guardar número original de filas de datos para fórmulas Excel
                    original_data_rows = len(df_marca)
                    if original_data_rows < 12:
                        console.print(
                            f"[yellow]Advertencia:[/] Hoja '{marca_sheet_name}' tiene < 12 meses de datos ({original_data_rows}). "
                            "Algunos calculos de Excel pueden fallar o dar NaN."
                        )
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
                    min_periods_for_layout = 42                                    # 36 meses + 6 pipelines para mantener layout
                    missing_periods = max(0, min_periods_for_layout - original_data_rows)

                    def build_if_no_zero(num_range: str, den_range: str, formula_body: str) -> str:
                        """
                        Envuelve una fórmula con validación de ceros: si hay 0 en alguna de las ventanas, devuelve "-".
                        'formula_body' no debe llevar '=' al inicio.
                        """
                        return f"=IF(OR(COUNTIF({num_range},0)>0,COUNTIF({den_range},0)>0),\"-\",{formula_body})"

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
                    var_wp = []
                    for i, j in zip([10, 4, 1], [11, 5, 2]):
                        num_start = original_data_rows + excel_row_offset - i - 2
                        num_end = original_data_rows + excel_row_offset - 1
                        den_start = original_data_rows + excel_row_offset - 2 * j - 2
                        den_end = original_data_rows + excel_row_offset - j - 2
                        num_range = f"C{num_start}:C{num_end}"
                        den_range = f"C{den_start}:C{den_end}"
                        formula_body = f"SUM({num_range})/SUM({den_range})-1"
                        var_wp.append(build_if_no_zero(num_range, den_range, formula_body))
                    var['WP by Numerator'] = var_wp

                    # Variaciones Cliente
                    for p in range(7):
                        cli_var = []
                        for i, j in zip([10, 4, 1], [11, 5, 2]):
                            num_start = original_data_rows + excel_row_offset - i - p - 2
                            num_end = original_data_rows + excel_row_offset - p - 1
                            den_start = original_data_rows + excel_row_offset - 2 * j - p - 2
                            den_end = original_data_rows + excel_row_offset - j - p - 2
                            num_range = f"L{num_start}:L{num_end}"
                            den_range = f"L{den_start}:L{den_end}"
                            formula_body = f"SUM({num_range})/SUM({den_range})-1"
                            cli_var.append(build_if_no_zero(num_range, den_range, formula_body))
                        var[f'Cliente P{p}'] = cli_var

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
                        num_range = f"{col}{num_ini}:{col}{num_fin}"
                        den_range = f"{col}{den_ini}:{col}{den_fin}"
                        formula_body = f"SUM({num_range})/SUM({den_range})-1"
                        return build_if_no_zero(num_range, den_range, formula_body)

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

                    # Limpiar variaciones sin sentido sin crear columnas negativas
                    if missing_periods > 0:
                        console.print(
                            f"[yellow]Advertencia:[/] Correlaciones/variaciones con periodos incompletos para "
                            f"[green]{marca_sheet_name}[/green] ({original_data_rows}/{min_periods_for_layout}); "
                            "se calculan correlaciones posibles; faltantes='-'."
                        )


                    # ---------- Unir Y-1 y Y-2 --------------------------------------
                    df_variations_excel = pd.concat([var, aux], ignore_index=True)



                    # --- 1.9) Cálculo de correlaciones en Excel (MAT) ---
                    # Se genera un diccionario con fórmulas de correlación para cada pipeline (P0 a P6)
                    # Se construyen fórmulas Excel que calculan la correlación Pearson entre dos rangos de 12 filas:
                    #   uno en la columna M y otro en la columna N, considerando el desplazamiento (pipeline).
                    # Los índices son base 1 y se garantiza que cada rango tenga exactamente 12 filas; de lo contrario, se asigna '-'.
            
                    # ---------- Correlaciones: 12m, 2 años antes (12m terminando hace 24m), 2 años (ventana 24m) ----------

                    series_sell_out = pd.to_numeric(df_marca[COL_SELL_OUT], errors="coerce")
                    series_sell_in = pd.to_numeric(df_marca[COL_SELL_IN], errors="coerce")

                    def _window_invalid(series: "pd.Series", start_row_excel: int, end_row_excel: int) -> bool:
                        """
                        Valida que la ventana exista, no tenga NaN y no tenga ceros.
                        Devuelve True si la ventana es inválida.
                        """
                        start_idx = start_row_excel - excel_row_offset
                        end_idx = end_row_excel - excel_row_offset
                        if start_idx < 0 or end_idx >= len(series):
                            return True
                        window = series.iloc[start_idx:end_idx+1]
                        return window.isna().any() or (window == 0).any()

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
                                # Cada pipeline consume filas adicionales; valida que haya datos suficientes
                                if (n_data - p) < (window + end_offset):
                                    row[f'P{p}'] = '-'
                                    continue

                                n_start = max(row_ini - p, 2)
                                n_end   = max(row_fin - p, 2)

                                # Ambas ventanas deben tener exactamente 'window' filas
                                if (m_end - m_start + 1 == window) and (n_end - n_start + 1 == window):
                                    # Si hay 0s o NaN en alguna ventana, se considera incompleto y se marca con '-'
                                    if _window_invalid(series_sell_out, m_start, m_end) or _window_invalid(series_sell_in, n_start, n_end):
                                        row[f'P{p}'] = "-"
                                    else:
                                        # Usa coma ',' en argumentos; función en inglés 'CORREL' como en tu flujo actual
                                        m_range = f"M{m_start}:M{m_end}"
                                        n_range = f"N{n_start}:N{n_end}"
                                        row[f'P{p}'] = (
                                            f"=IF(OR(COUNTBLANK({m_range})>0,COUNTBLANK({n_range})>0,"
                                            f"COUNTIF({m_range},0)>0,COUNTIF({n_range},0)>0),\"-\","
                                            f"CORREL({m_range},{n_range}))"
                                        )
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


                    # --- 1.10-bis) Variaciones YoY 12m para gráfico de evolución (columnas V y W) ---
                    evo_kantar_formulas: List[object] = []
                    evo_sellin_formulas: List[object] = []
                    for r_idx in range(original_data_rows):
                        row_excel = r_idx + excel_row_offset

                        # WP by Numerator YoY 12m: (SUM últimos 12) / (SUM 12 previos) - 1
                        if row_excel >= (excel_row_offset + 23):
                            num_range_c = f"C{row_excel - 11}:C{row_excel}"
                            den_range_c = f"C{row_excel - 23}:C{row_excel - 12}"
                            evo_kantar_formulas.append(
                                f"=IFERROR(SUM({num_range_c})/IF(SUM({den_range_c})=0,1,SUM({den_range_c}))-1,NA())"
                            )
                        else:
                            evo_kantar_formulas.append(np.nan)

                        # Sell-in YoY 12m SIN pipeline (pipeline se aplica al graficar)
                        if row_excel >= (excel_row_offset + 23):
                            sell_end = row_excel
                            sell_start = sell_end - 11
                            sell_prev_end = sell_end - 12
                            sell_prev_start = sell_prev_end - 11
                            num_range_l = f"L{sell_start}:L{sell_end}"
                            den_range_l = f"L{sell_prev_start}:L{sell_prev_end}"
                            evo_sellin_formulas.append(
                                f"=IFERROR(SUM({num_range_l})/IF(SUM({den_range_l})=0,1,SUM({den_range_l}))-1,NA())"
                            )
                        else:
                            evo_sellin_formulas.append(np.nan)

                    df_evolution_excel = pd.DataFrame(
                        {
                            COL_EVO_KANTAR_YOY: evo_kantar_formulas,
                            COL_EVO_SELLIN_YOY: evo_sellin_formulas,
                        },
                        index=df_excel.index[:original_data_rows],
                    )

                    # --- 1.11) Ensamblar DataFrame final para Excel ---
                    # Unir datos originales + coberturas + variaciones de evolución (V, W)
                    df_excel_final = pd.concat([df_excel, df_cov_excel_scaled, df_evolution_excel], axis=1)

                    # Crear la sección de resumen (Variaciones, Promedios, Correlación + Estabilidad)
                    # Añadir filas vacías y reorganizar
                    df_variations_excel['spacer1'] = np.nan
                    # df_averages_excel['spacer2'] = np.nan
                    df_correlations_excel['spacer3'] = np.nan

                    # Aplanar las tablas de resumen para concatenarlas horizontalmente
                    summary_part1 = df_variations_excel.T.reset_index().T # Variaciones
                    summary_part2 = df_averages_excel.T.reset_index().T   # Promedios
                    summary_part3 = df_correlations_excel.T.reset_index().T # Correlaciones

                    # Crear un DataFrame vacío con el número correcto de columnas para alinear
                    max_cols = df_excel_final.shape[1]
                    summary_placeholder = pd.DataFrame(np.nan, index=range(max(len(summary_part1), len(summary_part2), len(summary_part3))), columns=df_excel_final.columns)

                    # Rellenar el placeholder (esto requiere manejo cuidadoso de índices y columnas)
                    # Simplificación: Crear el df_excel_summary_part como antes y concatenar al final
                    df_excel_summary_part = pd.concat([df_variations_excel.reset_index(drop=True),
                                                      df_averages_excel.reset_index(drop=True),
                                                      df_correlations_excel.reset_index(drop=True)], axis=1)

                    # Añadir fila vacía de separación
                    df_excel_final.loc[len(df_excel_final)] = [np.nan] * len(df_excel_final.columns)

                    # Poner "Estabilidad" 2 filas arriba de la fila de encabezado "Correlacion":
                    #   Estabilidad
                    #   (fila en blanco)
                    #   Correlacion / P0..P6 (encabezado)
                    stab_row = {c: np.nan for c in df_excel_summary_part.columns}
                    if "Correlacion" in stab_row:
                        stab_row["Correlacion"] = "Estabilidad"
                        # Asume Cobertura P0-P6 en columnas O a U (después de escalonar)
                        coverage_start_col_idx = 15  # Col O es la 15 (1-based)
                        for p in range(7):
                            key = f"P{p}"
                            if key not in stab_row:
                                continue
                            col_letter = get_column_letter(coverage_start_col_idx + p)
                            # OJO: estos valores ya vienen "escalonados" (pipeline aplicado) por `escalona()`,
                            # así que la estabilidad se calcula con la misma fila para todos los pipelines (como P0).
                            row_last_cov = last_row_excel
                            row_prev_cov = last_row_excel - 12
                            # Requiere 12 meses hacia atrás y suficiente historia para que ambas coberturas existan
                            if row_last_cov >= excel_row_offset and row_prev_cov >= excel_row_offset and (original_data_rows >= (24 + p)):
                                stab_row[key] = f"=IFERROR({col_letter}{row_last_cov}-{col_letter}{row_prev_cov},NA())"
                            else:
                                stab_row[key] = "-"

                    df_stability_above = pd.DataFrame([stab_row], columns=df_excel_summary_part.columns)

                    # Añadir nombres de columnas del resumen como cabecera
                    summary_header = pd.DataFrame([df_excel_summary_part.columns], columns=df_excel_summary_part.columns)
                    df_excel_summary_part_with_header = pd.concat(
                        [df_stability_above, summary_header, df_excel_summary_part],
                        ignore_index=True,
                    )

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

            finally:
                status.stop()

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

                    # Estabilidad (2 decimales + mismos colores rojo/verde que variaciones)
                    # La fila "Estabilidad" queda justo arriba del header "Correlacion".
                    stab_cell = None
                    for row in ws.iter_rows(values_only=False):
                        for cell in row:
                            if isinstance(cell.value, str) and cell.value.strip().lower() == "estabilidad":
                                # Validar que debajo esté el header 'Correlacion' para evitar falsos positivos
                                below = ws.cell(row=cell.row + 1, column=cell.column).value
                                if isinstance(below, str) and below.strip().lower() == "correlacion":
                                    stab_cell = cell
                                    break
                        if stab_cell:
                            break

                    if stab_cell:
                        stab_row = stab_cell.row
                        start_col = stab_cell.column + 1  # P0
                        end_col = start_col + 6           # P6
                        for cc in range(start_col, end_col + 1):
                            ws.cell(row=stab_row, column=cc).number_format = "0.00"
                        stab_range = f"{_col_letter(start_col)}{stab_row}:{_col_letter(end_col)}{stab_row}"
                        ws.conditional_formatting.add(
                            stab_range,
                            _Rule(type="cellIs", operator="lessThan", formula=["0"], dxf=dxf_red),
                        )
                        ws.conditional_formatting.add(
                            stab_range,
                            _Rule(type="cellIs", operator="greaterThan", formula=["0"], dxf=dxf_green),
                        )

                wb2.save(xlsx_path)
            apply_variations_formatting(excel_temp_path)
            print(Fore.GREEN + "Formato de variaciones aplicado (0.0% + rojo/verde).")

            def apply_coverage_values_formatting(xlsx_path: str) -> None:
                """Formatea coberturas (P0..P6), variaciones (V,W) y resalta cortes clave."""
                from openpyxl import load_workbook as _load_wb3
                from openpyxl.styles import PatternFill as _PatternFill
                wb3 = _load_wb3(xlsx_path)
                current_cov_fill = _PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
                prev12_cov_fill = _PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                for ws in wb3.worksheets:
                    data_col = None
                    coverage_cols = []
                    evolution_var_cols = []
                    for col in range(1, ws.max_column + 1):
                        header_value = ws.cell(row=1, column=col).value
                        if header_value is None:
                            continue
                        header_text = str(header_value).strip()
                        if header_text == COL_DATA:
                            data_col = col
                        if header_text in {f"P{i}" for i in range(7)}:
                            coverage_cols.append(col)
                        if header_text in {COL_EVO_KANTAR_YOY, COL_EVO_SELLIN_YOY}:
                            evolution_var_cols.append(col)
                    if data_col is None or (not coverage_cols and not evolution_var_cols):
                        continue

                    row = 2
                    last_data_row = 1
                    while row <= ws.max_row:
                        value = ws.cell(row=row, column=data_col).value
                        if value is None or (isinstance(value, str) and value.strip() == ""):
                            break
                        last_data_row = row
                        row += 1
                    if last_data_row < 2:
                        continue

                    for rr in range(2, last_data_row + 1):
                        for cc in coverage_cols:
                            ws.cell(row=rr, column=cc).number_format = "0.0"
                        for cc in evolution_var_cols:
                            ws.cell(row=rr, column=cc).number_format = "0.0%"

                    # Resaltar coberturas del corte actual y de hace 12 meses para lectura rápida.
                    row_current = last_data_row
                    row_prev12 = last_data_row - 12 if (last_data_row - 12) >= 2 else None
                    for cc in coverage_cols:
                        ws.cell(row=row_current, column=cc).fill = current_cov_fill
                        if row_prev12 is not None:
                            ws.cell(row=row_prev12, column=cc).fill = prev12_cov_fill

                wb3.save(xlsx_path)

            apply_coverage_values_formatting(excel_temp_path)
            print(Fore.GREEN + "Formato aplicado: coberturas (1 decimal) y YoY evolución (0.0%).")

            add_native_excel_charts(
                excel_temp_path,
                coverage_label=coverage_label,
                trend_axis=trend_axis,
                evolution_slide_variant=evolution_slide_variant,
                include_english=include_english,
                pais_nombre=pais_nombre,
            )
            print(Fore.GREEN + "Graficos nativos de Excel insertados (editables).")
        except Exception as e:
            print(Fore.YELLOW + f"No se pudo aplicar el formato de correlaciones: {e}")

    except PermissionError:
        # Usualmente pasa cuando el Excel de salida esta abierto/bloqueado.
        if os.path.exists(excel_temp_path):
            try:
                os.remove(excel_temp_path)
            except Exception:
                pass
        raise
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
    except PermissionError:
        # Archivo destino abierto/bloqueado.
        if os.path.exists(excel_temp_path):
            try:
                os.remove(excel_temp_path)
            except Exception:
                pass
        raise
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
    marca_nombre: Optional[str] = None,
) -> "pd.DataFrame":
    """Calcula la cobertura rolling de 12 meses para cada pipeline."""
    acum_sell_out_py = df_marca[COL_SELL_OUT].rolling(window=12, min_periods=12).sum()
    acum_sell_out_py.index = df_marca[COL_DATA]
    df_coverage = pd.DataFrame(index=acum_sell_out_py.index)
    marca_label = (marca_nombre or "N/D").strip() or "N/D"
    needs_exception_warning = False
    for p in range(7):
        sell_in_shifted = df_marca[COL_SELL_IN].shift(p)
        acum_sell_in_shifted = sell_in_shifted.rolling(window=12, min_periods=12).sum()
        acum_sell_in_shifted.index = df_marca[COL_DATA]
        zero_mask = acum_sell_in_shifted == 0
        if zero_mask.any():
            needs_exception_warning = True
            acum_sell_in_shifted = acum_sell_in_shifted.copy()
            acum_sell_in_shifted.loc[zero_mask] = 1
        coverage_p = (acum_sell_out_py / acum_sell_in_shifted) * 100
        coverage_p = coverage_p.replace([np.inf, -np.inf], np.nan)
        df_coverage[f'P{p}'] = coverage_p
    pop_val_num = float(pop_coverage.get(pais_nombre, DEFAULT_POP_COVERAGE).replace('%', '')) / 100.0
    if coverage_type.lower() == "relativa" and pop_val_num > 0:
        df_coverage = df_coverage / pop_val_num
    if round_coverage:
        df_coverage = df_coverage.apply(_round_half_up_series)
    else:
        df_coverage = df_coverage.round(1)
    if needs_exception_warning:
        notify_zero_months_exception(marca_label)
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
        averages['Freq_MAT_Actual'] = df_marca[COL_FREQ].iloc[-12:].mean()
    else:
        averages['Penet_MAT_Actual'] = df_marca[COL_PENET].mean()
        averages['Buyers_MAT_Actual'] = df_marca[COL_BUYERS].mean()
        averages['Freq_MAT_Actual'] = df_marca[COL_FREQ].mean()
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


def build_evolution_figure(
    df_marca: "pd.DataFrame",
    pipeline: int,
    lang_index: int,
    marca_nombre: str,
    variant: str = "classic",
) -> Optional["plt.Figure"]:
    if len(df_marca) < 24:
        return None
    df_evol = df_marca[[COL_DATA, COL_SELL_IN, COL_SELL_OUT]].copy()
    df_evol[COL_DATA] = pd.to_datetime(df_evol[COL_DATA])
    return generar_grafico_evolucion_mensual(
        df_evol,
        pipeline,
        lang_index,
        marca_nombre=marca_nombre,
        variant=variant,
    )


def _coverage_value_for_year_month(coverage_series: "pd.Series", year: int, month: int) -> float:
    if coverage_series is None or coverage_series.empty:
        return np.nan
    idx = pd.to_datetime(coverage_series.index, errors="coerce")
    values = pd.to_numeric(coverage_series, errors="coerce")
    clean_series = pd.Series(values.to_numpy(dtype=float), index=idx).dropna()
    clean_series = clean_series[~clean_series.index.isna()]
    if clean_series.empty:
        return np.nan
    matched = clean_series[(clean_series.index.year == year) & (clean_series.index.month == month)]
    if matched.empty:
        return np.nan
    return float(matched.iloc[-1])


def _coverage_value_to_number(value: float, round_coverage: bool) -> float | int:
    if pd.notna(value):
        return int(np.floor(float(value) + 0.5)) if round_coverage else round(float(value), 1)
    return 0


def _coverage_value_to_text(value: float, round_coverage: bool) -> str:
    if pd.notna(value):
        return str(int(np.floor(float(value) + 0.5))) if round_coverage else f"{float(value):.1f}"
    return "0" if round_coverage else "0.0"



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
    summary_extra_months: Sequence[int],
    summary_extra_months_mode: str,
) -> Tuple[Dict[str, str], Dict[str, object], float, float, float, str]:
    ref_dt = dt.strptime(ref_month_year, '%m-%y')
    summary_columns, coverage_periods, _ = build_summary_columns(
        lang_index=lang_index,
        fabricante=fabricante,
        ref_dt=ref_dt,
        summary_extra_months=summary_extra_months,
        summary_extra_months_mode=summary_extra_months_mode,
    )
    _, _, base_prev, base_curr = build_summary_coverage_periods(
        ref_dt,
        summary_extra_months,
        summary_extra_months_mode,
    )
    coverage_anterior = _coverage_value_for_year_month(coverage_series, base_prev.year, base_prev.month)
    coverage_actual = _coverage_value_for_year_month(coverage_series, base_curr.year, base_curr.month)

    var_cliente_anual_y1 = df_variations.loc[df_variations['Tipo'] == 'Anual', f'Cliente P{pipeline}'].iloc[0]
    var_kantar_anual_y1 = df_variations.loc[df_variations['Tipo'] == 'Anual', 'WP by Numerator'].iloc[0]
    tendencia_alineada = "NO"
    if pd.notna(var_cliente_anual_y1) and pd.notna(var_kantar_anual_y1):
        if (var_cliente_anual_y1 * var_kantar_anual_y1) > 0:
            tendencia_alineada = "SI"
        elif var_cliente_anual_y1 == 0 and var_kantar_anual_y1 == 0:
            tendencia_alineada = "SI"

    cov_actual_val = _coverage_value_to_number(coverage_actual, round_coverage)
    cov_anterior_val = _coverage_value_to_number(coverage_anterior, round_coverage)
    estabilidad = (cov_actual_val - cov_anterior_val) if round_coverage else round(cov_actual_val - cov_anterior_val, 1)

    summary_row = {
        summary_columns[0]: marca_nombre_limpio,
        summary_columns[1]: pipeline,
        summary_columns[2]: f"{averages.get('Penet_MAT_Actual', 0):.1f}%",
        summary_columns[3]: f"{var_cliente_anual_y1*100:.1f}%" if pd.notna(var_cliente_anual_y1) else "0.0%",
        summary_columns[4]: f"{var_kantar_anual_y1*100:.1f}%" if pd.notna(var_kantar_anual_y1) else "0.0%",
    }

    coverage_col_idx = 5
    for period_dt in coverage_periods:
        cov_value = _coverage_value_for_year_month(coverage_series, period_dt.year, period_dt.month)
        summary_row[summary_columns[coverage_col_idx]] = _coverage_value_to_text(cov_value, round_coverage)
        coverage_col_idx += 1
    summary_row[summary_columns[-1]] = str(estabilidad) if round_coverage else f"{estabilidad:.1f}"

    banco_row = {
        'Periodo': ref_dt.date(),
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
        'Frecuencia Media Mensual': round(averages.get('Freq_MAT_Actual', 0), 1),
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
    'Raw Buyers Media Ano Mov Atual', 'Frecuencia Media Mensual', 'Pipeline', 'Cobertura Año Mov Actual',
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
    variations_box_style: str,
    coverage_slide_variant: str,
    evolution_slide_variant: str,
    round_coverage: bool,
    summary_extra_months: Sequence[int],
    summary_extra_months_mode: str,
) -> Tuple[str, "pd.DataFrame", "pd.DataFrame"]:
    chosen_lang, lang_index = determine_language(include_english, pais_nombre)
    ppt, tmp_ppt_path = copy_and_prune_template(root_dir, chosen_lang)
    labels = build_labels(lang_index, fabricante, ref_month_year, summary_extra_months, summary_extra_months_mode)
    builder = SlideBuilder(
        ppt,
        lang_index,
        labels,
        coverage_label,
        coverage_type=coverage_type,
        ref_month_year=ref_month_year,
        tipo_eje_tend=trend_axis,
        variations_box_style=variations_box_style,
        coverage_slide_variant=coverage_slide_variant,
    )
    builder.configure_cover(pais_nombre, fabricante, categoria_nombre, ref_month_year, chosen_lang)

    summary_rows: List[Dict[str, str]] = []
    bank_rows: List[Dict[str, object]] = []
    low_penetration_brands: List[str] = []

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
            issues_detected = detect_brand_data_issues(df_marca_ppt, window=0)
            if issues_detected:
                for issue in issues_detected:
                    if issue == "zero_dash":
                        notify_zero_months_exception(marca_nombre_limpio)
                    elif issue == "negative":
                        notify_negative_values_exception(marca_nombre_limpio)
            df_coverage = compute_coverage_dataframe(
                df_marca_ppt,
                pais_nombre,
                coverage_type,
                round_coverage,
                marca_nombre=marca_nombre_limpio,
            )
            df_variations = compute_variations_dataframe(df_marca_ppt)
            averages = compute_averages(df_marca_ppt)
            notify_buyers_threshold(marca_nombre_limpio, averages.get('Buyers_MAT_Actual'))
            try:
                buyers_val = averages.get('Buyers_MAT_Actual')
                if buyers_val is not None and not pd.isna(buyers_val) and float(buyers_val) < 200:
                    if marca_nombre_limpio not in low_penetration_brands:
                        low_penetration_brands.append(marca_nombre_limpio)
            except Exception:
                pass
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
                evolution_figure = build_evolution_figure(
                    df_marca_ppt,
                    pipeline,
                    lang_index,
                    marca_nombre_limpio,
                    variant=evolution_slide_variant,
                )
                assets = PipelineAssets(
                    pipeline=pipeline,
                    marca=marca_nombre_limpio,
                    coverage_series=coverage_series,
                    penetration_series=df_marca_ppt.set_index(COL_DATA)[COL_PENET].loc[coverage_series.dropna().index],
                    variation_table=variation_table,
                    trend_plot_df=df_trend_plot,
                    variations_detail=variations_detail,
                    evolution_figure=evolution_figure,
                    buyers_mat_actual=averages.get('Buyers_MAT_Actual'),
                    penet_mat_actual=averages.get('Penet_MAT_Actual'),
                    penet_mat_anterior=averages.get('Penet_MAT_Anterior'),
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
                    summary_extra_months=summary_extra_months,
                    summary_extra_months_mode=summary_extra_months_mode,
                )
                summary_rows.append(summary_row)
                bank_rows.append(bank_row)
        progress.update(task_id, advance=1)

    df_summary = pd.DataFrame(summary_rows)
    if not df_summary.empty:
        df_summary = df_summary[labels[(lang_index, 'Summary')]]
    df_bank = pd.DataFrame(bank_rows, columns=COVERAGE_BANK_COLUMNS)

    builder.add_summary_slide(df_summary, pais_nombre, categoria_nombre, low_penetration_brands=low_penetration_brands)
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
        self._script_start_monotonic: Optional[float] = None

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
            variations_box_style = normalize_variations_box_style(
                next((os.environ.get(k) for k in VARIATIONS_BOX_STYLE_ENV_KEYS if os.environ.get(k) is not None), None)
            )
            coverage_slide_variant = normalize_coverage_slide_variant(
                next((os.environ.get(k) for k in COVERAGE_SLIDE_VARIANT_ENV_KEYS if os.environ.get(k) is not None), None)
            )
            evolution_slide_variant = normalize_evolution_slide_variant(
                next((os.environ.get(k) for k in EVOLUTION_SLIDE_VARIANT_ENV_KEYS if os.environ.get(k) is not None), None)
            )
            include_english = False
            round_cov = False
            SELECTIONS['Razón'] = coverage_reason
            SELECTIONS['Eje tendencia'] = trend_axis
            SELECTIONS['Idioma PPT'] = 'ESPAÑOL'
            SELECTIONS['Inglés'] = 'No'
            SELECTIONS['Redondeo Cobertura'] = 'No'
            SELECTIONS["Estilo variaciones"] = "Bonito" if variations_box_style == "pretty" else "Clasico"
            SELECTIONS["Slide Cobertura"] = "Complementado" if coverage_slide_variant == "complemented" else "Clasico"
            SELECTIONS["Slide Evolucion"] = "Simple" if evolution_slide_variant == "simple" else "Clasico/Avanzado"
            summary_extra_months = get_summary_extra_months_from_env()
            SELECTIONS['Meses extra summary'] = format_summary_extra_months(summary_extra_months)
            summary_extra_months_mode = get_summary_extra_months_mode_from_env() or "recent"
            if summary_extra_months:
                SELECTIONS['Modo meses extra summary'] = "Mes más reciente" if summary_extra_months_mode == "recent" else "Año actual y anterior"
            else:
                # Evitar confusión y arrastre de estado de corridas anteriores.
                SELECTIONS.pop('Modo meses extra summary', None)
            clear_and_print_summary()
        else:
            coverage_type_value = coverage_type
            coverage_reason = razao_cov()
            trend_axis = tipo_eje_tendencia()
            variations_box_style = variations_box_style_option()
            coverage_slide_variant = coverage_slide_variant_option()
            evolution_slide_variant = evolution_slide_variant_option()
            include_english = include_english_flag()
            round_cov = round_coverage_flag()
            summary_extra_months = summary_extra_months_option()
            summary_extra_months_mode = summary_extra_months_mode_option(bool(summary_extra_months))
        return ExecutionOptions(
            coverage_type=coverage_type_value,
            coverage_reason=coverage_reason,
            trend_axis=trend_axis,
            variations_box_style=variations_box_style,
            include_english=include_english,
            round_coverage=round_cov,
            coverage_slide_variant=coverage_slide_variant,
            evolution_slide_variant=evolution_slide_variant,
            summary_extra_months=summary_extra_months,
            summary_extra_months_mode=summary_extra_months_mode,
            auto_mode=auto_mode,
        )


    def process_file(self, excel_file_name: str, options: ExecutionOptions, idx: int, total: int) -> None:
        global ROUND_COVERAGE
        ROUND_COVERAGE = options.round_coverage
        self.ensure_categories_loaded()
        excel_file_path = os.path.join(self.root_dir, excel_file_name)
        elapsed = None
        if self._script_start_monotonic is not None:
            try:
                elapsed = time.monotonic() - float(self._script_start_monotonic)
            except Exception:
                elapsed = None
        try:
            try:
                excel_file_obj = pd.ExcelFile(excel_file_path)
                marcas = excel_file_obj.sheet_names
            except FileNotFoundError:
                print(f"{Fore.RED}{Style.BRIGHT}Error: No se encontró el archivo seleccionado: {excel_file_path}")
                return
            except PermissionError:
                # Input bloqueado (raro) o sin permisos.
                print_file_locked_error(excel_file_path, elapsed_seconds=elapsed)
                return
            except Exception as exc:
                print(f"{Fore.RED}{Style.BRIGHT}Error al abrir el archivo Excel '{excel_file_name}': {exc}")
                return

            try:
                pais_nombre, cesta_nombre, categoria_nombre, categoria_nombre_corto, fabricante = parse_file_metadata(excel_file_name, self.categories)
            except ValueError as exc:
                print(f"{Fore.RED}{Style.BRIGHT}{exc}")
                return

            # Asegurar que el resumen refleje opciones también en modo AUTO (sin selección interactiva).
            SELECTIONS['Excel'] = excel_file_name
            SELECTIONS['Cobertura'] = options.coverage_type
            SELECTIONS['Razón'] = options.coverage_reason
            SELECTIONS['Eje tendencia'] = options.trend_axis
            SELECTIONS['Inglés'] = 'Sí' if options.include_english else 'No'
            SELECTIONS['Redondeo Cobertura'] = 'Sí' if options.round_coverage else 'No'
            SELECTIONS['Estilo variaciones'] = "Bonito" if normalize_variations_box_style(options.variations_box_style) == "pretty" else "Clasico"
            SELECTIONS['Slide Cobertura'] = "Complementado" if normalize_coverage_slide_variant(options.coverage_slide_variant) == "complemented" else "Clasico"
            SELECTIONS['Slide Evolucion'] = "Simple" if normalize_evolution_slide_variant(options.evolution_slide_variant) == "simple" else "Clasico/Avanzado"
            if options.summary_extra_months:
                SELECTIONS['Meses extra summary'] = format_summary_extra_months(options.summary_extra_months)
                SELECTIONS['Modo meses extra summary'] = "Mes más reciente" if options.summary_extra_months_mode == "recent" else "Año actual y anterior"
            else:
                SELECTIONS['Meses extra summary'] = "Ninguno"
                SELECTIONS.pop('Modo meses extra summary', None)

            SELECTIONS['Pais'] = pais_nombre
            clear_and_print_summary()
            print_file_header(idx, total, excel_file_name)

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
                options.trend_axis,
                options.evolution_slide_variant,
                options.include_english,
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
                variations_box_style=options.variations_box_style,
                coverage_slide_variant=options.coverage_slide_variant,
                evolution_slide_variant=options.evolution_slide_variant,
                round_coverage=options.round_coverage,
                summary_extra_months=options.summary_extra_months,
                summary_extra_months_mode=options.summary_extra_months_mode,
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
            print_file_summary(ruta_template_final, ruta_ppt_final, ruta_banco_final, elapsed_seconds=elapsed)
            report_zero_months_exceptions()
        except PermissionError as exc:
            locked_path = getattr(exc, "filename", None) or str(exc)
            print_file_locked_error(locked_path, elapsed_seconds=elapsed)
            return

    def run(self) -> None:
        if self._script_start_monotonic is None:
            self._script_start_monotonic = time.monotonic()
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
    start_mono = time.monotonic()
    try:
        app._script_start_monotonic = start_mono
        app.run()
    except PermissionError as exc:
        locked_path = getattr(exc, "filename", None) or str(exc)
        print_file_locked_error(locked_path, elapsed_seconds=(time.monotonic() - start_mono))
        try:
            cleanup_temp_dir(app.root_dir)
        except Exception:
            pass
    except KeyboardInterrupt:
        end_time = datetime.now().strftime("%I:%M:%S %p")
        elapsed = _format_elapsed(time.monotonic() - start_mono)
        msg = (
            "[bright_white]Programa terminado por el usuario[/bright_white]\n\n"
            f"[white]Hora de finalizacion: [bold]{end_time}[/bold][/white]\n"
            f"[white]Tiempo total: [bold]{elapsed}[/bold][/white]\n\n"
            "[grey]Hasta luego.[/grey]"
        )
        console.print()
        console.print(Panel.fit(msg, border_style="yellow", title="Coverages Latam"))
        console.print()
        try:
            cleanup_temp_dir(app.root_dir)
        except Exception:
            pass


if __name__ == "__main__":
    main()
