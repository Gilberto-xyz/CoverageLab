"""
===========================================
Automatizador de Nomenclatura de Archivos
===========================================

Este script automatiza la creación y nomenclatura de archivos Excel para el ejercicio de coberturas,
implementando un sistema estandarizado basado en códigos de países y categorías.

Funcionalidades Principales:
--------------------------
1. Selección asistida de país y obtención de su código
2. Búsqueda y selección de categorías de productos
3. Generación automática de nombres de archivo estandarizados
4. Creación de archivos Excel con estructura predefinida

Componentes Principales:
----------------------
- Diccionario de países (countries): Mapeo de países y sus códigos
- Lista de categorías (categorias): Categorías de productos con sus códigos
- Clase Colors: Códigos ANSI para formato de texto en consola

Funciones Principales:
--------------------
- obtener_codigo_pais(): Obtiene el código del país desde la entrada del usuario
- buscar_categorias(): Busca categorías basadas en palabras clave
- seleccionar_categoria(): Maneja la selección de categoría por el usuario
- crear_excel(): Genera el archivo Excel con la estructura requerida
- main(): Función principal que coordina el flujo del programa

Ejemplo de Uso:
-------------
1. Ejecutar el script
2. Ingresar país (ej: "mexico" o "mex")
3. Buscar categoría por palabra clave
4. Seleccionar categoría de la lista
5. Ingresar nombre del fabricante
6. El script generará automáticamente el archivo Excel con el nombre estandarizado

Formato del nombre de archivo resultante:
[código_país]_[código_categoría]_[fabricante].xlsx
Ejemplo: 52_SODA_COCACOLA.xlsx
"""
# script_optimizado.py
import os
import unicodedata
import sys

# Para la parte de colores ANSI (si lo deseas):
class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn').lower()

countries = {
    'latam': '10', 'lat': '10',
    'argentina': '54', 'arg': '54',
    'bolivia': '91', 'bol': '91',
    'brasil': '55', 'bra': '55', 'brazil': '55',
    'cam': '12',
    'chile': '56', 'chl': '56',
    'colombia': '57', 'col': '57',
    'ecuador': '93', 'ecu': '93',
    'mexico': '52', 'mex': '52', 'mx': '52',
    'peru': '51', 'per': '51'
}

def obtener_nombre_pais(key):
    mapping = {
        'latam': 'LatAm', 'lat': 'LatAm',
        'argentina': 'Argentina', 'arg': 'Argentina',
        'bolivia': 'Bolivia', 'bol': 'Bolivia',
        'brasil': 'Brasil', 'bra': 'Brasil', 'brazil': 'Brasil',
        'cam': 'CAM',
        'chile': 'Chile', 'chl': 'Chile',
        'colombia': 'Colombia', 'col': 'Colombia',
        'ecuador': 'Ecuador', 'ecu': 'Ecuador',
        'mexico': 'Mexico', 'mex': 'Mexico', 'mx': 'Mexico',
        'peru': 'Peru', 'per': 'Peru'
    }
    return mapping.get(key, "Desconocido")

def obtener_codigo_pais(input_pais):
    input_normalizado = strip_accents(input_pais)
    if input_normalizado in countries:
        return countries[input_normalizado], obtener_nombre_pais(input_normalizado)
    else:
        # Intentar buscar por substring
        matches = [key for key in countries.keys() if input_normalizado in key]
        if len(matches) == 1:
            return countries[matches[0]], obtener_nombre_pais(matches[0])
        elif len(matches) > 1:
            print(f"{Colors.WARNING}Se encontraron múltiples países:{Colors.ENDC}")
            for idx, key in enumerate(matches, 1):
                print(f"{Colors.OKGREEN}{idx}. {obtener_nombre_pais(key)}{Colors.ENDC}")
            try:
                seleccion = int(input(f"{Colors.OKBLUE}Seleccione el país (número): {Colors.ENDC}"))
                if 1 <= seleccion <= len(matches):
                    selected_key = matches[seleccion - 1]
                    return countries[selected_key], obtener_nombre_pais(selected_key)
                else:
                    print(f"{Colors.FAIL}Selección inválida.{Colors.ENDC}")
                    return None, None
            except ValueError:
                print(f"{Colors.FAIL}Entrada inválida.{Colors.ENDC}")
                return None, None
        else:
            return None, None

# Lista de categorías con sus descripciones y códigos
categorias = [
    {'categoria': 'Alimentos', 'descripcion': 'Carne Fresca', 'cod': 'MEAT'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Pañitos + Pañales', 'cod': 'CRDT'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Snacks)', 'cod': 'CRSN'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas Alcohólicas', 'cod': 'ALCB'},
    {'categoria': 'Bebidas', 'descripcion': 'Cervezas', 'cod': 'BEER'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas Gaseosas', 'cod': 'CARB'},
    {'categoria': 'Bebidas', 'descripcion': 'Agua Gasificada', 'cod': 'CWAT'},
    {'categoria': 'Bebidas', 'descripcion': 'Água de Coco', 'cod': 'COCW'},
    {'categoria': 'Bebidas', 'descripcion': 'Café_Consolidado de Café', 'cod': 'COFF'},
    {'categoria': 'Bebidas', 'descripcion': 'Cross Category (Bebidas)', 'cod': 'CRBE'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas Energéticas', 'cod': 'ENDR'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas Saborizadas Sin Gas', 'cod': 'FLBE'},
    {'categoria': 'Bebidas', 'descripcion': 'Café Tostado y Molido', 'cod': 'GCOF'},
    {'categoria': 'Bebidas', 'descripcion': 'Jugos Caseros', 'cod': 'HJUI'},
    {'categoria': 'Bebidas', 'descripcion': 'Té Helado', 'cod': 'ITEA'},
    {'categoria': 'Bebidas', 'descripcion': 'Café Instantáneo_Café Sucedáneo', 'cod': 'ICOF'},
    {'categoria': 'Bebidas', 'descripcion': 'Jugos y Nectares', 'cod': 'JUNE'},
    {'categoria': 'Bebidas', 'descripcion': 'Zumos de Vegetales', 'cod': 'VEJU'},
    {'categoria': 'Bebidas', 'descripcion': 'Agua Natural', 'cod': 'WATE'},
    {'categoria': 'Bebidas', 'descripcion': 'Gaseosas + Aguas', 'cod': 'CSDW'},
    {'categoria': 'Bebidas', 'descripcion': 'Mixta Café+Malta', 'cod': 'MXCM'},
    {'categoria': 'Bebidas', 'descripcion': 'Mixta Dolce Gusto_Mixta Té Helado + Café + Modificadores', 'cod': 'MXDG'},
    {'categoria': 'Bebidas', 'descripcion': 'Mixta Jugos y Leches', 'cod': 'MXJM'},
    {'categoria': 'Bebidas', 'descripcion': 'Mixta Jugos Líquidos + Bebidas de Soja', 'cod': 'MXJS'},
    {'categoria': 'Bebidas', 'descripcion': 'Mixta Té+Café', 'cod': 'MXTC'},
    {'categoria': 'Bebidas', 'descripcion': 'Jugos Liquidos_Jugos Polvo', 'cod': 'JUIC'},
    {'categoria': 'Bebidas', 'descripcion': 'Refrescos en Polvo_Jugos _ Bebidas Instantáneas En Polvo _ Jugos Polvo', 'cod': 'PWDJ'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas Refrescantes', 'cod': 'RFDR'},
    {'categoria': 'Bebidas', 'descripcion': 'Refrescos Líquidos_Jugos Líquidos', 'cod': 'RTDJ'},
    {'categoria': 'Bebidas', 'descripcion': 'Té Líquido _ Listo para Tomar', 'cod': 'RTEA'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas de Soja', 'cod': 'SOYB'},
    {'categoria': 'Bebidas', 'descripcion': 'Bebidas Isotónicas', 'cod': 'SPDR'},
    {'categoria': 'Bebidas', 'descripcion': 'Té e Infusiones_Te_Infusión Hierbas', 'cod': 'TEAA'},
    {'categoria': 'Bebidas', 'descripcion': 'Yerba Mate', 'cod': 'YERB'},
    {'categoria': 'Lacteos', 'descripcion': 'Manteca', 'cod': 'BUTT'},
    {'categoria': 'Lacteos', 'descripcion': 'Queso Fresco y para Untar', 'cod': 'CHEE'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche Condensada', 'cod': 'CMLK'},
    {'categoria': 'Lacteos', 'descripcion': 'Queso Untable', 'cod': 'CRCH'},
    {'categoria': 'Lacteos', 'descripcion': 'Yoghurt p_beber', 'cod': 'DYOG'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche Culinaria_Leche Evaporada', 'cod': 'EMLK'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche Fermentada', 'cod': 'FRMM'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche Líquida Saborizada_Leche Líquida Con Sabor', 'cod': 'FMLK'},
    {'categoria': 'Lacteos', 'descripcion': 'Fórmulas Infantiles', 'cod': 'FRMK'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche Líquida', 'cod': 'LQDM'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche Larga Vida', 'cod': 'LLFM'},
    {'categoria': 'Lacteos', 'descripcion': 'Margarina', 'cod': 'MARG'},
    {'categoria': 'Lacteos', 'descripcion': 'Queso Fundido', 'cod': 'MCHE'},
    {'categoria': 'Lacteos', 'descripcion': 'Crema de Leche', 'cod': 'MKCR'},
    {'categoria': 'Lacteos', 'descripcion': 'Mixta Lácteos_Postre+Leches+Yogurt', 'cod': 'MXDI'},
    {'categoria': 'Lacteos', 'descripcion': 'Mixta Leches', 'cod': 'MXMI'},
    {'categoria': 'Lacteos', 'descripcion': 'Mixta Yoghurt+Postres', 'cod': 'MXYD'},
    {'categoria': 'Lacteos', 'descripcion': 'Petit Suisse', 'cod': 'PTSS'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche en Polvo', 'cod': 'PWDM'},
    {'categoria': 'Lacteos', 'descripcion': 'Yoghurt p_comer', 'cod': 'SYOG'},
    {'categoria': 'Lacteos', 'descripcion': 'Leche_Leche Líquida Blanca _ Leche Liq. Natural', 'cod': 'MILK'},
    {'categoria': 'Lacteos', 'descripcion': 'Yoghurt', 'cod': 'YOGH'},
    {'categoria': 'Ropas y Calzados', 'descripcion': 'Ropas', 'cod': 'CLOT'},
    {'categoria': 'Ropas y Calzados', 'descripcion': 'Calzados', 'cod': 'FOOT'},
    {'categoria': 'Ropas y Calzados', 'descripcion': 'Medias_Calcetines', 'cod': 'SOCK'},
    {'categoria': 'Alimentos', 'descripcion': 'Arepas', 'cod': 'AREP'},
    {'categoria': 'Alimentos', 'descripcion': 'Cereales Infantiles', 'cod': 'BCER'},
    {'categoria': 'Alimentos', 'descripcion': 'Nutrición Infantil_Colados y Picados', 'cod': 'BABF'},
    {'categoria': 'Alimentos', 'descripcion': 'Frijoles', 'cod': 'BEAN'},
    {'categoria': 'Alimentos', 'descripcion': 'Galletas', 'cod': 'BISC'},
    {'categoria': 'Alimentos', 'descripcion': 'Caldos_Caldos y Sazonadores', 'cod': 'BOUI'},
    {'categoria': 'Alimentos', 'descripcion': 'Pan', 'cod': 'BREA'},
    {'categoria': 'Alimentos', 'descripcion': 'Apanados_Empanizadores', 'cod': 'BRCR'},
    {'categoria': 'Alimentos', 'descripcion': 'Empanados', 'cod': 'BRDC'},
    {'categoria': 'Alimentos', 'descripcion': 'Cereales_Cereales Desayuno_Avenas y Cereales', 'cod': 'CERE'},
    {'categoria': 'Alimentos', 'descripcion': 'Hamburguesas', 'cod': 'BURG'},
    {'categoria': 'Alimentos', 'descripcion': 'Mezclas Listas para Tortas_Preparados Base Harina Trigo', 'cod': 'CCMX'},
    {'categoria': 'Alimentos', 'descripcion': 'Queques_Ponques Industrializados', 'cod': 'CAKE'},
    {'categoria': 'Alimentos', 'descripcion': 'Conservas De Pescado', 'cod': 'FISH'},
    {'categoria': 'Alimentos', 'descripcion': 'Conservas de Frutas y Verduras', 'cod': 'CFAV'},
    {'categoria': 'Alimentos', 'descripcion': 'Dulce de Leche_Manjar', 'cod': 'CRML'},
    {'categoria': 'Alimentos', 'descripcion': 'Alfajores', 'cod': 'CMLC'},
    {'categoria': 'Alimentos', 'descripcion': 'Barras de Cereal', 'cod': 'CBAR'},
    {'categoria': 'Alimentos', 'descripcion': 'Pollo', 'cod': 'CHCK'},
    {'categoria': 'Alimentos', 'descripcion': 'Chocolate', 'cod': 'CHOC'},
    {'categoria': 'Alimentos', 'descripcion': 'Chocolate de Taza_Achocolatados _ Cocoas', 'cod': 'COCO'},
    {'categoria': 'Alimentos', 'descripcion': 'Salsas Frías', 'cod': 'COLS'},
    {'categoria': 'Alimentos', 'descripcion': 'Compotas', 'cod': 'COMP'},
    {'categoria': 'Alimentos', 'descripcion': 'Condimentos y Especias', 'cod': 'SPIC'},
    {'categoria': 'Alimentos', 'descripcion': 'Chocolate de Mesa', 'cod': 'CKCH'},
    {'categoria': 'Alimentos', 'descripcion': 'Aceite_Aceites Comestibles', 'cod': 'COIL'},
    {'categoria': 'Alimentos', 'descripcion': 'Salsas Listas_Salsas Caseras Envasadas', 'cod': 'CSAU'},
    {'categoria': 'Alimentos', 'descripcion': 'Grano, Harina y Masa de Maíz', 'cod': 'CNML'},
    {'categoria': 'Alimentos', 'descripcion': 'Fécula de Maíz', 'cod': 'CNST'},
    {'categoria': 'Alimentos', 'descripcion': 'Harina De Maíz', 'cod': 'CNFL'},
    {'categoria': 'Alimentos', 'descripcion': 'Ayudantes Culinarios', 'cod': 'CAID'},
    {'categoria': 'Alimentos', 'descripcion': 'Postres Preparados', 'cod': 'DESS'},
    {'categoria': 'Alimentos', 'descripcion': 'Jamón Endiablado', 'cod': 'DHAM'},
    {'categoria': 'Alimentos', 'descripcion': 'Semillas y Frutos Secos', 'cod': 'DFNS'},
    {'categoria': 'Alimentos', 'descripcion': 'Pan de Pascua', 'cod': 'EBRE'},
    {'categoria': 'Alimentos', 'descripcion': 'Huevos de Páscua', 'cod': 'EEGG'},
    {'categoria': 'Alimentos', 'descripcion': 'Huevos', 'cod': 'EGGS'},
    {'categoria': 'Alimentos', 'descripcion': 'Flash Cecinas', 'cod': 'FLSS'},
    {'categoria': 'Alimentos', 'descripcion': 'Harinas', 'cod': 'FLOU'},
    {'categoria': 'Alimentos', 'descripcion': 'Platos Listos Congelados', 'cod': 'FRDS'},
    {'categoria': 'Alimentos', 'descripcion': 'Alimentos Congelados', 'cod': 'FRFO'},
    {'categoria': 'Alimentos', 'descripcion': 'Jamones', 'cod': 'HAMS'},
    {'categoria': 'Alimentos', 'descripcion': 'Cereales Calientes_Cereales Precocidos', 'cod': 'HCER'},
    {'categoria': 'Alimentos', 'descripcion': 'Salsas Picantes', 'cod': 'HOTS'},
    {'categoria': 'Alimentos', 'descripcion': 'Helados', 'cod': 'ICEC'},
    {'categoria': 'Alimentos', 'descripcion': 'Pan Industrializado', 'cod': 'IBRE'},
    {'categoria': 'Alimentos', 'descripcion': 'Puré Instantáneo', 'cod': 'IMPO'},
    {'categoria': 'Alimentos', 'descripcion': 'Fideos Instantáneos', 'cod': 'INOO'},
    {'categoria': 'Alimentos', 'descripcion': 'Mermeladas', 'cod': 'JAMS'},
    {'categoria': 'Alimentos', 'descripcion': 'Ketchup', 'cod': 'KETC'},
    {'categoria': 'Alimentos', 'descripcion': 'Jugo de Limon Adereso', 'cod': 'LJDR'},
    {'categoria': 'Alimentos', 'descripcion': 'Maltas', 'cod': 'MALT'},
    {'categoria': 'Alimentos', 'descripcion': 'Adobos _ Sazonadores', 'cod': 'SEAS'},
    {'categoria': 'Alimentos', 'descripcion': 'Mayonesa', 'cod': 'MAYO'},
    {'categoria': 'Alimentos', 'descripcion': 'Cárnicos', 'cod': 'MEAT'},
    {'categoria': 'Alimentos', 'descripcion': 'Modificadores de Leche_Saborizadores p_leche', 'cod': 'MLKM'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Cereales Infantiles+Avenas', 'cod': 'MXCO'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Caldos + Saborizantes', 'cod': 'MXBS'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Caldos + Sopas', 'cod': 'MXSB'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Cereales + Cereales Calientes', 'cod': 'MXCH'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Chocolate + Manjar', 'cod': 'MXCC'},
    {'categoria': 'Alimentos', 'descripcion': 'Galletas, snacks y mini tostadas', 'cod': 'MXSN'},
    {'categoria': 'Alimentos', 'descripcion': 'Aceites + Mantecas', 'cod': 'COBT'},
    {'categoria': 'Alimentos', 'descripcion': 'Aceites + Conservas De Pescado', 'cod': 'COCF'},
    {'categoria': 'Alimentos', 'descripcion': 'Ayudantes Culinarios + Bolsa de Hornear', 'cod': 'CABB'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Huevos de Páscua + Chocolates', 'cod': 'MXEC'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Platos Listos Congelados + Pasta', 'cod': 'MXDP'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Platos Congelados y Listos para Comer', 'cod': 'MXFR'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Alimentos Congelados + Margarina', 'cod': 'MXFM'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Modificadores + Cocoa', 'cod': 'MXMC'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Pastas', 'cod': 'MXPS'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Sopas+Cremas+Ramen', 'cod': 'MXSO'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Margarina + Mayonesa + Queso Crema', 'cod': 'MXSP'},
    {'categoria': 'Alimentos', 'descripcion': 'Mixta Azúcar+Endulzantes', 'cod': 'MXSW'},
    {'categoria': 'Alimentos', 'descripcion': 'Mostaza', 'cod': 'MUST'},
    {'categoria': 'Alimentos', 'descripcion': 'Sustitutos de Crema', 'cod': 'NDCR'},
    {'categoria': 'Alimentos', 'descripcion': 'Fideos', 'cod': 'NOOD'},
    {'categoria': 'Alimentos', 'descripcion': 'Nuggets', 'cod': 'NUGG'},
    {'categoria': 'Alimentos', 'descripcion': 'Avena en hojuelas_liquidas', 'cod': 'OAFL'},
    {'categoria': 'Alimentos', 'descripcion': 'Aceitunas', 'cod': 'OLIV'},
    {'categoria': 'Alimentos', 'descripcion': 'Tortilla', 'cod': 'PANC'},
    {'categoria': 'Alimentos', 'descripcion': 'Panetón', 'cod': 'PANE'},
    {'categoria': 'Alimentos', 'descripcion': 'Pastas', 'cod': 'PAST'},
    {'categoria': 'Alimentos', 'descripcion': 'Salsas para Pasta', 'cod': 'PSAU'},
    {'categoria': 'Alimentos', 'descripcion': 'Turrón de maní', 'cod': 'PNOU'},
    {'categoria': 'Alimentos', 'descripcion': 'Carne Porcina', 'cod': 'PORK'},
    {'categoria': 'Alimentos', 'descripcion': 'Postres en Polvo_Postres para Preparar _ Horneables-Gelificables', 'cod': 'PPMX'},
    {'categoria': 'Alimentos', 'descripcion': 'Leche de Soya en Polvo', 'cod': 'PWSM'},
    {'categoria': 'Alimentos', 'descripcion': 'Cereales Precocidos', 'cod': 'PCCE'},
    {'categoria': 'Alimentos', 'descripcion': 'Masas Frescas_Tapas Empanadas y Tarta', 'cod': 'DOUG'},
    {'categoria': 'Alimentos', 'descripcion': 'Pre-Pizzas', 'cod': 'PPIZ'},
    {'categoria': 'Alimentos', 'descripcion': 'Meriendas listas', 'cod': 'REFR'},
    {'categoria': 'Alimentos', 'descripcion': 'Arroz', 'cod': 'RICE'},
    {'categoria': 'Alimentos', 'descripcion': 'Galletas de Arroz', 'cod': 'RBIS'},
    {'categoria': 'Alimentos', 'descripcion': 'Frijoles Procesados', 'cod': 'RTEB'},
    {'categoria': 'Alimentos', 'descripcion': 'Pratos Prontos _ Comidas Listas', 'cod': 'RTEM'},
    {'categoria': 'Alimentos', 'descripcion': 'Aderezos para Ensalada', 'cod': 'SDRE'},
    {'categoria': 'Alimentos', 'descripcion': 'Sal', 'cod': 'SALT'},
    {'categoria': 'Alimentos', 'descripcion': 'Galletas Saladas_Galletas No Dulce', 'cod': 'SLTC'},
    {'categoria': 'Alimentos', 'descripcion': 'Sardina Envasada', 'cod': 'SARD'},
    {'categoria': 'Alimentos', 'descripcion': 'Cecinas', 'cod': 'SAUS'},
    {'categoria': 'Alimentos', 'descripcion': 'Milanesas', 'cod': 'SCHN'},
    {'categoria': 'Alimentos', 'descripcion': 'Snacks', 'cod': 'SNAC'},
    {'categoria': 'Alimentos', 'descripcion': 'Fideos Sopa', 'cod': 'SNOO'},
    {'categoria': 'Alimentos', 'descripcion': 'Sopas_Sopas Cremas', 'cod': 'SOUP'},
    {'categoria': 'Alimentos', 'descripcion': 'Siyau', 'cod': 'SOYS'},
    {'categoria': 'Alimentos', 'descripcion': 'Tallarines_Spaguetti', 'cod': 'SPAG'},
    {'categoria': 'Alimentos', 'descripcion': 'Chocolate para Untar', 'cod': 'SPCH'},
    {'categoria': 'Alimentos', 'descripcion': 'Azucar', 'cod': 'SUGA'},
    {'categoria': 'Alimentos', 'descripcion': 'Galletas Dulces', 'cod': 'SWCO'},
    {'categoria': 'Alimentos', 'descripcion': 'Untables Dulces', 'cod': 'SWSP'},
    {'categoria': 'Alimentos', 'descripcion': 'Endulzantes', 'cod': 'SWEE'},
    {'categoria': 'Alimentos', 'descripcion': 'Torradas _ Tostadas', 'cod': 'TOAS'},
    {'categoria': 'Alimentos', 'descripcion': 'Salsas de Tomate', 'cod': 'TOMA'},
    {'categoria': 'Alimentos', 'descripcion': 'Atún Envasado', 'cod': 'TUNA'},
    {'categoria': 'Alimentos', 'descripcion': 'Leche Vegetal', 'cod': 'VMLK'},
    {'categoria': 'Alimentos', 'descripcion': 'Harinas de trigo', 'cod': 'WFLO'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Ambientadores_Desodorante Ambiental', 'cod': 'AIRC'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Jabón en Barra_Jabón de lavar', 'cod': 'BARS'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Cloro_Lavandinas_Lejías_Blanqueadores', 'cod': 'BLEA'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Pastillas para Inodoro', 'cod': 'CBLK'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Guantes de látex', 'cod': 'CGLO'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Esponjas de Limpieza_Esponjas y paños', 'cod': 'CLSP'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Utensilios de Limpieza', 'cod': 'CLTO'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Filtros de Café', 'cod': 'FILT'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Cross Category (Limpiadores Domesticos)', 'cod': 'CRHC'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Cross Category (Lavandería)', 'cod': 'CRLA'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Cross Category (Productos de Papel)', 'cod': 'CRPA'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Lavavajillas_Lavaplatos _ Lavalozas mano', 'cod': 'DISH'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Empaques domésticos_Bolsas plásticas_Plástico Adherente_Papel encerado_Papel aluminio', 'cod': 'DPAC'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Destapacañerias', 'cod': 'DRUB'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Perfumantes para Ropa_Perfumes para Ropa', 'cod': 'FBRF'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Cera p_pisos', 'cod': 'FWAX'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Desodorante para Pies', 'cod': 'FDEO'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Lustramuebles', 'cod': 'FRNP'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Bolsas de Basura', 'cod': 'GBBG'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Limpiadores verdes', 'cod': 'GCLE'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Limpiadores_Limpiadores y Desinfectantes', 'cod': 'CLEA'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Insecticidas_Raticidas', 'cod': 'INSE'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Toallas de papel_Papel Toalla _ Toallas de Cocina _ Rollos Absorbentes de Papel', 'cod': 'KITT'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Detergentes para ropa', 'cod': 'LAUN'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Apresto', 'cod': 'LSTA'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Mixta Pastillas para Inodoro + Limpiadores', 'cod': 'MXBC'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Mixta Home Care_Cloro-Limpiadores-Ceras-Ambientadores', 'cod': 'MXHC'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Mixta Limpiadores + Cloro', 'cod': 'MXCB'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Mixta Detergentes + Cloro', 'cod': 'MXLB'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Mixta Detergentes + Lavavajillas', 'cod': 'MXLD'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Pañitos + Papel Higienico', 'cod': 'CRTO'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Servilletas', 'cod': 'NAPK'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Film plastico e papel aluminio', 'cod': 'PLWF'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Esponjas de Acero', 'cod': 'SCOU'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Suavizantes de Ropa', 'cod': 'SOFT'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Quitamanchas_Desmanchadores', 'cod': 'STRM'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Papel Higiénico', 'cod': 'TOIP'},
    {'categoria': 'Cuidado del Hogar', 'descripcion': 'Paños de Limpieza', 'cod': 'WIPE'},
    {'categoria': 'OTC', 'descripcion': 'Analgésicos_Painkillers', 'cod': 'ANLG'},
    {'categoria': 'OTC', 'descripcion': 'Suplementos alimentares', 'cod': 'FSUP'},
    {'categoria': 'OTC', 'descripcion': 'Gastrointestinales_Efervescentes', 'cod': 'GMED'},
    {'categoria': 'OTC', 'descripcion': 'Vitaminas y Calcio', 'cod': 'VITA'},
    {'categoria': 'Otros', 'descripcion': 'Categoría Desconocida', 'cod': 'nan'},
    {'categoria': 'Otros', 'descripcion': 'Pilas_Baterías', 'cod': 'BATT'},
    {'categoria': 'Otros', 'descripcion': 'Combustible Gas', 'cod': 'CGAS'},
    {'categoria': 'Otros', 'descripcion': 'Panel Financiero de Hogares', 'cod': 'PFIN'},
    {'categoria': 'Otros', 'descripcion': 'Panel Financiero de Hogares', 'cod': 'PFIN'},
    {'categoria': 'Otros', 'descripcion': 'Cartuchos de Tintas', 'cod': 'INKC'},
    {'categoria': 'Otros', 'descripcion': 'Alimento para Mascota_Alim.p _ perro _ gato', 'cod': 'PETF'},
    {'categoria': 'Otros', 'descripcion': 'Telecomunicaciones _ Convergencia', 'cod': 'TELE'},
    {'categoria': 'Otros', 'descripcion': 'Tickets _ Till Rolls', 'cod': 'TILL'},
    {'categoria': 'Otros', 'descripcion': 'Tabaco _ Cigarrillos', 'cod': 'TOBA'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Incontinencia de Adultos', 'cod': 'ADIP'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Shampoo Infantil', 'cod': 'BSHM'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Maquinas de Afeitar', 'cod': 'RAZO'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cremas Corporales', 'cod': 'BDCR'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Paños Húmedos', 'cod': 'CWIP'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cremas para Peinar', 'cod': 'COMB'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Acondicionador_Bálsamo', 'cod': 'COND'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cross Category (Higiene)', 'cod': 'CRHY'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cross Category (Personal Care)', 'cod': 'CRPC'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Desodorantes', 'cod': 'DEOD'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Pañales_Pañales Desechables', 'cod': 'DIAP'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cremas Faciales', 'cod': 'FCCR'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Pañuelos Faciales', 'cod': 'FTIS'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Protección Femenina_Toallas Femeninas', 'cod': 'FEMI'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Fragancias', 'cod': 'FRAG'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cuidado del Cabello_Hair Care', 'cod': 'HAIR'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Tintes para el Cabello_Tintes _ Tintura _ Tintes y Coloración para el cabello', 'cod': 'HRCO'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Depilación', 'cod': 'HREM'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Alisadores para el Cabello', 'cod': 'HRST'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Fijadores para el Cabello_Modeladores_Gel_Fijadores para el cabello', 'cod': 'HSTY'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Tratamientos para el Cabello', 'cod': 'HRTR'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Óleo Calcáreo', 'cod': 'LINI'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Maquillaje_Cosméticos', 'cod': 'MAKE'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Jabón Medicinal', 'cod': 'MEDS'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Mixta Make Up+Tinturas', 'cod': 'MXMH'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Enjuague Bucal_Refrescante Bucal', 'cod': 'MOWA'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cuidado Bucal', 'cod': 'ORAL'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Protectores Femeninos', 'cod': 'SPAD'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Toallas Femininas', 'cod': 'STOW'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Shampoo', 'cod': 'SHAM'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Afeitado_Crema afeitar_Loción de afeitar_Pord. Antes del afeitado', 'cod': 'SHAV'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cremas Faciales y Corporales_Cremas de Belleza _ Cremas Cuerp y Faciales', 'cod': 'SKCR'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Protección Solar', 'cod': 'SUNP'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Talcos_Talco para pies', 'cod': 'TALC'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Tampones Femeninos', 'cod': 'TAMP'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Jabón de Tocador', 'cod': 'TOIL'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Cepillos Dentales', 'cod': 'TOOB'},
    {'categoria': 'Cuidado Personal', 'descripcion': 'Pastas Dentales', 'cod': 'TOOT'},
    {'categoria': 'Material Escolar', 'descripcion': 'Morrales y MAletas Escoalres', 'cod': 'BAGS'},
    {'categoria': 'Material Escolar', 'descripcion': 'Lapices de Colores', 'cod': 'CLPC'},
    {'categoria': 'Material Escolar', 'descripcion': 'Lapices De Grafito', 'cod': 'GRPC'},
    {'categoria': 'Material Escolar', 'descripcion': 'Unidaddores', 'cod': 'MRKR'},
    {'categoria': 'Material Escolar', 'descripcion': 'Cuadernos', 'cod': 'NTBK'},
    {'categoria': 'Material Escolar', 'descripcion': 'Útiles Escolares', 'cod': 'SCHS'},
    {'categoria': 'Diversos', 'descripcion': 'Estudio de Categorías', 'cod': 'CSTD'},
    {'categoria': 'Diversos', 'descripcion': 'Corporativa', 'cod': 'CORP'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category', 'cod': 'CROS'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Bebés)', 'cod': 'CRBA'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Desayuno)_Yogurt, Cereal, Pan y Queso', 'cod': 'CRBR'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Diet y Light)', 'cod': 'CRDT'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Alimentos Secos)', 'cod': 'CRDF'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Alimentos)', 'cod': 'CRFO'},
    {'categoria': 'Diversos', 'descripcion': 'Cross Category (Salsas)_Mayonesas-Ketchup _ Salsas Frías', 'cod': 'CRSA'},
    {'categoria': 'Diversos', 'descripcion': 'Demo', 'cod': 'DEMO'},
    {'categoria': 'Diversos', 'descripcion': 'Flash', 'cod': 'FLSH'},
    {'categoria': 'Diversos', 'descripcion': 'Holistic View', 'cod': 'HLVW'},
    {'categoria': 'Diversos', 'descripcion': 'Mezcla para café instantaneo y crema no láctea', 'cod': 'COCP'},
    {'categoria': 'Diversos', 'descripcion': 'Mezclas nutricionales y suplementos', 'cod': 'CRSN'},
    {'categoria': 'Diversos', 'descripcion': 'Consolidado_Multicategory', 'cod': 'MULT'},
    {'categoria': 'Diversos', 'descripcion': 'Pantry Check', 'cod': 'PCHK'},
    {'categoria': 'Diversos', 'descripcion': 'Inventario', 'cod': 'STCK'},
    {'categoria': 'Diversos', 'descripcion': 'Leche y Cereales Calientes_Cereales Precocidos y Leche Líquida Blanca', 'cod': 'MIHC'}
]


def buscar_categorias(keyword):
    keyword_normalizado = strip_accents(keyword)
    # Filtra la lista buscando en 'descripcion' y en 'cod'
    return [cat for cat in categorias 
            if keyword_normalizado in strip_accents(cat['descripcion']) 
            or keyword_normalizado in strip_accents(cat['cod'])]

def seleccionar_categoria():
    while True:
        keyword = input(f"{Colors.OKBLUE}Ingrese palabra clave para la categoría: {Colors.ENDC}").strip()
        if not keyword:
            print(f"{Colors.FAIL}La palabra clave no puede estar vacía.{Colors.ENDC}")
            continue
        matches = buscar_categorias(keyword)
        if not matches:
            print(f"{Colors.FAIL}No se encontraron categorías. Intente de nuevo.{Colors.ENDC}")
            continue
        
        print(f"{Colors.OKCYAN}Categorías encontradas:{Colors.ENDC}")
        for idx, cat in enumerate(matches, 1):
            print(f"{Colors.OKGREEN}{idx}. {cat['descripcion']} ({cat['cod']}){Colors.ENDC}")
        seleccion = input(f"{Colors.OKBLUE}Número de categoría (o 'r' para reintentar): {Colors.ENDC}").strip()
        if seleccion.lower() == 'r':
            continue
        
        try:
            seleccion = int(seleccion)
            if 1 <= seleccion <= len(matches):
                return matches[seleccion - 1]
            else:
                print(f"{Colors.FAIL}Selección inválida.{Colors.ENDC}")
        except ValueError:
            print(f"{Colors.FAIL}Entrada inválida. Use números o 'r'.{Colors.ENDC}")

def sanitizar_nombre_hoja(nombre):
    # Reemplaza caracteres conflictivos en nombre de hoja de Excel
    for char in ['\\', '/', '*', '[', ']', ':', '?']:
        nombre = nombre.replace(char, '_')
    return nombre

def generar_nombre_unico(nombre_archivo):
    if not os.path.exists(nombre_archivo):
        return nombre_archivo
    base, ext = os.path.splitext(nombre_archivo)
    contador = 1
    while True:
        nuevo_nombre = f"{base} ({contador}){ext}"
        if not os.path.exists(nuevo_nombre):
            return nuevo_nombre
        contador += 1

def crear_excel(nombre_archivo, nombre_fabricante, nombre_pais, categoria_seleccionada):
    # IMPORTS DIFERIDOS
    import pandas as pd
    import numpy as np

    # Agregar el prefijo "P0_" al nombre de la hoja
    nombre_hoja = sanitizar_nombre_hoja("P0_" + nombre_fabricante)
    nombre_archivo_unico = generar_nombre_unico(nombre_archivo)
    
    headers = ["Unidad", "Weighted R_VOL1", "Weighted PENET", "Weighted VO1_BUY", 
               "Weighted VO1_DAY", "Weighted FREQ", "BUYERS", "Fabricante"]
    
    data = {
        "Unidad":            ["Linea PowerView", "Base"],
        "Weighted R_VOL1":  [np.nan, "Weighted R_VOL1"],
        "Weighted PENET":   [np.nan, "Weighted PENET"],
        "Weighted VO1_BUY": [np.nan, "Weighted VO1_BUY"],
        "Weighted VO1_DAY": [np.nan, "Weighted VO1_DAY"],
        "Weighted FREQ":    [np.nan, "Weighted FREQ"],
        "BUYERS":           [np.nan, "BUYERS"],
        "Fabricante":       [np.nan, "SELL-IN: " + nombre_fabricante]
    }
    df = pd.DataFrame(data, columns=headers)

    # Escribe el Excel
    with pd.ExcelWriter(nombre_archivo_unico, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False, header=False)
    
    if nombre_archivo_unico == nombre_archivo:
        msg = f"Archivo de Excel '{nombre_archivo_unico}' creado exitosamente."
    else:
        msg = (f"El archivo '{nombre_archivo}' ya existía.\n"
               f"Se creó un nuevo archivo '{nombre_archivo_unico}'.")
    print(f"{Colors.OKGREEN}{msg}{Colors.ENDC}")
    
    # Retorna el nombre real del archivo creado
    return nombre_archivo_unico

def limpiar_pantalla():
    os.system('cls' if os.name == 'nt' else 'clear')

def mostrar_encabezado(contador, lista_archivos):
    print(f"{Colors.HEADER}{Colors.BOLD}=== Automatizador de Archivos ==={Colors.ENDC}\n")
    print(f"{Colors.OKCYAN}Archivos creados: {contador}{Colors.ENDC}")
    if lista_archivos:
        print(f"{Colors.OKCYAN}Lista de archivos creados:{Colors.ENDC}")
        for idx, archivo in enumerate(lista_archivos, 1):
            print(f"  {idx}. {archivo}")
    print("\n" + "-"*50 + "\n")

def main():
    contador = 0
    lista_archivos = []

    try:
        while True:
            limpiar_pantalla()
            mostrar_encabezado(contador, lista_archivos)
            print(f"{Colors.OKGREEN}Nota: Puede salir del programa en cualquier momento presionando Ctrl+C.{Colors.ENDC}\n")
            
            # 1. Seleccionar país
            while True:
                input_pais = input(f"{Colors.OKBLUE}Ingrese nombre/abreviación del país: {Colors.ENDC}").strip()
                if not input_pais:
                    print(f"{Colors.FAIL}No puede estar vacío.{Colors.ENDC}")
                    continue
                codigo_pais, nombre_pais = obtener_codigo_pais(input_pais)
                if codigo_pais:
                    print(f"{Colors.OKGREEN}País: {nombre_pais} (Código: {codigo_pais}){Colors.ENDC}\n")
                    break
                else:
                    print(f"{Colors.FAIL}País no encontrado. Intente nuevamente.{Colors.ENDC}\n")
            
            # 2. Seleccionar categoría
            cat_sel = seleccionar_categoria()
            print(f"{Colors.OKGREEN}Categoría: {cat_sel['descripcion']} (Código: {cat_sel['cod']}){Colors.ENDC}\n")
            
            # 3. Nombre del fabricante
            while True:
                fabricante = input(f"{Colors.OKBLUE}Ingrese nombre del fabricante: {Colors.ENDC}").strip()
                if fabricante:
                    break
                else:
                    print(f"{Colors.FAIL}El fabricante no puede estar vacío.{Colors.ENDC}\n")
            
            # 4. Generar nombre y crear Excel
            nombre_archivo = f"{codigo_pais}_{cat_sel['cod']}_{fabricante}.xlsx"
            print(f"{Colors.OKGREEN}Nombre de archivo: {nombre_archivo}{Colors.ENDC}")
            nombre_archivo_creado = crear_excel(nombre_archivo, fabricante, nombre_pais, cat_sel)
            
            # Actualizar contador y lista de archivos
            contador += 1
            lista_archivos.append(nombre_archivo_creado)
            
            # Esperar a que el usuario esté listo para continuar
            input(f"\n{Colors.OKBLUE}Presione Enter para crear otro archivo o Ctrl+C para salir...{Colors.ENDC}\n")
    
    except KeyboardInterrupt:
        limpiar_pantalla()
        mostrar_encabezado(contador, lista_archivos)
        print(f"{Colors.OKCYAN}Programa finalizado por el usuario. ¡Hasta luego!{Colors.ENDC}\n")
        sys.exit()

if __name__ == "__main__":
    main()
