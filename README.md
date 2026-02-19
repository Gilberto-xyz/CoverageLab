# CoverageLab

![Bienvenida](https://i.imgur.com/rZoALwD.jpeg)

Automatiza el flujo de coberturas desde archivos Excel de entrada hasta entregables finales de negocio.

`coverage_studio.py` genera tres salidas por corrida:
- `Template_...xlsx` con cálculos y estructura de trabajo.
- `...pptx` con la presentación final.
- `Banco_...xlsx` con el banco de coberturas.

## Contenido
- [Scripts principales](#scripts-principales)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Inicio rápido](#inicio-rápido)
- [Formato del archivo de entrada](#formato-del-archivo-de-entrada)
- [Modo automático (variables de entorno)](#modo-automático-variables-de-entorno)
- [Salidas generadas](#salidas-generadas)
- [Flujo general](#flujo-general)
- [Troubleshooting](#troubleshooting)
- [Licencia](#licencia)

## Scripts principales
- `coverage_studio.py`: motor principal de análisis y generación de entregables (Excel + PPT + banco).
- `archivos_studio.py`: asistente para crear archivos base de captura (`<codPais>_<codCategoria>_<fabricante>.xlsx`).
- `scorecards_studio.py`: exportador de scorecards en Excel a partir de archivos de cobertura.
- `Modelo_PPT.pptx`: plantilla obligatoria para construir la presentación.

## Requisitos
- Python 3.8 o superior.
- `pip` actualizado.
- Dependencias Python definidas en `requirements.txt`.

## Instalación

### Opción recomendada
```bash
pip install -r requirements.txt
```

### Opción manual
```bash
pip install pandas numpy matplotlib openpyxl tqdm colorama rich dataframe_image scipy python-pptx pillow
```

## Inicio rápido

### Flujo A (recomendado): crear base y luego procesar
1. Genera un archivo base:
```bash
python archivos_studio.py
```
2. Completa el Excel con tus datos.
3. Ejecuta el procesamiento:
```bash
python coverage_studio.py
```

### Flujo B: procesar un Excel existente
1. Coloca el `.xlsx` en la misma carpeta del proyecto.
2. Ejecuta:
```bash
python coverage_studio.py
```

### Flujo C: exportar scorecards
```bash
python scorecards_studio.py
```

## Formato del archivo de entrada

### Nombre del archivo
- Formato esperado: `<codPais>_<codCategoria>_<fabricante>.xlsx`
- Ejemplo: `52_CARB_Coca Cola.xlsx`
- El parser usa solo las primeras tres secciones separadas por `_`.
- Recomendación: evita usar `_` dentro del nombre del fabricante.

### Hojas
- Cada hoja se interpreta como una marca.
- Si una hoja empieza con `P0_` a `P6_`, ese prefijo se usa como pipeline.
- Si no tiene prefijo `P#_`, se procesa como pipeline `0`.

## Modo automático (variables de entorno)
Si defines `AUTO_FILE`, `coverage_studio.py` corre sin preguntas interactivas.

Variables principales:
- `AUTO_FILE`: archivo `.xlsx` a procesar.
- `AUTO_COV_TYPE`: `Absoluta`, `relativa` o `AUTO`.
- `AUTO_RAZON`: razón del análisis.
- `AUTO_EJE`: `simple` o `doble`.
- `AUTO_ENGLISH`: `1/0` (incluye etiquetas en inglés).
- `AUTO_ROUND_COV`: `1/0` (redondeo de cobertura).
- `AUTO_VAR_BOX_STYLE` o `AUTO_VAR_STYLE`: `classic` o `pretty`.
- `AUTO_COV_SLIDE` o `AUTO_COV_SLIDE_STYLE`: `classic` o `complemented`.
- `AUTO_EVO_SLIDE` o `AUTO_EVO_SLIDE_STYLE`: `classic` o `simple`.
- `AUTO_EXTEA` o `AUTO_EXTRA_MONTHS`: meses extra para summary, por ejemplo `8,11`.
- `AUTO_EXTEA_MODE` o `AUTO_EXTRA_MONTHS_MODE`: `recent` o `both`.
- `AUTO_INDEX` y `AUTO_TOTAL`: opcionales para corridas batch.

Ejemplo (PowerShell):
```powershell
$env:AUTO_FILE = "52_CARB_Coca Cola.xlsx"
$env:AUTO_COV_TYPE = "Absoluta"
$env:AUTO_RAZON = "Otras"
$env:AUTO_EJE = "simple"
$env:AUTO_ENGLISH = "0"
$env:AUTO_ROUND_COV = "1"
python coverage_studio.py
```

## Salidas generadas
Por cada archivo procesado se crea una carpeta:
- `<Pais>-<CategoriaCorta>-<Fabricante>-<Ref>_<CoverageLabel>`

Dentro se generan:
1. `Template_<Pais>-<CategoriaCorta>-<Fabricante>-<Ref>_<CoverageLabel>.xlsx`
2. `<Pais>-<CategoriaCorta>-<Fabricante>-<Ref>_<CoverageLabel>.pptx`
3. `Banco_<Fabricante>_<Categoria>_<Pais>_<Ref>_<CoverageLabel>.xlsx`

Notas:
- Se usa la carpeta temporal `tmp/` durante la ejecución y se elimina al final.
- El archivo `file_temp_coverage.xlsx` es auxiliar interno del proceso.

## Flujo general

```mermaid
flowchart TD
    A[Inicio] --> B{Ya tienes Excel de entrada}
    B -- No --> C[Ejecutar python archivos_studio.py]
    C --> D[Llenar datos en el Excel generado]
    B -- Si --> E[Usar Excel existente]
    D --> F[Ejecutar python coverage_studio.py]
    E --> F
    F --> G[Seleccionar archivo(s) xlsx]
    G --> H[Configurar opciones de cobertura]
    H --> I[Procesamiento por marca y pipeline]
    I --> J[Crear carpeta de salida]
    J --> K[Guardar Template xlsx]
    J --> L[Guardar reporte pptx]
    J --> M[Guardar banco de coberturas xlsx]
    M --> N[Fin]
```

## Troubleshooting
- Error de plantilla: valida que `Modelo_PPT.pptx` exista en la raíz del proyecto.
- No aparecen archivos Excel: verifica que terminen en `.xlsx`, no empiecen con `~$` y no estén bloqueados.
- Error de metadata en nombre: revisa el formato `<codPais>_<codCategoria>_<fabricante>.xlsx`.
- Error por dependencias faltantes: ejecuta `pip install -r requirements.txt`.

## Licencia
Uso interno del equipo.
