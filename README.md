# CoverageLab

![Bienvenida](https://i.imgur.com/rZoALwD.jpeg)

`coverage_studio.py` automatiza el analisis de cobertura y genera 3 salidas por corrida:
- un Excel template con calculos (`Template_...xlsx`)
- un PowerPoint final (`...pptx`)
- un banco de coberturas (`Banco_...xlsx`)

## Scripts del proyecto
- `coverage_studio.py`: motor principal de analisis y generacion de entregables.
- `archivos_studio.py`: generador de archivos Excel base para captura manual.
- `scorecards_studio.py`: utilitario adicional del repositorio.
- `Modelo_PPT.pptx`: plantilla obligatoria para construir la presentacion.

## Requisitos
- Python 3.8+
- Dependencias:
  - `pandas`
  - `numpy`
  - `matplotlib`
  - `openpyxl`
  - `tqdm`
  - `colorama`
  - `rich`
  - `dataframe_image`
  - `scipy`
  - `python-pptx`
  - `pillow`

Instalacion:

```bash
pip install pandas numpy matplotlib openpyxl tqdm colorama rich dataframe_image scipy python-pptx pillow
```

## Inicio rapido

### Opcion A (recomendada): generar base primero
1. Ejecuta:

```bash
python archivos_studio.py
```

2. Completa el Excel generado con tus datos.
3. Ejecuta:

```bash
python coverage_studio.py
```

### Opcion B: usar un Excel ya existente
1. Coloca el archivo `.xlsx` en la misma carpeta del script.
2. Ejecuta:

```bash
python coverage_studio.py
```

## Formato esperado del archivo de entrada

### Nombre del archivo
- Formato: `<codPais>_<codCategoria>_<fabricante>.xlsx`
- Ejemplo: `52_CARB_Coca Cola.xlsx`

Nota importante:
- El parser usa solo las primeras 3 partes separadas por `_`.
- Para evitar truncamiento en fabricante, evita usar `_` dentro del nombre del fabricante.

### Hojas
- Cada hoja se interpreta como marca a procesar.
- Si una hoja comienza con `P0_` ... `P6_`, ese prefijo se usa como pipeline en graficos nativos de Excel.
- Si no tiene prefijo `P#_`, se procesa como pipeline `0`.

## Flujo interactivo de `coverage_studio.py`
Durante la ejecucion, el script solicita:
1. Archivo(s) Excel a procesar.
2. Tipo de cobertura:
   - `1` Absoluta
   - `2` Relativa
   - `3` AUTO (usa configuracion predeterminada)
3. Razon del analisis.
4. Tipo de eje para tendencia (simple o doble).
5. Estilo del cuadro de variaciones (clasico o bonito).
6. Modo del slide de cobertura (clasico o complementado).
7. Modo del slide de evolucion (clasico o simple).
8. Idioma ingles (si/no).
9. Redondeo de cobertura (si/no).
10. Meses extra para summary (opcional) y modo de comparacion.

## Variables de entorno (modo automatico)
Si defines `AUTO_FILE`, el script corre en modo automatico sin preguntas.

Variables principales:
- `AUTO_FILE`: archivo `.xlsx` a procesar.
- `AUTO_COV_TYPE`: `Absoluta`, `relativa` o `AUTO`.
- `AUTO_RAZON`: razon del analisis.
- `AUTO_EJE`: `simple` o `doble`.
- `AUTO_ENGLISH`: `1/0`.
- `AUTO_ROUND_COV`: `1/0`.
- `AUTO_VAR_BOX_STYLE` o `AUTO_VAR_STYLE`: `classic` o `pretty`.
- `AUTO_COV_SLIDE` o `AUTO_COV_SLIDE_STYLE`: `classic` o `complemented`.
- `AUTO_EVO_SLIDE` o `AUTO_EVO_SLIDE_STYLE`: `classic` o `simple`.
- `AUTO_EXTEA` o `AUTO_EXTRA_MONTHS`: meses extra para summary (ej. `8,11`).
- `AUTO_EXTEA_MODE` o `AUTO_EXTRA_MONTHS_MODE`: `recent` o `both`.

## Salidas generadas
Para cada archivo procesado se crea una carpeta:
- `<Pais>-<CategoriaCorta>-<Fabricante>-<Ref>_<CoverageLabel>`

Dentro de esa carpeta se guarda:
1. `Template_<Pais>-<CategoriaCorta>-<Fabricante>-<Ref>_<CoverageLabel>.xlsx`
2. `<Pais>-<CategoriaCorta>-<Fabricante>-<Ref>_<CoverageLabel>.pptx`
3. `Banco_<Fabricante>_<Categoria>_<Pais>_<Ref>_<CoverageLabel>.xlsx`

Adicional:
- Se usa una carpeta temporal `tmp/` durante el proceso y se elimina al final.

## Flujo general (Mermaid)

```mermaid
flowchart TD
    A[Inicio] --> B{Ya tienes Excel de entrada?}
    B -- No --> C[Ejecutar python archivos_studio.py]
    C --> D[Llenar datos en el Excel generado]
    B -- Si --> E[Usar Excel existente<br/>&lt;codPais&gt;_&lt;codCategoria&gt;_&lt;fabricante&gt;.xlsx]
    D --> F[Ejecutar python coverage_studio.py]
    E --> F
    F --> G[Seleccionar archivo(s) .xlsx]
    G --> H[Configurar opciones<br/>cobertura, razon, ejes, slides, idioma]
    H --> I[Procesamiento por marca y pipeline]
    I --> J[Genera carpeta de salida]
    J --> K[Guarda Template_...xlsx]
    J --> L[Guarda ...pptx]
    J --> M[Guarda Banco_...xlsx]
    M --> N[Fin]
```

## Troubleshooting rapido
- Error de plantilla: valida que `Modelo_PPT.pptx` exista en la misma carpeta.
- No aparecen archivos Excel: revisa que terminen en `.xlsx` y no esten abiertos/bloqueados.
- Error de metadata en nombre: revisa formato `<codPais>_<codCategoria>_<fabricante>.xlsx`.
- Inconsistencias de salida: verifica que las hojas tengan estructura de datos completa y fechas validas.

## Licencia
Uso interno del equipo.
