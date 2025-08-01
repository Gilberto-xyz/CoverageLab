# coverage_g_mod_v11.3_stable.py

![Bienvenida](welcome.png)
Hola Bruna, Saludos desde MEXICO.

## Descripción General

Este script es una herramienta avanzada para el procesamiento, análisis y visualización de datos de cobertura y penetración de marcas en mercados de consumo masivo, especialmente diseñada para equipos de inteligencia comercial, marketing y ventas. Automatiza la generación de reportes en Excel y presentaciones PowerPoint a partir de archivos de datos mensuales, permitiendo comparar la información de ventas (Sell-in) y consumo (Sell-out/Kantar) bajo diferentes pipelines y metodologías de cobertura.

---

## Alcances y Funcionalidades

- **Carga y preprocesamiento de datos**: Lee archivos Excel con múltiples hojas (cada una representando una marca), limpia y estandariza los datos, y extrae metadatos clave del nombre del archivo (país, categoría, fabricante).
- **Cálculo de indicadores clave**:
  - Cobertura absoluta y relativa (ajustada por cobertura poblacional del país).
  - Penetración y número de compradores (Buyers) en año móvil (MAT).
  - Variaciones interanuales (Y-1, Y-2) en ventas y cobertura (anual, semestral, trimestral).
  - Correlación entre Sell-in y Sell-out (Pearson, MAT).
  - Estabilidad de la cobertura (diferencia entre último valor y hace 12 meses).
- **Automatización de reportes**:
  - Genera un archivo Excel con fórmulas dinámicas, acumulados, coberturas escalonadas y resúmenes listos para análisis.
  - Crea una carpeta de salida organizada por país, categoría, fabricante y fecha de referencia.
- **Visualización avanzada**:
  - Presentación PowerPoint con gráficos de barras (Cobertura vs Penetración), líneas (Tendencia Sell-in/Sell-out), y evolución mensual con variaciones YOY.
  - Tablas resumen y banco de coberturas exportados como imágenes de alta calidad.
  - Personalización de idioma (ES/PT) y etiquetas según país.
- **Interactividad**:
  - Selección interactiva de archivo, tipo de cobertura, razón de análisis y configuración de gráficos.
  - Progreso visual con barra (rich/tqdm) y mensajes de advertencia/éxito en colores.

---

## Estructura del Script

1. **Configuración y dependencias**: Importa librerías clave (pandas, numpy, matplotlib, pptx, openpyxl, rich, etc.) y define constantes globales, colores y catálogos embebidos (países, categorías).
2. **Funciones utilitarias**: Incluye utilidades para limpieza, escalonamiento de datos, cálculo de variaciones, correlaciones y manejo de fechas.
3. **Procesamiento principal**:
   - Selección y validación del archivo Excel.
   - Extracción de metadatos del nombre del archivo.
   - Preprocesamiento de cada hoja/marca.
   - Generación de archivo Excel temporal con fórmulas y resúmenes.
   - Renombrado y organización de archivos de salida.
4. **Generación de PowerPoint**:
   - Para cada marca y pipeline, crea slides con gráficos y tablas.
   - Slide resumen con tabla consolidada y espacio para comentarios.
   - Banco de coberturas exportado a Excel.

---

## Cómo Funciona

1. **Ejecución**: Al correr el script, se listan los archivos Excel disponibles. El usuario selecciona el archivo y responde preguntas interactivas sobre el tipo de cobertura y razón del análisis.
2. **Procesamiento**: El script procesa cada hoja del archivo, calcula los indicadores y genera un Excel temporal con fórmulas listas para análisis y auditoría.
3. **Visualización**: Se generan gráficos y tablas, que se insertan automáticamente en una presentación PowerPoint basada en una plantilla.
4. **Salida**: Todos los archivos generados (Excel final, PPT, banco de coberturas) se guardan en una carpeta específica, nombrada con los metadatos clave.

---

## Personalización y Mejora

- **Fácil actualización**: El código está modularizado y documentado, facilitando la adición de nuevas métricas, gráficos o ajustes en la lógica de negocio.
- **Soporte para nuevos países/categorías**: Solo es necesario actualizar los catálogos embebidos.
- **Internacionalización**: Las etiquetas y textos pueden adaptarse fácilmente a otros idiomas.
- **Escalabilidad**: Permite procesar grandes volúmenes de datos y múltiples marcas en una sola ejecución.

---

## Guía de Ejecución Paso a Paso

### Paso 1: Generar Excel con archivos_studio.py
1. Ejecuta el archivo `archivos_studio.py`
```bash
python archivos_studio.py
```
2. Este script generará los archivos Excel base necesarios para el análisis
3. Los archivos Excel creados contendrán las plantillas y estructuras requeridas

### Paso 2: Llenar la Información en Excel
1. Abre los archivos Excel generados en el Paso 1
2. Completa toda la información requerida en las hojas correspondientes:
   - Datos de ventas (Sell-in)
   - Datos de consumo (Sell-out/Kantar)
   - Información de marcas y categorías
   - Datos poblacionales por país
3. Guarda los cambios en los archivos Excel

### Paso 3: Ejecutar Estudio de Cobertura con coverage_studio.py
1. Ejecuta el archivo `coverage_studio.py`
```bash
python coverage_studio.py
```
2. Selecciona el archivo Excel que llenaste en el Paso 2
3. Sigue las instrucciones interactivas para configurar:
   - Tipo de cobertura deseado
   - Razón del análisis
   - Configuración de gráficos
4. El script procesará los datos y generará:
   - Reporte Excel con análisis detallado
   - Presentación PowerPoint con visualizaciones
   - Banco de coberturas

## Requisitos

- Python 3.8+
- Bibliotecas: pandas, numpy, matplotlib, openpyxl, tqdm, colorama, rich, dataframe_image, scipy, python-pptx, pillow
- Plantilla PowerPoint: `Modelo_Revision.pptx` en el mismo directorio

Instalación de dependencias:
```bash
pip install pandas numpy matplotlib openpyxl tqdm colorama rich dataframe_image scipy python-pptx pillow
```

---

## Notas Técnicas

- El script es interactivo y requiere ejecución en terminal/IDE con soporte de entrada estándar.
- Los archivos temporales y de salida se gestionan automáticamente.
- El código maneja errores y advertencias para asegurar robustez y trazabilidad.

---

## Créditos y Contacto

Desarrollado por el equipo de inteligencia de coberturas. Para soporte, mejoras o reportar bugs, contactar a: [LatAmDQ.Coverage@kantar.com]

---

## Licencia

Uso interno. Adaptable bajo requerimiento del área de negocio.
