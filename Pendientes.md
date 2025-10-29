# Pendientes y Seguimiento

## Bugs, Comentarios y Sugerencias de Cambios
- [ ] Bug: Describe el bug aquí. 
- [ ] Mejora: Describe la mejora o nueva funcionalidad aquí.
- [ ] Cambio: Descripción del cambio a realizar

## Historial de Cambios

### Junio 2025  
- [x] Optimización del rendimiento, inicio y carga de módulos  
- [x] Inclusión de `.gitignore` para excluir archivos Excel y carpetas locales  
- [x] Mejoras en interfaz, procesamiento múltiple y visualización en terminal (con revert de cambios previos)  
- [x] Ajuste de posición de diapositiva de resumen  
- [x] Commit inicial con archivos base y README  

### Julio 2025  
- [x] Renombrado de archivos y detección inicial de *headers* erróneos (con bug pendiente)  
- [x] Mejora de nombres e inclusión del fabricante en plantilla con prefijo “SELL-IN”  
- [x] Solución de bug en banco de datos que impedía mostrar la marca  
- [x] Inclusión de marca del fabricante en títulos  
- [x] Ajustes y correcciones en variaciones (numerador y cliente), firma y mensaje de salida  
- [x] Corrección en cobertura y correlación en Excel  
- [x] Inclusión de metadatos al listar archivos Excel  
- [x] Nueva estructura “Worldpanel by numerator” y mejoras visuales en presentación de diapositivas  
- [x] Actualizaciones de identidad, categorías y paginación  

### Agosto 2025  
- [x] Implementación de versiones en portugués e inglés (versión unificada de idiomas y portada en español)  
- [x] Correlaciones ampliadas: cálculo de 2 últimos años móviles con ventanas de 12 y 24 meses  
- [x] Variaciones corregidas a 24 meses con descarte del 5wh1 en PPT  
- [x] Limpieza de la forma anterior de correlación  
- [x] Gráfico de variación mensual con ventana de 1 mes y ajustes visuales  
- [x] Resumen reordenado a Fabricante/Marca con imagen autoajustable al ancho del slide  
- [x] Ajuste visual en terminal y persistencia de opciones seleccionadas  
- [x] Complementación del formato multidioma  
- [x] Eliminación automática de la carpeta temporal de PPT  
- [x] Cambio de nombre en README  
- [x] Agregar colores en el template, asi como formato de salida en % con un decimal segun sea el caso
- [x] Mejora en el rendimiento en el menu. carga de modulos hasta el momento de procesar los archivos excel y generacion de graficos
- [x] Agregar la estabilidad en la tabla resumen

### Septiembre 2025  
- [x] Se Agrega fecha de ejecución en el archivo Excel para el Banco [Tomar el mes actual donde se esta ejecutando el script y Ajuste formato a fecha (día 1 del mes) para que en Excel se pueda formatear como mmm-yy]
- [x] Se modifico "periodo" en el archivo Excel para el Banco [Ajuste formato a fecha (día 1 del mes) para que en Excel se pueda formatear como mmm-yy] y sea mes-yy visualmente en Excel
- [x] Se agrego la opcion de redondear en el gráfico de cobertura y summary los datos a 0 digitos (Implementacion global en el menu del script / como la opcion de graficos en doble eje) 
- [x] Se acorto el nombre de la categoria solo para nombres de archivos y carpetas delimitado por el primer guion (Ejemplo: "Cuidado del Cabello - Shampoo y Acondicionador" a "Cuidado del Cabello") [Solo afecta nombres de archivos/rutas; títulos/textos continúan con la categoría completa.
Maneja dashes con o sin espacios alrededor]
- [x] Ampliacion de paises en el mapa de paises (COUNTRY_MAP) y ajuste en la construccion del DataFrame 'pais' para evitar errores si se modifica COUNTRY_MAP, Ajuste en la cobertura relativa en los paises "CAM" [Se muestra cobertura urbana en vez de Poblacional, debido a que los paises de CAM tienen muestras pequeñas y no representan la poblacion total]

### Octubre 2025 
- [x] Ampliacion de los paises y cobertura poblacional en COUNTRY_MAP (Colombia, Ecuador, Peru, Chile, Argentina, Uruguay, Paraguay, Bolivia) en "archivos_studio.py" y ajuste en la construccion del DataFrame 'pais' para evitar errores si se modifica COUNTRY_MAP
- [x] Se reemplazo la imagen de bienvenida del README por un recurso alojado en linea y eliminamos "welcome.png" para aligerar el repositorio
- [x] Habilitamos la personalizacion de estilos al crear imagenes de DataFrames y reutilizamos la tabla de variaciones en la slide de tendencias mensuales
- [x] Optimizamos el Modelo_PPT y ajustamos "PPT_LAYOUT_INDEX" al layout base para evitar desfases en los elementos
