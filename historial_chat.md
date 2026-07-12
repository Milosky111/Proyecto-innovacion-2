# 📋 Historial de Desarrollo

Registro completo de la evolución del proyecto sesión por sesión.

---

## Sesión 1

> 📅 Jueves 18 de Junio de 2026 — 6:58 PM UTC

<details>
<summary><strong>Ver hitos de la sesión 1</strong></summary>

| # | Hito | Descripción |
|---|---|---|
| 1 | **Inicio del proyecto** | Script básico en `pandas` para extraer Nombre, Pérdida/Sobrante, Fecha y N° de caja desde la hoja `CIERRE TGM`. Se creó `requirements.txt`. |
| 2 | **Implementación de GUI** | Se integró `tkinter` para la interfaz gráfica y se recomendó `pyinstaller` para compilar a `.exe`. |
| 3 | **Errores de entorno** | Errores al compilar `pandas` con Python 3.14 (versión en desarrollo). Se orientó a usar Python estable o instalar herramientas de C++. |
| 4 | **Generalización** | Se rediseñó el programa para leer cualquier Excel, extraer sus hojas y mostrar columnas dinámicamente. |
| 5 | **Mejoras visuales (UI)** | Interfaz actualizada con `ttk` (Treeview + LabelFrames). Se agregó filtrado de columnas fantasma (`Unnamed`). |
| 6 | **Corrección de datos vacíos** | Se eliminaron espacios en blanco en nombres de columnas que generaban exportaciones vacías. Se forzó eliminar filas nulas. |
| 7 | **Modularización** | Código dividido en `logic.py` (lógica) y `extractor.py` (interfaz). Exportación temporal a `.txt` para auditar datos. |
| 8 | **Limpieza del archivo principal** | Se eliminaron líneas huérfanas de `pandas` en la interfaz que causaban errores tras la modularización. |
| 9 | **Selector de fila de títulos** | El Excel tenía encabezados decorativos. Se implementó un campo para que el usuario indicara en qué fila comenzaban los datos. |
| 10 | **Generador de reportes + retorno a Excel** | Función que detecta columnas de "Pérdidas" y "Número" para generar automáticamente `REPORTE FINAL`. |
| 11 | **Detección automática de encabezado** | Se eliminó el selector manual. El programa ahora lee 30 filas y elige la que tiene más celdas de texto. Excel de salida con fondo azul, letra blanca y auto-ajuste de columnas. |
| 12 | **Documentación** | Se actualizó `README.md` y se creó este archivo de historial. |

**Estado al cierre:** Proyecto modularizado con auto-detección, estilizado profesional y reportes integrados. La lógica de reporte aún tenía keywords hardcodeados (`PERDIDA`, `CAJA`, `NUMERO`).

</details>

---

## Sesión 2

> 📅 Jueves 18 de Junio de 2026 · Claude (Anthropic)

<details>
<summary><strong>Ver hitos de la sesión 2</strong></summary>

| # | Hito | Descripción |
|---|---|---|
| 13 | **Diagnóstico** | Se identificó el problema principal: `exportar_a_excel` en `logic.py` buscaba keywords hardcodeados (`PERDIDA`, `SOBRANTE`, `NUMERO`, `CAJA`), haciendo la herramienta dependiente de esa nomenclatura. |
| 14 | **Decisión de diseño** | Se confirmó el objetivo: herramienta 100% genérica, sin lógica especializada, que funcione con cualquier Excel. |
| 15 | **Eliminación de hardcoding** | Se reescribió `exportar_a_excel`. El `REPORTE FINAL` ahora concatena genéricamente todas las columnas elegidas: `Columna: valor \| Columna: valor`. |
| 16 | **Mejoras de estilo** | Filas alternas con azul suave (`EBF3FB`). Ancho máximo de columna limitado a 60 caracteres. |
| 17 | **Limpieza de `extractor.py`** | Se eliminaron imports y referencias huérfanas a `pandas` en la capa de vista. |
| 18 | **Documentación** | `README.md` reescrito con estructura más clara. Sesión 2 añadida al historial. |

**Estado al cierre:** Herramienta completamente genérica. Código limpio, modularizado y documentado.

</details>

---

## Sesión 3

> 📅 Jueves 18 de Junio de 2026 · Claude (Anthropic)

### 🔬 Diagnóstico de archivos reales

Se subieron dos Excel con datos modificados para proteger información sensible (`cajas.xlsx`, `cierre_1.xlsx`). Se inspeccionaron directamente con `openpyxl` y se encontró lo siguiente:

> **`cajas.xlsx`** — La fila 0 es el encabezado correcto, pero sus valores son números de negocio (`8620000`, `2860000`). El algoritmo anterior solo contaba celdas con contenido sin distinguir texto de números, por lo que los aceptaba como encabezados válidos. El `REPORTE FINAL` resultaba ilegible: `8620000: 5494010`.
>
> **`cierre_1.xlsx`** — Estructura de dos bloques en paralelo. Misma situación: encabezados mezclados con valores numéricos.
>
> **Conclusión:** El problema no era la detección de la fila, sino que los encabezados del Excel son datos de negocio reales (números o códigos). La solución no es mejorar el algoritmo sino darle al usuario la herramienta para nombrarlos correctamente.

<details>
<summary><strong>Ver hitos de la sesión 3</strong></summary>

| # | Hito | Descripción |
|---|---|---|
| 19 | **Algoritmo de detección mejorado** | `_detectar_encabezado` ahora puntúa filas de forma diferenciada: texto con letras → peso 2, número puro → peso 0.3. Más robusto ante archivos con encabezados mixtos. |
| 20 | **Exclusión de `REPORTE FINAL` previo** | `obtener_columnas` ahora filtra explícitamente la columna `REPORTE FINAL` para evitar que exportaciones anteriores contaminen una nueva extracción del mismo archivo. |
| 21 | **`rename_map` en `exportar_a_excel`** | Nuevo parámetro `rename_map` (`dict {nombre_original: nombre_nuevo}`) para renombrar columnas durante la exportación. El `REPORTE FINAL` generado usa los nombres nuevos. |
| 22 | **Rediseño completo de la interfaz** | `extractor.py` reescrito desde cero con layout de dos paneles: sidebar azul marino con indicador de pasos + panel derecho con tarjetas. Clase `HoverButton` con efecto hover y cursor de manita. |
| 23 | **`RenameDialog`** | Nueva clase `tk.Toplevel` (modal) con un campo por cada columna seleccionada para asignarle un nombre legible. Si se deja en blanco, conserva el original. Solución directa al problema de encabezados numéricos. |
| 24 | **Botón "✏ Renombrar columnas…"** | Botón secundario en la sección de exportar. Flujo: seleccionar columnas → renombrar → confirmar → guardar → exportar. |
| 25 | **Botones de selección rápida** | Botones "Seleccionar todo" y "Limpiar" en la lista de columnas. Contador en tiempo real de columnas seleccionadas. |
| 26 | **Indicador de estado en sidebar** | Punto de color en la barra lateral: 🔵 procesando · 🟢 éxito · 🔴 error · 🟡 advertencia. Los pasos completados se marcan en verde. |
| 27 | **Excel de salida mejorado** | Bordes en todas las celdas con `openpyxl.styles.Border`, altura explícita de filas (30px encabezado, 20px datos), fuente `Segoe UI` y primera fila congelada (`freeze_panes`). `exportar_a_excel` ahora retorna el número de filas exportadas. |
| 28 | **Documentación** | `README.md` e `historial_chat.md` actualizados con formato GitHub (tablas, emojis, bloques `<details>`, callouts con `>`). |

**Estado al cierre:** Interfaz completamente rediseñada. Nueva funcionalidad de renombrado que resuelve el problema de encabezados numéricos o en código. Excel de salida con estilizado completo.

</details>

---

## Sesión 4 — Fase 2: ETL y Automatización

> 📅 Martes 23 de Junio de 2026

<details>
<summary><strong>Ver hitos de la sesión 4</strong></summary>

| # | Hito | Descripción |
|---|---|---|
| 29 | **Lógica ETL en `logic.py`** | Se añadieron métodos para extracción por rango (`re` + `pandas`), exportación a CSV y a base de datos SQLite nativa (`sqlite3`). |
| 30 | **UI moderna (CustomTkinter)** | Migración de `tkinter` a `customtkinter` para interfaz con modo oscuro y diseño estético. Rediseño de la lista de columnas a checkboxes. |
| 31 | **Sistema de automatización** | Creación de `runner.py` usando `schedule` para ejecución en segundo plano mediante archivos de configuración JSON. |
| 32 | **Efectos dinámicos** | Implementación de feedback visual en botones (cambio de color al exportar exitosamente) y alertas de error profesional (`PermissionError`). |
| 33 | **Auto-Pilot de encabezados** | Implementación de algoritmo de densidad (`dropna(thresh=4)`) en `logic.py` para deducir automáticamente la fila de encabezado real sin intervención manual. |
| 34 | **Compilación final** | Preparación de entorno para `pyinstaller` con `--onefile` y soporte para íconos `.ico`. |
| 35 | **Sistema de logs** | `runner.py` genera `historial_extracciones.log` para registro técnico de auditoría. |

**Estado al cierre:** ETL básico funcional (rango, CSV, SQLite), interfaz migrada a CustomTkinter, primera versión de automatización en segundo plano con `schedule` y log en texto plano.

</details>

---

## Sesión 5 — Rediseño visual y validación de seguridad

> 📅 Miércoles 24 de Junio de 2026 · Claude (Anthropic)

<details>
<summary><strong>Ver hitos de la sesión 5</strong></summary>

| # | Hito | Descripción |
|---|---|---|
| 36 | **Paleta Cyberpunk Corporativo** | `config.py` reescrito: fondo carbón (`#0a0c10`), acento cyan (`#00f2ff`), magenta para crítico (`#ff2ec4`), verde terminal para éxito (`#39ff8c`). Tipografía mono (`Consolas`) para datos, UI normal para texto general. |
| 37 | **`RenameDialog` migrado a CustomTkinter** | El modal de renombrar columnas, que seguía en `tkinter` puro, se reescribió como `CTkToplevel` para que la estética sea consistente en toda la app. |
| 38 | **Sidebar tipo log de sistema** | Indicador de pasos rediseñado como `[ ] 1. CARGAR_ARCHIVO` → `[✓]` al completarse, con punto de estado en la cabecera. |
| 39 | **Corrección de bugs heredados** | Se eliminó código duplicado en `ExtractorApp.__init__` y un bloque muerto en `exportar_datos` que abría un segundo diálogo "Guardar como" sin manejo de errores después de cada exportación exitosa. |
| 40 | **Validación de seguridad (`security.py`)** | Nuevo módulo que valida cada Excel **antes** de abrirlo: firma binaria real vs. extensión declarada, tamaño máximo, rechazo de macros (`.xlsm`), protección contra "zip bomb" (ratio de compresión y tamaño descomprimido), rutas internas sospechosas (path traversal), límite de hojas. |
| 41 | **Dependencias de seguridad** | `requirements.txt` actualizado: `openpyxl>=3.1.0` (parche de XXE histórico) y `defusedxml` (protección contra "billion laughs" en el XML interno del `.xlsx`, que `openpyxl` no cubre por defecto). |
| 42 | **Bug de persistencia en Automatizaciones** | Se detectó que el panel de automatizaciones solo mostraba tareas guardadas en la sesión actual: al reabrir la app, la lista aparecía vacía aunque `automatizaciones.json` tuviera tareas. Se corrigió para que siempre se lea desde el archivo en disco. |

**Estado al cierre:** Interfaz con identidad visual propia y consistente. Primera capa de seguridad activa contra archivos corruptos o maliciosos. Bug de persistencia de automatizaciones corregido.

</details>

---

## Sesión 6 — Alineación con flujo objetivo y automatización por patrón

> 📅 Miércoles 24 de Junio de 2026 · Claude (Anthropic)

### 🎯 Contexto

Se recibió un diagrama de flujo objetivo (de José) que detalla cómo debería funcionar la automatización: configuración inicial (abrir planilla → definir carpeta → seleccionar rangos con clic → guardar perfil) y ejecución diaria (trigger de Windows → resolver archivo por patrón → extraer → exportar → log). Esta sesión adapta el proyecto existente hacia ese flujo, sin descartar el trabajo previo.

<details>
<summary><strong>Ver hitos de la sesión 6</strong></summary>

| # | Hito | Descripción |
|---|---|---|
| 43 | **`RangeSelector` — selección visual de celdas** | Nuevo componente en `components.py`: grilla clicable (`tkinter.Canvas`) que muestra una vista previa real de la hoja. El usuario hace clic en la celda inicial y en la celda final para marcar el rango; reemplaza el campo de texto donde había que escribir `B5:H100` a mano. Genera el mismo formato de rango que ya consumía `extraer_por_rango`. |
| 44 | **`obtener_vista_previa()` en `logic.py`** | Nuevo método que lee una porción cruda de la hoja (sin pandas) para alimentar la grilla del selector visual. |
| 45 | **Automatización por patrón de archivo** | Las tareas guardadas ya no apuntan a una ruta fija. Ahora se guardan como `carpeta_origen` + `patron_archivo`, donde el patrón admite `{YYYYMM}`, `{YYYY}`, `{MM}`, `{DD}` y comodines `*`. Así la automatización encuentra sola el archivo correspondiente al mes/día de la ejecución. |
| 46 | **`resolver_patron.py`** | Nuevo módulo que expande los placeholders del patrón con la fecha actual y resuelve, vía `glob`, qué archivo real existe en la carpeta. Si hay varias coincidencias, usa la más reciente. |
| 47 | **`PerfilAutomatizacionDialog`** | Nuevo modal en `components.py` donde el usuario confirma el nombre de la tarea y el patrón de archivo antes de guardar, con una sugerencia automática derivada del nombre del archivo cargado. |
| 48 | **`runner.py` reescrito** | Se eliminó la dependencia de `schedule` y el loop interno: ahora el script se ejecuta una vez por invocación (pensado para ser disparado por el Programador de Tareas de Windows, no por sí mismo). Resuelve el archivo por patrón antes de procesar, y cada tarea se ejecuta de forma aislada (un error en una tarea no detiene las demás). |
| 49 | **Log estructurado** | El log de ejecuciones pasó de texto plano (`historial_extracciones.log`) a JSON estructurado (`historial_extracciones.json`), con `tarea`, `timestamp`, `estado`, `detalle` y `filas` por entrada — permite que la UI lo lea y muestre el resultado de la última ejecución de cada tarea. |
| 50 | **`registrar_tarea_windows.bat`** | Nuevo script que registra `runner.py` en el Programador de Tareas de Windows (`schtasks`) con un clic derecho → "Ejecutar como administrador". Verifica que Python esté en el `PATH` antes de registrar la tarea, y deja documentado que el resultado de cada ejecución debe revisarse en el log JSON (no hay consola visible cuando Task Scheduler dispara la tarea). |
| 51 | **Panel de Automatizaciones con resultado de ejecución** | El panel ahora cruza `automatizaciones.json` con `historial_extracciones.json` y muestra, junto a cada tarea, el resultado (✓/✗) y detalle de su última ejecución. Se agregó botón "↻ Actualizar" para refrescar sin reabrir la app. |
| 52 | **Corrección de bug post-edición** | Una edición previa había dejado el método `extraer_por_rango` sin su firma (`def extraer_por_rango(...)` faltante) tras insertar `obtener_vista_previa` antes de él. Se restauró; verificado con prueba de integración completa (guardar tarea → `runner.py` → exportación → log). |

**Estado al cierre:** Flujo de configuración alineado con el diagrama objetivo (selección visual de rango, perfiles con patrón de archivo). Automatización ejecutable vía Programador de Tareas de Windows, con log estructurado legible desde la UI. Pendiente para un sprint futuro: Servicio de Windows real (hoy es Task Scheduler, no un servicio en `services.msc`) y la capa de transformación ETL más avanzada.

</details>