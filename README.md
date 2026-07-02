# 📊 Extractor de Datos Excel

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python&logoColor=white" alt="Python Version">
  <img src="https://img.shields.io/badge/Framework-Tkinter-subtle?style=for-the-badge&logo=python&logoColor=white" alt="Tkinter">
  <img src="https://img.shields.io/badge/Data_Handling-Pandas-darkgreen?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas">
  <img src="https://img.shields.io/badge/Platform-Windows-0078D4?style=for-the-badge&logo=windows&logoColor=white" alt="Windows Support">
</p>

---

## 🚀 Uso Rápido (Para Usuarios Finales)

**No necesitas instalar Python ni ninguna librería para usar este programa.**

1. Ve a la carpeta `dist/` (o a la sección de *Releases* si estás en GitHub).
2. Descarga el archivo **`extractor.exe`**.
3. Haz doble clic en el archivo para abrir la aplicación. 
   > *Nota: Como lee configuraciones locales, asegúrate de que el ejecutable esté en la misma carpeta que los archivos `config_perfiles.json` y `extractor_log.db` (si ya existen).*
4. ¡Listo! Utiliza la interfaz para cargar tu Excel, seleccionar columnas/rangos y exportar.

---

## 📝 Descripción del Proyecto

**Extractor de Datos Excel** es una aplicación de escritorio robusta diseñada para optimizar y automatizar los flujos de trabajo con archivos analíticos de Excel. Permite realizar extracciones manuales precisas mediante una interfaz gráfica, así como programar tareas en segundo plano que extraen datos, generan reportes de auditoría y envían notificaciones por correo electrónico sin intervención humana.

---

## ✨ Características Principales

* **Extracción Manual Visual:** Selecciona columnas completas o usa el **selector visual de celdas por clic** para marcar rangos específicos directamente sobre una vista previa de tu hoja de cálculo.
* **Automatización Integrada:** Configura perfiles (rutas, rangos, horarios) que se integran nativamente con el **Programador de Tareas de Windows**, ejecutándose de forma invisible (*headless*).
* **Seguridad Avanzada:** Bloquea de inmediato archivos con macros, extensiones falsas o *zip bombs*, protegiendo el equipo antes de intentar procesar los datos.
* **Mapeo Dinámico de Nombres:** Renombra las columnas extraídas en tiempo real antes de exportarlas.
* **Múltiples Formatos de Salida:** Exporta los resultados limpios a Excel (`.xlsx`) estilizado, `.csv`, o bases de datos SQLite local (`.db`).
* **Auditoría y Alertas:** Genera PDFs gerenciales con el historial de extracciones y envía notificaciones automáticas por correo electrónico (SMTP) indicando el éxito o los errores de la tarea programada.

---

## 🧩 Arquitectura del Sistema

El proyecto está diseñado bajo una arquitectura modular limpia, separando estrictamente la interfaz gráfica (UI) de la lógica de negocio (Core).

### 🖥️ Módulos de Interfaz Gráfica (`/ui` y raíz)
Construida con `tkinter` y `ttk`, ofreciendo un diseño moderno y oscuro para los paneles de control.

* **`extractor.py` / `ui_views.py`**: El punto de entrada principal y el gestor de la maquetación (tarjetas, botones, layout general).
* **`click_range_selector.py`**: Un lienzo interactivo que renderiza una vista previa del Excel, permitiendo al usuario marcar el rango de inicio y fin haciendo clics, en lugar de escribir coordenadas.
* **`selector_rangos.py`**: Panel avanzado para previsualizar datos y renombrar las columnas del rango seleccionado.
* **`panel_automatizaciones.py`**: El *dashboard* principal de tareas programadas. Muestra un resumen del estado de cada perfil (Éxito, Error, Sin archivo) y permite forzar su ejecución.
* **`form_perfil.py`**: Formulario modal para configurar todos los parámetros de un perfil automatizado (origen, destino, programación, alertas SMTP).

### ⚙️ Motor Core y Procesamiento (`/core`)
Diseñado para funcionar tanto con la UI como en modo consola (*headless*) a través de tareas programadas.

* **`security.py`**: Módulo crítico. Inspecciona firmas binarias y previene ataques o bloqueos por archivos maliciosos (ej. *zip bombs* o macros ocultas).
* **`excel_reader.py`**: Motor de lectura híbrido. Usa `pandas` para columnas completas y `xlwings` para leer rangos crudos (preservando valores calculados y no las fórmulas subyacentes).
* **`runner.py` & `scheduler.py`**: 
  * `runner.py` orquesta la extracción automatizada (resolución de fecha, lectura, exportación, logging y notificación). 
  * `scheduler.py` inyecta estas tareas directamente en `schtasks.exe` (Windows).
* **`exporters.py`**: Formatea y exporta los *DataFrames*. Incluye estilización automática de celdas para archivos `.xlsx`.
* **`logger.py` & `audit_report.py`**: Mantienen un historial inmutable en `extractor_log.db` y utilizan `reportlab` para emitir reportes de auditoría en formato PDF.
* **`notifier.py`**: Cliente SMTP que compone correos HTML detallados, adjuntando los resultados de la extracción y/o el PDF de auditoría (controlando los límites de MB).
* **`config_store.py`**: Gestor de estado que persiste todos los perfiles en el archivo `config_perfiles.json`.

---

## 💻 Instalación para Desarrolladores

Si deseas modificar el código fuente o compilar tu propio ejecutable:

1. **Clonar el repositorio:**
   ```bash
   git clone [https://github.com/tu-usuario/extractor-datos-excel.git](https://github.com/tu-usuario/extractor-datos-excel.git)
   cd extractor-datos-excel
