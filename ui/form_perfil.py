# ui/form_perfil.py
"""
Formulario modal para crear o editar un perfil de automatización.
Permite configurar: nombre, carpeta origen, patrón de archivo, hoja,
rangos, destino, horario y notificación.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re

from config import (BG_MAIN, BG_CARD, BG_SIDEBAR, ACCENT, ACCENT_HOVER,
                    SUCCESS, ERROR, TEXT_DARK, TEXT_LIGHT, TEXT_MUTED, BORDER,
                    FORMATOS_DESTINO, F)
from components import (HoverButton, Tooltip, AyudaInline, crear_entry,
                        recordar_carpeta, ultima_carpeta, explicar_error)
from core.config_store import ConfigStore
from core.excel_reader  import ExcelReader
from core.scheduler     import registrar_tarea, eliminar_tarea


class FormPerfil(tk.Toplevel):
    def __init__(self, parent, store: ConfigStore, perfil_id: str):
        super().__init__(parent)
        self.title("Configurar automatización")
        self.geometry("680x700")
        self.configure(bg=BG_MAIN)
        self.resizable(False, True)
        self.grab_set()

        self.store     = store
        self.perfil_id = perfil_id
        self.perfil    = store.obtener_perfil(perfil_id)
        self.reader    = ExcelReader()

        self._vars = {}
        self._entries = {}
        # Guarda, por campo, el texto de ejemplo ("placeholder") que se
        # muestra en gris cuando el campo está vacío. Se usa para poder
        # distinguir "el usuario escribió esto" de "quedó el texto de
        # ejemplo sin tocar" — sin este registro, un perfil nuevo guardaría
        # literalmente el texto de ejemplo como si fuera un dato real.
        self._placeholders = {}
        self._build()
        self._cargar_datos()

        # Snapshot del estado inicial, para poder avisar si el usuario
        # intenta cerrar el formulario con cambios sin guardar.
        self._snapshot_inicial = self._snapshot()
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _snapshot(self):
        return {k: v.get() for k, v in self._vars.items()}

    def _hay_cambios_sin_guardar(self):
        return self._snapshot() != self._snapshot_inicial

    def _valor(self, key):
        """
        Devuelve el valor real de un campo: si lo que está escrito es
        exactamente el texto de ejemplo (placeholder) que nunca se tocó,
        lo trata como vacío en vez de guardarlo como dato real.
        """
        v = self._vars[key].get().strip()
        if self._placeholders.get(key) == v:
            return ""
        return v

    # ── Construcción ──────────────────────────────────────────────────────────

    def _build(self):
        # Header
        tk.Frame(self, bg=BG_SIDEBAR, height=48).pack(fill=tk.X)
        tk.Label(self, text="Configurar automatización",
                 font=F(13, "bold"),
                 bg=BG_SIDEBAR, fg=TEXT_LIGHT).place(x=16, y=12)

        # Scroll container
        container = tk.Frame(self, bg=BG_MAIN)
        container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container, bg=BG_MAIN, highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=BG_MAIN)
        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        pad = {"padx": 20, "pady": 6}

        tk.Label(
            self.scroll_frame,
            text="Esta automatización revisará periódicamente una carpeta, "
                 "extraerá datos del Excel más reciente que encuentre y los "
                 "exportará sola, sin que tengas que abrir la app. Completa "
                 "las secciones de abajo — pasa el mouse sobre cada campo "
                 "si tienes dudas de qué escribir.",
            font=F(9, "italic"), bg=BG_MAIN, fg=TEXT_MUTED,
            wraplength=600, justify="left"
        ).pack(anchor="w", padx=20, pady=(14, 0))

        # ── Sección: Identificación ────────────────────────────────────────
        self._seccion("Identificación")
        self._campo("nombre",  "Nombre del perfil",  "Ej: Cierre TGM Coquimbo",
                    ayuda="Un nombre para identificar esta automatización en "
                          "la lista (ej: 'Cierre TGM Coquimbo').", **pad)

        # ── Sección: Origen ────────────────────────────────────────────────
        self._seccion(
            "Archivo de origen",
            ayuda="De dónde se leen los datos: la carpeta donde se guardan "
                  "las planillas, cómo se llaman y qué hoja/rango usar."
        )
        self._campo_carpeta(
            "carpeta_origen", "Carpeta de planillas", **pad,
            ayuda="Carpeta donde se guarda el Excel que quieres procesar."
        )
        self._campo(
            "patron_archivo", "Patrón de nombre",
            "Ej: cierre_tgm_{YYYYMM}.xlsx  |  tokens: {YYYYMM} {YYYY} {MM}",
            ayuda="Cómo se llama el archivo dentro de esa carpeta. Usa "
                  "{YYYYMM}, {YYYY} o {MM} donde el nombre cambie cada mes "
                  "(ej: cierre_tgm_{YYYYMM}.xlsx encuentra "
                  "cierre_tgm_202607.xlsx en julio 2026).",
            **pad
        )
        self._campo_hoja(
            **pad,
            ayuda="Hoja del Excel que contiene los datos a extraer. Usa "
                  "'Detectar' para leerlas desde un archivo real."
        )
        self._campo(
            "rangos", "Rangos de celdas",
            "Ej: B3:H30,J3:J30  — vacío = columnas completas",
            ayuda="Opcional. Si quieres extraer solo un área exacta de la "
                  "hoja, escribe uno o más rangos separados por coma "
                  "(notación Excel, ej: B3:H30). Déjalo vacío para tomar "
                  "todas las columnas de la hoja.",
            **pad
        )

        # ── Sección: Destino ───────────────────────────────────────────────
        self._seccion(
            "Destino",
            ayuda="Dónde y en qué formato se guarda el resultado cada vez "
                  "que corre la automatización."
        )
        self._campo_carpeta(
            "carpeta_destino", "Carpeta de exportación", **pad,
            ayuda="Carpeta donde se guardará el archivo con los datos "
                  "extraídos."
        )
        self._campo(
            "nombre_archivo", "Nombre del archivo de salida",
            "Sin extensión — se agrega automáticamente",
            ayuda="Nombre del archivo final, sin extensión (ej: "
                  "'reporte_cierre'). La extensión (.csv, .xlsx, .db) se "
                  "agrega sola según el formato elegido abajo.",
            **pad
        )
        self._campo_combo(
            "tipo_destino",  "Formato", FORMATOS_DESTINO, **pad,
            ayuda="Formato del archivo de salida: csv (texto plano), "
                  "xlsx (Excel) o sqlite (base de datos)."
        )
        self._campo_combo(
            "modo_destino",  "Modo",
            ["append (acumula)", "replace (sobreescribe)"], **pad,
            ayuda="'append' agrega los datos nuevos a los ya exportados "
                  "antes; 'replace' reemplaza el archivo completo cada vez."
        )

        # ── Sección: Horario ───────────────────────────────────────────────
        self._seccion(
            "Horario de ejecución",
            ayuda="Cuándo debe correr esta automatización sola, cada día."
        )
        self._campo(
            "hora_ejecucion", "Hora diaria (HH:MM)", "07:00",
            ayuda="Hora del día en que se ejecutará automáticamente "
                  "(formato 24 horas, ej: 07:00).",
            **pad
        )
        self._campo_check(
            "activo", "Activar automatización diaria", **pad,
            ayuda="Si está marcada, se programa una tarea del sistema "
                  "para correr esto sola todos los días a la hora indicada."
        )

        # ── Sección: Notificación ──────────────────────────────────────────
        self._seccion(
            "Notificación por email",
            ayuda="Opcional: recibe un correo automático avisando si la "
                  "ejecución salió bien o falló."
        )
        self._campo("email_destino",  "Email destinatario",   "usuario@empresa.cl",
                    ayuda="A qué correo se envía el aviso de resultado.", **pad)
        self._campo("smtp_host",      "Servidor SMTP",        "smtp.gmail.com",
                    ayuda="Servidor de correo saliente de tu proveedor "
                          "(ej: smtp.gmail.com para Gmail).", **pad)
        self._campo("smtp_port",      "Puerto SMTP",          "587",
                    ayuda="Puerto del servidor SMTP. 587 es el más común.", **pad)
        self._campo("smtp_user",      "Usuario SMTP",         "",
                    ayuda="Cuenta de correo que envía la notificación.", **pad)
        self._campo("smtp_pass",      "Contraseña SMTP",      "", show="*",
                    ayuda="Contraseña (o clave de aplicación) de esa cuenta "
                          "de correo. Se guarda solo en este equipo.", **pad)
        self._campo_check("notif_error",  "Notificar en error / archivo no encontrado",
                           ayuda="Envía un correo cuando la automatización "
                                 "falla o no encuentra el archivo esperado.",
                           **pad)
        self._campo_check("notif_exito",  "Notificar en éxito",
                           ayuda="Envía también un correo cuando todo sale "
                                 "bien, no solo cuando hay error.",
                           **pad)

        # ── Botones ────────────────────────────────────────────────────────
        btn_frame = tk.Frame(self, bg=BG_MAIN, pady=12)
        btn_frame.pack(fill=tk.X)

        btn_guardar = HoverButton(btn_frame, bg_normal=SUCCESS, bg_hover="#176138",
                    text="Guardar", font=F(11, "bold"),
                    fg=TEXT_LIGHT, padx=20, pady=8,
                    command=self._guardar)
        btn_guardar.pack(side=tk.LEFT, padx=(20, 8))
        Tooltip(btn_guardar,
                "Guarda esta configuración. Si 'Activar automatización "
                "diaria' está marcada, también programa la tarea diaria.")

        btn_cancelar = HoverButton(btn_frame, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="Cancelar", font=F(10),
                    fg=TEXT_DARK, padx=16, pady=8,
                    command=self._cancelar)
        btn_cancelar.pack(side=tk.LEFT)
        Tooltip(btn_cancelar, "Cierra sin guardar los cambios de este formulario.")

        btn_probar = HoverButton(btn_frame, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="▶ Probar ahora", font=F(10),
                    fg=ACCENT, padx=16, pady=8,
                    command=self._probar)
        btn_probar.pack(side=tk.RIGHT, padx=20)
        Tooltip(btn_probar,
                "Guarda y ejecuta esta automatización de inmediato, para "
                "comprobar que la carpeta, hoja y destino están bien "
                "configurados antes de dejarla corriendo sola.")

    # ── Widgets helper ────────────────────────────────────────────────────────

    def _seccion(self, titulo, ayuda=None):
        f = tk.Frame(self.scroll_frame, bg=ACCENT, height=2)
        f.pack(fill=tk.X, padx=20, pady=(14, 0))
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, padx=20, pady=(4, 0))
        tk.Label(row, text=titulo, font=F(10, "bold"),
                 bg=BG_MAIN, fg=ACCENT).pack(side=tk.LEFT, anchor="w")
        if ayuda:
            AyudaInline(row, ayuda, bg=BG_MAIN).pack(side=tk.LEFT, padx=(6, 0))

    def _campo(self, key, label, placeholder="", show="", ayuda=None, **kw):
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text=label, width=32, anchor="w",
                 font=F(9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar()
        entry = crear_entry(row, textvariable=var, font=F(10), width=36, show=show)
        entry.pack(side=tk.LEFT, padx=(8, 0))
        if ayuda:
            Tooltip(entry, ayuda)
        if placeholder:
            self._placeholders[key] = placeholder
            entry.insert(0, placeholder)
            entry.config(fg=TEXT_MUTED)
            def _on_focus_in(e, ph=placeholder, v=var, en=entry):
                if v.get() == ph:
                    en.delete(0, tk.END)
                    en.config(fg=TEXT_DARK)
            def _on_focus_out(e, ph=placeholder, v=var, en=entry):
                if not v.get():
                    en.insert(0, ph)
                    en.config(fg=TEXT_MUTED)
            entry.bind("<FocusIn>",  _on_focus_in)
            entry.bind("<FocusOut>", _on_focus_out)
        self._vars[key] = var
        self._entries[key] = entry

    def _campo_carpeta(self, key, label, ayuda=None, **kw):
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text=label, width=32, anchor="w",
                 font=F(9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar()
        entry = crear_entry(row, textvariable=var, font=F(10), width=28)
        entry.pack(side=tk.LEFT, padx=(8, 4))
        if ayuda:
            Tooltip(entry, ayuda)
        def _elegir_carpeta(v=var):
            elegida = filedialog.askdirectory(parent=self, initialdir=ultima_carpeta())
            if elegida:
                v.set(elegida)
                recordar_carpeta(elegida)

        btn = HoverButton(row, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="…", font=F(10), fg=ACCENT, padx=6, pady=2,
                    command=_elegir_carpeta)
        btn.pack(side=tk.LEFT)
        Tooltip(btn, "Abre el explorador de archivos para elegir la carpeta.")
        self._vars[key] = var

    def _campo_combo(self, key, label, opciones, ayuda=None, **kw):
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text=label, width=32, anchor="w",
                 font=F(9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar(value=opciones[0])
        combo = ttk.Combobox(row, textvariable=var, values=opciones,
                     state="readonly", width=34)
        combo.pack(side=tk.LEFT, padx=(8, 0))
        if ayuda:
            Tooltip(combo, ayuda)
        self._vars[key] = var

    def _campo_check(self, key, label, ayuda=None, **kw):
        var = tk.BooleanVar()
        chk = ttk.Checkbutton(self.scroll_frame, text=label, variable=var,
                               style="Main.TCheckbutton", cursor="hand2")
        chk.pack(anchor="w", **kw)
        if ayuda:
            Tooltip(chk, ayuda)
        self._vars[key] = var

    def _campo_hoja(self, ayuda=None, **kw):
        """Combo de hojas con botón para detectarlas automáticamente."""
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text="Hoja de trabajo", width=32, anchor="w",
                 font=F(9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar()
        self._combo_hoja = ttk.Combobox(row, textvariable=var,
                                         state="readonly", width=28)
        self._combo_hoja.pack(side=tk.LEFT, padx=(8, 4))
        if ayuda:
            Tooltip(self._combo_hoja, ayuda)
        btn_detectar = HoverButton(row, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="Detectar", font=F(9), fg=ACCENT, padx=8, pady=2,
                    command=self._detectar_hojas)
        btn_detectar.pack(side=tk.LEFT)
        Tooltip(btn_detectar,
                "Abre el archivo más reciente que calce con el patrón de "
                "nombre de arriba y lista sus hojas disponibles.")
        self._vars["hoja"] = var

    # ── Cargar / Guardar ──────────────────────────────────────────────────────

    def _cargar_datos(self):
        p = self.perfil

        def _sv(key, val):
            if not val:
                return
            self._vars[key].set(val)
            # El valor real cargado no debe verse como texto de ejemplo:
            # restaura el color normal (si el campo es de los que tienen
            # placeholder — carpeta_origen/destino no lo usan y no están
            # en este diccionario).
            entry = self._entries.get(key)
            if entry is not None:
                entry.config(fg=TEXT_DARK)

        _sv("nombre",         p.get("nombre", ""))
        _sv("carpeta_origen", p.get("carpeta_origen", ""))
        _sv("patron_archivo", p.get("patron_archivo", ""))
        _sv("hoja",           p.get("hoja", ""))
        _sv("rangos",         ", ".join(p.get("rangos", [])))

        dest = p.get("destino", {})
        _sv("carpeta_destino", dest.get("carpeta", ""))
        _sv("nombre_archivo",  dest.get("nombre_archivo", ""))
        self._vars["tipo_destino"].set(dest.get("tipo", "csv"))
        modo = dest.get("modo", "append")
        self._vars["modo_destino"].set(
            "append (acumula)" if modo == "append" else "replace (sobreescribe)"
        )

        sched = p.get("schedule", {})
        _sv("hora_ejecucion", sched.get("hora", "07:00"))
        self._vars["activo"].set(sched.get("activo", False))

        notif = p.get("notificacion", {})
        _sv("email_destino", notif.get("email_destino", ""))
        _sv("smtp_host",     notif.get("smtp_host", ""))
        _sv("smtp_port",     str(notif.get("smtp_port", 587)))
        _sv("smtp_user",     notif.get("smtp_user", ""))
        _sv("smtp_pass",     notif.get("smtp_pass", ""))
        self._vars["notif_error"].set(notif.get("en_error", True))
        self._vars["notif_exito"].set(notif.get("en_exito", False))

    def _detectar_hojas(self):
        carpeta = self._valor("carpeta_origen")
        patron  = self._valor("patron_archivo")
        if not carpeta or not patron:
            messagebox.showwarning("Faltan datos",
                                   "Ingresa la carpeta y el patrón antes de detectar hojas.",
                                   parent=self)
            return
        try:
            from datetime import datetime
            ruta = self.reader.resolver_archivo_mensual(carpeta, patron)
            self.reader.abrir(ruta)
            hojas = self.reader.obtener_hojas()
            self._combo_hoja.config(values=hojas)
            if hojas:
                self._combo_hoja.current(0)
        except Exception as e:
            messagebox.showerror("Error", explicar_error(e), parent=self)

    def _guardar(self, cerrar=True, avisar=True):
        """
        Valida y guarda el perfil.
        cerrar=False se usa desde 'Probar ahora': guarda sin cerrar la
        ventana, para poder mostrar el resultado de la prueba en un
        diálogo que sigue perteneciendo a una ventana que todavía existe
        (si se cerrara antes, el diálogo de resultado fallaría).
        Devuelve True si guardó, False si la validación falló.
        """
        g  = self._valor
        nombre = g("nombre")
        if not nombre:
            messagebox.showwarning("Campo requerido",
                                   "El nombre del perfil es obligatorio.", parent=self)
            return False

        hora = g("hora_ejecucion") or "07:00"
        if not re.fullmatch(r"([01]?\d|2[0-3]):([0-5]\d)", hora):
            messagebox.showwarning(
                "Hora inválida",
                f"'{hora}' no es una hora válida.\n"
                "Usa el formato HH:MM en 24 horas, por ejemplo 07:00 o 18:30.",
                parent=self)
            return False

        puerto_raw = g("smtp_port") or "587"
        if not puerto_raw.isdigit():
            messagebox.showwarning(
                "Puerto SMTP inválido",
                f"'{puerto_raw}' no es un número válido de puerto.\n"
                "Debe ser solo dígitos, por ejemplo 587.",
                parent=self)
            return False

        rangos_raw = g("rangos")
        rangos = [r.strip() for r in rangos_raw.split(",") if r.strip()] if rangos_raw else []

        modo_raw = self._vars["modo_destino"].get()
        modo = "append" if "append" in modo_raw else "replace"

        datos = {
            "nombre":          nombre,
            "carpeta_origen":  g("carpeta_origen"),
            "patron_archivo":  g("patron_archivo"),
            "hoja":            g("hoja"),
            "rangos":          rangos,
            "destino": {
                "tipo":           self._vars["tipo_destino"].get(),
                "carpeta":        g("carpeta_destino"),
                "nombre_archivo": g("nombre_archivo"),
                "modo":           modo,
            },
            "schedule": {
                "hora":   hora,
                "activo": self._vars["activo"].get(),
            },
            "notificacion": {
                "email_destino": g("email_destino"),
                "smtp_host":     g("smtp_host"),
                "smtp_port":     int(puerto_raw),
                "smtp_user":     g("smtp_user"),
                "smtp_pass":     g("smtp_pass"),
                "en_error":      self._vars["notif_error"].get(),
                "en_exito":      self._vars["notif_exito"].get(),
            },
        }

        self.store.actualizar_perfil(self.perfil_id, datos)

        # Registrar o eliminar tarea en Windows
        if datos["schedule"]["activo"]:
            try:
                registrar_tarea(nombre, self.perfil_id,
                                hora=datos["schedule"]["hora"])
            except Exception:
                pass  # En desarrollo o no-Windows, se ignora silenciosamente
        else:
            try:
                eliminar_tarea(nombre)
            except Exception:
                pass

        # El snapshot se actualiza para reflejar que ya no hay cambios
        # pendientes (si luego el usuario cierra sin tocar nada más, no
        # se le preguntará innecesariamente).
        self._snapshot_inicial = self._snapshot()

        if avisar:
            messagebox.showinfo("Guardado", f"Perfil '{nombre}' guardado correctamente.",
                                parent=self)
        if cerrar:
            self.destroy()
        return True

    def _cancelar(self):
        if self._hay_cambios_sin_guardar():
            if not messagebox.askyesno(
                "Cambios sin guardar",
                "Tienes cambios sin guardar en este formulario.\n"
                "¿Salir sin guardarlos?",
                parent=self
            ):
                return
        self.destroy()

    def _probar(self):
        # cerrar=False y avisar=False: guarda en silencio y mantiene la
        # ventana abierta, para poder mostrar el resultado de la prueba
        # aquí mismo y dejar que el usuario corrija algo si algo falla.
        if not self._guardar(cerrar=False, avisar=False):
            return
        from core.runner import ejecutar_perfil
        from core.logger import RunLogger
        perfil = self.store.obtener_perfil(self.perfil_id)
        resultado = ejecutar_perfil(perfil, logger=RunLogger())
        if resultado["estado"] == "ok":
            messagebox.showinfo("Prueba exitosa",
                                f"✔ {resultado['filas']} filas exportadas.\n{resultado['archivo']}",
                                parent=self)
        else:
            messagebox.showerror("Error en prueba",
                                 resultado["error"], parent=self)
