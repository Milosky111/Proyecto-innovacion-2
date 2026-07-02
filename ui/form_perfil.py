# ui/form_perfil.py
"""
Formulario modal para crear o editar un perfil de automatización.
Permite configurar: nombre, carpeta origen, patrón de archivo, hoja,
rangos, destino, horario y notificación.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from config import (BG_MAIN, BG_CARD, BG_SIDEBAR, ACCENT, ACCENT_HOVER,
                    SUCCESS, ERROR, TEXT_DARK, TEXT_LIGHT, TEXT_MUTED, BORDER,
                    FORMATOS_DESTINO)
from components import HoverButton
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
        self._build()
        self._cargar_datos()

    # ── Construcción ──────────────────────────────────────────────────────────

    def _build(self):
        # Header
        tk.Frame(self, bg=BG_SIDEBAR, height=48).pack(fill=tk.X)
        tk.Label(self, text="Configurar automatización",
                 font=("Segoe UI", 13, "bold"),
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

        # ── Sección: Identificación ────────────────────────────────────────
        self._seccion("Identificación")
        self._campo("nombre",  "Nombre del perfil",  "Ej: Cierre TGM Coquimbo", **pad)

        # ── Sección: Origen ────────────────────────────────────────────────
        self._seccion("Archivo de origen")
        self._campo_carpeta("carpeta_origen", "Carpeta de planillas", **pad)
        self._campo("patron_archivo", "Patrón de nombre",
                    "Ej: cierre_tgm_{YYYYMM}.xlsx  |  tokens: {YYYYMM} {YYYY} {MM}", **pad)
        self._campo_hoja(**pad)
        self._campo("rangos", "Rangos de celdas",
                    "Ej: B3:H30,J3:J30  — vacío = columnas completas", **pad)

        # ── Sección: Destino ───────────────────────────────────────────────
        self._seccion("Destino")
        self._campo_carpeta("carpeta_destino", "Carpeta de exportación", **pad)
        self._campo("nombre_archivo", "Nombre del archivo de salida",
                    "Sin extensión — se agrega automáticamente", **pad)
        self._campo_combo("tipo_destino",  "Formato",     FORMATOS_DESTINO, **pad)
        self._campo_combo("modo_destino",  "Modo",
                          ["append (acumula)", "replace (sobreescribe)"], **pad)

        # ── Sección: Horario ───────────────────────────────────────────────
        self._seccion("Horario de ejecución")
        self._campo("hora_ejecucion", "Hora diaria (HH:MM)", "07:00", **pad)
        self._campo_check("activo", "Activar automatización diaria", **pad)

        # ── Sección: Notificación ──────────────────────────────────────────
        self._seccion("Notificación por email")
        self._campo("email_destino",  "Email destinatario",   "usuario@empresa.cl", **pad)
        self._campo("smtp_host",      "Servidor SMTP",        "smtp.gmail.com", **pad)
        self._campo("smtp_port",      "Puerto SMTP",          "587", **pad)
        self._campo("smtp_user",      "Usuario SMTP",         "", **pad)
        self._campo("smtp_pass",      "Contraseña SMTP",      "", show="*", **pad)
        self._campo_check("notif_error",  "Notificar en error / archivo no encontrado", **pad)
        self._campo_check("notif_exito",  "Notificar en éxito", **pad)

        # ── Botones ────────────────────────────────────────────────────────
        btn_frame = tk.Frame(self, bg=BG_MAIN, pady=12)
        btn_frame.pack(fill=tk.X)

        HoverButton(btn_frame, bg_normal=SUCCESS, bg_hover="#176138",
                    text="Guardar", font=("Segoe UI", 11, "bold"),
                    fg=TEXT_LIGHT, padx=20, pady=8,
                    command=self._guardar).pack(side=tk.LEFT, padx=(20, 8))

        HoverButton(btn_frame, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="Cancelar", font=("Segoe UI", 10),
                    fg=TEXT_DARK, padx=16, pady=8,
                    command=self.destroy).pack(side=tk.LEFT)

        HoverButton(btn_frame, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="▶ Probar ahora", font=("Segoe UI", 10),
                    fg=ACCENT, padx=16, pady=8,
                    command=self._probar).pack(side=tk.RIGHT, padx=20)

    # ── Widgets helper ────────────────────────────────────────────────────────

    def _seccion(self, titulo):
        f = tk.Frame(self.scroll_frame, bg=ACCENT, height=2)
        f.pack(fill=tk.X, padx=20, pady=(14, 0))
        tk.Label(self.scroll_frame, text=titulo,
                 font=("Segoe UI", 10, "bold"),
                 bg=BG_MAIN, fg=ACCENT).pack(anchor="w", padx=20, pady=(4, 0))

    def _campo(self, key, label, placeholder="", show="", **kw):
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text=label, width=32, anchor="w",
                 font=("Segoe UI", 9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar()
        entry = tk.Entry(row, textvariable=var, font=("Segoe UI", 10),
                         width=36, relief="solid", bd=1, show=show,
                         highlightcolor=ACCENT, highlightthickness=1)
        entry.pack(side=tk.LEFT, padx=(8, 0))
        if placeholder:
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

    def _campo_carpeta(self, key, label, **kw):
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text=label, width=32, anchor="w",
                 font=("Segoe UI", 9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar()
        tk.Entry(row, textvariable=var, font=("Segoe UI", 10),
                 width=28, relief="solid", bd=1).pack(side=tk.LEFT, padx=(8, 4))
        HoverButton(row, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="…", font=("Segoe UI", 10), fg=ACCENT, padx=6, pady=2,
                    command=lambda v=var: v.set(
                        filedialog.askdirectory(parent=self) or v.get()
                    )).pack(side=tk.LEFT)
        self._vars[key] = var

    def _campo_combo(self, key, label, opciones, **kw):
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text=label, width=32, anchor="w",
                 font=("Segoe UI", 9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar(value=opciones[0])
        ttk.Combobox(row, textvariable=var, values=opciones,
                     state="readonly", width=34).pack(side=tk.LEFT, padx=(8, 0))
        self._vars[key] = var

    def _campo_check(self, key, label, **kw):
        var = tk.BooleanVar()
        tk.Checkbutton(self.scroll_frame, text=label, variable=var,
                       font=("Segoe UI", 10), bg=BG_MAIN, fg=TEXT_DARK,
                       activebackground=BG_MAIN, selectcolor=BG_MAIN,
                       cursor="hand2").pack(anchor="w", **kw)
        self._vars[key] = var

    def _campo_hoja(self, **kw):
        """Combo de hojas con botón para detectarlas automáticamente."""
        row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        row.pack(fill=tk.X, **kw)
        tk.Label(row, text="Hoja de trabajo", width=32, anchor="w",
                 font=("Segoe UI", 9), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        var = tk.StringVar()
        self._combo_hoja = ttk.Combobox(row, textvariable=var,
                                         state="readonly", width=28)
        self._combo_hoja.pack(side=tk.LEFT, padx=(8, 4))
        HoverButton(row, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="Detectar", font=("Segoe UI", 9), fg=ACCENT, padx=8, pady=2,
                    command=self._detectar_hojas).pack(side=tk.LEFT)
        self._vars["hoja"] = var

    # ── Cargar / Guardar ──────────────────────────────────────────────────────

    def _cargar_datos(self):
        p = self.perfil
        _sv = lambda key, val: self._vars[key].set(val) if val else None

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
        carpeta = self._vars["carpeta_origen"].get()
        patron  = self._vars["patron_archivo"].get()
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
            messagebox.showerror("Error", str(e), parent=self)

    def _guardar(self):
        g  = lambda k: self._vars[k].get().strip()
        nombre = g("nombre")
        if not nombre:
            messagebox.showwarning("Campo requerido",
                                   "El nombre del perfil es obligatorio.", parent=self)
            return

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
                "hora":   g("hora_ejecucion") or "07:00",
                "activo": self._vars["activo"].get(),
            },
            "notificacion": {
                "email_destino": g("email_destino"),
                "smtp_host":     g("smtp_host"),
                "smtp_port":     int(g("smtp_port") or 587),
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

        messagebox.showinfo("Guardado", f"Perfil '{nombre}' guardado correctamente.",
                            parent=self)
        self.destroy()

    def _probar(self):
        self._guardar()
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
