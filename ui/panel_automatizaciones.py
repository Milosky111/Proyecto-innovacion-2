# ui/panel_automatizaciones.py
"""
Panel que muestra todos los perfiles configurados, su estado de último run
y permite crear, editar, eliminar y ejecutar manualmente cada uno.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

from config import (BG_MAIN, BG_CARD, BG_SIDEBAR, ACCENT, ACCENT_HOVER,
                    SUCCESS, ERROR, WARNING, TEXT_DARK, TEXT_LIGHT,
                    TEXT_MUTED, BORDER, F)
from components import (HoverButton, Tooltip, AyudaInline, abrir_carpeta_de,
                        recordar_carpeta, ultima_carpeta, explicar_error)
from core.config_store import ConfigStore
from core.logger       import RunLogger
from core.runner       import ejecutar_perfil
from core.audit_report import generar_informe


ESTADO_COLOR = {
    "ok":          "#1E8449",
    "sin_archivo": "#D4AC0D",
    "error":       "#C0392B",
    None:          "#6B7A8D",
}
ESTADO_LABEL = {
    "ok":          "✔ OK",
    "sin_archivo": "⚠ Sin archivo",
    "error":       "✖ Error",
    None:          "— Sin ejecutar",
}


class PanelAutomatizaciones(tk.Toplevel):
    def __init__(self, parent, store: ConfigStore, logger: RunLogger,
                 on_editar_callback=None):
        super().__init__(parent)
        self.title("Automatizaciones configuradas")
        self.geometry("860x560")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.minsize(700, 420)

        self.store    = store
        self.logger   = logger
        self.on_edit  = on_editar_callback  # llamado con perfil_id al editar

        self._build()
        self.actualizar_lista()

    # ── Construcción UI ───────────────────────────────────────────────────────

    def _build(self):
        # Toolbar
        toolbar = tk.Frame(self, bg=BG_SIDEBAR, pady=10)
        toolbar.pack(fill=tk.X)

        tk.Label(toolbar, text="⚙  Automatizaciones", font=F(14, "bold"),
                 bg=BG_SIDEBAR, fg=TEXT_LIGHT).pack(side=tk.LEFT, padx=16)
        AyudaInline(
            toolbar,
            "Una 'automatización' es una tarea programada: revisa una "
            "carpeta, extrae datos de un Excel y los exporta solo, sin "
            "que tengas que abrir la app cada vez. Aquí ves todas las "
            "que has creado y su último resultado.",
            bg=BG_SIDEBAR
        ).pack(side=tk.LEFT, padx=(6, 0))

        btn_nueva = HoverButton(toolbar, bg_normal=SUCCESS, bg_hover="#176138",
                    text="+ Nueva automatización", font=F(10, "bold"),
                    fg=TEXT_LIGHT, padx=14, pady=6,
                    command=self._nueva)
        btn_nueva.pack(side=tk.RIGHT, padx=12)
        Tooltip(btn_nueva,
                "Crea un perfil nuevo: carpeta de origen, hoja, "
                "horario y a dónde exportar. Se abre un formulario "
                "para completarlo.")

        btn_informe = HoverButton(toolbar, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="📄  Informe de auditoría…", font=F(10),
                    fg=ACCENT, padx=12, pady=6,
                    command=self._exportar_informe)
        btn_informe.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(btn_informe,
                "Genera un PDF con el historial de ejecuciones de todas "
                "las automatizaciones: qué corrió, cuándo y con qué "
                "resultado.")

        # Subtítulo explicativo, para que un usuario nuevo entienda de
        # inmediato qué está viendo en la tabla de abajo.
        sub = tk.Frame(self, bg=BG_MAIN)
        sub.pack(fill=tk.X)
        tk.Label(
            sub,
            text="Cada fila es una automatización configurada. Selecciona "
                 "una y usa los botones de abajo para editarla, probarla "
                 "o eliminarla.",
            font=F(9, "italic"), bg=BG_MAIN, fg=TEXT_MUTED,
            anchor="w", padx=16, pady=(8, 0)
        ).pack(fill=tk.X)

        # Banner de alertas — avisa de fallos recientes sin que el usuario
        # tenga que leer toda la tabla para darse cuenta.
        self.banner = tk.Frame(self, bg=BG_MAIN)
        self.banner.pack(fill=tk.X)
        self.lbl_banner = tk.Label(
            self.banner, text="", font=F(10, "bold"),
            bg=BG_MAIN, fg=TEXT_DARK, anchor="w", padx=16, pady=8
        )
        self.lbl_banner.pack(fill=tk.X)

        # Tabla
        frame_tabla = tk.Frame(self, bg=BG_MAIN)
        frame_tabla.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)

        cols = ("nombre", "hoja", "destino", "hora", "ultimo_run", "estado")
        self.tree = ttk.Treeview(frame_tabla, columns=cols,
                                  show="headings", selectmode="browse")

        anchos = {"nombre": 180, "hoja": 90, "destino": 80,
                  "hora": 60, "ultimo_run": 140, "estado": 100}
        titulos = {"nombre": "Nombre", "hoja": "Hoja",
                   "destino": "Formato", "hora": "Hora",
                   "ultimo_run": "Último run", "estado": "Estado"}

        for col in cols:
            self.tree.heading(col, text=titulos[col])
            self.tree.column(col, width=anchos[col], anchor="w")

        self.sb_tree = ttk.Scrollbar(frame_tabla, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.sb_tree.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sb_tree.pack(side=tk.RIGHT, fill=tk.Y)
        Tooltip(self.tree,
                "Estado: ✔ OK = la última ejecución exportó datos bien · "
                "⚠ Sin archivo = no encontró el Excel esperado · "
                "✖ Error = falló al ejecutarse · "
                "— Sin ejecutar = aún no ha corrido nunca.",
                wraplength=320)

        # Mensaje que se muestra en vez de la tabla cuando todavía no hay
        # ninguna automatización creada, para que la pantalla no se vea
        # como un error o un espacio vacío sin explicación.
        self.frame_vacio = tk.Frame(frame_tabla, bg=BG_MAIN)
        tk.Label(
            self.frame_vacio, text="📭", font=F(28), bg=BG_MAIN
        ).pack(pady=(30, 6))
        tk.Label(
            self.frame_vacio, text="Aún no tienes automatizaciones creadas",
            font=F(11, "bold"), bg=BG_MAIN, fg=TEXT_DARK
        ).pack()
        tk.Label(
            self.frame_vacio,
            text="Usa el botón '+ Nueva automatización' de arriba para "
                 "crear la primera.",
            font=F(9), bg=BG_MAIN, fg=TEXT_MUTED
        ).pack(pady=(2, 0))

        self.tree.tag_configure("ok",          foreground=ESTADO_COLOR["ok"])
        self.tree.tag_configure("error",        foreground=ESTADO_COLOR["error"])
        self.tree.tag_configure("sin_archivo",  foreground=ESTADO_COLOR["sin_archivo"])
        self.tree.tag_configure("sin_run",      foreground=ESTADO_COLOR[None])

        # Botones de acción
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 14))

        btn_editar = HoverButton(btn_frame, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="✏  Editar", font=F(10),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self._editar)
        btn_editar.pack(side=tk.LEFT, padx=(0, 8))
        Tooltip(btn_editar,
                "Abre el formulario de la automatización seleccionada "
                "para cambiar su carpeta, hoja, horario u otros datos.")

        btn_ejecutar = HoverButton(btn_frame, bg_normal="#1E8449", bg_hover="#176138",
                    text="▶  Ejecutar ahora", font=F(10),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self._ejecutar_ahora)
        btn_ejecutar.pack(side=tk.LEFT, padx=(0, 8))
        Tooltip(btn_ejecutar,
                "Corre la automatización seleccionada de inmediato, sin "
                "esperar a su horario programado. Útil para probar que "
                "todo funciona bien.")

        btn_eliminar = HoverButton(btn_frame, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="🗑  Eliminar", font=F(10),
                    fg=ERROR, padx=14, pady=7,
                    command=self._eliminar)
        btn_eliminar.pack(side=tk.LEFT, padx=(0, 8))
        Tooltip(btn_eliminar,
                "Borra la automatización seleccionada y desactiva su "
                "tarea programada. Pedirá confirmación antes de borrar.")

        btn_actualizar = HoverButton(btn_frame, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="↺  Actualizar", font=F(10),
                    fg=TEXT_MUTED, padx=14, pady=7,
                    command=self.actualizar_lista)
        btn_actualizar.pack(side=tk.RIGHT)
        Tooltip(btn_actualizar,
                "Refresca la tabla con el estado más reciente de cada "
                "automatización.")

        self.lbl_status = tk.Label(self, text="", font=F(9),
                                    bg=BG_MAIN, fg=TEXT_MUTED)
        self.lbl_status.pack(pady=(0, 8))

    # ── Datos ─────────────────────────────────────────────────────────────────

    def actualizar_lista(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        perfiles = self.store.listar_perfiles()

        if not perfiles:
            self.tree.pack_forget()
            self.sb_tree.pack_forget()
            self.frame_vacio.pack(fill=tk.BOTH, expand=True)
            self._actualizar_banner([], [], 0)
            return
        self.frame_vacio.pack_forget()
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sb_tree.pack(side=tk.RIGHT, fill=tk.Y)

        fallidas, sin_archivo, ok = [], [], 0

        for p in perfiles:
            ultimo = self.logger.ultimo_run(p["id"])
            estado = ultimo["estado"] if ultimo else None
            ts     = ultimo["timestamp"][:16] if ultimo else "—"

            if estado == "error":
                fallidas.append(p.get("nombre", "Sin nombre"))
            elif estado == "sin_archivo":
                sin_archivo.append(p.get("nombre", "Sin nombre"))
            elif estado == "ok":
                ok += 1

            tag = estado if estado else "sin_run"
            self.tree.insert("", tk.END, iid=p["id"], tags=(tag,), values=(
                p.get("nombre", "Sin nombre"),
                p.get("hoja",   "—"),
                p.get("destino", {}).get("tipo", "—"),
                p.get("schedule", {}).get("hora", "—"),
                ts,
                ESTADO_LABEL.get(estado, "—"),
            ))

        self._actualizar_banner(fallidas, sin_archivo, ok)

    def _actualizar_banner(self, fallidas, sin_archivo, ok):
        """
        Banner de alerta visible apenas se abre el panel: resume cuántas
        automatizaciones tienen su última ejecución en error o sin archivo,
        para que el usuario no tenga que leer toda la tabla.
        """
        total_problemas = len(fallidas) + len(sin_archivo)

        if total_problemas == 0:
            self.banner.configure(bg=BG_MAIN)
            self.lbl_banner.configure(
                bg=BG_MAIN, fg=SUCCESS,
                text=f"✔  Todas las automatizaciones con historial están al día ({ok} con última ejecución exitosa)."
                if ok else ""
            )
            return

        partes = []
        if fallidas:
            partes.append(f"{len(fallidas)} con error ({', '.join(fallidas[:3])}{'…' if len(fallidas) > 3 else ''})")
        if sin_archivo:
            partes.append(f"{len(sin_archivo)} sin archivo encontrado ({', '.join(sin_archivo[:3])}{'…' if len(sin_archivo) > 3 else ''})")

        self.banner.configure(bg="#FDEDEC")
        self.lbl_banner.configure(
            bg="#FDEDEC", fg=ERROR,
            text=f"⚠  Atención: {' · '.join(partes)}."
        )

    def _exportar_informe(self):
        """Genera el informe de auditoría en PDF y lo guarda donde el usuario elija."""
        ruta = filedialog.asksaveasfilename(
            title="Guardar informe de auditoría como…",
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")],
            initialfile="informe_auditoria.pdf",
            initialdir=ultima_carpeta(),
            parent=self
        )
        if not ruta:
            return
        recordar_carpeta(ruta)
        try:
            generar_informe(self.logger, ruta)
            if messagebox.askyesno(
                "Informe generado",
                f"El informe de auditoría se guardó en:\n{ruta}\n\n"
                "¿Abrir la carpeta donde se guardó?",
                parent=self
            ):
                abrir_carpeta_de(ruta)
        except Exception as e:
            messagebox.showerror("Error al generar informe", explicar_error(e), parent=self)

    def _get_seleccionado_id(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sin selección", "Selecciona un perfil de la lista.",
                                   parent=self)
            return None
        return sel[0]

    # ── Acciones ──────────────────────────────────────────────────────────────

    def _nueva(self):
        perfil = self.store.nuevo_perfil()
        if self.on_edit:
            self.on_edit(perfil["id"])
        self.actualizar_lista()

    def _editar(self):
        pid = self._get_seleccionado_id()
        if pid and self.on_edit:
            self.on_edit(pid)
            self.actualizar_lista()

    def _eliminar(self):
        pid = self._get_seleccionado_id()
        if not pid:
            return
        perfil = self.store.obtener_perfil(pid)
        if messagebox.askyesno(
            "Confirmar eliminación",
            f"¿Eliminar el perfil '{perfil['nombre']}'?\nEsta acción no se puede deshacer.",
            parent=self
        ):
            self.store.eliminar_perfil(pid)
            self.actualizar_lista()

    def _ejecutar_ahora(self):
        pid = self._get_seleccionado_id()
        if not pid:
            return
        perfil = self.store.obtener_perfil(pid)
        self.lbl_status.config(text=f"Ejecutando '{perfil['nombre']}'…",
                                fg=ACCENT)
        self.update()

        def _run():
            r = ejecutar_perfil(perfil, logger=self.logger)
            self.store.marcar_ultimo_run(pid)
            self.after(0, lambda: self._post_ejecucion(perfil["nombre"], r))

        threading.Thread(target=_run, daemon=True).start()

    def _post_ejecucion(self, nombre, resultado):
        self.actualizar_lista()
        estado = resultado["estado"]
        if estado == "ok":
            msg = f"✔ '{nombre}' exportó {resultado['filas']} filas correctamente."
            self.lbl_status.config(text=msg, fg=SUCCESS)
        else:
            msg = f"✖ Error en '{nombre}': {resultado['error']}"
            self.lbl_status.config(text=msg, fg=ERROR)
            messagebox.showerror("Error en ejecución", msg, parent=self)
