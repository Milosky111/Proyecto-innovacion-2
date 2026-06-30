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
                    TEXT_MUTED, BORDER)
from components import HoverButton
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

        tk.Label(toolbar, text="⚙  Automatizaciones", font=("Segoe UI", 14, "bold"),
                 bg=BG_SIDEBAR, fg=TEXT_LIGHT).pack(side=tk.LEFT, padx=16)

        HoverButton(toolbar, bg_normal=SUCCESS, bg_hover="#176138",
                    text="+ Nueva automatización", font=("Segoe UI", 10, "bold"),
                    fg=TEXT_LIGHT, padx=14, pady=6,
                    command=self._nueva).pack(side=tk.RIGHT, padx=12)

        HoverButton(toolbar, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="📄  Informe de auditoría…", font=("Segoe UI", 10),
                    fg=ACCENT, padx=12, pady=6,
                    command=self._exportar_informe).pack(side=tk.RIGHT, padx=(0, 8))

        # Banner de alertas — avisa de fallos recientes sin que el usuario
        # tenga que leer toda la tabla para darse cuenta.
        self.banner = tk.Frame(self, bg=BG_MAIN)
        self.banner.pack(fill=tk.X)
        self.lbl_banner = tk.Label(
            self.banner, text="", font=("Segoe UI", 10, "bold"),
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

        sb = ttk.Scrollbar(frame_tabla, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("ok",          foreground=ESTADO_COLOR["ok"])
        self.tree.tag_configure("error",        foreground=ESTADO_COLOR["error"])
        self.tree.tag_configure("sin_archivo",  foreground=ESTADO_COLOR["sin_archivo"])
        self.tree.tag_configure("sin_run",      foreground=ESTADO_COLOR[None])

        # Botones de acción
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 14))

        HoverButton(btn_frame, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="✏  Editar", font=("Segoe UI", 10),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self._editar).pack(side=tk.LEFT, padx=(0, 8))

        HoverButton(btn_frame, bg_normal="#1E8449", bg_hover="#176138",
                    text="▶  Ejecutar ahora", font=("Segoe UI", 10),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self._ejecutar_ahora).pack(side=tk.LEFT, padx=(0, 8))

        HoverButton(btn_frame, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="🗑  Eliminar", font=("Segoe UI", 10),
                    fg=ERROR, padx=14, pady=7,
                    command=self._eliminar).pack(side=tk.LEFT, padx=(0, 8))

        HoverButton(btn_frame, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="↺  Actualizar", font=("Segoe UI", 10),
                    fg=TEXT_MUTED, padx=14, pady=7,
                    command=self.actualizar_lista).pack(side=tk.RIGHT)

        self.lbl_status = tk.Label(self, text="", font=("Segoe UI", 9),
                                    bg=BG_MAIN, fg=TEXT_MUTED)
        self.lbl_status.pack(pady=(0, 8))

    # ── Datos ─────────────────────────────────────────────────────────────────

    def actualizar_lista(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        perfiles = self.store.listar_perfiles()
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
            parent=self
        )
        if not ruta:
            return
        try:
            generar_informe(self.logger, ruta)
            messagebox.showinfo(
                "Informe generado",
                f"El informe de auditoría se guardó en:\n{ruta}",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("Error al generar informe", str(e), parent=self)

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
