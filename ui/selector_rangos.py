# ui/selector_rangos.py
"""
Ventana modal para que el usuario defina rangos de celdas por hoja.
Flujo por hoja:
  1. Escribir rango (ej: B3:R35)
  2. Previsualizar datos extraídos
  3. Renombrar columnas del rango
  4. Guardar — queda asociado al perfil

El resultado es un dict:
  {
    "CAJAS": {
        "rango": "B3:R35",
        "encabezado_en_fila": 0,   # 0 = primera fila del rango es encabezado
        "rename_map": {"Col orig": "Nombre legible", ...}
    },
    ...
  }
"""

import tkinter as tk
from tkinter import ttk, messagebox

from config import (BG_MAIN, BG_CARD, BG_SIDEBAR, ACCENT, ACCENT_HOVER,
                    SUCCESS, ERROR, TEXT_DARK, TEXT_LIGHT, TEXT_MUTED,
                    BORDER, ROW_EVEN, ROW_ODD, F)
from components import HoverButton, Tooltip, AyudaInline, crear_entry, crear_spinbox, explicar_error
from core.excel_reader import ExcelReader
from ui.click_range_selector import ClickRangeSelector


class SelectorRangos(tk.Toplevel):
    """
    Ventana modal de configuración de rangos.
    Retorna los rangos configurados en self.resultado al cerrarse.
    """

    def __init__(self, parent, reader: ExcelReader, rangos_previos: dict = None):
        super().__init__(parent)
        self.title("Configurar rangos por hoja")
        self.geometry("900x620")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.minsize(780, 500)
        self.grab_set()

        self.reader        = reader
        self.hojas         = reader.obtener_hojas()
        self.resultado     = dict(rangos_previos or {})  # {hoja: {rango, rename_map}}
        # Copia aparte del estado original: 'Guardar rango de esta hoja'
        # modifica self.resultado al instante (hoja por hoja), así que sin
        # esta copia, cerrar con 'Cancelar' dejaría aplicados los rangos
        # que ya se guardaron en esta sesión — aunque el botón diga que
        # cierra sin guardar cambios.
        self._resultado_original = dict(rangos_previos or {})
        self._hoja_actual  = None
        self._rename_vars  = {}   # {col_original: StringVar}
        self._enc_var      = tk.IntVar(value=0)

        self._build()
        if self.hojas:
            self._seleccionar_hoja(self.hojas[0])
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build(self):
        # ── Layout principal: lista izquierda + panel derecho
        paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, bg=BG_MAIN,
                                sashwidth=4, sashrelief="flat")
        paned.pack(fill=tk.BOTH, expand=True)

        # Panel izquierdo: hojas
        left = tk.Frame(paned, bg=BG_SIDEBAR, width=190)
        paned.add(left, minsize=160)
        self._build_panel_hojas(left)

        # Panel derecho: configuración del rango
        right = tk.Frame(paned, bg=BG_MAIN)
        paned.add(right, minsize=560)
        self._build_panel_config(right)

        # Barra inferior
        bar = tk.Frame(self, bg=BG_MAIN, pady=10)
        bar.pack(fill=tk.X, side=tk.BOTTOM)

        btn_confirmar = HoverButton(bar, bg_normal=SUCCESS, bg_hover="#176138",
                    text="✔  Confirmar configuración",
                    font=F(11, "bold"),
                    fg=TEXT_LIGHT, padx=20, pady=8,
                    command=self._confirmar)
        btn_confirmar.pack(side=tk.RIGHT, padx=16)
        Tooltip(btn_confirmar,
                "Cierra esta ventana y guarda los rangos configurados en "
                "todas las hojas. La ventana principal quedará en modo "
                "'Rangos de celdas' lista para exportar.")

        btn_cancelar = HoverButton(bar, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="Cancelar", font=F(10),
                    fg=TEXT_DARK, padx=14, pady=8,
                    command=self._cancelar)
        btn_cancelar.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(btn_cancelar, "Cierra sin guardar cambios en esta ventana.")

        self.lbl_status = tk.Label(bar, text="", font=F(9),
                                    bg=BG_MAIN, fg=TEXT_MUTED)
        self.lbl_status.pack(side=tk.LEFT, padx=16)

    def _build_panel_hojas(self, parent):
        tk.Label(parent, text="Hojas del archivo",
                 font=F(10, "bold"),
                 bg=BG_SIDEBAR, fg=TEXT_LIGHT,
                 pady=12).pack(fill=tk.X, padx=12)
        tk.Label(parent,
                 text="Configura el rango de\ncada hoja por separado.\n"
                      "En verde: ya tiene rango.",
                 font=F(8, "italic"), justify="left",
                 bg=BG_SIDEBAR, fg="#A8C4E0").pack(fill=tk.X, padx=12, pady=(0, 10))

        self._btns_hoja = {}
        for hoja in self.hojas:
            # Se usa tk.Label (no tk.Button) a propósito: en macOS, un
            # tk.Button con fondo oscuro se dibuja igual con el botón
            # gris nativo del sistema, lo que rompe el look de la barra
            # lateral oscura. tk.Label sí respeta los colores en
            # cualquier plataforma, y con un bind de clic funciona igual
            # de bien como opción de lista.
            btn = tk.Label(
                parent, text=hoja,
                font=F(9), anchor="w",
                bg=BG_SIDEBAR, fg="#A8C4E0",
                cursor="hand2", padx=12, pady=6
            )
            btn.bind("<Button-1>", lambda e, h=hoja: self._seleccionar_hoja(h))
            btn.pack(fill=tk.X, padx=4, pady=1)
            Tooltip(btn, f"Configurar el rango de celdas de la hoja '{hoja}'.")
            self._btns_hoja[hoja] = btn

    def _build_panel_config(self, parent):
        # Header de hoja seleccionada
        self.lbl_hoja_header = tk.Label(
            parent, text="Selecciona una hoja",
            font=F(13, "bold"),
            bg=BG_MAIN, fg=TEXT_DARK, anchor="w"
        )
        self.lbl_hoja_header.pack(fill=tk.X, padx=20, pady=(16, 2))

        self.lbl_hoja_sub = tk.Label(
            parent, text="Define el rango de celdas a extraer",
            font=F(9, "italic"),
            bg=BG_MAIN, fg=TEXT_MUTED, anchor="w"
        )
        self.lbl_hoja_sub.pack(fill=tk.X, padx=20, pady=(0, 12))

        ttk.Separator(parent, orient="horizontal").pack(fill=tk.X, padx=20)

        # ── Sección rango ──────────────────────────────────────────────────
        sec_rango = tk.Frame(parent, bg=BG_MAIN)
        sec_rango.pack(fill=tk.X, padx=20, pady=(14, 0))

        titulo_rango = tk.Frame(sec_rango, bg=BG_MAIN)
        titulo_rango.pack(fill=tk.X, anchor="w")
        tk.Label(titulo_rango, text="Rango de celdas",
                 font=F(10, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(side=tk.LEFT)
        AyudaInline(
            titulo_rango,
            "Este modo abre el archivo con Microsoft Excel real (no solo "
            "lo lee como datos) para traer valores calculados de fórmulas "
            "y listas desplegables correctamente. Por eso necesita tener "
            "Microsoft Excel instalado en este computador — si no lo "
            "tienes, usa mejor 'Columnas completas' en la ventana "
            "principal.",
            bg=BG_MAIN
        ).pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(sec_rango,
                 text="Notación Excel: B3:R35  —  incluye fila de encabezados",
                 font=F(8, "italic"),
                 bg=BG_MAIN, fg=TEXT_MUTED).pack(anchor="w", pady=(2, 6))

        row_rango = tk.Frame(sec_rango, bg=BG_MAIN)
        row_rango.pack(fill=tk.X)

        self.entry_rango = crear_entry(row_rango, font=F(11), width=14)
        self.entry_rango.pack(side=tk.LEFT, ipady=4)
        Tooltip(self.entry_rango,
                "Escribe el rango en notación Excel, ej: B3:R35 "
                "(columna+fila inicial : columna+fila final). También "
                "puedes usar el botón 'Marcar con clic…' si prefieres "
                "no escribirlo a mano.")

        tk.Label(row_rango, text="Encabezado en fila",
                 font=F(9), bg=BG_MAIN, fg=TEXT_MUTED
                 ).pack(side=tk.LEFT, padx=(16, 4))

        self._enc_spin = crear_spinbox(
            row_rango, from_=1, to=10,
            textvariable=self._enc_var, width=4
        )
        self._enc_spin.pack(side=tk.LEFT)
        Tooltip(self._enc_spin,
                "¿En qué fila del rango están los nombres de columna? "
                "1 = la primera fila del rango. Si tus datos empiezan "
                "inmediatamente (sin encabezado), deja 1.")

        tk.Label(row_rango, text="del rango",
                 font=F(9), bg=BG_MAIN, fg=TEXT_MUTED
                 ).pack(side=tk.LEFT, padx=(4, 0))

        btn_prev = HoverButton(row_rango, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="▶  Previsualizar",
                    font=F(10, "bold"),
                    fg=TEXT_LIGHT, padx=14, pady=5,
                    command=self._previsualizar)
        btn_prev.pack(side=tk.LEFT, padx=(16, 0))
        Tooltip(btn_prev,
                "Lee el rango escrito y muestra abajo las primeras 5 "
                "filas, para que confirmes que es el área correcta "
                "antes de guardar.")

        btn_clic = HoverButton(row_rango, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="🖱  Marcar con clic…",
                    font=F(10, "bold"),
                    fg=ACCENT, padx=14, pady=5,
                    command=self._abrir_selector_clic)
        btn_clic.pack(side=tk.LEFT, padx=(8, 0))
        Tooltip(btn_clic,
                "Alternativa a escribir el rango a mano: abre una "
                "grilla con los datos reales de la hoja y marca el "
                "área haciendo clic en la celda inicial y luego en "
                "la celda final.")

        # ── Sección previsualización ───────────────────────────────────────
        tk.Label(parent, text="Previsualización (primeras 5 filas)",
                 font=F(10, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(anchor="w", padx=20, pady=(14, 4))

        frame_prev = tk.Frame(parent, bg=BG_MAIN)
        frame_prev.pack(fill=tk.X, padx=20)

        self.tree_prev = ttk.Treeview(frame_prev, show="headings",
                                       selectmode="none", height=5)
        sb_prev = ttk.Scrollbar(frame_prev, orient="horizontal",
                                  command=self.tree_prev.xview)
        self.tree_prev.configure(xscrollcommand=sb_prev.set)
        self.tree_prev.pack(fill=tk.X)
        sb_prev.pack(fill=tk.X)

        # ── Sección renombrado ─────────────────────────────────────────────
        tk.Label(parent, text="Renombrar columnas",
                 font=F(10, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(anchor="w", padx=20, pady=(14, 2))
        tk.Label(parent,
                 text="Deja en blanco para mantener el nombre original del Excel.",
                 font=F(8, "italic"),
                 bg=BG_MAIN, fg=TEXT_MUTED).pack(anchor="w", padx=20)

        # Frame scrollable para campos de renombrado
        outer = tk.Frame(parent, bg=BG_MAIN)
        outer.pack(fill=tk.BOTH, expand=True, padx=20, pady=(6, 0))

        canvas_r = tk.Canvas(outer, bg=BG_MAIN, highlightthickness=0, height=120)
        sb_r = ttk.Scrollbar(outer, orient="vertical", command=canvas_r.yview)
        self._frame_rename = tk.Frame(canvas_r, bg=BG_MAIN)
        self._frame_rename.bind(
            "<Configure>",
            lambda e: canvas_r.configure(scrollregion=canvas_r.bbox("all"))
        )
        canvas_r.create_window((0, 0), window=self._frame_rename, anchor="nw")
        canvas_r.configure(yscrollcommand=sb_r.set)
        canvas_r.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_r.pack(side=tk.RIGHT, fill=tk.Y)

        # Botón guardar rango
        btn_guardar = HoverButton(parent, bg_normal=SUCCESS, bg_hover="#176138",
                    text="💾  Guardar rango de esta hoja",
                    font=F(10, "bold"),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self._guardar_rango_hoja)
        btn_guardar.pack(anchor="e", padx=20, pady=10)
        Tooltip(btn_guardar,
                "Guarda el rango y los renombres definidos para ESTA "
                "hoja. Repite el proceso para cada hoja que necesites "
                "y luego presiona 'Confirmar configuración' abajo.")

    # ── Lógica ────────────────────────────────────────────────────────────────

    def _abrir_selector_clic(self):
        """
        Abre la grilla visual de selección por clic. Si el usuario confirma
        un rango, lo escribe en self.entry_rango y dispara la previsualización
        automáticamente — el resultado se integra al flujo normal del modo
        texto, no lo reemplaza.
        """
        if not self._hoja_actual:
            messagebox.showwarning("Sin hoja", "Selecciona una hoja primero.", parent=self)
            return

        try:
            filas_preview = self.reader.obtener_vista_previa(self._hoja_actual)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer la vista previa:\n{e}", parent=self)
            return

        if not filas_preview:
            messagebox.showinfo("Hoja vacía", "Esta hoja no tiene datos para mostrar.", parent=self)
            return

        selector = ClickRangeSelector(self, filas_preview)
        self.wait_window(selector)

        if selector.result:
            self.entry_rango.delete(0, tk.END)
            self.entry_rango.insert(0, selector.result)
            self._previsualizar()

    def _seleccionar_hoja(self, hoja):
        # Resaltar botón activo
        for h, btn in self._btns_hoja.items():
            if h == hoja:
                btn.config(bg=ACCENT, fg=TEXT_LIGHT, font=F(9, "bold"))
            else:
                tiene = hoja in self.resultado
                btn.config(
                    bg=BG_SIDEBAR,
                    fg="#5DCA8A" if h in self.resultado else "#A8C4E0",
                    font=F(9)
                )

        self._hoja_actual = hoja
        self.lbl_hoja_header.config(text=f"Hoja: {hoja}")
        self.lbl_hoja_sub.config(
            text="✔ Rango configurado" if hoja in self.resultado
            else "Sin rango configurado — define uno abajo"
        )

        # Cargar datos previos si existen
        prev = self.resultado.get(hoja, {})
        self.entry_rango.delete(0, tk.END)
        if prev.get("rango"):
            self.entry_rango.insert(0, prev["rango"])
            self._enc_var.set(prev.get("encabezado_en_fila", 0) + 1)
            self._previsualizar(rename_map=prev.get("rename_map", {}))
        else:
            self._limpiar_preview()
            self._limpiar_rename()

    def _previsualizar(self, rename_map=None):
        rango = self.entry_rango.get().strip().upper()
        if not rango:
            messagebox.showwarning("Rango vacío",
                                   "Escribe un rango antes de previsualizar.\nEjemplo: B3:R35",
                                   parent=self)
            return

        enc_fila = self._enc_var.get() - 1  # 0-based dentro del rango

        try:
            df = self.reader.leer_rango(self._hoja_actual, rango,
                                         fila_encabezado=enc_fila)
        except Exception as e:
            messagebox.showerror("Error al leer rango", explicar_error(e), parent=self)
            return

        if df.empty:
            messagebox.showwarning("Sin datos",
                                   "El rango no retornó datos. Verifica la notación.",
                                   parent=self)
            return

        # ── Previsualización
        self._limpiar_preview()
        cols = list(df.columns)
        self.tree_prev["columns"] = cols
        for col in cols:
            self.tree_prev.heading(col, text=col)
            ancho = max(len(str(col)) * 8 + 16, 80)
            self.tree_prev.column(col, width=min(ancho, 160), anchor="w")

        for i, (_, row) in enumerate(df.head(5).iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree_prev.insert("", tk.END,
                                   values=[str(v) if v is not None else "" for v in row],
                                   tags=(tag,))
        self.tree_prev.tag_configure("even", background=ROW_EVEN)
        self.tree_prev.tag_configure("odd",  background=ROW_ODD)

        # ── Campos de renombrado
        self._limpiar_rename()
        self._rename_vars = {}
        rename_map = rename_map or {}

        # Encabezado de la sección
        hdr = tk.Frame(self._frame_rename, bg=BG_MAIN)
        hdr.pack(fill=tk.X, pady=(0, 4))
        tk.Label(hdr, text="Columna original", width=28, anchor="w",
                 font=F(9, "bold"), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        tk.Label(hdr, text="→  Nuevo nombre (opcional)", anchor="w",
                 font=F(9, "bold"), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT, padx=(8, 0))

        for col in cols:
            row_f = tk.Frame(self._frame_rename, bg=BG_MAIN)
            row_f.pack(fill=tk.X, pady=2)

            tk.Label(row_f, text=str(col)[:30], width=28, anchor="w",
                     font=F(9), bg=BG_MAIN, fg=TEXT_DARK).pack(side=tk.LEFT)

            var = tk.StringVar(value=rename_map.get(str(col), ""))
            entry = crear_entry(row_f, textvariable=var, font=F(9), width=30)
            entry.pack(side=tk.LEFT, padx=(8, 0))
            self._rename_vars[str(col)] = var

        self.lbl_status.config(
            text=f"Rango {rango} → {len(df)} filas × {len(cols)} columnas",
            fg=SUCCESS
        )

    def _guardar_rango_hoja(self):
        if not self._hoja_actual:
            return
        rango = self.entry_rango.get().strip().upper()
        if not rango:
            messagebox.showwarning("Sin rango",
                                   "Define un rango antes de guardar.", parent=self)
            return
        if not self._rename_vars:
            messagebox.showwarning("Sin previsualización",
                                   "Haz clic en 'Previsualizar' primero para ver las columnas.",
                                   parent=self)
            return

        rename_map = {
            col: var.get().strip()
            for col, var in self._rename_vars.items()
            if var.get().strip()
        }

        self.resultado[self._hoja_actual] = {
            "rango":              rango,
            "encabezado_en_fila": self._enc_var.get() - 1,
            "rename_map":         rename_map,
        }

        # Actualizar color del botón de hoja
        btn = self._btns_hoja.get(self._hoja_actual)
        if btn:
            btn.config(fg="#5DCA8A")

        self.lbl_hoja_sub.config(text="✔ Rango configurado")
        self.lbl_status.config(
            text=f"✔ Rango guardado para '{self._hoja_actual}'",
            fg=SUCCESS
        )

    def _confirmar(self):
        if not self.resultado:
            messagebox.showwarning("Sin configuración",
                                   "Configura al menos un rango antes de confirmar.",
                                   parent=self)
            return
        self.destroy()

    def _cancelar(self):
        if self.resultado != self._resultado_original:
            if not messagebox.askyesno(
                "Descartar cambios",
                "Configuraste rangos en esta sesión que aún no se han "
                "confirmado.\n¿Cerrar sin guardarlos?",
                parent=self
            ):
                return
            # Se restaura el estado original para que la ventana que
            # abrió este diálogo no vea ningún cambio aplicado.
            self.resultado = self._resultado_original
        self.destroy()

    def _limpiar_preview(self):
        self.tree_prev["columns"] = ()
        for item in self.tree_prev.get_children():
            self.tree_prev.delete(item)

    def _limpiar_rename(self):
        for widget in self._frame_rename.winfo_children():
            widget.destroy()
        self._rename_vars = {}
