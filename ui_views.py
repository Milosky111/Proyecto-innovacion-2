# ui_views.py
import tkinter as tk
from tkinter import ttk
import os

from config import (
    BG_MAIN, BG_SIDEBAR, BG_CARD, ACCENT, ACCENT_HOVER, SUCCESS, 
    WARNING, ERROR, TEXT_LIGHT, TEXT_DARK, TEXT_MUTED, BORDER, ROW_EVEN, ROW_ODD
)
from components import HoverButton

class UIManager:
    def __init__(self, root, controlador):
        self.root = root
        self.controlador = controlador # Conecta la interfaz con la lógica de extractor.py
        
        self._build_ui()

    def _build_ui(self):
        self.sidebar = tk.Frame(self.root, bg=BG_SIDEBAR, width=220)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        self.panel = tk.Frame(self.root, bg=BG_MAIN)
        self.panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_sidebar()
        self._build_panel()

    def _build_sidebar(self):
        tk.Label(self.sidebar, text="📊", font=("Segoe UI", 28),
                 bg=BG_SIDEBAR, fg=TEXT_LIGHT).pack(pady=(28, 4))
        tk.Label(self.sidebar, text="Extractor\nde Datos",
                 font=("Segoe UI", 14, "bold"), bg=BG_SIDEBAR,
                 fg=TEXT_LIGHT, justify="center").pack()

        ttk.Separator(self.sidebar, orient="horizontal").pack(fill=tk.X, padx=20, pady=20)

        self.step_labels = []
        pasos = [("1", "Cargar archivo"), ("2", "Elegir hoja"), 
                 ("3", "Seleccionar\ncolumnas"), ("4", "Exportar")]
        
        for num, texto in pasos:
            f = tk.Frame(self.sidebar, bg=BG_SIDEBAR)
            f.pack(fill=tk.X, padx=18, pady=4)

            badge = tk.Label(f, text=num, width=2, font=("Segoe UI", 10, "bold"),
                             bg=ACCENT, fg=TEXT_LIGHT, padx=4, pady=2)
            badge.pack(side=tk.LEFT, padx=(0, 10))

            lbl = tk.Label(f, text=texto, font=("Segoe UI", 10),
                           bg=BG_SIDEBAR, fg="#A8C4E0", justify="left", anchor="w")
            lbl.pack(side=tk.LEFT, fill=tk.X)
            self.step_labels.append((badge, lbl))

        tk.Frame(self.sidebar, bg=BG_SIDEBAR).pack(fill=tk.BOTH, expand=True)

        self.lbl_estado_dot = tk.Label(self.sidebar, text="●", font=("Segoe UI", 10), bg=BG_SIDEBAR, fg="#A8C4E0")
        self.lbl_estado_dot.pack(pady=(0, 2))

        self.lbl_estado = tk.Label(self.sidebar, text="Esperando archivo…", font=("Segoe UI", 9), 
                                   bg=BG_SIDEBAR, fg="#A8C4E0", wraplength=180, justify="center")
        self.lbl_estado.pack(pady=(0, 20))

    def _build_panel(self):
        header = tk.Frame(self.panel, bg=BG_MAIN)
        header.pack(fill=tk.X, padx=20, pady=(20, 16))

        tk.Label(header, text="Extractor de Datos Excel", font=("Segoe UI", 18, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(anchor="w")
        tk.Label(header, text="Carga cualquier Excel, elige las columnas y exporta.",
                 font=("Segoe UI", 10), bg=BG_MAIN, fg=TEXT_MUTED).pack(anchor="w")

        self._card_archivo()
        self._card_hoja()
        self._card_columnas()
        self._card_exportar()

    def _make_card(self, title, emoji=""):
        outer = tk.Frame(self.panel, bg=BG_MAIN)
        outer.pack(fill=tk.X, padx=20, pady=(0, 12))

        card = tk.Frame(outer, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        tk.Frame(card, bg=ACCENT, height=4).pack(fill=tk.X)

        header = tk.Frame(card, bg=BG_CARD)
        header.pack(fill=tk.X, padx=16, pady=(10, 8))
        tk.Label(header, text=f"{emoji}  {title}", font=("Segoe UI", 10, "bold"), bg=BG_CARD, fg=TEXT_DARK).pack(anchor="w")

        body = tk.Frame(card, bg=BG_CARD)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 14))
        return body

    def _card_archivo(self):
        body = self._make_card("Cargar archivo Excel", "📂")
        row = tk.Frame(body, bg=BG_CARD)
        row.pack(fill=tk.X)

        HoverButton(row, bg_normal=ACCENT, bg_hover=ACCENT_HOVER, text="Buscar archivo…",
                    font=("Segoe UI", 10, "bold"), fg=TEXT_LIGHT, padx=16, pady=7,
                    command=self.controlador.cargar_archivo).pack(side=tk.LEFT)

        self.lbl_archivo = tk.Label(row, text="Ningún archivo seleccionado", font=("Segoe UI", 10, "italic"),
                                    bg=BG_CARD, fg=TEXT_MUTED, anchor="w")
        self.lbl_archivo.pack(side=tk.LEFT, padx=(14, 0), fill=tk.X, expand=True)

    def _card_hoja(self):
        body = self._make_card("Seleccionar hoja de trabajo", "📄")
        row = tk.Frame(body, bg=BG_CARD)
        row.pack(fill=tk.X)

        tk.Label(row, text="Hoja:", font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_DARK).pack(side=tk.LEFT, padx=(0, 8))

        self.combo_hojas = ttk.Combobox(row, state="disabled", font=("Segoe UI", 10), width=36)
        self.combo_hojas.pack(side=tk.LEFT)
        self.combo_hojas.bind("<<ComboboxSelected>>", self.controlador.cargar_columnas)

        self.lbl_info_hoja = tk.Label(row, text="", font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_info_hoja.pack(side=tk.LEFT, padx=(12, 0))

    def _card_columnas(self):
        outer = tk.Frame(self.panel, bg=BG_MAIN)
        outer.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 12))

        card = tk.Frame(outer, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        tk.Frame(card, bg=ACCENT, height=4).pack(fill=tk.X)

        # ── Tabs: modo columnas / modo rangos ─────────────────────────────
        tab_bar = tk.Frame(card, bg=BG_CARD)
        tab_bar.pack(fill=tk.X, padx=16, pady=(10, 0))

        self.var_modo = tk.StringVar(value="columnas")

        self.tab_col = tk.Radiobutton(
            tab_bar, text="🗂  Columnas completas",
            variable=self.var_modo, value="columnas",
            font=("Segoe UI", 10, "bold"),
            bg=BG_CARD, fg=TEXT_DARK,
            activebackground=BG_CARD, selectcolor=BG_CARD,
            cursor="hand2", indicatoron=False,
            relief="solid", bd=1,
            padx=12, pady=5,
            command=self.controlador._cambiar_modo
        )
        self.tab_col.pack(side=tk.LEFT, padx=(0, 4))

        self.tab_rango = tk.Radiobutton(
            tab_bar, text="⊞  Rangos de celdas",
            variable=self.var_modo, value="rangos",
            font=("Segoe UI", 10, "bold"),
            bg=BG_CARD, fg=TEXT_DARK,
            activebackground=BG_CARD, selectcolor=BG_CARD,
            cursor="hand2", indicatoron=False,
            relief="solid", bd=1,
            padx=12, pady=5,
            command=self.controlador._cambiar_modo
        )
        self.tab_rango.pack(side=tk.LEFT)

        # ── Contenedor de paneles intercambiables ──────────────────────────
        self.frame_modos = tk.Frame(card, bg=BG_CARD)
        self.frame_modos.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 14))

        # Panel modo columnas
        self.panel_columnas = tk.Frame(self.frame_modos, bg=BG_CARD)
        self._build_panel_columnas(self.panel_columnas)

        # Panel modo rangos
        self.panel_rangos = tk.Frame(self.frame_modos, bg=BG_CARD)
        self._build_panel_rangos(self.panel_rangos)

        # Mostrar columnas por defecto
        self.panel_columnas.pack(fill=tk.BOTH, expand=True)

    def _build_panel_columnas(self, parent):
        header = tk.Frame(parent, bg=BG_CARD)
        header.pack(fill=tk.X)

        tk.Label(header, text="Ctrl+clic o Shift+clic para seleccionar varias",
                 font=("Segoe UI", 8, "italic"),
                 bg=BG_CARD, fg=TEXT_MUTED).pack(side=tk.LEFT)

        btn_frame = tk.Frame(header, bg=BG_CARD)
        btn_frame.pack(side=tk.RIGHT)

        HoverButton(btn_frame, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="Seleccionar todo", font=("Segoe UI", 8),
                    fg=ACCENT, padx=8, pady=4,
                    command=self.controlador._seleccionar_todo).pack(side=tk.LEFT, padx=(0, 6))

        HoverButton(btn_frame, bg_normal="#FEE8E8", bg_hover="#FCC5C5",
                    text="Limpiar", font=("Segoe UI", 8),
                    fg=ERROR, padx=8, pady=4,
                    command=self.controlador._limpiar_seleccion).pack(side=tk.LEFT)

        tree_frame = tk.Frame(parent, bg=BG_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        style = ttk.Style()
        style.configure("Col.Treeview", font=("Segoe UI", 10), rowheight=28,
                         background=BG_CARD, fieldbackground=BG_CARD, borderwidth=0)
        style.configure("Col.Treeview.Heading", font=("Segoe UI", 10, "bold"),
                         background="#E8F0FE", foreground=TEXT_DARK, relief="flat")
        style.map("Col.Treeview",
                   background=[("selected", ACCENT)],
                   foreground=[("selected", TEXT_LIGHT)])

        self.lista_columnas = ttk.Treeview(
            tree_frame, columns=("Columna",),
            show="headings", selectmode="extended", style="Col.Treeview")
        self.lista_columnas.heading("Columna", text="Columnas disponibles")
        self.lista_columnas.column("Columna", anchor="w")
        self.lista_columnas.tag_configure("even", background=ROW_EVEN)
        self.lista_columnas.tag_configure("odd",  background=ROW_ODD)

        sb = ttk.Scrollbar(tree_frame, orient="vertical",
                             command=self.lista_columnas.yview)
        self.lista_columnas.configure(yscrollcommand=sb.set)
        self.lista_columnas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lista_columnas.bind("<<TreeviewSelect>>", self.controlador._actualizar_contador)

        self.lbl_contador = tk.Label(parent, text="0 columnas seleccionadas",
                                      font=("Segoe UI", 9), bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_contador.pack(anchor="e", pady=(4, 0))

    def _build_panel_rangos(self, parent):
        # Mensaje de estado de rangos configurados
        self.lbl_rangos_estado = tk.Label(
            parent,
            text="Sin rangos configurados — usa el menú  Rangos → Definir rangos de celdas…",
            font=("Segoe UI", 10, "italic"),
            bg=BG_CARD, fg=TEXT_MUTED,
            wraplength=540, justify="left"
        )
        self.lbl_rangos_estado.pack(anchor="w", pady=(8, 12))

        HoverButton(parent, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="⊞  Definir rangos de celdas…",
                    font=("Segoe UI", 10, "bold"),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self.controlador._abrir_selector_rangos).pack(anchor="w")

        # Lista de rangos configurados
        self.frame_rangos_lista = tk.Frame(parent, bg=BG_CARD)
        self.frame_rangos_lista.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

    def _card_exportar(self):
        body = self._make_card("Exportar resultado", "📊")

        row = tk.Frame(body, bg=BG_CARD)
        row.pack(fill=tk.X)

        self.var_reporte = tk.BooleanVar(value=True)
        tk.Checkbutton(row,
                        text="Incluir columna 'REPORTE FINAL' con resumen por fila",
                        variable=self.var_reporte,
                        font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_DARK,
                        activebackground=BG_CARD, selectcolor=BG_CARD,
                        cursor="hand2").pack(side=tk.LEFT)

        row2 = tk.Frame(body, bg=BG_CARD)
        row2.pack(fill=tk.X, pady=(10, 0))

        self.btn_exportar = HoverButton(
            row2, bg_normal="#1E8449", bg_hover="#176138",
            text="⬇  Exportar a Excel",
            font=("Segoe UI", 11, "bold"), fg=TEXT_LIGHT, padx=20, pady=9,
            command=self.controlador.exportar_datos, state=tk.DISABLED)
        self.btn_exportar.pack(side=tk.LEFT)

        self.btn_renombrar = HoverButton(
            row2, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
            text="✏  Renombrar columnas…",
            font=("Segoe UI", 10), fg=ACCENT, padx=14, pady=9,
            command=self.controlador._abrir_renombrar, state=tk.DISABLED)
        self.btn_renombrar.pack(side=tk.LEFT, padx=(10, 0))

        self.lbl_filas = tk.Label(row2, text="", font=("Segoe UI", 9),
                                   bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_filas.pack(side=tk.LEFT, padx=(16, 0))