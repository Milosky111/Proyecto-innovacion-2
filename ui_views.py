# ui_views.py
import tkinter as tk
from tkinter import ttk

from config import (
    BG_MAIN, BG_SIDEBAR, BG_CARD, ACCENT, ACCENT_HOVER,
    ERROR, TEXT_LIGHT, TEXT_DARK, TEXT_MUTED, BORDER, ROW_EVEN, ROW_ODD, F
)
from components import HoverButton, Tooltip, AyudaInline, TabButton

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
        tk.Label(self.sidebar, text="📊", font=F(28),
                 bg=BG_SIDEBAR, fg=TEXT_LIGHT).pack(pady=(28, 4))
        tk.Label(self.sidebar, text="Extractor\nde Datos",
                 font=F(14, "bold"), bg=BG_SIDEBAR,
                 fg=TEXT_LIGHT, justify="center").pack()

        ttk.Separator(self.sidebar, orient="horizontal").pack(fill=tk.X, padx=20, pady=20)

        self.step_labels = []
        pasos = [
            ("1", "Cargar archivo",
             "Elige el archivo Excel (.xlsx o .xls) desde tu computador."),
            ("2", "Elegir hoja",
             "Selecciona qué hoja del archivo quieres extraer, por si tiene varias."),
            ("3", "Seleccionar\ncolumnas",
             "Marca qué columnas de esa hoja quieres incluir en el resultado, "
             "o define un rango exacto de celdas."),
            ("4", "Exportar",
             "Genera el archivo Excel final con solo lo que seleccionaste."),
        ]

        for num, texto, ayuda in pasos:
            f = tk.Frame(self.sidebar, bg=BG_SIDEBAR)
            f.pack(fill=tk.X, padx=18, pady=4)

            badge = tk.Label(f, text=num, width=2, font=F(10, "bold"),
                             bg=ACCENT, fg=TEXT_LIGHT, padx=4, pady=2)
            badge.pack(side=tk.LEFT, padx=(0, 10))

            lbl = tk.Label(f, text=texto, font=F(10),
                           bg=BG_SIDEBAR, fg="#A8C4E0", justify="left", anchor="w")
            lbl.pack(side=tk.LEFT, fill=tk.X)
            self.step_labels.append((badge, lbl))

            # El paso completo (número + texto) explica su función al pasar
            # el mouse — así un usuario nuevo entiende el flujo sin salir
            # de la ventana principal.
            Tooltip(f, ayuda, wraplength=190)
            Tooltip(badge, ayuda, wraplength=190)
            Tooltip(lbl, ayuda, wraplength=190)

        tk.Frame(self.sidebar, bg=BG_SIDEBAR).pack(fill=tk.BOTH, expand=True)

        self.lbl_estado_dot = tk.Label(self.sidebar, text="●", font=F(10), bg=BG_SIDEBAR, fg="#A8C4E0")
        self.lbl_estado_dot.pack(pady=(0, 2))

        self.lbl_estado = tk.Label(self.sidebar, text="Esperando archivo…", font=F(9), 
                                   bg=BG_SIDEBAR, fg="#A8C4E0", wraplength=180, justify="center")
        self.lbl_estado.pack(pady=(0, 20))

    def _build_panel(self):
        header = tk.Frame(self.panel, bg=BG_MAIN)
        header.pack(fill=tk.X, padx=20, pady=(20, 16))

        tk.Label(header, text="Extractor de Datos Excel", font=F(18, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(anchor="w")
        tk.Label(header,
                 text="Carga cualquier Excel, elige las columnas y exporta. "
                      "Sigue los pasos 1 a 4 del menú de la izquierda — "
                      "cada uno se activa cuando el anterior está listo.",
                 font=F(10), bg=BG_MAIN, fg=TEXT_MUTED,
                 wraplength=640, justify="left").pack(anchor="w")

        self._card_archivo()
        self._card_hoja()
        self._card_columnas()
        self._card_exportar()

    def _make_card(self, title, emoji="", ayuda=None):
        outer = tk.Frame(self.panel, bg=BG_MAIN)
        outer.pack(fill=tk.X, padx=20, pady=(0, 12))

        card = tk.Frame(outer, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        tk.Frame(card, bg=ACCENT, height=4).pack(fill=tk.X)

        header = tk.Frame(card, bg=BG_CARD)
        header.pack(fill=tk.X, padx=16, pady=(10, 8))
        tk.Label(header, text=f"{emoji}  {title}", font=F(10, "bold"), bg=BG_CARD, fg=TEXT_DARK).pack(side=tk.LEFT)
        if ayuda:
            ico = AyudaInline(header, ayuda)
            ico.pack(side=tk.LEFT, padx=(8, 0))

        body = tk.Frame(card, bg=BG_CARD)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 14))
        return body

    def _card_archivo(self):
        body = self._make_card(
            "Cargar archivo Excel", "📂",
            ayuda="Este es el punto de partida: elige el archivo Excel "
                  "(.xlsx o .xls) del que quieres extraer datos. Los "
                  "siguientes pasos se activan solo después de este."
        )
        row = tk.Frame(body, bg=BG_CARD)
        row.pack(fill=tk.X)

        btn_buscar = HoverButton(
            row, bg_normal=ACCENT, bg_hover=ACCENT_HOVER, text="Buscar archivo…",
            font=F(10, "bold"), fg=TEXT_LIGHT, padx=16, pady=7,
            command=self.controlador.cargar_archivo)
        btn_buscar.pack(side=tk.LEFT)
        Tooltip(btn_buscar,
                "Abre el explorador de archivos de tu computador para "
                "elegir el Excel que quieres procesar.")

        self.lbl_archivo = tk.Label(row, text="Ningún archivo seleccionado", font=F(10, "italic"),
                                    bg=BG_CARD, fg=TEXT_MUTED, anchor="w")
        self.lbl_archivo.pack(side=tk.LEFT, padx=(14, 0), fill=tk.X, expand=True)

    def _card_hoja(self):
        body = self._make_card(
            "Seleccionar hoja de trabajo", "📄",
            ayuda="Un archivo Excel puede tener varias hojas (pestañas). "
                  "Aquí eliges cuál de ellas contiene los datos que "
                  "quieres extraer. Al elegirla se cargan sus columnas abajo."
        )
        row = tk.Frame(body, bg=BG_CARD)
        row.pack(fill=tk.X)

        tk.Label(row, text="Hoja:", font=F(10), bg=BG_CARD, fg=TEXT_DARK).pack(side=tk.LEFT, padx=(0, 8))

        self.combo_hojas = ttk.Combobox(row, state="disabled", font=F(10), width=36)
        self.combo_hojas.pack(side=tk.LEFT)
        self.combo_hojas.bind("<<ComboboxSelected>>", self.controlador.cargar_columnas)
        Tooltip(self.combo_hojas,
                "Lista de hojas encontradas dentro del archivo cargado. "
                "Se activa apenas cargas un archivo en el paso 1.")

        self.lbl_info_hoja = tk.Label(row, text="", font=F(9), bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_info_hoja.pack(side=tk.LEFT, padx=(12, 0))

    def _card_columnas(self):
        outer = tk.Frame(self.panel, bg=BG_MAIN)
        outer.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 12))

        card = tk.Frame(outer, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        tk.Frame(card, bg=ACCENT, height=4).pack(fill=tk.X)

        # ── Encabezado de la tarjeta con ayuda ──────────────────────────────
        header_paso3 = tk.Frame(card, bg=BG_CARD)
        header_paso3.pack(fill=tk.X, padx=16, pady=(10, 0))
        tk.Label(header_paso3, text="🧩  Elegir qué extraer", font=F(10, "bold"),
                  bg=BG_CARD, fg=TEXT_DARK).pack(side=tk.LEFT)
        AyudaInline(
            header_paso3,
            "Hay dos formas de elegir qué datos extraer:\n\n"
            "• Columnas completas: marca columnas enteras de la hoja "
            "(toma todas sus filas).\n\n"
            "• Rangos de celdas: define un área exacta (ej. B3:H30) "
            "por hoja, útil cuando los datos no empiezan en la fila 1 "
            "o hay varias tablas en la misma hoja.",
            font=F(8, "bold")
        ).pack(side=tk.LEFT, padx=(8, 0))

        # ── Tabs: modo columnas / modo rangos ─────────────────────────────
        tab_bar = tk.Frame(card, bg=BG_CARD)
        tab_bar.pack(fill=tk.X, padx=16, pady=(8, 0))

        self.var_modo = tk.StringVar(value="columnas")

        self.tab_col = TabButton(
            tab_bar, text="🗂  Columnas completas",
            command=lambda: self._seleccionar_modo("columnas"),
            bg_normal=BG_CARD, fg_normal=TEXT_DARK,
        )
        self.tab_col.pack(side=tk.LEFT, padx=(0, 4))
        Tooltip(self.tab_col,
                "Modo simple: elige columnas completas de la hoja "
                "(se incluyen todas sus filas de datos).")

        self.tab_rango = TabButton(
            tab_bar, text="⊞  Rangos de celdas",
            command=lambda: self._seleccionar_modo("rangos"),
            bg_normal=BG_CARD, fg_normal=TEXT_DARK,
        )
        self.tab_rango.pack(side=tk.LEFT)
        Tooltip(self.tab_rango,
                "Modo avanzado: define un rango exacto de celdas por "
                "hoja (ej. B3:H30). Útil si la tabla no parte en la "
                "fila 1 o si quieres combinar varias hojas en un mismo "
                "archivo de salida.")

        self.tab_col.set_selected(True)

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

    def _seleccionar_modo(self, modo):
        """
        Reemplaza lo que antes hacía tk.Radiobutton automáticamente
        (fijar la variable y redibujar el widget seleccionado): al ser
        ahora dos TabButton independientes, esta función actualiza el
        estado y el resaltado visual de ambas pestañas manualmente.
        """
        self.var_modo.set(modo)
        self.tab_col.set_selected(modo == "columnas")
        self.tab_rango.set_selected(modo == "rangos")
        self.controlador._cambiar_modo()

    def _build_panel_columnas(self, parent):
        header = tk.Frame(parent, bg=BG_CARD)
        header.pack(fill=tk.X)

        tk.Label(header, text="Ctrl+clic o Shift+clic para seleccionar varias",
                 font=F(8, "italic"),
                 bg=BG_CARD, fg=TEXT_MUTED).pack(side=tk.LEFT)

        btn_frame = tk.Frame(header, bg=BG_CARD)
        btn_frame.pack(side=tk.RIGHT)

        btn_todo = HoverButton(btn_frame, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
                    text="Seleccionar todo", font=F(8),
                    fg=ACCENT, padx=8, pady=4,
                    command=self.controlador._seleccionar_todo)
        btn_todo.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn_todo, "Marca todas las columnas de la lista de una vez.")

        btn_limpiar = HoverButton(btn_frame, bg_normal="#FEE8E8", bg_hover="#FCC5C5",
                    text="Limpiar", font=F(8),
                    fg=ERROR, padx=8, pady=4,
                    command=self.controlador._limpiar_seleccion)
        btn_limpiar.pack(side=tk.LEFT)
        Tooltip(btn_limpiar, "Quita la selección actual, sin borrar la lista de columnas.")

        tree_frame = tk.Frame(parent, bg=BG_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        style = ttk.Style()
        style.configure("Col.Treeview", font=F(10), rowheight=28,
                         background=BG_CARD, fieldbackground=BG_CARD, borderwidth=0)
        style.configure("Col.Treeview.Heading", font=F(10, "bold"),
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
        Tooltip(self.lista_columnas,
                "Clic para elegir una columna. Mantén Ctrl (o ⌘ en Mac) "
                "y haz clic para agregar más de una; usa Shift para "
                "elegir un rango seguido.")

        sb = ttk.Scrollbar(tree_frame, orient="vertical",
                             command=self.lista_columnas.yview)
        self.lista_columnas.configure(yscrollcommand=sb.set)
        self.lista_columnas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lista_columnas.bind("<<TreeviewSelect>>", self.controlador._actualizar_contador)

        self.lbl_contador = tk.Label(parent, text="0 columnas seleccionadas",
                                      font=F(9), bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_contador.pack(anchor="e", pady=(4, 0))

    def _build_panel_rangos(self, parent):
        # Mensaje de estado de rangos configurados
        self.lbl_rangos_estado = tk.Label(
            parent,
            text="Sin rangos configurados — usa el menú  Rangos → Definir rangos de celdas…",
            font=F(10, "italic"),
            bg=BG_CARD, fg=TEXT_MUTED,
            wraplength=540, justify="left"
        )
        self.lbl_rangos_estado.pack(anchor="w", pady=(8, 12))

        btn_def_rangos = HoverButton(parent, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="⊞  Definir rangos de celdas…",
                    font=F(10, "bold"),
                    fg=TEXT_LIGHT, padx=14, pady=7,
                    command=self.controlador._abrir_selector_rangos)
        btn_def_rangos.pack(anchor="w")
        Tooltip(btn_def_rangos,
                "Abre una ventana donde, hoja por hoja, escribes o marcas "
                "con el mouse el área exacta de celdas a extraer, y puedes "
                "renombrar sus columnas antes de exportar.")

        # Lista de rangos configurados
        self.frame_rangos_lista = tk.Frame(parent, bg=BG_CARD)
        self.frame_rangos_lista.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

    def _card_exportar(self):
        body = self._make_card(
            "Exportar resultado", "📊",
            ayuda="Último paso: genera el archivo Excel final con lo que "
                  "elegiste arriba. Te pedirá dónde guardarlo y con qué "
                  "nombre."
        )

        row = tk.Frame(body, bg=BG_CARD)
        row.pack(fill=tk.X)

        self.var_reporte = tk.BooleanVar(value=True)
        chk_reporte = ttk.Checkbutton(
            row, text="Incluir columna 'REPORTE FINAL' con resumen por fila",
            variable=self.var_reporte, style="Card.TCheckbutton", cursor="hand2"
        )
        chk_reporte.pack(side=tk.LEFT)
        Tooltip(chk_reporte,
                "Si está marcada, se agrega al final una columna extra "
                "con un resumen automático de cada fila.")

        row2 = tk.Frame(body, bg=BG_CARD)
        row2.pack(fill=tk.X, pady=(10, 0))

        self.btn_exportar = HoverButton(
            row2, bg_normal="#1E8449", bg_hover="#176138",
            text="⬇  Exportar a Excel",
            font=F(11, "bold"), fg=TEXT_LIGHT, padx=20, pady=9,
            command=self.controlador.exportar_datos, state=tk.DISABLED)
        self.btn_exportar.pack(side=tk.LEFT)
        Tooltip(self.btn_exportar,
                "Crea el archivo Excel final con las columnas o rangos "
                "que seleccionaste, y te pregunta dónde guardarlo. Se "
                "activa cuando hay al menos una columna o rango elegido.")

        self.btn_renombrar = HoverButton(
            row2, bg_normal="#E8F0FE", bg_hover="#C5D8FB",
            text="✏  Renombrar columnas…",
            font=F(10), fg=ACCENT, padx=14, pady=9,
            command=self.controlador._abrir_renombrar, state=tk.DISABLED)
        self.btn_renombrar.pack(side=tk.LEFT, padx=(10, 0))
        Tooltip(self.btn_renombrar,
                "Opcional: abre una ventana para ponerle un nombre más "
                "claro a cada columna seleccionada antes de exportar.")

        self.lbl_filas = tk.Label(row2, text="", font=F(9),
                                   bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_filas.pack(side=tk.LEFT, padx=(16, 0))