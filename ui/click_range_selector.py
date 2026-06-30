# ui/click_range_selector.py
"""
Selector visual de rango por clic en celdas — modo alternativo al campo de
texto ("B3:R35") del SelectorRangos. Pinta una grilla con la vista previa
real de la hoja; el usuario hace clic en la celda inicial y luego en la
celda final para marcar el rango. Al confirmar, retorna el string de rango
en self.result, en el mismo formato que ya usa el modo texto.
"""

import tkinter as tk
from tkinter import messagebox

from config import BG_MAIN, BG_CARD, ACCENT, ACCENT_HOVER, SUCCESS, ERROR, TEXT_LIGHT, TEXT_DARK, TEXT_MUTED, BORDER
from components import HoverButton


class ClickRangeSelector(tk.Toplevel):
    """
    Uso:
        sel = ClickRangeSelector(parent, filas_preview)
        parent.wait_window(sel)
        if sel.result:
            ...usar sel.result como rango, ej "B5:H100"...

    filas_preview: lista de listas de strings (vista previa cruda de la hoja).
    El ancho de cada columna se calcula dinámicamente según su contenido más
    largo, como en Excel — evita que columnas con texto corto desperdicien
    espacio y columnas con texto largo se vean cortadas a la mitad.
    """

    CELDA_ANCHO_MIN = 60
    CELDA_ANCHO_MAX = 170
    CELDA_ANCHO_PADDING = 16     # espacio extra alrededor del texto
    PX_POR_CARACTER = 6.5        # estimación para fuente Segoe UI 8pt
    CELDA_ALTO = 24
    COL_HEADER_ANCHO = 38
    TEXT_DATA_COLOR = "#1A1A1A"  # negro neutro, sin tinte azul (a diferencia
                                 # de TEXT_DARK), para que el texto de las
                                 # celdas se lea parejo en toda la grilla
    ROW_EVEN = "#F4F7FB"         # zebra striping, tono sutil

    def __init__(self, parent, filas_preview):
        super().__init__(parent)
        self.title("Marcar rango con clic")
        self.geometry("920x560")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.grab_set()
        self.lift()
        self.focus_force()

        self.result = None
        self.filas_preview = filas_preview or []
        self.n_filas = len(self.filas_preview)
        self.n_cols = max((len(f) for f in self.filas_preview), default=0)
        self._anchos_col = self._calcular_anchos_columna()

        self._celda_inicio = None
        self._celda_fin = None
        self._rect_seleccion = None

        # ── Cabecera ────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=BG_MAIN)
        header.pack(fill=tk.X, padx=20, pady=(16, 8))
        tk.Label(header, text="Marca el rango de datos",
                 font=("Segoe UI", 13, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(anchor="w")
        tk.Label(header,
                 text="Clic en la celda donde empieza la tabla, luego clic en la celda donde "
                      "termina. Ignora títulos o textos decorativos — marca solo la tabla real.",
                 font=("Segoe UI", 9), bg=BG_MAIN, fg=TEXT_MUTED,
                 wraplength=820, justify="left").pack(anchor="w", pady=(2, 0))

        self.lbl_rango_actual = tk.Label(
            header, text="Rango: (sin seleccionar)",
            font=("Segoe UI", 11, "bold"), bg=BG_MAIN, fg=SUCCESS
        )
        self.lbl_rango_actual.pack(anchor="w", pady=(8, 0))

        if self.n_filas == 0 or self.n_cols == 0:
            tk.Label(
                self, text="⚠ No se encontraron datos en esta hoja para mostrar.",
                font=("Segoe UI", 11, "bold"), bg=BG_MAIN, fg=ERROR
            ).pack(pady=40)
            self.canvas = None
        else:
            self._construir_grilla()

        # ── Botones ─────────────────────────────────────────────────────────
        btn_row = tk.Frame(self, bg=BG_MAIN)
        btn_row.pack(pady=(0, 16))

        HoverButton(btn_row, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="Usar este rango", font=("Segoe UI", 10, "bold"),
                    fg=TEXT_LIGHT, padx=18, pady=8,
                    command=self._confirmar).pack(side=tk.LEFT, padx=(0, 10))

        HoverButton(btn_row, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="Cancelar", font=("Segoe UI", 10),
                    fg=TEXT_DARK, padx=14, pady=8,
                    command=self.destroy).pack(side=tk.LEFT)

    # ── Construcción de la grilla ─────────────────────────────────────────────

    def _calcular_anchos_columna(self):
        """
        Calcula el ancho de cada columna según su contenido más largo,
        acotado entre CELDA_ANCHO_MIN y CELDA_ANCHO_MAX. Las columnas con
        texto largo quedan más anchas (sin truncar tanto); las columnas
        cortas o vacías quedan compactas, sin desperdiciar espacio.
        """
        anchos = [self.CELDA_ANCHO_MIN] * self.n_cols
        for fila in self.filas_preview:
            for col, texto in enumerate(fila):
                if col >= self.n_cols or not texto:
                    continue
                ancho_estimado = int(len(texto) * self.PX_POR_CARACTER) + self.CELDA_ANCHO_PADDING
                ancho_estimado = max(self.CELDA_ANCHO_MIN, min(ancho_estimado, self.CELDA_ANCHO_MAX))
                if ancho_estimado > anchos[col]:
                    anchos[col] = ancho_estimado
        return anchos

    def _offset_x(self, col):
        """Posición X acumulada donde empieza la columna `col` (0-based)."""
        return self.COL_HEADER_ANCHO + sum(self._anchos_col[:col])

    def _col_en_x(self, x):
        """Dado un X de canvas, retorna el índice de columna bajo ese punto, o None."""
        acumulado = self.COL_HEADER_ANCHO
        for col, ancho in enumerate(self._anchos_col):
            if acumulado <= x < acumulado + ancho:
                return col
            acumulado += ancho
        return None

    def _construir_grilla(self):
        container = tk.Frame(self, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        ancho_total = self.COL_HEADER_ANCHO + sum(self._anchos_col)
        alto_total = self.CELDA_ALTO + self.n_filas * self.CELDA_ALTO
        ancho_visible = min(ancho_total, 860)
        alto_visible = min(alto_total, 400)

        h_scroll = tk.Scrollbar(container, orient="horizontal")
        v_scroll = tk.Scrollbar(container, orient="vertical")

        self.canvas = tk.Canvas(
            container, bg=BG_MAIN, highlightthickness=0,
            width=ancho_visible, height=alto_visible,
            scrollregion=(0, 0, ancho_total, alto_total),
            xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set
        )

        h_scroll.config(command=self.canvas.xview)
        v_scroll.config(command=self.canvas.yview)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._dibujar_grilla()
        self.canvas.bind("<Button-1>", self._on_click_seguro, add="+")

    def _col_letra(self, idx0):
        idx = idx0
        letras = ""
        while True:
            idx, resto = divmod(idx, 26)
            letras = chr(65 + resto) + letras
            if idx == 0:
                break
            idx -= 1
        return letras

    def _dibujar_grilla(self):
        c = self.canvas
        ancho_total = self.COL_HEADER_ANCHO + sum(self._anchos_col)

        # Encabezados de columna (A, B, C…)
        for col in range(self.n_cols):
            x0 = self._offset_x(col)
            ancho = self._anchos_col[col]
            c.create_rectangle(x0, 0, x0 + ancho, self.CELDA_ALTO,
                                fill="#1F3864", outline="#16294A")
            c.create_text(x0 + ancho / 2, self.CELDA_ALTO / 2,
                           text=self._col_letra(col), fill="#FFFFFF",
                           font=("Segoe UI", 9, "bold"))

        # Esquina superior izquierda (sobre los encabezados de fila)
        c.create_rectangle(0, 0, self.COL_HEADER_ANCHO, self.CELDA_ALTO,
                            fill="#1F3864", outline="#16294A")

        for fila in range(self.n_filas):
            y0 = self.CELDA_ALTO + fila * self.CELDA_ALTO
            es_par = fila % 2 == 1
            color_fondo_fila = self.ROW_EVEN if es_par else "#FFFFFF"

            # Encabezado de fila (1, 2, 3…)
            c.create_rectangle(0, y0, self.COL_HEADER_ANCHO, y0 + self.CELDA_ALTO,
                                fill="#1F3864", outline="#16294A")
            c.create_text(self.COL_HEADER_ANCHO / 2, y0 + self.CELDA_ALTO / 2,
                           text=str(fila + 1), fill="#FFFFFF",
                           font=("Segoe UI", 9, "bold"))

            valores_fila = self.filas_preview[fila]
            for col in range(self.n_cols):
                x0 = self._offset_x(col)
                ancho = self._anchos_col[col]
                texto = valores_fila[col] if col < len(valores_fila) else ""

                c.create_rectangle(
                    x0, y0, x0 + ancho, y0 + self.CELDA_ALTO,
                    fill=color_fondo_fila, outline=BORDER,
                    tags=f"celda_{fila}_{col}"
                )
                if texto:
                    # Recortar visualmente solo si de verdad no cabe en el
                    # ancho ya calculado para esa columna (caso límite, ya
                    # que _calcular_anchos_columna deja espacio suficiente
                    # para la mayoría de los casos).
                    max_chars = max(int((ancho - 10) / self.PX_POR_CARACTER), 3)
                    texto_mostrado = texto if len(texto) <= max_chars else texto[:max_chars - 1] + "…"
                    c.create_text(
                        x0 + 6, y0 + self.CELDA_ALTO / 2, text=texto_mostrado,
                        fill=self.TEXT_DATA_COLOR, font=("Segoe UI", 8), anchor="w"
                    )

        # Línea de cierre a la derecha y abajo, para que la grilla no se
        # vea "abierta" cuando el contenido es menor al área visible.
        alto_total = self.CELDA_ALTO + self.n_filas * self.CELDA_ALTO
        c.create_line(0, 0, 0, alto_total, fill=BORDER)
        c.create_line(ancho_total, 0, ancho_total, alto_total, fill=BORDER)

    # ── Interacción ─────────────────────────────────────────────────────────

    def _coords_a_celda(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        if x < self.COL_HEADER_ANCHO or y < self.CELDA_ALTO:
            return None
        col = self._col_en_x(x)
        fila = int((y - self.CELDA_ALTO) // self.CELDA_ALTO)
        if col is not None and 0 <= fila < self.n_filas:
            return (fila, col)
        return None

    def _on_click_seguro(self, event):
        try:
            self._on_click(event)
        except Exception as e:
            messagebox.showerror(
                "Error en el selector de rango",
                f"Ocurrió un problema al registrar el clic:\n{type(e).__name__}: {e}",
                parent=self
            )

    def _on_click(self, event):
        celda = self._coords_a_celda(event)
        if celda is None:
            return

        if self._celda_inicio is None or self._celda_fin is not None:
            self._celda_inicio = celda
            self._celda_fin = None
        else:
            self._celda_fin = celda

        self._repintar_seleccion()

    def _repintar_seleccion(self):
        if self._rect_seleccion is not None:
            self.canvas.delete(self._rect_seleccion)
            self._rect_seleccion = None

        if self._celda_inicio is None:
            self.lbl_rango_actual.config(text="Rango: (sin seleccionar)")
            return

        f_ini, c_ini = self._celda_inicio
        if self._celda_fin is None:
            x0 = self._offset_x(c_ini)
            y0 = self.CELDA_ALTO + f_ini * self.CELDA_ALTO
            x1 = x0 + self._anchos_col[c_ini]
            y1 = y0 + self.CELDA_ALTO
            self._rect_seleccion = self.canvas.create_rectangle(
                x0, y0, x1, y1, outline=ACCENT, width=3
            )
            ref = f"{self._col_letra(c_ini)}{f_ini + 1}"
            self.lbl_rango_actual.config(text=f"Rango: {ref}  →  (clic en celda final)")
            return

        f_fin, c_fin = self._celda_fin
        f_top, f_bot = sorted((f_ini, f_fin))
        c_left, c_right = sorted((c_ini, c_fin))

        x0 = self._offset_x(c_left)
        y0 = self.CELDA_ALTO + f_top * self.CELDA_ALTO
        x1 = self._offset_x(c_right) + self._anchos_col[c_right]
        y1 = self.CELDA_ALTO + (f_bot + 1) * self.CELDA_ALTO

        self._rect_seleccion = self.canvas.create_rectangle(
            x0, y0, x1, y1, outline=ACCENT, width=3
        )

        ref = f"{self._col_letra(c_left)}{f_top + 1}:{self._col_letra(c_right)}{f_bot + 1}"
        self.lbl_rango_actual.config(text=f"Rango: {ref}")

    def _confirmar(self):
        if self._celda_inicio is None or self._celda_fin is None:
            messagebox.showwarning(
                "Rango incompleto",
                "Haz clic en la celda inicial y luego en la celda final antes de confirmar.",
                parent=self
            )
            return
        f_ini, c_ini = self._celda_inicio
        f_fin, c_fin = self._celda_fin
        f_top, f_bot = sorted((f_ini, f_fin))
        c_left, c_right = sorted((c_ini, c_fin))
        self.result = f"{self._col_letra(c_left)}{f_top + 1}:{self._col_letra(c_right)}{f_bot + 1}"
        self.destroy()
