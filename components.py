# components.py
import tkinter as tk
from tkinter import ttk
# Importamos las constantes necesarias desde config
from config import (BG_MAIN, BG_CARD, ACCENT, ACCENT_HOVER, TEXT_LIGHT,
                     TEXT_DARK, TEXT_MUTED, BORDER, F)


class HoverButton(ttk.Button):
    """
    Botón de color sólido con efecto hover, con la MISMA apariencia en
    Windows y macOS.

    Antes este componente heredaba de tk.Button (el botón clásico de Tk).
    En macOS, tk.Button se dibuja con el tema nativo "Aqua", que IGNORA
    los colores de fondo/texto que se le asignen: sin importar qué bg/fg
    se configure, siempre se ve como el botón gris del sistema. Por eso
    en Mac aparecían botones con texto blanco sobre fondo gris claro,
    prácticamente ilegible (como "Buscar archivo…").

    ttk.Button, en cambio, sí permite pintarse a medida si se usa un tema
    "no nativo" como "clam" (activado una sola vez en aplicar_estilo_ttk,
    más abajo en este archivo). Por eso este componente ahora es un
    ttk.Button con un estilo propio por instancia, y el hover lo maneja
    ttk automáticamente a través del estado "active" — ya no hace falta
    enlazar manualmente los eventos <Enter>/<Leave>.
    """
    _contador = 0

    def __init__(self, master, bg_normal, bg_hover, text="", font=None,
                 fg=TEXT_LIGHT, padx=12, pady=6, command=None,
                 state=tk.NORMAL, **kwargs):
        HoverButton._contador += 1
        estilo = f"HB{HoverButton._contador}.TButton"

        style = ttk.Style()
        style.configure(
            estilo,
            background=bg_normal, foreground=fg,
            font=font or F(10), padding=(padx, pady),
            borderwidth=1, bordercolor=bg_hover, relief="solid"
        )
        style.map(
            estilo,
            background=[("disabled", "#E5E5E5"), ("active", bg_hover)],
            foreground=[("disabled", "#A3A3A3")],
        )

        super().__init__(
            master, text=text, style=estilo, command=command,
            state=state, cursor="hand2", **kwargs
        )


# ─────────────────────────────────────────────────────────────────────────────
# Tooltips: pequeño globo de ayuda que aparece al dejar el mouse quieto sobre
# un botón/campo. Se usan en toda la app para explicar, sin que el usuario
# tenga que preguntar, qué hace cada botón o para qué sirve cada casilla.
# ─────────────────────────────────────────────────────────────────────────────
class Tooltip:
    """
    Agrega un globo de ayuda a cualquier widget de Tk/ttk.

    Uso:
        Tooltip(boton_exportar, "Genera el archivo Excel final con las "
                                 "columnas que elegiste arriba.")
    """
    RETARDO_MS = 450

    def __init__(self, widget, texto, wraplength=260):
        self.widget = widget
        self.texto = texto
        self.wraplength = wraplength
        self._after_id = None
        self._tip = None
        widget.bind("<Enter>", self._agendar, add="+")
        widget.bind("<Leave>", self._ocultar, add="+")
        widget.bind("<ButtonPress>", self._ocultar, add="+")

    def _agendar(self, _event=None):
        self._cancelar()
        self._after_id = self.widget.after(self.RETARDO_MS, self._mostrar)

    def _cancelar(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _mostrar(self):
        if self._tip is not None or not self.widget.winfo_exists():
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8

        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        try:
            self._tip.attributes("-topmost", True)
        except tk.TclError:
            pass

        frame = tk.Frame(self._tip, bg="#1A1A2E", padx=1, pady=1)
        frame.pack()
        tk.Label(
            frame, text=self.texto, justify="left",
            bg="#2B2B3D", fg="#FFFFFF", font=F(9),
            wraplength=self.wraplength, padx=10, pady=6
        ).pack()

    def _ocultar(self, _event=None):
        self._cancelar()
        if self._tip is not None:
            self._tip.destroy()
            self._tip = None


class AyudaInline(tk.Label):
    """
    Círculo con "?" que se coloca junto a un título de sección y muestra un
    tooltip explicativo al pasar el mouse. Útil para no recargar de texto
    la pantalla y aun así dejar la explicación a un clic de distancia visual.
    """
    def __init__(self, master, texto, bg=None, font=None, **kwargs):
        # El "chip" siempre usa el mismo gris neutro (BORDER) como fondo,
        # independiente del fondo del contenedor (tarjeta blanca o sidebar
        # oscuro) — así se mantiene legible en ambos casos.
        super().__init__(
            master, text=" ? ", font=font or F(8, "bold"),
            bg=BORDER, fg=TEXT_DARK, cursor="hand2",
            padx=1, **kwargs
        )
        Tooltip(self, texto)


class TabButton(ttk.Button):
    """
    Botón usado como "pestaña" seleccionable (ej. Columnas completas /
    Rangos de celdas). Reemplaza tk.Radiobutton(indicatoron=False):
    en macOS ese widget también se dibuja con apariencia nativa y no
    respeta de forma confiable los colores de fondo asignados. Al ser
    un ttk.Button con estilo propio (tema "clam"), el color sí se
    respeta igual en Windows y en macOS.

    Uso:
        tab = TabButton(parent, text="…", command=..., bg_selected=ACCENT)
        tab.set_selected(True)   # marca visualmente como pestaña activa
    """
    _contador = 0

    def __init__(self, master, text, command=None, font=None,
                 fg_normal=TEXT_DARK, fg_selected=TEXT_LIGHT,
                 bg_normal=BG_CARD, bg_selected=ACCENT, **kwargs):
        TabButton._contador += 1
        self._base = f"Tab{TabButton._contador}"
        font = font or F(10, "bold")

        style = ttk.Style()
        style.configure(
            f"{self._base}.TButton",
            background=bg_normal, foreground=fg_normal,
            font=font, padding=(12, 5), borderwidth=1, relief="solid"
        )
        style.configure(
            f"{self._base}Sel.TButton",
            background=bg_selected, foreground=fg_selected,
            font=font, padding=(12, 5), borderwidth=1, relief="solid"
        )

        super().__init__(
            master, text=text, style=f"{self._base}.TButton",
            command=command, cursor="hand2", **kwargs
        )

    def set_selected(self, selected: bool):
        self.configure(style=f"{self._base}{'Sel' if selected else ''}.TButton")


def aplicar_estilo_ttk(root):
    """
    Unifica la apariencia de los widgets ttk (Combobox, Treeview, Scrollbar,
    Separator, etc.) en Windows y macOS.

    Por defecto, ttk dibuja estos widgets usando el "tema nativo" de cada
    sistema operativo (Aqua en macOS, Vista/XP en Windows), por lo que la
    misma ventana se ve distinta según el equipo: cambia el color de fondo,
    el borde de los combos, el estilo de las barras de desplazamiento, etc.

    "clam" es un tema de ttk que NO depende del sistema operativo: se dibuja
    igual en cualquier plataforma, lo que permite pintarlo con los mismos
    colores de esta app (config.py) y lograr una interfaz idéntica en Mac
    y en Windows. Se llama una sola vez, apenas se crea la ventana principal.
    """
    style = ttk.Style(root)
    style.theme_use("clam")

    # Combobox
    style.configure(
        "TCombobox",
        fieldbackground="#FFFFFF", background="#FFFFFF",
        foreground=TEXT_DARK, arrowcolor=ACCENT,
        bordercolor=BORDER, lightcolor="#FFFFFF", darkcolor="#FFFFFF",
        padding=4, font=F(10)
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", "#FFFFFF"), ("disabled", "#F0F0F0")],
        foreground=[("disabled", TEXT_MUTED)],
    )
    root.option_add("*TCombobox*Listbox.font", F(10))
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", TEXT_LIGHT)

    # Treeview (tablas)
    style.configure(
        "Treeview",
        background="#FFFFFF", fieldbackground="#FFFFFF",
        foreground=TEXT_DARK, rowheight=26,
        bordercolor=BORDER, borderwidth=1, font=F(10)
    )
    style.configure(
        "Treeview.Heading",
        background="#E8F0FE", foreground=TEXT_DARK,
        relief="flat", font=F(10, "bold")
    )
    style.map(
        "Treeview",
        background=[("selected", ACCENT)],
        foreground=[("selected", TEXT_LIGHT)],
    )
    style.map("Treeview.Heading", background=[("active", "#D8E6FC")])

    # Scrollbar
    style.configure(
        "Vertical.TScrollbar", background=BORDER, troughcolor=BG_MAIN,
        bordercolor=BG_MAIN, arrowcolor=TEXT_MUTED, gripcount=0
    )
    style.configure(
        "Horizontal.TScrollbar", background=BORDER, troughcolor=BG_MAIN,
        bordercolor=BG_MAIN, arrowcolor=TEXT_MUTED, gripcount=0
    )
    style.map("Vertical.TScrollbar", background=[("active", ACCENT)])
    style.map("Horizontal.TScrollbar", background=[("active", ACCENT)])

    # Separator
    style.configure("TSeparator", background=BORDER)

    # Checkbutton — se definen dos variantes porque el fondo del
    # checkbox debe calzar con el contenedor donde se use (tarjeta
    # blanca o fondo general de ventana); un tk.Checkbutton clásico
    # ignora esos colores en macOS y deja un recuadro blanco/gris
    # que no combina con el resto de la pantalla.
    style.configure(
        "Card.TCheckbutton", background=BG_CARD, foreground=TEXT_DARK,
        font=F(10)
    )
    style.map("Card.TCheckbutton", background=[("active", BG_CARD)])

    style.configure(
        "Main.TCheckbutton", background=BG_MAIN, foreground=TEXT_DARK,
        font=F(10)
    )
    style.map("Main.TCheckbutton", background=[("active", BG_MAIN)])

    return style


def crear_entry(parent, textvariable=None, font=None, width=20, show="", **kwargs):
    """
    tk.Entry con estilo consistente: fondo blanco, texto oscuro y un
    borde sutil que se resalta en azul al hacer clic dentro.

    A diferencia de los Label/Frame de esta app (que sí fijan su color
    con BG_MAIN/BG_CARD desde config.py), un tk.Entry sin bg/fg
    explícitos hereda la apariencia del sistema operativo. En macOS con
    Modo Oscuro activado eso significa fondo negro y texto blanco —el
    "cuadro negro" que rompe visualmente el resto de la ventana, clara
    y en colores propios de la app. Por eso todos los campos de texto
    de la app se crean con esta función, que fija el color siempre.
    """
    entry = tk.Entry(
        parent, textvariable=textvariable, font=font or F(10),
        width=width, show=show,
        bg="#FFFFFF", fg=TEXT_DARK, insertbackground=ACCENT,
        disabledbackground="#F0F0F0", readonlybackground="#F0F0F0",
        relief="solid", bd=1,
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        **kwargs
    )
    return entry


def crear_spinbox(parent, from_, to, textvariable=None, font=None, width=4, **kwargs):
    """Spinbox con el mismo criterio de color fijo que crear_entry (ver su docstring)."""
    spin = tk.Spinbox(
        parent, from_=from_, to=to, textvariable=textvariable,
        font=font or F(10), width=width,
        bg="#FFFFFF", fg=TEXT_DARK, insertbackground=ACCENT,
        buttonbackground="#FFFFFF",
        relief="solid", bd=1,
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        **kwargs
    )
    return spin


class RenameDialog(tk.Toplevel):
    """Ventana para renombrar columnas seleccionadas antes de exportar."""
    def __init__(self, parent, columnas):
        super().__init__(parent)
        self.title("Renombrar columnas (opcional)")
        self.geometry("500x420")
        self.configure(bg=BG_MAIN)
        self.resizable(False, True)
        self.grab_set()  # modal

        self.result = None
        self.entries = {}

        tk.Label(self, text="Renombrar columnas",
                 font=F(13, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(pady=(18, 2))
        tk.Label(self,
                 text="Deja en blanco para mantener el nombre original.",
                 font=F(9, "italic"),
                 bg=BG_MAIN, fg=TEXT_MUTED).pack(pady=(0, 12))

        # Frame scrollable
        container = tk.Frame(self, bg=BG_MAIN)
        container.pack(fill=tk.BOTH, expand=True, padx=20)

        canvas = tk.Canvas(container, bg=BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=BG_MAIN)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Encabezados
        hdr = tk.Frame(self.scroll_frame, bg=BG_MAIN)
        hdr.pack(fill=tk.X, pady=(0, 6))
        tk.Label(hdr, text="Nombre actual", width=22, anchor="w",
                 font=F(9, "bold"), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        tk.Label(hdr, text="→  Nuevo nombre", anchor="w",
                 font=F(9, "bold"), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT, padx=(8,0))

        for col in columnas:
            row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
            row.pack(fill=tk.X, pady=3)

            tk.Label(row, text=col[:24], width=22, anchor="w",
                     font=F(9), bg=BG_MAIN, fg=TEXT_DARK).pack(side=tk.LEFT)

            entry = crear_entry(row, font=F(9), width=28)
            entry.pack(side=tk.LEFT, padx=(8, 0))
            self.entries[col] = entry

        # Botones
        btn_row = tk.Frame(self, bg=BG_MAIN)
        btn_row.pack(pady=14)

        HoverButton(btn_row, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="Aplicar y exportar",
                    font=F(10, "bold"),
                    fg=TEXT_LIGHT, padx=16, pady=7,
                    command=self._aplicar).pack(side=tk.LEFT, padx=(0, 8))

        HoverButton(btn_row, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="Cancelar",
                    font=F(10),
                    fg=TEXT_DARK, padx=16, pady=7,
                    command=self.destroy).pack(side=tk.LEFT)

    def _aplicar(self):
        self.result = {
            col: (entry.get().strip() or col)
            for col, entry in self.entries.items()
        }
        self.destroy()