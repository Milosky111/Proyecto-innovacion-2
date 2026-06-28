# components.py
import tkinter as tk
from tkinter import ttk
# Importamos las constantes necesarias desde config
from config import BG_MAIN, ACCENT, ACCENT_HOVER, TEXT_LIGHT, TEXT_DARK, TEXT_MUTED

class HoverButton(tk.Button):
    def __init__(self, master, bg_normal, bg_hover, **kwargs):
        super().__init__(master, bg=bg_normal, activebackground=bg_hover,
                         relief="flat", cursor="hand2", **kwargs)
        self._bg_normal = bg_normal
        self._bg_hover  = bg_hover
        self.bind("<Enter>", lambda e: self.config(bg=bg_hover))
        self.bind("<Leave>", lambda e: self.config(bg=bg_normal))


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
                 font=("Segoe UI", 13, "bold"),
                 bg=BG_MAIN, fg=TEXT_DARK).pack(pady=(18, 2))
        tk.Label(self,
                 text="Deja en blanco para mantener el nombre original.",
                 font=("Segoe UI", 9, "italic"),
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
                 font=("Segoe UI", 9, "bold"), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT)
        tk.Label(hdr, text="→  Nuevo nombre", anchor="w",
                 font=("Segoe UI", 9, "bold"), bg=BG_MAIN, fg=TEXT_MUTED).pack(side=tk.LEFT, padx=(8,0))

        for col in columnas:
            row = tk.Frame(self.scroll_frame, bg=BG_MAIN)
            row.pack(fill=tk.X, pady=3)

            tk.Label(row, text=col[:24], width=22, anchor="w",
                     font=("Segoe UI", 9), bg=BG_MAIN, fg=TEXT_DARK).pack(side=tk.LEFT)

            entry = tk.Entry(row, font=("Segoe UI", 9), width=28,
                             relief="solid", bd=1,
                             highlightthickness=1,
                             highlightcolor=ACCENT)
            entry.pack(side=tk.LEFT, padx=(8, 0))
            self.entries[col] = entry

        # Botones
        btn_row = tk.Frame(self, bg=BG_MAIN)
        btn_row.pack(pady=14)

        HoverButton(btn_row, bg_normal=ACCENT, bg_hover=ACCENT_HOVER,
                    text="Aplicar y exportar",
                    font=("Segoe UI", 10, "bold"),
                    fg=TEXT_LIGHT, padx=16, pady=7,
                    command=self._aplicar).pack(side=tk.LEFT, padx=(0, 8))

        HoverButton(btn_row, bg_normal="#E8E8E8", bg_hover="#D0D0D0",
                    text="Cancelar",
                    font=("Segoe UI", 10),
                    fg=TEXT_DARK, padx=16, pady=7,
                    command=self.destroy).pack(side=tk.LEFT)

    def _aplicar(self):
        self.result = {
            col: (entry.get().strip() or col)
            for col, entry in self.entries.items()
        }
        self.destroy()