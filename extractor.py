# extractor.py  (punto de entrada)
"""
Aplicación principal.
Mantiene la funcionalidad de extracción manual existente
y agrega el menú de automatizaciones.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox

from config import BG_MAIN, ACCENT, TEXT_LIGHT, TEXT_DARK, TEXT_MUTED, SUCCESS
from components import HoverButton, RenameDialog
from core.excel_reader  import ExcelReader
from core.config_store  import ConfigStore
from core.logger        import RunLogger
from ui_views           import UIManager


class ExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Datos Excel")
        self.root.geometry("900x700")
        self.root.configure(bg=BG_MAIN)
        self.root.resizable(True, True)
        self.root.minsize(700, 560)

        # Core
        self.reader = ExcelReader()
        self.store  = ConfigStore()
        self.logger = RunLogger()

        # Estado de rangos (debe existir antes de construir la UI)
        self._rangos_config = {}

        # UI
        self.ui = UIManager(self.root, self)
        self._agregar_menu()

    # ── Menú ──────────────────────────────────────────────────────────────────

    def _agregar_menu(self):
        menubar = tk.Menu(self.root)

        m_rango = tk.Menu(menubar, tearoff=0)
        m_rango.add_command(label="Definir rangos de celdas…",
                             command=self._abrir_selector_rangos)
        m_rango.add_command(label="Ver rangos configurados",
                             command=self._ver_rangos)
        menubar.add_cascade(label="Rangos", menu=m_rango)

        m_auto = tk.Menu(menubar, tearoff=0)
        m_auto.add_command(label="Ver automatizaciones…",
                            command=self._abrir_panel_automatizaciones)
        m_auto.add_separator()
        m_auto.add_command(label="Nueva automatización…",
                            command=self._nueva_automatizacion)
        menubar.add_cascade(label="Automatizaciones", menu=m_auto)
        self.root.config(menu=menubar)

    def _abrir_selector_rangos(self):
        if not self.reader.ruta_archivo:
            messagebox.showwarning(
                "Sin archivo", "Carga un archivo Excel primero (Paso 1).")
            return
        from ui.selector_rangos import SelectorRangos
        prev = getattr(self, "_rangos_config", {})
        dlg = SelectorRangos(self.root, self.reader, rangos_previos=prev)
        self.root.wait_window(dlg)
        if dlg.resultado:
            self._rangos_config = dlg.resultado
            n = len(dlg.resultado)
            self._set_estado(
                f"{n} hoja{'s' if n != 1 else ''} con rango configurado", "ok")
            # Cambiar automáticamente a modo rangos y actualizar panel
            self.ui.var_modo.set("rangos")
            self._cambiar_modo()

    def _ver_rangos(self):
        cfg = getattr(self, "_rangos_config", {})
        if not cfg:
            messagebox.showinfo("Sin rangos", "No hay rangos configurados aún.")
            return
        lineas = []
        for hoja, datos in cfg.items():
            ren = len(datos.get("rename_map", {}))
            lineas.append(
                f"• {hoja}: {datos['rango']}"
                + (f"  ({ren} cols renombradas)" if ren else ""))
        messagebox.showinfo("Rangos configurados", "\n".join(lineas))

    def _abrir_panel_automatizaciones(self):
        from ui.panel_automatizaciones import PanelAutomatizaciones
        PanelAutomatizaciones(
            self.root, self.store, self.logger,
            on_editar_callback=self._abrir_form_perfil
        )

    def _nueva_automatizacion(self):
        perfil = self.store.nuevo_perfil()
        self._abrir_form_perfil(perfil["id"])

    def _abrir_form_perfil(self, perfil_id: str):
        from ui.form_perfil import FormPerfil
        FormPerfil(self.root, self.store, perfil_id)

    # ── Estado y pasos ────────────────────────────────────────────────────────

    def _set_estado(self, texto, tipo="info"):
        colores = {"info": "#A8C4E0", "ok": "#5DCA8A",
                   "error": "#FF7B7B", "warn": "#F0C040"}
        self.ui.lbl_estado.config(text=texto)
        self.ui.lbl_estado_dot.config(fg=colores.get(tipo, "#A8C4E0"))

    def _activar_paso(self, paso):
        for i, (badge, lbl) in enumerate(self.ui.step_labels):
            if i + 1 <= paso:
                badge.config(bg="#5DCA8A")
                lbl.config(fg=TEXT_LIGHT,
                           font=("Segoe UI", 10,
                                 "bold" if i + 1 == paso else "normal"))
            else:
                badge.config(bg=ACCENT)
                lbl.config(fg="#A8C4E0", font=("Segoe UI", 10))

    # ── Selección de columnas ─────────────────────────────────────────────────

    def _seleccionar_todo(self):
        for item in self.ui.lista_columnas.get_children():
            self.ui.lista_columnas.selection_add(item)
        self._actualizar_contador()

    def _limpiar_seleccion(self):
        self.ui.lista_columnas.selection_remove(
            self.ui.lista_columnas.get_children())
        self._actualizar_contador()

    def _actualizar_contador(self, event=None):
        n = len(self.ui.lista_columnas.selection())
        self.ui.lbl_contador.config(
            text=f"{n} columna{'s' if n != 1 else ''} "
                 f"seleccionada{'s' if n != 1 else ''}",
            fg=ACCENT if n > 0 else TEXT_MUTED
        )
        self._validar_exportar()

    def _get_columnas_seleccionadas(self):
        items = self.ui.lista_columnas.selection()
        return [str(self.ui.lista_columnas.item(i, "values")[0]) for i in items]

    def _abrir_renombrar(self):
        cols = self._get_columnas_seleccionadas()
        if not cols:
            messagebox.showwarning("Sin selección",
                                   "Selecciona columnas antes de renombrar.")
            return
        dialog = RenameDialog(self.root, cols)
        self.root.wait_window(dialog)
        if dialog.result:
            self._exportar_con_nombres(dialog.result)

    # ── Carga de archivo ──────────────────────────────────────────────────────

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if not ruta:
            return

        self.ui.lbl_archivo.config(
            text=os.path.basename(ruta), fg=TEXT_DARK,
            font=("Segoe UI", 10, "bold")
        )
        self._set_estado("Leyendo hojas…", "info")
        self.root.update()

        try:
            self.reader.abrir(ruta)
            hojas = self.reader.obtener_hojas()
            self.ui.combo_hojas.config(state="readonly", values=hojas)
            if hojas:
                self.ui.combo_hojas.current(0)
                self.cargar_columnas()
            self._activar_paso(2)
            self._set_estado(
                f"Archivo cargado\n({len(hojas)} hoja{'s' if len(hojas)!=1 else ''})",
                "ok"
            )
        except Exception as e:
            messagebox.showerror("Error al cargar", str(e))
            self._set_estado("Error al cargar archivo", "error")

    def cargar_columnas(self, event=None):
        hoja = self.ui.combo_hojas.get()
        if not hoja:
            return

        self._set_estado("Detectando columnas…", "info")
        self.root.update()

        try:
            for item in self.ui.lista_columnas.get_children():
                self.ui.lista_columnas.delete(item)

            cols = self.reader.obtener_columnas(hoja)

            for idx, col in enumerate(cols):
                tag = "even" if idx % 2 == 0 else "odd"
                self.ui.lista_columnas.insert(
                    "", tk.END, values=(col,), tags=(tag,))

            self.ui.btn_exportar.config(state=tk.NORMAL)
            self.ui.btn_renombrar.config(state=tk.NORMAL)
            self.ui.lbl_info_hoja.config(
                text=f"{len(cols)} columnas encontradas", fg=SUCCESS)
            self._activar_paso(3)
            self._set_estado(f"Hoja '{hoja}' cargada\nSelecciona columnas", "ok")
            self._actualizar_contador()
        except Exception as e:
            messagebox.showerror("Error al leer hoja", str(e))
            self._set_estado("Error al leer hoja", "error")

    # ── Exportación manual ────────────────────────────────────────────────────

    def exportar_datos(self):
        modo = self.ui.var_modo.get()
        if modo == "columnas":
            columnas = self._get_columnas_seleccionadas()
            if not columnas:
                messagebox.showwarning(
                    "Sin selección",
                    "Selecciona al menos una columna de la lista para exportar.")
                return
            self._exportar_con_nombres({c: c for c in columnas})
        else:
            self._exportar_rangos()

    def _exportar_rangos(self):
        cfg = getattr(self, "_rangos_config", {})
        if not cfg:
            messagebox.showwarning(
                "Sin rangos configurados",
                "Define al menos un rango en el modo Rangos de celdas\n"
                "antes de exportar.")
            return
        self._exportar_con_rangos(cfg)

    def _exportar_con_rangos(self, cfg: dict):
        ruta = filedialog.asksaveasfilename(
            title="Guardar como…",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if not ruta:
            self._set_estado("Exportación cancelada", "warn")
            return
        try:
            self._set_estado("Extrayendo rangos…", "info")
            self.root.update()

            import pandas as pd
            dfs = []
            for hoja, datos in cfg.items():
                df = self.reader.leer_rango(
                    hoja,
                    datos["rango"],
                    fila_encabezado=datos.get("encabezado_en_fila", 0)
                )
                rename_map = datos.get("rename_map", {})
                if rename_map:
                    df = df.rename(columns={k: v for k, v in rename_map.items()
                                            if k in df.columns and v})
                df.insert(0, "Hoja", hoja)
                dfs.append(df)

            if not dfs:
                messagebox.showwarning("Sin datos", "No se extrajeron datos.")
                return

            df_final = pd.concat(dfs, ignore_index=True)

            from core.exporters import _estilizar_xlsx
            df_final.to_excel(ruta, index=False, engine="openpyxl")
            _estilizar_xlsx(ruta)
            n_filas = len(df_final)

            self.ui.lbl_filas.config(
                text=f"✔ {n_filas} filas exportadas",
                fg=SUCCESS, font=("Segoe UI", 9, "bold"))
            self._activar_paso(4)
            self._set_estado(f"Exportado\n{n_filas} filas ✔", "ok")
            messagebox.showinfo(
                "¡Listo!",
                f"Archivo guardado en:\n{ruta}\n\n{n_filas} filas exportadas.")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))
            self._set_estado("Error al exportar", "error")

    def _exportar_con_nombres(self, mapa_nombres):
        ruta = filedialog.asksaveasfilename(
            title="Guardar como…",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if not ruta:
            self._set_estado("Exportación cancelada", "warn")
            return

        try:
            self._set_estado("Generando Excel…", "info")
            self.root.update()

            df = self.reader.df_actual[list(mapa_nombres.keys())].copy()
            df = df.dropna(how="all")
            if mapa_nombres:
                df = df.rename(columns=mapa_nombres)

            from core.exporters import _a_xlsx, _estilizar_xlsx
            import os
            df.to_excel(ruta, index=False, engine="openpyxl")
            _estilizar_xlsx(ruta)
            n_filas = len(df)

            self.ui.lbl_filas.config(
                text=f"✔ {n_filas} filas exportadas",
                fg=SUCCESS, font=("Segoe UI", 9, "bold")
            )
            self._activar_paso(4)
            self._set_estado(f"Exportado\n{n_filas} filas ✔", "ok")
            messagebox.showinfo(
                "¡Listo!",
                f"Archivo guardado en:\n{ruta}\n\n{n_filas} filas exportadas."
            )
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))
            self._set_estado("Error al exportar", "error")



    # ── Modo tabs: columnas / rangos ──────────────────────────────────────────

    def _cambiar_modo(self):
        modo = self.ui.var_modo.get()
        if modo == "columnas":
            self.ui.panel_rangos.pack_forget()
            self.ui.panel_columnas.pack(fill=tk.BOTH, expand=True)
            tiene_cols = len(self.ui.lista_columnas.get_children()) > 0
            self.ui.btn_renombrar.config(
                state=tk.NORMAL if tiene_cols else tk.DISABLED)
        else:
            self.ui.panel_columnas.pack_forget()
            self.ui.panel_rangos.pack(fill=tk.BOTH, expand=True)
            self.ui.btn_renombrar.config(state=tk.DISABLED)
            self._actualizar_panel_rangos()
        self._validar_exportar()

    def _validar_exportar(self):
        modo = self.ui.var_modo.get()
        if modo == "columnas":
            ok = len(self.ui.lista_columnas.selection()) > 0
        else:
            ok = bool(self._rangos_config)
        self.ui.btn_exportar.config(state=tk.NORMAL if ok else tk.DISABLED)

    def _actualizar_panel_rangos(self):
        cfg = self._rangos_config
        for w in self.ui.frame_rangos_lista.winfo_children():
            w.destroy()

        if not cfg:
            self.ui.lbl_rangos_estado.config(
                text="Sin rangos configurados — usa el botón de arriba o "
                     "el menú Rangos → Definir rangos de celdas…",
                fg=TEXT_MUTED)
            return

        self.ui.lbl_rangos_estado.config(
            text=f"✔  {len(cfg)} hoja{'s' if len(cfg) != 1 else ''} con rangos configurados:",
            fg=SUCCESS)

        for hoja, datos in cfg.items():
            row = tk.Frame(self.ui.frame_rangos_lista, bg=BG_CARD)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=f"  • {hoja}",
                     font=("Segoe UI", 10, "bold"),
                     bg=BG_CARD, fg=TEXT_DARK,
                     width=22, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=datos.get("rango", ""),
                     font=("Segoe UI", 10),
                     bg=BG_CARD, fg=ACCENT).pack(side=tk.LEFT, padx=(0, 12))
            ren = len(datos.get("rename_map", {}))
            if ren:
                tk.Label(row, text=f"{ren} cols renombradas",
                         font=("Segoe UI", 9, "italic"),
                         bg=BG_CARD, fg=TEXT_MUTED).pack(side=tk.LEFT)


if __name__ == "__main__":
    root = tk.Tk()
    app  = ExtractorApp(root)
    root.mainloop()
