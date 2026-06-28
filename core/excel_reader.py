# core/excel_reader.py
"""
Lee archivos Excel sin necesitar que estén abiertos (openpyxl).
Soporta:
  - Detección automática de la fila de encabezados
  - Extracción por rango de celdas (ej: "B3:H20")
  - Resolución de archivo mensual por patrón (ej: "cierre_tgm_{YYYYMM}.xlsx")
"""

import os
import re
import glob
import pandas as pd
import openpyxl
from datetime import datetime


class ExcelReader:
    def inspeccionar_celda(self, nombre_hoja: str, celda: str):
        wb = openpyxl.load_workbook(self.ruta_archivo, data_only=False)
        ws = wb[nombre_hoja]
        c = ws[celda]
        print(f"Valor: {c.value}")
        print(f"Tipo:  {type(c.value)}")

    def __init__(self):
        self.ruta_archivo = None
        self.df_actual    = None
        self.excel_file   = None
        self.fila_encabezado = 0

    # ── Apertura ─────────────────────────────────────────────────────────────

    def abrir(self, ruta: str):
        """Carga el archivo y lo prepara para lectura."""
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"No se encontró el archivo: {ruta}")
        self.ruta_archivo = ruta
        self.excel_file   = pd.ExcelFile(ruta, engine="openpyxl")
        return self

    def resolver_archivo_mensual(self, carpeta: str, patron: str, fecha: datetime = None) -> str:
        """
        Dado un patrón con {YYYYMM} o {YYYY} y {MM}, encuentra el archivo
        del mes en la carpeta indicada.

        Ejemplo:
            patron = "cierre_tgm_{YYYYMM}.xlsx"
            → busca "cierre_tgm_202604.xlsx" en carpeta
        """
        fecha = fecha or datetime.now()

        nombre = (patron
                  .replace("{YYYYMM}", fecha.strftime("%Y%m"))
                  .replace("{YYYY}",   fecha.strftime("%Y"))
                  .replace("{MM}",     fecha.strftime("%m"))
                  .replace("{DD}",     fecha.strftime("%d")))

        ruta_completa = os.path.join(carpeta, nombre)

        if os.path.exists(ruta_completa):
            return ruta_completa

        # Intenta búsqueda con glob si el nombre tiene comodines residuales
        candidatos = glob.glob(os.path.join(carpeta, nombre))
        if candidatos:
            return max(candidatos, key=os.path.getmtime)

        raise FileNotFoundError(
            f"No se encontró '{nombre}' en '{carpeta}'.\n"
            f"Verifica que el patrón '{patron}' coincida con el nombre real del archivo."
        )

    # ── Hojas ────────────────────────────────────────────────────────────────

    def obtener_hojas(self):
        self._check_abierto()
        return self.excel_file.sheet_names

    # ── Columnas (modo libre, detección automática de encabezado) ─────────────

    def obtener_columnas(self, nombre_hoja: str):
        """Detecta encabezados automáticamente y carga el DataFrame."""
        self._check_abierto()
        mejor_fila = self._detectar_encabezado(nombre_hoja)
        self.fila_encabezado = mejor_fila

        self.df_actual = pd.read_excel(
            self.ruta_archivo, sheet_name=nombre_hoja,
            header=mejor_fila, engine="openpyxl"
        )
        self.df_actual.columns = [str(c).strip() for c in self.df_actual.columns]

        return [
            col for col in self.df_actual.columns
            if col
            and not col.lower().startswith("unnamed")
            and col.upper() != "REPORTE FINAL"
        ]

    # ── Rangos (modo presentación — celda a celda) ────────────────────────────

    def leer_rango(self, nombre_hoja: str, rango: str,
                fila_encabezado: int = 0) -> pd.DataFrame:
        """
        Extrae un rango usando xlwings (Excel real) para obtener valores
        calculados, fórmulas y listas desplegables correctamente.
        """
        import xlwings as xw

        with xw.App(visible=False, add_book=False) as app:
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(self.ruta_archivo)
            ws = wb.sheets[nombre_hoja]
            datos = ws.range(rango).value
            wb.close()

        if datos is None:
            return pd.DataFrame()

        # Celda única → xlwings devuelve el valor directo
        if not isinstance(datos, list):
            datos = [[datos]]

        # Fila única → xlwings devuelve lista plana
        if datos and not isinstance(datos[0], list):
            datos = [datos]

        if not datos:
            return pd.DataFrame()

        # Encabezados
        enc_idx = min(fila_encabezado, len(datos) - 1)
        encabezados = [
            str(v).strip() if v is not None else f"Col_{i}"
            for i, v in enumerate(datos[enc_idx])
        ]

        # Desduplicar
        seen = {}
        enc_unicos = []
        for nombre in encabezados:
            if nombre in seen:
                seen[nombre] += 1
                enc_unicos.append(f"{nombre}.{seen[nombre]}")
            else:
                seen[nombre] = 0
                enc_unicos.append(nombre)

        df = pd.DataFrame(datos[enc_idx + 1:], columns=enc_unicos)
        return df.dropna(how="all")

    def leer_multiples_rangos(self, nombre_hoja: str, rangos: list) -> pd.DataFrame:
        """
        Lee varios rangos de la misma hoja y los concatena horizontalmente.
        Útil cuando los datos de una fila están dispersos en bloques.
        """
        dfs = [self.leer_rango(nombre_hoja, r) for r in rangos]
        dfs = [d for d in dfs if not d.empty]
        if not dfs:
            return pd.DataFrame()
        return pd.concat(dfs, axis=1)

    # ── Helpers internos ─────────────────────────────────────────────────────

    def _check_abierto(self):
        if not self.ruta_archivo:
            raise RuntimeError("Debes llamar a abrir() antes de leer datos.")

    def _detectar_encabezado(self, nombre_hoja: str) -> int:
        """
        Analiza las primeras 30 filas y elige la que más parece encabezado:
        prioriza filas con texto (peso 2) sobre números puros (peso 0.3).
        """
        df_temp = pd.read_excel(
            self.ruta_archivo, sheet_name=nombre_hoja,
            header=None, nrows=30, engine="openpyxl"
        )

        mejor_fila, max_score = 0, -1

        for idx, row in df_temp.iterrows():
            celdas = [x for x in row if pd.notna(x) and str(x).strip() not in ("", " ")]
            if not celdas:
                continue

            score = sum(2 if any(c.isalpha() for c in str(x).strip()) else 0.3
                        for x in celdas)

            if score > max_score:
                max_score  = score
                mejor_fila = idx

        return mejor_fila
