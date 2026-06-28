# core/runner.py
"""
Runner headless — ejecuta un perfil completo sin interfaz gráfica.
Puede invocarse desde el scheduler, desde CLI o desde la UI (botón "Ejecutar ahora").

Flujo por perfil:
  1. Resolver archivo mensual
  2. Leer rangos configurados
  3. Aplicar rename_map
  4. Exportar al destino
  5. Registrar en log
  6. Notificar por email
"""

import os
import sys

# Permite importar desde la raíz del proyecto cuando se ejecuta directamente
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from datetime import datetime

from core.excel_reader  import ExcelReader
from core.exporters     import exportar
from core.logger        import RunLogger
from core.notifier      import notificar
from core.config_store  import ConfigStore


def ejecutar_perfil(perfil: dict, logger: RunLogger = None, fecha: datetime = None) -> dict:
    """
    Ejecuta un perfil y retorna un dict con el resultado:
    {
        "estado":  "ok" | "sin_archivo" | "error",
        "filas":   int,
        "archivo": str,
        "error":   str,
    }
    """
    logger = logger or RunLogger()
    fecha  = fecha  or datetime.now()
    reader = ExcelReader()

    resultado = {"estado": "ok", "filas": 0, "archivo": "", "error": ""}

    try:
        # 1. Resolver archivo del mes
        ruta = reader.resolver_archivo_mensual(
            perfil["carpeta_origen"],
            perfil["patron_archivo"],
            fecha=fecha,
        )
        resultado["archivo"] = ruta
        reader.abrir(ruta)

        # 2. Leer rangos o columnas completas
        rangos = perfil.get("rangos", [])

        if rangos:
            dfs = [reader.leer_rango(perfil["hoja"], r) for r in rangos]
            dfs = [d for d in dfs if not d.empty]
            df  = pd.concat(dfs, axis=1) if dfs else pd.DataFrame()
        else:
            # Modo libre: carga todas las columnas de la hoja
            cols = reader.obtener_columnas(perfil["hoja"])
            df   = reader.df_actual[cols].copy()
            df   = df.dropna(how="all")

        if df.empty:
            raise ValueError("No se extrajeron datos. Verifica los rangos configurados.")

        # 3. Renombrar columnas si hay mapa
        rename_map = perfil.get("rename_map", {})
        if rename_map:
            df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

        # 4. Exportar
        n_filas = exportar(df, perfil["destino"], perfil_nombre=perfil["nombre"])
        resultado["filas"] = n_filas

    except FileNotFoundError as e:
        resultado["estado"] = "sin_archivo"
        resultado["error"]  = str(e)

    except Exception as e:
        resultado["estado"] = "error"
        resultado["error"]  = str(e)

    # 5. Log
    logger.registrar(
        perfil_id     = perfil["id"],
        perfil_nombre = perfil["nombre"],
        estado        = resultado["estado"],
        filas         = resultado["filas"],
        archivo_usado = resultado["archivo"],
        destino       = str(perfil.get("destino", {})),
        error         = resultado["error"],
    )

    # 6. Notificar
    try:
        notificar(
            perfil  = perfil,
            estado  = resultado["estado"],
            filas   = resultado["filas"],
            archivo = resultado["archivo"],
            error   = resultado["error"],
        )
    except Exception as e_notif:
        # El fallo de notificación no cancela el resultado
        resultado["notif_error"] = str(e_notif)

    return resultado


def ejecutar_todos(store: ConfigStore = None, logger: RunLogger = None):
    """Ejecuta todos los perfiles activos. Usado por el scheduler."""
    store  = store  or ConfigStore()
    logger = logger or RunLogger()

    perfiles = [p for p in store.listar_perfiles()
                if p.get("schedule", {}).get("activo")]

    resultados = []
    for perfil in perfiles:
        r = ejecutar_perfil(perfil, logger=logger)
        store.marcar_ultimo_run(perfil["id"])
        resultados.append({"perfil": perfil["nombre"], **r})

    return resultados


# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse, json

    parser = argparse.ArgumentParser(description="Runner headless de extracción")
    parser.add_argument("--perfil-id", help="ID del perfil a ejecutar (omitir = todos los activos)")
    parser.add_argument("--todos",     action="store_true", help="Ejecuta todos los perfiles activos")
    args = parser.parse_args()

    store  = ConfigStore()
    logger = RunLogger()

    if args.todos or not args.perfil_id:
        resultados = ejecutar_todos(store, logger)
        print(json.dumps(resultados, ensure_ascii=False, indent=2))
    else:
        perfil = store.obtener_perfil(args.perfil_id)
        if not perfil:
            print(f"Perfil '{args.perfil_id}' no encontrado.")
            sys.exit(1)
        r = ejecutar_perfil(perfil, logger=logger)
        print(json.dumps(r, ensure_ascii=False, indent=2))
