# core/logger.py
"""
Registro persistente de cada ejecución en SQLite.
Permite consultar historial por perfil, fecha y estado.
"""

import sqlite3
import os
from datetime import datetime
from config import LOG_DB_PATH


class RunLogger:
    def __init__(self, db_path=LOG_DB_PATH):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        con = sqlite3.connect(self.db_path)
        con.execute("""
            CREATE TABLE IF NOT EXISTS run_log (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                perfil_id      TEXT NOT NULL,
                perfil_nombre  TEXT,
                timestamp      TEXT NOT NULL,
                estado         TEXT NOT NULL,
                filas_export   INTEGER DEFAULT 0,
                archivo_usado  TEXT,
                destino        TEXT,
                detalle_error  TEXT
            )
        """)
        con.commit()
        con.close()

    def registrar(self, perfil_id: str, perfil_nombre: str,
                  estado: str, filas: int = 0,
                  archivo_usado: str = "", archivo: str = "",
                  destino: str = "", error: str = ""):
        """
        estado: "ok" | "sin_archivo" | "error"
        """
        con = sqlite3.connect(self.db_path)
        arch = archivo_usado or archivo
        con.execute("""
            INSERT INTO run_log
              (perfil_id, perfil_nombre, timestamp, estado,
               filas_export, archivo_usado, destino, detalle_error)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            perfil_id,
            perfil_nombre,
            datetime.now().isoformat(timespec="seconds"),
            estado,
            filas,
            arch,
            destino,
            error,
        ))
        con.commit()
        con.close()

    def obtener_historial(self, perfil_id: str = None, limite: int = 100):
        """Retorna lista de dicts con las últimas ejecuciones."""
        con = sqlite3.connect(self.db_path)
        con.row_factory = sqlite3.Row
        if perfil_id:
            rows = con.execute(
                "SELECT * FROM run_log WHERE perfil_id=? ORDER BY id DESC LIMIT ?",
                (perfil_id, limite)
            ).fetchall()
        else:
            rows = con.execute(
                "SELECT * FROM run_log ORDER BY id DESC LIMIT ?",
                (limite,)
            ).fetchall()
        con.close()
        return [dict(r) for r in rows]

    def ultimo_run(self, perfil_id: str):
        """Retorna el dict del último run de un perfil, o None."""
        con = sqlite3.connect(self.db_path)
        con.row_factory = sqlite3.Row
        row = con.execute(
            "SELECT * FROM run_log WHERE perfil_id=? ORDER BY id DESC LIMIT 1",
            (perfil_id,)
        ).fetchone()
        con.close()
        return dict(row) if row else None
