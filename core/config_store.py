# core/config_store.py
"""
Gestiona perfiles de automatización persistidos en JSON.
Cada perfil define: qué Excel, qué rangos, dónde guardar, cuándo ejecutar y cómo notificar.
"""

import json
import uuid
import os
from datetime import datetime
from config import CONFIG_PATH


def _perfil_vacio():
    return {
        "id":              str(uuid.uuid4()),
        "nombre":          "",
        "carpeta_origen":  "",
        "patron_archivo":  "",        # ej: "cierre_tgm_{YYYYMM}.xlsx"
        "hoja":            "",
        "rangos":          [],        # lista de strings "B5:F20"
        "rename_map":      {},        # {"col_original": "col_nueva"}
        "destino": {
            "tipo":            "csv",  # csv | xlsx | sqlite
            "carpeta":         "",
            "nombre_archivo":  "",     # sin extensión; se agrega automáticamente
            "modo":            "append",  # append | replace
        },
        "schedule": {
            "hora":    "07:00",
            "activo":  False,
            "ultimo_run": None,
        },
        "notificacion": {
            "email_destino": "",
            "smtp_host":     "",
            "smtp_port":     587,
            "smtp_user":     "",
            "smtp_pass":     "",
            "en_error":      True,
            "en_exito":      False,
        },
        "creado_en": datetime.now().isoformat(),
    }


class ConfigStore:
    def __init__(self, ruta=CONFIG_PATH):
        self.ruta = ruta
        self._data = self._cargar()

    # ── I/O ──────────────────────────────────────────────────────────────────

    def _cargar(self):
        if os.path.exists(self.ruta):
            try:
                with open(self.ruta, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        return {"perfiles": [], "smtp_global": {}}

    def guardar(self):
        with open(self.ruta, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    # ── Perfiles ─────────────────────────────────────────────────────────────

    def listar_perfiles(self):
        return list(self._data.get("perfiles", []))

    def obtener_perfil(self, perfil_id):
        for p in self._data["perfiles"]:
            if p["id"] == perfil_id:
                return p
        return None

    def nuevo_perfil(self):
        """Crea un perfil vacío, lo agrega y retorna su id."""
        perfil = _perfil_vacio()
        self._data.setdefault("perfiles", []).append(perfil)
        self.guardar()
        return perfil

    def actualizar_perfil(self, perfil_id, datos: dict):
        for i, p in enumerate(self._data["perfiles"]):
            if p["id"] == perfil_id:
                self._data["perfiles"][i].update(datos)
                self.guardar()
                return True
        return False

    def eliminar_perfil(self, perfil_id):
        antes = len(self._data["perfiles"])
        self._data["perfiles"] = [
            p for p in self._data["perfiles"] if p["id"] != perfil_id
        ]
        if len(self._data["perfiles"]) < antes:
            self.guardar()
            return True
        return False

    def marcar_ultimo_run(self, perfil_id, timestamp=None):
        ts = timestamp or datetime.now().isoformat()
        for p in self._data["perfiles"]:
            if p["id"] == perfil_id:
                p["schedule"]["ultimo_run"] = ts
                self.guardar()
                return

    # ── SMTP global ──────────────────────────────────────────────────────────

    def get_smtp_global(self):
        return self._data.get("smtp_global", {})

    def set_smtp_global(self, datos: dict):
        self._data["smtp_global"] = datos
        self.guardar()
