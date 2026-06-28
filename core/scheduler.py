# core/scheduler.py
"""
Integración con Windows Task Scheduler (schtasks.exe).
Registra/actualiza/elimina tareas programadas para ejecutar runner.py headless.

Alternativa para desarrollo/pruebas: scheduler en proceso con la librería `schedule`.
"""

import os
import sys
import subprocess
import shutil
from config import CONFIG_PATH


# ── Windows Task Scheduler ────────────────────────────────────────────────────

def _runner_path() -> str:
    """Ruta absoluta al runner.py (o runner.exe si está compilado)."""
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    exe  = os.path.join(base, "runner.exe")
    if os.path.exists(exe):
        return exe
    return os.path.join(base, "core", "runner.py")

def _python_path() -> str:
    return sys.executable


def registrar_tarea(nombre_perfil: str, perfil_id: str, hora: str = "07:00") -> bool:
    """
    Crea o actualiza una tarea en Windows Task Scheduler para el perfil dado.
    La tarea ejecuta runner.py --perfil-id <id> a la hora indicada (HH:MM), diariamente.
    Retorna True si tuvo éxito.
    """
    if not shutil.which("schtasks"):
        raise EnvironmentError("schtasks.exe no disponible. Requiere Windows.")

    nombre_tarea = _nombre_tarea(nombre_perfil)
    runner       = _runner_path()
    python       = _python_path()

    if runner.endswith(".py"):
        cmd_accion = f'"{python}" "{runner}" --perfil-id {perfil_id}'
    else:
        cmd_accion = f'"{runner}" --perfil-id {perfil_id}'

    args = [
        "schtasks", "/Create", "/F",
        "/TN",  nombre_tarea,
        "/TR",  cmd_accion,
        "/SC",  "DAILY",
        "/ST",  hora,
        "/RL",  "HIGHEST",
    ]

    result = subprocess.run(args, capture_output=True, text=True)
    return result.returncode == 0


def eliminar_tarea(nombre_perfil: str) -> bool:
    """Elimina la tarea programada asociada al perfil."""
    if not shutil.which("schtasks"):
        return False

    nombre_tarea = _nombre_tarea(nombre_perfil)
    result = subprocess.run(
        ["schtasks", "/Delete", "/F", "/TN", nombre_tarea],
        capture_output=True, text=True
    )
    return result.returncode == 0


def listar_tareas_registradas() -> list:
    """Retorna los nombres de tareas del extractor registradas en Windows."""
    if not shutil.which("schtasks"):
        return []

    result = subprocess.run(
        ["schtasks", "/Query", "/FO", "CSV", "/NH"],
        capture_output=True, text=True
    )
    tareas = []
    for line in result.stdout.splitlines():
        if "ExtractorTGM" in line:
            tareas.append(line.split(",")[0].strip('"'))
    return tareas


def _nombre_tarea(nombre_perfil: str) -> str:
    import re
    safe = re.sub(r"[^a-zA-Z0-9_\- ]", "", nombre_perfil).strip()
    return f"ExtractorTGM\\{safe}"


# ── Scheduler en proceso (desarrollo / no-Windows) ────────────────────────────

def iniciar_scheduler_en_proceso(store, logger):
    """
    Fallback sin schtasks: usa la librería `schedule` para correr en proceso.
    Útil durante desarrollo. En producción se usa Task Scheduler de Windows.
    """
    try:
        import schedule
        import threading
        import time
    except ImportError:
        raise ImportError("Instala 'schedule': pip install schedule")

    def _job():
        from core.runner import ejecutar_todos
        ejecutar_todos(store, logger)

    # Agrupa todos los perfiles activos bajo la misma hora (07:00 por defecto)
    schedule.every().day.at("07:00").do(_job)

    def _loop():
        while True:
            schedule.run_pending()
            time.sleep(30)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    return t
