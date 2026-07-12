# config.py
import platform

# ─── Tipografía multiplataforma ───────────────────────────────────────────────
# "Segoe UI" solo existe en Windows. Si se usa tal cual en macOS o Linux,
# Tk la reemplaza por una fuente por defecto con métricas distintas y la
# interfaz se ve distinta (textos más anchos/angostos, botones que no
# coinciden, etc). Por eso elegimos la fuente según el sistema operativo:
# el resultado visual (tamaños, negritas, cursivas) queda equivalente en
# Windows y en Mac.
_SISTEMA = platform.system()
if _SISTEMA == "Windows":
    FONT_FAMILY = "Segoe UI"
elif _SISTEMA == "Darwin":       # macOS
    FONT_FAMILY = "Helvetica Neue"
else:                             # Linux / otros
    FONT_FAMILY = "DejaVu Sans"


def F(size, *estilos):
    """
    Atajo para construir tuplas de fuente consistentes en toda la app.
    Uso:  F(10)                -> (FONT_FAMILY, 10)
          F(10, "bold")        -> (FONT_FAMILY, 10, "bold")
          F(9, "bold","italic")-> (FONT_FAMILY, 9, "bold italic")
    """
    if not estilos:
        return (FONT_FAMILY, size)
    return (FONT_FAMILY, size, " ".join(estilos))


# ─── Paleta de colores ────────────────────────────────────────────────────────
BG_MAIN      = "#F0F4F8"
BG_SIDEBAR   = "#1F3864"
BG_CARD      = "#FFFFFF"
ACCENT       = "#2E75B6"
ACCENT_HOVER = "#1F5A9E"
SUCCESS      = "#1E8449"
WARNING      = "#D4AC0D"
ERROR        = "#C0392B"
TEXT_LIGHT   = "#FFFFFF"
TEXT_DARK    = "#1A1A2E"
TEXT_MUTED   = "#6B7A8D"
BORDER       = "#D0D8E4"
ROW_EVEN     = "#EDF3FB"
ROW_ODD      = "#FFFFFF"
# ──────────────────────────────────────────────────────────────────────────────
# ─── Configuración de automatización ─────────────────────────────────────────
CONFIG_PATH       = "config_perfiles.json"
LOG_DB_PATH       = "extractor_log.db"
HORA_DEFAULT      = "07:00"
FORMATOS_DESTINO  = ["csv", "xlsx", "sqlite"]
