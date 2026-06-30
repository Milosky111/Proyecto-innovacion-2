# security.py
"""
Validación de archivos Excel antes de procesarlos.

Diseñado para el escenario "archivo de origen desconocido" (terceros suben
el Excel), pero se ejecuta siempre — incluso en uso interno — porque no cuesta
nada y evita crashes con archivos corruptos.

Filosofía: RECHAZAR temprano y con un mensaje claro, nunca intentar "limpiar"
o "sanitizar" un archivo sospechoso. Si algo no cuadra, no se abre.
"""
import os
import zipfile

# --- Límites configurables ---------------------------------------------------
MAX_FILE_SIZE_MB = 50          # tamaño máximo del archivo subido
MAX_UNCOMPRESSED_MB = 500      # tamaño máximo permitido tras descomprimir (zip bomb)
MAX_COMPRESSION_RATIO = 100    # ratio descomprimido/comprimido sospechoso
MAX_SHEETS = 50                # nº máximo de hojas
MAX_ENTRIES_IN_ZIP = 5000      # nº máximo de archivos dentro del .xlsx (zip)

# Firmas binarias reales (los primeros bytes del archivo, sin importar la extensión)
FIRMA_ZIP = b"PK\x03\x04"          # .xlsx / .xlsm modernos (Office Open XML)
FIRMA_OLE2 = b"\xD0\xCF\x11\xE0"   # .xls legado (formato binario antiguo)

EXTENSIONES_PERMITIDAS = (".xlsx", ".xls")
EXTENSIONES_CON_MACROS = (".xlsm", ".xltm", ".xlsb")  # rechazadas explícitamente


class ArchivoInvalido(Exception):
    """Se lanza cuando un archivo no pasa la validación de seguridad."""
    pass


def _leer_firma(ruta, n=8):
    with open(ruta, "rb") as f:
        return f.read(n)


def validar_extension(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in EXTENSIONES_CON_MACROS:
        raise ArchivoInvalido(
            f"Los archivos con macros ({ext}) no están permitidos por seguridad. "
            "Guarda el Excel como .xlsx (sin macros) y vuelve a intentarlo."
        )
    if ext not in EXTENSIONES_PERMITIDAS:
        raise ArchivoInvalido(
            f"Extensión '{ext}' no soportada. Solo se aceptan archivos .xlsx o .xls."
        )
    return ext


def validar_tamano(ruta):
    if not os.path.exists(ruta):
        raise ArchivoInvalido("El archivo no existe o no se puede acceder a él.")
    size_mb = os.path.getsize(ruta) / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        raise ArchivoInvalido(
            f"El archivo pesa {size_mb:.1f} MB, supera el límite de {MAX_FILE_SIZE_MB} MB."
        )
    if size_mb == 0:
        raise ArchivoInvalido("El archivo está vacío (0 bytes).")


def validar_firma_binaria(ruta, ext):
    """
    Verifica que el contenido real del archivo coincida con su extensión.
    Evita que un archivo renombrado (ej. un .exe llamado 'datos.xlsx') pase
    como si fuera un Excel legítimo.
    """
    firma = _leer_firma(ruta)
    if ext == ".xlsx" and not firma.startswith(FIRMA_ZIP):
        raise ArchivoInvalido(
            "El archivo tiene extensión .xlsx pero su contenido no es un ZIP/Office válido. "
            "Puede estar corrupto o no ser un Excel real."
        )
    if ext == ".xls" and not firma.startswith(FIRMA_OLE2):
        raise ArchivoInvalido(
            "El archivo tiene extensión .xls pero su contenido no coincide con el "
            "formato binario de Excel. Puede estar corrupto o no ser un Excel real."
        )


def validar_zip_seguro(ruta, ext):
    """
    Solo aplica a .xlsx (que internamente es un ZIP). Protege contra:
    - Zip bombs: archivos pequeños que descomprimen a tamaños enormes
    - Excesivo número de entradas (otra forma de agotar recursos)
    - Macros embebidas (xl/vbaProject.bin), aunque la extensión diga .xlsx
    - Path traversal en nombres de entrada (../../etc/passwd)
    """
    if ext != ".xlsx":
        return  # los .xls legacy no son ZIP, no aplica esta validación

    try:
        with zipfile.ZipFile(ruta) as zf:
            entradas = zf.infolist()

            if len(entradas) > MAX_ENTRIES_IN_ZIP:
                raise ArchivoInvalido(
                    f"El archivo contiene {len(entradas)} elementos internos, "
                    f"más del límite permitido ({MAX_ENTRIES_IN_ZIP}). Parece corrupto o malicioso."
                )

            total_descomprimido = 0
            for info in entradas:
                nombre = info.filename

                # Path traversal: nombres que intentan escapar del zip
                if nombre.startswith("/") or ".." in nombre.replace("\\", "/").split("/"):
                    raise ArchivoInvalido(
                        "El archivo contiene rutas internas sospechosas. No se procesará."
                    )

                # Macros embebidas, sin importar la extensión declarada
                if nombre.lower() == "xl/vbaproject.bin":
                    raise ArchivoInvalido(
                        "El archivo contiene macros (VBA) embebidas y no se puede procesar "
                        "por seguridad, aunque su extensión sea .xlsx."
                    )

                total_descomprimido += info.file_size

                # Ratio de compresión sospechoso en una sola entrada (zip bomb clásico)
                if info.compress_size > 0:
                    ratio = info.file_size / info.compress_size
                    if ratio > MAX_COMPRESSION_RATIO and info.file_size > 10 * 1024 * 1024:
                        raise ArchivoInvalido(
                            "El archivo contiene datos con una compresión anormalmente alta "
                            "(posible 'zip bomb'). No se procesará."
                        )

            total_mb = total_descomprimido / (1024 * 1024)
            if total_mb > MAX_UNCOMPRESSED_MB:
                raise ArchivoInvalido(
                    f"El contenido descomprimido del archivo ({total_mb:.0f} MB) supera el "
                    f"límite permitido ({MAX_UNCOMPRESSED_MB} MB). No se procesará."
                )

    except zipfile.BadZipFile:
        raise ArchivoInvalido(
            "El archivo .xlsx está corrupto o no es un ZIP válido."
        )


def validar_excel(ruta):
    """
    Punto de entrada único. Lanza ArchivoInvalido con un mensaje claro
    en cuanto encuentra el primer problema. No intenta corregir nada.
    """
    ext = validar_extension(ruta)
    validar_tamano(ruta)
    validar_firma_binaria(ruta, ext)
    validar_zip_seguro(ruta, ext)
    return True
