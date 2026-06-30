# core/notifier.py
"""
Envía notificaciones por email al completar (o fallar) una ejecución.
Usa SMTP estándar; compatible con Gmail, Outlook, o servidor corporativo.

Puede adjuntar el archivo de salida de la extracción y/o un informe de
auditoría en PDF, como respaldo de la ejecución.
"""

import os
import smtplib
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

MAX_ADJUNTO_MB = 20  # límite razonable para no chocar con límites de SMTP


def notificar(perfil: dict, estado: str, filas: int = 0,
              archivo: str = "", error: str = "",
              adjuntos: list = None):
    """
    Envía email si el perfil tiene notificación configurada para ese estado.

    estado: "ok" | "sin_archivo" | "error"
    adjuntos: lista opcional de rutas de archivo a adjuntar (ej. el archivo
              de salida exportado y/o el informe de auditoría en PDF).
              Archivos inexistentes o que excedan MAX_ADJUNTO_MB se omiten
              silenciosamente (se avisa en el cuerpo del correo).
    """
    notif = perfil.get("notificacion", {})

    if not notif.get("email_destino"):
        return  # sin destinatario configurado

    if estado == "ok" and not notif.get("en_exito"):
        return
    if estado in ("error", "sin_archivo") and not notif.get("en_error"):
        return

    adjuntos = adjuntos or []
    adjuntos_validos, adjuntos_omitidos = _filtrar_adjuntos(adjuntos)

    asunto, cuerpo = _construir_mensaje(
        perfil["nombre"], estado, filas, archivo, error, adjuntos_omitidos
    )
    _enviar(notif, asunto, cuerpo, adjuntos_validos)


def _filtrar_adjuntos(rutas: list):
    """Separa adjuntos válidos (existen y no exceden el tamaño) de omitidos."""
    validos, omitidos = [], []
    for ruta in rutas:
        if not ruta or not os.path.exists(ruta):
            continue
        size_mb = os.path.getsize(ruta) / (1024 * 1024)
        if size_mb > MAX_ADJUNTO_MB:
            omitidos.append((os.path.basename(ruta), size_mb))
        else:
            validos.append(ruta)
    return validos, omitidos


def _construir_mensaje(nombre_perfil, estado, filas, archivo, error, adjuntos_omitidos=None):
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    adjuntos_omitidos = adjuntos_omitidos or []

    if estado == "ok":
        asunto = f"✅ Extracción completada — {nombre_perfil}"
        cuerpo = f"""
Ejecución exitosa del perfil <b>{nombre_perfil}</b><br><br>
<b>Fecha:</b> {ts}<br>
<b>Archivo procesado:</b> {archivo}<br>
<b>Filas exportadas:</b> {filas}<br>
"""
    elif estado == "sin_archivo":
        asunto = f"⚠️ Archivo no encontrado — {nombre_perfil}"
        cuerpo = f"""
La ejecución del perfil <b>{nombre_perfil}</b> no encontró el archivo esperado.<br><br>
<b>Fecha:</b> {ts}<br>
<b>Detalle:</b> {error}<br>
"""
    else:
        asunto = f"❌ Error en extracción — {nombre_perfil}"
        cuerpo = f"""
La ejecución del perfil <b>{nombre_perfil}</b> falló con un error.<br><br>
<b>Fecha:</b> {ts}<br>
<b>Error:</b> {error}<br>
"""

    if adjuntos_omitidos:
        lista = "<br>".join(
            f"- {nombre} ({size_mb:.1f} MB, supera el límite de {MAX_ADJUNTO_MB} MB)"
            for nombre, size_mb in adjuntos_omitidos
        )
        cuerpo += f"<br><b>⚠ Adjuntos omitidos por tamaño:</b><br>{lista}<br>"

    return asunto, cuerpo


def _enviar(notif: dict, asunto: str, cuerpo_html: str, adjuntos: list = None):
    adjuntos = adjuntos or []

    msg = MIMEMultipart("mixed")
    msg["Subject"] = asunto
    msg["From"]    = notif.get("smtp_user", "")
    msg["To"]      = notif["email_destino"]

    cuerpo_msg = MIMEMultipart("alternative")
    cuerpo_msg.attach(MIMEText(cuerpo_html, "html", "utf-8"))
    msg.attach(cuerpo_msg)

    for ruta in adjuntos:
        _adjuntar_archivo(msg, ruta)

    host = notif.get("smtp_host", "")
    port = int(notif.get("smtp_port", 587))
    user = notif.get("smtp_user", "")
    pwd  = notif.get("smtp_pass", "")

    if not host:
        raise ValueError("SMTP no configurado. Completa el host en la configuración del perfil.")

    with smtplib.SMTP(host, port, timeout=30) as server:
        server.ehlo()
        server.starttls()
        if user and pwd:
            server.login(user, pwd)
        server.sendmail(user, notif["email_destino"], msg.as_string())


def _adjuntar_archivo(msg: MIMEMultipart, ruta: str):
    nombre = os.path.basename(ruta)
    tipo, _ = mimetypes.guess_type(ruta)
    tipo_principal, tipo_secundario = (tipo.split("/", 1) if tipo else ("application", "octet-stream"))

    with open(ruta, "rb") as f:
        parte = MIMEBase(tipo_principal, tipo_secundario)
        parte.set_payload(f.read())

    encoders.encode_base64(parte)
    parte.add_header("Content-Disposition", f'attachment; filename="{nombre}"')
    msg.attach(parte)

