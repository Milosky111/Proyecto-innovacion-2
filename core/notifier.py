# core/notifier.py
"""
Envía notificaciones por email al completar (o fallar) una ejecución.
Usa SMTP estándar; compatible con Gmail, Outlook, o servidor corporativo.
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime


def notificar(perfil: dict, estado: str, filas: int = 0,
              archivo: str = "", error: str = ""):
    """
    Envía email si el perfil tiene notificación configurada para ese estado.

    estado: "ok" | "sin_archivo" | "error"
    """
    notif = perfil.get("notificacion", {})

    if not notif.get("email_destino"):
        return  # sin destinatario configurado

    if estado == "ok" and not notif.get("en_exito"):
        return
    if estado in ("error", "sin_archivo") and not notif.get("en_error"):
        return

    asunto, cuerpo = _construir_mensaje(perfil["nombre"], estado, filas, archivo, error)
    _enviar(notif, asunto, cuerpo)


def _construir_mensaje(nombre_perfil, estado, filas, archivo, error):
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")

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

    return asunto, cuerpo


def _enviar(notif: dict, asunto: str, cuerpo_html: str):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = notif.get("smtp_user", "")
    msg["To"]      = notif["email_destino"]
    msg.attach(MIMEText(cuerpo_html, "html", "utf-8"))

    host = notif.get("smtp_host", "")
    port = int(notif.get("smtp_port", 587))
    user = notif.get("smtp_user", "")
    pwd  = notif.get("smtp_pass", "")

    if not host:
        raise ValueError("SMTP no configurado. Completa el host en la configuración del perfil.")

    with smtplib.SMTP(host, port, timeout=15) as server:
        server.ehlo()
        server.starttls()
        if user and pwd:
            server.login(user, pwd)
        server.sendmail(user, notif["email_destino"], msg.as_string())
