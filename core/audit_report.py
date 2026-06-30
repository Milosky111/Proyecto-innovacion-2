# core/audit_report.py
"""
Genera un informe de auditoría en PDF a partir del historial de ejecuciones
(core/logger.py). Pensado para imprimir o presentar: resumen ejecutivo,
totales por estado, desglose por perfil, y detalle cronológico.
"""

import os
from datetime import datetime

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

from core.logger import RunLogger

ACCENT = colors.HexColor("#1F3864")
ACCENT_LIGHT = colors.HexColor("#D5E8F0")
SUCCESS = colors.HexColor("#1E8449")
ERROR = colors.HexColor("#C0392B")
WARNING = colors.HexColor("#D4AC0D")
TEXT_MUTED = colors.HexColor("#5A6B78")
ROW_EVEN = colors.HexColor("#EDF3FB")

ESTADO_COLOR = {"ok": SUCCESS, "error": ERROR, "sin_archivo": WARNING}
ESTADO_LABEL = {"ok": "Exitosa", "error": "Error", "sin_archivo": "Sin archivo"}


def generar_informe(logger: RunLogger, ruta_salida: str,
                     desde: str = None, hasta: str = None,
                     titulo: str = "Informe de Auditoría — Extractor de Datos") -> str:
    """
    Genera el PDF de auditoría en ruta_salida y retorna esa misma ruta.
    desde/hasta acotan el rango de fechas (formato ISO, ej "2026-06-01");
    si se omiten, se usa todo el historial disponible.
    """
    resumen = logger.obtener_resumen(desde=desde, hasta=hasta)
    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)) or ".", exist_ok=True)

    doc = SimpleDocTemplate(
        ruta_salida, pagesize=letter,
        topMargin=2 * cm, bottomMargin=2 * cm,
        leftMargin=2 * cm, rightMargin=2 * cm
    )

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(
        "TituloInforme", parent=styles["Title"],
        textColor=ACCENT, fontSize=18, spaceAfter=4
    )
    style_subtitle = ParagraphStyle(
        "Subtitulo", parent=styles["Normal"],
        textColor=TEXT_MUTED, fontSize=10, spaceAfter=16, fontName="Helvetica-Oblique"
    )
    style_h2 = ParagraphStyle(
        "Seccion", parent=styles["Heading2"],
        textColor=ACCENT, fontSize=13, spaceBefore=18, spaceAfter=8
    )
    style_body = ParagraphStyle(
        "Cuerpo", parent=styles["Normal"], fontSize=10, leading=14
    )

    story = []

    # ── Cabecera ──────────────────────────────────────────────────────────
    story.append(Paragraph(titulo, style_title))
    rango_txt = _formatear_rango(desde, hasta)
    generado = datetime.now().strftime("%d/%m/%Y %H:%M")
    story.append(Paragraph(f"{rango_txt} · Generado el {generado}", style_subtitle))

    # ── Resumen ejecutivo ────────────────────────────────────────────────
    story.append(Paragraph("Resumen ejecutivo", style_h2))

    totales = resumen["totales_estado"]
    ok = totales.get("ok", 0)
    error = totales.get("error", 0)
    sin_archivo = totales.get("sin_archivo", 0)
    total = resumen["total_ejecuciones"]
    tasa_exito = f"{(ok / total * 100):.0f}%" if total else "—"

    data_resumen = [
        ["Total de ejecuciones", "Exitosas", "Con error", "Sin archivo", "Tasa de éxito"],
        [str(total), str(ok), str(error), str(sin_archivo), tasa_exito],
    ]
    t_resumen = Table(data_resumen, colWidths=[3.2 * cm] * 5)
    t_resumen.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 1), (-1, 1), ROW_EVEN),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (-1, 1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D8E4")),
    ]))
    story.append(t_resumen)

    if total == 0:
        story.append(Spacer(1, 16))
        story.append(Paragraph(
            "No hay ejecuciones registradas en el rango seleccionado.", style_body
        ))
        doc.build(story)
        return ruta_salida

    # ── Desglose por perfil ──────────────────────────────────────────────
    story.append(Paragraph("Desglose por automatización", style_h2))

    data_perfiles = [["Automatización", "Ejecuciones", "Exitosas", "Fallidas",
                       "Filas exportadas", "Última ejecución"]]
    for p in resumen["por_perfil"]:
        data_perfiles.append([
            p["perfil_nombre"] or "(sin nombre)",
            str(p["ejecuciones"]),
            str(p["exitosas"]),
            str(p["fallidas"]),
            str(p["filas_totales"] or 0),
            _fmt_fecha(p["ultima_ejecucion"]),
        ])

    t_perfiles = Table(data_perfiles, colWidths=[4.5 * cm, 2.2 * cm, 2.0 * cm, 2.0 * cm, 2.8 * cm, 3.3 * cm])
    estilo_perfiles = [
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D8E4")),
    ]
    for i in range(1, len(data_perfiles)):
        if i % 2 == 0:
            estilo_perfiles.append(("BACKGROUND", (0, i), (-1, i), ROW_EVEN))
    t_perfiles.setStyle(TableStyle(estilo_perfiles))
    story.append(t_perfiles)

    # ── Detalle cronológico ───────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("Detalle de ejecuciones", style_h2))
    story.append(Paragraph(
        f"Listado completo, de la más reciente a la más antigua "
        f"({len(resumen['ejecuciones'])} registros).",
        style_subtitle
    ))

    data_detalle = [["Fecha/Hora", "Automatización", "Estado", "Filas", "Archivo / Detalle"]]
    for e in resumen["ejecuciones"]:
        detalle = e["archivo_usado"] or e["detalle_error"] or "—"
        if len(detalle) > 48:
            detalle = detalle[:45] + "…"
        data_detalle.append([
            _fmt_fecha(e["timestamp"]),
            (e["perfil_nombre"] or "—")[:22],
            ESTADO_LABEL.get(e["estado"], e["estado"]),
            str(e["filas_export"]),
            detalle,
        ])

    t_detalle = Table(data_detalle, colWidths=[2.8 * cm, 3.5 * cm, 2.2 * cm, 1.5 * cm, 6.8 * cm], repeatRows=1)
    estilo_detalle = [
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D0D8E4")),
    ]
    for i, e in enumerate(resumen["ejecuciones"], start=1):
        color = ESTADO_COLOR.get(e["estado"], TEXT_MUTED)
        estilo_detalle.append(("TEXTCOLOR", (2, i), (2, i), color))
        estilo_detalle.append(("FONTNAME", (2, i), (2, i), "Helvetica-Bold"))
        if i % 2 == 0:
            estilo_detalle.append(("BACKGROUND", (0, i), (-1, i), ROW_EVEN))
    t_detalle.setStyle(TableStyle(estilo_detalle))
    story.append(t_detalle)

    doc.build(story)
    return ruta_salida


def _formatear_rango(desde, hasta):
    if desde and hasta:
        return f"Periodo: {desde} a {hasta}"
    if desde:
        return f"Desde: {desde}"
    if hasta:
        return f"Hasta: {hasta}"
    return "Periodo: histórico completo"


def _fmt_fecha(iso_ts):
    if not iso_ts:
        return "—"
    try:
        dt = datetime.fromisoformat(iso_ts)
        return dt.strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso_ts[:16]
