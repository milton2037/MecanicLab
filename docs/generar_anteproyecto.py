# -*- coding: utf-8 -*-
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Ellipse, FancyArrowPatch, Rectangle
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Inches, Emu
from docx.oxml import OxmlElement

ROOT = Path(__file__).resolve().parent
FIG = ROOT / "figuras"
FIG.mkdir(exist_ok=True)
DOCX = ROOT / "Anteproyecto_MecanicLab.docx"

INK = "#1a1a1a"
GRAY = "#4a4a4a"
LINE = "#2b2b2b"


def set_run_font(run, name="Times New Roman", size=12, bold=False, italic=False, color=INK):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color.replace("#", ""))


def add_p(doc, text, *, size=12, bold=False, italic=False, align="justify", space_after=8, first_line=True, space_before=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if first_line and align == "justify":
        pf.first_line_indent = Cm(1.25)
    if align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        set_run_font(run, size=14 if level == 1 else 13, bold=True, color="1A1A1A")
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(8)
    return h


def caption(doc, text):
    p = add_p(doc, text, size=11, italic=True, align="center", first_line=False, space_after=12)
    return p


def add_hyper(p, text, url):
    part = p.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1A1A1A")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "22")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rPr.append(rFonts)
    rPr.append(u)
    rPr.append(color)
    rPr.append(sz)
    i = OxmlElement("w:i")
    rPr.append(i)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    p._p.append(hyperlink)


def shade_cell(cell, hex_color):
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_text(cell, text, *, bold=False, size=10, center=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)


def draw_actors():
    fig, ax = plt.subplots(figsize=(11.2, 7.4), dpi=160)
    ax.set_xlim(0, 11.2)
    ax.set_ylim(0, 7.4)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    border = Rectangle((2.55, 0.45), 8.35, 6.5, fill=False, lw=1.1, ec=LINE)
    ax.add_patch(border)
    ax.text(6.7, 6.72, "MecanicLab", ha="center", va="center", fontsize=11, fontweight="bold", color=INK)

    actors = [
        (0.35, 6.35, "Cliente"),
        (0.35, 5.35, "Recepcionista"),
        (0.35, 4.35, "Inspector CDA"),
        (0.35, 3.35, "Director técnico"),
        (0.35, 2.35, "Mecánico"),
        (0.35, 1.35, "Jefe de taller"),
        (0.35, 0.55, "Asesor / cajero"),
    ]
    for x, y, name in actors:
        ax.plot(x + 0.55, y + 0.38, "o", ms=9, color=INK, mfc="white", mew=1.1)
        ax.plot([x + 0.55, x + 0.55], [y + 0.28, y + 0.05], color=INK, lw=1.1)
        ax.plot([x + 0.38, x + 0.72], [y + 0.18, y + 0.18], color=INK, lw=1.1)
        ax.plot([x + 0.55, x + 0.42], [y + 0.05, y - 0.18], color=INK, lw=1.1)
        ax.plot([x + 0.55, x + 0.68], [y + 0.05, y - 0.18], color=INK, lw=1.1)
        ax.text(x + 1.15, y + 0.12, name, fontsize=8, color=INK, va="center")

    cases = [
        (3.4, 6.05, "Agendar cita"),
        (5.7, 6.05, "Consultar estado"),
        (8.0, 6.05, "Autorizar / pagar"),
        (3.4, 4.85, "Recepcionar ingreso"),
        (5.7, 4.85, "Código único / QR"),
        (8.0, 4.85, "Notificar cliente"),
        (3.4, 3.65, "Inspección TM"),
        (5.7, 3.65, "Emitir FUR"),
        (8.0, 3.65, "Reportar RUNT/SICOV"),
        (3.4, 2.45, "Diagnosticar"),
        (5.7, 2.45, "Cotizar ítems"),
        (8.0, 2.45, "Kanban / OT"),
        (3.4, 1.25, "Inventario"),
        (5.7, 1.25, "Facturar por negocio"),
        (8.0, 1.25, "Auditar roles"),
    ]
    for x, y, label in cases:
        e = Ellipse((x + 0.95, y), 1.95, 0.62, facecolor="white", edgecolor=LINE, lw=1.0)
        ax.add_patch(e)
        ax.text(x + 0.95, y, label, ha="center", va="center", fontsize=7.2, color=INK)

    links = [
        (1.85, 6.45, 3.4, 6.05),
        (1.85, 6.45, 5.7, 6.05),
        (1.85, 6.45, 8.0, 6.05),
        (1.85, 5.45, 3.4, 4.85),
        (1.85, 5.45, 5.7, 4.85),
        (1.85, 4.45, 3.4, 3.65),
        (1.85, 3.45, 5.7, 3.65),
        (1.85, 3.45, 8.0, 3.65),
        (1.85, 2.45, 3.4, 2.45),
        (1.85, 1.45, 5.7, 2.45),
        (1.85, 1.45, 8.0, 2.45),
        (1.85, 0.65, 3.4, 1.25),
        (1.85, 0.65, 5.7, 1.25),
        (1.85, 5.45, 8.0, 4.85),
        (1.85, 0.65, 8.0, 1.25),
    ]
    for x1, y1, x2, y2 in links:
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-", color=GRAY, lw=0.7))

    ax.text(5.6, 0.18, "Figura 1. Diagrama de actores y casos de uso (dinámico).", ha="center", fontsize=8, color=GRAY)
    fig.tight_layout()
    out = FIG / "diagrama_actores.png"
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close()
    return out


def class_box(ax, x, y, w, h, title, attrs):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="square,pad=0", facecolor="white", edgecolor=LINE, lw=0.9))
    ax.plot([x, x + w], [y + h - 0.28, y + h - 0.28], color=LINE, lw=0.8)
    ax.text(x + w / 2, y + h - 0.15, title, ha="center", va="center", fontsize=7.2, fontweight="bold", color=INK)
    ax.text(x + 0.08, y + h - 0.42, attrs, ha="left", va="top", fontsize=6.1, color=INK, linespacing=1.25)


def draw_classes():
    fig, ax = plt.subplots(figsize=(11.2, 7.6), dpi=160)
    ax.set_xlim(0, 11.2)
    ax.set_ylim(0, 7.6)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    class_box(ax, 0.25, 5.7, 2.15, 1.55, "Usuario", "id\nusername\narea\nrol")
    class_box(ax, 2.6, 5.7, 2.15, 1.55, "Cliente", "id\nidentidad\ncontacto\notp")
    class_box(ax, 4.95, 5.7, 2.15, 1.55, "Vehiculo", "id\nplaca\nmarca\nkm")
    class_box(ax, 7.3, 5.7, 2.15, 1.55, "Sede", "id\nnombre\ncda_activo\ntaller_activo")
    class_box(ax, 0.25, 3.55, 2.4, 1.7, "Cita", "id\nfecha/hora\ntipo_servicio\nestado")
    class_box(ax, 2.9, 3.55, 2.55, 1.7, "Ingreso", "codigo_unico\nqr\nmixto\nchequeo")
    class_box(ax, 5.7, 3.55, 2.4, 1.7, "Seguimiento", "negocio\nestado\nvisible_cliente")
    class_box(ax, 8.35, 3.55, 2.5, 1.7, "Pago/Factura", "negocio\nmedio\ncomprobante\nFE")
    class_box(ax, 0.4, 1.15, 2.5, 1.85, "InspeccionCDA", "pruebas\ndefecto A/B\nequipo")
    class_box(ax, 3.15, 1.15, 2.5, 1.85, "ResultadoFUR", "numero_fur\naprobado\nrunt/sicov")
    class_box(ax, 5.9, 1.15, 2.5, 1.85, "DiagnosticoOT", "informe\nhallazgos\netapa kanban")
    class_box(ax, 8.6, 1.15, 2.3, 1.85, "Inventario", "sku\nstock\nreserva_web")

    def assoc(x1, y1, x2, y2):
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-", color=GRAY, lw=0.8))

    assoc(1.32, 5.7, 1.45, 5.25)
    assoc(3.67, 5.7, 4.15, 5.25)
    assoc(6.02, 5.7, 4.3, 5.25)
    assoc(8.37, 5.7, 6.9, 5.25)
    assoc(1.45, 3.55, 1.65, 3.0)
    assoc(4.15, 3.55, 4.4, 3.0)
    assoc(6.9, 3.55, 7.15, 3.0)
    assoc(4.15, 3.55, 1.65, 3.0)
    assoc(4.15, 3.55, 9.75, 3.0)

    ax.text(5.6, 0.35, "Figura 2. Diagrama de clases simplificado (estático). Núcleo compartido; CDA y Taller no se asocian entre sí.",
            ha="center", fontsize=8, color=GRAY)
    fig.tight_layout()
    out = FIG / "diagrama_clases.png"
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close()
    return out


def draw_timeline():
    fig, ax = plt.subplots(figsize=(10.8, 3.6), dpi=160)
    fases = [
        ("Levantamiento\ny arquitectura", 0, 2),
        ("Vistas Portal\ny Core", 2, 4.2),
        ("Vistas CDA\ny Taller", 4.2, 6.2),
        ("Modelo de datos\ny prototipo", 6.2, 8.2),
        ("Validación\ndocente", 8.2, 10),
    ]
    ax.set_xlim(-0.2, 10.4)
    ax.set_ylim(0, 3.2)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.plot([0, 10], [1.35, 1.35], color=LINE, lw=1.2)
    for i, (name, a, b) in enumerate(fases):
        ax.add_patch(Rectangle((a, 1.12), b - a - 0.08, 0.46, facecolor="white", edgecolor=LINE, lw=1))
        ax.text((a + b) / 2, 1.35, str(i + 1), ha="center", va="center", fontsize=9, fontweight="bold")
        ax.text((a + b) / 2, 2.15, name, ha="center", va="bottom", fontsize=7.4, color=INK)
    ax.text(0, 0.55, "Sem. 1-2", fontsize=7.2, color=GRAY)
    ax.text(9.15, 0.55, "Cierre corte", fontsize=7.2, color=GRAY, ha="right")
    ax.text(5, 0.2, "Figura 3. Línea de tiempo del anteproyecto (actividades).", ha="center", fontsize=8, color=GRAY)
    fig.tight_layout()
    out = FIG / "linea_tiempo.png"
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close()
    return out


def build_doc(img_actores, img_clases, img_tiempo):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2.5)
        section.page_width = Cm(21.59)
        section.page_height = Cm(27.94)

        header = section.header.paragraphs[0]
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = header.add_run("FESC — Metodologías de software · Anteproyecto MecanicLab")
        set_run_font(r, size=9, italic=True, color="4A4A4A")

        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = footer.add_run("Documento de anteproyecto · uso académico")
        set_run_font(r2, size=9, color="4A4A4A")

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    add_p(doc, "FUNDACIÓN DE ESTUDIOS SUPERIORES COMFANORTE — FESC", size=12, bold=True, align="center", first_line=False, space_after=2)
    add_p(doc, "Programa académico de Metodologías de Software", size=12, align="center", first_line=False, space_after=2)
    add_p(doc, "San José de Cúcuta, Norte de Santander, Colombia", size=12, align="center", first_line=False, space_after=18)

    add_p(doc, "ANTEPROYECTO", size=14, bold=True, align="center", first_line=False, space_after=10)
    add_p(
        doc,
        "MecanicLab: sistema de información para la gestión integrada de un Centro de Diagnóstico Automotor (CDA) y un taller de servicio automotriz, con núcleo compartido y portal del cliente",
        size=13,
        bold=True,
        align="center",
        first_line=False,
        space_after=16,
    )
    add_p(doc, "Equipo de desarrollo: integrantes del curso de Metodologías de Software", size=11, italic=True, align="center", first_line=False, space_after=4)
    add_p(doc, "Octubre de 2026", size=11, align="center", first_line=False, space_after=20)

    heading(doc, "1. Título del proyecto (anteproyecto)", 1)
    add_p(
        doc,
        "MecanicLab es un sistema web que articula cuatro bloques: Portal del cliente, Núcleo compartido (Core), CDA y Taller. Los dos últimos son independientes entre sí y solo dependen del Core. El título expresa tanto el objeto tecnológico (sistema de información) como el dominio normativo y operativo (inspección técnico-mecánica y reparación).",
    )

    heading(doc, "1.1 Qué: problemática", 2)
    add_p(
        doc,
        "En Colombia, la revisión técnico-mecánica y de emisiones contaminantes (RTMyEC) es un trámite obligatorio para gran parte del parque automotor y debe realizarse en un CDA acreditado. En paralelo, el mismo vehículo suele requerir diagnóstico, cotización, autorización del propietario, repuestos y pago en un taller. En la práctica, estos dos negocios se gestionan con papeles, hojas de cálculo o software aislado: el cliente no ve el estado del proceso en tiempo real; el código de seguimiento, si existe, es secuencial y fácil de adivinar; las evidencias (fotos, videos, FUR) no quedan ligadas a un ingreso único; CDA y taller duplican datos del propietario y de la placa; y cada negocio factura sin una hoja de vida vehicular común. El resultado es opacidad para el ciudadano, retrabajo para el personal y riesgo de incumplimiento frente a RUNT, SICOV y la facturación electrónica.",
    )

    heading(doc, "1.2 Por qué: justificación breve", 2)
    add_p(
        doc,
        "Se justifica el desarrollo porque la fragmentación de la información aumenta tiempos de espera, errores de recepción y desconfianza del cliente, en un país donde más de la mitad de los vehículos obligados a RTM no tienen el certificado vigente (RUNT, 2024/2025). Un núcleo compartido con código único no secuencial, dos líneas de tiempo (CDA y taller) y roles auditables permite trazabilidad sin mezclar las operaciones de inspección y de reparación.",
    )

    heading(doc, "1.3 Cómo: plataformas (medio de desarrollo)", 2)
    add_p(
        doc,
        "El medio de desarrollo es una aplicación web. El backend se implementa en Java con Spring Boot (versión 4.x) y el empaquetado se realiza con Gradle. Las vistas se construyen con HTML, Thymeleaf y CSS: un tema oscuro para el sistema interno (Core, CDA y Taller) y un tema claro e independiente para el portal del cliente. El modelado estructural se documenta en dbdiagram.io y el comportamiento (actores y casos de uso) en un tablero Figma. El control de versiones y la colaboración se realizan en repositorio Git; la evidencia del curso se adjunta en Atlassian (Jira/Confluence), según indicación docente. No se emplea, en esta fase de anteproyecto, un instalador de escritorio: el despliegue local se hace con el servidor embebido (Tomcat) en el puerto 8080.",
    )

    heading(doc, "1.4 Cuándo: diagrama de actividades (línea de tiempo)", 2)
    add_p(
        doc,
        "El cronograma del anteproyecto se organiza en cinco frentes sucesivos, compatibles con un corte académico de metodologías: levantamiento y arquitectura de cuatro bloques; prototipo de portal y Core; prototipo de CDA y Taller; consolidación del modelo de datos; y validación documental. La Figura 3 resume la línea de tiempo.",
        first_line=True,
    )
    doc.add_picture(str(img_tiempo), width=Inches(6.1))
    last = doc.paragraphs[-1]
    last.alignment = WD_ALIGN_PARAGRAPH.CENTER

    tabla = doc.add_table(rows=6, cols=3)
    tabla.style = "Table Grid"
    headers = ["Fase", "Actividades", "Entregable"]
    for i, h in enumerate(headers):
        set_cell_text(tabla.rows[0].cells[i], h, bold=True, size=10, center=True)
        shade_cell(tabla.rows[0].cells[i], "E8E8E8")
    filas = [
        ("1. Levantamiento", "Problema, actores, estados CDA/Taller, alcance sin persistencia inicial", "Arquitectura de módulos"),
        ("2. Portal y Core", "Login, mapa interno, citas, seguimiento, recepción, pagos (vistas)", "Prototipo navegable interno/cliente"),
        ("3. CDA y Taller", "Pistas, inspección, FUR, diagnóstico, Kanban, inventario (vistas)", "Prototipo de ambos negocios"),
        ("4. Datos", "Adaptación del modelo relacional al dominio MecanicLab", "Diagrama de clases / DBML"),
        ("5. Cierre de corte", "Anteproyecto, bibliografía APA, carga en Atlassian", "Documento formal"),
    ]
    for r, fila in enumerate(filas, start=1):
        for c, val in enumerate(fila):
            set_cell_text(tabla.rows[r].cells[c], val, size=10)
    doc.add_paragraph()

    heading(doc, "1.5 Dónde: lugar geográfico", 2)
    add_p(
        doc,
        "El proyecto se formula en San José de Cúcuta, capital de Norte de Santander, Colombia, en el marco académico de la Fundación de Estudios Superiores Comfanorte (FESC). El dominio de aplicación es el parque automotor que circula en el área metropolitana (Cúcuta, Los Patios, Villa del Rosario y municipios aledaños) y que debe cumplir RTMyEC ante CDA acreditados, además de atenderse en talleres de mecánica ligera y mediana. Fuentes locales estiman un parque departamental del orden de 170.000 a 180.000 vehículos matriculados, con un volumen relevante de automotores de placa extranjera registrados en Cúcuta (La Opinión, 2024/2025). El sistema se concibe para operar en sedes urbanas con conectividad a internet, sin restringirse a un único predio: el Core admite varias sedes, pistas de CDA y bahías de taller.",
    )

    heading(doc, "1.6 Quiénes: cliente final", 2)
    add_p(
        doc,
        "El cliente final es doble. Primero, el ciudadano propietario o conductor que agenda, consulta con código y placa (sin sesión), autoriza ítems y paga. Segundo, la organización prestadora: personal de recepción, inspectores y director técnico del CDA, mecánicos, jefe de taller, asesor comercial y cajero, todos bajo roles del Core. El usuario académico (docente y pares) es destinatario de la evidencia metodológica, no del servicio productivo.",
    )

    heading(doc, "1.7 Para qué: solución", 2)
    add_p(
        doc,
        "Para unificar, sin fusionar indebidamente, la operación de CDA y taller alrededor de un ingreso con código único alfanumérico o QR no secuencial; publicar estados al cliente en tiempo real; conservar evidencias, costos y pagos bajo autenticación; facturar por negocio; y dejar una hoja de vida del vehículo con origen (CDA, taller o mixto). La solución no sustituye a RUNT ni a SICOV: se alinea a ellos como destinos de reporte del FUR.",
    )

    heading(doc, "2. Descripción del problema", 1)
    add_p(
        doc,
        "El problema se dimensiona con cifras oficiales, hechos normativos y prácticas observadas en la operación de CDA y talleres.",
    )
    add_p(
        doc,
        "Hecho 1 — magnitud del parque y de la evasión de RTM. Al 31 de diciembre de 2024, el RUNT reportó 19.972.482 vehículos registrados en Colombia (62 % motocicletas). De 17.191.567 unidades obligadas a revisión técnico-mecánica, 9.641.743 (56,1 %) no tenían certificado vigente y 7.549.824 (43,9 %) sí lo tenían (RUNT, 2025). En 2023 la proporción de no vigencia era del 57 % sobre 15.904.250 vehículos obligados (RUNT, 2024). Es decir, más de la mitad del parque obligado opera sin evidencia vigente de aptitud técnica y de emisiones. Un sistema que reduzca fricción de agendamiento, reinspección y comunicación no elimina la evasión, pero sí ataca una causa operativa: procesos opacos, filas y pérdida de confianza.",
    )
    add_p(
        doc,
        "Hecho 2 — el CDA es un organismo de apoyo altamente regulado. La RTMyEC se realiza en CDA acreditados por ONAC y registrados en el RUNT. La Superintendencia de Transporte precisa que el pago del servicio no obliga a expedir certificado: este depende de aprobar los ítems evaluados; al valor del CDA se suman IVA, tarifa RUNT, servicio SICOV y recaudo de la Agencia Nacional de Seguridad Vial (Superintendencia de Transporte, 2022). El registro de un CDA exige, entre otros, acreditación como organismo de inspección, contrato SICOV, póliza, personal formado (mínimo 155 horas) y responsable autorizado para firmar el Formato Uniforme de Resultados — FUR (Ministerio de Transporte, 2020, Resolución 20203040011355, compilada en la Resolución 20223040045295 de 2022). Un software de taller genérico no cubre pistas, defectos tipo A o B, calibración de equipos ni reporte a RUNT/SICOV.",
    )
    add_p(
        doc,
        "Hecho 3 — coexistencia de dos lógicas de negocio sobre el mismo vehículo. El cliente frecuentemente lleva el automotor a inspección y, el mismo día o en reinspección, a reparación (gases, luces, frenos, clutch). Si cada negocio tiene su propia ficha, se duplican placa, propietario y evidencias; si se usa un consecutivo correlativo, el tercero puede inferir el estado de otros ingresos. MecanicLab plantea un ingreso mixto con un solo código y dos líneas de tiempo, publicadas por inspector/director técnico (CDA) y por mecánico/jefe de taller/sistema (taller).",
    )
    add_p(
        doc,
        "Hecho 4 — asimetría de información con el propietario. Sin portal, el cliente desconoce si el vehículo está en inspección, esperando autorización, esperando repuesto o listo para entrega. Las cotizaciones llegan por WhatsApp informal, sin registro de consentimiento ni decisión ítem a ítem. Las evidencias (fotos de hallazgo) no quedan bloqueadas para el anónimo: o se ocultan todas o se exponen de más. El problema no es solo “falta de un sistema”, sino falta de un contrato de visibilidad: línea de tiempo pública por código; costos, evidencias y pagos solo con cuenta (OTP).",
    )
    add_p(
        doc,
        "Hecho 5 — territorio. Norte de Santander concentra un parque mixto (nacional y de placa extranjera) con presión sobre citas, impuestos y circulación urbana (La Opinión, 2024/2025). Cúcuta, como ciudad de frontera, incrementa la demanda de servicios automotrices y de inspección. Un prototipo académico situado aquí es pertinente y transferible a otras sedes del país.",
    )

    heading(doc, "3. Objetivo general", 1)
    add_p(
        doc,
        "Diseñar e implementar el prototipo de un sistema de información web (MecanicLab) que gestione de forma integrada, mediante un núcleo compartido, los procesos de un Centro de Diagnóstico Automotor y de un taller de servicio automotriz, ofreciendo al cliente trazabilidad por código único y al personal control de roles, estados y evidencias, sin que CDA y Taller dependan entre sí.",
    )

    heading(doc, "3.1 Objetivos específicos", 2)
    objetivos = [
        "Especificar la arquitectura de cuatro bloques (Portal, Core, CDA y Taller) y los estados de proceso visibles al cliente, a partir de la normativa de RTMyEC y de la operación de un taller.",
        "Construir las vistas del portal del cliente (agendamiento, consulta por código y placa, aprobaciones, tienda de repuestos, pagos y cuenta con OTP) con una identidad visual distinta a la del sistema interno.",
        "Construir las vistas del núcleo (usuarios y seguridad, clientes y vehículos, sedes y agenda, recepción con código único, seguimiento, notificaciones, pagos y facturación, hoja de vida y reportes).",
        "Construir las vistas del CDA (programación de pista, inspección con defectos A/B, equipos y calibración, FUR, RUNT/SICOV y reinspección) y del taller (diagnóstico, cotización ítem a ítem, Kanban, bahías, inventario y proveedores).",
        "Documentar el modelo estructural (clases / base de datos) y el modelo de comportamiento (actores y casos de uso), y consolidar el anteproyecto con fuentes en normas APA para su archivo en Atlassian.",
    ]
    for i, o in enumerate(objetivos, 1):
        add_p(doc, f"{i}. {o}")

    heading(doc, "4. Antecedentes", 1)
    add_p(
        doc,
        "Los antecedentes se organizan en tres capas: marco normativo colombiano; literatura académica sobre gestión de servicios automotrices y sistemas de información vehicular; y soluciones de software existentes, para delimitar el vacío que ocupa MecanicLab.",
    )
    add_p(
        doc,
        "En el plano normativo, la Ley 769 de 2002 (Código Nacional de Tránsito) y sus modificaciones establecen la RTM y el papel de los organismos de apoyo. El Ministerio de Transporte reglamentó el registro de CDA ante el RUNT (Resolución 20203040011355 de 2020, compilada en 2022), con referencias a las NTC 5375, 5385, 6218 y 6282 y a la operación de SICOV. La Superintendencia de Transporte publica la Guía de revisión segura y tableros de control de CDA alimentados por el RUNT (Superintendencia de Transporte, 2022, 2025). ASOCDA consolida boletines de certificados expedidos y de evasión de RTMyEC. Estos documentos no son software, pero definen qué debe quedar trazado: FUR, responsable de firma, equipos, reinspección y reporte.",
    )
    add_p(
        doc,
        "En el plano académico, Vásquez Fajardo et al. (2024), en revista indexada y disponible en SciELO, analizan la gestión comercial y administrativa de una empresa de servicio automotriz (lavado y engrase) y evidencian que la incertidumbre en la toma de decisiones y la débil orientación al cliente degradan la operación. Aunque el caso no es un CDA, sí respalda la tesis de que los servicios vehiculares requieren control de procesos, participación de actores y seguimiento. Velasco Pinzón y Huérfano Daza (2017), en trabajo de grado de la Universidad Distrital Francisco José de Caldas (repositorio institucional; línea de ingeniería telemática), proponen un sistema telemático para consultar información vehicular por servicios web, con identificación de propietario y placa. Ese antecedente es próximo al módulo de consulta de MecanicLab, pero no cubre inspección normativa, Kanban de taller ni facturación por negocio. La literatura de mantenimiento informático sobre arquitecturas web muestra la viabilidad de registrar estados y evidencias de equipos, sin resolver el dualismo CDA–taller (Velásquez Rodríguez y Custodio Ruiz, 2011).",
    )
    add_p(
        doc,
        "En el mercado colombiano existen DMS y software de taller (órdenes de trabajo, inventario, factura electrónica DIAN, historial por placa). Productos como plataformas de concesionario o de taller digitalizan la OT y el stock, pero se orientan al taller o al concesionario, no a la cola de pistas, defectos A/B, FUR y SICOV. Inversamente, los sistemas de CDA suelen ser cerrados y acoplados a la línea de inspección, sin portal de autorización de cotizaciones ni bodega de repuestos. El vacío identificado es precisamente el de MecanicLab: un Core compartido con dos negocios independientes y un portal ciudadano con reglas de visibilidad.",
    )
    add_p(
        doc,
        "Desde la ingeniería de software, el anteproyecto se apoya en el modelado UML de comportamiento (actores y casos de uso) y de estructura (clases), tal como lo formulan Booch, Rumbaugh y Jacobson, y en el ciclo de prototipado evolutivo descrito por Sommerville (2016): primero vistas y arquitectura, después persistencia. Esa decisión metodológica explica por qué el prototipo actual navega módulos con datos de ejemplo y deja la base de datos como siguiente incremento.",
    )

    heading(doc, "5. Justificación", 1)
    add_p(
        doc,
        "¿Por qué desarrollar esta problemática? Porque el país tiene un parque masivo con alta no vigencia de RTM, un CDA sujeto a acreditación y reporte externo, y un taller que convive con el mismo cliente sin un lenguaje de estados común. Digitalizar solo el taller deja ciego el proceso de inspección; digitalizar solo el CDA deja al propietario sin cotización ni inventario. La frontera cucuteña agrava la densidad de atenciones y la necesidad de códigos no correlativos y de historial.",
    )
    add_p(
        doc,
        "¿Para qué? Para que el cliente sepa, con un código y una placa, en qué etapa está cada negocio; para que autorizar un ítem deje rastro; para que el inspector publique “en inspección” o “rechazado” y el mecánico publique “esperando repuestos y pago”, sin que un módulo invada al otro; para que la factura electrónica salga a nombre del negocio que prestó el servicio; y para formar, en un curso de metodologías, un artefacto completo (problema, objetivos, UML, prototipo y referencias APA).",
    )
    add_p(
        doc,
        "Resultados esperados. (a) Prototipo web navegable de los cuatro bloques, con estilos diferenciados interno/cliente. (b) Código único de seguimiento y mapa de estados CDA y taller. (c) Modelo de datos adaptado al dominio (usuarios, roles, clientes, vehículos, ingresos, inspección, OT, inventario, pagos). (d) Documentación de anteproyecto lista para Atlassian. (e) Base para un segundo incremento: persistencia, autenticación real, pasarela y conectores RUNT/SICOV. No se espera, en este corte, la certificación ONAC del CDA ni la habilitación DIAN en producción: eso excede el alcance académico.",
    )

    heading(doc, "6. Metodología", 1)
    add_p(
        doc,
        "El trabajo se desarrolló con un enfoque de prototipado evolutivo e ingeniería dirigida por modelos. Primero se descompuso el dominio en cuatro bloques y se fijó la regla de independencia CDA–Taller. Después se implementaron controladores y plantillas (Spring Boot + Thymeleaf) para validar la navegación con el docente y el equipo, sin base de datos. En paralelo se construyeron dos vistas UML exigidas por la guía: una dinámica (actores y casos de uso) y una estática (clases / esquema de datos). Las herramientas de modelado oficiales del equipo son Figma (comportamiento) y dbdiagram.io (estructura); en este documento se reproducen versiones de trabajo en escala de grises y se citan los enlaces canónicos.",
    )
    add_p(doc, "Enlace del diagrama de actores y casos de uso (Figma):", first_line=False, space_after=2)
    p = add_p(doc, "", first_line=False, align="left", space_after=8)
    add_hyper(
        p,
        "https://www.figma.com/board/fJ1XVwWWbvuM8IKN7q7lXg/MecanicLab-Diagrama-de-Actores-y-Casos-de-Uso",
        "https://www.figma.com/board/fJ1XVwWWbvuM8IKN7q7lXg/MecanicLab-%E2%80%94-Diagrama-de-Actores-y-Casos-de-Uso?node-id=0-1",
    )
    add_p(doc, "Enlace del diagrama estructural / modelo de datos (dbdiagram.io):", first_line=False, space_after=2)
    p = add_p(doc, "", first_line=False, align="left", space_after=10)
    add_hyper(p, "https://dbdiagram.io/d/6ac637a69068dd16e269d18c", "https://dbdiagram.io/d/6ac637a69068dd16e269d18c")

    heading(doc, "6.1 Diagrama de actores (comportamiento) — dinámico", 2)
    add_p(
        doc,
        "El diagrama dinámico representa quién inicia qué caso de uso. El actor Cliente interactúa con el portal: agendar (sede, fecha, hora y tipo de servicio), consultar estado con código único y placa sin iniciar sesión, y —una vez autenticado con OTP— autorizar o rechazar ítems, pagar y ver evidencias. El Recepcionista registra el ingreso, genera el código no secuencial y la lista de chequeo. El Inspector y el Director técnico publican la línea CDA (inspección, FUR, RUNT/SICOV, reinspección). El Mecánico y el Jefe de taller publican la línea de taller (diagnóstico, Kanban, pruebas, listo para entrega). El Asesor convierte hallazgos en cotización; el Cajero y el módulo de pagos del Core registran caja o pasarela y emiten factura por negocio. El Administrador gestiona usuarios, roles y auditoría. Los sistemas RUNT y SICOV se modelan como actores externos de frontera, no como usuarios humanos.",
    )
    doc.add_picture(str(img_actores), width=Inches(6.2))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption(doc, "Figura 1. Diagrama de actores y casos de uso de MecanicLab (elaboración propia a partir del tablero Figma del equipo).")

    heading(doc, "6.2 Diagrama de clase (estructural) — estático", 2)
    add_p(
        doc,
        "El diagrama estático fija las clases persistentes y sus relaciones. El Core concentra Usuario, Rol, Cliente, Vehiculo, Sede, Recurso (pista o bahía), Cita, Ingreso (código único), SeguimientoEvento, Notificacion, Pago, Factura, HojaVidaVehiculo y ArchivoAdjunto. El CDA añade Inspeccion, InspeccionPrueba (defecto A/B), EquipoMedicion, ResultadoFUR y Reinspeccion. El Taller añade Diagnostico, Hallazgo, Cotizacion, OrdenTrabajo, Repuesto y OrdenCompra. No existe asociación directa CDA–Taller: ambos apuntan a Ingreso. Un ingreso mixto reutiliza el mismo codigo_unico y duplica eventos de seguimiento con atributo negocio = cda | taller. Los catálogos (estados, tipos de identidad, canales, medios de pago) evitan enumeraciones rígidas en código.",
    )
    doc.add_picture(str(img_clases), width=Inches(6.2))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption(doc, "Figura 2. Diagrama de clases simplificado (elaboración propia a partir del modelo en dbdiagram.io).")

    heading(doc, "7. Resultados", 1)
    add_p(
        doc,
        "Al cierre de este corte, el alcance alcanzado es el de un prototipo funcional de interfaz, no el de un sistema en producción con persistencia. Se considera un resultado válido de metodologías porque materializa arquitectura, actores, estados y pantallas verificables.",
    )
    add_p(
        doc,
        "Alcance logrado. (1) Aplicación Spring Boot ejecutable en http://localhost:8080. (2) Acceso interno con mapa de módulos y menú por bloque. (3) Portal del cliente con cabecera propia y hoja de estilo independiente (portal.css frente a app.css). (4) Pantallas de Core, CDA y Taller alineadas a los módulos del planteamiento (citas, seguimiento mixto, recepción con código de ejemplo K7M-9Q2X, inspección, FUR, Kanban, inventario, etc.). (5) Modelo de datos adaptado desde un esquema de recursos humanos hacia el dominio MecanicLab. (6) El presente anteproyecto, con bibliografía APA y diagramas UML.",
    )
    add_p(
        doc,
        "Respuestas del trabajo a la pregunta de investigación implícita —¿es viable un sistema que no fusione CDA y taller y aun así ofrezca una sola cara al cliente?—: sí, a nivel de diseño y prototipo de vistas. La viabilidad se demuestra porque el menú y los controladores separan /cda y /taller, el Core concentra /core, y el portal consulta ambas líneas bajo un código. Lo que aún no responde el prototipo es la persistencia transaccional, la autenticación OTP real, la pasarela y los web services de RUNT/SICOV; esas ausencias se declaran como trabajo futuro, no como defectos ocultos.",
    )
    add_p(
        doc,
        "Limitaciones. Datos de demostración en las plantillas; sin motor de base de datos conectado; sin pruebas de carga ni evaluación de usabilidad formal (ISO/IEC 25010) más allá de la revisión de navegación del equipo.",
    )

    heading(doc, "8. Bibliografía", 1)
    add_p(
        doc,
        "American Psychological Association. (2020). Publication manual of the American Psychological Association (7.ª ed.).",
        first_line=False,
        italic=False,
    )

    refs = [
        "Asociación Nacional de Centros de Diagnóstico Automotor. (s. f.). Cifras. https://www.aso-cda.org/cifras/",
        "Booch, G., Rumbaugh, J. y Jacobson, I. (2006). El lenguaje unificado de modelado (2.ª ed.). Pearson Addison-Wesley.",
        "Congreso de la República de Colombia. (2002). Ley 769 de 2002, por la cual se expide el Código Nacional de Tránsito Terrestre. Diario Oficial n.º 44.932.",
        "La Opinión. (2024). Van 30.000 embargos en Norte de Santander por no pagar impuesto vehicular. https://laopinion.co/economia/van-30000-embargos-en-norte-de-santander-por-no-pagar-impuesto-vehicular",
        "La Opinión. (2025). Impuesto a vehículos impulsa recaudación de Norte de Santander. https://laopinion.co/region/impuesto-vehiculos-impulsa-recaudacion-de-norte-de-santander",
        "Ministerio de Transporte. (2020). Resolución 20203040011355 de 2020, por la cual se reglamenta el registro de los Organismos de Apoyo al Tránsito ante el sistema RUNT. https://www.runt.gov.co/sites/default/files/normas/RESOLUCION%20ORGANISMOS%20DE%20APOYO%20AL%20TRANSITO%20f.pdf",
        "Ministerio de Transporte. (2022). Resolución 20223040045295 de 2022 (compilación de organismos de apoyo al tránsito).",
        "Registro Único Nacional de Tránsito. (2024). RUNT en cifras: balance del sector tránsito y transporte 2023. https://www.runt.gov.co/sites/default/files/documentos/Balance%20de%20cifras%20del%20sector%202023%20RUNT.pdf",
        "Registro Único Nacional de Tránsito. (2025). Balance del sector tránsito y transporte 2024. https://www.runt.gov.co/sites/default/files/documentos/Balance%20del%20Sector%20Tránsito%20y%20Transporte%202024.pdf",
        "Registro Único Nacional de Tránsito. (2025). Boletín de prensa 001 de 2025: balance del sector tránsito y transporte. https://www.runt.gov.co/sites/default/files/Boletín%20de%20Prensa%20001%20de%202025.pdf",
        "Velásquez Rodríguez, E. y Custodio Ruiz, Á. (2011). Sistema para la gestión del mantenimiento para un control supervisorio basado en software libre. Universidad, Ciencia y Tecnología, 15(59), 103-113. https://ve.scielo.org/scielo.php?script=sci_arttext&pid=S1316-48212011000200007",
        "Sommerville, I. (2016). Ingeniería de software (10.ª ed.). Pearson.",
        "Superintendencia de Transporte. (2022). Guía de revisión segura. https://www.supertransporte.gov.co/documentos/2022/Noviembre/Comunicaciones_16/GuiaRevisionSegura.pdf",
        "Superintendencia de Transporte. (2025). Tablero de control Centro de Diagnóstico Automotor CDA. https://transformaciondigital.supertransporte.gov.co/documentos/2025/tableros/TyT/Ficha-Tablero-CDA.pdf",
        "Vásquez Fajardo, C. E., Aguayo Rossignoli, L. F., Vásquez Fajardo, X. y Fajardo Vaca, L. M. (2024). Criterios estratégicos para el mejoramiento de la gestión comercial y administrativa en una empresa de servicio. Prohominum, 6(2), 60-76. https://doi.org/10.47606/ACVEN/PH0252",
        "Velasco Pinzón, E. D. y Huérfano Daza, J. A. (2017). Sistema telemático para la integración de información vehicular [Trabajo de grado, Universidad Distrital Francisco José de Caldas]. Repositorio institucional. https://hdl.handle.net/11349/5326",
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-1.25)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(ref)
        set_run_font(run, size=12)

    heading(doc, "Anexo. Carga en Atlassian", 1)
    add_p(
        doc,
        "El docente solicitó adjuntar el paquete en Atlassian. Se recomienda: (1) crear o abrir el espacio de Confluence (o el epic de Jira) del equipo MecanicLab; (2) adjuntar este archivo Word; (3) insertar los dos enlaces de diagramas (Figma y dbdiagram) como marcadores; (4) adjuntar capturas del prototipo (login, portal, dashboard, CDA, taller); (5) dejar comentario de revisión para el docente. El presente archivo es el entregable documental; Atlassian es el repositorio de evidencias del curso, no el entorno de ejecución del software.",
    )

    doc.save(DOCX)
    return DOCX


def main():
    a = draw_actors()
    c = draw_classes()
    t = draw_timeline()
    path = build_doc(a, c, t)
    print(path)


if __name__ == "__main__":
    main()
