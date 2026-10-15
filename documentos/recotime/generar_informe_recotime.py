#!/usr/bin/env python3
"""Informe de perfil RecoTime con membrete de The Company Dev."""

import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT.parent / "campus" / "assets"
OUT = ROOT / "Informe The Company Dev - Perfil RecoTime.docx"

NAVY = RGBColor(0x00, 0x30, 0x60)
BLACK = RGBColor(0x1B, 0x33, 0x25)
DARK = "1B3325"

# TRM del 14 de octubre de 2026, publicada por la Superintendencia Financiera
# y reproducida ese día en trmhoy.co.
TRM = 3181.74

KEYWORDS = (
    "EL CLIENTE",
    "EL CONTRATISTA",
    "THE COMPANY DEV",
    "RECUPERACO",
    "RECOROUTE",
    "RecuperaCo",
    "RecoRoute",
    "RECOTIME",
    "RecoTime",
    "SEO",
    "Seo",
    "CRM",
    "ERP",
)


def cop(value):
    amount = int(round(value))
    sign = "-" if amount < 0 else ""
    return sign + "$" + f"{abs(amount):,}".replace(",", ".")


def usd_to_cop(dollars):
    return dollars * TRM


def usd(dollars):
    if abs(dollars - round(dollars)) < 0.001:
        text = f"{int(round(dollars)):,}".replace(",", ".")
    else:
        text = f"{dollars:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "US$ " + text


def set_run_font(run, size=13, bold=False, color=None, name="Poppins"):
    run.bold = bold
    run.italic = False
    run.font.size = Pt(size)
    run.font.name = name
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def set_paragraph_format(paragraph, before=0, after=8, align="justify", keep_next=False):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = 1.08
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if align == "justify":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pPr = paragraph._p.get_or_add_pPr()
    if keep_next:
        if pPr.find(qn("w:keepNext")) is None:
            pPr.append(OxmlElement("w:keepNext"))
    if pPr.find(qn("w:widowControl")) is None:
        pPr.append(OxmlElement("w:widowControl"))


def split_keywords(text):
    pattern = re.compile("(" + "|".join(re.escape(word) for word in KEYWORDS) + ")")
    return [piece for piece in pattern.split(text) if piece]


def add_runs(paragraph, parts, size=13, color=None):
    for text, bold in parts:
        for piece in split_keywords(text):
            run = paragraph.add_run(piece)
            set_run_font(
                run,
                size=size,
                bold=bold or piece in KEYWORDS,
                color=color if color else BLACK,
            )


def add_body(doc, parts, size=13, before=0, after=8, align="justify", keep_next=False, color=None):
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, before=before, after=after, align=align, keep_next=keep_next)
    add_runs(paragraph, parts, size=size, color=color)
    return paragraph


def add_heading(doc, text):
    return add_body(
        doc,
        [(text, True)],
        size=13,
        before=12,
        after=4,
        align="left",
        keep_next=True,
    )


def add_subheading(doc, text):
    return add_body(
        doc,
        [(text, True)],
        size=13,
        before=8,
        after=2,
        align="left",
        keep_next=True,
    )


def add_bullet(doc, parts, size=13):
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, before=1, after=2, align="left")
    fmt = paragraph.paragraph_format
    fmt.left_indent = Cm(0.75)
    fmt.first_line_indent = Cm(-0.4)
    add_runs(paragraph, [("•  ", False)] + parts, size=size)
    return paragraph


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, color="1B3325", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), sz)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        tcBorders.append(element)
    tcPr.append(tcBorders)


def set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def prevent_row_split(row, keep_with_next=False):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))
    if keep_with_next:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                pPr = paragraph._p.get_or_add_pPr()
                if pPr.find(qn("w:keepNext")) is None:
                    pPr.append(OxmlElement("w:keepNext"))


def write_cell(cell, text, bold=False, align="left", fill=None, size=12, color=None):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    set_paragraph_format(paragraph, before=0, after=0, align=align)
    for piece in split_keywords(text):
        run = paragraph.add_run(piece)
        set_run_font(run, size=size, bold=bold or piece in KEYWORDS, color=color or BLACK)
    set_cell_borders(cell)
    set_cell_margins(cell)
    if fill:
        shade_cell(cell, fill)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)


def add_table(doc, headers, rows, widths, aligns=None, size=10, header_size=10):
    if aligns is None:
        aligns = ["left"] * len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    total = sum(widths)
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        write_cell(
            cell,
            header,
            bold=True,
            align="center",
            fill=DARK,
            size=header_size,
            color=RGBColor(255, 255, 255),
        )
        cell.width = Twips(widths[index])
    trPr = table.rows[0]._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))
    keep_block = len(rows) <= 7
    prevent_row_split(table.rows[0], keep_with_next=True)

    for r_index, row in enumerate(rows):
        fill = "F7FBF8" if r_index % 2 else "FFFFFF"
        for c_index, value in enumerate(row):
            cell = table.rows[r_index + 1].cells[c_index]
            write_cell(cell, value, bold=False, align=aligns[c_index], fill=fill, size=size)
            cell.width = Twips(widths[c_index])
        prevent_row_split(
            table.rows[r_index + 1],
            keep_with_next=keep_block and r_index < len(rows) - 1,
        )

    spacer = doc.add_paragraph()
    set_paragraph_format(spacer, before=0, after=2, align="left")
    return table


def clear_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)
    tblPr.append(borders)


def set_table_indent(table, indent):
    tblPr = table._tbl.tblPr
    tblInd = tblPr.find(qn("w:tblInd"))
    if tblInd is None:
        tblInd = OxmlElement("w:tblInd")
        tblPr.append(tblInd)
    tblInd.set(qn("w:w"), str(indent))
    tblInd.set(qn("w:type"), "dxa")


def add_note(doc, text):
    return add_body(doc, [(text, False)], size=11, before=0, after=6, align="left")


def add_contacts(doc):
    contacts = [
        (ASSETS / "icon_telefono.png", "3107862438 - 3118620020"),
        (ASSETS / "icon_correo.png", "info@thecompanydev.com"),
        (ASSETS / "icon_web.png", "www.thecompanydev.com"),
    ]
    anchor = doc.add_paragraph()
    set_paragraph_format(anchor, before=14, after=0, align="right", keep_next=True)

    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.autofit = False
    widths = [420, 3480]
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(sum(widths)))
    tblW.set(qn("w:type"), "dxa")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    set_table_indent(table, 9648 - sum(widths))
    clear_table_borders(table)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for index, (icon, text) in enumerate(contacts):
        row = table.rows[index]
        prevent_row_split(row, keep_with_next=index < len(contacts) - 1)
        icon_cell = row.cells[0]
        text_cell = row.cells[1]
        for cell in (icon_cell, text_cell):
            set_cell_margins(cell, top=20, bottom=20, left=40, right=40)
        icon_cell.text = ""
        paragraph = icon_cell.paragraphs[0]
        set_paragraph_format(paragraph, before=0, after=0, align="right")
        run = paragraph.add_run()
        run.add_picture(str(icon), width=Inches(0.2), height=Inches(0.2))
        text_cell.text = ""
        paragraph = text_cell.paragraphs[0]
        set_paragraph_format(paragraph, before=0, after=0, align="left")
        text_run = paragraph.add_run(text)
        set_run_font(text_run, size=11, bold=False, color=BLACK)


def add_background(header, image_path):
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_format(paragraph, before=0, after=0, align="left")
    run = paragraph.add_run()
    set_run_font(run, size=1)
    run.add_picture(image_path, width=Inches(8.5), height=Inches(11))
    drawing = run._element.find(qn("w:drawing"))
    inline = drawing.find(qn("wp:inline"))
    extent = inline.find(qn("wp:extent"))
    cx = extent.get("cx")
    cy = extent.get("cy")
    doc_pr = inline.find(qn("wp:docPr"))
    graphic = inline.find(qn("a:graphic"))
    anchor = parse_xml(
        f'''
        <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   distT="0" distB="0" distL="0" distR="0"
                   simplePos="0" relativeHeight="0" behindDoc="1"
                   locked="0" layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="{cx}" cy="{cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:wrapNone/>
          <wp:docPr id="{doc_pr.get("id")}" name="membrete"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
        </wp:anchor>
        '''
    )
    anchor.append(graphic)
    drawing.replace(inline, anchor)


def configure_section(section):
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1.72)
    section.bottom_margin = Inches(1.08)
    section.left_margin = Inches(0.95)
    section.right_margin = Inches(0.85)
    section.header_distance = Inches(0.15)
    section.footer_distance = Inches(0.15)
    section.different_first_page_header_footer = False
    sectPr = section._sectPr
    for tag in ("w:titlePg",):
        node = sectPr.find(qn(tag))
        if node is not None:
            sectPr.remove(node)


def add_label_line(doc, label, value):
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, before=0, after=1, align="left")
    label_run = paragraph.add_run(label)
    set_run_font(label_run, size=13, bold=True, color=NAVY)
    for piece in split_keywords(value):
        value_run = paragraph.add_run(piece)
        set_run_font(value_run, size=13, bold=piece in KEYWORDS, color=BLACK)
    return paragraph


def build():
    doc = Document()
    configure_section(doc.sections[0])
    add_background(doc.sections[0].header, str(ASSETS / "membrete_fondo_carta_sin_contacto.png"))

    normal = doc.styles["Normal"]
    normal.font.name = "Poppins"
    normal.font.size = Pt(13)
    normal.font.color.rgb = BLACK
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), "Poppins")

    # Mercado leído el 14 de octubre de 2026. Las equivalencias en pesos
    # usan la TRM de ese día y son cálculo de este informe.
    odoo_co_std = 8.95
    odoo_co_std_promo = 7.25
    odoo_co_custom_first = 13.40
    odoo_co_custom_list_oec = 16.40
    odoo_co_custom_aidoo = 13.60
    odoo_co_custom_base = 20.40
    odoo_us_std_first = 24.90
    odoo_us_std = 31.10
    odoo_us_std_month_first = 31.10
    odoo_us_std_month = 38.90
    odoo_us_custom_first = 49.00
    odoo_us_custom = 61.00
    odoo_us_custom_month = 76.20
    odoo_light = 8.90
    team = 30
    # La directriz de $20.000 por módulo sumaba $60.000 y quedaba por
    # encima de una suite de una sola licencia. Cada módulo queda ahora
    # en el orden de su referencia: Odoo para la gestión, Factorial para
    # la jornada y WorkDrive o SharePoint para el archivo.
    price_crm = 15000
    price_attendance = 8000
    price_drive = 10000
    full_user = price_crm + price_attendance + price_drive
    google_standard_cop = 58400
    kvm1_offer = 19900
    kvm1_renew = 38900
    kvm2_renew = 50900
    kvm4_renew = 97900
    kvm8_renew = 168900

    bitrix = [
        ("Free", "1 a 2 en la tabla", 0, 0),
        ("Basic", "5", 69, 49),
        ("Standard", "50", 144, 99),
        ("Professional", "100", 289, 199),
        ("Enterprise", "250", 579, 399),
    ]

    add_label_line(doc, "Fecha     ", "14 de octubre de 2026")
    add_label_line(
        doc,
        "Asunto   ",
        "Perfil del servicio RecoTime: identidad, alcance, mercado y precios",
    )
    add_label_line(doc, "Para      ", "Directivos de THE COMPANY DEV")

    add_body(
        doc,
        [("INFORME DE PERFIL DE SERVICIO", True)],
        size=13,
        before=12,
        after=0,
        align="left",
        keep_next=True,
    )
    add_body(
        doc,
        [("RECOTIME", True)],
        size=14,
        before=2,
        after=0,
        align="left",
        keep_next=True,
        color=NAVY,
    )
    add_body(
        doc,
        [("Trazabilidad de clientes, chat interno, jornada e intranet Drive", True)],
        size=13,
        before=0,
        after=2,
        align="left",
        keep_next=True,
    )
    add_body(
        doc,
        [
            (
                "Quién atendió al cliente, qué se dijo y cómo fue la jornada.",
                False,
            )
        ],
        size=12,
        before=2,
        after=8,
        align="left",
        color=NAVY,
    )

    add_heading(doc, "RESUMEN EJECUTIVO")
    add_body(
        doc,
        [
            (
                "Este informe define RecoTime, fija el alcance de esta versión y sostiene el precio contra tarifas públicas leídas el 14 de octubre de 2026. Lo usa la directiva de ",
                False,
            ),
            ("THE COMPANY DEV", True),
            (
                " para aprobar la oferta. Las secciones 2 y 3 dicen para qué existe el sistema. La sección 9 trae el mercado. La sección 11 trae el precio. La sección 17 deja, en corto, la parte social frente a Fondo Emprender.",
                False,
            ),
        ],
    )
    add_subheading(doc, "Qué es")
    add_body(
        doc,
        [
            (
                "RecoTime es el sistema de gestión de la empresa. El módulo 1 lleva la trazabilidad de la gestión de cada cliente y el chat interno de la compañía, junto con el ERP de áreas, cargos y procesos. El módulo 2 registra la jornada: llegada, almuerzo, descansos y salida. El módulo 3 es el Drive, la intranet de documentos. La nómina y la contabilidad quedan en el programa que la empresa ya use.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Módulo", "Qué queda registrado", "Por usuario al mes"],
        [
            [
                "1. CRM y ERP de gestión",
                "Expediente del cliente, chat interno, áreas, cargos, procesos y trazabilidad del trabajador.",
                cop(price_crm),
            ],
            [
                "2. Asistencia",
                "Llegada, inicio y fin de almuerzo, descansos y salida, con fecha y hora.",
                cop(price_attendance),
            ],
            [
                "3. Drive",
                "Intranet de documentos. THE COMPANY DEV entrega 2 TB para toda la cuenta.",
                cop(price_drive),
            ],
            [
                "Los tres",
                "Suma de los módulos. Puesta en marcha y capacitación incluidas.",
                cop(full_user),
            ],
        ],
        [2500, 5100, 2000],
        aligns=["left", "left", "center"],
        size=9,
        header_size=9,
    )
    add_subheading(doc, "Precio y límites")
    add_body(
        doc,
        [
            (
                f"Para 30 usuarios, los tres módulos son {cop(full_user * team)} al mes. La licencia de Odoo Standard en Colombia queda en {cop(usd_to_cop(odoo_co_std))} por usuario, sin implementación. La directriz de $20.000 por módulo sumaba $60.000 y dejaba el sistema completo por encima de esa licencia de una sola vez. El precio de esta tabla es el ajuste de la sección 10.1. Hay tres formas de pago: mensual, anualidad de 10 meses por 12, y licencia vitalicia de 22 mensualidades.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            ("La VPS la coloca el cliente. ", True),
            (
                f"La contrata, la paga, entrega el acceso y la renueva. El dominio también lo coloca el cliente. La mensualidad es solo el software. La renovación del KVM 1 de Hostinger, usada como referencia, sale en {cop(kvm1_renew * 12)} al año. Los 2 TB del Drive los entrega ",
                False,
            ),
            ("THE COMPANY DEV", True),
            (", sí o sí, con ese módulo.", False),
        ],
    )

    add_heading(doc, "1. QUÉ ES RECOTIME")
    add_body(
        doc,
        [
            (
                "RecoTime reúne la gestión de clientes, el chat interno de la empresa, la operación del equipo y la jornada. El CRM es el expediente de la gestión de cada cliente: quién lo atendió, qué se hizo, en qué estado quedó y qué se escribió en el chat interno. Cada cambio guarda usuario, fecha y hora. El ERP de gestión ordena áreas, cargos y procesos, y comparte las mismas personas del expediente.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "Sobre esa base, RecoTime lleva la trazabilidad de los trabajadores y un módulo de asistencia. La trazabilidad es el expediente de gestión de cada persona: área, cargo, cliente o proceso asignado, actividad registrada y estado. La asistencia muestra, en una línea de tiempo por trabajador, la llegada, el inicio y el fin del almuerzo, los descansos y la salida. Cada marca guarda fecha y hora.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "El módulo Drive es la intranet de la empresa. Ahí se guarda la información de la gestión: documentos, manuales, actas, formatos y las carpetas de cada área. Cada persona ve lo que su rol permite. ",
                False,
            ),
            ("THE COMPANY DEV", True),
            (
                " entrega los 2 TB sí o sí. Son 2 TB para toda la cuenta, no por usuario.",
                False,
            ),
        ],
    )
    add_body(
        doc,
        [
            (
                "La contabilidad, la factura electrónica, la PILA y la liquidación de nómina permanecen en el programa contable de la empresa. RecoTime cubre el expediente del cliente, el chat interno, la operación del equipo, la jornada y la intranet. Por eso se ofrece al lado de Alegra o Siigo, y se compara con Odoo, Bitrix24 y Google Workspace en el tramo que hoy una pyme termina pagando por partes.",
                False,
            )
        ],
    )

    add_heading(doc, "2. MISIÓN")
    add_body(
        doc,
        [
            (
                "RecoTime existe para que una empresa deje constancia de tres hechos que hoy se reconstruyen en chats personales y hojas sueltas: cómo se gestionó cada cliente, qué se dijo en el chat interno de la compañía y cómo fue la jornada de cada trabajador. El expediente del cliente guarda responsable, estado, fecha y hora. El chat interno guarda autor, fecha y hora, y permanece en el sistema cuando una persona deja la empresa. La asistencia separa llegada, almuerzo, descansos y salida. El Drive guarda los documentos de esa gestión. La empresa consulta el expediente. El trabajador consulta su línea del día y el chat de su trabajo. La nómina y la contabilidad siguen en el programa que la empresa ya tenga.",
                False,
            )
        ],
    )

    add_heading(doc, "3. VISIÓN")
    add_body(
        doc,
        [
            (
                "En 2030, una empresa que use RecoTime debe poder abrir el expediente de un cliente y ver quién lo gestionó, qué se acordó en el chat interno y en qué estado quedó, y debe poder abrir el día de un trabajador y ver la llegada, el almuerzo, los descansos y la salida sin armar el reporte a mano. Esa es la meta, medida en consultas que hoy no existen en un solo sistema. El lugar que RecoTime busca es el de las pymes colombianas de 8 a 80 trabajadores que ya liquidan nómina y contabilidad en otro programa. El criterio de avance es concreto: la orden de trabajo y el historial del cliente viven en RecoTime, y el teléfono personal deja de ser el archivo de la empresa.",
                False,
            )
        ],
    )

    add_heading(doc, "4. VALORES")
    add_table(
        doc,
        ["Valor", "Qué significa en el servicio"],
        [
            ["Claridad", "Gerencia ve clientes, procesos y la jornada en el mismo sistema."],
            ["Trazabilidad", "Cada actividad guarda persona, asignación, hora y estado."],
            ["Respeto del tiempo", "La llegada, el almuerzo y los descansos quedan visibles y separados."],
            ["Expediente del cliente", "Cada gestión guarda responsable, estado, fecha, hora y el hilo del chat interno."],
            ["Acompañamiento", "La puesta en marcha y la capacitación por rol van en el precio."],
            ["Memoria de la empresa", "El Drive guarda la gestión en carpetas, con una cuota de 2 TB para la cuenta."],
        ],
        [2800, 6800],
    )

    add_heading(doc, "5. PROBLEMA QUE RESUELVE")
    add_body(
        doc,
        [
            (
                "Cuando la gestión de un cliente vive en el WhatsApp de quien lo atendió, la orden del equipo en otro chat personal y la asistencia en un mensaje de «ya llegué», gerencia reconstruye el día después de que pasó. RecoTime separa esos registros así:",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Hoy, sin RecoTime", "Con RecoTime"],
        [
            [
                "La gestión del cliente está en el chat personal de quien lo atendió.",
                "Expediente del cliente: responsable, estado, fecha, hora e historial.",
            ],
            [
                "Nadie sabe con certeza qué trabajador atendió qué proceso.",
                "Trazabilidad: asignación, actividad, hora y estado en el expediente.",
            ],
            [
                "La llegada se confirma por mensaje. El almuerzo y los descansos no quedan.",
                "Línea del día por persona: llegada, almuerzo, descansos y salida.",
            ],
            [
                "El reporte de jornada se arma el viernes, con datos incompletos.",
                "El supervisor abre el día en curso y ve las marcas ya registradas.",
            ],
            [
                "Las órdenes del equipo se escriben en el teléfono personal.",
                "Chat interno de la empresa: autor, fecha, hora, cliente o área.",
            ],
            [
                "Los manuales y las actas viven en chats, correos y carpetas sueltas.",
                "Drive: intranet con carpetas por área y una cuota de 2 TB para la cuenta.",
            ],
        ],
        [4600, 5000],
        size=10,
    )

    add_heading(doc, "6. MERCADO OBJETIVO")
    add_bullet(
        doc,
        [
            (
                "Pymes de servicios y operación que necesitan el expediente de cada cliente, un chat interno de la empresa y la jornada del equipo en el mismo sistema.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Empresas de 8 a 80 trabajadores activos en las que gerencia todavía cruza una hoja de clientes con un chat de «ya llegué».",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Negocios que ya tienen contabilidad y nómina en Alegra, Siigo u otro programa, y buscan la capa de gestión y asistencia.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "El comprador de esta versión es la empresa que dirige personas y clientes. Las cifras de nómina y de contabilidad de la sección 9 sirven para mostrar qué problema cubre otro producto. RecoTime puede crecer después, por cotización, hacia varias sedes, marcación desde un reloj o una aplicación publicada en tienda.",
                False,
            )
        ],
    )

    add_heading(doc, "7. CÓMO FUNCIONA")
    add_body(
        doc,
        [
            (
                "RecoTime tiene cuatro accesos. Cada persona entra con su usuario y su rol: gerencia, gestión de clientes, supervisor y trabajador. El chat interno está disponible para los roles que la empresa habilite. Un mensaje no sale del sistema hacia el teléfono personal.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Acceso", "Qué hace en RecoTime"],
        [
            [
                "Gerencia",
                "Ve el tablero, el expediente de los clientes, el chat interno, la asistencia del día y el Drive.",
            ],
            [
                "Gestión de clientes",
                "Actualiza el expediente: responsable, estado y lo tratado. Escribe en el chat interno del cliente o del área.",
            ],
            [
                "Supervisor",
                "Asigna trabajadores a clientes o procesos, revisa la trazabilidad y las novedades de la jornada.",
            ],
            [
                "Trabajador",
                "Marca llegada, almuerzo, descansos y salida. Consulta su día, sus asignaciones y el Drive de su área.",
            ],
        ],
        [2400, 7200],
    )
    add_body(
        doc,
        [
            (
                "La jornada de una persona tiene cinco marcas: llegada, inicio de almuerzo, fin de almuerzo, cada descanso y salida. La pantalla las muestra en orden, con la hora. Un descanso se distingue del almuerzo. Si falta una marca, el supervisor la ve en la lista del día y la novedad queda en el expediente.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "La trazabilidad del trabajador corre en paralelo. Una actividad se abre con trabajador, cliente o proceso, hora de inicio y estado. Al cerrarla, queda la hora de fin. Gerencia puede abrir a una persona y ver qué atendió y cómo fue su día.",
                False,
            )
        ],
    )

    add_subheading(doc, "Qué guarda cada registro")
    add_body(
        doc,
        [
            (
                "Esta versión no publica una base de datos ni una interfaz de programación. Sí fija qué dato mínimo queda en cada registro, para que la puesta en marcha y los reportes hablen de los mismos campos.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Registro", "Campos que quedan"],
        [
            ["Expediente del cliente", "Cliente, responsable, estado de la gestión, fecha y hora de cada cambio."],
            ["Mensaje del chat interno", "Autor, fecha, hora, cliente o área, y el texto. El hilo queda en la empresa."],
            ["Actividad del trabajador", "Trabajador, cliente o proceso, hora de inicio, hora de fin y estado."],
            ["Marca de asistencia", "Trabajador, tipo de marca, fecha y hora. Los tipos son llegada, inicio de almuerzo, fin de almuerzo, descanso y salida."],
            ["Documento del Drive", "Archivo, carpeta de área y rol que puede verlo. La cuota de la cuenta es de 2 TB."],
        ],
        [2800, 6800],
        size=9,
        header_size=9,
    )

    add_heading(doc, "8. ALCANCE DE ESTA VERSIÓN")
    add_table(
        doc,
        ["Módulo", "Qué incluye"],
        [
            [
                "1. CRM y ERP de gestión",
                "Expediente del cliente: responsable, estado, fecha y hora de cada cambio. Chat interno por cliente y por área, con autor, fecha y hora. Áreas, cargos, procesos y tablero. Trazabilidad del trabajador: ficha, asignación, actividad y estado. Reportes de la operación.",
            ],
            [
                "2. Asistencia",
                "Llegada, inicio y fin de almuerzo, descansos y salida. Línea de tiempo por trabajador. Lista del día para el supervisor. Reportes de la jornada.",
            ],
            [
                "3. Drive",
                "Intranet de la empresa. Carpetas por área, documentos de la gestión y permisos por rol. Cuota de 2 TB para toda la cuenta.",
            ],
        ],
        [2800, 6800],
    )
    add_subheading(doc, "Puesta en marcha")
    add_body(
        doc,
        [
            (
                "La configuración, la carga inicial acordada, la estructura de carpetas del Drive, la capacitación por rol y la documentación básica van incluidas en las tres modalidades. El calendario propuesto es de 8 a 12 semanas, contadas desde el día en que la empresa entrega el organigrama, la lista de trabajadores, los clientes activos, el horario de llegada, almuerzo y descansos, y el acceso a la VPS y al dominio. Si esa información llega después, el calendario se corre.",
                False,
            )
        ],
    )
    add_subheading(doc, "Fuera de esta versión")
    add_body(
        doc,
        [
            (
                "Esta versión deja para una cotización posterior la nómina, la PILA, la nómina electrónica ante la DIAN, la factura electrónica, la contabilidad, el inventario, las compras y el POS. También se cotizan el reloj o el biométrico, la publicación en Play Store o App Store, varias sedes, el espacio por encima de la cuota de 2 TB y el consumo de WhatsApp. Las marcas de esta versión las registra el trabajador o el supervisor dentro de RecoTime. La VPS y el dominio no se cotizan aquí: los coloca el cliente.",
                False,
            )
        ],
    )

    add_heading(doc, "9. PRECIOS PUBLICADOS EN EL MERCADO")
    add_body(
        doc,
        [
            (
                "Las cifras de esta sección se leyeron el 14 de octubre de 2026, salvo donde se indica otra fecha. Pueden cambiar. Donde la página no dice si el IVA va incluido, este informe tampoco lo afirma. Las equivalencias en pesos usan la TRM de $3.181,74 del 14 de octubre de 2026 y son cálculo de este informe, redondeado al peso. Un ejemplo de 30 trabajadores es una multiplicación propia, no una cotización de ese proveedor.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "Odoo cambia el precio según el país. La consulta directa a odoo.com/pricing desde este informe mostró la lista de Estados Unidos. La tarifa que varias fuentes asignan a Colombia es más baja y se presenta aparte, con su fuente. Las dos listas se necesitan para no cotizar RecoTime contra un dólar que el cliente colombiano no paga.",
                False,
            )
        ],
    )

    add_subheading(doc, "9.1. Odoo")
    add_body(
        doc,
        [
            (
                "Fuente de la lista internacional: odoo.com/pricing, abierta el 14 de octubre de 2026. El plan de una aplicación es gratis, con usuarios sin límite, en Odoo en línea. Standard y Personalizado incluyen todas las aplicaciones, entre ellas CRM, RR. HH. y asistencia, por un solo valor por usuario. La propia página dice que el descuento dura 12 meses y aplica a los usuarios pedidos al inicio. El alojamiento de Odoo.sh y la implementación se cobran aparte. Un empleado cargado en RR. HH., aunque no tenga usuario, cuenta como usuario de pago.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Plan en la lista de EE. UU.", "Primer año", "Después", "COP / usuario / mes, después"],
        [
            [
                "Standard, pago anual",
                usd(odoo_us_std_first),
                usd(odoo_us_std),
                cop(usd_to_cop(odoo_us_std)),
            ],
            [
                "Standard, pago mensual",
                usd(odoo_us_std_month_first),
                usd(odoo_us_std_month),
                cop(usd_to_cop(odoo_us_std_month)),
            ],
            [
                "Personalizado, pago anual",
                usd(odoo_us_custom_first),
                usd(odoo_us_custom),
                cop(usd_to_cop(odoo_us_custom)),
            ],
            [
                "Personalizado, pago mensual",
                usd(61.00),
                usd(odoo_us_custom_month),
                cop(usd_to_cop(odoo_us_custom_month)),
            ],
        ],
        [2800, 1800, 1800, 3200],
        aligns=["left", "center", "center", "center"],
        size=9,
        header_size=9,
    )
    add_note(
        doc,
        "Valores por usuario al mes, en dólares de la lista que mostró la página. La columna en pesos convierte solo el precio de renovación. La página en español muestra además un usuario light por US$ 8,90, para quien solo marca asistencia, vacaciones o gastos. Ese usuario se contrata con un asesor y esta tabla no lo mezcla con el usuario completo.",
    )
    add_body(
        doc,
        [
            (
                "Fuentes de la tarifa que se reporta para Colombia: OEC.sh, con lectura desde Colombia fechada el 26 de septiembre de 2026; la calculadora de Aidoo sobre precios públicos de Odoo en 2026; y la nota de un partner de Odoo Colombia del 24 de septiembre de 2026 sobre la actualización del plan Personalizado. Standard de lista anual coincide en US$ 8,95. El Personalizado no coincide entre fuentes, y este informe deja las tres lecturas.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Lectura para Colombia", "US$ / usuario / mes", "COP / usuario / mes", "Fuente"],
        [
            [
                "Standard, lista anual",
                usd(odoo_co_std),
                cop(usd_to_cop(odoo_co_std)),
                "OEC.sh y Aidoo",
            ],
            [
                "Standard, primer año anual",
                usd(odoo_co_std_promo),
                cop(usd_to_cop(odoo_co_std_promo)),
                "OEC.sh",
            ],
            [
                "Standard, pago mensual",
                usd(9.10),
                cop(usd_to_cop(9.10)),
                "OEC.sh",
            ],
            [
                "Personalizado, primer año anual",
                usd(odoo_co_custom_first),
                cop(usd_to_cop(odoo_co_custom_first)),
                "OEC.sh y nota del 24 sep 2026",
            ],
            [
                "Personalizado, lista anual",
                usd(odoo_co_custom_list_oec),
                cop(usd_to_cop(odoo_co_custom_list_oec)),
                "OEC.sh",
            ],
            [
                "Personalizado, lista anual",
                usd(odoo_co_custom_aidoo),
                cop(usd_to_cop(odoo_co_custom_aidoo)),
                "Aidoo",
            ],
            [
                "Personalizado, tarifa base",
                usd(odoo_co_custom_base),
                cop(usd_to_cop(odoo_co_custom_base)),
                "Nota del 24 sep 2026",
            ],
        ],
        [3000, 2000, 2200, 2400],
        aligns=["left", "center", "center", "left"],
        size=8,
        header_size=8,
    )
    add_note(
        doc,
        f"Para el ejemplo de este informe se usan dos cifras que se repiten: Standard de lista anual a {usd(odoo_co_std)} y Personalizado de primera contratación anual a {usd(odoo_co_custom_first)}. Con 30 usuarios, la licencia Standard queda en {cop(usd_to_cop(odoo_co_std * team))} al mes ({cop(usd_to_cop(odoo_co_std * team * 12))} al año). La de Personalizado, en esa primera contratación, queda en {cop(usd_to_cop(odoo_co_custom_first * team))} al mes. La implementación va por fuera.",
    )
    add_body(
        doc,
        [
            (
                "Un artículo de Catalizadora de 2026, sobre una pyme colombiana de 20 empleados, ubica el primer año de Odoo entre $54.700.000 y $149.900.000 cuando se suman licencia, implementación, localización, hosting y soporte. Es la estimación de ese artículo, no un precio de catálogo. Sirve para recordar que la tabla de arriba es la licencia.",
                False,
            )
        ],
    )

    add_subheading(doc, "9.2. Bitrix24")
    add_body(
        doc,
        [
            (
                "Fuente: bitrix24.com/prices, abierta el 14 de octubre de 2026. Bitrix24 cobra un fijo por organización, con un tope de usuarios. CRM, tareas y seguimiento de tiempo comercial entran desde los planes pagos. Las herramientas de RR. HH. y el registro de horas del empleado aparecen en Professional, no en Standard. La página advierte que puede sumarse IVA según el país. La oferta Vibe+, vigente del 1 de septiembre al 10 de diciembre de 2026 para suscripciones de 12 meses, baja Professional a US$ 240 al mes. Las tablas de este informe usan la línea estable, sin esa oferta.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Plan", "Usuarios", "Pago mensual", "Pago anual, por mes", "Equivalente COP del anual"],
        [
            [
                name,
                users,
                usd(month) if month else "US$ 0",
                usd(year) if year else "US$ 0",
                cop(usd_to_cop(year)) if year else "$0",
            ]
            for name, users, month, year in bitrix
        ],
        [1800, 1700, 1900, 2100, 2100],
        aligns=["left", "center", "center", "center", "center"],
        size=8,
        header_size=8,
    )
    add_note(
        doc,
        "La tabla de planes muestra Free para 1 a 2 usuarios. La pregunta frecuente de la misma página dice que el plan gratis permite usuarios sin límite, con herramientas recortadas. El pago anual de Professional son US$ 2.388 al año, equivalentes a "
        + cop(usd_to_cop(199 * 12))
        + ". Enterprise sigue en la misma página hasta 10.000 usuarios, a US$ 7.999 al mes con pago anual.",
    )

    add_subheading(doc, "9.3. Zoho CRM y Zoho People")
    add_body(
        doc,
        [
            (
                "Fuente directa del 14 de octubre de 2026: zoho.com/crm. La página confirma el plan gratis hasta 3 usuarios y publica Zoho CRM Plus desde US$ 20 por usuario al mes. Los valores de Standard, Professional y Enterprise no quedaron impresos en esa consulta porque la página los carga según el país. Una lectura del 6 de agosto de 2026 de la lista anual de Estados Unidos los ubica en US$ 14, US$ 23, US$ 40 y US$ 52 por usuario al mes, de Standard a Ultimate. Este informe usa CRM Plus, que sí se leyó hoy, y cita la otra escala como lectura de esa fecha.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "Zoho People, en zoho.com/people, confirma el 14 de octubre de 2026 los planes Essential HR, Professional, Premium y Enterprise, y ubica la asistencia, las hojas de tiempo y los turnos en la línea de pago. Los dólares no salieron impresos. Una revisión del 22 de septiembre de 2026 de la lista de Estados Unidos (hriscost.com, sobre la página de Zoho) registra Essential HR a US$ 3, Workforce y Talent a US$ 5, y Enterprise a US$ 8 por usuario al mes con pago anual. El plan gratis cubre hasta 3 usuarios en esa lista. La nómina de Zoho se vende aparte.",
                False,
            )
        ],
    )
    add_note(
        doc,
        "Ejemplo de este informe, con la lectura de agosto y septiembre, para 5 usuarios de CRM Professional y 30 de People Essential: (5 × US$ 23) + (30 × US$ 3) = US$ 205 al mes, equivalentes a "
        + cop(usd_to_cop(205))
        + ". Cubre CRM y personas. El ERP de gestión de RecoTime y la línea de almuerzo y descansos no están en ese paquete. Zoho One, en una guía de 2026, licencia a todo el personal desde US$ 37 por empleado al mes con pago anual: 30 empleados equivalen a "
        + cop(usd_to_cop(30 * 37))
        + " al mes.",
    )

    add_subheading(doc, "9.4. Factorial, Buk y HubSpot")
    add_body(
        doc,
        [
            (
                "Factorial, en factorial.com/pricing-plans el 14 de octubre de 2026, publica un punto de partida de US$ 8 por usuario al mes y arma el plan a la medida. El registro de tiempo es un módulo que se cotiza encima del núcleo. Para 30 usuarios, solo el punto de partida son US$ 240 al mes, equivalentes a "
                + cop(usd_to_cop(240))
                + ". Es un sistema de personas, no un CRM ni un ERP de gestión.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "Buk Colombia, en buk.co/precios, calcula la mensualidad según los colaboradores vigentes y los módulos, entre ellos nómina y control de asistencia. La página no imprime una tarifa. Este informe no le asigna una cifra. El cobro es mensual e incluye las cuentas de los colaboradores y los administradores que la empresa necesite.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "La tabla viva de HubSpot no cargó el 14 de octubre de 2026. La guía comercial de HubSpot y el catálogo leído a inicios de octubre de 2026 coinciden en Sales Hub Professional a US$ 90 por asiento al mes con pago anual, u US$ 100 mes a mes, más un onboarding obligatorio de US$ 1.500. Enterprise figura en US$ 150 por asiento al mes. Es un CRM por asiento comercial. La jornada del equipo queda por fuera.",
                False,
            )
        ],
    )

    add_subheading(doc, "9.5. Alegra en Colombia")
    add_body(
        doc,
        [
            (
                "Fuente: alegra.com/colombia/contabilidad/precios y alegra.com/colombia/precios, abiertas el 14 de octubre de 2026. Alegra es contabilidad, factura, POS y nómina. En la misma página de contabilidad, las tarjetas muestran el valor con pago anual y, más abajo, el valor del mes a mes. La pregunta frecuente de esa página todavía imprime una escala anterior ($69.900, $149.900, $219.900 y $279.900). Este informe usa las tarjetas y los valores mensuales rotulados.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Plan de contabilidad", "Mes a mes", "Equivale con pago anual", "Se factura al año"],
        [
            ["Emprendedor", cop(74900), cop(56175), cop(674100)],
            ["Pyme", cop(163900), cop(122925), cop(1475100)],
            ["Pro", cop(250900), cop(188175), cop(2258100)],
            ["Plus", cop(319900), cop(239925), cop(2879100)],
        ],
        [2400, 2200, 2600, 2400],
        aligns=["left", "center", "center", "center"],
        size=9,
        header_size=9,
    )
    add_note(
        doc,
        "Emprendedor admite ingresos hasta $10.000.000 al mes y 1 usuario, más el contador. Pyme sube a $40.000.000 y 2 usuarios. Pro, a $180.000.000 y 3 usuarios. Plus llega a 8 usuarios y 100 bodegas. El certificado digital figura en $0 dentro del plan y en $150.000 al año como referencia tachada.",
    )
    add_table(
        doc,
        ["Nómina Alegra", "Mes a mes", "Equivale con pago anual", "Se factura al año"],
        [
            ["De 1 a 10 empleados", cop(29900), cop(22425), cop(269100)],
            ["De 11 a 20 empleados", cop(69000), cop(51750), cop(621000)],
        ],
        [2400, 2200, 2600, 2400],
        aligns=["left", "center", "center", "center"],
        size=9,
        header_size=9,
    )
    add_note(
        doc,
        "La página también lista una banda de 21 a 45 empleados. El texto extraído el 14 de octubre de 2026 no imprimió su precio, y este informe no lo inventa. De 1 a 10 empleados el plan es emisión electrónica. De 11 a 20 suma cálculo de nómina, colillas y reportes. Sigue siendo nómina, no asistencia ni CRM.",
    )

    add_subheading(doc, "9.6. Dynamics 365 Business Central y la nómina local")
    add_body(
        doc,
        [
            (
                "Fuente: la página de precios de Microsoft Dynamics 365 Business Central, consultada el 14 de octubre de 2026. Essentials cuesta US$ 80 por usuario al mes y Premium US$ 110, con pago anual. Premium agrega manufactura y gestión de servicio. Para 30 usuarios, Essentials son US$ 2.400 al mes, equivalentes a "
                + cop(usd_to_cop(2400))
                + ". La implementación se contrata con un partner. Es un ERP contable y operativo completo, en otro rango de inversión.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "En nómina colombiana, la guía NominaCheck del 21 de septiembre de 2026 lee precios públicos así: Siigo Nómina Electrónica Lite, para 3 a 10 empleados, en $239.920 al año con descuento de cliente nuevo (lista $299.900), solo emisión. Aleluya, antes Nominapp, en $220.000 al mes para 10 empleados en el plan Tranqui, con liquidación. Siigo Nube, el contable, se contrata aparte de la nómina. Se citan para situar el piso de un software de pago de personal, que RecoTime no reemplaza.",
                False,
            )
        ],
    )

    add_subheading(doc, "9.7. Intranet y almacenamiento: Google, Zoho y SharePoint")
    add_body(
        doc,
        [
            (
                "Fuente de Google: la ayuda de facturación de Google Workspace, consultada el 14 de octubre de 2026. Business Standard incluye 2 TB de almacenamiento agrupado por usuario, más correo, documentos y reuniones. En pesos, el plan anual figura en $58.400 por usuario al mes y el plan flexible en $70.000. Para 30 usuarios, el anual son "
                + cop(google_standard_cop * team)
                + " al mes. Ese producto entrega 2 TB por persona, no una cuota única de 2 TB, y no es un CRM de operación ni un control de almuerzo y descansos.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                "Zoho WorkDrive, en la guía de planes de zoho.com, cobra por usuario con un mínimo de 3. Con pago anual, Starter queda en US$ 2,50, Team en US$ 4,50 y Business en US$ 9. Entre 3 y 10 usuarios el almacenamiento del equipo es 1 TB, 3 TB y 5 TB. Por encima de 10 usuarios se suma espacio por persona. Team, a US$ 4,50, equivale a "
                + cop(usd_to_cop(4.50))
                + " por usuario al mes. Es solo el archivo. SharePoint Plan 1, en la página de Microsoft, cuesta US$ 5 por usuario al mes con pago anual, equivalentes a "
                + cop(usd_to_cop(5))
                + ". Es la intranet de archivos, sin el CRM ni la asistencia de RecoTime.",
                False,
            )
        ],
    )

    add_subheading(doc, "9.8. VPS de Hostinger")
    add_body(
        doc,
        [
            (
                "Fuente: hostinger.com/co/precios/vps-hosting, abierta el 14 de octubre de 2026. La página muestra una oferta al mes y una renovación. La pregunta frecuente usa la oferta: KVM 1 en $19.900, KVM 2 en $28.900, KVM 4 en $38.900 y KVM 8 en $77.900. La tarjeta de cada plan dice a cuánto se renueva. El disco publicado es NVMe. El ancho de banda de 4 TB del KVM 1 es tráfico, no espacio para guardar archivos. Los 2 TB del Drive no dependen de ese disco: los entrega THE COMPANY DEV.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Plan", "Oferta al mes", "Renovación al mes", "Renovación al año", "Disco NVMe"],
        [
            ["KVM 1", cop(kvm1_offer), cop(kvm1_renew), cop(kvm1_renew * 12), "50 GB"],
            ["KVM 2", cop(28900), cop(kvm2_renew), cop(kvm2_renew * 12), "100 GB"],
            ["KVM 4", cop(38900), cop(kvm4_renew), cop(kvm4_renew * 12), "200 GB"],
            ["KVM 8", cop(77900), cop(kvm8_renew), cop(kvm8_renew * 12), "400 GB"],
        ],
        [1400, 1800, 2100, 2300, 2000],
        aligns=["left", "center", "center", "center", "center"],
        size=8,
        header_size=8,
    )
    add_note(
        doc,
        "Renovación al año = renovación al mes × 12, cálculo de este informe. KVM 1 se renueva en $466.800 al año. Es la cifra que se acerca a los $500.000 al año que la operación usa como referencia. La oferta del primer periodo, a $19.900 al mes, son $238.800 si se mantienen doce meses. Hostinger puede cambiar la campaña. Esta tabla es referencia: la VPS la coloca el cliente y no se factura dentro de RecoTime. El dominio tampoco va aquí.",
    )

    add_heading(doc, "10. LECTURA PARA RECOTIME")
    add_body(
        doc,
        [
            (
                "Una empresa de 30 trabajadores activos, con las tarifas de la sección 9, queda en estos órdenes. La columna mensual es licencia o suscripción. La puesta en marcha, cuando el proveedor la cobra aparte, no está sumada.",
                False,
            )
        ],
    )
    rows_30 = [
        (
            "RecoTime, los tres módulos",
            cop(full_user * team),
            f"30 × {cop(full_user)}. Puesta en marcha incluida. La VPS la coloca el cliente.",
        ),
        (
            "RecoTime, solo CRM y ERP",
            cop(price_crm * team),
            f"30 × {cop(price_crm)}. Asistencia y Drive se suman si el cliente los toma.",
        ),
        (
            "Odoo Standard, tarifa Colombia",
            cop(usd_to_cop(odoo_co_std * team)),
            "30 usuarios a US$ 8,95. Todas las apps. Implementación aparte.",
        ),
        (
            "Odoo Personalizado, primer año Colombia",
            cop(usd_to_cop(odoo_co_custom_first * team)),
            "30 usuarios a US$ 13,40. Studio, API y multiempresa.",
        ),
        (
            "Odoo Standard, lista de EE. UU. al renovar",
            cop(usd_to_cop(odoo_us_std * team)),
            "La misma suscripción si el precio que aplica es el de Estados Unidos.",
        ),
        (
            "Bitrix24 Standard, pago anual",
            cop(usd_to_cop(99)),
            "Tope de 50 usuarios. CRM y tareas. Sin las horas de empleado de Professional.",
        ),
        (
            "Bitrix24 Professional, pago anual",
            cop(usd_to_cop(199)),
            "Tope de 100 usuarios. Incluye RR. HH. y horas del empleado. El partner va aparte.",
        ),
        (
            "Zoho CRM y People, lectura 2026",
            cop(usd_to_cop(205)),
            "5 asientos de CRM Professional y 30 de People Essential. Dos productos.",
        ),
        (
            "Factorial, punto de partida",
            cop(usd_to_cop(240)),
            "30 × US$ 8. El módulo de tiempo se cotiza encima.",
        ),
        (
            "Business Central Essentials",
            cop(usd_to_cop(2400)),
            "30 × US$ 80. ERP de finanzas y operación. Partner aparte.",
        ),
        (
            "Alegra Nómina, 11 a 20 empleados",
            cop(69000),
            "Nómina electrónica y liquidación. Otro problema, otro sistema.",
        ),
    ]
    add_table(
        doc,
        ["Referencia para 30 personas", "Orden mensual", "Qué está pagando"],
        [list(row) for row in rows_30],
        [2800, 1900, 4900],
        aligns=["left", "center", "left"],
        size=8,
        header_size=8,
    )
    bitrix_pro_month = usd_to_cop(199)
    users_under_bitrix = int(bitrix_pro_month // full_user)
    add_subheading(doc, "10.1. Precio acoplado al mercado")
    add_body(
        doc,
        [
            (
                "La directriz de $20.000 por módulo sumaba $60.000 por usuario. Para 30 personas eso eran $1.800.000 al mes, por encima de la licencia de Odoo Standard en Colombia. El ajuste deja cada módulo en el orden de su propia referencia y los tres juntos entre las dos licencias de Odoo. Los pesos de las referencias en dólares son cálculo de este informe con la TRM de $3.181,74. La puesta en marcha va incluida en estos valores. La VPS no va: la coloca el cliente.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Pieza", "Referencia de mercado", "RecoTime", "Lectura"],
        [
            [
                "CRM y ERP de gestión",
                "Odoo Standard, Colombia: " + cop(usd_to_cop(odoo_co_std)) + " por usuario, solo licencia",
                cop(price_crm) + " por usuario, puesta en marcha incluida",
                "Cabe. " + cop(usd_to_cop(odoo_co_std) - price_crm) + " por debajo de esa licencia.",
            ],
            [
                "Asistencia",
                "Factorial desde " + cop(usd_to_cop(8)) + " por usuario. El tiempo se cotiza encima.",
                cop(price_attendance) + " por usuario",
                "Cabe. Cubre llegada, almuerzo, descansos y salida.",
            ],
            [
                "Drive",
                "WorkDrive Team " + cop(usd_to_cop(4.50)) + ". SharePoint Plan 1 " + cop(usd_to_cop(5)) + ".",
                cop(price_drive) + " por usuario",
                "Cabe. Por debajo de las dos referencias de archivos.",
            ],
            [
                "Los tres módulos",
                "Odoo Standard " + cop(usd_to_cop(odoo_co_std)) + ". Personalizado de primer año " + cop(usd_to_cop(odoo_co_custom_first)) + ".",
                cop(full_user) + " por usuario",
                "Cabe entre las dos licencias. La puesta en marcha va incluida.",
            ],
            [
                "30 usuarios, los tres",
                "Odoo Standard " + cop(usd_to_cop(odoo_co_std * team)) + " al mes, solo licencia",
                cop(full_user * team) + " al mes",
                cop(full_user * team - usd_to_cop(odoo_co_std * team)) + " de diferencia. En Odoo la implementación se paga aparte.",
            ],
            [
                "30 usuarios frente a Bitrix24 Professional",
                cop(bitrix_pro_month) + " al mes, tope de 100 usuarios",
                cop(full_user * team) + " al mes",
                "Bitrix cobra un fijo. Hasta " + str(users_under_bitrix) + " usuarios, RecoTime completo queda más bajo.",
            ],
            [
                "VPS y dominio",
                "KVM 1: 50 GB por " + cop(kvm1_renew * 12) + " al año",
                "Fuera de la mensualidad",
                "Los coloca el cliente. Los 2 TB del Drive los entrega THE COMPANY DEV.",
            ],
        ],
        [2100, 2800, 2500, 2200],
        aligns=["left", "left", "left", "left"],
        size=8,
        header_size=8,
    )
    add_body(
        doc,
        [
            (
                f"El Drive queda en {cop(price_drive)} por usuario. Zoho WorkDrive Team, solo archivos, equivale a {cop(usd_to_cop(4.50))} por usuario. SharePoint Plan 1 equivale a {cop(usd_to_cop(5))}. Google Workspace Business Standard, en la lista de Colombia, cuesta {cop(google_standard_cop)} por usuario e incluye correo, documentos y 2 TB por persona. Los tres módulos de RecoTime juntos son {cop(full_user)} por usuario. Con el Drive, THE COMPANY DEV entrega 2 TB para toda la cuenta.",
                False,
            )
        ],
    )
    add_body(
        doc,
        [
            (
                f"La mensualidad no incluye servidor. La VPS la coloca el cliente: la contrata en Hostinger, la paga, entrega el acceso a ",
                False,
            ),
            ("THE COMPANY DEV", True),
            (
                f" y la renueva. El dominio también lo coloca el cliente. La instalación de RecoTime se hace en esa VPS y no se factura como hosting. La referencia de $500.000 al año corresponde a la renovación del KVM 1, {cop(kvm1_renew * 12)}. Los 2 TB del Drive los entrega THE COMPANY DEV, sí o sí, con ese módulo.",
                False,
            ),
        ],
    )

    add_heading(doc, "11. MODALIDADES Y PRECIOS PROPUESTOS")
    add_body(
        doc,
        [
            (
                "Los valores de esta sección son la recomendación de este informe para aprobación de la directiva. No son precios publicados del mercado. Van en pesos colombianos. El IVA se liquida según el régimen de ",
                False,
            ),
            ("THE COMPANY DEV", True),
            (" al momento de facturar, y no está sumado en las tablas.", False),
        ],
    )
    add_body(
        doc,
        [
            (
                "RecoTime se ofrece en tres modalidades y por módulos. En ninguna modalidad se cobra una implementación aparte. La configuración, la puesta en marcha y la capacitación por rol están incluidas. El precio es por usuario activo al mes y es la suma de los módulos que el cliente tome. La VPS no entra en ese precio: la coloca el cliente. También coloca el dominio.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Concepto", "Mensual", "Anualidad", "Licencia vitalicia"],
        [
            ["Forma de pago", "Cada mes", "Un pago al año", "Pago único, 50 % y 50 %"],
            ["Implementación", "Incluida", "Incluida", "Incluida"],
            ["Uso del sistema", "Mientras pague", "12 meses", "Indefinido, versión entregada"],
            ["Soporte de 3 h/mes", "Incluido", "Incluido", "3 meses; luego plan opcional"],
            ["Actualizaciones", "Incluidas", "Incluidas", "Con el plan de soporte"],
            ["VPS, disco y dominio", "Los coloca el cliente", "Los coloca el cliente", "Los coloca el cliente"],
            ["Cuota Drive", "2 TB por cuenta", "2 TB por cuenta", "2 TB por cuenta"],
            ["Código fuente", "THE COMPANY DEV", "THE COMPANY DEV", "THE COMPANY DEV"],
            ["Beneficio", "1 mes gratis por año cumplido", "Se pagan 10 meses y se usan 12", "Sin canon de licencia"],
        ],
        [2300, 2430, 2430, 2440],
        aligns=["left", "center", "center", "center"],
        size=8,
        header_size=8,
    )

    add_subheading(doc, "11.1. Precio de cada módulo")
    add_body(
        doc,
        [
            (
                f"Cada módulo tiene su propio valor, tomado del orden de mercado de la sección 10.1. Los tres juntos cuestan {cop(full_user)} por usuario: es la suma. Quien toma un módulo paga ese módulo. Quien toma dos paga la suma de esos dos.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Módulo", "Por usuario al mes", "Anualidad por usuario", "Vitalicia por usuario"],
        [
            ["1. CRM y ERP de gestión", cop(price_crm), cop(price_crm * 10), cop(price_crm * 22)],
            ["2. Asistencia", cop(price_attendance), cop(price_attendance * 10), cop(price_attendance * 22)],
            ["3. Drive", cop(price_drive), cop(price_drive * 10), cop(price_drive * 22)],
            ["Los tres juntos", cop(full_user), cop(full_user * 10), cop(full_user * 22)],
        ],
        [2800, 2200, 2300, 2300],
        aligns=["left", "center", "center", "center"],
        size=9,
        header_size=9,
    )
    add_note(
        doc,
        "Anualidad por usuario = mensualidad × 10, y cubre 12 meses. "
        f"El CRM y ERP al año son {cop(price_crm * 10)} por usuario. La asistencia, {cop(price_attendance * 10)}. El Drive, {cop(price_drive * 10)}. Los tres, {cop(full_user * 10)}. "
        "La vitalicia es la mensualidad × 22: "
        f"{cop(price_crm * 22)}, {cop(price_attendance * 22)} y {cop(price_drive * 22)} por usuario, y {cop(full_user * 22)} los tres. "
        "El soporte opcional, después de los 3 meses de garantía, es el 40 % de lo que el cliente paga por usuario.",
    )

    add_subheading(doc, "11.2. Módulos sueltos o combinados")
    add_body(
        doc,
        [
            (
                "El cliente elige uno, dos o los tres. Dos módulos se cobran como la suma de sus tarifas. Los reportes de cada módulo van con ese módulo. La cuota de 2 TB va solo con el Drive, así se tome suelto o con los demás. La VPS la coloca el cliente en cualquier combinación.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Combinación", "Por usuario al mes", "Anualidad por usuario", "Vitalicia por usuario"],
        [
            [
                "CRM y ERP + asistencia",
                cop(price_crm + price_attendance),
                cop((price_crm + price_attendance) * 10),
                cop((price_crm + price_attendance) * 22),
            ],
            [
                "CRM y ERP + Drive",
                cop(price_crm + price_drive),
                cop((price_crm + price_drive) * 10),
                cop((price_crm + price_drive) * 22),
            ],
            [
                "Asistencia + Drive",
                cop(price_attendance + price_drive),
                cop((price_attendance + price_drive) * 10),
                cop((price_attendance + price_drive) * 22),
            ],
            ["Los tres módulos", cop(full_user), cop(full_user * 10), cop(full_user * 22)],
        ],
        [2800, 2200, 2300, 2300],
        aligns=["left", "center", "center", "center"],
        size=9,
        header_size=9,
    )
    add_note(
        doc,
        f"Para 30 usuarios, el CRM y ERP solo son {cop(price_crm * team)} al mes, la asistencia sola {cop(price_attendance * team)} y el Drive solo {cop(price_drive * team)}. Los tres juntos son {cop(full_user * team)} al mes.",
    )

    add_subheading(doc, "11.3. La VPS y el dominio los coloca el cliente")
    add_body(
        doc,
        [
            (
                "La mensualidad, la anualidad y la licencia vitalicia son el software de RecoTime. No incluyen servidor ni dominio. La VPS la coloca el cliente antes de empezar: la contrata, la paga, entrega el acceso y la renueva. El dominio lo coloca del mismo modo. ",
                False,
            ),
            ("THE COMPANY DEV", True),
            (
                " instala el sistema en esa VPS. Si el cliente no ha colocado la VPS y el dominio, el calendario de puesta en marcha no arranca.",
                False,
            ),
        ],
    )
    add_table(
        doc,
        ["Concepto", "Quién lo coloca", "Referencia leída el 14 de octubre de 2026"],
        [
            [
                "VPS Hostinger",
                "EL CLIENTE la coloca",
                "Referencia KVM 1: oferta $19.900 al mes. Renovación $38.900 al mes, $466.800 al año. Disco de 50 GB.",
            ],
            [
                "Dominio",
                "EL CLIENTE lo coloca",
                "No tiene un solo precio. Depende de la extensión y el registrador.",
            ],
            [
                "Cuota Drive de 2 TB",
                "Incluida en RecoTime",
                "THE COMPANY DEV los entrega sí o sí con el módulo Drive. Son 2 TB para toda la cuenta.",
            ],
        ],
        [2200, 2500, 4900],
        aligns=["left", "left", "left"],
        size=8,
        header_size=8,
    )
    add_note(
        doc,
        "KVM 8, el plan más alto de esa página, se renueva en $2.026.800 al año. Esa tabla es el costo de la VPS que coloca el cliente. Los 2 TB del Drive los entrega THE COMPANY DEV, sí o sí.",
    )

    add_subheading(doc, "11.4. Licencia vitalicia")
    add_body(
        doc,
        [
            (
                "El pago único equivale a 22 mensualidades del alcance que el cliente contrate, sistema completo o módulos sueltos. Cubre el derecho de uso de la versión entregada.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Incluye el derecho de uso indefinido de la versión entregada, la instalación en la VPS del cliente, la capacitación y 3 meses de soporte de garantía.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Después de la garantía, el cliente asume su propio soporte. Si toma el plan opcional, paga el 40 % de su mensualidad por usuario, recibe 3 horas al mes y las actualizaciones.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            ("El código fuente sigue siendo de ", False),
            ("THE COMPANY DEV", True),
            (
                ". La licencia da derecho de uso. No autoriza reventa, copia ni modificación por terceros.",
                False,
            ),
        ],
    )

    add_subheading(doc, "11.5. Servicios que se cotizan aparte")
    add_table(
        doc,
        ["Servicio", "Valor propuesto"],
        [
            [
                "Hora adicional de soporte o ajustes",
                "$90.000. Es la misma hora ya planteada en los perfiles de RecoRoute y de Campus by TCD.",
            ],
            ["Nómina, PILA o factura electrónica DIAN", "Se cotizan, o permanecen en el software contable del cliente."],
            ["Inventario, compras o POS", "Se cotizan según el alcance."],
            ["Reloj, biométrico o app en tiendas", "Se cotiza. Esta versión registra la marca dentro de RecoTime."],
            ["Sede adicional o espacio por encima de 2 TB", "Se cotiza. Los 2 TB del Drive los entrega THE COMPANY DEV."],
            ["Consumo de WhatsApp", "Lo paga el cliente al proveedor de la línea. No va en la mensualidad."],
            ["VPS y dominio", "Los coloca el cliente. No se facturan dentro de RecoTime."],
        ],
        [4000, 5600],
    )

    add_heading(doc, "12. COMPARACIÓN DE ALCANCE FRENTE AL MERCADO")
    add_table(
        doc,
        ["Proceso", "RecoTime, esta versión", "Dónde aparece en el mercado citado"],
        [
            ["Expediente del cliente y chat interno", "Módulo 1, suelto o con los demás", "Odoo y Zoho CRM llevan el cliente. Bitrix24 suma el chat de la empresa en el mismo producto."],
            ["Llegada, almuerzo, descansos y salida", "Módulo 2, suelto o con los demás", "Odoo Asistencias, Zoho People y Factorial."],
            ["Drive, intranet de 2 TB por cuenta", "Módulo 3, suelto o con los demás", "Google Workspace, Zoho WorkDrive y SharePoint."],
            ["Puesta en marcha y capacitación", "Incluidas", "En Odoo, Bitrix24, HubSpot y Business Central la implementación va aparte o se cotiza."],
            ["VPS y dominio", "Los coloca EL CLIENTE", "Hostinger publica el KVM. RecoTime se instala ahí y no incluye el servidor."],
            ["Contabilidad y factura DIAN", "Fuera de esta versión", "Alegra, Siigo, Odoo y Business Central."],
            ["Nómina y PILA", "Fuera de esta versión", "Alegra Nómina, Siigo Nómina, Aleluya y Buk."],
        ],
        [2300, 3300, 4000],
        size=8,
        header_size=8,
    )

    add_heading(doc, "13. EJEMPLO DE INVERSIÓN")
    add_body(
        doc,
        [
            (
                f"Empresa de 30 usuarios activos con los tres módulos, a {cop(full_user)} por usuario. La factura es solo el software. La VPS y el dominio los coloca el cliente y no entran en estos pagos. Cálculo de este informe.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Modalidad", "Primer pago", "Pagos siguientes", "Total del primer año"],
        [
            ["Mensual", cop(full_user * team), cop(full_user * team) + " cada mes", cop(full_user * team * 12)],
            ["Anualidad", cop(full_user * team * 10), cop(full_user * team * 10) + " al renovar", cop(full_user * team * 10)],
            [
                "Licencia vitalicia",
                cop(full_user * team * 22 / 2) + " (50 %)",
                cop(full_user * team * 22 / 2) + " a la entrega",
                cop(full_user * team * 22),
            ],
            [
                "Vitalicia con soporte",
                cop(full_user * team * 22 / 2) + " (50 %)",
                "Segundo 50 % a la entrega y " + cop(full_user * 0.4 * team) + " al mes desde el mes 4",
                cop(full_user * team * 22 + full_user * 0.4 * team * 9),
            ],
        ],
        [1750, 2300, 3250, 2300],
        aligns=["left", "center", "left", "center"],
        size=8,
        header_size=8,
    )
    add_note(
        doc,
        f"Los dos pagos de la licencia vitalicia caen en el primer año: {cop(full_user * team * 22)}. El soporte opcional son nueve meses si el arranque es en enero: 9 × {cop(full_user * 0.4 * team)} = {cop(full_user * 0.4 * team * 9)}. Sumados a la licencia dan {cop(full_user * team * 22 + full_user * 0.4 * team * 9)}. La VPS queda fuera de esta tabla: la coloca el cliente. Si renueva un KVM 1 todo el año, paga {cop(kvm1_renew * 12)} a Hostinger.",
    )
    add_body(
        doc,
        [
            (
                f"Si esa empresa toma solo el CRM y ERP, paga {cop(price_crm)} por usuario: {cop(price_crm * team)} al mes, o {cop(price_crm * team * 10)} en la anualidad. La asistencia sola son {cop(price_attendance * team)} al mes y el Drive solo {cop(price_drive * team)}. La licencia regional de Odoo Standard sigue en {cop(usd_to_cop(odoo_co_std))} por usuario, {cop(usd_to_cop(odoo_co_std * team))} al mes para 30 usuarios, sin implementación. Los tres módulos de RecoTime son {cop(full_user * team)} al mes. Google Workspace Business Standard, para las mismas 30 personas, son {cop(google_standard_cop * team)} al mes e incluyen correo y 2 TB por persona.",
                False,
            )
        ],
    )

    add_heading(doc, "14. CONDICIONES DEL SERVICIO")
    add_table(
        doc,
        ["Condición", "Detalle"],
        [
            ["Horario de soporte", "Lunes a viernes, de 9:00 a.m. a 5:00 p.m., por correo y WhatsApp."],
            ["Tiempo de respuesta", "Primera respuesta en máximo 24 horas hábiles."],
            ["Horas incluidas", "3 horas al mes, no acumulables. Aplican a la mensualidad, a la anualidad y al plan de soporte."],
            ["Pago mensual", "Dentro de los primeros 5 días de cada mes."],
            ["Pago anual", "Un solo pago al iniciar el periodo."],
            ["Pago de la licencia vitalicia", "50 % al iniciar y 50 % a la entrega."],
            ["Medios de pago", "Nequi, transferencia o efectivo."],
            ["Fidelización", "1 mes gratis por cada año completo en la suscripción mensual, entregado al final del año. No aplica a la anualidad ni a la licencia vitalicia."],
            [
                "Crecimiento",
                f"Se cobra por usuario activo y por módulo. CRM y ERP {cop(price_crm)}, asistencia {cop(price_attendance)}, Drive {cop(price_drive)}.",
            ],
            ["VPS y dominio", "Los coloca EL CLIENTE: los contrata, los paga, entrega el acceso y los renueva."],
            ["2 TB del Drive", "Los entrega THE COMPANY DEV, sí o sí, cuando el cliente toma el módulo Drive."],
            ["Información de la empresa", "El calendario de 8 a 12 semanas corre cuando el cliente ya colocó la VPS y el dominio, y entregó organigrama, trabajadores, clientes y horario."],
        ],
        [2800, 6800],
    )

    add_heading(doc, "15. RESPONSABILIDADES Y PROPIEDAD INTELECTUAL")
    add_bullet(
        doc,
        [
            ("VPS y dominio: los coloca ", False),
            ("EL CLIENTE", True),
            (". Los contrata, los paga, entrega el acceso y los renueva. ", False),
            ("THE COMPANY DEV", True),
            (
                " instala RecoTime en la VPS que el cliente colocó y no responde por caídas, lentitud, pérdida de información, vencimientos, renovaciones ni fallas del proveedor. Los 2 TB del Drive los entrega THE COMPANY DEV, sí o sí.",
                False,
            ),
        ],
    )
    add_bullet(
        doc,
        [
            ("Certificado SSL y correos institucionales: a cargo de ", False),
            ("EL CLIENTE", True),
            (".", False),
        ],
    )
    add_bullet(
        doc,
        [
            ("Información de la operación: ", False),
            ("EL CLIENTE", True),
            (
                " entrega a tiempo el organigrama, los trabajadores, los clientes, los procesos y el horario de llegada, almuerzo y descansos. Los retrasos mueven el calendario.",
                False,
            ),
        ],
    )
    add_bullet(
        doc,
        [
            ("Nómina, factura electrónica y consumo de mensajería: a cargo de ", False),
            ("EL CLIENTE", True),
            (", en el proveedor que elija.", False),
        ],
    )
    add_bullet(
        doc,
        [
            ("Código fuente: es propiedad de ", False),
            ("THE COMPANY DEV", True),
            (" en todas las modalidades.", False),
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Datos personales de trabajadores y clientes: se tratan conforme a la Ley 1581 de 2012. La política de tratamiento y el aviso de privacidad de la empresa son responsabilidad de ",
                False,
            ),
            ("EL CLIENTE", True),
            (".", False),
        ],
    )

    add_heading(doc, "16. PUNTOS PARA APROBACIÓN DE LA DIRECTIVA")
    add_bullet(
        doc,
        [
            (
                f"Aprobar el precio acoplado al mercado: CRM y ERP {cop(price_crm)}, asistencia {cop(price_attendance)} y Drive {cop(price_drive)} por usuario activo al mes. Los tres juntos son {cop(full_user)} por usuario, que es la suma.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                f"Confirmar el reparto: el módulo 1 es la trazabilidad de la gestión de clientes, el chat interno y el ERP de áreas y procesos; el módulo 2 es la asistencia; el módulo 3 es el Drive. Para 30 usuarios, el módulo 1 son {cop(price_crm * team)} al mes y los tres módulos son {cop(full_user * team)}.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Usar la sección 10.1 con el cliente: cada módulo cabe en su referencia y los tres juntos quedan entre las dos licencias de Odoo en Colombia. Los 2 TB del Drive los entrega THE COMPANY DEV, sí o sí.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [
            (
                "Definir si los tres primeros clientes congelan por 24 meses el valor del plan que firmen, sin bajar la tarifa.",
                False,
            )
        ],
    )
    add_bullet(
        doc,
        [("Aprobar el calendario de 8 a 12 semanas y las 3 horas mensuales de soporte.", False)],
    )
    add_bullet(
        doc,
        [("Aprobar la hora adicional de $90.000, en la misma referencia de RecoRoute y de Campus by TCD.", False)],
    )
    add_bullet(
        doc,
        [
            (
                "Dejar por escrito, en cada propuesta, que la VPS y el dominio los coloca el cliente, que la mensualidad es solo el software, que la nómina y la factura electrónica quedan por fuera, y que las marcas las registra el trabajador o el supervisor.",
                False,
            )
        ],
    )
    add_heading(doc, "17. PARTE SOCIAL")
    add_body(
        doc,
        [
            (
                "Más adelante, RecoTime puede acompañar a un emprendimiento que busque o ya tenga recursos del Fondo Emprender del SENA. Esta versión no incluye ese trámite. El fondo, en su reglamento de 2026, financia por dos líneas: Crear, para una iniciativa nueva, y Fortalecimiento productivo, para un negocio ya registrado. En las dos mira ingresos, autoempleo y puestos de trabajo. Esos son los temas sociales que este servicio toca, y de forma leve.",
                False,
            )
        ],
    )
    add_table(
        doc,
        ["Tema social", "En qué puede ayudar RecoTime"],
        [
            [
                "Autoempleo y puestos de trabajo",
                "La asistencia deja quién llegó, almorzó, descansó y salió. Sirve para ver el equipo que la iniciativa dice tener o retener.",
            ],
            [
                "Orden de la unidad productiva",
                "El expediente del cliente y el chat interno dejan cómo se atiende el negocio, dentro de la empresa.",
            ],
            [
                "Continuidad del emprendimiento",
                "El Drive guarda actas, formatos y la gestión. Los 2 TB los entrega THE COMPANY DEV.",
            ],
        ],
        [2800, 6800],
        size=9,
        header_size=9,
    )
    add_body(
        doc,
        [
            (
                "La ayuda es de operación. La postulación, los indicadores ante el fondo y la condonación siguen a cargo de quien presente el plan.",
                False,
            )
        ],
    )

    add_body(
        doc,
        [("Elaborado por", True)],
        before=14,
        after=0,
        align="left",
        keep_next=True,
    )
    add_body(
        doc,
        [("Jhon Javier Angulo Hilarion", False)],
        before=0,
        after=0,
        align="left",
        keep_next=True,
    )
    add_body(
        doc,
        [("Ejecutivo Comercial", False)],
        before=0,
        after=0,
        align="left",
        keep_next=True,
    )
    add_body(
        doc,
        [("THE COMPANY DEV", True)],
        before=0,
        after=0,
        align="left",
        keep_next=True,
    )
    add_contacts(doc)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
