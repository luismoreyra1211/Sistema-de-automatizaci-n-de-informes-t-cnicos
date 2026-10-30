import os
import random
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# carpeta donde guardarás informes
OUTPUT_DIR = os.path.join(Path.home(),"Downloads")
imagenes = [
    os.path.join(BASE_DIR, "imagen_1.png"),
    os.path.join(BASE_DIR, "imagen_2.png"),
    os.path.join(BASE_DIR, "imagen_3.png"),
    os.path.join(BASE_DIR, "imagen_4.png"),
    os.path.join(BASE_DIR, "imagen_5.png"),
]


def configurar_estilos(doc):
    style = doc.styles['Heading 1']
    font = style.font

    font.name = 'Calibri'
    font.size = Pt(10)
    font.bold = True
    font.color.rgb = RGBColor(0, 0, 0)

    pf = style.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)

def aplicar_formato_parrafo(paragraph, size=10, bold=False):
    for run in paragraph.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.font.bold = bold

def configurar_estilos(doc):
    style = doc.styles['Heading 1']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(10)
    font.bold = True
    font.color.rgb = RGBColor(0, 0, 0)  # negro
    pf = style.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)

def agregar_titulo_1(doc, texto):
    return doc.add_heading(texto, level=1)

def agregar_titulo_2(doc, texto):
    return doc.add_heading(texto, level=2)

def set_paragraph_border(paragraph, position="top"):
    p = paragraph._p
    pPr = p.get_or_add_pPr()

    pbdr = OxmlElement('w:pBdr')
    border = OxmlElement(f'w:{position}')

    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), '8')
    border.set(qn('w:space'), '1')
    border.set(qn('w:color'), '000000')

    pbdr.append(border)
    pPr.append(pbdr)

def add_page_number(paragraph):

    paragraph.add_run("Página ")

    # PAGE
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

    paragraph.add_run(" de ")

    # NUMPAGES
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "NUMPAGES"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

def crear_header(doc, info, titulo, cliente, codigo, ruta_logo=None):

    section = doc.sections[0]
    header = section.header

    tabla = header.add_table(rows=2, cols=3, width=Inches(3))

    # =========================
    # UTILIDADES
    # =========================
    def set_font_size(cell, size=9):
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(size)

    def set_borders(table):
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblBorders = OxmlElement('w:tblBorders')

        for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            element = OxmlElement(f'w:{edge}')
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), '8')
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), '000000')
            tblBorders.append(element)

        tblPr.append(tblBorders)

    def set_cell_width(cell, width):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(width * 1440)))
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)

    # =========================
    # FILA 1
    # =========================

    # LOGO (ahora opcional en local)
    cell_logo = tabla.rows[0].cells[0]
    cell_logo.text = ""

    if ruta_logo:
        p = cell_logo.paragraphs[0]

        # CONTROL DE ESPACIADO
        p.paragraph_format.space_before = Pt(0.1)
        p.paragraph_format.space_after = Pt(0.1)
        p.paragraph_format.line_spacing = 1

        # OPCIONAL: centrar
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p.add_run()
        run.add_picture(ruta_logo, width=Inches(0.8))

    # TÍTULO
    cell_centro = tabla.rows[0].cells[1]
    cell_centro.text = titulo

    p = cell_centro.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)

    aplicar_formato_parrafo(p, size=9, bold=True)

    # CÓDIGO
    cell_derecha = tabla.rows[0].cells[2]
    cell_derecha.text = codigo

    p2 = cell_derecha.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(6)
    p2.paragraph_format.space_after = Pt(0)

    aplicar_formato_parrafo(p2, size=9, bold=True)
    # =========================
    # FILA 2
    # =========================

    # INFO
    cell = tabla.rows[1].cells[0]
    cell.text = info
    p_info = cell.paragraphs[0]
    aplicar_formato_parrafo(p_info, size=7, bold=True)

    # CLIENTE
    cell_cliente = tabla.rows[1].cells[1]
    cell_cliente.text = cliente
    p_cliente = cell_cliente.paragraphs[0]
    p_cliente.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_formato_parrafo(p_cliente, size=8, bold=True)

    # PÁGINA
    cell_pagina = tabla.rows[1].cells[2]
    cell_pagina.text = ""

    p3 = cell_pagina.paragraphs[0]
    add_page_number(p3)
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # =========================
    # FORMATO FINAL
    # =========================
    for row in tabla.rows:
        for cell in row.cells:
            set_font_size(cell, 9)

    set_borders(tabla)

    tabla.autofit = False
    tabla.allow_autofit = False

    set_cell_width(tabla.rows[0].cells[0], 1)
    set_cell_width(tabla.rows[0].cells[1], 4.0)
    set_cell_width(tabla.rows[0].cells[2], 1)

def crear_footer(doc, ruta_img_izq, ruta_img_der):

    section = doc.sections[0]
    footer = section.footer

    p = footer.paragraphs[0]
    p.clear()

    # =========================
    # CALCULAR ANCHO REAL
    # =========================
    page_width = section.page_width
    margen_izq = section.left_margin
    margen_der = section.right_margin

    ancho_util = page_width - margen_izq - margen_der

    # =========================
    # TABULACIÓN DERECHA EXACTA
    # =========================
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.clear_all()  # importante
    tab_stops.add_tab_stop(ancho_util)

    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # =========================
    # IMAGEN IZQUIERDA
    # =========================
    run = p.add_run()
    run.add_picture(ruta_img_izq,width=Inches(1.5),height=Inches(0.6))

    # =========================
    # SALTO A DERECHA
    # =========================
    p.add_run("\t")

    # =========================
    # IMAGEN DERECHA
    # =========================
    run = p.add_run()
    run.add_picture(ruta_img_der,width=Inches(1.8),height=Inches(0.6))

def agregar_imagen_a_celda(cell, ruta_imagen):
    paragraph = cell.paragraphs[0]

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1

    run = paragraph.add_run()
    run.add_picture(ruta_imagen, width=Inches(0.6))

    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def set_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr

    borders = OxmlElement('w:tblBorders')

    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '8')
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), '000000')
        borders.append(element)

    tblPr.append(borders)

def agregar_portada(
    doc,
    inf,
    sede,
    cliente,
    direccion,
    cotizacion,
    atencion,
    servicio,
    fecha_inicio,
    fecha_fin,
    elaborado_por,
    revisado_por
):

    # =========================
    # LÍNEA SUPERIOR
    # =========================
    linea_arriba = doc.add_paragraph()
    set_paragraph_border(linea_arriba, "bottom")
    linea_arriba.paragraph_format.space_after = Pt(6)

    # =========================
    # TÍTULO 1
    # =========================
    titulo1 = doc.add_paragraph()
    run1 = titulo1.add_run(f"INFORME TÉCNICO: {inf} - 2026")
    run1.font.name = "Calibri"
    run1.bold = True
    run1.font.size = Pt(12)

    # =========================
    # TÍTULO 2
    # =========================
    titulo2 = doc.add_paragraph()
    texto_titulo2 = f"INFORME TÉCNICO DE MANTENIMIENTO PREVENTIVO DE TABLEROS ELECTRICOS\n{sede}"
    run2 = titulo2.add_run(texto_titulo2.upper())

    run2.font.name = "Calibri"
    run2.bold = True
    run2.font.size = Pt(11)

    titulo2.paragraph_format.space_after = Pt(6)

    # =========================
    # LÍNEA INFERIOR
    # =========================
    linea_abajo = doc.add_paragraph()
    set_paragraph_border(linea_abajo, "top")
    linea_abajo.paragraph_format.space_before = Pt(6)

    # =========================
    # TABLA DE DATOS
    # =========================
    tabla = doc.add_table(rows=8, cols=2)
    tabla.autofit = False

    tabla.columns[0].width = Inches(3)
    tabla.columns[1].width = Inches(7)

    datos = [
        ("Cliente", cliente),
        ("Dirección", direccion),
        ("Cotización", cotizacion),
        ("Atención", atencion),
        ("Servicio", f"MANTENIMIENTO PREVENTIVO DE TABLEROS ELÉCTRICOS – {sede}"),
        ("Fecha de ejecución", f"{fecha_inicio}\n{fecha_fin}"),
        ("Elaborado por", elaborado_por),
        ("Revisado por", revisado_por)
    ]

    for j, (campo, valor) in enumerate(datos):

        # Columna izquierda
        cell1 = tabla.rows[j].cells[0]
        cell1.text = campo
        aplicar_formato_parrafo(cell1.paragraphs[0], size=12, bold=True)

        # Columna derecha
        cell2 = tabla.rows[j].cells[1]
        texto = f":  {valor}"

        if j == 4:  # SERVICIO en mayúscula
            texto = texto.upper()

        cell2.text = texto
        aplicar_formato_parrafo(cell2.paragraphs[0], size=11, bold=(j == 4))


    # =========================
    # ESTILOS
    # =========================
    configurar_estilos(doc)

def agregar_objetivo(doc, direccion, cotizacion):
    texto = (
        f"El presente informe describe y detalla mediante un panel fotográfico "
        f"las actividades que se realizaron según número de cotización {cotizacion}. "
        f"El trabajo se llevó a cabo en {direccion}."
    )

    p = doc.add_paragraph(texto)

    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)

def agregar_alcance(doc, servicio):
    p1 = doc.add_paragraph("El presente servicio comprende lo siguiente:")
    p2 = doc.add_paragraph(servicio.upper())

    for p in [p1, p2]:
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(10)

    for run in p2.runs:
        run.bold = True

def agregar_ejecucion(doc, fecha_inicio, fecha_fin):

    f1 = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    f2 = datetime.strptime(fecha_fin, "%Y-%m-%d")

    meses = [
        "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]

    texto = f"El servicio se realizó durante el {f1.day} al {f2.day} de {meses[f1.month-1]} del {f1.year} y se realizaron las siguientes actividades:"
    p = doc.add_paragraph(texto)

    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)

    actividades = [
        "Verificación externa e interna de tablero.",
        "Verificación de temperatura de los componentes del tablero eléctrico con cámara termográfica.",
        "Revisar y ordenar el cableado eléctrico interno en el tablero eléctrico.",
        "Verificación de los componentes internos de los tableros eléctricos.",
        "Realización de la limpieza interna y externa de los tableros, cables y pletinas de cobre usando herramientas y solvente dieléctrico.",
        "Verificación de funcionamiento de los equipos de medición y lámparas de señalización del tablero.",
        "Verificación de par de apriete en bornes y conexiones del tablero.",
        "Verificación de la existencia de directorios y diagramas unifilares dentro del tablero.",
        "Revisión del sistema de puesta a tierra y verificación de continuidad del circuito PAT.",
        "Verificación de empalmes, rótulos y terminales de conductores eléctricos dentro del tablero."
    ]

    for act in actividades:
        p = doc.add_paragraph(act, style='List Bullet')
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(10)

def limpiar_parrafo_tabla(cell):
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

def agregar_tabla_epp(doc):

    p = doc.add_paragraph("Equipos de Protección Personal (EPP):")
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.bold = True

    tabla = doc.add_table(rows=6, cols=3)
    tabla.autofit = False

    set_table_borders(tabla)

    ancho_col = [Inches(0.8), Inches(1.5), Inches(4)]

    for row in tabla.rows:
        for i, cell in enumerate(row.cells):
            cell.width = ancho_col[i]

    headers = ["EPP", "DESCRIPCIÓN", "FUNCIÓN"]

    for i, h in enumerate(headers):
        cell = tabla.rows[0].cells[i]
        cell.text = h
        limpiar_parrafo_tabla(cell)

        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            for run in paragraph.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(10)
                run.font.bold = True
    
    for cell in tabla.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), "F2F2F2")

        tcPr.append(shd)

    # -------------------------
    # DATOS (SOLO TEXTO)
    # -------------------------
    datos = [
        ["1","Guantes",
         "Protección contra el riesgo a sufrir lesiones por rozamiento con material abrasivo o superficies salientes e irregulares. Protección de riesgo eléctrico según nivel de tensión."],

        ["2","Zapatos de seguridad dieléctricos",
         "Protección contra el riesgo de contacto con cables de tensión en áreas eléctricas."],

        ["3","Uniforme de trabajo completo",
         "Protección contra el riesgo de contacto con sustancias peligrosas o material abrasivo."],

        ["4","Casco de seguridad",
         "Protección contra golpes por caída de objetos en trabajos en altura y riesgo eléctrico."],

        ["5","Lentes de seguridad",
         "Protección contra partículas y polvo proyectado."]
    ]

    for i, fila in enumerate(datos):
        for j, val in enumerate(fila):
            cell = tabla.rows[i+1].cells[j]

            if j == 0:
                cell.text = ""
                agregar_imagen_a_celda(cell, imagenes[i])
            else:
                cell.text = val

                limpiar_parrafo_tabla(cell)

                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Calibri"
                        run.font.size = Pt(10)

def agregar_requisitos_seguridad(doc):

    # -------------------------
    # TEXTO INICIAL
    # -------------------------
    p = doc.add_paragraph(
        "Todo el personal de campo involucrado en el desarrollo de los trabajos:"
    )

    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)

    # -------------------------
    # LISTA CON VIÑETAS
    # -------------------------
    lista = [
        "Cuenta con las pólizas de seguros SCTR contra todo riesgo.",
        "Cuenta con inducción de trabajos a realizar por parte del jefe encargado de la cuadrilla.",
        "Todos los equipos, herramientas y materiales utilizados durante la inspección eléctrica fueron revisados y se comprobó su operatividad mediante formatos de verificación (check list)."
    ]

    for item in lista:
        p = doc.add_paragraph(item, style='List Bullet')
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(10)

    # -------------------------
    # TEXTO HERRAMIENTAS
    # -------------------------
    p = doc.add_paragraph("Para los trabajos se utilizó lo siguiente:")
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)

    p = doc.add_paragraph("Herramientas utilizadas:")
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.bold = True

    herramientas = [
        "Cámara termográfica portátil marca Guide modelo BRITIR.",
        "Voltímetro.",
        "Pinza amperimétrica.",
        "Desarmadores aislados.",
        "Alicates aislados.",
        "Fluido dieléctrico.",
        "Trapos.",
        "Soplete."
    ]

    for h in herramientas:
        p = doc.add_paragraph(h, style='List Bullet')
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(10)

    agregar_tabla_epp(doc)

def agregar_norma(doc):

    # =========================
    # PÁRRAFO 1
    # =========================
    p1 = doc.add_paragraph(
        "Para un correcto mantenimiento preventivo del sistema eléctrico, se aplicó para el presente trabajo el análisis termográfico mediante una cámara portátil marca GUIDE. "
        "Con el análisis termográfico podemos supervisar el funcionamiento y estado de los equipos a intervenir, permitiendo dar soluciones y/o recomendaciones eficientes a los diversos datos que se obtienen durante la inspección."
    )

    for run in p1.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)

    # =========================
    # PÁRRAFO 2
    # =========================
    p2 = doc.add_paragraph(
        "Las especificaciones de temperatura varían según el tipo exacto del equipo. El calentamiento generalmente está relacionado con la emisión de corriente; por lo tanto, la corriente de carga tendrá un gran impacto en ΔT. "
        "Teniendo en cuenta lo anteriormente mencionado, para el mantenimiento de los equipos nos basaremos en la norma ANSI/NETA ATS-2011, en donde nos proporciona una tabla de valores razonables para poder diagnosticar el estado y funcionamiento de los equipos."
    )

    for run in p2.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)

    # =========================
    # TABLA
    # =========================
    tabla = doc.add_table(rows=7, cols=3)
    tabla.autofit = False

    set_table_borders(tabla)

    # =========================
    # ANCHOS
    # =========================
    ancho_col = [Inches(2), Inches(2), Inches(2)]

    for row in tabla.rows:
        for i, cell in enumerate(row.cells):
            cell.width = ancho_col[i]

    # =========================
    # FILA 1: TABLA 100.18 (MERGE)
    # =========================
    cell1 = tabla.rows[0].cells[0]
    cell1.text = "TABLA 100.18"
    cell1.merge(tabla.rows[0].cells[1])
    cell1.merge(tabla.rows[0].cells[2])

    for p in cell1.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.bold = True

    # =========================
    # FILA 2: SUBTÍTULO (MERGE)
    # =========================
    cell2 = tabla.rows[1].cells[0]
    cell2.text = "ACCIONES SUGERIDAS BASADAS EN EL AUMENTO DE LA TEMPERATURA"
    cell2.merge(tabla.rows[1].cells[1])
    cell2.merge(tabla.rows[1].cells[2])

    for p in cell2.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.bold = True

    # =========================
    # ENCABEZADOS (FILA 3)
    # =========================
    tabla.rows[2].cells[0].text = "Diferencia de temperatura (∆T) basada en comparaciones entre componentes similares bajo carga igual"
    tabla.rows[2].cells[1].text = "Diferencia de temperatura (∆T) basada en comparaciones entre la temperatura del componente y del medio ambiente"
    tabla.rows[2].cells[2].text = "Acción sugerida"

    # =========================
    # DATOS
    # =========================
    datos = [
        ["1° C – 3° C", "1° C – 10° C", "Posible deficiencia; justifica la investigación."],
        ["4° C – 15° C", "11° C – 20° C", "Indica deficiencia probable; reparar cuando el tiempo lo permita."],
        ["------", "21° C – 40° C", "Monitorear hasta que se puedan realizar las medidas correctivas."],
        [">15° C", ">40° C", "Discrepancia mayor; reparar inmediatamente."]
    ]

    for i, fila in enumerate(datos):
        for j, val in enumerate(fila):
            tabla.rows[i + 3].cells[j].text = val

    # =========================
    # FORMATO LIMPIO
    # =========================
    for row in tabla.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1

                for r in p.runs:
                    r.font.name = "Calibri"
                    r.font.size = Pt(10)

    p3 = doc.add_paragraph(
        "Fuente:(ANSI/NETA ATS-2011)."
        )

    for run in p3.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.bold = True

    for i in range(3):  # filas 0, 1, 2
        for cell in tabla.rows[i].cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), "F2F2F2")

            tcPr.append(shd)

def obtener_termografia_random():

    carpeta = Path(
        r"C:\Users\USER\Documents\AIO SAC\codigo_python\camara_termografica"
    )

    # verificar existencia

    if not carpeta.exists():

        print(

            "Carpeta termográfica no encontrada"

        )

        return None

    extensiones = {

        ".jpg",

        ".jpeg",

        ".png",

        ".bmp"

    }

    imagenes = []

    for archivo in carpeta.iterdir():

        if not archivo.is_file():

            continue

        if archivo.suffix.lower() not in extensiones:

            continue

        imagenes.append(

            str(
                archivo
            )

        )

    if len(imagenes) == 0:

        print(

            "No existen imágenes termográficas"

        )

        return None

    return random.choice(

        imagenes

    )

def agregar_resumen_actividades(doc, tableros):

    # =========================
    # TEXTO PREVIO
    # =========================
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run1 = p1.add_run("RELACIÓN DE TABLEROS A INTERVENIR")
    run1.font.name = "Calibri"
    run1.bold = True
    run1.font.size = Pt(10)

    p2 = doc.add_paragraph()
    run2 = p2.add_run("A continuación, se detallan los tableros a intervenir:")
    run2.font.name = "Calibri"
    run2.font.size = Pt(10)

    # =========================
    # CANTIDAD DE FILAS
    # =========================
    n = len(tableros)
    filas_totales = 2 + n

    # =========================
    # TABLA BASE
    # =========================
    tabla = doc.add_table(rows=filas_totales, cols=6)
    tabla.style = "Table Grid"

    # =========================
    # ANCHOS
    # =========================
    widths = [1, 6, 6, 5, 5, 2]

    for row in tabla.rows:
        for i, width in enumerate(widths):
            row.cells[i].width = Cm(width)

    # =========================
    # FILA 0: TÍTULO
    # =========================
    fila_titulo = tabla.rows[0].cells
    fila_titulo[0].text = "RELACION DE TABLEROS ELECTRICO INTERVENIDOS"

    for i in range(1, 6):
        fila_titulo[0].merge(fila_titulo[i])

    for p in fila_titulo[0].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(10)
            run.bold = True

    # Color filas 0 y 1
    for i in range(2):
        for cell in tabla.rows[i].cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), "F2F2F2")

            tcPr.append(shd)

    # =========================
    # FILA 1: HEADERS
    # =========================
    fila_header = tabla.rows[1].cells

    fila_header[0].text = "ITEM"
    fila_header[1].text = "DENOMINACIÓN"
    fila_header[3].text = "TIPO"
    fila_header[4].text = "CAPACIDAD"
    fila_header[5].text = "NIVEL"

    fila_header[1].merge(fila_header[2])

    for i in [0, 1, 3, 4, 5]:
        for p in fila_header[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.name = "Calibri"
                run.font.size = Pt(10)

    # =========================
    # FILAS DE DATOS (DESDE JSON)
    # =========================
    for i, t in enumerate(tableros):

        fila = tabla.rows[i + 2].cells

        # usar .get para evitar errores si falta algo
        denominacion = str(t.get("DENOMINACIÓN", "Tablero Eléctrico")).strip()
        nombre_tablero = str(t.get("NOMBRE_TABLERO", t.get("tablero", ""))).strip()
        tipo = str(t.get("TIPO", "")).strip()
        capacidad = str(t.get("CAPACIDAD", "")).strip()
        nivel = str(t.get("NIVEL", "")).strip()

        item = t.get(
            "ITEM",
            i + 1
        )

        fila[0].text = str(item)
        fila[1].text = denominacion
        fila[2].text = nombre_tablero
        fila[3].text = tipo
        fila[4].text = capacidad
        fila[5].text = nivel

        # Centrar columnas clave
        for col in [0, 4, 5]:
            for p in fila[col].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # =========================
    # FORMATO GENERAL
    # =========================
    for row in tabla.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Calibri"

    # =========================
    # TEXTO FINAL
    # =========================
    p3 = doc.add_paragraph()
    run3 = p3.add_run(
        "Cuadro N°1- Relación de Tableros \nFuente: Acciona Ingenieros y Obras S.A.C"
    )
    run3.font.name = "Calibri"
    run3.bold = True
    run3.font.size = Pt(10)

def agregar_panel_fotografico(doc, tableros):

    doc.add_paragraph()

    agregar_titulo_1(doc, "7. PANEL FOTOGRÁFICO")

    n = len(tableros)

    for t in range(n):

        if t > 0:
            doc.add_page_break()
            doc.add_paragraph()

        datos_tablero = tableros[t]

        # =========================
        # TABLA
        # =========================
        tabla = doc.add_table(rows=13, cols=8)
        tabla.autofit = False

        set_table_borders(tabla)

        for row in tabla.rows:
            for cell in row.cells:
                cell.width = Inches(0.8)

        # =========================
        # FORMATO CELDA
        # =========================
        def formatear_celda(celda, bold=True):

            for p in celda.paragraphs:

                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1

                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)
                    run.font.bold = bold

        # =========================
        # FORMATO VALOR
        # =========================
        def formatear_valor(valor, unidad):

            if valor is None:
                return ""

            valor_str = str(valor).strip()

            if valor_str == "" or valor_str == "-":
                return valor_str

            return f"{valor_str}{unidad}"

        # =========================
        # INSERTAR IMAGEN
        # =========================
        def insertar_imagen_en_celda(

            celda,

            ruta,

            ancho=3.48,

            alto=6.21

        ):

            try:

                celda.text = ""

                p = celda.paragraphs[0]

                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                p.paragraph_format.space_before = Pt(0)

                p.paragraph_format.space_after = Pt(0)

                p.paragraph_format.line_spacing = 1

                run = p.add_run()

                run.add_picture(

                    ruta,

                    width=Cm(ancho),

                    height=Cm(alto)

                )

            except Exception as e:

                print(

                    "Error insertando imagen:",

                    ruta,

                    e

                )

        # =========================
        # DATOS TABLERO
        # =========================
        nombre_tablero = str(
            datos_tablero.get("NOMBRE_TABLERO", datos_tablero.get("tablero", ""))
        ).upper()

        capacidad = str(datos_tablero.get("CAPACIDAD", "")).strip()

        # =========================
        # FILA 0
        # =========================
        fila0 = tabla.rows[0].cells

        item = datos_tablero.get(
            "ITEM",
            t + 1
        )

        fila0[0].text = f"{item}. {nombre_tablero}"

        fila0[4].text = capacidad

        fila0[0].merge(fila0[3])
        fila0[4].merge(fila0[7])

        for cell in fila0:
            formatear_celda(cell, True)

        # =========================
        # FILA 1
        # =========================
        fila1 = tabla.rows[1].cells

        fila1[0].text = "IMÁGENES ANTES DEL MANTENIMIENTO"
        fila1[4].text = "IMÁGENES DURANTE DEL MANTENIMIENTO"

        fila1[0].merge(fila1[3])
        fila1[4].merge(fila1[7])

        for cell in fila1:
            formatear_celda(cell, True)

        # =========================
        # FILA 2
        # =========================
        tabla.rows[2].cells[0].merge(tabla.rows[2].cells[1])
        tabla.rows[2].cells[2].merge(tabla.rows[2].cells[3])

        tabla.rows[2].cells[4].merge(tabla.rows[2].cells[5])
        tabla.rows[2].cells[6].merge(tabla.rows[2].cells[7])

        # =========================
        # FILA 3
        # =========================
        fila3 = tabla.rows[3].cells

        fila3[0].text = "IMÁGENES DESPUÉS DEL MANTENIMIENTO"
        fila3[4].text = "CÁMARA TERMOGRÁFICA"

        fila3[0].merge(fila3[3])
        fila3[4].merge(fila3[7])

        for cell in fila3:
            formatear_celda(cell, True)

        # =========================
        # FILAS 4-6
        # =========================
        tabla.cell(4, 0).merge(tabla.cell(6, 1))
        tabla.cell(4, 2).merge(tabla.cell(6, 3))

        tabla.cell(4, 4).merge(tabla.cell(4, 7))

        tabla.cell(5, 5).merge(tabla.cell(5, 6))
        tabla.cell(6, 5).merge(tabla.cell(6, 6))

        tabla.rows[5].cells[4].text = "T° comp"
        tabla.rows[5].cells[5].text = "T° amb"
        tabla.rows[5].cells[7].text = "ΔT"

        for i in [4, 5, 6]:
            for cell in tabla.rows[i].cells:
                formatear_celda(cell, True)

        # =========================
        # FILA 7
        # =========================
        fila7 = tabla.rows[7].cells

        fila7[0].text = "MEDICIONES"
        fila7[4].text = "OBSERVACIONES"

        fila7[0].merge(fila7[3])
        fila7[4].merge(fila7[7])

        for cell in fila7:
            formatear_celda(cell, True)

        # =========================
        # MEDICIONES
        # =========================
        datos = [
            ("L1-L2", "V", ["L1-N", "L1-T"], "V"),
            ("L2-L3", "V", ["L2-N", "L2-T"], "V"),
            ("L1-L3", "V", ["L3-N", "L3-T"], "V"),
            ("IG", "A", ["Ir"], "A"),
            ("Is", "A", ["It"], "A"),
        ]

        for i, fila_data in enumerate(datos):

            fila = tabla.rows[8 + i].cells

            clave1, unidad1, claves2, unidad2 = fila_data

            # =========================
            # IZQUIERDA
            # =========================
            valor1 = formatear_valor(
                datos_tablero.get(clave1, ""),
                unidad1
            )

            fila[0].text = clave1
            fila[1].text = valor1

            # =========================
            # DERECHA
            # =========================
            fila[2].text = "/".join(claves2)

            celda_valor = fila[3]
            celda_valor.text = ""

            p = celda_valor.paragraphs[0]

            for idx, c in enumerate(claves2):

                v = formatear_valor(
                    datos_tablero.get(c, ""),
                    unidad2
                )

                if v != "":

                    run = p.add_run()
                    run.text = v

                    if idx < len(claves2) - 1:
                        run.add_break()

            # =========================
            # FORMATO
            # =========================
            for j in range(4):
                formatear_celda(fila[j], False)

        # =========================
        # OBSERVACIONES
        # =========================
        tabla.cell(8, 4).merge(tabla.cell(12, 7))

        obs = str(datos_tablero.get("Observaciones", ""))

        tabla.cell(8, 4).text = obs

        formatear_celda(tabla.cell(8, 4), False)

        # =========================
        # IMÁGENES
        # =========================
        try:

            antes = datos_tablero.get("antes", [])
            durante = datos_tablero.get("durante", [])
            despues = datos_tablero.get("despues", [])


            antes_sel = random.sample(
                antes,
                min(2, len(antes))
            ) if antes else []

            durante_sel = random.sample(
                durante,
                min(2, len(durante))
            ) if durante else []

            despues_sel = random.sample(
                despues,
                min(2, len(despues))
            ) if despues else []

            # =========================
            # ANTES
            # =========================
            for i, img in enumerate(antes_sel):

                insertar_imagen_en_celda(
                    tabla.rows[2].cells[[0, 2][i]],
                    img
                )

            # =========================
            # DURANTE
            # =========================
            for i, img in enumerate(durante_sel):

                insertar_imagen_en_celda(
                    tabla.rows[2].cells[[4, 6][i]],
                    img
                )

            # =========================
            # DESPUÉS
            # =========================
            celdas_despues = [
                tabla.cell(4, 0),
                tabla.cell(4, 2)
            ]

            for i, img in enumerate(despues_sel):

                insertar_imagen_en_celda(
                    celdas_despues[i],
                    img
                )

            # =========================
            # TERMOGRAFÍA
            # =========================

            ##img_termografia = obtener_termografia_random()

            ##if img_termografia:

            ##    insertar_imagen_en_celda(

            ##        tabla.cell(4,4),

            ##       img_termografia,

            ##        ancho=3.75,

            ##        alto=5.25

            ##    )

        except Exception as e:

            print(
                f"Error en imágenes tablero {t+1}:",
                e
            )

        # =========================
        # FORMATO FINAL
        # =========================
        for row in tabla.rows:

            for cell in row.cells:

                for p in cell.paragraphs:

                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    for run in p.runs:
                        run.font.size = Pt(9)
                        run.font.name = "Calibri"


def generar_informe(datos, tableros):


    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    logo = os.path.join(BASE_DIR, "logo.png")
    pie_izq = os.path.join(BASE_DIR, "pie_pagina_1.jpg")
    pie_der = os.path.join(BASE_DIR, "pie_pagina_2.png")

    doc = Document()

    # =========================
    # HEADER
    # =========================
    crear_header(
        doc,
        info=f"INFO: {datos['inf']}-2026",
        titulo=f"INFORME TÉCNICO DE MANTENIMIENTO PREVENTIVO DE TABLEROS ELECTRICOS - {str(datos['sede']['SEDE']).upper()}",
        cliente=str(datos["sede"]["CLIENTE"]).upper(),
        codigo=str(f"AIO-OPE-{datos['aio']}").upper(),
        ruta_logo=logo
    )

    # =========================
    # FOOTER
    # =========================
    crear_footer(
        doc,
        ruta_img_izq=pie_izq,
        ruta_img_der=pie_der
    )

    # =========================
    # PORTADA
    # =========================
    agregar_portada(
        doc,
        inf=datos["inf"],
        sede=datos["sede"]["SEDE"],
        cliente=datos["sede"]["CLIENTE"],
        direccion=datos["sede"]["DIRECCIÓN"],
        cotizacion=datos["cot"],
        atencion=datos["aten"],
        servicio="MANTENIMIENTO",
        fecha_inicio=str(datos["fecha_inicio"]),
        fecha_fin=str(datos["fecha_fin"]),
        elaborado_por=datos["elaborado"],
        revisado_por=datos["revisado"]
    )


    # =========================
    # CONTENIDO
    # =========================
    doc.add_page_break()
    doc.add_paragraph()
    agregar_titulo_1(doc, "1. OBJETIVO")
    agregar_objetivo(doc, datos["sede"]["DIRECCIÓN"], datos["cot"])

    agregar_titulo_1(doc, "2. ALCANCE DEL SERVICIO")
    agregar_alcance(doc, "Mantenimiento preventivo de tableros eléctricos")

    agregar_titulo_1(doc, "3. EJECUCIÓN DEL SERVICIO")
    agregar_ejecucion(
        doc,
        fecha_inicio=str(datos["fecha_inicio"]),
        fecha_fin=str(datos["fecha_fin"])
    )

    doc.add_page_break()
    doc.add_paragraph()
    agregar_titulo_1(doc, "4. REQUISITOS BÁSICOS DE SEGURIDAD")
    agregar_requisitos_seguridad(doc)

    doc.add_paragraph()
    agregar_titulo_1(doc, "5. NORMA ANSI/NETA ATS-2011")
    agregar_norma(doc)

    doc.add_page_break()
    doc.add_paragraph()
    agregar_titulo_1(doc, "6. RESUMEN DE ACTIVIDADES REALIZADAS")
    agregar_resumen_actividades(doc, tableros)

    doc.add_page_break()
    agregar_panel_fotografico(doc, tableros)

    # =========================
    # SALTO DE PÁGINA
    # =========================
    doc.add_page_break()
    doc.add_paragraph()  # párrafo vacío (equivale a salto de línea)
    agregar_titulo_1(doc, "8. CONCLUSIONES")
    doc.add_paragraph("Conclusiones del trabajo...")

    agregar_titulo_1(doc, "9. RECOMENDACIONES")
    doc.add_paragraph("Recomendaciones finales...")

    # =========================
    # GUARDAR
    # =========================
    output_path = os.path.join(OUTPUT_DIR, f"INFORME TÉCNICO {datos['inf']}-2026 {datos['sede']['CLIENTE']}.docx")
    doc.save(output_path)
    return output_path
