import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import sqlite3
from datetime import datetime
import os

def generar_formulario_oficial_bys(ids_solicitudes, departamento="Departamento de Servicios Generales", estado_firma="Pendiente de Firma", hash_transaccion="", consecutivo_forzado=None):
    os.makedirs("pedimentos_generados", exist_ok=True)
    
    # 1. Obtener consecutivo (usar el forzado para regenerar o incrementar en pedidos nuevos)
    anio_actual = datetime.now().year
    
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import sqlite3
from datetime import datetime
import os

def generar_formulario_oficial_bys(ids_solicitudes, departamento="Departamento de Servicios Generales", estado_firma="Pendiente de Firma", hash_transaccion="", consecutivo_forzado=None):
    os.makedirs("pedimentos_generados", exist_ok=True)
    
    # 1. Obtener consecutivo (usar el forzado para regenerar o incrementar en pedidos nuevos)
    anio_actual = datetime.now().year
    
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS control_consecutivos (
                anio INTEGER PRIMARY KEY,
                ultimo_numero INTEGER NOT NULL
            )
        """)
        cursor.execute("SELECT ultimo_numero FROM control_consecutivos WHERE anio = ?", (anio_actual,))
        row = cursor.fetchone()
        ultimo_num = row[0] if row else 50
        nuevo_num = ultimo_num + 1
        cursor.execute("INSERT OR REPLACE INTO control_consecutivos (anio, ultimo_numero) VALUES (?, ?)", (anio_actual, nuevo_num))
        conn.commit()
        conn.close()
        consecutivo = f"{nuevo_num:03d}-{anio_actual}"
    
    # 2. Obtener datos de las solicitudes
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    placeholders = ",".join("?" * len(ids_solicitudes))
    cursor.execute(f"""
        SELECT s.id, a.codigo_grupal, a.descripcion, s.cantidad, f.nombre, f.espacio_asignado
        FROM solicitudes s
        JOIN articulos a ON s.articulo_id = a.codigo_grupal
        JOIN funcionarios f ON s.funcionario_id = f.id
        WHERE s.id IN ({placeholders})
    """, ids_solicitudes)
    items = cursor.fetchall()
    
    solicitante = items[0][4] if items else "Funcionario"
    espacio = items[0][5] if items else "General"
    conn.close()

    # 3. Crear Libro de Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Pedimento_{consecutivo}"
    ws.views.sheetView[0].showGridLines = True

    # Estilos oficiales
    font_titulo = Font(name="Calibri", size=11, bold=True, color="002060")
    font_subtitulo = Font(name="Calibri", size=9, bold=True, color="333333")
    font_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=8.5)
    font_small = Font(name="Calibri", size=7.5, italic=True)
    font_bold = Font(name="Calibri", size=8.5, bold=True)

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    thin = Side(border_style="thin", color="808080")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # Anchos de columna optimizados para 1 página Carta
    anchos = {
        'A': 10,   # Partida
        'B': 11,   # Subpartida
        'C': 11,   # Código
        'D': 9,    # Subcódigo
        'E': 44,   # Descripción de artículo
        'F': 9,    # Solic.
        'G': 9     # Desp.
    }
    for col, width in anchos.items():
        ws.column_dimensions[col].width = width

    # ENCABEZADO OFICIAL
    ws.merge_cells("A1:G1")
    ws["A1"] = "DIRECCIÓN NACIONAL DE DESARROLLO DE LA COMUNIDAD"
    ws["A1"].font = font_titulo
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:G2")
    ws["A2"] = "DEPARTAMENTO DE BIENES Y SUMINISTROS — BODEGA INSTITUCIONAL"
    ws["A2"].font = font_subtitulo
    ws["A2"].alignment = align_center

    ws.merge_cells("A3:E3")
    ws["A3"] = f"Pedimento de Materiales y Suministros: {departamento}"
    ws["A3"].font = font_bold
    ws["A3"].alignment = align_left

    ws.merge_cells("F3:G3")
    ws["F3"] = f"Pedimento N°: {consecutivo}"
    ws["F3"].font = Font(name="Calibri", size=10, bold=True, color="C00000")
    ws["F3"].alignment = align_center

    ws.merge_cells("A4:E4")
    ws["A4"] = f"Solicitante: {solicitante}  |  Espacio: {espacio}"
    ws["A4"].font = font_data
    ws["A4"].alignment = align_left

    ws.merge_cells("F4:G4")
    ws["F4"] = f"Fecha: {datetime.now().strftime('%d/%m/%Y')}"
    ws["F4"].font = font_data
    ws["F4"].alignment = align_center

    # TABLA DE PRODUCTOS (Fila 6)
    headers = [
        ("A6", "Partida"),
        ("B6", "Subpartida"),
        ("C6", "Código"),
        ("D6", "Subcódigo"),
        ("E6", "Descripción de Artículo"),
        ("F6", "Solic."),
        ("G6", "Desp.")
    ]
    for cell_ref, text in headers:
        ws[cell_ref] = text
        ws[cell_ref].font = font_header
        ws[cell_ref].fill = fill_header
        ws[cell_ref].alignment = align_center
        ws[cell_ref].border = border_cell
    ws.row_dimensions[6].height = 20

    # Llenado de ítems (filas 7 a 16 -> 10 renglones exactos)
    for idx in range(10):
        row_num = 7 + idx
        ws.row_dimensions[row_num].height = 18

        if idx < len(items):
            it = items[idx]
            cod_g = str(it[1]).strip()
            desc = str(it[2]).strip()
            cant = int(it[3])

            partida = cod_g[:5] if len(cod_g) >= 5 else "29905"
            subpartida = cod_g[5:10] if len(cod_g) >= 10 else "01900"
            codigo = cod_g[10:16] if len(cod_g) >= 16 else cod_g
            subcodigo = cod_g[16:18] if len(cod_g) >= 18 else "00"

            ws[f"A{row_num}"] = partida
            ws[f"B{row_num}"] = subpartida
            ws[f"C{row_num}"] = codigo
            ws[f"D{row_num}"] = subcodigo
            ws[f"E{row_num}"] = desc
            ws[f"F{row_num}"] = cant
            ws[f"G{row_num}"] = ""
        else:
            ws[f"A{row_num}"] = ""
            ws[f"B{row_num}"] = ""
            ws[f"C{row_num}"] = ""
            ws[f"D{row_num}"] = ""
            ws[f"E{row_num}"] = ""
            ws[f"F{row_num}"] = ""
            ws[f"G{row_num}"] = ""

        ws[f"A{row_num}"].alignment = align_center
        ws[f"B{row_num}"].alignment = align_center
        ws[f"C{row_num}"].alignment = align_center
        ws[f"D{row_num}"].alignment = align_center
        ws[f"E{row_num}"].alignment = align_left
        ws[f"F{row_num}"].alignment = align_center
        ws[f"G{row_num}"].alignment = align_center

        for col_letter in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{col_letter}{row_num}"].font = font_data
            ws[f"{col_letter}{row_num}"].border = border_cell

    # SECCIÓN INFERIOR: SELLOS Y FIRMAS (Filas 18 a 26)
    ws.row_dimensions[17].height = 6

    # Bloque 1: Firma Jefatura
    ws.merge_cells("A18:D18")
    ws["A18"] = "AUTORIZACIÓN DE JEFATURA"
    ws["A18"].font = font_subtitulo
    ws["A18"].fill = fill_sub
    ws["A18"].alignment = align_center
    ws["A18"].border = border_cell

    ws.merge_cells("A19:D22")
    ws["A19"] = f"Estado: {estado_firma}\nJenier Camacho Picado\nNombre y firma jefe departamento"
    ws["A19"].font = font_data
    ws["A19"].alignment = align_center
    ws["A19"].border = border_cell

    # Bloque 2: Sellos de Bodega
    ws.merge_cells("E18:G18")
    ws["E18"] = "CONTROL Y RECEPCIÓN DE BODEGA"
    ws["E18"].font = font_subtitulo
    ws["E18"].fill = fill_sub
    ws["E18"].alignment = align_center
    ws["E18"].border = border_cell

    ws.merge_cells("E19:G22")
    ws["E19"] = "Espacio para sellos físicos:\nMERCADERÍA DESPACHADA / RETIRADA\nFirma y Fecha de Entrega"
    ws["E19"].font = font_small
    ws["E19"].alignment = align_center
    ws["E19"].border = border_cell

    # Pie institucional normativo
    ws.merge_cells("A24:D24")
    ws["A24"] = "Original: Control de Bodega | Copia: Solicitante"
    ws["A24"].font = font_small
    ws["A24"].alignment = align_left

    ws.merge_cells("E24:G24")
    ws["E24"] = f"Seguridad BCCR: {hash_transaccion[:16]}..." if hash_transaccion else "Validación Digital DINADECO"
    ws["E24"].font = font_small
    ws["E24"].alignment = align_right

    # ===============================================================
    # CONFIGURACIÓN ESTRICTA DE PÁGINA (1 SOLA PÁGINA EXACTA)
    # ===============================================================
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER

    # Forzar a 1 sola página exacta de alto y ancho
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    # Márgenes compactos
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    # Eliminar completamente el pie de página predeterminado de Excel
    ws.oddHeader.left.text = ""
    ws.oddHeader.center.text = ""
    ws.oddHeader.right.text = ""
    ws.oddFooter.left.text = ""
    ws.oddFooter.center.text = ""
    ws.oddFooter.right.text = ""
    ws.evenHeader.left.text = ""
    ws.evenHeader.center.text = ""
    ws.evenHeader.right.text = ""
    ws.evenFooter.left.text = ""
    ws.evenFooter.center.text = ""
    ws.evenFooter.right.text = ""

    # Área de impresión delimitada (Filas 1 a 24)
    ws.print_area = "A1:G24"

    # 4. Guardar archivo
    ruta_excel = os.path.join("pedimentos_generados", f"Pedimento_{consecutivo}.xlsx")
    wb.save(ruta_excel)
    return consecutivo, ruta_excel

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS control_consecutivos (
                anio INTEGER PRIMARY KEY,
                ultimo_numero INTEGER NOT NULL
            )
        """)
        cursor.execute("SELECT ultimo_numero FROM control_consecutivos WHERE anio = ?", (anio_actual,))
        row = cursor.fetchone()
        ultimo_num = row[0] if row else 50
        nuevo_num = ultimo_num + 1
        cursor.execute("INSERT OR REPLACE INTO control_consecutivos (anio, ultimo_numero) VALUES (?, ?)", (anio_actual, nuevo_num))
        conn.commit()
        conn.close()
        consecutivo = f"{nuevo_num:03d}-{anio_actual}"
    
    # 2. Obtener datos de las solicitudes
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    placeholders = ",".join("?" * len(ids_solicitudes))
    cursor.execute(f"""
        SELECT s.id, a.codigo_grupal, a.descripcion, s.cantidad, f.nombre, f.espacio_asignado
        FROM solicitudes s
        JOIN articulos a ON s.articulo_id = a.codigo_grupal
        JOIN funcionarios f ON s.funcionario_id = f.id
        WHERE s.id IN ({placeholders})
    """, ids_solicitudes)
    items = cursor.fetchall()
    
    solicitante = items[0][4] if items else "Funcionario"
    espacio = items[0][5] if items else "General"
    conn.close()

    # 3. Crear Libro de Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Pedimento_{consecutivo}"
    ws.views.sheetView[0].showGridLines = True

    # Estilos oficiales
    font_titulo = Font(name="Calibri", size=11, bold=True, color="002060")
    font_subtitulo = Font(name="Calibri", size=9, bold=True, color="333333")
    font_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=8.5)
    font_small = Font(name="Calibri", size=7.5, italic=True)
    font_bold = Font(name="Calibri", size=8.5, bold=True)

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    thin = Side(border_style="thin", color="808080")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # Anchos de columna optimizados para 1 página Carta
    anchos = {
        'A': 10,   # Partida
        'B': 11,   # Subpartida
        'C': 11,   # Código
        'D': 9,    # Subcódigo
        'E': 44,   # Descripción de artículo
        'F': 9,    # Solic.
        'G': 9     # Desp.
    }
    for col, width in anchos.items():
        ws.column_dimensions[col].width = width

    # ENCABEZADO OFICIAL
    ws.merge_cells("A1:G1")
    ws["A1"] = "DIRECCIÓN NACIONAL DE DESARROLLO DE LA COMUNIDAD"
    ws["A1"].font = font_titulo
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:G2")
    ws["A2"] = "DEPARTAMENTO DE BIENES Y SUMINISTROS — BODEGA INSTITUCIONAL"
    ws["A2"].font = font_subtitulo
    ws["A2"].alignment = align_center

    ws.merge_cells("A3:E3")
    ws["A3"] = f"Pedimento de Materiales y Suministros: {departamento}"
    ws["A3"].font = font_bold
    ws["A3"].alignment = align_left

    ws.merge_cells("F3:G3")
    ws["F3"] = f"Pedimento N°: {consecutivo}"
    ws["F3"].font = Font(name="Calibri", size=10, bold=True, color="C00000")
    ws["F3"].alignment = align_center

    ws.merge_cells("A4:E4")
    ws["A4"] = f"Solicitante: {solicitante}  |  Espacio: {espacio}"
    ws["A4"].font = font_data
    ws["A4"].alignment = align_left

    ws.merge_cells("F4:G4")
    ws["F4"] = f"Fecha: {datetime.now().strftime('%d/%m/%Y')}"
    ws["F4"].font = font_data
    ws["F4"].alignment = align_center

    # TABLA DE PRODUCTOS (Fila 6)
    headers = [
        ("A6", "Partida"),
        ("B6", "Subpartida"),
        ("C6", "Código"),
        ("D6", "Subcódigo"),
        ("E6", "Descripción de Artículo"),
        ("F6", "Solic."),
        ("G6", "Desp.")
    ]
    for cell_ref, text in headers:
        ws[cell_ref] = text
        ws[cell_ref].font = font_header
        ws[cell_ref].fill = fill_header
        ws[cell_ref].alignment = align_center
        ws[cell_ref].border = border_cell
    ws.row_dimensions[6].height = 20

    # Llenado de ítems (filas 7 a 16 -> 10 renglones exactos)
    for idx in range(10):
        row_num = 7 + idx
        ws.row_dimensions[row_num].height = 18

        if idx < len(items):
            it = items[idx]
            cod_g = str(it[1]).strip()
            desc = str(it[2]).strip()
            cant = int(it[3])

            partida = cod_g[:5] if len(cod_g) >= 5 else "29905"
            subpartida = cod_g[5:10] if len(cod_g) >= 10 else "01900"
            codigo = cod_g[10:16] if len(cod_g) >= 16 else cod_g
            subcodigo = cod_g[16:18] if len(cod_g) >= 18 else "00"

            ws[f"A{row_num}"] = partida
            ws[f"B{row_num}"] = subpartida
            ws[f"C{row_num}"] = codigo
            ws[f"D{row_num}"] = subcodigo
            ws[f"E{row_num}"] = desc
            ws[f"F{row_num}"] = cant
            ws[f"G{row_num}"] = ""
        else:
            ws[f"A{row_num}"] = ""
            ws[f"B{row_num}"] = ""
            ws[f"C{row_num}"] = ""
            ws[f"D{row_num}"] = ""
            ws[f"E{row_num}"] = ""
            ws[f"F{row_num}"] = ""
            ws[f"G{row_num}"] = ""

        ws[f"A{row_num}"].alignment = align_center
        ws[f"B{row_num}"].alignment = align_center
        ws[f"C{row_num}"].alignment = align_center
        ws[f"D{row_num}"].alignment = align_center
        ws[f"E{row_num}"].alignment = align_left
        ws[f"F{row_num}"].alignment = align_center
        ws[f"G{row_num}"].alignment = align_center

        for col_letter in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{col_letter}{row_num}"].font = font_data
            ws[f"{col_letter}{row_num}"].border = border_cell

    # SECCIÓN INFERIOR: SELLOS Y FIRMAS (Filas 18 a 26)
    ws.row_dimensions[17].height = 6

    # Bloque 1: Firma Jefatura
    ws.merge_cells("A18:D18")
    ws["A18"] = "AUTORIZACIÓN DE JEFATURA"
    ws["A18"].font = font_subtitulo
    ws["A18"].fill = fill_sub
    ws["A18"].alignment = align_center
    ws["A18"].border = border_cell

    ws.merge_cells("A19:D22")
    ws["A19"] = f"Estado: {estado_firma}\nJenier Camacho Picado\nNombre y firma jefe departamento"
    ws["A19"].font = font_data
    ws["A19"].alignment = align_center
    ws["A19"].border = border_cell

    # Bloque 2: Sellos de Bodega
    ws.merge_cells("E18:G18")
    ws["E18"] = "CONTROL Y RECEPCIÓN DE BODEGA"
    ws["E18"].font = font_subtitulo
    ws["E18"].fill = fill_sub
    ws["E18"].alignment = align_center
    ws["E18"].border = border_cell

    ws.merge_cells("E19:G22")
    ws["E19"] = "Espacio para sellos físicos:\nMERCADERÍA DESPACHADA / RETIRADA\nFirma y Fecha de Entrega"
    ws["E19"].font = font_small
    ws["E19"].alignment = align_center
    ws["E19"].border = border_cell

    # Pie institucional normativo
    ws.merge_cells("A24:D24")
    ws["A24"] = "Original: Control de Bodega | Copia: Solicitante"
    ws["A24"].font = font_small
    ws["A24"].alignment = align_left

    ws.merge_cells("E24:G24")
    ws["E24"] = f"Seguridad BCCR: {hash_transaccion[:16]}..." if hash_transaccion else "Validación Digital DINADECO"
    ws["E24"].font = font_small
    ws["E24"].alignment = align_right

    # ===============================================================
    # CONFIGURACIÓN ESTRICTA DE PÁGINA (1 SOLA PÁGINA EXACTA)
    # ===============================================================
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER

    # Forzar a 1 sola página exacta de alto y ancho
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    # Márgenes compactos
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    # Eliminar completamente el pie de página predeterminado de Excel
    ws.oddHeader.left.text = ""
    ws.oddHeader.center.text = ""
    ws.oddHeader.right.text = ""
    ws.oddFooter.left.text = ""
    ws.oddFooter.center.text = ""
    ws.oddFooter.right.text = ""
    ws.evenHeader.left.text = ""
    ws.evenHeader.center.text = ""
    ws.evenHeader.right.text = ""
    ws.evenFooter.left.text = ""
    ws.evenFooter.center.text = ""
    ws.evenFooter.right.text = ""

    # Área de impresión delimitada (Filas 1 a 24)
    ws.print_area = "A1:G24"

    # 4. Guardar archivo
    ruta_excel = os.path.join("pedimentos_generados", f"Pedimento_{consecutivo}.xlsx")
    wb.save(ruta_excel)
    return consecutivo, ruta_excel
