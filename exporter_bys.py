import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import sqlite3
from datetime import datetime
import os

def generar_formulario_oficial_bys(ids_solicitudes, departamento="Departamento de Servicios Generales B", estado_firma="Pendiente de Firma", hash_transaccion="", consecutivo_forzado=None):
    os.makedirs("pedimentos_generados", exist_ok=True)
    
    # 1. Consecutivo oficial
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
    
    # 2. Extracción de solicitudes
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
    conn.close()

    # 3. Construcción del Libro Excel con plantilla formal Ver 2.01.2021
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pedimento"
    ws.views.sheetView[0].showGridLines = True

    # Tipografías oficiales sobrias
    font_titulo = Font(name="Arial", size=13, bold=True)
    font_sub = Font(name="Arial", size=8.5, bold=True)
    font_label = Font(name="Arial", size=8)
    font_table_hdr = Font(name="Arial", size=8, bold=True)
    font_item = Font(name="Arial", size=7.5)
    font_item_qty = Font(name="Arial", size=8.5, bold=True)
    font_small = Font(name="Arial", size=6.5)

    thin = Side(border_style="thin", color="000000")
    border_box = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # Anchos de columna oficiales
    ws.column_dimensions['A'].width = 8.5   # Partida
    ws.column_dimensions['B'].width = 8.5   # Sub-partida
    ws.column_dimensions['C'].width = 9.0   # Código
    ws.column_dimensions['D'].width = 8.5   # Sub-código
    ws.column_dimensions['E'].width = 46.0  # Descripción de artículo
    ws.column_dimensions['F'].width = 6.5   # Solic.
    ws.column_dimensions['G'].width = 6.5   # Desp.

    # -------------------------------------------------------------
    # ENCABEZADO INSTITUCIONAL OFICIAL
    # -------------------------------------------------------------
    ws.merge_cells("A1:C2")
    ws["A1"] = "DIRECCIÓN NACIONAL\nDE DESARROLLO DE LA COMUNIDAD\nGOBIERNO DE COSTA RICA"
    ws["A1"].font = Font(name="Arial", size=7.5, bold=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("D1:G1")
    ws["D1"] = "Pedimento de materiales y suministros"
    ws["D1"].font = font_titulo
    ws["D1"].alignment = align_right

    ws.merge_cells("D2:G2")
    ws["D2"] = f"DSG-{consecutivo}"
    ws["D2"].font = Font(name="Arial", size=9, bold=True)
    ws["D2"].alignment = align_right

    ws["A3"] = "Departamento/Región:"
    ws["A3"].font = font_sub
    ws.merge_cells("B3:D3")
    ws["B3"] = departamento
    ws["B3"].font = font_label

    ws["E3"] = "Inic. DSG"
    ws["E3"].font = font_sub
    ws["E3"].alignment = align_right

    ws["F3"] = "Fecha pedido:"
    ws["F3"].font = font_sub
    ws["G3"] = datetime.now().strftime("%d/%m/%Y")
    ws["G3"].font = font_label
    ws["G3"].alignment = align_center

    ws["F4"] = "Consecutivo:"
    ws["F4"].font = font_sub
    ws["G4"] = consecutivo
    ws["G4"].font = Font(name="Arial", size=9, bold=True)
    ws["G4"].alignment = align_center

    # -------------------------------------------------------------
    # TABLA FORMAL DE ARTÍCULOS (Filas 5 y 6)
    # -------------------------------------------------------------
    ws.merge_cells("A5:A6")
    ws["A5"] = "Partida"
    ws.merge_cells("B5:B6")
    ws["B5"] = "Sub-\npartida"
    ws.merge_cells("C5:C6")
    ws["C5"] = "Código"
    ws.merge_cells("D5:D6")
    ws["D5"] = "Sub-\ncódigo"
    ws.merge_cells("E5:E6")
    ws["E5"] = "Descripción de artículo"
    
    ws.merge_cells("F5:G5")
    ws["F5"] = "Cantidad"
    ws["F6"] = "Solic."
    ws["G6"] = "Desp."

    for r in range(5, 7):
        for c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws[f"{c}{r}"]
            cell.font = font_table_hdr
            cell.alignment = align_center
            cell.border = border_box

    # -------------------------------------------------------------
    # CUERPO DE INSUMOS (20 Renglones Oficiales: Filas 7 a 26)
    # -------------------------------------------------------------
    total_filas_tabla = 20
    num_items = len(items)

    for i in range(total_filas_tabla):
        row = 7 + i
        ws.row_dimensions[row].height = 14

        if i < num_items:
            it = items[i]
            cod_g = str(it[1]).strip()
            desc = str(it[2]).strip()
            cant = int(it[3])

            partida = cod_g[:5] if len(cod_g) >= 5 else "29905"
            subpartida = cod_g[5:10] if len(cod_g) >= 10 else "01900"
            codigo = cod_g[10:16] if len(cod_g) >= 16 else cod_g
            subcodigo = cod_g[16:18] if len(cod_g) >= 18 else "00"

            ws[f"A{row}"] = partida
            ws[f"B{row}"] = subpartida
            ws[f"C{row}"] = codigo
            ws[f"D{row}"] = subcodigo
            ws[f"E{row}"] = desc
            ws[f"F{row}"] = cant
            ws[f"G{row}"] = ""
        elif i == num_items:
            # Marca de fin de lista reglamentaria
            ws[f"A{row}"] = ""
            ws[f"B{row}"] = ""
            ws[f"C{row}"] = ""
            ws[f"D{row}"] = ""
            ws[f"E{row}"] = "***** UL *****"
            ws[f"F{row}"] = ""
            ws[f"G{row}"] = ""
        else:
            # Puntos en líneas vacías oficiales
            ws[f"A{row}"] = ""
            ws[f"B{row}"] = ""
            ws[f"C{row}"] = ""
            ws[f"D{row}"] = ""
            ws[f"E{row}"] = "."
            ws[f"F{row}"] = ""
            ws[f"G{row}"] = ""

        ws[f"A{row}"].alignment = align_center
        ws[f"B{row}"].alignment = align_center
        ws[f"C{row}"].alignment = align_center
        ws[f"D{row}"].alignment = align_center
        ws[f"E{row}"].alignment = align_left
        ws[f"F{row}"].alignment = align_center
        ws[f"G{row}"].alignment = align_center

        for col in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{col}{row}"].font = font_item_qty if col == "F" else font_item
            ws[f"{col}{row}"].border = border_box

    # Fila de separación
    ws.merge_cells("A27:G27")
    ws["A27"] = "Productos no registrados"
    ws["A27"].font = Font(name="Arial", size=8, italic=True)
    ws["A27"].alignment = align_center
    ws["A27"].border = border_box
    ws.row_dimensions[27].height = 13

    # -------------------------------------------------------------
    # LOS 4 CUADROS DE SELLOS Y FIRMAS OFICIALES (Filas 28 a 34)
    # -------------------------------------------------------------
    # Cuadro 1: Jefatura Solicitante
    ws.merge_cells("A28:B31")
    ws["A28"] = "Poner sello aquí"
    ws["A28"].font = Font(name="Arial", size=7, italic=True)
    ws["A28"].alignment = align_center

    ws.merge_cells("A32:B32")
    ws["A32"] = "Jenier Camacho Picado"
    ws["A32"].font = Font(name="Arial", size=8, bold=True)
    ws["A32"].alignment = align_center

    ws.merge_cells("A33:B34")
    ws["A33"] = "Nombre y firma jefe departamento"
    ws["A33"].font = Font(name="Arial", size=7, bold=True)
    ws["A33"].alignment = Alignment(horizontal="center", vertical="top")

    # Cuadro 2: Recibido Conforme
    ws.merge_cells("C28:D32")
    ws["C28"] = ""
    ws.merge_cells("C33:D34")
    ws["C33"] = "Nombre de quién recibe conforme"
    ws["C33"].font = Font(name="Arial", size=7, bold=True)
    ws["C33"].alignment = Alignment(horizontal="center", vertical="top")

    # Cuadro 3: Entrega de Bodega
    ws.merge_cells("E28:E32")
    ws["E28"] = "Sello y Fecha"
    ws["E28"].font = Font(name="Arial", size=7, italic=True)
    ws["E28"].alignment = align_center

    ws.merge_cells("E33:E34")
    ws["E33"] = "Encargado de entrega"
    ws["E33"].font = Font(name="Arial", size=7, bold=True)
    ws["E33"].alignment = Alignment(horizontal="center", vertical="top")

    # Cuadro 4: Seguridad (Regional)
    ws.merge_cells("F28:G32")
    ws["F28"] = ""
    ws.merge_cells("F33:G34")
    ws["F33"] = "Nombre del oficial de seguridad\nOficial de seguridad (En caso de regional)"
    ws["F33"].font = Font(name="Arial", size=6.5, bold=True)
    ws["F33"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # Bordes en los cuadros inferiores
    for r in range(28, 35):
        for col in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{col}{r}"].border = border_box

    # -------------------------------------------------------------
    # PIE DE PÁGINA NORMATIVO OFICIAL
    # -------------------------------------------------------------
    ws.merge_cells("A35:D35")
    ws["A35"] = "Original: Control de bodega.  Ver 2.01.2021"
    ws["A35"].font = font_small
    ws["A35"].alignment = align_left

    ws.merge_cells("E35:G35")
    ws["E35"] = "COPIA: departamento solicitante"
    ws["E35"].font = font_small
    ws["E35"].alignment = align_right

    # =============================================================
    # CONFIGURACIÓN DE IMPRESIÓN (1 PÁGINA CARTA EXACTA)
    # =============================================================
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    # Quitar cualquier encabezado o pie automático de Excel
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

    ws.print_area = "A1:G35"

    ruta_excel = os.path.join("pedimentos_generados", f"Pedimento_{consecutivo}.xlsx")
    wb.save(ruta_excel)
    return consecutivo, ruta_excel
