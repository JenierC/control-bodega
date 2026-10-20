import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import sqlite3
from datetime import datetime
import os

def generar_formulario_oficial_bys(ids_solicitudes, departamento="Departamento de Servicios Generales B", estado_firma="Pendiente de Firma", hash_transaccion="", consecutivo_forzado=None):
    os.makedirs("pedimentos_generados", exist_ok=True)
    
    # 1. Control del consecutivo oficial
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
    
    # 2. Extracción de solicitudes desde la base de datos
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

    # 3. Construcción del libro en Excel con la estructura de la Ver 2.01.2021
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pedimento"
    ws.views.sheetView[0].showGridLines = True

    # Tipografía institucional sobria
    font_org = Font(name="Arial", size=8, bold=True)
    font_titulo = Font(name="Arial", size=13, bold=True)
    font_sub_bold = Font(name="Arial", size=8.5, bold=True)
    font_sub_norm = Font(name="Arial", size=8.5)
    font_tbl_hdr = Font(name="Arial", size=8, bold=True)
    font_item = Font(name="Arial", size=7.5)
    font_item_qty = Font(name="Arial", size=8.5, bold=True)
    font_small = Font(name="Arial", size=7)
    font_small_bold = Font(name="Arial", size=7, bold=True)
    font_mini_bold = Font(name="Arial", size=6.5, bold=True)

    thin = Side(border_style="thin", color="000000")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

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
    # ENCABEZADO OFICIAL
    # -------------------------------------------------------------
    ws.merge_cells("A1:C2")
    ws["A1"] = "DIRECCIÓN NACIONAL\nDE DESARROLLO DE\nLA COMUNIDAD"
    ws["A1"].font = font_org
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("D1:G1")
    ws["D1"] = "Pedimento de materiales y suministros"
    ws["D1"].font = font_titulo
    ws["D1"].alignment = align_right

    ws.merge_cells("E2:G2")
    ws["E2"] = "GOBIERNO DE COSTA RICA"
    ws["E2"].font = Font(name="Arial", size=8.5, bold=True)
    ws["E2"].alignment = align_right

    ws["A3"] = "Departamento/Región:"
    ws["A3"].font = font_sub_bold
    ws.merge_cells("B3:D3")
    ws["B3"] = departamento
    ws["B3"].font = font_sub_norm

    ws["E3"] = "Inic. DSG"
    ws["E3"].font = font_sub_bold
    ws["E3"].alignment = align_right

    ws.merge_cells("F3:G3")
    ws["F3"] = f"DSG-{consecutivo}"
    ws["F3"].font = Font(name="Arial", size=9, bold=True)
    ws["F3"].alignment = align_right

    ws["E4"] = ""
    ws["F4"] = "Fecha pedido:"
    ws["F4"].font = font_sub_bold
    ws["F4"].alignment = align_right
    ws["G4"] = datetime.now().strftime("%d/%m/%Y")
    ws["G4"].font = font_sub_norm
    ws["G4"].alignment = align_center

    ws["F5"] = "Consecutivo:"
    ws["F5"].font = font_sub_bold
    ws["F5"].alignment = align_right
    ws["G5"] = consecutivo
    ws["G5"].font = Font(name="Arial", size=9, bold=True)
    ws["G5"].alignment = align_center

    # -------------------------------------------------------------
    # ENCABEZADO DE LA TABLA
    # -------------------------------------------------------------
    ws.merge_cells("A6:A7")
    ws["A6"] = "Partida"
    ws.merge_cells("B6:B7")
    ws["B6"] = "Sub-\npartida"
    ws.merge_cells("C6:C7")
    ws["C6"] = "Código"
    ws.merge_cells("D6:D7")
    ws["D6"] = "Sub-\ncódigo"
    ws.merge_cells("E6:E7")
    ws["E6"] = "Descripción de artículo"
    
    ws.merge_cells("F6:G6")
    ws["F6"] = "Cantidad"
    ws["F7"] = "Solic."
    ws["G7"] = "Desp."

    for r in range(6, 8):
        for c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws[f"{c}{r}"]
            cell.font = font_tbl_hdr
            cell.alignment = align_center
            cell.border = border_cell

    # -------------------------------------------------------------
    # CUERPO DE ARTÍCULOS (Filas 8 a 27)
    # -------------------------------------------------------------
    total_filas = 20
    num_items = len(items)

    for i in range(total_filas):
        row = 8 + i
        ws.row_dimensions[row].height = 13.5

        if i < num_items:
            it = items[i]
            cod_g = str(it[1]).strip()
            desc = str(it[2]).strip()
            cant = int(it[3])

            partida = cod_g[:5] if len(cod_g) >= 5 else "29905"
            subpartida = str(int(cod_g[5:10])) if len(cod_g) >= 10 and cod_g[5:10].isdigit() else (cod_g[5:10] if len(cod_g) >= 10 else "1050")
            codigo = str(int(cod_g[10:16])) if len(cod_g) >= 16 and cod_g[10:16].isdigit() else (cod_g[10:16] if len(cod_g) >= 16 else cod_g)
            subcodigo = str(int(cod_g[16:18])) if len(cod_g) >= 18 and cod_g[16:18].isdigit() else (cod_g[16:18] if len(cod_g) >= 18 else "0")

            ws[f"A{row}"] = partida
            ws[f"B{row}"] = subpartida
            ws[f"C{row}"] = codigo
            ws[f"D{row}"] = subcodigo
            ws[f"E{row}"] = desc
            ws[f"F{row}"] = cant
            ws[f"G{row}"] = ""
        elif i == num_items:
            ws[f"A{row}"] = ""
            ws[f"B{row}"] = ""
            ws[f"C{row}"] = ""
            ws[f"D{row}"] = ""
            ws[f"E{row}"] = "****** UL ******"
            ws[f"F{row}"] = ""
            ws[f"G{row}"] = ""
        else:
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
            ws[f"{col}{row}"].border = border_cell

    # Fila de separación oficial
    ws.merge_cells("A28:G28")
    ws["A28"] = "Productos no registrados"
    ws["A28"].font = Font(name="Arial", size=7.5, italic=True)
    ws["A28"].alignment = align_center
    ws["A28"].border = border_cell
    ws.row_dimensions[28].height = 13

    # -------------------------------------------------------------
    # LOS 4 CUADROS DE SELLOS Y FIRMAS REGLAMENTARIOS (Filas 29 a 35)
    # -------------------------------------------------------------
    # Cuadro 1: Jefatura (Jenier Camacho Picado)
    ws.merge_cells("A29:B32")
    ws["A29"] = "Poner sello aquí"
    ws["A29"].font = Font(name="Arial", size=7, italic=True)
    ws["A29"].alignment = align_center

    ws.merge_cells("A33:B33")
    ws["A33"] = "Jenier Camacho Picado"
    ws["A33"].font = Font(name="Arial", size=8, bold=True)
    ws["A33"].alignment = align_center

    ws.merge_cells("A34:B35")
    ws["A34"] = "Nombre y firma jefe departamento"
    ws["A34"].font = font_small_bold
    ws["A34"].alignment = Alignment(horizontal="center", vertical="top")

    # Cuadro 2: Recibo Conforme
    ws.merge_cells("C29:D33")
    ws["C29"] = ""
    ws.merge_cells("C34:D35")
    ws["C34"] = "Nombre de quién recibe conforme"
    ws["C34"].font = font_small_bold
    ws["C34"].alignment = Alignment(horizontal="center", vertical="top")

    # Cuadro 3: Entrega de Bodega
    ws.merge_cells("E29:E32")
    ws["E29"] = "Sello y Fecha"
    ws["E29"].font = Font(name="Arial", size=7, italic=True)
    ws["E29"].alignment = align_center

    ws.merge_cells("E33:E33")
    ws["E33"] = ""

    ws.merge_cells("E34:E35")
    ws["E34"] = "Encargado de entrega"
    ws["E34"].font = font_small_bold
    ws["E34"].alignment = Alignment(horizontal="center", vertical="top")

    # Cuadro 4: Oficial de Seguridad (Regional)
    ws.merge_cells("F29:G32")
    ws["F29"] = ""
    ws.merge_cells("F33:G33")
    ws["F33"] = "Nombre del oficial de seguridad"
    ws["F33"].font = font_mini_bold
    ws["F33"].alignment = align_center

    ws.merge_cells("F34:G35")
    ws["F34"] = "Oficial de seguridad (En caso de regional)"
    ws["F34"].font = font_mini_bold
    ws["F34"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # Bordes de los 4 cuadros
    for r in range(29, 36):
        for col in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{col}{r}"].border = border_cell

    # -------------------------------------------------------------
    # PIE DE PÁGINA NORMATIVO OFICIAL (Fila 36)
    # -------------------------------------------------------------
    ws.merge_cells("A36:D36")
    ws["A36"] = "Original: Control de bodega. Ver 2.01.2021"
    ws["A36"].font = font_small
    ws["A36"].alignment = align_left

    ws.merge_cells("E36:G36")
    ws["E36"] = "COPIA: departamento solicitante"
    ws["E36"].font = font_small
    ws["E36"].alignment = align_right

    # -------------------------------------------------------------
    # CONFIGURACIÓN EXACTA DE IMPRESIÓN (1 SOLA PÁGINA CARTA)
    # -------------------------------------------------------------
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    # Remueve el pie automático de Excel para que no se corte
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

    ws.print_area = "A1:G36"

    ruta_excel = os.path.join("pedimentos_generados", f"Pedimento_{consecutivo}.xlsx")
    wb.save(ruta_excel)
    return consecutivo, ruta_excel
