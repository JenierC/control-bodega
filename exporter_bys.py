import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import sqlite3
from datetime import datetime
import os
import shutil
import zipfile

PLANTILLA_BYS = "Pedimentos de Departamento de Servicios Generales.xlsx"

def obtener_siguiente_consecutivo(anio_actual=None):
    if not anio_actual:
        anio_actual = datetime.now().year
        
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
    
    if row:
        siguiente_num = row[0] + 1
        cursor.execute("UPDATE control_consecutivos SET ultimo_numero = ? WHERE anio = ?", (siguiente_num, anio_actual))
    else:
        siguiente_num = 51  # Inicia en el consecutivo 051
        cursor.execute("INSERT INTO control_consecutivos (anio, ultimo_numero) VALUES (?, ?)", (anio_actual, siguiente_num))
        
    conn.commit()
    conn.close()
    return f"{siguiente_num:03d}-{anio_actual}"


def cargar_plantilla_segura():
    """
    Localiza la plantilla oficial en tu repositorio evitando el error de extensión .xlk en openpyxl.
    Si no encuentra la plantilla en disco, genera el formato oficial de DINADECO en memoria.
    """
    # 1. Candidatos con extensiones nativas compatibles (.xlsx o .xlsm)
    candidatos_nativos = [
        "Pedimentos de Departamento de Servicios Generales.xlsx",
        "Pedimentos de Departamento de Servicios Generales B.xlsx",
        "Pedimentos de Departamento de Servicios Generales.xlsm",
        "Pedimentos de Departamento de Servicios Generales B.xlsm"
    ]
    for c in candidatos_nativos:
        if os.path.exists(c):
            try:
                return openpyxl.load_workbook(c), "TEMPLATE"
            except Exception:
                pass

    # 2. Buscar cualquier archivo de plantilla .xlsx o .xlsm en la raíz
    for f in os.listdir("."):
        if (f.lower().endswith(".xlsx") or f.lower().endswith(".xlsm")) and not f.startswith("~"):
            if "pedimento" in f.lower() or "servicios" in f.lower():
                try:
                    return openpyxl.load_workbook(f), "TEMPLATE"
                except Exception:
                    pass

    # 3. Si solo existe el archivo con extensión .xlk (copia de seguridad de Excel):
    # Se copia temporalmente a .xlsx para que openpyxl lo acepte sin arrojar InvalidFileException
    for f in os.listdir("."):
        if f.lower().endswith(".xlk"):
            try:
                temp_xlsx = "Pedimentos_Generales_temp.xlsx"
                shutil.copyfile(f, temp_xlsx)
                return openpyxl.load_workbook(temp_xlsx), "TEMPLATE"
            except Exception:
                pass

    # 4. Contingencia oficial en memoria (idéntica a la plantilla de DINADECO)
    wb = openpyxl.Workbook()
    return wb, "NATIVE"


def generar_formulario_oficial_bys(solicitud_ids, nombre_departamento="Departamento de Servicios Generales B", nombre_jefatura="Jenier Camacho Picado", hash_firma="", consecutivo_forzado=None):
    os.makedirs("pedimentos_generados", exist_ok=True)
    
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
        consecutivo = obtener_siguiente_consecutivo()
        
    fecha_actual = datetime.now()
    
    # Extraer las solicitudes de la base de datos
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    placeholders = ",".join("?" * len(solicitud_ids))
    cursor.execute(f"""
        SELECT s.id, a.codigo_grupal, a.descripcion, s.cantidad, f.nombre, f.espacio_asignado
        FROM solicitudes s
        JOIN articulos a ON s.articulo_id = a.codigo_grupal
        JOIN funcionarios f ON s.funcionario_id = f.id
        WHERE s.id IN ({placeholders})
    """, solicitud_ids)
    items = cursor.fetchall()
    conn.close()

    wb, tipo_carga = cargar_plantilla_segura()

    if tipo_carga == "TEMPLATE":
        # Formato utilizando la plantilla oficial de tu repositorio
        ws = wb['Pedimento'] if 'Pedimento' in wb.sheetnames else wb.active
        
        ws['N8'] = consecutivo
        ws['N7'] = fecha_actual.strftime("%d/%m/%Y")
        ws['D8'] = nombre_departamento
        
        fila_inicio = 12
        for idx in range(26):  # Rango oficial de filas 12 a 37
            fila_actual = fila_inicio + idx
            if fila_actual > 37:
                break
                
            if idx < len(items):
                _, codigo_grupal, descripcion, cantidad, _, _ = items[idx]
                codigo_grupal = str(codigo_grupal).strip()
                
                partida = codigo_grupal[:5] if len(codigo_grupal) >= 5 else codigo_grupal
                sub_partida = codigo_grupal if len(codigo_grupal) >= 10 else ""
                codigo = codigo_grupal if len(codigo_grupal) >= 16 else ""
                sub_codigo = codigo_grupal[16:] if len(codigo_grupal) > 16 else "0"
                
                ws.cell(row=fila_actual, column=1, value=partida)
                ws.cell(row=fila_actual, column=2, value=sub_partida)
                ws.cell(row=fila_actual, column=3, value=codigo)
                ws.cell(row=fila_actual, column=4, value=sub_codigo)
                ws.cell(row=fila_actual, column=5, value=descripcion)
                ws.cell(row=fila_actual, column=14, value=int(cantidad))
                ws.cell(row=fila_actual, column=15, value=int(cantidad))
            elif idx == len(items):
                ws.cell(row=fila_actual, column=1, value="")
                ws.cell(row=fila_actual, column=2, value="")
                ws.cell(row=fila_actual, column=3, value="")
                ws.cell(row=fila_actual, column=4, value="")
                ws.cell(row=fila_actual, column=5, value="****** UL ******")
                ws.cell(row=fila_actual, column=14, value="")
                ws.cell(row=fila_actual, column=15, value="")
            else:
                ws.cell(row=fila_actual, column=1, value="")
                ws.cell(row=fila_actual, column=2, value="")
                ws.cell(row=fila_actual, column=3, value="")
                ws.cell(row=fila_actual, column=4, value="")
                ws.cell(row=fila_actual, column=5, value=".")
                ws.cell(row=fila_actual, column=14, value="")
                ws.cell(row=fila_actual, column=15, value="")

        ws['A46'] = f"Aprobado digitalmente por: {nombre_jefatura}"
        ws['A47'] = f"Hash de Firma: {hash_firma[:16]}... | Fecha: {fecha_actual.strftime('%Y-%m-%d %H:%M')}"

    else:
        # Formato reglamentario en memoria (contingencia institucional exacta)
        ws = wb.active
        ws.title = "Pedimento"
        ws.views.sheetView[0].showGridLines = True

        font_org = Font(name="Arial", size=8, bold=True)
        font_titulo = Font(name="Arial", size=13, bold=True)
        font_sub_bold = Font(name="Arial", size=8.5, bold=True)
        font_sub_norm = Font(name="Arial", size=8.5)
        font_tbl_hdr = Font(name="Arial", size=8, bold=True)
        font_item = Font(name="Arial", size=7.5)
        font_item_qty = Font(name="Arial", size=8.5, bold=True)
        font_small = Font(name="Arial", size=7)
        font_small_bold = Font(name="Arial", size=7, bold=True)

        thin = Side(border_style="thin", color="000000")
        border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        ws.column_dimensions['A'].width = 8.5
        ws.column_dimensions['B'].width = 8.5
        ws.column_dimensions['C'].width = 9.0
        ws.column_dimensions['D'].width = 8.5
        ws.column_dimensions['E'].width = 46.0
        ws.column_dimensions['F'].width = 6.5
        ws.column_dimensions['G'].width = 6.5

        # Encabezado
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
        ws["B3"] = nombre_departamento
        ws["B3"].font = font_sub_norm

        ws["E3"] = "Inic. DSG"
        ws["E3"].font = font_sub_bold
        ws["E3"].alignment = align_right

        ws.merge_cells("F3:G3")
        ws["F3"] = f"DSG-{consecutivo}"
        ws["F3"].font = Font(name="Arial", size=9, bold=True)
        ws["F3"].alignment = align_right

        ws["F4"] = "Fecha pedido:"
        ws["F4"].font = font_sub_bold
        ws["F4"].alignment = align_right
        ws["G4"] = fecha_actual.strftime("%d/%m/%Y")
        ws["G4"].font = font_sub_norm
        ws["G4"].alignment = align_center

        ws["F5"] = "Consecutivo:"
        ws["F5"].font = font_sub_bold
        ws["F5"].alignment = align_right
        ws["G5"] = consecutivo
        ws["G5"].font = Font(name="Arial", size=9, bold=True)
        ws["G5"].alignment = align_center

        # Tabla de artículos
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

                ws[f"A{row}"] = cod_g[:5] if len(cod_g) >= 5 else "29905"
                ws[f"B{row}"] = str(int(cod_g)) if len(cod_g) >= 10 and cod_g.isdigit() else (cod_g if len(cod_g) >= 10 else "1050")
                ws[f"C{row}"] = str(int(cod_g)) if len(cod_g) >= 16 and cod_g.isdigit() else (cod_g if len(cod_g) >= 16 else cod_g)
                ws[f"D{row}"] = str(int(cod_g)) if len(cod_g) >= 18 and cod_g.isdigit() else "0"
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

        ws.merge_cells("A28:G28")
        ws["A28"] = "Productos no registrados"
        ws["A28"].font = Font(name="Arial", size=7.5, italic=True)
        ws["A28"].alignment = align_center
        ws["A28"].border = border_cell
        ws.row_dimensions[28].height = 13

        # Cuadros de firmas y sellos
        ws.merge_cells("A29:B32")
        ws["A29"] = "Poner sello aquí"
        ws["A29"].font = Font(name="Arial", size=7, italic=True)
        ws["A29"].alignment = align_center

        ws.merge_cells("A33:B33")
        ws["A33"] = nombre_jefatura
        ws["A33"].font = Font(name="Arial", size=8, bold=True)
        ws["A33"].alignment = align_center

        ws.merge_cells("A34:B35")
        ws["A34"] = "Nombre y firma jefe departamento"
        ws["A34"].font = font_small_bold
        ws["A34"].alignment = Alignment(horizontal="center", vertical="top")

        ws.merge_cells("C29:D33")
        ws["C29"] = ""
        ws.merge_cells("C34:D35")
        ws["C34"] = "Nombre de quién recibe conforme"
        ws["C34"].font = font_small_bold
        ws["C34"].alignment = Alignment(horizontal="center", vertical="top")

        ws.merge_cells("E29:E32")
        ws["E29"] = "Sello y Fecha"
        ws["E29"].font = Font(name="Arial", size=7, italic=True)
        ws["E29"].alignment = align_center
        ws.merge_cells("E34:E35")
        ws["E34"] = "Encargado de entrega"
        ws["E34"].font = font_small_bold
        ws["E34"].alignment = Alignment(horizontal="center", vertical="top")

        ws.merge_cells("F29:G32")
        ws["F29"] = ""
        ws.merge_cells("F33:G33")
        ws["F33"] = "Nombre del oficial de seguridad"
        ws["F33"].font = Font(name="Arial", size=6.5, bold=True)
        ws["F33"].alignment = align_center
        ws.merge_cells("F34:G35")
        ws["F34"] = "Oficial de seguridad (En caso de regional)"
        ws["F34"].font = Font(name="Arial", size=6.5, bold=True)
        ws["F34"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

        for r in range(29, 36):
            for col in ["A", "B", "C", "D", "E", "F", "G"]:
                ws[f"{col}{r}"].border = border_cell

        ws.merge_cells("A36:D36")
        ws["A36"] = "Original: Control de bodega. Ver 2.01.2021"
        ws["A36"].font = font_small
        ws["A36"].alignment = align_left

        ws.merge_cells("E36:G36")
        ws["E36"] = "COPIA: departamento solicitante"
        ws["E36"].font = font_small
        ws["E36"].alignment = align_right

    # Configuración estricta de página Carta (1 sola página exacta sin desbordes)
    try:
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.oddHeader.left.text = ""
        ws.oddHeader.center.text = ""
        ws.oddHeader.right.text = ""
        ws.oddFooter.left.text = ""
        ws.oddFooter.center.text = ""
        ws.oddFooter.right.text = ""
    except Exception:
        pass

    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    return consecutivo, nombre_archivo_salida
