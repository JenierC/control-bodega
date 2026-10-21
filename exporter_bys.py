import openpyxl
import sqlite3
from datetime import datetime
import os
import shutil

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
        siguiente_num = 51
        cursor.execute("INSERT INTO control_consecutivos (anio, ultimo_numero) VALUES (?, ?)", (anio_actual, siguiente_num))
        
    conn.commit()
    conn.close()
    return f"{siguiente_num:03d}-{anio_actual}"


def cargar_plantilla_segura():
    # 1. Buscar archivos con extensión nativa compatible (.xlsx / .xlsm)
    candidatos = [
        "Pedimentos de Departamento de Servicios Generales.xlsx",
        "Pedimentos de Departamento de Servicios Generales.xlsm",
        "Pedimentos de Departamento de Servicios Generales B.xlsx",
        "Pedimentos de Departamento de Servicios Generales B.xlsm"
    ]
    for c in candidatos:
        if os.path.exists(c):
            try:
                return openpyxl.load_workbook(c)
            except Exception:
                pass

    for f in os.listdir("."):
        if (f.lower().endswith(".xlsx") or f.lower().endswith(".xlsm")) and not f.startswith("~"):
            if "pedimento" in f.lower() or "servicios" in f.lower():
                try:
                    return openpyxl.load_workbook(f)
                except Exception:
                    pass

    # 2. Si solo existe el archivo .xlk, se copia a .xlsx para evitar InvalidFileException
    for f in os.listdir("."):
        if f.lower().endswith(".xlk"):
            temp_xlsx = "plantilla_desde_xlk.xlsx"
            shutil.copyfile(f, temp_xlsx)
            try:
                return openpyxl.load_workbook(temp_xlsx)
            except Exception:
                pass

    raise FileNotFoundError("No se encontró la plantilla oficial de pedimentos en formato Excel (.xlsx, .xlsm o .xlk)")


def generar_formulario_oficial_bys(solicitud_ids, nombre_departamento="Departamento de Servicios Generales B", nombre_jefatura="Jenier Camacho Picado", hash_firma="", consecutivo_forzado=None):
    os.makedirs("pedimentos_generados", exist_ok=True)
    
    wb = cargar_plantilla_segura()
    ws = wb['Pedimento'] if 'Pedimento' in wb.sheetnames else wb.active
    wb.active = ws  # Asegura que LibreOffice convierta la hoja correcta
    
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
        consecutivo = obtener_siguiente_consecutivo()
        
    fecha_actual = datetime.now()
    
    # Encabezado oficial en las celdas exactas de la plantilla
    ws['N8'] = consecutivo
    ws['N7'] = fecha_actual.strftime("%d/%m/%Y")
    ws['D8'] = nombre_departamento
    
    # Cargar artículos de la base de datos
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
    
    fila_inicio = 12
    for idx in range(26):  # Renglones oficiales 12 a 37
        fila_actual = fila_inicio + idx
        if fila_actual > 37:
            break
            
        if idx < len(items):
            _, codigo_grupal, descripcion, cantidad, _, _ = items[idx]
            codigo_grupal = str(codigo_grupal).strip()
            
            partida = codigo_grupal[:5] if len(codigo_grupal) >= 5 else codigo_grupal
            sub_partida = codigo_grupal[5:10] if len(codigo_grupal) >= 10 else ""
            codigo = codigo_grupal[10:16] if len(codigo_grupal) >= 16 else ""
            sub_codigo = codigo_grupal[16:] if len(codigo_grupal) > 16 else "0"
            
            ws.cell(row=fila_actual, column=1, value=partida)
            ws.cell(row=fila_actual, column=2, value=sub_partida)
            ws.cell(row=fila_actual, column=3, value=codigo)
            ws.cell(row=fila_actual, column=4, value=sub_codigo)
            ws.cell(row=fila_actual, column=5, value=descripcion)
            ws.cell(row=fila_actual, column=14, value=int(cantidad))
            ws.cell(row=fila_actual, column=15, value="")
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

    # ===============================================================
    # ACTUALIZACIÓN EXACTA: JENIER CAMACHO PICADO EN LA JEFATURA
    # ===============================================================
    nombre_jefe_titular = "Jenier Camacho Picado"
    
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    txt = str(cell.value).strip()
                    # 1. Reemplazar cualquier variante del nombre antiguo
                    if any(k in txt.lower() for k in ["edgar", "vargas", "brenes"]):
                        cell.value = nombre_jefe_titular
                    # 2. Asegurar que la celda inmediatamente sobre el cargo tenga a Jenier Camacho Picado
                    if "nombre y firma" in txt.lower() and "jefe" in txt.lower():
                        r_jefe = cell.row - 1
                        sheet.cell(row=r_jefe, column=cell.column, value=nombre_jefe_titular)

    # ===============================================================
    # CONFIGURACIÓN DE PÁGINA (CONSERVA EL PIE DE PÁGINA COMPLETO)
    # ===============================================================
    try:
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        
        ws.page_margins.top = 0.35
        ws.page_margins.bottom = 0.85
        ws.page_margins.left = 0.4
        ws.page_margins.right = 0.4
        ws.page_margins.footer = 0.45
        ws.page_margins.header = 0.2
    except Exception:
        pass
    
    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    return consecutivo, nombre_archivo_salida
