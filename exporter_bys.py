import openpyxl
import sqlite3
from datetime import datetime
import os

PLANTILLA_BYS = "Pedimentos de Departamento de Servicios Generales.xlk"

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


def generar_formulario_oficial_bys(solicitud_ids, nombre_departamento, nombre_jefatura, hash_firma, consecutivo_forzado=None):
    # Detección de la plantilla institucional oficial en el repositorio
    posibles_plantillas = [
        "Pedimentos de Departamento de Servicios Generales.xlk",
        "Pedimentos de Departamento de Servicios Generales.xlsx",
        "Pedimentos de Departamento de Servicios Generales.xlsm",
        "Pedimentos de Departamento de Servicios Generales B.xlsm",
        "Pedimentos de Departamento de Servicios Generales B.xlsx"
    ]
    
    plantilla = None
    for p in posibles_plantillas:
        if os.path.exists(p):
            plantilla = p
            break
            
    if not plantilla:
        for f in os.listdir("."):
            if f.startswith("Pedimentos de Departamento") and any(f.endswith(ext) for ext in [".xlk", ".xlsx", ".xlsm"]):
                plantilla = f
                break
                
    if not plantilla:
        raise FileNotFoundError(f"No se encuentra la plantilla oficial: {PLANTILLA_BYS}")

    wb = openpyxl.load_workbook(plantilla)
    
    if 'Pedimento' in wb.sheetnames:
        ws = wb['Pedimento']
    else:
        ws = wb.active
    
    # Si se pide regenerar un pedimento viejo, mantiene su consecutivo; si es nuevo, incrementa
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
        consecutivo = obtener_siguiente_consecutivo()
        
    fecha_actual = datetime.now()
    
    # Celdas exactas del formato institucional original
    ws['N8'] = consecutivo
    ws['N7'] = fecha_actual.strftime("%d/%m/%Y")
    ws['D8'] = nombre_departamento
    
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    
    fila_inicio = 12
    for idx, sol_id in enumerate(solicitud_ids):
        if fila_inicio + idx > 37:
            break
            
        cursor.execute("""
            SELECT a.codigo_grupal, a.descripcion, s.cantidad 
            FROM solicitudes s
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            WHERE s.id = ?
        """, (sol_id,))
        
        item = cursor.fetchone()
        if item:
            codigo_grupal, descripcion, cantidad = item
            
            partida = codigo_grupal[:5] if len(codigo_grupal) >= 5 else codigo_grupal
            sub_partida = codigo_grupal[5:10] if len(codigo_grupal) >= 10 else ""
            codigo = codigo_grupal[10:16] if len(codigo_grupal) >= 16 else ""
            sub_codigo = codigo_grupal[16:] if len(codigo_grupal) > 16 else "0"
            
            fila_actual = fila_inicio + idx
            ws.cell(row=fila_actual, column=1, value=partida)
            ws.cell(row=fila_actual, column=2, value=sub_partida)
            ws.cell(row=fila_actual, column=3, value=codigo)
            ws.cell(row=fila_actual, column=4, value=sub_codigo)
            ws.cell(row=fila_actual, column=14, value=cantidad)
            ws.cell(row=fila_actual, column=15, value=cantidad)
    
    conn.close()
    
    ws['A46'] = f"Aprobado digitalmente por: {nombre_jefatura}"
    ws['A47'] = f"Hash de Firma: {hash_firma[:16]}... | Fecha: {fecha_actual.strftime('%Y-%m-%d %H:%M')}"
    
    # Regla estricta para forzar a 1 sola página exacta sin desbordes ni hojas huérfanas
    try:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
    except Exception:
        pass
    
    os.makedirs("pedimentos_generados", exist_ok=True)
    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    
    return consecutivo, nombre_archivo_salida
