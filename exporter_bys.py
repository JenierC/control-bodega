import openpyxl
import sqlite3
from datetime import datetime
import os
import shutil

PLANTILLA_XLK = "Pedimentos de Departamento de Servicios Generales.xlk"
PLANTILLA_XLSX = "Pedimentos_de_Departamento_de_Servicios_Generales.xlsx"

def obtener_siguiente_consecutivo(anio_actual=None):
    if not anio_actual:
        anio_actual = datetime.now().year
        
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS control_consecutivos (
            anio INTEGER PRIMARY KEY,
            ultimo_numero INTEGER NOT NULL
        )
    ''')
    
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


def generar_formulario_oficial_bys(solicitud_ids, nombre_departamento, nombre_jefatura, hash_firma):
    if os.path.exists(PLANTILLA_XLSX):
        plantilla = PLANTILLA_XLSX
    elif os.path.exists(PLANTILLA_XLK):
        shutil.copy(PLANTILLA_XLK, PLANTILLA_XLSX)
        plantilla = PLANTILLA_XLSX
    elif os.path.exists("Pedimentos de Departamento de Servicios Generales.xlsx"):
        plantilla = "Pedimentos de Departamento de Servicios Generales.xlsx"
    else:
        raise FileNotFoundError(f"No se encontró la plantilla oficial en formato Excel.")

    wb = openpyxl.load_workbook(plantilla)
    ws = wb['Pedimento']
    
    consecutivo = obtener_siguiente_consecutivo()
    fecha_actual = datetime.now()
    
    ws['N8'] = consecutivo
    ws['N7'] = fecha_actual.strftime("%d/%m/%Y")
    ws['D8'] = nombre_departamento
    
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    
    fila_inicio = 12
    for idx, sol_id in enumerate(solicitud_ids):
        if fila_inicio + idx > 37:
            break
            
        cursor.execute('''
            SELECT a.codigo_grupal, a.descripcion, s.cantidad 
            FROM solicitudes s
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            WHERE s.id = ?
        ''', (sol_id,))
        
        item = cursor.fetchone()
        if item:
            codigo_grupal, descripcion, cantidad = item
            
            partida = codigo_grupal[:5] if len(codigo_grupal) >= 5 else codigo_grupal
            sub_partida = codigo_grupal[5:10] if len(codigo_grupal) >= 10 else ""
            codigo = codigo_grupal[10:16] if len(codigo_grupal) >= 16 else ""
            sub_codigo = codigo_grupal[16:] if len(codigo_grupal) > 16 else "00"
            
            fila_actual = fila_inicio + idx
            ws.cell(row=fila_actual, column=1, value=partida)
            ws.cell(row=fila_actual, column=2, value=sub_partida)
            ws.cell(row=fila_actual, column=3, value=codigo)
            ws.cell(row=fila_actual, column=4, value=sub_codigo)
            ws.cell(row=fila_actual, column=5, value=descripcion)
            ws.cell(row=fila_actual, column=14, value=cantidad)
            ws.cell(row=fila_actual, column=15, value=cantidad)
    
    conn.close()
    
    ws['A46'] = f"Aprobado digitalmente por: {nombre_jefatura}"
    ws['A47'] = f"Hash de Firma: {hash_firma[:16]}... | Fecha: {fecha_actual.strftime('%Y-%m-%d %H:%M')}"
    
    os.makedirs("pedimentos_generados", exist_ok=True)
    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    
    return consecutivo, nombre_archivo_salida
