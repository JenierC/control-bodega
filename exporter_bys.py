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
        raise FileNotFoundError("No se encontró la plantilla oficial en formato Excel.")

    wb = openpyxl.load_workbook(plantilla)
    ws = wb['Pedimento']
    
    consecutivo = obtener_siguiente_consecutivo()
    
    # 1. Asignar fecha en formato datetime exacto
    ws['N7'] = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    # 2. Consecutivo y departamento
    ws['N8'] = consecutivo
    ws['D8'] = "Departamento de Servicios Generales B"
    
    # 3. LIMPIAR COMPLETAMENTE LAS FILAS 12 A 37 (Elimina artículos viejos de prueba)
    for r in range(12, 38):
        ws.cell(row=r, column=1, value=None)
        ws.cell(row=r, column=2, value=None)
        ws.cell(row=r, column=3, value=None)
        ws.cell(row=r, column=4, value=None)
        ws.cell(row=r, column=5, value=None)
        ws.cell(row=r, column=14, value=None)
        ws.cell(row=r, column=15, value=None)
    
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    
    # 4. Escribir únicamente los artículos solicitados
    fila_inicio = 12
    total_items = len(solicitud_ids)
    
    for idx, sol_id in enumerate(solicitud_ids):
        if fila_inicio + idx > 37:
            break
            
        cursor.execute('''
            SELECT a.partida, a.subpartida, a.codigo, a.subcodigo, a.descripcion, s.cantidad 
            FROM solicitudes s
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            WHERE s.id = ?
        ''', (sol_id,))
        
        item = cursor.fetchone()
        if item:
            partida, subpartida, codigo, subcodigo, descripcion, cantidad = item
            
            fila_actual = fila_inicio + idx
            ws.cell(row=fila_actual, column=1, value=str(partida or "29905"))
            ws.cell(row=fila_actual, column=2, value=str(subpartida or "01040"))
            ws.cell(row=fila_actual, column=3, value=str(codigo or "000060"))
            ws.cell(row=fila_actual, column=4, value=str(subcodigo or "00"))
            ws.cell(row=fila_actual, column=5, value=str(descripcion))
            ws.cell(row=fila_actual, column=14, value=int(cantidad))
            ws.cell(row=fila_actual, column=15, value=int(cantidad))
    
    # Marcar marca de fin de lista UL (Última Línea) inmediatamente después del último ítem
    fila_ul = fila_inicio + total_items
    if fila_ul <= 37:
        ws.cell(row=fila_ul, column=5, value="***** UL *****")

    conn.close()

    os.makedirs("pedimentos_generados", exist_ok=True)
    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    
    return consecutivo, nombre_archivo_salida
