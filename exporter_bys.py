import openpyxl
import sqlite3
from datetime import datetime
import os
import shutil

PLANTILLA_XLK = "Pedimentos de Departamento de Servicios Generales.xlk"
PLANTILLA_XLSX = "Pedimentos_de_Departamento_de_Servicios_Generales.xlsx"
EXCEL_MAESTRO = "Control_Maestro_238_Articulos_DINADECO.xlsx"

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
        cursor.execute("INSERT INTO control_consecutivos (anio, ultimo_numero) VALUES (?, ?)", (anio_actual, 51))
        
    conn.commit()
    conn.close()
    
    return f"{siguiente_num:03d}-{anio_actual}"


def cargar_mapa_master():
    mapa = {}
    if os.path.exists(EXCEL_MAESTRO):
        try:
            wb = openpyxl.load_workbook(EXCEL_MAESTRO, data_only=True)
            if 'Master_238_Articulos' in wb.sheetnames:
                ws = wb['Master_238_Articulos']
                for r in range(2, ws.max_row + 1):
                    grupal = str(ws.cell(r, 1).value or "").strip()
                    if grupal:
                        mapa[grupal] = {
                            "partida": str(ws.cell(r, 3).value or "29905").strip(),
                            "subpartida": str(ws.cell(r, 4).value or "01900").strip(),
                            "codigo": str(ws.cell(r, 5).value or "000000").strip(),
                            "subcodigo": str(ws.cell(r, 6).value or "00").strip(),
                        }
        except Exception:
            pass
    return mapa


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
    
    # 1. Asignar fecha sin hora para no activar la alerta de error en celda A5
    ws['N7'] = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    # 2. Asignar consecutivo y departamento oficial
    ws['N8'] = consecutivo
    ws['D8'] = "Departamento de Servicios Generales B"
    
    # 3. LIMPIAR RIGUROSAMENTE FILAS 12 A 37 USANDO CADENA VACÍA ""
    # (Esto vacía completamente las celdas de prueba sin tocar la fórmula de la columna E)
    for r in range(12, 38):
        ws.cell(row=r, column=1, value="")
        ws.cell(row=r, column=2, value="")
        ws.cell(row=r, column=3, value="")
        ws.cell(row=r, column=4, value="")
        ws.cell(row=r, column=14, value="")
        ws.cell(row=r, column=15, value="")
    
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    
    # Automigración de estructura si la tabla es antigua
    cursor.execute("PRAGMA table_info(articulos)")
    cols_existentes = [row[1] for row in cursor.fetchall()]
    for col in ['partida', 'subpartida', 'codigo', 'subcodigo']:
        if col not in cols_existentes:
            try:
                cursor.execute(f"ALTER TABLE articulos ADD COLUMN {col} TEXT")
            except Exception:
                pass
    conn.commit()

    mapa_master = cargar_mapa_master()
    
    fila_inicio = 12
    
    # 4. Escribir únicamente los artículos solicitados
    for idx, sol_id in enumerate(solicitud_ids):
        if fila_inicio + idx > 37:
            break
            
        cursor.execute('''
            SELECT a.codigo_grupal, a.descripcion, s.cantidad, a.partida, a.subpartida, a.codigo, a.subcodigo
            FROM solicitudes s
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            WHERE s.id = ?
        ''', (sol_id,))
        
        item = cursor.fetchone()
        if item:
            cod_grupal, descripcion, cantidad, partida, subpartida, codigo, subcodigo = item
            
            # Obtener códigos de presupuesto desde el catálogo Maestro o fallback
            if not (partida and subpartida and codigo):
                if cod_grupal in mapa_master:
                    partida = mapa_master[cod_grupal]['partida']
                    subpartida = mapa_master[cod_grupal]['subpartida']
                    codigo = mapa_master[cod_grupal]['codigo']
                    subcodigo = mapa_master[cod_grupal]['subcodigo']
                else:
                    partida = cod_grupal[:5] if len(cod_grupal) >= 5 else "29905"
                    subpartida = "01040" if "040" in cod_grupal else "01900"
                    codigo = cod_grupal[10:16] if len(cod_grupal) >= 16 else "000000"
                    subcodigo = "00"

            fila_actual = fila_inicio + idx
            ws.cell(row=fila_actual, column=1, value=str(partida))
            ws.cell(row=fila_actual, column=2, value=str(subpartida))
            ws.cell(row=fila_actual, column=3, value=str(codigo))
            ws.cell(row=fila_actual, column=4, value=str(subcodigo))
            ws.cell(row=fila_actual, column=14, value=int(cantidad))
            ws.cell(row=fila_actual, column=15, value=int(cantidad))

    conn.close()

    os.makedirs("pedimentos_generados", exist_ok=True)
    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    
    return consecutivo, nombre_archivo_salida
