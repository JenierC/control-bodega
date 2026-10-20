import openpyxl
import sqlite3
from datetime import datetime
import os
import shutil
import subprocess

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


def generar_formulario_oficial_bys(solicitud_ids, nombre_departamento, nombre_jefatura, hash_firma, consecutivo_forzado=None):
    # 1. Selección segura de la plantilla original de DINADECO
    # Primero buscamos archivos nativos .xlsx o .xlsm para evitar el error de extensión .xlk en openpyxl
    plantilla = None
    candidatos_nativos = [
        "Pedimentos de Departamento de Servicios Generales.xlsx",
        "Pedimentos de Departamento de Servicios Generales.xlsm",
        "Pedimentos de Departamento de Servicios Generales B.xlsx",
        "Pedimentos de Departamento de Servicios Generales B.xlsm"
    ]
    for c in candidatos_nativos:
        if os.path.exists(c):
            try:
                openpyxl.load_workbook(c)
                plantilla = c
                break
            except Exception:
                pass

    if not plantilla:
        for f in os.listdir("."):
            if (f.endswith(".xlsx") or f.endswith(".xlsm")) and ("pedimento" in f.lower() or "servicios" in f.lower()):
                try:
                    openpyxl.load_workbook(f)
                    plantilla = f
                    break
                except Exception:
                    pass

    # Si solo existe el archivo con extensión .xlk (backup de Excel), lo convertimos a .xlsx
    if not plantilla:
        for f in os.listdir("."):
            if f.endswith(".xlk"):
                temp_xlsx = f.replace(".xlk", "_plantilla.xlsx")
                shutil.copyfile(f, temp_xlsx)
                try:
                    openpyxl.load_workbook(temp_xlsx)
                    plantilla = temp_xlsx
                    break
                except Exception:
                    soffice = shutil.which("libreoffice") or shutil.which("soffice")
                    if soffice:
                        try:
                            subprocess.run([soffice, "--headless", "--convert-to", "xlsx", f], check=True, timeout=15)
                            conv = f.replace(".xlk", ".xlsx")
                            if os.path.exists(conv):
                                plantilla = conv
                                break
                        except Exception:
                            pass

    if not plantilla:
        if os.path.exists(PLANTILLA_BYS):
            plantilla = PLANTILLA_BYS
        else:
            raise FileNotFoundError("No se encuentra la plantilla oficial 'Pedimentos de Departamento de Servicios Generales'")

    wb = openpyxl.load_workbook(plantilla)
    if 'Pedimento' in wb.sheetnames:
        ws = wb['Pedimento']
    else:
        ws = wb.active
    
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
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
    
    # Asegurar que se imprima/exporte en 1 sola página exacta sin cortar
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
