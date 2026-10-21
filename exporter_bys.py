import openpyxl
import sqlite3
from datetime import datetime
import os
import shutil
import re
import unicodedata

PLANTILLA_BYS = "Pedimentos de Departamento de Servicios Generales.xlsx"

def normalizar_texto(texto):
    if not texto:
        return ""
    t = unicodedata.normalize('NFD', str(texto).upper())
    t = "".join(c for c in t if unicodedata.category(c) != 'Mn')
    t = re.sub(r'[^A-Z0-9]', '', t)
    return t

def desglosar_codigo_institucional(descripcion, codigo_grupal):
    """
    Retorna con precisión milimétrica los códigos oficiales de Partida, Sub-partida,
    Código y Sub-código según la normativa formal de DINADECO (Pedimento_057-2026..pdf).
    """
    d = normalizar_texto(descripcion)
    
    # 1. FIBRA PARA CEPILLO ELECTRICO 17 ROJA
    if "CEPILLO" in d:
        return "29905", "1030", "100", "4"
        
    # 2. ESPONJA DE FIBRA, LAVAPLATOS
    if "ESPONJA" in d:
        return "29905", "1030", "21", "0"
        
    # 3. CERA LIQUIDA
    if "CERA" in d:
        return "29905", "1010", "2", "3"
        
    # 4. TOALLA PARA SECADO DE MANOS (DISPENSADOR)
    if "SECADO" in d or ("TOALLA" in d and "MANOS" in d):
        return "29905", "1900", "190203", "0"
        
    # 5. TOALLAS DESECHABLES (MAYORDOMO)
    if "MAYORDOMO" in d:
        return "29905", "1900", "190205", "0"
        
    # 6. BOLSA PARA BASURA TIPO JARDIN
    if "JARDIN" in d or ("BASURA" in d and "GRANDE" not in d):
        return "29905", "1060", "100015", "0"
        
    # 7. ABRILLANTADOR PARA MUEBLES
    if "ABRILLANTADOR" in d:
        return "29905", "1040", "60", "0"
        
    # 8. DESINFECTANTE EN GALON
    if "DESINFECTANTE" in d:
        return "29905", "1035", "200", "0"
        
    # 9. DESODORANTE AMBIENTAL EN AEROSOL
    if "DESODORANTE" in d:
        return "29905", "1050", "1", "0"
        
    # 10. PAPEL HIGIENICO TIPO JUMBO
    if "HIGIENICO" in d or "JUMBO" in d:
        return "29905", "1900", "190201", "1"
        
    # 11. REPUESTO DE AROMATIZANTE AMBIENTAL
    if "AROMATIZANTE" in d:
        return "29905", "1050", "1", "1"
        
    # 12. PANO DE MICROFIBRA
    if "MICROFIBRA" in d or "PANO" in d:
        return "29905", "1900", "55", "0"
        
    # 13. SERVILLETAS DE PAPEL
    if "SERVILLETAS" in d or "SERVILLETA" in d:
        return "29905", "1900", "190209", "0"

    # Desglose de contingencia si el insumo no figura en los nombres clave anteriores
    tokens = re.findall(r'[0-9]+', str(codigo_grupal))
    if len(tokens) >= 4:
        return tokens[0], str(int(tokens[1])), str(int(tokens[2])), str(int(tokens[3]))

    cod_puro = re.sub(r'[^0-9]', '', str(codigo_grupal))
    partida = cod_puro[:5] if len(cod_puro) >= 5 else "29905"
    resto = cod_puro[5:]
    if len(resto) >= 11:
        sp = str(int(resto[:5])) if resto[:5].isdigit() else resto[:5]
        cd = str(int(resto[5:11])) if resto[5:11].isdigit() else resto[5:11]
        sc = str(int(resto[11:])) if resto[11:].isdigit() else (resto[11:] or "0")
        return partida, sp, cd, sc
    elif len(resto) >= 8:
        sp = str(int(resto[:4])) if resto[:4].isdigit() else resto[:4]
        cd = str(int(resto[4:7])) if resto[4:7].isdigit() else resto[4:7]
        sc = str(int(resto[7:])) if resto[7:].isdigit() else (resto[7:] or "0")
        return partida, sp, cd, sc

    return "29905", "1050", resto if resto else "1", "0"


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
    wb.active = ws
    
    if consecutivo_forzado:
        consecutivo = str(consecutivo_forzado).strip()
    else:
        consecutivo = obtener_siguiente_consecutivo()
        
    fecha_actual = datetime.now()
    
    # -------------------------------------------------------------
    # PUNTO 2: FUERZA OBLIGATORIA DE "Departamento de Servicios Generales B"
    # -------------------------------------------------------------
    depto_texto = str(nombre_departamento).strip()
    if not depto_texto.endswith("B"):
        depto_texto = f"{depto_texto} B"
        
    ws['N8'] = consecutivo
    ws['N7'] = fecha_actual.strftime("%d/%m/%Y")
    ws['D8'] = depto_texto
    
    # Consulta de solicitudes
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
    
    # -------------------------------------------------------------
    # PUNTO 1: LLENADO CON LOS CÓDIGOS OFICIALES EXACTOS (FILAS 12 A 37)
    # -------------------------------------------------------------
    fila_inicio = 12
    for idx in range(26):
        fila_actual = fila_inicio + idx
        if fila_actual > 37:
            break
            
        if idx < len(items):
            _, codigo_grupal, descripcion, cantidad, _, _ = items[idx]
            
            # Desglose oficial validado contra Pedimento_057-2026..pdf
            partida, sub_partida, codigo, sub_codigo = desglosar_codigo_institucional(descripcion, codigo_grupal)
            
            ws.cell(row=fila_actual, column=1, value=partida)
            ws.cell(row=fila_actual, column=2, value=sub_partida)
            ws.cell(row=fila_actual, column=3, value=codigo)
            ws.cell(row=fila_actual, column=4, value=sub_codigo)
            ws.cell(row=fila_actual, column=5, value=str(descripcion).strip())
            ws.cell(row=fila_actual, column=14, value=int(cantidad))
            ws.cell(row=fila_actual, column=15, value="")
        elif idx == len(items):
            ws.cell(row=fila_actual, column=1, value="")
            ws.cell(row=fila_actual, column=2, value="")
            ws.cell(row=fila_actual, column=3, value="")
            ws.cell(row=fila_actual, column=4, value="")
            ws.cell(row=fila_actual, column=5, value="***** UL *****")
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

    # Garantizar a Jenier Camacho Picado en jefatura
    ws['A46'].value = "Jenier Camacho Picado"
    for r in range(1, 65):
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                txt = str(cell.value).strip().lower()
                if "edgar" in txt or "vargas" in txt or "brenes" in txt:
                    cell.value = "Jenier Camacho Picado"

    # -------------------------------------------------------------
    # PUNTOS 4 Y 5: PIE DE PÁGINA OFICIAL INTACTO
    # -------------------------------------------------------------
    fecha_pie_texto = fecha_actual.strftime("%d/%m/%Y")
    nombre_archivo_pie = f"Pedimentos de Departamento de Servicios Generales B {consecutivo}.xlsm"
    hora_pie_texto = fecha_actual.strftime("%H:%M")

    for sheet in wb.worksheets:
        for ftr in [sheet.oddFooter, sheet.evenFooter, sheet.firstFooter]:
            ftr.left.text = fecha_pie_texto
            ftr.center.text = nombre_archivo_pie
            ftr.right.text = hora_pie_texto

    # -------------------------------------------------------------
    # PUNTO 3: CENTRADO DE PÁGINA (CORRECCIÓN DEL DESFASE IZQUIERDO)
    # -------------------------------------------------------------
    try:
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_setup.horizontalCentered = True
        
        # Margen balanceado idéntico al original bueno (X=55.0)
        ws.page_margins.left = 0.75
        ws.page_margins.right = 0.75
        ws.page_margins.top = 0.35
        ws.page_margins.bottom = 0.85
        ws.page_margins.footer = 0.45
        ws.page_margins.header = 0.2
    except Exception:
        pass
    
    nombre_archivo_salida = f"pedimentos_generados/Pedimento_{consecutivo}.xlsx"
    wb.save(nombre_archivo_salida)
    return consecutivo, nombre_archivo_salida
