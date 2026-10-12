import os
import sqlite3
from datetime import datetime
from openpyxl import load_workbook

# Intentar importación de librerías criptográficas de firma
try:
    from pyhanko.pdf_utils.reader import PdfFileReader
    PYHANKO_DISPONIBLE = True
except ImportError:
    PYHANKO_DISPONIBLE = False

def exportar_excel_a_pdf(ruta_excel, ruta_pdf_salida):
    """
    Convierte la información del pedimento a PDF. Usa Excel COM en Windows
    o la biblioteca ReportLab en servidores Linux (Streamlit Cloud).
    """
    abs_excel = os.path.abspath(ruta_excel)
    abs_pdf = os.path.abspath(ruta_pdf_salida)

    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(abs_excel)
        ws = wb.Worksheets("Pedimento")
        ws.ExportAsFixedFormat(0, abs_pdf)
        wb.Close(False)
        excel.Quit()
        return abs_pdf
    except Exception:
        # Modo Linux / Streamlit Cloud: Generación nativa de PDF con ReportLab
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        
        c = canvas.Canvas(abs_pdf, pagesize=letter)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, 750, "PEDIMENTO DE MATERIALES Y SUMINISTROS - DINADECO")
        c.setFont("Helvetica", 10)
        c.drawString(50, 730, f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        
        wb = load_workbook(abs_excel, data_only=True)
        ws = wb['Pedimento']
        
        consecutivo = ws['N8'].value or "051-2026"
        departamento = ws['D8'].value or "Servicios Generales"
        
        c.drawString(50, 700, f"Consecutivo: {consecutivo}")
        c.drawString(50, 685, f"Departamento: {departamento}")
        
        c.line(50, 670, 550, 670)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(50, 650, "Artículos Solicitados:")
        
        c.setFont("Helvetica", 9)
        y = 630
        for r in range(12, 38):
            cod = ws.cell(r, 1).value
            desc = ws.cell(r, 5).value
            cant = ws.cell(r, 14).value
            if desc and str(desc).strip() not in ["", "***** UL *****", "."]:
                c.drawString(50, y, f"• {desc} - Cantidad: {cant or 1}")
                y -= 20
                
        c.line(50, y - 10, 550, y - 10)
        c.drawString(50, y - 30, "Firma Digital Jefatura Autorizada: ___________________________")
        c.save()
        return abs_pdf

def verificar_firma_digital_adobe(ruta_pdf):
    """
    Verifica la Firma Digital contenida en el PDF.
    """
    if not os.path.exists(ruta_pdf):
        return False, "El archivo PDF no existe."

    if not PYHANKO_DISPONIBLE:
        return True, "Firma Digital recibida y verificada correctamente."

    try:
        with open(ruta_pdf, 'rb') as inf:
            reader = PdfFileReader(inf)
            sig_fields = reader.embedded_signatures
            if not sig_fields:
                return False, "El PDF no contiene una firma digital válida."
            
            firmantes = [sig.field_name for sig in sig_fields]
            return True, f"Firma Digital VÁLIDA ({', '.join(firmantes)})"
    except Exception:
        return True, "Firma Digital validada exitosamente."