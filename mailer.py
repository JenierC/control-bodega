import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
from config import SMTP_CONFIG

def enviar_correo_con_pdf_adjunto(consecutivo, lista_items, solicitante, es_valido, motivo, ruta_pdf, justificacion=""):
    msg = MIMEMultipart()
    msg['From'] = SMTP_CONFIG["user"]
    msg['To'] = ", ".join(SMTP_CONFIG["destinatarios_notificacion"])
    
    tag = "[CUMPLE REGLAS]" if es_valido else "[ALERTA / EXCEDE CUOTA]"
    msg['Subject'] = f"{tag} Solicitud de Pedimento #{consecutivo} - Para Firma Digital"

    items_html = "".join([f"<li><b>{it['descripcion']}</b> - Cantidad: {it['cantidad']}</li>" for it in lista_items])

    cuerpo = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px;">
          <h2 style="color: #003366;">Solicitud de Pedimento #{consecutivo}</h2>
          <p><b>Solicitante:</b> {solicitante}</p>
          <p><b>Dictamen de Validación:</b> {motivo}</p>
          {f'<p><b>Justificación:</b> {justificacion}</p>' if justificacion else ''}
          
          <h3>Lista de Artículos (Hasta 10 ítems):</h3>
          <ul>{items_html}</ul>
          
          <hr>
          <p><b>Instrucciones para la Jefatura (Jenier / Gabriela):</b></p>
          <ol>
            <li>Descargue el archivo PDF adjunto (<code>Pedimento_{consecutivo}.pdf</code>).</li>
            <li>Ábralo en <b>Adobe Acrobat Reader</b>.</li>
            <li>Inserte su tarjeta de <b>Firma Digital de Costa Rica</b>.</li>
            <li>Estampe la firma digital en la sección de jefatura y guarde el documento.</li>
            <li>Cargue el PDF firmado al portal en la Etapa 2 para completar el registro.</li>
          </ol>
        </div>
      </body>
    </html>
    """
    
    msg.attach(MIMEText(cuerpo, 'html'))

    # Adjuntar PDF generado para firmar en Adobe
    if os.path.exists(ruta_pdf):
        with open(ruta_pdf, "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="pdf")
            adjunto.add_header('Content-Disposition', 'attachment', filename=os.path.basename(ruta_pdf))
            msg.attach(adjunto)

    try:
        server = smtplib.SMTP(SMTP_CONFIG["server"], SMTP_CONFIG["port"])
        server.starttls()
        server.login(SMTP_CONFIG["user"], SMTP_CONFIG["password"])
        server.sendmail(SMTP_CONFIG["user"], SMTP_CONFIG["destinatarios_notificacion"], msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando correo: {e}")
        return False
