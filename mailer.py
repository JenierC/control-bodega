import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

def obtener_config_smtp():
    try:
        import streamlit as st
        if "smtp" in st.secrets:
            return st.secrets["smtp"]
    except Exception:
        pass

    try:
        from config import SMTP_CONFIG
        return SMTP_CONFIG
    except Exception:
        return None


def enviar_correo_con_pdf_adjunto(consecutivo, lista_items, solicitante, es_valido, motivo, ruta_pdf, justificacion=""):
    cfg = obtener_config_smtp()
    if not cfg:
        return False, "No se encontró la configuración SMTP."

    user = cfg.get("user", "")
    password = cfg.get("password", "")
    server_host = cfg.get("server", "smtp.gmail.com")
    port = int(cfg.get("port", 587))
    dest_principal = cfg.get("destinatario_principal", "agonzalez@dinadeco.go.cr")
    dest_cc = cfg.get("destinatarios_cc", ["ggutierrez@dinadeco.go.cr", "jcamacho@dinadeco.go.cr"])

    if not password or password == "TU_CONTRASEÑA_AQUI":
        return False, "No se ha configurado la contraseña de correo en Secrets."

    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = dest_principal
    msg['Cc'] = ", ".join(dest_cc) if isinstance(dest_cc, list) else str(dest_cc)
    
    tag = "[CUMPLE REGLAS]" if es_valido else "[ALERTA / EXCEDE CUOTA]"
    msg['Subject'] = f"{tag} Solicitud de Pedimento #{consecutivo} - Para Firma Digital"

    items_html = "".join([f"<li><b>{it['descripcion']}</b> - Cantidad: {it['cantidad']}</li>" for it in lista_items])

    cuerpo = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; border: 1px solid #ddd;">
          <h2 style="color: #003366;">Solicitud de Pedimento #{consecutivo}</h2>
          <p><b>Solicitante:</b> {solicitante}</p>
          <p><b>Dictamen:</b> {motivo}</p>
          {f'<p><b>Justificación:</b> {justificacion}</p>' if justificacion else ''}
          
          <h3>Lista de Artículos Solicitados:</h3>
          <ul>{items_html}</ul>
          
          <hr>
          <p><b>Instrucciones para la Autorización / Firma Digital:</b></p>
          <ol>
            <li>Descargue el archivo PDF adjunto (<code>Pedimento_{consecutivo}.pdf</code>).</li>
            <li>Ábralo en <b>Adobe Acrobat Reader</b>.</li>
            <li>Inserte su tarjeta de <b>Firma Digital de Costa Rica</b>.</li>
            <li>Estampe la firma digital en la casilla correspondiente y guarde el documento.</li>
            <li>Cargue el PDF firmado al portal web para completar la salida de bodega.</li>
          </ol>
        </div>
      </body>
    </html>
    """
    
    msg.attach(MIMEText(cuerpo, 'html'))

    if os.path.exists(ruta_pdf):
        with open(ruta_pdf, "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="pdf")
            adjunto.add_header('Content-Disposition', 'attachment', filename=os.path.basename(ruta_pdf))
            msg.attach(adjunto)

    todos_destinatarios = [dest_principal] + (dest_cc if isinstance(dest_cc, list) else [dest_cc])

    try:
        server = smtplib.SMTP(server_host, port)
        server.starttls()
        server.login(user, password)
        server.sendmail(user, todos_destinatarios, msg.as_string())
        server.quit()
        return True, "Correo enviado exitosamente."
    except Exception as e:
        return False, f"Fallo al conectar con el servidor de correo ({type(e).__name__}): {str(e)}"


def enviar_correo_pedimento_firmado(consecutivo, ruta_pdf_firmado, detalle_firma="Firma Digital Verificada"):
    """
    Envía el PDF con Firma Digital de BCCR al Encargado de Bodega (agomez@dinadeco.go.cr)
    con copia a jcamacho@dinadeco.go.cr y agonzalez@dinadeco.go.cr para autorizar despacho.
    """
    cfg = obtener_config_smtp()
    if not cfg:
        return False, "No se encontró la configuración SMTP."

    user = cfg.get("user", "")
    password = cfg.get("password", "")
    server_host = cfg.get("server", "smtp.gmail.com")
    port = int(cfg.get("port", 587))
    
    # Destinatarios oficiales para la salida de bodega
    dest_principal = "agomez@dinadeco.go.cr"
    dest_cc = ["jcamacho@dinadeco.go.cr", "agonzalez@dinadeco.go.cr"]

    if not password or password == "TU_CONTRASEÑA_AQUI":
        return False, "No se ha configurado la contraseña de correo en Secrets."

    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = dest_principal
    msg['Cc'] = ", ".join(dest_cc)
    msg['Subject'] = f"✅ [PEDIMENTO FIRMADO Y AUTORIZADO] Consecutivo #{consecutivo} - Salida de Bodega"

    cuerpo = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <div style="background-color: #f4f6f9; padding: 20px; border-radius: 8px; border: 1px solid #003366;">
          <h2 style="color: #003366;">✅ Salida de Bodega Autorizada - Pedimento #{consecutivo}</h2>
          <p>Estimado(a) Encargado(a) de Bodega (<b>agomez@dinadeco.go.cr</b>),</p>
          <p>Se ha verificado la <b>Firma Digital</b> del documento y se autoriza oficialmente la salida de los materiales correspondientes al pedimento <b>#{consecutivo}</b>.</p>

          <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
            <tr>
              <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Consecutivo:</b></td>
              <td style="padding: 8px; border-bottom: 1px solid #ddd;">#{consecutivo}</td>
            </tr>
            <tr>
              <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Estado:</b></td>
              <td style="padding: 8px; border-bottom: 1px solid #ddd; color: green; font-weight: bold;">AUTORIZADO PARA ENTREGA / DESPACHO</td>
            </tr>
            <tr>
              <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Validación de Firma:</b></td>
              <td style="padding: 8px; border-bottom: 1px solid #ddd;">{detalle_firma}</td>
            </tr>
          </table>

          <hr style="margin-top: 20px;">
          <p>Se adjunta el Pedimento oficial en formato PDF debidamente firmado para proceder con el despacho físico y archivo de control.</p>
        </div>
      </body>
    </html>
    """
    
    msg.attach(MIMEText(cuerpo, 'html'))

    if os.path.exists(ruta_pdf_firmado):
        with open(ruta_pdf_firmado, "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="pdf")
            adjunto.add_header('Content-Disposition', 'attachment', filename=f"Pedimento_{consecutivo}_Firmado.pdf")
            msg.attach(adjunto)

    todos_destinatarios = [dest_principal] + dest_cc

    try:
        server = smtplib.SMTP(server_host, port)
        server.starttls()
        server.login(user, password)
        server.sendmail(user, todos_destinatarios, msg.as_string())
        server.quit()
        return True, "Correo enviado al encargado de bodega exitosamente."
    except Exception as e:
        return False, f"Fallo al enviar correo a bodega ({type(e).__name__}): {str(e)}"
