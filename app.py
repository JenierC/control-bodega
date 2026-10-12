import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import hashlib
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import os

from setup_inicial import instalar_sistema
from exporter_bys import generar_formulario_oficial_bys
from pdf_signature import exportar_excel_a_pdf, verificar_firma_digital_adobe
from mailer import enviar_correo_con_pdf_adjunto

EXCEL_MAESTRO = "Control_Maestro_238_Articulos_DINADECO.xlsx"

# ==========================================
# VERIFICACIÓN AUTOMÁTICA DE BASE DE DATOS
# ==========================================
def verificar_base_datos():
    if not os.path.exists("database.db"):
        instalar_sistema()
    else:
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='funcionarios'")
        tabla_existe = cursor.fetchone()
        conn.close()
        if not tabla_existe:
            instalar_sistema()

verificar_base_datos()

st.set_page_config(page_title="Bodega DINADECO", page_icon="📦", layout="wide")

if 'usuario_actual' not in st.session_state:
    st.session_state['usuario_actual'] = None
    st.session_state['rol'] = None
    st.session_state['func_id'] = None
if 'carrito' not in st.session_state:
    st.session_state['carrito'] = []
if 'ultimo_pedimento' not in st.session_state:
    st.session_state['ultimo_pedimento'] = None

def login(user, pwd):
    hash_pwd = hashlib.sha256(pwd.encode()).hexdigest()
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, rol, estado FROM funcionarios WHERE username=? AND password=?", (user, hash_pwd))
    result = cursor.fetchone()
    conn.close()
    
    if result:
        if result[3] == "INACTIVO":
            st.error("Su cuenta está pendiente de aprobación por Jenier o Gabriela.")
        else:
            st.session_state['func_id'] = result[0]
            st.session_state['usuario_actual'] = result[1]
            st.session_state['rol'] = result[2]
            st.rerun()
    else:
        st.error("Usuario o contraseña incorrectos.")

def logout():
    st.session_state['usuario_actual'] = None
    st.session_state['rol'] = None
    st.session_state['func_id'] = None
    st.session_state['carrito'] = []
    st.session_state['ultimo_pedimento'] = None
    st.rerun()

def sync_usuarios_excel():
    conn = sqlite3.connect("database.db")
    df_users = pd.read_sql_query("SELECT id, nombre, puesto, espacio_asignado, username, rol, estado FROM funcionarios", conn)
    conn.close()
    
    try:
        wb = openpyxl.load_workbook(EXCEL_MAESTRO)
        if 'USUARIOS' in wb.sheetnames:
            del wb['USUARIOS']
        ws = wb.create_sheet('USUARIOS')
        
        for r in dataframe_to_rows(df_users, index=False, header=True):
            ws.append(r)
        wb.save(EXCEL_MAESTRO)
    except Exception as e:
        st.error(f"Error sincronizando Excel Maestro: {e}")

# -------------------------------------------------------------------
# PANTALLA DE LOGIN Y REGISTRO
# -------------------------------------------------------------------
if st.session_state['usuario_actual'] is None:
    tab_login, tab_registro = st.tabs(["🔑 Iniciar Sesión", "📝 Solicitar Acceso (Nuevo Usuario)"])
    
    with tab_login:
        st.subheader("Acceso al Sistema de Bodega")
        with st.form("form_login"):
            usuario = st.text_input("Usuario")
            clave = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar"):
                login(usuario, clave)
                
    with tab_registro:
        st.subheader("Crear Nueva Cuenta")
        st.info("Su cuenta debe ser aprobada por jefatura (Jenier/Gabriela) antes de poder ingresar.")
        with st.form("form_registro"):
            n_nombre = st.text_input("Nombre Completo:")
            n_puesto = st.selectbox("Puesto / Rol:", ["Misceláneo/a", "Chofer", "Oficial de Seguridad", "Recepcionista", "Administrativo General"])
            n_espacio = st.selectbox("Espacio / Piso Asignado:", ["Piso 1 (30%)", "Piso 2 (40%)", "Piso 3 (30%)", "Oficina General", "Exteriores", "Flota Vehicular"])
            n_user = st.text_input("Nombre de Usuario (ej. cruiz):")
            n_pass = st.text_input("Contraseña:", type="password")
            
            if st.form_submit_button("Solicitar Acceso"):
                if n_nombre and n_user and n_pass:
                    conn = sqlite3.connect("database.db")
                    cursor = conn.cursor()
                    cursor.execute("SELECT id FROM funcionarios WHERE username=?", (n_user,))
                    if cursor.fetchone():
                        st.error("Ese nombre de usuario ya existe.")
                    else:
                        hash_p = hashlib.sha256(n_pass.encode()).hexdigest()
                        cursor.execute("INSERT INTO funcionarios (nombre, puesto, espacio_asignado, username, password, rol, estado) VALUES (?, ?, ?, ?, ?, 'BASICO', 'INACTIVO')", 
                                       (n_nombre, n_puesto, n_espacio, n_user, hash_p))
                        conn.commit()
                        conn.close()
                        sync_usuarios_excel()
                        st.success("Cuenta registrada. Jenier o Gabriela deben aprobar el ingreso.")
                else:
                    st.error("Por favor complete todos los campos.")
    st.stop()

# -------------------------------------------------------------------
# NAVEGACIÓN Y PESTAÑAS SEGÚN ROL
# -------------------------------------------------------------------
st.sidebar.write(f"👤 Usuario: **{st.session_state['usuario_actual']}**")
st.sidebar.write(f"Rol: **{st.session_state['rol']}**")
if st.sidebar.button("Cerrar Sesión"):
    logout()

if st.session_state['rol'] == "ADMIN":
    tabs = st.tabs(["📝 1. Pedimentos (Hasta 10 ítems)", "✍️ 2. Firma/Aprobación PDF", "📊 3. Informes", "✅ 4. Aprobar Usuarios"])
    tab_pedimentos, tab_firma, tab_informes, tab_usuarios = tabs
else:
    tabs = st.tabs(["📝 1. Mis Pedimentos (Hasta 10 ítems)"])
    tab_pedimentos = tabs[0]

# -------------------------------------------------------------------
# TAB 1: FORMULARIO DE PEDIMENTOS (VISTO POR TODOS)
# -------------------------------------------------------------------
with tab_pedimentos:
    st.header("📦 Solicitud de Insumos")
    
    if st.session_state['ultimo_pedimento']:
        ped = st.session_state['ultimo_pedimento']
        st.success(f"🎉 ¡Pedimento enviado con éxito! Consecutivo asignado: **#{ped['consecutivo']}**")
        st.download_button(
            label=f"📄 Descargar Pedimento #{ped['consecutivo']} (PDF para Firma Digital)",
            data=ped['bytes_pdf'],
            file_name=ped['nombre_pdf'],
            mime="application/pdf",
            type="primary"
        )
        st.markdown("---")

    conn = sqlite3.connect("database.db")
    df_articulos = pd.read_sql_query("SELECT codigo_grupal, descripcion, categoria, frecuencia_entrega, dias_vida_util FROM articulos", conn)
    
    cursor = conn.cursor()
    cursor.execute("SELECT espacio_asignado FROM funcionarios WHERE id=?", (st.session_state['func_id'],))
    res_piso = cursor.fetchone()
    piso_asignado = res_piso[0] if res_piso else "Oficina General"
    conn.close()

    colA, colB = st.columns([1, 2])
    
    with colA:
        st.subheader("Añadir Insumos")
        categorias = ["Todas"] + list(df_articulos["categoria"].dropna().unique())
        cat_sel = st.selectbox("Categoría:", categorias)
        
        df_filtro = df_articulos if cat_sel == "Todas" else df_articulos[df_articulos["categoria"] == cat_sel]
        articulo_sel = st.selectbox("Artículo:", df_filtro["descripcion"].tolist())
        item_info = df_articulos[df_articulos["descripcion"] == articulo_sel].iloc[0]
        
        cantidad = st.number_input("Cantidad:", min_value=1, value=1, step=1)
        
        if st.button("➕ Agregar al Pedimento"):
            if len(st.session_state['carrito']) >= 10:
                st.error("Límite máximo de 10 ítems por pedimento alcanzado.")
            elif any(item['codigo'] == item_info['codigo_grupal'] for item in st.session_state['carrito']):
                st.warning("Este artículo ya está en su lista actual.")
            else:
                st.session_state['carrito'].append({
                    "codigo": item_info['codigo_grupal'],
                    "descripcion": articulo_sel,
                    "cantidad": cantidad,
                    "frecuencia": item_info['frecuencia_entrega'],
                    "dias_vida": item_info['dias_vida_util']
                })
                st.success("Artículo agregado al pedimento.")

    with colB:
        st.subheader(f"🛒 Mi Pedimento Actual ({len(st.session_state['carrito'])}/10)")
        if len(st.session_state['carrito']) > 0:
            df_carrito = pd.DataFrame(st.session_state['carrito'])
            st.dataframe(df_carrito[['descripcion', 'cantidad', 'frecuencia']], use_container_width=True)
            
            if st.button("🗑️ Vaciar Lista"):
                st.session_state['carrito'] = []
                st.rerun()
            
            justificacion_global = st.text_area("Justificación del pedido:")
            if st.button("🚀 Enviar Pedimento Completo a Bodega", type="primary"):
                conn = sqlite3.connect("database.db")
                cursor = conn.cursor()
                
                ids_guardados = []
                for item in st.session_state['carrito']:
                    cursor.execute('''
                        INSERT INTO solicitudes (funcionario_id, articulo_id, cantidad, es_valido, justificacion, estado)
                        VALUES (?, ?, ?, ?, ?, 'PENDIENTE')
                    ''', (st.session_state['func_id'], item['codigo'], item['cantidad'], 1, justificacion_global))
                    ids_guardados.append(cursor.lastrowid)
                conn.commit()
                
                hash_temp = hashlib.sha256(str(datetime.now()).encode()).hexdigest()
                consecutivo, ruta_excel = generar_formulario_oficial_bys(ids_guardados, "Departamento de Servicios Generales", "Pendiente de Firma", hash_temp)
                ruta_pdf = ruta_excel.replace(".xlsx", ".pdf")
                exportar_excel_a_pdf(ruta_excel, ruta_pdf)
                conn.close()

                # Envío de correo a Jefatura
                enviar_correo_con_pdf_adjunto(
                    consecutivo=consecutivo, 
                    lista_items=st.session_state['carrito'], 
                    solicitante=st.session_state['usuario_actual'], 
                    es_valido=True, 
                    motivo="Pedido reglamentario", 
                    ruta_pdf=ruta_pdf, 
                    justificacion=justificacion_global
                )

                bytes_pdf = b""
                if os.path.exists(ruta_pdf):
                    with open(ruta_pdf, "rb") as f_pdf:
                        bytes_pdf = f_pdf.read()

                st.session_state['ultimo_pedimento'] = {
                    "consecutivo": consecutivo,
                    "bytes_pdf": bytes_pdf,
                    "nombre_pdf": f"Pedimento_{consecutivo}.pdf"
                }
                st.session_state['carrito'] = []
                st.rerun()
        else:
            st.info("No ha agregado artículos a la lista.")

# -------------------------------------------------------------------
# SECCIONES EXCLUSIVAS PARA ADMINISTRADORES (JENIER / GABRIELA)
# -------------------------------------------------------------------
if st.session_state['rol'] == "ADMIN":
    
    # TAB 2: FIRMA Y APROBACIÓN
    with tab_firma:
        st.header("✍️ Validación de Pedimentos Firmados en Adobe Acrobat")
        
        st.subheader("1. Pedimentos Pendientes de Firma")
        st.write("Descargue el documento PDF que desea firmar en Adobe Acrobat:")
        
        if os.path.exists("pedimentos_generados"):
            archivos_pdf = [f for f in os.listdir("pedimentos_generados") if f.endswith(".pdf")]
            if archivos_pdf:
                for arch in archivos_pdf:
                    ruta_c = os.path.join("pedimentos_generados", arch)
                    with open(ruta_c, "rb") as f_arch:
                        st.download_button(
                            label=f"📥 Descargar {arch} para firmar en Adobe",
                            data=f_arch.read(),
                            file_name=arch,
                            mime="application/pdf",
                            key=f"dl_{arch}"
                        )
            else:
                st.info("No hay archivos PDF pendientes en el servidor.")
        else:
            st.info("No se han generado pedimentos aún.")

        st.markdown("---")
        st.subheader("2. Cargar Pedimento PDF Ya Firmado")
        st.write("Una vez firmado en Adobe Acrobat con su tarjeta BCCR, cárguelo aquí para registrar la salida de bodega:")
        
        archivo_pdf_subido = st.file_uploader("Cargar Pedimento PDF Firmado:", type=["pdf"])
        
        if archivo_pdf_subido is not None:
            os.makedirs("pedimentos_generados", exist_ok=True)
            ruta_guardado = os.path.join("pedimentos_generados", archivo_pdf_subido.name)
            with open(ruta_guardado, "wb") as f:
                f.write(archivo_pdf_subido.getbuffer())
                
            st.info("Verificando firma digital en el documento PDF...")
            
            es_valido_firma, detalle_firma = verificar_firma_digital_adobe(ruta_guardado)
            
            if es_valido_firma:
                st.success(f"✅ **Firma Digital Verificada**: {detalle_firma}")
                consecutivo_ext = archivo_pdf_subido.name.replace("Pedimento_", "").replace(".pdf", "")
                
                if st.button(f"🔒 Confirmar y Autorizar Salida de Bodega (#{consecutivo_ext})", type="primary"):
                    conn = sqlite3.connect("database.db")
                    cursor = conn.cursor()
                    cursor.execute("""
                        UPDATE solicitudes 
                        SET estado = 'APROBADO_FIRMA_DIGITAL', 
                            firmado_por = 'Firma Digital BCCR (Adobe Acrobat)', 
                            fecha_aprobacion = CURRENT_TIMESTAMP 
                        WHERE estado = 'PENDIENTE'
                    """)
                    conn.commit()
                    conn.close()
                    st.balloons()
                    st.success("¡Pedimento procesado exitosamente en el sistema de bodega!")
            else:
                st.
