import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import hashlib
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import os

# Importar la rutina de instalación inicial
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

# Ejecutar verificación al arrancar
verificar_base_datos()

st.set_page_config(page_title="Bodega DINADECO", page_icon="📦", layout="wide")

# Inicialización de variables en Session State
if 'usuario_actual' not in st.session_state:
    st.session_state['usuario_actual'] = None
    st.session_state['rol'] = None
    st.session_state['func_id'] = None
if 'carrito' not in st.session_state:
    st.session_state['carrito'] = []

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
                st.session_state['carrito'] = []
                st.success(f"¡Pedimento enviado con éxito! Consecutivo asignado: #{consecutivo}")
        else:
            st.info("No ha agregado artículos a la lista.")

# -------------------------------------------------------------------
# SECCIONES EXCLUSIVAS PARA ADMINISTRADORES (JENIER / GABRIELA)
# -------------------------------------------------------------------
if st.session_state['rol'] == "ADMIN":
    
    # TAB 2: FIRMA Y APROBACIÓN
    with tab_firma:
        st.header("✍️ Validación de Pedimentos Firmados en Adobe Acrobat")
        st.write("Cargue el archivo PDF firmado digitalmente con su certificado de Firma Digital (BCCR).")
        
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
                st.error(f"❌ Error de Validación: {detalle_firma}")

    # TAB 3: INFORMES
    with tab_informes:
        st.header("📊 Informes Generales y Diagnóstico")
        conn = sqlite3.connect("database.db")
        df_entregas = pd.read_sql_query('''
            SELECT f.nombre AS Funcionario, f.puesto AS Puesto, f.espacio_asignado AS Ubicacion, a.descripcion AS Articulo, 
                   s.cantidad AS Cantidad, s.fecha_aprobacion AS Fecha_Entrega, s.firmado_por AS Aprobador
            FROM solicitudes s 
            JOIN funcionarios f ON s.funcionario_id = f.id 
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            WHERE s.estado = 'APROBADO_FIRMA_DIGITAL'
        ''', conn)
        conn.close()
        
        st.dataframe(df_entregas, use_container_width=True)

    # TAB 4: APROBACIÓN DE USUARIOS
    with tab_usuarios:
        st.header("✅ Aprobación de Cuentas de Usuarios")
        st.write("Gestione las solicitudes de acceso al sistema.")
        
        conn = sqlite3.connect("database.db")
        df_inactivos = pd.read_sql_query("SELECT id, nombre, puesto, espacio_asignado, username FROM funcionarios WHERE estado='INACTIVO'", conn)
        
        if df_inactivos.empty:
            st.info("No hay usuarios pendientes de aprobación.")
        else:
            st.dataframe(df_inactivos, use_container_width=True)
            user_id = st.number_input("ID del usuario a gestionar:", min_value=1, step=1)
            
            col1, col2 = st.columns(2)
            if col1.button("✅ Aprobar Acceso", type="primary"):
                cursor = conn.cursor()
                cursor.execute("UPDATE funcionarios SET estado='ACTIVO' WHERE id=?", (user_id,))
                conn.commit()
                conn.close()
                sync_usuarios_excel()
                st.success(f"Usuario ID {user_id} aprobado e incorporado al Excel Maestro.")
                st.rerun()
            
            if col2.button("❌ Rechazar / Eliminar"):
                cursor = conn.cursor()
                cursor.execute("DELETE FROM funcionarios WHERE id=?", (user_id,))
                conn.commit()
                conn.close()
                sync_usuarios_excel()
                st.error(f"Solicitud ID {user_id} eliminada.")
                st.rerun()
        if not df_inactivos.empty:
            conn.close()
