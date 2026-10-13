import streamlit as st
import streamlit.components.v1 as components
import sqlite3
import pandas as pd
from datetime import datetime
import time
import hashlib
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import glob

from setup_inicial import instalar_sistema
from exporter_bys import generar_formulario_oficial_bys
from pdf_signature import exportar_excel_a_pdf, verificar_firma_digital_adobe
from mailer import enviar_correo_con_pdf_adjunto, enviar_correo_pedimento_firmado

EXCEL_MAESTRO = "Control_Maestro_238_Articulos_DINADECO.xlsx"
MAX_INACTIVIDAD = 300  # 5 minutos de inactividad (300 segundos)

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
if 'last_activity' not in st.session_state:
    st.session_state['last_activity'] = None

now = time.time()

def logout_inactividad():
    st.session_state['usuario_actual'] = None
    st.session_state['rol'] = None
    st.session_state['func_id'] = None
    st.session_state['carrito'] = []
    st.session_state['ultimo_pedimento'] = None
    st.session_state['last_activity'] = None
    st.query_params.clear()

def obtener_timestamp_query():
    if 't' in st.query_params:
        raw_t = st.query_params['t']
        if isinstance(raw_t, (list, tuple)):
            raw_t = raw_t[0]
        try:
            return float(raw_t)
        except (ValueError, TypeError):
            return None
    return None

last_t = obtener_timestamp_query()
last_active = st.session_state.get('last_activity') or last_t

# 1. Recuperar sesión válida tras un refresco de página (F5)
if st.session_state['usuario_actual'] is None:
    if 'u' in st.query_params and last_t is not None:
        if (now - last_t) <= MAX_INACTIVIDAD:
            st.session_state['usuario_actual'] = st.query_params['u']
            st.session_state['rol'] = st.query_params.get('r', 'BASICO')
            st.session_state['func_id'] = int(st.query_params.get('fid', 0))
            st.session_state['last_activity'] = now
            st.query_params['t'] = str(int(now))
        else:
            logout_inactividad()
            st.warning("⏱️ Su sesión ha expirado por inactividad (más de 5 minutos). Inicie sesión nuevamente.")
    else:
        st.query_params.clear()
else:
    # 2. Si ya hay sesión abierta, verificar inactividad
    if last_active is not None and (now - last_active) > MAX_INACTIVIDAD:
        logout_inactividad()
        st.warning("⏱️ Su sesión se ha cerrado automáticamente por inactividad de más de 5 minutos.")
        st.rerun()
    else:
        st.session_state['last_activity'] = now
        st.query_params['t'] = str(int(now))

# Temporizador en el navegador (cliente) para cierre automático sin clics
def inyectar_temporizador_inactividad():
    js_codigo = """
    <script>
    (function() {
        const LIMITE_MS = 300000; // 5 minutos
        let ultimaActividad = Date.now();
        let reloj;

        function cerrarSesionAuto() {
            try {
                window.top.location.href = window.top.location.origin + window.top.location.pathname;
            } catch(e) {
                window.location.href = window.location.origin + window.location.pathname;
            }
        }

        function registrarAccion() {
            const momento = Date.now();
            if (momento - ultimaActividad > LIMITE_MS) {
                cerrarSesionAuto();
                return;
            }
            ultimaActividad = momento;
            clearTimeout(reloj);
            reloj = setTimeout(cerrarSesionAuto, LIMITE_MS);
        }

        const eventos = ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll', 'click'];
        eventos.forEach(ev => {
            try {
                window.top.addEventListener(ev, registrarAccion, true);
            } catch(e) {}
            try {
                window.addEventListener(ev, registrarAccion, true);
            } catch(e) {}
        });

        try {
            window.top.addEventListener('focus', registrarAccion);
        } catch(e) {}

        registrarAccion();
    })();
    </script>
    """
    components.html(js_codigo, height=0, width=0)

# Ocultar barra de desarrollador (Share, GitHub, Menú) a usuarios básicos
if st.session_state['rol'] != "ADMIN":
    st.markdown("""
        <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .stAppHeader {display: none;}
        </style>
    """, unsafe_allow_html=True)

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
            st.session_state['last_activity'] = time.time()
            
            st.query_params['u'] = result[1]
            st.query_params['r'] = result[2]
            st.query_params['fid'] = str(result[0])
            st.query_params['t'] = str(int(time.time()))
            st.rerun()
    else:
        st.error("Usuario o contraseña incorrectos.")

def logout():
    logout_inactividad()
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

def obtener_ultimo_consecutivo_db():
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
    conn.close()
    return row[0] if row else 50

def fijar_ultimo_consecutivo_db(nuevo_num):
    anio_actual = datetime.now().year
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS control_consecutivos (
            anio INTEGER PRIMARY KEY,
            ultimo_numero INTEGER NOT NULL
        )
    """)
    cursor.execute("INSERT OR REPLACE INTO control_consecutivos (anio, ultimo_numero) VALUES (?, ?)", (anio_actual, nuevo_num))
    conn.commit()
    conn.close()

def limpiar_historial_y_archivos():
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM solicitudes")
    conn.commit()
    conn.close()

    folder = "pedimentos_generados"
    if os.path.exists(folder):
        archivos = glob.glob(os.path.join(folder, "*"))
        for f in archivos:
            try:
                os.remove(f)
            except Exception:
                pass

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

# Inyectar el temporizador activo en el navegador una vez autenticado
inyectar_temporizador_inactividad()

# -------------------------------------------------------------------
# NAVEGACIÓN Y BOTÓN DE CIERRE DE SESIÓN DESTACADO
# -------------------------------------------------------------------
st.sidebar.markdown("### 👤 Sesión Activa")
st.sidebar.write(f"Usuario: **{st.session_state['usuario_actual']}**")
st.sidebar.write(f"Rol: **{st.session_state['rol']}**")

if st.sidebar.button("🔴 Cerrar Sesión", type="primary", use_container_width=True):
    logout()

st.sidebar.markdown("---")

if st.session_state['rol'] == "ADMIN":
    tabs = st.tabs(["📝 1. Pedimentos", "✍️ 2. Firma/Aprobación PDF", "📊 3. Informes", "✅ 4. Aprobar Usuarios", "⚙️ 5. Control Master"])
    tab_pedimentos, tab_firma, tab_informes, tab_usuarios, tab_master = tabs
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
        st.success(f"🎉 ¡Pedimento procesado con éxito! Consecutivo asignado: **#{ped['consecutivo']}**")
        
        if ped.get("exito_mail"):
            st.info("📩 El documento ha sido enviado por correo para la respectiva Firma Digital.")
        else:
            st.error(f"⚠️ **Atención:** Ocurrió un inconveniente al enviar el correo automático: **{ped.get('msg_mail')}**")
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
            if st.button("🚀 Enviar Pedimento Completo a Jefatura", type="primary"):
                conn = sqlite3.connect("database.db")
                cursor = conn.cursor()
                
                ids_guardados = []
                for item in st.session_state['carrito']:
                    cursor.execute("""
                        INSERT INTO solicitudes (funcionario_id, articulo_id, cantidad, es_valido, justificacion, estado)
                        VALUES (?, ?, ?, ?, ?, 'PENDIENTE')
                    """, (st.session_state['func_id'], item['codigo'], item['cantidad'], 1, justificacion_global))
                    ids_guardados.append(cursor.lastrowid)
                conn.commit()
                
                hash_temp = hashlib.sha256(str(datetime.now()).encode()).hexdigest()
                consecutivo, ruta_excel = generar_formulario_oficial_bys(ids_guardados, "Departamento de Servicios Generales", "Pendiente de Firma", hash_temp)
                ruta_pdf = ruta_excel.replace(".xlsx", ".pdf")
                exportar_excel_a_pdf(ruta_excel, ruta_pdf)
                conn.close()

                exito_mail, msg_mail = enviar_correo_con_pdf_adjunto(
                    consecutivo=consecutivo, 
                    lista_items=st.session_state['carrito'], 
                    solicitante=st.session_state['usuario_actual'], 
                    es_valido=True, 
                    motivo="Pedido reglamentario", 
                    ruta_pdf=ruta_pdf, 
                    justificacion=justificacion_global
                )

                st.session_state['ultimo_pedimento'] = {
                    "consecutivo": consecutivo,
                    "exito_mail": exito_mail,
                    "msg_mail": msg_mail
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

                    exito_envio, msg_envio = enviar_correo_pedimento_firmado(
                        consecutivo=consecutivo_ext,
                        ruta_pdf_firmado=ruta_guardado,
                        detalle_firma=detalle_firma
                    )

                    st.balloons()
                    st.success(f"¡Pedimento #{consecutivo_ext} procesado y autorizado exitosamente en bodega!")
                    
                    if exito_envio:
                        st.info("📩 El pedimento firmado ha sido enviado por correo a **agomez@dinadeco.go.cr** (Encargado de Bodega) con copia a **jcamacho@dinadeco.go.cr** y **agonzalez@dinadeco.go.cr**.")
                    else:
                        st.warning(f"⚠️ El pedimento fue autorizado en el sistema, pero ocurrió un aviso con el correo: {msg_envio}")
            else:
                st.error(f"❌ Error de Validación: {detalle_firma}")

    # TAB 3: INFORMES
    with tab_informes:
        st.header("📊 Informes Generales y Diagnóstico")
        conn = sqlite3.connect("database.db")
        df_entregas = pd.read_sql_query("""
            SELECT f.nombre AS Funcionario, f.puesto AS Puesto, f.espacio_asignado AS Ubicacion, a.descripcion AS Articulo, 
                   s.cantidad AS Cantidad, s.fecha_aprobacion AS Fecha_Entrega, s.firmado_por AS Aprobador
            FROM solicitudes s 
            JOIN funcionarios f ON s.funcionario_id = f.id 
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            WHERE s.estado = 'APROBADO_FIRMA_DIGITAL'
        """, conn)
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

    # TAB 5: CONTROL MASTER, USUARIOS Y REINICIO
    with tab_master:
        st.header("⚙️ Control Master del Sistema")
        st.write("Herramientas exclusivas para el usuario Administrador Master.")
        
        st.markdown("---")
        st.subheader("1. Gestión de Usuarios y Contraseñas")
        
        conn = sqlite3.connect("database.db")
        df_todos_usuarios = pd.read_sql_query("SELECT id, nombre, puesto, espacio_asignado, username, rol, estado FROM funcionarios", conn)
        
        total_users = len(df_todos_usuarios)
        activos_count = len(df_todos_usuarios[df_todos_usuarios['estado'] == 'ACTIVO'])
        pendientes_count = len(df_todos_usuarios[df_todos_usuarios['estado'] == 'INACTIVO'])
        
        col1, col2, col
