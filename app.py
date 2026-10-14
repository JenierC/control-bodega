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
import json
import base64
import re
import urllib.request
import pypdf
import io

from setup_inicial import instalar_sistema
from exporter_bys import generar_formulario_oficial_bys
from pdf_signature import exportar_excel_a_pdf, verificar_firma_digital_adobe
from mailer import enviar_correo_con_pdf_adjunto, enviar_correo_pedimento_firmado

EXCEL_MAESTRO = "Control_Maestro_238_Articulos_DINADECO.xlsx"
MAX_INACTIVIDAD = 300  # 5 minutos de inactividad
CLAVE_TEMPORAL_DEFAULT = "Dinadeco2026*"

LISTA_USUARIOS_OFICIALES = [
    {"nombre": "Emmanuel Alvarez", "email": "ealvarez@dinadeco.go.cr", "username": "ealvarez", "puesto": "Misceláneo/a", "espacio": "Piso 3 (30%)", "rol": "BASICO"},
    {"nombre": "Ana Orozco Sánchez", "email": "aorozco@dinadeco.go.cr", "username": "aorozco", "puesto": "Misceláneo/a", "espacio": "Piso 1 (30%)", "rol": "BASICO"},
    {"nombre": "Ana Yancy González", "email": "agonzalez@dinadeco.go.cr", "username": "agonzalez", "puesto": "Recepcionista", "espacio": "Oficina General", "rol": "ADMIN"},
    {"nombre": "Pamela Quirós", "email": "pquiros@dinadeco.go.cr", "username": "pquiros", "puesto": "Administrativa", "espacio": "Oficina General", "rol": "ADMIN"},
    {"nombre": "Glenda Loría", "email": "gloria@dinadeco.go.cr", "username": "gloria", "puesto": "Administrativa", "espacio": "Oficina General", "rol": "ADMIN"},
    {"nombre": "Ronald Montero", "email": "rmontero@dinadeco.go.cr", "username": "rmontero", "puesto": "Chofer", "espacio": "Flota Vehicular", "rol": "BASICO"},
    {"nombre": "Randall Solano", "email": "rsolano@dinadeco.go.cr", "username": "rsolano", "puesto": "Chofer", "espacio": "Flota Vehicular", "rol": "BASICO"},
    {"nombre": "Nelson Levey", "email": "nlevey@dinadeco.go.cr", "username": "nlevey", "puesto": "Chofer", "espacio": "Flota Vehicular", "rol": "BASICO"},
    {"nombre": "Hans Arrieta Morales", "email": "harrieta@dinadeco.go.cr", "username": "harrieta", "puesto": "Oficial de Seguridad", "espacio": "Exteriores", "rol": "BASICO"},
    {"nombre": "Henry Guevara", "email": "hguevara@dinadeco.go.cr", "username": "hguevara", "puesto": "Oficial de Seguridad", "espacio": "Exteriores", "rol": "BASICO"},
    {"nombre": "Jorge Barquero", "email": "jbarquero@dinadeco.go.cr", "username": "jbarquero", "puesto": "Oficial de Seguridad", "espacio": "Exteriores", "rol": "BASICO"},
    {"nombre": "Miguel Granados", "email": "mgranados@dinadeco.go.cr", "username": "mgranados", "puesto": "Oficial de Seguridad", "espacio": "Exteriores", "rol": "BASICO"}
]

# ==========================================
# VERIFICACIÓN AUTOMÁTICA Y MIGRACIÓN BD
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

    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA table_info(solicitudes)")
    cols_sol = [row[1] for row in cursor.fetchall()]
    nuevas_cols_sol = {
        'consecutivo': 'TEXT',
        'cantidad_despachada': 'INTEGER',
        'diferencia': 'INTEGER',
        'retirado_por': 'TEXT',
        'fecha_despacho_real': 'TEXT',
        'ruta_pdf_escaneado': 'TEXT'
    }
    for col, col_type in nuevas_cols_sol.items():
        if col not in cols_sol:
            try:
                cursor.execute(f"ALTER TABLE solicitudes ADD COLUMN {col} {col_type}")
            except Exception:
                pass

    cursor.execute("PRAGMA table_info(funcionarios)")
    cols_func = [row[1] for row in cursor.fetchall()]
    if 'debe_cambiar_password' not in cols_func:
        try:
            cursor.execute("ALTER TABLE funcionarios ADD COLUMN debe_cambiar_password INTEGER DEFAULT 0")
        except Exception:
            pass
    if 'email' not in cols_func:
        try:
            cursor.execute("ALTER TABLE funcionarios ADD COLUMN email TEXT")
        except Exception:
            pass

    hash_temp = hashlib.sha256(CLAVE_TEMPORAL_DEFAULT.encode()).hexdigest()
    for u in LISTA_USUARIOS_OFICIALES:
        cursor.execute("SELECT id FROM funcionarios WHERE username=?", (u["username"],))
        if not cursor.fetchone():
            cursor.execute("""
                INSERT INTO funcionarios (nombre, puesto, espacio_asignado, username, email, password, rol, estado, debe_cambiar_password)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'ACTIVO', 1)
            """, (u["nombre"], u["puesto"], u["espacio"], u["username"], u["email"], hash_temp, u["rol"]))

    conn.commit()
    conn.close()

verificar_base_datos()

# ===================================================================
# FUNCIÓN DE EXTRACCIÓN MEJORADA: TEXTO DIGITAL Y VISIÓN AI
# ===================================================================
def extraer_datos_pedimento_pdf(pdf_bytes, filename):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    full_text = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            full_text += t + "\n"
            
    match_cons = re.search(r'Consecutivo:\s*([0-9]{3}-[0-9]{4})', full_text, re.IGNORECASE)
    if not match_cons:
        match_cons = re.search(r'([0-9]{3}-[0-9]{4})', filename)
    consecutivo = match_cons.group(1) if match_cons else "051-2026"
    
    articulos_encontrados = []
    
    # 1. BÚSQUEDA ROBUSTA EN EL TEXTO DIGITAL (Cruce contra catálogo de la base de datos)
    if len(full_text.strip()) > 20:
        try:
            conn_cat = sqlite3.connect("database.db")
            cursor_cat = conn_cat.cursor()
            cursor_cat.execute("SELECT codigo_grupal, descripcion FROM articulos")
            catalogo = cursor_cat.fetchall()
            conn_cat.close()
            
            lines = [l.strip() for l in full_text.split("\n") if l.strip()]
            
            for cod, desc in catalogo:
                desc_norm = desc.strip().upper()
                # Si el insumo está presente en el documento
                if len(desc_norm) > 5 and desc_norm in full_text.upper():
                    cant_sol = 1
                    cant_desp = 1
                    for idx, line in enumerate(lines):
                        if desc_norm[:18] in line.upper():
                            # Buscar la cantidad en las siguientes 5 líneas generadas por Excel
                            nums_next = []
                            for nxt in lines[idx+1:idx+6]:
                                if nxt.isdigit() and int(nxt) < 1000:
                                    nums_next.append(int(nxt))
                            
                            if nums_next:
                                cant_sol = nums_next[0]
                                cant_desp = nums_next[1] if len(nums_next) > 1 else cant_sol
                            else:
                                nums_same = re.findall(r'\b\d+\b', line)
                                if nums_same and nums_same[-1].isdigit():
                                    cant_sol = int(nums_same[-1])
                                    cant_desp = cant_sol
                            break
                            
                    articulos_encontrados.append({
                        "codigo": cod,
                        "descripcion": desc,
                        "cantidad_solicitada": cant_sol,
                        "cantidad_despachada": cant_desp
                    })
        except Exception:
            pass

    # 2. SI NO HUBO TEXTO DIGITAL (Escaneo plano), RECURRIR A GEMINI VISION
    if not articulos_encontrados:
        img_bytes = None
        for page in reader.pages:
            if page.images:
                img_bytes = page.images[0].data
                break
                
        api_key = st.secrets.get("gemini_api_key", None) if "gemini_api_key" in st.secrets else os.environ.get("GEMINI_API_KEY", None)
        if api_key and img_bytes:
            try:
                b64_img = base64.b64encode(img_bytes).decode("utf-8")
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
                prompt = """
                Analiza este pedimento oficial de DINADECO. Extrae en JSON estricto:
                {
                  "consecutivo": "ej: 051-2026",
                  "articulos": [
                    {
                      "codigo": "Código numérico si aparece",
                      "descripcion": "Nombre del insumo",
                      "cantidad_solicitada": 4,
                      "cantidad_despachada": 4
                    }
                  ]
                }
                Devuelve ÚNICAMENTE el JSON.
                """
                payload = {
                    "contents": [{
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": "image/jpeg", "data": b64_img}}
                        ]
                    }]
                }
                req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers={'Content-Type': 'application/json'})
                with urllib.request.urlopen(req, timeout=15) as resp:
                    data = json.loads(resp.read().decode('utf-8'))
                    raw = data['candidates'][0]['content']['parts'][0]['text']
                    raw = re.sub(r'```json\s*', '', raw).replace('```', '').strip()
                    parsed = json.loads(raw)
                    if parsed.get("consecutivo"):
                        consecutivo = parsed["consecutivo"]
                    for a in parsed.get("articulos", []):
                        articulos_encontrados.append({
                            "codigo": a.get("codigo", ""),
                            "descripcion": a.get("descripcion", ""),
                            "cantidad_solicitada": int(a.get("cantidad_solicitada", 1)),
                            "cantidad_despachada": int(a.get("cantidad_despachada", a.get("cantidad_solicitada", 1)))
                        })
            except Exception:
                pass

    return consecutivo, articulos_encontrados

st.set_page_config(page_title="Bodega DINADECO", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

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
if 'debe_cambiar_password' not in st.session_state:
    st.session_state['debe_cambiar_password'] = False
if 'articulos_manuales_tab2' not in st.session_state:
    st.session_state['articulos_manuales_tab2'] = []

now = time.time()

def logout_inactividad():
    st.session_state['usuario_actual'] = None
    st.session_state['rol'] = None
    st.session_state['func_id'] = None
    st.session_state['carrito'] = []
    st.session_state['ultimo_pedimento'] = None
    st.session_state['last_activity'] = None
    st.session_state['debe_cambiar_password'] = False
    st.session_state['articulos_manuales_tab2'] = []
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

if st.session_state['usuario_actual'] is None:
    if 'u' in st.query_params and last_t is not None:
        if (now - last_t) <= MAX_INACTIVIDAD:
            st.session_state['usuario_actual'] = st.query_params['u']
            st.session_state['rol'] = st.query_params.get('r', 'BASICO')
            st.session_state['func_id'] = int(st.query_params.get('fid', 0))
            st.session_state['last_activity'] = now
            st.query_params['t'] = str(int(now))
            
            conn = sqlite3.connect("database.db")
            cursor = conn.cursor()
            cursor.execute("SELECT COALESCE(debe_cambiar_password, 0) FROM funcionarios WHERE id=?", (st.session_state['func_id'],))
            row_cp = cursor.fetchone()
            conn.close()
            st.session_state['debe_cambiar_password'] = bool(row_cp[0]) if row_cp else False
        else:
            logout_inactividad()
            st.warning("⏱️ Su sesión ha expirado por inactividad. Inicie sesión de nuevo.")
    else:
        st.query_params.clear()
else:
    if last_active is not None and (now - last_active) > MAX_INACTIVIDAD:
        logout_inactividad()
        st.warning("⏱️ Su sesión se ha cerrado automáticamente por inactividad (5 minutos).")
        st.rerun()
    else:
        st.session_state['last_activity'] = now
        st.query_params['t'] = str(int(now))

def inyectar_temporizador_inactividad():
    js_codigo = """
    <script>
    (function() {
        const LIMITE_MS = 300000;
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
            try { window.top.addEventListener(ev, registrarAccion, true); } catch(e) {}
            try { window.addEventListener(ev, registrarAccion, true); } catch(e) {}
        });

        try { window.top.addEventListener('focus', registrarAccion); } catch(e) {}
        registrarAccion();
    })();
    </script>
    """
    components.html(js_codigo, height=0, width=0)

if st.session_state['rol'] != "ADMIN":
    st.markdown("""
        <style>
        [data-testid="stToolbarActions"],
        .stToolbarActionButton,
        [data-testid="stToolbar"] a,
        [data-testid="stToolbar"] button:not([kind="headerNoPadding"]),
        .stAppDeployButton,
        div[class*="viewerBadge"],
        #MainMenu,
        footer {
            display: none !important;
            visibility: hidden !important;
        }
        [data-testid="stSidebarCollapsedControl"] {
            display: block !important;
            visibility: visible !important;
        }
        [data-testid="stSidebarCollapsedControl"] button {
            display: inline-flex !important;
            visibility: visible !important;
        }
        </style>
    """, unsafe_allow_html=True)

def login(user, pwd):
    hash_pwd = hashlib.sha256(pwd.encode()).hexdigest()
    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, rol, estado, COALESCE(debe_cambiar_password, 0) FROM funcionarios WHERE username=? AND password=?", (user, hash_pwd))
    result = cursor.fetchone()
    conn.close()
    
    if result:
        if result[3] == "INACTIVO":
            st.error("Su cuenta está pendiente de aprobación por Jenier o Gabriela.")
        else:
            st.session_state['func_id'] = result[0]
            st.session_state['usuario_actual'] = result[1]
            st.session_state['rol'] = result[2]
            st.session_state['debe_cambiar_password'] = bool(result[4])
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
    df_users = pd.read_sql_query("SELECT id, nombre, puesto, espacio_asignado, username, email, rol, estado FROM funcionarios", conn)
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

    for folder in ["pedimentos_generados", "pedimentos_escaneados"]:
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
                        cursor.execute("INSERT INTO funcionarios (nombre, puesto, espacio_asignado, username, password, rol, estado, debe_cambiar_password) VALUES (?, ?, ?, ?, ?, 'BASICO', 'INACTIVO', 0)", 
                                       (n_nombre, n_puesto, n_espacio, n_user, hash_p))
                        conn.commit()
                        conn.close()
                        sync_usuarios_excel()
                        st.success("Cuenta registrada. Jenier o Gabriela deben aprobar el ingreso.")
                else:
                    st.error("Por favor complete todos los campos.")
    st.stop()

# ===================================================================
# PANTALLA OBLIGATORIA: CAMBIO DE CONTRASEÑA EN EL PRIMER INICIO
# ===================================================================
if st.session_state.get('debe_cambiar_password', False):
    st.warning("🔒 **Primer Inicio de Sesión:** Por motivos de seguridad institucional, debe definir una contraseña personal antes de continuar.")
    with st.form("form_cambio_clave_obligatorio"):
        st.write(f"Usuario: **{st.session_state['usuario_actual']}**")
        nueva_clave = st.text_input("Nueva Contraseña:", type="password")
        conf_clave = st.text_input("Confirmar Nueva Contraseña:", type="password")
        
        col_c1, col_c2 = st.columns([1, 4])
        if col_c1.form_submit_button("💾 Guardar y Entrar", type="primary"):
            if not nueva_clave or len(nueva_clave) < 6:
                st.error("La nueva contraseña debe tener al menos 6 caracteres.")
            elif nueva_clave != conf_clave:
                st.error("Las contraseñas no coinciden.")
            elif nueva_clave == CLAVE_TEMPORAL_DEFAULT:
                st.error("No puede utilizar la misma contraseña temporal provisional.")
            else:
                hash_nueva = hashlib.sha256(nueva_clave.encode()).hexdigest()
                conn = sqlite3.connect("database.db")
                cursor = conn.cursor()
                cursor.execute("UPDATE funcionarios SET password=?, debe_cambiar_password=0 WHERE id=?", (hash_nueva, st.session_state['func_id']))
                conn.commit()
                conn.close()
                sync_usuarios_excel()
                st.session_state['debe_cambiar_password'] = False
                st.success("✅ ¡Contraseña actualizada exitosamente! Ingresando al sistema...")
                time.sleep(1)
                st.rerun()
        if col_c2.form_submit_button("Cerrar Sesión"):
            logout()
    st.stop()

inyectar_temporizador_inactividad()

# -------------------------------------------------------------------
# NAVEGACIÓN Y BOTÓN DE CIERRE DE SESIÓN EN SIDEBAR
# -------------------------------------------------------------------
st.sidebar.markdown("### 👤 Sesión Activa")
st.sidebar.write(f"Usuario: **{st.session_state['usuario_actual']}**")
st.sidebar.write(f"Rol: **{st.session_state['rol']}**")

if st.sidebar.button("🔴 Cerrar Sesión", type="primary", use_container_width=True, key="btn_logout_sidebar"):
    logout()

st.sidebar.markdown("---")

if st.session_state['rol'] == "ADMIN":
    tabs = st.tabs([
        "📝 1. Pedimentos", 
        "✍️ 2. Firma/Aprobación PDF", 
        "📦 3. Despacho/Recepción Bodega", 
        "📊 4. Informes", 
        "✅ 5. Aprobar Usuarios", 
        "⚙️ 6. Control Master"
    ])
    tab_pedimentos, tab_firma, tab_recepcion, tab_informes, tab_usuarios, tab_master = tabs
else:
    tabs = st.tabs(["📝 1. Mis Pedimentos (Hasta 10 ítems)"])
    tab_pedimentos = tabs[0]

# -------------------------------------------------------------------
# TAB 1: FORMULARIO DE PEDIMENTOS (CON CANDADO DE VIDA ÚTIL)
# -------------------------------------------------------------------
with tab_pedimentos:
    col_u1, col_u2 = st.columns([4, 1])
    with col_u1:
        st.caption(f"👤 Conectado: **{st.session_state['usuario_actual']}** ({st.session_state['rol']})")
    with col_u2:
        if st.button("🔴 Salir", key="btn_logout_top", use_container_width=True):
            logout()

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

    colA, colB = st.columns([1, 2])
    
    with colA:
        st.subheader("Añadir Insumos")
        categorias = ["Todas"] + list(df_articulos["categoria"].dropna().unique())
        cat_sel = st.selectbox("Categoría:", categorias)
        
        df_filtro = df_articulos if cat_sel == "Todas" else df_articulos[df_articulos["categoria"] == cat_sel]
        articulo_sel = st.selectbox("Artículo:", df_filtro["descripcion"].tolist())
        item_info = df_articulos[df_articulos["descripcion"] == articulo_sel].iloc[0]
        
        cantidad = st.number_input("Cantidad:", min_value=1, value=1, step=1)
        
        cursor.execute('''
            SELECT COALESCE(fecha_despacho_real, fecha_aprobacion)
            FROM solicitudes
            WHERE funcionario_id = ? AND articulo_id = ? AND estado LIKE 'ENTREGADO%'
            ORDER BY id DESC LIMIT 1
        ''', (st.session_state['func_id'], item_info['codigo_grupal']))
        last_req = cursor.fetchone()
        
        bloqueado = False
        msg_bloqueo = ""
        dias_faltantes = 0
        
        if last_req and last_req[0] and pd.notna(item_info['dias_vida_util']):
            try:
                fecha_str = str(last_req[0]).split()[0]
                if "/" in fecha_str:
                    last_date = datetime.strptime(fecha_str, "%d/%m/%Y").date()
                else:
                    last_date = datetime.strptime(fecha_str, "%Y-%m-%d").date()
                    
                dias_transcurridos = (datetime.now().date() - last_date).days
                vida_util = int(item_info['dias_vida_util'])
                
                if dias_transcurridos < vida_util:
                    bloqueado = True
                    dias_faltantes = vida_util - dias_transcurridos
                    msg_bloqueo = f"⛔ **Frecuencia de Entrega No Cumplida:**\nEste artículo se le despachó físicamente el **{last_date.strftime('%d/%m/%Y')}**.\nDebe esperar **{dias_faltantes} días** más para volver a solicitarlo (Vida útil reglamentaria: {vida_util} días)."
            except Exception:
                pass
                
        if bloqueado:
            st.error(msg_bloqueo)
            st.info("💡 Si ocurrió una emergencia (ej. derrame, daño, rotura), justifíquelo para habilitar la solicitud por excepción:")
            just_excepcion = st.text_input("Justificación de Excepción:", key="just_exc")
            if just_excepcion.strip():
                if st.button("⚠️ Solicitar por Excepción"):
                    if len(st.session_state['carrito']) >= 10:
                        st.error("Límite máximo de 10 ítems por pedimento alcanzado.")
                    elif any(item['codigo'] == item_info['codigo_grupal'] for item in st.session_state['carrito']):
                        st.warning("Este artículo ya está en su lista actual.")
                    else:
                        st.session_state['carrito'].append({
                            "codigo": item_info['codigo_grupal'],
                            "descripcion": f"{articulo_sel} [EXCEPCIÓN: {just_excepcion}]",
                            "cantidad": cantidad,
                            "frecuencia": item_info['frecuencia_entrega'],
                            "dias_vida": item_info['dias_vida_util']
                        })
                        st.success("Artículo agregado al pedimento como excepción.")
        else:
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
        
    conn.close()

    with colB:
        st.subheader(f"🛒 Mi Pedimento Actual ({len(st.session_state['carrito'])}/10)")
        if len(st.session_state['carrito']) > 0:
            df_carrito = pd.DataFrame(st.session_state['carrito'])
            st.dataframe(df_carrito[['descripcion', 'cantidad', 'frecuencia']], use_container_width=True)
            
            if st.button("🗑️ Vaciar Lista"):
                st.session_state['carrito'] = []
                st.rerun()
            
            justificacion_global = st.text_area("Justificación general del pedido (Opcional):")
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
                
                placeholders = ",".join("?" * len(ids_guardados))
                cursor.execute(f"UPDATE solicitudes SET consecutivo = ? WHERE id IN ({placeholders})", [consecutivo] + ids_guardados)
                conn.commit()
                
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
# SECCIONES EXCLUSIVAS PARA ADMINISTRADORES
# -------------------------------------------------------------------
if st.session_state['rol'] == "ADMIN":
    
    # ===============================================================
    # TAB 2: FIRMA Y APROBACIÓN
    # ===============================================================
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
        
        st.subheader("2. Cargar Pedimento PDF Firmado (Trámite Normal - Con Envío a Bodega)")
        st.write("Para pedidos nuevos: cárguelo aquí para registrar la firma digital y **enviar el correo automático a Bodega**:")
        
        archivo_pdf_subido = st.file_uploader("Cargar Pedimento PDF Firmado:", type=["pdf"], key="uploader_firmado")
        
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

        st.markdown("---")
        
        st.subheader("3. Cargar Pedimento Ya Firmado (Histórico / Sin Envío a Bodega)")
        st.write("Suba un pedimento que **ya fue firmado y solicitado a bodega anteriormente**. El sistema lo registrará como aprobado **SIN enviar correos a bodega**, dejándolo listo para la Pestaña 3.")
        
        conn = sqlite3.connect("database.db")
        df_func_tab2 = pd.read_sql_query("SELECT id, nombre, puesto FROM funcionarios WHERE estado='ACTIVO' ORDER BY nombre ASC", conn)
        df_art_tab2 = pd.read_sql_query("SELECT codigo_grupal, descripcion FROM articulos ORDER BY descripcion ASC", conn)
        conn.close()
        
        pdf_historico_firmado = st.file_uploader("📂 Subir PDF ya firmado por Jefatura:", type=["pdf"], key="uploader_hist_firmado")
        
        if pdf_historico_firmado is not None:
            bytes_hist_firmado = pdf_historico_firmado.read()
            
            with st.spinner("🔍 Analizando documento PDF..."):
                consecutivo_leido, articulos_leidos = extraer_datos_pedimento_pdf(bytes_hist_firmado, pdf_historico_firmado.name)
            
            col_h1, col_h2 = st.columns(2)
            with col_h1:
                func_hist_solicitante_id = st.selectbox(
                    "👤 ¿Quién solicitó este pedimento?:", 
                    df_func_tab2['id'].tolist(),
                    format_func=lambda x: f"{df_func_tab2[df_func_tab2['id'] == x]['nombre'].values[0]} ({df_func_tab2[df_func_tab2['id'] == x]['puesto'].values[0]})",
                    key="sel_func_hist_tab2"
                )
            with col_h2:
                consecutivo_hist_final = st.text_input("Número de Consecutivo:", value=consecutivo_leido, key="txt_consecutivo_hist_tab2")
            
            st.markdown("#### Artículos del Pedimento Firmado:")
            
            if articulos_leidos:
                st.success(f"✅ Se leyeron {len(articulos_leidos)} artículo(s) automáticamente del PDF:")
                df_art_view = pd.DataFrame(articulos_leidos)[['codigo', 'descripcion', 'cantidad_solicitada']]
                st.dataframe(df_art_view, use_container_width=True)
                lista_articulos_final = articulos_leidos
            else:
                st.info("ℹ️ **Documento Escaneado Plano:** Si es una imagen escaneada sin texto seleccionable, elija los artículos a continuación:")
                
                col_add1, col_add2, col_add3 = st.columns([3, 1, 1])
                with col_add1:
                    art_esc_sel = st.selectbox("Insumo:", df_art_tab2['descripcion'].tolist(), key="sel_art_esc_tab2")
                with col_add2:
                    cant_esc_sel = st.number_input("Cantidad Solicitada:", min_value=1, value=4, step=1, key="num_cant_esc_tab2")
                with col_add3:
                    st.write("")
                    st.write("")
                    if st.button("➕ Agregar Artículo", key="btn_add_esc_tab2"):
                        cod_esc = df_art_tab2[df_art_tab2['descripcion'] == art_esc_sel].iloc[0]['codigo_grupal']
                        st.session_state['articulos_manuales_tab2'].append({
                            "codigo": cod_esc,
                            "descripcion": art_esc_sel,
                            "cantidad_solicitada": cant_esc_sel
                        })
                        st.success("Artículo agregado a la lista histórica.")
                        
                if st.session_state['articulos_manuales_tab2']:
                    st.dataframe(pd.DataFrame(st.session_state['articulos_manuales_tab2'])[['codigo', 'descripcion', 'cantidad_solicitada']], use_container_width=True)
                    if st.button("🗑️ Limpiar Artículos", key="btn_clear_esc_tab2"):
                        st.session_state['articulos_manuales_tab2'] = []
                        st.rerun()
                
                lista_articulos_final = st.session_state['articulos_manuales_tab2']
                
            os.makedirs("pedimentos_generados", exist_ok=True)
            ruta_guardado_hist = os.path.join("pedimentos_generados", f"Pedimento_{consecutivo_hist_final}.pdf")
            with open(ruta_guardado_hist, "wb") as f_h:
                f_h.write(bytes_hist_firmado)
                
            if st.button(f"💾 Registrar Pedimento #{consecutivo_hist_final} como Aprobado (Sin Enviar Correos)", type="primary", key="btn_save_hist_tab2"):
                if not lista_articulos_final:
                    st.error("Por favor agregue al menos un artículo a la lista antes de guardar.")
                else:
                    conn = sqlite3.connect("database.db")
                    cursor = conn.cursor()
                    
                    cursor.execute("DELETE FROM solicitudes WHERE consecutivo = ?", (consecutivo_hist_final,))
                    
                    for a_item in lista_articulos_final:
                        cod_a = a_item.get("codigo", "")
                        desc_a = a_item.get("descripcion", "")
                        cant_a = int(a_item.get("cantidad_solicitada", 1))
                        
                        cursor.execute("SELECT codigo_grupal FROM articulos WHERE codigo_grupal = ? OR descripcion LIKE ?", (cod_a, f"%{desc_a[:15]}%"))
                        row_c = cursor.fetchone()
                        cod_final = row_c[0] if row_c else cod_a
                        
                        cursor.execute("""
                            INSERT INTO solicitudes (
                                funcionario_id, articulo_id, cantidad, es_valido,
                                justificacion, estado, consecutivo, firmado_por, fecha_aprobacion
                            ) VALUES (?, ?, ?, 1, 'Pedimento Histórico Registrado en Pestaña 2', 'APROBADO_FIRMA_DIGITAL', ?, 'Firma Digital BCCR (Histórico)', CURRENT_TIMESTAMP)
                        """, (func_hist_solicitante_id, cod_final, cant_a, consecutivo_hist_final))
                        
                    conn.commit()
                    conn.close()
                    st.session_state['articulos_manuales_tab2'] = []
                    st.balloons()
                    st.success(f"🎉 ¡Pedimento #{consecutivo_hist_final} registrado como APROBADO sin enviar correos! Ahora diríjase a la **Pestaña 3** para cotejar lo entregado físicamente.")

    # ===============================================================
    # TAB 3: RECEPCIÓN Y DESPACHO FÍSICO DE BODEGA
    # ===============================================================
    with tab_recepcion:
        st.header("📦 Recepción de Bodega y Verificación de Entregas")
        st.write("Suba el documento PDF escaneado con los sellos físicos de bodega (*MERCADERÍA DESPACHADA* y *MERCADERÍA RETIRADA*) para realizar el cotejo automático.")
        
        pdf_escaneado_file = st.file_uploader("📂 Subir Pedimento Escaneado devuelto por Bodega:", type=["pdf"], key="uploader_escaneado")
        
        if pdf_escaneado_file is not None:
            bytes_escaneados = pdf_escaneado_file.read()
            match = re.search(r'(\d{3}-\d{4})', pdf_escaneado_file.name)
            consecutivo_detectado = match.group(1) if match else "053-2026"
            
            st.success(f"✅ **Documento Detectado:** Pedimento #{consecutivo_detectado}")
            
            conn = sqlite3.connect("database.db")
            df_solicitud_exacta = pd.read_sql_query("""
                SELECT s.id, a.codigo_grupal, a.descripcion, s.cantidad AS cantidad_solicitada,
                       COALESCE(s.cantidad_despachada, s.cantidad) AS cantidad_despachada,
                       f.nombre AS solicitante
                FROM solicitudes s
                JOIN articulos a ON s.articulo_id = a.codigo_grupal
                JOIN funcionarios f ON s.funcionario_id = f.id
                WHERE s.consecutivo = ?
            """, conn, params=(consecutivo_detectado,))
            conn.close()
            
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                consecutivo_final = st.text_input("Consecutivo:", value=consecutivo_detectado, key="r_consecutivo")
            with col_m2:
                fecha_desp = st.text_input("Fecha en el Sello (DD/MM/AAAA):", value="13/10/2026", key="r_fecha")
            with col_m3:
                persona_retiro = st.text_input("Retirado por (del Sello):", value="Pamela CA", key="r_retiro")
                
            st.markdown("---")
            st.subheader("📋 Cotejo de Insumos: Solicitado vs. Despachado Real")
            
            if df_solicitud_exacta.empty:
                st.warning(f"⚠️ El pedimento #{consecutivo_detectado} aún no ha sido cargado en la **Pestaña 2 (Punto 3)**. Por favor cárguelo allí primero para que el sistema conozca los artículos pedidos y a quién pertenecía.")
            else:
                st.info(f"👤 Solicitante registrado en sistema: **{df_solicitud_exacta.iloc[0]['solicitante']}**")
                
                items_cotejo = []
                for _, r in df_solicitud_exacta.iterrows():
                    c_sol = int(r['cantidad_solicitada'])
                    c_desp = int(r['cantidad_despachada'])
                    items_cotejo.append({
                        "id": r['id'],
                        "Artículo": r['descripcion'],
                        "Cant. Solicitada": c_sol,
                        "Cant. Despachada Real": c_desp,
                        "Diferencia": c_sol - c_desp
                    })
                
                df_editor_base = pd.DataFrame(items_cotejo)
                
                df_editado = st.data_editor(
                    df_editor_base,
                    column_config={
                        "id": None,
                        "Artículo": st.column_config.TextColumn(disabled=True),
                        "Cant. Solicitada": st.column_config.NumberColumn(disabled=True),
                        "Cant. Despachada Real": st.column_config.NumberColumn(min_value=0, step=1),
                        "Diferencia": st.column_config.NumberColumn(disabled=True)
                    },
                    hide_index=True,
                    use_container_width=True,
                    key="editor_cotejo_tab3"
                )
                
                os.makedirs("pedimentos_escaneados", exist_ok=True)
                ruta_pdf_final = os.path.join("pedimentos_escaneados", f"Escaneado_{consecutivo_final}.pdf")
                
                if st.button("💾 Confirmar Recepción y Registrar en Inventario Oficial", type="primary", use_container_width=True):
                    with open(ruta_pdf_final, "wb") as f_out:
                        f_out.write(bytes_escaneados)
                        
                    conn = sqlite3.connect("database.db")
                    cursor = conn.cursor()
                    
                    for _, fila in df_editado.iterrows():
                        fila_id = int(fila['id'])
                        sol_cant = int(fila['Cant. Solicitada'])
                        desp_cant = int(fila['Cant. Despachada Real'])
                        dif_calc = sol_cant - desp_cant
                        est_fila = "ENTREGADO_CON_DIFERENCIAS" if dif_calc != 0 else "ENTREGADO_CONFORME"
                        
                        cursor.execute("""
                            UPDATE solicitudes
                            SET cantidad_despachada = ?,
                                diferencia = ?,
                                retirado_por = ?,
                                fecha_despacho_real = ?,
                                ruta_pdf_escaneado = ?,
                                estado = ?
                            WHERE id = ?
                        """, (desp_cant, dif_calc, persona_retiro, fecha_desp, ruta_pdf_final, est_fila, fila_id))
                        
                    conn.commit()
                    conn.close()
                    st.balloons()
                    st.success(f"🎉 ¡Despacho del Pedimento #{consecutivo_final} registrado con éxito! Se cotejó contra lo pedido y se actualizaron los inventarios.")

    # TAB 4: INFORMES GENERALES
    with tab_informes:
        st.header("📊 Informes Generales y Diagnóstico de Bodega")
        conn = sqlite3.connect("database.db")
        
        df_entregas = pd.read_sql_query("""
            SELECT s.consecutivo AS Pedimento,
                   f.nombre AS Funcionario, 
                   f.espacio_asignado AS Ubicacion, 
                   a.descripcion AS Articulo, 
                   s.cantidad AS Solicitado, 
                   COALESCE(s.cantidad_despachada, s.cantidad) AS Despachado_Real,
                   COALESCE(s.diferencia, 0) AS Faltante,
                   COALESCE(s.retirado_por, 'Pendiente') AS Retirado_Por,
                   COALESCE(s.fecha_despacho_real, s.fecha_aprobacion) AS Fecha_Despacho,
                   s.estado AS Estado,
                   COALESCE(s.ruta_pdf_escaneado, 'Sin archivo') AS Archivo_Escaneado
            FROM solicitudes s 
            JOIN funcionarios f ON s.funcionario_id = f.id 
            JOIN articulos a ON s.articulo_id = a.codigo_grupal
            ORDER BY s.id DESC
        """, conn)
        conn.close()
        
        st.dataframe(df_entregas, use_container_width=True)

    # TAB 5: APROBACIÓN DE USUARIOS
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

    # TAB 6: CONTROL MASTER, USUARIOS Y REINICIO
    with tab_master:
        st.header("⚙️ Control Master del Sistema")
        st.write("Herramientas exclusivas para el usuario Administrador Master.")
        
        st.markdown("---")
        st.subheader("1. Gestión de Usuarios, Contraseñas y Eliminación de Perfiles")
        
        conn = sqlite3.connect("database.db")
        df_todos_usuarios = pd.read_sql_query("SELECT id, nombre, puesto, espacio_asignado, username, email, rol, estado, debe_cambiar_password FROM funcionarios", conn)
        
        total_users = len(df_todos_usuarios)
        activos_count = len(df_todos_usuarios[df_todos_usuarios['estado'] == 'ACTIVO'])
        pendientes_count = len(df_todos_usuarios[df_todos_usuarios['estado'] == 'INACTIVO'])
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Usuarios Registrados", total_users)
        col2.metric("Usuarios Activos", activos_count)
        col3.metric("Pendientes de Aprobar", pendientes_count)
        
        st.dataframe(df_todos_usuarios, use_container_width=True)
        
        st.write("**Editar datos, restablecer contraseñas o eliminar un usuario:**")
        user_sel_id = st.selectbox(
            "Seleccione el Usuario a gestionar:", 
            df_todos_usuarios['id'].tolist(),
            format_func=lambda x: f"ID {x} - {df_todos_usuarios[df_todos_usuarios['id'] == x]['nombre'].values[0]} (@{df_todos_usuarios[df_todos_usuarios['id'] == x]['username'].values[0]})"
        )
        
        user_data = df_todos_usuarios[df_todos_usuarios['id'] == user_sel_id].iloc[0]
        
        with st.form("form_edit_user"):
            e_nombre = st.text_input("Nombre Completo:", value=user_data['nombre'])
            e_username = st.text_input("Nombre de Usuario (username):", value=user_data['username'])
            e_email = st.text_input("Correo Institucional:", value=user_data['email'] if pd.notna(user_data['email']) else "")
            e_rol = st.selectbox("Rol de Sistema:", ["BASICO", "ADMIN"], index=0 if user_data['rol'] == "BASICO" else 1)
            e_pass = st.text_input("Nueva Contraseña (Dejar en blanco para mantener la actual):", type="password")
            e_forzar_cambio = st.checkbox("Forzar cambio de contraseña en su próximo inicio de sesión", value=bool(user_data['debe_cambiar_password']))
            
            if st.form_submit_button("💾 Guardar Cambios de Usuario", type="primary"):
                cursor = conn.cursor()
                flag_cambio = 1 if e_forzar_cambio else 0
                if e_pass.strip():
                    hash_nueva_pass = hashlib.sha256(e_pass.encode()).hexdigest()
                    cursor.execute("""
                        UPDATE funcionarios 
                        SET nombre=?, username=?, email=?, rol=?, password=?, debe_cambiar_password=? 
                        WHERE id=?
                    """, (e_nombre, e_username, e_email, e_rol, hash_nueva_pass, flag_cambio, user_sel_id))
                else:
                    cursor.execute("""
                        UPDATE funcionarios 
                        SET nombre=?, username=?, email=?, rol=?, debe_cambiar_password=? 
                        WHERE id=?
                    """, (e_nombre, e_username, e_email, e_rol, flag_cambio, user_sel_id))
                conn.commit()
                sync_usuarios_excel()
                st.success(f"¡Datos del usuario ID {user_sel_id} actualizados correctamente!")
                st.rerun()

        st.markdown("#### 🗑️ Eliminar Perfil de Usuario")
        confirmar_eliminar = st.checkbox(f"Confirmo que deseo eliminar de forma definitiva al usuario ID {user_sel_id} ({user_data['nombre']})")
        if confirmar_eliminar:
            if st.button("❌ Eliminar Perfil Definitivamente", type="secondary"):
                if user_sel_id == st.session_state['func_id']:
                    st.error("No puede eliminar su propio perfil mientras tiene la sesión activa.")
                else:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM funcionarios WHERE id=?", (user_sel_id,))
                    conn.commit()
                    sync_usuarios_excel()
                    st.success(f"El perfil de {user_data['nombre']} ha sido eliminado del sistema.")
                    st.rerun()
        
        conn.close()

        st.markdown("---")
        st.subheader("2. Ajuste Manual de Consecutivo de Pedimentos")
        
        num_actual = obtener_ultimo_consecutivo_db()
        anio_actual = datetime.now().year
        
        st.info(f"El último pedimento registrado fue el **#{num_actual:03d}-{anio_actual}**. El siguiente será el **#{(num_actual + 1):03d}-{anio_actual}**.")
        
        nuevo_ultimo = st.number_input("Indique el ÚLTIMO número generado:", min_value=0, max_value=999, value=num_actual, step=1)
        
        if st.button("💾 Guardar Consecutivo de Pedimentos", type="primary"):
            fijar_ultimo_consecutivo_db(int(nuevo_ultimo))
            siguiente = int(nuevo_ultimo) + 1
            st.success(f"¡Consecutivo actualizado! El próximo pedimento será el **#{siguiente:03d}-{anio_actual}**.")
            st.rerun()

        st.markdown("---")
        st.subheader("3. Limpieza de Pedimentos y Archivos de Prueba")
        st.write("Elimina la base de datos de solicitudes y **elimina todos los archivos PDF/Excel de la lista de firmas pendientes y escaneos**.")
        
        if st.button("🗑️ Borrar Historial de Solicitudes y PDFs de Prueba"):
            limpiar_historial_y_archivos()
            st.success("¡Se ha limpiado el historial y se borraron todos los archivos PDF pendientes de firma!")
            st.rerun()

        st.markdown("---")
        st.subheader("4. Reinicio Completo del Sistema")
        st.warning("⚠️ **Atención:** Esta opción recreará la base de datos completa a su estado inicial de fábrica.")
        
        confirmar_reset = st.checkbox("Confirmo que deseo reinstalar el sistema por completo")
        if confirmar_reset:
            if st.button("🔥 Reinstalar Base de Datos Completa", type="primary"):
                instalar_sistema()
                limpiar_historial_y_archivos()
                st.success("¡El sistema ha sido reinstalado a su estado de fábrica!")
                st.rerun()
