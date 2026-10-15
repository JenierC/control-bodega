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
# GESTIÓN Y PERSISTENCIA DE API KEY GEMINI
# ==========================================
def obtener_gemini_api_key():
    try:
        conn = sqlite3.connect("database.db")
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS config_sistema (clave TEXT PRIMARY KEY, valor TEXT)")
        c.execute("SELECT valor FROM config_sistema WHERE clave = 'gemini_api_key'")
        row = c.fetchone()
        conn.close()
        if row and row[0] and row[0].strip():
            return row[0].strip()
    except Exception:
        pass

    if hasattr(st, "secrets"):
        if "gemini_api_key" in st.secrets and str(st.secrets["gemini_api_key"]).strip():
            return str(st.secrets["gemini_api_key"]).strip()
        if "GEMINI_API_KEY" in st.secrets and str(st.secrets["GEMINI_API_KEY"]).strip():
            return str(st.secrets["GEMINI_API_KEY"]).strip()
        for sec in st.secrets:
            try:
                val = st.secrets[sec]
                if isinstance(val, dict) or hasattr(val, "get"):
                    for k in ["gemini_api_key", "GEMINI_API_KEY", "api_key", "key"]:
                        if k in val and str(val[k]).strip():
                            return str(val[k]).strip()
            except Exception:
                pass

    return os.environ.get("GEMINI_API_KEY", None)

def guardar_gemini_api_key(nueva_key):
    try:
        conn = sqlite3.connect("database.db")
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS config_sistema (clave TEXT PRIMARY KEY, valor TEXT)")
        c.execute("INSERT OR REPLACE INTO config_sistema (clave, valor) VALUES ('gemini_api_key', ?)", (nueva_key.strip(),))
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False

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
# FUNCIÓN DE EXTRACCIÓN ROBUSTA DE PEDIMENTOS (DIGITAL O ESCANEO)
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
    
    # 1. BÚSQUEDA EXACTA EN TABLA: Partida, Subpartida, Código, Subcódigo, Cantidad
    pattern_bys = re.compile(r'(\d{5})\s+(\d{4,5})\s+(\d{5,6})\s+(\d{2})\s+(\d+)')
    matches_bys = pattern_bys.findall(full_text)
    
    if matches_bys:
        try:
            conn_cat = sqlite3.connect("database.db")
            c = conn_cat.cursor()
            for partida, subpartida, codigo, subcodigo, cantidad in matches_bys:
                c.execute("""
                    SELECT codigo_grupal, descripcion 
                    FROM articulos 
                    WHERE codigo_grupal LIKE ? AND codigo_grupal LIKE ?
                """, (f"{partida}%", f"%{codigo}{subcodigo}"))
                row = c.fetchone()
                if not row:
                    c.execute("""
                        SELECT codigo_grupal, descripcion 
                        FROM articulos 
                        WHERE codigo_grupal LIKE ? AND codigo_grupal LIKE ?
                    """, (f"{partida}%", f"%{codigo}%"))
                    row = c.fetchone()
                    
                if row:
                    articulos_encontrados.append({
                        "codigo": row[0],
                        "descripcion": row[1],
                        "cantidad_solicitada": int(cantidad),
                        "cantidad_despachada": int(cantidad)
                    })
            conn_cat.close()
        except Exception:
            pass

    # 2. VISIÓN IA SI ES ESCANEO PLANO
    if not articulos_encontrados:
        img_bytes = None
        for page in reader.pages:
            if page.images:
                img_bytes = page.images[0].data
                break
                
        api_key = obtener_gemini_api_key()
        if api_key and img_bytes:
            try:
                b64_img = base64.b64encode(img_bytes).decode("utf-8")
                prompt_ai = """
                Analiza este pedimento oficial de DINADECO. Extrae en JSON estricto:
                {
                  "consecutivo": "ej: 051-2026",
                  "articulos": [
                    {
                      "codigo": "Código numérico",
                      "descripcion": "Nombre del insumo",
                      "cantidad_solicitada": 4,
                      "cantidad_despachada": 4
                    }
                  ]
                }
                Devuelve ÚNICAMENTE el JSON.
                """
                payload_ai = {
                    "contents": [{
                        "parts": [
                            {"text": prompt_ai},
                            {"inlineData": {"mimeType": "image/jpeg", "data": b64_img}}
                        ]
                    }],
                    "generationConfig": {
                        "temperature": 0.1,
                        "responseMimeType": "application/json"
                    }
                }
                for mod in ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]:
                    try:
                        url_ai = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent?key={api_key}"
                        req_ai = urllib.request.Request(url_ai, data=json.dumps(payload_ai).encode('utf-8'), headers={'Content-Type': 'application/json'})
                        with urllib.request.urlopen(req_ai, timeout=20) as resp_ai:
                            data_ai = json.loads(resp_ai.read().decode('utf-8'))
                            raw_ai = data_ai['candidates'][0]['content']['parts'][0]['text']
                            raw_ai = re.sub(r'```json\s*', '', raw_ai).replace('```', '').strip()
                            parsed_ai = json.loads(raw_ai)
                            if parsed_ai.get("consecutivo"):
                                consecutivo = parsed_ai["consecutivo"]
                            items_ai = parsed_ai.get("articulos") or parsed_ai.get("items") or []
                            for a in items_ai:
                                articulos_encontrados.append({
                                    "codigo": a.get("codigo", ""),
                                    "descripcion": a.get("descripcion", ""),
                                    "cantidad_solicitada": int(a.get("cantidad_solicitada", 1)),
                                    "cantidad_despachada": int(a.get("cantidad_despachada", a.get("cantidad_solicitada", 1)))
                                })
                            break
                    except Exception:
                        continue
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
            n_pass = st.text_input("Contraseña:",
