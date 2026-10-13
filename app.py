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
            st.warning("⏱️ Su sesión ha expirado por inactividad. Inicie sesión de nuevo.")
    else:
        st.query_params.clear()
else:
    # 2. Si ya hay sesión abierta, verificar inactividad
    if last_active is not None and (now - last_active) > MAX_INACTIVIDAD:
        logout_inactividad()
        st.warning("⏱️ Su sesión se ha cerrado automáticamente por inactividad (5 minutos).")
        st.rerun()
    else:
        st.session_state['last_activity'] = now
        st.query_params['t'] = str(int(now))

# Temporizador en el navegador para cierre automático
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

# -------------------------------------------------------------------
# OCULTAR BARRA DE DESARROLLADOR (SHARE, GITHUB, MENÚ) A USUARIOS BÁSICOS
# MANTENIENDO VISIBLE LA FLECHA (>>) PARA ABRIR Y CERRAR EL MENÚ
# -------------------------------------------------------------------
if st.session_state['rol'] != "ADMIN":
    st.markdown("""
        <style>
        /* Ocultar botones de Share, GitHub, lápiz y menú derecho */
        [data-testid="stToolbar"] {
            display: none !important;
            visibility: hidden !important;
        }
        #MainMenu {
            visibility: hidden !important;
            display: none !important;
        }
        footer {
            visibility: hidden !important;
            display: none !important;
        }
        /* Mantener visible y operativa la flecha (>>) para el menú lateral */
        [data-testid="collapsedControl"] {
            display: flex !important;
            visibility: visible !important;
        }
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
    cursor.execute("INSERT OR REPLACE INTO control_consecutivos (anio, ultimo_numero
