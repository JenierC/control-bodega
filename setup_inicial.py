import sqlite3
import pandas as pd
import os
import hashlib

BD_PATH = "database.db"
EXCEL_MAESTRO = "Control_Maestro_238_Articulos_DINADECO.xlsx"
EXCEL_PEDIMENTOS = "Pedimentos de Departamento de Servicios Generales.xlk"

def instalar_sistema():
    print("🚀 Iniciando instalación y configuración del Sistema de Bodega DINADECO...")
    
    os.makedirs("pedimentos_generados", exist_ok=True)
    print("  ✓ Carpeta 'pedimentos_generados/' creada.")

    conn = sqlite3.connect(BD_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS articulos (
            codigo_grupal TEXT PRIMARY KEY,
            descripcion TEXT NOT NULL,
            categoria TEXT,
            frecuencia_entrega TEXT,
            dias_vida_util INTEGER
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS control_pisos (
            codigo_articulo TEXT PRIMARY KEY,
            articulo TEXT,
            total_mensual REAL,
            piso_1_tope REAL,
            piso_2_tope REAL,
            piso_3_tope REAL,
            criterio_operativo TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS departamentos (
            nombre TEXT PRIMARY KEY,
            iniciales TEXT,
            jefatura TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS funcionarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            puesto TEXT DEFAULT 'Misceláneo/a',
            espacio_asignado TEXT NOT NULL,
            username TEXT UNIQUE,
            password TEXT,
            rol TEXT DEFAULT 'BASICO',
            estado TEXT DEFAULT 'INACTIVO'
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS solicitudes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            funcionario_id INTEGER,
            articulo_id TEXT,
            cantidad INTEGER,
            es_valido INTEGER,
            motivo_validacion TEXT,
            justificacion TEXT,
            estado TEXT DEFAULT 'PENDIENTE',
            firmado_por TEXT,
            hash_firma TEXT,
            fecha_solicitud DATETIME DEFAULT CURRENT_TIMESTAMP,
            fecha_aprobacion DATETIME,
            FOREIGN KEY (funcionario_id) REFERENCES funcionarios(id),
            FOREIGN KEY (articulo_id) REFERENCES articulos(codigo_grupal)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS control_consecutivos (
            anio INTEGER PRIMARY KEY,
            ultimo_numero INTEGER NOT NULL
        )
    ''')

    cursor.execute("INSERT OR IGNORE INTO control_consecutivos (anio, ultimo_numero) VALUES (2026, 50)")

    # Insertar usuarios administradores (Jenier y Gabriela) con clave inicial 'admin123'
    hash_admin = hashlib.sha256("admin123".encode()).hexdigest()
    cursor.execute("INSERT OR IGNORE INTO funcionarios (nombre, puesto, espacio_asignado, username, password, rol, estado) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                   ("Jenier Camacho Picado", "Supervisor", "Oficina Central", "jenier", hash_admin, "ADMIN", "ACTIVO"))
    cursor.execute("INSERT OR IGNORE INTO funcionarios (nombre, puesto, espacio_asignado, username, password, rol, estado) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                   ("Gabriela Gutiérrez Araya", "Supervisora", "Oficina Central", "gabriela", hash_admin, "ADMIN", "ACTIVO"))

    if os.path.exists(EXCEL_MAESTRO):
        try:
            df_master = pd.read_excel(EXCEL_MAESTRO, sheet_name='Master_238_Articulos')
            mapa_dias = {
                'Mensual': 30, 'Bimensual': 60, 'Trimestral': 90, 
                'Semestral': 180, 'Anual': 365, 'Bienal': 730, 
                'Trienal': 1095, 'Quinquenal': 1825, 'Eventual': 0
            }
            for _, row in df_master.iterrows():
                frec = str(row['Frecuencia Sugerida de Entrega']).strip()
                dias = mapa_dias.get(frec, 30)
                cursor.execute('''
                    INSERT OR REPLACE INTO articulos (codigo_grupal, descripcion, categoria, frecuencia_entrega, dias_vida_util)
                    VALUES (?, ?, ?, ?, ?)
                ''', (str(row['Grupal']), row['Descripción de Artículo'], row['Categoría de Bodega'], frec, dias))
            
            df_pisos = pd.read_excel(EXCEL_MAESTRO, sheet_name='Limpieza_por_Piso')
            for _, row in df_pisos.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO control_pisos 
                    (codigo_articulo, articulo, total_mensual, piso_1_tope, piso_2_tope, piso_3_tope, criterio_operativo)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    str(row['Código']), row['Artículo'], row['Total Mensual'],
                    row['Piso 1 (30%)'], row['Piso 2 (40%)'], row['Piso 3 (30%)'], row['Criterio Operativo']
                ))
            print("  ✓ Catálogo de 238 artículos y cuotas por piso cargados.")
        except Exception as e:
            print("  ⚠ Alerta cargando hojas del Excel Maestro:", e)

    if os.path.exists(EXCEL_PEDIMENTOS):
        try:
            df_deps = pd.read_excel(EXCEL_PEDIMENTOS, sheet_name='DEPARTAMENTOS')
            for _, row in df_deps.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO departamentos (nombre, iniciales, jefatura)
                    VALUES (?, ?, ?)
                ''', (str(row['DEPARTAMENTOS']).strip(), str(row['INICIALES ']).strip(), str(row['JEFATURAS o Autorizados']).strip()))
            print("  ✓ Lista oficial de Departamentos y Jefaturas cargada.")
        except Exception as e:
            print("  ⚠ Alerta cargando hoja DEPARTAMENTOS:", e)

    conn.commit()
    conn.close()
    print("✅ ¡Instalación inicial completada con éxito!")

if __name__ == "__main__":
    instalar_sistema()