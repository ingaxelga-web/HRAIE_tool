import io
import sqlite3
import pandas as pd
import unicodedata
import re
from datetime import date
from flask import Flask, render_template, request, jsonify, send_file
from contextlib import closing

app = Flask(__name__)
DB_NAME = "hospital.db"

# ==========================================
# UTILIDADES Y SEGURIDAD
# ==========================================
def remove_accents(input_str):
    if pd.isna(input_str) or input_str is None: return ""
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).upper()

def get_db_connection():
    conexion = sqlite3.connect(DB_NAME, timeout=15)
    conexion.row_factory = sqlite3.Row
    return conexion

def safe_int(value, default=0):
    try: 
        # Tolerar '2.0' convirtiendo a float primero
        return int(float(value)) if value else default
    except (ValueError, TypeError): 
        return default

def clean_col_name(c):
    return re.sub(r'[^A-Z0-9]', '', remove_accents(str(c)))

def normalizar_fecha(fecha_str):
    """Limpia los ceros de tiempo de Pandas si viene como 2026-08-20 00:00:00"""
    fecha = str(fecha_str).strip()
    if fecha.endswith("00:00:00"):
        return fecha.replace("00:00:00", "").strip()
    return fecha

def inicializar_bd():
    with closing(get_db_connection()) as conexion:
        cursor = conexion.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS capturas_diarias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_alta TEXT, codigo TEXT, nota TEXT, destino TEXT,
                expediente TEXT, ap_paterno TEXT, ap_materno TEXT, nombres TEXT, curp TEXT,
                fecha_nacimiento TEXT, sexo TEXT, seguridad_social TEXT,
                intervencion TEXT, diagnostico TEXT, fase TEXT, estatus TEXT,
                fecha_dx TEXT, edad_dx INTEGER, inicio TEXT, termino TEXT, edad_inicio INTEGER,
                clave_cnis TEXT, descripcion_insumo TEXT, cantidad INTEGER, presentacion TEXT,
                fuente_fin TEXT, anio_recepcion TEXT, lote TEXT, orden_suministro TEXT,
                dx_medicos TEXT, intervencion_sadmi TEXT, dx_sadmi TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS capturas_borrador (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_alta TEXT, codigo TEXT, nota TEXT, destino TEXT,
                expediente TEXT, ap_paterno TEXT, ap_materno TEXT, nombres TEXT, curp TEXT,
                fecha_nacimiento TEXT, sexo TEXT, seguridad_social TEXT,
                intervencion TEXT, diagnostico TEXT, fase TEXT, estatus TEXT,
                fecha_dx TEXT, edad_dx INTEGER, inicio TEXT, termino TEXT, edad_inicio INTEGER,
                clave_cnis TEXT, descripcion_insumo TEXT, cantidad INTEGER, presentacion TEXT,
                fuente_fin TEXT, anio_recepcion TEXT, lote TEXT, orden_suministro TEXT,
                dx_medicos TEXT, intervencion_sadmi TEXT, dx_sadmi TEXT
            )
        ''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS catalogo_66_intervenciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            intervencion TEXT, diagnostico_completo TEXT,
            busqueda_intervencion TEXT, busqueda_diagnostico TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS catalogo_cie10 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT, diagnostico_completo TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS catalogo_sadmi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            intervencion TEXT, diagnostico_completo TEXT)''')
        
        # Índices para consultas masivas (Modo Producción)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_cap_curp ON capturas_diarias(curp)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_cap_exp ON capturas_diarias(expediente)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_bor_curp ON capturas_borrador(curp)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_bor_exp ON capturas_borrador(expediente)")
        conexion.commit()
inicializar_bd()

# ==========================================
# VISTAS
# ==========================================
@app.route('/')
def inicio(): return render_template('index.html')

@app.route('/admin')
def admin(): return render_template('admin.html')

@app.route('/borradores')
def vista_borradores():
    with closing(get_db_connection()) as conexion:
        cursor = conexion.cursor()
        cursor.execute("SELECT * FROM capturas_borrador ORDER BY id ASC")
        registros = [dict(f) for f in cursor.fetchall()]
    return render_template('borradores.html', registros=registros)

# ==========================================
# IMPORTAR BORRADORES DESDE EXCEL
# ==========================================
@app.route('/cargar_borradores_excel', methods=['POST'])
def cargar_borradores_excel():
    if 'archivo' not in request.files: return jsonify({"estatus": "error", "mensaje": "Archivo faltante"})
    try:
        archivo = request.files['archivo']
        df_raw = pd.read_excel(archivo, header=None, dtype=str) # Cargar como string para evitar floats en expedientes
        
        header_idx = 0
        mejor_conteo = 0
        palabras_clave = ['CURP', 'EXPEDIENTE', 'DESTINO', 'CNIS', 'DIAGNOSTICO',
                          'INTERVENCION', 'FASE', 'ESTATUS', 'NOMBRE', 'SEXO', 'PATERNO']
        for i, row in df_raw.iterrows():
            row_str = "".join([remove_accents(str(x)).upper() for x in row.values if pd.notna(x)])
            conteo = sum(1 for p in palabras_clave if p in row_str)
            if conteo > mejor_conteo:
                mejor_conteo = conteo
                header_idx = i
        if mejor_conteo < 3: header_idx = 2

        archivo.seek(0)
        df = pd.read_excel(archivo, header=header_idx, dtype=str)
        cols_cleaned = {clean_col_name(c): c for c in df.columns}

        db_to_excel = {
            'fecha_alta': ['FECHAALTA'], 'codigo': ['CODIGO'], 'nota': ['NOTA'], 'destino': ['DESTINO'],
            'expediente': ['EXPEDIENTE'], 'ap_paterno': ['PATERNO'], 'ap_materno': ['MATERNO'],
            'nombres': ['NOMBRE'], 'curp': ['CURP'], 'fecha_nacimiento': ['NACIMIENTO'],
            'sexo': ['SEXO'], 'seguridad_social': ['SEGURIDAD'], 'fase': ['FASE'], 'estatus': ['ESTATUS'],
            'fecha_dx': ['FECHADX', 'FECHADEDX'], 'edad_dx': ['EDADDX'], 'inicio': ['INICIO'],
            'termino': ['TERMINO'], 'edad_inicio': ['EDADINICIO'], 'cantidad': ['CANTIDAD'],
            'presentacion': ['PRESENTA'], 'lote': ['LOTE'], 'orden_suministro': ['ORDEN'],
            'descripcion_insumo': ['DESCRIP'], 'dx_medicos': ['DIAGNOSTICOSMEDICOS', 'DXMEDICOS'],
            'intervencion_sadmi': ['INTERVENCIONSADMI'], 'dx_sadmi': ['DXSADMI'],
            'intervencion': ['INTERVENCION'], 'diagnostico': ['DIAGNOSTICO'],
            'clave_cnis': ['CNIS'], 'fuente_fin': ['FUENTE'], 'anio_recepcion': ['ANODERECEPCION', 'RECEP']
        }

        final_mapping = {}
        for db_col, keywords in db_to_excel.items():
            for keyword in keywords:
                for clean, real_col in cols_cleaned.items():
                    if keyword in clean:
                        if db_col == 'intervencion' and 'SADMI' in clean: continue
                        if db_col == 'diagnostico' and ('SADMI' in clean or 'MEDICOS' in clean): continue
                        final_mapping[db_col] = real_col
                        break
                if db_col in final_mapping: break

        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            insertados = 0
            duplicados = 0
            
            for _, row in df.iterrows():
                col_exp = final_mapping.get('expediente', '')
                col_nom = final_mapping.get('nombres', '')
                if (col_exp in df.columns and pd.isna(row[col_exp])) and (col_nom in df.columns and pd.isna(row[col_nom])):
                    continue
                    
                d = {}
                for db_col in db_to_excel.keys():
                    real_col = final_mapping.get(db_col)
                    if real_col and real_col in df.columns:
                        val = row[real_col]
                        if pd.notna(val) and str(val).lower() != 'nan' and str(val).strip() != '':
                            d[db_col] = normalizar_fecha(val)
                        else:
                            d[db_col] = ""
                    else:
                        d[db_col] = ""

                # Regla de duplicación CURP/Expediente
                curp_import = d.get('curp', '').strip().upper()
                exp_import = d.get('expediente', '').strip().upper()
                identificador = curp_import if curp_import else exp_import
                d['curp'] = identificador
                d['expediente'] = identificador
                
                # Antiduplicados en lote activo
                cursor.execute("""
                    SELECT id FROM capturas_borrador 
                    WHERE expediente = ? AND clave_cnis = ? AND codigo = ?
                """, (identificador, d.get('clave_cnis',''), d.get('codigo','')))
                
                if cursor.fetchone():
                    duplicados += 1
                    continue

                cursor.execute('''
                    INSERT INTO capturas_borrador (
                        fecha_alta, codigo, nota, destino, expediente, ap_paterno, ap_materno, nombres, curp,
                        fecha_nacimiento, sexo, seguridad_social, intervencion, diagnostico, fase, estatus,
                        fecha_dx, edad_dx, inicio, termino, edad_inicio, clave_cnis, descripcion_insumo,
                        cantidad, presentacion, fuente_fin, anio_recepcion, lote, orden_suministro,
                        dx_medicos, intervencion_sadmi, dx_sadmi
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (
                    d['fecha_alta'], d['codigo'], d['nota'], d['destino'], d['expediente'], d['ap_paterno'],
                    d['ap_materno'], d['nombres'], d['curp'], d['fecha_nacimiento'], d['sexo'],
                    d['seguridad_social'], d['intervencion'], d['diagnostico'], d['fase'], d['estatus'],
                    d['fecha_dx'], safe_int(d['edad_dx']), d['inicio'], d['termino'], safe_int(d['edad_inicio']),
                    d['clave_cnis'], d['descripcion_insumo'], safe_int(d['cantidad'], 1), d['presentacion'],
                    d['fuente_fin'], d['anio_recepcion'], d['lote'], d['orden_suministro'], d['dx_medicos'],
                    d['intervencion_sadmi'], d['dx_sadmi']
                ))
                insertados += 1
                
            conexion.commit()
            msg = f"Se insertaron {insertados} registros."
            if duplicados > 0: msg += f" (Se ignoraron {duplicados} ya existentes)"
        return jsonify({"estatus": "exito", "mensaje": msg})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

# ==========================================
# ENDPOINTS RESTAURADOS DE BORRADORES
# ==========================================
@app.route('/obtener_borradores_json', methods=['GET'])
def obtener_borradores_json():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT * FROM capturas_borrador ORDER BY id ASC")
            registros = [dict(f) for f in cursor.fetchall()]
        return jsonify({"estatus": "exito", "datos": registros})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/api/siguiente_borrador', methods=['GET'])
def siguiente_borrador():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT * FROM capturas_borrador ORDER BY id ASC LIMIT 1")
            fila = cursor.fetchone()
            if fila:
                cursor.execute("SELECT COUNT(*) as restantes FROM capturas_borrador")
                restantes = cursor.fetchone()['restantes']
                return jsonify({"estatus": "exito", "datos": dict(fila), "restantes": restantes})
            return jsonify({"estatus": "vacio"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/api/borradores/vaciar', methods=['POST'])
def vaciar_borradores():
    try:
        with closing(get_db_connection()) as conexion:
            conexion.cursor().execute("DELETE FROM capturas_borrador")
            conexion.commit()
        return jsonify({"estatus": "exito"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

# Endpoints requeridos por el Tablero de Borradores (borradores.html)
@app.route('/api/borradores/<int:bid>', methods=['PATCH'])
def actualizar_campo_borrador(bid):
    data = request.json or {}
    campo = data.get('campo')
    valor = data.get('valor', '')
    
    campos_validos = {
        'expediente','curp','nombres','ap_paterno','ap_materno','clave_cnis',
        'intervencion','diagnostico','intervencion_sadmi','dx_sadmi','fase','estatus','fecha_dx'
    }
    if campo not in campos_validos:
        return jsonify({"estatus": "error", "mensaje": "Campo no permitido"})
    
    try:
        with closing(get_db_connection()) as conexion:
            conexion.cursor().execute(f"UPDATE capturas_borrador SET {campo} = ? WHERE id = ?", (valor, bid))
            conexion.commit()
        return jsonify({"estatus": "exito"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/api/borradores/<int:bid>/lote', methods=['PATCH'])
def actualizar_lote_borrador(bid):
    data = request.json or {}
    permitidos = ['nombres','ap_paterno','ap_materno','expediente','fecha_nacimiento','sexo']
    sets, vals = [], []
    for k in permitidos:
        if k in data:
            sets.append(f"{k} = ?"); vals.append(data[k])
    
    if not sets:
        return jsonify({"estatus": "exito"})
        
    vals.append(bid)
    try:
        with closing(get_db_connection()) as conexion:
            conexion.cursor().execute(f"UPDATE capturas_borrador SET {', '.join(sets)} WHERE id = ?", tuple(vals))
            conexion.commit()
        return jsonify({"estatus": "exito"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/api/borradores/<int:bid>/validar', methods=['POST'])
def validar_borrador(bid):
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT * FROM capturas_borrador WHERE id = ?", (bid,))
            r = cursor.fetchone()
            if not r: return jsonify({"estatus": "error", "mensaje": "Registro no encontrado"})
            r = dict(r)

            exp = (r.get('expediente') or '').strip()
            nom = (r.get('nombres') or '').strip()
            if not exp or not nom:
                faltan = []
                if not exp: faltan.append('Expediente/CURP')
                if not nom: faltan.append('Nombre(s)')
                return jsonify({"estatus": "error", "mensaje": "Falta: " + ", ".join(faltan)})

            cursor.execute('''
                INSERT INTO capturas_diarias (
                    fecha_alta, codigo, nota, destino, expediente, ap_paterno, ap_materno, nombres, curp,
                    fecha_nacimiento, sexo, seguridad_social, intervencion, diagnostico, fase, estatus,
                    fecha_dx, edad_dx, inicio, termino, edad_inicio, clave_cnis, descripcion_insumo,
                    cantidad, presentacion, fuente_fin, anio_recepcion, lote, orden_suministro,
                    dx_medicos, intervencion_sadmi, dx_sadmi
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', (
                r['fecha_alta'], r['codigo'], r['nota'], r['destino'], r['expediente'], r['ap_paterno'],
                r['ap_materno'], r['nombres'], r['curp'], r['fecha_nacimiento'], r['sexo'],
                r['seguridad_social'], r['intervencion'], r['diagnostico'], r['fase'], r['estatus'],
                r['fecha_dx'], r['edad_dx'], r['inicio'], r['termino'], r['edad_inicio'], r['clave_cnis'],
                r['descripcion_insumo'], r['cantidad'], r['presentacion'], r['fuente_fin'],
                r['anio_recepcion'], r['lote'], r['orden_suministro'], r['dx_medicos'],
                r['intervencion_sadmi'], r['dx_sadmi']
            ))
            cursor.execute("DELETE FROM capturas_borrador WHERE id = ?", (bid,))
            conexion.commit()
        return jsonify({"estatus": "exito"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

# ==========================================
# PACIENTES E HISTORIAL
# ==========================================
@app.route('/obtener_pacientes_historial', methods=['GET'])
def obtener_pacientes_historial():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute('''
                SELECT DISTINCT curp, expediente, nombres, ap_paterno, ap_materno
                FROM capturas_diarias
                WHERE curp != '' OR expediente != ''
            ''')
            registros = cursor.fetchall()
            nombres_completos, curps, expedientes = set(), set(), set()
            for r in registros:
                if r['curp']: curps.add(r['curp'])
                if r['expediente']: expedientes.add(r['expediente'])
                nombre = f"{r['ap_paterno'] or ''} {r['ap_materno'] or ''} {r['nombres'] or ''}".strip().upper()
                if nombre: nombres_completos.add(nombre)
        return jsonify({
            "estatus": "exito",
            "nombres": sorted(nombres_completos),
            "curps": sorted(curps),
            "expedientes": sorted(expedientes)
        })
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/buscar_paciente_por_nombre', methods=['POST'])
def buscar_paciente_por_nombre():
    nombre_completo = (request.json or {}).get('nombre', '').strip().upper()
    if not nombre_completo: return jsonify({"estatus": "error"})
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("""
                SELECT * FROM capturas_diarias
                WHERE (UPPER(TRIM(ap_paterno)) || ' ' || UPPER(TRIM(ap_materno)) || ' ' || UPPER(TRIM(nombres))) LIKE ?
                ORDER BY id DESC LIMIT 1
            """, (f"%{nombre_completo}%",))
            paciente = cursor.fetchone()
            if paciente: return jsonify({"estatus": "exito", "datos": dict(paciente)})
            return jsonify({"estatus": "no_encontrado"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/buscar_paciente/<identificador>', methods=['GET'])
def buscar_paciente(identificador):
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            id_limpio = identificador.strip().upper()
            cursor.execute("SELECT * FROM capturas_diarias WHERE curp = ? OR expediente = ? ORDER BY id DESC LIMIT 1",
                           (id_limpio, id_limpio))
            paciente_local = cursor.fetchone()
            if paciente_local:
                return jsonify({"estatus": "exito", "fuente": "Capturas Anteriores", "datos": dict(paciente_local)})
        return jsonify({"estatus": "no_encontrado"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

# ==========================================
# CATÁLOGOS (lectura y carga admin)
# ==========================================
@app.route('/sugerir_cie10', methods=['POST'])
def sugerir_cie10():
    texto_doctor = (request.json or {}).get('texto', '')
    if not texto_doctor: return jsonify({"opciones": [], "intervencion": "", "diagnostico_exacto": ""})
    texto_limpio = remove_accents(texto_doctor.replace("_", " "))
    palabras = [p for p in texto_limpio.split() if len(p) > 2]
    if not palabras: return jsonify({"opciones": [], "intervencion": "", "diagnostico_exacto": ""})
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            condiciones = " AND ".join(["(busqueda_intervencion LIKE ? OR busqueda_diagnostico LIKE ?)"] * len(palabras))
            parametros = []
            for p in palabras: parametros.extend([f"%{p}%", f"%{p}%"])
            cursor.execute(f"SELECT diagnostico_completo, intervencion FROM catalogo_66_intervenciones WHERE {condiciones}", parametros)
            rows = cursor.fetchall()
            if not rows: return jsonify({"opciones": [], "intervencion": "", "diagnostico_exacto": ""})
            intervenciones = list({r['intervencion'] for r in rows})
            diagnosticos = [r['diagnostico_completo'] for r in rows]
            if len(rows) == 1:
                return jsonify({"intervencion": intervenciones[0], "diagnostico_exacto": diagnosticos[0], "opciones": diagnosticos})
            elif len(intervenciones) == 1:
                return jsonify({"intervencion": intervenciones[0], "diagnostico_exacto": "", "opciones": diagnosticos})
            return jsonify({"intervencion": "", "diagnostico_exacto": "", "opciones": diagnosticos})
    except Exception: return jsonify({"opciones": [], "intervencion": "", "diagnostico_exacto": ""})

@app.route('/obtener_diagnosticos', methods=['GET'])
def obtener_diagnosticos():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT diagnostico_completo FROM catalogo_66_intervenciones")
            return jsonify({"estatus": "exito", "datos": [f['diagnostico_completo'] for f in cursor.fetchall()]})
    except Exception: return jsonify({"estatus": "error", "datos": []})

@app.route('/obtener_intervenciones_gc', methods=['GET'])
def obtener_intervenciones_gc():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT DISTINCT intervencion FROM catalogo_66_intervenciones ORDER BY intervencion")
            return jsonify({"estatus": "exito", "datos": [f['intervencion'] for f in cursor.fetchall()]})
    except Exception: return jsonify({"estatus": "error", "datos": []})

@app.route('/obtener_categorias_cie10', methods=['GET'])
def obtener_categorias_cie10():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT DISTINCT categoria FROM catalogo_cie10 ORDER BY categoria")
            return jsonify({"estatus": "exito", "datos": [f['categoria'] for f in cursor.fetchall()]})
    except Exception: return jsonify({"estatus": "error", "datos": []})

@app.route('/obtener_subcategorias_cie10', methods=['POST'])
def obtener_subcategorias_cie10():
    categoria = (request.json or {}).get('categoria', '')
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT diagnostico_completo FROM catalogo_cie10 WHERE categoria = ?", (categoria,))
            return jsonify({"estatus": "exito", "datos": [f['diagnostico_completo'] for f in cursor.fetchall()]})
    except Exception: return jsonify({"estatus": "error", "datos": []})

@app.route('/obtener_catalogo_sadmi', methods=['GET'])
def obtener_catalogo_sadmi():
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT DISTINCT intervencion FROM catalogo_sadmi ORDER BY intervencion")
            return jsonify({"estatus": "exito", "datos": [f['intervencion'] for f in cursor.fetchall()]})
    except Exception: return jsonify({"estatus": "error", "datos": []})

@app.route('/obtener_subcategorias_sadmi', methods=['POST'])
def obtener_subcategorias_sadmi():
    intervencion = (request.json or {}).get('intervencion', '')
    try:
        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()
            cursor.execute("SELECT diagnostico_completo FROM catalogo_sadmi WHERE intervencion = ? AND diagnostico_completo IS NOT NULL AND diagnostico_completo != ''", (intervencion,))
            return jsonify({"estatus": "exito", "datos": [f['diagnostico_completo'] for f in cursor.fetchall() if f['diagnostico_completo']]})
    except Exception: return jsonify({"estatus": "error", "datos": []})

def _leer_excel_flexible(archivo, col_objetivo_keywords):
    df = pd.read_excel(archivo, dtype=str)
    cols_clean = {clean_col_name(c): c for c in df.columns}
    encontrados = {}
    for key, keywords in col_objetivo_keywords.items():
        for kw in keywords:
            for clean, real in cols_clean.items():
                if kw in clean:
                    encontrados[key] = real
                    break
            if key in encontrados: break
    return df, encontrados

@app.route('/cargar_66_intervenciones', methods=['POST'])
def cargar_66_intervenciones():
    if 'archivo' not in request.files: return jsonify({"estatus": "error", "mensaje": "Archivo faltante"})
    try:
        df = pd.read_excel(request.files['archivo'], sheet_name=0, header=3, dtype=str)
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        for col in ['Grupo', 'Intervención', 'Categoría', 'Clave PT']:
            if col in df.columns: df[col] = df[col].ffill()
            
        df = df.dropna(subset=['Clave CIE -10'])
        df['Clave CIE -10'] = df['Clave CIE -10'].astype(str).str.strip().str.replace("_", " ")
        df['Descripción diagnóstico CIE-10'] = df['Descripción diagnóstico CIE-10'].astype(str).str.strip().str.replace("_", " ")
        df['diagnostico_completo'] = df['Clave CIE -10'] + " " + df['Descripción diagnóstico CIE-10']
        df['Intervención'] = df['Intervención'].astype(str).str.replace("_", " ")
        
        df_limpio = df[['diagnostico_completo', 'Intervención', 'Grupo', 'Clave CIE -10']].rename(columns={'Intervención': 'intervencion'})
        df_limpio['busqueda_intervencion'] = df_limpio['intervencion'].apply(remove_accents)
        df_limpio['busqueda_diagnostico'] = df_limpio['diagnostico_completo'].apply(remove_accents)
        
        with closing(get_db_connection()) as conexion:
            df_limpio.to_sql('catalogo_66_intervenciones', conexion, if_exists='replace', index=False)
            
        return jsonify({"estatus": "exito", "mensaje": f"Se cargaron {len(df_limpio)} diagnósticos de GC."})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

@app.route('/cargar_catalogo_cie10', methods=['POST'])
def cargar_catalogo_cie10():
    if 'archivo' not in request.files: return jsonify({"estatus": "error", "mensaje": "Archivo faltante"})
    try:
        df = pd.read_excel(request.files['archivo'], header=None, dtype=str)
        df['categoria'] = df[2].ffill().astype(str).str.upper().str.replace("_", " ")
        df['diagnostico_completo'] = df[0].astype(str).str.upper().str.replace("_", " ")
        df_catalogo = df[['categoria', 'diagnostico_completo']].dropna()
        df_catalogo = df_catalogo[df_catalogo['diagnostico_completo'] != 'NAN']
        
        with closing(get_db_connection()) as conexion:
            df_catalogo.to_sql('catalogo_cie10', conexion, if_exists='replace', index=False)
        return jsonify({"estatus": "exito", "mensaje": f"CIE-10 estructurado por categorías ({len(df_catalogo)} registros)."})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": f"Error procesando Excel: {str(e)}"})

@app.route('/cargar_catalogo_sadmi', methods=['POST'])
def cargar_catalogo_sadmi():
    if 'archivo' not in request.files: return jsonify({"estatus": "error", "mensaje": "Archivo faltante"})
    try:
        df = pd.read_excel(request.files['archivo'], dtype=str)
        df.dropna(how='all', inplace=True)
        if 'Intervención SADMI' in df.columns and 'Subclasificación CIE-10' in df.columns:
            df_limpio = df[['Intervención SADMI', 'Subclasificación CIE-10']].rename(
                columns={'Intervención SADMI': 'intervencion', 'Subclasificación CIE-10': 'diagnostico_completo'}
            )
        else:
            return jsonify({"estatus": "error", "mensaje": "Las columnas deben llamarse 'Intervención SADMI' y 'Subclasificación CIE-10'."})
            
        df_limpio['intervencion'] = df_limpio['intervencion'].astype(str).str.strip().str.replace("_", " ").str.upper()
        df_limpio['diagnostico_completo'] = df_limpio['diagnostico_completo'].astype(str).str.strip().str.replace("_", " ").str.upper()
        df_limpio.loc[df_limpio['diagnostico_completo'].isin(['NAN', 'NONE', '']), 'diagnostico_completo'] = None
        
        with closing(get_db_connection()) as conexion:
            df_limpio.to_sql('catalogo_sadmi', conexion, if_exists='replace', index=False)
        return jsonify({"estatus": "exito", "mensaje": f"Catálogo SADMI cargado ({len(df_limpio)} relaciones)."})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

# ==========================================
# GUARDADO DE CAPTURAS DESDE INDEX
# ==========================================
@app.route('/guardar', methods=['POST'])
def guardar():
    datos = request.json or {}
    borrador_id = datos.get('borrador_id')
    try:
        if not datos.get('fecha_alta'): datos['fecha_alta'] = date.today().isoformat()
        if not datos.get('anio_recepcion'): datos['anio_recepcion'] = str(date.today().year)

        curp_input = datos.get('curp', '').strip().upper()
        exp_input = datos.get('expediente', '').strip().upper()
        identificador_final = curp_input if curp_input else exp_input

        with closing(get_db_connection()) as conexion:
            cursor = conexion.cursor()

            orig_curp = ''
            orig_exp = ''
            if borrador_id:
                cursor.execute("SELECT curp, expediente FROM capturas_borrador WHERE id = ?", (borrador_id,))
                orig = cursor.fetchone()
                if orig:
                    orig_curp = orig['curp'] if orig['curp'] else ''
                    orig_exp = orig['expediente'] if orig['expediente'] else ''
                cursor.execute("DELETE FROM capturas_borrador WHERE id = ?", (borrador_id,))

            cursor.execute('''
                INSERT INTO capturas_diarias (
                    fecha_alta, codigo, nota, destino, expediente, ap_paterno, ap_materno, nombres, curp,
                    fecha_nacimiento, sexo, seguridad_social, intervencion, diagnostico, fase, estatus,
                    fecha_dx, edad_dx, inicio, termino, edad_inicio, clave_cnis, descripcion_insumo,
                    cantidad, presentacion, fuente_fin, anio_recepcion, lote, orden_suministro,
                    dx_medicos, intervencion_sadmi, dx_sadmi
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', (
                datos.get('fecha_alta',''), datos.get('codigo',''), datos.get('nota',''), datos.get('destino',''),
                identificador_final, datos.get('ap_paterno',''), datos.get('ap_materno',''),
                datos.get('nombres',''), identificador_final,
                datos.get('fecha_nacimiento',''), datos.get('sexo',''), datos.get('seguridad_social',''),
                datos.get('intervencion',''), datos.get('diagnostico',''), datos.get('fase',''), datos.get('estatus',''),
                datos.get('fecha_dx',''), safe_int(datos.get('edad_dx')), datos.get('inicio',''), datos.get('termino',''),
                safe_int(datos.get('edad_inicio')), datos.get('clave_cnis',''), datos.get('descripcion_insumo',''),
                safe_int(datos.get('cantidad'), 1), datos.get('presentacion',''),
                datos.get('fuente_fin',''), datos.get('anio_recepcion',''), datos.get('lote',''),
                datos.get('orden_suministro',''), datos.get('dx_medicos',''),
                datos.get('intervencion_sadmi',''), datos.get('dx_sadmi','')
            ))

            if orig_curp or orig_exp or identificador_final:
                condiciones = []
                parametros_where = []

                if orig_curp:
                    condiciones.append("curp = ?"); parametros_where.append(orig_curp)
                if orig_exp:
                    condiciones.append("expediente = ?"); parametros_where.append(orig_exp)
                if identificador_final and identificador_final not in [orig_curp, orig_exp]:
                    condiciones.append("curp = ?"); parametros_where.append(identificador_final)
                    condiciones.append("expediente = ?"); parametros_where.append(identificador_final)

                if condiciones:
                    where_clause = " OR ".join(condiciones)
                    cursor.execute(f"SELECT * FROM capturas_borrador WHERE {where_clause}", tuple(parametros_where))
                    hermanos = cursor.fetchall()
                    
                    for hermano in hermanos:
                        h = dict(hermano)
                        cursor.execute('''
                            INSERT INTO capturas_diarias (
                                fecha_alta, codigo, nota, destino, expediente, ap_paterno, ap_materno, nombres, curp,
                                fecha_nacimiento, sexo, seguridad_social, intervencion, diagnostico, fase, estatus,
                                fecha_dx, edad_dx, inicio, termino, edad_inicio, clave_cnis, descripcion_insumo,
                                cantidad, presentacion, fuente_fin, anio_recepcion, lote, orden_suministro,
                                dx_medicos, intervencion_sadmi, dx_sadmi
                            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ''', (
                            h['fecha_alta'] if h['fecha_alta'] else datos.get('fecha_alta',''),
                            h['codigo'] if h['codigo'] else datos.get('codigo',''),
                            h['nota'] if h['nota'] else datos.get('nota',''),
                            h['destino'] if h['destino'] else datos.get('destino',''),
                            identificador_final, datos.get('ap_paterno',''), datos.get('ap_materno',''),
                            datos.get('nombres',''), identificador_final,
                            datos.get('fecha_nacimiento',''), datos.get('sexo',''), datos.get('seguridad_social',''),
                            datos.get('intervencion',''), datos.get('diagnostico',''), datos.get('fase',''), datos.get('estatus',''),
                            datos.get('fecha_dx',''), safe_int(datos.get('edad_dx')), datos.get('inicio',''), datos.get('termino',''),
                            safe_int(datos.get('edad_inicio')), 
                            h['clave_cnis'], h['descripcion_insumo'], safe_int(h['cantidad'], 1), h['presentacion'],
                            h['fuente_fin'] if h['fuente_fin'] else datos.get('fuente_fin',''), 
                            h['anio_recepcion'] if h['anio_recepcion'] else datos.get('anio_recepcion',''),
                            h['lote'], h['orden_suministro'], 
                            datos.get('dx_medicos',''), datos.get('intervencion_sadmi',''), datos.get('dx_sadmi','')
                        ))
                    if hermanos:
                        cursor.execute(f"DELETE FROM capturas_borrador WHERE {where_clause}", tuple(parametros_where))
            conexion.commit()
        return jsonify({"estatus": "exito"})
    except Exception as e: return jsonify({"estatus": "error", "mensaje": str(e)})

# ==========================================
# EXPORTACIÓN EXCEL
# ==========================================
@app.route('/exportar_excel', methods=['GET'])
def exportar_excel():
    try:
        with closing(get_db_connection()) as conexion:
            df = pd.read_sql_query("SELECT * FROM capturas_diarias ORDER BY id ASC", conexion)
        if df.empty: return "No hay datos capturados.", 404

        df.insert(0, 'No_secuencial', '###')
        
        columnas = {
            'No_secuencial': 'No.', 
            'codigo': 'Codigo', 
            'nota': 'Nota', 
            'fecha_alta': 'Fecha alta',
            'destino': 'Destino', 
            'expediente': 'No. expediente', 
            'ap_paterno': 'Apellido paterno',
            'ap_materno': 'Apellido materno', 
            'nombres': 'Nombres (s)', 
            'curp': 'Curp',
            'intervencion': 'INTERVENCION',
            'diagnostico': 'Diagnostico', 
            'fecha_dx': 'Fecha dx', 
            'clave_cnis': 'Clave de CNIS (Compendio Nacional de Insumos para la Salud) \u200b',
            'descripcion_insumo': 'Descripcion', 
            'cantidad': 'Cantidad', 
            'presentacion': 'Presentacion',
            'fuente_fin': 'Fuente de Financiamiento\u200b', 
            'lote': 'Lote', 
            'orden_suministro': 'Orden de suministro', 
            'dx_medicos': 'DIAGNOSTICOS MEDICOS ',
            'intervencion_sadmi': 'INTERVENCION SADMI', 
            'dx_sadmi': 'DX SADMI'
        }
        
        if 'id' in df.columns: df = df.drop(columns=['id'])
        orden_final = [k for k in columnas.keys() if k in df.columns]
        df = df[orden_final].rename(columns=columnas)

        COLORES_BLOQUE = {
            'control':     'FFF9C4', 'paciente':    'BBDEFB',
            'diagnostico': 'E1BEE7', 'insumo':      'C8E6C9',
        }
        
        BLOQUES = {
            'control':     ['No.', 'Codigo', 'Nota', 'Fecha alta', 'Destino'],
            'paciente':    ['No. expediente', 'Apellido paterno', 'Apellido materno', 'Nombres (s)', 'Curp'],
            'diagnostico': ['INTERVENCION', 'Diagnostico', 'Fecha dx', 'DIAGNOSTICOS MEDICOS ', 'INTERVENCION SADMI', 'DX SADMI'],
            'insumo':      ['Clave de CNIS (Compendio Nacional de Insumos para la Salud) \u200b', 'Descripcion', 'Cantidad', 'Presentacion', 'Fuente de Financiamiento\u200b', 'Lote', 'Orden de suministro'],
        }
        
        color_por_columna = {}
        for bloque, cols in BLOQUES.items():
            for c in cols:
                color_por_columna[c] = COLORES_BLOQUE[bloque]

        meses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
        hoy = date.today()
        nombre_hoja = f"{meses[hoy.month - 1]} {hoy.year}"

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=nombre_hoja, startrow=3)
            hoja = writer.sheets[nombre_hoja]

            from openpyxl.styles import Font, PatternFill, Alignment
            for celda in hoja[4]:
                nombre = celda.value
                color = color_por_columna.get(nombre, 'FFFFFF')
                celda.font = Font(bold=True, color='000000')
                celda.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                celda.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            for col_cells in hoja.columns:
                longitud_max = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
                letra_col = col_cells[0].column_letter
                hoja.column_dimensions[letra_col].width = min(max(longitud_max + 2, 10), 45)
            
            hoja.freeze_panes = "A5"

        output.seek(0)
        return send_file(output, download_name="Reporte_Capturas.xlsx", as_attachment=True)
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    # Producción (debug en falso para evitar intrusiones en la red)
    app.run(host="0.0.0.0", port=5000, debug=False)