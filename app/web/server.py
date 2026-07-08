import os
import uuid
import queue
import threading
import tempfile
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response, send_file
from werkzeug.utils import secure_filename

from app.models.sensor import Sensor
from app.models.proyecto import ProyectoConfig
from app.core.parser_primarios import detectar_sensores, cargar_datos_primarios, obtener_rango_fechas as rango_primarios
from app.core.parser_fallas import cargar_datos_fallas, obtener_rango_fechas as rango_fallas
from app.core.data_cleaner import limpiar_datos, limpiar_primarios, clip_calificacion
from app.core.data_simulator import simular_sensores_faltantes
from app.core.excel_writer import llenar_plantilla

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'calificadoria_uploads')
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), 'calificadoria_output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Plantilla bundleada — hasta nuevo aviso, no se sube manualmente.
_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLANTILLA_PROYECTO = os.path.join(_APP_DIR, 'data', 'plantillaProyecto.xlsx')

_sessions: dict = {}


# ─────────────────────────────────────────────────────────────────
#  PÁGINAS
# ─────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    resp = app.make_response(render_template('index.html'))
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate'
    resp.headers['Pragma'] = 'no-cache'
    return resp


# ─────────────────────────────────────────────────────────────────
#  API — SUBIDA DE ARCHIVOS
# ─────────────────────────────────────────────────────────────────

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió ningún archivo.'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'Solo se aceptan archivos .xlsx'}), 400
    filename = secure_filename(f.filename)
    unique = f"{uuid.uuid4().hex}_{filename}"
    path = os.path.join(UPLOAD_DIR, unique)
    f.save(path)
    return jsonify({'path': path, 'filename': filename})


# ─────────────────────────────────────────────────────────────────
#  API — DETECCIÓN DE SENSORES
# ─────────────────────────────────────────────────────────────────

def _fmt_rango(ts_min, ts_max) -> dict:
    fmt = '%d/%m/%Y %H:%M'
    return {
        'desde': ts_min.strftime(fmt) if ts_min else None,
        'hasta': ts_max.strftime(fmt) if ts_max else None,
    }


@app.route('/api/detect-sensors', methods=['POST'])
def detect_sensors():
    data = request.json
    ruta = data.get('ruta_primarios', '')
    ruta_f = data.get('ruta_fallas', '')
    num = int(data.get('num_sensores', 9))
    try:
        detectados = detectar_sensores(ruta)
        result = [
            {
                'serial': s.serial,
                'nombre': s.nombre,
                'tiene_temperatura': s.tiene_temperatura,
                'tiene_humedad': s.tiene_humedad,
                'posicion': s.posicion,
                'simulado': False,
                'col_idx_temp': s.col_idx_temp,
                'col_idx_hum': s.col_idx_hum,
                'usar': True,
            }
            for s in detectados
        ]
        posiciones = {s['posicion'] for s in result}
        pos = max(posiciones, default=0) + 1
        while len(result) < num:
            while pos in posiciones:
                pos += 1
            result.append({
                'serial': f'SIM-{pos:02d}',
                'nombre': 'RHTemp101A',
                'tiene_temperatura': True,
                'tiene_humedad': True,
                'posicion': pos,
                'simulado': True,
                'col_idx_temp': None,
                'col_idx_hum': None,
                'usar': True,
            })
            posiciones.add(pos)
            pos += 1

        rp_min, rp_max = rango_primarios(ruta)
        rf_min, rf_max = rango_fallas(ruta_f) if ruta_f else (None, None)

        return jsonify({
            'sensores': result,
            'total_archivo': len(detectados),
            'rango_primarios': _fmt_rango(rp_min, rp_max),
            'rango_fallas':    _fmt_rango(rf_min, rf_max),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ─────────────────────────────────────────────────────────────────
#  API — GENERACIÓN
# ─────────────────────────────────────────────────────────────────

@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.json
    sid = uuid.uuid4().hex
    q = queue.Queue()
    _sessions[sid] = {'queue': q, 'output': None, 'done': False}
    threading.Thread(target=_run_pipeline, args=(sid, data), daemon=True).start()
    return jsonify({'session_id': sid})


@app.route('/api/stream/<sid>')
def stream(sid):
    def events():
        if sid not in _sessions:
            yield "data: ERROR: sesión no encontrada\n\n"
            return
        q = _sessions[sid]['queue']
        while True:
            msg = q.get()
            if msg is None:
                ok = 1 if _sessions[sid]['output'] else 0
                yield f"data: __DONE__:{sid}:{ok}\n\n"
                break
            # escape newlines for SSE
            for line in str(msg).splitlines():
                yield f"data: {line}\n\n"
    return Response(events(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.route('/api/download/<sid>')
def download(sid):
    if sid not in _sessions:
        return 'Not found', 404
    output = _sessions[sid].get('output')
    if not output or not os.path.exists(output):
        return 'File not ready', 404
    base = os.path.basename(output)
    # strip the uuid suffix: "ensayo_<hex>.xlsx" → "ensayo.xlsx"
    name_parts = base.rsplit('_', 1)
    download_name = name_parts[0] + '.xlsx' if len(name_parts) == 2 else base
    return send_file(output, as_attachment=True, download_name=download_name)


# ─────────────────────────────────────────────────────────────────
#  PIPELINE
# ─────────────────────────────────────────────────────────────────

def _run_pipeline(sid: str, data: dict):
    q = _sessions[sid]['queue']

    def log(msg: str):
        q.put(msg)

    try:
        log('Validando configuración...')
        config = _build_config(data)

        # Validar que las fechas ingresadas estén dentro del rango de los archivos
        rp_min, rp_max = rango_primarios(config.ruta_primarios)
        if rp_min and rp_max:
            fin_24h = config.inicio_24h + __import__('datetime').timedelta(hours=24)
            if config.inicio_24h < rp_min or config.inicio_24h > rp_max:
                raise ValueError(
                    f"La fecha de inicio '{config.inicio_24h.strftime('%d/%m/%Y %H:%M')}' "
                    f"no está en el archivo de Primarios.\n"
                    f"Rango disponible: {rp_min.strftime('%d/%m/%Y %H:%M')} — {rp_max.strftime('%d/%m/%Y %H:%M')}"
                )

        rf_min, rf_max = rango_fallas(config.ruta_fallas)
        if rf_min and rf_max:
            if config.inicio_falla < rf_min or config.inicio_falla > rf_max:
                raise ValueError(
                    f"La fecha de inicio de falla '{config.inicio_falla.strftime('%d/%m/%Y %H:%M')}' "
                    f"no está en el archivo de Fallas.\n"
                    f"Rango disponible: {rf_min.strftime('%d/%m/%Y %H:%M')} — {rf_max.strftime('%d/%m/%Y %H:%M')}"
                )

        log('Construyendo lista de sensores...')
        sensores = _build_sensores(data.get('sensores', []))

        log(f"Leyendo primarios: {os.path.basename(config.ruta_primarios)}")
        sensores = cargar_datos_primarios(config.ruta_primarios, sensores, config.inicio_24h)
        log(f"  {len(sensores)} sensores cargados")

        log('Limpiando errores de sensor (rango amplio para primarios)...')
        sensores = limpiar_primarios(sensores, config.setpoint_temp, config.setpoint_hum)

        timestamps = None
        for s in sensores:
            if s.datos is not None and len(s.datos) > 0:
                timestamps = s.datos['timestamp']
                break
        if timestamps is None:
            raise ValueError("No se encontraron datos en el rango de 24 horas indicado.")

        for s in sensores:
            if s.tiene_temperatura and not s.tiene_datos_temp():
                log(f"  ⚠ {s.serial}: sin temperatura — se simulará")
            if s.tiene_humedad and not s.tiene_datos_hum():
                log(f"  ⚠ {s.serial}: sin HR — se simulará")

        sensores = simular_sensores_faltantes(
            sensores, timestamps,
            config.rango_temp_min, config.rango_temp_max,
            config.rango_hum_min,  config.rango_hum_max,
            num_sensores=config.num_sensores,
            setpoint_temp=config.setpoint_temp,
            setpoint_hum=config.setpoint_hum,
        )

        log('Aplicando tolerancia de calificación (±2 °C / ±5 %HR)...')
        sensores = clip_calificacion(sensores, config.setpoint_temp, config.setpoint_hum)

        log(f"Leyendo fallas ({config.tipo_prueba}): {os.path.basename(config.ruta_fallas)}")
        df_ft, df_fh = cargar_datos_fallas(config.ruta_fallas, config.tipo_prueba, config.inicio_falla)
        log(f"  {len(df_ft)} registros temp  |  {len(df_fh)} registros HR")

        log('Generando Excel final...')
        nombre = (config.ensayo.strip() or 'reporte').replace('/', '-').replace('\\', '-')
        output_path = os.path.join(OUTPUT_DIR, f"{nombre}_{sid}.xlsx")
        llenar_plantilla(config.ruta_plantilla, output_path,
                         sensores, timestamps, df_ft, df_fh, config)

        _sessions[sid]['output'] = output_path
        log(f"\n✓  REPORTE GENERADO EXITOSAMENTE: {nombre}.xlsx")

    except Exception as e:
        import traceback
        log(f"\n✗  ERROR: {e}")
        log(traceback.format_exc())
    finally:
        _sessions[sid]['done'] = True
        q.put(None)


# ─────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────

def _parsear_fecha_hora(fecha: str, hora: str, ctx: str) -> datetime:
    fecha = fecha.strip().replace('-', '/')
    hora = hora.strip().lower().replace(' ', '').replace('.', '')
    texto = f"{fecha} {hora}"
    for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y %I:%M%p",
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %I:%M:%S%p"]:
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    raise ValueError(f"{ctx}: formato inválido '{fecha} {hora}'. Usa DD/MM/AAAA HH:MM")


def _build_config(data: dict) -> ProyectoConfig:
    c = ProyectoConfig()
    c.ruta_primarios = data['ruta_primarios']
    c.ruta_fallas    = data['ruta_fallas']
    c.ruta_plantilla = data.get('ruta_plantilla') or PLANTILLA_PROYECTO
    c.inicio_24h     = _parsear_fecha_hora(data['ini_fecha'], data['ini_hora'], "Inicio 24h")
    c.inicio_falla   = _parsear_fecha_hora(data['falla_fecha'], data['falla_hora'], "Inicio falla")
    c.tipo_prueba    = data['tipo_prueba']
    c.rango_temp_min = float(data['tmin'])
    c.rango_temp_max = float(data['tmax'])
    c.rango_hum_min  = float(data['hmin'])
    c.rango_hum_max  = float(data['hmax'])
    c.setpoint_temp  = float(data['sp_temp'])
    c.setpoint_hum   = float(data['sp_hum'])
    c.num_sensores   = max(1, min(120, int(data.get('num_sensores', 9))))
    vc = data.get('variable_calificacion', 'ambas')
    c.variable_calificacion = vc if vc in ('temperatura', 'humedad', 'ambas') else 'ambas'
    c.empresa        = data.get('empresa', '')
    c.marca_equipo   = data.get('marca', '')
    c.ubicacion      = data.get('ubicacion', '')
    c.codigo_equipo  = data.get('codigo', '')
    c.ensayo         = data.get('ensayo', '')
    lec_t = data.get('lec_temp', '').strip()
    lec_h = data.get('lec_hum', '').strip()
    c.lectura_equipo_temp = float(lec_t) if lec_t else c.setpoint_temp
    c.lectura_equipo_hum  = float(lec_h) if lec_h else c.setpoint_hum
    return c


def _build_sensores(sensores_data: list) -> list:
    result = []
    for s in sensores_data:
        if not s.get('usar', True):
            continue
        result.append(Sensor(
            serial=s['serial'],
            nombre=s['nombre'],
            descripcion='',
            posicion=int(s['posicion']),
            tiene_temperatura=s['tiene_temperatura'],
            tiene_humedad=s['tiene_humedad'],
            col_idx_temp=s.get('col_idx_temp'),
            col_idx_hum=s.get('col_idx_hum'),
        ))
    return result
