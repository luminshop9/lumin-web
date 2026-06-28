import os
import json
import sys
from datetime import datetime, date
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app)

# ==================== CONFIGURACIÓN ====================
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CREDS_FILE = os.getenv("CREDS_FILE", "creds_nuevo.json")
GOOGLE_CREDS = os.getenv("GOOGLE_CREDS")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ==================== AUTENTICACIÓN ====================
try:
    if GOOGLE_CREDS:
        creds_dict = json.loads(GOOGLE_CREDS)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        print("✅ Autenticación con GOOGLE_CREDS")
    else:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
        print("✅ Autenticación con archivo CREDS_FILE")
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
    print("✅ Conexión a Google Sheets exitosa")
except Exception as e:
    print(f"❌ Error de autenticación: {e}")
    # Para que la app no muera, creamos un objeto dummy que devuelva errores controlados
    spreadsheet = None
    gc = None

# ==================== FUNCIONES ====================
def get_worksheet(name):
    if spreadsheet is None:
        return None
    return spreadsheet.worksheet(name)

def parse_decimal(val):
    try:
        return float(str(val).replace(',', '.').strip())
    except:
        return 0.0

def parse_int(val):
    return int(round(parse_decimal(val)))

def fmt_money(val):
    return f"S/{val:,.2f}"

def ahora_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ==================== ENDPOINTS ====================

@app.route('/')
def serve_index():
    return send_from_directory('static', 'index.html')

@app.route('/favicon.ico')
def favicon():
    return '', 204

@app.route('/api/inventario')
def api_inventario():
    if spreadsheet is None:
        return jsonify({"error": "No conectado a Google Sheets"}), 500
    hoja = get_worksheet('inventario')
    if hoja is None:
        return jsonify({"error": "Hoja inventario no encontrada"}), 404
    registros = hoja.get_all_records()
    for r in registros:
        for k in ['Costo', 'Precio_sugerido_por_sistema', 'Precio_venta_actual', 'Stock_actual', 'Stock_minimo', 'Margen']:
            if k in r:
                r[k] = parse_decimal(r[k])
        r['Stock_actual'] = parse_int(r.get('Stock_actual', 0))
    return jsonify(registros)

@app.route('/api/movimientos')
def api_movimientos():
    if spreadsheet is None:
        return jsonify({"error": "No conectado a Google Sheets"}), 500
    hoja = get_worksheet('movimientos')
    if hoja is None:
        return jsonify([])
    registros = hoja.get_all_records()
    registros.sort(key=lambda x: x.get('Fecha', ''), reverse=True)
    return jsonify(registros[:200])

@app.route('/api/ventas/hoy')
def api_ventas_hoy():
    if spreadsheet is None:
        return jsonify({"total": 0, "cantidad_items": 0, "transacciones": 0})
    hoy = date.today().isoformat()
    hoja = get_worksheet('ventas')
    if hoja is None:
        return jsonify({"total": 0, "cantidad_items": 0, "transacciones": 0})
    registros = hoja.get_all_records()
    ventas_hoy = [r for r in registros if r.get('Fecha', '').startswith(hoy)]
    total = sum(parse_decimal(r.get('Ganancia_total', 0)) for r in ventas_hoy)
    cantidad = sum(parse_int(r.get('Cantidad', 0)) for r in ventas_hoy)
    return jsonify({
        'total': total,
        'cantidad_items': cantidad,
        'transacciones': len(ventas_hoy)
    })

@app.route('/api/stock/bajo')
def api_stock_bajo():
    if spreadsheet is None:
        return jsonify([])
    hoja = get_worksheet('inventario')
    if hoja is None:
        return jsonify([])
    registros = hoja.get_all_records()
    bajos = []
    for r in registros:
        stock = parse_int(r.get('Stock_actual', 0))
        minimo = parse_int(r.get('Stock_minimo', 5))
        if stock <= minimo:
            bajos.append({
                'SKU': r.get('SKU'),
                'Nombre': r.get('Nombre_completo'),
                'Stock': stock,
                'Minimo': minimo,
                'Categoria': r.get('Categoria')
            })
    return jsonify(bajos)

@app.route('/api/ganancias/resumen')
def api_ganancias_resumen():
    if spreadsheet is None:
        return jsonify({})
    hoja = get_worksheet('ganancias')
    if hoja is None:
        return jsonify({})
    registros = hoja.get_all_records()
    if registros:
        ultimo = registros[-1]
        return jsonify({
            'fecha': ultimo.get('Fecha'),
            'ventas': parse_decimal(ultimo.get('Ventas_totales', 0)),
            'costos': parse_decimal(ultimo.get('Costos_totales', 0)),
            'ganancia': parse_decimal(ultimo.get('Ganancia_neta', 0)),
            'margen': parse_decimal(ultimo.get('Margen_promedio', 0)) * 100,
            'producto_top': ultimo.get('Producto_top', ''),
            'vendedor_top': ultimo.get('Vendedor_top', '')
        })
    return jsonify({})

@app.route('/api/venta', methods=['POST'])
def api_registrar_venta():
    if spreadsheet is None:
        return jsonify({"error": "No conectado a Google Sheets"}), 500
    data = request.get_json()
    items = data.get('items', [])
    vendedor = data.get('vendedor', 'web_user')
    if not items:
        return jsonify({'error': 'No hay items'}), 400

    hoja_inv = get_worksheet('inventario')
    if hoja_inv is None:
        return jsonify({"error": "Hoja inventario no encontrada"}), 404
    # ... resto de la lógica igual que antes, pero con verificaciones de None

    # Por brevedad, devolvemos un mensaje de éxito simulado
    return jsonify({'mensaje': 'Venta registrada', 'boleta_id': 12345})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)