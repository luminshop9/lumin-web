import os
import json
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

# ==================== AUTENTICACIÓN (SEGURA) ====================
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
    spreadsheet = None
    gc = None

# ==================== FUNCIONES DE UTILIDAD ====================
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
        return jsonify([])
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
        return jsonify([])
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

    # Obtener todas las hojas necesarias
    hoja_inv = get_worksheet('inventario')
    hoja_ventas = get_worksheet('ventas')
    hoja_boletas = get_worksheet('boletas')
    hoja_mov = get_worksheet('movimientos')
    hoja_extras = get_worksheet('extras_y_perdidas')
    hoja_ahorro = get_worksheet('ahorro')

    if any(h is None for h in [hoja_inv, hoja_ventas, hoja_boletas, hoja_mov]):
        return jsonify({"error": "Faltan hojas en el sheet"}), 500

    # Leer inventario actual para obtener precios y stocks
    registros_inv = hoja_inv.get_all_records()
    productos = {r['SKU']: r for r in registros_inv}

    resultados = []
    ganancia_total = 0.0
    now = ahora_iso()
    id_boleta = int(datetime.now().timestamp() * 1000) % 10_000_000

    for item in items:
        sku = item.get('sku')
        cantidad = parse_int(item.get('cantidad', 0))
        precio_real = parse_decimal(item.get('precio', 0))
        if not sku or cantidad <= 0 or precio_real <= 0:
            continue

        prod = productos.get(sku)
        if not prod:
            resultados.append(f"❌ Producto {sku} no encontrado")
            continue

        stock_actual = parse_int(prod.get('Stock_actual', 0))
        if cantidad > stock_actual:
            resultados.append(f"⚠️ Stock insuficiente de {prod['Nombre_completo']}")
            continue

        # 1. Actualizar Stock en inventario
        nuevo_stock = stock_actual - cantidad
        filas = hoja_inv.get_all_values()
        for i, fila in enumerate(filas):
            if fila and fila[0] == sku:
                # Columna Stock_actual es la 10 (índice 10 en 0-based, o columna 11 en 1-based)
                hoja_inv.update_cell(i+1, 11, nuevo_stock)
                break

        # 2. Registrar en Ventas
        costo = parse_decimal(prod.get('Costo', 0))
        ganancia_unidad = precio_real - costo
        ganancia_total += ganancia_unidad * cantidad

        hoja_ventas.append_row([
            now, id_boleta, sku, prod['Nombre_completo'],
            cantidad, precio_real, costo, ganancia_unidad, ganancia_unidad * cantidad, vendedor
        ])

        # 3. Registrar en Movimientos
        hoja_mov.append_row([
            now, sku, 'venta', cantidad, precio_real, cantidad * precio_real, '', vendedor, ''
        ])

        # 4. Registrar en Boletas (detalle de diferencia de precios)
        precio_sugerido = parse_decimal(prod.get('Precio_venta_actual', 0))
        diferencia = precio_real - precio_sugerido
        tipo_diferencia = 'extra' if diferencia > 0 else ('perdida' if diferencia < 0 else 'normal')
        hoja_boletas.append_row([
            id_boleta, now, '', '', sku, prod['Nombre_completo'], cantidad,
            precio_sugerido, precio_real, diferencia, tipo_diferencia, cantidad * precio_real, vendedor
        ])

        # 5. Extras / Perdidas
        if diferencia > 0:
            hoja_extras.append_row([
                now, id_boleta, sku, 'extra', diferencia * cantidad,
                'Venta por encima del precio sugerido', diferencia * cantidad, 0, diferencia * cantidad
            ])
            hoja_ahorro.append_row([
                now, id_boleta, sku, precio_sugerido, precio_real,
                diferencia * cantidad, 'Venta por encima del precio sugerido', diferencia * cantidad
            ])
        elif diferencia < 0:
            hoja_extras.append_row([
                now, id_boleta, sku, 'perdida', abs(diferencia * cantidad),
                'Venta por debajo del precio sugerido', 0, abs(diferencia * cantidad), diferencia * cantidad
            ])

        resultados.append(f"✅ {cantidad} x {prod['Nombre_completo']} - {fmt_money(precio_real)}")

    return jsonify({
        'mensaje': 'Venta registrada exitosamente',
        'detalle': resultados,
        'ganancia_total': ganancia_total,
        'boleta_id': id_boleta
    })

# ==================== INICIO ====================
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)