import os
import json
import re
import random
from datetime import datetime, date, timedelta
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app)

# ============================================================
# CONFIGURACIÓN
# ============================================================
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CREDS_FILE = os.getenv("CREDS_FILE", "creds_nuevo.json")
GOOGLE_CREDS = os.getenv("GOOGLE_CREDS")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ============================================================
# AUTENTICACIÓN
# ============================================================
gc = None
spreadsheet = None
autenticado = False

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
    autenticado = True
    print("✅ Conexión a Google Sheets exitosa")
except Exception as e:
    print(f"❌ Error de autenticación: {e}")
    autenticado = False

# ============================================================
# FUNCIONES DE UTILIDAD
# ============================================================
def get_worksheet(name):
    if spreadsheet is None:
        return None
    return spreadsheet.worksheet(name)

def parse_decimal(val):
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r'[^\d,.]', '', s)
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

def parse_int(val):
    return int(round(parse_decimal(val)))

def fmt_money(val):
    return f"S/{val:,.2f}"

def ahora_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def generar_sku(nombre):
    base = re.sub(r"[^A-Za-z0-9]", "", nombre.upper())[:5] or "PROD"
    sufijo = datetime.now().strftime("%d%H%M") + str(random.randint(10, 99))
    return f"{base}-{sufijo}"

def obtener_margenes():
    try:
        hoja = get_worksheet('configuracion_margenes')
        if hoja is None:
            return {}
        registros = hoja.get_all_records()
        margenes = {}
        for r in registros:
            cat = r.get('Categoria')
            if cat:
                margenes[cat] = {
                    'min': parse_decimal(r.get('Margen_minimo', 0.3)),
                    'sugerido': parse_decimal(r.get('Margen_sugerido', 0.5)),
                    'max': parse_decimal(r.get('Margen_maximo', 0.8))
                }
        return margenes
    except Exception as e:
        print("Error leyendo configuracion_margenes:", e)
        return {}

def sugerir_precio(costo, categoria):
    if costo <= 0:
        return None
    margenes = obtener_margenes()
    margen = margenes.get(categoria, {}).get('sugerido', 0.5)
    return round(costo * (1 + margen), 2)

def getToday():
    return date.today().isoformat()

# ============================================================
# CONFIGURACIÓN DE LA TIENDA (leer/escribir en hoja configuracion)
# ============================================================
def leer_configuracion():
    """Lee la hoja 'configuracion' y devuelve un dict con clave:valor"""
    try:
        hoja = get_worksheet('configuracion')
        if hoja is None:
            return {}
        registros = hoja.get_all_records()
        config = {}
        for r in registros:
            clave = r.get('clave', '').strip()
            valor = r.get('valor', '').strip()
            if clave:
                config[clave] = valor
        return config
    except Exception as e:
        print("Error leyendo configuracion:", e)
        return {}

def guardar_configuracion(config_dict):
    """Guarda o actualiza la hoja 'configuracion' con los valores del dict"""
    try:
        hoja = get_worksheet('configuracion')
        if hoja is None:
            return False
        # Obtener filas existentes
        filas = hoja.get_all_values()
        # Si no hay filas, crear encabezados
        if not filas:
            hoja.append_row(["clave", "valor"])
            filas = [["clave", "valor"]]
        # Actualizar cada clave
        for clave, valor in config_dict.items():
            encontrado = False
            for i, fila in enumerate(filas):
                if fila and len(fila) > 0 and fila[0] == clave:
                    hoja.update_cell(i+1, 2, valor)
                    encontrado = True
                    break
            if not encontrado:
                hoja.append_row([clave, valor])
        return True
    except Exception as e:
        print("Error guardando configuracion:", e)
        return False

# ============================================================
# ENDPOINTS
# ============================================================
@app.route('/')
def serve_index():
    return send_from_directory('static', 'index.html')

@app.route('/favicon.ico')
def favicon():
    return '', 204

@app.route('/api/configuracion', methods=['GET'])
def api_get_configuracion():
    config = leer_configuracion()
    return jsonify(config)

@app.route('/api/configuracion', methods=['POST'])
def api_set_configuracion():
    data = request.get_json()
    if not data:
        return jsonify({"error": "Datos inválidos"}), 400
    if guardar_configuracion(data):
        return jsonify({"mensaje": "Configuración guardada"})
    else:
        return jsonify({"error": "Error al guardar configuración"}), 500

@app.route('/api/inventario')
def api_inventario():
    if not autenticado or spreadsheet is None:
        return jsonify([])
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
    if not autenticado or spreadsheet is None:
        return jsonify([])
    hoja = get_worksheet('movimientos')
    if hoja is None:
        return jsonify([])
    registros = hoja.get_all_records()
    registros.sort(key=lambda x: x.get('Fecha', ''), reverse=True)
    return jsonify(registros)

@app.route('/api/ventas/hoy')
def api_ventas_hoy():
    if not autenticado or spreadsheet is None:
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
    if not autenticado or spreadsheet is None:
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
    if not autenticado or spreadsheet is None:
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

@app.route('/api/sugerir-precio', methods=['POST'])
def api_sugerir_precio():
    data = request.get_json()
    costo = parse_decimal(data.get('costo', 0))
    categoria = data.get('categoria', 'General')
    if costo <= 0:
        return jsonify({"error": "Costo inválido"}), 400
    precio = sugerir_precio(costo, categoria)
    return jsonify({"precio_sugerido": precio})

@app.route('/api/producto', methods=['POST'])
def api_crear_producto():
    if not autenticado or spreadsheet is None:
        return jsonify({"error": "No conectado a Google Sheets"}), 500
    data = request.get_json()
    sku = data.get('sku', '').strip()
    nombre = data.get('nombre', '').strip()
    categoria = data.get('categoria', 'General')
    costo = parse_decimal(data.get('costo', 0))
    stock = parse_int(data.get('stock', 0))
    proveedor = data.get('proveedor', '')
    precio_venta = parse_decimal(data.get('precio_venta', 0))

    if not nombre:
        return jsonify({"error": "Nombre es obligatorio"}), 400
    if costo <= 0:
        return jsonify({"error": "Costo debe ser mayor a 0"}), 400

    hoja_inv = get_worksheet('inventario')
    if hoja_inv is None:
        return jsonify({"error": "Hoja inventario no encontrada"}), 404
    registros = hoja_inv.get_all_records()
    if not sku:
        sku = generar_sku(nombre)
    elif any(r.get('SKU') == sku for r in registros):
        return jsonify({"error": "SKU ya existe"}), 400

    if precio_venta <= 0:
        precio_venta = sugerir_precio(costo, categoria)
        if precio_venta is None:
            precio_venta = round(costo * 1.6, 2)

    margen = round((precio_venta - costo) / costo, 6) if costo else 0
    estado = "OK" if stock > 5 else "Bajo"
    ahora = ahora_iso()

    fila = [
        sku, nombre, "estandar", nombre, categoria, "", "",
        costo, precio_venta, precio_venta, stock, 5,
        estado, margen, "unidad", proveedor,
        ahora, "web_user"
    ]
    hoja_inv.append_row(fila, value_input_option="USER_ENTERED")
    return jsonify({"mensaje": "Producto creado", "sku": sku, "precio_sugerido": precio_venta})

@app.route('/api/actualizar-producto', methods=['PUT'])
def api_actualizar_producto():
    if not autenticado or spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    sku = data.get('sku')
    precio_venta = parse_decimal(data.get('precio_venta', 0))
    stock = parse_int(data.get('stock', 0))

    hoja_inv = get_worksheet('inventario')
    if hoja_inv is None:
        return jsonify({"error": "Hoja inventario no encontrada"}), 404
    filas = hoja_inv.get_all_values()
    for i, fila in enumerate(filas):
        if fila and fila[0] == sku:
            if precio_venta > 0:
                hoja_inv.update_cell(i+1, 10, precio_venta)
            if stock >= 0:
                hoja_inv.update_cell(i+1, 11, stock)
                estado = "OK" if stock > 5 else "Bajo"
                hoja_inv.update_cell(i+1, 13, estado)
            return jsonify({"mensaje": "Producto actualizado"})
    return jsonify({"error": "Producto no encontrado"}), 404

@app.route('/api/boleta/<int:boleta_id>')
def api_boleta(boleta_id):
    if not autenticado or spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    hoja = get_worksheet('boletas')
    if hoja is None:
        return jsonify({"error": "Hoja boletas no encontrada"}), 404
    registros = hoja.get_all_records()
    items = [r for r in registros if parse_int(r.get('ID_Boleta', 0)) == boleta_id]
    if not items:
        return jsonify({"error": "Boleta no encontrada"}), 404
    subtotal = sum(parse_decimal(r.get('Total_fila', 0)) for r in items)
    igv = round(subtotal - (subtotal / 1.18), 2)
    total = subtotal
    return jsonify({
        'boleta_id': boleta_id,
        'fecha': items[0].get('Fecha') if items else '',
        'items': items,
        'subtotal': subtotal,
        'igv': igv,
        'total': total,
        'vendedor': items[0].get('Vendedor', '') if items else '',
        'cliente': items[0].get('Cliente', '') if items else ''
    })

@app.route('/api/clientes')
def api_clientes():
    if not autenticado or spreadsheet is None:
        return jsonify([])
    try:
        hoja = get_worksheet('clientes')
        if hoja is None:
            return jsonify([])
        registros = hoja.get_all_records()
        return jsonify(registros)
    except Exception as e:
        print("Error en /api/clientes:", e)
        return jsonify([])

@app.route('/api/cliente', methods=['POST'])
def api_guardar_cliente():
    if not autenticado or spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    nombre = data.get('nombre', '').strip()
    telefono = data.get('telefono', '').strip()
    email = data.get('email', '').strip()
    cliente_id = data.get('id', '')

    if not nombre:
        return jsonify({"error": "Nombre es obligatorio"}), 400

    try:
        hoja = get_worksheet('clientes')
        if hoja is None:
            return jsonify({"error": "Hoja clientes no encontrada"}), 404

        if cliente_id:
            registros = hoja.get_all_records()
            for i, r in enumerate(registros, start=2):
                if str(r.get('id')) == str(cliente_id):
                    hoja.update_cell(i, 2, nombre)
                    hoja.update_cell(i, 3, telefono)
                    hoja.update_cell(i, 4, email)
                    return jsonify({"mensaje": "Cliente actualizado"})
            return jsonify({"error": "Cliente no encontrado"}), 404
        else:
            nuevo_id = str(int(datetime.now().timestamp() * 1000))
            hoja.append_row([nuevo_id, nombre, telefono, email, getToday()])
            return jsonify({"mensaje": "Cliente creado", "id": nuevo_id})
    except Exception as e:
        print("Error en /api/cliente:", e)
        return jsonify({"error": str(e)}), 500

@app.route('/api/anular-venta', methods=['POST'])
def api_anular_venta():
    if not autenticado or spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    boleta_id = data.get('boleta_id')
    razon = data.get('razon', 'Anulación manual')

    if not boleta_id:
        return jsonify({"error": "Falta número de boleta"}), 400

    hoja_boletas = get_worksheet('boletas')
    if hoja_boletas is None:
        return jsonify({"error": "Hoja boletas no encontrada"}), 404
    registros = hoja_boletas.get_all_records()
    items = [r for r in registros if parse_int(r.get('ID_Boleta', 0)) == boleta_id]
    if not items:
        return jsonify({"error": "Boleta no encontrada"}), 404

    hoja_inv = get_worksheet('inventario')
    if hoja_inv is None:
        return jsonify({"error": "Hoja inventario no encontrada"}), 404
    for item in items:
        sku = item.get('SKU')
        cantidad = parse_int(item.get('Cantidad', 0))
        filas_inv = hoja_inv.get_all_values()
        for i, fila in enumerate(filas_inv):
            if fila and fila[0] == sku:
                stock_actual = parse_int(fila[10])
                nuevo_stock = stock_actual + cantidad
                hoja_inv.update_cell(i+1, 11, nuevo_stock)
                estado = "OK" if nuevo_stock > 5 else "Bajo"
                hoja_inv.update_cell(i+1, 13, estado)
                break

    hoja_mov = get_worksheet('movimientos')
    if hoja_mov is not None:
        hoja_mov.append_row([
            ahora_iso(), '', 'venta_anulada', 0, 0, 0, f'Boleta #{boleta_id}', 'web_user', razon
        ])

    return jsonify({"mensaje": "Venta anulada correctamente"})

@app.route('/api/venta', methods=['POST'])
def api_registrar_venta():
    if not autenticado or spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500

    data = request.get_json()
    items = data.get('items', [])
    vendedor = data.get('vendedor', 'web_user')
    cliente = data.get('cliente', '')

    if not items:
        return jsonify({'error': 'No hay items'}), 400

    hoja_inv = get_worksheet('inventario')
    hoja_ventas = get_worksheet('ventas')
    hoja_boletas = get_worksheet('boletas')
    hoja_mov = get_worksheet('movimientos')
    hoja_extras = get_worksheet('extras_y_perdidas')
    hoja_ahorro = get_worksheet('ahorro')

    if any(h is None for h in [hoja_inv, hoja_ventas, hoja_boletas, hoja_mov]):
        return jsonify({"error": "Faltan hojas en el sheet"}), 500

    registros_inv = hoja_inv.get_all_records()
    productos = {r['SKU']: r for r in registros_inv}

    resultados = []
    ganancia_total = 0.0
    subtotal = 0.0
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

        nuevo_stock = stock_actual - cantidad
        filas = hoja_inv.get_all_values()
        for i, fila in enumerate(filas):
            if fila and fila[0] == sku:
                hoja_inv.update_cell(i+1, 11, nuevo_stock)
                break

        costo = parse_decimal(prod.get('Costo', 0))
        ganancia_unidad = precio_real - costo
        ganancia_total += ganancia_unidad * cantidad
        subtotal += precio_real * cantidad

        hoja_ventas.append_row([
            now, id_boleta, sku, prod['Nombre_completo'],
            cantidad, precio_real, costo, ganancia_unidad, ganancia_unidad * cantidad, vendedor
        ])
        hoja_mov.append_row([
            now, sku, 'venta', cantidad, precio_real, cantidad * precio_real, id_boleta, vendedor, ''
        ])
        precio_sugerido = parse_decimal(prod.get('Precio_venta_actual', 0))
        diferencia = precio_real - precio_sugerido
        tipo_diferencia = 'extra' if diferencia > 0 else ('perdida' if diferencia < 0 else 'normal')
        hoja_boletas.append_row([
            id_boleta, now, cliente, '', sku, prod['Nombre_completo'], cantidad,
            precio_sugerido, precio_real, diferencia, tipo_diferencia, cantidad * precio_real, vendedor
        ])
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
        'boleta_id': id_boleta,
        'subtotal': subtotal
    })

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)