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
import groq

load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app)

# ==================== CONFIGURACIÓN ====================
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CREDS_FILE = os.getenv("CREDS_FILE", "creds_nuevo.json")
GOOGLE_CREDS = os.getenv("GOOGLE_CREDS")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ==================== AUTENTICACIÓN GOOGLE ====================
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

# ==================== GROQ ====================
groq_client = None
if GROQ_API_KEY:
    try:
        groq_client = groq.Groq(api_key=GROQ_API_KEY)
        print("✅ Groq client inicializado")
    except Exception as e:
        print(f"❌ Error inicializando Groq: {e}")

# ==================== FUNCIONES DE UTILIDAD ====================
def get_worksheet(name):
    if spreadsheet is None:
        return None
    try:
        return spreadsheet.worksheet(name)
    except:
        return None

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
        return jsonify({"error": "No conectado"}), 500
    hoja = get_worksheet('inventario')
    if hoja is None:
        return jsonify([])
    registros = hoja.get_all_records()
    for r in registros:
        for k in ['Costo', 'Precio_sugerido_por_sistema', 'Precio_venta_actual', 'Stock_actual', 'Stock_minimo', 'Margen']:
            if k in r:
                r[k] = parse_decimal(r[k])
        r['Stock_actual'] = parse_int(r.get('Stock_actual', 0))
        # Calcular ventas desde movimientos
        ventas = [m for m in movimientos_cache if m.get('SKU') == r['SKU'] and m.get('Tipo') == 'venta']
        r['ventas'] = sum(parse_int(m.get('Cantidad', 0)) for m in ventas)
    return jsonify(registros)

# Cache para movimientos (para no leer cada vez)
movimientos_cache = []

@app.route('/api/movimientos')
def api_movimientos():
    if spreadsheet is None:
        return jsonify([])
    hoja = get_worksheet('movimientos')
    if hoja is None:
        return jsonify([])
    registros = hoja.get_all_records()
    registros.sort(key=lambda x: x.get('Fecha', ''), reverse=True)
    global movimientos_cache
    movimientos_cache = registros
    return jsonify(registros[:500])

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
    if spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    sku = data.get('sku', '').strip()
    nombre = data.get('nombre', '').strip()
    categoria = data.get('categoria', 'General')
    costo = parse_decimal(data.get('costo', 0))
    stock = parse_int(data.get('stock', 0))
    stockMinimo = parse_int(data.get('stockMinimo', 5))
    proveedor = data.get('proveedor', '')
    precio_venta = parse_decimal(data.get('precio_venta', 0))
    alias = data.get('alias', '')
    imagen = data.get('imagen', '')

    if not nombre:
        return jsonify({"error": "Nombre es obligatorio"}), 400
    if costo <= 0:
        return jsonify({"error": "Costo debe ser mayor a 0"}), 400

    hoja_inv = get_worksheet('inventario')
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
        sku, nombre, "estandar", nombre, categoria, "", alias,
        costo, precio_venta, precio_venta, stock, stockMinimo,
        estado, margen, "unidad", proveedor,
        ahora, "web_user", imagen
    ]
    hoja_inv.append_row(fila, value_input_option="USER_ENTERED")
    return jsonify({"mensaje": "Producto creado", "sku": sku, "precio_sugerido": precio_venta})

@app.route('/api/actualizar-producto', methods=['PUT'])
def api_actualizar_producto():
    if spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    sku = data.get('sku')
    nombre = data.get('nombre')
    categoria = data.get('categoria')
    costo = parse_decimal(data.get('costo', 0))
    precio_venta = parse_decimal(data.get('precio_venta', 0))
    stock = parse_int(data.get('stock', 0))
    stockMinimo = parse_int(data.get('stockMinimo', 0))
    proveedor = data.get('proveedor', '')
    alias = data.get('alias', '')
    imagen = data.get('imagen', '')

    hoja_inv = get_worksheet('inventario')
    filas = hoja_inv.get_all_values()
    for i, fila in enumerate(filas):
        if fila and fila[0] == sku:
            # Actualizar columnas: 2 nombre, 5 categoria, 8 costo, 10 precio, 11 stock, 12 stockMin, 16 proveedor, 7 alias, 19 imagen
            if nombre: hoja_inv.update_cell(i+1, 2, nombre)
            if categoria: hoja_inv.update_cell(i+1, 5, categoria)
            if costo > 0: hoja_inv.update_cell(i+1, 8, costo)
            if precio_venta > 0: hoja_inv.update_cell(i+1, 10, precio_venta)
            if stock >= 0: hoja_inv.update_cell(i+1, 11, stock)
            if stockMinimo >= 0: hoja_inv.update_cell(i+1, 12, stockMinimo)
            if proveedor: hoja_inv.update_cell(i+1, 16, proveedor)
            if alias: hoja_inv.update_cell(i+1, 7, alias)
            if imagen: hoja_inv.update_cell(i+1, 19, imagen)
            # Actualizar estado stock
            if stock >= 0:
                estado = "OK" if stock > 5 else "Bajo"
                hoja_inv.update_cell(i+1, 13, estado)
            return jsonify({"mensaje": "Producto actualizado"})
    return jsonify({"error": "Producto no encontrado"}), 404

@app.route('/api/venta', methods=['POST'])
def api_registrar_venta():
    if spreadsheet is None:
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
    hoja_settings = get_worksheet('configuracion')  # para número de boleta

    if any(h is None for h in [hoja_inv, hoja_ventas, hoja_boletas, hoja_mov]):
        return jsonify({"error": "Faltan hojas en el sheet"}), 500

    # Leer siguiente número de boleta
    next_boleta = 934416062
    try:
        settings_rows = hoja_settings.get_all_values() if hoja_settings else []
        for row in settings_rows:
            if row and row[0] == 'next_boleta':
                next_boleta = int(row[1])
                break
    except:
        pass

    registros_inv = hoja_inv.get_all_records()
    productos = {r['SKU']: r for r in registros_inv}

    resultados = []
    ganancia_total = 0.0
    subtotal = 0.0
    now = ahora_iso()
    id_boleta = next_boleta

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
            now, sku, 'venta', cantidad, precio_real, cantidad * precio_real, '', vendedor, ''
        ])
        precio_sugerido = parse_decimal(prod.get('Precio_venta_actual', 0))
        diferencia = precio_real - precio_sugerido
        tipo_diferencia = 'extra' if diferencia > 0 else ('perdida' if diferencia < 0 else 'normal')
        hoja_boletas.append_row([
            id_boleta, now, '', cliente, sku, prod['Nombre_completo'], cantidad,
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

    # Actualizar next_boleta
    next_boleta += 1
    try:
        if hoja_settings:
            for i, row in enumerate(hoja_settings.get_all_values()):
                if row and row[0] == 'next_boleta':
                    hoja_settings.update_cell(i+1, 2, next_boleta)
                    break
            else:
                hoja_settings.append_row(['next_boleta', next_boleta])
    except:
        pass

    return jsonify({
        'mensaje': 'Venta registrada exitosamente',
        'detalle': resultados,
        'ganancia_total': ganancia_total,
        'boleta_id': id_boleta,
        'subtotal': subtotal
    })

@app.route('/api/boleta/<int:boleta_id>')
def api_boleta(boleta_id):
    if spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    hoja = get_worksheet('boletas')
    if hoja is None:
        return jsonify({"error": "Hoja no encontrada"}), 404
    registros = hoja.get_all_records()
    items = [r for r in registros if parse_int(r.get('ID_Boleta', 0)) == boleta_id]
    if not items:
        return jsonify({"error": "Boleta no encontrada"}), 404
    subtotal = sum(parse_decimal(r.get('Total_fila', 0)) for r in items)
    igv = round(subtotal / 1.18 * 0.18, 2)
    total = round(subtotal, 2)
    cliente = items[0].get('Cliente', '') if items else ''
    return jsonify({
        'boleta_id': boleta_id,
        'fecha': items[0].get('Fecha') if items else '',
        'items': items,
        'subtotal': subtotal,
        'igv': igv,
        'total': total,
        'vendedor': items[0].get('Vendedor', '') if items else '',
        'cliente': cliente
    })

@app.route('/api/anular', methods=['POST'])
def api_anular_venta():
    if spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    fecha = data.get('fecha')
    sku = data.get('sku')
    if not fecha or not sku:
        return jsonify({"error": "Faltan datos"}), 400

    # Buscar la venta en movimientos por fecha y sku
    hoja_mov = get_worksheet('movimientos')
    registros = hoja_mov.get_all_records()
    for i, r in enumerate(registros, start=2):
        if r.get('Fecha') == fecha and r.get('SKU') == sku and r.get('Tipo') == 'venta':
            # Revertir stock
            cantidad = parse_int(r.get('Cantidad', 0))
            if cantidad > 0:
                hoja_inv = get_worksheet('inventario')
                inv_rows = hoja_inv.get_all_values()
                for j, row in enumerate(inv_rows):
                    if row and row[0] == sku:
                        stock_actual = parse_int(row[10])
                        nuevo_stock = stock_actual + cantidad
                        hoja_inv.update_cell(j+1, 11, nuevo_stock)
                        break
            # Marcar como anulada en movimientos (o eliminar fila)
            hoja_mov.update_cell(i, 3, 'venta_anulada')  # tipo
            hoja_mov.update_cell(i, 9, 'Anulada por web')  # nota
            return jsonify({"mensaje": "Venta anulada correctamente"})
    return jsonify({"error": "Venta no encontrada"}), 404

@app.route('/api/reset', methods=['POST'])
def api_reset_mensual():
    if spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    data = request.get_json()
    password = data.get('password')
    if password != os.getenv("RESET_PASSWORD", "Lumin2024"):
        return jsonify({"error": "Contraseña incorrecta"}), 403

    mes_actual = datetime.now().strftime("%Y-%m")
    fecha_backup = ahora_iso()
    sheets_to_reset = ["movimientos", "compras", "ventas", "boletas", "extras_y_perdidas", "ahorro", "proformas"]
    resumen = []
    for nombre in sheets_to_reset:
        hoja = get_worksheet(nombre)
        if hoja is None:
            continue
        filas = hoja.get_all_values()
        if not filas:
            continue
        encabezado, datos = filas[0], filas[1:]
        # Buscar columna de fecha
        idx_fecha = 0
        for i, col in enumerate(encabezado):
            if "fecha" in col.lower() and "confirmacion" not in col.lower():
                idx_fecha = i
                break
        filas_mes = [f for f in datos if len(f) > idx_fecha and str(f[idx_fecha]).startswith(mes_actual)]
        filas_resto = [f for f in datos if not (len(f) > idx_fecha and str(f[idx_fecha]).startswith(mes_actual))]
        if not filas_mes:
            resumen.append(f"• {nombre}: sin datos de {mes_actual}")
            continue
        backup_json = json.dumps({"hoja": nombre, "encabezado": encabezado, "filas": filas_mes}, ensure_ascii=False)
        hoja_backup = get_worksheet('backups_mensuales')
        if hoja_backup:
            hoja_backup.append_row([fecha_backup, mes_actual, "web_user", len(filas_mes), backup_json])
        hoja.clear()
        hoja.update([encabezado] + filas_resto, value_input_option="USER_ENTERED")
        resumen.append(f"• {nombre}: {len(filas_mes)} filas archivadas y borradas")
    return jsonify({"mensaje": "Reset completado\n" + "\n".join(resumen)})

@app.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    if spreadsheet is None:
        return jsonify({"error": "No conectado"}), 500
    hoja = get_worksheet('configuracion')
    if hoja is None:
        return jsonify({})
    if request.method == 'GET':
        rows = hoja.get_all_values()
        settings = {}
        for row in rows:
            if len(row) >= 2:
                settings[row[0]] = row[1]
        # Valores por defecto
        return jsonify({
            'storeName': settings.get('storeName', 'LUMIN STORE'),
            'phone': settings.get('phone', '999 999 999'),
            'nextBoleta': int(settings.get('nextBoleta', 934416062)),
            'metaDiaria': float(settings.get('metaDiaria', 500)),
            'logo': settings.get('logo', '')
        })
    else:  # POST
        data = request.get_json()
        # Guardar cada clave en la hoja
        for key, value in data.items():
            # Buscar si existe
            rows = hoja.get_all_values()
            found = False
            for i, row in enumerate(rows, start=1):
                if row and row[0] == key:
                    hoja.update_cell(i, 2, str(value))
                    found = True
                    break
            if not found:
                hoja.append_row([key, str(value)])
        return jsonify({"mensaje": "Configuración guardada"})

@app.route('/api/chat', methods=['POST'])
def api_chat():
    data = request.get_json()
    pregunta = data.get('pregunta', '')
    if not pregunta:
        return jsonify({"respuesta": "No entendí tu pregunta."})

    # Si no hay Groq, responder con lógica local básica
    if not groq_client:
        return jsonify({"respuesta": generate_local_response(pregunta)})

    # Construir contexto del negocio
    inventario_valor = sum(parse_decimal(p.get('Costo', 0)) * parse_int(p.get('Stock_actual', 0)) for p in inventario_cache)
    hoy = date.today().isoformat()
    ventas_hoy = sum(parse_decimal(m.get('Total', 0)) for m in movimientos_cache if m.get('Tipo') == 'venta' and m.get('Fecha', '').startswith(hoy))
    ventas_mes = sum(parse_decimal(m.get('Total', 0)) for m in movimientos_cache if m.get('Tipo') == 'venta' and m.get('Fecha', '').startswith(datetime.now().strftime("%Y-%m")))
    compras_mes = sum(parse_decimal(m.get('Total', 0)) for m in movimientos_cache if m.get('Tipo') == 'compra' and m.get('Fecha', '').startswith(datetime.now().strftime("%Y-%m")))
    ganancia_mes = ventas_mes - compras_mes
    productos_bajos = [p['Nombre_completo'] for p in inventario_cache if parse_int(p.get('Stock_actual', 0)) <= parse_int(p.get('Stock_minimo', 5))]

    prompt = f"""Eres un consultor financiero y de inventario para un negocio peruano llamado LUMIN.
    Contexto actual:
    - Valor del inventario: S/{inventario_valor:,.2f}
    - Ventas de hoy: S/{ventas_hoy:,.2f}
    - Ventas del mes: S/{ventas_mes:,.2f}
    - Compras del mes: S/{compras_mes:,.2f}
    - Ganancia neta del mes: S/{ganancia_mes:,.2f}
    - Productos con stock bajo: {', '.join(productos_bajos) if productos_bajos else 'Ninguno'}
    - Total de productos: {len(inventario_cache)}

    El usuario pregunta: "{pregunta}"

    Responde de manera clara, directa, en español peruano, con recomendaciones accionables. Si es una consulta de números, dame cifras exactas. Si es de recomendaciones, sugiere acciones concretas.
    """
    try:
        response = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
            max_tokens=500
        )
        respuesta = response.choices[0].message.content.strip()
        return jsonify({"respuesta": respuesta})
    except Exception as e:
        print("Error en chat Groq:", e)
        return jsonify({"respuesta": generate_local_response(pregunta)})

# Cache para inventario en chat
inventario_cache = []

def generate_local_response(pregunta):
    # Respuesta local si no hay Groq
    if 'stock' in pregunta.lower() or 'inventario' in pregunta.lower():
        total = sum(parse_int(p.get('Stock_actual', 0)) for p in inventario_cache)
        return f"Hay {total} unidades en total en el inventario."
    if 'venta' in pregunta.lower():
        hoy = date.today().isoformat()
        ventas_hoy = sum(parse_decimal(m.get('Total', 0)) for m in movimientos_cache if m.get('Tipo') == 'venta' and m.get('Fecha', '').startswith(hoy))
        return f"Las ventas de hoy suman {fmt_money(ventas_hoy)}."
    if 'ganancia' in pregunta.lower() or 'rentabilidad' in pregunta.lower():
        ventas_mes = sum(parse_decimal(m.get('Total', 0)) for m in movimientos_cache if m.get('Tipo') == 'venta' and m.get('Fecha', '').startswith(datetime.now().strftime("%Y-%m")))
        compras_mes = sum(parse_decimal(m.get('Total', 0)) for m in movimientos_cache if m.get('Tipo') == 'compra' and m.get('Fecha', '').startswith(datetime.now().strftime("%Y-%m")))
        ganancia = ventas_mes - compras_mes
        return f"La ganancia neta del mes es {fmt_money(ganancia)} (Ventas: {fmt_money(ventas_mes)}, Compras: {fmt_money(compras_mes)})."
    if 'recomend' in pregunta.lower():
        bajos = [p['Nombre_completo'] for p in inventario_cache if parse_int(p.get('Stock_actual', 0)) <= parse_int(p.get('Stock_minimo', 5))]
        if bajos:
            return f"Te recomiendo reponer: {', '.join(bajos[:5])}."
        return "Todo el stock está en niveles adecuados."
    return "Puedo ayudarte con consultas de stock, ventas, rentabilidad y recomendaciones. ¿Qué necesitas saber?"

# ==================== INICIO ====================
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)