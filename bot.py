#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bot de Inventario Lumin - VERSIÓN WEBHOOK PARA RENDER
Fecha: 2026-06-23

CORRECCIONES APLICADAS (2026-06-24):
1. normalizar() ahora preserva puntos (.) y dos puntos (:) para decimales y etiquetas c:/v:
2. extraer_numeros() ahora captura decimales con punto o coma
3. parsear_compra() acepta decimales, formato "a" para costo, y etiquetas c/v
4. parsear_items_multiples() acepta decimales en cantidad y precio
5. detectar_intencion() usa word boundaries (\b) para evitar matching parcial
   (ej: "mas" en "camas" ya no coincide con intención "agregar")
"""

import os
import re
import json
import time
import random
import threading
import unicodedata
import logging
import asyncio
import difflib
from datetime import datetime, date, timedelta
from typing import List, Dict, Optional, Tuple, Any
from collections import defaultdict

import gspread
from google.oauth2.service_account import Credentials
from groq import Groq
from flask import Flask, request

from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes, filters
)

# ==================== CONFIGURACIÓN ====================
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CREDS_FILE = os.getenv("CREDS_FILE", "creds_nuevo.json")
RESET_PASSWORD = os.getenv("RESET_PASSWORD")
ACCESS_PASSWORD = os.getenv("ACCESS_PASSWORD")
WEBHOOK_URL = os.getenv("WEBHOOK_URL")

GROQ_MODEL = "llama-3.1-8b-instant"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
log = logging.getLogger("inventario_bot")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
gc = gspread.authorize(_creds)
spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
groq_client = Groq(api_key=GROQ_API_KEY)

_WS_CACHE = {}
USUARIOS_AUTENTICADOS = {}
CONTEXTO_SELECCION = {}
PROFORMAS_PENDIENTES = {}


def ws(nombre_hoja):
    if nombre_hoja not in _WS_CACHE:
        _WS_CACHE[nombre_hoja] = spreadsheet.worksheet(nombre_hoja)
    return _WS_CACHE[nombre_hoja]


EXCEL_EPOCH = datetime(1899, 12, 30)


def serial_excel_a_datetime(serial):
    try:
        if serial is None or serial == "":
            return None
        if isinstance(serial, (int, float)):
            return EXCEL_EPOCH + timedelta(days=float(serial))
        s = str(serial).strip().replace(',', '.')
        return EXCEL_EPOCH + timedelta(days=float(s))
    except (ValueError, TypeError, OverflowError):
        return None


def fecha_a_iso(fecha_valor):
    if fecha_valor is None or fecha_valor == "":
        return ""
    if isinstance(fecha_valor, str):
        s = fecha_valor.strip()
        if len(s) >= 10 and s[4] == '-' and s[7] == '-':
            return s[:10]
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
            try:
                return datetime.strptime(s[:19] if len(s) > 19 else s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        try:
            return serial_excel_a_datetime(float(s.replace(',', '.'))).strftime("%Y-%m-%d")
        except (ValueError, TypeError, AttributeError):
            return ""
    if isinstance(fecha_valor, (int, float)):
        dt = serial_excel_a_datetime(fecha_valor)
        if dt:
            return dt.strftime("%Y-%m-%d")
        return ""
    if isinstance(fecha_valor, datetime):
        return fecha_valor.strftime("%Y-%m-%d")
    return ""


def ahora_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_decimal(valor):
    if valor is None or valor == "":
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    s = re.sub(r"[Ss]/\.?", "", s)
    s = s.replace(" ", "")
    if s == "":
        return 0.0
    tiene_coma = "," in s
    tiene_punto = "." in s
    if tiene_coma and tiene_punto:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif tiene_coma and not tiene_punto:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_int(valor):
    return int(round(parse_decimal(valor)))


# ==================== CORRECCIÓN 1: normalizar() ====================
def normalizar(texto):
    """Normaliza texto pero PRESERVA puntos decimales y comas decimales."""
    if texto is None:
        return ""
    texto = str(texto).lower().strip()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    # CORRECCIÓN: PRESERVAR puntos (.), dos puntos (:), y comas (,)
    texto = re.sub(r"[^a-z0-9\s/,.:]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def limpiar_segmento(texto):
    return texto.strip(" ,") if texto else ""


def limpiar_nombre_completo(nombre):
    nombre = re.sub(r',\s*,', ',', nombre)
    nombre = re.sub(r',\s*$', '', nombre)
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    return nombre


def limpiar_nombre_creacion(texto):
    t = re.sub(r"\s+", " ", texto.strip(" ,"))
    return t.title()


def singularizar(palabra):
    if len(palabra) > 3 and palabra.endswith("s"):
        return palabra[:-1]
    return palabra


def fmt_money(valor):
    return f"S/{valor:,.2f}"


FILLER_WORDS = {"estandar", "unidad", "unidades"}

STOPWORDS = {
    "a", "de", "del", "el", "la", "los", "las", "un", "una", "al",
    "x", "s", "soles", "sol", "para", "que", "por", "y", "e",
}


def tokens_significativos(texto, eliminar_stopwords=True):
    crudos = normalizar(texto).split()
    tokens = []
    for t in crudos:
        t = t.strip(",")
        if not t or t in FILLER_WORDS:
            continue
        if eliminar_stopwords and t in STOPWORDS:
            continue
        tokens.append(singularizar(t))
    return tokens


try:
    from rapidfuzz import fuzz, process
    HAS_FUZZ = True
    log.info("Usando rapidfuzz para fuzzy matching")
except ImportError:
    try:
        from thefuzz import fuzz, process
        HAS_FUZZ = True
        log.info("Usando thefuzz para fuzzy matching")
    except ImportError:
        HAS_FUZZ = False
        log.warning("thefuzz/rapidfuzz no instalado. Usando difflib (mas lento)")


def _mejor_score_token(token_query, tokens_candidato):
    mejor = 0.0
    for tc in tokens_candidato:
        if token_query == tc:
            return 1.0
        if len(token_query) >= 4 and (token_query in tc or tc in token_query):
            mejor = max(mejor, 0.9)
        else:
            if HAS_FUZZ:
                score = fuzz.ratio(token_query, tc) / 100.0
            else:
                score = difflib.SequenceMatcher(None, token_query, tc).ratio()
            mejor = max(mejor, score)
    return mejor


def _score_producto(tokens_query, tokens_candidato):
    if not tokens_query or not tokens_candidato:
        return 0.0
    puntajes = [_mejor_score_token(tq, tokens_candidato) for tq in tokens_query]
    minimo = min(puntajes)
    promedio = sum(puntajes) / len(puntajes)
    if minimo < 0.55:
        return promedio * 0.25
    return promedio


def _tokens_de_producto(p, eliminar_stopwords=True):
    candidatos_texto = [p["Nombre_base"], p["Nombre_completo"], p["SKU"]]
    if p.get("Alias"):
        candidatos_texto += [a for a in p["Alias"].split(";") if a.strip()]
    tokens = []
    for c in candidatos_texto:
        tokens += tokens_significativos(c, eliminar_stopwords)
    return tokens


def buscar_producto(query, productos, umbral=0.6):
    query_norm = normalizar(query)
    for p in productos:
        alias = p.get("Alias", "")
        if alias and normalizar(alias) == query_norm:
            return "ok", p

    tokens_q = tokens_significativos(query, eliminar_stopwords=True)
    if not tokens_q:
        return "no_encontrado", None

    numeros_query = [t for t in tokens_q if t.isdigit() or re.match(r'\d+/\d+', t)]

    puntajes = [(_score_producto(tokens_q, _tokens_de_producto(p, eliminar_stopwords=True)), p)
                for p in productos]
    puntajes.sort(key=lambda x: x[0], reverse=True)
    mejor_score, _ = puntajes[0]

    if mejor_score < umbral:
        return "no_encontrado", None

    UMBRAL_MULTIPLE = 0.35
    candidatos_top = [p for score, p in puntajes if score >= UMBRAL_MULTIPLE]

    if numeros_query:
        candidatos_filtrados = []
        for p in candidatos_top:
            nombre_prod = p["Nombre_completo"].lower()
            if all(n in nombre_prod for n in numeros_query):
                candidatos_filtrados.append(p)
        if candidatos_filtrados:
            candidatos_top = candidatos_filtrados

    candidatos_top = sorted(candidatos_top, key=lambda p: p["Precio_venta_actual"], reverse=True)

    skus_distintos = {p["SKU"] for p in candidatos_top}
    if len(skus_distintos) > 1:
        return "multiple", candidatos_top[:10]

    return "ok", candidatos_top[0] if candidatos_top else None


def formatear_lista_productos(candidatos):
    lineas = []
    for i, p in enumerate(candidatos, start=1):
        nombre_fmt = f"`{p['Nombre_completo']}`"
        alias = (p.get("Alias") or "").strip()
        if alias and ";" in alias:
            alias = alias.split(";")[0].strip()
        alias_fmt = f" (alias: `{alias}`)" if alias else ""
        lineas.append(
            f"{i}. {nombre_fmt}{alias_fmt} - Precio: {fmt_money(p['Precio_venta_actual'])} - "
            f"Stock: {parse_int(p['Stock_actual'])}"
        )
    lineas.append("")
    lineas.append("💡 Copia el nombre o alias para usarlo en tus comandos.")
    return "\n".join(lineas)


def generar_sku(nombre):
    base = re.sub(r"[^A-Za-z0-9]", "", nombre.upper())[:5] or "PROD"
    existentes = {p["SKU"] for p in leer_inventario()}
    while True:
        sufijo = datetime.now().strftime("%d%H%M") + str(random.randint(10, 99))
        sku = f"{base}-{sufijo}"
        if sku not in existentes:
            return sku


COLS_INVENTARIO = [
    "SKU", "Nombre_base", "Variante", "Nombre_completo", "Categoria", "Categoria2",
    "Alias", "Costo", "Precio_sugerido_por_sistema", "Precio_venta_actual",
    "Stock_actual", "Stock_minimo", "Estado_stock", "Margen", "Unidad",
    "Proveedor", "Ultima_actualizacion", "Actualizado_por",
]
COL_IDX = {nombre: i + 1 for i, nombre in enumerate(COLS_INVENTARIO)}
NUMERIC_INVENTARIO = {
    "Costo", "Precio_sugerido_por_sistema", "Precio_venta_actual",
    "Stock_actual", "Stock_minimo", "Margen",
}


_CACHE_INVENTARIO = {"datos": None, "timestamp": 0.0}
CACHE_INVENTARIO_TTL = 5


def invalidar_cache_inventario():
    _CACHE_INVENTARIO["timestamp"] = 0.0


def leer_inventario():
    ahora = time.time()
    if (
        _CACHE_INVENTARIO["datos"] is not None
        and (ahora - _CACHE_INVENTARIO["timestamp"]) < CACHE_INVENTARIO_TTL
    ):
        return _CACHE_INVENTARIO["datos"]

    hoja = ws("inventario")
    filas = hoja.get_all_values()
    productos = []
    for i, fila in enumerate(filas[1:], start=2):
        if not any(fila):
            continue
        d = {}
        for col_nombre, idx in COL_IDX.items():
            valor = fila[idx - 1] if idx - 1 < len(fila) else ""
            if col_nombre in NUMERIC_INVENTARIO:
                d[col_nombre] = parse_decimal(valor)
            else:
                d[col_nombre] = valor
        d["_row"] = i
        productos.append(d)

    _CACHE_INVENTARIO["datos"] = productos
    _CACHE_INVENTARIO["timestamp"] = ahora
    return productos


def estado_stock(stock_actual, stock_minimo):
    return "OK" if stock_actual > stock_minimo else "Bajo"


def actualizar_celda(hoja_nombre, fila, columna_nombre, valor):
    ws(hoja_nombre).update_cell(fila, COL_IDX[columna_nombre], valor)
    if hoja_nombre == "inventario":
        invalidar_cache_inventario()


def generar_alias(nombre_base, productos_existentes):
    nombre_norm = normalizar(nombre_base)
    sin_espacios = nombre_norm.replace(" ", "")

    if len(sin_espacios) <= 8:
        alias_base = sin_espacios
    else:
        palabras = [p for p in nombre_norm.split() if p not in FILLER_WORDS and p not in STOPWORDS]
        if not palabras:
            palabras = nombre_norm.split()
        partes = []
        for p in palabras:
            if p.isdigit():
                partes.append(p[:2])
            else:
                partes.append(p[:4] if len(palabras) <= 2 else p[:2])
        alias_base = "".join(partes)[:10] or sin_espacios[:10]

    if not alias_base:
        alias_base = "prod"

    alias_existentes = set()
    for p in productos_existentes:
        a = normalizar(p.get("Alias", ""))
        if a:
            alias_existentes.add(a)

    if alias_base not in alias_existentes:
        return alias_base

    contador = 2
    while f"{alias_base}{contador}" in alias_existentes:
        contador += 1
    return f"{alias_base}{contador}"


def crear_producto(nombre, costo, stock_inicial, precio_venta=None, usuario=""):
    nombre = limpiar_nombre_creacion(nombre)
    sku = generar_sku(nombre)
    precio_sugerido = round(costo * 1.6, 2)
    precio_final = precio_venta if precio_venta else precio_sugerido
    margen = round((precio_final - costo) / costo, 6) if costo else 0
    ahora = ahora_iso()
    try:
        alias_generado = generar_alias(nombre, leer_inventario())
    except Exception:
        alias_generado = ""
    fila = [
        sku, nombre, "estandar", nombre, "General", "", alias_generado,
        costo, precio_sugerido, precio_final, stock_inicial, 5,
        estado_stock(stock_inicial, 5), margen, "unidad", "Telegram Bot",
        ahora, usuario,
    ]
    ws("inventario").append_row(fila, value_input_option="USER_ENTERED")
    invalidar_cache_inventario()
    return dict(zip(COLS_INVENTARIO, fila))


def registrar_movimiento(sku, tipo, cantidad, precio_unitario, usuario, nota=""):
    fecha = ahora_iso()
    total = round(cantidad * precio_unitario, 2)
    ws("movimientos").append_row(
        [fecha, sku, tipo, cantidad, precio_unitario, total, "", usuario, nota],
        value_input_option="USER_ENTERED",
    )


def registrar_compra(sku, cantidad, costo_unitario, usuario):
    fecha = ahora_iso()
    total = round(cantidad * costo_unitario, 2)
    ws("compras").append_row(
        [int(time.time() * 1000) % 10_000_000, fecha, sku, cantidad, costo_unitario, total, "", "", "Si", ""],
        value_input_option="USER_ENTERED",
    )


def registrar_venta_y_boleta(producto, cantidad, precio_real, usuario):
    fecha = ahora_iso()
    costo = producto["Costo"]
    precio_sugerido = producto["Precio_venta_actual"]
    ganancia_unidad = round(precio_real - costo, 2)
    ganancia_total = round(ganancia_unidad * cantidad, 2)
    id_boleta = int(time.time() * 1000) % 10_000_000

    ws("ventas").append_row(
        [fecha, id_boleta, producto["SKU"], producto["Nombre_completo"], cantidad,
         precio_real, costo, ganancia_unidad, ganancia_total, usuario],
        value_input_option="USER_ENTERED",
    )

    diferencia = round(precio_real - precio_sugerido, 2)
    tipo_diferencia = "extra" if diferencia > 0 else ("perdida" if diferencia < 0 else "normal")
    total_fila = round(precio_real * cantidad, 2)
    ws("boletas").append_row(
        [id_boleta, fecha, "", "", producto["SKU"], producto["Nombre_completo"], cantidad,
         precio_sugerido, precio_real, diferencia, tipo_diferencia, total_fila, usuario],
        value_input_option="USER_ENTERED",
    )

    monto_dif_total = round(diferencia * cantidad, 2)
    if diferencia > 0:
        ws("extras_y_perdidas").append_row(
            [fecha, id_boleta, producto["SKU"], "extra", monto_dif_total,
             "Venta por encima del precio sugerido", monto_dif_total, 0, monto_dif_total],
            value_input_option="USER_ENTERED",
        )
        ws("ahorro").append_row(
            [fecha, id_boleta, producto["SKU"], precio_sugerido, precio_real,
             monto_dif_total, "Venta por encima del precio sugerido", monto_dif_total],
            value_input_option="USER_ENTERED",
        )
    elif diferencia < 0:
        ws("extras_y_perdidas").append_row(
            [fecha, id_boleta, producto["SKU"], "perdida", abs(monto_dif_total),
             "Venta por debajo del precio sugerido", 0, abs(monto_dif_total), monto_dif_total],
            value_input_option="USER_ENTERED",
        )

    return ganancia_total


def registrar_log(usuario_id, usuario_nombre, comando, entrada, respuesta, tiempo):
    try:
        ahora = ahora_iso()
        ws("logs_bot").append_row(
            [ahora, usuario_id, usuario_nombre, comando, entrada, respuesta[:300], round(tiempo, 2)],
            value_input_option="USER_ENTERED",
        )
    except Exception as e:
        log.warning(f"No se pudo registrar log: {e}")


def actualizar_hoja_ganancias():
    try:
        hoy = date.today().isoformat()

        ventas_raw = ws("ventas").get_all_values(value_render_option="UNFORMATTED_VALUE")
        ahorro_raw = ws("ahorro").get_all_values(value_render_option="UNFORMATTED_VALUE")
        extras_raw = ws("extras_y_perdidas").get_all_values(value_render_option="UNFORMATTED_VALUE")

        ventas_hoy = [r for r in ventas_raw[1:] if r and r[0] and fecha_a_iso(r[0]) == hoy]
        ahorro_hoy = [r for r in ahorro_raw[1:] if r and r[0] and fecha_a_iso(r[0]) == hoy]
        extras_hoy = [r for r in extras_raw[1:] if r and r[0] and fecha_a_iso(r[0]) == hoy]

        ventas_totales = 0.0
        costos_totales = 0.0
        ganancia_neta = 0.0

        for r in ventas_hoy:
            try:
                cantidad = parse_decimal(r[4]) if len(r) > 4 else 0
                precio_venta = parse_decimal(r[5]) if len(r) > 5 else 0
                costo = parse_decimal(r[6]) if len(r) > 6 else 0
                ganancia_total_fila = parse_decimal(r[8]) if len(r) > 8 else 0
                ventas_totales += cantidad * precio_venta
                costos_totales += cantidad * costo
                ganancia_neta += ganancia_total_fila
            except Exception:
                continue

        margen_promedio = round(ganancia_neta / ventas_totales, 4) if ventas_totales else 0

        descuentos_dados = 0.0
        extra_ahorro = 0.0

        for r in extras_hoy:
            try:
                tipo = str(r[3]).lower().strip() if len(r) > 3 else ""
                monto = parse_decimal(r[4]) if len(r) > 4 else 0
                if tipo == "perdida":
                    descuentos_dados += monto
                elif tipo == "extra":
                    extra_ahorro += monto
            except Exception:
                continue

        for r in ahorro_hoy:
            try:
                extra_val = parse_decimal(r[5]) if len(r) > 5 else 0
                extra_ahorro += extra_val
            except Exception:
                continue

        producto_top, vendedor_top = "", ""
        if ventas_hoy:
            acum_p, acum_v = {}, {}
            for r in ventas_hoy:
                try:
                    producto = str(r[3]) if len(r) > 3 else ""
                    vendedor = str(r[9]) if len(r) > 9 else ""
                    ganancia = parse_decimal(r[8]) if len(r) > 8 else 0
                    if producto:
                        acum_p[producto] = acum_p.get(producto, 0) + ganancia
                    if vendedor:
                        acum_v[vendedor] = acum_v.get(vendedor, 0) + ganancia
                except Exception:
                    continue
            if acum_p:
                producto_top = max(acum_p, key=acum_p.get)
            if acum_v:
                vendedor_top = max(acum_v, key=acum_v.get)

        fila_nueva = [
            hoy, round(ventas_totales, 2), round(costos_totales, 2), round(ganancia_neta, 2),
            margen_promedio, round(descuentos_dados, 2), round(extra_ahorro, 2),
            producto_top, vendedor_top
        ]

        hoja = ws("ganancias")
        filas = hoja.get_all_values(value_render_option="UNFORMATTED_VALUE")

        fila_idx = None
        for i, f in enumerate(filas[1:], start=2):
            if f and fecha_a_iso(f[0]) == hoy:
                fila_idx = i
                break

        if fila_idx:
            hoja.update(f"A{fila_idx}:I{fila_idx}", [fila_nueva], value_input_option="USER_ENTERED")
        else:
            hoja.append_row(fila_nueva, value_input_option="USER_ENTERED")

        return {
            'ventas': ventas_totales, 'costos': costos_totales,
            'ganancia_neta': ganancia_neta, 'margen': margen_promedio,
            'descuentos': descuentos_dados, 'extras': extra_ahorro,
            'producto_top': producto_top, 'vendedor_top': vendedor_top
        }

    except Exception as e:
        log.warning(f"No se pudo actualizar hoja ganancias: {e}")
        return None


def calcular_resumen_periodo(fecha_inicio_iso, fecha_fin_iso):
    try:
        hoja_ganancias = ws("ganancias")
        filas_ganancias = hoja_ganancias.get_all_values(value_render_option="UNFORMATTED_VALUE")

        ventas_totales = 0.0
        costos_totales = 0.0
        ganancia_neta = 0.0
        descuentos_dados = 0.0
        extra_ahorro = 0.0
        dias_con_datos = 0

        for f in filas_ganancias[1:]:
            if not f or not f[0]:
                continue
            fecha_fila = fecha_a_iso(f[0])
            if not fecha_fila or not (fecha_inicio_iso <= fecha_fila <= fecha_fin_iso):
                continue
            dias_con_datos += 1
            ventas_totales += parse_decimal(f[1]) if len(f) > 1 else 0
            costos_totales += parse_decimal(f[2]) if len(f) > 2 else 0
            ganancia_neta += parse_decimal(f[3]) if len(f) > 3 else 0
            descuentos_dados += parse_decimal(f[5]) if len(f) > 5 else 0
            extra_ahorro += parse_decimal(f[6]) if len(f) > 6 else 0

        margen_periodo = round(ganancia_neta / ventas_totales, 4) if ventas_totales else 0.0

        producto_top, vendedor_top = "", ""
        try:
            filas_ventas = ws("ventas").get_all_values(value_render_option="UNFORMATTED_VALUE")
            acum_p, acum_v = {}, {}
            for r in filas_ventas[1:]:
                if not r or not r[0]:
                    continue
                fecha_r = fecha_a_iso(r[0])
                if not fecha_r or not (fecha_inicio_iso <= fecha_r <= fecha_fin_iso):
                    continue
                producto = str(r[3]) if len(r) > 3 else ""
                vendedor = str(r[9]) if len(r) > 9 else ""
                ganancia_fila = parse_decimal(r[8]) if len(r) > 8 else 0
                if producto:
                    acum_p[producto] = acum_p.get(producto, 0) + ganancia_fila
                if vendedor:
                    acum_v[vendedor] = acum_v.get(vendedor, 0) + ganancia_fila
            if acum_p:
                producto_top = max(acum_p, key=acum_p.get)
            if acum_v:
                vendedor_top = max(acum_v, key=acum_v.get)
        except Exception as e:
            log.warning(f"No se pudo calcular top del periodo desde ventas: {e}")

        return {
            'dias_con_datos': dias_con_datos,
            'ventas': ventas_totales, 'costos': costos_totales,
            'ganancia_neta': ganancia_neta, 'margen': margen_periodo,
            'descuentos': descuentos_dados, 'extras': extra_ahorro,
            'producto_top': producto_top, 'vendedor_top': vendedor_top,
        }
    except Exception as e:
        log.warning(f"No se pudo calcular el resumen del periodo: {e}")
        return None


SHEETS_PARA_RESET = ["movimientos", "compras", "ventas", "boletas", "extras_y_perdidas", "ahorro", "proformas"]


def ejecutar_reset_mensual(usuario_nombre):
    mes_actual = datetime.now().strftime("%Y-%m")
    fecha_backup = ahora_iso()
    resumen = []

    for nombre_hoja in SHEETS_PARA_RESET:
        hoja = ws(nombre_hoja)
        filas = hoja.get_all_values()
        if not filas:
            continue
        encabezado, datos = filas[0], filas[1:]

        idx_fecha = 0
        for i, col in enumerate(encabezado):
            if "fecha" in col.lower() and "confirmacion" not in col.lower():
                idx_fecha = i
                break

        filas_mes = [f for f in datos if len(f) > idx_fecha and str(f[idx_fecha]).startswith(mes_actual)]
        filas_resto = [f for f in datos if not (len(f) > idx_fecha and str(f[idx_fecha]).startswith(mes_actual))]

        if not filas_mes:
            resumen.append(f"• {nombre_hoja}: sin datos de {mes_actual}")
            continue

        backup_json = json.dumps({"hoja": nombre_hoja, "encabezado": encabezado, "filas": filas_mes},
                                  ensure_ascii=False)
        ws("backups_mensuales").append_row(
            [fecha_backup, mes_actual, usuario_nombre, len(filas_mes), backup_json],
            value_input_option="USER_ENTERED",
        )

        hoja.clear()
        hoja.update([encabezado] + filas_resto, value_input_option="USER_ENTERED")
        resumen.append(f"• {nombre_hoja}: {len(filas_mes)} filas archivadas y borradas")

    return "Reset mensual completado (el inventario NO fue tocado):\n" + "\n".join(resumen)


INTENCIONES = {
    "reset": ["reset", "resetear"],
    "confirmar": ["confirmar", "confirma", "acepto", "aceptar"],
    "agregar": ["agregar", "anadir", "sumar", "mas", "mas"],
    "actualizar": ["actualizar", "actualiza", "cambiar", "cambia", "modificar", "modifica"],
    "recomendar": ["recomiendas", "recomienda", "recomendacion", "sugerencia", "sugieres", "sugiere"],
    "proforma": ["proforma", "cotizacion", "cotiza"],
    "ganancias": ["ganancias", "ganancia", "resumen", "caja", "balance"],
    "stock": ["stock", "inventario", "existencias"],
    "compra": ["compre", "compra", "comprar", "compre"],
    "venta": ["vendi", "vende", "vender", "vendi"],
    "precio": ["precio", "precios", "cuanto", "vale", "cuesta"],
    "ayuda": ["ayuda", "help", "comandos", "menu"],
    "alias": ["alias"],
    "stock_bajo": ["bajo", "alertas", "alerta"],
    "historial": ["historial", "historico"],
}

ORDEN_PRIORIDAD = ["reset", "confirmar", "agregar", "actualizar", "recomendar", "proforma", "ganancias",
                    "stock", "alias", "stock_bajo", "historial", "compra", "venta", "precio", "ayuda"]


# ==================== CORRECCIÓN 2: detectar_intencion() ====================
def detectar_intencion(texto_normalizado):
    tokens = texto_normalizado.split()
    encontradas = set()

    texto_lower = texto_normalizado.lower().strip()
    for intencion, palabras_clave in INTENCIONES.items():
        for palabra in palabras_clave:
            palabra_limpia = palabra.replace("\\s+", " ").strip()
            # CORRECCIÓN: usar word boundaries (\b) para evitar matching parcial
            # Ejemplo: "mas" no debe coincidir con "camas", "sumar" no con "luminar"
            if re.search(rf'\b{re.escape(palabra_limpia)}\b', texto_lower):
                encontradas.add(intencion)
                break

    for token in tokens:
        token = token.strip(",")
        if not token or token.isdigit() or len(token) < 4:
            continue
        for intencion, palabras_clave in INTENCIONES.items():
            if intencion in encontradas:
                continue
            if HAS_FUZZ:
                matches = process.extract(token, palabras_clave, limit=1)
                if matches and matches[0][1] >= 85:
                    encontradas.add(intencion)
            else:
                if difflib.get_close_matches(token, palabras_clave, n=1, cutoff=0.85):
                    encontradas.add(intencion)

    for intencion in ORDEN_PRIORIDAD:
        if intencion in encontradas:
            return intencion
    return None


# ==================== CORRECCIÓN 3: extraer_numeros() ====================
def extraer_numeros(texto):
    """Extrae números incluyendo decimales con punto o coma."""
    return [parse_decimal(n) for n in re.findall(r"\d+(?:[.,]\d+)?", texto)]


def quitar_palabras_intencion(texto_normalizado, lista_palabras):
    resultado = texto_normalizado
    for kw in lista_palabras:
        resultado = re.sub(rf"\b{kw}\w*\b", "", resultado)
    return re.sub(r"\s+", " ", resultado).strip()


def extraer_nombre_producto(texto_normalizado, palabras_a_quitar):
    todas_las_keywords = set()
    for lista in INTENCIONES.values():
        todas_las_keywords.update(lista)

    tokens = texto_normalizado.split()
    resultado = []
    for token in tokens:
        token = token.strip(",")
        if not token:
            continue
        if re.fullmatch(r"\d+(?:[.,]\d+)?", token):
            continue
        if token in STOPWORDS or token in palabras_a_quitar:
            continue
        if difflib.get_close_matches(token, todas_las_keywords, n=1, cutoff=0.78):
            continue
        resultado.append(token)
    return " ".join(resultado).strip()


# ==================== CORRECCIÓN 4: parsear_compra() ====================
def parsear_compra(texto_normalizado):
    resto = quitar_palabras_intencion(texto_normalizado, INTENCIONES["compra"])

    forzar_nuevo = False
    m_marca = re.match(r"^(n|nuevo|nueva)\b\s*(.*)$", resto)
    if m_marca:
        forzar_nuevo = True
        resto = m_marca.group(2).strip()

    costo = None
    precio_venta = None

    # Detectar c<numero> y v<numero> (acepta punto o coma decimal)
    m_c = re.search(r'\bc[:-]?\s*(\d+(?:[.,]\d+)?)', resto, re.IGNORECASE)
    m_v = re.search(r'\bv[:-]?\s*(\d+(?:[.,]\d+)?)', resto, re.IGNORECASE)
    if m_c:
        costo = parse_decimal(m_c.group(1))
    if m_v:
        precio_venta = parse_decimal(m_v.group(1))

    # Eliminar las etiquetas c<numero> y v<numero> del texto
    resto_sin_cv = re.sub(r'\b[cv][:-]?\s*\d+(?:[.,]\d+)?', '', resto, flags=re.IGNORECASE)
    resto_sin_cv = re.sub(r'\s+', ' ', resto_sin_cv).strip()

    # Dividir por comas (respetando espacios alrededor)
    partes_raw = re.split(r',\s*', resto_sin_cv)
    partes = [limpiar_segmento(p) for p in partes_raw if limpiar_segmento(p)]

    if not partes:
        return None

    primera = partes[0]

    # CORRECCIÓN: Intentar extraer "cantidad nombre a precio" de la primera parte
    m_cant_precio = re.match(r"^(\d+(?:[.,]\d+)?)\s+(.+?)\s+a\s+(\d+(?:[.,]\d+)?)$", primera)
    if m_cant_precio:
        cantidad = parse_decimal(m_cant_precio.group(1))
        nombre = m_cant_precio.group(2).strip()
        precio_en_a = parse_decimal(m_cant_precio.group(3))
        numeros_sueltos = [precio_en_a]
    else:
        m_cant = re.match(r"^(\d+(?:[.,]\d+)?)\s+(.+)$", primera)
        if m_cant:
            cantidad = parse_decimal(m_cant.group(1))
            nombre = m_cant.group(2).strip()
        else:
            cantidad = 1.0
            nombre = primera
        numeros_sueltos = []

    if not nombre:
        return None

    # Buscar números sueltos en las partes restantes (para costo y precio)
    for seg in partes[1:]:
        seg_limpio = seg.strip()
        m_num = re.match(r"^(\d+(?:[.,]\d+)?)$", seg_limpio)
        if m_num:
            numeros_sueltos.append(parse_decimal(m_num.group(1)))

    # Asignar costo y precio venta de los números sueltos
    if costo is None and numeros_sueltos:
        costo = numeros_sueltos[0] if len(numeros_sueltos) > 0 else None
    if precio_venta is None and len(numeros_sueltos) > 1:
        precio_venta = numeros_sueltos[1]

    return {
        "cantidad": cantidad,
        "nombre": nombre,
        "costo": costo,
        "precio_venta": precio_venta,
        "forzar_nuevo": forzar_nuevo
    }


# ==================== CORRECCIÓN 5: parsear_items_multiples() ====================
def parsear_items_multiples(resto):
    """Parsea texto como: '1 tubo a 12.30, 3 galletas a 9.60' o '1 tubo a 13, 3 spray'"""
    if not resto or not resto.strip():
        return []

    items = []
    # Separar por "y" o por coma
    resto = re.sub(r'\s+y\s+', '|||', resto)
    partes = re.split(r',\s*|\|\|\|', resto)

    for seg in partes:
        seg = limpiar_segmento(seg)
        if not seg:
            continue

        # Regex: ^(cantidad) (nombre) (?: a (precio))?$
        m = re.match(r"^(\d+(?:[.,]\d+)?)\s+(.+?)(?:\s+a\s+(\d+(?:[.,]\d+)?))?$", seg)
        if m:
            cantidad = parse_decimal(m.group(1))
            nombre = m.group(2).strip()
            precio = parse_decimal(m.group(3)) if m.group(3) else None
            items.append({"cantidad": cantidad, "nombre": nombre, "precio": precio})
        else:
            # Si no hay cantidad, intentar con "nombre a precio" (cantidad=1)
            m2 = re.match(r"^(.+?)(?:\s+a\s+(\d+(?:[.,]\d+)?))?$", seg)
            if m2:
                nombre = m2.group(1).strip()
                precio = parse_decimal(m2.group(2)) if m2.group(2) else None
                items.append({"cantidad": 1, "nombre": nombre, "precio": precio})

    return items


def manejar_historial(texto_norm, chat_id=None):
    nombre_producto = extraer_nombre_producto(texto_norm, INTENCIONES["historial"])
    if not nombre_producto:
        return "Historial de que producto? Ej: 'historial arboles'."

    productos = leer_inventario()
    estado, resultado = buscar_producto(nombre_producto, productos, umbral=0.75)

    if estado == "multiple":
        if chat_id is not None:
            CONTEXTO_SELECCION[chat_id] = {
                "productos": resultado,
                "callback": procesar_historial_seleccionado,
                "args": None,
            }
        return ("Encontre varias opciones, responde con el numero:\n" +
                formatear_lista_productos(resultado) +
                "\n\nEscribe el numero del producto que quieres.")
    if estado == "no_encontrado":
        return f"No encontre '{nombre_producto}'."

    p = resultado
    return _generar_historial_por_sku(p["SKU"], p["Nombre_completo"])


def manejar_ayuda():
    return (
        "Como hablarme (lenguaje natural, sin comandos):\n\n"
        "Compras:\n"
        "• 'compre 5 arboles a 3' -> compra 5 unid. a S/3\n"
        "• 'compre 5 arboles, 3, 10' -> compra (cant, costo, precio venta)\n"
        "• 'compre iphone 10, c8, v20' -> costo S/8, precio venta S/20\n"
        "• 'compre n iphone 10, c8, v20' -> fuerza crear producto nuevo\n\n"
        "Ventas:\n"
        "• 'vendi 2 arboles a 10' -> venta simple\n"
        "• 'vendi 1 laptop, 1 laptop pro' -> venta de varios productos\n"
        "• 'vendi 1 laptop a 15, 2 arboles a 12' -> venta con precios personalizados\n\n"
        "Consultas:\n"
        "• 'stock' -> todo el inventario\n"
        "• 'precio de arboles' -> precio y stock (si hay varios, los lista)\n"
        "• 'arbol' -> busca directamente por nombre\n\n"
        "Proformas:\n"
        "• 'proforma 2 arboles, 3 galletas' -> cotizacion\n"
        "• 'agregar 1 cel, 3 codos' -> anade a la proforma actual\n"
        "• 'confirmar' / 'confirmar 60' / 'confirmar 57 60' -> registra la "
        "proforma como venta (con pago y vuelto)\n\n"
        "Actualizar:\n"
        "• 'actualizar arboles precio 12' -> cambia precio\n"
        "• 'actualizar arboles stock 8' -> fija el stock exacto\n"
        "• 'actualizar arboles costo 4 precio 18' -> cambia costo y precio\n\n"
        "Recomendacion:\n"
        "• 'que recomiendas para arboles' -> sugerencia de precio\n\n"
        "Ganancias:\n"
        "• 'ganancias' -> resumen del dia\n\n"
        "Reset:\n"
        "• '/reset <contrasena>' -> archiva y limpia el mes actual (no toca inventario)\n"
    )


async def handle_seleccion_numero(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    texto = update.message.text.strip()

    if chat_id not in CONTEXTO_SELECCION:
        return False

    try:
        idx = int(texto) - 1
        datos = CONTEXTO_SELECCION[chat_id]
        if 0 <= idx < len(datos["productos"]):
            producto = datos["productos"][idx]
            await datos["callback"](update, context, producto, datos["args"])
            if "cola_venta" not in datos and "cola_proforma" not in datos:
                del CONTEXTO_SELECCION[chat_id]
            return True
    except Exception:
        pass
    return False


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    inicio = time.time()
    texto = update.message.text or ""
    usuario = update.effective_user
    usuario_nombre = usuario.username or usuario.first_name or "desconocido"
    usuario_id = usuario.id
    chat_id = update.effective_chat.id

    if ACCESS_PASSWORD is not None:
        if chat_id not in USUARIOS_AUTENTICADOS or not USUARIOS_AUTENTICADOS[chat_id]:
            password_ingresada = texto.strip()
            if password_ingresada == ACCESS_PASSWORD:
                USUARIOS_AUTENTICADOS[chat_id] = True
                respuesta = (
                    "Acceso concedido\n\n"
                    "Bienvenido! Soy tu asistente de inventario.\n\n"
                    "Como hablarme:\n"
                    "Escribe naturalmente, tolero errores de tipeo y plurales.\n\n"
                    "Ejemplos rapidos:\n"
                    "• 'compre 5 arboles a 3' -> compra\n"
                    "• 'vendi 2 arboles a 10' -> venta\n"
                    "• 'stock' -> ver inventario\n"
                    "• 'precio de arboles' -> consultar precio\n"
                    "• 'proforma 4 cel, 3 arbol' -> cotizar\n"
                    "• 'ganancias' -> resumen del dia\n\n"
                    "Escribe 'ayuda' para ver todos los comandos."
                )
                await update.message.reply_text(respuesta, parse_mode=None)
                return
            else:
                await update.message.reply_text(
                    "Acceso restringido\n\n"
                    "Ingresa la contrasena de acceso para continuar.",
                    parse_mode=None
                )
                return

    if chat_id in CONTEXTO_SELECCION and texto.strip().isdigit():
        if await handle_seleccion_numero(update, context):
            return

    texto_norm = normalizar(texto)
    intencion = detectar_intencion(texto_norm)
    origen = "keywords"

    if intencion == "agregar":
        respuesta = await asyncio.to_thread(manejar_proforma, texto_norm, usuario_nombre, chat_id)
        await update.message.reply_text(respuesta, parse_mode=None)
        return

    datos_groq = None
    if intencion is None:
        productos_rapido = await asyncio.to_thread(leer_inventario)
        estado_rapido, _ = buscar_producto(texto_norm, productos_rapido)
        if estado_rapido in ("ok", "multiple"):
            intencion = "precio"
            origen = "match_directo"

    if intencion is None:
        datos_groq = await asyncio.to_thread(groq_interpretar, texto)
        intencion = datos_groq.get("intencion")
        origen = "groq"
        if intencion == "desconocido" or not intencion:
            intencion = None

    try:
        if intencion == "confirmar":
            respuesta = await asyncio.to_thread(manejar_confirmar, texto_norm, chat_id, usuario_nombre)

        elif intencion == "compra":
            respuesta = await asyncio.to_thread(manejar_compra, texto_norm, usuario_nombre, chat_id)
            await asyncio.to_thread(actualizar_hoja_ganancias)

        elif intencion == "venta":
            respuesta = await asyncio.to_thread(manejar_venta, texto_norm, usuario_nombre, chat_id)
            if respuesta is None:
                if chat_id in CONTEXTO_SELECCION and "cola_venta" in CONTEXTO_SELECCION[chat_id]:
                    await _continuar_venta_pendiente(update, context, chat_id, usuario_nombre)
                return

        elif intencion == "precio":
            respuesta = await asyncio.to_thread(manejar_precio, texto_norm, chat_id)

        elif intencion == "stock":
            respuesta = await asyncio.to_thread(manejar_stock)

        elif intencion == "proforma":
            respuesta = await asyncio.to_thread(manejar_proforma, texto_norm, usuario_nombre, chat_id)
            if respuesta is None:
                if chat_id in CONTEXTO_SELECCION and "cola_proforma" in CONTEXTO_SELECCION[chat_id]:
                    await _continuar_proforma_pendiente(update, context, chat_id, 
                        CONTEXTO_SELECCION[chat_id]["cola_proforma"], usuario_nombre,
                        CONTEXTO_SELECCION[chat_id]["items_json_acum"],
                        CONTEXTO_SELECCION[chat_id]["subtotal_acum"],
                        CONTEXTO_SELECCION[chat_id]["detalle_acum"])
                return

        elif intencion == "ganancias":
            respuesta = await asyncio.to_thread(manejar_ganancias, texto_norm)

        elif intencion == "actualizar":
            respuesta = await asyncio.to_thread(manejar_actualizar, texto_norm, usuario_nombre, chat_id)

        elif intencion == "recomendar":
            respuesta = await asyncio.to_thread(manejar_recomendar, texto_norm)

        elif intencion == "alias":
            respuesta = await asyncio.to_thread(manejar_alias, texto, usuario_nombre)

        elif intencion == "stock_bajo":
            respuesta = await asyncio.to_thread(manejar_stock_bajo)

        elif intencion == "historial":
            respuesta = await asyncio.to_thread(manejar_historial, texto_norm, chat_id)

        elif intencion == "ayuda":
            respuesta = manejar_ayuda()

        elif intencion == "reset":
            respuesta = "Para resetear el mes usa el comando: /reset <contrasena>"

        else:
            respuesta = "No entendi tu mensaje. Escribe 'ayuda' para ver ejemplos."

    except Exception as e:
        log.exception("Error procesando mensaje")
        respuesta = f"Ocurrio un error procesando tu mensaje: {e}"

    await update.message.reply_text(respuesta, parse_mode=None)

    tiempo_total = time.time() - inicio
    asyncio.create_task(
        asyncio.to_thread(registrar_log, usuario_id, usuario_nombre, intencion or "desconocido",
                           texto, respuesta, tiempo_total)
    )


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ACCESS_PASSWORD is None:
        await update.message.reply_text(
            "Bienvenido al Bot de Inventario Lumin!\n\n"
            "Escribe 'ayuda' para ver como usarme.",
            parse_mode=None
        )
        return
    await update.message.reply_text(
        "Bot de Inventario Lumin\n\n"
        "Ingresa la contrasena de acceso para continuar.",
        parse_mode=None
    )


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(manejar_ayuda(), parse_mode=None)


async def cmd_reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or args[0] != RESET_PASSWORD:
        await update.message.reply_text("Contrasena incorrecta o faltante. Uso: /reset <contrasena>")
        return
    usuario = update.effective_user
    usuario_nombre = usuario.username or usuario.first_name or "desconocido"
    respuesta = await asyncio.to_thread(ejecutar_reset_mensual, usuario_nombre)
    await update.message.reply_text(respuesta, parse_mode=None)


# ==================== WEBHOOK + FLASK ====================
app = None
loop = None

flask_app = Flask(__name__)

@flask_app.route('/')
def home():
    return "Bot de Inventario Lumin funcionando con Webhook ✅"

@flask_app.route('/ping')
def ping():
    return "pong", 200

@flask_app.route('/webhook', methods=['POST'])
def webhook():
    global loop
    try:
        json_data = request.get_json(force=True)
        update = Update.de_json(json_data, app.bot)

        future = asyncio.run_coroutine_threadsafe(
            app.process_update(update),
            loop
        )
        future.result(timeout=30)

        return 'OK', 200
    except Exception as e:
        log.error(f"Error en webhook: {e}")
        return 'Error', 500

def run_bot():
    global app, loop

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler(["ayuda", "help"], cmd_ayuda))
    app.add_handler(CommandHandler("reset", cmd_reset))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    loop.run_until_complete(app.initialize())
    loop.run_until_complete(app.start())

    webhook_endpoint = f"{WEBHOOK_URL}/webhook"
    log.info(f"Configurando webhook en: {webhook_endpoint}")
    loop.run_until_complete(app.bot.set_webhook(
        url=webhook_endpoint,
        drop_pending_updates=True
    ))
    log.info("✅ Webhook configurado correctamente.")
    log.info("✅ Bot listo y esperando mensajes via webhook...")

    loop.run_forever()

def main():
    bot_thread = threading.Thread(target=run_bot, daemon=True)
    bot_thread.start()

    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port)

if __name__ == "__main__":
    import threading
    main()