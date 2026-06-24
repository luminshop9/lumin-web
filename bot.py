#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bot de Inventario Lumin - VERSIÓN WEBHOOK PARA RENDER
Fecha: 2026-06-23
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

# ==================== CONFIGURACIÓN (desde variables de entorno) ====================
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CREDS_FILE = os.getenv("CREDS_FILE", "creds_nuevo.json")
RESET_PASSWORD = os.getenv("RESET_PASSWORD")
ACCESS_PASSWORD = os.getenv("ACCESS_PASSWORD")
WEBHOOK_URL = os.getenv("WEBHOOK_URL")  # ej: https://lumin-bot-deploy.onrender.com

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


def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).lower().strip()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    texto = re.sub(r"[^a-z0-9\s/,]", " ", texto)
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
    # CORRECCION 1: Alias exacto primero (case-insensitive)
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

    # CORRECCION 3: umbral por defecto subido a 0.7
    if mejor_score < umbral:
        return "no_encontrado", None

    # CORRECCION 2: UMBRAL_MULTIPLE bajado a 0.35
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


def detectar_intencion(texto_normalizado):
    tokens = texto_normalizado.split()
    encontradas = set()
    
    texto_lower = texto_normalizado.lower().strip()
    for intencion, palabras_clave in INTENCIONES.items():
        for palabra in palabras_clave:
            palabra_limpia = palabra.replace("\\s+", " ").strip()
            if palabra_limpia in texto_lower:
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


def extraer_numeros(texto):
    texto_sin_comas = texto.replace(",", " ")
    return [parse_decimal(n) for n in re.findall(r"\d+(?:\.\d+)?", texto_sin_comas)]


def quitar_palabras_intencion(texto_normalizado, lista_palabras):
    resultado = texto_normalizado
    for kw in lista_palabras:
        resultado = re.sub(rf"\b{kw}\w*\b", "", resultado)
    return resultado.strip()


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
        if re.fullmatch(r"\d+(\.\d+)?", token):
            continue
        if token in STOPWORDS or token in palabras_a_quitar:
            continue
        if difflib.get_close_matches(token, todas_las_keywords, n=1, cutoff=0.78):
            continue
        resultado.append(token)
    return " ".join(resultado).strip()


# CORRECCION 9 (+ coma decimal): Mejorar parsear_compra para extraer nombre con numeros sueltos

def parsear_compra(texto_normalizado):
    resto = quitar_palabras_intencion(texto_normalizado, INTENCIONES["compra"])

    forzar_nuevo = False
    m_marca = re.match(r"^(n|nuevo|nueva)\b\s*(.*)$", resto)
    if m_marca:
        forzar_nuevo = True
        resto = m_marca.group(2).strip()

    costo = None
    precio_venta = None

    # Detectar c<numero> y v<numero> (con o sin espacio después de c/v)
    m_c = re.search(r'\bc[:\-]?\s*(\d+(?:[.,]\d+)?)', resto, re.IGNORECASE)
    m_v = re.search(r'\bv[:\-]?\s*(\d+(?:[.,]\d+)?)', resto, re.IGNORECASE)
    if m_c:
        costo = parse_decimal(m_c.group(1))
    if m_v:
        precio_venta = parse_decimal(m_v.group(1))

    # Eliminar las etiquetas c<numero> y v<numero> del texto
    resto_sin_cv = re.sub(r'\b[cv][:\-]?\s*\d+(?:[.,]\d+)?', '', resto, flags=re.IGNORECASE)
    resto_sin_cv = re.sub(r'\s+', ' ', resto_sin_cv).strip()

    # Dividir por comas (pero no si la coma está dentro de un número)
    # Usamos una separación más inteligente: dividir por coma, pero respetando que los números no se corten
    partes_raw = re.split(r',\s*', resto_sin_cv)
    partes = [limpiar_segmento(p) for p in partes_raw if limpiar_segmento(p)]

    if not partes:
        return None

    primera = partes[0]
    # Intentar extraer "cantidad nombre" de la primera parte
    m_cant = re.match(r"^(\d+(?:\.\d+)?)\s+(.+)$", primera)
    if m_cant:
        cantidad = parse_decimal(m_cant.group(1))
        nombre = m_cant.group(2).strip()
    else:
        # Si no hay cantidad, tomar todo como nombre
        cantidad = 1.0
        nombre = primera

    if not nombre:
        return None

    # Buscar números sueltos en las partes restantes (para costo y precio)
    numeros_sueltos = []
    for seg in partes[1:]:
        seg_limpio = seg.strip()
        # Intentar extraer número puro
        m_num = re.match(r"^(\d+(?:\.\d+)?)$", seg_limpio)
        if m_num:
            numeros_sueltos.append(parse_decimal(m_num.group(1)))

    # Asignar costo y precio venta de los números sueltos (orden: primero costo, luego precio)
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


def parsear_items_multiples(resto):
    # Separar por " y " o por coma
    # Primero reemplazar " y " por un separador único
    resto = re.sub(r'\s+y\s+', '|||', resto)
    # También separar por coma
    partes = re.split(r',\s*|\|\|\|', resto)
    items = []
    for seg in partes:
        seg = limpiar_segmento(seg)
        if not seg:
            continue
        # Intentar extraer "cantidad nombre a precio"
        m = re.match(r"^(\d+(?:\.\d+)?)\s+(.+?)(?:\s+a\s+(\d+(?:\.\d+)?))?$", seg)
        if m:
            cantidad = parse_decimal(m.group(1))
            nombre = m.group(2).strip()
            precio = parse_decimal(m.group(3)) if m.group(3) else None
            items.append({"cantidad": cantidad, "nombre": nombre, "precio": precio})
    return items


def groq_interpretar(texto):
    prompt = f"""Eres el clasificador de un bot de inventario para un negocio peruano.
El personal escribe rapido y con errores. Clasifica este mensaje:

"{texto}"

Responde UNICAMENTE un JSON valido, sin texto adicional, con este formato exacto:
{{"intencion": "compra|venta|precio|stock|proforma|ganancias|actualizar|recomendar|ayuda|desconocido",
  "producto": "nombre del producto mencionado o vacio",
  "cantidad": numero o null,
  "costo": numero o null,
  "precio": numero o null}}"""
    try:
        resp = groq_client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=200,
        )
        texto_resp = resp.choices[0].message.content.strip()
        texto_resp = re.sub(r"^```json|```$", "", texto_resp.strip(), flags=re.MULTILINE).strip()
        return json.loads(texto_resp)
    except Exception as e:
        log.warning(f"Groq interpretar fallo: {e}")
        return {"intencion": "desconocido", "producto": "", "cantidad": None, "costo": None, "precio": None}


def groq_recomendar_precio(producto, margenes_categoria):
    costo = producto["Costo"]
    categoria = producto.get("Categoria", "General")
    margen_info = ""
    if margenes_categoria:
        margen_info = (
            f"Para la categoria '{categoria}' el margen sugerido suele ser entre "
            f"{margenes_categoria['Margen_minimo']*100:.0f}% y {margenes_categoria['Margen_maximo']*100:.0f}%."
        )
    prompt = f"""Producto: {producto['Nombre_completo']}
Costo: S/{costo:.2f}
Precio actual de venta: S/{producto['Precio_venta_actual']:.2f}
{margen_info}

Da una recomendacion de precio de venta MUY breve y directa (maximo 2 lineas,
en espanol peruano informal), incluyendo el precio sugerido en soles."""
    try:
        resp = groq_client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4,
            max_tokens=120,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        log.warning(f"Groq recomendar fallo: {e}")
        sugerido = round(costo * 1.6, 2)
        return f"Sugerido: S/{sugerido:.2f} (margen 60% sobre costo S/{costo:.2f})."


async def procesar_compra_seleccionada(update, context, producto, args):
    usuario_nombre = args[0]
    datos = args[1]

    cantidad = parse_int(datos["cantidad"])
    costo = datos["costo"]
    precio_venta = datos["precio_venta"]
    forzar_nuevo = datos["forzar_nuevo"]

    if forzar_nuevo:
        nuevo = crear_producto(producto["Nombre_base"], costo, cantidad, precio_venta, usuario_nombre)
        registrar_movimiento(nuevo["SKU"], "compra", cantidad, costo, usuario_nombre, "nuevo (forzado)")
        registrar_compra(nuevo["SKU"], cantidad, costo, usuario_nombre)
        respuesta = (f"🛒 Producto nuevo creado: 📦 {nuevo['Nombre_completo']} (SKU {nuevo['SKU']})\n"
                     f"Stock: {cantidad} | Costo: {fmt_money(costo)} | "
                     f"Precio venta: {fmt_money(nuevo['Precio_venta_actual'])}")
        await update.message.reply_text(respuesta, parse_mode=None)
        return

    nuevo_stock = parse_int(producto["Stock_actual"]) + cantidad
    actualizar_celda("inventario", producto["_row"], "Stock_actual", nuevo_stock)

    costo_final = costo if costo is not None else producto["Costo"]
    if costo is not None:
        actualizar_celda("inventario", producto["_row"], "Costo", costo_final)

    precio_final = precio_venta if precio_venta is not None else producto["Precio_venta_actual"]
    if precio_venta is not None:
        actualizar_celda("inventario", producto["_row"], "Precio_venta_actual", precio_final)

    margen = round((precio_final - costo_final) / costo_final, 6) if costo_final else 0
    actualizar_celda("inventario", producto["_row"], "Margen", margen)
    actualizar_celda("inventario", producto["_row"], "Estado_stock",
                     estado_stock(nuevo_stock, producto["Stock_minimo"]))
    actualizar_celda("inventario", producto["_row"], "Ultima_actualizacion", ahora_iso())
    actualizar_celda("inventario", producto["_row"], "Actualizado_por", usuario_nombre)

    registrar_movimiento(producto["SKU"], "compra", cantidad, costo_final, usuario_nombre)
    registrar_compra(producto["SKU"], cantidad, costo_final, usuario_nombre)

    respuesta = (f"🛒 Compra registrada: {cantidad} x 📦 {producto['Nombre_completo']} a {fmt_money(costo_final)}\n"
                 f"Stock nuevo: {nuevo_stock} | Precio venta: {fmt_money(precio_final)}")
    await update.message.reply_text(respuesta, parse_mode=None)


def manejar_compra(texto_norm, usuario_nombre, chat_id):
    datos = parsear_compra(texto_norm)
    if not datos:
        return ("No entendi bien. Usa: 'compre 5 arboles a 3', 'compre 5 arboles, 3, 10' "
                "o 'compre iphone 10, c8, v20' (c=costo, v=precio venta).")

    cantidad = parse_int(datos["cantidad"])
    nombre_producto = datos["nombre"]
    costo = datos["costo"]
    precio_venta = datos["precio_venta"]
    forzar_nuevo = datos["forzar_nuevo"]

    productos = leer_inventario()

    if forzar_nuevo:
        if costo is None:
            return f"Para crear '{nombre_producto}' como producto nuevo necesito el costo (ej: c8)."
        nuevo = crear_producto(nombre_producto, costo, cantidad, precio_venta, usuario_nombre)
        registrar_movimiento(nuevo["SKU"], "compra", cantidad, costo, usuario_nombre, "nuevo (forzado)")
        registrar_compra(nuevo["SKU"], cantidad, costo, usuario_nombre)
        return (f"🛒 Producto nuevo creado: 📦 {nuevo['Nombre_completo']} (SKU {nuevo['SKU']})\n"
                f"Stock: {cantidad} | Costo: {fmt_money(costo)} | "
                f"Precio venta: {fmt_money(nuevo['Precio_venta_actual'])}")

    estado, resultado = buscar_producto(nombre_producto, productos)

    if estado == "multiple":
        CONTEXTO_SELECCION[chat_id] = {
            "productos": resultado,
            "callback": procesar_compra_seleccionada,
            "args": (usuario_nombre, datos)
        }
        return ("Encontre varias opciones, responde con el numero:\n" +
                formatear_lista_productos(resultado) +
                "\n\nEscribe el numero del producto que quieres.")

    # CORRECCION 7: Crear producto automaticamente sin "n" cuando se da costo y precio
    if estado == "no_encontrado":
        if costo is not None and precio_venta is not None:
            nuevo = crear_producto(nombre_producto, costo, cantidad, precio_venta, usuario_nombre)
            registrar_movimiento(nuevo["SKU"], "compra", cantidad, costo, usuario_nombre, "producto nuevo")
            registrar_compra(nuevo["SKU"], cantidad, costo, usuario_nombre)
            return (f"🛒 Producto nuevo creado: 📦 {nuevo['Nombre_completo']} (SKU {nuevo['SKU']})\n"
                    f"Stock: {cantidad} | Costo: {fmt_money(costo)} | "
                    f"Precio venta: {fmt_money(nuevo['Precio_venta_actual'])}")
        if costo is None:
            return (f"'{nombre_producto}' no existe todavia. Dime el costo para crearlo: "
                    f"'compre {cantidad} {nombre_producto}, c<costo>' o agrega 'nuevo' si ya existe algo parecido.")
        # costo existe pero precio_venta es None -> crear con precio sugerido
        nuevo = crear_producto(nombre_producto, costo, cantidad, None, usuario_nombre)
        registrar_movimiento(nuevo["SKU"], "compra", cantidad, costo, usuario_nombre, "producto nuevo")
        registrar_compra(nuevo["SKU"], cantidad, costo, usuario_nombre)
        return (f"🛒 Producto nuevo creado: 📦 {nuevo['Nombre_completo']} (SKU {nuevo['SKU']})\n"
                f"Stock: {cantidad} | Costo: {fmt_money(costo)} | "
                f"Precio venta: {fmt_money(nuevo['Precio_venta_actual'])}")

    producto = resultado
    nuevo_stock = parse_int(producto["Stock_actual"]) + cantidad
    actualizar_celda("inventario", producto["_row"], "Stock_actual", nuevo_stock)

    costo_final = costo if costo is not None else producto["Costo"]
    if costo is not None:
        actualizar_celda("inventario", producto["_row"], "Costo", costo_final)

    precio_final = precio_venta if precio_venta is not None else producto["Precio_venta_actual"]
    if precio_venta is not None:
        actualizar_celda("inventario", producto["_row"], "Precio_venta_actual", precio_final)

    margen = round((precio_final - costo_final) / costo_final, 6) if costo_final else 0
    actualizar_celda("inventario", producto["_row"], "Margen", margen)
    actualizar_celda("inventario", producto["_row"], "Estado_stock",
                     estado_stock(nuevo_stock, producto["Stock_minimo"]))
    actualizar_celda("inventario", producto["_row"], "Ultima_actualizacion", ahora_iso())
    actualizar_celda("inventario", producto["_row"], "Actualizado_por", usuario_nombre)

    registrar_movimiento(producto["SKU"], "compra", cantidad, costo_final, usuario_nombre)
    registrar_compra(producto["SKU"], cantidad, costo_final, usuario_nombre)

    return (f"🛒 Compra registrada: {cantidad} x 📦 {producto['Nombre_completo']} a {fmt_money(costo_final)}\n"
            f"Stock nuevo: {nuevo_stock} | Precio venta: {fmt_money(precio_final)}")


def _vender_un_item(producto, cantidad, precio_real, usuario_nombre):
    stock_actual = parse_int(producto["Stock_actual"])
    if cantidad > stock_actual:
        return (f"⚠️ Stock insuficiente de 📦 {producto['Nombre_completo']} "
                f"(disponible {stock_actual}, pedido {cantidad}). No se vendio.", 0.0)

    nuevo_stock = stock_actual - cantidad
    actualizar_celda("inventario", producto["_row"], "Stock_actual", nuevo_stock)
    actualizar_celda("inventario", producto["_row"], "Estado_stock",
                     estado_stock(nuevo_stock, producto["Stock_minimo"]))
    registrar_movimiento(producto["SKU"], "venta", cantidad, precio_real, usuario_nombre)
    ganancia = registrar_venta_y_boleta(producto, cantidad, precio_real, usuario_nombre)

    linea = f"{cantidad} x 📦 {producto['Nombre_completo']} a {fmt_money(precio_real)} (stock: {nuevo_stock})"
    diferencia = round(precio_real - producto["Precio_venta_actual"], 2)
    if diferencia > 0:
        linea += f" 💚 extra {fmt_money(diferencia * cantidad)}"
    elif diferencia < 0:
        linea += f" 📉 descuento {fmt_money(abs(diferencia) * cantidad)}"
    return linea, ganancia


# CORRECCION 4: Funciones auxiliares para seleccion individual en ventas y proformas

async def procesar_venta_seleccionada(update, context, producto, args):
    """Callback cuando el usuario selecciona un producto de una lista ambigua en venta."""
    usuario_nombre = args[0]
    item = args[1]
    chat_id = update.effective_chat.id
    
    precio_real = item["precio"] if item["precio"] is not None else producto["Precio_venta_actual"]
    cantidad = parse_int(item["cantidad"])
    linea, ganancia = _vender_un_item(producto, cantidad, precio_real, usuario_nombre)
    
    if chat_id not in CONTEXTO_SELECCION:
        CONTEXTO_SELECCION[chat_id] = {"resultados_venta": []}
    if "resultados_venta" not in CONTEXTO_SELECCION[chat_id]:
        CONTEXTO_SELECCION[chat_id]["resultados_venta"] = []
    CONTEXTO_SELECCION[chat_id]["resultados_venta"].append((linea, ganancia))
    
    await _continuar_venta_pendiente(update, context, chat_id, usuario_nombre)


async def _continuar_venta_pendiente(update, context, chat_id, usuario_nombre):
    """Continua procesando la cola de items de venta pendientes."""
    if chat_id not in CONTEXTO_SELECCION or "cola_venta" not in CONTEXTO_SELECCION[chat_id]:
        if chat_id in CONTEXTO_SELECCION and "resultados_venta" in CONTEXTO_SELECCION[chat_id]:
            resultados = CONTEXTO_SELECCION[chat_id]["resultados_venta"]
            lineas = [r[0] for r in resultados]
            ganancia_total = sum(r[1] for r in resultados)
            actualizar_hoja_ganancias()
            respuesta = "💰 Venta registrada:\n" + "\n".join(lineas) + f"\n\n📈 Ganancia total: {fmt_money(ganancia_total)}"
            await update.message.reply_text(respuesta, parse_mode=None)
            if chat_id in CONTEXTO_SELECCION:
                del CONTEXTO_SELECCION[chat_id]
        return
    
    cola = CONTEXTO_SELECCION[chat_id]["cola_venta"]
    if not cola:
        if "resultados_venta" in CONTEXTO_SELECCION[chat_id]:
            resultados = CONTEXTO_SELECCION[chat_id]["resultados_venta"]
            lineas = [r[0] for r in resultados]
            ganancia_total = sum(r[1] for r in resultados)
            actualizar_hoja_ganancias()
            respuesta = "💰 Venta registrada:\n" + "\n".join(lineas) + f"\n\n📈 Ganancia total: {fmt_money(ganancia_total)}"
            await update.message.reply_text(respuesta, parse_mode=None)
            if chat_id in CONTEXTO_SELECCION:
                del CONTEXTO_SELECCION[chat_id]
        return
    
    item = cola.pop(0)
    nombre_limpio = limpiar_nombre_completo(item["nombre"])
    productos = leer_inventario()
    
    # CORRECCION 5: Mejorar busqueda de "codo" en ventas
    estado, resultado = buscar_producto(nombre_limpio, productos, umbral=0.7)
    
    if estado == "no_encontrado":
        nombre_singular = singularizar(nombre_limpio)
        if nombre_singular != nombre_limpio:
            estado, resultado = buscar_producto(nombre_singular, productos, umbral=0.7)
        if estado == "no_encontrado":
            estado, resultado = buscar_producto(nombre_limpio, productos, umbral=0.4)
        if estado == "no_encontrado":
            nombre_norm = normalizar(nombre_limpio)
            for p in productos:
                if normalizar(p["Nombre_base"]) == nombre_norm:
                    estado, resultado = "ok", p
                    break
    
    if estado == "multiple":
        CONTEXTO_SELECCION[chat_id]["cola_venta"] = cola
        CONTEXTO_SELECCION[chat_id]["callback"] = procesar_venta_seleccionada
        CONTEXTO_SELECCION[chat_id]["args"] = (usuario_nombre, item)
        CONTEXTO_SELECCION[chat_id]["productos"] = resultado
        respuesta = (f"Encontre varias opciones para '{item['nombre']}', responde con el numero:\n" +
                    formatear_lista_productos(resultado) +
                    "\n\nEscribe el numero del producto que quieres.")
        await update.message.reply_text(respuesta, parse_mode=None)
        return
    
    if estado == "no_encontrado":
        if "resultados_venta" not in CONTEXTO_SELECCION[chat_id]:
            CONTEXTO_SELECCION[chat_id]["resultados_venta"] = []
        CONTEXTO_SELECCION[chat_id]["resultados_venta"].append((f"No encontre '{item['nombre']}'.", 0.0))
        CONTEXTO_SELECCION[chat_id]["cola_venta"] = cola
        await _continuar_venta_pendiente(update, context, chat_id, usuario_nombre)
        return
    
    producto = resultado
    precio_real = item["precio"] if item["precio"] is not None else producto["Precio_venta_actual"]
    cantidad = parse_int(item["cantidad"])
    linea, ganancia = _vender_un_item(producto, cantidad, precio_real, usuario_nombre)
    
    if "resultados_venta" not in CONTEXTO_SELECCION[chat_id]:
        CONTEXTO_SELECCION[chat_id]["resultados_venta"] = []
    CONTEXTO_SELECCION[chat_id]["resultados_venta"].append((linea, ganancia))
    CONTEXTO_SELECCION[chat_id]["cola_venta"] = cola
    await _continuar_venta_pendiente(update, context, chat_id, usuario_nombre)


def manejar_venta(texto_norm, usuario_nombre, chat_id):
    resto = quitar_palabras_intencion(texto_norm, INTENCIONES["venta"])
    items = parsear_items_multiples(resto)
    if not items:
        return "Usa: 'vendi 2 arboles a 10' o 'vendi 1 laptop, 1 laptop pro'."

    productos = leer_inventario()
    lineas = []
    ganancia_total = 0.0
    cola_pendientes = []
    
    for item in items:
        nombre_limpio = limpiar_nombre_completo(item["nombre"])
        
        # CORRECCION 5: Mejorar busqueda de "codo" en ventas
        estado, resultado = buscar_producto(nombre_limpio, productos, umbral=0.7)
        
        if estado == "no_encontrado":
            nombre_singular = singularizar(nombre_limpio)
            if nombre_singular != nombre_limpio:
                estado, resultado = buscar_producto(nombre_singular, productos, umbral=0.7)
            if estado == "no_encontrado":
                estado, resultado = buscar_producto(nombre_limpio, productos, umbral=0.4)
            if estado == "no_encontrado":
                nombre_norm = normalizar(nombre_limpio)
                for p in productos:
                    if normalizar(p["Nombre_base"]) == nombre_norm:
                        estado, resultado = "ok", p
                        break
        
        if estado == "multiple":
            cola_pendientes.append(item)
            continue
        
        if estado == "no_encontrado":
            lineas.append(f"No encontre '{item['nombre']}'.")
            continue

        producto = resultado
        precio_real = item["precio"] if item["precio"] is not None else producto["Precio_venta_actual"]
        cantidad = parse_int(item["cantidad"])
        linea, ganancia = _vender_un_item(producto, cantidad, precio_real, usuario_nombre)
        lineas.append(linea)
        ganancia_total += ganancia
    
    if cola_pendientes:
        CONTEXTO_SELECCION[chat_id] = {
            "cola_venta": cola_pendientes,
            "resultados_venta": [(l, g) for l, g in zip(lineas, [ganancia_total] * len(lineas)) if l] if lineas else [],
            "callback": procesar_venta_seleccionada,
            "args": (usuario_nombre, cola_pendientes[0]),
            "productos": None,
        }
        return None
    
    actualizar_hoja_ganancias()
    return "💰 Venta registrada:\n" + "\n".join(lineas) + f"\n\n📈 Ganancia total: {fmt_money(ganancia_total)}"


async def procesar_precio_seleccionado(update, context, producto, args):
    p = producto
    respuesta = (f"{p['Nombre_completo']} (SKU {p['SKU']})\n"
                 f"Precio venta: {fmt_money(p['Precio_venta_actual'])}\n"
                 f"Costo: {fmt_money(p['Costo'])}\n"
                 f"Stock: {parse_int(p['Stock_actual'])} {p['Estado_stock']}")
    await update.message.reply_text(respuesta, parse_mode=None)


def manejar_precio(texto_norm, chat_id=None):
    nombre_producto = extraer_nombre_producto(texto_norm, INTENCIONES["precio"])
    if not nombre_producto:
        return "Precio de que producto? Ej: 'precio de arboles'."

    productos = leer_inventario()
    estado, resultado = buscar_producto(nombre_producto, productos)

    if estado == "multiple":
        if chat_id is not None:
            CONTEXTO_SELECCION[chat_id] = {
                "productos": resultado,
                "callback": procesar_precio_seleccionado,
                "args": None,
            }
        return ("Encontre varias opciones, responde con el numero:\n" +
                formatear_lista_productos(resultado) +
                "\n\nEscribe el numero del producto que quieres.")
    if estado == "no_encontrado":
        return f"No encontre '{nombre_producto}' en el inventario."

    p = resultado
    return (f"{p['Nombre_completo']} (SKU {p['SKU']})\n"
            f"Precio venta: {fmt_money(p['Precio_venta_actual'])}\n"
            f"Costo: {fmt_money(p['Costo'])}\n"
            f"Stock: {parse_int(p['Stock_actual'])} {p['Estado_stock']}")


def manejar_stock():
    productos = leer_inventario()
    if not productos:
        return "No hay productos registrados todavia."
    lineas = ["Inventario actual:"]
    for p in sorted(productos, key=lambda x: x["Nombre_completo"]):
        lineas.append(
            f"• {p['Nombre_completo']}: {parse_int(p['Stock_actual'])} unid. "
            f"@ {fmt_money(p['Precio_venta_actual'])} {p['Estado_stock']}"
        )
    return "\n".join(lineas)


def manejar_stock_bajo():
    productos = leer_inventario()
    bajos = [p for p in productos if p.get("Estado_stock") in ("Bajo", "Agotado")]

    if not bajos:
        return "✅ Todos los productos tienen stock suficiente."

    bajos = sorted(bajos, key=lambda x: parse_int(x["Stock_actual"]))
    lineas = ["⚠️ Productos con stock bajo:"]
    for p in bajos:
        alias = (p.get("Alias") or "").strip()
        alias_fmt = f" (alias: {alias})" if alias else ""
        lineas.append(
            f"• {p['Nombre_completo']}{alias_fmt} - Stock: {parse_int(p['Stock_actual'])} "
            f"(minimo {parse_int(p['Stock_minimo'])})"
        )
    return "\n".join(lineas)


def calcular_descuento(subtotal):
    if subtotal > 300:
        return 0.10
    if subtotal > 100:
        return 0.08
    if subtotal > 50:
        return 0.05
    return 0.0


def manejar_proforma(texto_norm, usuario_nombre, chat_id):
    es_agregar = detectar_intencion(texto_norm) == "agregar"

    if es_agregar:
        pendiente = PROFORMAS_PENDIENTES.get(chat_id)
        if not pendiente:
            return "🧾 No tienes una proforma activa. Crea una con 'proforma 2 arboles, 3 galletas'."

        resto = quitar_palabras_intencion(texto_norm, INTENCIONES["agregar"])
        items_nuevos = parsear_items_multiples(resto)
        if not items_nuevos:
            return "🧾 Que quieres agregar? Ej: 'agregar 1 cel, 3 codos'."

        return _agregar_a_proforma(chat_id, items_nuevos, usuario_nombre)

    resto = quitar_palabras_intencion(texto_norm, INTENCIONES["proforma"])
    items_pedidos = parsear_items_multiples(resto)
    if not items_pedidos:
        return "🧾 Usa: 'proforma 2 arboles, 3 galletas'."

    return _crear_proforma(items_pedidos, usuario_nombre, chat_id)


# CORRECCION 4: Funciones auxiliares para proforma con seleccion individual

async def procesar_proforma_seleccionada(update, context, producto, args):
    """Callback cuando el usuario selecciona un producto de una lista ambigua en proforma."""
    chat_id = update.effective_chat.id
    item = args[0]
    items_pedidos = args[1]
    usuario_nombre = args[2]
    items_json_acum = args[3] if len(args) > 3 else []
    subtotal_acum = args[4] if len(args) > 4 else 0.0
    detalle_acum = args[5] if len(args) > 5 else []
    
    cantidad = parse_int(item["cantidad"])
    precio = item["precio"] if item["precio"] is not None else producto["Precio_venta_actual"]
    sub = round(cantidad * precio, 2)
    subtotal_acum += sub
    detalle_acum.append(f"{cantidad} x {producto['Nombre_completo']} @ {fmt_money(precio)} = {fmt_money(sub)}")
    items_json_acum.append({"sku": producto["SKU"], "producto": producto["Nombre_completo"],
                           "cantidad": cantidad, "precio": precio, "subtotal": sub})
    
    await _continuar_proforma_pendiente(update, context, chat_id, items_pedidos, usuario_nombre, 
                                         items_json_acum, subtotal_acum, detalle_acum)


async def _continuar_proforma_pendiente(update, context, chat_id, items_pedidos, usuario_nombre,
                                         items_json_acum, subtotal_acum, detalle_acum):
    """Continua procesando la cola de items de proforma pendientes."""
    if not items_pedidos:
        if not items_json_acum:
            await update.message.reply_text("No pude armar la proforma. Revisa los nombres.", parse_mode=None)
            if chat_id in CONTEXTO_SELECCION:
                del CONTEXTO_SELECCION[chat_id]
            return
        
        porcentaje_descuento = calcular_descuento(subtotal_acum)
        descuento_soles = round(subtotal_acum * porcentaje_descuento, 2)
        total = round(subtotal_acum - descuento_soles, 2)

        nuevo_id = int(time.time())
        fecha = ahora_iso()
        ws("proformas").append_row(
            [nuevo_id, fecha, "", json.dumps(items_json_acum, ensure_ascii=False),
             subtotal_acum, descuento_soles, porcentaje_descuento, total, "pendiente", "", usuario_nombre],
            value_input_option="USER_ENTERED",
        )

        PROFORMAS_PENDIENTES[chat_id] = {
            "id": nuevo_id,
            "items": items_json_acum,
            "subtotal": subtotal_acum,
            "total": total,
            "modificable": True
        }

        respuesta = "🧾 Proforma:\n" + "\n".join(detalle_acum)
        respuesta += f"\n\nSubtotal: {fmt_money(subtotal_acum)}"
        if porcentaje_descuento > 0:
            respuesta += f"\n📉 Descuento sugerido ({porcentaje_descuento*100:.0f}%): -{fmt_money(descuento_soles)}"
        respuesta += f"\nTotal: {fmt_money(total)}"
        respuesta += "\n\nPara agregar mas: 'agregar 1 cel, 3 codos'"
        respuesta += "\nPara confirmar: 'confirmar' (o 'confirmar <pago>' / 'confirmar <total_real> <pago>')"
        await update.message.reply_text(respuesta, parse_mode=None)
        if chat_id in CONTEXTO_SELECCION:
            del CONTEXTO_SELECCION[chat_id]
        return
    
    item = items_pedidos.pop(0)
    nombre_limpio = limpiar_nombre_completo(item["nombre"])
    productos = leer_inventario()
    
    estado, resultado = buscar_producto(nombre_limpio, productos)
    
    if estado == "multiple":
        CONTEXTO_SELECCION[chat_id] = {
            "cola_proforma": items_pedidos,
            "items_json_acum": items_json_acum,
            "subtotal_acum": subtotal_acum,
            "detalle_acum": detalle_acum,
            "callback": procesar_proforma_seleccionada,
            "args": (item, items_pedidos, usuario_nombre, items_json_acum, subtotal_acum, detalle_acum),
            "productos": resultado,
        }
        respuesta = (f"Encontre varias opciones para '{item['nombre']}', responde con el numero:\n" +
                    formatear_lista_productos(resultado) +
                    "\n\nEscribe el numero del producto que quieres.")
        await update.message.reply_text(respuesta, parse_mode=None)
        return
    
    if estado == "no_encontrado":
        detalle_acum.append(f"No encontre '{item['nombre']}'")
        await _continuar_proforma_pendiente(update, context, chat_id, items_pedidos, usuario_nombre,
                                             items_json_acum, subtotal_acum, detalle_acum)
        return
    
    p = resultado
    cantidad = parse_int(item["cantidad"])
    precio = item["precio"] if item["precio"] is not None else p["Precio_venta_actual"]
    sub = round(cantidad * precio, 2)
    subtotal_acum += sub
    detalle_acum.append(f"{cantidad} x {p['Nombre_completo']} @ {fmt_money(precio)} = {fmt_money(sub)}")
    items_json_acum.append({"sku": p["SKU"], "producto": p["Nombre_completo"],
                           "cantidad": cantidad, "precio": precio, "subtotal": sub})
    
    await _continuar_proforma_pendiente(update, context, chat_id, items_pedidos, usuario_nombre,
                                         items_json_acum, subtotal_acum, detalle_acum)


def _crear_proforma(items_pedidos, usuario_nombre, chat_id):
    detalle = []
    items_json = []
    subtotal = 0.0
    cola_pendientes = []
    
    productos = leer_inventario()
    
    for item in items_pedidos:
        nombre_limpio = limpiar_nombre_completo(item["nombre"])
        estado, resultado = buscar_producto(nombre_limpio, productos)

        if estado == "multiple":
            cola_pendientes.append(item)
            continue

        if estado == "no_encontrado":
            detalle.append(f"No encontre '{item['nombre']}'")
            continue

        p = resultado
        cantidad = parse_int(item["cantidad"])
        precio = item["precio"] if item["precio"] is not None else p["Precio_venta_actual"]
        sub = round(cantidad * precio, 2)
        subtotal += sub
        detalle.append(f"{cantidad} x {p['Nombre_completo']} @ {fmt_money(precio)} = {fmt_money(sub)}")
        items_json.append({"sku": p["SKU"], "producto": p["Nombre_completo"],
                           "cantidad": cantidad, "precio": precio, "subtotal": sub})

    if cola_pendientes:
        CONTEXTO_SELECCION[chat_id] = {
            "cola_proforma": cola_pendientes,
            "items_json_acum": items_json,
            "subtotal_acum": subtotal,
            "detalle_acum": detalle,
            "callback": procesar_proforma_seleccionada,
            "args": (cola_pendientes[0], cola_pendientes, usuario_nombre, items_json, subtotal, detalle),
            "productos": None,
        }
        return None

    if not items_json:
        return "No pude armar la proforma. Revisa los nombres.\n" + "\n".join(detalle)

    porcentaje_descuento = calcular_descuento(subtotal)
    descuento_soles = round(subtotal * porcentaje_descuento, 2)
    total = round(subtotal - descuento_soles, 2)

    nuevo_id = int(time.time())
    fecha = ahora_iso()
    ws("proformas").append_row(
        [nuevo_id, fecha, "", json.dumps(items_json, ensure_ascii=False),
         subtotal, descuento_soles, porcentaje_descuento, total, "pendiente", "", usuario_nombre],
        value_input_option="USER_ENTERED",
    )

    PROFORMAS_PENDIENTES[chat_id] = {
        "id": nuevo_id,
        "items": items_json,
        "subtotal": subtotal,
        "total": total,
        "modificable": True
    }

    respuesta = "🧾 Proforma:\n" + "\n".join(detalle)
    respuesta += f"\n\nSubtotal: {fmt_money(subtotal)}"
    if porcentaje_descuento > 0:
        respuesta += f"\n📉 Descuento sugerido ({porcentaje_descuento*100:.0f}%): -{fmt_money(descuento_soles)}"
    respuesta += f"\nTotal: {fmt_money(total)}"
    respuesta += "\n\nPara agregar mas: 'agregar 1 cel, 3 codos'"
    respuesta += "\nPara confirmar: 'confirmar' (o 'confirmar <pago>' / 'confirmar <total_real> <pago>')"
    return respuesta


def _agregar_a_proforma(chat_id, items_nuevos, usuario_nombre):
    pendiente = PROFORMAS_PENDIENTES.get(chat_id)
    if not pendiente:
        return "🧾 No tienes una proforma activa."

    productos = leer_inventario()
    detalle = []
    subtotal_extra = 0.0
    items_extra = []

    for item in items_nuevos:
        nombre_limpio = limpiar_nombre_completo(item["nombre"])
        estado, resultado = buscar_producto(nombre_limpio, productos)

        if estado == "multiple":
            resultado = resultado[0]
            estado = "ok"

        if estado == "no_encontrado":
            detalle.append(f"No encontre '{item['nombre']}'")
            continue

        p = resultado
        cantidad = parse_int(item["cantidad"])
        precio = item["precio"] if item["precio"] is not None else p["Precio_venta_actual"]
        sub = round(cantidad * precio, 2)
        subtotal_extra += sub
        detalle.append(f"{cantidad} x {p['Nombre_completo']} @ {fmt_money(precio)} = {fmt_money(sub)}")
        items_extra.append({"sku": p["SKU"], "producto": p["Nombre_completo"],
                            "cantidad": cantidad, "precio": precio, "subtotal": sub})

    if not items_extra:
        return "⚠️ No pude agregar nada. Revisa los nombres."

    pendiente["items"] += items_extra
    nuevo_subtotal = pendiente["subtotal"] + subtotal_extra
    pendiente["subtotal"] = nuevo_subtotal
    nuevo_total = round(nuevo_subtotal * (1 - calcular_descuento(nuevo_subtotal)), 2)
    pendiente["total"] = nuevo_total

    hoja_pro = ws("proformas")
    filas_pro = hoja_pro.get_all_values()
    for i, fila in enumerate(filas_pro[1:], start=2):
        if fila and str(fila[0]).strip() == str(pendiente["id"]):
            hoja_pro.update_cell(i, 4, json.dumps(pendiente["items"], ensure_ascii=False))
            hoja_pro.update_cell(i, 5, nuevo_subtotal)
            hoja_pro.update_cell(i, 6, round(nuevo_subtotal * calcular_descuento(nuevo_subtotal), 2))
            hoja_pro.update_cell(i, 8, nuevo_total)
            break

    respuesta = "🧾 Proforma actualizada\n\n"
    for it in pendiente["items"]:
        respuesta += f"{it['cantidad']} x {it['producto']} @ {fmt_money(it['precio'])} = {fmt_money(it['subtotal'])}\n"
    errores_no_encontrados = [d for d in detalle if d.startswith("No encontre")]
    if errores_no_encontrados:
        respuesta += "\n" + "\n".join(errores_no_encontrados) + "\n"
    respuesta += f"\nNuevo subtotal: {fmt_money(nuevo_subtotal)}"
    respuesta += f"\nNuevo total: {fmt_money(nuevo_total)}"
    respuesta += "\n\nPuedes seguir agregando con 'agregar ...'"
    respuesta += "\nO confirmar con 'confirmar'"
    return respuesta


def actualizar_proforma_estado(proforma_id, estado):
    hoja = ws("proformas")
    filas = hoja.get_all_values()
    for i, fila in enumerate(filas[1:], start=2):
        if fila and str(fila[0]).strip() == str(proforma_id):
            hoja.update_cell(i, 9, estado)
            hoja.update_cell(i, 10, ahora_iso())
            break


def registrar_ajuste_proforma(proforma_id, tipo, monto_abs, subtotal, total_real):
    fecha = ahora_iso()
    if tipo == "extra":
        ws("extras_y_perdidas").append_row(
            [fecha, proforma_id, "", "extra", monto_abs, "Ajuste en confirmacion de proforma",
             monto_abs, 0, monto_abs],
            value_input_option="USER_ENTERED",
        )
        ws("ahorro").append_row(
            [fecha, proforma_id, "", subtotal, total_real, monto_abs,
             "Ajuste en confirmacion de proforma", monto_abs],
            value_input_option="USER_ENTERED",
        )
    else:
        ws("extras_y_perdidas").append_row(
            [fecha, proforma_id, "", "perdida", monto_abs, "Ajuste en confirmacion de proforma",
             0, monto_abs, monto_abs],
            value_input_option="USER_ENTERED",
        )


def manejar_confirmar(texto_norm, chat_id, usuario_nombre):
    pendiente = PROFORMAS_PENDIENTES.get(chat_id)
    if not pendiente:
        return "🧾 No tienes ninguna proforma pendiente. Crea una primero: 'proforma 2 arboles'."

    # Extraer números (incluye decimales)
    numeros = extraer_numeros(quitar_palabras_intencion(texto_norm, INTENCIONES["confirmar"]))
    
    # Ahora soporta decimales
    total_real = None
    pago = None
    if len(numeros) >= 2:
        total_real = numeros[0]
        pago = numeros[1]
    elif len(numeros) == 1:
        # Si solo hay un número, lo tomamos como pago (mantiene compatibilidad con versiones anteriores)
        pago = numeros[0]

    productos = leer_inventario()
    lineas = []
    ganancia_total = 0.0

    stock_ok = True
    for item in pendiente["items"]:
        producto = next((p for p in productos if p["SKU"] == item["sku"]), None)
        if not producto:
            lineas.append(f"{item['producto']} ya no existe en el inventario.")
            stock_ok = False
            continue
        if producto["Stock_actual"] < item["cantidad"]:
            lineas.append(f"Stock insuficiente de {producto['Nombre_completo']} "
                          f"(disponible {parse_int(producto['Stock_actual'])}, pedido {item['cantidad']}).")
            stock_ok = False
            continue

    if not stock_ok:
        return "⚠️ No se puede confirmar la proforma porque hay productos sin stock suficiente.\n" + "\n".join(lineas)

    lineas = []
    ganancia_total = 0.0

    for item in pendiente["items"]:
        producto = next((p for p in productos if p["SKU"] == item["sku"]), None)
        if not producto:
            lineas.append(f"{item['producto']} ya no existe en el inventario.")
            continue
        linea, ganancia = _vender_un_item(producto, item["cantidad"], item["precio"], usuario_nombre)
        lineas.append(linea)
        ganancia_total += ganancia

    actualizar_proforma_estado(pendiente["id"], "confirmada")
    del PROFORMAS_PENDIENTES[chat_id]
    actualizar_hoja_ganancias()

    respuesta = "🧾 Proforma confirmada como venta:\n" + "\n".join(lineas)
    respuesta += f"\n📈 Ganancia: {fmt_money(ganancia_total)}"

    subtotal_catalogo = pendiente["subtotal"]
    if total_real is not None:
        diferencia = round(total_real - subtotal_catalogo, 2)
        if diferencia > 0:
            respuesta += f"\n💚 Extra cobrado vs catalogo: {fmt_money(diferencia)}"
            registrar_ajuste_proforma(pendiente["id"], "extra", abs(diferencia), subtotal_catalogo, total_real)
        elif diferencia < 0:
            respuesta += f"\n📉 Descuento/perdida vs catalogo: {fmt_money(abs(diferencia))}"
            registrar_ajuste_proforma(pendiente["id"], "perdida", abs(diferencia), subtotal_catalogo, total_real)

    if pago is not None:
        total_a_cobrar = total_real if total_real is not None else pendiente["total"]
        vuelto = round(pago - total_a_cobrar, 2)
        if vuelto > 0:
            respuesta += f"\n💵 Pago {fmt_money(pago)} -> Vuelto: {fmt_money(vuelto)}"
        elif vuelto < 0:
            respuesta += f"\n💵 Pago {fmt_money(pago)} -> Falta cobrar: {fmt_money(abs(vuelto))}"
        else:
            respuesta += f"\n💵 Pago exacto: {fmt_money(pago)}"

    return respuesta


async def procesar_actualizacion_seleccionada(update, context, producto, args):
    usuario_nombre = args[0]
    campo = args[1]
    numeros = args[2]
    costo_explicito = args[3] if len(args) > 3 else None
    precio_explicito = args[4] if len(args) > 4 else None

    ahora = ahora_iso()

    if campo == "Stock_actual":
        nuevo_stock = parse_int(numeros[-1]) if numeros else parse_int(costo_explicito or precio_explicito or 0)
        actualizar_celda("inventario", producto["_row"], "Stock_actual", nuevo_stock)
        actualizar_celda("inventario", producto["_row"], "Estado_stock",
                         estado_stock(nuevo_stock, producto["Stock_minimo"]))
        actualizar_celda("inventario", producto["_row"], "Ultima_actualizacion", ahora)
        actualizar_celda("inventario", producto["_row"], "Actualizado_por", usuario_nombre)
        respuesta = f"Actualizado: {producto['Nombre_completo']} -> Stock: {nuevo_stock}"
        await update.message.reply_text(respuesta, parse_mode=None)
        return

    if costo_explicito is not None and precio_explicito is not None:
        costo_nuevo, precio_nuevo = costo_explicito, precio_explicito
    elif costo_explicito is not None and precio_explicito is None:
        costo_nuevo = costo_explicito
        precio_nuevo = numeros[0] if numeros else producto["Precio_venta_actual"]
    elif precio_explicito is not None and costo_explicito is None:
        precio_nuevo = precio_explicito
        costo_nuevo = numeros[0] if numeros else producto["Costo"]
    elif len(numeros) >= 2:
        costo_nuevo, precio_nuevo = numeros[0], numeros[1]
    else:
        if campo == "Costo":
            costo_nuevo, precio_nuevo = numeros[0], producto["Precio_venta_actual"]
        else:
            costo_nuevo, precio_nuevo = producto["Costo"], numeros[0]

    actualizar_celda("inventario", producto["_row"], "Costo", costo_nuevo)
    actualizar_celda("inventario", producto["_row"], "Precio_venta_actual", precio_nuevo)
    margen = round((precio_nuevo - costo_nuevo) / costo_nuevo, 6) if costo_nuevo else 0
    actualizar_celda("inventario", producto["_row"], "Margen", margen)
    actualizar_celda("inventario", producto["_row"], "Ultima_actualizacion", ahora)
    actualizar_celda("inventario", producto["_row"], "Actualizado_por", usuario_nombre)

    respuesta = (f"Actualizado: {producto['Nombre_completo']} -> "
                 f"Costo: {fmt_money(costo_nuevo)} | Precio venta: {fmt_money(precio_nuevo)}")
    await update.message.reply_text(respuesta, parse_mode=None)


PATRON_ALIAS_COMANDO = re.compile(r'^alias\s+"([^"]+)"\s+(\S+)\s*$', re.IGNORECASE)


def manejar_alias(texto_original, usuario_nombre):
    m = PATRON_ALIAS_COMANDO.match(texto_original.strip())
    if not m:
        return ('⚠️ Usa el formato: alias "Nombre completo exacto" nuevo_alias\n'
                'Ejemplo: alias "Tubo 1/2 con rosca" tr')

    nombre_exacto = m.group(1).strip()
    nuevo_alias = m.group(2).strip()

    productos = leer_inventario()

    producto = next(
        (p for p in productos if p["Nombre_completo"].strip().lower() == nombre_exacto.lower()),
        None,
    )
    if not producto:
        return (f"❌ No encontré el producto exacto '{nombre_exacto}'. "
                f"Revisa el nombre o usa `tubo` para ver la lista con nombres copiables.")

    nuevo_alias_norm = normalizar(nuevo_alias)
    if not nuevo_alias_norm:
        return "⚠️ El alias no puede estar vacio."

    conflicto = next(
        (p for p in productos
         if p["SKU"] != producto["SKU"] and normalizar(p.get("Alias", "")) == nuevo_alias_norm),
        None,
    )
    if conflicto:
        return (f"⚠️ El alias '{nuevo_alias}' ya esta asignado a "
                f"'{conflicto['Nombre_completo']}'. Usa otro.")

    actualizar_celda("inventario", producto["_row"], "Alias", nuevo_alias_norm)
    actualizar_celda("inventario", producto["_row"], "Ultima_actualizacion", ahora_iso())
    actualizar_celda("inventario", producto["_row"], "Actualizado_por", usuario_nombre)

    return f"✅ Alias actualizado: {producto['Nombre_completo']} -> alias: `{nuevo_alias_norm}`"


# CORRECCION 8 (+ coma decimal): Mejorar manejar_actualizar para aceptar orden de palabras variable
# y formatos c<valor> y v<valor> (acepta coma o punto como separador decimal)

def manejar_actualizar(texto_norm, usuario_nombre, chat_id):
    if re.search(r"\bprecio\b", texto_norm):
        campo = "Precio_venta_actual"
    elif re.search(r"\bcosto\b", texto_norm):
        campo = "Costo"
    elif re.search(r"\bstock\b", texto_norm):
        campo = "Stock_actual"
    else:
        campo = "Precio_venta_actual"
    
    # Extraer c<valor> y v<valor> explicitos (acepta coma o punto decimal)
    costo_explicito = None
    precio_explicito = None
    m_c = re.search(r'\bc[:\-]?\s*(\d+(?:[.,]\d+)?)', texto_norm, re.IGNORECASE)
    m_v = re.search(r'\bv[:\-]?\s*(\d+(?:[.,]\d+)?)', texto_norm, re.IGNORECASE)
    if m_c:
        costo_explicito = parse_decimal(m_c.group(1))
    if m_v:
        precio_explicito = parse_decimal(m_v.group(1))
    
    # Extraer numeros sueltos (excluyendo los ya capturados como c/v)
    # CORRECCION COMA DECIMAL: regex acepta coma o punto
    texto_para_numeros = re.sub(r'\b[cv][:\-]?\s*\d+(?:[.,]\d+)?', '', texto_norm, flags=re.IGNORECASE)
    numeros = extraer_numeros(texto_para_numeros)
    
    if not numeros and costo_explicito is None and precio_explicito is None:
        return "Usa: 'actualizar arboles precio 12' o 'actualizar arboles precio 5 19' (costo y precio)."

    # Extraer nombre del producto - eliminar c<valor> y v<valor> primero
    texto_sin_cv = re.sub(r'\b[cv][:\-]?\s*\d+(?:[.,]\d+)?', '', texto_norm, flags=re.IGNORECASE)
    texto_sin_cv = re.sub(r'\s+', ' ', texto_sin_cv).strip()
    palabras_extra = set(INTENCIONES["actualizar"]) | {"precio", "costo", "stock", "c", "v"}
    nombre_producto = extraer_nombre_producto(texto_sin_cv, palabras_extra)
    if not nombre_producto:
        return "Que producto quieres actualizar?"

    productos = leer_inventario()
    estado, resultado = buscar_producto(nombre_producto, productos)
    if estado == "multiple":
        CONTEXTO_SELECCION[chat_id] = {
            "productos": resultado,
            "callback": procesar_actualizacion_seleccionada,
            "args": (usuario_nombre, campo, numeros, costo_explicito, precio_explicito)
        }
        return ("Encontre varias opciones, responde con el numero:\n" +
                formatear_lista_productos(resultado) +
                "\n\nEscribe el numero del producto que quieres actualizar.")

    if estado == "no_encontrado":
        return f"No encontre '{nombre_producto}'."

    p = resultado
    ahora = ahora_iso()

    if campo == "Stock_actual":
        nuevo_stock = parse_int(numeros[-1]) if numeros else parse_int(costo_explicito or precio_explicito or 0)
        actualizar_celda("inventario", p["_row"], "Stock_actual", nuevo_stock)
        actualizar_celda("inventario", p["_row"], "Estado_stock",
                         estado_stock(nuevo_stock, p["Stock_minimo"]))
        actualizar_celda("inventario", p["_row"], "Ultima_actualizacion", ahora)
        actualizar_celda("inventario", p["_row"], "Actualizado_por", usuario_nombre)
        return f"Actualizado: {p['Nombre_completo']} -> Stock: {nuevo_stock}"

    if costo_explicito is not None and precio_explicito is not None:
        costo_nuevo, precio_nuevo = costo_explicito, precio_explicito
    elif costo_explicito is not None and precio_explicito is None:
        costo_nuevo = costo_explicito
        precio_nuevo = numeros[0] if numeros else p["Precio_venta_actual"]
    elif precio_explicito is not None and costo_explicito is None:
        precio_nuevo = precio_explicito
        costo_nuevo = numeros[0] if numeros else p["Costo"]
    elif len(numeros) >= 2:
        costo_nuevo, precio_nuevo = numeros[0], numeros[1]
    else:
        if campo == "Costo":
            costo_nuevo, precio_nuevo = numeros[0], p["Precio_venta_actual"]
        else:
            costo_nuevo, precio_nuevo = p["Costo"], numeros[0]

    actualizar_celda("inventario", p["_row"], "Costo", costo_nuevo)
    actualizar_celda("inventario", p["_row"], "Precio_venta_actual", precio_nuevo)
    margen = round((precio_nuevo - costo_nuevo) / costo_nuevo, 6) if costo_nuevo else 0
    actualizar_celda("inventario", p["_row"], "Margen", margen)
    actualizar_celda("inventario", p["_row"], "Ultima_actualizacion", ahora)
    actualizar_celda("inventario", p["_row"], "Actualizado_por", usuario_nombre)

    return (f"Actualizado: {p['Nombre_completo']} -> "
            f"Costo: {fmt_money(costo_nuevo)} | Precio venta: {fmt_money(precio_nuevo)}")


def manejar_ganancias(texto_norm=None):
    if texto_norm:
        hoy_dt = date.today()
        if re.search(r"\bsemana\b", texto_norm):
            fecha_inicio = (hoy_dt - timedelta(days=6)).isoformat()
            fecha_fin = hoy_dt.isoformat()
            etiqueta = f"la semana ({fecha_inicio} al {fecha_fin})"
            return _formatear_resumen_periodo(etiqueta, fecha_inicio, fecha_fin)
        if re.search(r"\bmes\b", texto_norm):
            fecha_inicio = hoy_dt.replace(day=1).isoformat()
            fecha_fin = hoy_dt.isoformat()
            etiqueta = f"el mes ({fecha_inicio} al {fecha_fin})"
            return _formatear_resumen_periodo(etiqueta, fecha_inicio, fecha_fin)
        if re.search(r"\bano\b", texto_norm):
            fecha_inicio = hoy_dt.replace(month=1, day=1).isoformat()
            fecha_fin = hoy_dt.isoformat()
            etiqueta = f"el año ({fecha_inicio} al {fecha_fin})"
            return _formatear_resumen_periodo(etiqueta, fecha_inicio, fecha_fin)

    resumen = actualizar_hoja_ganancias()
    hoy = date.today().isoformat()

    if resumen:
        return (f"📈 Resumen de hoy ({hoy}):\n"
                f"Ventas: {fmt_money(resumen['ventas'])}\n"
                f"Costos: {fmt_money(resumen['costos'])}\n"
                f"Ganancia neta: {fmt_money(resumen['ganancia_neta'])}\n"
                f"Margen: {resumen['margen']*100:.1f}%\n"
                f"📉 Descuentos: {fmt_money(resumen['descuentos'])}\n"
                f"💚 Extra: {fmt_money(resumen['extras'])}\n"
                f"Producto mas vendido: {resumen['producto_top'] if resumen['producto_top'] else 'Ninguno'}\n"
                f"Vendedor top: {resumen['vendedor_top'] if resumen['vendedor_top'] else 'Ninguno'}")

    hoja = ws("ganancias")
    filas = hoja.get_all_values(value_render_option="UNFORMATTED_VALUE")

    fila_hoy = None
    for f in filas[1:]:
        if f and fecha_a_iso(f[0]) == hoy:
            fila_hoy = f
            break

    if not fila_hoy:
        return f"No hay datos para hoy ({hoy})"

    total_ventas = parse_decimal(fila_hoy[1]) if len(fila_hoy) > 1 else 0
    total_costos = parse_decimal(fila_hoy[2]) if len(fila_hoy) > 2 else 0
    ganancia_neta = parse_decimal(fila_hoy[3]) if len(fila_hoy) > 3 else 0
    margen = parse_decimal(fila_hoy[4]) if len(fila_hoy) > 4 else 0
    descuentos = parse_decimal(fila_hoy[5]) if len(fila_hoy) > 5 else 0
    extra = parse_decimal(fila_hoy[6]) if len(fila_hoy) > 6 else 0
    producto_top = fila_hoy[7] if len(fila_hoy) > 7 else ""
    vendedor_top = fila_hoy[8] if len(fila_hoy) > 8 else ""

    return (f"📈 Resumen de hoy ({hoy}):\n"
            f"Ventas: {fmt_money(total_ventas)}\n"
            f"Costos: {fmt_money(total_costos)}\n"
            f"Ganancia neta: {fmt_money(ganancia_neta)}\n"
            f"Margen: {margen*100:.1f}%\n"
            f"📉 Descuentos: {fmt_money(descuentos)}\n"
            f"💚 Extra: {fmt_money(extra)}\n"
            f"Producto mas vendido: {producto_top if producto_top else 'Ninguno'}\n"
            f"Vendedor top: {vendedor_top if vendedor_top else 'Ninguno'}")


def _formatear_resumen_periodo(etiqueta, fecha_inicio, fecha_fin):
    resumen = calcular_resumen_periodo(fecha_inicio, fecha_fin)
    if not resumen or resumen["dias_con_datos"] == 0:
        return f"📊 No hay datos registrados para {etiqueta}."

    return (f"📊 Resumen de {etiqueta}:\n"
            f"Ventas: {fmt_money(resumen['ventas'])}\n"
            f"Costos: {fmt_money(resumen['costos'])}\n"
            f"Ganancia neta: {fmt_money(resumen['ganancia_neta'])}\n"
            f"Margen: {resumen['margen']*100:.1f}%\n"
            f"📉 Descuentos: {fmt_money(resumen['descuentos'])}\n"
            f"💚 Extra: {fmt_money(resumen['extras'])}\n"
            f"Producto mas vendido: {resumen['producto_top'] if resumen['producto_top'] else 'Ninguno'}\n"
            f"Vendedor top: {resumen['vendedor_top'] if resumen['vendedor_top'] else 'Ninguno'}")


def manejar_recomendar(texto_norm):
    nombre_producto = extraer_nombre_producto(texto_norm, INTENCIONES["recomendar"])
    if not nombre_producto:
        return "Para que producto quieres una recomendacion?"

    productos = leer_inventario()
    estado, resultado = buscar_producto(nombre_producto, productos)
    if estado == "multiple":
        return "Encontre varias opciones:\n" + formatear_lista_productos(resultado)
    if estado == "no_encontrado":
        return f"No encontre '{nombre_producto}'."

    p = resultado
    costo = p["Costo"]
    precio_actual = p["Precio_venta_actual"]
    sugerido_60 = round(costo * 1.6, 2)
    sugerido_80 = round(costo * 1.8, 2)
    
    respuesta_base = (f"Recomendacion para {p['Nombre_completo']}:\n"
                      f"Costo actual: {fmt_money(costo)}\n"
                      f"Precio actual: {fmt_money(precio_actual)}\n"
                      f"Sugerido (60% margen): {fmt_money(sugerido_60)}\n"
                      f"Sugerido (80% margen): {fmt_money(sugerido_80)}")
    
    try:
        margenes = ws("configuracion_margenes").get_all_records(value_render_option="UNFORMATTED_VALUE")
        margen_categoria = next((m for m in margenes if m.get("Categoria") == p.get("Categoria")), None)
        respuesta_ia = groq_recomendar_precio(p, margen_categoria)
        return respuesta_base + f"\n\nIA: {respuesta_ia}"
    except Exception:
        return respuesta_base


async def procesar_historial_seleccionado(update, context, producto, args):
    p = producto
    respuesta = _generar_historial_por_sku(p["SKU"], p["Nombre_completo"])
    await update.message.reply_text(respuesta, parse_mode=None)


def _generar_historial_por_sku(sku, nombre_completo):
    try:
        filas = ws("movimientos").get_all_values(value_render_option="UNFORMATTED_VALUE")
    except Exception as e:
        log.warning(f"No se pudo leer movimientos para historial: {e}")
        return f"No pude leer el historial de {nombre_completo} en este momento."

    compras = []
    ventas = []
    for fila in filas[1:]:
        if len(fila) < 5 or str(fila[1]).strip() != sku:
            continue
        fecha_fila = fecha_a_iso(fila[0])
        tipo = str(fila[2]).strip().lower()
        cantidad = parse_decimal(fila[3])
        precio_unitario = parse_decimal(fila[4])
        registro = (fecha_fila, cantidad, precio_unitario)
        if tipo == "compra":
            compras.append(registro)
        elif tipo == "venta":
            ventas.append(registro)

    compras.sort(key=lambda r: r[0], reverse=True)
    ventas.sort(key=lambda r: r[0], reverse=True)
    compras = compras[:5]
    ventas = ventas[:5]

    lineas = [f"📜 Historial de {nombre_completo} (SKU {sku}):"]

    lineas.append("📦 Compras:")
    if compras:
        for fecha_f, cantidad, precio in compras:
            lineas.append(f"• {fecha_f}: {parse_int(cantidad)} unid. @ {fmt_money(precio)}")
    else:
        lineas.append("• Sin compras registradas.")

    lineas.append("💰 Ventas:")
    if ventas:
        for fecha_f, cantidad, precio in ventas:
            lineas.append(f"• {fecha_f}: {parse_int(cantidad)} unid. @ {fmt_money(precio)}")
    else:
        lineas.append("• Sin ventas registradas.")

    return "\n".join(lineas)


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
            # CORRECCION 4: No borrar CONTEXTO_SELECCION aqui para ventas/proformas con cola
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
app = None  # Variable global para la aplicación de Telegram
loop = None  # Event loop principal para el bot

flask_app = Flask(__name__)

@flask_app.route('/')
def home():
    return "Bot de Inventario Lumin funcionando con Webhook ✅"

@flask_app.route('/ping')
def ping():
    return "pong", 200

@flask_app.route('/webhook', methods=['POST'])
def webhook():
    """Recibe los updates de Telegram vía webhook."""
    global loop
    try:
        json_data = request.get_json(force=True)
        update = Update.de_json(json_data, app.bot)
        
        # Enviar la tarea al loop principal
        future = asyncio.run_coroutine_threadsafe(
            app.process_update(update),
            loop
        )
        # Esperar a que termine (opcional, pero recomendado)
        future.result(timeout=30)
        
        return 'OK', 200
    except Exception as e:
        log.error(f"Error en webhook: {e}")
        return 'Error', 500

def run_bot():
    """Ejecuta el bot en un loop separado."""
    global app, loop
    
    # Crear un nuevo loop para el bot
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    # Crear la aplicación de Telegram
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    # Registrar todos los handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler(["ayuda", "help"], cmd_ayuda))
    app.add_handler(CommandHandler("reset", cmd_reset))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    # Inicializar y arrancar la aplicación
    loop.run_until_complete(app.initialize())
    loop.run_until_complete(app.start())
    
    # Configurar webhook
    webhook_endpoint = f"{WEBHOOK_URL}/webhook"
    log.info(f"Configurando webhook en: {webhook_endpoint}")
    loop.run_until_complete(app.bot.set_webhook(
        url=webhook_endpoint,
        drop_pending_updates=True
    ))
    log.info("✅ Webhook configurado correctamente.")
    log.info("✅ Bot listo y esperando mensajes vía webhook...")
    
    # Mantener el loop corriendo para siempre
    loop.run_forever()

def main():
    # Iniciar el bot en un hilo separado
    bot_thread = threading.Thread(target=run_bot, daemon=True)
    bot_thread.start()
    
    # Iniciar Flask en el hilo principal
    port = int(os.environ.get("PORT", 10000))
    flask_app.run(host="0.0.0.0", port=port)

if __name__ == "__main__":
    import threading
    main()