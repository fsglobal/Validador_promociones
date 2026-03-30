import os
import re
import sys
import xml.etree.ElementTree as ET
from io import StringIO

import pandas as pd
from colorama import Fore, init

init(autoreset=True)

BASE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
EXCEL_PATH = os.path.join(BASE_PATH, "Excel")
EXPORT_PATH = os.path.join(BASE_PATH, "Export")


# ============================================================
# SISTEMA DE ESCRITURA DUAL (CONSOLA + TXT)
# ============================================================

class DualWriter:
    def __init__(self, console, buffer):
        self.console = console
        self.buffer = buffer

    def write(self, text):
        self.console.write(text)
        clean = re.sub(r"\x1b\[[0-9;]*m", "", text)
        self.buffer.write(clean)

    def flush(self):
        self.console.flush()


# ============================================================
# UTILIDADES GENERALES
# ============================================================

def limpiar_fecha(f):
    if f is None:
        return None
    s = str(f).strip()
    if not s:
        return None
    if " " in s:
        return s.split(" ")[0]
    if "T" in s:
        return s.split("T")[0]
    return s


def normalizar_fecha_excel(f):
    try:
        dt = pd.to_datetime(f)
        if pd.isna(dt):
            return None
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def normalizar_texto(v):
    if v is None:
        return ""
    s = str(v).strip().upper()
    if s in {"", "NAN", "NONE", "NULL"}:
        return ""
    return s


def normalizar_local(v):
    if v is None:
        return None
    s = str(v).strip()
    if s.upper() in {"", "NAN", "NONE", "NULL"}:
        return None
    try:
        return str(int(float(s)))
    except Exception:
        return s


def normalizar_encabezado(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\n", " ")).strip()


def es_id_promocion_valido(v):
    raw = normalizar_texto(v)
    if raw in {"", "NO CONSIDERAR", "NOCONSIDERAR", "N/A", "NA", "-", "--", "SIN ID"}:
        return False

    normalizado = normalizar_local(v)
    if normalizado is None:
        return False

    if normalizar_texto(normalizado) in {"", "NO CONSIDERAR", "NOCONSIDERAR", "N/A", "NA", "-", "--", "SIN ID"}:
        return False

    return bool(re.fullmatch(r"\d+", str(normalizado).strip()))


def limpiar_dataframe_columnas(df):
    if df is None or df.empty:
        return df

    columnas_limpias = []
    usados = {}

    for col in df.columns:
        base = normalizar_encabezado(col) or str(col).strip() or "COLUMNA"
        if base.upper().startswith("UNNAMED:"):
            base = base.split(":", 1)[0]

        clave = base.upper()
        usados.setdefault(clave, 0)
        usados[clave] += 1

        if usados[clave] > 1:
            columnas_limpias.append(f"{base}__{usados[clave]}")
        else:
            columnas_limpias.append(base)

    df = df.copy()
    df.columns = columnas_limpias
    return df


def normalizar_sku(v):
    if v is None:
        return None
    s = str(v).strip()
    if s.upper() in {"", "NAN", "NONE", "NULL"}:
        return None
    try:
        return str(int(float(s)))
    except Exception:
        return normalizar_texto(s)


def normalizar_lista_skus(valores):
    resultado = set()
    if not valores:
        return resultado
    for v in valores:
        sku = normalizar_sku(v)
        if sku:
            resultado.add(sku)
    return resultado


def es_vacio(v):
    return normalizar_texto(v) == ""


def a_float(v):
    if v is None:
        return None
    try:
        if isinstance(v, str):
            v = v.replace("%", "").replace(".", "") if re.fullmatch(r"\d+\.\d{3}", v.strip()) else v
            v = v.replace(",", ".")
        return float(v)
    except Exception:
        return None


def parsear_porcentaje_excel(v):
    if v is None or es_vacio(v):
        return None
    raw = str(v).strip().replace("%", "").replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", raw)
    if not m:
        return None
    num = float(m.group(1))
    return num / 100.0 if num > 1 else num


def floats_iguales(a, b, tol=0.0001):
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) < tol


def money_iguales(a, b, tol=0.01):
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) < tol


def formatear_porcentaje(p):
    if p is None:
        return "-"
    return f"{float(p) * 100:.2f}%"


def formatear_monto(v):
    if v is None:
        return "-"
    try:
        return f"{float(v):.2f}"
    except Exception:
        return str(v)


def formatear_numero(v):
    if v is None:
        return "-"
    try:
        return f"{float(v):.2f}"
    except Exception:
        return str(v)


def formatear_cantidad(v):
    if v is None:
        return "-"
    try:
        n = float(v)
        return str(int(n)) if n.is_integer() else str(n)
    except Exception:
        return str(v)


def _primer_no_vacio(*valores):
    for valor in valores:
        if valor not in (None, "", [], (), set()):
            return valor
    return ""


def _texto_origen_limpio(lista_nombre="", skus=None, etiqueta_lista="Lista", etiqueta_sku="SKU"):
    lista_nombre = normalizar_texto(lista_nombre)
    if lista_nombre:
        return f"{etiqueta_lista}: <span class='text-blue'>({lista_nombre})</span>"
    skus_norm = sorted(normalizar_lista_skus(skus or []))
    if skus_norm:
        return f"{etiqueta_sku}: <span class='text-blue'>({', '.join(skus_norm)})</span>"
    return f"{etiqueta_sku}: <span class='text-blue'>(-)</span>"


def construir_leyenda_excel_compat(tipo_desc_raw, productos_excel, cantidad_excel, porcentaje_excel,
                                   descuento_bruto_excel_q, monto_pack_excel, nombre_lista_excel=""):
    origen = _texto_origen_limpio(nombre_lista_excel, productos_excel, etiqueta_lista="Lista", etiqueta_sku="Productos")
    return (
        f"Excel → Tipo: <span class='text-blue'>({tipo_desc_raw or '-'})</span> | "
        f"{origen} | "
        f"Unidades: <span class='text-blue'>({formatear_cantidad(cantidad_excel)})</span> | "
        f"% comercial Excel: <span class='text-blue'>({formatear_porcentaje(porcentaje_excel)})</span> | "
        f"Dcto bruto Excel(Q): <span class='text-blue'>({formatear_numero(descuento_bruto_excel_q)})</span> | "
        f"PVPOfertaPack: <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span>"
    )


def construir_leyenda_condicion_compat(condition_skus, condition_quantity, nombre_lista_export=""):
    origen = _texto_origen_limpio(nombre_lista_export, condition_skus, etiqueta_lista="Lista", etiqueta_sku="SKU")
    return (
        f"Condición Export → {origen} | "
        f"Cantidad: <span class='text-blue'>({formatear_cantidad(condition_quantity)})</span>"
    )


def construir_leyenda_applier_compat(tipo_desc, promo, applier_skus, applier_product_lists,
                                     porcentaje_excel, applier_pct_nodo, applier_pct_tecnico,
                                     descuento_bruto_excel_q, monto_pack_excel):
    nombre_lista_applier = normalizar_texto(_primer_no_vacio(applier_product_lists[0] if applier_product_lists else "", ""))
    origen = _texto_origen_limpio(nombre_lista_applier, applier_skus, etiqueta_lista="Lista", etiqueta_sku="SKU")
    base = (
        f"Applier Export → Tipo: <span class='text-blue'>({promo.get('applier_type') or '-'})</span> | "
        f"{origen} | "
        f"Cantidad: <span class='text-blue'>({formatear_cantidad(promo.get('applier_quantity'))})</span>"
    )

    if tipo_desc == "2DA":
        return (
            base + " | "
            f"% técnico export: <span class='text-blue'>({formatear_porcentaje(applier_pct_tecnico)})</span> | "
            f"% nodo export: <span class='text-blue'>({formatear_porcentaje(applier_pct_nodo)})</span> | "
            f"% comercial Excel: <span class='text-blue'>({formatear_porcentaje(porcentaje_excel)})</span> | "
            f"Dcto bruto Excel(Q): <span class='text-blue'>({formatear_numero(descuento_bruto_excel_q)})</span> | "
            f"Monto: <span class='text-blue'>({formatear_monto(promo.get('applier_amount'))})</span>"
        )

    if tipo_desc == "NOMINAL":
        return (
            base + " | "
            f"PVPOfertaPack Excel: <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span> | "
            f"Dcto bruto Excel(Q): <span class='text-blue'>({formatear_numero(descuento_bruto_excel_q)})</span> | "
            f"Monto export: <span class='text-blue'>({formatear_monto(promo.get('applier_amount'))})</span>"
        )

    return (
        base + " | "
        f"%: <span class='text-blue'>({formatear_porcentaje(applier_pct_nodo)})</span> | "
        f"Monto: <span class='text-blue'>({formatear_monto(promo.get('applier_amount'))})</span>"
    )


def construir_resumen_excel_limpio(tipo_desc_raw, tipo_desc, nombre_lista_excel, productos_excel,
                                   cantidad_excel, porcentaje_excel, descuento_bruto_excel_q, monto_pack_excel):
    partes = [f"Tipo <span class='text-blue'>({tipo_desc_raw or '-'})</span>"]

    if nombre_lista_excel:
        partes.append(f"Lista <span class='text-blue'>({nombre_lista_excel})</span>")
    elif productos_excel:
        partes.append(f"SKU <span class='text-blue'>({', '.join(sorted(productos_excel))})</span>")

    if cantidad_excel is not None:
        partes.append(f"Cantidad <span class='text-blue'>({formatear_cantidad(cantidad_excel)})</span>")

    if tipo_desc == "2DA" and porcentaje_excel is not None:
        partes.append(f"Descuento comercial <span class='text-blue'>({formatear_porcentaje(porcentaje_excel)})</span>")
    elif tipo_desc == "PORCENTUAL" and porcentaje_excel is not None:
        partes.append(f"Porcentaje <span class='text-blue'>({formatear_porcentaje(porcentaje_excel)})</span>")
    elif tipo_desc in {"NOMINAL", "PACK", "PACK_NOMINAL"}:
        if monto_pack_excel is not None:
            partes.append(f"PVPOfertaPack <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span>")
        if descuento_bruto_excel_q is not None:
            partes.append(f"Dcto bruto Q <span class='text-blue'>({formatear_numero(descuento_bruto_excel_q)})</span>")

    return " | ".join(partes)


def construir_resumen_condicion_limpio(condition_skus, condition_quantity, nombre_lista_export=""):
    partes = []
    if nombre_lista_export:
        partes.append(f"Lista <span class='text-blue'>({nombre_lista_export})</span>")
    elif condition_skus:
        partes.append(f"SKU <span class='text-blue'>({', '.join(sorted(condition_skus))})</span>")
    else:
        partes.append("SKU <span class='text-blue'>(-)</span>")

    if condition_quantity is not None:
        partes.append(f"Cantidad <span class='text-blue'>({formatear_cantidad(condition_quantity)})</span>")

    return " | ".join(partes)


def construir_resumen_applier_limpio(tipo_desc, promo, applier_skus, applier_product_lists,
                                     porcentaje_excel, applier_pct_nodo, applier_pct_tecnico):
    partes = [f"Tipo <span class='text-blue'>({promo.get('applier_type') or '-'})</span>"]

    nombre_lista_applier = normalizar_texto(applier_product_lists[0]) if applier_product_lists else ""
    if nombre_lista_applier:
        partes.append(f"Lista <span class='text-blue'>({nombre_lista_applier})</span>")
    elif applier_skus:
        partes.append(f"SKU <span class='text-blue'>({', '.join(sorted(applier_skus))})</span>")
    else:
        partes.append("SKU <span class='text-blue'>(-)</span>")

    if promo.get("applier_quantity") is not None:
        partes.append(f"Cantidad <span class='text-blue'>({formatear_cantidad(promo.get('applier_quantity'))})</span>")

    if tipo_desc == "2DA":
        if applier_pct_nodo is not None:
            partes.append(f"% nodo export <span class='text-blue'>({formatear_porcentaje(applier_pct_nodo)})</span>")
        if porcentaje_excel is not None:
            partes.append(f"% comercial Excel <span class='text-blue'>({formatear_porcentaje(porcentaje_excel)})</span>")
        if applier_pct_tecnico is not None:
            partes.append(f"% técnico visible <span class='text-blue'>({formatear_porcentaje(applier_pct_tecnico)})</span>")
    elif tipo_desc == "PORCENTUAL":
        if applier_pct_nodo is not None:
            partes.append(f"Porcentaje <span class='text-blue'>({formatear_porcentaje(applier_pct_nodo)})</span>")
    else:
        if promo.get("applier_amount") is not None:
            partes.append(f"Monto <span class='text-blue'>({formatear_monto(promo.get('applier_amount'))})</span>")

    return " | ".join(partes)


def agregar_detalle(detalles, tipo, grupo, mensaje):
    detalles.append((tipo, f"[{grupo}] {mensaje}"))


def validar_multivalor(vals_excel, vals_export, etiqueta):
    detalles = []
    ok = True

    vals_excel = sorted(set(vals_excel))
    vals_export = sorted(set(vals_export))

    for v in vals_excel:
        if v in vals_export:
            detalles.append(("OK", f"{etiqueta} (<span class='text-blue'>{v}</span>) Correcto entre archivo Excel y Export"))
        else:
            ok = False
            detalles.append(("ERR", f"{etiqueta} (<span class='text-blue'>{v}</span>) está en el Excel y no se encuentra asociado en el archivo Export"))

    for v in vals_export:
        if v not in vals_excel:
            ok = False
            detalles.append(("ERR", f"{etiqueta} (<span class='text-blue'>{v}</span>) está en el archivo Export y no se encuentra en el Excel"))

    return ok, detalles


# ============================================================
# UTILIDADES DE NEGOCIO
# ============================================================

def extraer_mecanica_pack(valor):
    """
    Devuelve tupla (lleva, paga) para textos como:
    '2X1', '3x2', 'PACK 4X3', etc.
    Si no encuentra patrón pack válido, devuelve None.
    """
    txt = normalizar_texto(valor)
    txt = txt.replace(" ", "")
    m = re.search(r"(\d+)\s*[Xx]\s*(\d+)", txt)
    if not m:
        return None

    lleva = int(m.group(1))
    paga = int(m.group(2))

    if lleva > paga:
        return (lleva, paga)

    return None


def extraer_mecanica_combo_precio(valor):
    """
    Detecta mecánicas comerciales tipo 3x23990, donde el primer número es cantidad
    y el segundo un precio total del combo.
    """
    txt = normalizar_texto(valor)
    txt = txt.replace(" ", "")
    m = re.search(r"(\d+)\s*[Xx]\s*(\d{4,})", txt)
    if not m:
        return None

    cantidad = int(m.group(1))
    monto = float(m.group(2))
    if cantidad >= 2 and monto > 999:
        return (cantidad, monto)
    return None


def es_nominal_un_producto(tipo_desc, productos_excel, cantidad_excel):
    if tipo_desc != "NOMINAL":
        return False
    if len(productos_excel or set()) != 1:
        return False
    return cantidad_excel in {None, 0, 1}


def obtener_competencia_esperada(area_responsable, tipo_desc, productos_excel, cantidad_excel):
    area = normalizar_texto(area_responsable)
    if area == "BYCP":
        return None if es_nominal_un_producto(tipo_desc, productos_excel, cantidad_excel) else "Comp. X Promociones"
    if area in {"FARMA", "BIENESTAR"}:
        return "Comp. X Producto"
    return None


def validar_competencia_por_area(detalles, promo, area_responsable, tipo_desc, productos_excel, cantidad_excel):
    area = normalizar_texto(area_responsable)
    if not area:
        agregar_detalle(detalles, "WARN", "ÁREA", "No se pudo determinar AreaResponsable desde hoja IMPUT/Input")
        return True

    agregar_detalle(detalles, "INFO", "ÁREA", f"AreaResponsable detectada: <span class='text-blue'>({area})</span>")
    esperada = obtener_competencia_esperada(area, tipo_desc, productos_excel, cantidad_excel)
    if esperada is None:
        if area == "BYCP" and es_nominal_un_producto(tipo_desc, productos_excel, cantidad_excel):
            agregar_detalle(detalles, "INFO", "COMPETENCIA", "Excepción BYCP detectada: NOMINAL de 1 producto. No se fuerza competencia por promoción")
        return True

    actual = promo.get("__tipo_competencia") or "-"
    if actual == esperada:
        agregar_detalle(detalles, "OK", "COMPETENCIA", f"Competencia correcta por regla de área: esperada <span class='text-blue'>({esperada})</span> y Export trae <span class='text-blue'>({actual})</span>")
        return True

    agregar_detalle(detalles, "ERR", "COMPETENCIA", f"Competencia incorrecta por regla de área. Area <span class='text-blue'>({area})</span> esperaba <span class='text-blue'>({esperada})</span> pero Export trae <span class='text-blue'>({actual})</span>")
    return False


def construir_resumen_mensaje(promo, nombre_lista_export, condition_skus):
    origen = f"Lista <span class='text-blue'>({nombre_lista_export})</span>" if nombre_lista_export else (
        f"SKU <span class='text-blue'>({', '.join(sorted(condition_skus))})</span>" if condition_skus else "SKU <span class='text-blue'>(-)</span>"
    )
    cantidad = promo.get("condition_item_quantity")
    if cantidad is None:
        cantidad = promo.get("condition_quantity")
    return (
        f"Tipo <span class='text-blue'>(MSJE)</span> | {origen} | "
        f"Item cada <span class='text-blue'>({formatear_cantidad(cantidad)})</span> | "
        f"Mensaje <span class='text-blue'>({promo.get('message_applier_name') or '-'})</span> | "
        f"Salida <span class='text-blue'>({promo.get('message_output') or '-'})</span> | "
        f"Texto <span class='text-blue'>({promo.get('message_text') or '-'})</span>"
    )


def validar_promocion_mensaje(detalles, grupo, promo, productos_excel, nombre_lista_excel, nombre_lista_export, listas_productos_export):
    ok = True
    condition_skus = normalizar_lista_skus(promo.get("condition_skus", []))
    cantidad_msg = promo.get("condition_item_quantity")
    if cantidad_msg is None:
        cantidad_msg = promo.get("condition_quantity")

    agregar_detalle(detalles, "INFO", "MSJE", construir_resumen_mensaje(promo, nombre_lista_export, condition_skus))

    if promo.get("applier_type") != "MESSAGE":
        ok = False
        agregar_detalle(detalles, "ERR", "MSJE", "La campaña de mensaje debe viajar con MessageApplier")
    else:
        agregar_detalle(detalles, "OK", "MSJE", "MessageApplier detectado correctamente")

    if nombre_lista_excel or nombre_lista_export:
        if nombre_lista_excel == nombre_lista_export and nombre_lista_export:
            agregar_detalle(detalles, "OK", "CONDICIÓN", f"Lista de condición del MSJE coincide: <span class='text-blue'>({nombre_lista_export})</span>")
        else:
            ok = False
            agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Lista de condición del MSJE no coincide. Excel <span class='text-blue'>({nombre_lista_excel or '-'})</span> vs Export <span class='text-blue'>({nombre_lista_export or '-'})</span>")
    else:
        if condition_skus == productos_excel:
            agregar_detalle(detalles, "OK", "CONDICIÓN", "Los SKU de la condición del MSJE coinciden con columna C del Excel")
        else:
            ok = False
            faltan = sorted(productos_excel - condition_skus)
            sobran = sorted(condition_skus - productos_excel)
            if faltan:
                agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Faltan SKU en condición MSJE respecto a columna C: <span class='text-blue'>({', '.join(faltan)})</span>")
            if sobran:
                agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Sobran SKU en condición MSJE respecto a columna C: <span class='text-blue'>({', '.join(sobran)})</span>")

    if cantidad_msg == 1:
        agregar_detalle(detalles, "OK", "MSJE", "Cantidad de condición del mensaje correcta: item cada <span class='text-blue'>(1)</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "MSJE", f"La campaña de mensaje debe viajar con item cada = 1, pero Export trae <span class='text-blue'>({formatear_cantidad(cantidad_msg)})</span>")

    if promo.get("message_applier_name"):
        agregar_detalle(detalles, "OK", "MSJE", f"Nombre del mensaje detectado: <span class='text-blue'>({promo.get('message_applier_name')})</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "MSJE", "MessageApplier no informa messageName")

    if normalizar_texto(promo.get("message_output")) == "SCREEN":
        agregar_detalle(detalles, "OK", "MSJE", "Salida del mensaje correcta: <span class='text-blue'>(SCREEN)</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "MSJE", f"Salida del mensaje incorrecta. Export trae <span class='text-blue'>({promo.get('message_output') or '-'})</span>")

    if promo.get("message_text"):
        agregar_detalle(detalles, "INFO", "MSJE", f"Texto del mensaje: <span class='text-blue'>({promo.get('message_text')})</span>")
    else:
        agregar_detalle(detalles, "WARN", "MSJE", "No se encontró el texto asociado al messageName dentro de messageList")

    return ok


def validar_promocion_farma_combo_precio(detalles, promo, cantidad_excel, combo_excel, nombre_lista_excel, nombre_lista_export):
    ok = True
    combo_qty, combo_amount = combo_excel
    actual_qty_cond = promo.get("condition_quantity")
    actual_qty_applier = promo.get("applier_quantity")
    actual_amount = promo.get("applier_amount")
    strategy = promo.get("applier_strategy")

    agregar_detalle(detalles, "INFO", "APPLIER", f"Nueva lógica FARMA_COMBO_PRECIO detectada: combo <span class='text-blue'>({combo_qty}x{formatear_monto(combo_amount)})</span>")

    if promo.get("applier_type") != "FIX_AMOUNT":
        ok = False
        agregar_detalle(detalles, "ERR", "APPLIER", "La promo FARMA_COMBO_PRECIO debe viajar con FixAmountDiscountApplier")
    else:
        agregar_detalle(detalles, "OK", "APPLIER", "Tipo de applier correcto para FARMA_COMBO_PRECIO: <span class='text-blue'>(FIX_AMOUNT)</span>")

    if actual_qty_cond == combo_qty:
        agregar_detalle(detalles, "OK", "CONDICIÓN", f"Cantidad de condición correcta para FARMA_COMBO_PRECIO: <span class='text-blue'>({combo_qty})</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Cantidad de condición incorrecta para FARMA_COMBO_PRECIO. Esperado <span class='text-blue'>({combo_qty})</span> pero Export trae <span class='text-blue'>({formatear_cantidad(actual_qty_cond)})</span>")

    if actual_qty_applier is not None and int(float(actual_qty_applier)) == int(combo_qty):
        agregar_detalle(detalles, "OK", "APPLIER", f"Cantidad del applier correcta para FARMA_COMBO_PRECIO: <span class='text-blue'>({combo_qty})</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "APPLIER", f"Cantidad del applier incorrecta para FARMA_COMBO_PRECIO. Esperado <span class='text-blue'>({combo_qty})</span> pero Export trae <span class='text-blue'>({formatear_cantidad(actual_qty_applier)})</span>")

    if money_iguales(combo_amount, actual_amount):
        agregar_detalle(detalles, "OK", "APPLIER", f"Monto total del combo correcto para FARMA_COMBO_PRECIO: <span class='text-blue'>({formatear_monto(combo_amount)})</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "APPLIER", f"Monto incorrecto para FARMA_COMBO_PRECIO. Esperado <span class='text-blue'>({formatear_monto(combo_amount)})</span> pero Export trae <span class='text-blue'>({formatear_monto(actual_amount)})</span>")

    if strategy == 1:
        agregar_detalle(detalles, "OK", "APPLIER", "Strategy correcta para FARMA_COMBO_PRECIO: <span class='text-blue'>(1 - menor)</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "APPLIER", f"Strategy incorrecta para FARMA_COMBO_PRECIO. Esperado <span class='text-blue'>(1)</span> pero Export trae <span class='text-blue'>({strategy if strategy is not None else '-'})</span>")

    if nombre_lista_excel or nombre_lista_export:
        if nombre_lista_excel == nombre_lista_export and nombre_lista_export:
            agregar_detalle(detalles, "OK", "LISTA PRODUCTOS", f"Lista de combo FARMA correcta: <span class='text-blue'>({nombre_lista_export})</span>")
        else:
            ok = False
            agregar_detalle(detalles, "ERR", "LISTA PRODUCTOS", f"Lista de combo FARMA incorrecta. Excel <span class='text-blue'>({nombre_lista_excel or '-'})</span> vs Export <span class='text-blue'>({nombre_lista_export or '-'})</span>")

    return ok


def reconstruir_skus_desde_listas(nombres_listas, listas_productos_export):
    """
    Recibe una colección de nombres de listas y devuelve el conjunto total de SKU
    reconstruidos desde el export.
    """
    resultado = set()
    listas_no_reconstruidas = []

    for nombre in nombres_listas or []:
        nn = normalizar_texto(nombre)
        if not nn:
            continue
        skus = listas_productos_export.get(nn, set())
        if skus:
            resultado.update(normalizar_lista_skus(skus))
        else:
            listas_no_reconstruidas.append(nn)

    return resultado, listas_no_reconstruidas


def calcular_porcentaje_tecnico_2da(porcentaje_export, cantidad_applier):
    """
    En 2DA, el porcentaje técnico visible es el porcentaje del nodo
    dividido por la cantidad del applier.
    Ejemplo:
      nodo export = 0.30
      cantidad     = 2
      técnico      = 0.15
    """
    if porcentaje_export is None:
        return None
    try:
        qty = int(float(cantidad_applier)) if cantidad_applier is not None else None
    except Exception:
        qty = None

    if qty and qty > 1:
        return float(porcentaje_export) / qty

    return float(porcentaje_export)


def obtener_valor_descuento_bruto_q(grupo):
    """
    Intenta leer el descuento bruto desde la columna Q del Excel.
    Primero busca por nombre probable; si no existe, intenta por posición 17 (índice 16).
    """
    col_q = buscar_columna(grupo, [
        "DESCUENTO BRUTO",
        "DESCUENTO BRUTO EXCEL",
        "DESCUENTO BRUTO Q",
        "COLUMNA Q",
        "Q"
    ])

    valor = None

    if col_q:
        try:
            celda = grupo[col_q].iloc[0]
            if not es_vacio(celda):
                valor = float(str(celda).replace(",", "."))
                return valor
        except Exception:
            pass

    try:
        if len(grupo.columns) >= 17:
            celda = grupo.iloc[0, 16]
            if not es_vacio(celda):
                valor = float(str(celda).replace(",", "."))
                return valor
    except Exception:
        pass

    return None


def extraer_id_msje_asociado_desde_grupo(grupo):
    col_lista_locales = buscar_columna(grupo, [
        "ID LISTA LOCALES",
        "ID Lista Locales",
        "ID LISTA LOCAL",
        "ID Lista Local",
        "AJ"
    ])
    if not col_lista_locales:
        return None

    valor = grupo[col_lista_locales].iloc[0]
    return normalizar_local(valor) if es_id_promocion_valido(valor) else None


def obtener_promo_msje_asociada(grupo, promos_por_id):
    id_msje = extraer_id_msje_asociado_desde_grupo(grupo)
    if not id_msje:
        return None, None

    promo_msje = promos_por_id.get(id_msje) if promos_por_id else None
    return id_msje, promo_msje


def _formatear_fecha_msje(fecha):
    fecha_limpia = limpiar_fecha(fecha)
    if not fecha_limpia:
        return ""
    try:
        return pd.to_datetime(fecha_limpia).strftime("%d-%m-%Y")
    except Exception:
        return fecha_limpia


def construir_msje_popup_data(id_padre, id_msje_asociado=None, promo_msje_asociada=None,
                              productos_excel=None, nombre_lista_excel=""):
    return {
        "hay": bool(id_msje_asociado),
        "id_msje": normalizar_local(id_msje_asociado) if id_msje_asociado else "",
        "id_padre": normalizar_local(id_padre) if id_padre else "",
        "mensaje": promo_msje_asociada.get("message_applier_name", "") if promo_msje_asociada else "",
        "salida": promo_msje_asociada.get("message_output", "") if promo_msje_asociada else "",
        "texto": promo_msje_asociada.get("message_text", "") if promo_msje_asociada else "",
        "resumen": f"ID MSJE - {normalizar_local(id_msje_asociado)}" if id_msje_asociado else "No hay",
        "resumen_condicion": "-",
        "nombre_lista_excel": nombre_lista_excel or "",
        "nombre_lista_export": normalizar_texto(promo_msje_asociada["productLists"][0]) if promo_msje_asociada and promo_msje_asociada.get("productLists") else "",
        "fecha_inicio": _formatear_fecha_msje(promo_msje_asociada.get("startDate")) if promo_msje_asociada else "",
        "fecha_fin": _formatear_fecha_msje(promo_msje_asociada.get("endDate")) if promo_msje_asociada else "",
    }


# ============================================================
# VALIDACIÓN APPLIER VS CONDICIÓN
# ============================================================

def validar_applier_vs_condicion(detalles, promo, listas_productos_export):
    """
    Valida que el applier viaje exactamente con los mismos productos que la condición.
    Puede venir por SKU explícito o por lista de productos.
    Si no trae ni SKU ni lista, debe ser error.
    """
    ok = True

    condition_skus = normalizar_lista_skus(promo.get("condition_skus", []))
    condition_lists = [normalizar_texto(x) for x in promo.get("condition_product_lists", []) if normalizar_texto(x)]

    applier_skus = normalizar_lista_skus(promo.get("applier_skus", []))
    applier_lists = [normalizar_texto(x) for x in promo.get("applier_product_lists", []) if normalizar_texto(x)]

    condition_list_skus, condition_lists_no_recon = reconstruir_skus_desde_listas(condition_lists, listas_productos_export)
    applier_list_skus, applier_lists_no_recon = reconstruir_skus_desde_listas(applier_lists, listas_productos_export)

    expected_condition_skus = set(condition_skus)
    if condition_list_skus:
        expected_condition_skus.update(condition_list_skus)

    actual_applier_skus = set(applier_skus)
    if applier_list_skus:
        actual_applier_skus.update(applier_list_skus)

    if condition_lists:
        if condition_lists_no_recon:
            agregar_detalle(
                detalles,
                "WARN",
                "CONDICIÓN",
                f"No se pudo reconstruir la composición SKU de las listas de condición: <span class='text-blue'>({', '.join(condition_lists_no_recon)})</span>"
            )

    if applier_lists:
        if applier_lists_no_recon:
            agregar_detalle(
                detalles,
                "WARN",
                "APPLIER",
                f"No se pudo reconstruir la composición SKU de las listas del applier: <span class='text-blue'>({', '.join(applier_lists_no_recon)})</span>"
            )

    if not expected_condition_skus:
        ok = False
        agregar_detalle(
            detalles,
            "ERR",
            "CONDICIÓN",
            "La condición no informa SKU explícitos ni fue posible reconstruir SKU desde lista de productos"
        )
        return ok

    if not actual_applier_skus:
        ok = False
        agregar_detalle(
            detalles,
            "ERR",
            "APPLIER",
            "El applier no informa SKU explícitos ni lista de productos. No puede quedar válido porque debe coincidir exactamente con la condición"
        )
        return ok

    faltan = sorted(expected_condition_skus - actual_applier_skus)
    sobran = sorted(actual_applier_skus - expected_condition_skus)

    if not faltan and not sobran:
        if applier_lists and condition_lists:
            listas_ok = sorted(set(applier_lists) & set(condition_lists))
            nombre_lista = listas_ok[0] if listas_ok else applier_lists[0]
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"Aplicador coincide con lista de productos <span class='text-blue'>({nombre_lista})</span>"
            )
        else:
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"El applier coincide exactamente con la condición en productos: <span class='text-blue'>({', '.join(sorted(expected_condition_skus))})</span>"
            )
    else:
        ok = False
        agregar_detalle(
            detalles,
            "ERR",
            "APPLIER",
            "Los productos del applier no coinciden exactamente con los productos de la condición"
        )
        if faltan:
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                f"Faltan SKU en applier respecto a condición: <span class='text-blue'>({', '.join(faltan)})</span>"
            )
        if sobran:
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                f"Sobran SKU en applier respecto a condición: <span class='text-blue'>({', '.join(sobran)})</span>"
            )

    return ok


# ============================================================
# FECHAS
# ============================================================

def evaluar_fechas(fi_excel, ff_excel, fi_export, ff_export, detalles, grupo="FECHAS"):
    inicio_ok = (fi_excel == fi_export)
    fin_ok = (ff_excel == ff_export)

    if inicio_ok:
        agregar_detalle(detalles, "OK", grupo,
            f"Fecha Inicio Excel ({fi_excel}) y Export ({fi_export}) coinciden")
    else:
        if fin_ok:
            agregar_detalle(detalles, "WARN", grupo,
                f"Fecha Inicio Excel ({fi_excel}) y Export ({fi_export}) no coinciden, pero la Fecha Fin sí coincide. Posible extensión")
        else:
            agregar_detalle(detalles, "ERR", grupo,
                f"Fecha Inicio Excel ({fi_excel}) no coincide con Export ({fi_export})")

    if fin_ok:
        agregar_detalle(detalles, "OK", grupo,
            f"Fecha Fin Excel ({ff_excel}) y Export ({ff_export}) coinciden")
    else:
        agregar_detalle(detalles, "ERR", grupo,
            f"Fecha Fin Excel ({ff_excel}) no coincide con Export ({ff_export})")

    return inicio_ok, fin_ok


def buscar_columna(df, nombres_posibles):
    columnas_norm = {
        c: str(c).replace("\n", " ").replace("  ", " ").strip().upper()
        for c in df.columns
    }
    nombres_norm = [
        str(n).replace("\n", " ").replace("  ", " ").strip().upper()
        for n in nombres_posibles
    ]
    for col_real, col_norm in columnas_norm.items():
        if col_norm in nombres_norm:
            return col_real
    return None


def extraer_productos_excel(grupo):
    col_prod = buscar_columna(grupo, [
        "CÓDIGO PRODUCTO", "Código Producto", "Código\nProducto", "SKU", "COLUMNA C"
    ])
    productos = set()
    if not col_prod:
        return productos
    for v in grupo[col_prod]:
        if es_vacio(v):
            continue
        sku = normalizar_sku(v)
        if sku:
            productos.add(sku)
    return productos


def inferir_tipo_descuento(valor):
    txt = normalizar_texto(valor)
    txt = txt.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")

    if "PACK NOMINAL" in txt:
        return "PACK_NOMINAL"

    pack = extraer_mecanica_pack(txt)
    if pack:
        return "PACK"

    if "2DA" in txt or "DCTO 2DA" in txt or "DECTO 2DA" in txt or "DESCUENTO 2DA" in txt:
        return "2DA"
    if "NOMINAL" in txt or "PVP FIJO" in txt or "PRECIO FIJO" in txt:
        return "NOMINAL"
    if "PORCENTUAL" in txt or txt == "%" or txt == "PORCENTAJE":
        return "PORCENTUAL"
    return txt


# ============================================================
# PARSEO EXPORT
# ============================================================

def convertir_txt_a_xml_con_root(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        contenido = f.read()

    pattern = r"<uy\.com\.geocom\.geopromotion\.service\.promotion\.Promotion>(.*?)</uy\.com\.geocom\.geopromotion\.service\.promotion\.Promotion>"
    bloques = re.findall(pattern, contenido, re.DOTALL)

    root = ET.Element("Root")
    for bloque in bloques:
        try:
            nodo = ET.fromstring("<Promotion>" + bloque + "</Promotion>")
            root.append(nodo)
        except Exception:
            pass

    return ET.ElementTree(root), contenido


def parsear_listas_productos(raw_text):
    """
    Intenta reconstruir listas de productos desde el export.
    Si no encuentra composición SKU, devuelve la lista vacía para ese nombre.
    """
    resultado = {}

    for m in re.finditer(r"<uy\.com\.geocom\.geopromotion\.service\.list\.ProductList>.*?<name>(.*?)</name>.*?</uy\.com\.geocom\.geopromotion\.service\.list\.ProductList>", raw_text, re.DOTALL):
        nombre = normalizar_texto(m.group(1))
        if nombre and nombre not in resultado:
            resultado[nombre] = set()

    patrones = [
        re.compile(r"<productListName>(.*?)</productListName>.*?<sku(?:Field)?>(\d+)</sku(?:Field)?>", re.DOTALL),
        re.compile(r"<name>(.*?)</name>.*?<productCode>(\d+)</productCode>", re.DOTALL),
        re.compile(r"<name>(.*?)</name>.*?<skuField>(\d+)</skuField>", re.DOTALL),
    ]

    for patron in patrones:
        for m in patron.finditer(raw_text):
            nombre = normalizar_texto(m.group(1))
            sku = normalizar_texto(m.group(2))
            if nombre:
                resultado.setdefault(nombre, set()).add(sku)

    return resultado


def parsear_mensajes(raw_text):
    mensajes = {}
    patron = re.compile(r"<uy\.com\.geocom\.geopromotion\.service\.message\.Message>.*?<name>(.*?)</name>.*?<message>(.*?)</message>.*?<output>(.*?)</output>.*?</uy\.com\.geocom\.geopromotion\.service\.message\.Message>", re.DOTALL)
    for m in patron.finditer(raw_text):
        nombre = normalizar_texto(m.group(1))
        if not nombre:
            continue
        mensajes[nombre] = {
            "name": m.group(1).strip(),
            "text": re.sub(r"\s+", " ", m.group(2)).strip(),
            "output": re.sub(r"\s+", " ", m.group(3)).strip(),
        }
    return mensajes


# ============================================================
# PARSEO PROMOS DESDE EXPORT
# ============================================================

def parsear_promos(tree, export_name=None):
    promos = []

    def clean(tag):
        return tag.split("}")[-1].split(".")[-1]

    def get_text(node, path):
        n = node.find(path) if node is not None else None
        return n.text.strip() if n is not None and n.text else None

    def get_bool_from_head(promo_node, tag):
        head = promo_node.find(".//promotionHead")
        if head is None:
            return False
        n = head.find(tag)
        return bool(n is not None and n.text and n.text.strip().lower() == "true")

    FILTER = "uy.com.geocom.geopromotion.service.promotion.Filter"

    for promo in tree.getroot().findall(".//{*}Promotion"):
        head = promo.find(".//promotionHead")
        creation_user = get_text(head, "creationUser") or "-"
        enabled_raw = get_text(head, "enabled")
        enabled = bool(enabled_raw and enabled_raw.lower() == "true")

        d = {
            "id": get_text(promo, ".//id"),
            "creationUser": creation_user,
            "enabled": enabled,
            "startDate": limpiar_fecha(get_text(promo, ".//startDate")),
            "endDate": limpiar_fecha(get_text(promo, ".//endDate")),
            "dontCompete": get_bool_from_head(promo, "dontCompete"),
            "competesByPromotion": get_bool_from_head(promo, "competesByPromotion"),
            "unitCompetence": get_bool_from_head(promo, "unitCompetence"),
            "__tipo_competencia": (
                "No Compite" if get_bool_from_head(promo, "dontCompete") else
                "Comp. X Promociones" if get_bool_from_head(promo, "competesByPromotion") else
                "Comp. X Unidades" if get_bool_from_head(promo, "unitCompetence") else
                "Comp. X Producto"
            ),
            "__export_origen": export_name,
            "__xml": ET.tostring(promo, encoding="unicode"),
            "name": get_text(promo, ".//name"),
            "description": get_text(promo, ".//description"),
            "area_name": get_text(head, "areaName") or "-",
            "productLists": [],
            "locales": [],
            "localLists": [],
            "skuFields": [],
            "condition_skus": [],
            "condition_quantity": None,
            "condition_item_quantity": None,
            "condition_product_lists": [],
            "fixAmount": None,
            "percentage": None,
            "applier_type": None,
            "applier_amount": None,
            "applier_percentage": None,
            "applier_quantity": None,
            "applier_to_quantity": False,
            "applier_strategy": None,
            "applier_skus": [],
            "applier_product_lists": [],
            "message_applier_name": None,
            "message_output": None,
            "message_text": None,
        }

        for flt in promo.findall(f".//{FILTER}"):
            field = get_text(flt, "fieldID")
            value = get_text(flt, "value")
            if not field or value is None:
                continue
            f = field.lower()
            nv = normalizar_texto(value)

            if "productlist" in f:
                d["productLists"].append(nv)
                d["condition_product_lists"].append(nv)
            elif f == "localfield":
                d["locales"].append(normalizar_local(value))
            elif f == "locallistfield":
                d["localLists"].append(nv)
            elif "sku" in f:
                d["skuFields"].append(nv)
                d["condition_skus"].append(nv)
            elif f == "quantityfield":
                try:
                    d["condition_quantity"] = int(float(value))
                except Exception:
                    pass
            elif f == "itemquantityfield":
                try:
                    d["condition_item_quantity"] = int(float(value))
                except Exception:
                    pass

        for ap in promo.iter():
            tag_name = clean(ap.tag)

            if tag_name == "FixAmountDiscountApplier":
                d["applier_type"] = "FIX_AMOUNT"
                amount_node = ap.find("./amount")
                qty_node = ap.find("./quantity")
                tq_node = ap.find("./toQuantity")
                strategy_node = ap.find("./strategy")
                if amount_node is not None and amount_node.text:
                    d["applier_amount"] = float(amount_node.text)
                    d["fixAmount"] = int(float(amount_node.text))
                if qty_node is not None and qty_node.text:
                    d["applier_quantity"] = float(qty_node.text)
                if tq_node is not None and tq_node.text:
                    d["applier_to_quantity"] = tq_node.text.strip().lower() == "true"
                if strategy_node is not None and strategy_node.text:
                    try:
                        d["applier_strategy"] = int(float(strategy_node.text))
                    except Exception:
                        pass
                for flt in ap.findall(f".//{FILTER}"):
                    fid = get_text(flt, "fieldID")
                    val = get_text(flt, "value")
                    if not fid or val is None:
                        continue
                    fid_norm = fid.lower()
                    val_norm = normalizar_texto(val)
                    if fid == "skuField" or "sku" in fid_norm:
                        d["applier_skus"].append(val_norm)
                    elif "productlist" in fid_norm:
                        d["applier_product_lists"].append(val_norm)

            elif tag_name == "PercentageDiscountApplier":
                d["applier_type"] = "PERCENTAGE"
                pn = ap.find("./percentage")
                qty_node = ap.find("./quantity")
                tq_node = ap.find("./toQuantity")
                strategy_node = ap.find("./strategy")
                if pn is not None and pn.text:
                    d["applier_percentage"] = float(pn.text)
                    d["percentage"] = float(pn.text)
                if qty_node is not None and qty_node.text:
                    d["applier_quantity"] = float(qty_node.text)
                if tq_node is not None and tq_node.text:
                    d["applier_to_quantity"] = tq_node.text.strip().lower() == "true"
                if strategy_node is not None and strategy_node.text:
                    try:
                        d["applier_strategy"] = int(float(strategy_node.text))
                    except Exception:
                        pass
                for flt in ap.findall(f".//{FILTER}"):
                    fid = get_text(flt, "fieldID")
                    val = get_text(flt, "value")
                    if not fid or val is None:
                        continue
                    fid_norm = fid.lower()
                    val_norm = normalizar_texto(val)
                    if fid == "skuField" or "sku" in fid_norm:
                        d["applier_skus"].append(val_norm)
                    elif "productlist" in fid_norm:
                        d["applier_product_lists"].append(val_norm)

            elif tag_name == "AmountDiscountApplier":
                d["applier_type"] = "AMOUNT"
                amount_node = ap.find("./amount")
                qty_node = ap.find("./quantity")
                tq_node = ap.find("./toQuantity")
                strategy_node = ap.find("./strategy")
                if amount_node is not None and amount_node.text:
                    d["applier_amount"] = float(amount_node.text)
                if qty_node is not None and qty_node.text:
                    d["applier_quantity"] = float(qty_node.text)
                if tq_node is not None and tq_node.text:
                    d["applier_to_quantity"] = tq_node.text.strip().lower() == "true"
                if strategy_node is not None and strategy_node.text:
                    try:
                        d["applier_strategy"] = int(float(strategy_node.text))
                    except Exception:
                        pass
                for flt in ap.findall(f".//{FILTER}"):
                    fid = get_text(flt, "fieldID")
                    val = get_text(flt, "value")
                    if not fid or val is None:
                        continue
                    fid_norm = fid.lower()
                    val_norm = normalizar_texto(val)
                    if fid == "skuField" or "sku" in fid_norm:
                        d["applier_skus"].append(val_norm)
                    elif "productlist" in fid_norm:
                        d["applier_product_lists"].append(val_norm)

            elif tag_name == "MessageApplier":
                d["applier_type"] = "MESSAGE"
                d["message_applier_name"] = get_text(ap, "messageName")
                d["message_output"] = get_text(ap, "messageOutput")

        d["productLists"] = sorted(set(filter(None, d["productLists"])))
        d["locales"] = sorted(set(filter(None, d["locales"])))
        d["localLists"] = sorted(set(filter(None, d["localLists"])))
        d["skuFields"] = sorted(set(filter(None, d["skuFields"])))
        d["condition_skus"] = sorted(set(filter(None, d["condition_skus"])))
        d["condition_product_lists"] = sorted(set(filter(None, d["condition_product_lists"])))
        d["applier_skus"] = sorted(set(filter(None, d["applier_skus"])))
        d["applier_product_lists"] = sorted(set(filter(None, d["applier_product_lists"])))

        promos.append(d)

    return promos


def cargar_promos_desde_exports(export_dir):
    promos = []
    listas_productos_export = {}
    mensajes_export = {}

    for exp in os.listdir(export_dir):
        if not exp.endswith(".txt"):
            continue
        ruta = os.path.join(export_dir, exp)
        tree, raw_text = convertir_txt_a_xml_con_root(ruta)
        promos_archivo = parsear_promos(tree, export_name=exp)
        promos.extend(promos_archivo)
        listas = parsear_listas_productos(raw_text)
        for nombre, skus in listas.items():
            listas_productos_export.setdefault(nombre, set()).update(skus)
        mensajes_export.update(parsear_mensajes(raw_text))

    for promo in promos:
        nombre_msg = normalizar_texto(promo.get("message_applier_name"))
        if nombre_msg and nombre_msg in mensajes_export:
            promo["message_text"] = mensajes_export[nombre_msg].get("text")
            promo["message_output"] = promo.get("message_output") or mensajes_export[nombre_msg].get("output")

    return promos, listas_productos_export


# ============================================================
# DETECCIÓN EXCEL TRADICIONAL
# ============================================================

COLUMNAS_EVENTOS = ["RC", "ID GEO", "LOCAL", "MARCA", "LISTA PRODUCTOS", "LISTA LOCAL"]


def detectar_fila_encabezado(df):
    for i in range(len(df)):
        fila = df.iloc[i].astype(str).str.upper()
        count = sum(1 for c in COLUMNAS_EVENTOS if any(c in x for x in fila))
        if count >= 2:
            return i
    return None


def leer_hoja_eventos(path_excel):
    try:
        xls = pd.ExcelFile(path_excel)
    except Exception as e:
        print("ERROR leyendo Excel:", path_excel, e)
        return None

    eventos = []

    for hoja in xls.sheet_names:
        try:
            df_raw = pd.read_excel(path_excel, sheet_name=hoja, header=None)
        except Exception as e:
            print(f"ERROR hoja {hoja}: {e}")
            continue

        header_row = detectar_fila_encabezado(df_raw)
        if header_row is None:
            continue

        df = pd.read_excel(path_excel, sheet_name=hoja, header=header_row)
        df.columns = [normalizar_texto(c) for c in df.columns]

        if "RC" in df.columns and "ID GEO" in df.columns:
            eventos.append(df)

    if eventos:
        return pd.concat(eventos, ignore_index=True)
    return None


# ============================================================
# VALIDACIÓN TRADICIONAL
# ============================================================

def validar_promocion_tradicional(id_geo, df_promo, promo_export, productos_excel, locales_excel, listas_local_excel):
    detalles = []
    ok = True

    id_excel = normalizar_local(id_geo)
    id_export = normalizar_local(promo_export["id"])

    if id_excel != id_export:
        ok = False
        agregar_detalle(detalles, "ERR", "ID", f"ID GEO Excel <span class='text-blue'>({id_excel})</span> es distinto al ID Export <span class='text-blue'>({id_export})</span>")
    else:
        agregar_detalle(detalles, "OK", "ID", f"ID GEO Excel <span class='text-blue'>({id_excel})</span> y Export coinciden")

    fi_excel = normalizar_fecha_excel(df_promo["FECHA DE INICIO EVENTO"].iloc[0])
    ff_excel = normalizar_fecha_excel(df_promo["FECHA TERMINO EVENTO"].iloc[0])
    inicio_ok, fin_ok = evaluar_fechas(fi_excel, ff_excel, promo_export["startDate"], promo_export["endDate"], detalles)
    if not inicio_ok or not fin_ok:
        ok = False

    if promo_export.get("productLists"):
        for lp in promo_export["productLists"]:
            agregar_detalle(detalles, "INFO", "LISTA PRODUCTOS",
                f"Lista de Productos <span class='text-blue'>({lp})</span> utilizada en Export. En flujo tradicional la validación de lista es informativa")

    listas_excel = sorted({
        normalizar_texto(v)
        for v in df_promo.get("LISTA LOCAL", [])
        if not es_vacio(v)
    })

    locales_sin_lista = sorted({
        normalizar_local(row["LOCAL"])
        for _, row in df_promo.iterrows()
        if es_vacio(row.get("LISTA LOCAL"))
    })

    listas_local_export = [normalizar_texto(v) for v in promo_export.get("localLists", [])]

    if listas_excel:
        for lista in listas_excel:
            if lista in listas_local_export:
                agregar_detalle(detalles, "OK", "LOCALES", f"LISTA LOCAL <span class='text-blue'>({lista})</span> del Excel coincide con Export")
            else:
                ok = False
                agregar_detalle(detalles, "ERR", "LOCALES", f"LISTA LOCAL <span class='text-blue'>({lista})</span> del Excel no coincide con Export")

        for loc in locales_sin_lista:
            ok = False
            agregar_detalle(detalles, "ERR", "LOCALES", f"LOCAL <span class='text-blue'>({loc})</span> no se encuentra asociado a LISTA LOCAL en el Excel")
    else:
        locales_excel_vals = sorted({normalizar_local(v) for v in df_promo["LOCAL"] if not es_vacio(v)})
        locales_export = sorted({normalizar_local(v) for v in promo_export.get("locales", [])})
        ok_l, det_l = validar_multivalor(locales_excel_vals, locales_export, "LOCAL")
        if not ok_l:
            ok = False
        for t, m in det_l:
            agregar_detalle(detalles, t, "LOCALES", m)

    cell_value = df_promo["DESCUENTO"].iloc[0]
    excel_val = parsear_porcentaje_excel(cell_value)
    export_val = promo_export.get("percentage")
    if excel_val is None:
        ok = False
        agregar_detalle(detalles, "ERR", "DESCUENTO", f"No se pudo interpretar descuento del Excel <span class='text-blue'>({cell_value})</span>")
    elif export_val is not None and floats_iguales(excel_val, export_val):
        agregar_detalle(detalles, "OK", "DESCUENTO",
            f"Descuento porcentual Excel <span class='text-blue'>({excel_val*100:.0f}%)</span> coincide con Export <span class='text-blue'>({export_val*100:.0f}%)</span>")
    else:
        ok = False
        agregar_detalle(detalles, "ERR", "DESCUENTO",
            f"Descuento porcentual Excel <span class='text-blue'>({excel_val*100:.0f}%)</span> no coincide con Export <span class='text-blue'>({(export_val or 0)*100:.0f}%)</span>")

    return ok, detalles


# ============================================================
# HOJA COMPLETAR
# ============================================================

COLUMNAS_CLAVE_COMPLETAR = {
    "N°", "CÓDIGO PRODUCTO", "CÓDIGO PRODUCTO", "DESCRIPTOR", "PVP FIJO UNITARIO",
    "# UNIDADES PACK", "PVP OFERTA PACK", "TIPO DE DESCUENTO", "DESCUENTO PORCENTUAL",
    "DESCUENTO NOMINAL PACK BRUTO", "COBERTURA LOCALES", "F. INICIO", "F. TÉRMINO",
    "ID A FACTURAR", "ID LISTA CLIENTE", "ID GEOCOM", "ID LISTA GEO", "ID LISTA LOCALES",
    "ID LISTA PRODUCTOS"
}


def leer_hoja_completar(path_excel):
    try:
        xls = pd.ExcelFile(path_excel)
        if "Completar" not in xls.sheet_names:
            return None
    except Exception:
        return None

    try:
        df_raw = pd.read_excel(path_excel, sheet_name="Completar", header=None)
    except Exception:
        return None

    if df_raw.empty:
        return None

    df_norm = df_raw.fillna("").astype(str).map(lambda x: x.replace("\n", " ").strip())
    max_filas = min(10, len(df_norm))

    for fila in range(max_filas):
        fila_vals = set(df_norm.iloc[fila].tolist())
        coincidencias = 0
        for esperado in COLUMNAS_CLAVE_COMPLETAR:
            for celda in fila_vals:
                if esperado.replace(" ", "").lower() == celda.replace(" ", "").lower():
                    coincidencias += 1
                    break

        if coincidencias >= 3:
            df = pd.read_excel(path_excel, sheet_name="Completar", header=fila)
            df = limpiar_dataframe_columnas(df)
            print(Fore.GREEN + f"✓ Hoja COMPLETAR detectada: {os.path.basename(path_excel)} (fila {fila+1})")
            return df

    return None


def leer_hoja_imput(path_excel):
    try:
        xls = pd.ExcelFile(path_excel)
    except Exception:
        return None

    hoja_objetivo = None
    for hoja in xls.sheet_names:
        hoja_norm = normalizar_texto(hoja)
        if hoja_norm in {"IMPUT", "INPUT"}:
            hoja_objetivo = hoja
            break

    if not hoja_objetivo:
        return None

    try:
        df_raw = pd.read_excel(path_excel, sheet_name=hoja_objetivo, header=None)
    except Exception:
        return None

    if df_raw.empty:
        return None

    df_norm = df_raw.fillna("").astype(str).map(lambda x: x.replace("\n", " ").strip())
    max_filas = min(12, len(df_norm))

    for fila in range(max_filas):
        fila_vals = {normalizar_texto(v) for v in df_norm.iloc[fila].tolist()}
        if "AREARESPONSABLE" in fila_vals and any("GEOCOM" in v for v in fila_vals):
            try:
                df = pd.read_excel(path_excel, sheet_name=hoja_objetivo, header=fila)
                df = limpiar_dataframe_columnas(df)
                print(Fore.GREEN + f"✓ Hoja IMPUT detectada: {os.path.basename(path_excel)} (fila {fila+1})")
                return df
            except Exception:
                return None

    return None


def construir_mapa_area_responsable(excel_files):
    mapa = {}
    for file in excel_files:
        df_imput = leer_hoja_imput(file)
        if df_imput is None or df_imput.empty:
            continue

        col_area = buscar_columna(df_imput, ["AREARESPONSABLE", "AREA RESPONSABLE", "AREA"])
        col_id_geo = obtener_columna_id_geocom(df_imput)
        col_id_fact = buscar_columna(df_imput, ["ID A FACTURAR", "ID FACTURAR", "ID A Facturar"])

        if not col_area:
            continue

        for _, row in df_imput.iterrows():
            area = normalizar_texto(row.get(col_area))
            if not area:
                continue

            id_geo = normalizar_local(row.get(col_id_geo)) if col_id_geo else None
            id_fact = normalizar_local(row.get(col_id_fact)) if col_id_fact else None

            if id_geo:
                mapa[id_geo] = area
            if id_fact:
                mapa[id_fact] = area

    return mapa


# ============================================================
# VALIDACIÓN COMPLETAR CONTRA EXPORT
# ============================================================

def validar_promocion_completar(id_geo, grupo, promo, listas_productos_export, mapa_area_responsable=None, promos_por_id=None, retornar_msje_data=False):
    detalles = []
    ok = True

    id_excel = normalizar_local(id_geo)
    msje_popup_data = construir_msje_popup_data(id_excel)
    id_export = normalizar_local(promo["id"])

    if id_excel != id_export:
        ok = False
        agregar_detalle(detalles, "ERR", "ID", f"ID Geocom Excel <span class='text-blue'>({id_excel})</span> es distinto al ID Export <span class='text-blue'>({id_export})</span>")
    else:
        agregar_detalle(detalles, "OK", "ID", f"ID Geocom Excel y Export coinciden <span class='text-blue'>({id_excel})</span>")

    col_fact = buscar_columna(grupo, ["ID A FACTURAR", "ID a Facturar", "ID FACTURAR"])
    id_fact = normalizar_local(grupo[col_fact].iloc[0]) if col_fact else None
    if id_fact and id_fact != id_excel:
        agregar_detalle(detalles, "WARN", "FACTURAR", f"ID a Facturar <span class='text-blue'>({id_fact})</span> es distinto de ID Geocom <span class='text-blue'>({id_excel})</span>")
    elif id_fact:
        agregar_detalle(detalles, "OK", "FACTURAR", f"ID a Facturar coincide <span class='text-blue'>({id_fact})</span>")

    col_fi = buscar_columna(grupo, ["F. INICIO", "F. Inicio", "Fecha Inicio", "FECHA DE INICIO", "F.Inicio"])
    col_ff = buscar_columna(grupo, ["F. TÉRMINO", "F. Término", "Fecha Término", "FECHA DE TÉRMINO", "F.Termino"])

    fi_excel = normalizar_fecha_excel(grupo[col_fi].iloc[0]) if col_fi else None
    ff_excel = normalizar_fecha_excel(grupo[col_ff].iloc[0]) if col_ff else None

    if not col_fi:
        ok = False
        agregar_detalle(detalles, "ERR", "FECHAS", "No existe columna Fecha Inicio en Excel")
    if not col_ff:
        ok = False
        agregar_detalle(detalles, "ERR", "FECHAS", "No existe columna Fecha Fin en Excel")
    if col_fi and col_ff:
        inicio_ok, fin_ok = evaluar_fechas(fi_excel, ff_excel, promo["startDate"], promo["endDate"], detalles)
        if not inicio_ok or not fin_ok:
            ok = False

    col_tipo = buscar_columna(grupo, ["TIPO DE DESCUENTO", "DESCUENTO"])
    tipo_desc_raw = grupo[col_tipo].iloc[0] if col_tipo else ""
    tipo_desc = inferir_tipo_descuento(tipo_desc_raw)
    mecanica_pack = extraer_mecanica_pack(tipo_desc_raw)
    pack_label = f"{mecanica_pack[0]}x{mecanica_pack[1]}" if mecanica_pack else "PACK"

    col_desc = buscar_columna(grupo, ["DESCUENTO PORCENTUAL", "DESCUENTO P", "DESCUENTOP", "DESCUENTO\nPORCENTUAL"])
    col_pack = buscar_columna(grupo, ["# UNIDADES PACK", "#\nUNIDADES\nPACK", "UNIDADES PACK"])
    col_pvp = buscar_columna(grupo, ["PVP OFERTA PACK", "PVP\nOFERTA\nPACK"])
    col_desc_nom_pack = buscar_columna(grupo, [
        "DESCUENTO NOMINAL PACK BRUTO",
        "DESCUENTO NOMINAL PACK",
        "DESCUENTO NOMINAL"
    ])
    col_lista_prod = buscar_columna(grupo, ["ID LISTA PRODUCTOS", "ID LISTA GEO", "ID Lista Productos", "ID Lista Geo"])

    porcentaje_excel = parsear_porcentaje_excel(grupo[col_desc].iloc[0]) if col_desc else None
    cantidad_excel = None
    if col_pack and not es_vacio(grupo[col_pack].iloc[0]):
        try:
            cantidad_excel = int(float(grupo[col_pack].iloc[0]))
        except Exception:
            cantidad_excel = None

    monto_pack_excel = None
    if col_pvp and not es_vacio(grupo[col_pvp].iloc[0]):
        try:
            monto_pack_excel = float(str(grupo[col_pvp].iloc[0]).replace(",", "."))
        except Exception:
            monto_pack_excel = None

    descuento_nominal_pack_excel = None
    if col_desc_nom_pack and not es_vacio(grupo[col_desc_nom_pack].iloc[0]):
        try:
            descuento_nominal_pack_excel = float(str(grupo[col_desc_nom_pack].iloc[0]).replace(",", "."))
        except Exception:
            descuento_nominal_pack_excel = None

    descuento_bruto_excel_q = obtener_valor_descuento_bruto_q(grupo)

    productos_excel = extraer_productos_excel(grupo)
    condition_skus = normalizar_lista_skus(promo.get("condition_skus", []))
    applier_skus = normalizar_lista_skus(promo.get("applier_skus", []))
    applier_product_lists = [normalizar_texto(x) for x in promo.get("applier_product_lists", []) if normalizar_texto(x)]

    col_area_excel = buscar_columna(grupo, ["AREARESPONSABLE", "AREA RESPONSABLE", "AREA"])
    area_responsable = normalizar_texto(grupo[col_area_excel].iloc[0]) if col_area_excel and not es_vacio(grupo[col_area_excel].iloc[0]) else ""
    if not area_responsable and mapa_area_responsable:
        area_responsable = normalizar_texto(mapa_area_responsable.get(id_excel))
        if not area_responsable and id_fact:
            area_responsable = normalizar_texto(mapa_area_responsable.get(id_fact))
    if not area_responsable:
        area_responsable = normalizar_texto(promo.get("area_name"))

    combo_precio_excel = extraer_mecanica_combo_precio(tipo_desc_raw)
    es_mensaje_popup = promo.get("applier_type") == "MESSAGE"
    es_farma_combo_precio = (
        area_responsable == "FARMA" and
        combo_precio_excel is not None and
        promo.get("applier_type") == "FIX_AMOUNT"
    )

    applier_pct_nodo = promo.get("applier_percentage")
    applier_pct_tecnico = (
        calcular_porcentaje_tecnico_2da(applier_pct_nodo, promo.get("applier_quantity"))
        if tipo_desc == "2DA"
        else applier_pct_nodo
    )

    # --------------------------------------------------------
    # LEYENDAS BASE
    # --------------------------------------------------------
    nombre_lista_excel = normalizar_texto(grupo[col_lista_prod].iloc[0]) if col_lista_prod and not es_vacio(grupo[col_lista_prod].iloc[0]) else ""
    nombre_lista_export = normalizar_texto(promo["productLists"][0]) if promo.get("productLists") else ""
    msje_popup_data["nombre_lista_excel"] = nombre_lista_excel

    ok_competencia = validar_competencia_por_area(detalles, promo, area_responsable, tipo_desc, productos_excel, cantidad_excel)
    if not ok_competencia:
        ok = False

    agregar_detalle(
        detalles,
        "INFO",
        "LEYENDA",
        construir_leyenda_excel_compat(
            tipo_desc_raw,
            productos_excel,
            cantidad_excel,
            porcentaje_excel,
            descuento_bruto_excel_q,
            monto_pack_excel,
            nombre_lista_excel
        )
    )

    agregar_detalle(
        detalles,
        "INFO",
        "LEYENDA",
        construir_leyenda_condicion_compat(
            condition_skus,
            promo.get("condition_quantity"),
            nombre_lista_export
        )
    )

    agregar_detalle(
        detalles,
        "INFO",
        "LEYENDA",
        construir_leyenda_applier_compat(
            tipo_desc,
            promo,
            applier_skus,
            applier_product_lists,
            porcentaje_excel,
            applier_pct_nodo,
            applier_pct_tecnico,
            descuento_bruto_excel_q,
            monto_pack_excel
        )
    )

    agregar_detalle(
        detalles,
        "INFO",
        "EXCEL",
        construir_resumen_excel_limpio(
            tipo_desc_raw,
            tipo_desc,
            nombre_lista_excel,
            productos_excel,
            cantidad_excel,
            porcentaje_excel,
            descuento_bruto_excel_q,
            monto_pack_excel
        )
    )

    agregar_detalle(
        detalles,
        "INFO",
        "CONDICIÓN",
        construir_resumen_condicion_limpio(
            condition_skus,
            promo.get("condition_quantity"),
            nombre_lista_export
        )
    )

    agregar_detalle(
        detalles,
        "INFO",
        "APPLIER",
        construir_resumen_applier_limpio(
            tipo_desc,
            promo,
            applier_skus,
            applier_product_lists,
            porcentaje_excel,
            applier_pct_nodo,
            applier_pct_tecnico
        )
    )

    if es_mensaje_popup:
        ok_mensaje = validar_promocion_mensaje(detalles, grupo, promo, productos_excel, nombre_lista_excel, nombre_lista_export, listas_productos_export)
        if not ok_mensaje:
            ok = False
        return (ok, detalles, msje_popup_data) if retornar_msje_data else (ok, detalles)

    id_msje_asociado, promo_msje_asociada = obtener_promo_msje_asociada(grupo, promos_por_id)
    msje_popup_data = construir_msje_popup_data(
        id_excel,
        id_msje_asociado=id_msje_asociado,
        promo_msje_asociada=promo_msje_asociada,
        productos_excel=productos_excel,
        nombre_lista_excel=nombre_lista_excel,
    )

    # --------------------------------------------------------
    # LISTA PRODUCTOS / CONDICIÓN
    # --------------------------------------------------------
    if nombre_lista_excel or nombre_lista_export:
        if nombre_lista_excel == nombre_lista_export and nombre_lista_excel:
            agregar_detalle(detalles, "OK", "LISTA PRODUCTOS", f"ID Lista Productos Excel <span class='text-blue'>({nombre_lista_excel})</span> coincide con la lista usada en Export")
        else:
            ok = False
            agregar_detalle(detalles, "ERR", "LISTA PRODUCTOS", f"ID Lista Productos Excel <span class='text-blue'>({nombre_lista_excel or '-'})</span> no coincide con la lista de Export <span class='text-blue'>({nombre_lista_export or '-'})</span>")

        productos_lista_export = listas_productos_export.get(nombre_lista_export, set())
        if productos_lista_export:
            productos_lista_export = normalizar_lista_skus(productos_lista_export)
            faltan_en_lista = sorted(productos_excel - productos_lista_export)
            sobran_en_lista = sorted(productos_lista_export - productos_excel)
            if faltan_en_lista:
                ok = False
                agregar_detalle(detalles, "ERR", "LISTA PRODUCTOS", f"La lista del Export no contiene todos los productos de columna C. Faltan: <span class='text-blue'>({', '.join(faltan_en_lista)})</span>")
            else:
                agregar_detalle(detalles, "OK", "LISTA PRODUCTOS", "La lista del Export contiene los productos esperados desde columna C")
            if sobran_en_lista:
                agregar_detalle(detalles, "WARN", "LISTA PRODUCTOS", f"La lista del Export contiene productos adicionales: <span class='text-blue'>({', '.join(sobran_en_lista)})</span>")
        else:
            agregar_detalle(detalles, "WARN", "LISTA PRODUCTOS", f"No se pudo reconstruir la composición SKU de la lista <span class='text-blue'>({nombre_lista_export or nombre_lista_excel})</span> desde el Export. Se valida nombre de lista, no composición")
    else:
        if condition_skus == productos_excel:
            agregar_detalle(detalles, "OK", "CONDICIÓN", "Los SKU de la condición coinciden con columna C del Excel")
        else:
            ok = False
            faltan = sorted(productos_excel - condition_skus)
            extras = sorted(condition_skus - productos_excel)
            if faltan:
                agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Faltan SKU en condición respecto a columna C: <span class='text-blue'>({', '.join(faltan)})</span>")
            if extras:
                agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Sobran SKU en condición respecto a columna C: <span class='text-blue'>({', '.join(extras)})</span>")

    ok_applier = validar_applier_vs_condicion(detalles, promo, listas_productos_export)
    if not ok_applier:
        ok = False

    if es_farma_combo_precio:
        ok_combo = validar_promocion_farma_combo_precio(detalles, promo, cantidad_excel, combo_precio_excel, nombre_lista_excel, nombre_lista_export)
        if not ok_combo:
            ok = False
        return (ok, detalles, msje_popup_data) if retornar_msje_data else (ok, detalles)

    # --------------------------------------------------------
    # REGLA PORCENTUAL
    # --------------------------------------------------------
    if tipo_desc == "PORCENTUAL":
        esperado = porcentaje_excel
        actual = promo.get("applier_percentage")
        unidades_pack = cantidad_excel if cantidad_excel not in {None, 0} else 1

        if esperado is None:
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", "No se pudo interpretar Descuento Porcentual del Excel")

        elif promo.get("applier_type") != "PERCENTAGE":
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", "La promo porcentual debe viajar con PercentageDiscountApplier")

        elif actual is None:
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", "El applier porcentual no informa porcentaje en el Export")

        elif floats_iguales(esperado, actual):
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"Porcentaje del applier coincide con columna P: Excel <span class='text-blue'>({formatear_porcentaje(esperado)})</span> = Export <span class='text-blue'>({formatear_porcentaje(actual)})</span>"
            )

        elif unidades_pack > 1 and floats_iguales(esperado, actual / unidades_pack):
            porcentaje_real = actual / unidades_pack
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"Porcentaje del applier validado por pack: Export <span class='text-blue'>({formatear_porcentaje(actual)})</span> / #UnidadesPack <span class='text-blue'>({unidades_pack})</span> = <span class='text-blue'>({formatear_porcentaje(porcentaje_real)})</span>, que coincide con Excel <span class='text-blue'>({formatear_porcentaje(esperado)})</span>"
            )
        else:
            ok = False
            if unidades_pack > 1:
                porcentaje_real = actual / unidades_pack
                agregar_detalle(
                    detalles,
                    "ERR",
                    "APPLIER",
                    f"Porcentaje del applier no coincide con columna P. Excel <span class='text-blue'>({formatear_porcentaje(esperado)})</span> vs Export directo <span class='text-blue'>({formatear_porcentaje(actual)})</span> y también vs Export/#UnidadesPack <span class='text-blue'>({formatear_porcentaje(porcentaje_real)})</span>"
                )
            else:
                agregar_detalle(
                    detalles,
                    "ERR",
                    "APPLIER",
                    f"Porcentaje del applier no coincide con columna P: Excel <span class='text-blue'>({formatear_porcentaje(esperado)})</span> vs Export <span class='text-blue'>({formatear_porcentaje(actual)})</span>"
                )

    # --------------------------------------------------------
    # REGLA 2DA UNIDAD
    # --------------------------------------------------------
    elif tipo_desc == "2DA":
        esperado_pct_comercial = porcentaje_excel
        esperado_qty = cantidad_excel
        area_2da_sin_division = area_responsable in {"BYCP", "FIDELIZACION"}

        actual_pct_nodo = promo.get("applier_percentage")
        actual_qty_cond = promo.get("condition_quantity")
        actual_qty_applier = promo.get("applier_quantity")

        qty_text = (
            str(int(float(actual_qty_applier)))
            if actual_qty_applier is not None else "-"
        )

        esperado_pct_nodo = None
        if esperado_pct_comercial is not None:
            try:
                if area_2da_sin_division:
                    esperado_pct_nodo = float(esperado_pct_comercial)
                elif esperado_qty not in {None, 0}:
                    esperado_pct_nodo = float(esperado_pct_comercial) / float(esperado_qty)
            except Exception:
                esperado_pct_nodo = None

        actual_pct_tecnico = calcular_porcentaje_tecnico_2da(actual_pct_nodo, actual_qty_applier)

        if promo.get("applier_type") != "PERCENTAGE":
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                "Descuento 2da unidad debe viajar con PercentageDiscountApplier"
            )

        if esperado_pct_comercial is None:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                "No se pudo interpretar Descuento Porcentual del Excel para 2da unidad"
            )
        elif esperado_pct_nodo is None:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                "No se pudo calcular el porcentaje esperado del nodo para 2da unidad"
            )
        elif floats_iguales(actual_pct_nodo, esperado_pct_nodo):
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"Aplicador 2da unidad correcto. "
                f"Excel comercial <span class='text-blue'>({formatear_porcentaje(esperado_pct_comercial)})</span> "
                f"con cantidad <span class='text-blue'>({esperado_qty})</span> "
                f"equivale a nodo export <span class='text-blue'>({formatear_porcentaje(esperado_pct_nodo)})</span>. "
                f"Export trae nodo <span class='text-blue'>({formatear_porcentaje(actual_pct_nodo)})</span> "
                f"y porcentaje técnico visible <span class='text-blue'>({formatear_porcentaje(actual_pct_tecnico)})</span>"
            )
        else:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                f"El aplicador de 2da unidad no coincide. "
                f"Excel comercial <span class='text-blue'>({formatear_porcentaje(esperado_pct_comercial)})</span> "
                f"con cantidad <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span> "
                f"debería viajar como nodo export <span class='text-blue'>({formatear_porcentaje(esperado_pct_nodo)})</span>, "
                f"pero Export trae nodo <span class='text-blue'>({formatear_porcentaje(actual_pct_nodo)})</span> "
                f"y técnico visible <span class='text-blue'>({formatear_porcentaje(actual_pct_tecnico)})</span>"
            )

        if esperado_qty is not None and actual_qty_cond == esperado_qty:
            agregar_detalle(
                detalles,
                "OK",
                "CONDICIÓN",
                f"Cantidad de la condición coincide con #UnidadesPack: "
                f"<span class='text-blue'>({esperado_qty})</span>"
            )
        else:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "CONDICIÓN",
                f"Cantidad de la condición no coincide con #UnidadesPack. "
                f"Excel <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span> "
                f"vs Export <span class='text-blue'>({actual_qty_cond if actual_qty_cond is not None else '-'})</span>"
            )

        qty_applier_ok = False
        if esperado_qty is not None and actual_qty_applier is not None:
            qty_applier_int = int(float(actual_qty_applier))
            if area_2da_sin_division:
                qty_applier_ok = qty_applier_int in {1, int(esperado_qty)}
            else:
                qty_applier_ok = qty_applier_int == int(esperado_qty)

        if qty_applier_ok:
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"Cantidad del applier correcta para 2da unidad: "
                f"<span class='text-blue'>({int(float(actual_qty_applier))})</span>"
            )
        else:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                f"Cantidad del applier no coincide con la regla esperada de 2da unidad. "
                f"Excel <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span> "
                f"vs Export <span class='text-blue'>({actual_qty_applier if actual_qty_applier is not None else '-'})</span>"
            )

    # --------------------------------------------------------
    # REGLA PACK
    # --------------------------------------------------------
    elif tipo_desc == "PACK":
        esperado_qty = cantidad_excel
        esperado_qty_applier = 1 if area_responsable == "BYCP" else cantidad_excel
        if area_responsable == "BYCP":
            esperado_amount = monto_pack_excel
            regla_pack = f"Regla BYCP aplicada: no se divide descuento/monto. PVPOfertaPack esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span>"
        else:
            esperado_amount = (monto_pack_excel / cantidad_excel) if monto_pack_excel is not None and cantidad_excel not in {None, 0} else None
            regla_pack = (
                f"Regla aplicada: PVPOfertaPack <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span> "
                f"/ #UnidadesPack <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span> "
                f"= Monto unitario esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span>"
            )
        actual_amount = promo.get("applier_amount")
        actual_qty = promo.get("condition_quantity")
        actual_qty_applier = promo.get("applier_quantity")

        agregar_detalle(
            detalles,
            "INFO",
            "APPLIER",
            f"Promo PACK detectada <span class='text-blue'>({pack_label})</span>. {regla_pack}"
        )

        if mecanica_pack and cantidad_excel is not None and mecanica_pack[0] != cantidad_excel:
            agregar_detalle(
                detalles,
                "WARN",
                "CONDICIÓN",
                f"La mecánica del texto <span class='text-blue'>({pack_label})</span> "
                f"no coincide con #UnidadesPack del Excel <span class='text-blue'>({cantidad_excel})</span>. "
                f"Se valida con #UnidadesPack."
            )

        if promo.get("applier_type") not in {"FIX_AMOUNT", "AMOUNT"}:
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", f"La promo PACK <span class='text-blue'>({pack_label})</span> debe viajar con FixAmountDiscountApplier o AmountDiscountApplier")

        if esperado_amount is None:
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", f"No se pudo calcular monto esperado del applier para PACK <span class='text-blue'>({pack_label})</span>")
        elif money_iguales(esperado_amount, actual_amount):
            agregar_detalle(
                detalles,
                "OK",
                "APPLIER",
                f"Monto del applier correcto para PACK <span class='text-blue'>({pack_label})</span>: "
                f"PVPOfertaPack <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span> "
                f"/ #UnidadesPack <span class='text-blue'>({esperado_qty})</span> "
                f"= <span class='text-blue'>({formatear_monto(esperado_amount)})</span>"
            )
        else:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                f"Monto del applier incorrecto para PACK <span class='text-blue'>({pack_label})</span>. "
                f"Esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span> "
                f"según PVPOfertaPack <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span> "
                f"/ #UnidadesPack <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span>, "
                f"pero Export trae <span class='text-blue'>({formatear_monto(actual_amount)})</span>"
            )

        if esperado_qty is not None and actual_qty == esperado_qty:
            agregar_detalle(detalles, "OK", "CONDICIÓN", f"Cantidad de condición coincide con #UnidadesPack: <span class='text-blue'>({esperado_qty})</span>")
        else:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "CONDICIÓN",
                f"Cantidad de condición no coincide con #UnidadesPack. "
                f"Excel <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span> "
                f"vs Export <span class='text-blue'>({actual_qty if actual_qty is not None else '-'})</span>"
            )

        if esperado_qty_applier is not None and actual_qty_applier is not None and int(float(actual_qty_applier)) == int(esperado_qty_applier):
            agregar_detalle(detalles, "OK", "APPLIER", f"Cantidad del applier correcta para PACK: <span class='text-blue'>({int(float(actual_qty_applier))})</span>")
        else:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                f"Cantidad del applier no coincide con la regla esperada de PACK. "
                f"Esperado <span class='text-blue'>({esperado_qty_applier if esperado_qty_applier is not None else '-'})</span> "
                f"vs Export <span class='text-blue'>({actual_qty_applier if actual_qty_applier is not None else '-'})</span>"
            )

    # --------------------------------------------------------
    # REGLA PACK NOMINAL
    # --------------------------------------------------------
    elif tipo_desc == "PACK_NOMINAL":
        esperado_qty = cantidad_excel
        actual_qty = promo.get("condition_quantity")
        actual_qty_applier = promo.get("applier_quantity")
        actual_amount = promo.get("applier_amount")
        applier_type = promo.get("applier_type")

        if applier_type not in {"AMOUNT", "FIX_AMOUNT"}:
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", "La promo PACK NOMINAL debe viajar con AmountDiscountApplier o FixAmountDiscountApplier")

        if esperado_qty is not None and actual_qty == esperado_qty:
            agregar_detalle(detalles, "OK", "CONDICIÓN", f"Cantidad de condición coincide con #UnidadesPack: <span class='text-blue'>({esperado_qty})</span>")
        else:
            ok = False
            agregar_detalle(detalles, "ERR", "CONDICIÓN", f"Cantidad de condición no coincide con #UnidadesPack. Excel <span class='text-blue'>({esperado_qty if esperado_qty is not None else '-'})</span> vs Export <span class='text-blue'>({actual_qty if actual_qty is not None else '-'})</span>")

        esperado_qty_applier = None
        esperado_amount = None
        regla_pack_nom = None

        if descuento_nominal_pack_excel is None or cantidad_excel in {None, 0}:
            ok = False
            agregar_detalle(detalles, "ERR", "APPLIER", "No se pudo calcular el descuento nominal esperado para PACK NOMINAL")
        else:
            if area_responsable == "BYCP":
                if applier_type == "AMOUNT":
                    esperado_qty_applier = 1
                    esperado_amount = descuento_nominal_pack_excel
                    regla_pack_nom = (
                        f"Regla BYCP aplicada con applier AMOUNT: no se divide el descuento. "
                        f"Descuento Nominal Pack Bruto esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span> "
                        f"con cantidad de applier <span class='text-blue'>(1)</span>"
                    )
                elif applier_type == "FIX_AMOUNT":
                    esperado_qty_applier = cantidad_excel
                    esperado_amount = monto_pack_excel
                    regla_pack_nom = (
                        f"Regla BYCP aplicada con applier FIX_AMOUNT: se valida precio oferta pack. "
                        f"PVPOfertaPack esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span> "
                        f"con cantidad de applier <span class='text-blue'>({cantidad_excel})</span>"
                    )
                else:
                    esperado_qty_applier = None
                    esperado_amount = None
                    regla_pack_nom = ""
            else:
                esperado_qty_applier = cantidad_excel
                esperado_amount = descuento_nominal_pack_excel / cantidad_excel
                regla_pack_nom = f"Regla aplicada: Descuento Nominal Pack Bruto <span class='text-blue'>({formatear_monto(descuento_nominal_pack_excel)})</span> / #UnidadesPack <span class='text-blue'>({cantidad_excel})</span> = Descuento unitario esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span>"

            if esperado_qty_applier is not None and actual_qty_applier is not None and int(float(actual_qty_applier)) == int(esperado_qty_applier):
                agregar_detalle(detalles, "OK", "APPLIER", f"Cantidad del applier correcta para PACK NOMINAL: <span class='text-blue'>({int(float(actual_qty_applier))})</span>")
            else:
                ok = False
                agregar_detalle(detalles, "ERR", "APPLIER", f"Cantidad del applier no coincide con la regla esperada de PACK NOMINAL. Esperado <span class='text-blue'>({esperado_qty_applier if esperado_qty_applier is not None else '-'})</span> vs Export <span class='text-blue'>({actual_qty_applier if actual_qty_applier is not None else '-'})</span>")

            agregar_detalle(
                detalles,
                "INFO",
                "APPLIER",
                f"Promo PACK NOMINAL detectada. {regla_pack_nom}"
            )

            if money_iguales(esperado_amount, actual_amount):
                agregar_detalle(
                    detalles,
                    "OK",
                    "APPLIER",
                    f"Monto del applier correcto para PACK NOMINAL: esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span> y Export <span class='text-blue'>({formatear_monto(actual_amount)})</span>"
                )
            else:
                ok = False
                agregar_detalle(
                    detalles,
                    "ERR",
                    "APPLIER",
                    f"Monto del applier incorrecto para PACK NOMINAL. Esperado <span class='text-blue'>({formatear_monto(esperado_amount)})</span> pero Export trae <span class='text-blue'>({formatear_monto(actual_amount)})</span>"
                )

    # --------------------------------------------------------
    # REGLA NOMINAL
    # --------------------------------------------------------
    elif tipo_desc == "NOMINAL":
        esperado_skus = productos_excel
        applier_skus_nominal = normalizar_lista_skus(promo.get("applier_skus", []))
        applier_list_skus_nominal, _ = reconstruir_skus_desde_listas(
            promo.get("applier_product_lists", []),
            listas_productos_export
        )
        applier_skus_nominal_total = set(applier_skus_nominal)
        applier_skus_nominal_total.update(applier_list_skus_nominal)

        if monto_pack_excel is None:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                "No se pudo interpretar PVPOfertaPack para validación nominal"
            )

        elif promo.get("applier_type") not in {"FIX_AMOUNT", "AMOUNT"}:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                "La promo nominal debe viajar con FixAmountDiscountApplier o AmountDiscountApplier"
            )

        elif not applier_skus_nominal_total:
            ok = False
            agregar_detalle(
                detalles,
                "ERR",
                "APPLIER",
                "La promo nominal no puede quedar válida si el applier no informa SKU ni lista de productos"
            )

        else:
            monto_ok = money_iguales(monto_pack_excel, promo.get("applier_amount"))
            skus_ok = applier_skus_nominal_total == esperado_skus

            if monto_ok and skus_ok:
                agregar_detalle(
                    detalles,
                    "OK",
                    "APPLIER",
                    "Monto y productos del applier coinciden con Excel"
                )
            else:
                ok = False

                if not monto_ok:
                    agregar_detalle(
                        detalles,
                        "ERR",
                        "APPLIER",
                        f"El monto del applier no coincide con el PVPOfertaPack del Excel. "
                        f"Excel PVPOfertaPack = <span class='text-blue'>({formatear_monto(monto_pack_excel)})</span> "
                        f"pero el Export trae <span class='text-blue'>({formatear_monto(promo.get('applier_amount'))})</span>. "
                        f"Posible error: se cargó el descuento bruto en vez del precio oferta."
                    )

                if not skus_ok:
                    faltan_nominal = sorted(esperado_skus - applier_skus_nominal_total)
                    sobran_nominal = sorted(applier_skus_nominal_total - esperado_skus)

                    if faltan_nominal:
                        agregar_detalle(
                            detalles,
                            "ERR",
                            "APPLIER",
                            f"Faltan SKU nominales en applier: <span class='text-blue'>({', '.join(faltan_nominal)})</span>"
                        )

                    if sobran_nominal:
                        agregar_detalle(
                            detalles,
                            "ERR",
                            "APPLIER",
                            f"Sobran SKU nominales en applier: <span class='text-blue'>({', '.join(sobran_nominal)})</span>"
                        )

    else:
        agregar_detalle(detalles, "WARN", "TIPO", f"Tipo de descuento <span class='text-blue'>({tipo_desc_raw})</span> no tiene regla específica nueva. Se conserva validación general")

    return (ok, detalles, msje_popup_data) if retornar_msje_data else (ok, detalles)


# ============================================================
# GENERAR ARCHIVOS TXT
# ============================================================

def generar_txt(nombre_archivo, buffer):
    try:
        with open(nombre_archivo, "w", encoding="utf-8") as f:
            f.write(buffer.getvalue())
        print(Fore.GREEN + f"\nArchivo TXT generado: {nombre_archivo}\n")
    except Exception as e:
        print(Fore.RED + f"\nError al generar TXT: {e}\n")


# ============================================================
# SISTEMA DE SELECCIÓN DE RC (WEB O CONSOLA)
# ============================================================

def obtener_rc():
    if len(sys.argv) > 1:
        return sys.argv[1].upper().strip()
    print(Fore.CYAN + "========================================")
    print(Fore.CYAN + "Buscar promociones del usuario:", end=" ")
    return input().upper().strip()


# ============================================================
# FLUJO TRADICIONAL
# ============================================================

def ejecutar_flujo_tradicional(excel_files, rc_externo=None):
    df_eventos_total = pd.DataFrame()
    df_codigos_total = pd.DataFrame()
    archivos_tradicional = []
    eventos_ok = False

    for file in excel_files:
        df_ev = leer_hoja_eventos(file)
        if df_ev is not None:
            eventos_ok = True
            archivos_tradicional.append(os.path.splitext(os.path.basename(file))[0])
            df_eventos_total = pd.concat([df_eventos_total, df_ev], ignore_index=True)

        try:
            xls = pd.ExcelFile(file)
            for hoja in xls.sheet_names:
                if "CODIGO" in hoja.upper():
                    df_cod = pd.read_excel(file, sheet_name=hoja)
                    df_cod.columns = [normalizar_texto(c) for c in df_cod.columns]
                    df_codigos_total = pd.concat([df_codigos_total, df_cod], ignore_index=True)
        except Exception:
            pass

    if not eventos_ok:
        return None, None, None, None

    rc = rc_externo.upper().strip() if rc_externo is not None else obtener_rc()
    print(Fore.CYAN + f"\n      FLUJO TRADICIONAL (USUARIO: {rc})")
    print(Fore.CYAN + " Archivos analizados:")
    for a in archivos_tradicional:
        print(Fore.CYAN + f"  - {a}")
    print(Fore.CYAN + "========================================\n")

    df_eventos_total["ID GEO"] = df_eventos_total["ID GEO"].apply(normalizar_local)
    df_usuario = df_eventos_total[df_eventos_total["RC"].astype(str).str.upper() == rc]
    return df_usuario, df_codigos_total, rc, archivos_tradicional


def obtener_columna_id_geocom(df):
    if df is None or df.empty:
        return None

    candidatos = []
    for c in df.columns:
        c_norm = normalizar_texto(normalizar_encabezado(c))
        if "ID GEOCOM" in c_norm or c_norm == "IDGEOCOM" or ("GEOCOM" in c_norm and "ID" in c_norm):
            candidatos.append(c)

    if candidatos:
        return candidatos[0]

    return None


def preparar_df_completar_para_validacion(df_completar_total):
    if df_completar_total is None or df_completar_total.empty:
        return pd.DataFrame(), None, []

    df = limpiar_dataframe_columnas(df_completar_total).copy()
    nombre_col_id_geo = obtener_columna_id_geocom(df)

    if not nombre_col_id_geo:
        return df, None, []

    df[nombre_col_id_geo] = df[nombre_col_id_geo].apply(normalizar_local)

    mascara_valida = df[nombre_col_id_geo].apply(es_id_promocion_valido)
    ids_descartados = sorted({
        str(v).strip()
        for v in df.loc[~mascara_valida, nombre_col_id_geo].tolist()
        if str(v).strip() not in {"", "None", "nan", "NaN"}
    })

    df = df[mascara_valida].copy()
    df[nombre_col_id_geo] = df[nombre_col_id_geo].apply(normalizar_local)

    return df, nombre_col_id_geo, ids_descartados


# ============================================================
# MAIN
# ============================================================

def main():
    print(Fore.CYAN + "\n========= VALIDADOR PROMOCIONES =========\n")

    if not os.path.isdir(EXCEL_PATH):
        print(Fore.RED + f"No existe carpeta Excel: {EXCEL_PATH}")
        return
    if not os.path.isdir(EXPORT_PATH):
        print(Fore.RED + f"No existe carpeta Export: {EXPORT_PATH}")
        return

    excel_files = [
        os.path.join(EXCEL_PATH, f)
        for f in os.listdir(EXCEL_PATH)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    promos, listas_productos_export = cargar_promos_desde_exports(EXPORT_PATH)
    promos_por_id = {normalizar_local(p.get("id")): p for p in promos if es_id_promocion_valido(p.get("id"))}
    mapa_area_responsable = construir_mapa_area_responsable(excel_files)

    df_usuario, df_codigos_total, rc, archivos_tradicional = ejecutar_flujo_tradicional(excel_files)

    if df_usuario is not None:
        promos_usuario = [p for p in promos if p["creationUser"] == rc]

        buffer_tradicional = StringIO()
        original_stdout = sys.stdout
        sys.stdout = DualWriter(original_stdout, buffer_tradicional)

        print(Fore.BLUE + "========= RESULTADOS TRADICIONAL =========")
        print("ID GEO | ID EXPORT | Resultado")

        productos_excel = {}
        locales_excel = {}
        listas_local_excel = {}

        for id_geo, grupo in df_usuario.groupby("ID GEO"):
            listas_p = {normalizar_texto(v) for v in grupo["LISTA PRODUCTOS"] if not es_vacio(v)}
            if listas_p:
                productos_excel[id_geo] = sorted(list(listas_p))
            else:
                marcas = {normalizar_texto(v) for v in grupo["MARCA"] if not es_vacio(v)}
                skus = []
                for marca in marcas:
                    if not df_codigos_total.empty and "MARCA" in df_codigos_total.columns and "CÓDIGO PRODUCTO" in df_codigos_total.columns:
                        df_m = df_codigos_total[df_codigos_total["MARCA"].astype(str).str.upper() == marca]
                        skus.extend([normalizar_texto(x) for x in df_m["CÓDIGO PRODUCTO"].tolist()])
                productos_excel[id_geo] = sorted(set(skus))

            ll = {normalizar_texto(v) for v in grupo["LISTA LOCAL"] if not es_vacio(v)}
            listas_local_excel[id_geo] = next(iter(ll)) if ll else ""
            locales_excel[id_geo] = sorted({normalizar_local(v) for v in grupo["LOCAL"] if not es_vacio(v)})

            promo = next((p for p in promos_usuario if normalizar_local(p["id"]) == str(id_geo)), None)

            if not promo:
                print(Fore.RED + f"{id_geo} | No encontrada | Campo en Rojo No coinciden")
                print(Fore.YELLOW + "   ! [EXPORT] No existe en export")
                print("-" * 55)
                continue

            ok, detalles = validar_promocion_tradicional(id_geo, grupo, promo, productos_excel, locales_excel, listas_local_excel)

            print((Fore.GREEN if ok else Fore.RED) + f"{id_geo} | {promo['id']} | {'Coinciden' if ok else 'Campo en Rojo No coinciden'}")
            for tipo, msg in detalles:
                color = Fore.GREEN if tipo == "OK" else (Fore.YELLOW if tipo in {"WARN", "INFO"} else Fore.RED)
                pref = "+" if tipo == "OK" else ("!" if tipo in {"WARN", "INFO"} else "-")
                print(color + f"   {pref} {msg}")
            print("-" * 55)

        print(Fore.CYAN + "\n====== FIN FLUJO TRADICIONAL ======\n")
        sys.stdout = original_stdout

        resp = input("¿Desea generar archivo TXT del flujo TRADICIONAL? (S/N): ").strip().upper()
        if resp == "S":
            generar_txt("resultado_tradicional.txt", buffer_tradicional)

    archivos_completar = []
    df_completar_total = pd.DataFrame()
    for file in excel_files:
        df_c = leer_hoja_completar(file)
        if df_c is not None:
            df_completar_total = pd.concat([df_completar_total, df_c], ignore_index=True)
            archivos_completar.append(os.path.splitext(os.path.basename(file))[0])

    if not df_completar_total.empty:
        print(Fore.CYAN + "\n========================================")
        print(Fore.CYAN + "            FLUJO COMPLETAR")
        print(Fore.CYAN + " Archivos analizados:")
        for a in archivos_completar:
            print(Fore.CYAN + f"  - {a}")
        print(Fore.CYAN + "========================================\n")

        buffer_completar = StringIO()
        original_stdout = sys.stdout
        sys.stdout = DualWriter(original_stdout, buffer_completar)

        print(Fore.BLUE + "========= RESULTADOS COMPLETAR =========")
        print("ID GEO | ID EXPORT | Resultado")

        df_completar_validacion, nombre_col_id_geo, ids_descartados = preparar_df_completar_para_validacion(df_completar_total)

        if not nombre_col_id_geo:
            print(Fore.RED + "\n❌ ERROR: No existe columna ID GEOCOM en la hoja COMPLETAR.\n")
            sys.stdout = original_stdout
            return

        if ids_descartados:
            print(Fore.YELLOW + f"⚠ IDs/valores descartados en Excel COMPLETAR por no ser válidos: {', '.join(ids_descartados)}")

        ids_excel_detectados = sorted({normalizar_local(v) for v in df_completar_validacion[nombre_col_id_geo].tolist() if es_id_promocion_valido(v)})
        print(Fore.CYAN + f"IDs válidos detectados en Excel COMPLETAR: {', '.join(ids_excel_detectados) if ids_excel_detectados else '-'}")

        ids_export_detectados = sorted({normalizar_local(p['id']) for p in promos if es_id_promocion_valido(p.get('id'))})
        faltantes_en_export = [i for i in ids_excel_detectados if i not in ids_export_detectados]
        if faltantes_en_export:
            print(Fore.YELLOW + f"IDs presentes en Excel pero no encontrados en Export: {', '.join(faltantes_en_export)}")

        for id_geo, grupo in df_completar_validacion.groupby(nombre_col_id_geo, dropna=False):
            id_geo_norm = normalizar_local(id_geo)
            promo = next((p for p in promos if normalizar_local(p["id"]) == id_geo_norm), None)
            if not promo:
                print(Fore.RED + f"{id_geo_norm} | No encontrada | No coinciden")
                print(Fore.YELLOW + "   ! [EXPORT] No existe en export")
                print("-" * 55)
                continue

            ok, detalles = validar_promocion_completar(id_geo, grupo, promo, listas_productos_export, mapa_area_responsable, promos_por_id)
            etiqueta_tipo = " | Tipo: MSJE" if promo.get("applier_type") == "MESSAGE" else ""
            print((Fore.GREEN if ok else Fore.RED) + f"{normalizar_local(id_geo)} | {promo['id']} | {'Coinciden' if ok else 'No coinciden'}{etiqueta_tipo}")
            for tipo, msg in detalles:
                color = Fore.GREEN if tipo == "OK" else (Fore.YELLOW if tipo in {"WARN", "INFO"} else Fore.RED)
                pref = "+" if tipo == "OK" else ("!" if tipo in {"WARN", "INFO"} else "-")
                print(color + f"   {pref} {msg}")
            print("-" * 55)

        print(Fore.CYAN + "\n========= FIN FLUJO COMPLETAR =========\n")
        sys.stdout = original_stdout

        resp = input("¿Desea generar archivo TXT del flujo COMPLETAR? (S/N): ").strip().upper()
        if resp == "S":
            generar_txt("resultado_completar.txt", buffer_completar)
    else:
        print(Fore.YELLOW + "\nNo se encontró Excel tipo COMPLETAR.\n")

    print(Fore.CYAN + "\n========= PROCESO COMPLETADO =========\n")


if __name__ == "__main__":
    print("Ejecutando validador en modo consola...")
    main()
