import os
import re
import sys
from datetime import datetime
from io import StringIO
import xml.etree.ElementTree as ET

import pandas as pd
from flask import Flask, render_template, request, jsonify, redirect, send_file


# ============================================================
# RUTAS BASE Y CONFIGURACIÓN INICIAL
# ============================================================

BASE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
MODULOS_PATH = os.path.join(BASE_PATH, "modulos")
EXCEL_PATH = os.path.join(BASE_PATH, "Excel")
EXPORT_PATH = os.path.join(BASE_PATH, "Export")
LOG_PATH = os.path.join(BASE_PATH, "logs")

os.makedirs(LOG_PATH, exist_ok=True)
sys.path.append(MODULOS_PATH)

# Limpieza automática al iniciar el servidor
for carpeta in [EXCEL_PATH, EXPORT_PATH]:
    try:
        os.makedirs(carpeta, exist_ok=True)
        for archivo in os.listdir(carpeta):
            ruta = os.path.join(carpeta, archivo)
            if os.path.isfile(ruta):
                os.remove(ruta)
    except Exception:
        pass


# ============================================================
# IMPORTACIÓN DE MÓDULOS DEL PROYECTO
# ============================================================

from validador import (
    leer_hoja_eventos,
    leer_hoja_completar,
    ejecutar_flujo_tradicional,
    validar_promocion_tradicional,
    validar_promocion_completar,
    normalizar_local,
    parsear_promos,
    convertir_txt_a_xml_con_root
)

from parser_listas_export import parsear_listas_productos_export
from gestor import registrar_rutas_gestor
from repositorio import registrar_rutas_repositorio


# ============================================================
# CONFIGURACIÓN FLASK
# ============================================================

app = Flask(
    __name__,
    template_folder=os.path.join(os.path.dirname(__file__), "templates"),
    static_folder=os.path.join(os.path.dirname(__file__), "static")
)

app.secret_key = "ClaveUltraSecretaParaMensajesWeb"
registrar_rutas_gestor(app)
registrar_rutas_repositorio(app)


# ============================================================
# LOGGING
# ============================================================

def escribir_log(linea):
    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    archivo = os.path.join(LOG_PATH, f"log_{fecha}.txt")
    with open(archivo, "a", encoding="utf-8") as f:
        f.write(linea + "\n")


# ============================================================
# UTILIDADES GENERALES DE ARCHIVOS
# ============================================================

def limpiar_carpeta(path):
    errores = []
    for f in os.listdir(path):
        try:
            fp = os.path.join(path, f)
            if os.path.isfile(fp):
                os.remove(fp)
        except Exception as e:
            errores.append(str(e))
    return errores


def listar_archivos():
    return os.listdir(EXCEL_PATH), os.listdir(EXPORT_PATH)


# ============================================================
# UTILIDADES DE LIMPIEZA Y FORMATEO
# ============================================================

def _strip_html(texto):
    return re.sub(r"<[^>]+>", "", str(texto or "")).strip()


def _extraer_entre_parentesis(texto, etiqueta):
    """
    Busca patrones como 'Etiqueta: (valor)' dentro del texto.
    """
    patron = rf"{re.escape(etiqueta)}\s*:\s*\((.*?)\)"
    m = re.search(patron, texto)
    return m.group(1).strip() if m else ""

def _extraer_area_desde_detalles(detalles):
    """
    Toma el AreaResponsable detectada por el validador desde los detalles técnicos.
    La fuente válida es Excel/IMPUT, no el Export.
    """
    for d in detalles:
        if isinstance(d, tuple):
            _, msg = d
        else:
            msg = d.get("msg", "")
        msg_plain = _strip_html(msg)
        if msg_plain.startswith("[ÁREA]") and "AreaResponsable detectada" in msg_plain:
            area = _extraer_entre_parentesis(msg_plain, "AreaResponsable detectada")
            if area:
                return area
    return "-"


def _normalizar_lista_valores(valor):
    """
    Convierte 'a, b, c' -> 'a - b - c'
    """
    if not valor or valor == "-":
        return "-"
    partes = [p.strip() for p in valor.split(",") if p.strip()]
    return " - ".join(partes) if partes else "-"


def _formatear_monto_limpio(valor):
    """
    Convierte 11190.00 -> $11190
    Si no se puede, devuelve el valor original.
    """
    if not valor or valor == "-":
        return "-"
    try:
        num = float(str(valor).replace(",", "."))
        if num.is_integer():
            return f"${int(num)}"
        return f"${num:.2f}"
    except Exception:
        return str(valor)


def _formatear_numero_limpio(valor):
    if not valor or valor == "-":
        return "-"
    try:
        num = float(str(valor).replace(",", "."))
        return f"{num:.2f}"
    except Exception:
        return str(valor)


def _formatear_porcentaje_limpio(valor):
    if not valor or valor == "-":
        return "-"
    return valor


# ============================================================
# ANÁLISIS DE DETALLES PARA RESUMEN WEB
# ============================================================


def analizar_detalles(detalles):
    """
    A partir de los detalles técnicos genera:
    - estados resumidos
    - tipo_promocion
    - resumen_condicion
    - resumen_aplicador
    - mensaje principal amigable
    """
    mensajes = []
    for d in detalles:
        if isinstance(d, tuple):
            tipo, msg = d
        else:
            tipo, msg = d.get("tipo"), d.get("msg")
        mensajes.append({
            "tipo": tipo,
            "msg": msg,
            "msg_plain": _strip_html(msg)
        })

    resumen = {
        "estado_id": "No evaluado",
        "estado_facturar": "No evaluado",
        "estado_fechas": "No evaluado",
        "estado_condicion": "No evaluado",
        "estado_applier": "No evaluado",
        "fecha_inicio_ok": None,
        "fecha_fin_ok": None,
        "tipo_promocion": "-",
        "resumen_condicion": "-",
        "resumen_aplicador": "-",
        "mensaje_principal": "No coinciden",
        "aviso_principal": "",
    }

    # --------------------------------------------------------
    # Agrupar mensajes por tipo
    # --------------------------------------------------------
    id_items = [x for x in mensajes if x["msg_plain"].startswith("[ID]")]
    fact_items = [x for x in mensajes if x["msg_plain"].startswith("[FACTURAR]")]
    fechas_items = [x for x in mensajes if x["msg_plain"].startswith("[FECHAS]")]
    condicion_items = [x for x in mensajes if x["msg_plain"].startswith("[CONDICIÓN]")]
    applier_items = [x for x in mensajes if x["msg_plain"].startswith("[APPLIER]")]
    leyenda_items = [x for x in mensajes if x["msg_plain"].startswith("[LEYENDA]")]
    descuento_items = [x for x in mensajes if x["msg_plain"].startswith("[DESCUENTO]")]
    lista_items = [x for x in mensajes if x["msg_plain"].startswith("[LISTA PRODUCTOS]")]

    # --------------------------------------------------------
    # Estados base
    # --------------------------------------------------------
    if id_items:
        resumen["estado_id"] = "Coinciden" if all(x["tipo"] == "OK" for x in id_items) else "No coinciden"

    if fact_items:
        if any(x["tipo"] == "ERR" for x in fact_items):
            resumen["estado_facturar"] = "No coinciden"
        elif any(x["tipo"] == "WARN" for x in fact_items):
            resumen["estado_facturar"] = "Advertencia"
        else:
            resumen["estado_facturar"] = "Coinciden"

    # --------------------------------------------------------
    # Fechas
    # --------------------------------------------------------
    inicio_item = next(
        (x for x in fechas_items if x["msg_plain"].startswith("[FECHAS] Fecha Inicio Excel")),
        None
    )
    fin_item = next(
        (x for x in fechas_items if x["msg_plain"].startswith("[FECHAS] Fecha Fin Excel")),
        None
    )

    inicio_tipo = inicio_item["tipo"] if inicio_item else None
    fin_tipo = fin_item["tipo"] if fin_item else None

    fecha_inicio_excel = ""
    fecha_inicio_export = ""
    fecha_fin_excel = ""
    fecha_fin_export = ""

    if inicio_item:
        resumen["fecha_inicio_ok"] = (inicio_tipo == "OK")
        txt = inicio_item["msg_plain"]
        m = re.search(r"Fecha Inicio Excel \((.*?)\).*?Export \((.*?)\)", txt, re.IGNORECASE)
        if m:
            fecha_inicio_excel = m.group(1).strip()
            fecha_inicio_export = m.group(2).strip()

    if fin_item:
        resumen["fecha_fin_ok"] = (fin_tipo == "OK")
        txt = fin_item["msg_plain"]
        m = re.search(r"Fecha Fin Excel \((.*?)\).*?Export \((.*?)\)", txt, re.IGNORECASE)
        if m:
            fecha_fin_excel = m.group(1).strip()
            fecha_fin_export = m.group(2).strip()

    estado_fechas_base = "No evaluado"
    if inicio_tipo == "OK" and fin_tipo == "OK":
        estado_fechas_base = "OK"
    elif inicio_tipo in {"WARN", "ERR"} and fin_tipo == "OK":
        estado_fechas_base = "Posible Extensión"
    elif fin_tipo == "ERR":
        estado_fechas_base = "No coinciden"
    elif inicio_item or fin_item:
        estado_fechas_base = "No coinciden"

    detalle_fechas = []
    if fecha_inicio_excel:
        detalle_fechas.append(f"Inicio: {fecha_inicio_excel}")
    if fecha_fin_excel:
        detalle_fechas.append(f"Fin: {fecha_fin_excel}")

    if detalle_fechas and estado_fechas_base != "No evaluado":
        resumen["estado_fechas"] = f"{estado_fechas_base} | " + " | ".join(detalle_fechas)
    else:
        resumen["estado_fechas"] = estado_fechas_base

    # --------------------------------------------------------
    # Condición / Aplicador
    # --------------------------------------------------------
    if condicion_items:
        if any(x["tipo"] == "ERR" for x in condicion_items):
            resumen["estado_condicion"] = "No coinciden"
        elif any(x["tipo"] == "WARN" for x in condicion_items):
            resumen["estado_condicion"] = "Advertencia"
        else:
            resumen["estado_condicion"] = "Coinciden"

    applier_sin_sku_explicito = any(
        "no informa sku explícitos" in x["msg_plain"].lower()
        or "no informa sku explicitos" in x["msg_plain"].lower()
        for x in applier_items
    )

    if applier_items:
        if applier_sin_sku_explicito:
            resumen["estado_applier"] = "No coinciden"
        elif any(x["tipo"] == "ERR" for x in applier_items):
            resumen["estado_applier"] = "No coinciden"
        elif any(x["tipo"] == "WARN" for x in applier_items):
            resumen["estado_applier"] = "Advertencia"
        else:
            resumen["estado_applier"] = "Coinciden"

    # --------------------------------------------------------
    # Obtener leyendas
    # --------------------------------------------------------
    leyenda_excel = next((x for x in leyenda_items if "Excel → Tipo:" in x["msg_plain"]), None)
    leyenda_cond = next((x for x in leyenda_items if "Condición Export →" in x["msg_plain"]), None)
    leyenda_applier = next((x for x in leyenda_items if "Applier Export →" in x["msg_plain"]), None)

    # --------------------------------------------------------
    # Tipo promoción (flujo COMPLETAR)
    # --------------------------------------------------------
    if leyenda_excel:
        tipo = _extraer_entre_parentesis(leyenda_excel["msg_plain"], "Tipo")
        resumen["tipo_promocion"] = tipo if tipo else "-"

    tipo_prom = (resumen["tipo_promocion"] or "").upper()
    es_2da = "2DA" in tipo_prom
    es_pack_nominal = "PACK NOMINAL" in tipo_prom
    es_pack = bool(re.search(r"\bPACK\b", tipo_prom) or re.search(r"\d+\s*X\s*\d+", tipo_prom))
    es_porcentual = ("PORCENT" in tipo_prom or "%" in tipo_prom)
    es_nominal = ("NOMINAL" in tipo_prom and not es_pack_nominal)

    # --------------------------------------------------------
    # Resumen condición (flujo COMPLETAR)
    # --------------------------------------------------------
    if leyenda_cond:
        sku_val = _extraer_entre_parentesis(leyenda_cond["msg_plain"], "SKU")
        lista_val = _extraer_entre_parentesis(leyenda_cond["msg_plain"], "Lista")
        cantidad_cond_val = _extraer_entre_parentesis(leyenda_cond["msg_plain"], "Cantidad")

        sku_fmt = _normalizar_lista_valores(sku_val)
        lista_fmt = lista_val if lista_val and lista_val != "-" else "-"

        partes_cond = []

        if sku_fmt != "-":
            partes_cond.append(f"SKU: {sku_fmt}")
        elif lista_fmt != "-":
            partes_cond.append(f"SKU Lista = {lista_fmt}")

        if es_2da and cantidad_cond_val and cantidad_cond_val not in {"-", "0", "0.0", "0.00"}:
            try:
                q = float(cantidad_cond_val)
                q_txt = str(int(q)) if q.is_integer() else str(q)
            except Exception:
                q_txt = cantidad_cond_val
            partes_cond.append(f"Cada {q_txt} unidades")

        resumen["resumen_condicion"] = " | ".join(partes_cond) if partes_cond else "-"

    # --------------------------------------------------------
    # Resumen aplicador (flujo COMPLETAR)
    # --------------------------------------------------------
    if leyenda_applier:
        sku_val = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "SKU")
        lista_val = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "Lista")
        cantidad_val = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "Cantidad")

        porcentaje_val = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "%")
        monto_val = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "Monto")
        monto_export_val = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "Monto export")

        pct_nodo_export = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "% nodo export")
        pct_comercial_excel = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "% comercial Excel")
        dcto_bruto_q = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "Dcto bruto Excel(Q)")
        pvp_pack_excel = _extraer_entre_parentesis(leyenda_applier["msg_plain"], "PVPOfertaPack Excel")

        if not pvp_pack_excel and leyenda_excel:
            pvp_pack_excel = _extraer_entre_parentesis(leyenda_excel["msg_plain"], "PVPOfertaPack")

        unidades_excel = _extraer_entre_parentesis(leyenda_excel["msg_plain"], "Unidades") if leyenda_excel else ""

        sku_fmt = _normalizar_lista_valores(sku_val)
        lista_fmt = _normalizar_lista_valores(lista_val)
        partes = []

        if sku_fmt != "-":
            partes.append(f"SKU: {sku_fmt}")
        elif lista_fmt != "-":
            partes.append(f"Lista: {lista_fmt}")
        else:
            if resumen["resumen_condicion"].startswith("SKU Lista ="):
                partes.append(resumen["resumen_condicion"])

        if es_2da:
            if cantidad_val and cantidad_val not in {"-", "0", "0.0", "0.00"}:
                try:
                    q = float(cantidad_val)
                    q_txt = str(int(q)) if q.is_integer() else str(q)
                except Exception:
                    q_txt = cantidad_val
                partes.append(f"Cada {q_txt} unidades")

            pct_aplicador_fmt = _formatear_porcentaje_limpio(pct_nodo_export)
            pct_comercial_fmt = _formatear_porcentaje_limpio(pct_comercial_excel)
            dcto_bruto_fmt = _formatear_numero_limpio(dcto_bruto_q)

            if pct_aplicador_fmt != "-":
                partes.append(f"{pct_aplicador_fmt} aplicador")
            if pct_comercial_fmt != "-":
                partes.append(f"{pct_comercial_fmt} comercial")
            if dcto_bruto_fmt != "-":
                partes.append(f"Dcto bruto Q: {dcto_bruto_fmt}")

        elif es_pack:
            monto_fmt = _formatear_monto_limpio(monto_val or monto_export_val)
            pvp_fmt = _formatear_monto_limpio(pvp_pack_excel)

            if cantidad_val and cantidad_val not in {"-", "0", "0.0", "0.00"}:
                try:
                    q = float(cantidad_val)
                    q_txt = str(int(q)) if q.is_integer() else str(q)
                except Exception:
                    q_txt = cantidad_val
                partes.append(f"Cantidad: {q_txt}")

            if monto_fmt != "-":
                partes.append(f"Monto unitario: {monto_fmt}")

            if pvp_fmt != "-" and unidades_excel and unidades_excel != "-":
                try:
                    q_pack = float(unidades_excel)
                    q_pack_txt = str(int(q_pack)) if q_pack.is_integer() else str(q_pack)
                except Exception:
                    q_pack_txt = unidades_excel
                partes.append(f"Pack: {pvp_fmt} / {q_pack_txt}")

        elif es_nominal:
            pvp_fmt = _formatear_monto_limpio(pvp_pack_excel)
            monto_export_fmt = _formatear_monto_limpio(monto_export_val or monto_val)
            dcto_bruto_fmt = _formatear_numero_limpio(dcto_bruto_q)

            if pvp_fmt != "-":
                partes.append(f"PVPOfertaPack: {pvp_fmt}")
            if monto_export_fmt != "-":
                partes.append(f"Monto export: {monto_export_fmt}")
            if dcto_bruto_fmt != "-":
                partes.append(f"Dcto bruto Q: {dcto_bruto_fmt}")

        elif es_porcentual:
            pct_fmt = _formatear_porcentaje_limpio(porcentaje_val or pct_nodo_export)
            if pct_fmt != "-":
                partes.append(f"%: {pct_fmt}")

        else:
            if cantidad_val and cantidad_val not in {"-", "0", "0.0", "0.00"}:
                try:
                    q = float(cantidad_val)
                    q_txt = str(int(q)) if q.is_integer() else str(q)
                except Exception:
                    q_txt = cantidad_val
                partes.append(f"Cantidad: {q_txt}")

            monto_fmt = _formatear_monto_limpio(monto_export_val or monto_val)
            if monto_fmt != "-":
                partes.append(f"Monto: {monto_fmt}")

        resumen["resumen_aplicador"] = " | ".join(partes) if partes else "-"

    if applier_sin_sku_explicito:
        resumen["resumen_aplicador"] = "ERROR: applier sin SKU explícito"

    # --------------------------------------------------------
    # FALLBACK EVENTOS
    # --------------------------------------------------------
    porcentaje_eventos = ""
    porcentaje_export_eventos = ""

    if resumen["tipo_promocion"] == "-":
        for x in descuento_items:
            txt = x["msg_plain"]

            m_ambos = re.search(r"Excel\s*\((.*?)\).*?Export\s*\((.*?)\)", txt, re.IGNORECASE)
            if m_ambos:
                porcentaje_eventos = m_ambos.group(1).strip()
                porcentaje_export_eventos = m_ambos.group(2).strip()
                break

            m_uno = re.search(r"\((\d+(?:[\.,]\d+)?%)\)", txt)
            if m_uno:
                porcentaje_eventos = m_uno.group(1).strip()
                porcentaje_export_eventos = porcentaje_eventos
                break

        if porcentaje_eventos:
            resumen["tipo_promocion"] = f"PORCENTUAL - {porcentaje_eventos}"

    if resumen["resumen_condicion"] == "-":
        for x in condicion_items:
            txt = x["msg_plain"]
            m_lista = re.search(r"misma lista de productos del Excel:\s*\((.*?)\)", txt, re.IGNORECASE)
            if m_lista:
                resumen["resumen_condicion"] = f"Lista: {m_lista.group(1).strip()}"
                break

            m_cond = re.search(r"Excel\s*\((.*?)\)\s*vs\s*Condición\s*\((.*?)\)", txt, re.IGNORECASE)
            if m_cond:
                resumen["resumen_condicion"] = f"Lista Excel: {m_cond.group(1).strip()} | Condición: {m_cond.group(2).strip()}"
                break

    if resumen["resumen_condicion"] == "-":
        for x in lista_items:
            txt = x["msg_plain"]
            m = re.search(r"LISTA PRODUCTOS Excel\s*\((.*?)\)", txt, re.IGNORECASE)
            if m:
                resumen["resumen_condicion"] = f"Lista: {m.group(1).strip()}"
                break

    if resumen["resumen_aplicador"] == "-":
        lista_applier = ""
        for x in applier_items:
            txt = x["msg_plain"]
            m_lista = re.search(r"misma lista de productos del Excel:\s*\((.*?)\)", txt, re.IGNORECASE)
            if m_lista:
                lista_applier = m_lista.group(1).strip()
                break

            m_apl = re.search(r"Excel\s*\((.*?)\)\s*vs\s*Applier\s*\((.*?)\)", txt, re.IGNORECASE)
            if m_apl:
                lista_excel = m_apl.group(1).strip()
                lista_exp = m_apl.group(2).strip()
                partes_apl = [f"Lista Excel: {lista_excel}", f"Applier: {lista_exp}"]
                if porcentaje_export_eventos or porcentaje_eventos:
                    partes_apl.append(f"%: {porcentaje_export_eventos or porcentaje_eventos}")
                resumen["resumen_aplicador"] = " | ".join(partes_apl)
                break

        if resumen["resumen_aplicador"] == "-":
            if not lista_applier and resumen["resumen_condicion"].startswith("Lista: "):
                lista_applier = resumen["resumen_condicion"].replace("Lista: ", "", 1).strip()

            partes_apl = []
            if lista_applier:
                partes_apl.append(f"Lista: {lista_applier}")
            if porcentaje_export_eventos or porcentaje_eventos:
                partes_apl.append(f"%: {porcentaje_export_eventos or porcentaje_eventos}")
            resumen["resumen_aplicador"] = " | ".join(partes_apl) if partes_apl else "-"

    # fallback extra por seguridad para Tipo
    if resumen["tipo_promocion"] == "-" and (porcentaje_export_eventos or porcentaje_eventos):
        resumen["tipo_promocion"] = f"PORCENTUAL - {porcentaje_export_eventos or porcentaje_eventos}"

    # --------------------------------------------------------
    # Mensaje principal
    # --------------------------------------------------------
    hay_err_id = any(x["tipo"] == "ERR" for x in id_items)
    hay_err_fact = any(x["tipo"] == "ERR" for x in fact_items)
    hay_err_cond = any(x["tipo"] == "ERR" for x in condicion_items)
    hay_err_applier = applier_sin_sku_explicito or any(x["tipo"] == "ERR" for x in applier_items)

    solo_ext_fecha_inicio = (
        inicio_tipo == "WARN" and
        fin_tipo == "OK" and
        not hay_err_id and
        not hay_err_fact and
        not hay_err_cond and
        not hay_err_applier
    )

    if solo_ext_fecha_inicio:
        resumen["mensaje_principal"] = "Coinciden"
        resumen["aviso_principal"] = "Posible extensión: fecha inicio diferente"
    else:
        if (
            hay_err_id or
            hay_err_fact or
            hay_err_cond or
            hay_err_applier or
            fin_tipo == "ERR" or
            (inicio_tipo == "ERR" and fin_tipo != "OK")
        ):
            resumen["mensaje_principal"] = "No coinciden"
        else:
            if inicio_tipo == "OK" and fin_tipo == "OK":
                resumen["mensaje_principal"] = "Coinciden"
            else:
                resumen["mensaje_principal"] = "No coinciden"

    return resumen



# ============================================================
# RUTA PRINCIPAL
# ============================================================

@app.route("/")
@app.route("/validPromotion/")                            
def inicio():
    excel, export = listar_archivos()
    return render_template("index.html", excel_files=excel, export_files=export)


# ============================================================
# SUBIR ARCHIVOS
# ============================================================

@app.route("/upload", methods=["POST"])
def upload_files():
    cargados_excel = 0
    cargados_export = 0
    export_rechazados = []

    FIRMA_EXPORT = (
        "<?xml version='1.0' encoding='UTF-8'?>"
        "<uy.com.geocom.geopromotion.service.promotion.PromotionBlockList>"
        "<promotionTypeList>"
    )

    for file in request.files.getlist("excel_files"):
        if file.filename.lower().endswith(".xlsx"):
            file.save(os.path.join(EXCEL_PATH, file.filename))
            cargados_excel += 1

    for file in request.files.getlist("export_files"):
        if not file.filename.lower().endswith(".txt"):
            continue

        try:
            contenido = file.stream.read(4096)
            file.stream.seek(0)

            texto = None
            for enc in ("utf-8", "utf-16", "latin-1"):
                try:
                    texto = contenido.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue

            if texto is None or FIRMA_EXPORT not in texto.replace("\n", "").replace("\r", ""):
                export_rechazados.append(file.filename)
                continue

            file.save(os.path.join(EXPORT_PATH, file.filename))
            cargados_export += 1

        except Exception:
            export_rechazados.append(file.filename)

    return jsonify({
        "mensaje": "Carga finalizada",
        "excel": cargados_excel,
        "export": cargados_export,
        "excel_cargados": cargados_excel,
        "export_validos": cargados_export,
        "export_rechazados": export_rechazados,
        "lista_excel": os.listdir(EXCEL_PATH),
        "lista_export": os.listdir(EXPORT_PATH)
    })


# ============================================================
# BORRAR ARCHIVOS
# ============================================================

@app.route("/borrar", methods=["POST"])
def borrar_archivos():
    tipo = request.form.get("tipo")

    if tipo == "excel":
        errores = limpiar_carpeta(EXCEL_PATH)
        msg = "Se borraron TODOS los Excel."
    elif tipo == "export":
        errores = limpiar_carpeta(EXPORT_PATH)
        msg = "Se borraron TODOS los Export."
    else:
        return jsonify({"error": "Tipo inválido"})

    return jsonify({
        "mensaje": msg,
        "errores": errores,
        "lista_excel": os.listdir(EXCEL_PATH),
        "lista_export": os.listdir(EXPORT_PATH)
    })


# ============================================================
# PROCESAR VALIDACIÓN
# ============================================================

@app.route("/procesar", methods=["POST"])
def procesar():
    rc_web = request.form.get("rc", "").strip().upper()

    excel_files = [
        os.path.join(EXCEL_PATH, f)
        for f in os.listdir(EXCEL_PATH)
        if f.endswith(".xlsx")
    ]

    resultados_tradicional = []
    resultados_completar = []

    export_files = [
        os.path.join(EXPORT_PATH, f)
        for f in os.listdir(EXPORT_PATH)
        if f.endswith(".txt")
    ]

    promo_info_por_id = {}

    # --------------------------------------------------------
    # Info base desde export
    # --------------------------------------------------------
    for exp in export_files:
        export_name = os.path.basename(exp)
        tree, raw_text = convertir_txt_a_xml_con_root(exp)
        promos = parsear_promos(tree, export_name=export_name)

        for p in promos:
            pid = normalizar_local(str(p.get("id")).split(".")[0])
            promo_info_por_id[pid] = {
                "creationUser": p.get("creationUser", "-"),
                "enabled": p.get("enabled", False),
                "area_responsable": "-",
                "__tipo_competencia": p.get("__tipo_competencia", "-"),
                "__export_origen": p.get("__export_origen", "-"),
                "__tipo_descuento": "-"
            }

    # --------------------------------------------------------
    # Listas de productos export
    # --------------------------------------------------------
    listas_productos_export = {}
    for exp in export_files:
        listas_tmp = parsear_listas_productos_export(exp)
        for nombre, productos in listas_tmp.items():
            listas_productos_export.setdefault(nombre, set()).update(productos)

    # --------------------------------------------------------
    # Promos indexadas
    # --------------------------------------------------------
    promos_por_id = {}

    for exp in export_files:
        tree, raw_text = convertir_txt_a_xml_con_root(exp)
        promos = parsear_promos(tree, export_name=os.path.basename(exp))

        for promo_dict in promos:
            pid = normalizar_local(str(promo_dict["id"]).split(".")[0])

            promo_node = tree.getroot().find(f".//Promotion[id='{promo_dict['id']}']")

            if promo_node is None:
                for nodo in tree.getroot().findall(".//Promotion"):
                    id_node = nodo.find("id")
                    if id_node is not None and str(id_node.text).strip() == str(promo_dict["id"]).strip():
                        promo_node = nodo
                        break

            if promo_node is not None:
                promo_dict["__xml"] = ET.tostring(promo_node, encoding="unicode")
            else:
                promo_dict["__xml"] = None

            if pid not in promos_por_id:
                promos_por_id[pid] = promo_dict

    # ============================================================
    # FLUJO TRADICIONAL
    # ============================================================
    if rc_web:
        df_usuario, _, _, archivos_tradicional = ejecutar_flujo_tradicional(
            excel_files, rc_externo=rc_web
        )

        if df_usuario is not None and not df_usuario.empty:
            excel_origen_trad = ", ".join(
                sorted({os.path.basename(f) for f in archivos_tradicional})
            )

            for id_geo, grupo in df_usuario.groupby("ID GEO"):
                id_geo_norm = normalizar_local(str(id_geo).split(".")[0])
                promo = promos_por_id.get(id_geo_norm)
                info = promo_info_por_id.get(id_geo_norm, {}).copy()

                info["__tipo_descuento"] = "-"

                if "DESCUENTO" in grupo.columns:
                    val = grupo["DESCUENTO"].iloc[0]
                    texto_desc = "-"

                    if isinstance(val, (int, float)):
                        if val <= 1:
                            texto_desc = f"{int(val * 100)}%"
                        else:
                            texto_desc = f"{int(val)}%"
                    else:
                        texto_desc = str(val).strip()

                    info["__tipo_descuento"] = f"PORCENTUAL - {texto_desc}"

                if promo is None:
                    resultados_tradicional.append({
                        "id_geo": id_geo,
                        "mensaje": "No existe en export",
                        "aviso_principal": "",
                        "excel_origen": excel_origen_trad,
                        "export_origen": "-",
                        "promo_info": info,
                        "detalle": [{"tipo": "ERR", "msg": "No encontrada en export"}],
                        "estado_id": "No coinciden",
                        "estado_facturar": "No evaluado",
                        "estado_fechas": "No evaluado",
                        "estado_condicion": "No evaluado",
                        "estado_applier": "No evaluado",
                        "fecha_inicio_ok": None,
                        "fecha_fin_ok": None,
                        "tipo_promocion": "-",
                        "resumen_condicion": "-",
                        "resumen_aplicador": "-",
                    })
                    continue

                ok, detalles = validar_promocion_tradicional(
                    id_geo, grupo, promo, {}, {}, {}
                )

                analisis = analizar_detalles(detalles)
                info["area_responsable"] = _extraer_area_desde_detalles(detalles)

                resultados_tradicional.append({
                    "id_geo": id_geo,
                    "mensaje": analisis["mensaje_principal"],
                    "aviso_principal": analisis["aviso_principal"],
                    "excel_origen": excel_origen_trad,
                    "export_origen": info.get("__export_origen", "-"),
                    "promo_info": info,
                    "detalle": [{"tipo": d[0], "msg": d[1]} for d in detalles],
                    "estado_id": analisis["estado_id"],
                    "estado_facturar": analisis["estado_facturar"],
                    "estado_fechas": analisis["estado_fechas"],
                    "estado_condicion": analisis["estado_condicion"],
                    "estado_applier": analisis["estado_applier"],
                    "fecha_inicio_ok": analisis["fecha_inicio_ok"],
                    "fecha_fin_ok": analisis["fecha_fin_ok"],
                    "tipo_promocion": analisis["tipo_promocion"],
                    "resumen_condicion": analisis["resumen_condicion"],
                    "resumen_aplicador": analisis["resumen_aplicador"],
                })

    # ============================================================
    # PREPARAR DATAFRAME COMPLETAR
    # ============================================================
    df_completar_total = pd.DataFrame()

    for file in excel_files:
        df_c = leer_hoja_completar(file)
        if df_c is not None and not df_c.empty:
            df_c["__excel_origen"] = os.path.basename(file)
            df_completar_total = pd.concat([df_completar_total, df_c], ignore_index=True)

    # ============================================================
    # FLUJO COMPLETAR
    # ============================================================
    if not df_completar_total.empty:
        col_id_geo = [c for c in df_completar_total.columns if "GEOCOM" in c.upper()][0]

        for id_geo, grupo in df_completar_total.groupby(col_id_geo):
            id_geo_norm = normalizar_local(str(id_geo).split(".")[0])
            promo = promos_por_id.get(id_geo_norm)
            info = promo_info_por_id.get(id_geo_norm, {}).copy()

            if promo is None:
                resultados_completar.append({
                    "id_geo": id_geo,
                    "mensaje": "No existe en export",
                    "aviso_principal": "",
                    "excel_origen": grupo["__excel_origen"].iloc[0],
                    "export_origen": "-",
                    "promo_info": {},
                    "detalle": [{"tipo": "ERR", "msg": "No encontrada en export"}],
                    "estado_id": "No coinciden",
                    "estado_facturar": "No evaluado",
                    "estado_fechas": "No evaluado",
                    "estado_condicion": "No evaluado",
                    "estado_applier": "No evaluado",
                    "fecha_inicio_ok": None,
                    "fecha_fin_ok": None,
                    "tipo_promocion": "-",
                    "resumen_condicion": "-",
                    "resumen_aplicador": "-",
                })
                continue

            ok, detalles = validar_promocion_completar(
                id_geo,
                grupo,
                promo,
                listas_productos_export
            )

            analisis = analizar_detalles(detalles)

            info["area_responsable"] = _extraer_area_desde_detalles(detalles)
            info["__tipo_descuento"] = analisis["tipo_promocion"] or promo.get("__tipo_descuento", "-")

            resultados_completar.append({
                "id_geo": id_geo,
                "mensaje": analisis["mensaje_principal"],
                "aviso_principal": analisis["aviso_principal"],
                "excel_origen": grupo["__excel_origen"].iloc[0],
                "export_origen": info.get("__export_origen", "-"),
                "promo_info": info,
                "detalle": [{"tipo": d[0], "msg": d[1]} for d in detalles],
                "estado_id": analisis["estado_id"],
                "estado_facturar": analisis["estado_facturar"],
                "estado_fechas": analisis["estado_fechas"],
                "estado_condicion": analisis["estado_condicion"],
                "estado_applier": analisis["estado_applier"],
                "fecha_inicio_ok": analisis["fecha_inicio_ok"],
                "fecha_fin_ok": analisis["fecha_fin_ok"],
                "tipo_promocion": analisis["tipo_promocion"],
                "resumen_condicion": analisis["resumen_condicion"],
                "resumen_aplicador": analisis["resumen_aplicador"],
            })

    # ============================================================
    # DASHBOARD
    # ============================================================
    todos_los_resultados = resultados_tradicional + resultados_completar

    total_promos = len(todos_los_resultados)
    total_ok = sum(
        1 for r in todos_los_resultados
        if r.get("mensaje") == "Coinciden" and not r.get("aviso_principal")
    )
    total_warn = sum(
        1 for r in todos_los_resultados
        if r.get("mensaje") == "Coinciden" and r.get("aviso_principal")
    )
    total_err = sum(
        1 for r in todos_los_resultados
        if r.get("mensaje") != "Coinciden"
    )

    return render_template(
        "resultado.html",
        rc=rc_web,
        resultados_tradicional=resultados_tradicional,
        resultados_completar=resultados_completar,
        total_promos=total_promos,
        total_ok=total_ok,
        total_warn=total_warn,
        total_err=total_err
    )


# ============================================================
# DESCARGAR RESULTADOS
# ============================================================

@app.route("/descargar_resultados", methods=["POST"])
def descargar_resultados():
    fecha_txt = datetime.now().strftime("%d-%m-%Y")
    fecha_archivo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    rc = request.form.get("rc", "").strip()

    resultados_trad = request.form.get("tradicional_data", "")
    resultados_comp = request.form.get("completar_data", "")

    buffer = StringIO()
    buffer.write("========================================\n")
    buffer.write("VALIDACIÓN DE PROMOCIONES\n")
    buffer.write("========================================\n\n")
    buffer.write(f"Fecha validación: {fecha_txt}\n")
    buffer.write(f"Usuario creador: {rc if rc else '-'}\n\n")

    def limpiar_bloque_serializado(texto, titulo):
        if not texto or not texto.strip():
            return f"=== {titulo} ===\n\nSin promociones en esta sección.\n\n"

        salida = [f"=== {titulo} ===", ""]

        bloques = [b.strip() for b in texto.split("-----------------------------------------") if b.strip()]

        for bloque in bloques:
            lineas = [re.sub(r"<[^>]+>", "", x).strip() for x in bloque.splitlines()]
            lineas = [x for x in lineas if x]

            datos = {}
            detalles = []

            for ln in lineas:
                if ln.startswith("ID GEO:"):
                    datos["id_geo"] = ln.replace("ID GEO:", "").strip()
                elif ln.startswith("Excel:"):
                    datos["excel"] = ln.replace("Excel:", "").strip()
                elif ln.startswith("Export:"):
                    datos["export"] = ln.replace("Export:", "").strip()
                elif ln.startswith("Usuario creador:"):
                    datos["usuario"] = ln.replace("Usuario creador:", "").strip()
                elif ln.startswith("Área responsable:"):
                    datos["area"] = ln.replace("Área responsable:", "").strip()
                elif ln.startswith("Tipo descuento:"):
                    datos["tipo"] = ln.replace("Tipo descuento:", "").strip()
                elif ln.startswith("Resultado:"):
                    datos["resultado"] = ln.replace("Resultado:", "").strip()
                elif ln.startswith("Aviso:"):
                    datos["aviso"] = ln.replace("Aviso:", "").strip()
                elif ln.startswith("Estado ID:"):
                    datos["estado_id"] = ln.replace("Estado ID:", "").strip()
                elif ln.startswith("Estado Facturar:"):
                    datos["estado_facturar"] = ln.replace("Estado Facturar:", "").strip()
                elif ln.startswith("Estado Fechas:"):
                    datos["estado_fechas"] = ln.replace("Estado Fechas:", "").strip()
                elif ln.startswith("Estado Condición:"):
                    datos["estado_condicion"] = ln.replace("Estado Condición:", "").strip()
                elif ln.startswith("Estado Aplicador:"):
                    datos["estado_aplicador"] = ln.replace("Estado Aplicador:", "").strip()
                elif ln.startswith("Condición limpia:"):
                    datos["condicion"] = ln.replace("Condición limpia:", "").strip()
                elif ln.startswith("Aplicador limpio:"):
                    datos["aplicador"] = ln.replace("Aplicador limpio:", "").strip()
                elif ln.startswith("- "):
                    detalles.append(ln)

            estado_final = datos.get("resultado", "-")
            aviso = datos.get("aviso", "")
            if aviso and aviso != "-":
                estado_final = f"{estado_final} - {aviso}"

            salida.append("----------------------------------------")
            salida.append(f"PROMOCIÓN: {datos.get('id_geo', '-')}")
            salida.append("----------------------------------------")
            salida.append("")
            salida.append(f"Archivo Excel: {datos.get('excel', '-')}")
            salida.append(f"Archivo Export: {datos.get('export', '-')}")
            salida.append(f"Área responsable: {datos.get('area', '-')}")
            salida.append("")
            salida.append(f"Estado final: {estado_final}  |  Tipo promoción: {datos.get('tipo', '-')}")
            salida.append("")
            salida.append(f"Condición: {datos.get('condicion', '-')}")
            salida.append(f"Aplicador: {datos.get('aplicador', '-')}")
            salida.append("")
            salida.append("Validaciones:")

            id_geo = datos.get("id_geo", "-")
            estado_id = datos.get("estado_id", "-")
            estado_fact = datos.get("estado_facturar", "-")
            estado_fechas = datos.get("estado_fechas", "-")
            estado_cond = datos.get("estado_condicion", "-")
            estado_apl = datos.get("estado_aplicador", "-")

            fecha_inicio = None
            fecha_fin = None
            for det in detalles:
                m1 = re.search(r"Fecha Inicio Excel \((.*?)\)", det)
                if m1:
                    fecha_inicio = m1.group(1).strip()
                m2 = re.search(r"Fecha Fin Excel \((.*?)\)", det)
                if m2:
                    fecha_fin = m2.group(1).strip()

            salida.append(
                f"{'✓' if estado_id == 'Coinciden' else '✗'} ID promoción {'coincide' if estado_id == 'Coinciden' else 'no coincide'}: {id_geo}"
            )
            salida.append(
                f"{'✓' if estado_fact == 'Coinciden' else ('⚠' if estado_fact == 'Advertencia' else '✗')} "
                f"ID a facturar "
                f"{'coincide' if estado_fact == 'Coinciden' else ('con advertencia' if estado_fact == 'Advertencia' else 'no coincide')}: {id_geo}"
            )

            if estado_fechas == "OK":
                salida.append(f"✓ Fecha inicio coincide: {fecha_inicio or '-'}")
                salida.append(f"✓ Fecha fin coincide: {fecha_fin or '-'}")
            elif estado_fechas == "Posible Extensión":
                salida.append(f"⚠ Fecha inicio distinta / posible extensión: {fecha_inicio or '-'}")
                salida.append(f"✓ Fecha fin coincide: {fecha_fin or '-'}")
            else:
                salida.append(f"✗ Fecha inicio no coincide: {fecha_inicio or '-'}")
                salida.append(f"✗ Fecha fin no coincide: {fecha_fin or '-'}")

            salida.append(
                f"{'✓' if estado_cond == 'Coinciden' else ('⚠' if estado_cond == 'Advertencia' else '✗')} "
                f"Condición "
                f"{'correcta' if estado_cond == 'Coinciden' else ('con advertencia' if estado_cond == 'Advertencia' else 'incorrecta')}: "
                f"{datos.get('condicion', '-')}"
            )

            salida.append(
                f"{'✓' if estado_apl == 'Coinciden' else ('⚠' if estado_apl == 'Advertencia' else '✗')} "
                f"Aplicador "
                f"{'correcto' if estado_apl == 'Coinciden' else ('con advertencia' if estado_apl == 'Advertencia' else 'incorrecto')}: "
                f"{datos.get('aplicador', '-')}"
            )

            salida.append("")
        return "\n".join(salida) + "\n"

    buffer.write(limpiar_bloque_serializado(resultados_trad, "PROMOCIONES TRADICIONAL"))
    buffer.write("\n")
    buffer.write(limpiar_bloque_serializado(resultados_comp, "PROMOCIONES COMPLETAR"))

    filename = f"Resultados_Validador_{fecha_archivo}.txt"

    with open(filename, "w", encoding="utf-8") as f:
        f.write(buffer.getvalue())

    return send_file(
        filename,
        as_attachment=True,
        download_name=filename,
        mimetype="text/plain"
    )


# ============================================================
# RESET
# ============================================================

@app.route("/reset", methods=["GET"])
def reset():
    try:
        for var in ["df_usuario", "df_codigos_total", "df_completar_total"]:
            if var in globals():
                del globals()[var]

        import gc
        gc.collect()
        escribir_log("Memoria liberada correctamente.")

    except Exception as e:
        escribir_log(f"Error en reset: {e}")

    return redirect("/")


# ============================================================
# NUEVA VALIDACIÓN
# ============================================================

@app.route("/nueva_validacion")
def nueva_validacion():
    try:
        for carpeta in [EXCEL_PATH, EXPORT_PATH]:
            for archivo in os.listdir(carpeta):
                ruta = os.path.join(carpeta, archivo)
                if os.path.isfile(ruta):
                    os.remove(ruta)

        escribir_log("Nueva validación iniciada: carpetas Excel y Export limpiadas.")

    except Exception as e:
        escribir_log(f"Error en nueva_validacion: {e}")

    return redirect("/")


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    app.run(debug=True)
