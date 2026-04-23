"""
consultor_carga.py
Capa de servicio del Consultor de Cargas.

Objetivo:
- Consumir reglas_consultor.py
- Normalizar entradas del formulario
- Resolver la regla correcta
- Construir:
    * resumen
    * vista previa por pestañas
    * guía de carga
    * alertas
    * checklist final

Este archivo NO toca el validador.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any

from .reglas_consultor import (
    MODALIDAD_CLUB,
    MODALIDAD_MASIVA,
    MODALIDAD_EVENTOS,
    SUBMODO_CLON,
    SUBMODO_ORIGINAL,
    AREA_BYCP,
    AREA_BIENESTAR,
    AREA_FARMA,
    TIPO_DCTO_2DA_UNIDAD,
    TIPO_NOMINAL,
    TIPO_PACK_NOMINAL,
    TIPO_PACK_PRECIO_FIJO,
    TIPO_PORCENTUAL,
    TIPO_PACK_ESPECIAL_BYCP,
    resolver_regla,
)


# ============================================================
# NORMALIZACIÓN
# ============================================================

ALIASES_MODALIDAD = {
    "MASIVA": MODALIDAD_MASIVA,
    "MASSIVA": MODALIDAD_MASIVA,
    "NORMAL": MODALIDAD_MASIVA,
    "CLUB": MODALIDAD_CLUB,
    "EVENTOS": MODALIDAD_EVENTOS,
}

ALIASES_SUBMODO = {
    "ORIGINAL": SUBMODO_ORIGINAL,
    "ID ORIGINAL": SUBMODO_ORIGINAL,
    "CLON": SUBMODO_CLON,
    "ID CLON": SUBMODO_CLON,
}

ALIASES_AREA = {
    "FARMA": AREA_FARMA,
    "BIENESTAR": AREA_BIENESTAR,
    "BYCP": AREA_BYCP,
}

ALIASES_TIPO = {
    "NOMINAL": TIPO_NOMINAL,
    "PORCENTUAL": TIPO_PORCENTUAL,
    "PACK NOMINAL": TIPO_PACK_NOMINAL,
    "PACK_NOMINAL": TIPO_PACK_NOMINAL,
    "DCTO 2DA UNIDAD": TIPO_DCTO_2DA_UNIDAD,
    "DCTO_2DA_UNIDAD": TIPO_DCTO_2DA_UNIDAD,
    "DESCUENTO 2DA UNIDAD": TIPO_DCTO_2DA_UNIDAD,
    "PACK NOMINAL A PRECIO FIJO": TIPO_PACK_PRECIO_FIJO,
    "PACK PRECIO FIJO": TIPO_PACK_PRECIO_FIJO,
    "PACK_NOMINAL_PRECIO_FIJO": TIPO_PACK_PRECIO_FIJO,
    "PACK ESPECIAL BYCP": TIPO_PACK_ESPECIAL_BYCP,
    "PACK_ESPECIAL_BYCP": TIPO_PACK_ESPECIAL_BYCP,
    "BYCP 3X2 ESPECIAL": TIPO_PACK_ESPECIAL_BYCP,
    "BYCP_3X2_ESPECIAL": TIPO_PACK_ESPECIAL_BYCP,
    "3X2 BYCP": TIPO_PACK_ESPECIAL_BYCP,
    "BYCP 3X2": TIPO_PACK_ESPECIAL_BYCP,
}


def _normalizar_string(valor: Any) -> str:
    if valor is None:
        return ""
    return str(valor).strip().upper()


def normalizar_entrada(payload: dict[str, Any]) -> dict[str, Any]:
    """
    Normaliza un payload del formulario/UI a valores internos del consultor.

    Espera, como mínimo:
    - modalidad
    - area_funcional
    - tipo_descuento

    Opcional:
    - submodo_club
    - unidades_pack
    - pvp_oferta_pack
    - descuento_pack_nominal_bruto
    - descuento_porcentual
    - habilitada
    - fecha_inicio
    - fecha_fin
    - productos_o_lista
    - usa_mensaje
    """
    modalidad_raw = _normalizar_string(payload.get("modalidad"))
    area_raw = _normalizar_string(payload.get("area_funcional"))
    tipo_raw = _normalizar_string(payload.get("tipo_descuento"))
    submodo_raw = _normalizar_string(payload.get("submodo_club"))

    modalidad = ALIASES_MODALIDAD.get(modalidad_raw, modalidad_raw)
    area_funcional = ALIASES_AREA.get(area_raw, area_raw)
    tipo_descuento = ALIASES_TIPO.get(tipo_raw, tipo_raw)
    submodo_club = ALIASES_SUBMODO.get(submodo_raw, submodo_raw) if submodo_raw else None

    unidades_pack = payload.get("unidades_pack", 2)
    try:
        unidades_pack = int(unidades_pack) if unidades_pack not in (None, "") else 2
    except Exception:
        unidades_pack = 2

    habilitada = payload.get("habilitada", True)
    if isinstance(habilitada, str):
        habilitada = habilitada.strip().lower() in {"1", "true", "si", "sí", "yes", "on"}

    if tipo_descuento == TIPO_PACK_ESPECIAL_BYCP:
        unidades_pack = 3

    normalizado = {
        "modalidad": modalidad,
        "submodo_club": submodo_club,
        "area_funcional": area_funcional,
        "tipo_descuento": tipo_descuento,
        "unidades_pack": unidades_pack,
        "pvp_oferta_pack": payload.get("pvp_oferta_pack"),
        "descuento_pack_nominal_bruto": payload.get("descuento_pack_nominal_bruto"),
        "descuento_porcentual": payload.get("descuento_porcentual"),
        "habilitada": habilitada,
        "fecha_inicio": payload.get("fecha_inicio"),
        "fecha_fin": payload.get("fecha_fin"),
        "productos_o_lista": payload.get("productos_o_lista"),
        "usa_mensaje": payload.get("usa_mensaje"),
    }
    return normalizado


# ============================================================
# CÁLCULO / ARMADO
# ============================================================

def _safe_float(valor: Any) -> float | None:
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return None


def _formatear_numero(valor: float | None) -> str:
    if valor is None:
        return ""
    if float(valor).is_integer():
        return str(int(valor))
    return f"{valor:.2f}"


def _calcular_valor_aplicador(data: dict[str, Any], regla: dict[str, Any]) -> dict[str, Any]:
    """
    Calcula el valor orientativo que debería usarse en aplicadores.
    No modifica reglas de negocio; solo materializa el valor para la UI.
    """
    salida = {
        "valor_original": None,
        "valor_final": None,
        "detalle": "",
    }

    origen = regla["aplicadores"]["origen_valor"]
    divide_valor = regla["aplicadores"]["divide_valor"]
    unidades_pack = max(int(data.get("unidades_pack") or 2), 1)

    if origen == "PVPOfertaPack":
        valor_original = _safe_float(data.get("pvp_oferta_pack"))
    elif origen == "Descuento Pack Nominal Bruto":
        valor_original = _safe_float(data.get("descuento_pack_nominal_bruto"))
    elif origen == "Descuento Porcentual":
        valor_original = _safe_float(data.get("descuento_porcentual"))
        if data.get("tipo_descuento") == TIPO_PACK_ESPECIAL_BYCP and valor_original is None:
            valor_original = 100.0
    else:
        valor_original = None

    valor_final = valor_original
    detalle = "Sin cálculo adicional."

    if valor_original is not None and divide_valor:
        valor_final = valor_original / unidades_pack
        detalle = (
            f"Se divide el valor original ({_formatear_numero(valor_original)}) "
            f"entre #UnidadesPack ({unidades_pack})."
        )
    elif valor_original is not None and not divide_valor:
        detalle = (
            f"El valor se usa sin división: {_formatear_numero(valor_original)}."
        )

    salida["valor_original"] = valor_original
    salida["valor_final"] = valor_final
    salida["detalle"] = detalle
    return salida


def construir_basico(data: dict[str, Any], regla: dict[str, Any]) -> dict[str, Any]:
    basico = deepcopy(regla["basico"])
    basico["habilitada"] = data.get("habilitada", True)
    basico["fecha_inicio"] = data.get("fecha_inicio")
    basico["fecha_fin"] = data.get("fecha_fin")
    return basico


def construir_tiempo(data: dict[str, Any], regla: dict[str, Any]) -> dict[str, Any]:
    tiempo = deepcopy(regla["tiempo"])
    tiempo["fecha_inicio"] = data.get("fecha_inicio")
    tiempo["fecha_fin"] = data.get("fecha_fin")

    if data.get("modalidad") == MODALIDAD_EVENTOS:
        tiempo["usa_dias_especificos"] = True
        tiempo["observaciones"].append(
            "En EVENTOS se deben revisar fechas y días específicos."
        )

    return tiempo


def construir_condiciones(data: dict[str, Any], regla: dict[str, Any]) -> dict[str, Any]:
    condiciones = deepcopy(regla["condiciones"])

    if data.get("productos_o_lista"):
        condiciones["productos_o_lista"] = data["productos_o_lista"]

    if data.get("modalidad") == MODALIDAD_EVENTOS:
        condiciones["locales_regla"] = "Locales asignados"
        condiciones["observaciones"].append(
            "En EVENTOS usar locales asignados en lugar de EXC_LOCALES."
        )

    return condiciones


def construir_aplicadores(data: dict[str, Any], regla: dict[str, Any]) -> dict[str, Any]:
    aplicadores = deepcopy(regla["aplicadores"])
    calculo = _calcular_valor_aplicador(data, regla)

    aplicadores["valor_original"] = calculo["valor_original"]
    aplicadores["valor_final"] = calculo["valor_final"]
    aplicadores["detalle_calculo"] = calculo["detalle"]

    return aplicadores


def construir_camino(data: dict[str, Any], regla: dict[str, Any]) -> list[str]:
    pasos = list(regla["camino"])

    if data.get("modalidad") == MODALIDAD_CLUB:
        if data.get("submodo_club") == SUBMODO_ORIGINAL:
            pasos.insert(
                0,
                "En CLUB ORIGINAL, usar el ID original en ID a Facturar e ID Geocom."
            )
        elif data.get("submodo_club") == SUBMODO_CLON:
            pasos.insert(
                0,
                "En CLUB CLON, usar el ID clon en ID Lista Cliente y colocar el ID original en ID Alternativo."
            )

    if data.get("modalidad") == MODALIDAD_EVENTOS:
        pasos.append("Como es EVENTOS, revisar también días y locales asignados.")

    pasos.append("Antes de guardar, verificar que condición y aplicador usen la misma lista o productos.")
    return pasos


def construir_alertas(data: dict[str, Any], regla: dict[str, Any]) -> list[str]:
    alertas = list(regla["alertas"])

    tipo_descuento = data.get("tipo_descuento")
    area_funcional = data.get("area_funcional")

    if tipo_descuento == TIPO_PACK_NOMINAL and area_funcional == AREA_BYCP:
        alertas.append("BYCP + Pack nominal: no dividir descuento y usar aplicador cantidad 1.")

    if tipo_descuento == TIPO_DCTO_2DA_UNIDAD and area_funcional == AREA_BYCP:
        alertas.append("BYCP + 2da unidad: no dividir porcentaje y usar aplicador cantidad 1.")

    if tipo_descuento in {TIPO_PACK_NOMINAL, TIPO_DCTO_2DA_UNIDAD} and area_funcional in {AREA_FARMA, AREA_BIENESTAR}:
        alertas.append("Farma/Bienestar: dividir valor o porcentaje entre las unidades.")

    if data.get("modalidad") == MODALIDAD_CLUB:
        alertas.append("En CLUB, el área GEO final no es el área funcional: siempre va FIDELIZACION.")

    if data.get("modalidad") == MODALIDAD_EVENTOS:
        alertas.append("En EVENTOS usar locales asignados y revisar fechas/días específicos.")

    if tipo_descuento == TIPO_PACK_ESPECIAL_BYCP:
        alertas.append("PACK ESPECIAL BYCP: no tratar como pack a precio fijo.")
        alertas.append("Usar condición cantidad 3, aplicador porcentaje 100 a 1 unidad y strategy menor.")
        alertas.append("Competencia por producto porque el beneficio cae sobre el producto de menor valor.")

    return alertas


def construir_checklist(data: dict[str, Any], regla: dict[str, Any], aplicadores: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "titulo": "Competencia",
            "valor": regla["basico"]["competencia"],
            "detalle": "Confirmar que en GEO coincida exactamente con la lógica esperada.",
        },
        {
            "titulo": "Aplicador",
            "valor": regla["aplicadores"]["tipo"],
            "detalle": "Verificar que el tipo de aplicador en GEO sea el correcto.",
        },
        {
            "titulo": "Cantidad en condiciones",
            "valor": regla["condiciones"]["cantidad_cada"],
            "detalle": "Revisar el 'cantidad cada' esperado en condiciones.",
        },
        {
            "titulo": "Cantidad en aplicadores",
            "valor": regla["aplicadores"]["cantidad"],
            "detalle": "Revisar cantidad del aplicador.",
        },
        {
            "titulo": "Por unidad",
            "valor": regla["aplicadores"]["por_unidad"],
            "detalle": "Confirmar si debe marcarse o no.",
        },
        {
            "titulo": "Valor a cargar",
            "valor": _formatear_numero(aplicadores.get("valor_final")),
            "detalle": aplicadores.get("detalle_calculo", ""),
        },
        {
            "titulo": "Consistencia condición/aplicador",
            "valor": "Obligatoria",
            "detalle": "La condición y el aplicador deben usar los mismos productos o la misma lista.",
        },
    ]


# ============================================================
# SALIDA COMPLETA
# ============================================================

def construir_consulta(payload: dict[str, Any]) -> dict[str, Any]:
    """
    Función principal del servicio.
    Recibe payload del formulario y devuelve una salida lista
    para renderizar en UI.
    """
    data = normalizar_entrada(payload)

    regla = resolver_regla(
        modalidad=data["modalidad"],
        area_funcional=data["area_funcional"],
        tipo_descuento=data["tipo_descuento"],
        submodo_club=data.get("submodo_club"),
    )

    basico = construir_basico(data, regla)
    tiempo = construir_tiempo(data, regla)
    condiciones = construir_condiciones(data, regla)
    aplicadores = construir_aplicadores(data, regla)
    camino = construir_camino(data, regla)
    alertas = construir_alertas(data, regla)
    checklist = construir_checklist(data, regla, aplicadores)

    salida = {
        "entrada_normalizada": data,
        "resumen": deepcopy(regla["resumen"]),
        "preview": {
            "basico": basico,
            "tiempo": tiempo,
            "condiciones": condiciones,
            "aplicadores": aplicadores,
            "camino": camino,
        },
        "alertas": alertas,
        "checklist": checklist,
        "regla_original": regla,
    }

    # Completar resumen con datos calculados visibles
    salida["resumen"]["valor_referencia"] = aplicadores.get("valor_final")
    salida["resumen"]["detalle_calculo"] = aplicadores.get("detalle_calculo")
    salida["resumen"]["area_geo_final"] = basico.get("area_geo_final")
    salida["resumen"]["grupo_geo"] = basico.get("grupo_geo")

    return salida


# ============================================================
# FORMATEO OPCIONAL PARA TEXTO / DEBUG
# ============================================================

def construir_guia_textual(payload: dict[str, Any]) -> str:
    """
    Devuelve una guía textual simple para mostrar o depurar.
    """
    consulta = construir_consulta(payload)
    preview = consulta["preview"]
    resumen = consulta["resumen"]

    lineas: list[str] = []
    lineas.append("GUÍA DE CARGA")
    lineas.append("=" * 60)
    lineas.append(f"Tipo de descuento: {consulta['entrada_normalizada']['tipo_descuento']}")
    lineas.append(f"Área GEO final: {resumen['area_geo_final']}")
    lineas.append(f"Competencia: {resumen['competencia']}")
    lineas.append(f"Aplicador GEO: {resumen['aplicador_geo']}")
    if consulta['entrada_normalizada']['tipo_descuento'] == TIPO_PACK_ESPECIAL_BYCP:
        lineas.append("Caso especial: PACK ESPECIAL BYCP se interpreta como 3x2 fijo.")
    lineas.append("")

    lineas.append("BÁSICO")
    lineas.append("-" * 60)
    lineas.append(f"Modalidad: {preview['basico']['modalidad']}")
    lineas.append(f"Área funcional: {preview['basico']['area_funcional']}")
    lineas.append(f"Área GEO final: {preview['basico']['area_geo_final']}")
    lineas.append(f"Grupo GEO: {preview['basico']['grupo_geo']}")
    lineas.append(f"Nombre general: {preview['basico']['nombre_general']}")
    lineas.append(f"Habilitada: {preview['basico']['habilitada']}")
    lineas.append("")

    lineas.append("TIEMPO")
    lineas.append("-" * 60)
    lineas.append(f"Fecha inicio: {preview['tiempo'].get('fecha_inicio')}")
    lineas.append(f"Fecha fin: {preview['tiempo'].get('fecha_fin')}")
    lineas.append("")

    lineas.append("CONDICIONES")
    lineas.append("-" * 60)
    lineas.append(f"Tipo: {preview['condiciones']['tipo']}")
    lineas.append(f"Cantidad cada: {preview['condiciones']['cantidad_cada']}")
    lineas.append(f"Productos/lista: {preview['condiciones']['productos_o_lista']}")
    lineas.append(f"Locales: {preview['condiciones']['locales_regla']}")
    lineas.append("")

    lineas.append("APLICADORES")
    lineas.append("-" * 60)
    lineas.append(f"Tipo: {preview['aplicadores']['tipo']}")
    lineas.append(f"Origen del valor: {preview['aplicadores']['origen_valor']}")
    lineas.append(f"Valor final orientativo: {_formatear_numero(preview['aplicadores'].get('valor_final'))}")
    lineas.append(f"Cantidad: {preview['aplicadores']['cantidad']}")
    lineas.append(f"Estrategia: {preview['aplicadores']['estrategia']}")
    lineas.append(f"Por unidad: {preview['aplicadores']['por_unidad']}")
    lineas.append("")

    lineas.append("CAMINO")
    lineas.append("-" * 60)
    for paso in preview["camino"]:
        lineas.append(f"- {paso}")
    lineas.append("")

    lineas.append("ALERTAS")
    lineas.append("-" * 60)
    for alerta in consulta["alertas"]:
        lineas.append(f"- {alerta}")

    return "\n".join(lineas)


# ============================================================
# EJEMPLO
# ============================================================

if __name__ == "__main__":
    ejemplo = {
        "modalidad": "club",
        "submodo_club": "clon",
        "area_funcional": "BYCP",
        "tipo_descuento": "Pack nominal",
        "unidades_pack": 2,
        "descuento_pack_nominal_bruto": 21990,
        "habilitada": True,
        "fecha_inicio": "2026-04-22",
        "fecha_fin": "2026-04-30",
        "productos_o_lista": "Lista: CLUB_EJEMPLO",
    }

    print(construir_guia_textual(ejemplo))
