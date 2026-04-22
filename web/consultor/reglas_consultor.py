"""
reglas_consultor.py
Base de reglas del Consultor de Cargas.

Objetivo:
- Centralizar la lógica funcional del consultor.
- No tocar la lógica del validador.
- Dejar una estructura robusta y legible para que luego
  consultor_carga.py la consuma y construya:
    * guía de carga
    * vista previa del sistema
    * alertas
    * camino de configuración

Notas:
- Este archivo refleja la matriz funcional cerrada con el usuario.
- "Porcentual BYCP" compite por producto.
- CLUB usa siempre area GEO "FIDELIZACION" y grupo "CLUB_NECESIDADES".
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any


# ============================================================
# CATÁLOGOS
# ============================================================

MODALIDAD_MASIVA = "MASIVA"
MODALIDAD_CLUB = "CLUB"

SUBMODO_ORIGINAL = "ORIGINAL"
SUBMODO_CLON = "CLON"

AREA_FARMA = "FARMA"
AREA_BIENESTAR = "BIENESTAR"
AREA_BYCP = "BYCP"

AREA_GEO_FIDELIZACION = "FIDELIZACION"
GRUPO_CLUB_NECESIDADES = "CLUB_NECESIDADES"

TIPO_NOMINAL = "NOMINAL"
TIPO_PORCENTUAL = "PORCENTUAL"
TIPO_PACK_NOMINAL = "PACK_NOMINAL"
TIPO_DCTO_2DA_UNIDAD = "DCTO_2DA_UNIDAD"
TIPO_PACK_PRECIO_FIJO = "PACK_NOMINAL_PRECIO_FIJO"

TIPOS_DESCUENTO_SOPORTADOS = (
    TIPO_NOMINAL,
    TIPO_PORCENTUAL,
    TIPO_PACK_NOMINAL,
    TIPO_DCTO_2DA_UNIDAD,
    TIPO_PACK_PRECIO_FIJO,
)

AREAS_FUNCIONALES = (
    AREA_FARMA,
    AREA_BIENESTAR,
    AREA_BYCP,
)


# ============================================================
# TEXTOS NORMALIZADOS
# ============================================================

COMP_POR_PRODUCTO = "Comp. por Producto"
COMP_POR_PROMOCION = "Comp. por Promoción"

APLICADOR_PRECIO_FIJO = "Precio fijo a producto"
APLICADOR_PORCENTAJE = "Porcentaje a producto"
APLICADOR_MONTO = "Monto a producto"

ESTRATEGIA_MENOR = "menor"

NOMBRE_GENERAL_NORMAL = "DESCUENTO FCV"
NOMBRE_GENERAL_CLUB = "DCTO EXCLUSIVO CLUB"

REQUIERE_MENSAJE_TIPOS = {
    TIPO_PACK_NOMINAL,
    TIPO_DCTO_2DA_UNIDAD,
    TIPO_PACK_PRECIO_FIJO,
}


# ============================================================
# REGLAS BASE
# ============================================================

def _base_regla() -> dict[str, Any]:
    return {
        "resumen": {
            "tipo_descuento": "",
            "competencia": "",
            "aplicador_geo": "",
            "requiere_mensaje": False,
            "divide_valor": False,
            "origen_valor": "",
        },
        "basico": {
            "modalidad": "",
            "submodo_club": None,
            "area_funcional": "",
            "area_geo_final": "",
            "grupo_geo": None,
            "competencia": "",
            "habilitada": True,
            "nombre_general": "",
            "observaciones": [],
        },
        "tiempo": {
            "usa_fechas": True,
            "usa_dias_especificos": False,
            "observaciones": [],
        },
        "condiciones": {
            "tipo": "Boleta",
            "cantidad_cada": 1,
            "productos_o_lista": "Mismos productos/lista que aplicadores",
            "locales_regla": "EXC_LOCALES",
            "incluye_convenios": False,
            "observaciones": [],
        },
        "aplicadores": {
            "tipo": "",
            "origen_valor": "",
            "divide_valor": False,
            "formula_division": None,
            "cantidad": 1,
            "estrategia": None,
            "por_unidad": None,
            "observaciones": [],
        },
        "camino": [],
        "alertas": [],
    }


def _aplicar_club(regla: dict[str, Any], submodo_club: str) -> dict[str, Any]:
    r = deepcopy(regla)

    r["basico"]["modalidad"] = MODALIDAD_CLUB
    r["basico"]["submodo_club"] = submodo_club
    r["basico"]["area_geo_final"] = AREA_GEO_FIDELIZACION
    r["basico"]["grupo_geo"] = GRUPO_CLUB_NECESIDADES
    r["basico"]["nombre_general"] = NOMBRE_GENERAL_CLUB
    r["basico"]["observaciones"].append(
        "En CLUB el área GEO final va siempre como FIDELIZACION."
    )
    r["basico"]["observaciones"].append(
        "En CLUB se debe seleccionar siempre el grupo CLUB_NECESIDADES."
    )
    r["alertas"].append("CLUB usa FIDELIZACION + CLUB_NECESIDADES.")

    if submodo_club == SUBMODO_ORIGINAL:
        r["alertas"].append(
            "ID ORIGINAL: usar en ID a Facturar e ID Geocom. No lleva ID Alternativo."
        )
    elif submodo_club == SUBMODO_CLON:
        r["condiciones"]["incluye_convenios"] = True
        r["condiciones"]["observaciones"].append(
            "El CLON agrega convenios en condiciones."
        )
        r["alertas"].append(
            "ID CLON: usar en ID Lista Cliente y colocar el ID original en ID Alternativo."
        )

    return r


def _aplicar_masiva(regla: dict[str, Any], area_funcional: str) -> dict[str, Any]:
    r = deepcopy(regla)
    r["basico"]["modalidad"] = MODALIDAD_MASIVA
    r["basico"]["submodo_club"] = None
    r["basico"]["area_geo_final"] = area_funcional
    r["basico"]["grupo_geo"] = None
    r["basico"]["nombre_general"] = NOMBRE_GENERAL_NORMAL
    return r


# ============================================================
# REGLAS POR TIPO / ÁREA
# ============================================================

REGLAS_POR_TIPO_Y_AREA: dict[str, dict[str, dict[str, Any]]] = {
    TIPO_NOMINAL: {
        AREA_FARMA: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PRECIO_FIJO,
            "origen_valor": "PVPOfertaPack",
            "divide_valor": False,
            "cantidad_condicion": 1,
            "cantidad_aplicador": 1,
            "estrategia": None,
            "por_unidad": True,
            "notas": [
                "Nominal se carga como Precio fijo a producto.",
                "Tomar monto desde PVPOfertaPack.",
                "Marcar Por unidad.",
            ],
        },
        AREA_BIENESTAR: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PRECIO_FIJO,
            "origen_valor": "PVPOfertaPack",
            "divide_valor": False,
            "cantidad_condicion": 1,
            "cantidad_aplicador": 1,
            "estrategia": None,
            "por_unidad": True,
            "notas": [
                "Nominal se carga como Precio fijo a producto.",
                "Tomar monto desde PVPOfertaPack.",
                "Marcar Por unidad.",
            ],
        },
        AREA_BYCP: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PRECIO_FIJO,
            "origen_valor": "PVPOfertaPack",
            "divide_valor": False,
            "cantidad_condicion": 1,
            "cantidad_aplicador": 1,
            "estrategia": None,
            "por_unidad": True,
            "notas": [
                "Nominal se carga como Precio fijo a producto.",
                "Tomar monto desde PVPOfertaPack.",
                "Marcar Por unidad.",
            ],
        },
    },
    TIPO_PORCENTUAL: {
        AREA_FARMA: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PORCENTAJE,
            "origen_valor": "Descuento Porcentual",
            "divide_valor": False,
            "cantidad_condicion": 1,
            "cantidad_aplicador": 1,
            "estrategia": None,
            "por_unidad": None,
            "notas": [
                "Porcentual va en Porcentaje a producto.",
                "Cargar el porcentaje tal como viene.",
            ],
        },
        AREA_BIENESTAR: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PORCENTAJE,
            "origen_valor": "Descuento Porcentual",
            "divide_valor": False,
            "cantidad_condicion": 1,
            "cantidad_aplicador": 1,
            "estrategia": None,
            "por_unidad": None,
            "notas": [
                "Porcentual va en Porcentaje a producto.",
                "Cargar el porcentaje tal como viene.",
            ],
        },
        AREA_BYCP: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PORCENTAJE,
            "origen_valor": "Descuento Porcentual",
            "divide_valor": False,
            "cantidad_condicion": 1,
            "cantidad_aplicador": 1,
            "estrategia": None,
            "por_unidad": None,
            "notas": [
                "Porcentual BYCP compite por producto.",
                "Porcentual va en Porcentaje a producto.",
                "Cargar el porcentaje tal como viene.",
            ],
        },
    },
    TIPO_PACK_NOMINAL: {
        AREA_FARMA: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_MONTO,
            "origen_valor": "Descuento Pack Nominal Bruto",
            "divide_valor": True,
            "formula_division": "Descuento Pack Nominal Bruto / #UnidadesPack",
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": True,
            "notas": [
                "Dividir el descuento entre las unidades.",
                "Usar Monto a producto.",
                "Aplicador cantidad 2 y estrategia menor.",
            ],
        },
        AREA_BIENESTAR: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_MONTO,
            "origen_valor": "Descuento Pack Nominal Bruto",
            "divide_valor": True,
            "formula_division": "Descuento Pack Nominal Bruto / #UnidadesPack",
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": True,
            "notas": [
                "Dividir el descuento entre las unidades.",
                "Usar Monto a producto.",
                "Aplicador cantidad 2 y estrategia menor.",
            ],
        },
        AREA_BYCP: {
            "competencia": COMP_POR_PROMOCION,
            "aplicador": APLICADOR_MONTO,
            "origen_valor": "Descuento Pack Nominal Bruto",
            "divide_valor": False,
            "formula_division": None,
            "cantidad_condicion": 2,
            "cantidad_aplicador": 1,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": True,
            "notas": [
                "BYCP no divide el descuento.",
                "Usar Monto a producto.",
                "Aplicador cantidad 1 y estrategia menor.",
            ],
        },
    },
    TIPO_DCTO_2DA_UNIDAD: {
        AREA_FARMA: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PORCENTAJE,
            "origen_valor": "Descuento Porcentual",
            "divide_valor": True,
            "formula_division": "Descuento Porcentual / #UnidadesPack",
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": None,
            "notas": [
                "Dividir el porcentaje entre las unidades.",
                "Usar Porcentaje a producto.",
                "Aplicador cantidad 2 y estrategia menor.",
            ],
        },
        AREA_BIENESTAR: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PORCENTAJE,
            "origen_valor": "Descuento Porcentual",
            "divide_valor": True,
            "formula_division": "Descuento Porcentual / #UnidadesPack",
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": None,
            "notas": [
                "Dividir el porcentaje entre las unidades.",
                "Usar Porcentaje a producto.",
                "Aplicador cantidad 2 y estrategia menor.",
            ],
        },
        AREA_BYCP: {
            "competencia": COMP_POR_PROMOCION,
            "aplicador": APLICADOR_PORCENTAJE,
            "origen_valor": "Descuento Porcentual",
            "divide_valor": False,
            "formula_division": None,
            "cantidad_condicion": 2,
            "cantidad_aplicador": 1,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": None,
            "notas": [
                "BYCP no divide el porcentaje.",
                "Usar Porcentaje a producto.",
                "Aplicador cantidad 1 y estrategia menor.",
            ],
        },
    },
    TIPO_PACK_PRECIO_FIJO: {
        AREA_FARMA: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PRECIO_FIJO,
            "origen_valor": "PVPOfertaPack",
            "divide_valor": False,
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": False,
            "notas": [
                "Usar el precio final del pack.",
                "No marcar Por unidad.",
                "Aplicador cantidad 2 y estrategia menor.",
            ],
        },
        AREA_BIENESTAR: {
            "competencia": COMP_POR_PRODUCTO,
            "aplicador": APLICADOR_PRECIO_FIJO,
            "origen_valor": "PVPOfertaPack",
            "divide_valor": False,
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": False,
            "notas": [
                "Usar el precio final del pack.",
                "No marcar Por unidad.",
                "Aplicador cantidad 2 y estrategia menor.",
            ],
        },
        AREA_BYCP: {
            "competencia": COMP_POR_PROMOCION,
            "aplicador": APLICADOR_PRECIO_FIJO,
            "origen_valor": "PVPOfertaPack",
            "divide_valor": False,
            "cantidad_condicion": 2,
            "cantidad_aplicador": 2,
            "estrategia": ESTRATEGIA_MENOR,
            "por_unidad": False,
            "notas": [
                "Usar el precio final del pack.",
                "No marcar Por unidad.",
                "BYCP compite por promoción en pack a precio fijo.",
            ],
        },
    },
}


# ============================================================
# FUNCIÓN PRINCIPAL DE RESOLUCIÓN
# ============================================================

def resolver_regla(
    modalidad: str,
    area_funcional: str,
    tipo_descuento: str,
    submodo_club: str | None = None,
) -> dict[str, Any]:
    """
    Devuelve una estructura completa de reglas para el consultor.

    Parámetros
    ----------
    modalidad:
        "MASIVA" o "CLUB"
    area_funcional:
        "FARMA", "BIENESTAR", "BYCP"
    tipo_descuento:
        Uno de los tipos soportados
    submodo_club:
        "ORIGINAL" o "CLON" cuando la modalidad sea CLUB
    """
    if modalidad not in (MODALIDAD_MASIVA, MODALIDAD_CLUB):
        raise ValueError(f"Modalidad no soportada: {modalidad}")

    if area_funcional not in AREAS_FUNCIONALES:
        raise ValueError(f"Área funcional no soportada: {area_funcional}")

    if tipo_descuento not in TIPOS_DESCUENTO_SOPORTADOS:
        raise ValueError(f"Tipo de descuento no soportado: {tipo_descuento}")

    if modalidad == MODALIDAD_CLUB and submodo_club not in (SUBMODO_ORIGINAL, SUBMODO_CLON):
        raise ValueError("En modalidad CLUB debe indicar submodo_club='ORIGINAL' o 'CLON'.")

    config = REGLAS_POR_TIPO_Y_AREA[tipo_descuento][area_funcional]
    regla = _base_regla()

    # Resumen
    regla["resumen"]["tipo_descuento"] = tipo_descuento
    regla["resumen"]["competencia"] = config["competencia"]
    regla["resumen"]["aplicador_geo"] = config["aplicador"]
    regla["resumen"]["requiere_mensaje"] = tipo_descuento in REQUIERE_MENSAJE_TIPOS
    regla["resumen"]["divide_valor"] = config["divide_valor"]
    regla["resumen"]["origen_valor"] = config["origen_valor"]

    # Básico
    regla["basico"]["area_funcional"] = area_funcional
    regla["basico"]["competencia"] = config["competencia"]

    # Condiciones
    regla["condiciones"]["cantidad_cada"] = config["cantidad_condicion"]

    # Aplicadores
    regla["aplicadores"]["tipo"] = config["aplicador"]
    regla["aplicadores"]["origen_valor"] = config["origen_valor"]
    regla["aplicadores"]["divide_valor"] = config["divide_valor"]
    regla["aplicadores"]["formula_division"] = config.get("formula_division")
    regla["aplicadores"]["cantidad"] = config["cantidad_aplicador"]
    regla["aplicadores"]["estrategia"] = config["estrategia"]
    regla["aplicadores"]["por_unidad"] = config["por_unidad"]

    # Observaciones base
    regla["aplicadores"]["observaciones"].extend(config["notas"])

    if regla["resumen"]["requiere_mensaje"]:
        regla["alertas"].append(
            "Este tipo de promoción normalmente requiere campaña de mensaje asociada."
        )

    # Camino narrado
    regla["camino"] = [
        f"En Básico, definir competencia como '{config['competencia']}'.",
        f"En Condiciones, configurar cantidad cada {config['cantidad_condicion']}.",
        f"En Aplicadores, usar '{config['aplicador']}'.",
        f"Tomar el valor desde '{config['origen_valor']}'.",
    ]

    if config["divide_valor"]:
        regla["camino"].append(
            f"Dividir el valor usando: {config['formula_division']}."
        )
    else:
        regla["camino"].append("No dividir el valor.")

    regla["camino"].append(
        f"En Aplicadores, usar cantidad {config['cantidad_aplicador']}."
    )

    if config["estrategia"]:
        regla["camino"].append(
            f"Definir estrategia '{config['estrategia']}'."
        )

    if config["por_unidad"] is True:
        regla["camino"].append("Marcar 'Por unidad'.")
    elif config["por_unidad"] is False:
        regla["camino"].append("No marcar 'Por unidad'.")

    # Modalidad
    if modalidad == MODALIDAD_MASIVA:
        regla = _aplicar_masiva(regla, area_funcional)
    else:
        regla = _aplicar_club(regla, submodo_club=submodo_club)  # type: ignore[arg-type]

    # Observaciones universales
    regla["alertas"].append(
        "La condición y el aplicador deben apuntar a los mismos productos o la misma lista."
    )

    return regla


# ============================================================
# ESTRUCTURA AGRUPADA PARA CONSUMO RÁPIDO
# ============================================================

REGLAS_CONSULTOR: dict[str, Any] = {
    "catalogos": {
        "modalidades": [MODALIDAD_MASIVA, MODALIDAD_CLUB],
        "submodos_club": [SUBMODO_ORIGINAL, SUBMODO_CLON],
        "areas_funcionales": list(AREAS_FUNCIONALES),
        "tipos_descuento": list(TIPOS_DESCUENTO_SOPORTADOS),
    },
    "textos": {
        "competencia_producto": COMP_POR_PRODUCTO,
        "competencia_promocion": COMP_POR_PROMOCION,
        "nombre_general_normal": NOMBRE_GENERAL_NORMAL,
        "nombre_general_club": NOMBRE_GENERAL_CLUB,
        "area_geo_club": AREA_GEO_FIDELIZACION,
        "grupo_club": GRUPO_CLUB_NECESIDADES,
    },
    "requiere_mensaje_tipos": list(REQUIERE_MENSAJE_TIPOS),
    "reglas_por_tipo_y_area": REGLAS_POR_TIPO_Y_AREA,
}


if __name__ == "__main__":
    ejemplo = resolver_regla(
        modalidad=MODALIDAD_CLUB,
        area_funcional=AREA_BYCP,
        tipo_descuento=TIPO_PACK_NOMINAL,
        submodo_club=SUBMODO_CLON,
    )
    from pprint import pprint
    pprint(ejemplo)
