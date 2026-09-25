"""Avisos sobre cargas de horas sospechosas. Nunca frenan la generación."""
from __future__ import annotations

import calendar

from cunix_horas.agregador import Reporte

LIMITE_HORAS_POR_DIA = 12.0

# Cada celda de día se redondea a 2 decimales, así que aporta como mucho
# 0.005 h de desvío contra las horas crudas del export. El desvío del total
# NO es "unos centésimos": escala con la cantidad de celdas de día no vacías
# (un mes de 154 h con muchas celdas ya muestra 153.72: 0.28 h de desvío).
#
# El desvío se informa SIEMPRE como dato (ver `dato_de_desvio`), y a partir de
# este umbral además se avisa. 0.25 h es el punto en que deja de ser ruido de
# presentación: un dev-mes realista de unas 110 celdas de día (duraciones de
# Kimai como 50 min, que valen 0.8333 h y pierden 0.0033 h al redondearse) se
# aparta 0.37 h. A tarifa de software factory eso es plata todos los meses, y
# con el umbral viejo de 0.5 h pasaba en silencio.
UMBRAL_DESVIO_REDONDEO = 0.25


def _fecha_legible(dia: int, mes: int, anio: int) -> str:
    return f"{dia}/{mes}/{anio}"


def validar(reporte: Reporte) -> list[str]:
    """Devuelve la lista de avisos de un reporte."""
    avisos: list[str] = []

    for registro in reporte.descartados:
        fecha = registro.fecha
        avisos.append(
            f"Registro fuera del mes, EXCLUIDO del Excel: "
            f"{_fecha_legible(fecha.day, fecha.month, fecha.year)} "
            f"({registro.horas:.2f} h, proyecto {registro.cod_proyecto})"
        )

    for dia in range(1, reporte.dias_del_mes + 1):
        horas = reporte.total_del_dia(dia)
        legible = _fecha_legible(dia, reporte.mes, reporte.anio)
        es_fin_de_semana = calendar.weekday(reporte.anio, reporte.mes, dia) >= 5

        if horas > LIMITE_HORAS_POR_DIA:
            avisos.append(
                f"Más de {LIMITE_HORAS_POR_DIA:.0f} h en un día: {legible} "
                f"tiene {horas:.1f} h"
            )
        if es_fin_de_semana and horas > 0:
            avisos.append(f"Horas cargadas en fin de semana: {legible} ({horas:.1f} h)")
        if not es_fin_de_semana and horas == 0:
            avisos.append(f"Día hábil sin carga: {legible}")

    # El total que el cliente ve es la suma de celdas ya redondeadas, así que
    # puede apartarse de las horas reales del export: el error de cada celda se
    # acumula. Comparar redondeado contra redondeado sería una tautología que
    # no puede dispararse nunca; lo que hay que vigilar es esta otra diferencia.
    desvio = _desvio_por_redondeo(reporte)
    if desvio > UMBRAL_DESVIO_REDONDEO:
        avisos.append(
            f"Desvío por redondeo de {desvio:.2f} h: el Excel totaliza "
            f"{reporte.total_redondeado:.2f} h y las horas del export que caen "
            f"en {reporte.mes}/{reporte.anio} suman {reporte.total:.2f} h. "
            f"Ese segundo número NO es el total que muestra Kimai si hay "
            f"registros fuera del mes: esos se descartan y no cuentan acá. "
            f"No es un error de carga: cada celda de día se redondea a 2 "
            f"decimales y esas diferencias se suman. Decidí vos si esa "
            f"diferencia importa para facturar."
        )

    return avisos


def _desvio_por_redondeo(reporte: Reporte) -> float:
    """Cuánto se aparta el total que muestra el Excel de las horas del mes."""
    return abs(reporte.total_redondeado - reporte.total)


def dato_de_desvio(reporte: Reporte) -> str:
    """Línea informativa con el desvío por redondeo, supere o no el umbral.

    Va SIEMPRE al informe, aparte de los avisos: un desvío por debajo del
    umbral no amerita alarma, pero el dueño factura con estos números y tiene
    derecho a verlos todos los meses en vez de enterarse sólo cuando saltan.
    """
    desvio = _desvio_por_redondeo(reporte)
    return (
        f"Desvío por redondeo: {desvio:.2f} h. El Excel totaliza "
        f"{reporte.total_redondeado:.2f} h y las horas del export que caen en "
        f"{reporte.mes}/{reporte.anio} suman {reporte.total:.2f} h "
        f"(los registros fuera del mes no cuentan acá)."
        + ("" if desvio > UMBRAL_DESVIO_REDONDEO else " Por debajo del umbral"
           f" de {UMBRAL_DESVIO_REDONDEO:.2f} h: sólo informativo.")
    )
