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
# A partir de acá el desvío se avisa. Media hora es el punto en que la
# diferencia empieza a ser discutible en una factura; por debajo es ruido de
# presentación que no vale la pena poner delante del dueño todos los meses.
# 0.5 h equivale al peor caso de unas 100 celdas de día no vacías.
UMBRAL_DESVIO_REDONDEO = 0.5


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
    total_mostrado = reporte.total_redondeado
    total_crudo = reporte.total
    desvio = abs(total_mostrado - total_crudo)
    if desvio > UMBRAL_DESVIO_REDONDEO:
        avisos.append(
            f"Desvío por redondeo: el Excel totaliza {total_mostrado:.2f} h y el "
            f"export de Kimai trae {total_crudo:.2f} h, una diferencia de "
            f"{desvio:.2f} h. No es un error de carga: cada celda de día se "
            f"redondea a 2 decimales y esas diferencias se suman. Decidí vos si "
            f"esa diferencia importa para facturar."
        )

    return avisos
