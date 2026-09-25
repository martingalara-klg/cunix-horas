"""Avisos sobre cargas de horas sospechosas. Nunca frenan la generación."""
from __future__ import annotations

import calendar

from cunix_horas.agregador import Reporte

LIMITE_HORAS_POR_DIA = 12.0
TOLERANCIA = 0.01


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

    # Comparación sobre los valores REDONDEADOS, que son los que el cliente ve
    # y suma en el Excel. Compararlos sin redondear dejaba la verificación
    # ciega justo al riesgo que tiene que cubrir: un Excel cuyas filas no
    # cierran a la vista aunque el total interno sea exacto.
    total_mostrado = reporte.total_redondeado
    suma_por_dia = sum(
        reporte.total_redondeado_del_dia(d)
        for d in range(1, reporte.dias_del_mes + 1)
    )
    if abs(suma_por_dia - total_mostrado) > TOLERANCIA:
        avisos.append(
            f"Descuadre de horas: el total del mes es {total_mostrado:.2f} h pero la "
            f"suma de los días da {suma_por_dia:.2f} h. No envíes este Excel y reportá "
            f"el problema al equipo que mantiene la herramienta."
        )

    return avisos
