"""Avisos sobre cargas de horas sospechosas. Nunca frenan la generación.

Se miran los registros crudos de cada persona, tal como salieron de Kimai, y
los avisos van agrupados bajo su nombre en `output/<mes>/_validacion.txt`.

Son cuatro, y los cuatro son para que el dueño mire, no para que la
herramienta decida: un registro fuera del mes (que no entra a los anexos),
más de 12 h en un día, horas cargadas en fin de semana y un día hábil sin
ninguna carga.

El aviso de desvío por redondeo que había acá se eliminó junto con el Excel
pivoteado por desarrollador: era el único formato que redondeaba cada celda de
día. Los anexos escriben las horas exactas y no hay nada que redondear.
"""
from __future__ import annotations

import calendar

from cunix_horas.lector_kimai import Registro

LIMITE_HORAS_POR_DIA = 12.0


def _fecha_legible(dia: int, mes: int, anio: int) -> str:
    return f"{dia}/{mes}/{anio}"


def avisos_de_desarrollador(
    del_mes: list[Registro],
    descartados: list[Registro],
    anio: int,
    mes: int,
) -> list[str]:
    """Los avisos de una persona sobre sus registros crudos del mes.

    No muta lo que recibe: sólo lee.
    """
    avisos: list[str] = []

    for registro in descartados:
        fecha = registro.fecha
        avisos.append(
            f"Registro fuera del mes, EXCLUIDO de los anexos: "
            f"{_fecha_legible(fecha.day, fecha.month, fecha.year)} "
            f"({registro.horas:.2f} h, proyecto {registro.cod_proyecto})"
        )

    horas_por_dia: dict[int, float] = {}
    for registro in del_mes:
        dia = registro.fecha.day
        horas_por_dia[dia] = horas_por_dia.get(dia, 0.0) + registro.horas

    for dia in range(1, calendar.monthrange(anio, mes)[1] + 1):
        horas = horas_por_dia.get(dia, 0.0)
        legible = _fecha_legible(dia, mes, anio)
        es_fin_de_semana = calendar.weekday(anio, mes, dia) >= 5

        if horas > LIMITE_HORAS_POR_DIA:
            avisos.append(
                f"Más de {LIMITE_HORAS_POR_DIA:.0f} h en un día: {legible} "
                f"tiene {horas:.1f} h"
            )
        if es_fin_de_semana and horas > 0:
            avisos.append(f"Horas cargadas en fin de semana: {legible} ({horas:.1f} h)")
        if not es_fin_de_semana and horas == 0:
            avisos.append(f"Día hábil sin carga: {legible}")

    return avisos
