"""Pivot de registros de Kimai a la jerarquía Cliente > Proyecto > Actividad."""
from __future__ import annotations

import calendar
from collections import defaultdict
from dataclasses import dataclass, field

from cunix_horas.lector_kimai import Registro
from cunix_horas.mapeo import Mapeo


@dataclass(frozen=True)
class Fila:
    """Una actividad de un proyecto de un cliente, con sus horas por día."""

    cliente: str
    proyecto: str
    actividad: str
    horas_por_dia: dict[int, float] = field(default_factory=dict)

    @property
    def total(self) -> float:
        return sum(self.horas_por_dia.values())


@dataclass(frozen=True)
class Reporte:
    """Todo lo que hace falta para escribir el Excel de un desarrollador."""

    nombre_dev: str
    nombre_archivo: str
    anio: int
    mes: int
    filas: tuple[Fila, ...]
    descartados: tuple[Registro, ...]

    @property
    def dias_del_mes(self) -> int:
        return calendar.monthrange(self.anio, self.mes)[1]

    @property
    def total(self) -> float:
        return sum(f.total for f in self.filas)

    def total_del_dia(self, dia: int) -> float:
        return sum(f.horas_por_dia.get(dia, 0.0) for f in self.filas)


def agregar(
    registros: list[Registro], mapeo: Mapeo, anio: int, mes: int, archivo: str
) -> Reporte:
    """Convierte registros planos en un Reporte jerárquico del mes indicado.

    Los registros con fecha fuera del mes se excluyen y quedan en `descartados`.
    Lanza ErrorMapeo si aparece un código de proyecto o un username sin mapear.
    """
    del_mes: list[Registro] = []
    descartados: list[Registro] = []
    for registro in registros:
        if (registro.fecha.year, registro.fecha.month) == (anio, mes):
            del_mes.append(registro)
        else:
            descartados.append(registro)

    acumulado: dict[tuple[str, str, str], dict[int, float]] = defaultdict(dict)
    for registro in del_mes:
        destino = mapeo.resolver_proyecto(registro.cod_proyecto, "", archivo)
        clave = (destino.cliente, destino.proyecto, registro.actividad)
        dia = registro.fecha.day
        acumulado[clave][dia] = acumulado[clave].get(dia, 0.0) + registro.horas

    filas = tuple(
        Fila(cliente, proyecto, actividad, dict(sorted(horas.items())))
        for (cliente, proyecto, actividad), horas in sorted(acumulado.items())
    )

    usernames = {r.username for r in registros}
    if usernames:
        persona = mapeo.resolver_persona(sorted(usernames)[0], archivo)
        nombre_dev, nombre_archivo = persona.nombre, persona.archivo
    else:
        nombre_dev, nombre_archivo = "", ""

    return Reporte(
        nombre_dev=nombre_dev,
        nombre_archivo=nombre_archivo,
        anio=anio,
        mes=mes,
        filas=filas,
        descartados=tuple(descartados),
    )
