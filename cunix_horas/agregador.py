"""Pivot de registros de Kimai a la jerarquía Cliente > Proyecto > Actividad."""
from __future__ import annotations

import calendar
from collections import defaultdict
from dataclasses import dataclass, field

from cunix_horas.lector_kimai import Registro
from cunix_horas.mapeo import ErrorMapeo, Mapeo, Persona

# Decimales con los que se muestra cada hora en el Excel del partner.
DECIMALES = 2


def redondear(horas: float) -> float:
    """Redondeo único de horas: el que se escribe en la celda de día."""
    return round(horas, DECIMALES)


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

    @property
    def horas_por_dia_redondeadas(self) -> dict[int, float]:
        """Las horas tal como se escriben en la celda de día del Excel.

        Este es el único redondeo del sistema: todo lo demás que se muestra
        (totales de fila, de proyecto, de cliente y fila Total) se deriva
        sumando estos valores ya redondeados, para que el Excel cierre a la
        vista de quien suma una fila o una columna.
        """
        return {dia: redondear(horas) for dia, horas in self.horas_por_dia.items()}

    @property
    def total_redondeado(self) -> float:
        """Total de la actividad: suma de sus celdas de día ya redondeadas."""
        return redondear(sum(self.horas_por_dia_redondeadas.values()))


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

    @property
    def total_redondeado(self) -> float:
        """Total general tal como se escribe: suma de los totales de actividad."""
        return redondear(sum(f.total_redondeado for f in self.filas))

    def total_redondeado_del_dia(self, dia: int) -> float:
        """Total del día tal como se escribe: suma de las celdas de ese día."""
        return redondear(
            sum(f.horas_por_dia_redondeadas.get(dia, 0.0) for f in self.filas)
        )


def _verificar_export(registros: list[Registro], archivo: str) -> str:
    """Comprueba que el export sea de un solo desarrollador y tenga datos.

    Devuelve el único username encontrado. Lanza ErrorMapeo si el export está
    vacío o mezcla varios desarrolladores: en los dos casos el Excel resultante
    sería creíble y estaría mal, así que el archivo no se genera.
    """
    if not registros:
        raise ErrorMapeo(
            f"Export sin ningún registro de horas en {archivo}: el archivo tiene "
            f"los encabezados de Kimai pero ninguna fila de datos.\n"
            f"  Revisá el rango de fechas del export en Kimai: lo más probable es "
            f"que esté puesto en un período sin horas cargadas.\n"
            f"  Volvé a exportar el mes que querés generar y reemplazá el archivo."
        )

    usernames = sorted({r.username for r in registros})
    if len(usernames) > 1:
        raise ErrorMapeo(
            f"Export con horas de más de un desarrollador en {archivo}: "
            f"{', '.join(usernames)}\n"
            f"  Kimai tiene que exportarse filtrando por un solo desarrollador: "
            f"un archivo por persona.\n"
            f"  Si se genera igual, las horas de todos se le facturan a uno solo.\n"
            f"  Revisá el filtro de usuario en Kimai, volvé a exportar y dejá en "
            f"input/ un archivo por desarrollador."
        )
    return usernames[0]


def resolver_identidad(
    registros: list[Registro], mapeo: Mapeo, archivo: str
) -> Persona:
    """Verifica el export y devuelve la persona a la que corresponde.

    Es la primera parte de `agregar()`, expuesta aparte para que el CLI pueda
    conocer el nombre del Excel de salida ANTES de intentar generarlo. Sin eso
    no hay forma de saber qué archivo de `output/` va a ser reemplazado, y una
    falla posterior (por ejemplo un proyecto sin mapear) dejaría el Excel de la
    corrida anterior haciéndose pasar por el del mes.

    Lanza ErrorMapeo si el export está vacío, mezcla desarrolladores, o el
    username no está en mapeo.yaml.
    """
    username = _verificar_export(registros, archivo)
    return mapeo.resolver_persona(username, archivo)


def _verificar_hay_registros_del_mes(
    del_mes: list[Registro], descartados: list[Registro], anio: int, mes: int,
    archivo: str,
) -> None:
    """Falla si TODOS los registros del export quedaron fuera del mes.

    El Excel saldría con 0.0 h y código de salida 0: un reporte en blanco, con
    nombre y aspecto legítimos, que el dueño le manda al cliente sin que nada
    le haya avisado. Es el mismo error de rango de fechas en Kimai que el
    export vacío, por otro camino.
    """
    if del_mes or not descartados:
        return
    fechas = sorted(r.fecha for r in descartados)
    primera, ultima = fechas[0], fechas[-1]
    raise ErrorMapeo(
        f"Ningún registro del export cae dentro de {anio}-{mes:02d} en {archivo}: "
        f"los {len(descartados)} registros del archivo se descartaron por fecha.\n"
        f"  Las fechas del export van del {primera.day}/{primera.month}/{primera.year} "
        f"al {ultima.day}/{ultima.month}/{ultima.year}.\n"
        f"  Lo más probable es que el export se haya hecho con otro rango de fechas, "
        f"o que corresponda a otro mes del que estás generando.\n"
        f"  Volvé a exportar {anio}-{mes:02d} desde Kimai, o generá el mes que "
        f"realmente trae el archivo."
    )


def agregar(
    registros: list[Registro], mapeo: Mapeo, anio: int, mes: int, archivo: str
) -> Reporte:
    """Convierte registros planos en un Reporte jerárquico del mes indicado.

    Los registros con fecha fuera del mes se excluyen y quedan en `descartados`.
    Lanza ErrorMapeo si aparece un código de proyecto o un username sin mapear,
    si el export no tiene ningún registro, si mezcla varios desarrolladores, o
    si todos sus registros caen fuera del mes que se está generando.
    """
    persona = resolver_identidad(registros, mapeo, archivo)

    del_mes: list[Registro] = []
    descartados: list[Registro] = []
    for registro in registros:
        if (registro.fecha.year, registro.fecha.month) == (anio, mes):
            del_mes.append(registro)
        else:
            descartados.append(registro)

    _verificar_hay_registros_del_mes(del_mes, descartados, anio, mes, archivo)

    acumulado: dict[tuple[str, str, str], dict[int, float]] = defaultdict(dict)
    for registro in del_mes:
        destino = mapeo.resolver_proyecto(
            registro.cod_proyecto, registro.texto_proyecto, archivo
        )
        clave = (destino.cliente, destino.proyecto, registro.actividad)
        dia = registro.fecha.day
        acumulado[clave][dia] = acumulado[clave].get(dia, 0.0) + registro.horas

    filas = tuple(
        Fila(cliente, proyecto, actividad, dict(sorted(horas.items())))
        for (cliente, proyecto, actividad), horas in sorted(acumulado.items())
    )

    return Reporte(
        nombre_dev=persona.nombre,
        nombre_archivo=persona.archivo,
        anio=anio,
        mes=mes,
        filas=filas,
        descartados=tuple(descartados),
    )
