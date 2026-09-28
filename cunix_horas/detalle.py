"""Armado del detalle plano: una fila por registro de tiempo de Kimai.

Es el entregable que recibe el partner desde que dejó de pedir el Excel
pivoteado por desarrollador: **un solo archivo por mes, con todos los
desarrolladores juntos**, una fila por cada carga de horas.

Ocho de las diez columnas son passthrough directo del export (fecha con hora
de inicio, duración, nombre, usuario, mail, actividad, descripción y número
de proyecto): salen tal cual vinieron, sin transformar.

Las otras dos, `Customer` y `Project`, salen de `config/mapeo.yaml` cuando el
proyecto está declarado, y si no se derivan del texto crudo de Kimai. Un
proyecto sin mapear **ya no frena nada**: el mapeo pasó de ser obligatorio a
ser un pulido opcional de nombres. Se puede hacer porque cada fila lleva su
`Project number`, así que la trazabilidad no depende del mapeo; lo único que
se pierde sin él es el nombre prolijo, y para eso alcanza con listarlo en el
informe y que el dueño decida.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time, timedelta

from cunix_horas.lector_kimai import Registro
from cunix_horas.mapeo import Mapeo, derivar_cliente, derivar_proyecto

SEGUNDOS_POR_HORA = 3600

# Hora que se usa cuando el registro no trae `From`. Kimai puede no traerla y
# no es un dato que se pueda inventar: la celda muestra sólo la fecha igual,
# porque su formato es `yyyy-mm-dd`.
HORA_SIN_INICIO = time(0, 0)


def segundos_de(horas: float) -> int:
    """Las horas de un registro, en segundos enteros.

    Las duraciones de Kimai llegan como `float` (el .xlsx las trae como
    fracción de día, y 1 h vale 1.000000000000008). El archivo del partner
    lleva duraciones reales, así que el redondeo a segundos se hace una sola
    vez, acá, y tanto la celda escrita como la verificación de integridad
    parten de este mismo número.
    """
    return round(horas * SEGUNDOS_POR_HORA)


@dataclass(frozen=True)
class FilaDetalle:
    """Una fila del archivo del partner, en el orden de sus columnas."""

    fecha_hora: datetime
    duracion: timedelta
    nombre: str
    username: str
    email: str
    cliente: str
    proyecto: str
    actividad: str
    descripcion: str
    numero_proyecto: str

    @property
    def valores(self) -> tuple:
        """Las diez celdas de la fila, en el orden de las columnas."""
        return (
            self.fecha_hora,
            self.duracion,
            self.nombre,
            self.username,
            self.email,
            self.cliente,
            self.proyecto,
            self.actividad,
            # El partner recibe la celda vacía, no un texto vacío: 112 de las
            # 160 filas de su archivo de referencia vienen así.
            self.descripcion or None,
            self.numero_proyecto,
        )


@dataclass(frozen=True)
class ProyectoSinMapear:
    """Un proyecto que salió con el nombre derivado de Kimai."""

    codigo: str
    cliente: str
    proyecto: str


@dataclass(frozen=True)
class Detalle:
    """Todas las filas del mes, ya ordenadas, con lo que el informe necesita."""

    filas: tuple[FilaDetalle, ...]
    sin_mapear: tuple[ProyectoSinMapear, ...]
    numeros_ambiguos: tuple[tuple[str, tuple[str, ...]], ...]

    @property
    def segundos(self) -> int:
        return sum(int(f.duracion.total_seconds()) for f in self.filas)

    @property
    def horas(self) -> float:
        return self.segundos / SEGUNDOS_POR_HORA


def nombre_para_mostrar(registro: Registro) -> str:
    """El nombre con el que se agrupa y ordena a un desarrollador.

    Es el `Name` de Kimai. Si el export no lo trae, se cae al username: es
    preferible agrupar por algo que dejar a una persona sin nombre y mezclada
    con las demás.
    """
    return registro.nombre.strip() or registro.username.strip()


def _clave_de_orden(par: tuple[int, Registro]) -> tuple:
    """Agrupado por desarrollador (alfabético) y cronológico dentro de cada uno.

    El índice original entra al final para que dos registros del mismo momento
    salgan siempre en el mismo orden: sin eso el archivo cambiaría de una
    corrida a la siguiente sin que cambien los datos, y comparar dos envíos
    dejaría de servir.
    """
    indice, registro = par
    nombre = nombre_para_mostrar(registro)
    return (
        nombre.casefold(),
        nombre,
        registro.fecha,
        registro.hora_inicio or HORA_SIN_INICIO,
        indice,
    )


def _fila_de(registro: Registro, mapeo: Mapeo) -> tuple[FilaDetalle, str, str]:
    """La fila del partner de un registro. Devuelve además cliente y proyecto."""
    destino = mapeo.proyecto_opcional(registro.cod_proyecto)
    if destino is not None:
        cliente, proyecto = destino.cliente, destino.proyecto
    else:
        cliente = derivar_cliente(registro.texto_cliente)
        proyecto = derivar_proyecto(registro.texto_proyecto) or registro.cod_proyecto

    fila = FilaDetalle(
        fecha_hora=datetime.combine(
            registro.fecha, registro.hora_inicio or HORA_SIN_INICIO
        ),
        duracion=timedelta(seconds=segundos_de(registro.horas)),
        nombre=nombre_para_mostrar(registro),
        username=registro.username,
        email=registro.email,
        cliente=cliente,
        proyecto=proyecto,
        actividad=registro.actividad,
        descripcion=registro.descripcion,
        numero_proyecto=registro.numero_proyecto,
    )
    return fila, cliente, proyecto


def construir(registros: list[Registro], mapeo: Mapeo) -> Detalle:
    """Convierte los registros de todos los desarrolladores en el detalle plano."""
    ordenados = [par[1] for par in sorted(enumerate(registros), key=_clave_de_orden)]

    filas: list[FilaDetalle] = []
    sin_mapear: dict[str, ProyectoSinMapear] = {}
    nombres_por_numero: dict[str, list[str]] = {}

    for registro in ordenados:
        fila, cliente, proyecto = _fila_de(registro, mapeo)
        filas.append(fila)

        if mapeo.proyecto_opcional(registro.cod_proyecto) is None:
            sin_mapear.setdefault(
                registro.cod_proyecto,
                ProyectoSinMapear(registro.cod_proyecto, cliente, proyecto),
            )

        numero = registro.numero_proyecto.strip()
        if numero:
            nombres = nombres_por_numero.setdefault(numero, [])
            if proyecto not in nombres:
                nombres.append(proyecto)

    ambiguos = tuple(
        (numero, tuple(sorted(nombres)))
        for numero, nombres in sorted(nombres_por_numero.items())
        if len(nombres) > 1
    )

    return Detalle(
        filas=tuple(filas),
        sin_mapear=tuple(sin_mapear[codigo] for codigo in sorted(sin_mapear)),
        numeros_ambiguos=ambiguos,
    )
