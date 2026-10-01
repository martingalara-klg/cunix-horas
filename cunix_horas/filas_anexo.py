"""Las filas de la hoja `Detalle` del Anexo II-A, y lo que se agrupa de ellas.

Una fila por registro de tiempo, con las siete columnas que KLG completa:

    Fecha | Inicio | Fin | Persona | Proyecto | Descripción | Horas

Las otras seis columnas de la hoja (`Alertas`, `Revisión C.UNIX`, `Horas
aprobadas`, `Horas a pagar`, `Observación C.UNIX` y `Día nuevo`) son de
C.UNIX: tres son fórmulas que ya vienen escritas en la plantilla y tres las
completa él al revisar. La herramienta no las toca.

`Fin` **no** lo trae el export de Kimai: se calcula como `Inicio + Horas`. Es
aritmética sobre datos registrados, no una reconstrucción: Kimai guarda cada
registro como un bloque continuo. Cuando el export no trae hora de inicio (el
resumen mensual no la trae), `Inicio` y `Fin` quedan vacíos, que es lo que la
plantilla espera.

De estas mismas filas salen las tres tablas del informe, para que informe y
anexo no puedan discrepar: no hay dos caminos de cálculo.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, replace
from datetime import date, time

from cunix_horas.kimai_comun import Registro
from cunix_horas.mapeo import Mapeo, derivar_proyecto

MINUTOS_POR_HORA = 60
MINUTOS_POR_DIA = 24 * 60

# Tolerancia al comparar horas. Las duraciones del .xlsx de Kimai llegan como
# fracción de día, y 1 h vale 1.000000000000008.
EPSILON = 1e-6

# Los tickets que C.UNIX reconoce: iTop ('R-012532'), ClickUp ('CU-8xyz') y los
# BUG que el equipo usó en agosto. La fórmula de `Alertas` de la plantilla
# busca 'R-0' y 'CU-'; acá se agrega BUG para que el borrador de la tabla 3
# agrupe por algo útil en vez de tirar todo a «Sin ticket».
_TICKET = re.compile(r"\b(?:R-\d+|CU-[A-Za-z0-9]+|BUG-\d+)\b")

SIN_TICKET = "Sin ticket"


@dataclass(frozen=True)
class FilaAnexo:
    """Una fila de la hoja `Detalle`."""

    fecha: date
    inicio: time | None
    fin: time | None
    persona: str
    proyecto: str
    descripcion: str
    horas: float
    # Columna `Nota KLG`, que la herramienta agrega a la derecha de las de
    # C.UNIX cuando una fila necesita explicación (hoy, las horas separadas de
    # una persona sin usuario de Kimai). Vacía en la mayoría de las filas.
    nota: str = ""


@dataclass(frozen=True)
class ProyectoSinMapear:
    """Un proyecto que salió con el nombre derivado del texto de Kimai."""

    codigo: str
    proyecto: str


@dataclass(frozen=True)
class Armado:
    """Las filas del mes y lo que el informe de validación necesita saber."""

    filas: tuple[FilaAnexo, ...]
    sin_mapear: tuple[ProyectoSinMapear, ...]
    # persona -> usuario de Kimai, para la hoja `Datos`. Una persona sin
    # usuario (las horas que llegan por una fuente manual) queda en "".
    usuarios: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class TotalProyecto:
    """Una fila de la tabla 1 del informe: horas e importe de un proyecto."""

    proyecto: str
    cliente: str
    horas: float
    valor_hora: float | None

    @property
    def importe(self) -> float | None:
        return None if self.valor_hora is None else self.horas * self.valor_hora


@dataclass(frozen=True)
class TotalPersona:
    """Una fila de la tabla 2 del informe: horas, días y promedio."""

    persona: str
    perfil: str
    horas: float
    dias: int

    @property
    def promedio(self) -> float:
        return self.horas / self.dias if self.dias else 0.0


@dataclass(frozen=True)
class TrabajoDestacado:
    """Una fila del BORRADOR de la tabla 3 del informe."""

    proyecto: str
    ticket: str
    descripcion: str
    horas: float


def nombre_para_mostrar(registro: Registro) -> str:
    """El nombre con el que se agrupa a una persona en los dos anexos.

    Es el `Name` de Kimai. Si el export no lo trae, se cae al username: es
    preferible agrupar por algo que dejar a una persona sin nombre.
    """
    return registro.nombre.strip() or registro.username.strip()


def _a_minutos(hora: time) -> int:
    return hora.hour * MINUTOS_POR_HORA + hora.minute


def _a_hora(minutos: int) -> time:
    minutos %= MINUTOS_POR_DIA
    return time(minutos // MINUTOS_POR_HORA, minutos % MINUTOS_POR_HORA)


def sumar(hora: time, horas: float) -> time:
    """`hora` más una duración en horas, dando la vuelta a la medianoche.

    Un bloque que cruza la medianoche es real y la plantilla lo marca con
    «Cruza la medianoche»: no se corrige, se deja ver.
    """
    return _a_hora(_a_minutos(hora) + round(horas * MINUTOS_POR_HORA))


def restar(hora: time, horas: float) -> time:
    """`hora` menos una duración en horas, dando la vuelta a la medianoche."""
    return _a_hora(_a_minutos(hora) - round(horas * MINUTOS_POR_HORA))


def _nombre_de_proyecto(registro: Registro, mapeo: Mapeo) -> tuple[str, bool]:
    """El nombre del proyecto para el anexo, y si salió del mapeo o no."""
    destino = mapeo.proyecto_opcional(registro.cod_proyecto)
    if destino is not None:
        return destino.proyecto, True
    derivado = derivar_proyecto(registro.texto_proyecto) or registro.cod_proyecto
    return derivado, False


def _fila_de(registro: Registro, mapeo: Mapeo) -> tuple[FilaAnexo, bool]:
    proyecto, mapeado = _nombre_de_proyecto(registro, mapeo)
    inicio = registro.hora_inicio
    return (
        FilaAnexo(
            fecha=registro.fecha,
            inicio=inicio,
            fin=None if inicio is None else sumar(inicio, registro.horas),
            persona=nombre_para_mostrar(registro),
            proyecto=proyecto,
            descripcion=(registro.descripcion or "").strip(),
            horas=registro.horas,
        ),
        mapeado,
    )


def clave_de_orden(fila: FilaAnexo) -> tuple:
    """Cronológico y, dentro del día, por persona: el orden de la hoja."""
    return (
        fila.fecha,
        fila.persona.casefold(),
        fila.persona,
        _a_minutos(fila.inicio) if fila.inicio is not None else -1,
        fila.proyecto,
    )


def ordenar(filas) -> tuple[FilaAnexo, ...]:
    return tuple(sorted(filas, key=clave_de_orden))


def construir(registros: list[Registro], mapeo: Mapeo) -> Armado:
    """Las filas del anexo a partir de los registros ya leídos de Kimai.

    No muta nada de lo que recibe.
    """
    filas: list[FilaAnexo] = []
    sin_mapear: dict[str, ProyectoSinMapear] = {}
    usuarios: dict[str, str] = {}

    for registro in registros:
        fila, mapeado = _fila_de(registro, mapeo)
        filas.append(fila)
        if not mapeado:
            sin_mapear.setdefault(
                registro.cod_proyecto,
                ProyectoSinMapear(registro.cod_proyecto, fila.proyecto),
            )
        usuarios.setdefault(fila.persona, registro.username.strip())

    return Armado(
        filas=ordenar(filas),
        sin_mapear=tuple(sin_mapear[codigo] for codigo in sorted(sin_mapear)),
        usuarios=tuple(sorted(usuarios.items())),
    )


def total_de(filas) -> float:
    return sum(f.horas for f in filas)


def por_proyecto(filas, valores_hora: dict[str, tuple[str, float]]):
    """Las filas de la tabla 1, ordenadas de más a menos horas.

    `valores_hora` es lo que declara la hoja `Datos` de la plantilla:
    proyecto -> (cliente, valor hora). Un proyecto que no esté ahí sale con
    las horas y sin importe, y el informe de validación lo nombra: inventar un
    valor hora sería inventar plata.
    """
    acumulado: dict[str, float] = {}
    for fila in filas:
        acumulado[fila.proyecto] = acumulado.get(fila.proyecto, 0.0) + fila.horas

    totales = []
    for proyecto, horas in acumulado.items():
        cliente, valor = valores_hora.get(proyecto, ("", None))
        totales.append(TotalProyecto(proyecto, cliente, horas, valor))
    return tuple(sorted(totales, key=lambda t: (-t.horas, t.proyecto)))


def por_persona(filas, perfiles: dict[str, str], perfil_por_defecto: str):
    """Las filas de la tabla 2, ordenadas de más a menos horas."""
    horas: dict[str, float] = {}
    dias: dict[str, set] = {}
    for fila in filas:
        horas[fila.persona] = horas.get(fila.persona, 0.0) + fila.horas
        dias.setdefault(fila.persona, set()).add(fila.fecha)

    totales = [
        TotalPersona(
            persona,
            perfiles.get(persona, perfil_por_defecto),
            total,
            len(dias[persona]),
        )
        for persona, total in horas.items()
    ]
    return tuple(sorted(totales, key=lambda t: (-t.horas, t.persona)))


def ticket_de(descripcion: str) -> str:
    """El primer ticket que cita la descripción, o 'Sin ticket'."""
    coincidencia = _TICKET.search(descripcion or "")
    return coincidencia.group(0) if coincidencia else SIN_TICKET


def principales_trabajos(filas):
    """BORRADOR de la tabla 3: agrupado por proyecto y ticket, mecánicamente.

    OJO, ESTO NO ES EL ENTREGABLE FINAL. Es un punto de partida para que el
    dueño lo edite; la prosa la refina él todos los meses.

    La agrupación es mecánica: por proyecto y por el ticket que cita cada
    descripción, con las horas sumadas y las descripciones concatenadas. La
    tabla que C.UNIX armó para agosto de 2026 estaba agrupada **por tema**,
    leyendo las descripciones una por una: partió las 45,5 h «sin ticket» en
    tres líneas temáticas, fusionó tres tickets en una sola línea, y le sumó a
    un ticket trabajo relacionado que no lo citaba. Eso es criterio editorial
    —saber qué trabajos cuentan la misma historia— y no se automatiza sin
    inventar.

    Lo que sí garantiza esta función es la aritmética: cada fila del mes cae en
    exactamente un grupo, así que los subtotales cierran exactos y el total da
    el total del mes. El dueño puede reescribir los textos y fusionar líneas
    sumando sus horas, sin que el total deje de cuadrar.
    """
    acumulado: dict[tuple[str, str], list[FilaAnexo]] = {}
    for fila in filas:
        acumulado.setdefault(
            (fila.proyecto, ticket_de(fila.descripcion)), []
        ).append(fila)

    trabajos = []
    for (proyecto, ticket), agrupadas in acumulado.items():
        descripciones: list[str] = []
        for fila in sorted(agrupadas, key=clave_de_orden):
            texto = fila.descripcion.strip()
            if texto and texto not in descripciones:
                descripciones.append(texto)
        trabajos.append(
            TrabajoDestacado(
                proyecto=proyecto,
                ticket=ticket,
                descripcion="; ".join(descripciones),
                horas=total_de(agrupadas),
            )
        )
    return tuple(sorted(trabajos, key=lambda t: (t.proyecto, -t.horas, t.ticket)))


def sin_descripcion(filas) -> tuple[tuple[str, int], ...]:
    """Cuántos registros sin descripción tiene cada persona.

    El contrato de C.UNIX exige la descripción para aprobar las horas, así que
    esto va al informe de validación con nombre y cantidad. No frena nada: el
    dueño decide si las completa o las manda así.
    """
    cuenta: dict[str, int] = {}
    for fila in filas:
        if not fila.descripcion.strip():
            cuenta[fila.persona] = cuenta.get(fila.persona, 0) + 1
    return tuple(sorted(cuenta.items()))


def con_descripcion(fila: FilaAnexo, descripcion: str) -> FilaAnexo:
    """La misma fila con otra descripción. No muta la original."""
    return replace(fila, descripcion=descripcion)
