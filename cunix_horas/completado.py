"""Completa, desde `config/mapeo.yaml`, lo que el resumen mensual no trae.

Kimai da dos reportes y los dos sirven, pero no traen lo mismo.

El **reporte de detalle** trae todo. Sus registros pasan por acá **sin
tocarse**: su usuario es el de Kimai, aunque el mapeo declare otra cosa. El
mapeo es un respaldo para lo que la fuente no trae, no una corrección de lo
que sí trae.

El **resumen mensual** es la grilla de días: trae el nombre del
desarrollador, las horas y el proyecto, y nada más. De lo que le falta, los
dos anexos sólo escriben una cosa: el **usuario de Kimai** de la columna
«Usuario Kimai» de la hoja `Datos`. Ese usuario es la clave de `personas:`.

El mail y el número de proyecto ya no se completan ni se piden: ninguna de
las catorce columnas del Anexo II-A ni ninguna tabla del Anexo II los
escribe. Pedir un dato que después no se escribe en ningún lado sólo sirve
para trabar una entrega.

**Nada de esto frena un archivo.** Si la persona no está en el mapeo, o si su
nombre está declarado en dos entradas, el usuario queda vacío —igual que el
de quien todavía no tiene usuario de Kimai, que es un hecho legítimo— y el
informe de validación lo avisa con nombre y apellido. C.UNIX pide que cada
persona cargue sus horas con su propio usuario, así que eso tiene que verse;
pero verse es un aviso, no un bloqueo: las horas y las filas salen completas
igual.
"""
from __future__ import annotations

from dataclasses import dataclass, replace

from cunix_horas.kimai_comun import ORIGEN_RESUMEN_MENSUAL, Registro
from cunix_horas.mapeo import Mapeo, Persona

_POR_QUE_IMPORTA = (
    "  C.UNIX pide que cada persona cargue sus horas con su propio usuario de "
    "Kimai, así que esa celda vacía puede hacer que pregunte por esas horas "
    "antes de aprobarlas."
)

_SALIDA_POR_EL_DETALLE = (
    "  La otra salida es volver a exportar a esa persona desde Kimai con el "
    "reporte de detalle: ese export trae el usuario en cada fila y el mapeo "
    "no interviene."
)

_CIERRE = "  Las horas y las filas de esa persona entran igual: esto no frena nada."


@dataclass(frozen=True)
class Completado:
    """Los registros ya completados y los avisos que dejó completarlos.

    Los avisos no frenan: van tal cual a `output/<mes>/_validacion.txt`.
    """

    registros: tuple[Registro, ...]
    avisos: tuple[str, ...] = ()


def _aviso_de_nombre_repetido(
    nombre_dev: str, usernames: tuple[str, ...], archivo: str
) -> str:
    """El nombre del resumen mensual está declarado en dos entradas."""
    return (
        f"{nombre_dev} ({archivo}) va a la hoja «Datos» SIN usuario de Kimai: "
        f"en config/mapeo.yaml hay más de una persona con ese 'nombre:' "
        f"({', '.join(usernames)}).\n"
        f"  Ese export es un resumen mensual: no trae el username, sólo el "
        f"nombre, así que no hay forma de saber cuál de las dos es sin "
        f"inventar.\n" + _POR_QUE_IMPORTA + "\n"
        f"  Cambiá el 'nombre:' de una de ellas para que no se repita y volvé "
        f"a correr.\n" + _CIERRE
    )


def _aviso_de_persona_sin_mapear(nombre_dev: str, archivo: str) -> str:
    """La persona del resumen mensual no está declarada en `personas:`."""
    return (
        f"{nombre_dev} ({archivo}) va a la hoja «Datos» SIN usuario de Kimai: "
        f"no está en config/mapeo.yaml.\n"
        f"  Ese export es un resumen mensual, que no trae el usuario: el "
        f"único lugar de donde puede salir es el mapeo, y ahí ninguna persona "
        f"coincide con ese nombre.\n" + _POR_QUE_IMPORTA + "\n"
        f"  Si la persona ya está cargada, hacé que su 'nombre:' coincida "
        f"exactamente con lo de arriba. Si no está, agregá bajo personas:, "
        f"usando como clave su usuario de Kimai:\n"
        f"  AJUSTAR-username-de-kimai:\n"
        f'    nombre: "{nombre_dev}"\n'
        f'    archivo: "AJUSTAR - apellido, va en el nombre del archivo"\n'
        + _SALIDA_POR_EL_DETALLE
        + "\n"
        + _CIERRE
    )


def _usuario_de(
    nombre_dev: str, mapeo: Mapeo, archivo: str, avisos: list[str]
) -> str:
    """El usuario de Kimai de esa persona, o vacío con el aviso correspondiente."""
    persona: Persona | None = mapeo.persona_opcional(nombre_dev)
    if persona is not None:
        return persona.username

    usernames = mapeo.usernames_con_el_nombre(nombre_dev)
    if len(usernames) > 1:
        avisos.append(_aviso_de_nombre_repetido(nombre_dev, usernames, archivo))
    else:
        avisos.append(_aviso_de_persona_sin_mapear(nombre_dev, archivo))
    return ""


def completar_desde_mapeo(
    registros: list[Registro], mapeo: Mapeo, archivo: str
) -> Completado:
    """Los mismos registros, con el usuario de Kimai ya resuelto, y los avisos.

    Devuelve registros nuevos: no muta ni los que recibe (son `frozen`) ni la
    lista. Los del reporte de detalle salen tal como entraron, sin mirar el
    mapeo.

    Un desarrollador del resumen mensual que el mapeo no resuelve sale con el
    usuario vacío y deja un aviso. Nunca lanza: lo que el mapeo no tiene ya no
    frena ningún archivo.
    """
    usuarios: dict[str, str] = {}
    avisos: list[str] = []
    completados: list[Registro] = []

    for registro in registros:
        if registro.origen != ORIGEN_RESUMEN_MENSUAL:
            completados.append(registro)
            continue

        # El lector del resumen mensual deja el nombre para mostrar en
        # `username`, porque ese export no tiene otra cosa. Acá se separan las
        # dos columnas: `nombre` es ese nombre, y `username` es la clave de la
        # persona en el mapeo.
        nombre_dev = (registro.nombre or registro.username).strip()
        if nombre_dev not in usuarios:
            usuarios[nombre_dev] = _usuario_de(nombre_dev, mapeo, archivo, avisos)
        completados.append(
            replace(registro, nombre=nombre_dev, username=usuarios[nombre_dev])
        )

    return Completado(tuple(completados), tuple(avisos))
