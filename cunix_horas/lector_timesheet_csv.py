"""Lectura del export .csv de timesheet plano de Kimai.

Mismo contenido que el timesheet .xlsx, pero con dos diferencias que importan:
la fecha viene en ISO (`2026-08-31`) en vez de serial de Excel, y la duración
viene en `H:MM` (`2:00`) en vez de fracción de día.

Las columnas se leen **por nombre de encabezado**, no por posición: Kimai
agrega y reordena columnas de una versión a otra, y leer por posición haría
que un export nuevo imputara horas equivocadas sin avisar.

Tres cosas se detectan por archivo en vez de asumirse, porque asumirlas hacía
desaparecer horas sin ruido:

1. **El separador de columnas**, que según el locale de Kimai es `,` o `;`.
2. **Los encabezados repetidos**: con dos columnas `Duration` ganaba la última.
3. **Las filas sin fecha**: si traen datos son horas que no se facturan.

Los encabezados se normalizan (se les saca el espacio de los bordes) **una
sola vez**, y tanto la verificación como la lectura usan esas mismas claves.
Antes la verificación normalizaba y la lectura no, así que un archivo con
`" Date "` pasaba la verificación y devolvía cero registros sin error.
"""
from __future__ import annotations

import csv
import io
import re
from datetime import date
from pathlib import Path

from cunix_horas.kimai_comun import ErrorLectura, Registro, codigo_de_proyecto

COL_FECHA = "Date"
COL_DURACION = "Duration"
COL_USERNAME = "User"
COL_PROYECTO = "Project"
COL_ACTIVIDAD = "Activity"

ENCABEZADOS_ESPERADOS = (
    COL_FECHA,
    COL_DURACION,
    COL_USERNAME,
    COL_PROYECTO,
    COL_ACTIVIDAD,
)

# Kimai emite uno u otro según el locale de la instalación.
DELIMITADORES = (",", ";")
NOMBRE_DE_DELIMITADOR = {",": "coma", ";": "punto y coma"}

_DURACION = re.compile(r"^\s*(\d+):([0-5]\d)(?::([0-5]\d))?\s*$")

MINUTOS_POR_HORA = 60
SEGUNDOS_POR_HORA = 3600


def _encabezados_con(primera_linea: str, delimitador: str) -> list[str]:
    """La fila 1 leída con un delimitador, ya normalizada."""
    for fila in csv.reader([primera_linea], delimiter=delimitador):
        return [celda.strip() for celda in fila]
    return []


def _elegir_delimitador(primera_linea: str) -> str:
    """El delimitador con el que la fila 1 da más columnas de las esperadas."""

    def aciertos(delimitador: str) -> int:
        encabezados = set(_encabezados_con(primera_linea, delimitador))
        return sum(1 for esperado in ENCABEZADOS_ESPERADOS if esperado in encabezados)

    mejor = max(DELIMITADORES, key=aciertos)
    # Sin ningún acierto el archivo no es un export de Kimai: se devuelve la
    # coma para que el error lo dé `_verificar_encabezados`, que sabe decir
    # qué columnas faltan.
    return mejor if aciertos(mejor) else ","


def _verificar_encabezados(
    encabezados: list[str], delimitador: str, ruta: Path
) -> None:
    presentes = set(encabezados)
    faltantes = [e for e in ENCABEZADOS_ESPERADOS if e not in presentes]
    if faltantes:
        raise ErrorLectura(
            f"{ruta.name} no tiene el formato de export de Kimai. "
            f"En la fila 1 faltan las columnas: {', '.join(faltantes)}\n"
            f"  Las columnas se leyeron separadas por "
            f"{NOMBRE_DE_DELIMITADOR[delimitador]}; también se probó con "
            f"{NOMBRE_DE_DELIMITADOR[';' if delimitador == ',' else ',']}."
        )


def _verificar_sin_repetidos(encabezados: list[str], ruta: Path) -> None:
    """Dos columnas con el mismo nombre: ganaría la última y nadie se entera."""
    repetidos = [e for e in ENCABEZADOS_ESPERADOS if encabezados.count(e) > 1]
    if repetidos:
        raise ErrorLectura(
            f"{ruta.name}, fila 1: hay más de una columna con el mismo nombre "
            f"({', '.join(repetidos)}).\n"
            f"  No se puede saber cuál de las dos trae las horas buenas, así "
            f"que no se lee nada.\n"
            f"  Exportá de nuevo desde Kimai sin agregar columnas a mano."
        )


def _fecha_de(valor: str, nro_fila: int, ruta: Path) -> date:
    """Fecha ISO de una fila, o ErrorLectura en español si no lo es."""
    try:
        return date.fromisoformat(valor.strip())
    except (ValueError, AttributeError):
        raise ErrorLectura(
            f"{ruta.name}, fila {nro_fila}, columna {COL_FECHA}: {valor!r} no "
            f"es una fecha con formato AAAA-MM-DD.\n"
            f"  Exportá de nuevo desde Kimai sin editar el archivo a mano."
        ) from None


def _horas_de(valor: str, nro_fila: int, ruta: Path) -> float:
    """Duración 'H:MM' (o 'H:MM:SS') en horas, o ErrorLectura en español."""
    coincidencia = _DURACION.match(valor or "")
    if coincidencia is None:
        raise ErrorLectura(
            f"{ruta.name}, fila {nro_fila}, columna {COL_DURACION}: {valor!r} "
            f"no es una duración con formato H:MM.\n"
            f"  Exportá de nuevo desde Kimai sin editar el archivo a mano."
        )
    horas, minutos, segundos = coincidencia.groups()
    return (
        int(horas)
        + int(minutos) / MINUTOS_POR_HORA
        + int(segundos or 0) / SEGUNDOS_POR_HORA
    )


def _error_de_fila_sin_fecha(fila: dict[str, str], nro_fila: int, ruta: Path):
    """Una fila con datos pero sin fecha son horas que no se facturarían."""
    duracion = (fila.get(COL_DURACION) or "").strip()
    detalle = f" (trae {duracion!r} en la columna {COL_DURACION})" if duracion else ""
    return ErrorLectura(
        f"{ruta.name}, fila {nro_fila}: la columna {COL_FECHA} está vacía "
        f"pero la fila trae datos{detalle}.\n"
        f"  Saltearla dejaría esas horas afuera del Excel del cliente sin "
        f"avisar, así que no se genera nada con este archivo.\n"
        f"  Completá la fecha en Kimai y exportá de nuevo, o borrá la fila "
        f"entera si no corresponde."
    )


def _leer_texto(ruta: Path) -> str:
    try:
        # utf-8-sig: el export puede traer BOM, y sin esto el primer
        # encabezado sería '﻿Date' y no matchearía por nombre.
        return ruta.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        raise ErrorLectura(
            f"{ruta.name} no está en UTF-8. Exportá de nuevo desde Kimai sin "
            f"abrirlo ni guardarlo con otro programa."
        ) from None
    except OSError as error:
        raise ErrorLectura(f"No se pudo leer {ruta.name}: {error}") from None


def leer_timesheet_csv(ruta: Path) -> list[Registro]:
    """Devuelve los registros de tiempo de un timesheet plano .csv."""
    texto = _leer_texto(ruta)
    primera_linea = texto.splitlines()[0] if texto else ""
    delimitador = _elegir_delimitador(primera_linea)

    with io.StringIO(texto, newline="") as flujo:
        crudas = list(csv.reader(flujo, delimiter=delimitador))

    encabezados = [celda.strip() for celda in crudas[0]] if crudas else []
    _verificar_sin_repetidos(encabezados, ruta)
    _verificar_encabezados(encabezados, delimitador, ruta)

    registros: list[Registro] = []
    for nro_fila, valores in enumerate(crudas[1:], start=2):
        if not any(valor.strip() for valor in valores):
            continue  # Fila completamente vacía: no hay nada que perder.
        fila = dict(zip(encabezados, valores))
        if not (fila.get(COL_FECHA) or "").strip():
            raise _error_de_fila_sin_fecha(fila, nro_fila, ruta)
        texto_proyecto = (fila.get(COL_PROYECTO) or "").strip()
        registros.append(
            Registro(
                fecha=_fecha_de(fila[COL_FECHA], nro_fila, ruta),
                horas=_horas_de(fila.get(COL_DURACION, ""), nro_fila, ruta),
                username=(fila.get(COL_USERNAME) or "").strip(),
                cod_proyecto=codigo_de_proyecto(
                    texto_proyecto, f"{ruta.name}, fila {nro_fila}"
                ),
                actividad=(fila.get(COL_ACTIVIDAD) or "").strip(),
                texto_proyecto=texto_proyecto,
            )
        )
    return registros
