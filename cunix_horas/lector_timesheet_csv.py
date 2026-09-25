"""Lectura del export .csv de timesheet plano de Kimai.

Mismo contenido que el timesheet .xlsx, pero con dos diferencias que importan:
la fecha viene en ISO (`2026-08-31`) en vez de serial de Excel, y la duración
viene en `H:MM` (`2:00`) en vez de fracción de día.

Las columnas se leen **por nombre de encabezado**, no por posición: Kimai
agrega y reordena columnas de una versión a otra, y leer por posición haría
que un export nuevo imputara horas equivocadas sin avisar.
"""
from __future__ import annotations

import csv
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

_DURACION = re.compile(r"^\s*(\d+):([0-5]\d)(?::([0-5]\d))?\s*$")

MINUTOS_POR_HORA = 60
SEGUNDOS_POR_HORA = 3600


def _verificar_encabezados(encabezados: list[str] | None, ruta: Path) -> None:
    presentes = {e.strip() for e in (encabezados or [])}
    faltantes = [e for e in ENCABEZADOS_ESPERADOS if e not in presentes]
    if faltantes:
        raise ErrorLectura(
            f"{ruta.name} no tiene el formato de export de Kimai. "
            f"En la fila 1 faltan las columnas: {', '.join(faltantes)}"
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


def leer_timesheet_csv(ruta: Path) -> list[Registro]:
    """Devuelve los registros de tiempo de un timesheet plano .csv."""
    try:
        # utf-8-sig: el export puede traer BOM, y sin esto el primer
        # encabezado sería '﻿Date' y no matchearía por nombre.
        with ruta.open(encoding="utf-8-sig", newline="") as archivo:
            lector = csv.DictReader(archivo)
            _verificar_encabezados(lector.fieldnames, ruta)
            filas = list(enumerate(lector, start=2))
    except UnicodeDecodeError:
        raise ErrorLectura(
            f"{ruta.name} no está en UTF-8. Exportá de nuevo desde Kimai sin "
            f"abrirlo ni guardarlo con otro programa."
        ) from None
    except OSError as error:
        raise ErrorLectura(f"No se pudo leer {ruta.name}: {error}") from None

    registros: list[Registro] = []
    for nro_fila, fila in filas:
        if not (fila.get(COL_FECHA) or "").strip():
            continue
        texto_proyecto = (fila.get(COL_PROYECTO) or "").strip()
        registros.append(
            Registro(
                fecha=_fecha_de(fila[COL_FECHA], nro_fila, ruta),
                horas=_horas_de(fila.get(COL_DURACION, ""), nro_fila, ruta),
                username=(fila.get(COL_USERNAME) or "").strip(),
                cod_proyecto=codigo_de_proyecto(texto_proyecto),
                actividad=(fila.get(COL_ACTIVIDAD) or "").strip(),
                texto_proyecto=texto_proyecto,
            )
        )
    return registros
