"""Lectura del export .xlsx de timesheet plano de Kimai.

Fila 1 = encabezados (`A1='Date'`), una fila por registro de tiempo. Las
columnas se toman por posición porque es lo que este export viene emitiendo
desde siempre, y el encabezado se verifica antes de leer nada.
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

from cunix_horas.kimai_comun import (
    ErrorLectura,
    HojaXlsx,
    Registro,
    codigo_de_proyecto,
    serial_a_fecha,
)

COL_FECHA = "A"
COL_DURACION = "D"
COL_USERNAME = "F"
COL_PROYECTO = "J"
COL_ACTIVIDAD = "K"

ENCABEZADOS_ESPERADOS = {
    COL_FECHA: "Date",
    COL_DURACION: "Duration",
    COL_USERNAME: "User",
    COL_PROYECTO: "Project",
    COL_ACTIVIDAD: "Activity",
}


def _verificar_encabezados(encabezado: dict[str, str], ruta: Path) -> None:
    faltantes = [
        f"{col}={esperado!r}"
        for col, esperado in ENCABEZADOS_ESPERADOS.items()
        if encabezado.get(col) != esperado
    ]
    if faltantes:
        raise ErrorLectura(
            f"{ruta.name} no tiene el formato de export de Kimai. "
            f"Se esperaba en la fila 1: {', '.join(faltantes)}"
        )


def _fecha_de(valor: str, nro_fila: int, ruta: Path) -> date:
    """Fecha de una fila, o ErrorLectura en español si la celda no es una fecha."""
    try:
        return serial_a_fecha(valor)
    except (ValueError, TypeError, OverflowError):
        raise ErrorLectura(
            f"{ruta.name}, fila {nro_fila}, columna {COL_FECHA} "
            f"({ENCABEZADOS_ESPERADOS[COL_FECHA]}): {valor!r} no es una fecha "
            f"que Excel pueda interpretar.\n"
            f"  Exportá de nuevo desde Kimai sin editar el archivo a mano: la "
            f"columna {COL_FECHA} tiene que quedar con formato de fecha."
        ) from None


def _horas_de(valor: str | float, nro_fila: int, ruta: Path) -> float:
    """Duración de una fila, o ErrorLectura en español si la celda no es un número."""
    try:
        return float(valor)
    except (ValueError, TypeError):
        raise ErrorLectura(
            f"{ruta.name}, fila {nro_fila}, columna {COL_DURACION} "
            f"({ENCABEZADOS_ESPERADOS[COL_DURACION]}): {valor!r} no es una "
            f"duración numérica.\n"
            f"  Exportá de nuevo desde Kimai sin editar el archivo a mano: la "
            f"columna {COL_DURACION} tiene que quedar con formato de hora."
        ) from None


def leer_timesheet_xlsx(ruta: Path, hoja: HojaXlsx) -> list[Registro]:
    """Devuelve los registros de tiempo de un timesheet plano .xlsx."""
    filas = hoja.filas
    _verificar_encabezados(filas[0][1], ruta)

    registros: list[Registro] = []
    for nro_fila, fila in filas[1:]:
        if COL_FECHA not in fila:
            continue
        texto_proyecto = fila.get(COL_PROYECTO, "")
        registros.append(
            Registro(
                fecha=_fecha_de(fila[COL_FECHA], nro_fila, ruta),
                horas=_horas_de(fila.get(COL_DURACION, 0), nro_fila, ruta) * 24,
                username=fila.get(COL_USERNAME, ""),
                cod_proyecto=codigo_de_proyecto(texto_proyecto),
                actividad=fila.get(COL_ACTIVIDAD, ""),
                texto_proyecto=texto_proyecto,
            )
        )
    return registros
