"""Punto de entrada de la lectura: detecta el formato del export y delega.

El dueño exporta desde Kimai con distintos reportes, y cada uno da un archivo
distinto. `leer()` mira el archivo y elige el lector; los tres devuelven
`list[Registro]`, así que agregador, validador y escritor no se enteran.

    .csv                    -> timesheet plano en CSV
    .xlsx con A1='Date'     -> timesheet plano en XLSX
    .xlsx con B1='Total'    -> resumen mensual
    cualquier otra cosa     -> ErrorLectura

Nunca se adivina: si la fila 1 no es ninguna de esas, se falla diciendo qué
se encontró y qué formatos se reconocen.
"""
from __future__ import annotations

from pathlib import Path

from cunix_horas.kimai_comun import (
    ErrorLectura,
    Registro,
    codigo_de_proyecto,
    leer_hoja,
    serial_a_fecha,
)
from cunix_horas.lector_resumen_mensual import leer_resumen_mensual
from cunix_horas.lector_timesheet_csv import leer_timesheet_csv
from cunix_horas.lector_timesheet_xlsx import leer_timesheet_xlsx

# Re-exportados para que el resto del programa siga importando de acá.
__all__ = [
    "ErrorLectura",
    "Registro",
    "codigo_de_proyecto",
    "leer",
    "serial_a_fecha",
]

EXTENSION_CSV = ".csv"
MARCA_TIMESHEET_XLSX = ("A", "Date")
MARCA_RESUMEN_MENSUAL = ("B", "Total")

EXTENSIONES_DE_ENTRADA = (".xlsx", EXTENSION_CSV)


def _descripcion_de_celda(encabezado: dict[str, str], columna: str) -> str:
    valor = encabezado.get(columna, "").strip()
    return f"{columna}1={valor!r}" if valor else f"{columna}1 vacía"


def _error_de_formato_desconocido(
    ruta: Path, encabezado: dict[str, str]
) -> ErrorLectura:
    return ErrorLectura(
        f"{ruta.name} no tiene el formato de ningún export de Kimai conocido.\n"
        f"  En la fila 1 se encontró: "
        f"{_descripcion_de_celda(encabezado, 'A')}, "
        f"{_descripcion_de_celda(encabezado, 'B')}.\n"
        f"  Formatos que se reconocen:\n"
        f"  - timesheet plano .xlsx: A1='Date'\n"
        f"  - resumen mensual .xlsx: B1='Total'\n"
        f"  - timesheet plano .csv: encabezados Date, Duration, User, "
        f"Project, Activity en la fila 1"
    )


def leer(ruta: Path) -> list[Registro]:
    """Lee un export de Kimai, en cualquiera de sus tres formatos."""
    if ruta.suffix.lower() == EXTENSION_CSV:
        return leer_timesheet_csv(ruta)

    hoja = leer_hoja(ruta)
    if not hoja.filas:
        raise ErrorLectura(f"{ruta.name} está vacío")

    encabezado = hoja.encabezado
    columna, texto = MARCA_TIMESHEET_XLSX
    if encabezado.get(columna, "").strip() == texto:
        return leer_timesheet_xlsx(ruta, hoja)

    columna, texto = MARCA_RESUMEN_MENSUAL
    if encabezado.get(columna, "").strip() == texto:
        return leer_resumen_mensual(ruta, hoja)

    raise _error_de_formato_desconocido(ruta, encabezado)
