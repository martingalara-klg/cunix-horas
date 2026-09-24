"""Lectura de los exports .xlsx de Kimai.

No se usa openpyxl: Kimai emite el atributo `showZeroes` donde el esquema
OOXML define `showZeros`, y openpyxl.load_workbook() explota con
`TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'`.
Se parsea el XML del .xlsx directamente.
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
EPOCA_EXCEL = date(1899, 12, 30)

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

_CODIGO = re.compile(r"^\s*\[([^\]]+)\]")
_SOLO_LETRAS = re.compile(r"[A-Z]+")


class ErrorLectura(Exception):
    """El archivo de input no se pudo leer o no tiene el formato esperado."""


@dataclass(frozen=True)
class Registro:
    """Un registro de tiempo individual de Kimai."""

    fecha: date
    horas: float
    username: str
    cod_proyecto: str
    actividad: str


def serial_a_fecha(serial: str | float) -> date:
    """Convierte un serial de fecha de Excel a date. 46262.5 -> 2026-08-31."""
    return EPOCA_EXCEL + timedelta(days=float(serial))


def codigo_de_proyecto(texto: str) -> str:
    """Extrae el código entre corchetes del campo Project de Kimai.

    '[CO2610170] Aduana-Subastas | ...' -> 'CO2610170'
    """
    coincidencia = _CODIGO.match(texto)
    if coincidencia is None:
        raise ErrorLectura(f"Proyecto sin código entre corchetes: {texto!r}")
    return coincidencia.group(1)


def _letra_de_columna(referencia: str) -> str:
    """'AG12' -> 'AG'."""
    coincidencia = _SOLO_LETRAS.match(referencia)
    return coincidencia.group(0) if coincidencia else ""


def _cadenas_compartidas(archivo: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archivo.namelist():
        return []
    raiz = ET.fromstring(archivo.read("xl/sharedStrings.xml"))
    return [
        "".join(t.text or "" for t in si.iter(f"{NS}t"))
        for si in raiz.findall(f"{NS}si")
    ]


def _filas(archivo: zipfile.ZipFile) -> list[dict[str, str]]:
    """Devuelve cada fila como {letra_de_columna: valor}, saltando las vacías."""
    hojas = [n for n in archivo.namelist() if n.startswith("xl/worksheets/sheet")]
    if not hojas:
        raise ErrorLectura("El archivo no tiene ninguna hoja de cálculo")
    compartidas = _cadenas_compartidas(archivo)
    raiz = ET.fromstring(archivo.read(sorted(hojas)[0]))

    filas: list[dict[str, str]] = []
    for elemento in raiz.iter(f"{NS}row"):
        fila: dict[str, str] = {}
        for celda in elemento:
            columna = _letra_de_columna(celda.get("r", ""))
            if not columna:
                continue
            valor_numerico = celda.find(f"{NS}v")
            valor_inline = celda.find(f"{NS}is")
            if valor_numerico is not None:
                if celda.get("t") == "s":
                    fila[columna] = compartidas[int(valor_numerico.text)]
                else:
                    fila[columna] = valor_numerico.text or ""
            elif valor_inline is not None:
                fila[columna] = "".join(
                    t.text or "" for t in valor_inline.iter(f"{NS}t")
                )
        if fila:
            filas.append(fila)
    return filas


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


def leer(ruta: Path) -> list[Registro]:
    """Lee un export de Kimai y devuelve sus registros de tiempo."""
    if not zipfile.is_zipfile(ruta):
        raise ErrorLectura(f"{ruta.name} no es un archivo .xlsx válido")

    with zipfile.ZipFile(ruta) as archivo:
        filas = _filas(archivo)

    if not filas:
        raise ErrorLectura(f"{ruta.name} está vacío")

    _verificar_encabezados(filas[0], ruta)

    registros: list[Registro] = []
    for fila in filas[1:]:
        if COL_FECHA not in fila:
            continue
        registros.append(
            Registro(
                fecha=serial_a_fecha(fila[COL_FECHA]),
                horas=float(fila.get(COL_DURACION, 0)) * 24,
                username=fila.get(COL_USERNAME, ""),
                cod_proyecto=codigo_de_proyecto(fila.get(COL_PROYECTO, "")),
                actividad=fila.get(COL_ACTIVIDAD, ""),
            )
        )
    return registros
