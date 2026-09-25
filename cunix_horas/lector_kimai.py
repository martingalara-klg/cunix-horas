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
    # Al final para no romper las construcciones posicionales existentes.
    texto_proyecto: str = ""


def serial_a_fecha(serial: str | float) -> date:
    """Convierte un serial de fecha de Excel a date. 46265.708333333 -> 2026-08-31."""
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


def _filas(archivo: zipfile.ZipFile) -> list[tuple[int, dict[str, str]]]:
    """Devuelve cada fila como (nro_de_fila_en_el_Excel, {letra: valor}).

    El número de fila es el del .xlsx, no el del índice en la lista: las filas
    vacías se saltan, y el número tiene que servirle al dueño para abrir el
    archivo y mirar esa fila.
    """
    hojas = [n for n in archivo.namelist() if n.startswith("xl/worksheets/sheet")]
    if not hojas:
        raise ErrorLectura("El archivo no tiene ninguna hoja de cálculo")
    compartidas = _cadenas_compartidas(archivo)
    raiz = ET.fromstring(archivo.read(sorted(hojas)[0]))

    filas: list[tuple[int, dict[str, str]]] = []
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
            nro = int(elemento.get("r") or len(filas) + 1)
            filas.append((nro, fila))
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


def leer(ruta: Path) -> list[Registro]:
    """Lee un export de Kimai y devuelve sus registros de tiempo."""
    if not zipfile.is_zipfile(ruta):
        raise ErrorLectura(f"{ruta.name} no es un archivo .xlsx válido")

    with zipfile.ZipFile(ruta) as archivo:
        filas = _filas(archivo)

    if not filas:
        raise ErrorLectura(f"{ruta.name} está vacío")

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
