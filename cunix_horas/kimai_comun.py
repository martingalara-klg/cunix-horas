"""Piezas compartidas por los tres lectores de export de Kimai.

Acá vive lo que no depende del formato: el `Registro` que todos devuelven, el
error que todos lanzan, y el parseo del .xlsx.

Los .xlsx de Kimai **no se leen con openpyxl**: Kimai emite el atributo
`showZeroes` donde el esquema OOXML define `showZeros`, y
`openpyxl.load_workbook()` explota con
`TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'`.
Se parsea el XML del .xlsx directamente. Esto vale para los dos formatos
.xlsx de entrada (timesheet plano y resumen mensual), así que el parseo vive
una sola vez, acá.
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


@dataclass(frozen=True)
class HojaXlsx:
    """La primera hoja de un .xlsx, ya parseada.

    `filas` es ((nro_de_fila_en_el_Excel, {letra_de_columna: texto}), ...). El
    número de fila es el del .xlsx, no el del índice en la tupla: las filas
    vacías se saltan, y el número tiene que servirle al dueño para abrir el
    archivo y mirar esa fila.

    `filas_con_dias_mergeados` son los números de fila cuyas celdas de día
    están mergeadas. Es la señal por la que el resumen mensual distingue la
    fila de cliente, igual que `escritor_excel._fila_cliente` del lado de la
    escritura.
    """

    filas: tuple[tuple[int, dict[str, str]], ...]
    filas_con_dias_mergeados: frozenset[int]

    @property
    def encabezado(self) -> dict[str, str]:
        return self.filas[0][1] if self.filas else {}


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


def tiene_codigo(texto: str) -> bool:
    """True si el texto empieza con un código entre corchetes."""
    return _CODIGO.match(texto) is not None


def letra_de_columna(referencia: str) -> str:
    """'AG12' -> 'AG'."""
    coincidencia = _SOLO_LETRAS.match(referencia)
    return coincidencia.group(0) if coincidencia else ""


def indice_de_columna(letra: str) -> int:
    """'A' -> 1, 'B' -> 2, 'AG' -> 33."""
    indice = 0
    for caracter in letra:
        indice = indice * 26 + (ord(caracter) - ord("A") + 1)
    return indice


def _cadenas_compartidas(archivo: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archivo.namelist():
        return []
    raiz = ET.fromstring(archivo.read("xl/sharedStrings.xml"))
    return [
        "".join(t.text or "" for t in si.iter(f"{NS}t"))
        for si in raiz.findall(f"{NS}si")
    ]


def _filas_de(
    raiz: ET.Element, compartidas: list[str]
) -> list[tuple[int, dict[str, str]]]:
    filas: list[tuple[int, dict[str, str]]] = []
    for elemento in raiz.iter(f"{NS}row"):
        fila: dict[str, str] = {}
        for celda in elemento:
            columna = letra_de_columna(celda.get("r", ""))
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


def _filas_mergeadas_desde(
    raiz: ET.Element, primera_columna_de_dia: str
) -> frozenset[int]:
    """Filas cuyo merge arranca en la primera columna de día y no cambia de fila."""
    desde = indice_de_columna(primera_columna_de_dia)
    filas: set[int] = set()
    for merge in raiz.iter(f"{NS}mergeCell"):
        referencia = merge.get("ref", "")
        if ":" not in referencia:
            continue
        inicio, fin = referencia.split(":", 1)
        letra_inicio = letra_de_columna(inicio)
        if not letra_inicio or indice_de_columna(letra_inicio) != desde:
            continue
        digitos_inicio = inicio[len(letra_inicio):]
        digitos_fin = fin[len(letra_de_columna(fin)):]
        if digitos_inicio.isdigit() and digitos_inicio == digitos_fin:
            filas.add(int(digitos_inicio))
    return frozenset(filas)


def leer_hoja(ruta: Path, primera_columna_de_dia: str = "C") -> HojaXlsx:
    """Parsea la primera hoja de un .xlsx de Kimai."""
    if not zipfile.is_zipfile(ruta):
        raise ErrorLectura(f"{ruta.name} no es un archivo .xlsx válido")

    with zipfile.ZipFile(ruta) as archivo:
        hojas = [n for n in archivo.namelist() if n.startswith("xl/worksheets/sheet")]
        if not hojas:
            raise ErrorLectura("El archivo no tiene ninguna hoja de cálculo")
        compartidas = _cadenas_compartidas(archivo)
        raiz = ET.fromstring(archivo.read(sorted(hojas)[0]))

    return HojaXlsx(
        filas=tuple(_filas_de(raiz, compartidas)),
        filas_con_dias_mergeados=_filas_mergeadas_desde(raiz, primera_columna_de_dia),
    )
