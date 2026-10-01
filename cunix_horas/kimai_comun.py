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
from datetime import date, time, timedelta
from pathlib import Path

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
EPOCA_EXCEL = date(1899, 12, 30)

# Primera columna de día del resumen mensual: A es la etiqueta, B el total.
# Vive acá y no en el lector porque `leer_hoja` la necesita para detectar las
# filas mergeadas: si hubiera dos copias y alguien cambiara una, la detección
# de la fila de cliente dejaría de coincidir con la lectura de la grilla.
PRIMERA_COL_DE_DIA = "C"

# De qué reporte de Kimai salió un registro. El detalle trae una fila por
# registro con todas sus columnas; el resumen mensual trae la grilla de días y
# hay que completarle el usuario de Kimai desde el mapeo.
ORIGEN_DETALLE = "reporte de detalle"
ORIGEN_RESUMEN_MENSUAL = "resumen mensual"

_CODIGO = re.compile(r"^\s*\[([^\]]+)\]")
_HORA = re.compile(r"^\s*(\d{1,2}):([0-5]\d)(?::([0-5]\d))?\s*$")
_SOLO_LETRAS = re.compile(r"[A-Z]+")


class ErrorLectura(Exception):
    """El archivo de input no se pudo leer o no tiene el formato esperado."""


@dataclass(frozen=True)
class Registro:
    """Un registro de tiempo individual de Kimai.

    Los primeros cinco campos son los que necesitaba el Excel pivoteado por
    desarrollador. Los de abajo son el resto de lo que trae el reporte de
    detalle de Kimai.

    `email` y `numero_proyecto` se conservan tal como vienen del export, pero
    hoy ningún anexo los escribe: el Anexo II-A no tiene esas columnas. Se
    leen porque están en el archivo; nunca se exigen ni frenan nada.

    Todos los agregados van **al final y con valor por defecto**, para no
    romper las construcciones posicionales que ya existen.
    """

    fecha: date
    horas: float
    username: str
    cod_proyecto: str
    actividad: str
    # Al final para no romper las construcciones posicionales existentes.
    texto_proyecto: str = ""
    # Hora de inicio (columna `From` de Kimai). El entregable la combina con
    # `fecha` en una sola celda de fecha y hora.
    hora_inicio: time | None = None
    # Nombre para mostrar ('Matias Zalazar'), distinto de `username`
    # ('mzalazar'). Los dos viajan: el username es la clave del mapeo.
    nombre: str = ""
    email: str = ""
    # Vacía en la mayoría de los registros: es un campo opcional de Kimai.
    descripcion: str = ""
    # El `Project number` de Kimai, que **no** es el código entre corchetes:
    # el proyecto '[AD2690002] ...' tiene número de proyecto '210'. El código
    # es la clave del mapeo; el número va tal cual al entregable.
    numero_proyecto: str = ""
    # Texto crudo de la columna `Customer`, por el mismo motivo por el que se
    # conserva el del proyecto: de ahí sale el nombre a mostrar cuando el
    # proyecto no está en el mapeo.
    texto_cliente: str = ""
    # De qué reporte de Kimai salió este registro. No es un dato de Kimai: es
    # lo que permite saber qué columnas trae el export y cuáles hay que
    # completar desde `config/mapeo.yaml`. Los registros del detalle nunca se
    # completan: sus valores son los de Kimai, aunque el mapeo diga otra cosa.
    origen: str = ORIGEN_DETALLE


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


def codigo_de_proyecto(texto: str, ubicacion: str = "") -> str:
    """Extrae el código entre corchetes del campo Project de Kimai.

    '[CO2610170] Aduana-Subastas | ...' -> 'CO2610170'

    `ubicacion` es el archivo y la fila de donde salió el texto. Lo aporta
    quien llama, que es el único que lo sabe: sin eso el dueño lee que hay un
    proyecto sin código pero no en qué archivo ni en qué fila mirarlo.
    """
    coincidencia = _CODIGO.match(texto)
    if coincidencia is None:
        prefijo = f"{ubicacion}: " if ubicacion else ""
        raise ErrorLectura(
            f"{prefijo}Proyecto sin código entre corchetes: {texto!r}\n"
            f"  Se esperaba que la columna de proyecto empezara con el código "
            f"entre corchetes, como '[CO2610170] Nombre del proyecto'."
        )
    return coincidencia.group(1)


def hora_de_inicio(texto: str, ubicacion: str = "") -> time | None:
    """Convierte la columna `From` de Kimai a `time`. '13:00' -> time(13, 0).

    Vacía devuelve `None`: Kimai puede no traerla y no es un dato que se
    pueda inventar. Pero si trae algo que no es una hora **se falla**, en vez
    de dejarlo pasar: en el entregable esa hora forma parte de la celda de
    fecha, y una hora mal leída movería el registro de día.

    `ubicacion` es el archivo y la fila de donde salió el texto, que sólo
    conoce quien llama.
    """
    limpio = (texto or "").strip()
    if not limpio:
        return None
    coincidencia = _HORA.match(limpio)
    if coincidencia is None or int(coincidencia.group(1)) > 23:
        prefijo = f"{ubicacion}: " if ubicacion else ""
        raise ErrorLectura(
            f"{prefijo}La hora de inicio {texto!r} no tiene formato de hora.\n"
            f"  Se esperaba la columna From como 'HH:MM', por ejemplo '13:00'.\n"
            f"  Exportá de nuevo desde Kimai sin editar el archivo a mano."
        )
    horas, minutos, segundos = coincidencia.groups()
    return time(int(horas), int(minutos), int(segundos or 0))


def letra_desde_indice(indice: int) -> str:
    """1 -> 'A', 2 -> 'B', 33 -> 'AG'. Inversa de `indice_de_columna`."""
    letras = ""
    while indice > 0:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(ord("A") + resto) + letras
    return letras


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


def leer_hoja(
    ruta: Path, primera_columna_de_dia: str = PRIMERA_COL_DE_DIA
) -> HojaXlsx:
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
