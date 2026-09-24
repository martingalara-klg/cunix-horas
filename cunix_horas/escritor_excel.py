"""Generación del Excel con el formato que recibe el partner.

La plantilla es la fuente de estilos, no un archivo que se muta: la cantidad
de filas (proyectos por dev) y de columnas (días del mes) es variable, e
insertar filas en un .xlsx existente rompe estilos de forma silenciosa.
"""
from __future__ import annotations

from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from cunix_horas.agregador import Reporte

# Nunca strftime("%b"): depende del locale de la máquina.
MESES_ABREVIADOS = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}

# Celdas de la plantilla de las que se toma el estilo de cada tipo de fila.
MODELO_ENCABEZADO = "A1"
MODELO_CLIENTE = "A2"
MODELO_PROYECTO = "A3"
MODELO_ACTIVIDAD = "A4"
MODELO_TOTAL = "A8"

PRIMERA_COLUMNA_DE_DIA = 3  # C


def nombre_de_archivo(reporte: Reporte) -> str:
    """'Oct Dodera.xlsx'."""
    return f"{MESES_ABREVIADOS[reporte.mes]} {reporte.nombre_archivo}.xlsx"


def _columna_del_dia(dia: int) -> int:
    return PRIMERA_COLUMNA_DE_DIA + dia - 1


def _estilos_de(plantilla: Path) -> dict:
    """Lee de la plantilla el estilo de cada tipo de fila, el título y los anchos."""
    libro = openpyxl.load_workbook(plantilla)
    hoja = libro.active
    modelos = {
        "encabezado": MODELO_ENCABEZADO,
        "cliente": MODELO_CLIENTE,
        "proyecto": MODELO_PROYECTO,
        "actividad": MODELO_ACTIVIDAD,
        "total": MODELO_TOTAL,
    }
    estilos: dict = {
        nombre: {
            "font": copy(hoja[coord].font),
            "alignment": copy(hoja[coord].alignment),
            "number_format": hoja[coord].number_format,
        }
        for nombre, coord in modelos.items()
    }
    estilos["_titulo"] = hoja.title
    estilos["_anchos"] = {
        letra: dim.width
        for letra, dim in hoja.column_dimensions.items()
        if dim.width is not None
    }
    libro.close()
    return estilos


def _aplicar(celda: Cell, estilo: dict) -> None:
    celda.font = copy(estilo["font"])
    celda.alignment = copy(estilo["alignment"])
    celda.number_format = estilo["number_format"]


def _escribir_fila(
    hoja: Worksheet, nro_fila: int, valores: dict[int, object], estilo: dict, ancho: int
) -> None:
    for columna in range(1, ancho + 1):
        celda = hoja.cell(row=nro_fila, column=columna)
        if columna in valores:
            celda.value = valores[columna]
        _aplicar(celda, estilo)


def escribir(reporte: Reporte, plantilla: Path, destino: Path) -> Path:
    """Escribe el Excel del reporte en `destino` y devuelve esa ruta."""
    estilos = _estilos_de(plantilla)
    dias = reporte.dias_del_mes
    ancho = _columna_del_dia(dias)
    ultima_letra = get_column_letter(ancho)

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = estilos["_titulo"]

    # Fila 1: nombre del dev, 'Total', y un encabezado por día como TEXTO.
    encabezado: dict[int, object] = {1: reporte.nombre_dev, 2: "Total"}
    for dia in range(1, dias + 1):
        encabezado[_columna_del_dia(dia)] = f"{reporte.mes}/{dia}/{reporte.anio}"
    _escribir_fila(hoja, 1, encabezado, estilos["encabezado"], ancho)

    nro_fila = 2
    cliente_actual: str | None = None
    proyecto_actual: tuple[str, str] | None = None

    for fila in reporte.filas:
        if fila.cliente != cliente_actual:
            total_cliente = sum(
                f.total for f in reporte.filas if f.cliente == fila.cliente
            )
            _escribir_fila(
                hoja,
                nro_fila,
                {1: fila.cliente, 2: round(total_cliente, 2)},
                estilos["cliente"],
                ancho,
            )
            hoja.merge_cells(
                f"{get_column_letter(PRIMERA_COLUMNA_DE_DIA)}{nro_fila}:"
                f"{ultima_letra}{nro_fila}"
            )
            cliente_actual = fila.cliente
            proyecto_actual = None
            nro_fila += 1

        if (fila.cliente, fila.proyecto) != proyecto_actual:
            hermanas = [
                f
                for f in reporte.filas
                if (f.cliente, f.proyecto) == (fila.cliente, fila.proyecto)
            ]
            valores: dict[int, object] = {
                1: fila.proyecto,
                2: round(sum(f.total for f in hermanas), 2),
            }
            for dia in range(1, dias + 1):
                horas = sum(f.horas_por_dia.get(dia, 0.0) for f in hermanas)
                if horas:
                    valores[_columna_del_dia(dia)] = round(horas, 2)
            _escribir_fila(hoja, nro_fila, valores, estilos["proyecto"], ancho)
            proyecto_actual = (fila.cliente, fila.proyecto)
            nro_fila += 1

        valores = {1: fila.actividad, 2: round(fila.total, 2)}
        for dia, horas in fila.horas_por_dia.items():
            valores[_columna_del_dia(dia)] = round(horas, 2)
        _escribir_fila(hoja, nro_fila, valores, estilos["actividad"], ancho)
        nro_fila += 1

    # Fila Total: con 0.0 explícito en los días sin horas, como en la plantilla.
    totales: dict[int, object] = {1: "Total", 2: round(reporte.total, 2)}
    for dia in range(1, dias + 1):
        totales[_columna_del_dia(dia)] = round(reporte.total_del_dia(dia), 2)
    _escribir_fila(hoja, nro_fila, totales, estilos["total"], ancho)

    for letra, ancho_columna in estilos["_anchos"].items():
        hoja.column_dimensions[letra].width = ancho_columna

    destino.parent.mkdir(parents=True, exist_ok=True)
    libro.save(destino)
    return destino
