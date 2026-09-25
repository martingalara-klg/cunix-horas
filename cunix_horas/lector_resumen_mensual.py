"""Lectura del export .xlsx de resumen mensual de Kimai.

Tiene la misma forma que el Excel de salida: fila 1 con el nombre del dev,
`Total` y un encabezado por día; después una fila por cliente, una por
proyecto y una por actividad; y una fila `Total` al final.

Tres cosas hay que resolver acá, y ninguna se asume: se detectan por archivo.

1. **Qué es cada fila.** Igual que en `escritor_excel`, se distinguen por
   estructura: la de cliente tiene las celdas de día mergeadas, la de
   proyecto trae el código entre corchetes y las horas por día, y la de
   actividad viene abajo. Sólo las de actividad emiten registros: las de
   proyecto son subtotales y duplicarían las horas.
2. **El orden de la fecha** en el encabezado. Están todas las columnas del
   mes, así que uno de los dos componentes es constante y ése es el mes.
3. **El separador decimal**, que puede ser coma o punto.

Y hay una red de seguridad que los otros formatos no permiten: el archivo
declara su propio total en la columna B de la fila `Total`. Si la suma de lo
parseado no coincide, se falla. Un error leyendo esta grilla no puede
terminar en el Excel del cliente.
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from cunix_horas.kimai_comun import (
    ErrorLectura,
    HojaXlsx,
    Registro,
    codigo_de_proyecto,
    indice_de_columna,
    tiene_codigo,
)

COL_ETIQUETA = "A"
COL_TOTAL = "B"
PRIMERA_COL_DE_DIA = "C"
TEXTO_FILA_TOTAL = "Total"

_FECHA = re.compile(r"^\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})\s*$")
_DECIMAL_CON_COMA = re.compile(r",\d+\s*$")
_DECIMAL_CON_PUNTO = re.compile(r"\.\d+\s*$")

# El archivo redondea cada celda a 2 decimales y su propio total también, así
# que la suma de las celdas puede apartarse hasta medio centésimo por celda.
# Por encima de eso ya no es redondeo: es un error de lectura.
TOLERANCIA_BASE = 0.01
TOLERANCIA_POR_CELDA = 0.005


def _columnas_de_dia(encabezado: dict[str, str]) -> dict[str, str]:
    """{letra: texto del encabezado} de la primera columna de día en adelante."""
    desde = indice_de_columna(PRIMERA_COL_DE_DIA)
    return {
        letra: texto
        for letra, texto in encabezado.items()
        if indice_de_columna(letra) >= desde and texto.strip()
    }


def _fechas_de_encabezado(encabezado: dict[str, str], ruta: Path) -> dict[str, date]:
    """Fecha de cada columna de día, decidiendo el orden por el conjunto entero.

    Como están todas las columnas del mes, el componente que es constante en
    todos los encabezados es el mes. Si no se puede decidir sin ambigüedad se
    falla: adivinar mal movería las horas de día y de mes sin avisar.
    """
    crudos = _columnas_de_dia(encabezado)
    if not crudos:
        raise ErrorLectura(
            f"{ruta.name}: la fila 1 no tiene ninguna columna de día a partir "
            f"de la columna {PRIMERA_COL_DE_DIA}."
        )

    partes: dict[str, tuple[int, int, int]] = {}
    for letra, texto in crudos.items():
        coincidencia = _FECHA.match(texto)
        if coincidencia is None:
            raise ErrorLectura(
                f"{ruta.name}, fila 1, columna {letra}: {texto!r} no es un "
                f"encabezado de día con formato D/M/AAAA o M/D/AAAA."
            )
        partes[letra] = tuple(int(g) for g in coincidencia.groups())

    primeros = {p[0] for p in partes.values()}
    segundos = {p[1] for p in partes.values()}
    anios = {p[2] for p in partes.values()}
    if len(anios) != 1:
        raise ErrorLectura(
            f"{ruta.name}: los encabezados de día de la fila 1 no son todos "
            f"del mismo año (se encontraron {sorted(anios)})."
        )

    if len(segundos) == 1 and len(primeros) > 1:
        orden_dia_mes = True
    elif len(primeros) == 1 and len(segundos) > 1:
        orden_dia_mes = False
    else:
        muestra = ", ".join(list(crudos.values())[:4])
        raise ErrorLectura(
            f"{ruta.name}: no se puede saber si los encabezados de día de la "
            f"fila 1 vienen como día/mes o como mes/día. Se esperaba que uno "
            f"de los dos componentes fuera constante (el mes) y el otro "
            f"variara (el día), pero se encontró: {muestra} "
            f"({len(crudos)} columnas de día en total).\n"
            f"  Exportá de nuevo el resumen mensual desde Kimai con el mes "
            f"completo, sin editar el archivo a mano."
        )

    fechas: dict[str, date] = {}
    for letra, (primero, segundo, anio) in partes.items():
        dia, mes = (primero, segundo) if orden_dia_mes else (segundo, primero)
        try:
            fechas[letra] = date(anio, mes, dia)
        except ValueError:
            raise ErrorLectura(
                f"{ruta.name}, fila 1, columna {letra}: "
                f"{crudos[letra]!r} no es una fecha válida "
                f"(se interpretó como día {dia}, mes {mes}, año {anio})."
            ) from None
    return fechas


def _detectar_separador(valores: list[str]) -> str:
    """Separador decimal del archivo, deducido de sus propios números."""
    if any(_DECIMAL_CON_COMA.search(v) for v in valores):
        return ","
    if any(_DECIMAL_CON_PUNTO.search(v) for v in valores):
        return "."
    # Sin parte decimal en ninguna celda no hay ambigüedad posible.
    return "."


def _numero(valor: str, separador: str, ubicacion: str) -> float:
    texto = valor.strip()
    if separador == ",":
        texto = texto.replace(".", "").replace(",", ".")
    else:
        texto = texto.replace(",", "")
    try:
        return float(texto)
    except ValueError:
        raise ErrorLectura(
            f"{ubicacion}: {valor!r} no es una cantidad de horas numérica.\n"
            f"  Exportá de nuevo el resumen mensual desde Kimai sin editar el "
            f"archivo a mano."
        ) from None


def _valores_numericos(
    filas: tuple[tuple[int, dict[str, str]], ...], columnas_de_dia: set[str]
) -> list[str]:
    """Todo lo que en la grilla pretende ser un número: total y celdas de día."""
    interesantes = columnas_de_dia | {COL_TOTAL}
    return [
        texto
        for _, fila in filas
        for letra, texto in fila.items()
        if letra in interesantes and texto.strip()
    ]


def _fila_total(
    filas: tuple[tuple[int, dict[str, str]], ...], ruta: Path
) -> tuple[int, dict[str, str]]:
    nro, fila = filas[-1]
    if fila.get(COL_ETIQUETA, "").strip() != TEXTO_FILA_TOTAL:
        raise ErrorLectura(
            f"{ruta.name}: la última fila con datos (fila {nro}) no es la de "
            f"totales: se esperaba {TEXTO_FILA_TOTAL!r} en la columna "
            f"{COL_ETIQUETA} y se encontró "
            f"{fila.get(COL_ETIQUETA, '')!r}.\n"
            f"  Sin esa fila no se puede verificar que lo leído coincida con "
            f"el total que declara el archivo."
        )
    if not fila.get(COL_TOTAL, "").strip():
        raise ErrorLectura(
            f"{ruta.name}, fila {nro}: la fila {TEXTO_FILA_TOTAL!r} no "
            f"declara un total en la columna {COL_TOTAL}."
        )
    return nro, fila


def _verificar_total_declarado(
    registros: list[Registro], declarado: float, ruta: Path, nro_fila: int
) -> None:
    """Red de seguridad: lo parseado tiene que dar el total que declara el archivo."""
    leido = sum(r.horas for r in registros)
    tolerancia = TOLERANCIA_BASE + TOLERANCIA_POR_CELDA * len(registros)
    if abs(leido - declarado) > tolerancia:
        raise ErrorLectura(
            f"{ruta.name}: las horas leídas no cierran con el total que "
            f"declara el archivo.\n"
            f"  Total declarado en {COL_TOTAL}{nro_fila}: {declarado:.2f} h\n"
            f"  Total leído de las filas de actividad: {leido:.2f} h "
            f"({len(registros)} celdas con horas)\n"
            f"  No se genera nada con este archivo: el Excel del cliente "
            f"saldría con horas que no son las del export."
        )


def leer_resumen_mensual(ruta: Path, hoja: HojaXlsx) -> list[Registro]:
    """Devuelve los registros de tiempo de un resumen mensual .xlsx.

    `Registro.username` queda con lo que venga en A1, que en este formato es
    el nombre para mostrar y no el username: es el mapeo el que sabe resolver
    las dos cosas, no el lector.
    """
    filas = hoja.filas
    if len(filas) < 2:
        raise ErrorLectura(f"{ruta.name} no tiene filas de datos")

    encabezado = filas[0][1]
    nombre_dev = encabezado.get(COL_ETIQUETA, "").strip()
    if not nombre_dev:
        raise ErrorLectura(
            f"{ruta.name}: la celda {COL_ETIQUETA}1 no trae el nombre del "
            f"desarrollador."
        )

    fechas = _fechas_de_encabezado(encabezado, ruta)
    nro_total, fila_total = _fila_total(filas, ruta)
    separador = _detectar_separador(_valores_numericos(filas, set(fechas)))

    registros: list[Registro] = []
    texto_proyecto: str | None = None
    for nro_fila, fila in filas[1:-1]:
        etiqueta = fila.get(COL_ETIQUETA, "").strip()
        celdas = {
            letra: texto
            for letra, texto in fila.items()
            if letra in fechas and texto.strip()
        }

        es_cliente = nro_fila in hoja.filas_con_dias_mergeados or (
            not celdas and tiene_codigo(etiqueta)
        )
        if es_cliente:
            texto_proyecto = None
            continue

        if tiene_codigo(etiqueta):
            # Fila de proyecto: subtotales. Las horas salen de sus actividades.
            texto_proyecto = etiqueta
            continue

        if texto_proyecto is None:
            raise ErrorLectura(
                f"{ruta.name}, fila {nro_fila}: la actividad {etiqueta!r} no "
                f"viene debajo de ninguna fila de proyecto, así que no se "
                f"sabe a qué proyecto imputarle las horas."
            )

        for letra, texto in celdas.items():
            horas = _numero(
                texto, separador, f"{ruta.name}, fila {nro_fila}, columna {letra}"
            )
            if horas == 0:
                continue
            registros.append(
                Registro(
                    fecha=fechas[letra],
                    horas=horas,
                    username=nombre_dev,
                    cod_proyecto=codigo_de_proyecto(texto_proyecto),
                    actividad=etiqueta,
                    texto_proyecto=texto_proyecto,
                )
            )

    declarado = _numero(
        fila_total[COL_TOTAL],
        separador,
        f"{ruta.name}, fila {nro_total}, columna {COL_TOTAL}",
    )
    _verificar_total_declarado(registros, declarado, ruta, nro_total)
    return registros
