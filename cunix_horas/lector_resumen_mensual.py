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
3. **El separador decimal**, que puede ser coma o punto. Si el archivo mezcla
   los dos se falla: `"2.00"` leído con coma da 200.

Y hay una red de seguridad que los otros formatos no permiten: el archivo
declara sus propios totales en la fila `Total`, el general en la columna B y
uno por día en cada columna. Se comparan los dos: el general ataja las horas
que se perdieron y los de día atajan las que se leyeron en la columna
equivocada. Un error leyendo esta grilla no puede terminar en el Excel del
cliente.
"""
from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path

from cunix_horas.kimai_comun import (
    PRIMERA_COL_DE_DIA,
    ErrorLectura,
    HojaXlsx,
    Registro,
    codigo_de_proyecto,
    indice_de_columna,
    letra_desde_indice,
    tiene_codigo,
)

COL_ETIQUETA = "A"
COL_TOTAL = "B"
TEXTO_FILA_TOTAL = "Total"

_FECHA = re.compile(r"^\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})\s*$")
_DECIMAL_CON_COMA = re.compile(r",\d+\s*$")
_DECIMAL_CON_PUNTO = re.compile(r"\.\d+\s*$")

# Tolerancia de la comparación contra los totales que declara el archivo.
#
# El archivo redondea a 2 decimales cada celda y también sus totales, así que
# lo leído nunca da exactamente lo declarado. El error de cada celda cae en
# ±0.005, pero sumar n celdas NO suma n veces ese máximo: los errores son
# independientes y se cancelan entre sí. El desvío de la suma crece como √n,
# no como n (desvío estándar de un error uniforme en ±0.005: 0.005/√3 ≈
# 0.0029; el de la suma de n, 0.0029·√n).
#
# Criterio: 0.005 (el redondeo del total declarado, que es una sola celda) más
# 0.015·√n, o sea algo más de 5 desvíos estándar. Un archivo legítimo la
# supera con probabilidad despreciable, y una hora perdida no.
#
# Por qué no la cota del peor caso (0.005·n): crece lineal y se vuelve inútil
# justo donde más hace falta. En un resumen de 93 celdas daba 0.47 h, así que
# borrar una celda de 0.25 h no disparaba nada. Con ésta, 93 celdas dan 0.15 h
# y esa celda se detecta.
TOLERANCIA_DEL_TOTAL_DECLARADO = 0.005
TOLERANCIA_POR_RAIZ_DE_CELDAS = 0.015


def _tolerancia(celdas: int) -> float:
    """Cuánto puede apartarse legítimamente la suma de `celdas` de lo declarado."""
    return TOLERANCIA_DEL_TOTAL_DECLARADO + TOLERANCIA_POR_RAIZ_DE_CELDAS * math.sqrt(
        celdas
    )


def _columnas_de_dia(encabezado: dict[str, str]) -> dict[str, str]:
    """{letra: texto del encabezado} de la primera columna de día en adelante."""
    desde = indice_de_columna(PRIMERA_COL_DE_DIA)
    return {
        letra: texto
        for letra, texto in encabezado.items()
        if indice_de_columna(letra) >= desde and texto.strip()
    }


def _verificar_columnas_de_dia_contiguas(
    encabezado: dict[str, str],
    filas: tuple[tuple[int, dict[str, str]], ...],
    ruta: Path,
) -> None:
    """Un encabezado de día en blanco borraría esa columna y todas sus horas.

    Las columnas de día del resumen son el mes entero, así que van seguidas
    desde la primera hasta la última con datos. Si en el medio hay una sin
    encabezado, sus horas desaparecerían sin error propio: se falla.
    """
    desde = indice_de_columna(PRIMERA_COL_DE_DIA)
    con_encabezado = {indice_de_columna(letra) for letra in _columnas_de_dia(encabezado)}
    ocupadas = con_encabezado | {
        indice_de_columna(letra)
        for _, fila in filas
        for letra, texto in fila.items()
        if indice_de_columna(letra) >= desde and texto.strip()
    }
    if not ocupadas:
        return  # Sin ninguna columna de día: lo dice `_fechas_de_encabezado`.

    huecos = sorted(set(range(desde, max(ocupadas) + 1)) - con_encabezado)
    if huecos:
        letras = ", ".join(letra_desde_indice(i) for i in huecos)
        ultima = letra_desde_indice(max(ocupadas))
        raise ErrorLectura(
            f"{ruta.name}, fila 1: hay columnas sin encabezado de día entre "
            f"{PRIMERA_COL_DE_DIA} y {ultima}: {letras}.\n"
            f"  Sin encabezado no se sabe a qué día corresponden, y las horas "
            f"de esas columnas quedarían afuera del Excel del cliente sin "
            f"avisar.\n"
            f"  Exportá de nuevo el resumen mensual desde Kimai con el mes "
            f"completo, sin editar el archivo a mano."
        )


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


def _detectar_separador(valores: list[str], ruta: Path) -> str:
    """Separador decimal del archivo, deducido de sus propios números.

    Si el archivo mezcla los dos separadores no se elige uno: leer `"2.00"`
    como si el punto fuera de miles da 200, y el dueño terminaría viendo un
    mensaje sobre el total que no explica el problema real.
    """
    con_coma = [v.strip() for v in valores if _DECIMAL_CON_COMA.search(v)]
    con_punto = [v.strip() for v in valores if _DECIMAL_CON_PUNTO.search(v)]
    if con_coma and con_punto:
        raise ErrorLectura(
            f"{ruta.name}: las horas del archivo vienen con los dos "
            f"separadores decimales mezclados, así que no se sabe cuál usar.\n"
            f"  Con coma: {', '.join(sorted(set(con_coma))[:4])}\n"
            f"  Con punto: {', '.join(sorted(set(con_punto))[:4])}\n"
            f"  Leer una de las dos formas con el separador equivocado "
            f"multiplicaría esas horas por cien.\n"
            f"  Exportá de nuevo el resumen mensual desde Kimai sin editar el "
            f"archivo a mano."
        )
    if con_coma:
        return ","
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
    if abs(leido - declarado) > _tolerancia(len(registros)):
        raise ErrorLectura(
            f"{ruta.name}: las horas leídas no cierran con el total que "
            f"declara el archivo.\n"
            f"  Total declarado en {COL_TOTAL}{nro_fila}: {declarado:.2f} h\n"
            f"  Total leído de las filas de actividad: {leido:.2f} h "
            f"({len(registros)} celdas con horas)\n"
            f"  No se genera nada con este archivo: el Excel del cliente "
            f"saldría con horas que no son las del export."
        )


def _verificar_totales_por_dia(
    leido_por_columna: dict[str, float],
    celdas_por_columna: dict[str, int],
    fechas: dict[str, date],
    fila_total: dict[str, str],
    separador: str,
    ruta: Path,
    nro_fila: int,
) -> None:
    """El total general no ve un corrimiento de columnas; el de cada día sí.

    Si las horas se leyeran una columna corrida, la suma seguiría dando lo
    mismo y quedarían imputadas al día equivocado. El archivo declara también
    el total de cada día en su fila `Total`: se comparan uno por uno.
    """
    for letra in fechas:
        declarado_texto = fila_total.get(letra, "").strip()
        if not declarado_texto:
            continue  # El archivo no declara total para ese día: nada que comparar.
        declarado = _numero(
            declarado_texto,
            separador,
            f"{ruta.name}, fila {nro_fila}, columna {letra}",
        )
        leido = leido_por_columna.get(letra, 0.0)
        if abs(leido - declarado) > _tolerancia(celdas_por_columna.get(letra, 0)):
            raise ErrorLectura(
                f"{ruta.name}: las horas leídas del día "
                f"{fechas[letra].strftime('%d/%m/%Y')} (columna {letra}) no "
                f"cierran con lo que declara el archivo.\n"
                f"  Total declarado en {letra}{nro_fila}: {declarado:.2f} h\n"
                f"  Total leído de las filas de actividad: {leido:.2f} h\n"
                f"  Las horas están cayendo en un día que no es el suyo. No "
                f"se genera nada con este archivo: el Excel del cliente "
                f"saldría con las horas en las fechas equivocadas."
            )


def _error_de_actividad_huerfana(
    etiqueta: str,
    nro_fila: int,
    fila_sin_horas: tuple[int, str] | None,
    ruta: Path,
) -> ErrorLectura:
    """La actividad no tiene proyecto arriba. Casi siempre la culpa es de arriba.

    Una fila de proyecto que no trae horas por día se clasifica como cliente,
    y recién falla la actividad de abajo. El mensaje tiene que nombrar la fila
    que está realmente mal, que es la de arriba.
    """
    if fila_sin_horas is not None:
        nro_sospechosa, texto = fila_sin_horas
        return ErrorLectura(
            f"{ruta.name}, fila {nro_sospechosa}: {texto!r} trae código de "
            f"proyecto pero ninguna hora por día, así que se tomó como fila "
            f"de cliente y la actividad {etiqueta!r} de la fila {nro_fila} "
            f"quedó sin proyecto al que imputarle las horas.\n"
            f"  Si la fila {nro_sospechosa} es la del proyecto, sus horas por "
            f"día se perdieron en el export.\n"
            f"  Exportá de nuevo el resumen mensual desde Kimai sin editar el "
            f"archivo a mano."
        )
    return ErrorLectura(
        f"{ruta.name}, fila {nro_fila}: la actividad {etiqueta!r} no viene "
        f"debajo de ninguna fila de proyecto, así que no se sabe a qué "
        f"proyecto imputarle las horas."
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

    _verificar_columnas_de_dia_contiguas(encabezado, filas, ruta)
    fechas = _fechas_de_encabezado(encabezado, ruta)
    nro_total, fila_total = _fila_total(filas, ruta)
    separador = _detectar_separador(_valores_numericos(filas, set(fechas)), ruta)

    registros: list[Registro] = []
    leido_por_columna: dict[str, float] = {}
    celdas_por_columna: dict[str, int] = {}
    texto_proyecto: str | None = None
    fila_sin_horas: tuple[int, str] | None = None
    for nro_fila, fila in filas[1:-1]:
        etiqueta = fila.get(COL_ETIQUETA, "").strip()
        celdas = {
            letra: texto
            for letra, texto in fila.items()
            if letra in fechas and texto.strip()
        }

        mergeada = nro_fila in hoja.filas_con_dias_mergeados
        es_cliente = mergeada or (not celdas and tiene_codigo(etiqueta))
        if es_cliente:
            if not mergeada:
                # Se clasificó como cliente por descarte, no por estructura:
                # si abajo hay una actividad huérfana, el problema es ésta.
                fila_sin_horas = (nro_fila, etiqueta)
            texto_proyecto = None
            continue

        if tiene_codigo(etiqueta):
            # Fila de proyecto: subtotales. Las horas salen de sus actividades.
            texto_proyecto = etiqueta
            continue

        if texto_proyecto is None:
            raise _error_de_actividad_huerfana(
                etiqueta, nro_fila, fila_sin_horas, ruta
            )

        for letra, texto in celdas.items():
            horas = _numero(
                texto, separador, f"{ruta.name}, fila {nro_fila}, columna {letra}"
            )
            if horas == 0:
                continue
            leido_por_columna[letra] = leido_por_columna.get(letra, 0.0) + horas
            celdas_por_columna[letra] = celdas_por_columna.get(letra, 0) + 1
            registros.append(
                Registro(
                    fecha=fechas[letra],
                    horas=horas,
                    username=nombre_dev,
                    cod_proyecto=codigo_de_proyecto(
                        texto_proyecto, f"{ruta.name}, fila {nro_fila}"
                    ),
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
    _verificar_totales_por_dia(
        leido_por_columna,
        celdas_por_columna,
        fechas,
        fila_total,
        separador,
        ruta,
        nro_total,
    )
    return registros
