"""Lectura del export .xlsx de timesheet plano de Kimai.

Fila 1 = encabezados (`A1='Date'`), una fila por registro de tiempo. Las
columnas se toman por posición porque es lo que este export viene emitiendo
desde siempre, y el encabezado se verifica antes de leer nada.

Una fila sin fecha pero con datos no se saltea: son horas que no llegarían al
Excel del cliente y nadie se enteraría. Se falla nombrando el archivo y la
fila. Una fila completamente vacía sí se saltea, que no pierde nada.

Además de las cinco columnas mínimas se leen también `From` (B), `Name` (E),
`E-mail` (G), `Customer` (I), `Description` (L) y `Project number` (R). Ojo
con las dos últimas: `Description` viene vacía en la mayoría de los registros
y eso es válido, y `Project number` **no** es el código entre corchetes de
`Project` (el proyecto `[AD2690002] ...` tiene número `210`). El del corchete
es la clave del mapeo; el número y el mail se conservan porque el export los
trae, pero hoy ningún anexo los escribe.
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

from cunix_horas.kimai_comun import (
    ErrorLectura,
    HojaXlsx,
    Registro,
    codigo_de_proyecto,
    hora_de_inicio,
    serial_a_fecha,
)

COL_FECHA = "A"
COL_HORA_INICIO = "B"
COL_DURACION = "D"
COL_NOMBRE = "E"
COL_USERNAME = "F"
COL_EMAIL = "G"
COL_CLIENTE = "I"
COL_PROYECTO = "J"
COL_ACTIVIDAD = "K"
COL_DESCRIPCION = "L"
COL_NUMERO_PROYECTO = "R"

# Las que tienen que estar sí o sí: sin ellas no hay ni horas ni a quién
# imputárselas, y el archivo no se lee.
ENCABEZADOS_ESPERADOS = {
    COL_FECHA: "Date",
    COL_DURACION: "Duration",
    COL_USERNAME: "User",
    COL_PROYECTO: "Project",
    COL_ACTIVIDAD: "Activity",
}

# Las demás columnas del reporte de detalle. Como este lector toma las
# columnas por posición, una versión de Kimai que las reordene metería el
# dato equivocado en cada una. Por eso se verifican igual, pero **sólo si la
# columna existe en la fila 1**: un archivo que directamente no las trae se
# lee con esos campos vacíos, en vez de fallar entero.
ENCABEZADOS_DEL_DETALLE = {
    COL_HORA_INICIO: "From",
    COL_NOMBRE: "Name",
    COL_EMAIL: "E-mail",
    COL_CLIENTE: "Customer",
    COL_DESCRIPCION: "Description",
    COL_NUMERO_PROYECTO: "Project number",
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
    _verificar_encabezados_del_detalle(encabezado, ruta)


def _verificar_encabezados_del_detalle(
    encabezado: dict[str, str], ruta: Path
) -> None:
    """Las columnas del detalle, si están, tienen que estar donde van."""
    corridas = [
        f"{col}={encabezado[col]!r} (se esperaba {esperado!r})"
        for col, esperado in ENCABEZADOS_DEL_DETALLE.items()
        if col in encabezado and encabezado[col].strip() != esperado
    ]
    if corridas:
        raise ErrorLectura(
            f"{ruta.name}: las columnas de la fila 1 no están donde se "
            f"esperaba: {', '.join(corridas)}.\n"
            f"  Leerlas igual pondría el dato equivocado en cada columna de "
            f"los anexos que recibe C.UNIX.\n"
            f"  Exportá de nuevo desde Kimai con el reporte de detalle, sin "
            f"agregar ni mover columnas a mano."
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


def _error_de_fila_sin_fecha(
    fila: dict[str, str], nro_fila: int, ruta: Path
) -> ErrorLectura:
    """Una fila con datos pero sin fecha son horas que no se facturarían."""
    duracion = str(fila.get(COL_DURACION, "")).strip()
    detalle = (
        f" (trae {duracion!r} en la columna {COL_DURACION})" if duracion else ""
    )
    return ErrorLectura(
        f"{ruta.name}, fila {nro_fila}: la columna {COL_FECHA} "
        f"({ENCABEZADOS_ESPERADOS[COL_FECHA]}) está vacía pero la fila trae "
        f"datos{detalle}.\n"
        f"  Saltearla dejaría esas horas afuera del Excel del cliente sin "
        f"avisar, así que no se genera nada con este archivo.\n"
        f"  Completá la fecha en Kimai y exportá de nuevo, o borrá la fila "
        f"entera si no corresponde."
    )


def leer_timesheet_xlsx(ruta: Path, hoja: HojaXlsx) -> list[Registro]:
    """Devuelve los registros de tiempo de un timesheet plano .xlsx."""
    filas = hoja.filas
    _verificar_encabezados(filas[0][1], ruta)

    registros: list[Registro] = []
    for nro_fila, fila in filas[1:]:
        if not str(fila.get(COL_FECHA, "")).strip():
            if any(str(texto).strip() for texto in fila.values()):
                raise _error_de_fila_sin_fecha(fila, nro_fila, ruta)
            continue  # Fila completamente vacía: no hay nada que perder.
        texto_proyecto = fila.get(COL_PROYECTO, "")
        registros.append(
            Registro(
                fecha=_fecha_de(fila[COL_FECHA], nro_fila, ruta),
                horas=_horas_de(fila.get(COL_DURACION, 0), nro_fila, ruta) * 24,
                username=fila.get(COL_USERNAME, ""),
                cod_proyecto=codigo_de_proyecto(
                    texto_proyecto, f"{ruta.name}, fila {nro_fila}, columna "
                    f"{COL_PROYECTO} ({ENCABEZADOS_ESPERADOS[COL_PROYECTO]})"
                ),
                actividad=fila.get(COL_ACTIVIDAD, ""),
                texto_proyecto=texto_proyecto,
                hora_inicio=hora_de_inicio(
                    fila.get(COL_HORA_INICIO, ""),
                    f"{ruta.name}, fila {nro_fila}, columna {COL_HORA_INICIO} "
                    f"({ENCABEZADOS_DEL_DETALLE[COL_HORA_INICIO]})",
                ),
                nombre=fila.get(COL_NOMBRE, "").strip(),
                email=fila.get(COL_EMAIL, "").strip(),
                descripcion=fila.get(COL_DESCRIPCION, "").strip(),
                numero_proyecto=fila.get(COL_NUMERO_PROYECTO, "").strip(),
                texto_cliente=fila.get(COL_CLIENTE, "").strip(),
            )
        )
    return registros
