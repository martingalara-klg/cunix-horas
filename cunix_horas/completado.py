"""Completa, desde `config/mapeo.yaml`, lo que el resumen mensual no trae.

Kimai da dos reportes y los dos sirven, pero no traen lo mismo.

El **reporte de detalle** trae las diez columnas que el partner factura. Sus
registros pasan por acá **sin tocarse**: su usuario, su mail y su número de
proyecto son los de Kimai, aunque el mapeo declare otra cosa. El mapeo es un
respaldo para lo que la fuente no trae, no una corrección de lo que sí trae.

El **resumen mensual** es la grilla de días: trae el nombre del desarrollador,
las horas y el proyecto, y nada más. De las cinco columnas que le faltan:

- `Description` vacía y `Date` sin hora **son aceptables**. En el archivo de
  referencia del partner, 112 de sus 160 filas no tienen descripción y una
  tiene la hora en 00:00. Se emiten así y el informe lo avisa.
- `User`, `E-mail` y `Project number` **no pueden ir vacías**: las 160 filas
  del partner las tienen llenas, sin una sola excepción. Salen del mapeo, y
  si el mapeo no las tiene, **este archivo no entra**, con el bloque YAML
  listo para pegar. Nunca se inventan y nunca salen en blanco.

Frenar un archivo entero por un dato de configuración es deliberado: el
partner factura sobre estas filas, y una columna en blanco en un archivo de
setenta filas no la ve nadie. Frena sólo ese archivo: los demás
desarrolladores siguen su camino.
"""
from __future__ import annotations

from dataclasses import replace

from cunix_horas.kimai_comun import ORIGEN_RESUMEN_MENSUAL, Registro
from cunix_horas.mapeo import (
    ErrorMapeo,
    Mapeo,
    Persona,
    derivar_cliente,
    derivar_proyecto,
)

CIERRE = "  Los demás desarrolladores se procesan igual: sólo falta éste."

_SALIDA_POR_EL_DETALLE = (
    "  La otra salida es volver a exportar a esa persona desde Kimai con el "
    "reporte de detalle: ese export trae este dato en cada fila y el mapeo no "
    "interviene."
)


def _error_de_mail(persona: Persona, nombre_dev: str, archivo: str) -> ErrorMapeo:
    """Falta el `mail:` de una persona que exportó con el resumen mensual."""
    return ErrorMapeo(
        f"Falta el E-mail de {nombre_dev} para {archivo}, que es un export de "
        f"resumen mensual.\n"
        f"  Ese reporte no trae el mail del desarrollador, así que sale de "
        f"config/mapeo.yaml, y la persona '{persona.username}' no lo declara. "
        f"El partner necesita esa columna llena: sus 160 filas de referencia "
        f"la tienen.\n"
        f"  Completá en config/mapeo.yaml, bajo personas:, esa entrada:\n"
        f"  {persona.username}:\n"
        f'    nombre: "{persona.nombre}"\n'
        f'    archivo: "{persona.archivo}"\n'
        f'    mail: "AJUSTAR - el mail de Kimai de esta persona"\n'
        + _SALIDA_POR_EL_DETALLE
        + "\n"
        + CIERRE
    )


def _error_de_numero_de_proyecto(
    registro: Registro, mapeo: Mapeo, archivo: str
) -> ErrorMapeo:
    """Falta el `numero_proyecto:` de un proyecto del resumen mensual."""
    codigo = registro.cod_proyecto
    destino = mapeo.proyecto_opcional(codigo)
    if destino is not None:
        cliente, proyecto = destino.cliente, destino.proyecto
        que_hacer = "Completá en config/mapeo.yaml, bajo proyectos:, esa entrada:"
    else:
        cliente = (
            derivar_cliente(registro.texto_cliente)
            or "AJUSTAR - nombre del cliente para el partner"
        )
        proyecto = derivar_proyecto(registro.texto_proyecto) or codigo
        que_hacer = (
            "Ese proyecto tampoco está declarado. Agregá a config/mapeo.yaml, "
            "bajo proyectos:"
        )
    return ErrorMapeo(
        f"Falta el Project number del proyecto {codigo} ({proyecto}) para "
        f"{archivo}, que es un export de resumen mensual.\n"
        f"  Ese reporte no trae el número de proyecto, así que sale de "
        f"config/mapeo.yaml. El partner necesita esa columna llena: sus 160 "
        f"filas de referencia la tienen, sin una sola excepción.\n"
        f"  OJO: el Project number NO es el código entre corchetes. El "
        f"proyecto [AD2690002], por ejemplo, tiene número de proyecto 210. "
        f"Miralo en Kimai, en la ficha del proyecto.\n"
        f"  {que_hacer}\n"
        f"  {codigo}:\n"
        f'    cliente: "{cliente}"\n'
        f'    proyecto: "{proyecto}"\n'
        f'    numero_proyecto: "AJUSTAR - el Project number que Kimai le da a '
        f'este proyecto"\n'
        + _SALIDA_POR_EL_DETALLE
        + "\n"
        + CIERRE
    )


def _completar_uno(
    registro: Registro, mapeo: Mapeo, archivo: str, personas: dict[str, Persona]
) -> Registro:
    """Un registro de resumen mensual, con sus tres columnas ya resueltas."""
    # El lector del resumen mensual deja el nombre para mostrar en `username`,
    # porque ese export no tiene otra cosa. Acá se separan las dos columnas:
    # `Name` es ese nombre, y `User` es la clave de la persona en el mapeo.
    nombre_dev = (registro.nombre or registro.username).strip()

    persona = personas.get(nombre_dev)
    if persona is None:
        # Lanza ErrorMapeo con el bloque YAML si la persona no está declarada.
        persona = mapeo.resolver_persona(nombre_dev, archivo)
        personas[nombre_dev] = persona

    if not persona.mail:
        raise _error_de_mail(persona, nombre_dev, archivo)

    destino = mapeo.proyecto_opcional(registro.cod_proyecto)
    numero = destino.numero_proyecto if destino is not None else ""
    if not numero:
        raise _error_de_numero_de_proyecto(registro, mapeo, archivo)

    return replace(
        registro,
        nombre=nombre_dev,
        username=persona.username,
        email=persona.mail,
        numero_proyecto=numero,
    )


def completar_desde_mapeo(
    registros: list[Registro], mapeo: Mapeo, archivo: str
) -> list[Registro]:
    """Los mismos registros, con lo que el resumen mensual no trajo ya resuelto.

    Devuelve una lista nueva: no muta ni los registros (son `frozen`) ni la
    lista que recibe. Los registros del reporte de detalle salen tal como
    entraron.

    Lanza `ErrorMapeo` en cuanto a un registro del resumen mensual le falta el
    `User`, el `E-mail` o el `Project number` y el mapeo no los tiene. El
    mensaje nombra al desarrollador o al proyecto y trae el YAML pegable.
    """
    personas: dict[str, Persona] = {}
    return [
        _completar_uno(registro, mapeo, archivo, personas)
        if registro.origen == ORIGEN_RESUMEN_MENSUAL
        else registro
        for registro in registros
    ]
