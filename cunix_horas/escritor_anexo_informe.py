"""Escritura del **Anexo II**, el informe mensual en Word.

Se parte de `templates/Anexo-II-Informe-mensual-horas-KLG.docx`, la plantilla
vacía de C.UNIX, y se escribe una copia en `output/<mes>/`. **La plantilla
nunca se modifica.**

Las tres tablas que la herramienta llena se arman **clonando la fila modelo**
que dejó la plantilla: así cada fila nueva hereda sus bordes, su tipografía y
su alineación, y el documento sigue siendo el de C.UNIX.

| Qué | Cómo sale |
|---|---|
| Encabezado, período | automático |
| Contrato de fecha, fecha de emisión | configurables; si no están, queda el marcador |
| Tabla 0, plazo de entrega | configurable; si no está, queda `[●]` |
| Tabla 1, horas e importe por proyecto | automático |
| Tabla 2, horas por persona | automático |
| Tabla 3, principales trabajos | **borrador mecánico**, para que el dueño lo edite |
| Tabla 4, observaciones | se deja con `[●]`: la escribe el dueño |
| Tabla 5, firmas | intacta |

El párrafo «EJEMPLO: agosto 2026…» con el que C.UNIX mandó el formato se
**quita del documento generado**: es la nota de quien envió la plantilla, y no
corresponde en un informe que entrega KLG. En la plantilla queda.
"""
from __future__ import annotations

import copy
from dataclasses import dataclass
from pathlib import Path

from docx import Document

from cunix_horas.anexos import (
    MARCA_PENDIENTE,
    Periodo,
    formato_horas,
    formato_importe,
)

# Índice de cada tabla dentro del documento, en el orden en que aparecen.
TABLA_CONDICIONES = 0
TABLA_PROYECTOS = 1
TABLA_PERSONAS = 2
TABLA_TRABAJOS = 3
TABLA_OBSERVACIONES = 4
TABLA_FIRMAS = 5

# La fila que la plantilla dejó como modelo: la 0 es el encabezado y la última
# es la de totales.
FILA_MODELO = 1

MARCADOR_PERIODO = "[Mes AAAA]"
MARCADOR_FECHA = "[DD/MM/AAAA]"
INICIO_PARRAFO_EJEMPLO = "EJEMPLO"

# Marca interna para dejar un '[DD/MM/AAAA]' sin reemplazar sin que el paso
# siguiente lo vuelva a encontrar. Se deshace antes de guardar.
_RESERVADO = "\x00"


class ErrorPlantillaInforme(Exception):
    """La plantilla del informe no tiene la forma que el escritor espera."""


@dataclass(frozen=True)
class TextosInforme:
    """Lo que el dueño configura del informe. Vacío = queda el marcador."""

    dias_habiles_entrega: str = ""
    contrato_de_fecha: str = ""
    fecha_de_emision: str = ""


def _escribir_parrafo(parrafo, texto: str) -> None:
    """Reemplaza el texto de un párrafo conservando el formato del primer run."""
    if not parrafo.runs:
        parrafo.add_run("")
    parrafo.runs[0].text = texto
    for run in parrafo.runs[1:]:
        run._element.getparent().remove(run._element)


def _escribir_celda(celda, texto: str) -> None:
    """Reemplaza el texto de una celda conservando el formato del primer run."""
    parrafo = celda.paragraphs[0]
    for extra in celda.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    _escribir_parrafo(parrafo, texto)


def _clonar_fila(tabla, indice_modelo: int, indice_destino: int):
    """Inserta una copia de una fila existente y devuelve la fila nueva."""
    nueva = copy.deepcopy(tabla.rows[indice_modelo]._tr)
    tabla.rows[indice_destino - 1]._tr.addnext(nueva)
    return tabla.rows[indice_destino]


def _preparar_filas(tabla, cantidad: int):
    """Deja la tabla con `cantidad` filas de datos, clonando la fila modelo.

    La plantilla trae encabezado + una fila modelo + la fila de totales. Para
    un mes de cinco proyectos se clonan cuatro veces esa fila modelo, de modo
    que todas las filas nuevas salgan con el formato que puso C.UNIX.
    """
    if len(tabla.rows) < 3:
        raise ErrorPlantillaInforme(
            f"La plantilla del informe tiene una tabla con "
            f"{len(tabla.rows)} fila/s y se esperaban al menos tres "
            f"(encabezado, fila modelo y total).\n"
            f"  Si C.UNIX mandó una plantilla nueva, hay que revisar el "
            f"generador contra ella antes de usarla."
        )
    for desplazamiento in range(1, cantidad):
        _clonar_fila(tabla, FILA_MODELO, FILA_MODELO + desplazamiento)
    return [tabla.rows[FILA_MODELO + i] for i in range(cantidad)]


def _reemplazar_fechas(texto: str, pendientes: list) -> str:
    """Cambia cada '[DD/MM/AAAA]' por el valor que le toca, en orden.

    El primero es la fecha del contrato y el segundo la de emisión. El que no
    esté configurado **conserva el marcador**: un informe con `[DD/MM/AAAA]` a
    la vista le dice al dueño qué le falta; uno con una fecha inventada, no.
    """
    partes = texto.split(MARCADOR_FECHA)
    salida = partes[0]
    for parte in partes[1:]:
        valor = pendientes.pop(0) if pendientes else ""
        salida += (valor or MARCADOR_FECHA) + parte
    return salida


def _escribir_encabezado(documento, periodo: Periodo, textos: TextosInforme) -> None:
    """El período, y las dos fechas si están configuradas."""
    pendientes = [textos.contrato_de_fecha, textos.fecha_de_emision]
    for parrafo in documento.paragraphs:
        if MARCADOR_PERIODO not in parrafo.text and MARCADOR_FECHA not in parrafo.text:
            continue
        for run in parrafo.runs:
            texto = run.text.replace(MARCADOR_PERIODO, periodo.texto)
            if MARCADOR_FECHA in texto:
                texto = _reemplazar_fechas(texto, pendientes)
            if texto != run.text:
                run.text = texto


def _quitar_parrafo_de_ejemplo(documento) -> None:
    """Saca la nota con la que C.UNIX mandó el formato."""
    for parrafo in list(documento.paragraphs):
        if parrafo.text.strip().startswith(INICIO_PARRAFO_EJEMPLO):
            parrafo._element.getparent().remove(parrafo._element)


def _escribir_condiciones(tabla, dias_habiles: str) -> None:
    """El plazo de entrega en días hábiles, si está configurado."""
    if not dias_habiles:
        return
    for parrafo in tabla.rows[0].cells[0].paragraphs:
        if MARCA_PENDIENTE in parrafo.text:
            _escribir_parrafo(
                parrafo, parrafo.text.replace(MARCA_PENDIENTE, dias_habiles)
            )


def _escribir_proyectos(tabla, totales) -> None:
    """Tabla 1: cliente, proyecto, horas, valor hora e importe."""
    if not totales:
        return
    filas = _preparar_filas(tabla, len(totales))
    for fila, total in zip(filas, totales):
        _escribir_celda(fila.cells[0], total.cliente)
        _escribir_celda(fila.cells[1], total.proyecto)
        _escribir_celda(fila.cells[2], formato_horas(total.horas))
        _escribir_celda(
            fila.cells[3],
            "" if total.valor_hora is None else formato_importe(total.valor_hora),
        )
        _escribir_celda(
            fila.cells[4],
            "" if total.importe is None else formato_importe(total.importe),
        )

    total_fila = tabla.rows[-1]
    _escribir_celda(total_fila.cells[2], formato_horas(sum(t.horas for t in totales)))
    # Si a algún proyecto le falta el valor hora en la hoja `Datos`, el total a
    # facturar quedaría por debajo del real. Antes que un número creíble y
    # bajo, se deja vacío: el informe de validación nombra el proyecto.
    if all(t.valor_hora is not None for t in totales):
        _escribir_celda(
            total_fila.cells[4], formato_importe(sum(t.importe for t in totales))
        )
    else:
        _escribir_celda(total_fila.cells[4], "")


def _escribir_personas(tabla, totales) -> None:
    """Tabla 2: persona, perfil, horas, días con registro y promedio."""
    if not totales:
        return
    filas = _preparar_filas(tabla, len(totales))
    for fila, total in zip(filas, totales):
        _escribir_celda(fila.cells[0], total.persona)
        _escribir_celda(fila.cells[1], total.perfil)
        _escribir_celda(fila.cells[2], formato_horas(total.horas))
        _escribir_celda(fila.cells[3], str(total.dias))
        _escribir_celda(fila.cells[4], formato_horas(total.promedio))

    _escribir_celda(
        tabla.rows[-1].cells[2], formato_horas(sum(t.horas for t in totales))
    )


def _escribir_trabajos(tabla, trabajos) -> None:
    """Tabla 3: el BORRADOR de los principales trabajos.

    `Estado al cierre` se deja con la marca `[●]` que trae la fila modelo: no
    es un dato que la herramienta pueda saber.
    """
    if not trabajos:
        return
    filas = _preparar_filas(tabla, len(trabajos))
    for fila, trabajo in zip(filas, trabajos):
        _escribir_celda(fila.cells[0], trabajo.proyecto)
        _escribir_celda(fila.cells[1], trabajo.ticket)
        _escribir_celda(fila.cells[2], trabajo.descripcion)
        _escribir_celda(fila.cells[4], formato_horas(trabajo.horas))

    _escribir_celda(
        tabla.rows[-1].cells[4], formato_horas(sum(t.horas for t in trabajos))
    )


def escribir(
    plantilla: Path,
    destino: Path,
    periodo: Periodo,
    totales_proyecto,
    totales_persona,
    trabajos,
    textos: TextosInforme = TextosInforme(),
) -> None:
    """Escribe el Anexo II sobre una copia de la plantilla."""
    documento = Document(plantilla)
    if len(documento.tables) <= TABLA_FIRMAS:
        raise ErrorPlantillaInforme(
            f"La plantilla del informe tiene {len(documento.tables)} tablas y "
            f"se esperaban {TABLA_FIRMAS + 1}.\n"
            f"  Si C.UNIX mandó una plantilla nueva, hay que revisar el "
            f"generador contra ella antes de usarla."
        )

    _escribir_encabezado(documento, periodo, textos)
    _quitar_parrafo_de_ejemplo(documento)
    _escribir_condiciones(
        documento.tables[TABLA_CONDICIONES], textos.dias_habiles_entrega
    )
    _escribir_proyectos(documento.tables[TABLA_PROYECTOS], totales_proyecto)
    _escribir_personas(documento.tables[TABLA_PERSONAS], totales_persona)
    _escribir_trabajos(documento.tables[TABLA_TRABAJOS], trabajos)
    # Tabla 4 (observaciones) y tabla 5 (firmas): no se tocan a propósito.

    documento.save(destino)
