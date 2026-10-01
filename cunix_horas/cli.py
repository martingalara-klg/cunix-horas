"""Orquestación: de los exports de `input/<mes>/` a los dos anexos del mes.

    input/AAAA-MM/*.xlsx|*.csv        exports de Kimai, uno por desarrollador
    input/AAAA-MM/manual/*.xlsx       horas de quien no está en Kimai, y
                                      descripciones que alguien entrega aparte
            |
            v
    python -m cunix_horas AAAA-MM
            |
            v
    output/AAAA-MM/  Anexo II (.docx) + Anexo II-A (.xlsx) + _validacion.txt

Tres reglas que explican casi todas las decisiones de este módulo:

- **Un export que falla no frena a los demás.** Se registra el motivo, los
  anexos se generan igual con lo que sí se leyó, y salen con la marca de
  INCOMPLETO en el **nombre del archivo**, que es lo único que el dueño ve al
  adjuntarlos a un mail. Con todas las personas en un solo par de documentos,
  una que falta es invisible adentro.
- **Una fuente manual que no cierra sí frena todo.** Esas planillas mueven
  horas de una persona a otra: un reparto que no cuadra se factura mal y no se
  ve. Ahí no se genera nada.
- **El informe describe la carpeta, no la corrida.** El dueño no envía lo que
  hizo la corrida: envía lo que hay en `output/<mes>/`, así que todo archivo
  que quede ahí y no sea de esta corrida se nombra uno por uno.
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path

from cunix_horas import escritor_anexo_detalle as detalle_xlsx
from cunix_horas import escritor_anexo_informe as informe_docx
from cunix_horas import filas_anexo, fuentes_manuales
from cunix_horas.anexos import ErrorPeriodo, Periodo, nombre_incompleto
from cunix_horas.completado import completar_desde_mapeo
from cunix_horas.kimai_comun import ORIGEN_RESUMEN_MENSUAL
from cunix_horas.lector_kimai import (
    EXTENSIONES_DE_ENTRADA,
    ErrorLectura,
    Registro,
    leer,
)
from cunix_horas.mapeo import ErrorMapeo, Mapeo
from cunix_horas.validador import avisos_de_desarrollador

# Las plantillas vacías que mandó C.UNIX. Son fuente de estructura y estilos y
# nunca se modifican: se trabaja sobre copias.
PLANTILLA_DETALLE = "Anexo-II-A-Detalle-horas-KLG.xlsx"
PLANTILLA_INFORME = "Anexo-II-Informe-mensual-horas-KLG.docx"

# Marca del archivo que quedó de una corrida anterior y que esta corrida no
# pudo reemplazar. Va en el nombre para que no se confunda con el del mes.
SUFIJO_CORRIDA_ANTERIOR = " (CORRIDA ANTERIOR - NO ENVIAR)"

NOMBRE_INFORME = "_validacion.txt"

# Si _validacion.txt no se puede escribir, el que quedó en disco es el de otra
# corrida y describe otra cosa. El informe de ésta va entonces a este otro
# nombre, que dice en la cara cuál hay que leer.
NOMBRE_INFORME_ALTERNATIVO = (
    "_validacion (NO SE PUDO ESCRIBIR _validacion.txt - LEER ESTE).txt"
)

EXTENSIONES_DE_SALIDA = (".xlsx", ".docx")


@dataclass(frozen=True)
class Leido:
    """Un export que se pudo leer, ya separado en registros del mes y de fuera."""

    archivo: str
    del_mes: tuple[Registro, ...]
    descartados: tuple[Registro, ...]
    # Lo que el mapeo no pudo completar y no frena nada: hoy, el usuario de
    # Kimai de alguien que exportó con el resumen mensual y no está declarado.
    avisos: tuple[str, ...] = ()


def _reconfigurar_salida_utf8() -> None:
    """Evita que los acentos salgan rotos en la consola de Windows.

    No depende de que generar.bat sea el único punto de entrada: si alguien
    corre `python -m cunix_horas ...` desde otra terminal, esto reconfigura
    stdout/stderr a UTF-8 igual. Si la plataforma no lo soporta, no rompe.
    """
    for flujo in (sys.stdout, sys.stderr):
        reconfigurar = getattr(flujo, "reconfigure", None)
        if reconfigurar is not None:
            try:
                reconfigurar(encoding="utf-8")
            except (ValueError, OSError):
                pass


def _separar_por_mes(
    registros: list[Registro], periodo: Periodo
) -> tuple[list[Registro], list[Registro]]:
    del_mes = [r for r in registros if periodo.contiene(r.fecha)]
    fuera = [r for r in registros if not periodo.contiene(r.fecha)]
    return del_mes, fuera


def _error_de_export_vacio(archivo: str) -> str:
    return (
        f"Export sin ningún registro de horas en {archivo}: el archivo tiene "
        f"los encabezados de Kimai pero ninguna fila de datos.\n"
        f"  Revisá el rango de fechas del export en Kimai: lo más probable es "
        f"que esté puesto en un período sin horas cargadas.\n"
        f"  Volvé a exportar el mes que querés generar y reemplazá el archivo."
    )


def _error_de_mes_equivocado(
    descartados: list[Registro], periodo: Periodo, archivo: str
) -> str:
    fechas = sorted(r.fecha for r in descartados)
    primera, ultima = fechas[0], fechas[-1]
    return (
        f"Ningún registro del export cae dentro de {periodo.carpeta} en "
        f"{archivo}: los {len(descartados)} registros del archivo se "
        f"descartaron por fecha.\n"
        f"  Las fechas del export van del "
        f"{primera.day}/{primera.month}/{primera.year} al "
        f"{ultima.day}/{ultima.month}/{ultima.year}.\n"
        f"  Lo más probable es que el export se haya hecho con otro rango de "
        f"fechas, o que corresponda a otro mes del que estás generando.\n"
        f"  Volvé a exportar {periodo.carpeta} desde Kimai, o generá el mes "
        f"que realmente trae el archivo."
    )


def _escribir_atomico(destino: Path, escribir, verificar=None) -> None:
    """Genera el archivo en un temporal y recién ahí lo mueve sobre `destino`.

    La verificación corre sobre el temporal, **antes** del `os.replace`: si lo
    escrito no tiene las mismas horas que los exports, no llega a ocupar el
    nombre bueno.

    `os.replace` es atómico y pisa el destino existente, así que nunca hay una
    ventana en la que `destino` sea un archivo a medio escribir, y una falla
    durante la generación no toca el que ya estaba.

    El temporal vive en la misma carpeta que el destino: mover entre volúmenes
    no es atómico.
    """
    temporal = destino.with_name(f"~tmp-{os.getpid()}-{destino.name}")
    try:
        escribir(temporal)
        if verificar is not None:
            verificar(temporal)
        os.replace(temporal, destino)
    finally:
        try:
            temporal.unlink(missing_ok=True)
        except OSError:
            pass


def _apartar_previo(destino: Path) -> str:
    """Saca de en medio el archivo limpio que haya quedado de otra corrida.

    Se llama cuando el anexo limpio del mes NO se generó: si en output/ quedó
    el de una corrida anterior con ese mismo nombre, no puede seguir ahí
    haciéndose pasar por el de esta corrida. Se lo renombra en vez de borrarlo:
    el dato sigue estando por si el dueño lo necesita, pero el nombre ya no
    engaña.

    Devuelve el texto a agregar al informe (vacío si no había nada).
    """
    if not destino.exists():
        return ""

    def apartado(intento: int) -> Path:
        numero = "" if intento == 1 else f" {intento}"
        return destino.with_name(
            f"{destino.stem}{SUFIJO_CORRIDA_ANTERIOR}{numero}{destino.suffix}"
        )

    intento = 1
    while apartado(intento).exists():
        intento += 1
    try:
        destino.rename(apartado(intento))
    except OSError:
        return (
            f"\n  OJO: en output/ quedó {destino.name} de una corrida anterior "
            f"y no se pudo apartar (permiso denegado; es probable que lo tengas "
            f"abierto). NO lo envíes: no es el archivo de esta corrida."
        )
    return (
        f"\n  El {destino.name} que había quedado de una corrida anterior se "
        f"renombró a {apartado(intento).name} para que no se confunda con el "
        f"del mes."
    )


def _archivos_ajenos(carpeta_salida: Path, generados: list[str]) -> list[str]:
    """Los entregables que están en output/<mes>/ y NO salieron de esta corrida.

    El dueño no envía lo que hizo la corrida: envía lo que hay en la carpeta.
    Todo .xlsx o .docx que quede ahí sin que esta corrida lo haya generado —un
    anexo de hace tres corridas, un archivo de un formato anterior, algo que el
    dueño dejó a mano— se le va adjunto al cliente si nadie se lo nombra.

    Se ignoran los temporales y los archivos de bloqueo: no son del dueño y no
    sobreviven.
    """
    generados_ahora = set(generados)
    presentes: list[str] = []
    for extension in EXTENSIONES_DE_SALIDA:
        try:
            presentes.extend(p.name for p in carpeta_salida.glob(f"*{extension}"))
        except OSError:
            return []
    return [
        nombre
        for nombre in sorted(presentes)
        if nombre not in generados_ahora
        and not nombre.startswith("~")
        and not nombre.startswith(".~")
    ]


def _seccion_de_ajenos(mes: str, ajenos: list[str]) -> list[str]:
    """Las líneas del informe que enumeran lo que esta corrida no generó."""
    if not ajenos:
        return []

    apartados = [n for n in ajenos if SUFIJO_CORRIDA_ANTERIOR in n]
    otros = [n for n in ajenos if SUFIJO_CORRIDA_ANTERIOR not in n]

    lineas = [
        f"{len(ajenos)} archivo/s más en output/{mes}/ que esta corrida NO generó.",
        "NO LOS ENVÍES: no forman parte de la entrega de este mes.",
    ]
    if apartados:
        lineas.append("")
        lineas.append("Apartados por la herramienta (eran de una corrida anterior):")
        lineas.extend(f"  - {nombre}" for nombre in apartados)
    if otros:
        lineas.append("")
        lineas.append(
            "Nunca generados por esta corrida (quedaron de una corrida "
            "anterior, o los pusiste vos ahí):"
        )
        lineas.extend(f"  - {nombre}" for nombre in otros)
        lineas.append("")
        lineas.append(
            "Si alguno es de un formato anterior, o un anexo marcado como "
            "INCOMPLETO, está viejo: borralo o movelo fuera de la carpeta "
            "antes de armar el envío."
        )
    lineas.append("")
    return lineas


def _origenes_de(registros: list[Registro]) -> str:
    """De qué reporte de Kimai salieron las filas de un desarrollador.

    Casi siempre es uno solo, pero un export puede traer horas de dos personas
    y una persona puede aparecer en dos exports, así que se nombran todos.
    """
    return ", ".join(sorted({r.origen for r in registros}))


def _seccion_de_desarrolladores(leidos: list[Leido]) -> list[str]:
    """Quiénes entraron a los anexos, con cuántos registros, horas y origen.

    Va arriba de todo: es lo primero que el dueño tiene que poder contar
    contra la lista de su equipo antes de mandar nada.
    """
    entradas: list[tuple[str, str, int, float, str]] = []
    for leido in leidos:
        por_nombre: dict[str, list[Registro]] = {}
        for registro in leido.del_mes:
            por_nombre.setdefault(
                filas_anexo.nombre_para_mostrar(registro), []
            ).append(registro)
        for nombre in sorted(por_nombre, key=lambda n: (n.casefold(), n)):
            registros = por_nombre[nombre]
            entradas.append(
                (
                    nombre,
                    leido.archivo,
                    len(registros),
                    sum(r.horas for r in registros),
                    _origenes_de(registros),
                )
            )

    entradas.sort(key=lambda e: (e[0].casefold(), e[0], e[1]))
    lineas = [f"{len(entradas)} desarrollador/es en los anexos:"]
    if not entradas:
        lineas.append("  (ninguno)")
    lineas.extend(
        f"  - {nombre} ({archivo}): {cantidad} registro/s, {horas:.2f} h  [{origen}]"
        for nombre, archivo, cantidad, horas, origen in entradas
    )
    lineas.append("")

    resumidos = [e for e in entradas if ORIGEN_RESUMEN_MENSUAL in e[4]]
    if resumidos:
        lineas.append(
            f"{len(resumidos)} de ellos exportaron con el reporte de resumen "
            f"mensual. SUS FILAS VAN SIN DESCRIPCIÓN Y SIN HORARIO (las "
            f"columnas Inicio y Fin quedan vacías):"
        )
        lineas.extend(f"  - {e[0]} ({e[1]})" for e in resumidos)
        lineas.append("")
        lineas.append(
            "El resumen mensual no trae esas columnas. Para que esas filas "
            "lleven la descripción de lo que hizo cada uno y su horario, "
            "volvé a exportar a esas personas desde Kimai con el reporte de "
            "detalle y corré de nuevo."
        )
        lineas.append("")
    return lineas


def _seccion_de_fallos_de_escritura(fallos: list[tuple[str, str]]) -> list[str]:
    """Los anexos que no se pudieron escribir, con el motivo de cada uno.

    Es distinto de un export que no entró: acá los datos están bien y lo que
    falló fue la escritura. El otro anexo se genera igual —ningún fallo aborta
    la corrida de los demás—, pero la entrega del mes no está completa hasta
    que los dos salgan.
    """
    if not fallos:
        return []
    lineas = ["No se pudo generar:"]
    for nombre, motivo in fallos:
        lineas.append(f"  - {nombre}:")
        lineas.extend(f"    {linea}" for linea in motivo.splitlines())
    lineas.append("")
    lineas.append(
        "La entrega del mes son los DOS anexos: no envíes uno solo. Corregí lo "
        "de arriba y volvé a correr."
    )
    lineas.append("")
    return lineas


def _seccion_de_faltantes(no_leidos: list[tuple[str, str]]) -> list[str]:
    """Los exports que no entraron, con el motivo de cada uno."""
    if not no_leidos:
        return []
    lineas = [
        f"{len(no_leidos)} archivo/s que NO entraron. "
        "A LOS ANEXOS LES FALTAN ESOS DESARROLLADORES:"
    ]
    for entrada, motivo in no_leidos:
        lineas.append(f"  - {entrada}:")
        lineas.extend(f"    {linea}" for linea in motivo.splitlines())
    lineas.append("")
    lineas.append(
        "Los anexos de este mes salieron con la marca de INCOMPLETO en el "
        "nombre. NO los envíes así: corregí lo de arriba y volvé a correr. "
        "Si los mandás igual, C.UNIX aprueba de menos y nada adentro de los "
        "documentos lo delata."
    )
    lineas.append("")
    return lineas


def _seccion_de_sin_mapear(armado) -> list[str]:
    """Los proyectos que salieron con el nombre derivado de Kimai."""
    if armado is None or not armado.sin_mapear:
        return []
    lineas = [
        "--- Proyectos sin mapear ---",
        "",
        "No frenan nada: el proyecto salió con el nombre que trae Kimai. Si "
        "querés que C.UNIX vea otro nombre —y que coincida con el de la hoja "
        "«Datos», que es de donde sale el valor hora—, pegá esto en "
        "config/mapeo.yaml bajo proyectos: y ajustalo.",
        "",
    ]
    for proyecto in armado.sin_mapear:
        lineas.append(f"  {proyecto.codigo}:")
        lineas.append('    cliente: "AJUSTAR - nombre del cliente para C.UNIX"')
        lineas.append(f'    proyecto: "{proyecto.proyecto}"')
    lineas.append("")
    return lineas


def _seccion_de_sin_valor_hora(totales_proyecto) -> list[str]:
    """Los proyectos del mes que la hoja `Datos` no tarifa."""
    sin_valor = [t for t in totales_proyecto if t.valor_hora is None]
    if not sin_valor:
        return []
    lineas = [
        "--- Proyectos sin valor hora ---",
        "",
        f"{len(sin_valor)} proyecto/s del mes no tienen valor hora en la hoja "
        f"«Datos» de templates/{PLANTILLA_DETALLE}, así que en la tabla 1 del "
        f"informe salen con sus horas y SIN importe, y el total a facturar "
        f"quedó vacío:",
        "",
    ]
    lineas.extend(f"  - {t.proyecto} ({t.horas:.2f} h)" for t in sin_valor)
    lineas.append("")
    lineas.append(
        "Cargá el valor hora de esos proyectos en la hoja «Datos» de la "
        "plantilla (columnas Proyecto / Cliente / Valor hora, filas 8 a 22), "
        "con el nombre del proyecto escrito igual que acá arriba, y volvé a "
        "correr."
    )
    lineas.append("")
    return lineas


def _seccion_de_sin_descripcion(filas) -> list[str]:
    """Quiénes tienen registros sin descripción, y cuántos.

    El contrato de C.UNIX exige que cada registro diga qué se hizo para poder
    aprobar esas horas, así que esto va con nombre y cantidad. No frena nada:
    el dueño decide si las completa o las manda así.
    """
    faltantes = filas_anexo.sin_descripcion(filas)
    if not faltantes:
        return []
    total = sum(cantidad for _, cantidad in faltantes)
    lineas = [
        "--- Registros sin descripción ---",
        "",
        f"{total} registro/s del mes van SIN DESCRIPCIÓN. El contrato de "
        f"C.UNIX pide que cada registro diga qué se hizo, empezando por el "
        f"ticket de iTop o la tarea de ClickUp: esas horas puede no "
        f"aprobarlas.",
        "",
    ]
    lineas.extend(
        f"  - {persona}: {cantidad} registro/s" for persona, cantidad in faltantes
    )
    lineas.append("")
    lineas.append(
        "Pedile a esas personas que completen la descripción en Kimai y volvé "
        "a exportar, o cargala a mano en la hoja «Detalle» antes de enviar."
    )
    lineas.append("")
    return lineas


def _seccion_de_sin_usuario(leidos: list[Leido]) -> list[str]:
    """Quiénes van a la hoja «Datos» sin usuario de Kimai, y por qué.

    No frena nada: el usuario de Kimai es una sola celda de la hoja `Datos` y
    las horas de esa persona entran completas igual. Pero C.UNIX pide que cada
    uno cargue con su propio usuario, así que la celda vacía tiene que estar
    explicada acá y no aparecer sola en el anexo.
    """
    avisos = [aviso for leido in leidos for aviso in leido.avisos]
    if not avisos:
        return []
    lineas = [
        "--- Personas sin usuario de Kimai en la hoja «Datos» ---",
        "",
        f"{len(avisos)} persona/s del mes van con la columna «Usuario Kimai» "
        f"vacía. NO FRENA NADA: sus horas y sus filas están completas en los "
        f"dos anexos.",
        "",
    ]
    for aviso in avisos:
        lineas.extend(f"  {linea}" for linea in aviso.splitlines())
        lineas.append("")
    return lineas


def _resumen_de_validacion(
    mes: str,
    leidos: list[Leido],
    no_leidos: list[tuple[str, str]],
    generados: list[str],
    fallos: list[tuple[str, str]],
    armado,
    filas,
    totales_proyecto,
    notas_de_apartado: list[str],
    ajenos: list[str],
    avisos_por_dev: list[str],
    error_fatal: str,
) -> str:
    """Arma el contenido de `_validacion.txt`.

    Este archivo es el informe que el dueño lee antes de mandar nada, así que
    tiene que alcanzar por sí solo para decidir. Arriba de todo va quién entró
    y quién no: con todas las personas en un solo par de documentos, esa lista
    es lo único que distingue un envío completo de uno que no lo es.
    """
    lineas: list[str] = [f"Corrida de output/{mes}/", ""]

    if error_fatal:
        lineas.append("NO SE GENERÓ NINGÚN ANEXO. El motivo:")
        lineas.append("")
        lineas.extend(f"  {linea}" for linea in error_fatal.splitlines())
        lineas.append("")

    lineas.extend(_seccion_de_desarrolladores(leidos))
    lineas.extend(_seccion_de_faltantes(no_leidos))
    lineas.extend(_seccion_de_fallos_de_escritura(fallos))

    if generados:
        lineas.append("Para enviarle a C.UNIX:")
        lineas.extend(f"  - {nombre}" for nombre in generados)
        if filas:
            lineas.append(
                f"  {len(filas)} fila/s en la hoja Detalle, "
                f"{filas_anexo.total_de(filas):.2f} h en total."
            )
    elif not error_fatal:
        lineas.append("NO SE GENERÓ NINGÚN ANEXO.")
    for nota in notas_de_apartado:
        lineas.extend(nota.strip("\n").splitlines())
    lineas.append("")

    if generados:
        lineas.append("Lo que hay que completar a mano antes de enviar:")
        lineas.append(
            "  - Anexo II, tabla 3: el borrador de los principales trabajos. "
            "Está agrupado por proyecto y ticket y los totales cierran, pero "
            "la redacción es mecánica: reescribila y fusioná líneas sumando "
            "sus horas."
        )
        lineas.append("  - Anexo II, tabla 4: las tres observaciones ([●]).")
        lineas.append(
            "  - Anexo II, tabla 3, columna «Estado al cierre»: Terminado, En "
            "curso, En revisión de C.UNIX o Bloqueado."
        )
        lineas.append("  - Anexo II, firmas: nombre, cargo y fecha de quien emite.")
        lineas.append("")

    lineas.extend(_seccion_de_ajenos(mes, ajenos))
    lineas.extend(_seccion_de_sin_mapear(armado))
    lineas.extend(_seccion_de_sin_valor_hora(totales_proyecto))
    lineas.extend(_seccion_de_sin_descripcion(filas))
    lineas.extend(_seccion_de_sin_usuario(leidos))

    lineas.append("--- Avisos de validación ---")
    lineas.append("")
    if avisos_por_dev:
        lineas.extend(avisos_por_dev)
    elif no_leidos or fallos:
        # Nunca "Sin avisos." a secas cuando faltó alguien: sería el mensaje
        # más tranquilizador posible al lado de un envío incompleto.
        lineas.append(
            "Ningún aviso sobre los desarrolladores que sí entraron, pero los "
            "anexos NO están completos: mirá la lista de arriba."
        )
    elif ajenos:
        lineas.append(
            "Ningún aviso sobre lo generado, pero la carpeta tiene "
            f"{len(ajenos)} archivo/s que no son de esta corrida: mirá la "
            "lista de arriba antes de armar el envío."
        )
    else:
        lineas.append("Sin avisos.")
    return "\n".join(lineas).rstrip("\n") + "\n"


def _escribir_informe_de_validacion(
    carpeta_salida: Path, mes: str, contenido: str
) -> bool:
    """Deja el informe de ESTA corrida en disco. Devuelve si fue en su nombre.

    Si `_validacion.txt` no se puede escribir, el que quedó en disco es el de
    una corrida anterior y describe otra cosa. Escribirlo igual en otro lado es
    mejor que dejar al dueño con el viejo y sin saberlo, así que se escribe en
    un nombre que dice a la vista cuál de los dos hay que leer.
    """
    try:
        (carpeta_salida / NOMBRE_INFORME).write_text(contenido, encoding="utf-8")
        return True
    except PermissionError:
        print(
            f"ERROR: no se pudo escribir output/{mes}/{NOMBRE_INFORME}: "
            f"permiso denegado."
        )
        print("  Es probable que lo tengas abierto en el Bloc de notas.")
    except OSError as error:
        print(f"ERROR: no se pudo escribir output/{mes}/{NOMBRE_INFORME}: {error}")

    print(
        f"  OJO: el {NOMBRE_INFORME} que hay en output/{mes}/ es de una corrida "
        "anterior y NO describe lo que hay ahora en la carpeta."
    )
    alternativo = carpeta_salida / NOMBRE_INFORME_ALTERNATIVO
    try:
        alternativo.write_text(contenido, encoding="utf-8")
    except OSError as error:
        print(f"  Tampoco se pudo escribir {NOMBRE_INFORME_ALTERNATIVO}: {error}")
        print("  No hay informe de esta corrida: no envíes nada de esa carpeta.")
        return False
    print(f"  El informe de esta corrida quedó en output/{mes}/{alternativo.name}")
    print("  Leé ESE y borrá el viejo.")
    return False


def _avisos_agrupados_por_dev(leidos: list[Leido], periodo: Periodo):
    """Los avisos de cada persona, bajo su nombre, en un solo informe."""
    del_mes: dict[str, list[Registro]] = {}
    fuera: dict[str, list[Registro]] = {}
    for leido in leidos:
        for registro in leido.del_mes:
            del_mes.setdefault(
                filas_anexo.nombre_para_mostrar(registro), []
            ).append(registro)
        for registro in leido.descartados:
            fuera.setdefault(
                filas_anexo.nombre_para_mostrar(registro), []
            ).append(registro)

    lineas: list[str] = []
    cantidad = 0
    for nombre in sorted(set(del_mes) | set(fuera), key=lambda n: (n.casefold(), n)):
        avisos = avisos_de_desarrollador(
            del_mes.get(nombre, []), fuera.get(nombre, []), periodo.anio, periodo.mes
        )
        cantidad += len(avisos)
        if avisos:
            lineas.append(f"=== {nombre} ===")
            lineas.extend(f"  {a}" for a in avisos)
            lineas.append("")
    return lineas, cantidad


def _leer_entradas(
    entradas: list[Path], mapeo: Mapeo, periodo: Periodo, registrar_fallo
) -> list[Leido]:
    """Lee cada export y le completa lo que el mapeo tenga que poner.

    Un archivo que falla no frena a los demás.
    """
    leidos: list[Leido] = []
    for entrada in entradas:
        try:
            completado = completar_desde_mapeo(leer(entrada), mapeo, entrada.name)
            registros = list(completado.registros)
        except (ErrorLectura, ErrorMapeo) as error:
            registrar_fallo(entrada.name, str(error))
            continue
        except (OSError, ValueError) as error:
            registrar_fallo(entrada.name, f"Error al leer {entrada.name}: {error}")
            continue

        if not registros:
            registrar_fallo(entrada.name, _error_de_export_vacio(entrada.name))
            continue

        del_mes, descartados = _separar_por_mes(registros, periodo)
        if not del_mes:
            registrar_fallo(
                entrada.name,
                _error_de_mes_equivocado(descartados, periodo, entrada.name),
            )
            continue

        leidos.append(
            Leido(
                entrada.name,
                tuple(del_mes),
                tuple(descartados),
                completado.avisos,
            )
        )
    return leidos


def _equipo_de(totales_persona, usuarios: dict[str, str]):
    """Las filas de la hoja `Datos`, en el mismo orden que la tabla 2."""
    return tuple(
        (total.persona, total.perfil, usuarios.get(total.persona, ""))
        for total in totales_persona
    )


def procesar_mes(mes: str, raiz: Path) -> int:
    """Procesa todos los exports de input/<mes>/. Devuelve el código de salida."""
    try:
        periodo = Periodo.parsear(mes)
    except ErrorPeriodo as error:
        print(f"ERROR: {error}")
        return 1

    carpeta_entrada = raiz / "input" / mes
    if not carpeta_entrada.is_dir():
        print(f"ERROR: no existe la carpeta input/{mes}")
        print(f"  Creála y poné adentro los exports de Kimai: {carpeta_entrada}")
        return 1

    try:
        mapeo = Mapeo.cargar(raiz / "config" / "mapeo.yaml")
    except ErrorMapeo as error:
        print(f"ERROR: {error}")
        return 1

    plantilla_detalle = raiz / "templates" / PLANTILLA_DETALLE
    plantilla_informe = raiz / "templates" / PLANTILLA_INFORME
    for plantilla in (plantilla_detalle, plantilla_informe):
        if not plantilla.is_file():
            print(f"ERROR: falta la plantilla {plantilla.name}")
            print(f"  Se esperaba encontrarla en {plantilla}")
            print(
                "  Es la plantilla vacía que mandó C.UNIX: sin ella no hay de "
                "dónde sacar el formato del anexo."
            )
            return 1

    # Los formatos de export de Kimai: .xlsx y .csv, sin entrar en subcarpetas
    # (manual/ se lee aparte). '~$' es el archivo de bloqueo que deja Excel.
    entradas = sorted(
        p
        for extension in EXTENSIONES_DE_ENTRADA
        for p in carpeta_entrada.glob(f"*{extension}")
        if not p.name.startswith("~$")
    )
    if not entradas:
        print(
            f"ERROR: no hay ningún export de Kimai "
            f"({', '.join(EXTENSIONES_DE_ENTRADA)}) en input/{mes}"
        )
        return 1

    carpeta_salida = raiz / "output" / mes
    # Único punto de salida que NO escribe _validacion.txt: si la carpeta no se
    # pudo crear, no se tocó nada de output/ y tampoco hay dónde escribirlo.
    try:
        carpeta_salida.mkdir(parents=True, exist_ok=True)
    except PermissionError:
        print(f"ERROR: no se pudo crear output/{mes}: permiso denegado.")
        return 1
    except OSError as error:
        print(f"ERROR: no se pudo preparar output/{mes}: {error}")
        return 1

    print(f"Procesando input/{mes}/ ...")
    no_leidos: list[tuple[str, str]] = []

    def registrar_fallo(entrada: str, motivo: str) -> None:
        """Deja el fallo en consola y en la lista que va a _validacion.txt."""
        print(f"  {entrada}: NO ENTRA A LOS ANEXOS")
        print(motivo)
        no_leidos.append((entrada, motivo))

    leidos = _leer_entradas(entradas, mapeo, periodo, registrar_fallo)

    for leido in leidos:
        nombres = sorted({filas_anexo.nombre_para_mostrar(r) for r in leido.del_mes})
        horas = sum(r.horas for r in leido.del_mes)
        origen = _origenes_de(list(leido.del_mes))
        print(
            f"  {leido.archivo}  ->  {', '.join(nombres)}"
            f"  ({len(leido.del_mes)} registro/s, {horas:.2f} h)  [{origen}]"
        )
        if ORIGEN_RESUMEN_MENSUAL in origen:
            print(
                "    OJO: sus filas van sin descripción y sin horario. Está "
                "en el informe."
            )
        for aviso in leido.avisos:
            print(f"    AVISO: {aviso.splitlines()[0]}")

    registros = [r for leido in leidos for r in leido.del_mes]
    armado = filas_anexo.construir(registros, mapeo) if leidos else None
    filas: tuple = () if armado is None else armado.filas
    usuarios: tuple = () if armado is None else armado.usuarios

    error_fatal = ""
    if armado is not None and mapeo.fuentes_manuales.hay_alguna:
        try:
            filas, usuarios = fuentes_manuales.aplicar(
                filas, usuarios, mapeo.fuentes_manuales, carpeta_entrada, periodo
            )
        except fuentes_manuales.ErrorFuenteManual as error:
            error_fatal = str(error)
            print(f"ERROR: {error_fatal}")
            filas, usuarios = (), ()

    valores_hora = detalle_xlsx.leer_valores_hora(plantilla_detalle)
    totales_proyecto = filas_anexo.por_proyecto(filas, valores_hora)
    totales_persona = filas_anexo.por_persona(
        filas, mapeo.anexos.perfiles_por_persona, mapeo.anexos.perfil_por_defecto
    )
    trabajos = filas_anexo.principales_trabajos(filas)
    equipo = _equipo_de(totales_persona, dict(usuarios))

    limpios = {
        "detalle": mapeo.anexos.nombre_detalle(periodo),
        "informe": mapeo.anexos.nombre_informe(periodo),
    }
    nombres_salida = {
        clave: nombre_incompleto(nombre, len(no_leidos)) if no_leidos else nombre
        for clave, nombre in limpios.items()
    }

    generados: list[str] = []
    # Los anexos que no se pudieron escribir. Van aparte de `no_leidos`: el
    # nombre de los archivos lo decide lo que no entró de input/, y no una
    # falla de escritura posterior, que si no haría que el nombre del segundo
    # anexo dependiera del orden en que se generan.
    fallos: list[tuple[str, str]] = []
    notas: list[str] = []

    def registrar_fallo_de_escritura(nombre: str, motivo: str) -> None:
        print(f"  {nombre}: NO SE PUDO GENERAR")
        print(motivo)
        fallos.append((nombre, motivo))

    def apartar(clave: str) -> None:
        nota = _apartar_previo(carpeta_salida / limpios[clave])
        if nota:
            notas.append(nota)
            print(nota.strip())

    def generar(clave: str, escribir, verificar=None) -> None:
        """Genera un anexo. Si falla, lo registra y sigue con el otro."""
        destino = carpeta_salida / nombres_salida[clave]
        # Si el anexo sale marcado como incompleto, el limpio de una corrida
        # anterior no puede quedar ahí: es el que el dueño adjuntaría sin
        # pensarlo.
        if no_leidos:
            apartar(clave)
        try:
            _escribir_atomico(destino, escribir, verificar)
            generados.append(nombres_salida[clave])
        except PermissionError:
            registrar_fallo_de_escritura(
                nombres_salida[clave],
                f"No se pudo escribir {nombres_salida[clave]}: permiso "
                f"denegado.\n"
                f"  Es probable que tengas ese archivo abierto. Cerralo y "
                f"volvé a correr.",
            )
            apartar(clave)
        except (
            detalle_xlsx.ErrorIntegridad,
            detalle_xlsx.ErrorCapacidad,
            informe_docx.ErrorPlantillaInforme,
        ) as error:
            registrar_fallo_de_escritura(nombres_salida[clave], str(error))
            apartar(clave)
        except (OSError, ValueError) as error:
            registrar_fallo_de_escritura(
                nombres_salida[clave],
                f"Error al generar {nombres_salida[clave]}: {error}",
            )
            apartar(clave)

    if not filas:
        # Sin una sola fila no hay anexo que generar: uno vacío con nombre de
        # entregable es exactamente lo que esto busca evitar.
        for clave in limpios:
            apartar(clave)
        if not error_fatal:
            print("  No se generó ningún anexo: no se pudo leer ningún export.")
    else:
        horas_esperadas = filas_anexo.total_de(filas)
        generar(
            "detalle",
            lambda destino: detalle_xlsx.escribir(
                plantilla_detalle,
                destino,
                periodo,
                filas,
                equipo,
                mapeo.anexos.contrato_de_fecha,
            ),
            lambda destino: detalle_xlsx.verificar_integridad(
                destino, horas_esperadas
            ),
        )
        generar(
            "informe",
            lambda destino: informe_docx.escribir(
                plantilla_informe,
                destino,
                periodo,
                totales_proyecto,
                totales_persona,
                trabajos,
                informe_docx.TextosInforme(
                    dias_habiles_entrega=mapeo.anexos.dias_habiles_entrega,
                    contrato_de_fecha=mapeo.anexos.contrato_de_fecha,
                    fecha_de_emision=mapeo.anexos.fecha_de_emision,
                ),
            ),
        )

    for nombre in generados:
        print(f"  Generado: {nombre}")

    avisos_por_dev, cantidad_avisos = _avisos_agrupados_por_dev(leidos, periodo)
    ajenos = _archivos_ajenos(carpeta_salida, generados)
    contenido = _resumen_de_validacion(
        mes,
        leidos,
        no_leidos,
        generados,
        fallos,
        armado,
        filas,
        totales_proyecto,
        notas,
        ajenos,
        avisos_por_dev,
        error_fatal,
    )
    informe_en_su_nombre = _escribir_informe_de_validacion(
        carpeta_salida, mes, contenido
    )

    if ajenos:
        print(
            f"  OJO: en output/{mes}/ hay {len(ajenos)} archivo/s que esta "
            "corrida NO generó. No los envíes; están listados en el informe."
        )
    print(
        f"{len(leidos)} export/s leído/s, {len(no_leidos)} no leído/s, "
        f"{cantidad_avisos} aviso/s en output/{mes}/{NOMBRE_INFORME}"
    )
    if no_leidos:
        print(
            "  Los anexos salieron marcados como INCOMPLETO: NO los envíes "
            "hasta corregir lo de arriba y volver a correr."
        )
    if fallos:
        print(
            f"  {len(fallos)} anexo/s NO se pudieron generar. La entrega del "
            "mes son los dos: no envíes uno solo."
        )

    if not informe_en_su_nombre or error_fatal or fallos:
        return 1
    # Los archivos ajenos NO cambian el código de salida. Que la carpeta tenga
    # archivos de más no significa que la corrida haya fallado: la única acción
    # pendiente es del dueño, sobre su propia carpeta. Un código 1 recurrente
    # enseñaría a ignorar el código de salida, que es lo que distingue una
    # corrida incompleta de una completa.
    return 1 if no_leidos else 0


def main(argv: list[str] | None = None) -> int:
    _reconfigurar_salida_utf8()
    argumentos = sys.argv[1:] if argv is None else argv
    if len(argumentos) != 1:
        print("Uso: python -m cunix_horas AAAA-MM")
        print("Ejemplo: python -m cunix_horas 2026-09")
        return 1
    return procesar_mes(argumentos[0], Path.cwd())
