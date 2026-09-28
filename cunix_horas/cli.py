"""Orquestación: de los exports de input/<mes>/ al archivo del partner.

El entregable es **uno solo por mes**: `output/<mes>/Horas KLG-<Mes><Año>.xlsx`,
con el detalle plano de todos los desarrolladores juntos.

Ese cambio trae un riesgo que el formato anterior no tenía. Antes, si el
archivo de un desarrollador fallaba, faltaba un Excel entero en la carpeta:
imposible no notarlo. Ahora todo va en un solo archivo, así que **un
desarrollador que falta es invisible**: el archivo se ve completo y no lo es.

Por eso, cuando algún export falla:

- el consolidado **se genera igual** con los que sí se pudieron leer, porque
  no generarlo dejaría al dueño sin nada y sin forma de revisar;
- pero sale con la marca de incompleto en el **nombre del archivo**, que es lo
  único que el dueño ve cuando lo adjunta a un mail;
- y el archivo limpio que hubiera quedado de una corrida anterior se aparta,
  para que no se pueda enviar en su lugar.

El escritor pivoteado por desarrollador (`escritor_excel.py`) y su plantilla
siguen en el repo, probados, pero este CLI ya no los llama.
"""
from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from cunix_horas.completado import completar_desde_mapeo
from cunix_horas.detalle import Detalle, construir, nombre_para_mostrar, segundos_de
from cunix_horas.escritor_detalle import (
    ErrorIntegridad,
    escribir_detalle,
    nombre_de_archivo,
    nombre_incompleto,
    verificar_integridad,
)
from cunix_horas.kimai_comun import ORIGEN_RESUMEN_MENSUAL
from cunix_horas.lector_kimai import (
    EXTENSIONES_DE_ENTRADA,
    ErrorLectura,
    Registro,
    leer,
)
from cunix_horas.mapeo import ErrorMapeo, Mapeo
from cunix_horas.validador import avisos_de_desarrollador

FORMATO_MES = re.compile(r"^(\d{4})-(\d{2})$")

SEGUNDOS_POR_HORA = 3600

# Marca del archivo que quedó de una corrida anterior y que esta corrida no
# pudo reemplazar. Va en el nombre para que no se confunda con el del mes.
SUFIJO_CORRIDA_ANTERIOR = " (CORRIDA ANTERIOR - NO ENVIAR)"

NOMBRE_INFORME = "_validacion.txt"

# Si _validacion.txt no se puede escribir, el que quedó en disco es el de
# otra corrida y describe otra carpeta. El informe de esta corrida va
# entonces a este otro nombre, que dice en la cara cuál hay que leer.
NOMBRE_INFORME_ALTERNATIVO = "_validacion (NO SE PUDO ESCRIBIR _validacion.txt - LEER ESTE).txt"


@dataclass(frozen=True)
class Leido:
    """Un export que se pudo leer, ya separado en registros del mes y de fuera."""

    archivo: str
    del_mes: tuple[Registro, ...]
    descartados: tuple[Registro, ...]


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


def _parsear_mes(mes: str) -> tuple[int, int]:
    coincidencia = FORMATO_MES.match(mes)
    if coincidencia is None:
        raise ValueError(
            f"El mes debe tener el formato AAAA-MM (ej: 2025-10), no {mes!r}"
        )
    anio, numero = int(coincidencia.group(1)), int(coincidencia.group(2))
    if not 1 <= numero <= 12:
        raise ValueError(
            f"Mes fuera de rango en {mes!r}: AAAA-MM con MM entre 01 y 12"
        )
    return anio, numero


def _separar_por_mes(
    registros: list[Registro], anio: int, mes: int
) -> tuple[list[Registro], list[Registro]]:
    del_mes = [r for r in registros if (r.fecha.year, r.fecha.month) == (anio, mes)]
    fuera = [r for r in registros if (r.fecha.year, r.fecha.month) != (anio, mes)]
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
    descartados: list[Registro], anio: int, mes: int, archivo: str
) -> str:
    fechas = sorted(r.fecha for r in descartados)
    primera, ultima = fechas[0], fechas[-1]
    return (
        f"Ningún registro del export cae dentro de {anio}-{mes:02d} en {archivo}: "
        f"los {len(descartados)} registros del archivo se descartaron por fecha.\n"
        f"  Las fechas del export van del {primera.day}/{primera.month}/{primera.year} "
        f"al {ultima.day}/{ultima.month}/{ultima.year}.\n"
        f"  Lo más probable es que el export se haya hecho con otro rango de fechas, "
        f"o que corresponda a otro mes del que estás generando.\n"
        f"  Volvé a exportar {anio}-{mes:02d} desde Kimai, o generá el mes que "
        f"realmente trae el archivo."
    )


def _escribir_atomico(
    detalle: Detalle, destino: Path, segundos_esperados: int
) -> None:
    """Genera el archivo en un temporal y recién ahí lo mueve sobre `destino`.

    La verificación de integridad corre sobre el temporal, **antes** del
    `os.replace`: si lo escrito no tiene las mismas horas que los exports, no
    llega a ocupar el nombre bueno.

    `os.replace` es atómico y pisa el destino existente, así que nunca hay una
    ventana en la que `destino` sea un archivo a medio escribir, y una falla
    durante la generación no toca el que ya estaba.

    El temporal vive en la misma carpeta que el destino: mover entre volúmenes
    no es atómico.
    """
    temporal = destino.with_name(f"~tmp-{os.getpid()}-{destino.name}")
    try:
        escribir_detalle(detalle, temporal)
        verificar_integridad(temporal, segundos_esperados)
        os.replace(temporal, destino)
    finally:
        try:
            temporal.unlink(missing_ok=True)
        except OSError:
            pass


def _apartar_excel_previo(destino: Path) -> str:
    """Saca de en medio el archivo viejo que haya con el nombre `destino`.

    Se llama cuando el archivo limpio del mes NO se generó: si en output/ quedó
    el de una corrida anterior con ese mismo nombre, no puede seguir ahí
    haciéndose pasar por el de esta corrida. Se lo renombra en vez de borrarlo:
    el dato sigue estando por si el dueño lo necesita, pero el nombre ya no
    engaña.

    Devuelve el texto a agregar al motivo del fallo (vacío si no había nada).
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


def _xlsx_que_esta_corrida_no_genero(
    carpeta_salida: Path, generados: list[str]
) -> list[str]:
    """Los .xlsx que están en output/<mes>/ y NO salieron de esta corrida.

    El dueño no envía lo que hizo la corrida: envía lo que hay en la carpeta.
    Todo .xlsx que quede ahí sin que esta corrida lo haya generado —los Excel
    por desarrollador del formato anterior, un consolidado incompleto de hace
    tres corridas, un archivo que el dueño dejó a mano— se le va adjunto al
    cliente si nadie se lo nombra. Por eso el informe mira el disco, no la
    corrida.

    Se ignoran los temporales: no son archivos del dueño y no sobreviven.
    """
    generados_ahora = set(generados)
    try:
        presentes = sorted(p.name for p in carpeta_salida.glob("*.xlsx"))
    except OSError:
        return []
    return [
        nombre
        for nombre in presentes
        if nombre not in generados_ahora
        and not nombre.startswith("~tmp-")
        and not nombre.startswith("~$")
    ]


def _seccion_de_ajenos(mes: str, ajenos: list[str]) -> list[str]:
    """Las líneas del informe que enumeran lo que esta corrida no generó."""
    if not ajenos:
        return []

    apartados = [n for n in ajenos if SUFIJO_CORRIDA_ANTERIOR in n]
    otros = [n for n in ajenos if SUFIJO_CORRIDA_ANTERIOR not in n]

    lineas = [
        f"{len(ajenos)} .xlsx más en output/{mes}/ que esta corrida NO generó.",
        "NO LOS ENVÍES: no forman parte de la entrega de este mes.",
    ]
    if apartados:
        lineas.append("")
        lineas.append(
            "Apartados por la herramienta (eran de una corrida anterior):"
        )
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
            "Si alguno es un Excel del formato anterior (uno por "
            "desarrollador) o un consolidado marcado como INCOMPLETO, está "
            "viejo: borralo o movelo fuera de la carpeta antes de armar el "
            "envío."
        )
    lineas.append("")
    return lineas


def _horas(segundos: int) -> float:
    return segundos / SEGUNDOS_POR_HORA


def _origenes_de(registros: list[Registro]) -> str:
    """De qué reporte de Kimai salieron las filas de un desarrollador.

    Casi siempre es uno solo, pero un export puede traer horas de dos
    personas y una persona puede aparecer en dos exports, así que se nombran
    todos los que haya.
    """
    return ", ".join(sorted({r.origen for r in registros}))


def _seccion_de_desarrolladores(leidos: list[Leido]) -> list[str]:
    """Quiénes entraron al archivo, con cuántos registros, horas y origen.

    Va arriba de todo: es lo primero que el dueño tiene que poder contar
    contra la lista de su equipo antes de mandar nada. Y de cada uno se dice
    **de qué reporte de Kimai salieron sus filas**, porque el resumen mensual
    da filas sin descripción y sin hora de inicio: el archivo es válido, pero
    el dueño tiene que poder decidir si lo manda así.
    """
    entradas: list[tuple[str, str, int, int, str]] = []
    for leido in leidos:
        por_nombre: dict[str, list[Registro]] = {}
        for registro in leido.del_mes:
            por_nombre.setdefault(nombre_para_mostrar(registro), []).append(registro)
        for nombre in sorted(por_nombre, key=lambda n: (n.casefold(), n)):
            registros = por_nombre[nombre]
            entradas.append(
                (
                    nombre,
                    leido.archivo,
                    len(registros),
                    sum(segundos_de(r.horas) for r in registros),
                    _origenes_de(registros),
                )
            )

    entradas.sort(key=lambda e: (e[0].casefold(), e[0], e[1]))
    lineas = [f"{len(entradas)} desarrollador/es en el archivo:"]
    if not entradas:
        lineas.append("  (ninguno)")
    lineas.extend(
        f"  - {nombre} ({archivo}): {cantidad} registro/s, "
        f"{_horas(segundos):.2f} h  [{origen}]"
        for nombre, archivo, cantidad, segundos, origen in entradas
    )
    lineas.append("")

    resumidos = [e for e in entradas if ORIGEN_RESUMEN_MENSUAL in e[4]]
    if resumidos:
        lineas.append(
            f"{len(resumidos)} de ellos exportaron con el reporte de resumen "
            f"mensual. SUS FILAS VAN SIN DESCRIPCIÓN Y SIN HORA DE INICIO "
            f"(la hora queda en 00:00):"
        )
        lineas.extend(f"  - {e[0]} ({e[1]})" for e in resumidos)
        lineas.append("")
        lineas.append(
            "Esto NO frena nada y el archivo es válido: el resumen mensual no "
            "trae esas dos columnas, y el partner las acepta vacías (112 de "
            "las 160 filas de su archivo de referencia no tienen "
            "descripción). El usuario, el mail y el número de proyecto de "
            "esas filas salieron de config/mapeo.yaml."
        )
        lineas.append(
            "Si querés que esas filas lleven la descripción de lo que hizo "
            "cada uno y su hora de inicio, volvé a exportar a esas personas "
            "desde Kimai con el reporte de detalle y corré de nuevo."
        )
        lineas.append("")
    return lineas


def _seccion_de_faltantes(no_leidos: list[tuple[str, str]]) -> list[str]:
    """Los exports que no entraron, con el motivo de cada uno."""
    if not no_leidos:
        return []
    lineas = [
        f"{len(no_leidos)} archivo/s de input/ que NO entraron. "
        "AL ARCHIVO LE FALTAN ESOS DESARROLLADORES:"
    ]
    for entrada, motivo in no_leidos:
        lineas.append(f"  - {entrada}:")
        lineas.extend(f"    {linea}" for linea in motivo.splitlines())
    lineas.append("")
    lineas.append(
        "El archivo de este mes salió con la marca de INCOMPLETO en el "
        "nombre. NO lo envíes así: corregí lo de arriba y volvé a correr. "
        "Si lo mandás igual, el partner factura de menos y nada adentro del "
        "archivo lo delata."
    )
    lineas.append("")
    return lineas


def _seccion_de_sin_mapear(detalle: Detalle) -> list[str]:
    """Los proyectos que salieron con el nombre derivado de Kimai."""
    if not detalle.sin_mapear:
        return []
    lineas = [
        "--- Proyectos sin mapear ---",
        "",
        "No frenan nada: cada fila lleva su Project number y el proyecto salió "
        "con el nombre que trae Kimai. Si querés que el partner vea otro "
        "nombre, pegá esto en config/mapeo.yaml bajo proyectos: y ajustalo.",
        "",
    ]
    for proyecto in detalle.sin_mapear:
        lineas.append(f"  {proyecto.codigo}:")
        lineas.append(f'    cliente: "{proyecto.cliente}"')
        lineas.append(f'    proyecto: "{proyecto.proyecto}"')
    lineas.append("")
    return lineas


def _avisos_de_numeros_ambiguos(detalle: Detalle) -> list[str]:
    """Un mismo Project number con más de un nombre de proyecto en el mes."""
    return [
        f"El Project number {numero} aparece con {len(nombres)} nombres de "
        f"proyecto distintos: {', '.join(nombres)}. El partner ve dos nombres "
        f"para lo mismo. Lo más probable es que alguien haya renombrado el "
        f"proyecto en Kimai a mitad de mes; si querés un solo nombre, "
        f"declaralo en config/mapeo.yaml bajo proyectos:."
        for numero, nombres in detalle.numeros_ambiguos
    ]


def _resumen_de_validacion(
    mes: str,
    leidos: list[Leido],
    no_leidos: list[tuple[str, str]],
    detalle: Detalle | None,
    generado: str | None,
    nota_de_apartado: str,
    ajenos: list[str],
    avisos_generales: list[str],
    avisos_por_dev: list[str],
) -> str:
    """Arma el contenido de _validacion.txt.

    Este archivo es el informe que el dueño lee antes de mandar nada, así que
    tiene que alcanzar por sí solo para decidir. Arriba de todo va quién entró
    al archivo y quién no: con todos los desarrolladores en un solo Excel, esa
    lista es lo único que distingue un envío completo de uno que no lo es.
    """
    lineas: list[str] = [f"Corrida de output/{mes}/", ""]

    lineas.extend(_seccion_de_desarrolladores(leidos))
    lineas.extend(_seccion_de_faltantes(no_leidos))

    if generado is not None and detalle is not None:
        lineas.append(f"Archivo para el partner: {generado}")
        lineas.append(
            f"  {len(detalle.filas)} fila/s, {detalle.horas:.2f} h en total."
        )
    else:
        lineas.append("NO SE GENERÓ NINGÚN ARCHIVO para el partner.")
    if nota_de_apartado:
        lineas.extend(nota_de_apartado.strip("\n").splitlines())
    lineas.append("")

    lineas.extend(_seccion_de_ajenos(mes, ajenos))

    if detalle is not None:
        lineas.extend(_seccion_de_sin_mapear(detalle))

    lineas.append("--- Avisos de validación ---")
    lineas.append("")
    if avisos_generales:
        lineas.extend(f"  {a}" for a in avisos_generales)
        lineas.append("")
    if avisos_por_dev:
        lineas.extend(avisos_por_dev)
    elif not avisos_generales:
        if no_leidos:
            # Nunca "Sin avisos." a secas cuando faltó alguien: sería el
            # mensaje más tranquilizador posible al lado de un envío al que
            # le faltan desarrolladores.
            lineas.append(
                "Ningún aviso sobre los desarrolladores que sí entraron, pero "
                "el archivo NO está completo: mirá la lista de arriba."
            )
        elif ajenos:
            lineas.append(
                "Ningún aviso sobre el archivo generado, pero la carpeta tiene "
                f"{len(ajenos)} .xlsx que no son de esta corrida: mirá la lista "
                "de arriba antes de armar el envío."
            )
        else:
            lineas.append("Sin avisos.")
    return "\n".join(lineas).rstrip("\n") + "\n"


def _escribir_informe(carpeta_salida: Path, mes: str, contenido: str) -> bool:
    """Deja el informe de ESTA corrida en disco. Devuelve si fue en su nombre.

    Si `_validacion.txt` no se puede escribir, el que quedó en disco es el de
    una corrida anterior y describe otra cosa. Escribirlo igual en otro lado
    es mejor que dejar al dueño con el viejo y sin saberlo, así que se escribe
    en un nombre que dice a la vista cuál de los dos hay que leer.
    """
    try:
        (carpeta_salida / NOMBRE_INFORME).write_text(contenido, encoding="utf-8")
        return True
    except PermissionError:
        print(f"ERROR: no se pudo escribir output/{mes}/{NOMBRE_INFORME}: permiso denegado.")
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


def _avisos_agrupados_por_dev(
    leidos: list[Leido], anio: int, mes: int
) -> tuple[list[str], int]:
    """Los avisos de cada persona, bajo su nombre, en un solo informe."""
    del_mes: dict[str, list[Registro]] = {}
    fuera: dict[str, list[Registro]] = {}
    for leido in leidos:
        for registro in leido.del_mes:
            del_mes.setdefault(nombre_para_mostrar(registro), []).append(registro)
        for registro in leido.descartados:
            fuera.setdefault(nombre_para_mostrar(registro), []).append(registro)

    lineas: list[str] = []
    cantidad = 0
    for nombre in sorted(set(del_mes) | set(fuera), key=lambda n: (n.casefold(), n)):
        avisos = avisos_de_desarrollador(
            del_mes.get(nombre, []), fuera.get(nombre, []), anio, mes
        )
        cantidad += len(avisos)
        if avisos:
            lineas.append(f"=== {nombre} ===")
            lineas.extend(f"  {a}" for a in avisos)
            lineas.append("")
    return lineas, cantidad


def _leer_entradas(
    entradas: list[Path], mapeo: Mapeo, anio: int, numero_mes: int, registrar_fallo
) -> list[Leido]:
    """Lee cada export y le completa lo que el mapeo tenga que poner.

    Un archivo que falla no frena a los demás, y eso incluye el archivo al que
    el mapeo no le puede completar el usuario, el mail o el número de
    proyecto: frena ése solo, con el bloque YAML para arreglarlo.
    """
    leidos: list[Leido] = []
    for entrada in entradas:
        try:
            registros = completar_desde_mapeo(leer(entrada), mapeo, entrada.name)
        except (ErrorLectura, ErrorMapeo) as error:
            registrar_fallo(entrada.name, str(error))
            continue
        except (OSError, ValueError) as error:
            registrar_fallo(entrada.name, f"Error al leer {entrada.name}: {error}")
            continue

        if not registros:
            registrar_fallo(entrada.name, _error_de_export_vacio(entrada.name))
            continue

        del_mes, descartados = _separar_por_mes(registros, anio, numero_mes)
        if not del_mes:
            registrar_fallo(
                entrada.name,
                _error_de_mes_equivocado(descartados, anio, numero_mes, entrada.name),
            )
            continue

        leidos.append(Leido(entrada.name, tuple(del_mes), tuple(descartados)))
    return leidos


def procesar_mes(mes: str, raiz: Path) -> int:
    """Procesa todos los exports de input/<mes>/. Devuelve el código de salida."""
    try:
        anio, numero_mes = _parsear_mes(mes)
    except ValueError as error:
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

    # Los formatos de export de Kimai: .xlsx y .csv. '~$' es el archivo de
    # bloqueo que deja Excel cuando el dueño tiene un export abierto.
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
    # Único punto de salida que NO escribe _validacion.txt: si la carpeta no
    # se pudo crear, no se tocó nada de output/ y tampoco hay dónde escribirlo.
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
        print(f"  {entrada}: NO ENTRA AL ARCHIVO")
        print(motivo)
        no_leidos.append((entrada, motivo))

    leidos = _leer_entradas(entradas, mapeo, anio, numero_mes, registrar_fallo)

    for leido in leidos:
        nombres = sorted({nombre_para_mostrar(r) for r in leido.del_mes})
        horas = _horas(sum(segundos_de(r.horas) for r in leido.del_mes))
        origen = _origenes_de(list(leido.del_mes))
        print(
            f"  {leido.archivo}  ->  {', '.join(nombres)}"
            f"  ({len(leido.del_mes)} registro/s, {horas:.2f} h)  [{origen}]"
        )
        if ORIGEN_RESUMEN_MENSUAL in origen:
            print(
                "    OJO: sus filas van sin descripción y sin hora de inicio. "
                "Está en el informe."
            )

    registros = [r for leido in leidos for r in leido.del_mes]
    detalle = construir(registros, mapeo)
    segundos_esperados = sum(segundos_de(r.horas) for r in registros)

    nombre_limpio = nombre_de_archivo(mapeo.archivo_salida, anio, numero_mes)
    nombre_salida = (
        nombre_incompleto(nombre_limpio, len(no_leidos))
        if no_leidos
        else nombre_limpio
    )

    generado: str | None = None
    nota_de_apartado = ""
    if not leidos:
        # Sin un solo export legible no hay archivo que generar: uno vacío con
        # nombre de entregable es exactamente lo que este cambio busca evitar.
        nota_de_apartado = _apartar_excel_previo(carpeta_salida / nombre_limpio)
        print("  No se generó ningún archivo: no se pudo leer ningún export.")
        if nota_de_apartado:
            print(nota_de_apartado.strip())
    else:
        destino = carpeta_salida / nombre_salida
        # Si el consolidado sale marcado como incompleto, el limpio de una
        # corrida anterior no puede quedar ahí: es el que el dueño adjuntaría
        # sin pensarlo.
        nota_de_apartado = (
            _apartar_excel_previo(carpeta_salida / nombre_limpio) if no_leidos else ""
        )
        try:
            _escribir_atomico(detalle, destino, segundos_esperados)
            generado = nombre_salida
        except ErrorIntegridad as error:
            registrar_fallo(nombre_salida, str(error))
        except PermissionError:
            registrar_fallo(
                nombre_salida,
                f"No se pudo escribir {nombre_salida}: permiso denegado.\n"
                f"  Es probable que tengas ese archivo abierto. Cerralo y "
                f"volvé a correr.",
            )
        except (OSError, ValueError) as error:
            registrar_fallo(
                nombre_salida, f"Error al generar {nombre_salida}: {error}"
            )
        if nota_de_apartado:
            print(nota_de_apartado.strip())

    if generado is not None:
        print(
            f"  Archivo para el partner: {generado}"
            f"  ({len(detalle.filas)} fila/s, {detalle.horas:.2f} h)"
        )

    avisos_generales = _avisos_de_numeros_ambiguos(detalle)
    avisos_por_dev, cantidad_avisos = _avisos_agrupados_por_dev(
        leidos, anio, numero_mes
    )
    cantidad_avisos += len(avisos_generales)

    ajenos = _xlsx_que_esta_corrida_no_genero(
        carpeta_salida, [generado] if generado else []
    )
    contenido = _resumen_de_validacion(
        mes,
        leidos,
        no_leidos,
        detalle if leidos else None,
        generado,
        nota_de_apartado,
        ajenos,
        avisos_generales,
        avisos_por_dev,
    )
    informe_en_su_nombre = _escribir_informe(carpeta_salida, mes, contenido)

    if ajenos:
        print(
            f"  OJO: en output/{mes}/ hay {len(ajenos)} .xlsx que esta corrida "
            "NO generó. No los envíes; están listados en el informe."
        )
    print(
        f"{len(leidos)} export/s leído/s, {len(no_leidos)} no leído/s, "
        f"{cantidad_avisos} aviso/s en output/{mes}/{NOMBRE_INFORME}"
    )
    if no_leidos:
        print(
            "  El archivo salió marcado como INCOMPLETO: NO lo envíes hasta "
            "corregir lo de arriba y volver a correr."
        )

    if not informe_en_su_nombre:
        return 1
    # Los .xlsx ajenos NO cambian el código de salida. Que la carpeta tenga
    # archivos de más no significa que la corrida haya fallado: la única
    # acción pendiente es del dueño, sobre su propia carpeta. Un código 1
    # recurrente enseñaría a ignorar el código de salida, que es lo que
    # distingue una corrida incompleta de una completa.
    return 1 if no_leidos else 0


def main(argv: list[str] | None = None) -> int:
    _reconfigurar_salida_utf8()
    argumentos = sys.argv[1:] if argv is None else argv
    if len(argumentos) != 1:
        print("Uso: python -m cunix_horas AAAA-MM")
        print("Ejemplo: python -m cunix_horas 2025-10")
        return 1
    return procesar_mes(argumentos[0], Path.cwd())
