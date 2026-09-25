"""Orquestación: de input/<mes>/*.xlsx a output/<mes>/*.xlsx."""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

from cunix_horas.agregador import Reporte, agregar, resolver_identidad
from cunix_horas.escritor_excel import escribir, nombre_de_archivo_de
from cunix_horas.lector_kimai import ErrorLectura, leer
from cunix_horas.mapeo import ErrorMapeo, Mapeo
from cunix_horas.validador import validar

FORMATO_MES = re.compile(r"^(\d{4})-(\d{2})$")

# Marca del Excel que quedó de una corrida anterior y que esta corrida no pudo
# reemplazar. Va en el nombre para que sea imposible confundirlo con el del mes.
SUFIJO_CORRIDA_ANTERIOR = " (CORRIDA ANTERIOR - NO ENVIAR)"


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


def _escribir_atomico(reporte: Reporte, plantilla: Path, destino: Path) -> None:
    """Genera el Excel en un temporal y recién ahí lo mueve sobre `destino`.

    `os.replace` es atómico y pisa el destino existente, así que nunca hay una
    ventana en la que `destino` sea un Excel a medio escribir, y una falla
    durante la generación no toca el archivo que ya estaba.

    El temporal vive en la misma carpeta que el destino: mover entre volúmenes
    no es atómico.
    """
    temporal = destino.with_name(f"~tmp-{os.getpid()}-{destino.name}")
    try:
        escribir(reporte, plantilla, temporal)
        os.replace(temporal, destino)
    finally:
        try:
            temporal.unlink(missing_ok=True)
        except OSError:
            pass


def _apartar_excel_previo(destino: Path) -> str:
    """Saca de en medio el Excel viejo que haya con el nombre `destino`.

    Se llama cuando un archivo de entrada falló y su Excel NO se generó. Si en
    output/ quedó el del mes pasado con ese mismo nombre, no puede seguir ahí
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
            f"abierto). NO lo envíes: no es el Excel de esta corrida."
        )
    return (
        f"\n  El {destino.name} que había quedado de una corrida anterior se "
        f"renombró a {apartado(intento).name} para que no se confunda con el "
        f"del mes."
    )


def _resumen_de_validacion(
    mes: str,
    generados: list[str],
    no_generados: list[tuple[str, str]],
    avisos: list[str],
) -> str:
    """Arma el contenido de _validacion.txt: primero el estado de la corrida.

    Este archivo es el informe que el dueño lee antes de mandar nada, así que
    tiene que alcanzar por sí solo para decidir. Nunca puede decir "Sin
    avisos." si hubo archivos que no se generaron.
    """
    lineas: list[str] = [f"Corrida de output/{mes}/", ""]

    lineas.append(f"{len(generados)} Excel generado/s:")
    lineas.extend(f"  - {nombre}" for nombre in generados)
    if not generados:
        lineas.append("  (ninguno)")
    lineas.append("")

    if no_generados:
        lineas.append(
            f"{len(no_generados)} archivo/s NO GENERADO/S. "
            "Esos Excel NO están en esta carpeta:"
        )
        for entrada, motivo in no_generados:
            lineas.append(f"  - {entrada}:")
            lineas.extend(f"    {linea}" for linea in motivo.splitlines())
        lineas.append("")
        lineas.append(
            "Corregí lo de arriba y volvé a correr ANTES de enviar nada: faltan "
            "los Excel listados como NO GENERADOS, y los que sí están cubren "
            "sólo a los desarrolladores listados como generados."
        )
        lineas.append("")

    lineas.append("--- Avisos de validación de los Excel generados ---")
    lineas.append("")
    if avisos:
        lineas.extend(avisos)
    elif no_generados:
        # Nunca "Sin avisos." a secas cuando hubo errores: sería el mensaje
        # más tranquilizador posible al lado de una corrida que falló.
        lineas.append(
            "Ningún aviso sobre los Excel que sí se generaron, pero la corrida "
            "NO está completa: mirá la lista de NO GENERADOS de arriba."
        )
    else:
        lineas.append("Sin avisos.")
    return "\n".join(lineas).rstrip("\n") + "\n"


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

    plantilla = raiz / "templates" / "plantilla.xlsx"
    if not plantilla.is_file():
        print(f"ERROR: falta la plantilla {plantilla}")
        return 1

    entradas = sorted(
        p for p in carpeta_entrada.glob("*.xlsx") if not p.name.startswith("~$")
    )
    if not entradas:
        print(f"ERROR: no hay ningún .xlsx en input/{mes}")
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
    avisos_totales: list[str] = []
    generados: list[str] = []
    no_generados: list[tuple[str, str]] = []
    # Nombre de salida -> archivo de entrada que lo generó, para detectar
    # colisiones (dos exports que resuelven al mismo "<Mes> <Apellido>.xlsx").
    destinos_generados: dict[Path, str] = {}

    def registrar_fallo(entrada: str, motivo: str) -> None:
        """Deja el fallo en consola y en la lista que va a _validacion.txt."""
        print(f"  {entrada}: NO GENERADO")
        for linea in motivo.splitlines():
            print(f"    {linea}")
        no_generados.append((entrada, motivo))

    def registrar_fallo_de_mapeo(entrada: str, error: Exception, extra: str) -> None:
        """Como registrar_fallo, pero sin indentar el mensaje en consola.

        El mensaje de ErrorMapeo (en especial el bloque YAML sugerido para
        mapeo.yaml) ya trae el sangrado pegable listo.
        """
        print(f"  {entrada}: NO GENERADO")
        print(str(error) + extra)
        no_generados.append((entrada, str(error) + extra))

    for entrada in entradas:
        # Primero se resuelve la identidad del export: eso ya alcanza para
        # saber qué archivo de output/ va a reemplazar esta entrada, antes de
        # que pueda fallar cualquier otra cosa.
        try:
            registros = leer(entrada)
            persona = resolver_identidad(registros, mapeo, entrada.name)
        except (ErrorLectura, ErrorMapeo) as error:
            # Destino todavía desconocido: no hay archivo viejo que apartar.
            registrar_fallo_de_mapeo(entrada.name, error, "")
            continue
        except (OSError, ValueError) as error:
            registrar_fallo(entrada.name, f"Error al leer {entrada.name}: {error}")
            continue

        destino = carpeta_salida / nombre_de_archivo_de(numero_mes, persona.archivo)
        if destino in destinos_generados:
            # El destino lo acaba de generar otra entrada de esta misma
            # corrida: no es un archivo viejo, no se aparta.
            registrar_fallo(
                entrada.name,
                f"{destino.name} ya fue generado en esta corrida a partir de "
                f"{destinos_generados[destino]}.\n"
                f"Dejá en input/{mes}/ un solo export por desarrollador "
                "y volvé a correr.",
            )
            continue

        try:
            reporte = agregar(registros, mapeo, anio, numero_mes, entrada.name)
            _escribir_atomico(reporte, plantilla, destino)
        except (ErrorLectura, ErrorMapeo) as error:
            registrar_fallo_de_mapeo(
                entrada.name, error, _apartar_excel_previo(destino)
            )
            continue
        except PermissionError:
            registrar_fallo(
                entrada.name,
                f"No se pudo escribir {destino.name}: permiso denegado.\n"
                "Es probable que tengas ese Excel abierto. Cerralo y volvé a "
                "correr." + _apartar_excel_previo(destino),
            )
            continue
        except (OSError, ValueError) as error:
            registrar_fallo(
                entrada.name,
                f"Error al generar el Excel de {entrada.name}: {error}"
                + _apartar_excel_previo(destino),
            )
            continue

        destinos_generados[destino] = entrada.name
        clientes = len({f.cliente for f in reporte.filas})
        print(
            f"  {entrada.name}  ->  {destino.name}"
            f"  ({reporte.total_redondeado:.1f} h, {clientes} cliente/s)"
        )
        generados.append(destino.name)

        avisos = validar(reporte)
        if avisos:
            avisos_totales.append(f"=== {destino.name} ===")
            avisos_totales.extend(f"  {a}" for a in avisos)
            avisos_totales.append("")

    cantidad_avisos = sum(1 for a in avisos_totales if a.startswith("  "))
    contenido = _resumen_de_validacion(mes, generados, no_generados, avisos_totales)
    try:
        (carpeta_salida / "_validacion.txt").write_text(contenido, encoding="utf-8")
    except PermissionError:
        print(f"ERROR: no se pudo escribir output/{mes}/_validacion.txt: permiso denegado.")
        print("  Es probable que lo tengas abierto en el Bloc de notas. Cerralo y volvé a correr.")
        return 1
    except OSError as error:
        print(f"ERROR: no se pudo escribir output/{mes}/_validacion.txt: {error}")
        return 1

    print(
        f"{len(generados)} archivo/s generado/s, {len(no_generados)} no generado/s, "
        f"{cantidad_avisos} aviso/s en output/{mes}/_validacion.txt"
    )
    return 1 if no_generados else 0


def main(argv: list[str] | None = None) -> int:
    _reconfigurar_salida_utf8()
    argumentos = sys.argv[1:] if argv is None else argv
    if len(argumentos) != 1:
        print("Uso: python -m cunix_horas AAAA-MM")
        print("Ejemplo: python -m cunix_horas 2025-10")
        return 1
    return procesar_mes(argumentos[0], Path.cwd())
