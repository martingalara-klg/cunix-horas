"""Orquestación: de input/<mes>/*.xlsx a output/<mes>/*.xlsx."""
from __future__ import annotations

import re
import sys
from pathlib import Path

from cunix_horas.agregador import agregar
from cunix_horas.escritor_excel import escribir, nombre_de_archivo
from cunix_horas.lector_kimai import ErrorLectura, leer
from cunix_horas.mapeo import ErrorMapeo, Mapeo
from cunix_horas.validador import validar

FORMATO_MES = re.compile(r"^(\d{4})-(\d{2})$")


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


def _limpiar_xlsx_previos(carpeta: Path) -> None:
    """Saca de output/<mes>/ los .xlsx que dejó una corrida anterior.

    Si no se limpian, una corrida que falla deja ahí el Excel de la corrida
    previa: mismo nombre, aspecto legítimo, datos viejos o incompletos. El
    dueño no tiene cómo distinguirlo del Excel del mes. Limpiando antes de
    generar, output/ nunca queda con una mezcla de dos corridas.
    """
    for archivo in sorted(carpeta.glob("*.xlsx")):
        if archivo.name.startswith("~$") or not archivo.is_file():
            continue
        archivo.unlink()


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
    try:
        carpeta_salida.mkdir(parents=True, exist_ok=True)
        _limpiar_xlsx_previos(carpeta_salida)
    except PermissionError:
        print(f"ERROR: no se pudo preparar output/{mes}: permiso denegado.")
        print(
            "  Es probable que tengas abierto alguno de los Excel de la corrida "
            "anterior. Cerralos y volvé a correr."
        )
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

    for entrada in entradas:
        destino: Path | None = None
        try:
            registros = leer(entrada)
            reporte = agregar(registros, mapeo, anio, numero_mes, entrada.name)
            destino = carpeta_salida / nombre_de_archivo(reporte)
            if destino in destinos_generados:
                registrar_fallo(
                    entrada.name,
                    f"{destino.name} ya fue generado en esta corrida a partir de "
                    f"{destinos_generados[destino]}.\n"
                    f"Dejá en input/{mes}/ un solo export por desarrollador "
                    "y volvé a correr.",
                )
                continue
            escribir(reporte, plantilla, destino)
        except (ErrorLectura, ErrorMapeo) as error:
            # En consola, sin indentación propia: el mensaje (en especial el
            # bloque YAML sugerido para mapeo.yaml) ya trae el sangrado
            # pegable listo.
            print(f"  {entrada.name}: NO GENERADO")
            print(str(error))
            no_generados.append((entrada.name, str(error)))
            continue
        except PermissionError:
            nombre_destino = destino.name if destino is not None else entrada.name
            registrar_fallo(
                entrada.name,
                f"No se pudo escribir {nombre_destino}: permiso denegado.\n"
                "Es probable que tengas ese Excel abierto. Cerralo y volvé a correr.",
            )
            continue
        except (OSError, ValueError) as error:
            registrar_fallo(
                entrada.name,
                f"Error al generar el Excel de {entrada.name}: {error}",
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
