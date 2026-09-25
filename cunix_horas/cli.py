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
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    print(f"Procesando input/{mes}/ ...")
    avisos_totales: list[str] = []
    generados = 0
    hubo_errores = False
    # Nombre de salida -> archivo de entrada que lo generó, para detectar
    # colisiones (dos exports que resuelven al mismo "<Mes> <Apellido>.xlsx").
    destinos_generados: dict[Path, str] = {}

    for entrada in entradas:
        destino: Path | None = None
        try:
            registros = leer(entrada)
            reporte = agregar(registros, mapeo, anio, numero_mes, entrada.name)
            destino = carpeta_salida / nombre_de_archivo(reporte)
            if destino in destinos_generados:
                print(f"  {entrada.name}: NO GENERADO")
                print(
                    f"    {destino.name} ya fue generado en esta corrida a partir de "
                    f"{destinos_generados[destino]}."
                )
                print(
                    f"    Dejá en input/{mes}/ un solo export por desarrollador "
                    "y volvé a correr."
                )
                hubo_errores = True
                continue
            escribir(reporte, plantilla, destino)
        except (ErrorLectura, ErrorMapeo) as error:
            # Sin indentación propia: el mensaje (en especial el bloque YAML
            # sugerido para mapeo.yaml) ya trae el sangrado pegable listo.
            print(f"  {entrada.name}: NO GENERADO")
            print(str(error))
            hubo_errores = True
            continue
        except PermissionError:
            nombre_destino = destino.name if destino is not None else entrada.name
            print(f"  {entrada.name}: NO GENERADO")
            print(f"    No se pudo escribir {nombre_destino}: permiso denegado.")
            print(
                "    Es probable que tengas ese Excel abierto. "
                "Cerralo y volvé a correr."
            )
            hubo_errores = True
            continue
        except (OSError, ValueError) as error:
            print(f"  {entrada.name}: NO GENERADO")
            print(f"    Error al generar el Excel de {entrada.name}: {error}")
            hubo_errores = True
            continue

        destinos_generados[destino] = entrada.name
        clientes = len({f.cliente for f in reporte.filas})
        print(
            f"  {entrada.name}  ->  {destino.name}"
            f"  ({reporte.total:.1f} h, {clientes} cliente/s)"
        )
        generados += 1

        avisos = validar(reporte)
        if avisos:
            avisos_totales.append(f"=== {destino.name} ===")
            avisos_totales.extend(f"  {a}" for a in avisos)
            avisos_totales.append("")

    cantidad_avisos = sum(1 for a in avisos_totales if a.startswith("  "))
    (carpeta_salida / "_validacion.txt").write_text(
        "\n".join(avisos_totales) if avisos_totales else "Sin avisos.\n",
        encoding="utf-8",
    )

    print(
        f"{generados} archivo/s generado/s, {cantidad_avisos} aviso/s en "
        f"output/{mes}/_validacion.txt"
    )
    return 1 if hubo_errores else 0


def main(argv: list[str] | None = None) -> int:
    _reconfigurar_salida_utf8()
    argumentos = sys.argv[1:] if argv is None else argv
    if len(argumentos) != 1:
        print("Uso: python -m cunix_horas AAAA-MM")
        print("Ejemplo: python -m cunix_horas 2025-10")
        return 1
    return procesar_mes(argumentos[0], Path.cwd())
