"""Tests del lector de timesheet plano en CSV."""
import csv
from datetime import date
from pathlib import Path

import pytest
from conftest import FIXTURES

from cunix_horas.lector_kimai import ErrorLectura, Registro, leer

FRANCO = FIXTURES / "kimai-timesheet-csv-franco.csv"


def _columnas_de(origen: Path) -> list[str]:
    with origen.open(encoding="utf-8-sig", newline="") as archivo:
        return list(csv.DictReader(archivo).fieldnames)


def _reescribir(
    origen: Path, destino: Path, orden: list[str] | None = None, bom: bool = False
) -> Path:
    """Copia el CSV con las columnas en otro orden y/o con BOM al principio."""
    with origen.open(encoding="utf-8-sig", newline="") as archivo:
        lector = csv.DictReader(archivo)
        columnas = list(orden or lector.fieldnames)
        filas = list(lector)
    codificacion = "utf-8-sig" if bom else "utf-8"
    with destino.open("w", encoding=codificacion, newline="") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=columnas, extrasaction="ignore")
        escritor.writeheader()
        escritor.writerows(filas)
    return destino


def test_lee_los_15_registros_del_csv():
    registros = leer(FRANCO)
    assert len(registros) == 15
    assert all(isinstance(r, Registro) for r in registros)


def test_suma_62_horas_y_media():
    assert round(sum(r.horas for r in leer(FRANCO)), 2) == 62.5


def test_trae_el_username_y_el_codigo_de_proyecto():
    registros = leer(FRANCO)
    assert {r.username for r in registros} == {"franco"}
    assert {r.cod_proyecto for r in registros} == {"PR2510126"}


def test_todas_las_fechas_caen_en_agosto_de_2026():
    fechas = [r.fecha for r in leer(FRANCO)]
    assert all(f.year == 2026 and f.month == 8 for f in fechas)
    assert min(fechas) == date(2026, 8, 3)
    assert max(fechas) == date(2026, 8, 31)


def test_la_duracion_viene_en_horas_y_minutos():
    """'2:00' son 2 horas; nada de fracción de día como en el .xlsx."""
    del_31 = [r for r in leer(FRANCO) if r.fecha == date(2026, 8, 31)]
    assert round(sum(r.horas for r in del_31), 2) == 2.0


def test_lee_por_nombre_de_columna_y_no_por_posicion(tmp_path):
    """Con las columnas dadas vuelta, el resultado tiene que ser el mismo."""
    orden = list(reversed(_columnas_de(FRANCO)))
    revuelto = _reescribir(FRANCO, tmp_path / "revuelto.csv", orden=orden)
    assert leer(revuelto) == leer(FRANCO)


def test_tolera_el_bom_al_principio_del_archivo(tmp_path):
    con_bom = _reescribir(FRANCO, tmp_path / "con-bom.csv", bom=True)
    assert con_bom.read_bytes().startswith(b"\xef\xbb\xbf")
    assert leer(con_bom) == leer(FRANCO)


def test_falla_si_falta_una_columna_esperada(tmp_path):
    ruta = tmp_path / "sin-duration.csv"
    ruta.write_text(
        "Date,User,Project,Activity\n2026-08-31,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura, match="no tiene el formato"):
        leer(ruta)


def test_una_duracion_ilegible_da_un_error_en_espanol(tmp_path):
    ruta = tmp_path / "roto.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "2026-08-31,dos horas,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "roto.csv" in mensaje
    assert "fila 2" in mensaje
    assert "no es una duración" in mensaje


def test_una_fecha_ilegible_da_un_error_en_espanol(tmp_path):
    ruta = tmp_path / "roto.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "31/08/2026,2:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    assert "no es una fecha" in str(excepcion.value)
