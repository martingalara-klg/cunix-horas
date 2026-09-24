from datetime import date

import pytest
from conftest import FIXTURES

from cunix_horas.lector_kimai import (
    ErrorLectura,
    Registro,
    codigo_de_proyecto,
    leer,
    serial_a_fecha,
)


def test_serial_a_fecha_convierte_el_serial_de_excel():
    assert serial_a_fecha("46265.708333333") == date(2026, 8, 31)
    assert serial_a_fecha(46243.333333333) == date(2026, 8, 9)


def test_codigo_de_proyecto_extrae_lo_que_esta_entre_corchetes():
    texto = "[CO2610170] Aduana-Subastas | Servicio Nacional de Aduanas - Soporte"
    assert codigo_de_proyecto(texto) == "CO2610170"


def test_codigo_de_proyecto_falla_si_no_hay_corchetes():
    with pytest.raises(ErrorLectura, match="sin código"):
        codigo_de_proyecto("Aduana-Subastas")


def test_leer_devuelve_los_24_registros_del_fixture():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert len(registros) == 24
    assert all(isinstance(r, Registro) for r in registros)


def test_leer_suma_las_horas_correctas():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert round(sum(r.horas for r in registros), 2) == 76.5


def test_leer_cubre_el_rango_de_fechas_esperado():
    fechas = [r.fecha for r in leer(FIXTURES / "kimai-mzalazar.xlsx")]
    assert min(fechas) == date(2026, 8, 3)
    assert max(fechas) == date(2026, 8, 31)


def test_leer_extrae_username_actividad_y_codigo():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert {r.username for r in registros} == {"mzalazar"}
    assert {r.actividad for r in registros} == {"Desarrollo"}
    assert {r.cod_proyecto for r in registros} == {
        "CO2610170",
        "CO2510115",
        "PR2510126",
    }


def test_leer_agrupa_las_horas_por_codigo_de_proyecto():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    por_codigo = {}
    for r in registros:
        por_codigo[r.cod_proyecto] = por_codigo.get(r.cod_proyecto, 0.0) + r.horas
    assert round(por_codigo["CO2610170"], 2) == 61.0
    assert round(por_codigo["CO2510115"], 2) == 10.0
    assert round(por_codigo["PR2510126"], 2) == 5.5


def test_leer_falla_con_un_archivo_que_no_es_xlsx(tmp_path):
    falso = tmp_path / "roto.xlsx"
    falso.write_text("esto no es un xlsx", encoding="utf-8")
    with pytest.raises(ErrorLectura, match="no es un archivo"):
        leer(falso)


def test_leer_falla_si_faltan_las_columnas_esperadas(tmp_path):
    import zipfile

    ruta = tmp_path / "sin-columnas.xlsx"
    hoja = (
        '<?xml version="1.0"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1">'
        '<c r="A1" t="inlineStr"><is><t>Otra cosa</t></is></c>'
        "</row></sheetData></worksheet>"
    )
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("xl/worksheets/sheet1.xml", hoja)
    with pytest.raises(ErrorLectura, match="no tiene el formato"):
        leer(ruta)
