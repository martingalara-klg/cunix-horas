"""El período, los nombres de archivo y el formato de los números."""
import pytest

from cunix_horas.anexos import (
    ErrorPeriodo,
    Periodo,
    formato_horas,
    formato_importe,
    marca_de_incompleto,
    nombre_incompleto,
)


def test_el_periodo_se_lee_de_aaaa_mm():
    periodo = Periodo.parsear("2026-09")
    assert (periodo.anio, periodo.mes) == (2026, 9)
    assert periodo.carpeta == "2026-09"
    assert periodo.dias == 30


@pytest.mark.parametrize("texto", ["2026/09", "sep-2026", "202609", "", "2026-9"])
def test_un_periodo_mal_escrito_falla_en_espanol(texto):
    with pytest.raises(ErrorPeriodo, match="AAAA-MM"):
        Periodo.parsear(texto)


def test_un_mes_fuera_de_rango_falla():
    with pytest.raises(ErrorPeriodo, match="fuera de rango"):
        Periodo.parsear("2026-13")


def test_el_texto_del_periodo_va_en_espanol_sin_depender_del_locale():
    """Nunca strftime('%B'): en Windows sale en inglés."""
    assert Periodo(2026, 8).texto == "Agosto 2026"
    assert Periodo(2026, 9).texto == "Septiembre 2026"
    assert Periodo(2026, 12).texto == "Diciembre 2026"


def test_el_periodo_sabe_que_fechas_le_pertenecen():
    import datetime as dt

    periodo = Periodo(2026, 8)
    assert periodo.contiene(dt.date(2026, 8, 31))
    assert not periodo.contiene(dt.date(2026, 9, 1))


def test_la_marca_de_incompleto_concuerda_en_numero():
    assert "FALTA 1 DESARROLLADOR " in marca_de_incompleto(1)
    assert "FALTAN 2 DESARROLLADORES " in marca_de_incompleto(2)


def test_la_marca_va_antes_de_la_extension():
    assert nombre_incompleto("Anexo-II-A.xlsx", 1) == (
        "Anexo-II-A (INCOMPLETO - FALTA 1 DESARROLLADOR - NO ENVIAR).xlsx"
    )


def test_las_horas_salen_con_coma_y_un_decimal():
    assert formato_horas(306.0) == "306,0"
    assert formato_horas(76.5) == "76,5"


def test_las_horas_redondean_medio_hacia_arriba():
    """Con el redondeo por defecto de Python, 2,25 daría 2,2."""
    assert formato_horas(2.25) == "2,3"


def test_los_importes_llevan_punto_de_miles_y_dos_decimales():
    assert formato_importe(3706.0) == "3.706,00"
    assert formato_importe(17) == "17,00"
    assert formato_importe(1234567.5) == "1.234.567,50"
