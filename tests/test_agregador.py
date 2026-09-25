from datetime import date

import pytest
from conftest import FIXTURES

from cunix_horas.agregador import agregar
from cunix_horas.lector_kimai import Registro, leer
from cunix_horas.mapeo import ErrorMapeo, Mapeo


def mapeo():
    return Mapeo.cargar(FIXTURES / "mapeo-test.yaml")


def reg(dia, horas, codigo="CO2610170", actividad="Desarrollo", mes=8):
    return Registro(
        fecha=date(2026, mes, dia),
        horas=horas,
        username="mzalazar",
        cod_proyecto=codigo,
        actividad=actividad,
    )


def test_agrega_las_horas_del_mismo_dia_y_proyecto():
    reporte = agregar([reg(3, 2.0), reg(3, 1.5)], mapeo(), 2026, 8, "x.xlsx")
    assert len(reporte.filas) == 1
    assert reporte.filas[0].horas_por_dia == {3: 3.5}


def test_separa_por_proyecto():
    reporte = agregar(
        [reg(3, 2.0, "CO2610170"), reg(3, 1.0, "CO2510115")], mapeo(), 2026, 8, "x.xlsx"
    )
    assert len(reporte.filas) == 2
    assert {f.proyecto for f in reporte.filas} == {"Subastas", "SIAC-OIRS"}


def test_separa_por_actividad_dentro_del_mismo_proyecto():
    reporte = agregar(
        [reg(3, 2.0, actividad="Desarrollo"), reg(3, 1.0, actividad="Testing")],
        mapeo(),
        2026,
        8,
        "x.xlsx",
    )
    assert len(reporte.filas) == 2
    assert {f.actividad for f in reporte.filas} == {"Desarrollo", "Testing"}


def test_las_filas_salen_ordenadas_por_cliente():
    reporte = agregar(
        [reg(3, 1.0, "PR2510126"), reg(3, 1.0, "CO2510115"), reg(3, 1.0, "CO2610170")],
        mapeo(),
        2026,
        8,
        "x.xlsx",
    )
    clientes = [f.cliente for f in reporte.filas]
    assert clientes == sorted(clientes)


def test_el_total_de_la_fila_suma_sus_dias():
    reporte = agregar([reg(3, 2.0), reg(5, 4.5)], mapeo(), 2026, 8, "x.xlsx")
    assert reporte.filas[0].total == 6.5


def test_el_total_del_reporte_suma_todas_las_filas():
    reporte = agregar(
        [reg(3, 2.0, "CO2610170"), reg(4, 1.0, "CO2510115")], mapeo(), 2026, 8, "x.xlsx"
    )
    assert reporte.total == 3.0


def test_total_del_dia_suma_todas_las_filas_de_ese_dia():
    reporte = agregar(
        [reg(3, 2.0, "CO2610170"), reg(3, 1.0, "CO2510115"), reg(4, 5.0)],
        mapeo(),
        2026,
        8,
        "x.xlsx",
    )
    assert reporte.total_del_dia(3) == 3.0
    assert reporte.total_del_dia(4) == 5.0
    assert reporte.total_del_dia(10) == 0.0


def test_descarta_los_registros_fuera_del_mes():
    reporte = agregar([reg(3, 2.0), reg(15, 9.0, mes=7)], mapeo(), 2026, 8, "x.xlsx")
    assert reporte.total == 2.0
    assert len(reporte.descartados) == 1
    assert reporte.descartados[0].fecha == date(2026, 7, 15)


def test_dias_del_mes():
    assert agregar([], mapeo(), 2026, 2, "x.xlsx").dias_del_mes == 28
    assert agregar([], mapeo(), 2024, 2, "x.xlsx").dias_del_mes == 29
    assert agregar([], mapeo(), 2025, 4, "x.xlsx").dias_del_mes == 30
    assert agregar([], mapeo(), 2025, 10, "x.xlsx").dias_del_mes == 31


def test_toma_el_nombre_del_dev_del_mapeo():
    reporte = agregar([reg(3, 2.0)], mapeo(), 2026, 8, "x.xlsx")
    assert reporte.nombre_dev == "Matias Zalazar"
    assert reporte.nombre_archivo == "Zalazar"


def test_proyecto_sin_mapear_sugiere_el_alias_no_el_codigo(tmp_path):
    """El texto original de Kimai (columna Project) debe llegar hasta el error.

    Antes, agregar() llamaba a resolver_proyecto con texto_kimai="" porque
    Registro no lo conservaba, y el mensaje repetía el código como alias.
    """
    ruta_mapeo = tmp_path / "mapeo.yaml"
    ruta_mapeo.write_text(
        "personas:\n"
        "  mzalazar:\n"
        '    nombre: "Matias Zalazar"\n'
        '    archivo: "Zalazar"\n'
        "proyectos:\n"
        "  CO2610170:\n"
        '    cliente: "Servicio Nacional de Aduanas"\n'
        '    proyecto: "Subastas"\n',
        encoding="utf-8",
    )
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    with pytest.raises(ErrorMapeo) as excepcion:
        agregar(registros, Mapeo.cargar(ruta_mapeo), 2026, 8, "kimai-mzalazar.xlsx")

    mensaje = str(excepcion.value)
    assert "CO2510115" in mensaje
    assert "ISPCH-SopEvo-SIAC" in mensaje
    assert 'proyecto: "ISPCH-SopEvo-SIAC"' in mensaje


def test_sobre_el_fixture_real_cierran_los_totales():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    reporte = agregar(registros, mapeo(), 2026, 8, "kimai-mzalazar.xlsx")
    assert round(reporte.total, 2) == 76.5
    assert round(sum(f.total for f in reporte.filas), 2) == 76.5
    suma_dias = sum(reporte.total_del_dia(d) for d in range(1, reporte.dias_del_mes + 1))
    assert round(suma_dias, 2) == 76.5
    assert {f.proyecto for f in reporte.filas} == {"Subastas", "SIAC-OIRS", "SELICO"}
