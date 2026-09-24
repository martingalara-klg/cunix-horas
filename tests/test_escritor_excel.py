import openpyxl
import pytest
from conftest import FIXTURES

from cunix_horas.agregador import Fila, Reporte
from cunix_horas.escritor_excel import escribir, nombre_de_archivo


def reporte_de_ejemplo(anio=2025, mes=10):
    return Reporte(
        nombre_dev="Franco Dodera",
        nombre_archivo="Dodera",
        anio=anio,
        mes=mes,
        filas=(
            Fila("Club Atlético Talleres", "CRM", "Desarrollo", {4: 2.0, 31: 3.0}),
            Fila("Sistemas - C.UNIX", "Fan Player", "Desarrollo", {31: 4.0}),
        ),
        descartados=(),
    )


@pytest.fixture
def generado(tmp_path):
    destino = tmp_path / "salida.xlsx"
    escribir(reporte_de_ejemplo(), FIXTURES / "plantilla.xlsx", destino)
    return openpyxl.load_workbook(destino).active


def test_nombre_de_archivo():
    assert nombre_de_archivo(reporte_de_ejemplo()) == "Oct Dodera.xlsx"
    assert nombre_de_archivo(reporte_de_ejemplo(mes=1)) == "Jan Dodera.xlsx"


def test_encabezado(generado):
    assert generado["A1"].value == "Franco Dodera"
    assert generado["B1"].value == "Total"
    assert generado["C1"].value == "10/1/2025"
    assert generado["D1"].value == "10/2/2025"
    assert generado["AG1"].value == "10/31/2025"


def test_la_cantidad_de_columnas_sigue_los_dias_del_mes(tmp_path):
    for anio, mes, ultima_columna in [(2025, 2, 30), (2024, 2, 31), (2025, 4, 32)]:
        destino = tmp_path / f"{anio}-{mes}.xlsx"
        escribir(
            Reporte("X", "X", anio, mes, (), ()),
            FIXTURES / "plantilla.xlsx",
            destino,
        )
        hoja = openpyxl.load_workbook(destino).active
        assert hoja.max_column == ultima_columna


def test_jerarquia_de_filas(generado):
    assert generado["A2"].value == "Club Atlético Talleres"
    assert generado["A3"].value == "CRM"
    assert generado["A4"].value == "Desarrollo"
    assert generado["A5"].value == "Sistemas - C.UNIX"
    assert generado["A6"].value == "Fan Player"
    assert generado["A7"].value == "Desarrollo"
    assert generado["A8"].value == "Total"


def test_totales_de_la_columna_b(generado):
    assert generado["B2"].value == 5.0
    assert generado["B3"].value == 5.0
    assert generado["B4"].value == 5.0
    assert generado["B5"].value == 4.0
    assert generado["B8"].value == 9.0


def test_horas_por_dia(generado):
    assert generado["F3"].value == 2.0   # día 4 -> columna 2+4 = F
    assert generado["AG3"].value == 3.0  # día 31 -> columna 33 = AG
    assert generado["AG6"].value == 4.0
    assert generado["D3"].value is None  # día 2, sin horas -> vacío


def test_la_fila_total_pone_cero_en_los_dias_sin_horas(generado):
    assert generado["C8"].value == 0.0
    assert generado["F8"].value == 2.0
    assert generado["AG8"].value == 7.0


def test_las_filas_de_cliente_no_tienen_horas_por_dia(generado):
    assert generado["F2"].value is None


def test_las_filas_de_proyecto_van_en_negrita(generado):
    assert generado["A3"].font.b is True
    assert generado["A6"].font.b is True
    assert generado["A2"].font.b is not True
    assert generado["A4"].font.b is not True


def test_merge_en_las_filas_de_cliente(generado):
    rangos = {str(r) for r in generado.merged_cells.ranges}
    assert "C2:AG2" in rangos
    assert "C5:AG5" in rangos


def test_anchos_de_columna_copiados_de_la_plantilla(generado):
    assert generado.column_dimensions["A"].width == pytest.approx(34.14, abs=0.01)
    assert generado.column_dimensions["B"].width == pytest.approx(9.29, abs=0.01)


def test_el_nombre_de_la_hoja_es_el_de_la_plantilla(generado):
    assert generado.title == "Worksheet"


def test_las_horas_se_redondean_a_dos_decimales(tmp_path):
    destino = tmp_path / "r.xlsx"
    escribir(
        Reporte("X", "X", 2025, 10, (Fila("C", "P", "A", {1: 1 / 3}),), ()),
        FIXTURES / "plantilla.xlsx",
        destino,
    )
    hoja = openpyxl.load_workbook(destino).active
    assert hoja["C4"].value == 0.33
