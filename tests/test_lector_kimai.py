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


def test_leer_conserva_el_texto_original_de_proyecto():
    """El texto completo de la columna Project viaja hasta agregador.agregar()."""
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    textos = {r.cod_proyecto: r.texto_proyecto for r in registros}
    assert textos["CO2610170"].startswith("[CO2610170] Aduana-Subastas")
    assert textos["CO2510115"].startswith("[CO2510115] ISPCH-SopEvo-SIAC")


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


# --- Regresión: un valor no numérico no debe llegar como ValueError crudo ---
# (el dueño leía "could not convert string to float: '2026-08-31'" y no sabía
# qué archivo, qué fila ni qué columna mirar.)

import zipfile

NS_HOJA = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ENCABEZADOS = {"A": "Date", "D": "Duration", "F": "User", "J": "Project", "K": "Activity"}


def _xlsx_de_kimai(ruta, filas_de_datos):
    """Arma un .xlsx mínimo con el formato de export de Kimai (inline strings)."""

    def fila_xml(nro, celdas):
        cel = "".join(
            f'<c r="{col}{nro}" t="inlineStr"><is><t>{texto}</t></is></c>'
            for col, texto in sorted(celdas.items())
        )
        return f'<row r="{nro}">{cel}</row>'

    filas = [fila_xml(1, ENCABEZADOS)]
    filas += [fila_xml(n, celdas) for n, celdas in enumerate(filas_de_datos, start=2)]
    hoja = (
        f'<?xml version="1.0"?><worksheet xmlns="{NS_HOJA}">'
        f'<sheetData>{"".join(filas)}</sheetData></worksheet>'
    )
    with zipfile.ZipFile(ruta, "w") as archivo:
        archivo.writestr("xl/worksheets/sheet1.xml", hoja)
    return ruta


def _fila(fecha="46265.708333333", duracion="0.0416666666", usuario="mzalazar"):
    return {
        "A": fecha,
        "D": duracion,
        "F": usuario,
        "J": "[CO2610170] Subastas | largo",
        "K": "Desarrollo",
    }


def test_una_fecha_no_numerica_da_un_error_en_espanol(tmp_path):
    ruta = _xlsx_de_kimai(tmp_path / "kimai.xlsx", [_fila(fecha="2026-08-31")])
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "kimai.xlsx" in mensaje
    assert "fila 2" in mensaje
    assert "columna A" in mensaje
    assert "no es una fecha" in mensaje


def test_una_duracion_no_numerica_da_un_error_en_espanol(tmp_path):
    ruta = _xlsx_de_kimai(tmp_path / "kimai.xlsx", [_fila(), _fila(duracion="8 hs")])
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "fila 3" in mensaje
    assert "columna D" in mensaje
    assert "no es una duración numérica" in mensaje


def test_un_export_sin_filas_de_datos_se_lee_como_lista_vacia(tmp_path):
    """El lector no opina: es agregar() quien frena el export vacío."""
    assert leer(_xlsx_de_kimai(tmp_path / "kimai.xlsx", [])) == []


# --- El despachador: leer() elige el lector según el archivo ---------------


def test_el_xlsx_plano_de_otro_dev_se_sigue_leyendo_igual():
    """El formato que ya andaba no cambia: segundo caso real del timesheet."""
    registros = leer(FIXTURES / "kimai-timesheet-xlsx-lcarducci.xlsx")
    assert len(registros) == 8
    assert round(sum(r.horas for r in registros), 2) == 8.0
    assert {r.username for r in registros} == {"lcarducci"}
    assert {r.cod_proyecto for r in registros} == {"AD2690002"}
    assert {r.actividad for r in registros} == {"Desarrollo", "Gestión"}
    assert all(r.fecha.year == 2026 and r.fecha.month == 8 for r in registros)


def test_un_xlsx_de_formato_desconocido_dice_que_encontro_y_que_espera(tmp_path):
    ruta = tmp_path / "otra-cosa.xlsx"
    hoja = (
        '<?xml version="1.0"?>'
        f'<worksheet xmlns="{NS_HOJA}"><sheetData><row r="1">'
        '<c r="A1" t="inlineStr"><is><t>Resumen de horas</t></is></c>'
        '<c r="B1" t="inlineStr"><is><t>Horas</t></is></c>'
        "</row></sheetData></worksheet>"
    )
    with zipfile.ZipFile(ruta, "w") as archivo:
        archivo.writestr("xl/worksheets/sheet1.xml", hoja)

    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "otra-cosa.xlsx" in mensaje
    assert "Resumen de horas" in mensaje
    assert "Horas" in mensaje
    assert "A1='Date'" in mensaje
    assert "B1='Total'" in mensaje
    assert ".csv" in mensaje


# --- I4: en el .xlsx plano, una fila con horas pero sin fecha tampoco se saltea


def test_una_fila_del_xlsx_con_duracion_pero_sin_fecha_falla(tmp_path):
    sin_fecha = _fila()
    del sin_fecha["A"]
    ruta = _xlsx_de_kimai(tmp_path / "kimai.xlsx", [_fila(), sin_fecha])
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "kimai.xlsx" in mensaje
    assert "fila 3" in mensaje
    assert "columna A" in mensaje


def test_una_fila_del_xlsx_con_todas_las_celdas_en_blanco_se_saltea(tmp_path):
    """Ahí no hay horas que perder: saltearla está bien."""
    ruta = _xlsx_de_kimai(tmp_path / "kimai.xlsx", [_fila(), {"D": "", "F": ""}])
    assert len(leer(ruta)) == 1


# --- M7: `EPOCA_EXCEL` no lo importaba nadie -------------------------------


def test_lector_kimai_no_reexporta_la_epoca_de_excel():
    """El re-export estaba muerto: `serial_a_fecha` es lo único que la usa."""
    from cunix_horas import lector_kimai

    assert "EPOCA_EXCEL" not in lector_kimai.__all__
    assert not hasattr(lector_kimai, "EPOCA_EXCEL")
