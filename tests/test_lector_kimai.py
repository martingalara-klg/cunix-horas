from datetime import date, time

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


def _xlsx_de_kimai(ruta, filas_de_datos, encabezados=None):
    """Arma un .xlsx mínimo con el formato de export de Kimai (inline strings)."""

    def fila_xml(nro, celdas):
        cel = "".join(
            f'<c r="{col}{nro}" t="inlineStr"><is><t>{texto}</t></is></c>'
            for col, texto in sorted(celdas.items())
        )
        return f'<row r="{nro}">{cel}</row>'

    filas = [fila_xml(1, encabezados or ENCABEZADOS)]
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


# --- El detalle de cada registro -------------------------------------------
# La hoja «Detalle» del Anexo II-A lleva una fila por registro, con la hora de
# inicio, el nombre para mostrar, el mail, la descripción y el número de
# proyecto. Kimai los trae y el lector los conserva, aunque los anexos de hoy
# no escriban todos los campos.

ENCABEZADOS_CON_DETALLE = {
    **ENCABEZADOS,
    "B": "From",
    "E": "Name",
    "G": "E-mail",
    "I": "Customer",
    "L": "Description",
    "R": "Project number",
}


def _fila_con_detalle(descripcion="Ticket R-000001", desde="09:30"):
    fila = _fila()
    fila.update(
        {
            "B": desde,
            "E": "Matias Zalazar",
            "G": "matias.zalazar@cunix.net",
            "I": "[616050001] Instituto de Salud Publica de Chile",
            "L": descripcion,
            "R": "210",
        }
    )
    if descripcion is None:
        del fila["L"]
    return fila


def test_el_primer_registro_del_xlsx_real_trae_todo_el_detalle():
    primero = leer(FIXTURES / "kimai-mzalazar.xlsx")[0]
    assert primero.fecha == date(2026, 8, 31)
    assert primero.hora_inicio == time(17, 0)
    assert primero.nombre == "Matias Zalazar"
    assert primero.username == "mzalazar"
    assert primero.email == "matias.zalazar@cunix.net"
    assert primero.descripcion == "Ticket R-012528"
    assert primero.numero_proyecto == "CO2510115"
    assert primero.texto_cliente == (
        "[616050001] Instituto de Salud Publica de Chile"
    )


def test_el_nombre_para_mostrar_y_el_username_son_dos_campos_distintos():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert {r.username for r in registros} == {"mzalazar"}
    assert {r.nombre for r in registros} == {"Matias Zalazar"}


def test_el_numero_de_proyecto_no_es_el_codigo_entre_corchetes():
    """Luciano: el corchete dice AD2690002 y el Project number de Kimai es 210.

    Son dos campos distintos y los dos hacen falta: el del corchete es la
    clave del mapeo, y el número es un campo propio de Kimai (que en
    septiembre mostró justo 210 para este proyecto). Los anexos de hoy no
    escriben el número, pero el lector no puede confundirlos.
    """
    registros = leer(FIXTURES / "kimai-timesheet-xlsx-lcarducci.xlsx")
    assert {r.cod_proyecto for r in registros} == {"AD2690002"}
    assert {r.numero_proyecto for r in registros} == {"210"}
    assert all(r.numero_proyecto != r.cod_proyecto for r in registros)


def test_el_primer_registro_de_luciano_trae_todo_el_detalle():
    primero = leer(FIXTURES / "kimai-timesheet-xlsx-lcarducci.xlsx")[0]
    assert primero.fecha == date(2026, 8, 31)
    assert primero.hora_inicio == time(11, 0)
    assert primero.nombre == "Luciano Carducci"
    assert primero.username == "lcarducci"
    assert primero.email == "luiciano.carducci@cunix.net"
    assert primero.descripcion.startswith("Revisi")
    assert primero.descripcion.endswith("EFS Aruba.")
    assert primero.numero_proyecto == "210"
    assert primero.texto_cliente == "CUNIX"


def test_un_registro_del_xlsx_sin_descripcion_queda_vacio_y_no_rompe(tmp_path):
    """En el archivo de septiembre del partner 112 de 160 filas no la traen."""
    ruta = _xlsx_de_kimai(
        tmp_path / "sin-descripcion.xlsx",
        [_fila_con_detalle(descripcion=None), _fila_con_detalle(descripcion="")],
        encabezados=ENCABEZADOS_CON_DETALLE,
    )
    registros = leer(ruta)
    assert len(registros) == 2
    assert [r.descripcion for r in registros] == ["", ""]
    assert all(r.nombre == "Matias Zalazar" for r in registros)


def test_un_xlsx_sin_las_columnas_del_detalle_se_sigue_leyendo(tmp_path):
    """Los campos nuevos quedan vacíos; las horas, que es lo que factura, no."""
    ruta = _xlsx_de_kimai(tmp_path / "viejo.xlsx", [_fila()])
    registro = leer(ruta)[0]
    assert registro.horas > 0
    assert registro.hora_inicio is None
    assert (registro.nombre, registro.email, registro.numero_proyecto) == (
        "",
        "",
        "",
    )


def test_una_columna_del_detalle_corrida_de_lugar_falla(tmp_path):
    """Este lector toma las columnas por posición: leerlas igual mezclaría datos."""
    encabezados = {**ENCABEZADOS_CON_DETALLE, "E": "E-mail", "G": "Name"}
    ruta = _xlsx_de_kimai(
        tmp_path / "corrido.xlsx",
        [_fila_con_detalle()],
        encabezados=encabezados,
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "corrido.xlsx" in mensaje
    assert "no están donde se esperaba" in mensaje
    assert "'Name'" in mensaje


def test_una_hora_de_inicio_ilegible_da_un_error_en_espanol(tmp_path):
    """Esa hora va en la celda de fecha del entregable: mal leída mueve el día."""
    ruta = _xlsx_de_kimai(
        tmp_path / "hora-rota.xlsx",
        [_fila_con_detalle(desde="las nueve")],
        encabezados=ENCABEZADOS_CON_DETALLE,
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "hora-rota.xlsx" in mensaje
    assert "fila 2" in mensaje
    assert "columna B" in mensaje
    assert "formato de hora" in mensaje


def test_los_campos_del_detalle_van_al_final_y_con_valor_por_defecto():
    """Las construcciones posicionales que ya existían no se rompen."""
    registro = Registro(date(2026, 8, 3), 1.0, "mzalazar", "CO2510115", "Desarrollo")
    assert registro.texto_proyecto == ""
    assert registro.hora_inicio is None
    assert registro.nombre == ""
    assert registro.email == ""
    assert registro.descripcion == ""
    assert registro.numero_proyecto == ""
    assert registro.texto_cliente == ""
