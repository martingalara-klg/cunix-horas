"""Tests del lector de resumen mensual (.xlsx con B1='Total')."""
import zipfile
from datetime import date

import pytest
from conftest import FIXTURES

from cunix_horas.lector_kimai import ErrorLectura, leer

LAUTARO = FIXTURES / "kimai-resumen-mensual-lautaro.xlsx"
ALEXIS = FIXTURES / "kimai-resumen-mensual-alexis.xlsx"

NS_HOJA = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
LETRAS = [chr(c) for c in range(ord("A"), ord("Z") + 1)] + [
    "A" + chr(c) for c in range(ord("A"), ord("Z") + 1)
]


def _resumen_xlsx(ruta, encabezado, filas, fila_mergeada=2):
    """Arma un resumen mensual mínimo: cada fila es {letra: texto}."""

    def fila_xml(nro, celdas):
        cel = "".join(
            f'<c r="{col}{nro}" t="inlineStr"><is><t>{texto}</t></is></c>'
            for col, texto in celdas.items()
        )
        return f'<row r="{nro}">{cel}</row>'

    cuerpo = [fila_xml(1, encabezado)]
    cuerpo += [fila_xml(n, celdas) for n, celdas in enumerate(filas, start=2)]
    ultima = LETRAS[len(encabezado) - 1]
    merge = (
        f'<mergeCells count="1"><mergeCell ref="C{fila_mergeada}:'
        f'{ultima}{fila_mergeada}"/></mergeCells>'
        if fila_mergeada
        else ""
    )
    hoja = (
        f'<?xml version="1.0"?><worksheet xmlns="{NS_HOJA}">'
        f'<sheetData>{"".join(cuerpo)}</sheetData>{merge}</worksheet>'
    )
    with zipfile.ZipFile(ruta, "w") as archivo:
        archivo.writestr("xl/worksheets/sheet1.xml", hoja)
    return ruta


def _encabezado(dias=5, dia_primero=True, mes=8, anio=2026):
    celdas = {"A": "Dev Ficticio", "B": "Total"}
    for indice in range(dias):
        dia = indice + 1
        texto = f"{dia}/{mes}/{anio}" if dia_primero else f"{mes}/{dia}/{anio}"
        celdas[LETRAS[2 + indice]] = texto
    return celdas


def _resumen_simple(ruta, total="3.00", horas=("1.00", "2.00"), **kwargs):
    """Un cliente, un proyecto, una actividad con horas en C y D."""
    dias = {"C": horas[0], "D": horas[1]}
    return _resumen_xlsx(
        ruta,
        _encabezado(**kwargs),
        [
            {"A": "[000000000] Cliente Ficticio", "B": total},
            {"A": "[XX0000000] Proyecto | largo", "B": total, **dias},
            {"A": "Desarrollo", "B": total, **dias},
            {"A": "Total", "B": total, **dias},
        ],
    )


# --- Lautaro: día/mes, decimales con coma -----------------------------------


def test_lautaro_suma_9_horas():
    assert round(sum(r.horas for r in leer(LAUTARO)), 2) == 9.0


def test_lautaro_interpreta_las_fechas_como_dia_mes():
    registros = leer(LAUTARO)
    horas_por_fecha = {r.fecha: r.horas for r in registros}
    assert horas_por_fecha == {
        date(2026, 8, 3): 0.5,
        date(2026, 8, 10): 1.5,
        date(2026, 8, 26): 4.0,
        date(2026, 8, 31): 3.0,
    }


def test_lautaro_lee_los_decimales_con_coma():
    """'0,50' son media hora, no 50."""
    assert min(r.horas for r in leer(LAUTARO)) == 0.5


def test_lautaro_trae_el_nombre_para_mostrar_como_username():
    """El resumen mensual no trae username: resolverlo es cosa del mapeo."""
    assert {r.username for r in leer(LAUTARO)} == {"Lautaro Zalazar"}


def test_lautaro_trae_proyecto_y_actividad():
    registros = leer(LAUTARO)
    assert {r.cod_proyecto for r in registros} == {"GI2680001"}
    assert {r.actividad for r in registros} == {"Coordinación interna"}


# --- Alexis: mes/día, decimales con punto -----------------------------------


def test_alexis_suma_150_horas():
    assert round(sum(r.horas for r in leer(ALEXIS)), 2) == 150.0


def test_alexis_interpreta_las_fechas_como_mes_dia():
    fechas = [r.fecha for r in leer(ALEXIS)]
    assert all(f.year == 2026 and f.month == 8 for f in fechas)
    assert min(fechas) == date(2026, 8, 2)
    assert max(fechas) == date(2026, 8, 28)


def test_alexis_lee_los_decimales_con_punto():
    horas = {r.fecha: r.horas for r in leer(ALEXIS)}
    assert horas[date(2026, 8, 6)] == 13.5


def test_alexis_trae_el_nombre_el_proyecto_y_la_actividad():
    registros = leer(ALEXIS)
    assert {r.username for r in registros} == {"Alexis Carnero"}
    assert {r.cod_proyecto for r in registros} == {"PR2510126"}
    assert {r.actividad for r in registros} == {"Desarrollo"}


# --- Las filas de proyecto son subtotales: no se cuentan --------------------


@pytest.mark.parametrize("ruta, esperado", [(LAUTARO, 9.0), (ALEXIS, 150.0)])
def test_no_duplica_las_horas_de_las_filas_de_proyecto(ruta, esperado):
    """Si se contaran proyecto Y actividad, el total saldría el doble."""
    registros = leer(ruta)
    assert round(sum(r.horas for r in registros), 2) == esperado
    assert round(sum(r.horas for r in registros), 2) != round(2 * esperado, 2)


def test_una_celda_por_registro_y_ninguna_fecha_repetida():
    registros = leer(LAUTARO)
    assert len(registros) == 4
    assert len({r.fecha for r in registros}) == 4


# --- Red de seguridad: el total declarado por el archivo --------------------


def test_lee_bien_un_resumen_fabricado(tmp_path):
    registros = leer(_resumen_simple(tmp_path / "ok.xlsx"))
    assert [(r.fecha, r.horas) for r in registros] == [
        (date(2026, 8, 1), 1.0),
        (date(2026, 8, 2), 2.0),
    ]


def test_falla_si_el_total_declarado_no_cierra_con_lo_leido(tmp_path):
    """Total adulterado: dice 30 h y las celdas suman 3 h."""
    ruta = _resumen_xlsx(
        tmp_path / "adulterado.xlsx",
        _encabezado(),
        [
            {"A": "[000000000] Cliente Ficticio", "B": "30.00"},
            {
                "A": "[XX0000000] Proyecto | largo",
                "B": "30.00",
                "C": "1.00",
                "D": "2.00",
            },
            {"A": "Desarrollo", "B": "30.00", "C": "1.00", "D": "2.00"},
            {"A": "Total", "B": "30.00", "C": "1.00", "D": "2.00"},
        ],
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "no cierran con el total" in mensaje
    assert "30.00" in mensaje
    assert "3.00" in mensaje


def test_falla_si_la_ultima_fila_no_es_la_de_total(tmp_path):
    ruta = _resumen_xlsx(
        tmp_path / "sin-total.xlsx",
        _encabezado(),
        [
            {"A": "[000000000] Cliente Ficticio", "B": "3.00"},
            {
                "A": "[XX0000000] Proyecto | largo",
                "B": "3.00",
                "C": "1.00",
                "D": "2.00",
            },
            {"A": "Desarrollo", "B": "3.00", "C": "1.00", "D": "2.00"},
        ],
    )
    with pytest.raises(ErrorLectura, match="no es la de"):
        leer(ruta)


# --- Ambigüedades: se detectan, no se adivinan ------------------------------


def test_falla_si_el_encabezado_de_fechas_es_ambiguo(tmp_path):
    """Con una sola columna de día no hay forma de saber cuál es el mes."""
    encabezado = {"A": "Dev Ficticio", "B": "Total", "C": "8/8/2026"}
    ruta = _resumen_xlsx(
        tmp_path / "ambiguo.xlsx",
        encabezado,
        [
            {"A": "[000000000] Cliente Ficticio", "B": "1.00"},
            {"A": "[XX0000000] Proyecto | largo", "B": "1.00", "C": "1.00"},
            {"A": "Desarrollo", "B": "1.00", "C": "1.00"},
            {"A": "Total", "B": "1.00", "C": "1.00"},
        ],
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    assert "día/mes o como mes/día" in str(excepcion.value)


def test_el_orden_de_la_fecha_se_detecta_por_archivo(tmp_path):
    """El mismo contenido con los dos órdenes da las mismas fechas."""
    dia_mes = leer(_resumen_simple(tmp_path / "dm.xlsx", dia_primero=True))
    mes_dia = leer(_resumen_simple(tmp_path / "md.xlsx", dia_primero=False))
    assert [r.fecha for r in dia_mes] == [r.fecha for r in mes_dia]


def test_el_separador_decimal_se_detecta_por_archivo(tmp_path):
    con_coma = leer(
        _resumen_simple(tmp_path / "coma.xlsx", total="3,00", horas=("1,00", "2,00"))
    )
    con_punto = leer(_resumen_simple(tmp_path / "punto.xlsx"))
    assert [r.horas for r in con_coma] == [r.horas for r in con_punto] == [1.0, 2.0]
