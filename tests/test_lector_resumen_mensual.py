"""Tests del lector de resumen mensual (.xlsx con B1='Total').

Este formato **vuelve a aceptarse** como entrada. No trae el usuario de
Kimai, ni el mail, ni la descripción, ni la hora de inicio, ni el número de
proyecto: de esos cinco, la descripción y la hora se emiten vacías (el
partner las acepta así), y los otros tres los completa `completado.py` desde
`config/mapeo.yaml`.

Lo de acá es sólo el parseo de la grilla, así que casi todos los tests llaman
al lector interno `leer_resumen_mensual`. El bloque del final comprueba que
el punto de entrada lo reconoce y lo marca con su origen.
"""
import zipfile
from datetime import date
from pathlib import Path

import pytest
from conftest import FIXTURES

from cunix_horas.kimai_comun import (
    ORIGEN_RESUMEN_MENSUAL,
    ErrorLectura,
    leer_hoja,
)
from cunix_horas.lector_kimai import leer as leer_por_el_punto_de_entrada
from cunix_horas.lector_resumen_mensual import leer_resumen_mensual


def leer(ruta):
    """El lector interno, sin pasar por el punto de entrada."""
    ruta = Path(ruta)
    return leer_resumen_mensual(ruta, leer_hoja(ruta))

LAUTARO = FIXTURES / "kimai-resumen-mensual-lautaro.xlsx"
ALEXIS = FIXTURES / "kimai-resumen-mensual-alexis.xlsx"
MULTI = FIXTURES / "kimai-resumen-mensual-multi.xlsx"

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


def test_falla_si_el_archivo_mezcla_los_separadores_decimales(tmp_path):
    """'2.00' leído con coma da 200: no se elige uno, se falla."""
    ruta = _resumen_simple(
        tmp_path / "mezclado.xlsx", total="3,00", horas=("1,00", "2.00")
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "separadores decimales mezclados" in mensaje
    assert "1,00" in mensaje
    assert "2.00" in mensaje


def test_falla_si_un_encabezado_de_dia_esta_vacio(tmp_path):
    """Sin encabezado, esa columna y todas sus horas desaparecían en silencio."""
    encabezado = _encabezado()
    encabezado["D"] = ""
    ruta = _resumen_xlsx(
        tmp_path / "sin-encabezado.xlsx",
        encabezado,
        [
            {"A": "[000000000] Cliente Ficticio", "B": "3.00"},
            {
                "A": "[XX0000000] Proyecto | largo",
                "B": "3.00",
                "C": "1.00",
                "D": "2.00",
            },
            {"A": "Desarrollo", "B": "3.00", "C": "1.00", "D": "2.00"},
            {"A": "Total", "B": "3.00", "C": "1.00", "D": "2.00"},
        ],
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "sin encabezado de día" in mensaje
    assert ": D." in mensaje


def test_una_fila_de_proyecto_sin_horas_se_nombra_a_si_misma(tmp_path):
    """El mensaje tiene que apuntar a la fila de proyecto, no a la actividad."""
    ruta = _resumen_xlsx(
        tmp_path / "proyecto-sin-horas.xlsx",
        _encabezado(),
        [
            {"A": "[000000000] Cliente Ficticio", "B": "3.00"},
            {"A": "[XX0000000] Proyecto | largo", "B": "3.00"},
            {"A": "Desarrollo", "B": "3.00", "C": "1.00", "D": "2.00"},
            {"A": "Total", "B": "3.00", "C": "1.00", "D": "2.00"},
        ],
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "fila 3" in mensaje
    assert "[XX0000000] Proyecto | largo" in mensaje
    assert "ninguna hora por día" in mensaje


def test_la_primera_columna_de_dia_no_esta_duplicada():
    """Si hubiera dos copias, la detección de filas mergeadas podría divergir."""
    import inspect

    from cunix_horas import kimai_comun, lector_resumen_mensual

    assert lector_resumen_mensual.PRIMERA_COL_DE_DIA is kimai_comun.PRIMERA_COL_DE_DIA
    parametro = inspect.signature(kimai_comun.leer_hoja).parameters[
        "primera_columna_de_dia"
    ]
    assert parametro.default is kimai_comun.PRIMERA_COL_DE_DIA


# --- Un cliente con dos proyectos y un proyecto con dos actividades --------
# `kimai-resumen-mensual-multi.xlsx` es el único fixture con más de un
# proyecto por cliente y más de una actividad por proyecto, que es justo lo
# que hace no trivial la clasificación de filas. Los reales traen uno de cada.


def test_multi_suma_las_21_horas_del_archivo():
    assert round(sum(r.horas for r in leer(MULTI)), 2) == 21.0


def test_multi_separa_los_tres_proyectos_de_los_dos_clientes():
    horas: dict[str, float] = {}
    for registro in leer(MULTI):
        horas[registro.cod_proyecto] = round(
            horas.get(registro.cod_proyecto, 0.0) + registro.horas, 2
        )
    assert horas == {"PR2510126": 11.0, "CO2610170": 8.25, "GI2680001": 1.75}


def test_multi_no_mezcla_las_dos_actividades_del_mismo_proyecto():
    horas: dict[tuple[str, str], float] = {}
    for registro in leer(MULTI):
        clave = (registro.cod_proyecto, registro.actividad)
        horas[clave] = round(horas.get(clave, 0.0) + registro.horas, 2)
    assert horas[("PR2510126", "Desarrollo")] == 6.5
    assert horas[("PR2510126", "Gestión")] == 4.5


def test_multi_imputa_cada_hora_a_su_dia():
    horas: dict[date, float] = {}
    for registro in leer(MULTI):
        horas[registro.fecha] = round(horas.get(registro.fecha, 0.0) + registro.horas, 2)
    assert horas == {
        date(2026, 8, 3): 5.5,
        date(2026, 8, 5): 8.5,
        date(2026, 8, 12): 3.0,
        date(2026, 8, 20): 2.25,
        date(2026, 8, 31): 1.75,
    }


def test_multi_emite_un_registro_por_celda_de_actividad():
    """Ni las filas de cliente ni las de proyecto emiten registros."""
    assert len(leer(MULTI)) == 7


# --- Un mes completo fabricado, para meterle los defectos de a uno ---------

DIAS_DE_AGOSTO = 31


def _columna_del_dia(dia: int, corrimiento: int = 0) -> str:
    return LETRAS[2 + dia - 1 + corrimiento]


def _resumen_del_mes(ruta, actividades, quitar=(), corrimiento=0):
    """Un resumen del mes entero, con los totales calculados de los datos.

    `actividades` es [(nombre, {día: horas}), ...] bajo un mismo proyecto. Los
    totales por día y el general se calculan ANTES de aplicar `quitar` (celdas
    que se borran de las filas de actividad) y `corrimiento` (columnas que se
    desplazan), así el archivo queda igual al que emite Kimai salvo por el
    defecto que el test quiere probar.
    """
    por_dia: dict[int, float] = {}
    for _, horas in actividades:
        for dia, valor in horas.items():
            por_dia[dia] = round(por_dia.get(dia, 0.0) + valor, 2)
    general = round(sum(por_dia.values()), 2)

    def celdas(horas, nombre=None):
        return {
            _columna_del_dia(dia, corrimiento): f"{valor:.2f}"
            for dia, valor in sorted(horas.items())
            if (nombre, dia) not in quitar
        }

    encabezado = {"A": "Dev Ficticio", "B": "Total"}
    for dia in range(1, DIAS_DE_AGOSTO + 1):
        encabezado[_columna_del_dia(dia)] = f"{dia}/8/2026"

    filas = [
        {"A": "[000000000] Cliente Ficticio", "B": f"{general:.2f}"},
        {
            "A": "[XX0000000] Proyecto | largo",
            "B": f"{general:.2f}",
            **celdas(por_dia),
        },
    ]
    filas += [
        {"A": nombre, "B": f"{sum(horas.values()):.2f}", **celdas(horas, nombre)}
        for nombre, horas in actividades
    ]
    filas.append(
        {
            "A": "Total",
            "B": f"{general:.2f}",
            **{
                _columna_del_dia(dia): f"{por_dia.get(dia, 0.0):.2f}"
                for dia in range(1, DIAS_DE_AGOSTO + 1)
            },
        }
    )
    return _resumen_xlsx(ruta, encabezado, filas)


def _mes_de_93_celdas():
    """93 celdas que declaran 184.25 h: 92 de 2 h y una de 0.25 h."""
    actividades = []
    for indice, nombre in enumerate(("Desarrollo", "Gestión", "Soporte")):
        horas = {dia: 2.0 for dia in range(1, DIAS_DE_AGOSTO + 1)}
        if indice == 0:
            horas[1] = 0.25
        actividades.append((nombre, horas))
    return actividades


def test_un_mes_completo_fabricado_se_lee_entero(tmp_path):
    registros = leer(_resumen_del_mes(tmp_path / "mes.xlsx", _mes_de_93_celdas()))
    assert len(registros) == 93
    assert round(sum(r.horas for r in registros), 2) == 184.25


def test_falla_si_se_pierde_una_celda_chica_de_un_mes_grande(tmp_path):
    """Con 93 celdas, borrar una de 0.25 h tiene que disparar el control.

    La tolerancia vieja (0.01 + 0.005 por celda) llegaba a 0.47 h, así que
    esta celda pasaba sin que nadie se enterara.
    """
    ruta = _resumen_del_mes(
        tmp_path / "celda-borrada.xlsx",
        _mes_de_93_celdas(),
        quitar={("Desarrollo", 1)},
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "no cierran con el total" in mensaje
    assert "184.25" in mensaje
    assert "184.00" in mensaje


def test_el_redondeo_real_de_un_mes_grande_no_da_falso_positivo(tmp_path):
    """93 celdas redondeadas a 2 decimales no pueden hacer saltar el control."""
    actividades = []
    for indice, nombre in enumerate(("Desarrollo", "Gestión", "Soporte")):
        horas = {
            dia: round(1.0 + ((dia * 7 + indice) % 13) / 3, 2)
            for dia in range(1, DIAS_DE_AGOSTO + 1)
        }
        actividades.append((nombre, horas))
    ruta = _resumen_del_mes(tmp_path / "redondeado.xlsx", actividades)
    assert len(leer(ruta)) == 93


def test_falla_si_las_horas_estan_corridas_una_columna(tmp_path):
    """El total general no cambia, pero cada hora quedó en el día siguiente."""
    ruta = _resumen_del_mes(
        tmp_path / "corrido.xlsx",
        [("Desarrollo", {3: 4.0, 5: 2.0})],
        corrimiento=1,
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "03/08/2026" in mensaje
    assert "no cierran" in mensaje
    assert "fechas equivocadas" in mensaje


def test_un_mes_sin_corrimiento_pasa_el_control_por_dia(tmp_path):
    ruta = _resumen_del_mes(
        tmp_path / "derecho.xlsx", [("Desarrollo", {3: 4.0, 5: 2.0})]
    )
    assert [(r.fecha, r.horas) for r in leer(ruta)] == [
        (date(2026, 8, 3), 4.0),
        (date(2026, 8, 5), 2.0),
    ]


# --- El punto de entrada acepta este formato y lo marca ---------------------
# El resumen mensual volvió a aceptarse: el dueño puede usar el export que ya
# tiene. Lo que no trae no se inventa, se completa desde el mapeo, y eso lo
# hace `completado.py` a partir del `origen` que el lector deja marcado acá.


@pytest.mark.parametrize("ruta", [LAUTARO, ALEXIS, MULTI])
def test_el_punto_de_entrada_acepta_el_resumen_mensual(ruta):
    assert leer_por_el_punto_de_entrada(ruta)


def test_el_punto_de_entrada_devuelve_lo_mismo_que_el_lector_interno():
    assert leer_por_el_punto_de_entrada(LAUTARO) == leer(LAUTARO)


@pytest.mark.parametrize("ruta", [LAUTARO, ALEXIS, MULTI])
def test_los_registros_quedan_marcados_como_resumen_mensual(ruta):
    """Sin esta marca, `completado.py` no sabría a cuáles completar."""
    assert {r.origen for r in leer_por_el_punto_de_entrada(ruta)} == {
        ORIGEN_RESUMEN_MENSUAL
    }


def test_el_resumen_mensual_no_trae_las_columnas_que_completa_el_mapeo():
    """Salen vacías del lector: llenarlas es cosa del mapeo, no del parseo."""
    for registro in leer_por_el_punto_de_entrada(LAUTARO):
        assert registro.email == ""
        assert registro.numero_proyecto == ""
        assert registro.hora_inicio is None
        assert registro.descripcion == ""


def test_el_lector_interno_sigue_leyendo_los_totales_de_los_dos_reales():
    """Lo que el parseo ya sabía hacer se conserva entero."""
    assert round(sum(r.horas for r in leer(LAUTARO)), 2) == 9.0
    assert round(sum(r.horas for r in leer(ALEXIS)), 2) == 150.0
