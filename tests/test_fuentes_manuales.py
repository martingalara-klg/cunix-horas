"""Las planillas de input/<mes>/manual/.

Dos propiedades son el motivo de este módulo y están cubiertas una por una:

- la lectura es **tolerante** (no se asume ni la hoja, ni la fila del
  encabezado, ni el formato de la fecha) y, cuando no alcanza, el error dice
  qué se esperaba encontrar;
- una resta que no cierra **no genera nada**.
"""
from datetime import date, time

import openpyxl
import pytest

from cunix_horas.anexos import (
    DescripcionesAparte,
    FuentesManuales,
    HorasSinKimai,
    Periodo,
)
from cunix_horas.filas_anexo import FilaAnexo
from cunix_horas.fuentes_manuales import (
    ErrorFuenteManual,
    aplicar,
    leer_planilla_de_horas,
)

PERIODO = Periodo(2026, 9)
GABRIEL = HorasSinKimai(
    persona="Gabriel Denis",
    planilla="gabriel.xlsx",
    restar_a="Alexis Carnero",
    proyecto="SELICO",
)


def planilla(tmp_path, filas, encabezados=("Fecha", "Horas", "Descripción"), titulo=True):
    """Escribe una planilla de horas con los encabezados que se le pidan."""
    carpeta = tmp_path / "manual"
    carpeta.mkdir(parents=True, exist_ok=True)
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Septiembre 2026"
    if titulo:
        hoja.append(["REPORTE DE HORAS"])
        hoja.append([])
    hoja.append(list(encabezados))
    for fila_de_datos in filas:
        hoja.append(list(fila_de_datos))
    ruta = carpeta / "gabriel.xlsx"
    libro.save(ruta)
    return ruta


def fila(dia, horas, persona="Alexis Carnero", inicio=None, fin=None, descripcion=""):
    return FilaAnexo(
        fecha=date(2026, 9, dia),
        inicio=inicio,
        fin=fin,
        persona=persona,
        proyecto="SELICO",
        descripcion=descripcion,
        horas=horas,
    )


def aplicar_horas(tmp_path, filas, fuente=GABRIEL):
    return aplicar(
        tuple(filas),
        (("Alexis Carnero", "acarnero"),),
        FuentesManuales(horas_sin_kimai=(fuente,)),
        tmp_path,
        PERIODO,
    )


# --- Lectura tolerante ------------------------------------------------------


def test_lee_la_fecha_como_texto_en_espanol_sin_anio(tmp_path):
    ruta = planilla(tmp_path, [["Mié 02 Sep", 8, "algo"], ["Jue 03 Sep", 4, "otra"]])
    renglones = leer_planilla_de_horas(ruta, PERIODO)
    assert [(r.dia, r.horas) for r in renglones] == [(2, 8.0), (3, 4.0)]


def test_lee_la_fecha_con_barras_y_como_fecha_de_excel(tmp_path):
    ruta = planilla(tmp_path, [["02/09/2026", 8, "a"], [date(2026, 9, 3), 4, "b"]])
    assert [r.dia for r in leer_planilla_de_horas(ruta, PERIODO)] == [2, 3]


def test_saltea_las_filas_de_subtotal_y_total(tmp_path):
    ruta = planilla(
        tmp_path,
        [
            ["Mié 02 Sep", 8, "algo"],
            [None, None, None],
            [None, 8, "Subtotal semana 01-05 Sep"],
            ["TOTAL SEPTIEMBRE 2026", None, None],
        ],
    )
    assert [r.dia for r in leer_planilla_de_horas(ruta, PERIODO)] == [2]


def test_no_depende_del_orden_ni_de_los_acentos_de_los_encabezados(tmp_path):
    ruta = planilla(
        tmp_path,
        [["algo que se hizo", 8, "02/09/2026"]],
        encabezados=("DETALLE", "HS", "FECHA"),
    )
    (renglon,) = leer_planilla_de_horas(ruta, PERIODO)
    assert (renglon.dia, renglon.horas) == (2, 8.0)
    assert renglon.descripcion == "algo que se hizo"


def test_una_planilla_que_no_se_entiende_dice_que_esperaba_encontrar(tmp_path):
    ruta = planilla(tmp_path, [["x", "y"]], encabezados=("Columna A", "Columna B"))
    with pytest.raises(ErrorFuenteManual) as excepcion:
        leer_planilla_de_horas(ruta, PERIODO)
    mensaje = str(excepcion.value)
    assert "«fecha»" in mensaje and "«horas»" in mensaje
    assert "columna a" in mensaje.lower()


def test_una_planilla_de_otro_mes_frena_diciendolo(tmp_path):
    ruta = planilla(tmp_path, [["02/08/2026", 8, "algo"]])
    with pytest.raises(ErrorFuenteManual, match="es de otro mes"):
        leer_planilla_de_horas(ruta, PERIODO)


def test_una_planilla_sin_renglones_de_datos_lo_dice(tmp_path):
    ruta = planilla(tmp_path, [])
    with pytest.raises(ErrorFuenteManual, match="ningún renglón de datos"):
        leer_planilla_de_horas(ruta, PERIODO)


# --- La resta que cierra ----------------------------------------------------


def test_la_resta_separa_las_horas_dia_por_dia(tmp_path):
    planilla(tmp_path, [["Mié 02 Sep", 3, "lo de Gabriel"]])
    filas, usuarios = aplicar_horas(tmp_path, [fila(2, 8.0), fila(3, 5.0)])

    por_persona = {}
    for f in filas:
        por_persona[f.persona] = por_persona.get(f.persona, 0.0) + f.horas
    assert por_persona == {"Alexis Carnero": 10.0, "Gabriel Denis": 3.0}
    assert sum(f.horas for f in filas) == 13.0
    assert dict(usuarios)["Gabriel Denis"] == ""


def test_un_dia_entero_deja_a_la_otra_persona_sin_fila(tmp_path):
    planilla(tmp_path, [["Mié 02 Sep", 8, "todo el día"]])
    filas, _ = aplicar_horas(tmp_path, [fila(2, 8.0)])
    assert [(f.persona, f.horas) for f in filas] == [("Gabriel Denis", 8.0)]


def test_el_corte_dentro_del_dia_es_por_horario(tmp_path):
    """La primera parte del bloque queda para quien lo tenía; la última pasa."""
    planilla(tmp_path, [["Mié 02 Sep", 3, "lo de Gabriel"]])
    filas, _ = aplicar_horas(
        tmp_path, [fila(2, 8.0, inicio=time(8, 0), fin=time(16, 0))]
    )
    por_persona = {f.persona: f for f in filas}
    assert por_persona["Alexis Carnero"].inicio == time(8, 0)
    assert por_persona["Alexis Carnero"].fin == time(13, 0)
    assert por_persona["Gabriel Denis"].inicio == time(13, 0)
    assert por_persona["Gabriel Denis"].fin == time(16, 0)


def test_las_filas_movidas_llevan_la_nota_que_lo_explica(tmp_path):
    planilla(tmp_path, [["Mié 02 Sep", 3, "lo de Gabriel"]])
    filas, _ = aplicar_horas(tmp_path, [fila(2, 8.0)])
    assert all("no tiene usuario de Kimai" in f.nota for f in filas)


def test_la_descripcion_de_la_planilla_viaja_a_la_fila_nueva(tmp_path):
    planilla(tmp_path, [["Mié 02 Sep", 3, "R-01 lo que hizo"]])
    filas, _ = aplicar_horas(tmp_path, [fila(2, 8.0)])
    nueva = next(f for f in filas if f.persona == "Gabriel Denis")
    assert nueva.descripcion == "R-01 lo que hizo"
    assert nueva.proyecto == "SELICO"


# --- La resta que NO cierra: no se genera nada ------------------------------


def test_un_dia_que_el_otro_no_tiene_frena_sin_generar(tmp_path):
    planilla(tmp_path, [["Mié 09 Sep", 3, "algo"]])
    with pytest.raises(ErrorFuenteManual) as excepcion:
        aplicar_horas(tmp_path, [fila(2, 8.0)])
    mensaje = str(excepcion.value)
    assert "NO SE GENERÓ NADA" in mensaje
    assert "no tiene ninguna hora cargada ese día" in mensaje
    assert "9/09" in mensaje


def test_un_dia_que_quedaria_en_negativo_frena_sin_generar(tmp_path):
    planilla(tmp_path, [["Mié 02 Sep", 12, "algo"]])
    with pytest.raises(ErrorFuenteManual) as excepcion:
        aplicar_horas(tmp_path, [fila(2, 8.0)])
    mensaje = str(excepcion.value)
    assert "NO SE GENERÓ NADA" in mensaje
    assert "negativo" in mensaje


def test_una_planilla_que_falta_frena_diciendo_donde_se_la_buscaba(tmp_path):
    (tmp_path / "manual").mkdir(parents=True)
    with pytest.raises(ErrorFuenteManual) as excepcion:
        aplicar_horas(tmp_path, [fila(2, 8.0)])
    mensaje = str(excepcion.value)
    assert "Falta la planilla manual gabriel.xlsx" in mensaje
    assert "manual" in mensaje
    assert "NO SE GENERÓ NADA" in mensaje


def test_si_la_persona_ya_carga_en_kimai_la_declaracion_frena(tmp_path):
    """Si no, sus horas se contarían dos veces."""
    planilla(tmp_path, [["Mié 02 Sep", 3, "algo"]])
    with pytest.raises(ErrorFuenteManual, match="dos veces"):
        aplicar_horas(tmp_path, [fila(2, 8.0), fila(2, 3.0, persona="Gabriel Denis")])


# --- Descripciones entregadas aparte ----------------------------------------


def planilla_de_descripciones(tmp_path, filas):
    carpeta = tmp_path / "manual"
    carpeta.mkdir(parents=True, exist_ok=True)
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(
        [
            "Fecha",
            "Inicio",
            "Fin",
            "Persona",
            "Proyecto",
            "Descripción (ticket)",
            "Horas",
        ]
    )
    for fila_de_datos in filas:
        hoja.append(list(fila_de_datos))
    libro.save(carpeta / "alexis.xlsx")


def aplicar_descripciones(tmp_path, filas):
    return aplicar(
        tuple(filas),
        (("Alexis Carnero", "acarnero"),),
        FuentesManuales(
            descripciones=(DescripcionesAparte("Alexis Carnero", "alexis.xlsx"),)
        ),
        tmp_path,
        PERIODO,
    )


def test_la_descripcion_entregada_aparte_se_carga_en_la_fila(tmp_path):
    planilla_de_descripciones(
        tmp_path,
        [["02/09/2026", "08:15", "16:15", "Alexis Carnero", "SELICO", "R-01 algo", 8]],
    )
    filas, _ = aplicar_descripciones(tmp_path, [fila(2, 8.0)])
    assert filas[0].descripcion == "R-01 algo"
    assert filas[0].inicio == time(8, 15)


def test_si_las_horas_no_coinciden_no_se_genera_nada(tmp_path):
    planilla_de_descripciones(
        tmp_path,
        [["02/09/2026", "08:15", "16:15", "Alexis Carnero", "SELICO", "R-01 algo", 7]],
    )
    with pytest.raises(ErrorFuenteManual) as excepcion:
        aplicar_descripciones(tmp_path, [fila(2, 8.0)])
    assert "NO SE GENERÓ NADA" in str(excepcion.value)


def test_un_dia_de_la_planilla_que_la_persona_no_tiene_frena(tmp_path):
    planilla_de_descripciones(
        tmp_path,
        [["09/09/2026", "08:15", "16:15", "Alexis Carnero", "SELICO", "R-01 algo", 8]],
    )
    with pytest.raises(ErrorFuenteManual, match="no tiene horas cargadas ese día"):
        aplicar_descripciones(tmp_path, [fila(2, 8.0)])


# --- Un mes sin fuentes declaradas ------------------------------------------


def test_sin_fuentes_declaradas_las_filas_salen_tal_cual(tmp_path):
    entrada = (fila(2, 8.0),)
    filas, usuarios = aplicar(
        entrada, (("Alexis Carnero", "acarnero"),), FuentesManuales(), tmp_path, PERIODO
    )
    assert filas == entrada
    assert usuarios == (("Alexis Carnero", "acarnero"),)
