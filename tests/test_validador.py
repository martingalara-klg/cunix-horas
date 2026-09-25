import calendar
from datetime import date

from cunix_horas.agregador import Fila, Reporte
from cunix_horas.lector_kimai import Registro
from cunix_horas.validador import dato_de_desvio, validar


def dias_habiles_completos():
    """Octubre 2025 con 8 h en cada día hábil y nada los fines de semana."""
    return {
        dia: 8.0 for dia in range(1, 32) if calendar.weekday(2025, 10, dia) < 5
    }


def reporte(horas_por_dia, descartados=()):
    return Reporte(
        nombre_dev="Matias Zalazar",
        nombre_archivo="Zalazar",
        anio=2025,
        mes=10,
        filas=(Fila("Cliente", "Proyecto", "Desarrollo", dict(horas_por_dia)),),
        descartados=tuple(descartados),
    )


def test_sin_problemas_no_hay_avisos():
    assert validar(reporte(dias_habiles_completos())) == []


def test_avisa_si_un_dia_supera_las_12_horas():
    horas = dias_habiles_completos()
    horas[1] = 14.0
    avisos = validar(reporte(horas))
    assert any("14.0" in a and "1/10/2025" in a for a in avisos)


def test_avisa_por_horas_en_fin_de_semana():
    horas = dias_habiles_completos()
    horas[4] = 3.0  # 2025-10-04 es sábado
    avisos = validar(reporte(horas))
    assert any("fin de semana" in a and "4/10/2025" in a for a in avisos)


def test_avisa_por_dias_habiles_sin_carga():
    horas = dias_habiles_completos()
    del horas[1]
    del horas[2]
    avisos = validar(reporte(horas))
    assert any("sin carga" in a and "1/10/2025" in a for a in avisos)
    assert any("sin carga" in a and "2/10/2025" in a for a in avisos)


def test_avisa_por_registros_descartados():
    descartado = Registro(date(2025, 9, 30), 4.0, "mzalazar", "CO2610170", "Desarrollo")
    avisos = validar(reporte(dias_habiles_completos(), [descartado]))
    assert any("fuera del mes" in a and "30/9/2025" in a for a in avisos)


def reporte_de_varias_filas(horas_por_dia, cantidad_filas):
    """Mismo mes, con N filas de actividad idénticas: muchas celdas de día."""
    return Reporte(
        nombre_dev="Matias Zalazar",
        nombre_archivo="Zalazar",
        anio=2025,
        mes=10,
        filas=tuple(
            Fila("Cliente", "Proyecto", f"Actividad {n}", dict(horas_por_dia))
            for n in range(cantidad_filas)
        ),
        descartados=(),
    )


def test_no_hay_falsos_positivos_de_desvio_por_redondeo():
    """Horas exactas: el redondeo no aparta nada, no hay nada que avisar."""
    assert not any(
        "Desvío por redondeo" in a for a in validar(reporte(dias_habiles_completos()))
    )


# Una duración de Kimai se carga en minutos, así que las horas que puede
# producir son múltiplos de 1/60. 50 minutos son 0.8333... h, que redondeadas a
# la celda de día valen 0.83: se pierden 0.0033 h por celda. Un dev-mes de
# muchas celdas así acumula décimas de hora, y eso es lo que el aviso vigila.
MINUTOS_POR_CELDA = 50
HORAS_POR_CELDA = MINUTOS_POR_CELDA / 60


def test_avisa_cuando_el_redondeo_acumula_mas_que_el_umbral():
    """El aviso vigila el total del Excel contra las horas crudas del export.

    Regresión: el aviso anterior comparaba valores redondeados contra valores
    redondeados, que es la misma cuenta dos veces y no puede dispararse nunca.
    Lo que sí puede pasar es esto: cada celda de día pierde hasta 0.005 h al
    redondearse y el error se acumula con la cantidad de celdas.

    Segunda regresión: este test usaba 8.00499 h por celda, un valor que
    ninguna duración de Kimai puede producir. Un control que sólo se dispara
    con un input imposible no prueba nada; ahora son 124 celdas de 50 minutos.
    """
    # 0.0033 h de pérdida por celda x 31 días x 4 filas = 0.41 h.
    horas = {dia: HORAS_POR_CELDA for dia in range(1, 32)}
    avisos = validar(reporte_de_varias_filas(horas, 4))

    desvios = [a for a in avisos if "Desvío por redondeo" in a]
    assert len(desvios) == 1
    assert "0.41 h" in desvios[0]
    assert "redondea" in desvios[0]
    assert "no es un error de carga" in desvios[0].lower()


def test_el_aviso_de_desvio_no_llama_a_las_horas_del_mes_el_total_de_kimai():
    """Los registros descartados no entran en ese número; el texto lo aclara."""
    horas = {dia: HORAS_POR_CELDA for dia in range(1, 32)}
    descartado = Registro(
        date(2025, 9, 30), 4.0, "mzalazar", "CO2610170", "Desarrollo"
    )
    reporte_con_descartes = Reporte(
        nombre_dev="Matias Zalazar",
        nombre_archivo="Zalazar",
        anio=2025,
        mes=10,
        filas=tuple(
            Fila("Cliente", "Proyecto", f"Actividad {n}", dict(horas))
            for n in range(4)
        ),
        descartados=(descartado,),
    )
    desvio = next(
        a for a in validar(reporte_con_descartes) if "Desvío por redondeo" in a
    )
    assert "export de Kimai trae" not in desvio
    assert "caen en 10/2025" in desvio
    assert "fuera del mes" in desvio


def test_el_desvio_por_redondeo_no_se_avisa_por_debajo_del_umbral():
    """Un desvío chico es ruido de presentación, no se pone delante del dueño."""
    # 0.0033 h x 31 días x 1 fila = 0.10 h, por debajo de las 0.25 h del umbral.
    horas = {dia: HORAS_POR_CELDA for dia in range(1, 32)}
    assert not any(
        "Desvío por redondeo" in a for a in validar(reporte_de_varias_filas(horas, 1))
    )


def test_el_desvio_se_informa_siempre_aunque_no_supere_el_umbral():
    """El dueño factura con estos números: los ve todos los meses."""
    horas = {dia: HORAS_POR_CELDA for dia in range(1, 32)}
    dato = dato_de_desvio(reporte_de_varias_filas(horas, 1))
    assert "0.10 h" in dato
    assert "informativo" in dato


def test_el_dato_de_desvio_tambien_esta_cuando_el_desvio_es_cero():
    dato = dato_de_desvio(reporte(dias_habiles_completos()))
    assert "0.00 h" in dato
