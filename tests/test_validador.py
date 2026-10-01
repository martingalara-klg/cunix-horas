"""Los avisos de carga de horas, sobre los registros crudos de cada persona.

Antes había dos juegos de avisos: éste y uno que trabajaba sobre el `Reporte`
pivoteado del Excel por desarrollador. Ese formato se eliminó, y con él el
aviso de desvío por redondeo, que sólo tenía sentido ahí: los anexos escriben
las horas exactas. Los cuatro avisos que quedan son los que siguen importando,
y están cubiertos acá uno por uno.
"""
import calendar
from datetime import date

from cunix_horas.lector_kimai import Registro
from cunix_horas.validador import avisos_de_desarrollador

ANIO, MES = 2025, 10


def registro(dia, horas, mes=MES, anio=ANIO):
    return Registro(date(anio, mes, dia), horas, "mzalazar", "CO2610170", "Desarrollo")


def mes_habil_completo():
    """Octubre 2025 con 8 h en cada día hábil y nada los fines de semana."""
    return [
        registro(dia, 8.0)
        for dia in range(1, 32)
        if calendar.weekday(ANIO, MES, dia) < 5
    ]


def avisos(del_mes, descartados=()):
    return avisos_de_desarrollador(list(del_mes), list(descartados), ANIO, MES)


def test_sin_problemas_no_hay_avisos():
    assert avisos(mes_habil_completo()) == []


def test_avisa_si_un_dia_supera_las_12_horas():
    del_mes = mes_habil_completo() + [registro(1, 6.0)]
    assert any("14.0" in a and "1/10/2025" in a for a in avisos(del_mes))


def test_no_avisa_con_exactamente_12_horas():
    """El límite es «más de 12», no «12»: una jornada larga no es un hallazgo."""
    del_mes = mes_habil_completo() + [registro(1, 4.0)]
    assert not any("Más de 12" in a for a in avisos(del_mes))


def test_avisa_por_horas_en_fin_de_semana():
    del_mes = mes_habil_completo() + [registro(4, 3.0)]  # 2025-10-04, sábado
    assert any("fin de semana" in a and "4/10/2025" in a for a in avisos(del_mes))


def test_avisa_por_dias_habiles_sin_carga():
    del_mes = [r for r in mes_habil_completo() if r.fecha.day not in (1, 2)]
    resultado = avisos(del_mes)
    assert any("sin carga" in a and "1/10/2025" in a for a in resultado)
    assert any("sin carga" in a and "2/10/2025" in a for a in resultado)


def test_no_avisa_por_un_fin_de_semana_sin_carga():
    assert not any("sin carga" in a and "4/10/2025" in a for a in avisos(mes_habil_completo()))


def test_avisa_por_registros_descartados():
    descartado = registro(30, 4.0, mes=9)
    resultado = avisos(mes_habil_completo(), [descartado])
    assert any("fuera del mes" in a and "30/9/2025" in a for a in resultado)
    assert any("EXCLUIDO de los anexos" in a for a in resultado)


def test_las_horas_del_dia_se_suman_entre_registros():
    """El límite diario mira el día entero, no cada carga por separado."""
    del_mes = mes_habil_completo() + [registro(1, 3.0), registro(1, 2.0)]
    assert any("13.0" in a and "1/10/2025" in a for a in avisos(del_mes))


def test_no_muta_lo_que_recibe():
    del_mes = mes_habil_completo()
    copia = list(del_mes)
    avisos(del_mes)
    assert del_mes == copia
