import calendar
from datetime import date

from cunix_horas.agregador import Fila, Reporte
from cunix_horas.lector_kimai import Registro
from cunix_horas.validador import validar


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


def test_no_hay_falsos_positivos_de_descuadre():
    assert not any("Descuadre" in a for a in validar(reporte(dias_habiles_completos())))


def test_avisa_por_descuadre_de_horas():
    # Horas en un día fuera del rango del mes (día 32 no existe en octubre)
    # Esto causa que el total de la fila (8.0) no coincida con la suma de días válidos (0)
    horas = {32: 8.0}
    avisos = validar(reporte(horas))
    assert any("Descuadre de horas" in a and "8.00" in a and "0.00" in a for a in avisos)
    assert any("No envíes este Excel" in a for a in avisos)
