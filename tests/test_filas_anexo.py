"""Las filas de la hoja `Detalle` y las tres tablas que salen de ellas."""
from datetime import date, time

from cunix_horas.filas_anexo import (
    FilaAnexo,
    construir,
    por_persona,
    por_proyecto,
    principales_trabajos,
    sin_descripcion,
    ticket_de,
    total_de,
)
from cunix_horas.kimai_comun import Registro
from cunix_horas.mapeo import DestinoProyecto, Mapeo

MAPEO = Mapeo(
    personas={},
    proyectos={"PR01": DestinoProyecto("MINVU", "SELICO", "PR01")},
)


def registro(dia, horas, descripcion="", inicio=time(9, 0), codigo="PR01", nombre="Ana Perez"):
    return Registro(
        fecha=date(2026, 9, dia),
        horas=horas,
        username="aperez",
        cod_proyecto=codigo,
        actividad="Desarrollo",
        texto_proyecto=f"[{codigo}] Proyecto crudo | descripcion larga",
        hora_inicio=inicio,
        nombre=nombre,
        descripcion=descripcion,
    )


def fila(dia, horas, persona="Ana Perez", proyecto="SELICO", descripcion=""):
    return FilaAnexo(
        fecha=date(2026, 9, dia),
        inicio=None,
        fin=None,
        persona=persona,
        proyecto=proyecto,
        descripcion=descripcion,
        horas=horas,
    )


# --- La hoja Detalle --------------------------------------------------------


def test_cada_registro_da_una_fila_con_el_nombre_del_proyecto_mapeado():
    armado = construir([registro(1, 4.0, "R-012345 algo")], MAPEO)
    (unica,) = armado.filas
    assert unica.persona == "Ana Perez"
    assert unica.proyecto == "SELICO"
    assert unica.horas == 4.0
    assert unica.descripcion == "R-012345 algo"


def test_el_fin_se_calcula_sumandole_las_horas_al_inicio():
    (unica,) = construir([registro(1, 4.5, inicio=time(8, 15))], MAPEO).filas
    assert unica.inicio == time(8, 15)
    assert unica.fin == time(12, 45)


def test_un_bloque_que_cruza_la_medianoche_da_la_vuelta():
    """Es real y la plantilla lo marca: no se corrige, se deja ver."""
    (unica,) = construir([registro(7, 3.0, inicio=time(23, 0))], MAPEO).filas
    assert unica.fin == time(2, 0)


def test_sin_hora_de_inicio_las_dos_columnas_quedan_vacias():
    """El resumen mensual de Kimai no trae horario y la plantilla lo acepta."""
    (unica,) = construir([registro(1, 4.0, inicio=None)], MAPEO).filas
    assert unica.inicio is None and unica.fin is None


def test_un_proyecto_sin_mapear_sale_con_el_nombre_derivado_y_se_lista():
    armado = construir([registro(1, 4.0, codigo="ZZ99")], MAPEO)
    assert armado.filas[0].proyecto == "Proyecto crudo"
    assert [p.codigo for p in armado.sin_mapear] == ["ZZ99"]


def test_las_filas_salen_en_orden_cronologico_y_por_persona():
    registros = [
        registro(5, 1.0, nombre="Beto Suarez"),
        registro(1, 1.0, nombre="Ana Perez"),
        registro(5, 1.0, nombre="Ana Perez"),
    ]
    filas = construir(registros, MAPEO).filas
    assert [(f.fecha.day, f.persona) for f in filas] == [
        (1, "Ana Perez"),
        (5, "Ana Perez"),
        (5, "Beto Suarez"),
    ]


def test_el_usuario_de_kimai_de_cada_persona_viaja_para_la_hoja_datos():
    assert construir([registro(1, 4.0)], MAPEO).usuarios == (("Ana Perez", "aperez"),)


# --- Tabla 1: horas e importe por proyecto ----------------------------------


def test_la_tabla_de_proyectos_suma_las_horas_y_calcula_el_importe():
    filas = [fila(1, 4.0), fila(2, 6.0, proyecto="Subastas")]
    totales = por_proyecto(filas, {"SELICO": ("MINVU", 17.0), "Subastas": ("SNA", 20.0)})
    assert [(t.proyecto, t.horas, t.importe) for t in totales] == [
        ("Subastas", 6.0, 120.0),
        ("SELICO", 4.0, 68.0),
    ]


def test_la_tabla_de_proyectos_cierra_con_el_total_del_mes():
    filas = [fila(1, 4.0), fila(2, 6.0, proyecto="Subastas"), fila(3, 1.5)]
    totales = por_proyecto(filas, {})
    assert sum(t.horas for t in totales) == total_de(filas)


def test_un_proyecto_sin_valor_hora_sale_sin_importe_en_vez_de_inventarlo():
    (unico,) = por_proyecto([fila(1, 4.0)], {})
    assert unico.valor_hora is None and unico.importe is None


# --- Tabla 2: horas por persona ---------------------------------------------


def test_la_tabla_de_personas_cuenta_horas_dias_y_promedio():
    filas = [fila(1, 4.0), fila(1, 2.0), fila(2, 3.0)]
    (unica,) = por_persona(filas, {}, "Desarrollador")
    assert (unica.horas, unica.dias) == (9.0, 2)
    assert unica.promedio == 4.5
    assert unica.perfil == "Desarrollador"


def test_el_perfil_declarado_de_una_persona_gana_sobre_el_general():
    (unica,) = por_persona([fila(1, 4.0)], {"Ana Perez": "Líder técnica"}, "Desarrollador")
    assert unica.perfil == "Líder técnica"


def test_la_tabla_de_personas_cierra_con_el_total_del_mes():
    filas = [fila(1, 4.0), fila(1, 2.0, persona="Beto Suarez"), fila(2, 3.0)]
    totales = por_persona(filas, {}, "Desarrollador")
    assert sum(t.horas for t in totales) == total_de(filas)


# --- Tabla 3: el borrador de los principales trabajos -----------------------


def test_el_ticket_se_saca_de_la_descripcion():
    assert ticket_de("R-012532 – corrección del error") == "R-012532"
    assert ticket_de("CU-8abc tarea de ClickUp") == "CU-8abc"
    assert ticket_de("BUG-107 mal clasificado") == "BUG-107"
    assert ticket_de("reunión de coordinación") == "Sin ticket"
    assert ticket_de("") == "Sin ticket"


def test_los_trabajos_se_agrupan_por_proyecto_y_ticket():
    filas = [
        fila(1, 4.0, descripcion="R-01 una cosa"),
        fila(2, 2.0, descripcion="R-01 otra cosa"),
        fila(3, 1.0, descripcion="sin ticket acá"),
    ]
    trabajos = principales_trabajos(filas)
    por_ticket = {t.ticket: t for t in trabajos}
    assert por_ticket["R-01"].horas == 6.0
    assert "una cosa" in por_ticket["R-01"].descripcion
    assert "otra cosa" in por_ticket["R-01"].descripcion
    assert por_ticket["Sin ticket"].horas == 1.0


def test_el_borrador_de_trabajos_cierra_con_el_total_del_mes():
    """Cada fila cae en exactamente un grupo: los subtotales tienen que cerrar."""
    filas = [
        fila(1, 4.0, descripcion="R-01 algo"),
        fila(2, 2.5, descripcion="sin ticket"),
        fila(3, 1.25, proyecto="Subastas", descripcion="R-02 otra"),
        fila(4, 0.25, proyecto="Subastas", descripcion=""),
    ]
    trabajos = principales_trabajos(filas)
    assert sum(t.horas for t in trabajos) == total_de(filas) == 8.0


def test_las_descripciones_repetidas_no_se_repiten_en_el_borrador():
    filas = [fila(1, 1.0, descripcion="R-01 lo mismo"), fila(2, 1.0, descripcion="R-01 lo mismo")]
    (unico,) = principales_trabajos(filas)
    assert unico.descripcion == "R-01 lo mismo"


# --- Aviso de registros sin descripción -------------------------------------


def test_se_cuentan_los_registros_sin_descripcion_por_persona():
    filas = [
        fila(1, 1.0, descripcion=""),
        fila(2, 1.0, descripcion="   "),
        fila(3, 1.0, descripcion="R-01 algo"),
        fila(4, 1.0, persona="Beto Suarez", descripcion=""),
    ]
    assert sin_descripcion(filas) == (("Ana Perez", 2), ("Beto Suarez", 1))


def test_sin_registros_vacios_no_hay_nada_que_avisar():
    assert sin_descripcion([fila(1, 1.0, descripcion="R-01 algo")]) == ()
