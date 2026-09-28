"""Armado del detalle plano: orden, nombres derivados y avisos de proyecto."""
from datetime import date, datetime, time, timedelta

from cunix_horas.detalle import construir, segundos_de
from cunix_horas.kimai_comun import Registro
from cunix_horas.mapeo import DestinoProyecto, Mapeo, derivar_cliente, derivar_proyecto

PROYECTOS = {
    "CO2510115": DestinoProyecto("Instituto de Salud Pública de Chile", "SIAC-OIRS"),
}


def _mapeo(proyectos=None):
    return Mapeo({}, PROYECTOS if proyectos is None else proyectos)


def _registro(
    dia=3,
    horas=1.0,
    username="mzalazar",
    cod_proyecto="CO2510115",
    texto_proyecto="[CO2510115] ISPCH-SopEvo-SIAC | descripción larga",
    hora_inicio=time(9, 0),
    nombre="Matias Zalazar",
    numero_proyecto="CO2510115",
    texto_cliente="[616050001] Instituto de Salud Publica de Chile",
    descripcion="Ticket R-012528",
):
    return Registro(
        fecha=date(2026, 8, dia),
        horas=horas,
        username=username,
        cod_proyecto=cod_proyecto,
        actividad="Desarrollo",
        texto_proyecto=texto_proyecto,
        hora_inicio=hora_inicio,
        nombre=nombre,
        email=f"{username}@cunix.net",
        descripcion=descripcion,
        numero_proyecto=numero_proyecto,
        texto_cliente=texto_cliente,
    )


# --- Passthrough y derivación de las dos columnas que no son passthrough ----


def test_la_fila_combina_la_fecha_con_la_hora_de_inicio():
    fila = construir([_registro(dia=5, hora_inicio=time(14, 30))], _mapeo()).filas[0]
    assert fila.fecha_hora == datetime(2026, 8, 5, 14, 30)


def test_la_duracion_es_un_timedelta_y_no_un_numero():
    fila = construir([_registro(horas=1.5)], _mapeo()).filas[0]
    assert fila.duracion == timedelta(hours=1, minutes=30)
    assert not isinstance(fila.duracion, float)


def test_un_proyecto_mapeado_sale_con_el_nombre_del_mapeo():
    fila = construir([_registro()], _mapeo()).filas[0]
    assert fila.cliente == "Instituto de Salud Pública de Chile"
    assert fila.proyecto == "SIAC-OIRS"


def test_un_proyecto_sin_mapear_sale_con_el_nombre_derivado_de_kimai():
    detalle = construir([_registro()], _mapeo(proyectos={}))
    fila = detalle.filas[0]
    assert fila.cliente == "Instituto de Salud Publica de Chile"
    assert fila.proyecto == "ISPCH-SopEvo-SIAC"
    assert [p.codigo for p in detalle.sin_mapear] == ["CO2510115"]
    assert detalle.sin_mapear[0].proyecto == "ISPCH-SopEvo-SIAC"


def test_un_proyecto_mapeado_no_aparece_como_sin_mapear():
    assert construir([_registro()], _mapeo()).sin_mapear == ()


def test_el_cliente_sin_corchetes_se_deriva_igual():
    """El caso de Luciano: su Customer en Kimai viene como 'CUNIX', sin código.

    Si la derivación asumiera el prefijo, el partner recibiría la columna
    Customer en blanco para todo ese desarrollador.
    """
    assert derivar_cliente("CUNIX") == "CUNIX"
    fila = construir(
        [_registro(texto_cliente="CUNIX")], _mapeo(proyectos={})
    ).filas[0]
    assert fila.cliente == "CUNIX"


def test_el_proyecto_derivado_corta_en_la_barra():
    assert derivar_proyecto("[AD2690002] C.UNIX - Internos | largo") == (
        "C.UNIX - Internos"
    )


def test_el_project_number_sale_del_campo_propio_y_no_del_corchete():
    """El caso de Luciano: código '[AD2690002]' y Project number '210'.

    Son dos campos distintos de Kimai que se parecen. El del corchete es la
    clave del mapeo; el que ve el partner es el otro.
    """
    fila = construir(
        [
            _registro(
                cod_proyecto="AD2690002",
                texto_proyecto="[AD2690002] C.UNIX - Proyectos Internos | VictoriusCP2",
                numero_proyecto="210",
            )
        ],
        _mapeo(proyectos={}),
    ).filas[0]
    assert fila.numero_proyecto == "210"
    assert "AD2690002" not in fila.numero_proyecto


def test_la_descripcion_vacia_deja_la_celda_vacia():
    fila = construir([_registro(descripcion="")], _mapeo()).filas[0]
    assert fila.valores[8] is None


# --- Orden ------------------------------------------------------------------


def test_las_filas_se_agrupan_por_dev_alfabetico_y_van_cronologicas():
    registros = [
        _registro(dia=5, nombre="Matias Zalazar", username="mzalazar"),
        _registro(dia=9, nombre="Alexis Carnero", username="acarnero"),
        _registro(dia=2, nombre="Matias Zalazar", username="mzalazar"),
        _registro(dia=1, nombre="Alexis Carnero", username="acarnero"),
    ]
    filas = construir(registros, _mapeo()).filas
    assert [(f.nombre, f.fecha_hora.day) for f in filas] == [
        ("Alexis Carnero", 1),
        ("Alexis Carnero", 9),
        ("Matias Zalazar", 2),
        ("Matias Zalazar", 5),
    ]


def test_dentro_de_un_dia_ordena_por_hora_de_inicio():
    registros = [
        _registro(dia=3, hora_inicio=time(17, 0)),
        _registro(dia=3, hora_inicio=time(9, 0)),
    ]
    filas = construir(registros, _mapeo()).filas
    assert [f.fecha_hora.hour for f in filas] == [9, 17]


def test_un_registro_sin_hora_de_inicio_no_rompe_el_orden():
    registros = [
        _registro(dia=3, hora_inicio=time(9, 0)),
        _registro(dia=3, hora_inicio=None),
    ]
    filas = construir(registros, _mapeo()).filas
    assert [f.fecha_hora.hour for f in filas] == [0, 9]


# --- Integridad y avisos ----------------------------------------------------


def test_el_total_en_segundos_no_pierde_horas_por_el_float_de_kimai():
    """Kimai trae 1 h como 1.000000000000008: el redondeo es a segundos."""
    detalle = construir([_registro(horas=1.000000000000008)] * 3, _mapeo())
    assert detalle.segundos == 3 * 3600
    assert detalle.horas == 3.0
    assert segundos_de(1.000000000000008) == 3600


def test_un_project_number_con_dos_nombres_de_proyecto_se_avisa():
    """Alguien renombró el proyecto en Kimai a mitad de mes."""
    registros = [
        _registro(
            cod_proyecto="AA1",
            texto_proyecto="[AA1] Portal viejo | largo",
            numero_proyecto="210",
        ),
        _registro(
            cod_proyecto="AA2",
            texto_proyecto="[AA2] Portal nuevo | largo",
            numero_proyecto="210",
        ),
    ]
    detalle = construir(registros, _mapeo(proyectos={}))
    assert detalle.numeros_ambiguos == (("210", ("Portal nuevo", "Portal viejo")),)


def test_un_project_number_con_un_solo_nombre_no_se_avisa():
    detalle = construir([_registro(), _registro(dia=4)], _mapeo())
    assert detalle.numeros_ambiguos == ()
