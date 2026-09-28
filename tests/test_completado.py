"""Lo que el resumen mensual no trae y el mapeo completa.

Dos reglas que no se cruzan:

- Un registro del **resumen mensual** se completa desde el mapeo, y si al
  mapeo le falta el usuario, el mail o el número de proyecto, ese archivo no
  entra, con el bloque YAML listo para pegar.
- Un registro del **reporte de detalle** no mira el mapeo ni aunque el mapeo
  declare esos campos: sus valores son los de Kimai.
"""
from datetime import date, time

import pytest
import yaml

from cunix_horas.completado import completar_desde_mapeo
from cunix_horas.detalle import construir
from cunix_horas.kimai_comun import (
    ORIGEN_DETALLE,
    ORIGEN_RESUMEN_MENSUAL,
    Registro,
)
from cunix_horas.mapeo import DestinoProyecto, ErrorMapeo, Mapeo, Persona

ARCHIVO = "lautaro.xlsx"

PERSONA = Persona(
    nombre="Lautaro Zalazar",
    archivo="L Zalazar",
    mail="lautaro.zalazar@cunix.net",
    username="lzalazar",
)

PROYECTO = DestinoProyecto("Sistemas - C.UNIX", "Victorius 3", "GI-123")


def mapeo(persona=PERSONA, proyecto=PROYECTO):
    return Mapeo({persona.username: persona}, {"GI2680001": proyecto})


def del_resumen(**cambios):
    """Un registro tal como lo deja el lector de resumen mensual."""
    base = dict(
        fecha=date(2026, 8, 3),
        horas=2.0,
        # El lector deja ahí el nombre para mostrar: es lo único que trae.
        username="Lautaro Zalazar",
        cod_proyecto="GI2680001",
        actividad="Coordinación interna",
        texto_proyecto="[GI2680001] C.UNIX - Victorius 3",
        origen=ORIGEN_RESUMEN_MENSUAL,
    )
    base.update(cambios)
    return Registro(**base)


def _bloque_yaml(mensaje: str, clave: str) -> str:
    """Las líneas del mensaje desde `clave` mientras sigan indentadas."""
    lineas = mensaje.splitlines()
    inicio = next(i for i, linea in enumerate(lineas) if linea.strip() == clave)
    bloque = [lineas[inicio].strip()]
    for linea in lineas[inicio + 1:]:
        if not linea.startswith("    "):
            break
        bloque.append("  " + linea.strip())
    return "\n".join(bloque)


# --- El resumen mensual se completa desde el mapeo --------------------------


def test_completa_el_user_el_mail_y_el_project_number():
    completado = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO)[0]
    assert completado.username == "lzalazar"
    assert completado.email == "lautaro.zalazar@cunix.net"
    assert completado.numero_proyecto == "GI-123"


def test_el_name_sale_del_propio_archivo():
    """El nombre para mostrar lo trae el export; el mapeo no lo pisa."""
    completado = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO)[0]
    assert completado.nombre == "Lautaro Zalazar"


def test_no_muta_el_registro_que_recibe():
    original = del_resumen()
    completar_desde_mapeo([original], mapeo(), ARCHIVO)
    assert original.username == "Lautaro Zalazar"
    assert original.email == ""
    assert original.numero_proyecto == ""


def test_la_fila_sale_con_la_descripcion_vacia_y_la_fecha_sin_hora():
    """Las dos columnas que el partner acepta vacías."""
    registros = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO)
    fila = construir(registros, mapeo()).filas[0]

    assert fila.descripcion == ""
    assert fila.valores[8] is None  # la celda va vacía, no con un texto vacío
    assert fila.fecha_hora.date() == date(2026, 8, 3)
    assert fila.fecha_hora.time() == time(0, 0)


def test_la_fila_sale_con_las_diez_columnas_llenas_salvo_la_descripcion():
    registros = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO)
    fila = construir(registros, mapeo()).filas[0]

    assert fila.nombre == "Lautaro Zalazar"
    assert fila.username == "lzalazar"
    assert fila.email == "lautaro.zalazar@cunix.net"
    assert fila.cliente == "Sistemas - C.UNIX"
    assert fila.proyecto == "Victorius 3"
    assert fila.numero_proyecto == "GI-123"


# --- Lo que el mapeo no tiene frena ese archivo -----------------------------


def test_sin_mail_el_archivo_no_entra_y_el_mensaje_nombra_a_la_persona():
    sin_mail = Persona("Lautaro Zalazar", "L Zalazar", "", "lzalazar")
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo([del_resumen()], mapeo(persona=sin_mail), ARCHIVO)

    mensaje = str(excepcion.value)
    assert "Lautaro Zalazar" in mensaje
    assert "E-mail" in mensaje
    assert ARCHIVO in mensaje
    assert "resumen mensual" in mensaje


def test_el_error_de_mail_trae_el_bloque_yaml_pegable():
    sin_mail = Persona("Lautaro Zalazar", "L Zalazar", "", "lzalazar")
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo([del_resumen()], mapeo(persona=sin_mail), ARCHIVO)

    datos = yaml.safe_load(_bloque_yaml(str(excepcion.value), "lzalazar:"))
    assert datos["lzalazar"]["nombre"] == "Lautaro Zalazar"
    assert "mail" in datos["lzalazar"]


def test_sin_numero_de_proyecto_el_archivo_no_entra():
    sin_numero = DestinoProyecto("Sistemas - C.UNIX", "Victorius 3")
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo([del_resumen()], mapeo(proyecto=sin_numero), ARCHIVO)

    mensaje = str(excepcion.value)
    assert "GI2680001" in mensaje
    assert "Victorius 3" in mensaje
    assert "Project number" in mensaje
    assert ARCHIVO in mensaje


def test_el_error_de_numero_de_proyecto_trae_el_bloque_yaml_pegable():
    sin_numero = DestinoProyecto("Sistemas - C.UNIX", "Victorius 3")
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo([del_resumen()], mapeo(proyecto=sin_numero), ARCHIVO)

    datos = yaml.safe_load(_bloque_yaml(str(excepcion.value), "GI2680001:"))
    assert datos["GI2680001"]["proyecto"] == "Victorius 3"
    assert "numero_proyecto" in datos["GI2680001"]


def test_un_proyecto_que_no_esta_en_el_mapeo_tambien_frena():
    """Sin entrada no hay número, y el número no se deriva de nada."""
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo(
            [del_resumen(cod_proyecto="XX9999999")], mapeo(), ARCHIVO
        )
    assert "XX9999999" in str(excepcion.value)


def test_una_persona_que_no_esta_en_el_mapeo_frena_con_su_nombre():
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo(
            [del_resumen(username="Alexis Carnero")], mapeo(), ARCHIVO
        )
    mensaje = str(excepcion.value)
    assert "Alexis Carnero" in mensaje
    assert "mail:" in mensaje


def test_el_mensaje_aclara_que_los_demas_siguen():
    sin_mail = Persona("Lautaro Zalazar", "L Zalazar", "", "lzalazar")
    with pytest.raises(ErrorMapeo) as excepcion:
        completar_desde_mapeo([del_resumen()], mapeo(persona=sin_mail), ARCHIVO)
    assert "Los demás desarrolladores se procesan igual" in str(excepcion.value)


# --- El reporte de detalle no mira el mapeo ---------------------------------


DEL_DETALLE = Registro(
    fecha=date(2026, 8, 3),
    horas=2.0,
    username="el-de-kimai",
    cod_proyecto="GI2680001",
    actividad="Desarrollo",
    texto_proyecto="[GI2680001] C.UNIX - Victorius 3",
    hora_inicio=time(9, 30),
    nombre="Lautaro Zalazar",
    email="el-mail-de-kimai@cunix.net",
    descripcion="Lo que hizo ese día",
    numero_proyecto="el-numero-de-kimai",
    origen=ORIGEN_DETALLE,
)


def test_el_detalle_sale_tal_cual_aunque_el_mapeo_diga_otra_cosa():
    """El mapeo declara usuario, mail y número, y Kimai dice otros tres."""
    completado = completar_desde_mapeo([DEL_DETALLE], mapeo(), "matias.xlsx")[0]

    assert completado is DEL_DETALLE
    assert completado.username == "el-de-kimai"
    assert completado.email == "el-mail-de-kimai@cunix.net"
    assert completado.numero_proyecto == "el-numero-de-kimai"


def test_el_detalle_no_frena_aunque_al_mapeo_le_falte_todo():
    """Un mapeo vacío no es problema del detalle: sus datos vienen de Kimai."""
    vacio = Mapeo({}, {})
    assert completar_desde_mapeo([DEL_DETALLE], vacio, "matias.xlsx") == [DEL_DETALLE]


def test_los_dos_origenes_conviven_en_la_misma_lista():
    completados = completar_desde_mapeo(
        [DEL_DETALLE, del_resumen()], mapeo(), "mezcla.xlsx"
    )
    assert completados[0].username == "el-de-kimai"
    assert completados[1].username == "lzalazar"
