"""Lo que el resumen mensual no trae y el mapeo completa.

Dos reglas que no se cruzan:

- Un registro del **resumen mensual** saca del mapeo el usuario de Kimai, que
  es lo único que le falta y que los anexos escriben. Si el mapeo no lo
  resuelve, el usuario queda vacío y sale un aviso: ese archivo entra igual.
- Un registro del **reporte de detalle** no mira el mapeo: sus valores son los
  de Kimai.

Ni el mail ni el número de proyecto se completan: el Anexo II-A no tiene esas
columnas, así que el mapeo ya no los declara.
"""
from datetime import date, time

import yaml

from cunix_horas.completado import completar_desde_mapeo
from cunix_horas.filas_anexo import construir
from cunix_horas.kimai_comun import (
    ORIGEN_DETALLE,
    ORIGEN_RESUMEN_MENSUAL,
    Registro,
)
from cunix_horas.mapeo import DestinoProyecto, Mapeo, Persona

ARCHIVO = "lautaro.xlsx"

PERSONA = Persona(
    nombre="Lautaro Zalazar",
    archivo="L Zalazar",
    username="lzalazar",
)

PROYECTO = DestinoProyecto("Sistemas - C.UNIX", "Victorius 3")


def mapeo(personas=(PERSONA,), proyecto=PROYECTO):
    return Mapeo({p.username: p for p in personas}, {"GI2680001": proyecto})


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


# --- El resumen mensual saca del mapeo el usuario de Kimai ------------------


def test_completa_el_usuario_de_kimai():
    completado = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO)
    assert completado.registros[0].username == "lzalazar"
    assert completado.avisos == ()


def test_el_name_sale_del_propio_archivo():
    """El nombre para mostrar lo trae el export; el mapeo no lo pisa."""
    completado = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO)
    assert completado.registros[0].nombre == "Lautaro Zalazar"


def test_no_muta_el_registro_que_recibe():
    original = del_resumen()
    completar_desde_mapeo([original], mapeo(), ARCHIVO)
    assert original.username == "Lautaro Zalazar"
    assert original.nombre == ""


def test_la_fila_del_anexo_sale_sin_descripcion_y_sin_horario():
    """Las tres columnas que el resumen mensual no trae y la plantilla acepta."""
    registros = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO).registros
    fila = construir(list(registros), mapeo()).filas[0]

    assert fila.descripcion == ""
    assert fila.inicio is None and fila.fin is None
    assert fila.fecha == date(2026, 8, 3)


def test_la_fila_del_anexo_sale_con_la_persona_y_el_proyecto_resueltos():
    registros = completar_desde_mapeo([del_resumen()], mapeo(), ARCHIVO).registros
    armado = construir(list(registros), mapeo())

    assert armado.filas[0].persona == "Lautaro Zalazar"
    assert armado.filas[0].proyecto == "Victorius 3"
    # El usuario lo completó el mapeo sobre el registro: va a la hoja Datos,
    # no a la fila del Detalle.
    assert armado.usuarios == (("Lautaro Zalazar", "lzalazar"),)


# --- Lo que el mapeo no tiene AVISA, no frena -------------------------------


def test_una_persona_que_no_esta_en_el_mapeo_entra_igual_sin_usuario():
    completado = completar_desde_mapeo(
        [del_resumen(username="Alexis Carnero")], mapeo(), ARCHIVO
    )
    registro = completado.registros[0]
    assert registro.nombre == "Alexis Carnero"
    assert registro.username == ""
    assert registro.horas == 2.0
    assert len(completado.avisos) == 1


def test_el_aviso_nombra_a_la_persona_el_archivo_y_por_que_importa():
    completado = completar_desde_mapeo(
        [del_resumen(username="Alexis Carnero")], mapeo(), ARCHIVO
    )
    (aviso,) = completado.avisos
    assert "Alexis Carnero" in aviso
    assert ARCHIVO in aviso
    assert "resumen mensual" in aviso
    assert "su propio usuario" in aviso
    assert "no frena nada" in aviso


def test_el_aviso_trae_el_bloque_yaml_pegable():
    completado = completar_desde_mapeo(
        [del_resumen(username="Alexis Carnero")], mapeo(), ARCHIVO
    )
    datos = yaml.safe_load(
        _bloque_yaml(completado.avisos[0], "AJUSTAR-username-de-kimai:")
    )
    entrada = datos["AJUSTAR-username-de-kimai"]
    assert entrada["nombre"] == "Alexis Carnero"
    assert "archivo" in entrada
    # El mail ya no se pide: no hay columna donde escribirlo.
    assert "mail" not in entrada


def test_la_fila_de_una_persona_sin_mapear_sale_completa_igual():
    registros = completar_desde_mapeo(
        [del_resumen(username="Alexis Carnero")], mapeo(), ARCHIVO
    ).registros
    armado = construir(list(registros), mapeo())

    assert armado.filas[0].persona == "Alexis Carnero"
    assert armado.filas[0].horas == 2.0
    # Igual que Gabriel Denis, que no tiene usuario de Kimai: celda vacía.
    assert armado.usuarios == (("Alexis Carnero", ""),)


def test_avisa_una_sola_vez_por_persona_aunque_tenga_muchos_registros():
    sin_mapear = [del_resumen(username="Alexis Carnero") for _ in range(5)]
    assert len(completar_desde_mapeo(sin_mapear, mapeo(), ARCHIVO).avisos) == 1


def test_dos_personas_con_el_mismo_nombre_avisan_en_vez_de_frenar():
    repetidas = (
        Persona("Lautaro Zalazar", "L Zalazar", "lzalazar"),
        Persona("Lautaro Zalazar", "Zalazar L", "lzalazar2"),
    )
    completado = completar_desde_mapeo(
        [del_resumen()], mapeo(personas=repetidas), ARCHIVO
    )
    (aviso,) = completado.avisos
    assert completado.registros[0].username == ""
    assert "lzalazar" in aviso and "lzalazar2" in aviso
    assert "no frena nada" in aviso


def test_un_proyecto_que_no_esta_en_el_mapeo_tampoco_frena():
    """El Anexo II-A no tiene columna de número de proyecto: nada que pedir."""
    completado = completar_desde_mapeo(
        [del_resumen(cod_proyecto="XX9999999")], mapeo(), ARCHIVO
    )
    assert completado.avisos == ()
    assert completado.registros[0].horas == 2.0


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
    """El mapeo declara otro usuario para esa persona y Kimai manda."""
    completado = completar_desde_mapeo([DEL_DETALLE], mapeo(), "matias.xlsx")

    assert completado.registros[0] is DEL_DETALLE
    assert completado.registros[0].username == "el-de-kimai"
    assert completado.avisos == ()


def test_el_detalle_no_avisa_ni_frena_aunque_el_mapeo_este_vacio():
    """Un mapeo vacío no es problema del detalle: sus datos vienen de Kimai."""
    completado = completar_desde_mapeo([DEL_DETALLE], Mapeo({}, {}), "matias.xlsx")
    assert completado.registros == (DEL_DETALLE,)
    assert completado.avisos == ()


def test_los_dos_origenes_conviven_en_la_misma_lista():
    completado = completar_desde_mapeo(
        [DEL_DETALLE, del_resumen()], mapeo(), "mezcla.xlsx"
    )
    assert completado.registros[0].username == "el-de-kimai"
    assert completado.registros[1].username == "lzalazar"
