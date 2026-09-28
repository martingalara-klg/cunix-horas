import re

import pytest
import yaml
from conftest import FIXTURES

from cunix_horas.mapeo import ErrorMapeo, Mapeo


def cargar():
    return Mapeo.cargar(FIXTURES / "mapeo-test.yaml")


CUERPO_MINIMO = """
personas:
  mzalazar:
    nombre: "Matias Zalazar"
    archivo: "Zalazar"
proyectos:
  CO2610170:
    cliente: "Servicio Nacional de Aduanas"
    proyecto: "Subastas"
"""



def test_resolver_proyecto_conocido():
    destino = cargar().resolver_proyecto("CO2610170", "[CO2610170] Aduana", "x.xlsx")
    assert destino.cliente == "Servicio Nacional de Aduanas"
    assert destino.proyecto == "Subastas"


def test_resolver_proyecto_conserva_acentos():
    destino = cargar().resolver_proyecto("CO2510115", "[CO2510115] ISPCH", "x.xlsx")
    assert destino.cliente == "Instituto de Salud Pública de Chile"


def test_resolver_persona_conocida():
    persona = cargar().resolver_persona("mzalazar", "x.xlsx")
    assert persona.nombre == "Matias Zalazar"
    assert persona.archivo == "Zalazar"


def test_proyecto_desconocido_sugiere_la_linea_yaml():
    texto = "[PR2610199] MINVU-Portal2 | Subsecretaría de Vivienda - Portal"
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_proyecto("PR2610199", texto, "kimai-mzalazar.xlsx")
    mensaje = str(excepcion.value)
    assert "PR2610199" in mensaje
    assert "kimai-mzalazar.xlsx" in mensaje
    assert "config/mapeo.yaml" in mensaje
    assert "MINVU-Portal2" in mensaje
    assert "cliente:" in mensaje
    assert "proyecto:" in mensaje


def test_persona_desconocida_sugiere_la_linea_yaml():
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_persona("fdodera", "kimai-fdodera.xlsx")
    mensaje = str(excepcion.value)
    assert "fdodera" in mensaje
    assert "config/mapeo.yaml" in mensaje
    assert "nombre:" in mensaje
    assert "archivo:" in mensaje


def test_falla_si_el_yaml_no_existe(tmp_path):
    with pytest.raises(ErrorMapeo, match="No se encontró"):
        Mapeo.cargar(tmp_path / "no-existe.yaml")


def test_falla_si_al_yaml_le_falta_una_seccion(tmp_path):
    ruta = tmp_path / "incompleto.yaml"
    ruta.write_text("personas: {}\n", encoding="utf-8")
    with pytest.raises(ErrorMapeo, match="proyectos"):
        Mapeo.cargar(ruta)


def test_falla_si_a_un_proyecto_le_falta_el_cliente(tmp_path):
    ruta = tmp_path / "malo.yaml"
    ruta.write_text(
        "personas: {}\nproyectos:\n  AB123:\n    proyecto: 'X'\n", encoding="utf-8"
    )
    with pytest.raises(ErrorMapeo, match="cliente"):
        Mapeo.cargar(ruta)


def test_proyecto_desconocido_sugiere_yaml_pegable(tmp_path):
    """Verifica que el YAML sugerido en el error es pegable y válido."""
    texto = "[PR2610199] MINVU-Portal2 | Subsecretaría de Vivienda - Portal"
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_proyecto("PR2610199", texto, "test.xlsx")

    mensaje = str(excepcion.value)
    # Extrae las líneas del YAML sugerido (desde la clave hasta el final)
    # Busca líneas que comienzan con "  " (2 espacios) o más
    lineas = mensaje.split("\n")
    inicio_bloque = None
    for i, linea in enumerate(lineas):
        if "PR2610199:" in linea:
            inicio_bloque = i
            break

    assert inicio_bloque is not None, "No se encontró la clave en el mensaje"

    # Extrae el bloque YAML sugerido
    bloque_yaml = "\n".join(lineas[inicio_bloque:])

    # Crea un archivo de prueba con la estructura base
    ruta = tmp_path / "test.yaml"
    contenido_base = """personas:
  mzalazar:
    nombre: "Test"
    archivo: "Test"

proyectos:
  CO2610170:
    cliente: "Existente"
    proyecto: "Existente"
"""
    ruta.write_text(contenido_base + bloque_yaml + "\n", encoding="utf-8")

    # Carga el YAML y verifica que la entrada nueva está como hermana
    datos = yaml.safe_load(ruta.read_text(encoding="utf-8"))
    assert "PR2610199" in datos["proyectos"], "La clave nueva no aparece en proyectos"
    assert isinstance(
        datos["proyectos"]["PR2610199"], dict
    ), "La entrada no es un dict"
    assert "cliente" in datos["proyectos"]["PR2610199"]
    assert "proyecto" in datos["proyectos"]["PR2610199"]


def test_persona_desconocida_sugiere_yaml_pegable(tmp_path):
    """Verifica que el YAML sugerido en el error es pegable y válido."""
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_persona("fjohnson", "test.xlsx")

    mensaje = str(excepcion.value)
    # Extrae las líneas del YAML sugerido
    lineas = mensaje.split("\n")
    inicio_bloque = None
    for i, linea in enumerate(lineas):
        if "fjohnson:" in linea:
            inicio_bloque = i
            break

    assert inicio_bloque is not None, "No se encontró la clave en el mensaje"

    # Extrae el bloque YAML sugerido
    bloque_yaml = "\n".join(lineas[inicio_bloque:])

    # Crea un archivo de prueba con la estructura base
    # Importante: insertar el bloque al final de la sección personas, antes de proyectos
    ruta = tmp_path / "test.yaml"
    contenido_base = """personas:
  mzalazar:
    nombre: "Test"
    archivo: "Test"
"""
    # Agrega el nuevo bloque a la sección personas
    contenido = contenido_base + bloque_yaml + "\n"
    # Agrega la sección proyectos después
    contenido += """
proyectos:
  CO2610170:
    cliente: "Test"
    proyecto: "Test"
"""
    ruta.write_text(contenido, encoding="utf-8")

    # Carga el YAML y verifica que la entrada nueva está como hermana
    datos = yaml.safe_load(ruta.read_text(encoding="utf-8"))
    assert "fjohnson" in datos["personas"], "La clave nueva no aparece en personas"
    assert isinstance(
        datos["personas"]["fjohnson"], dict
    ), "La entrada no es un dict"
    assert "nombre" in datos["personas"]["fjohnson"]
    assert "archivo" in datos["personas"]["fjohnson"]


# --- El resumen mensual trae el nombre para mostrar, no el username --------


def _mapeo_con(personas: str, tmp_path):
    ruta = tmp_path / "mapeo.yaml"
    ruta.write_text(
        "personas:\n" + personas + "\nproyectos:\n"
        "  CO2610170:\n"
        '    cliente: "Aduanas"\n'
        '    proyecto: "Subastas"\n',
        encoding="utf-8",
    )
    return Mapeo.cargar(ruta)


def test_resolver_persona_por_el_nombre_para_mostrar():
    """El resumen mensual identifica al dev por el 'nombre:' del mapeo."""
    persona = cargar().resolver_persona("Matias Zalazar", "resumen.xlsx")
    assert persona.archivo == "Zalazar"


def test_el_username_gana_sobre_el_nombre(tmp_path):
    """Primero se busca por clave; el nombre es el recurso de después."""
    mapeo = _mapeo_con(
        "  mzalazar:\n"
        '    nombre: "Matias Zalazar"\n'
        '    archivo: "Zalazar"\n'
        "  Matias Zalazar:\n"
        '    nombre: "Otro Dev"\n'
        '    archivo: "Otro"\n',
        tmp_path,
    )
    assert mapeo.resolver_persona("Matias Zalazar", "x.xlsx").archivo == "Otro"


def test_el_nombre_se_compara_sin_importar_espacios_ni_mayusculas():
    persona = cargar().resolver_persona("  matias   zalazar ", "resumen.xlsx")
    assert persona.archivo == "Zalazar"


def test_dos_personas_con_el_mismo_nombre_no_eligen_una(tmp_path):
    mapeo = _mapeo_con(
        "  mzalazar:\n"
        '    nombre: "Matias Zalazar"\n'
        '    archivo: "Zalazar"\n'
        "  mzalazar2:\n"
        '    nombre: "Matias Zalazar"\n'
        '    archivo: "Zalazar M"\n',
        tmp_path,
    )
    with pytest.raises(ErrorMapeo) as excepcion:
        mapeo.resolver_persona("Matias Zalazar", "resumen.xlsx")
    mensaje = str(excepcion.value)
    assert "ambiguo" in mensaje
    assert "mzalazar" in mensaje
    assert "mzalazar2" in mensaje


def test_un_nombre_sin_mapear_sugiere_un_yaml_pegable_con_ese_nombre(tmp_path):
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_persona("Lautaro Zalazar", "resumen.xlsx")
    mensaje = str(excepcion.value)
    assert "Lautaro Zalazar" in mensaje
    assert "config/mapeo.yaml" in mensaje

    # El bloque sugerido tiene que ser YAML pegable bajo personas:
    lineas = mensaje.split("\n")
    inicio = next(
        i for i, linea in enumerate(lineas) if "AJUSTAR-username-de-kimai:" in linea
    )
    bloque = "\n".join(lineas[inicio:])
    ruta = tmp_path / "pegado.yaml"
    ruta.write_text("personas:\n" + bloque + "\n", encoding="utf-8")
    datos = yaml.safe_load(ruta.read_text(encoding="utf-8"))
    entrada = next(iter(datos["personas"].values()))
    assert entrada["nombre"] == "Lautaro Zalazar"
    assert "archivo" in entrada


# --- Texto de la única fila de actividad ---------------------------------


def test_actividad_por_defecto_cuando_el_yaml_no_la_declara():
    """mapeo-test.yaml no tiene 'actividad:': vale lo que el partner vio siempre."""
    assert cargar().actividad == "Desarrollo"


def test_actividad_declarada_en_el_yaml(tmp_path):
    ruta = tmp_path / "mapeo.yaml"
    ruta.write_text(
        'actividad: "Servicios profesionales"'
        + CUERPO_MINIMO,
        encoding="utf-8",
    )
    assert Mapeo.cargar(ruta).actividad == "Servicios profesionales"


@pytest.mark.parametrize("valor", ['""', '"   "', "", "7"])
def test_actividad_vacia_o_no_textual_falla(tmp_path, valor):
    ruta = tmp_path / "mapeo.yaml"
    ruta.write_text("actividad: " + valor + CUERPO_MINIMO, encoding="utf-8")
    with pytest.raises(ErrorMapeo) as excepcion:
        Mapeo.cargar(ruta)
    mensaje = str(excepcion.value)
    assert "actividad:" in mensaje
    assert "Desarrollo" in mensaje


# --- 'personas:' dejó de ser obligatorio ------------------------------------
# El archivo que recibe el partner trae el nombre, el usuario y el mail de
# cada desarrollador tal como vienen de Kimai: el mapeo ya no los necesita.


def test_el_mapeo_carga_sin_la_seccion_personas(tmp_path):
    ruta = tmp_path / "sin-personas.yaml"
    ruta.write_text(
        'proyectos:\n  AB123:\n    cliente: "Cliente"\n    proyecto: "Proyecto"\n',
        encoding="utf-8",
    )
    mapeo = Mapeo.cargar(ruta)
    assert mapeo.proyecto_opcional("AB123").proyecto == "Proyecto"


def test_un_proyecto_sin_declarar_devuelve_none_en_vez_de_frenar(tmp_path):
    ruta = tmp_path / "sin-personas.yaml"
    ruta.write_text("proyectos:\n", encoding="utf-8")
    assert Mapeo.cargar(ruta).proyecto_opcional("NO-ESTA") is None


# --- El nombre del archivo que recibe el partner ----------------------------


def test_el_patron_del_archivo_de_salida_tiene_un_valor_por_defecto(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text("proyectos:\n", encoding="utf-8")
    assert Mapeo.cargar(ruta).archivo_salida == "Horas KLG-{mes}{anio}.xlsx"


def test_el_patron_del_archivo_de_salida_se_puede_cambiar(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        'proyectos:\narchivo_salida: "Horas KLG-Sept{anio}.xlsx"\n', encoding="utf-8"
    )
    assert Mapeo.cargar(ruta).archivo_salida == "Horas KLG-Sept{anio}.xlsx"


def test_un_patron_con_un_reemplazo_inventado_falla_en_espanol(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text('proyectos:\narchivo_salida: "Horas {dia}.xlsx"\n', encoding="utf-8")
    with pytest.raises(ErrorMapeo, match=r"\{mes\} y \{anio\}"):
        Mapeo.cargar(ruta)


def test_un_patron_que_no_termina_en_xlsx_falla_en_espanol(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text('proyectos:\narchivo_salida: "Horas {mes}"\n', encoding="utf-8")
    with pytest.raises(ErrorMapeo, match="terminar en .xlsx"):
        Mapeo.cargar(ruta)
