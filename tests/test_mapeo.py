import re

import pytest
import yaml
from conftest import FIXTURES

from cunix_horas.mapeo import ErrorMapeo, Mapeo


def cargar():
    return Mapeo.cargar(FIXTURES / "mapeo-test.yaml")


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
