import pytest
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
