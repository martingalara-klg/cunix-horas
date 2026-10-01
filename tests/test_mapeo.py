import re

import pytest
import yaml
from conftest import FIXTURES, RAIZ

from cunix_horas.anexos import Periodo
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


# --- La sección 'anexos:' ---------------------------------------------------
# Todo opcional: lo que no esté se resuelve solo o queda como marcador a la
# vista en el documento. Nunca se inventa una fecha ni un plazo.


def test_los_nombres_de_los_anexos_tienen_valor_por_defecto(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text("proyectos:\n", encoding="utf-8")
    anexos = Mapeo.cargar(ruta).anexos
    assert anexos.nombre_detalle(Periodo(2026, 9)) == (
        "Anexo-II-A-Detalle-horas-KLG-2026-09.xlsx"
    )
    assert anexos.nombre_informe(Periodo(2026, 9)) == (
        "Anexo-II-Informe-mensual-horas-KLG-2026-09.docx"
    )


def test_los_nombres_de_los_anexos_se_pueden_cambiar(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        'proyectos:\nanexos:\n  archivo_detalle: "Detalle {periodo}.xlsx"\n',
        encoding="utf-8",
    )
    assert Mapeo.cargar(ruta).anexos.nombre_detalle(Periodo(2026, 9)) == (
        "Detalle 2026-09.xlsx"
    )


def test_un_patron_con_un_reemplazo_inventado_falla_en_espanol(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        'proyectos:\nanexos:\n  archivo_detalle: "Detalle {dia}.xlsx"\n',
        encoding="utf-8",
    )
    with pytest.raises(ErrorMapeo, match=r"\{periodo\}"):
        Mapeo.cargar(ruta)


def test_un_patron_que_no_termina_en_la_extension_correcta_falla(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        'proyectos:\nanexos:\n  archivo_informe: "Informe {periodo}.xlsx"\n',
        encoding="utf-8",
    )
    with pytest.raises(ErrorMapeo, match="terminar en .docx"):
        Mapeo.cargar(ruta)


def test_sin_configurar_no_hay_plazo_ni_fechas(tmp_path):
    """El informe sale con los marcadores a la vista, que es lo que se busca."""
    ruta = tmp_path / "m.yaml"
    ruta.write_text("proyectos:\n", encoding="utf-8")
    anexos = Mapeo.cargar(ruta).anexos
    assert anexos.dias_habiles_entrega == ""
    assert anexos.contrato_de_fecha == ""
    assert anexos.fecha_de_emision == ""
    assert anexos.perfil_por_defecto == "Desarrollador"


def test_el_plazo_de_entrega_se_puede_escribir_como_numero(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        "proyectos:\nanexos:\n  dias_habiles_entrega: 5\n", encoding="utf-8"
    )
    assert Mapeo.cargar(ruta).anexos.dias_habiles_entrega == "5"


def test_los_perfiles_distintos_del_general_se_declaran_por_persona(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        'proyectos:\nanexos:\n  perfil_por_defecto: "Desarrollador"\n'
        '  perfiles:\n    "Ana Perez": "Líder técnica"\n',
        encoding="utf-8",
    )
    anexos = Mapeo.cargar(ruta).anexos
    assert anexos.perfiles_por_persona == {"Ana Perez": "Líder técnica"}


# --- La sección 'fuentes_manuales:' -----------------------------------------


def test_un_mes_sin_fuentes_manuales_declaradas_carga_igual(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text("proyectos:\n", encoding="utf-8")
    fuentes = Mapeo.cargar(ruta).fuentes_manuales
    assert fuentes.hay_alguna is False
    assert fuentes.horas_sin_kimai == ()
    assert fuentes.descripciones == ()


def test_las_horas_de_quien_no_esta_en_kimai_se_declaran_con_a_quien_restarselas(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        "proyectos:\nfuentes_manuales:\n  horas_sin_kimai:\n"
        '    - persona: "Gabriel Denis"\n'
        '      planilla: "gabriel-denis.xlsx"\n'
        '      restar_a: "Alexis Carnero"\n'
        '      proyecto: "SELICO"\n',
        encoding="utf-8",
    )
    fuentes = Mapeo.cargar(ruta).fuentes_manuales
    assert fuentes.hay_alguna is True
    (entrada,) = fuentes.horas_sin_kimai
    assert entrada.persona == "Gabriel Denis"
    assert entrada.restar_a == "Alexis Carnero"
    assert entrada.proyecto == "SELICO"


def test_una_fuente_de_horas_sin_restar_a_falla_diciendo_que_falta(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        "proyectos:\nfuentes_manuales:\n  horas_sin_kimai:\n"
        '    - persona: "Gabriel Denis"\n'
        '      planilla: "gabriel-denis.xlsx"\n',
        encoding="utf-8",
    )
    with pytest.raises(ErrorMapeo, match="restar_a"):
        Mapeo.cargar(ruta)


def test_las_descripciones_aparte_se_declaran_con_persona_y_planilla(tmp_path):
    ruta = tmp_path / "m.yaml"
    ruta.write_text(
        "proyectos:\nfuentes_manuales:\n  descripciones:\n"
        '    - persona: "Alexis Carnero"\n'
        '      planilla: "alexis-carnero.xlsx"\n',
        encoding="utf-8",
    )
    (entrada,) = Mapeo.cargar(ruta).fuentes_manuales.descripciones
    assert entrada.persona == "Alexis Carnero"
    assert entrada.planilla == "alexis-carnero.xlsx"


def test_el_mapeo_del_repo_declara_las_horas_de_quien_no_tiene_usuario():
    """Regresión: septiembre tiene que salir con el mecanismo ya funcionando."""
    mapeo = Mapeo.cargar(RAIZ / "config" / "mapeo.yaml")
    personas = [e.persona for e in mapeo.fuentes_manuales.horas_sin_kimai]
    assert "Gabriel Denis" in personas
