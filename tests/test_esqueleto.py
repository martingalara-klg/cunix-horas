from conftest import FIXTURES, RAIZ


def test_el_paquete_es_importable():
    import cunix_horas  # noqa: F401


def test_los_fixtures_existen():
    assert (FIXTURES / "kimai-mzalazar.xlsx").is_file()
    assert (FIXTURES / "plantilla.xlsx").is_file()


def test_la_plantilla_de_produccion_existe():
    assert (RAIZ / "templates" / "plantilla.xlsx").is_file()
