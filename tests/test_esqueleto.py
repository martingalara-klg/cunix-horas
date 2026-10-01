from conftest import FIXTURES, RAIZ


def test_el_paquete_es_importable():
    import cunix_horas  # noqa: F401


def test_los_fixtures_existen():
    assert (FIXTURES / "kimai-mzalazar.xlsx").is_file()


def test_las_plantillas_de_los_anexos_existen():
    """Los dos anexos que KLG entrega todos los meses salen de acá."""
    plantillas = RAIZ / "templates"
    assert (plantillas / "Anexo-II-A-Detalle-horas-KLG.xlsx").is_file()
    assert (plantillas / "Anexo-II-Informe-mensual-horas-KLG.docx").is_file()
