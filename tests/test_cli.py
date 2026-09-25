import shutil

import openpyxl
from conftest import FIXTURES

from cunix_horas.cli import procesar_mes


def preparar(tmp_path, nombres=("kimai-mzalazar.xlsx",)):
    (tmp_path / "config").mkdir()
    (tmp_path / "templates").mkdir()
    entrada = tmp_path / "input" / "2026-08"
    entrada.mkdir(parents=True)
    shutil.copy(FIXTURES / "mapeo-test.yaml", tmp_path / "config" / "mapeo.yaml")
    shutil.copy(FIXTURES / "plantilla.xlsx", tmp_path / "templates" / "plantilla.xlsx")
    for nombre in nombres:
        shutil.copy(FIXTURES / "kimai-mzalazar.xlsx", entrada / nombre)
    return tmp_path


def test_genera_el_excel_del_mes(tmp_path):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    salida = raiz / "output" / "2026-08" / "Aug Zalazar.xlsx"
    assert salida.is_file()
    hoja = openpyxl.load_workbook(salida).active
    assert hoja["A1"].value == "Matias Zalazar"
    assert hoja["B1"].value == "Total"
    assert hoja.cell(row=hoja.max_row, column=1).value == "Total"


def test_los_totales_del_excel_cierran_en_76_5(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    hoja = openpyxl.load_workbook(
        raiz / "output" / "2026-08" / "Aug Zalazar.xlsx"
    ).active
    assert hoja.cell(row=hoja.max_row, column=2).value == 76.5


def test_escribe_el_archivo_de_validacion(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    validacion = raiz / "output" / "2026-08" / "_validacion.txt"
    assert validacion.is_file()
    assert "Día hábil sin carga" in validacion.read_text(encoding="utf-8")


def test_un_archivo_roto_no_frena_a_los_demas(tmp_path, capsys):
    raiz = preparar(tmp_path)
    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")
    codigo = procesar_mes("2026-08", raiz)
    assert (raiz / "output" / "2026-08" / "Aug Zalazar.xlsx").is_file()
    assert codigo == 1
    assert "roto.xlsx" in capsys.readouterr().out


def test_falla_si_no_existe_la_carpeta_del_mes(tmp_path, capsys):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-09", raiz) == 1
    assert "input/2026-09" in capsys.readouterr().out


def test_falla_con_un_mes_mal_escrito(tmp_path, capsys):
    raiz = preparar(tmp_path)
    assert procesar_mes("agosto", raiz) == 1
    assert "AAAA-MM" in capsys.readouterr().out


def test_un_proyecto_sin_mapear_frena_solo_a_ese_dev(tmp_path, capsys):
    raiz = preparar(tmp_path)
    (raiz / "config" / "mapeo.yaml").write_text(
        "personas:\n"
        "  mzalazar:\n"
        '    nombre: "Matias Zalazar"\n'
        '    archivo: "Zalazar"\n'
        "proyectos:\n"
        "  CO2610170:\n"
        '    cliente: "Aduanas"\n'
        '    proyecto: "Subastas"\n',
        encoding="utf-8",
    )
    assert procesar_mes("2026-08", raiz) == 1
    salida = capsys.readouterr().out
    assert "CO2510115" in salida
    assert "config/mapeo.yaml" in salida
    assert not (raiz / "output" / "2026-08" / "Aug Zalazar.xlsx").exists()
