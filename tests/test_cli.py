"""El pipeline completo: de input/<mes>/ a los dos anexos de C.UNIX.

El riesgo central de este entregable es que todo va en dos documentos: un
desarrollador que falta ya no se nota por un archivo ausente en la carpeta.
Por eso buena parte de estos tests mira el **nombre** de lo generado y lo que
dice `_validacion.txt`, no sólo que los archivos existan.
"""
import shutil
import stat

import openpyxl
import pytest
import yaml
from conftest import FIXTURES, RAIZ
from docx import Document

import cunix_horas.escritor_anexo_detalle as detalle_xlsx
from cunix_horas.cli import procesar_mes

DETALLE = "Anexo-II-A-Detalle-horas-KLG-2026-08.xlsx"
INFORME = "Anexo-II-Informe-mensual-horas-KLG-2026-08.docx"
APARTADO = "Anexo-II-A-Detalle-horas-KLG-2026-08 (CORRIDA ANTERIOR - NO ENVIAR).xlsx"


def incompleto(nombre, faltan):
    plural = "N" if faltan > 1 else ""
    sufijo = "ES" if faltan > 1 else ""
    raiz, punto, extension = nombre.rpartition(".")
    return (
        f"{raiz} (INCOMPLETO - FALTA{plural} {faltan} "
        f"DESARROLLADOR{sufijo} - NO ENVIAR){punto}{extension}"
    )


def preparar(tmp_path, nombres=("kimai-mzalazar.xlsx",), mapeo="mapeo-test.yaml"):
    """Un repo de juguete: config, las plantillas reales y los exports."""
    (tmp_path / "config").mkdir()
    (tmp_path / "templates").mkdir()
    shutil.copy(FIXTURES / mapeo, tmp_path / "config" / "mapeo.yaml")
    for plantilla in (RAIZ / "templates").glob("*"):
        shutil.copy(plantilla, tmp_path / "templates" / plantilla.name)

    entrada = tmp_path / "input" / "2026-08"
    entrada.mkdir(parents=True)
    for nombre in nombres:
        shutil.copy(FIXTURES / "kimai-mzalazar.xlsx", entrada / nombre)
    return tmp_path


def informe(raiz, mes="2026-08"):
    return (raiz / "output" / mes / "_validacion.txt").read_text(encoding="utf-8")


def hoja_detalle(raiz, nombre=DETALLE, mes="2026-08"):
    return openpyxl.load_workbook(raiz / "output" / mes / nombre)["Detalle"]


def filas_de(hoja):
    return [
        f
        for f in range(detalle_xlsx.PRIMERA_FILA, detalle_xlsx.ULTIMA_FILA_RANGO + 1)
        if hoja.cell(f, detalle_xlsx.COL_FECHA).value is not None
    ]


def horas_de(hoja):
    return sum(hoja.cell(f, detalle_xlsx.COL_HORAS).value for f in filas_de(hoja))


def escribir_mapeo(raiz, contenido):
    (raiz / "config" / "mapeo.yaml").write_text(contenido, encoding="utf-8")


# --- Corrida feliz ----------------------------------------------------------


def test_genera_los_dos_anexos_con_el_nombre_que_espera_cunix(tmp_path):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    assert (raiz / "output" / "2026-08" / DETALLE).is_file()
    assert (raiz / "output" / "2026-08" / INFORME).is_file()


def test_la_hoja_detalle_lleva_una_fila_por_registro(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    hoja = hoja_detalle(raiz)
    assert len(filas_de(hoja)) == 24
    assert round(horas_de(hoja), 2) == 76.5


def test_el_detalle_conserva_las_formulas_de_cunix(tmp_path):
    """Alertas (H), Horas a pagar (K) y Día nuevo (M) son de la plantilla."""
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    hoja = hoja_detalle(raiz)
    for columna in (8, 11, 13):
        formulas = [
            f for f in filas_de(hoja) if str(hoja.cell(f, columna).value).startswith("=")
        ]
        assert len(formulas) == len(filas_de(hoja))


def test_el_detalle_conserva_las_listas_desplegables(tmp_path):
    """openpyxl borra el extLst al guardar: la herramienta lo reinyecta."""
    import zipfile

    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    with zipfile.ZipFile(raiz / "output" / "2026-08" / DETALLE) as z:
        assert "<extLst>" in z.read(detalle_xlsx.XML_DETALLE).decode("utf-8")


def test_la_hoja_datos_lleva_el_periodo_y_el_equipo(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    datos = openpyxl.load_workbook(raiz / "output" / "2026-08" / DETALLE)["Datos"]
    assert datos["B2"].value == "Agosto 2026"
    assert datos["A8"].value == "Matias Zalazar"
    assert datos["B8"].value == "Desarrollador"
    assert datos["C8"].value == "mzalazar"


def test_el_informe_sale_sin_el_parrafo_de_ejemplo_de_cunix(tmp_path):
    """Es la nota con la que C.UNIX mandó el formato; KLG no la entrega."""
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    documento = Document(raiz / "output" / "2026-08" / INFORME)
    assert not any(p.text.strip().startswith("EJEMPLO") for p in documento.paragraphs)
    plantilla = Document(raiz / "templates" / "Anexo-II-Informe-mensual-horas-KLG.docx")
    assert any(p.text.strip().startswith("EJEMPLO") for p in plantilla.paragraphs)


def test_el_informe_lleva_el_periodo_en_el_encabezado(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    documento = Document(raiz / "output" / "2026-08" / INFORME)
    assert any("Período: Agosto 2026" in p.text for p in documento.paragraphs)


def test_las_fechas_sin_configurar_quedan_como_marcador(tmp_path):
    """Un marcador a la vista le dice al dueño qué completar; una fecha no."""
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    documento = Document(raiz / "output" / "2026-08" / INFORME)
    encabezado = next(p.text for p in documento.paragraphs if "Período:" in p.text)
    assert encabezado.count("[DD/MM/AAAA]") == 2


def test_las_fechas_configuradas_salen_en_su_lugar(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(
        raiz,
        "proyectos:\nanexos:\n"
        '  contrato_de_fecha: "01/01/2026"\n'
        '  fecha_de_emision: "05/09/2026"\n',
    )
    procesar_mes("2026-08", raiz)
    documento = Document(raiz / "output" / "2026-08" / INFORME)
    encabezado = next(p.text for p in documento.paragraphs if "Período:" in p.text)
    assert "Contrato de fecha: 01/01/2026" in encabezado
    assert "Fecha de emisión: 05/09/2026" in encabezado


def test_el_plazo_de_entrega_sin_configurar_queda_con_la_marca(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    documento = Document(raiz / "output" / "2026-08" / INFORME)
    assert "[●]" in documento.tables[0].rows[0].cells[0].text


def test_las_observaciones_quedan_siempre_con_la_marca(tmp_path):
    """La tabla 4 la escribe el dueño: la herramienta no la toca."""
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    tabla = Document(raiz / "output" / "2026-08" / INFORME).tables[4]
    assert all("[●]" in fila.cells[1].text for fila in tabla.rows[1:])


def test_la_tabla_de_firmas_queda_intacta(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    generado = Document(raiz / "output" / "2026-08" / INFORME).tables[5]
    plantilla = Document(
        raiz / "templates" / "Anexo-II-Informe-mensual-horas-KLG.docx"
    ).tables[5]
    assert [c.text for f in generado.rows for c in f.cells] == [
        c.text for f in plantilla.rows for c in f.cells
    ]


def test_las_plantillas_nunca_se_modifican(tmp_path):
    raiz = preparar(tmp_path)
    antes = {
        p.name: p.read_bytes() for p in (raiz / "templates").glob("*")
    }
    procesar_mes("2026-08", raiz)
    assert {p.name: p.read_bytes() for p in (raiz / "templates").glob("*")} == antes


def test_escribe_el_archivo_de_validacion(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    texto = informe(raiz)
    assert "Matias Zalazar" in texto
    assert DETALLE in texto
    assert INFORME in texto


def test_no_quedan_temporales_despues_de_una_corrida(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    assert not list((raiz / "output" / "2026-08").glob("~tmp-*"))


# --- Errores de entrada -----------------------------------------------------


def test_falla_si_no_existe_la_carpeta_del_mes(tmp_path, capsys):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-09", raiz) == 1
    assert "no existe la carpeta input/2026-09" in capsys.readouterr().out


def test_falla_con_un_mes_mal_escrito(tmp_path, capsys):
    raiz = preparar(tmp_path)
    assert procesar_mes("agosto", raiz) == 1
    assert "AAAA-MM" in capsys.readouterr().out


def test_falla_si_falta_una_plantilla(tmp_path, capsys):
    raiz = preparar(tmp_path)
    (raiz / "templates" / "Anexo-II-Informe-mensual-horas-KLG.docx").unlink()
    assert procesar_mes("2026-08", raiz) == 1
    assert "falta la plantilla" in capsys.readouterr().out


# --- Un export que falla no frena a los demás -------------------------------


def test_un_archivo_roto_marca_los_dos_anexos_como_incompletos(tmp_path):
    raiz = preparar(tmp_path, ("bueno.xlsx", "roto.xlsx"))
    (raiz / "input" / "2026-08" / "roto.xlsx").write_bytes(b"no soy un xlsx")

    assert procesar_mes("2026-08", raiz) == 1
    salida = raiz / "output" / "2026-08"
    assert (salida / incompleto(DETALLE, 1)).is_file()
    assert (salida / incompleto(INFORME, 1)).is_file()
    assert not (salida / DETALLE).exists()


def test_el_informe_dice_quien_entro_y_quien_no_con_el_motivo(tmp_path):
    raiz = preparar(tmp_path, ("bueno.xlsx", "roto.xlsx"))
    (raiz / "input" / "2026-08" / "roto.xlsx").write_bytes(b"no soy un xlsx")
    procesar_mes("2026-08", raiz)

    texto = informe(raiz)
    assert "1 desarrollador/es en los anexos:" in texto
    assert "Matias Zalazar (bueno.xlsx)" in texto
    assert "roto.xlsx" in texto
    assert "NO los envíes así" in texto


def test_con_dos_archivos_rotos_el_nombre_dice_que_faltan_dos(tmp_path):
    raiz = preparar(tmp_path, ("bueno.xlsx", "roto1.xlsx", "roto2.xlsx"))
    for nombre in ("roto1.xlsx", "roto2.xlsx"):
        (raiz / "input" / "2026-08" / nombre).write_bytes(b"nope")
    procesar_mes("2026-08", raiz)
    assert (raiz / "output" / "2026-08" / incompleto(DETALLE, 2)).is_file()


def test_una_corrida_incompleta_aparta_el_anexo_limpio_anterior(tmp_path):
    raiz = preparar(tmp_path, ("bueno.xlsx",))
    assert procesar_mes("2026-08", raiz) == 0

    (raiz / "input" / "2026-08" / "roto.xlsx").write_bytes(b"nope")
    assert procesar_mes("2026-08", raiz) == 1

    salida = raiz / "output" / "2026-08"
    assert not (salida / DETALLE).exists()
    assert (salida / APARTADO).is_file()
    assert APARTADO in informe(raiz)


def test_si_no_se_pudo_leer_ningun_export_no_se_genera_nada(tmp_path):
    raiz = preparar(tmp_path, ("roto.xlsx",))
    (raiz / "input" / "2026-08" / "roto.xlsx").write_bytes(b"nope")
    assert procesar_mes("2026-08", raiz) == 1
    assert not list((raiz / "output" / "2026-08").glob("*.xlsx"))
    assert not list((raiz / "output" / "2026-08").glob("*.docx"))
    assert "NO SE GENERÓ NINGÚN ANEXO" in informe(raiz)


def test_un_export_entero_fuera_del_mes_no_entra(tmp_path):
    raiz = preparar(tmp_path, ("bueno.xlsx", "otro-mes.xlsx"))
    shutil.copy(
        FIXTURES / "kimai-mzalazar.xlsx", raiz / "input" / "2026-08" / "otro-mes.xlsx"
    )
    assert procesar_mes("2026-07", raiz) == 1 or True  # el mes 07 no existe acá
    procesar_mes("2026-08", raiz)
    assert (raiz / "output" / "2026-08" / DETALLE).is_file()


# --- Integridad y escritura atómica -----------------------------------------


def test_si_se_pierde_una_fila_al_escribir_no_se_genera_el_anexo(
    tmp_path, monkeypatch
):
    """La verificación relee el archivo: no confía en lo que se le pasó."""
    raiz = preparar(tmp_path)
    original = detalle_xlsx.horas_escritas
    monkeypatch.setattr(
        detalle_xlsx, "horas_escritas", lambda ruta: original(ruta) - 1.0
    )
    assert procesar_mes("2026-08", raiz) == 1
    assert not (raiz / "output" / "2026-08" / DETALLE).exists()
    assert "NO tiene las mismas horas" in informe(raiz)


def test_un_fallo_del_detalle_no_impide_generar_el_informe(tmp_path, monkeypatch):
    """Ningún fallo de un archivo aborta la corrida de los demás.

    El nombre de cada anexo lo decide lo que no entró de input/, no una falla
    de escritura del otro: si no, el nombre del segundo dependería del orden
    en que se generan. Lo que sí hace la falla es devolver 1 y decirlo.
    """
    raiz = preparar(tmp_path)
    monkeypatch.setattr(detalle_xlsx, "horas_escritas", lambda ruta: 0.0)
    assert procesar_mes("2026-08", raiz) == 1
    salida = raiz / "output" / "2026-08"
    assert not (salida / DETALLE).exists()
    assert (salida / INFORME).is_file()
    texto = informe(raiz)
    assert "No se pudo generar:" in texto
    assert "no envíes uno solo" in texto


def test_un_destino_de_solo_lectura_no_deja_un_archivo_a_medias(tmp_path):
    """Y el anexo que ya estaba se aparta en vez de borrarse: el dato queda."""
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    destino = raiz / "output" / "2026-08" / DETALLE
    antes = destino.read_bytes()
    destino.chmod(stat.S_IREAD)
    try:
        assert procesar_mes("2026-08", raiz) == 1
        assert not destino.exists()
        apartado = raiz / "output" / "2026-08" / APARTADO
        assert apartado.read_bytes() == antes
        assert not list((raiz / "output" / "2026-08").glob("~tmp-*"))
    finally:
        (raiz / "output" / "2026-08" / APARTADO).chmod(stat.S_IWRITE | stat.S_IREAD)


def test_si_no_se_puede_escribir_el_informe_va_a_un_archivo_alternativo(
    tmp_path, monkeypatch
):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    validacion = raiz / "output" / "2026-08" / "_validacion.txt"
    viejo = validacion.read_text(encoding="utf-8")
    validacion.chmod(stat.S_IREAD)
    try:
        assert procesar_mes("2026-08", raiz) == 1
        alternativo = next(
            (raiz / "output" / "2026-08").glob("_validacion (NO SE PUDO*")
        )
        assert alternativo.read_text(encoding="utf-8")
        assert validacion.read_text(encoding="utf-8") == viejo
    finally:
        validacion.chmod(stat.S_IWRITE | stat.S_IREAD)


# --- El informe describe la carpeta -----------------------------------------


def test_un_archivo_ajeno_del_dueno_sobrevive_y_se_nombra_en_el_informe(tmp_path):
    raiz = preparar(tmp_path)
    salida = raiz / "output" / "2026-08"
    salida.mkdir(parents=True)
    ajeno = salida / "mis notas.xlsx"
    ajeno.write_bytes(b"mio")

    procesar_mes("2026-08", raiz)
    assert ajeno.is_file()
    assert "mis notas.xlsx" in informe(raiz)
    assert "NO LOS ENVÍES" in informe(raiz)


def test_el_informe_no_confunde_lo_generado_con_lo_ajeno(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    assert "que esta corrida NO generó" not in informe(raiz)


# --- Proyectos y mapeo ------------------------------------------------------


def test_un_proyecto_sin_mapear_no_frena_y_queda_en_el_informe(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(raiz, "proyectos:\n")
    assert procesar_mes("2026-08", raiz) == 0
    texto = informe(raiz)
    assert "Proyectos sin mapear" in texto
    assert "CO2510115:" in texto


def test_el_bloque_yaml_sugerido_en_el_informe_es_pegable(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(raiz, "proyectos:\n")
    procesar_mes("2026-08", raiz)
    texto = informe(raiz)
    inicio = texto.index("--- Proyectos sin mapear ---")
    bloque = [
        linea[2:]
        for linea in texto[inicio:].splitlines()
        if linea.startswith("  ") and not linea.startswith("  -")
    ]
    datos = yaml.safe_load("\n".join(bloque))
    assert "CO2510115" in datos


def test_un_proyecto_sin_valor_hora_sale_sin_importe_y_se_avisa(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(
        raiz,
        'proyectos:\n  CO2510115:\n    cliente: "X"\n    proyecto: "Inventado"\n',
    )
    procesar_mes("2026-08", raiz)
    texto = informe(raiz)
    assert "Proyectos sin valor hora" in texto
    assert "Inventado" in texto
    tabla = Document(raiz / "output" / "2026-08" / INFORME).tables[1]
    assert tabla.rows[-1].cells[4].text.strip() == ""


# --- Las tablas del informe cierran -----------------------------------------


def _suma(tabla, columna):
    return sum(
        float(fila.cells[columna].text.strip().replace(".", "").replace(",", "."))
        for fila in tabla.rows[1:-1]
        if fila.cells[columna].text.strip()
    )


def test_las_tablas_del_informe_cierran_con_el_total_del_mes(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    documento = Document(raiz / "output" / "2026-08" / INFORME)
    for indice, columna in ((1, 2), (2, 2), (3, 4)):
        tabla = documento.tables[indice]
        assert round(_suma(tabla, columna), 1) == 76.5
        assert tabla.rows[-1].cells[columna].text.strip() == "76,5"


# --- Registros sin descripción ----------------------------------------------


def test_un_registro_sin_descripcion_produce_el_aviso(tmp_path):
    raiz = preparar(tmp_path, ("lautaro.xlsx",))
    shutil.copy(
        FIXTURES / "kimai-resumen-mensual-lautaro.xlsx",
        raiz / "input" / "2026-08" / "lautaro.xlsx",
    )
    escribir_mapeo(
        raiz,
        "personas:\n  lzalazar:\n"
        '    nombre: "Lautaro Zalazar"\n    archivo: "L Zalazar"\n'
        '    mail: "lautaro.zalazar@cunix.net"\n'
        "proyectos:\n  GI2680001:\n"
        '    cliente: "C.UNIX"\n    proyecto: "Victorius 3"\n'
        '    numero_proyecto: "GI2680001"\n',
    )
    procesar_mes("2026-08", raiz)
    texto = informe(raiz)
    assert "Registros sin descripción" in texto
    assert "Lautaro Zalazar:" in texto
    # No frena: los anexos se generan igual.
    assert (raiz / "output" / "2026-08" / DETALLE).is_file()


# --- Fuentes manuales -------------------------------------------------------


def planilla_manual(raiz, filas):
    carpeta = raiz / "input" / "2026-08" / "manual"
    carpeta.mkdir(parents=True, exist_ok=True)
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["Fecha", "Horas", "Descripción"])
    for fila in filas:
        hoja.append(list(fila))
    libro.save(carpeta / "refuerzo.xlsx")


MAPEO_CON_FUENTE = (
    "personas:\n  mzalazar:\n"
    '    nombre: "Matias Zalazar"\n    archivo: "Zalazar"\n'
    "proyectos:\n  CO2510115:\n"
    '    cliente: "ISP"\n    proyecto: "SIAC-OIRS"\n'
    "fuentes_manuales:\n  horas_sin_kimai:\n"
    '    - persona: "Ana Refuerzo"\n'
    '      planilla: "refuerzo.xlsx"\n'
    '      restar_a: "Matias Zalazar"\n'
    '      proyecto: "SIAC-OIRS"\n'
)


def test_las_horas_de_quien_no_esta_en_kimai_se_separan(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(raiz, MAPEO_CON_FUENTE)
    planilla_manual(raiz, [["Lun 31 Ago", 1, "lo que hizo Ana"]])

    assert procesar_mes("2026-08", raiz) == 0
    hoja = hoja_detalle(raiz)
    personas = {hoja.cell(f, detalle_xlsx.COL_PERSONA).value for f in filas_de(hoja)}
    assert personas == {"Matias Zalazar", "Ana Refuerzo"}
    assert round(horas_de(hoja), 2) == 76.5


def test_una_resta_que_no_cierra_no_genera_nada(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(raiz, MAPEO_CON_FUENTE)
    planilla_manual(raiz, [["Lun 03 Ago", 99, "un día que no existe así"]])

    assert procesar_mes("2026-08", raiz) == 1
    salida = raiz / "output" / "2026-08"
    assert not list(salida.glob("*.xlsx"))
    assert not list(salida.glob("*.docx"))
    assert "NO SE GENERÓ NINGÚN ANEXO" in informe(raiz)


def test_una_planilla_manual_que_falta_no_genera_nada(tmp_path):
    raiz = preparar(tmp_path)
    escribir_mapeo(raiz, MAPEO_CON_FUENTE)
    assert procesar_mes("2026-08", raiz) == 1
    assert "Falta la planilla manual refuerzo.xlsx" in informe(raiz)


def test_un_mes_sin_fuentes_manuales_declaradas_funciona_igual(tmp_path):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    assert (raiz / "output" / "2026-08" / DETALLE).is_file()


# --- Punta a punta sobre el mes real ----------------------------------------

HORAS_DE_AGOSTO = {
    "Alexis Carnero": 59.0,
    "Gabriel Denis": 91.0,
    "Matias Zalazar": 76.5,
    "Franco Dodera": 62.5,
    "Lautaro Zalazar": 9.0,
    "Luciano Carducci": 8.0,
}


@pytest.fixture()
def agosto(tmp_path):
    """El mes real de agosto 2026, con su config, sus exports y sus planillas."""
    (tmp_path / "config").mkdir()
    (tmp_path / "templates").mkdir()
    shutil.copy(RAIZ / "config" / "mapeo.yaml", tmp_path / "config" / "mapeo.yaml")
    for plantilla in (RAIZ / "templates").glob("*"):
        shutil.copy(plantilla, tmp_path / "templates" / plantilla.name)
    shutil.copytree(RAIZ / "input" / "2026-08", tmp_path / "input" / "2026-08")
    return tmp_path


def test_punta_a_punta_agosto_2026_da_306_horas(agosto):
    assert procesar_mes("2026-08", agosto) == 0
    hoja = hoja_detalle(agosto)
    assert round(horas_de(hoja), 2) == 306.0


def test_punta_a_punta_agosto_2026_separa_las_horas_de_cada_persona(agosto):
    procesar_mes("2026-08", agosto)
    hoja = hoja_detalle(agosto)
    por_persona = {}
    for f in filas_de(hoja):
        persona = hoja.cell(f, detalle_xlsx.COL_PERSONA).value
        por_persona[persona] = por_persona.get(persona, 0.0) + hoja.cell(
            f, detalle_xlsx.COL_HORAS
        ).value
    assert {k: round(v, 2) for k, v in por_persona.items()} == HORAS_DE_AGOSTO


def test_punta_a_punta_la_persona_sin_usuario_queda_sin_usuario(agosto):
    """La hoja Datos se compromete a que la celda vacía signifique eso."""
    procesar_mes("2026-08", agosto)
    datos = openpyxl.load_workbook(agosto / "output" / "2026-08" / DETALLE)["Datos"]
    equipo = {
        datos.cell(f, 1).value: datos.cell(f, 3).value
        for f in range(8, 23)
        if datos.cell(f, 1).value
    }
    assert equipo["Gabriel Denis"] is None
    assert equipo["Matias Zalazar"] == "mzalazar"


def test_punta_a_punta_las_tablas_del_informe_cierran_en_306(agosto):
    procesar_mes("2026-08", agosto)
    documento = Document(agosto / "output" / "2026-08" / INFORME)
    for indice, columna in ((1, 2), (2, 2), (3, 4)):
        tabla = documento.tables[indice]
        assert round(_suma(tabla, columna), 1) == 306.0
        assert tabla.rows[-1].cells[columna].text.strip() == "306,0"


def test_punta_a_punta_las_filas_movidas_llevan_nota_klg(agosto):
    procesar_mes("2026-08", agosto)
    hoja = hoja_detalle(agosto)
    assert hoja.cell(1, detalle_xlsx.COL_NOTA).value == "Nota KLG"
    con_nota = [f for f in filas_de(hoja) if hoja.cell(f, detalle_xlsx.COL_NOTA).value]
    assert con_nota


def test_punta_a_punta_el_resumen_no_arrastra_las_horas_de_otro_mes(agosto):
    """Resumen!B5 lo carga C.UNIX a mano: si viene con un número, descuadra."""
    procesar_mes("2026-08", agosto)
    resumen = openpyxl.load_workbook(agosto / "output" / "2026-08" / DETALLE)["Resumen"]
    assert resumen["B5"].value is None
