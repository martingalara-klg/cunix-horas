"""El pipeline completo: de input/<mes>/ al único archivo del partner.

El riesgo central de este entregable es que todo va en un solo archivo: un
desarrollador que falta ya no se nota por un Excel ausente en la carpeta. Por
eso buena parte de estos tests mira el **nombre** del archivo generado y lo
que dice `_validacion.txt`, no sólo que el archivo exista.
"""
import shutil
import stat
from pathlib import Path

import openpyxl
import yaml
from conftest import FIXTURES, RAIZ

import cunix_horas.cli as cli
from cunix_horas.cli import procesar_mes
from cunix_horas.kimai_comun import leer_hoja
from cunix_horas.lector_kimai import leer, leer_resumen_mensual

LIMPIO = "Horas KLG-Aug2026.xlsx"
APARTADO = "Horas KLG-Aug2026 (CORRIDA ANTERIOR - NO ENVIAR).xlsx"


def incompleto(faltan):
    plural = "N" if faltan > 1 else ""
    sufijo = "ES" if faltan > 1 else ""
    return (
        f"Horas KLG-Aug2026 (INCOMPLETO - FALTA{plural} {faltan} "
        f"DESARROLLADOR{sufijo} - NO ENVIAR).xlsx"
    )


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


def informe(raiz, mes="2026-08"):
    return (raiz / "output" / mes / "_validacion.txt").read_text(encoding="utf-8")


def hoja_de(raiz, nombre, mes="2026-08"):
    return openpyxl.load_workbook(raiz / "output" / mes / nombre).active


def horas_de(hoja):
    return (
        sum(
            hoja.cell(row=f, column=2).value.total_seconds()
            for f in range(2, hoja.max_row + 1)
        )
        / 3600
    )


# --- Corrida feliz ----------------------------------------------------------


def test_genera_el_archivo_del_mes_con_el_nombre_del_partner(tmp_path):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    hoja = hoja_de(raiz, LIMPIO)
    assert hoja["A1"].value == "Date"
    assert hoja.max_row == 25  # 24 registros + encabezado
    assert horas_de(hoja) == 76.5


def test_escribe_el_archivo_de_validacion(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    texto = informe(raiz)
    assert texto.startswith("Corrida de output/2026-08/")
    assert "1 desarrollador/es en el archivo:" in texto
    assert "Matias Zalazar (kimai-mzalazar.xlsx)" in texto
    assert f"Archivo para el partner: {LIMPIO}" in texto
    assert "Día hábil sin carga" in texto


def test_no_quedan_temporales_despues_de_una_corrida(tmp_path):
    raiz = preparar(tmp_path)
    procesar_mes("2026-08", raiz)
    assert [p.name for p in (raiz / "output" / "2026-08").glob("~tmp-*")] == []


def test_falla_si_no_existe_la_carpeta_del_mes(tmp_path, capsys):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-09", raiz) == 1
    assert "input/2026-09" in capsys.readouterr().out


def test_falla_con_un_mes_mal_escrito(tmp_path, capsys):
    raiz = preparar(tmp_path)
    assert procesar_mes("agosto", raiz) == 1
    assert "AAAA-MM" in capsys.readouterr().out


# --- El riesgo central: un desarrollador que falta ---------------------------


def test_un_archivo_roto_marca_el_consolidado_como_incompleto(tmp_path, capsys):
    """Los demás entran igual, pero el nombre avisa que falta alguien."""
    raiz = preparar(tmp_path)
    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")

    assert procesar_mes("2026-08", raiz) == 1
    salida = raiz / "output" / "2026-08"
    assert not (salida / LIMPIO).exists()
    assert (salida / incompleto(1)).is_file()

    # El desarrollador que sí se leyó está completo.
    assert horas_de(hoja_de(raiz, incompleto(1))) == 76.5

    texto = informe(raiz)
    assert "roto.xlsx" in texto
    assert "NO entraron" in texto
    assert "NO lo envíes" in texto
    assert "roto.xlsx" in capsys.readouterr().out


def test_el_informe_dice_quien_entro_y_quien_no_con_el_motivo(tmp_path):
    raiz = preparar(tmp_path)
    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")
    procesar_mes("2026-08", raiz)

    texto = informe(raiz)
    posicion_devs = texto.index("desarrollador/es en el archivo")
    posicion_archivo = texto.index("Archivo para el partner")
    # Quién entró y quién no va ARRIBA del nombre del archivo generado.
    assert posicion_devs < texto.index("NO entraron") < posicion_archivo
    assert "no es un archivo .xlsx válido" in texto


def test_con_dos_archivos_rotos_el_nombre_dice_que_faltan_dos(tmp_path):
    raiz = preparar(tmp_path)
    for nombre in ("roto-a.xlsx", "roto-b.xlsx"):
        (raiz / "input" / "2026-08" / nombre).write_text("basura", encoding="utf-8")
    assert procesar_mes("2026-08", raiz) == 1
    assert (raiz / "output" / "2026-08" / incompleto(2)).is_file()


def test_una_corrida_incompleta_aparta_el_archivo_limpio_anterior(tmp_path):
    """El limpio del mes pasado no puede quedar ahí para enviarse por error."""
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    salida = raiz / "output" / "2026-08"
    assert (salida / LIMPIO).is_file()

    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")
    assert procesar_mes("2026-08", raiz) == 1

    assert not (salida / LIMPIO).exists()
    assert (salida / APARTADO).is_file()
    assert (salida / incompleto(1)).is_file()
    assert "CORRIDA ANTERIOR - NO ENVIAR" in informe(raiz)


def test_si_no_se_pudo_leer_ningun_export_no_se_genera_nada(tmp_path):
    raiz = preparar(tmp_path, nombres=())
    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")
    assert procesar_mes("2026-08", raiz) == 1
    assert list((raiz / "output" / "2026-08").glob("*.xlsx")) == []
    assert "NO SE GENERÓ NINGÚN ARCHIVO" in informe(raiz)


def test_un_export_entero_fuera_del_mes_no_entra_al_archivo(tmp_path):
    """Mismo disparador que el export vacío: el rango de fechas mal en Kimai."""
    raiz = preparar(tmp_path)
    septiembre = raiz / "input" / "2026-09"
    septiembre.mkdir(parents=True)
    shutil.copy(FIXTURES / "kimai-mzalazar.xlsx", septiembre / "kimai-mzalazar.xlsx")

    assert procesar_mes("2026-09", raiz) == 1
    assert list((raiz / "output" / "2026-09").glob("*.xlsx")) == []

    texto = informe(raiz, "2026-09")
    assert "kimai-mzalazar.xlsx" in texto
    assert "se descartaron por fecha" in texto
    assert "/8/2026" in texto
    assert "otro rango de fechas" in texto


def test_un_archivo_de_salida_abierto_queda_en_el_informe(tmp_path, monkeypatch):
    raiz = preparar(tmp_path)

    def replace_bloqueado(origen, destino):
        raise PermissionError(13, "Acceso denegado")

    monkeypatch.setattr(cli.os, "replace", replace_bloqueado)
    assert procesar_mes("2026-08", raiz) == 1
    texto = informe(raiz)
    assert "permiso denegado" in texto
    assert "abierto" in texto


# --- Integridad --------------------------------------------------------------


def test_si_se_pierde_una_fila_al_escribir_no_se_genera_el_archivo(
    tmp_path, monkeypatch
):
    """Fallo ruidoso, no aviso: el archivo se vería completo sin serlo."""
    raiz = preparar(tmp_path)
    escribir_real = cli.escribir_detalle

    def escribir_perdiendo_una_fila(detalle, destino):
        recortado = type(detalle)(
            filas=detalle.filas[:-1],
            sin_mapear=detalle.sin_mapear,
            numeros_ambiguos=detalle.numeros_ambiguos,
        )
        return escribir_real(recortado, destino)

    monkeypatch.setattr(cli, "escribir_detalle", escribir_perdiendo_una_fila)
    assert procesar_mes("2026-08", raiz) == 1
    assert list((raiz / "output" / "2026-08").glob("*.xlsx")) == []

    texto = informe(raiz)
    assert "NO tiene las mismas horas" in texto
    assert "No se generó nada" in texto


# --- Proyectos sin mapear y Project number ambiguo ---------------------------


def test_un_proyecto_sin_mapear_ya_no_frena_y_queda_en_el_informe(tmp_path):
    raiz = preparar(tmp_path)
    (raiz / "config" / "mapeo.yaml").write_text(
        'proyectos:\n  CO2610170:\n    cliente: "Aduanas"\n'
        '    proyecto: "Subastas"\n',
        encoding="utf-8",
    )
    assert procesar_mes("2026-08", raiz) == 0

    hoja = hoja_de(raiz, LIMPIO)
    proyectos = {
        hoja.cell(row=f, column=7).value for f in range(2, hoja.max_row + 1)
    }
    clientes = {hoja.cell(row=f, column=6).value for f in range(2, hoja.max_row + 1)}
    # El mapeado sale con el nombre del mapeo; el otro, con el de Kimai.
    assert "Subastas" in proyectos
    assert "ISPCH-SopEvo-SIAC" in proyectos
    assert "Instituto de Salud Publica de Chile" in clientes

    texto = informe(raiz)
    assert "--- Proyectos sin mapear ---" in texto
    assert "CO2510115:" in texto


def test_el_bloque_yaml_sugerido_en_el_informe_es_pegable(tmp_path):
    """Lo que el informe ofrece pegar en mapeo.yaml tiene que ser YAML válido."""
    raiz = preparar(tmp_path)
    (raiz / "config" / "mapeo.yaml").write_text("proyectos:\n", encoding="utf-8")
    assert procesar_mes("2026-08", raiz) == 0

    lineas = informe(raiz).splitlines()
    inicio = next(i for i, linea in enumerate(lineas) if "CO2510115:" in linea)
    bloque = "\n".join(lineas[inicio : inicio + 3])

    ruta = raiz / "config" / "mapeo.yaml"
    ruta.write_text(ruta.read_text(encoding="utf-8") + bloque + "\n", encoding="utf-8")
    datos = yaml.safe_load(ruta.read_text(encoding="utf-8"))
    assert "cliente" in datos["proyectos"]["CO2510115"]
    assert "proyecto" in datos["proyectos"]["CO2510115"]


def test_un_project_number_con_dos_nombres_se_avisa_en_el_informe(tmp_path):
    """Alguien renombró el proyecto en Kimai a mitad de mes."""
    from test_lector_kimai import _fila, _xlsx_de_kimai

    raiz = preparar(tmp_path, nombres=())
    (raiz / "config" / "mapeo.yaml").write_text("proyectos:\n", encoding="utf-8")
    filas = []
    for codigo, nombre in (("AA1", "Portal viejo"), ("AA2", "Portal nuevo")):
        fila = _fila()
        fila["J"] = f"[{codigo}] {nombre} | largo"
        fila["R"] = "210"
        filas.append(fila)
    _xlsx_de_kimai(
        raiz / "input" / "2026-08" / "kimai.xlsx",
        filas,
        encabezados={
            "A": "Date",
            "D": "Duration",
            "F": "User",
            "J": "Project",
            "K": "Activity",
            "R": "Project number",
        },
    )

    assert procesar_mes("2026-08", raiz) == 0
    texto = informe(raiz)
    assert "El Project number 210 aparece con 2 nombres" in texto
    assert "Portal nuevo" in texto and "Portal viejo" in texto


# --- Un export con dos desarrolladores --------------------------------------


def test_un_export_con_dos_devs_entra_igual_y_cada_fila_lleva_su_dueno(tmp_path):
    """En el detalle plano cada fila trae su Name, User y E-mail.

    El chequeo viejo (un export por persona) existía porque el Excel pivoteado
    le imputaba todas las horas a una sola persona. Acá eso no puede pasar.
    """
    from test_lector_kimai import _fila, _xlsx_de_kimai

    raiz = preparar(tmp_path, nombres=())
    _xlsx_de_kimai(
        raiz / "input" / "2026-08" / "kimai-mezclado.xlsx",
        [_fila(usuario="mzalazar"), _fila(usuario="zlopez")],
    )
    assert procesar_mes("2026-08", raiz) == 0
    hoja = hoja_de(raiz, LIMPIO)
    assert {hoja.cell(row=f, column=4).value for f in (2, 3)} == {
        "mzalazar",
        "zlopez",
    }


# --- El informe describe la CARPETA, no la corrida ---------------------------


def test_un_xlsx_ajeno_del_dueno_sobrevive_y_se_nombra_en_el_informe(tmp_path):
    raiz = preparar(tmp_path)
    salida = raiz / "output" / "2026-08"
    salida.mkdir(parents=True, exist_ok=True)
    ajeno = salida / "NOTAS DEL DUENO.xlsx"
    ajeno.write_text("notas del dueño", encoding="utf-8")

    assert procesar_mes("2026-08", raiz) == 0
    assert ajeno.read_text(encoding="utf-8") == "notas del dueño"
    texto = informe(raiz)
    assert "NOTAS DEL DUENO.xlsx" in texto
    assert "NO LOS ENVÍES" in texto


def test_los_excel_del_formato_anterior_quedan_listados_como_ajenos(tmp_path):
    """Los 'Aug Apellido.xlsx' de la época del Excel por desarrollador."""
    raiz = preparar(tmp_path)
    salida = raiz / "output" / "2026-08"
    salida.mkdir(parents=True, exist_ok=True)
    (salida / "Aug Zalazar.xlsx").write_text("formato viejo", encoding="utf-8")

    assert procesar_mes("2026-08", raiz) == 0
    texto = informe(raiz)
    assert "Aug Zalazar.xlsx" in texto
    assert "formato anterior" in texto
    assert "Sin avisos." not in texto


def test_el_informe_no_confunde_lo_generado_con_lo_ajeno(tmp_path):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    texto = informe(raiz)
    assert "NO LOS ENVÍES" not in texto
    assert f"Archivo para el partner: {LIMPIO}" in texto


def test_un_apartado_sigue_en_el_informe_en_las_corridas_siguientes(tmp_path):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")
    assert procesar_mes("2026-08", raiz) == 1

    (raiz / "input" / "2026-08" / "roto.xlsx").unlink()
    assert procesar_mes("2026-08", raiz) == 0
    texto = informe(raiz)
    assert "Apartados por la herramienta" in texto
    assert "CORRIDA ANTERIOR - NO ENVIAR" in texto


def test_si_el_archivo_viejo_no_se_puede_apartar_el_informe_lo_dice(
    tmp_path, monkeypatch
):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0

    def rename_bloqueado(self, destino):
        raise PermissionError(13, "Acceso denegado")

    (raiz / "input" / "2026-08" / "roto.xlsx").write_text("basura", encoding="utf-8")
    monkeypatch.setattr(Path, "rename", rename_bloqueado)
    assert procesar_mes("2026-08", raiz) == 1

    salida = raiz / "output" / "2026-08"
    assert (salida / LIMPIO).is_file()  # no se pudo mover, sigue ahí
    texto = informe(raiz)
    assert "no se pudo apartar" in texto
    assert "NO lo envíes" in texto


def test_un_destino_de_solo_lectura_no_deja_un_archivo_a_medias(tmp_path):
    raiz = preparar(tmp_path)
    salida = raiz / "output" / "2026-08"
    assert procesar_mes("2026-08", raiz) == 0
    previo = (salida / LIMPIO).read_bytes()

    (salida / LIMPIO).chmod(stat.S_IREAD)
    try:
        codigo = procesar_mes("2026-08", raiz)
    finally:
        for archivo in salida.glob("*.xlsx"):
            archivo.chmod(stat.S_IWRITE | stat.S_IREAD)

    assert codigo == 1
    assert (salida / LIMPIO).read_bytes() == previo


def test_si_no_se_puede_escribir_el_informe_va_a_un_archivo_alternativo(
    tmp_path, capsys, monkeypatch
):
    raiz = preparar(tmp_path)
    assert procesar_mes("2026-08", raiz) == 0
    salida = raiz / "output" / "2026-08"
    viejo = (salida / "_validacion.txt").read_text(encoding="utf-8")

    write_text_real = Path.write_text

    def write_text_bloqueado(self, contenido, *args, **kwargs):
        if self.name == "_validacion.txt":
            raise PermissionError(13, "Acceso denegado")
        return write_text_real(self, contenido, *args, **kwargs)

    monkeypatch.setattr(Path, "write_text", write_text_bloqueado)
    assert procesar_mes("2026-08", raiz) == 1

    alternativos = [
        p for p in salida.glob("_validacion*.txt") if p.name != "_validacion.txt"
    ]
    assert len(alternativos) == 1
    assert "NO SE PUDO ESCRIBIR" in alternativos[0].name
    assert (salida / "_validacion.txt").read_text(encoding="utf-8") == viejo
    assert "no se pudo escribir" in capsys.readouterr().out


# --- Punta a punta sobre los cinco exports reales de input/2026-08/ ----------


def test_punta_a_punta_con_un_mapeo_incompleto(tmp_path):
    """Los cinco archivos que el dueño exportó de verdad para agosto de 2026.

    Tres son el reporte de detalle y entran solos. Los otros dos son el
    resumen mensual, que se acepta, pero el mapeo de este test no tiene el
    usuario ni el mail de Alexis Carnero ni el número de proyecto de
    Victorius 3, y esas columnas no se inventan: esos dos archivos no entran
    y el nombre del archivo lo dice.

    El mapeo sale de `tests/fixtures/mapeo-incompleto-de-prueba.yaml` y no de
    `config/mapeo.yaml`: lo que se prueba es qué hace el programa cuando al
    mapeo le falta un dato, no el estado de la configuración de producción,
    que el dueño completa cuando consigue los datos.
    """
    entrada_real = RAIZ / "input" / "2026-08"
    (tmp_path / "config").mkdir()
    shutil.copy(
        FIXTURES / "mapeo-incompleto-de-prueba.yaml", tmp_path / "config" / "mapeo.yaml"
    )
    shutil.copytree(entrada_real, tmp_path / "input" / "2026-08")

    assert procesar_mes("2026-08", tmp_path) == 1

    detallados = ["franco.csv", "luciano.xlsx", "matias.xlsx"]
    resumidos = ["alexis.xlsx", "lautaro.xlsx"]
    # Los exports del mes son esos cinco y nada más. `manual/` es la carpeta
    # de las planillas que no salen de Kimai, y el programa no la mira.
    assert sorted(p.name for p in entrada_real.iterdir() if p.is_file()) == sorted(
        detallados + resumidos
    )

    esperadas = (
        sum(round(r.horas * 3600) for n in detallados for r in leer(entrada_real / n))
        / 3600
    )

    hoja = hoja_de(tmp_path, incompleto(2))
    assert horas_de(hoja) == esperadas == 147.0
    assert hoja.max_row - 1 == 47

    nombres = {hoja.cell(row=f, column=3).value for f in range(2, hoja.max_row + 1)}
    assert nombres == {"Franco Dodera", "Luciano Carducci", "Matias Zalazar"}

    # Y lo que falta es exactamente lo de los dos resúmenes mensuales.
    faltantes = sum(
        r.horas
        for n in resumidos
        for r in leer_resumen_mensual(entrada_real / n, leer_hoja(entrada_real / n))
    )
    assert faltantes == 159.0
    texto = informe(tmp_path)
    for nombre in resumidos:
        assert nombre in texto
    assert "resumen mensual" in texto

    # El motivo de cada uno nombra el dato que falta y trae el YAML pegable.
    assert "Alexis Carnero" in texto
    assert "GI2680001" in texto
    assert "numero_proyecto:" in texto
    assert "mail:" in texto


def test_punta_a_punta_completo_con_el_mapeo_de_prueba(tmp_path):
    """Los cinco desarrolladores adentro: 306.0 h exactas.

    Mismo input real, pero con un mapeo que sí tiene el usuario y el mail de
    Alexis y el número de proyecto de Victorius 3 (inventados: viven en el
    fixture, nunca en config/mapeo.yaml). Con eso, los dos resúmenes
    mensuales entran y el archivo sale limpio.
    """
    (tmp_path / "config").mkdir()
    shutil.copy(
        FIXTURES / "mapeo-completo-de-prueba.yaml", tmp_path / "config" / "mapeo.yaml"
    )
    shutil.copytree(RAIZ / "input" / "2026-08", tmp_path / "input" / "2026-08")

    assert procesar_mes("2026-08", tmp_path) == 0

    hoja = hoja_de(tmp_path, LIMPIO)
    assert horas_de(hoja) == 306.0
    assert hoja.max_row - 1 == 70

    nombres = {hoja.cell(row=f, column=3).value for f in range(2, hoja.max_row + 1)}
    assert nombres == {
        "Alexis Carnero",
        "Franco Dodera",
        "Lautaro Zalazar",
        "Luciano Carducci",
        "Matias Zalazar",
    }

    texto = informe(tmp_path)
    assert "5 desarrollador/es en el archivo:" in texto
    assert "NO ENTRARON" not in texto


def test_punta_a_punta_completo_las_filas_del_resumen_van_enteras(tmp_path):
    """Las tres columnas que el partner necesita llenas, llenas."""
    (tmp_path / "config").mkdir()
    shutil.copy(
        FIXTURES / "mapeo-completo-de-prueba.yaml", tmp_path / "config" / "mapeo.yaml"
    )
    shutil.copytree(RAIZ / "input" / "2026-08", tmp_path / "input" / "2026-08")
    assert procesar_mes("2026-08", tmp_path) == 0

    hoja = hoja_de(tmp_path, LIMPIO)
    filas = {}
    for f in range(2, hoja.max_row + 1):
        valores = [hoja.cell(row=f, column=c).value for c in range(1, 11)]
        filas.setdefault(valores[2], []).append(valores)

    for nombre, usuario, mail, numero in (
        ("Lautaro Zalazar", "lzalazar", "lautaro.zalazar@cunix.net", "GI2680001-DE-PRUEBA"),
        ("Alexis Carnero", "acarnero-de-prueba", "alexis.carnero@ejemplo-de-prueba.invalid", "PR2510126"),
    ):
        for valores in filas[nombre]:
            assert valores[3] == usuario
            assert valores[4] == mail
            assert valores[9] == numero
            # Lo que el resumen mensual no trae y el partner sí acepta vacío.
            assert valores[8] is None
            assert valores[0].hour == 0 and valores[0].minute == 0


def test_el_informe_dice_de_que_export_salio_cada_desarrollador(tmp_path):
    (tmp_path / "config").mkdir()
    shutil.copy(
        FIXTURES / "mapeo-completo-de-prueba.yaml", tmp_path / "config" / "mapeo.yaml"
    )
    shutil.copytree(RAIZ / "input" / "2026-08", tmp_path / "input" / "2026-08")
    assert procesar_mes("2026-08", tmp_path) == 0

    # Sólo las líneas del listado de desarrolladores, que llevan sus horas.
    lineas = [l for l in informe(tmp_path).splitlines() if " registro/s, " in l]
    origenes = {l.split(" (")[0].strip("  - "): l.split("[")[1].rstrip("]") for l in lineas}
    assert origenes == {
        "Alexis Carnero": "resumen mensual",
        "Franco Dodera": "reporte de detalle",
        "Lautaro Zalazar": "resumen mensual",
        "Luciano Carducci": "reporte de detalle",
        "Matias Zalazar": "reporte de detalle",
    }


def test_el_informe_avisa_de_las_columnas_vacias_del_resumen_mensual(tmp_path):
    """No frena nada: es lo que el dueño necesita para decidir si lo manda."""
    (tmp_path / "config").mkdir()
    shutil.copy(
        FIXTURES / "mapeo-completo-de-prueba.yaml", tmp_path / "config" / "mapeo.yaml"
    )
    shutil.copytree(RAIZ / "input" / "2026-08", tmp_path / "input" / "2026-08")
    assert procesar_mes("2026-08", tmp_path) == 0

    texto = informe(tmp_path)
    assert "SIN DESCRIPCIÓN Y SIN HORA DE INICIO" in texto
    assert "Alexis Carnero" in texto
    assert "Lautaro Zalazar" in texto
    assert "reporte de detalle y corré de nuevo" in texto
    # Los que exportaron con el detalle no aparecen en esa lista.
    aviso = texto.split("SIN DESCRIPCIÓN Y SIN HORA DE INICIO")[1].split("Esto NO")[0]
    assert "Franco Dodera" not in aviso


def test_los_tres_exports_de_detalle_siguen_dando_las_mismas_horas(tmp_path):
    """62.5, 76.5 y 8.0: lo que este cambio no puede haber tocado."""
    (tmp_path / "config").mkdir()
    shutil.copy(
        FIXTURES / "mapeo-completo-de-prueba.yaml", tmp_path / "config" / "mapeo.yaml"
    )
    shutil.copytree(RAIZ / "input" / "2026-08", tmp_path / "input" / "2026-08")
    assert procesar_mes("2026-08", tmp_path) == 0

    hoja = hoja_de(tmp_path, LIMPIO)
    horas = {}
    for f in range(2, hoja.max_row + 1):
        nombre = hoja.cell(row=f, column=3).value
        segundos = hoja.cell(row=f, column=2).value.total_seconds()
        horas[nombre] = horas.get(nombre, 0.0) + segundos / 3600

    assert horas["Franco Dodera"] == 62.5
    assert horas["Matias Zalazar"] == 76.5
    assert horas["Luciano Carducci"] == 8.0
    assert horas["Lautaro Zalazar"] == 9.0
    assert horas["Alexis Carnero"] == 150.0


def test_punta_a_punta_el_project_number_de_luciano_es_210(tmp_path):
    """El caso que distingue el Project number del código entre corchetes."""
    (tmp_path / "config").mkdir()
    shutil.copy(RAIZ / "config" / "mapeo.yaml", tmp_path / "config" / "mapeo.yaml")
    (tmp_path / "input" / "2026-08").mkdir(parents=True)
    shutil.copy(
        RAIZ / "input" / "2026-08" / "luciano.xlsx",
        tmp_path / "input" / "2026-08" / "luciano.xlsx",
    )

    assert procesar_mes("2026-08", tmp_path) == 0
    hoja = hoja_de(tmp_path, LIMPIO)
    numeros = {hoja.cell(row=f, column=10).value for f in range(2, hoja.max_row + 1)}
    assert numeros == {"210"}
    assert "AD2690002" not in numeros
