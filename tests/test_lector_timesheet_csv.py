"""Tests del lector de timesheet plano en CSV."""
import csv
from datetime import date, time
from pathlib import Path

import pytest
from conftest import FIXTURES

from cunix_horas.lector_kimai import ErrorLectura, Registro, leer

FRANCO = FIXTURES / "kimai-timesheet-csv-franco.csv"


def _columnas_de(origen: Path) -> list[str]:
    with origen.open(encoding="utf-8-sig", newline="") as archivo:
        return list(csv.DictReader(archivo).fieldnames)


def _reescribir(
    origen: Path, destino: Path, orden: list[str] | None = None, bom: bool = False
) -> Path:
    """Copia el CSV con las columnas en otro orden y/o con BOM al principio."""
    with origen.open(encoding="utf-8-sig", newline="") as archivo:
        lector = csv.DictReader(archivo)
        columnas = list(orden or lector.fieldnames)
        filas = list(lector)
    codificacion = "utf-8-sig" if bom else "utf-8"
    with destino.open("w", encoding=codificacion, newline="") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=columnas, extrasaction="ignore")
        escritor.writeheader()
        escritor.writerows(filas)
    return destino


def test_lee_los_15_registros_del_csv():
    registros = leer(FRANCO)
    assert len(registros) == 15
    assert all(isinstance(r, Registro) for r in registros)


def test_suma_62_horas_y_media():
    assert round(sum(r.horas for r in leer(FRANCO)), 2) == 62.5


def test_trae_el_username_y_el_codigo_de_proyecto():
    registros = leer(FRANCO)
    assert {r.username for r in registros} == {"franco"}
    assert {r.cod_proyecto for r in registros} == {"PR2510126"}


def test_todas_las_fechas_caen_en_agosto_de_2026():
    fechas = [r.fecha for r in leer(FRANCO)]
    assert all(f.year == 2026 and f.month == 8 for f in fechas)
    assert min(fechas) == date(2026, 8, 3)
    assert max(fechas) == date(2026, 8, 31)


def test_la_duracion_viene_en_horas_y_minutos():
    """'2:00' son 2 horas; nada de fracción de día como en el .xlsx."""
    del_31 = [r for r in leer(FRANCO) if r.fecha == date(2026, 8, 31)]
    assert round(sum(r.horas for r in del_31), 2) == 2.0


def test_lee_por_nombre_de_columna_y_no_por_posicion(tmp_path):
    """Columnas dadas vuelta Y con otras intercaladas: el resultado es el mismo.

    Darlas vuelta sola no alcanza: una versión nueva de Kimai agrega columnas
    en el medio, que es lo que corre de lugar a las que importan.
    """
    orden = []
    for indice, columna in enumerate(reversed(_columnas_de(FRANCO))):
        orden.append(f"Columna nueva {indice}")
        orden.append(columna)
    orden.append("Columna nueva final")
    revuelto = _reescribir(FRANCO, tmp_path / "revuelto.csv", orden=orden)
    assert leer(revuelto) == leer(FRANCO)


def test_tolera_el_bom_al_principio_del_archivo(tmp_path):
    con_bom = _reescribir(FRANCO, tmp_path / "con-bom.csv", bom=True)
    assert con_bom.read_bytes().startswith(b"\xef\xbb\xbf")
    assert leer(con_bom) == leer(FRANCO)


def test_falla_si_falta_una_columna_esperada(tmp_path):
    ruta = tmp_path / "sin-duration.csv"
    ruta.write_text(
        "Date,User,Project,Activity\n2026-08-31,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura, match="no tiene el formato"):
        leer(ruta)


def test_una_duracion_ilegible_da_un_error_en_espanol(tmp_path):
    ruta = tmp_path / "roto.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "2026-08-31,dos horas,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "roto.csv" in mensaje
    assert "fila 2" in mensaje
    assert "no es una duración" in mensaje


def test_una_fecha_ilegible_da_un_error_en_espanol(tmp_path):
    ruta = tmp_path / "roto.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "31/08/2026,2:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    assert "no es una fecha" in str(excepcion.value)


# --- I4: una fila con horas pero sin fecha son horas que no se facturan ----


def test_falla_si_una_fila_trae_duracion_pero_no_fecha(tmp_path):
    """Saltearla devolvía 1 registro y 1 h, con 2 h perdidas y sin ninguna señal."""
    ruta = tmp_path / "sin-fecha.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        ",2:00,franco,[PR2510126] x,Desarrollo\n"
        "2026-08-31,1:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "sin-fecha.csv" in mensaje
    assert "fila 2" in mensaje
    assert "'2:00'" in mensaje


def test_una_fila_completamente_vacia_se_saltea(tmp_path):
    """Ahí no hay horas que perder: saltearla está bien."""
    ruta = tmp_path / "con-vacia.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "2026-08-31,1:00,franco,[PR2510126] x,Desarrollo\n"
        ",,,,\n"
        "\n"
        "2026-08-30,2:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    registros = leer(ruta)
    assert len(registros) == 2
    assert sum(r.horas for r in registros) == 3.0


# --- I5: validación y lectura usan exactamente las mismas claves -----------


def test_lee_igual_con_los_encabezados_llenos_de_espacios(tmp_path):
    """Antes la validación pasaba y devolvía 0 registros y 0 horas sin error."""
    ruta = tmp_path / "con-espacios.csv"
    ruta.write_text(
        " Date , Duration , User , Project , Activity \n"
        "2026-08-31,2:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    registros = leer(ruta)
    assert len(registros) == 1
    assert registros[0].horas == 2.0
    assert registros[0].username == "franco"


# --- M5: dos columnas con el mismo nombre ----------------------------------


def test_falla_si_hay_una_columna_duplicada(tmp_path):
    """Con dos 'Duration' ganaba la última y devolvía 9 h sin chistar."""
    ruta = tmp_path / "duplicada.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity,Duration\n"
        "2026-08-31,2:00,franco,[PR2510126] x,Desarrollo,9:00\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "más de una columna con el mismo nombre" in mensaje
    assert "Duration" in mensaje


# --- M6: Kimai emite el CSV con punto y coma en algunos locales ------------


def test_lee_un_csv_separado_por_punto_y_coma(tmp_path):
    ruta = tmp_path / "punto-y-coma.csv"
    ruta.write_text(
        "Date;Duration;User;Project;Activity\n"
        "2026-08-31;2:00;franco;[PR2510126] x;Desarrollo\n"
        "2026-08-30;1:30;franco;[PR2510126] x;Gestión\n",
        encoding="utf-8",
    )
    registros = leer(ruta)
    assert len(registros) == 2
    assert sum(r.horas for r in registros) == 3.5
    assert {r.actividad for r in registros} == {"Desarrollo", "Gestión"}


def test_un_csv_que_no_es_de_kimai_nombra_los_dos_separadores(tmp_path):
    ruta = tmp_path / "otra-cosa.csv"
    ruta.write_text("Fecha;Horas\n2026-08-31;2\n", encoding="utf-8")
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "no tiene el formato" in mensaje
    assert "coma" in mensaje
    assert "punto y coma" in mensaje


# --- M3: el proyecto sin código dice en qué archivo y en qué fila ----------


def test_un_proyecto_sin_codigo_dice_donde_esta(tmp_path):
    ruta = tmp_path / "sin-codigo.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "2026-08-31,2:00,franco,Aduana-Subastas,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "sin-codigo.csv" in mensaje
    assert "fila 2" in mensaje
    assert "sin código" in mensaje


# --- El detalle plano: los datos que el partner pide en cada fila ----------


def test_el_primer_registro_del_csv_real_trae_todo_el_detalle():
    primero = leer(FRANCO)[0]
    assert primero.fecha == date(2026, 8, 31)
    assert primero.hora_inicio == time(13, 0)
    assert primero.nombre == "Franco Dodera"
    assert primero.username == "franco"
    assert primero.email == "fanco.dodera@cunix.net"
    assert primero.descripcion == (
        "Desarrollo script para parametrizacion, y solucion bug, no se pudo "
        "probar correctamente falta VPN"
    )
    assert primero.numero_proyecto == "PR2510126"
    assert primero.texto_cliente.startswith("[618010007] ")


def test_el_nombre_para_mostrar_y_el_username_son_dos_campos_distintos():
    registros = leer(FRANCO)
    assert {r.username for r in registros} == {"franco"}
    assert {r.nombre for r in registros} == {"Franco Dodera"}


def test_todos_los_registros_traen_hora_de_inicio_nombre_y_mail():
    registros = leer(FRANCO)
    assert all(r.hora_inicio is not None for r in registros)
    assert all(r.nombre and r.email for r in registros)


def test_un_registro_del_csv_sin_descripcion_queda_vacio_y_no_rompe(tmp_path):
    """En el archivo de septiembre del partner 112 de 160 filas no la traen."""
    ruta = tmp_path / "sin-descripcion.csv"
    ruta.write_text(
        "Date,From,Duration,Name,User,E-mail,Customer,Project,Activity,"
        "Description,Project number\n"
        "2026-08-31,09:30,2:00,Franco Dodera,franco,franco@cunix.net,"
        "[618010007] MINVU,[PR2510126] x,Desarrollo,,210\n",
        encoding="utf-8",
    )
    registro = leer(ruta)[0]
    assert registro.descripcion == ""
    assert registro.hora_inicio == time(9, 30)
    assert registro.nombre == "Franco Dodera"
    assert registro.numero_proyecto == "210"


def test_un_csv_sin_las_columnas_del_detalle_se_sigue_leyendo(tmp_path):
    """Los campos nuevos quedan vacíos; las horas, que es lo que factura, no."""
    ruta = tmp_path / "viejo.csv"
    ruta.write_text(
        "Date,Duration,User,Project,Activity\n"
        "2026-08-31,2:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    registro = leer(ruta)[0]
    assert registro.horas == 2.0
    assert registro.hora_inicio is None
    assert (registro.nombre, registro.email, registro.numero_proyecto) == (
        "",
        "",
        "",
    )


def test_una_hora_de_inicio_ilegible_da_un_error_en_espanol(tmp_path):
    """Esa hora va en la celda de fecha del entregable: mal leída mueve el día."""
    ruta = tmp_path / "hora-rota.csv"
    ruta.write_text(
        "Date,From,Duration,User,Project,Activity\n"
        "2026-08-31,mediodia,2:00,franco,[PR2510126] x,Desarrollo\n",
        encoding="utf-8",
    )
    with pytest.raises(ErrorLectura) as excepcion:
        leer(ruta)
    mensaje = str(excepcion.value)
    assert "hora-rota.csv" in mensaje
    assert "fila 2" in mensaje
    assert "From" in mensaje
    assert "formato de hora" in mensaje


def test_el_detalle_tambien_se_lee_por_nombre_de_columna(tmp_path):
    """Con las columnas dadas vuelta, el detalle tiene que salir igual."""
    revuelto = _reescribir(
        FRANCO, tmp_path / "revuelto.csv", orden=list(reversed(_columnas_de(FRANCO)))
    )
    assert leer(revuelto)[0] == leer(FRANCO)[0]
