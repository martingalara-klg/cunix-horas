# -*- coding: utf-8 -*-
"""Edicion puntual de los dos anexos de agosto 2026 que C.UNIX envio como ejemplo.

QUE HACE
    Lee los originales de la raiz del repositorio (nunca los modifica) y escribe
    dos copias corregidas en ``output/2026-08/`` con el mismo nombre:

      * ``Anexo-II-A-Detalle-horas-KLG-2026-08.xlsx``
      * ``Anexo-II-Informe-mensual-horas-KLG-2026-08 (1).docx``

    La correccion central es separar las horas de Gabriel Denis de las de
    Alexis Carnero. Gabriel trabajo 91 h en agosto 2026 pero no tiene usuario
    de Kimai: sus horas quedaron cargadas en la cuenta de Alexis. Las horas de
    Gabriel salen de ``horas_agosto_2026 (2).xlsx`` (hoja ``Agosto 2026``).
    Ademas se completan los pendientes que C.UNIX dejo marcados para que los
    llene KLG (perfiles, usuarios de Kimai, principales trabajos y
    observaciones).

POR QUE ACA Y NO EN LA HERRAMIENTA
    ``cunix_horas/`` genera los entregables a partir de la exportacion de Kimai.
    La correccion de Gabriel no es una regla general del proceso: es un arreglo
    de un mes concreto sobre datos que Kimai nunca registro, apoyado en una
    planilla manual que no forma parte del circuito. Meterlo en la herramienta
    obligaria a modelar "persona sin usuario de Kimai cuyas horas viven en la
    cuenta de otro", que es exactamente lo que el informe se compromete a que
    no vuelva a pasar. Por eso vive como script de un solo uso, versionado para
    poder auditarlo y repetirlo, y la herramienta queda intacta.

HORARIOS RECONSTRUIDOS
    Kimai guardo un unico bloque continuo por dia a nombre de Alexis. Al separar
    las horas, cada bloque compartido se parte por horario: Alexis toma el
    principio y Gabriel el final (04/08, bloque 08:15-23:15, Alexis 7 h y
    Gabriel 8 h -> Alexis 08:15-15:15 y Gabriel 15:15-23:15). Los dias que
    quedan enteros de Gabriel (2 y 20 de agosto) conservan el horario original
    y Alexis no tiene fila. Esos horarios son una RECONSTRUCCION, no un
    registro: las horas de cada uno si estan respaldadas, el reparto dentro del
    dia no. Queda declarado en la columna "Nota KLG" de la hoja Detalle y en la
    tabla de Observaciones del informe.

USO
    python scripts/anexos_agosto_2026.py
"""

from __future__ import annotations

import copy
import datetime as dt
import re
import shutil
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from docx import Document

# --- Constantes que el dueño puede querer cambiar ---------------------------

PERFIL = "Desarrollador"

USUARIOS_KIMAI = {
    "Alexis Carnero": "acarnero",
    "Franco Dodera": "franco",
    "Matias Zalazar": "mzalazar",
    "Lautaro Zalazar": "lzalazar",
    "Luciano Carducci": "lcarducci",
    # Gabriel Denis no tiene usuario de Kimai: se deja vacio a proposito.
    "Gabriel Denis": "",
}

GABRIEL = "Gabriel Denis"
ALEXIS = "Alexis Carnero"
PROYECTO_GABRIEL = "SELICO"
HORAS_ESPERADAS_ALEXIS = 59.0

MARCA_PENDIENTE = "[●]"

# Decisiones del dueño (2026-09-30).
# ESTADO_AL_CIERRE se pone igual en las 21 filas de la tabla 3 a pedido suyo;
# dijo que despues corrige a mano las que no correspondan.
ESTADO_AL_CIERRE = "Terminado"
# Plazo de entrega que KLG se compromete a cumplir, en la tabla de condiciones.
DIAS_HABILES_ENTREGA = "5"


ENCABEZADO_NOTA = "Nota KLG"
NOTA_DIA_COMPARTIDO = (
    "Horario reconstruido por KLG: Kimai registro un unico bloque diario en la "
    "cuenta de Alexis Carnero. El bloque se partio por horario entre Alexis "
    "Carnero (principio) y Gabriel Denis (final). Las horas estan respaldadas; "
    "el corte dentro del dia es una reconstruccion, no un registro."
)
NOTA_DIA_ENTERO = (
    "Dia completo de Gabriel Denis, que no tiene usuario de Kimai. Se conserva "
    "el horario del bloque original registrado en la cuenta de Alexis Carnero."
)

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "output" / "2026-08"
ORIGEN_XLSX = RAIZ / "Anexo-II-A-Detalle-horas-KLG-2026-08.xlsx"
ORIGEN_DOCX = RAIZ / "Anexo-II-Informe-mensual-horas-KLG-2026-08 (1).docx"
ORIGEN_GABRIEL = RAIZ / "horas_agosto_2026 (2).xlsx"

PRIMERA_FILA = 2
ULTIMA_FILA_RANGO = 601
COL_FECHA, COL_INICIO, COL_FIN, COL_PERSONA = 1, 2, 3, 4
COL_PROYECTO, COL_DESCRIPCION, COL_HORAS = 5, 6, 7
COL_NOTA = 14  # columna N, fuera de todos los rangos de formulas (A..M)


# --- Lectura ----------------------------------------------------------------


def leer_horas_de_gabriel():
    """Devuelve {dia del mes: (horas, descripcion)} desde la planilla manual.

    Solo las filas 5 a 20 son registros; las 22-26 son subtotales semanales y
    la 28 es el total. La fecha viene como texto en español sin año
    ("Dom 02 Ago"): el dia es lo unico que importa, el mes y el año son agosto
    de 2026.
    """
    hoja = openpyxl.load_workbook(ORIGEN_GABRIEL, data_only=True)["Agosto 2026"]
    registros = {}
    for fila in range(5, 21):
        texto_fecha = hoja.cell(fila, 1).value
        horas = hoja.cell(fila, 2).value
        descripcion = hoja.cell(fila, 3).value
        if not texto_fecha or not horas:
            raise SystemExit(f"ABORTADO: fila {fila} de la planilla de Gabriel incompleta.")
        coincidencia = re.search(r"\b(\d{1,2})\b", str(texto_fecha))
        if not coincidencia:
            raise SystemExit(f"ABORTADO: no se pudo leer el dia de «{texto_fecha}».")
        dia = int(coincidencia.group(1))
        if dia in registros:
            raise SystemExit(f"ABORTADO: dia {dia} repetido en la planilla de Gabriel.")
        registros[dia] = (float(horas), str(descripcion).strip())
    return registros


def leer_detalle(hoja):
    """Lee las filas con datos de la hoja Detalle como diccionarios."""
    filas = []
    for fila in range(PRIMERA_FILA, ULTIMA_FILA_RANGO + 1):
        fecha = hoja.cell(fila, COL_FECHA).value
        if fecha is None:
            continue
        filas.append(
            {
                "fecha": fecha,
                "inicio": hoja.cell(fila, COL_INICIO).value,
                "fin": hoja.cell(fila, COL_FIN).value,
                "persona": hoja.cell(fila, COL_PERSONA).value,
                "proyecto": hoja.cell(fila, COL_PROYECTO).value,
                "descripcion": hoja.cell(fila, COL_DESCRIPCION).value,
                "horas": float(hoja.cell(fila, COL_HORAS).value),
                "nota": None,
            }
        )
    return filas


# --- Reparto de horas -------------------------------------------------------


def verificar_aritmetica(filas, gabriel):
    """Aborta si el reparto no cierra. Se corre antes de escribir nada."""
    bloques_alexis = {}
    for f in filas:
        if f["persona"] != ALEXIS:
            continue
        bloques_alexis.setdefault(f["fecha"].day, []).append(f)

    total_alexis = sum(f["horas"] for b in bloques_alexis.values() for f in b)
    total_gabriel = sum(horas for horas, _ in gabriel.values())

    problemas = []
    if round(total_alexis - total_gabriel, 6) != HORAS_ESPERADAS_ALEXIS:
        problemas.append(
            f"Alexis {total_alexis} - Gabriel {total_gabriel} != {HORAS_ESPERADAS_ALEXIS}"
        )
    for dia, (horas, _) in sorted(gabriel.items()):
        if dia not in bloques_alexis:
            problemas.append(f"Gabriel tiene el dia {dia} y Alexis no.")
        elif round(sum(f["horas"] for f in bloques_alexis[dia]) - horas, 6) < 0:
            problemas.append(f"Dia {dia}: Alexis quedaria en negativo.")
    for dia, bloques in sorted(bloques_alexis.items()):
        if len(bloques) != 1:
            problemas.append(f"Dia {dia}: Alexis no tiene exactamente un bloque.")
    if problemas:
        raise SystemExit("ABORTADO, la aritmetica no cierra:\n  " + "\n  ".join(problemas))


def _a_minutos(hora):
    return hora.hour * 60 + hora.minute


def _a_hora(minutos):
    minutos %= 24 * 60
    return dt.time(minutos // 60, minutos % 60)


def repartir(filas, gabriel):
    """Separa las horas de Gabriel de las de Alexis, dia por dia.

    El corte es por horario: Alexis se queda con el principio del bloque y
    Gabriel con el final. Es una reconstruccion, no un registro.
    """
    resultado = []
    for f in filas:
        dia = f["fecha"].day
        if f["persona"] != ALEXIS or dia not in gabriel:
            resultado.append(f)
            continue

        horas_gabriel, descripcion_gabriel = gabriel[dia]
        horas_alexis = round(f["horas"] - horas_gabriel, 6)
        corte = _a_hora(_a_minutos(f["inicio"]) + int(round(horas_alexis * 60)))

        if horas_alexis > 0:
            fila_alexis = dict(f)
            fila_alexis["fin"] = corte
            fila_alexis["horas"] = horas_alexis
            fila_alexis["nota"] = NOTA_DIA_COMPARTIDO
            resultado.append(fila_alexis)
            inicio_gabriel, nota_gabriel = corte, NOTA_DIA_COMPARTIDO
        else:
            # El dia entero pasa a Gabriel: Alexis no tiene fila ese dia.
            inicio_gabriel, nota_gabriel = f["inicio"], NOTA_DIA_ENTERO

        resultado.append(
            {
                "fecha": f["fecha"],
                "inicio": inicio_gabriel,
                "fin": f["fin"],
                "persona": GABRIEL,
                "proyecto": PROYECTO_GABRIEL,
                "descripcion": descripcion_gabriel,
                "horas": horas_gabriel,
                "nota": nota_gabriel,
            }
        )

    # Mismo orden que ya tenia la hoja: por fecha y, dentro del dia, por persona.
    resultado.sort(key=lambda f: (f["fecha"], f["persona"]))
    return resultado


# --- Escritura del Excel ----------------------------------------------------


def escribir_excel(filas):
    destino = SALIDA / ORIGEN_XLSX.name
    # Sin data_only: las formulas se conservan como formulas.
    libro = openpyxl.load_workbook(ORIGEN_XLSX)

    datos = libro["Datos"]
    for fila in range(8, 13):
        persona = datos.cell(fila, 1).value
        datos.cell(fila, 2).value = PERFIL
        datos.cell(fila, 3).value = USUARIOS_KIMAI[persona] or None
    # Gabriel Denis se suma al equipo, sin usuario de Kimai.
    for columna in (1, 2, 3):
        datos.cell(13, columna)._style = copy.copy(datos.cell(12, columna)._style)
    datos.cell(13, 1).value = GABRIEL
    datos.cell(13, 2).value = PERFIL
    datos.cell(13, 3).value = None

    detalle = libro["Detalle"]
    detalle.cell(1, COL_NOTA).value = ENCABEZADO_NOTA
    detalle.cell(1, COL_NOTA)._style = copy.copy(detalle.cell(1, 13)._style)
    detalle.column_dimensions["N"].width = 60

    for indice, f in enumerate(filas):
        fila = PRIMERA_FILA + indice
        detalle.cell(fila, COL_FECHA).value = f["fecha"]
        detalle.cell(fila, COL_INICIO).value = f["inicio"]
        detalle.cell(fila, COL_FIN).value = f["fin"]
        detalle.cell(fila, COL_PERSONA).value = f["persona"]
        detalle.cell(fila, COL_PROYECTO).value = f["proyecto"]
        detalle.cell(fila, COL_DESCRIPCION).value = f["descripcion"]
        detalle.cell(fila, COL_HORAS).value = f["horas"]
        detalle.cell(fila, COL_NOTA).value = f["nota"]

    # Las filas sobrantes quedan vacias. Las formulas de Alertas (H),
    # Horas a pagar (K) y Dia nuevo (M) ya vienen escritas en todas las filas
    # del rango 2:601 del original, incluidas las que ahora estrenan datos, asi
    # que no hay que copiarlas ni adaptarlas: alcanza con no pisarlas.
    for fila in range(PRIMERA_FILA + len(filas), ULTIMA_FILA_RANGO + 1):
        for columna in list(range(COL_FECHA, COL_HORAS + 1)) + [COL_NOTA]:
            detalle.cell(fila, columna).value = None

    SALIDA.mkdir(parents=True, exist_ok=True)
    libro.save(destino)
    restaurar_validaciones_x14(ORIGEN_XLSX, destino)
    return destino


def restaurar_validaciones_x14(origen, destino):
    """Reinyecta el bloque extLst que openpyxl descarta al guardar.

    Son las listas desplegables de Persona y Proyecto de la hoja Detalle, que
    apuntan a la hoja Datos. openpyxl no soporta esa extension y la elimina.
    """
    with zipfile.ZipFile(origen) as z:
        original = z.read("xl/worksheets/sheet4.xml").decode("utf-8")
    coincidencia = re.search(r"<extLst>.*?</extLst>", original, re.S)
    if not coincidencia:
        return
    # xr:uid solo tiene sentido con el namespace xr, que openpyxl no declara en
    # la raiz de la hoja que genera. Sin quitarlo, el archivo no abre.
    bloque = re.sub(r'\s+xr:uid="[^"]*"', "", coincidencia.group(0))

    temporal = destino.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(destino) as entrada, zipfile.ZipFile(
        temporal, "w", zipfile.ZIP_DEFLATED
    ) as salida:
        for elemento in entrada.infolist():
            contenido = entrada.read(elemento.filename)
            if elemento.filename == "xl/worksheets/sheet4.xml":
                texto = contenido.decode("utf-8")
                if "<extLst>" not in texto:
                    texto = texto.replace("</worksheet>", bloque + "</worksheet>")
                contenido = texto.encode("utf-8")
            salida.writestr(elemento, contenido)
    shutil.move(str(temporal), str(destino))


# --- Contenido redactado para el informe ------------------------------------

# Agrupacion de las 91 h de Gabriel (a partir de las descripciones reales de su
# planilla) mas la linea de las 59 h de Alexis, que siguen sin descripcion en
# Kimai. Reemplazan a la unica fila que decia "Alexis Carnero: sin descripcion
# en Kimai ... 150,0". Suman 150,0, asi que el total de la tabla sigue en 306,0.
PRINCIPALES_TRABAJOS = [
    {
        "proyecto": "MinVu",
        "ticket": "BUG-102 a BUG-117",
        "trabajo": (
            "Registros Técnicos (RRTT), Gabriel Denis: corrección de los bugs "
            "reportados sobre el Informe Jurídico y el Informe Técnico, "
            "incluidos el BUG-107 (tipo de artículo 6 mal clasificado) y el "
            "incidente urgente de producción del 20/08; selector D.S. 135 del "
            "Módulo 1"
        ),
        "horas": 62.0,
    },
    {
        "proyecto": "MinVu",
        "ticket": "Sin ticket",
        "trabajo": (
            "Registros Técnicos (RRTT), Gabriel Denis: auditoría de las "
            "entregas v18 a v35 y de los ASPs entregados contra el historial "
            "del repositorio, recuperación de 16 scripts del Req17 y "
            "relevamiento de la deuda técnica documental"
        ),
        "horas": 16.0,
    },
    {
        "proyecto": "MinVu",
        "ticket": "Sin ticket",
        "trabajo": (
            "Registros Técnicos (RRTT), Gabriel Denis: traspaso del Req27 "
            "(CambioCategoria) y entrega v41 con la corrección urgente de la "
            "conexión faltante"
        ),
        "horas": 8.0,
    },
    {
        "proyecto": "MinVu",
        "ticket": "Sin ticket",
        "trabajo": (
            "Registros Técnicos (RRTT), Gabriel Denis: revisión final del "
            "manual de usuario y reuniones de coordinación del cierre del mes"
        ),
        "horas": 5.0,
    },
    {
        "proyecto": "MinVu",
        "ticket": "Sin ticket",
        "trabajo": (
            "Alexis Carnero: sin descripción en Kimai. KLG debe detallar el "
            "trabajo realizado."
        ),
        "horas": 59.0,
    },
]

EXPLICACIONES = {
    "alertas": (
        "Explicación de KLG: las alertas se originan en que 91,0 de las 150,0 h "
        "atribuidas a Alexis Carnero eran en realidad de Gabriel Denis, que "
        "trabajó sin usuario de Kimai. Separadas en este anexo, Alexis Carnero "
        "queda con 59,0 h en 17 días (promedio 3,5 h por día) y Gabriel Denis "
        "con 91,0 h en 16 días (promedio 5,7 h por día), y ningún día de "
        "ninguno de los dos supera las 9 h: la alerta de «más de 9 h en el "
        "día» del 04, 05, 06, 10, 11, 13 y 14/08 deja de corresponder. El "
        "domingo 02/08 (2 h) también era de Gabriel Denis y queda a su nombre. "
        "El cruce de medianoche del 07/08 de Franco Dodera (23:00 a 02:00) es "
        "real: fueron reuniones y revisiones que terminaron pasada la "
        "medianoche y Kimai las registró como un único bloque. Sobre «sin "
        "ticket»: en agosto los trabajos de SELICO, VictoriusCP2 y Victorius 3 "
        "no se venían gestionando con tickets de iTop; desde septiembre toda "
        "carga en Kimai empieza con el ticket de iTop o la tarea de ClickUp. "
        "Las 59,0 h de Alexis Carnero siguen sin descripción en Kimai: KLG las "
        "está reconstruyendo con él y las enviará como corrección de este "
        "anexo. Aclaración sobre los horarios: el inicio y el fin de los días "
        "compartidos entre Alexis Carnero y Gabriel Denis son una "
        "reconstrucción de KLG, no un registro de Kimai. Kimai guardó un único "
        "bloque continuo por día a nombre de Alexis, y KLG lo partió por "
        "horario asignando el principio del bloque a Alexis y el final a "
        "Gabriel (por ejemplo el 04/08: bloque de 08:15 a 23:15, Alexis de "
        "08:15 a 15:15 y Gabriel de 15:15 a 23:15). Las horas de cada uno sí "
        "están respaldadas: las de Gabriel salen de su reporte diario de "
        "agosto. Cada fila afectada del Anexo II-A lo indica en la columna "
        "«Nota KLG»."
    ),
    "equipo": (
        "Gabriel Denis se sumó al proyecto como un refuerzo previsto para "
        "unas tres semanas. Por las características "
        "del cliente final y del proyecto, su participación se extendió "
        "bastante más de lo previsto. Mientras se esperaba que fuera algo "
        "breve, sus horas se registraron en la cuenta de Alexis Carnero en "
        "lugar de tramitar un usuario propio en Kimai; al alargarse el "
        "trabajo, esa decisión inicial se arrastró, y es el origen de lo que "
        "C.UNIX observa en este período. No hubo altas ni bajas en el equipo: "
        "Gabriel Denis ya venía trabajando y lo que faltaba era su usuario. "
        "Este anexo separa sus horas día por día: Alexis Carnero queda con "
        "59,0 h y Gabriel Denis con 91,0 h, sin cambios en el total del mes "
        "(306,0 h) ni en el importe a facturar. KLG solicita a C.UNIX, junto "
        "con este informe, el alta del usuario de Kimai de Gabriel Denis, "
        "para que desde septiembre cargue sus propias horas. Hasta que el "
        "alta esté, ninguna hora de Gabriel Denis se imputa a la cuenta de "
        "otra persona. Antes de cada cierre mensual KLG verifica que toda "
        "persona con horas del mes tenga usuario propio en Kimai y que la "
        "hoja «Datos» del Anexo II-A no tenga ninguna fila con «Usuario "
        "Kimai» vacío."
    ),
    "esperas": "Sin novedades.",
}


# --- Escritura del informe --------------------------------------------------


def _formato(numero):
    """Un decimal con coma, redondeando medio hacia arriba como el original.

    Con el redondeo por defecto de Python, 2,25 daria 2,2 y cambiaria el
    promedio de Lautaro Zalazar, que esta fila no deberia tocar.
    """
    redondeado = Decimal(str(numero)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return f"{redondeado}".replace(".", ",")


def _escribir_celda(celda, texto):
    """Reemplaza el texto de una celda conservando el formato del primer run."""
    parrafo = celda.paragraphs[0]
    for extra in celda.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    if not parrafo.runs:
        parrafo.add_run("")
    parrafo.runs[0].text = texto
    for run in parrafo.runs[1:]:
        run._element.getparent().remove(run._element)


def _escribir_parrafo(parrafo, texto):
    """Reemplaza el texto de un parrafo conservando el formato del primer run."""
    if not parrafo.runs:
        parrafo.add_run("")
    parrafo.runs[0].text = texto
    for run in parrafo.runs[1:]:
        run._element.getparent().remove(run._element)


def _clonar_fila(tabla, indice_modelo, indice_destino):
    """Inserta una copia de una fila existente y devuelve la fila nueva."""
    nueva = copy.deepcopy(tabla.rows[indice_modelo]._tr)
    tabla.rows[indice_destino - 1]._tr.addnext(nueva)
    return tabla.rows[indice_destino]


def escribir_informe(filas, explicaciones):
    destino = SALIDA / ORIGEN_DOCX.name
    documento = Document(ORIGEN_DOCX)

    resumen = {}
    for f in filas:
        acumulado = resumen.setdefault(f["persona"], {"horas": 0.0, "dias": set()})
        acumulado["horas"] += f["horas"]
        acumulado["dias"].add(f["fecha"])

    # Tabla 2, horas por persona: la fila de Alexis se parte en dos. Gabriel
    # queda primero (91,0 h) y Alexis baja a la cuarta posicion, para mantener
    # el orden por horas descendente que ya tenia la tabla.
    tabla2 = documento.tables[2]
    nombres_informe = {
        "Gabriel Denis": GABRIEL,
        "Matías Zalazar": "Matias Zalazar",
        "Franco Dodera": "Franco Dodera",
        "Alexis Carnero": ALEXIS,
        "Lautaro Zalazar": "Lautaro Zalazar",
        "Luciano Carducci": "Luciano Carducci",
    }
    _escribir_celda(tabla2.rows[1].cells[0], "Gabriel Denis")
    _clonar_fila(tabla2, 1, 4)
    _escribir_celda(tabla2.rows[4].cells[0], "Alexis Carnero")
    for indice in range(1, 7):
        etiqueta = tabla2.rows[indice].cells[0].text.strip()
        if etiqueta not in nombres_informe:
            raise SystemExit(f"ABORTADO: persona inesperada en la tabla 2: «{etiqueta}».")
        datos = resumen[nombres_informe[etiqueta]]
        promedio = datos["horas"] / len(datos["dias"])
        _escribir_celda(tabla2.rows[indice].cells[1], PERFIL)
        _escribir_celda(tabla2.rows[indice].cells[2], _formato(datos["horas"]))
        _escribir_celda(tabla2.rows[indice].cells[3], str(len(datos["dias"])))
        _escribir_celda(tabla2.rows[indice].cells[4], _formato(promedio))

    # Tabla 3, principales trabajos: la primera fila (Alexis, 150,0 h sin
    # descripcion) se reemplaza por el trabajo real de Gabriel agrupado mas una
    # linea para las horas de Alexis que siguen sin descripcion.
    tabla3 = documento.tables[3]
    estado_original = tabla3.rows[1].cells[3].text.strip()
    for desplazamiento in range(1, len(PRINCIPALES_TRABAJOS)):
        _clonar_fila(tabla3, 1, 1 + desplazamiento)
    for desplazamiento, entrada in enumerate(PRINCIPALES_TRABAJOS):
        fila = tabla3.rows[1 + desplazamiento]
        _escribir_celda(fila.cells[0], entrada["proyecto"])
        _escribir_celda(fila.cells[1], entrada["ticket"])
        _escribir_celda(fila.cells[2], entrada["trabajo"])
        _escribir_celda(fila.cells[3], ESTADO_AL_CIERRE)
        _escribir_celda(fila.cells[4], _formato(entrada["horas"]))

    # Estado al cierre en TODAS las filas de la tabla 3, no solo en las nuevas:
    # las que venian de C.UNIX tambien traen la marca pendiente.
    for fila in tabla3.rows[1:]:
        if fila.cells[0].text.strip().lower().startswith("total"):
            continue
        if MARCA_PENDIENTE in fila.cells[3].text:
            _escribir_celda(fila.cells[3], ESTADO_AL_CIERRE)

    # Tabla 0, condiciones: se completa el plazo de entrega que dejo C.UNIX.
    celda_cond = documento.tables[0].rows[0].cells[0]
    if MARCA_PENDIENTE not in celda_cond.text:
        raise SystemExit("ABORTADO: la tabla de condiciones ya no tiene marca.")
    for parrafo in celda_cond.paragraphs:
        if MARCA_PENDIENTE in parrafo.text:
            _escribir_parrafo(parrafo, parrafo.text.replace(
                MARCA_PENDIENTE, DIAS_HABILES_ENTREGA))

    # Tabla 4, observaciones: se completan los tres pendientes que dejo C.UNIX.
    tabla4 = documento.tables[4]
    for indice, clave in ((1, "alertas"), (2, "equipo"), (3, "esperas")):
        celda = tabla4.rows[indice].cells[1]
        original = celda.text
        if MARCA_PENDIENTE not in original:
            raise SystemExit(f"ABORTADO: la fila {indice} de Observaciones ya no tiene marca.")
        _escribir_celda(celda, original.replace(MARCA_PENDIENTE, explicaciones[clave]))

    SALIDA.mkdir(parents=True, exist_ok=True)
    documento.save(destino)
    return destino


# --- Verificacion -----------------------------------------------------------


def verificar(destino_xlsx, destino_docx):
    libro = openpyxl.load_workbook(destino_xlsx)
    detalle = libro["Detalle"]
    filas = leer_detalle(detalle)

    por_persona = {}
    for f in filas:
        acumulado = por_persona.setdefault(
            f["persona"], {"horas": 0.0, "filas": 0, "dias": set()}
        )
        acumulado["horas"] += f["horas"]
        acumulado["filas"] += 1
        acumulado["dias"].add(f["fecha"])

    print("== Detalle ==")
    print(f"  filas con datos: {len(filas)}")
    print(f"  suma de la columna Horas: {sum(f['horas'] for f in filas):.1f}")
    for persona in sorted(por_persona):
        d = por_persona[persona]
        print(
            f"  {persona:<18} {d['horas']:>6.1f} h   {d['filas']:>2} filas   "
            f"{len(d['dias']):>2} dias   prom {d['horas'] / len(d['dias']):.1f}"
        )
    dias_alexis = sorted({f["fecha"].day for f in filas if f["persona"] == ALEXIS})
    print(f"  dias de Alexis: {dias_alexis}")
    print(f"  Alexis tiene filas el 2 o el 20: {bool({2, 20} & set(dias_alexis))}")
    print(f"  registros con horas <= 0: {sum(1 for f in filas if f['horas'] <= 0)}")
    cruces = [
        f"{f['fecha']:%d/%m} {f['persona']} {f['inicio']:%H:%M}-{f['fin']:%H:%M}"
        for f in filas
        if f["inicio"] and f["fin"] and f["fin"] < f["inicio"]
    ]
    print(f"  Fin anterior a Inicio (cruces de medianoche): {cruces}")

    print("== Formulas ==")
    resumen = libro["Resumen"]
    formulas_resumen = sum(
        1
        for fila in resumen.iter_rows(min_row=1, max_row=47)
        for c in fila
        if isinstance(c.value, str) and c.value.startswith("=")
    )
    pisadas = [
        c.coordinate
        for fila in resumen.iter_rows(min_row=11, max_row=47)
        for c in fila
        if c.value is not None and not isinstance(c.value, str)
    ]
    print(f"  Resumen: {formulas_resumen} celdas siguen siendo formula")
    print(f"  Resumen: celdas de calculo pisadas con un valor: {pisadas}")
    for columna, nombre in ((8, "Alertas"), (11, "Horas a pagar"), (13, "Dia nuevo")):
        con_formula = sum(
            1
            for fila in range(PRIMERA_FILA, PRIMERA_FILA + len(filas))
            if str(detalle.cell(fila, columna).value or "").startswith("=")
        )
        print(f"  Detalle.{nombre}: {con_formula}/{len(filas)} filas con formula")

    print("== Datos ==")
    datos = libro["Datos"]
    for fila in range(8, 24):
        if datos.cell(fila, 1).value:
            print(
                f"  {datos.cell(fila, 1).value:<18} perfil={datos.cell(fila, 2).value!r} "
                f"usuario={datos.cell(fila, 3).value!r}"
            )

    print("== Informe ==")
    documento = Document(destino_docx)
    for indice, etiqueta, columna in ((1, "tabla 1", 2), (2, "tabla 2", 2), (3, "tabla 3", 4)):
        tabla = documento.tables[indice]
        total = sum(
            float(fila.cells[columna].text.strip().replace(".", "").replace(",", "."))
            for fila in tabla.rows[1:-1]
        )
        print(
            f"  {etiqueta}: {len(tabla.rows) - 2} filas de datos, suman {total:.1f}, "
            f"total declarado {tabla.rows[-1].cells[columna].text.strip()}"
        )
    print(f"  tabla 1 importe total: {documento.tables[1].rows[-1].cells[4].text.strip()}")
    print(f"  tabla 0: {len(documento.tables[0].rows)} fila(s), intacta")
    print(f"  tabla 5: {len(documento.tables[5].rows)} filas, intacta")
    pendientes = [
        (i, ri, ci)
        for i, t in enumerate(documento.tables)
        for ri, r in enumerate(t.rows)
        for ci, c in enumerate(r.cells)
        if MARCA_PENDIENTE in c.text
    ]
    print(f"  celdas que todavia tienen marca pendiente: {pendientes}")


def main():
    gabriel = leer_horas_de_gabriel()
    libro = openpyxl.load_workbook(ORIGEN_XLSX)
    filas_originales = leer_detalle(libro["Detalle"])
    verificar_aritmetica(filas_originales, gabriel)

    filas = repartir(filas_originales, gabriel)
    destino_xlsx = escribir_excel(filas)
    destino_docx = escribir_informe(filas, EXPLICACIONES)
    print(f"Escrito: {destino_xlsx}")
    print(f"Escrito: {destino_docx}\n")
    verificar(destino_xlsx, destino_docx)


if __name__ == "__main__":
    main()
