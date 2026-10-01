"""Piezas que comparten los dos anexos: el período, los nombres y los números.

Los dos entregables del mes —el **Anexo II** (informe, Word) y el **Anexo II-A**
(detalle, Excel)— se arman del mismo período y se nombran con el mismo patrón,
así que eso vive una sola vez acá.

Los números se escriben como los escribe C.UNIX en su formato: coma decimal y
punto de miles. Nunca con `str(float)`, que daría `306.0` y `3706.0`.
"""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

# La marca que C.UNIX dejó en sus plantillas donde falta algo que completa KLG.
# Las celdas que la herramienta no sabe llenar la conservan: un hueco visible
# es preferible a un dato inventado.
MARCA_PENDIENTE = "[●]"

FORMATO_PERIODO = re.compile(r"^(\d{4})-(\d{2})$")

# Nunca `strftime("%B")`: depende del locale de la máquina y en Windows suele
# salir en inglés.
MESES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


class ErrorPeriodo(ValueError):
    """El período no tiene el formato AAAA-MM, o el mes está fuera de rango."""


@dataclass(frozen=True)
class Periodo:
    """El mes que se está generando."""

    anio: int
    mes: int

    @classmethod
    def parsear(cls, texto: str) -> "Periodo":
        coincidencia = FORMATO_PERIODO.match(texto or "")
        if coincidencia is None:
            raise ErrorPeriodo(
                f"El mes debe tener el formato AAAA-MM (ej: 2026-09), no {texto!r}"
            )
        anio, mes = int(coincidencia.group(1)), int(coincidencia.group(2))
        if not 1 <= mes <= 12:
            raise ErrorPeriodo(
                f"Mes fuera de rango en {texto!r}: AAAA-MM con MM entre 01 y 12"
            )
        return cls(anio, mes)

    @property
    def carpeta(self) -> str:
        """'2026-08': el nombre de la carpeta de input/ y de output/."""
        return f"{self.anio:04d}-{self.mes:02d}"

    @property
    def texto(self) -> str:
        """'Agosto 2026': como va en el encabezado del informe y en Datos!B2."""
        return f"{MESES_ES[self.mes].capitalize()} {self.anio}"

    @property
    def dias(self) -> int:
        return calendar.monthrange(self.anio, self.mes)[1]

    def contiene(self, fecha) -> bool:
        return (fecha.year, fecha.month) == (self.anio, self.mes)


# Nombres con los que C.UNIX espera recibir los dos anexos. `{periodo}` es el
# mes en el formato de la carpeta ('2026-09').
PATRON_DETALLE = "Anexo-II-A-Detalle-horas-KLG-{periodo}.xlsx"
PATRON_INFORME = "Anexo-II-Informe-mensual-horas-KLG-{periodo}.docx"

PERFIL_POR_DEFECTO = "Desarrollador"


@dataclass(frozen=True)
class ConfigAnexos:
    """Lo que `config/mapeo.yaml` declara bajo `anexos:`.

    Todo es opcional. Lo que no esté configurado se resuelve solo (los
    nombres de archivo, el perfil) o queda como marcador en el documento
    para que el dueño lo complete a mano (las fechas, el plazo de entrega).
    """

    archivo_detalle: str = PATRON_DETALLE
    archivo_informe: str = PATRON_INFORME
    perfil_por_defecto: str = PERFIL_POR_DEFECTO
    # persona -> perfil, para quien no sea el perfil por defecto.
    perfiles: tuple[tuple[str, str], ...] = ()
    dias_habiles_entrega: str = ""
    contrato_de_fecha: str = ""
    fecha_de_emision: str = ""

    def nombre_detalle(self, periodo: Periodo) -> str:
        return self.archivo_detalle.format(periodo=periodo.carpeta)

    def nombre_informe(self, periodo: Periodo) -> str:
        return self.archivo_informe.format(periodo=periodo.carpeta)

    @property
    def perfiles_por_persona(self) -> dict[str, str]:
        return dict(self.perfiles)


@dataclass(frozen=True)
class HorasSinKimai:
    """«Esta persona carga sus horas acá, y hay que restárselas a esta otra»."""

    persona: str
    planilla: str
    restar_a: str
    proyecto: str = ""
    nota: str = ""


@dataclass(frozen=True)
class DescripcionesAparte:
    """«Esta persona entrega sus descripciones en esta planilla»."""

    persona: str
    planilla: str


@dataclass(frozen=True)
class FuentesManuales:
    """Todo lo que `config/mapeo.yaml` declara bajo `fuentes_manuales:`.

    Vacío es lo normal: un mes en que todos cargan sus horas y sus
    descripciones en Kimai no necesita ninguna planilla manual.
    """

    horas_sin_kimai: tuple[HorasSinKimai, ...] = ()
    descripciones: tuple[DescripcionesAparte, ...] = ()

    @property
    def hay_alguna(self) -> bool:
        return bool(self.horas_sin_kimai or self.descripciones)


def marca_de_incompleto(faltan: int) -> str:
    """' (INCOMPLETO - FALTAN 2 DESARROLLADORES - NO ENVIAR)'.

    Va en el nombre de los dos anexos cuando algún export no se pudo leer. Los
    anexos se generan igual con lo que sí se leyó, para poder revisarlos, pero
    el nombre es lo único que se ve al adjuntarlos a un mail: sin esta marca,
    un mes al que le falta una persona se ve exactamente igual que uno completo.
    """
    if faltan == 1:
        return " (INCOMPLETO - FALTA 1 DESARROLLADOR - NO ENVIAR)"
    return f" (INCOMPLETO - FALTAN {faltan} DESARROLLADORES - NO ENVIAR)"


def nombre_incompleto(nombre: str, faltan: int) -> str:
    """El mismo nombre de archivo, con la marca antes de la extensión."""
    ruta = Path(nombre)
    return f"{ruta.stem}{marca_de_incompleto(faltan)}{ruta.suffix}"


def _redondear(numero: float, decimales: int) -> Decimal:
    """Redondeo medio hacia arriba, como el de la planilla de C.UNIX.

    Con el redondeo bancario de Python, 2,25 h daría 2,2 y el promedio de una
    persona dejaría de coincidir con el del anexo anterior.
    """
    paso = Decimal(1).scaleb(-decimales)
    return Decimal(str(numero)).quantize(paso, rounding=ROUND_HALF_UP)


def formato_horas(numero: float) -> str:
    """'306,0': un decimal con coma, como las horas del informe."""
    return f"{_redondear(numero, 1)}".replace(".", ",")


def formato_importe(numero: float) -> str:
    """'3.706,00': dos decimales, coma decimal y punto de miles."""
    texto = f"{_redondear(numero, 2):,.2f}"
    return texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
