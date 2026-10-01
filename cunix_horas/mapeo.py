"""Mapeo de códigos de Kimai a los nombres y usuarios que ve C.UNIX.

Guarda sólo lo que alguno de los dos anexos escribe: el nombre de cliente y de
proyecto de cada código, el nombre y el usuario de Kimai de cada persona, los
nombres de los dos archivos y los textos del informe, y las planillas
manuales del mes.

Lo que el entregable no escribe no se declara acá. Los anexos no tienen
columna de mail ni de número de proyecto, así que `mail:` y
`numero_proyecto:` dejaron de existir: eran dos datos que podían trabar una
entrega sin llegar nunca a ningún archivo.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import yaml

from cunix_horas.anexos import (
    ConfigAnexos,
    DescripcionesAparte,
    FuentesManuales,
    HorasSinKimai,
    PATRON_DETALLE,
    PATRON_INFORME,
    PERFIL_POR_DEFECTO,
)

# Prefijo '[codigo] ' del texto crudo de Kimai. Se saca para derivar el nombre
# que ve el partner cuando el proyecto no esta declarado en el mapeo.
_PREFIJO_CODIGO = re.compile(r"^\s*\[[^\]]+\]\s*")


class ErrorMapeo(Exception):
    """Falta una entrada en config/mapeo.yaml, o el archivo está mal formado."""


@dataclass(frozen=True)
class Persona:
    """Una persona de `personas:`, con lo que el mapeo sabe de ella.

    `username` es el respaldo de la única columna que el resumen mensual no
    trae y que los anexos sí escriben: «Usuario Kimai», en la hoja `Datos`.
    **Sólo se usa para los registros de ese export**: los del reporte de
    detalle traen su usuario de Kimai y no miran acá.

    No hay `mail:`. Ninguno de los dos anexos tiene columna de mail, así que
    el mapeo no guarda un dato que después no se escribe en ningún lado.
    Tampoco hay `archivo:` (el apellido del Excel por desarrollador del
    entregable anterior): si un `mapeo.yaml` viejo lo trae, se ignora.
    """

    nombre: str
    # La clave de `personas:`, que ES el usuario de Kimai que ve C.UNIX en la
    # hoja `Datos`. Viaja adentro de la Persona para que quien la resuelve por
    # `nombre` no tenga que volver a buscar cuál era su clave.
    username: str = ""


@dataclass(frozen=True)
class DestinoProyecto:
    """Los dos nombres que ve C.UNIX para un código de proyecto de Kimai.

    No hay `numero_proyecto:` por el mismo motivo que no hay `mail:`: el
    Anexo II-A no tiene columna de número de proyecto.
    """

    cliente: str
    proyecto: str


def _normalizar(texto: str) -> str:
    """Forma con la que se comparan los nombres para mostrar."""
    return " ".join(texto.split()).casefold()


def _sin_codigo(texto: str) -> str:
    """Le saca el prefijo '[código] ' al texto crudo de Kimai, si lo tiene."""
    return _PREFIJO_CODIGO.sub("", texto or "").strip()


def derivar_proyecto(texto_kimai: str) -> str:
    """El nombre de proyecto que ve el partner cuando no está mapeado.

    '[AD2690002] C.UNIX - Internos | C.UNIX - VictoriusCP2' -> 'C.UNIX - Internos'

    Lo que va después del '|' es la descripción larga de Kimai, que no entra
    en la columna Project.
    """
    return _sin_codigo(texto_kimai).split("|")[0].strip()


def _texto_opcional(valor, campo: str, entidad: str, ruta: Path) -> str:
    """Un campo opcional del YAML: ausente vale vacío, presente tiene que ser texto.

    No se acepta declarado y en blanco. Lo que queda escrito a medias conviene
    que frene ahora, con el nombre de la entrada, y no que salga vacío en un
    anexo que ya está camino a C.UNIX.
    """
    if valor is None:
        return ""
    if not isinstance(valor, str) or not valor.strip():
        raise ErrorMapeo(
            f"'{campo}:' de {entidad} en {ruta} está declarado pero vacío.\n"
            f"  Ese campo sale escrito en un anexo: o lo completás entre "
            f"comillas, o borrás la línea."
        )
    return valor.strip()


def _patron(
    contenido: dict, clave: str, por_defecto: str, extension: str, ruta: Path
) -> str:
    """Un patrón de nombre de archivo de `anexos:`, validado contra `{periodo}`.

    Se valida acá y no al escribir: si el patrón está mal, el error tiene que
    salir antes de leer ningún export y no después de haber procesado todo.
    """
    valor = (contenido or {}).get(clave)
    if valor is None:
        return por_defecto

    ayuda = (
        f"  Sacá la línea para usar el valor por defecto "
        f'("{por_defecto}"), o escribilo entre comillas usando {{periodo}}, '
        f"que vale el mes en formato AAAA-MM."
    )
    if not isinstance(valor, str) or not valor.strip():
        raise ErrorMapeo(
            f"'{clave}:' bajo 'anexos:' en {ruta} tiene que ser un texto no "
            f"vacío: es el nombre de un archivo que recibe C.UNIX.\n" + ayuda
        )
    patron = valor.strip()
    try:
        prueba = patron.format(periodo="2026-09")
    except (KeyError, IndexError, ValueError):
        raise ErrorMapeo(
            f"'{clave}:' bajo 'anexos:' en {ruta} usa algo que no se "
            f"entiende: {patron!r}.\n"
            f"  El único reemplazo que existe es {{periodo}}.\n" + ayuda
        ) from None
    if not prueba.lower().endswith(extension):
        raise ErrorMapeo(
            f"'{clave}:' bajo 'anexos:' en {ruta} tiene que terminar en "
            f"{extension}: con {patron!r} el archivo se llamaría {prueba!r} y "
            f"no se abriría.\n" + ayuda
        )
    return patron


def _leer_anexos(contenido: dict, ruta: Path) -> ConfigAnexos:
    """La sección `anexos:`, toda opcional."""
    datos = contenido.get("anexos") or {}
    if not isinstance(datos, dict):
        raise ErrorMapeo(
            f"'anexos:' en {ruta} tiene que ser un bloque de opciones, no "
            f"{type(datos).__name__}."
        )
    perfiles = datos.get("perfiles") or {}
    if not isinstance(perfiles, dict):
        raise ErrorMapeo(
            f"'perfiles:' bajo 'anexos:' en {ruta} tiene que ser una lista de "
            f'"Nombre Apellido": "Perfil".'
        )
    return ConfigAnexos(
        archivo_detalle=_patron(
            datos, "archivo_detalle", PATRON_DETALLE, ".xlsx", ruta
        ),
        archivo_informe=_patron(
            datos, "archivo_informe", PATRON_INFORME, ".docx", ruta
        ),
        perfil_por_defecto=_texto_opcional(
            datos.get("perfil_por_defecto"), "perfil_por_defecto", "'anexos:'", ruta
        )
        or PERFIL_POR_DEFECTO,
        perfiles=tuple(
            (str(persona), str(perfil)) for persona, perfil in sorted(perfiles.items())
        ),
        # Los tres de abajo quedan como marcador en el documento si no están.
        # Se aceptan números, para poder escribir `dias_habiles_entrega: 5`.
        dias_habiles_entrega=str(datos.get("dias_habiles_entrega") or "").strip(),
        contrato_de_fecha=str(datos.get("contrato_de_fecha") or "").strip(),
        fecha_de_emision=str(datos.get("fecha_de_emision") or "").strip(),
    )


def _entradas(datos, clave: str, ruta: Path) -> list:
    lista = (datos or {}).get(clave) or []
    if not isinstance(lista, list):
        raise ErrorMapeo(
            f"'{clave}:' bajo 'fuentes_manuales:' en {ruta} tiene que ser una "
            f"lista de entradas, cada una empezando con '- '."
        )
    return lista


def _campo(entrada, campo: str, clave: str, ruta: Path) -> str:
    valor = (entrada or {}).get(campo) if isinstance(entrada, dict) else None
    if not isinstance(valor, str) or not valor.strip():
        raise ErrorMapeo(
            f"A una entrada de '{clave}:' bajo 'fuentes_manuales:' en {ruta} "
            f"le falta '{campo}:'.\n"
            f"  Cada entrada tiene que decir, como mínimo, de quién son las "
            f"horas y en qué planilla de input/<mes>/manual/ están."
        )
    return valor.strip()


def _leer_fuentes_manuales(contenido: dict, ruta: Path) -> FuentesManuales:
    """La sección `fuentes_manuales:`, que es opcional entera.

    Un mes en que todos cargan sus horas y sus descripciones en Kimai no
    declara nada acá y corre igual.
    """
    datos = contenido.get("fuentes_manuales") or {}
    if not isinstance(datos, dict):
        raise ErrorMapeo(
            f"'fuentes_manuales:' en {ruta} tiene que ser un bloque con "
            f"'horas_sin_kimai:' y/o 'descripciones:'."
        )

    horas = tuple(
        HorasSinKimai(
            persona=_campo(entrada, "persona", "horas_sin_kimai", ruta),
            planilla=_campo(entrada, "planilla", "horas_sin_kimai", ruta),
            restar_a=_campo(entrada, "restar_a", "horas_sin_kimai", ruta),
            proyecto=str(entrada.get("proyecto") or "").strip(),
            nota=str(entrada.get("nota") or "").strip(),
        )
        for entrada in _entradas(datos, "horas_sin_kimai", ruta)
    )
    descripciones = tuple(
        DescripcionesAparte(
            persona=_campo(entrada, "persona", "descripciones", ruta),
            planilla=_campo(entrada, "planilla", "descripciones", ruta),
        )
        for entrada in _entradas(datos, "descripciones", ruta)
    )
    return FuentesManuales(horas, descripciones)


class Mapeo:
    """Traduce códigos de Kimai a los nombres que ve C.UNIX en los anexos."""

    def __init__(
        self,
        personas: dict[str, Persona],
        proyectos: dict[str, DestinoProyecto],
        anexos: ConfigAnexos = ConfigAnexos(),
        fuentes_manuales: FuentesManuales = FuentesManuales(),
    ) -> None:
        self._personas = personas
        self._proyectos = proyectos
        # Nombres de los dos anexos, perfiles y los textos del informe que el
        # dueño configura.
        self.anexos = anexos
        # Las planillas de input/<mes>/manual/ que este mes necesita.
        self.fuentes_manuales = fuentes_manuales
        # El resumen mensual de Kimai no trae el username, sólo el nombre para
        # mostrar. Este índice permite resolver la persona también por ahí, sin
        # agregar configuración nueva: el 'nombre:' ya está en cada entrada.
        self._por_nombre: dict[str, list[str]] = {}
        for username, persona in personas.items():
            self._por_nombre.setdefault(_normalizar(persona.nombre), []).append(
                username
            )

    @classmethod
    def cargar(cls, ruta: Path) -> "Mapeo":
        if not ruta.is_file():
            raise ErrorMapeo(f"No se encontró el archivo de configuración: {ruta}")

        contenido = yaml.safe_load(ruta.read_text(encoding="utf-8")) or {}

        # 'personas:' es opcional entera: quien exporta con el reporte de
        # detalle trae su nombre y su usuario de Kimai en cada fila, y a quien
        # exporta con el resumen mensual le falta sólo el usuario, que si no
        # está declarado queda vacío con un aviso. Nadie frena por esto.
        if "proyectos" not in contenido:
            raise ErrorMapeo(f"A {ruta} le falta la sección 'proyectos:'")

        personas: dict[str, Persona] = {}
        for username, datos in (contenido.get("personas") or {}).items():
            if "nombre" not in (datos or {}):
                raise ErrorMapeo(
                    f"A la persona '{username}' en {ruta} le falta 'nombre:'"
                )
            personas[username] = Persona(datos["nombre"], username)

        proyectos: dict[str, DestinoProyecto] = {}
        for codigo, datos in (contenido["proyectos"] or {}).items():
            for campo in ("cliente", "proyecto"):
                if campo not in (datos or {}):
                    raise ErrorMapeo(
                        f"Al proyecto '{codigo}' en {ruta} le falta '{campo}:'"
                    )
            proyectos[codigo] = DestinoProyecto(datos["cliente"], datos["proyecto"])

        return cls(
            personas,
            proyectos,
            _leer_anexos(contenido, ruta),
            _leer_fuentes_manuales(contenido, ruta),
        )

    def proyecto_opcional(self, codigo: str) -> DestinoProyecto | None:
        """El destino declarado para ese código, o None si no está declarado.

        Un proyecto sin mapear no frena nada: sale con el nombre derivado del
        texto de Kimai y queda listado en `_validacion.txt` con el bloque YAML
        listo para pegar. Por eso no existe una versión que falle.
        """
        return self._proyectos.get(codigo)

    def persona_opcional(self, identificador: str) -> Persona | None:
        """La persona de `personas:`, buscada por username y si no por nombre.

        El timesheet trae el username de Kimai; el resumen mensual trae sólo el
        nombre para mostrar. Se busca primero por username, que es la clave del
        mapeo, y recién si no aparece se busca por el 'nombre:' ya configurado.

        Devuelve `None` cuando no hay **una sola** persona que corresponda: ni
        la que no está declarada ni un nombre repetido en dos entradas frenan
        nada. Lo único que el mapeo aporta acá es el usuario de Kimai de la
        hoja `Datos`, y una celda vacía ahí es un hecho legítimo que C.UNIX ya
        conoce. Quien llama avisa; no se inventa un usuario.
        """
        persona = self._personas.get(identificador)
        if persona is not None:
            return persona

        usernames = self._por_nombre.get(_normalizar(identificador), [])
        if len(usernames) == 1:
            return self._personas[usernames[0]]
        return None

    def usernames_con_el_nombre(self, identificador: str) -> tuple[str, ...]:
        """Los usernames declarados con ese 'nombre:'. Vacío si ninguno.

        Más de uno significa que el nombre está repetido en `personas:` y que
        no hay forma de saber cuál es sin inventar. Lo usa el aviso, para
        poder nombrarlos.
        """
        return tuple(sorted(self._por_nombre.get(_normalizar(identificador), [])))
