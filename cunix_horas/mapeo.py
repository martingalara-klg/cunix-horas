"""Mapeo de códigos de Kimai a los nombres que ve el partner."""
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

_ALIAS = re.compile(r"^\s*\[[^\]]+\]\s*([^|]+)")

# Prefijo '[codigo] ' del texto crudo de Kimai. Se saca para derivar el nombre
# que ve el partner cuando el proyecto no esta declarado en el mapeo.
_PREFIJO_CODIGO = re.compile(r"^\s*\[[^\]]+\]\s*")


class ErrorMapeo(Exception):
    """Falta una entrada en config/mapeo.yaml, o el archivo está mal formado."""


@dataclass(frozen=True)
class Persona:
    """Una persona de `personas:`, con lo que el mapeo sabe de ella.

    `mail` y `username` son el respaldo de las dos columnas que el resumen
    mensual no trae. **Sólo se usan para los registros de ese export**: los
    del reporte de detalle traen su usuario y su mail de Kimai y no miran
    acá, aunque el mapeo declare otra cosa.
    """

    nombre: str
    archivo: str
    # Opcional. Vacío mientras nadie lo declare: es lo que hace frenar un
    # archivo de resumen mensual, con el bloque YAML para completarlo.
    mail: str = ""
    # La clave de `personas:`, que ES el `User` que ve el partner. Viaja
    # adentro de la Persona para que quien la resuelve por `nombre` no tenga
    # que volver a buscar cuál era su clave.
    username: str = ""


@dataclass(frozen=True)
class DestinoProyecto:
    cliente: str
    proyecto: str
    # Opcional, y es el `Project number` que ve el partner. Mismo respaldo que
    # `Persona.mail`: sólo lo usan los registros del resumen mensual.
    numero_proyecto: str = ""


def _normalizar(texto: str) -> str:
    """Forma con la que se comparan los nombres para mostrar."""
    return " ".join(texto.split()).casefold()


def _sin_codigo(texto: str) -> str:
    """Le saca el prefijo '[código] ' al texto crudo de Kimai, si lo tiene."""
    return _PREFIJO_CODIGO.sub("", texto or "").strip()


def derivar_cliente(texto_kimai: str) -> str:
    """El nombre de cliente que ve el partner cuando el proyecto no está mapeado.

    '[616050001] Instituto de Salud Pública' -> 'Instituto de Salud Pública'
    'CUNIX'                                  -> 'CUNIX'

    El segundo caso no es teórico: el cliente de uno de los desarrolladores
    viene sin corchetes. Si la derivación asumiera el prefijo, ese nombre
    saldría vacío y el partner recibiría una columna Customer en blanco.
    """
    return _sin_codigo(texto_kimai)


def derivar_proyecto(texto_kimai: str) -> str:
    """El nombre de proyecto que ve el partner cuando no está mapeado.

    '[AD2690002] C.UNIX - Internos | C.UNIX - VictoriusCP2' -> 'C.UNIX - Internos'

    Lo que va después del '|' es la descripción larga de Kimai, que no entra
    en la columna Project.
    """
    return _sin_codigo(texto_kimai).split("|")[0].strip()


def _alias_de_kimai(texto: str) -> str:
    """'[CO2610170] Aduana-Subastas | descripción larga' -> 'Aduana-Subastas'."""
    coincidencia = _ALIAS.match(texto)
    return coincidencia.group(1).strip() if coincidencia else texto.strip()


def _texto_opcional(valor, campo: str, entidad: str, ruta: Path) -> str:
    """Un campo opcional del YAML: ausente vale vacío, presente tiene que ser texto.

    No se acepta declarado y en blanco. Ese campo termina en una columna que
    el partner factura: si está escrito a medias conviene que frene ahora, con
    el nombre de la entrada, y no que salga vacío del otro lado.
    """
    if valor is None:
        return ""
    if not isinstance(valor, str) or not valor.strip():
        raise ErrorMapeo(
            f"'{campo}:' de {entidad} en {ruta} está declarado pero vacío.\n"
            f"  Ese campo va a una columna que el partner factura: o lo "
            f"completás entre comillas, o borrás la línea."
        )
    return valor.strip()


def _sugerencia_de_persona(identificador: str, archivo: str) -> str:
    """Mensaje de persona sin mapear, con el YAML listo para pegar.

    Lo que llega puede ser un username (timesheet) o un nombre para mostrar
    (resumen mensual), así que el bloque sugerido cambia según cuál sea: un
    nombre con espacios no puede ser la clave, que es siempre el username.
    """
    es_nombre = " " in identificador.strip()
    clave = "AJUSTAR-username-de-kimai" if es_nombre else identificador
    nombre = (
        identificador.strip()
        if es_nombre
        else "AJUSTAR - nombre completo, va en A1 del Excel"
    )
    return (
        f"Desarrollador sin mapear en {archivo}: {identificador}\n"
        f"  Así lo identifica este export: el timesheet trae el username de "
        f"Kimai, y el resumen mensual trae el nombre para mostrar. "
        f"config/mapeo.yaml resuelve las dos cosas, pero ninguna persona "
        f"coincide con eso.\n"
        f"  Si la persona ya está cargada, hacé que su username o su "
        f"'nombre:' coincidan exactamente con lo de arriba.\n"
        f"  La clave de la entrada ES el 'User' que ve el partner, y el "
        f"'mail:' es su columna E-mail: los dos hacen falta si esta persona "
        f"exportó con el reporte de resumen mensual, que no los trae. Con el "
        f"reporte de detalle salen de Kimai y el mapeo no interviene.\n"
        f"  Si no está, agregá bajo personas:\n"
        f"  {clave}:\n"
        f'    nombre: "{nombre}"\n'
        f'    archivo: "AJUSTAR - apellido, va en el nombre del archivo"\n'
        f'    mail: "AJUSTAR - el mail de Kimai de esta persona"'
    )


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

        # 'personas:' es opcional para quien exporta con el reporte de detalle:
        # ese export trae el nombre y el usuario de cada desarrollador. Sigue
        # haciendo falta para quien exporta con el resumen mensual.
        if "proyectos" not in contenido:
            raise ErrorMapeo(f"A {ruta} le falta la sección 'proyectos:'")

        personas: dict[str, Persona] = {}
        for username, datos in (contenido.get("personas") or {}).items():
            for campo in ("nombre", "archivo"):
                if campo not in (datos or {}):
                    raise ErrorMapeo(
                        f"A la persona '{username}' en {ruta} le falta '{campo}:'"
                    )
            personas[username] = Persona(
                datos["nombre"],
                datos["archivo"],
                _texto_opcional(
                    datos.get("mail"), "mail", f"la persona '{username}'", ruta
                ),
                username,
            )

        proyectos: dict[str, DestinoProyecto] = {}
        for codigo, datos in (contenido["proyectos"] or {}).items():
            for campo in ("cliente", "proyecto"):
                if campo not in (datos or {}):
                    raise ErrorMapeo(
                        f"Al proyecto '{codigo}' en {ruta} le falta '{campo}:'"
                    )
            proyectos[codigo] = DestinoProyecto(
                datos["cliente"],
                datos["proyecto"],
                _texto_opcional(
                    datos.get("numero_proyecto"),
                    "numero_proyecto",
                    f"el proyecto '{codigo}'",
                    ruta,
                ),
            )

        return cls(
            personas,
            proyectos,
            _leer_anexos(contenido, ruta),
            _leer_fuentes_manuales(contenido, ruta),
        )

    def proyecto_opcional(self, codigo: str) -> DestinoProyecto | None:
        """El destino declarado para ese código, o None si no está declarado.

        El detalle plano usa ésta y no `resolver_proyecto`: un proyecto sin
        mapear ya no frena nada, sale con el nombre derivado de Kimai y se
        lista en el informe. Cada fila lleva su `Project number`, así que la
        trazabilidad no depende del mapeo.
        """
        return self._proyectos.get(codigo)

    def resolver_proyecto(
        self, codigo: str, texto_kimai: str, archivo: str
    ) -> DestinoProyecto:
        destino = self._proyectos.get(codigo)
        if destino is None:
            alias = _alias_de_kimai(texto_kimai) or codigo
            raise ErrorMapeo(
                f"Proyecto sin mapear en {archivo}: [{codigo}] {alias}\n"
                f"  Agregá a config/mapeo.yaml, bajo proyectos:\n"
                f"  {codigo}:\n"
                f'    cliente: "AJUSTAR - nombre del cliente para el partner"\n'
                f'    proyecto: "{alias}"'
            )
        return destino

    def resolver_persona(self, identificador: str, archivo: str) -> Persona:
        """La persona de `personas:`, buscada por username y si no por nombre.

        El timesheet trae el username de Kimai; el resumen mensual trae sólo el
        nombre para mostrar. Se busca primero por username, que es la clave del
        mapeo, y recién si no aparece se busca por el 'nombre:' ya configurado.
        """
        persona = self._personas.get(identificador)
        if persona is not None:
            return persona

        usernames = self._por_nombre.get(_normalizar(identificador), [])
        if len(usernames) > 1:
            raise ErrorMapeo(
                f"Nombre ambiguo en {archivo}: {identificador}\n"
                f"  En config/mapeo.yaml hay más de una persona con ese "
                f"'nombre:' ({', '.join(sorted(usernames))}).\n"
                f"  Este export es un resumen mensual: no trae el username, "
                f"sólo el nombre, así que no hay forma de saber cuál de las "
                f"dos es sin inventar.\n"
                f"  Cambiá el 'nombre:' de una de ellas para que no se repita."
            )
        if usernames:
            return self._personas[usernames[0]]

        raise ErrorMapeo(_sugerencia_de_persona(identificador, archivo))
