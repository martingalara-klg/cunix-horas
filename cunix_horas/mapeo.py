"""Mapeo de códigos de Kimai a los nombres que ve el partner."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import yaml

_ALIAS = re.compile(r"^\s*\[[^\]]+\]\s*([^|]+)")

# Texto de la única fila de actividad de cada proyecto, cuando `config/mapeo.yaml`
# no lo declara. Es lo que el partner viene viendo desde siempre.
ACTIVIDAD_POR_DEFECTO = "Desarrollo"


class ErrorMapeo(Exception):
    """Falta una entrada en config/mapeo.yaml, o el archivo está mal formado."""


@dataclass(frozen=True)
class Persona:
    nombre: str
    archivo: str


@dataclass(frozen=True)
class DestinoProyecto:
    cliente: str
    proyecto: str


def _normalizar(texto: str) -> str:
    """Forma con la que se comparan los nombres para mostrar."""
    return " ".join(texto.split()).casefold()


def _alias_de_kimai(texto: str) -> str:
    """'[CO2610170] Aduana-Subastas | descripción larga' -> 'Aduana-Subastas'."""
    coincidencia = _ALIAS.match(texto)
    return coincidencia.group(1).strip() if coincidencia else texto.strip()


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
        f"  Si no está, agregá bajo personas:\n"
        f"  {clave}:\n"
        f'    nombre: "{nombre}"\n'
        f'    archivo: "AJUSTAR - apellido, va en el nombre del archivo"'
    )


def _leer_actividad(contenido: dict, ruta: Path) -> str:
    """Texto de la única fila de actividad, del YAML o el de por defecto.

    Es opcional: si no está declarado vale `Desarrollo`, que es lo que el
    partner recibió siempre. Si está pero vacío o no es texto, frena: un Excel
    con la fila de actividad en blanco o con un número adentro se vería raro
    del otro lado y nadie sabría de dónde salió.
    """
    if "actividad" not in contenido:
        return ACTIVIDAD_POR_DEFECTO

    actividad = contenido["actividad"]
    if not isinstance(actividad, str) or not actividad.strip():
        raise ErrorMapeo(
            f"'actividad:' en {ruta} tiene que ser un texto no vacío: es el "
            f"nombre de la única fila de actividad de cada proyecto en el "
            f"Excel del partner.\n"
            f"  Sacá la línea para usar el valor por defecto "
            f'("{ACTIVIDAD_POR_DEFECTO}"), o poné el texto entre comillas.'
        )
    return actividad.strip()


class Mapeo:
    """Traduce códigos de Kimai a los nombres del Excel del partner."""

    def __init__(
        self,
        personas: dict[str, Persona],
        proyectos: dict[str, DestinoProyecto],
        actividad: str = ACTIVIDAD_POR_DEFECTO,
    ) -> None:
        self._personas = personas
        self._proyectos = proyectos
        # El partner no ve cómo clasifican los desarrolladores en Kimai: cada
        # proyecto sale con una sola fila de actividad, siempre con este texto.
        self.actividad = actividad
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

        for seccion in ("personas", "proyectos"):
            if seccion not in contenido:
                raise ErrorMapeo(f"A {ruta} le falta la sección '{seccion}:'")

        personas: dict[str, Persona] = {}
        for username, datos in (contenido["personas"] or {}).items():
            for campo in ("nombre", "archivo"):
                if campo not in (datos or {}):
                    raise ErrorMapeo(
                        f"A la persona '{username}' en {ruta} le falta '{campo}:'"
                    )
            personas[username] = Persona(datos["nombre"], datos["archivo"])

        proyectos: dict[str, DestinoProyecto] = {}
        for codigo, datos in (contenido["proyectos"] or {}).items():
            for campo in ("cliente", "proyecto"):
                if campo not in (datos or {}):
                    raise ErrorMapeo(
                        f"Al proyecto '{codigo}' en {ruta} le falta '{campo}:'"
                    )
            proyectos[codigo] = DestinoProyecto(datos["cliente"], datos["proyecto"])

        return cls(personas, proyectos, _leer_actividad(contenido, ruta))

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
