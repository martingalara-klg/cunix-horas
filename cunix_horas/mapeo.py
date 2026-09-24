"""Mapeo de códigos de Kimai a los nombres que ve el partner."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import yaml

_ALIAS = re.compile(r"^\s*\[[^\]]+\]\s*([^|]+)")


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


def _alias_de_kimai(texto: str) -> str:
    """'[CO2610170] Aduana-Subastas | descripción larga' -> 'Aduana-Subastas'."""
    coincidencia = _ALIAS.match(texto)
    return coincidencia.group(1).strip() if coincidencia else texto.strip()


class Mapeo:
    """Traduce códigos de Kimai a los nombres del Excel del partner."""

    def __init__(
        self, personas: dict[str, Persona], proyectos: dict[str, DestinoProyecto]
    ) -> None:
        self._personas = personas
        self._proyectos = proyectos

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

        return cls(personas, proyectos)

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

    def resolver_persona(self, username: str, archivo: str) -> Persona:
        persona = self._personas.get(username)
        if persona is None:
            raise ErrorMapeo(
                f"Desarrollador sin mapear en {archivo}: {username}\n"
                f"  Agregá a config/mapeo.yaml, bajo personas:\n"
                f"  {username}:\n"
                f'    nombre: "AJUSTAR - nombre completo, va en A1 del Excel"\n'
                f'    archivo: "AJUSTAR - apellido, va en el nombre del archivo"'
            )
        return persona
