# cunix-horas — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convertir los exports mensuales `.xlsx` de Kimai (uno por desarrollador) en los Excel con el formato que recibe el partner, uno por desarrollador, ejecutando un solo comando.

**Architecture:** Pipeline de cinco etapas puras y desacopladas — `lector_kimai` → `mapeo` → `agregador` → `validador` → `escritor_excel` — orquestadas por `cli`. Cada etapa recibe y devuelve estructuras inmutables. El input se parsea con XML crudo (openpyxl no puede abrir los archivos de Kimai); el output se escribe con openpyxl tomando los estilos de una plantilla.

**Tech Stack:** Python 3.14, openpyxl 3.1.5, PyYAML, pytest. Sin otras dependencias.

**Spec:** `docs/superpowers/specs/2026-09-24-cunix-horas-design.md`

## Global Constraints

- El paquete vive en `cunix_horas/` en la raíz del proyecto, **no** bajo `src/`. Motivo: `python -m cunix_horas` debe funcionar sin `pip install -e .`, porque se invoca desde un `.bat` de doble clic.
- **Nunca** usar `openpyxl.load_workbook()` sobre un archivo de input de Kimai. Falla con `TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'`. El input se lee con `zipfile` + `xml.etree.ElementTree`.
- Namespace OOXML en todo el parseo: `{http://schemas.openxmlformats.org/spreadsheetml/2006/main}`.
- Época de los seriales de fecha de Excel: `date(1899, 12, 30)`. `46265.708333333` → `2026-08-31`.
- La duración de Kimai es fracción de día: horas = `Duration * 24`.
- Las horas se redondean a 2 decimales **sólo al escribir la celda**, nunca durante la acumulación.
- Todos los `dataclass` son `frozen=True`. Ninguna función muta sus argumentos.
- Todo identificador de código, mensaje de usuario y nombre de archivo del proyecto va en **español**.
- Los archivos de texto se leen y escriben siempre con `encoding="utf-8"` explícito (los nombres de cliente traen acentos: `Subsecretaría`, `Pública`).
- Nombres de mes para archivos de salida: tabla fija en inglés abreviado `{1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun", 7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec"}`. **Nunca** `strftime("%b")`: depende del locale de la máquina.
- Encabezados de día en el Excel de salida: **texto**, formato `M/D/YYYY` sin ceros a la izquierda (`10/1/2025`). Nunca objetos `date`.

## Estructura de archivos

| Archivo | Responsabilidad |
|---------|-----------------|
| `cunix_horas/__init__.py` | Vacío. Marca el paquete. |
| `cunix_horas/__main__.py` | Punto de entrada de `python -m cunix_horas`. Sólo llama a `cli.main()`. |
| `cunix_horas/lector_kimai.py` | `.xlsx` de Kimai → `list[Registro]`. Parseo XML, seriales de fecha, extracción del código de proyecto. |
| `cunix_horas/mapeo.py` | Carga `config/mapeo.yaml`. Resuelve código de proyecto → (cliente, proyecto) y username → (nombre, archivo). Lanza error con sugerencia YAML ante claves faltantes. |
| `cunix_horas/agregador.py` | `list[Registro]` + `Mapeo` → `Reporte` con jerarquía Cliente→Proyecto→Actividad y totales. |
| `cunix_horas/validador.py` | `Reporte` → `list[str]` de avisos. |
| `cunix_horas/escritor_excel.py` | `Reporte` + plantilla → archivo `.xlsx`. |
| `cunix_horas/cli.py` | Orquesta: recorre `input/<mes>/`, procesa cada archivo de forma aislada, escribe outputs y `_validacion.txt`, imprime resumen. |
| `config/mapeo.yaml` | Configuración editada por el usuario. |
| `templates/plantilla.xlsx` | Fuente de estilos (copia de `Oct Dodera.xlsx`). |
| `tests/conftest.py` | Pone la raíz del proyecto en `sys.path` y expone la ruta de fixtures. |
| `tests/fixtures/` | `kimai-mzalazar.xlsx`, `plantilla.xlsx`, `mapeo-test.yaml`. |
| `tests/test_*.py` | Un archivo de test por módulo. |
| `generar.bat` | Doble clic → ejecuta el mes que se le pase o pide uno. |
| `README.md` | Cómo usarlo, para vos dentro de seis meses. |

---

### Task 1: Esqueleto del proyecto, git y fixtures

**Files:**
- Create: `.gitignore`
- Create: `requirements.txt`
- Create: `cunix_horas/__init__.py`
- Create: `tests/conftest.py`
- Create: `templates/plantilla.xlsx` (copia de `Oct Dodera.xlsx`)
- Create: `tests/fixtures/kimai-mzalazar.xlsx` (copia de `20260924-kimai-export.xlsx`)
- Create: `tests/fixtures/plantilla.xlsx` (copia de `Oct Dodera.xlsx`)
- Create: `tests/test_esqueleto.py`

**Interfaces:**
- Consumes: nada.
- Produces: el paquete `cunix_horas` importable desde los tests; las constantes `RAIZ` y `FIXTURES` en `tests/conftest.py`.

- [ ] **Step 1: Inicializar git**

```bash
cd "C:/Users/Martin/Desktop/WORK/CUNIX/cunix-horas"
git init
git add docs/
git commit -m "docs: agregar spec y plan de cunix-horas"
```

- [ ] **Step 2: Crear `.gitignore`**

```gitignore
__pycache__/
*.pyc
.pytest_cache/
.venv/
venv/
output/
~$*.xlsx
```

Nota: `output/` se ignora a propósito — son archivos regenerables. `input/` **sí** se versiona: es la única copia de lo que se exportó de Kimai ese mes.

- [ ] **Step 3: Crear `requirements.txt`**

```
openpyxl>=3.1.5
PyYAML>=6.0
pytest>=8.0
```

- [ ] **Step 4: Instalar dependencias**

```bash
python -m pip install -r requirements.txt
```

Esperado: openpyxl ya está (3.1.5); instala PyYAML y pytest.

- [ ] **Step 5: Crear la estructura de carpetas y copiar los archivos originales**

```bash
cd "C:/Users/Martin/Desktop/WORK/CUNIX/cunix-horas"
mkdir -p cunix_horas config templates tests/fixtures input output
touch cunix_horas/__init__.py
cp "Oct Dodera.xlsx" templates/plantilla.xlsx
cp "Oct Dodera.xlsx" tests/fixtures/plantilla.xlsx
cp "20260924-kimai-export.xlsx" tests/fixtures/kimai-mzalazar.xlsx
```

Los dos archivos originales quedan en la raíz por ahora; se borran en la Task 7, una vez que todo funciona.

- [ ] **Step 6: Crear `tests/conftest.py`**

```python
"""Configuración compartida de los tests."""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).resolve().parent / "fixtures"

if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))
```

- [ ] **Step 7: Escribir el test del esqueleto**

`tests/test_esqueleto.py`:

```python
from conftest import FIXTURES, RAIZ


def test_el_paquete_es_importable():
    import cunix_horas  # noqa: F401


def test_los_fixtures_existen():
    assert (FIXTURES / "kimai-mzalazar.xlsx").is_file()
    assert (FIXTURES / "plantilla.xlsx").is_file()


def test_la_plantilla_de_produccion_existe():
    assert (RAIZ / "templates" / "plantilla.xlsx").is_file()
```

- [ ] **Step 8: Correr los tests**

Run: `python -m pytest tests/test_esqueleto.py -v`
Expected: 3 passed

- [ ] **Step 9: Commit**

```bash
git add .gitignore requirements.txt cunix_horas/ templates/ tests/
git commit -m "chore: esqueleto del proyecto, dependencias y fixtures"
```

---

### Task 2: `lector_kimai` — leer el export de Kimai

**Files:**
- Create: `cunix_horas/lector_kimai.py`
- Test: `tests/test_lector_kimai.py`

**Interfaces:**
- Consumes: nada.
- Produces:
  - `@dataclass(frozen=True) Registro(fecha: date, horas: float, username: str, cod_proyecto: str, actividad: str)`
  - `class ErrorLectura(Exception)`
  - `leer(ruta: Path) -> list[Registro]`
  - `serial_a_fecha(serial: str | float) -> date`
  - `codigo_de_proyecto(texto: str) -> str`

- [ ] **Step 1: Escribir los tests que fallan**

`tests/test_lector_kimai.py`:

```python
from datetime import date

import pytest
from conftest import FIXTURES

from cunix_horas.lector_kimai import (
    ErrorLectura,
    Registro,
    codigo_de_proyecto,
    leer,
    serial_a_fecha,
)


def test_serial_a_fecha_convierte_el_serial_de_excel():
    assert serial_a_fecha("46265.708333333") == date(2026, 8, 31)
    assert serial_a_fecha(46243.333333333) == date(2026, 8, 9)


def test_codigo_de_proyecto_extrae_lo_que_esta_entre_corchetes():
    texto = "[CO2610170] Aduana-Subastas | Servicio Nacional de Aduanas - Soporte"
    assert codigo_de_proyecto(texto) == "CO2610170"


def test_codigo_de_proyecto_falla_si_no_hay_corchetes():
    with pytest.raises(ErrorLectura, match="sin código"):
        codigo_de_proyecto("Aduana-Subastas")


def test_leer_devuelve_los_24_registros_del_fixture():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert len(registros) == 24
    assert all(isinstance(r, Registro) for r in registros)


def test_leer_suma_las_horas_correctas():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert round(sum(r.horas for r in registros), 2) == 76.5


def test_leer_cubre_el_rango_de_fechas_esperado():
    fechas = [r.fecha for r in leer(FIXTURES / "kimai-mzalazar.xlsx")]
    assert min(fechas) == date(2026, 8, 3)
    assert max(fechas) == date(2026, 8, 31)


def test_leer_extrae_username_actividad_y_codigo():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    assert {r.username for r in registros} == {"mzalazar"}
    assert {r.actividad for r in registros} == {"Desarrollo"}
    assert {r.cod_proyecto for r in registros} == {
        "CO2610170",
        "CO2510115",
        "PR2510126",
    }


def test_leer_agrupa_las_horas_por_codigo_de_proyecto():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    por_codigo = {}
    for r in registros:
        por_codigo[r.cod_proyecto] = por_codigo.get(r.cod_proyecto, 0.0) + r.horas
    assert round(por_codigo["CO2610170"], 2) == 61.0
    assert round(por_codigo["CO2510115"], 2) == 10.0
    assert round(por_codigo["PR2510126"], 2) == 5.5


def test_leer_falla_con_un_archivo_que_no_es_xlsx(tmp_path):
    falso = tmp_path / "roto.xlsx"
    falso.write_text("esto no es un xlsx", encoding="utf-8")
    with pytest.raises(ErrorLectura, match="no es un archivo"):
        leer(falso)


def test_leer_falla_si_faltan_las_columnas_esperadas(tmp_path):
    import zipfile

    ruta = tmp_path / "sin-columnas.xlsx"
    hoja = (
        '<?xml version="1.0"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1">'
        '<c r="A1" t="inlineStr"><is><t>Otra cosa</t></is></c>'
        "</row></sheetData></worksheet>"
    )
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("xl/worksheets/sheet1.xml", hoja)
    with pytest.raises(ErrorLectura, match="no tiene el formato"):
        leer(ruta)
```

- [ ] **Step 2: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_lector_kimai.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'cunix_horas.lector_kimai'`

- [ ] **Step 3: Escribir `cunix_horas/lector_kimai.py`**

```python
"""Lectura de los exports .xlsx de Kimai.

No se usa openpyxl: Kimai emite el atributo `showZeroes` donde el esquema
OOXML define `showZeros`, y openpyxl.load_workbook() explota con
`TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'`.
Se parsea el XML del .xlsx directamente.
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
EPOCA_EXCEL = date(1899, 12, 30)

COL_FECHA = "A"
COL_DURACION = "D"
COL_USERNAME = "F"
COL_PROYECTO = "J"
COL_ACTIVIDAD = "K"

ENCABEZADOS_ESPERADOS = {
    COL_FECHA: "Date",
    COL_DURACION: "Duration",
    COL_USERNAME: "User",
    COL_PROYECTO: "Project",
    COL_ACTIVIDAD: "Activity",
}

_CODIGO = re.compile(r"^\s*\[([^\]]+)\]")
_SOLO_LETRAS = re.compile(r"[A-Z]+")


class ErrorLectura(Exception):
    """El archivo de input no se pudo leer o no tiene el formato esperado."""


@dataclass(frozen=True)
class Registro:
    """Un registro de tiempo individual de Kimai."""

    fecha: date
    horas: float
    username: str
    cod_proyecto: str
    actividad: str


def serial_a_fecha(serial: str | float) -> date:
    """Convierte un serial de fecha de Excel a date. 46265.708333333 -> 2026-08-31."""
    return EPOCA_EXCEL + timedelta(days=float(serial))


def codigo_de_proyecto(texto: str) -> str:
    """Extrae el código entre corchetes del campo Project de Kimai.

    '[CO2610170] Aduana-Subastas | ...' -> 'CO2610170'
    """
    coincidencia = _CODIGO.match(texto)
    if coincidencia is None:
        raise ErrorLectura(f"Proyecto sin código entre corchetes: {texto!r}")
    return coincidencia.group(1)


def _letra_de_columna(referencia: str) -> str:
    """'AG12' -> 'AG'."""
    coincidencia = _SOLO_LETRAS.match(referencia)
    return coincidencia.group(0) if coincidencia else ""


def _cadenas_compartidas(archivo: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archivo.namelist():
        return []
    raiz = ET.fromstring(archivo.read("xl/sharedStrings.xml"))
    return [
        "".join(t.text or "" for t in si.iter(f"{NS}t"))
        for si in raiz.findall(f"{NS}si")
    ]


def _filas(archivo: zipfile.ZipFile) -> list[dict[str, str]]:
    """Devuelve cada fila como {letra_de_columna: valor}, saltando las vacías."""
    hojas = [n for n in archivo.namelist() if n.startswith("xl/worksheets/sheet")]
    if not hojas:
        raise ErrorLectura("El archivo no tiene ninguna hoja de cálculo")
    compartidas = _cadenas_compartidas(archivo)
    raiz = ET.fromstring(archivo.read(sorted(hojas)[0]))

    filas: list[dict[str, str]] = []
    for elemento in raiz.iter(f"{NS}row"):
        fila: dict[str, str] = {}
        for celda in elemento:
            columna = _letra_de_columna(celda.get("r", ""))
            if not columna:
                continue
            valor_numerico = celda.find(f"{NS}v")
            valor_inline = celda.find(f"{NS}is")
            if valor_numerico is not None:
                if celda.get("t") == "s":
                    fila[columna] = compartidas[int(valor_numerico.text)]
                else:
                    fila[columna] = valor_numerico.text or ""
            elif valor_inline is not None:
                fila[columna] = "".join(
                    t.text or "" for t in valor_inline.iter(f"{NS}t")
                )
        if fila:
            filas.append(fila)
    return filas


def _verificar_encabezados(encabezado: dict[str, str], ruta: Path) -> None:
    faltantes = [
        f"{col}={esperado!r}"
        for col, esperado in ENCABEZADOS_ESPERADOS.items()
        if encabezado.get(col) != esperado
    ]
    if faltantes:
        raise ErrorLectura(
            f"{ruta.name} no tiene el formato de export de Kimai. "
            f"Se esperaba en la fila 1: {', '.join(faltantes)}"
        )


def leer(ruta: Path) -> list[Registro]:
    """Lee un export de Kimai y devuelve sus registros de tiempo."""
    if not zipfile.is_zipfile(ruta):
        raise ErrorLectura(f"{ruta.name} no es un archivo .xlsx válido")

    with zipfile.ZipFile(ruta) as archivo:
        filas = _filas(archivo)

    if not filas:
        raise ErrorLectura(f"{ruta.name} está vacío")

    _verificar_encabezados(filas[0], ruta)

    registros: list[Registro] = []
    for fila in filas[1:]:
        if COL_FECHA not in fila:
            continue
        registros.append(
            Registro(
                fecha=serial_a_fecha(fila[COL_FECHA]),
                horas=float(fila.get(COL_DURACION, 0)) * 24,
                username=fila.get(COL_USERNAME, ""),
                cod_proyecto=codigo_de_proyecto(fila.get(COL_PROYECTO, "")),
                actividad=fila.get(COL_ACTIVIDAD, ""),
            )
        )
    return registros
```

- [ ] **Step 4: Correr los tests**

Run: `python -m pytest tests/test_lector_kimai.py -v`
Expected: 10 passed

- [ ] **Step 5: Commit**

```bash
git add cunix_horas/lector_kimai.py tests/test_lector_kimai.py
git commit -m "feat: lector de exports .xlsx de Kimai"
```

---

### Task 3: `mapeo` — resolver códigos a nombres del partner

**Files:**
- Create: `cunix_horas/mapeo.py`
- Create: `config/mapeo.yaml`
- Create: `tests/fixtures/mapeo-test.yaml`
- Test: `tests/test_mapeo.py`

**Interfaces:**
- Consumes: nada.
- Produces:
  - `@dataclass(frozen=True) Persona(nombre: str, archivo: str)`
  - `@dataclass(frozen=True) DestinoProyecto(cliente: str, proyecto: str)`
  - `class ErrorMapeo(Exception)`
  - `class Mapeo` con `Mapeo.cargar(ruta: Path) -> Mapeo` (classmethod), `resolver_proyecto(codigo: str, texto_kimai: str, archivo: str) -> DestinoProyecto`, `resolver_persona(username: str, archivo: str) -> Persona`

- [ ] **Step 1: Crear `tests/fixtures/mapeo-test.yaml`**

```yaml
personas:
  mzalazar:
    nombre: "Matias Zalazar"
    archivo: "Zalazar"

proyectos:
  CO2610170:
    cliente: "Servicio Nacional de Aduanas"
    proyecto: "Subastas"
  CO2510115:
    cliente: "Instituto de Salud Pública de Chile"
    proyecto: "SIAC-OIRS"
  PR2510126:
    cliente: "MINVU"
    proyecto: "SELICO"
```

- [ ] **Step 2: Escribir los tests que fallan**

`tests/test_mapeo.py`:

```python
import pytest
from conftest import FIXTURES

from cunix_horas.mapeo import ErrorMapeo, Mapeo


def cargar():
    return Mapeo.cargar(FIXTURES / "mapeo-test.yaml")


def test_resolver_proyecto_conocido():
    destino = cargar().resolver_proyecto("CO2610170", "[CO2610170] Aduana", "x.xlsx")
    assert destino.cliente == "Servicio Nacional de Aduanas"
    assert destino.proyecto == "Subastas"


def test_resolver_proyecto_conserva_acentos():
    destino = cargar().resolver_proyecto("CO2510115", "[CO2510115] ISPCH", "x.xlsx")
    assert destino.cliente == "Instituto de Salud Pública de Chile"


def test_resolver_persona_conocida():
    persona = cargar().resolver_persona("mzalazar", "x.xlsx")
    assert persona.nombre == "Matias Zalazar"
    assert persona.archivo == "Zalazar"


def test_proyecto_desconocido_sugiere_la_linea_yaml():
    texto = "[PR2610199] MINVU-Portal2 | Subsecretaría de Vivienda - Portal"
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_proyecto("PR2610199", texto, "kimai-mzalazar.xlsx")
    mensaje = str(excepcion.value)
    assert "PR2610199" in mensaje
    assert "kimai-mzalazar.xlsx" in mensaje
    assert "config/mapeo.yaml" in mensaje
    assert "MINVU-Portal2" in mensaje
    assert "cliente:" in mensaje
    assert "proyecto:" in mensaje


def test_persona_desconocida_sugiere_la_linea_yaml():
    with pytest.raises(ErrorMapeo) as excepcion:
        cargar().resolver_persona("fdodera", "kimai-fdodera.xlsx")
    mensaje = str(excepcion.value)
    assert "fdodera" in mensaje
    assert "config/mapeo.yaml" in mensaje
    assert "nombre:" in mensaje
    assert "archivo:" in mensaje


def test_falla_si_el_yaml_no_existe(tmp_path):
    with pytest.raises(ErrorMapeo, match="No se encontró"):
        Mapeo.cargar(tmp_path / "no-existe.yaml")


def test_falla_si_al_yaml_le_falta_una_seccion(tmp_path):
    ruta = tmp_path / "incompleto.yaml"
    ruta.write_text("personas: {}\n", encoding="utf-8")
    with pytest.raises(ErrorMapeo, match="proyectos"):
        Mapeo.cargar(ruta)


def test_falla_si_a_un_proyecto_le_falta_el_cliente(tmp_path):
    ruta = tmp_path / "malo.yaml"
    ruta.write_text(
        "personas: {}\nproyectos:\n  AB123:\n    proyecto: 'X'\n", encoding="utf-8"
    )
    with pytest.raises(ErrorMapeo, match="cliente"):
        Mapeo.cargar(ruta)
```

- [ ] **Step 3: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_mapeo.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'cunix_horas.mapeo'`

- [ ] **Step 4: Escribir `cunix_horas/mapeo.py`**

```python
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
                f"    {codigo}:\n"
                f'      cliente: "AJUSTAR - nombre del cliente para el partner"\n'
                f'      proyecto: "{alias}"'
            )
        return destino

    def resolver_persona(self, username: str, archivo: str) -> Persona:
        persona = self._personas.get(username)
        if persona is None:
            raise ErrorMapeo(
                f"Desarrollador sin mapear en {archivo}: {username}\n"
                f"  Agregá a config/mapeo.yaml, bajo personas:\n"
                f"    {username}:\n"
                f'      nombre: "AJUSTAR - nombre completo, va en A1 del Excel"\n'
                f'      archivo: "AJUSTAR - apellido, va en el nombre del archivo"'
            )
        return persona
```

- [ ] **Step 5: Correr los tests**

Run: `python -m pytest tests/test_mapeo.py -v`
Expected: 8 passed

- [ ] **Step 6: Crear `config/mapeo.yaml` de producción**

```yaml
# Mapeo de Kimai al Excel que recibe el partner.
#
# personas:  clave = username de Kimai (columna F del export)
# proyectos: clave = código entre corchetes del campo Project (columna J)
#
# Si aparece un código nuevo, el proceso frena y te sugiere la línea a pegar.

personas:
  mzalazar:
    nombre: "Matias Zalazar"
    archivo: "Zalazar"

proyectos:
  CO2610170:
    cliente: "Servicio Nacional de Aduanas"
    proyecto: "Subastas"
  CO2510115:
    cliente: "Instituto de Salud Pública de Chile"
    proyecto: "SIAC-OIRS"
  PR2510126:
    cliente: "MINVU"
    proyecto: "SELICO"
```

- [ ] **Step 7: Commit**

```bash
git add cunix_horas/mapeo.py config/mapeo.yaml tests/test_mapeo.py tests/fixtures/mapeo-test.yaml
git commit -m "feat: mapeo de codigos de Kimai a nombres del partner"
```

---

### Task 4: `agregador` — pivot a la jerarquía del Excel

**Files:**
- Create: `cunix_horas/agregador.py`
- Test: `tests/test_agregador.py`

**Interfaces:**
- Consumes: `Registro` de `lector_kimai`; `Mapeo`, `Persona` de `mapeo`.
- Produces:
  - `@dataclass(frozen=True) Fila(cliente: str, proyecto: str, actividad: str, horas_por_dia: dict[int, float])` con propiedad `total: float`
  - `@dataclass(frozen=True) Reporte(nombre_dev: str, nombre_archivo: str, anio: int, mes: int, filas: tuple[Fila, ...], descartados: tuple[Registro, ...])` con propiedades `dias_del_mes: int`, `total: float` y método `total_del_dia(dia: int) -> float`
  - `agregar(registros: list[Registro], mapeo: Mapeo, anio: int, mes: int, archivo: str) -> Reporte`

- [ ] **Step 1: Escribir los tests que fallan**

`tests/test_agregador.py`:

```python
from datetime import date

from conftest import FIXTURES

from cunix_horas.agregador import agregar
from cunix_horas.lector_kimai import Registro, leer
from cunix_horas.mapeo import Mapeo


def mapeo():
    return Mapeo.cargar(FIXTURES / "mapeo-test.yaml")


def reg(dia, horas, codigo="CO2610170", actividad="Desarrollo", mes=8):
    return Registro(
        fecha=date(2026, mes, dia),
        horas=horas,
        username="mzalazar",
        cod_proyecto=codigo,
        actividad=actividad,
    )


def test_agrega_las_horas_del_mismo_dia_y_proyecto():
    reporte = agregar([reg(3, 2.0), reg(3, 1.5)], mapeo(), 2026, 8, "x.xlsx")
    assert len(reporte.filas) == 1
    assert reporte.filas[0].horas_por_dia == {3: 3.5}


def test_separa_por_proyecto():
    reporte = agregar(
        [reg(3, 2.0, "CO2610170"), reg(3, 1.0, "CO2510115")], mapeo(), 2026, 8, "x.xlsx"
    )
    assert len(reporte.filas) == 2
    assert {f.proyecto for f in reporte.filas} == {"Subastas", "SIAC-OIRS"}


def test_separa_por_actividad_dentro_del_mismo_proyecto():
    reporte = agregar(
        [reg(3, 2.0, actividad="Desarrollo"), reg(3, 1.0, actividad="Testing")],
        mapeo(),
        2026,
        8,
        "x.xlsx",
    )
    assert len(reporte.filas) == 2
    assert {f.actividad for f in reporte.filas} == {"Desarrollo", "Testing"}


def test_las_filas_salen_ordenadas_por_cliente():
    reporte = agregar(
        [reg(3, 1.0, "PR2510126"), reg(3, 1.0, "CO2510115"), reg(3, 1.0, "CO2610170")],
        mapeo(),
        2026,
        8,
        "x.xlsx",
    )
    clientes = [f.cliente for f in reporte.filas]
    assert clientes == sorted(clientes)


def test_el_total_de_la_fila_suma_sus_dias():
    reporte = agregar([reg(3, 2.0), reg(5, 4.5)], mapeo(), 2026, 8, "x.xlsx")
    assert reporte.filas[0].total == 6.5


def test_el_total_del_reporte_suma_todas_las_filas():
    reporte = agregar(
        [reg(3, 2.0, "CO2610170"), reg(4, 1.0, "CO2510115")], mapeo(), 2026, 8, "x.xlsx"
    )
    assert reporte.total == 3.0


def test_total_del_dia_suma_todas_las_filas_de_ese_dia():
    reporte = agregar(
        [reg(3, 2.0, "CO2610170"), reg(3, 1.0, "CO2510115"), reg(4, 5.0)],
        mapeo(),
        2026,
        8,
        "x.xlsx",
    )
    assert reporte.total_del_dia(3) == 3.0
    assert reporte.total_del_dia(4) == 5.0
    assert reporte.total_del_dia(10) == 0.0


def test_descarta_los_registros_fuera_del_mes():
    reporte = agregar([reg(3, 2.0), reg(15, 9.0, mes=7)], mapeo(), 2026, 8, "x.xlsx")
    assert reporte.total == 2.0
    assert len(reporte.descartados) == 1
    assert reporte.descartados[0].fecha == date(2026, 7, 15)


def test_dias_del_mes():
    assert agregar([], mapeo(), 2026, 2, "x.xlsx").dias_del_mes == 28
    assert agregar([], mapeo(), 2024, 2, "x.xlsx").dias_del_mes == 29
    assert agregar([], mapeo(), 2025, 4, "x.xlsx").dias_del_mes == 30
    assert agregar([], mapeo(), 2025, 10, "x.xlsx").dias_del_mes == 31


def test_toma_el_nombre_del_dev_del_mapeo():
    reporte = agregar([reg(3, 2.0)], mapeo(), 2026, 8, "x.xlsx")
    assert reporte.nombre_dev == "Matias Zalazar"
    assert reporte.nombre_archivo == "Zalazar"


def test_sobre_el_fixture_real_cierran_los_totales():
    registros = leer(FIXTURES / "kimai-mzalazar.xlsx")
    reporte = agregar(registros, mapeo(), 2026, 8, "kimai-mzalazar.xlsx")
    assert round(reporte.total, 2) == 76.5
    assert round(sum(f.total for f in reporte.filas), 2) == 76.5
    suma_dias = sum(reporte.total_del_dia(d) for d in range(1, reporte.dias_del_mes + 1))
    assert round(suma_dias, 2) == 76.5
    assert {f.proyecto for f in reporte.filas} == {"Subastas", "SIAC-OIRS", "SELICO"}
```

- [ ] **Step 2: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_agregador.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'cunix_horas.agregador'`

- [ ] **Step 3: Escribir `cunix_horas/agregador.py`**

```python
"""Pivot de registros de Kimai a la jerarquía Cliente > Proyecto > Actividad."""
from __future__ import annotations

import calendar
from collections import defaultdict
from dataclasses import dataclass, field

from cunix_horas.lector_kimai import Registro
from cunix_horas.mapeo import Mapeo


@dataclass(frozen=True)
class Fila:
    """Una actividad de un proyecto de un cliente, con sus horas por día."""

    cliente: str
    proyecto: str
    actividad: str
    horas_por_dia: dict[int, float] = field(default_factory=dict)

    @property
    def total(self) -> float:
        return sum(self.horas_por_dia.values())


@dataclass(frozen=True)
class Reporte:
    """Todo lo que hace falta para escribir el Excel de un desarrollador."""

    nombre_dev: str
    nombre_archivo: str
    anio: int
    mes: int
    filas: tuple[Fila, ...]
    descartados: tuple[Registro, ...]

    @property
    def dias_del_mes(self) -> int:
        return calendar.monthrange(self.anio, self.mes)[1]

    @property
    def total(self) -> float:
        return sum(f.total for f in self.filas)

    def total_del_dia(self, dia: int) -> float:
        return sum(f.horas_por_dia.get(dia, 0.0) for f in self.filas)


def agregar(
    registros: list[Registro], mapeo: Mapeo, anio: int, mes: int, archivo: str
) -> Reporte:
    """Convierte registros planos en un Reporte jerárquico del mes indicado.

    Los registros con fecha fuera del mes se excluyen y quedan en `descartados`.
    Lanza ErrorMapeo si aparece un código de proyecto o un username sin mapear.
    """
    del_mes: list[Registro] = []
    descartados: list[Registro] = []
    for registro in registros:
        if (registro.fecha.year, registro.fecha.month) == (anio, mes):
            del_mes.append(registro)
        else:
            descartados.append(registro)

    acumulado: dict[tuple[str, str, str], dict[int, float]] = defaultdict(dict)
    for registro in del_mes:
        destino = mapeo.resolver_proyecto(registro.cod_proyecto, "", archivo)
        clave = (destino.cliente, destino.proyecto, registro.actividad)
        dia = registro.fecha.day
        acumulado[clave][dia] = acumulado[clave].get(dia, 0.0) + registro.horas

    filas = tuple(
        Fila(cliente, proyecto, actividad, dict(sorted(horas.items())))
        for (cliente, proyecto, actividad), horas in sorted(acumulado.items())
    )

    usernames = {r.username for r in registros}
    if usernames:
        persona = mapeo.resolver_persona(sorted(usernames)[0], archivo)
        nombre_dev, nombre_archivo = persona.nombre, persona.archivo
    else:
        nombre_dev, nombre_archivo = "", ""

    return Reporte(
        nombre_dev=nombre_dev,
        nombre_archivo=nombre_archivo,
        anio=anio,
        mes=mes,
        filas=filas,
        descartados=tuple(descartados),
    )
```

Nota sobre `resolver_proyecto(..., "", archivo)`: el segundo argumento es el texto original de Kimai y sólo se usa para armar la sugerencia YAML cuando falta el código. El agregador ya no lo tiene, así que pasa cadena vacía y `resolver_proyecto` cae al código como alias. El mensaje sigue siendo completo y accionable.

- [ ] **Step 4: Correr los tests**

Run: `python -m pytest tests/test_agregador.py -v`
Expected: 11 passed

- [ ] **Step 5: Commit**

```bash
git add cunix_horas/agregador.py tests/test_agregador.py
git commit -m "feat: agregador con jerarquia cliente/proyecto/actividad"
```

---

### Task 5: `validador` — avisos que no frenan

**Files:**
- Create: `cunix_horas/validador.py`
- Test: `tests/test_validador.py`

**Interfaces:**
- Consumes: `Reporte`, `Fila` de `agregador`; `Registro` de `lector_kimai`.
- Produces: `validar(reporte: Reporte) -> list[str]`, `LIMITE_HORAS_POR_DIA = 12.0`

- [ ] **Step 1: Escribir los tests que fallan**

`tests/test_validador.py`:

```python
import calendar
from datetime import date

from cunix_horas.agregador import Fila, Reporte
from cunix_horas.lector_kimai import Registro
from cunix_horas.validador import validar


def dias_habiles_completos():
    """Octubre 2025 con 8 h en cada día hábil y nada los fines de semana."""
    return {
        dia: 8.0 for dia in range(1, 32) if calendar.weekday(2025, 10, dia) < 5
    }


def reporte(horas_por_dia, descartados=()):
    return Reporte(
        nombre_dev="Matias Zalazar",
        nombre_archivo="Zalazar",
        anio=2025,
        mes=10,
        filas=(Fila("Cliente", "Proyecto", "Desarrollo", dict(horas_por_dia)),),
        descartados=tuple(descartados),
    )


def test_sin_problemas_no_hay_avisos():
    assert validar(reporte(dias_habiles_completos())) == []


def test_avisa_si_un_dia_supera_las_12_horas():
    horas = dias_habiles_completos()
    horas[1] = 14.0
    avisos = validar(reporte(horas))
    assert any("14.0" in a and "1/10/2025" in a for a in avisos)


def test_avisa_por_horas_en_fin_de_semana():
    horas = dias_habiles_completos()
    horas[4] = 3.0  # 2025-10-04 es sábado
    avisos = validar(reporte(horas))
    assert any("fin de semana" in a and "4/10/2025" in a for a in avisos)


def test_avisa_por_dias_habiles_sin_carga():
    horas = dias_habiles_completos()
    del horas[1]
    del horas[2]
    avisos = validar(reporte(horas))
    assert any("sin carga" in a and "1/10/2025" in a for a in avisos)
    assert any("sin carga" in a and "2/10/2025" in a for a in avisos)


def test_avisa_por_registros_descartados():
    descartado = Registro(date(2025, 9, 30), 4.0, "mzalazar", "CO2610170", "Desarrollo")
    avisos = validar(reporte(dias_habiles_completos(), [descartado]))
    assert any("fuera del mes" in a and "30/9/2025" in a for a in avisos)


def test_no_hay_falsos_positivos_de_descuadre():
    assert not any("Descuadre" in a for a in validar(reporte(dias_habiles_completos())))
```

- [ ] **Step 2: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_validador.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'cunix_horas.validador'`

- [ ] **Step 3: Escribir `cunix_horas/validador.py`**

```python
"""Avisos sobre cargas de horas sospechosas. Nunca frenan la generación."""
from __future__ import annotations

import calendar

from cunix_horas.agregador import Reporte

LIMITE_HORAS_POR_DIA = 12.0
TOLERANCIA = 0.01


def _fecha_legible(dia: int, mes: int, anio: int) -> str:
    return f"{dia}/{mes}/{anio}"


def validar(reporte: Reporte) -> list[str]:
    """Devuelve la lista de avisos de un reporte."""
    avisos: list[str] = []

    for registro in reporte.descartados:
        fecha = registro.fecha
        avisos.append(
            f"Registro fuera del mes, EXCLUIDO del Excel: "
            f"{_fecha_legible(fecha.day, fecha.month, fecha.year)} "
            f"({registro.horas:.2f} h, proyecto {registro.cod_proyecto})"
        )

    for dia in range(1, reporte.dias_del_mes + 1):
        horas = reporte.total_del_dia(dia)
        legible = _fecha_legible(dia, reporte.mes, reporte.anio)
        es_fin_de_semana = calendar.weekday(reporte.anio, reporte.mes, dia) >= 5

        if horas > LIMITE_HORAS_POR_DIA:
            avisos.append(
                f"Más de {LIMITE_HORAS_POR_DIA:.0f} h en un día: {legible} "
                f"tiene {horas:.1f} h"
            )
        if es_fin_de_semana and horas > 0:
            avisos.append(f"Horas cargadas en fin de semana: {legible} ({horas:.1f} h)")
        if not es_fin_de_semana and horas == 0:
            avisos.append(f"Día hábil sin carga: {legible}")

    suma_por_dia = sum(
        reporte.total_del_dia(d) for d in range(1, reporte.dias_del_mes + 1)
    )
    if abs(suma_por_dia - reporte.total) > TOLERANCIA:
        avisos.append(
            f"Descuadre interno: el total del mes es {reporte.total:.2f} h pero la "
            f"suma de los días da {suma_por_dia:.2f} h. Revisar el código."
        )

    return avisos
```

- [ ] **Step 4: Correr los tests**

Run: `python -m pytest tests/test_validador.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
git add cunix_horas/validador.py tests/test_validador.py
git commit -m "feat: validaciones informativas de carga de horas"
```

---

### Task 6: `escritor_excel` — generar el Excel del partner

**Files:**
- Create: `cunix_horas/escritor_excel.py`
- Test: `tests/test_escritor_excel.py`

**Interfaces:**
- Consumes: `Reporte`, `Fila` de `agregador`.
- Produces:
  - `MESES_ABREVIADOS: dict[int, str]`
  - `nombre_de_archivo(reporte: Reporte) -> str`
  - `escribir(reporte: Reporte, plantilla: Path, destino: Path) -> Path`

- [ ] **Step 1: Escribir los tests que fallan**

`tests/test_escritor_excel.py`:

```python
import openpyxl
import pytest
from conftest import FIXTURES

from cunix_horas.agregador import Fila, Reporte
from cunix_horas.escritor_excel import escribir, nombre_de_archivo


def reporte_de_ejemplo(anio=2025, mes=10):
    return Reporte(
        nombre_dev="Franco Dodera",
        nombre_archivo="Dodera",
        anio=anio,
        mes=mes,
        filas=(
            Fila("Club Atlético Talleres", "CRM", "Desarrollo", {4: 2.0, 31: 3.0}),
            Fila("Sistemas - C.UNIX", "Fan Player", "Desarrollo", {31: 4.0}),
        ),
        descartados=(),
    )


@pytest.fixture
def generado(tmp_path):
    destino = tmp_path / "salida.xlsx"
    escribir(reporte_de_ejemplo(), FIXTURES / "plantilla.xlsx", destino)
    return openpyxl.load_workbook(destino).active


def test_nombre_de_archivo():
    assert nombre_de_archivo(reporte_de_ejemplo()) == "Oct Dodera.xlsx"
    assert nombre_de_archivo(reporte_de_ejemplo(mes=1)) == "Jan Dodera.xlsx"


def test_encabezado(generado):
    assert generado["A1"].value == "Franco Dodera"
    assert generado["B1"].value == "Total"
    assert generado["C1"].value == "10/1/2025"
    assert generado["D1"].value == "10/2/2025"
    assert generado["AG1"].value == "10/31/2025"


def test_la_cantidad_de_columnas_sigue_los_dias_del_mes(tmp_path):
    for anio, mes, ultima_columna in [(2025, 2, 30), (2024, 2, 31), (2025, 4, 32)]:
        destino = tmp_path / f"{anio}-{mes}.xlsx"
        escribir(
            Reporte("X", "X", anio, mes, (), ()),
            FIXTURES / "plantilla.xlsx",
            destino,
        )
        hoja = openpyxl.load_workbook(destino).active
        assert hoja.max_column == ultima_columna


def test_jerarquia_de_filas(generado):
    assert generado["A2"].value == "Club Atlético Talleres"
    assert generado["A3"].value == "CRM"
    assert generado["A4"].value == "Desarrollo"
    assert generado["A5"].value == "Sistemas - C.UNIX"
    assert generado["A6"].value == "Fan Player"
    assert generado["A7"].value == "Desarrollo"
    assert generado["A8"].value == "Total"


def test_totales_de_la_columna_b(generado):
    assert generado["B2"].value == 5.0
    assert generado["B3"].value == 5.0
    assert generado["B4"].value == 5.0
    assert generado["B5"].value == 4.0
    assert generado["B8"].value == 9.0


def test_horas_por_dia(generado):
    assert generado["F3"].value == 2.0   # día 4 -> columna 2+4 = F
    assert generado["AG3"].value == 3.0  # día 31 -> columna 33 = AG
    assert generado["AG6"].value == 4.0
    assert generado["D3"].value is None  # día 2, sin horas -> vacío


def test_la_fila_total_pone_cero_en_los_dias_sin_horas(generado):
    assert generado["C8"].value == 0.0
    assert generado["F8"].value == 2.0
    assert generado["AG8"].value == 7.0


def test_las_filas_de_cliente_no_tienen_horas_por_dia(generado):
    assert generado["F2"].value is None


def test_las_filas_de_proyecto_van_en_negrita(generado):
    assert generado["A3"].font.b is True
    assert generado["A6"].font.b is True
    assert generado["A2"].font.b is not True
    assert generado["A4"].font.b is not True


def test_merge_en_las_filas_de_cliente(generado):
    rangos = {str(r) for r in generado.merged_cells.ranges}
    assert "C2:AG2" in rangos
    assert "C5:AG5" in rangos


def test_anchos_de_columna_copiados_de_la_plantilla(generado):
    assert generado.column_dimensions["A"].width == pytest.approx(34.14, abs=0.01)
    assert generado.column_dimensions["B"].width == pytest.approx(9.29, abs=0.01)


def test_el_nombre_de_la_hoja_es_el_de_la_plantilla(generado):
    assert generado.title == "Worksheet"


def test_las_horas_se_redondean_a_dos_decimales(tmp_path):
    destino = tmp_path / "r.xlsx"
    escribir(
        Reporte("X", "X", 2025, 10, (Fila("C", "P", "A", {1: 1 / 3}),), ()),
        FIXTURES / "plantilla.xlsx",
        destino,
    )
    hoja = openpyxl.load_workbook(destino).active
    assert hoja["C4"].value == 0.33
```

- [ ] **Step 2: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_escritor_excel.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'cunix_horas.escritor_excel'`

- [ ] **Step 3: Escribir `cunix_horas/escritor_excel.py`**

```python
"""Generación del Excel con el formato que recibe el partner.

La plantilla es la fuente de estilos, no un archivo que se muta: la cantidad
de filas (proyectos por dev) y de columnas (días del mes) es variable, e
insertar filas en un .xlsx existente rompe estilos de forma silenciosa.
"""
from __future__ import annotations

from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from cunix_horas.agregador import Reporte

# Nunca strftime("%b"): depende del locale de la máquina.
MESES_ABREVIADOS = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}

# Celdas de la plantilla de las que se toma el estilo de cada tipo de fila.
MODELO_ENCABEZADO = "A1"
MODELO_CLIENTE = "A2"
MODELO_PROYECTO = "A3"
MODELO_ACTIVIDAD = "A4"
MODELO_TOTAL = "A8"

PRIMERA_COLUMNA_DE_DIA = 3  # C


def nombre_de_archivo(reporte: Reporte) -> str:
    """'Oct Dodera.xlsx'."""
    return f"{MESES_ABREVIADOS[reporte.mes]} {reporte.nombre_archivo}.xlsx"


def _columna_del_dia(dia: int) -> int:
    return PRIMERA_COLUMNA_DE_DIA + dia - 1


def _estilos_de(plantilla: Path) -> dict:
    """Lee de la plantilla el estilo de cada tipo de fila, el título y los anchos."""
    libro = openpyxl.load_workbook(plantilla)
    hoja = libro.active
    modelos = {
        "encabezado": MODELO_ENCABEZADO,
        "cliente": MODELO_CLIENTE,
        "proyecto": MODELO_PROYECTO,
        "actividad": MODELO_ACTIVIDAD,
        "total": MODELO_TOTAL,
    }
    estilos: dict = {
        nombre: {
            "font": copy(hoja[coord].font),
            "alignment": copy(hoja[coord].alignment),
            "number_format": hoja[coord].number_format,
        }
        for nombre, coord in modelos.items()
    }
    estilos["_titulo"] = hoja.title
    estilos["_anchos"] = {
        letra: dim.width
        for letra, dim in hoja.column_dimensions.items()
        if dim.width is not None
    }
    libro.close()
    return estilos


def _aplicar(celda: Cell, estilo: dict) -> None:
    celda.font = copy(estilo["font"])
    celda.alignment = copy(estilo["alignment"])
    celda.number_format = estilo["number_format"]


def _escribir_fila(
    hoja: Worksheet, nro_fila: int, valores: dict[int, object], estilo: dict, ancho: int
) -> None:
    for columna in range(1, ancho + 1):
        celda = hoja.cell(row=nro_fila, column=columna)
        if columna in valores:
            celda.value = valores[columna]
        _aplicar(celda, estilo)


def escribir(reporte: Reporte, plantilla: Path, destino: Path) -> Path:
    """Escribe el Excel del reporte en `destino` y devuelve esa ruta."""
    estilos = _estilos_de(plantilla)
    dias = reporte.dias_del_mes
    ancho = _columna_del_dia(dias)
    ultima_letra = get_column_letter(ancho)

    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = estilos["_titulo"]

    # Fila 1: nombre del dev, 'Total', y un encabezado por día como TEXTO.
    encabezado: dict[int, object] = {1: reporte.nombre_dev, 2: "Total"}
    for dia in range(1, dias + 1):
        encabezado[_columna_del_dia(dia)] = f"{reporte.mes}/{dia}/{reporte.anio}"
    _escribir_fila(hoja, 1, encabezado, estilos["encabezado"], ancho)

    nro_fila = 2
    cliente_actual: str | None = None
    proyecto_actual: tuple[str, str] | None = None

    for fila in reporte.filas:
        if fila.cliente != cliente_actual:
            total_cliente = sum(
                f.total for f in reporte.filas if f.cliente == fila.cliente
            )
            _escribir_fila(
                hoja,
                nro_fila,
                {1: fila.cliente, 2: round(total_cliente, 2)},
                estilos["cliente"],
                ancho,
            )
            hoja.merge_cells(
                f"{get_column_letter(PRIMERA_COLUMNA_DE_DIA)}{nro_fila}:"
                f"{ultima_letra}{nro_fila}"
            )
            cliente_actual = fila.cliente
            proyecto_actual = None
            nro_fila += 1

        if (fila.cliente, fila.proyecto) != proyecto_actual:
            hermanas = [
                f
                for f in reporte.filas
                if (f.cliente, f.proyecto) == (fila.cliente, fila.proyecto)
            ]
            valores: dict[int, object] = {
                1: fila.proyecto,
                2: round(sum(f.total for f in hermanas), 2),
            }
            for dia in range(1, dias + 1):
                horas = sum(f.horas_por_dia.get(dia, 0.0) for f in hermanas)
                if horas:
                    valores[_columna_del_dia(dia)] = round(horas, 2)
            _escribir_fila(hoja, nro_fila, valores, estilos["proyecto"], ancho)
            proyecto_actual = (fila.cliente, fila.proyecto)
            nro_fila += 1

        valores = {1: fila.actividad, 2: round(fila.total, 2)}
        for dia, horas in fila.horas_por_dia.items():
            valores[_columna_del_dia(dia)] = round(horas, 2)
        _escribir_fila(hoja, nro_fila, valores, estilos["actividad"], ancho)
        nro_fila += 1

    # Fila Total: con 0.0 explícito en los días sin horas, como en la plantilla.
    totales: dict[int, object] = {1: "Total", 2: round(reporte.total, 2)}
    for dia in range(1, dias + 1):
        totales[_columna_del_dia(dia)] = round(reporte.total_del_dia(dia), 2)
    _escribir_fila(hoja, nro_fila, totales, estilos["total"], ancho)

    for letra, ancho_columna in estilos["_anchos"].items():
        hoja.column_dimensions[letra].width = ancho_columna

    destino.parent.mkdir(parents=True, exist_ok=True)
    libro.save(destino)
    return destino
```

- [ ] **Step 4: Correr los tests**

Run: `python -m pytest tests/test_escritor_excel.py -v`
Expected: 13 passed

- [ ] **Step 5: Commit**

```bash
git add cunix_horas/escritor_excel.py tests/test_escritor_excel.py
git commit -m "feat: escritor del Excel con formato del partner"
```

---

### Task 7: `cli`, `generar.bat` y README — el pipeline completo

**Files:**
- Create: `cunix_horas/cli.py`
- Create: `cunix_horas/__main__.py`
- Create: `generar.bat`
- Create: `README.md`
- Create: `input/.gitkeep`
- Test: `tests/test_cli.py`
- Delete: `Oct Dodera.xlsx`, `20260924-kimai-export.xlsx` (ya copiados a `templates/` y `tests/fixtures/`)

**Interfaces:**
- Consumes: `leer`, `ErrorLectura` de `lector_kimai`; `Mapeo`, `ErrorMapeo` de `mapeo`; `agregar` de `agregador`; `validar` de `validador`; `escribir`, `nombre_de_archivo` de `escritor_excel`.
- Produces: `procesar_mes(mes: str, raiz: Path) -> int` (código de salida), `main(argv: list[str] | None = None) -> int`

- [ ] **Step 1: Escribir los tests que fallan**

`tests/test_cli.py`:

```python
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
```

- [ ] **Step 2: Correr los tests para verificar que fallan**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'cunix_horas.cli'`

- [ ] **Step 3: Escribir `cunix_horas/cli.py`**

```python
"""Orquestación: de input/<mes>/*.xlsx a output/<mes>/*.xlsx."""
from __future__ import annotations

import re
import sys
from pathlib import Path

from cunix_horas.agregador import agregar
from cunix_horas.escritor_excel import escribir, nombre_de_archivo
from cunix_horas.lector_kimai import ErrorLectura, leer
from cunix_horas.mapeo import ErrorMapeo, Mapeo
from cunix_horas.validador import validar

FORMATO_MES = re.compile(r"^(\d{4})-(\d{2})$")


def _parsear_mes(mes: str) -> tuple[int, int]:
    coincidencia = FORMATO_MES.match(mes)
    if coincidencia is None:
        raise ValueError(
            f"El mes debe tener el formato AAAA-MM (ej: 2025-10), no {mes!r}"
        )
    anio, numero = int(coincidencia.group(1)), int(coincidencia.group(2))
    if not 1 <= numero <= 12:
        raise ValueError(
            f"Mes fuera de rango en {mes!r}: AAAA-MM con MM entre 01 y 12"
        )
    return anio, numero


def procesar_mes(mes: str, raiz: Path) -> int:
    """Procesa todos los exports de input/<mes>/. Devuelve el código de salida."""
    try:
        anio, numero_mes = _parsear_mes(mes)
    except ValueError as error:
        print(f"ERROR: {error}")
        return 1

    carpeta_entrada = raiz / "input" / mes
    if not carpeta_entrada.is_dir():
        print(f"ERROR: no existe la carpeta input/{mes}")
        print(f"  Creála y poné adentro los exports de Kimai: {carpeta_entrada}")
        return 1

    try:
        mapeo = Mapeo.cargar(raiz / "config" / "mapeo.yaml")
    except ErrorMapeo as error:
        print(f"ERROR: {error}")
        return 1

    plantilla = raiz / "templates" / "plantilla.xlsx"
    if not plantilla.is_file():
        print(f"ERROR: falta la plantilla {plantilla}")
        return 1

    entradas = sorted(
        p for p in carpeta_entrada.glob("*.xlsx") if not p.name.startswith("~$")
    )
    if not entradas:
        print(f"ERROR: no hay ningún .xlsx en input/{mes}")
        return 1

    carpeta_salida = raiz / "output" / mes
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    print(f"Procesando input/{mes}/ ...")
    avisos_totales: list[str] = []
    generados = 0
    hubo_errores = False

    for entrada in entradas:
        try:
            registros = leer(entrada)
            reporte = agregar(registros, mapeo, anio, numero_mes, entrada.name)
            destino = carpeta_salida / nombre_de_archivo(reporte)
            escribir(reporte, plantilla, destino)
        except (ErrorLectura, ErrorMapeo) as error:
            print(f"  {entrada.name}: NO GENERADO")
            for linea in str(error).splitlines():
                print(f"    {linea}")
            hubo_errores = True
            continue

        clientes = len({f.cliente for f in reporte.filas})
        print(
            f"  {entrada.name}  ->  {destino.name}"
            f"  ({reporte.total:.1f} h, {clientes} cliente/s)"
        )
        generados += 1

        avisos = validar(reporte)
        if avisos:
            avisos_totales.append(f"=== {destino.name} ===")
            avisos_totales.extend(f"  {a}" for a in avisos)
            avisos_totales.append("")

    cantidad_avisos = sum(1 for a in avisos_totales if a.startswith("  "))
    (carpeta_salida / "_validacion.txt").write_text(
        "\n".join(avisos_totales) if avisos_totales else "Sin avisos.\n",
        encoding="utf-8",
    )

    print(
        f"{generados} archivo/s generado/s, {cantidad_avisos} aviso/s en "
        f"output/{mes}/_validacion.txt"
    )
    return 1 if hubo_errores else 0


def main(argv: list[str] | None = None) -> int:
    argumentos = sys.argv[1:] if argv is None else argv
    if len(argumentos) != 1:
        print("Uso: python -m cunix_horas AAAA-MM")
        print("Ejemplo: python -m cunix_horas 2025-10")
        return 1
    return procesar_mes(argumentos[0], Path.cwd())
```

- [ ] **Step 4: Escribir `cunix_horas/__main__.py`**

```python
import sys

from cunix_horas.cli import main

sys.exit(main())
```

- [ ] **Step 5: Correr los tests**

Run: `python -m pytest tests/test_cli.py -v`
Expected: 7 passed

- [ ] **Step 6: Correr toda la suite**

Run: `python -m pytest -v`
Expected: 58 passed

- [ ] **Step 7: Escribir `generar.bat`**

```bat
@echo off
cd /d "%~dp0"
if "%~1"=="" (
    set /p MES="Mes a generar (AAAA-MM, ej 2025-10): "
) else (
    set MES=%~1
)
python -m cunix_horas %MES%
echo.
pause
```

- [ ] **Step 8: Probar el pipeline de punta a punta con datos reales**

```bash
cd "C:/Users/Martin/Desktop/WORK/CUNIX/cunix-horas"
mkdir -p input/2026-08
cp tests/fixtures/kimai-mzalazar.xlsx input/2026-08/
python -m cunix_horas 2026-08
```

Esperado en consola:

```
Procesando input/2026-08/ ...
  kimai-mzalazar.xlsx  ->  Aug Zalazar.xlsx  (76.5 h, 3 cliente/s)
1 archivo/s generado/s, N aviso/s en output/2026-08/_validacion.txt
```

Verificar a ojo `output/2026-08/Aug Zalazar.xlsx`: 31 columnas de día, jerarquía Cliente→Proyecto→Actividad, y `B` de la fila Total = 76.5.

- [ ] **Step 9: Escribir `README.md`**

````markdown
# cunix-horas

Convierte los exports de Kimai en los Excel mensuales que recibe el partner.

## Uso mensual

1. En Kimai, exportar a **Excel** las horas del mes, **un archivo por desarrollador**.
2. Crear la carpeta del mes y poner los archivos adentro:
   `input/2025-10/`  (el nombre de cada archivo da igual)
3. Doble clic en `generar.bat`, o desde una terminal:
   ```
   python -m cunix_horas 2025-10
   ```
4. Los Excel quedan en `output/2025-10/`, listos para enviar.
5. Leer `output/2025-10/_validacion.txt` antes de mandar nada.

## Cuando aparece un proyecto o un dev nuevo

El proceso frena para ese archivo y muestra la línea exacta a pegar en
`config/mapeo.yaml`. Se pega, se ajusta el nombre que va a ver el partner,
y se vuelve a correr.

## Si el partner cambia el formato del Excel

Reemplazar `templates/plantilla.xlsx` por el archivo nuevo. Los estilos
(fuentes, negritas, anchos de columna, nombre de la hoja) se toman de ahí.
Si cambia la *estructura* (otro orden de filas, otra columna de totales),
hay que tocar `cunix_horas/escritor_excel.py`.

## Tests

```
python -m pytest -v
```

## Estructura

```
config/mapeo.yaml        configuración editada a mano
templates/plantilla.xlsx fuente de estilos del Excel de salida
input/AAAA-MM/           exports de Kimai (se versionan)
output/AAAA-MM/          Excel generados (NO se versionan)
cunix_horas/             el código
tests/                   los tests
docs/superpowers/        spec y plan de implementación
```
````

- [ ] **Step 10: Limpiar los archivos originales de la raíz**

Ya están copiados en `templates/plantilla.xlsx` y `tests/fixtures/kimai-mzalazar.xlsx`.

```bash
cd "C:/Users/Martin/Desktop/WORK/CUNIX/cunix-horas"
rm "Oct Dodera.xlsx" "20260924-kimai-export.xlsx"
touch input/.gitkeep
```

- [ ] **Step 11: Correr la suite completa una última vez**

Run: `python -m pytest -v`
Expected: 58 passed

- [ ] **Step 12: Commit**

```bash
git add -A
git commit -m "feat: CLI, lanzador .bat y documentacion de uso"
```

---

## Self-review

**Cobertura de la spec:**

| Requisito de la spec | Task |
|---|---|
| Lectura del export de Kimai sin openpyxl | 2 |
| Serial de fecha, duración × 24, código de proyecto | 2 |
| Validación del formato del input | 2 |
| Mapeo YAML, match exacto por código | 3 |
| Error con sugerencia YAML lista para pegar | 3 |
| Jerarquía Cliente→Proyecto→Actividad y totales | 4 |
| Exclusión de registros fuera del mes | 4 |
| Los tipos de aviso + red de seguridad de descuadre | 5 |
| Encabezados de día como texto, columnas según el mes | 6 |
| Estilos tomados de la plantilla, negrita en proyecto | 6 |
| Merge `C:AG` en filas de cliente | 6 |
| `0.0` explícito en la fila Total | 6 |
| Redondeo a 2 decimales sólo al escribir | 6 |
| Nombre `<Mes> <Apellido>.xlsx` sin depender del locale | 6 |
| Un archivo que falla no frena a los demás | 7 |
| `_validacion.txt` en la carpeta de salida | 7 |
| `python -m cunix_horas AAAA-MM` y `.bat` | 7 |

**Consistencia de tipos:** `Registro` (Task 2) se consume en Tasks 4 y 5 con los mismos campos. `Mapeo.resolver_proyecto` / `resolver_persona` (Task 3) se llaman con la misma firma en Task 4. `Reporte` y `Fila` (Task 4) se consumen en Tasks 5, 6 y 7 con las mismas propiedades (`dias_del_mes`, `total`, `total_del_dia`, `horas_por_dia`). `escribir` y `nombre_de_archivo` (Task 6) se usan en Task 7 con la firma declarada.

**Deuda conocida y aceptada:** en Task 4, `agregar` llama a `resolver_proyecto` con el texto original de Kimai vacío, así que la sugerencia YAML de un proyecto sin mapear usa el código como alias en vez del nombre que le puso el dev en Kimai. El mensaje sigue siendo accionable. Si molesta en la práctica, la corrección es agregar el texto original a `Registro` y propagarlo.
