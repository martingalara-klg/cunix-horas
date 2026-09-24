# cunix-horas — Conversión de exports Kimai a Excel mensual del partner

**Fecha:** 2026-09-24
**Estado:** Diseño aprobado, pendiente de plan de implementación

## Problema

CUNIX (software factory, outsourcing) debe reportar mensualmente a su partner las horas de cada desarrollador en un formato Excel específico. Los desarrolladores cargan sus horas en **Kimai**. Hoy la conversión del export de Kimai al Excel del partner es manual.

## Objetivo

Un proceso repetible: se depositan los exports mensuales de Kimai en una carpeta, se ejecuta un comando, y salen los Excel listos para enviar — uno por desarrollador, con el formato exacto que el partner ya recibe.

## Alcance

**Entra:** lectura del export `.xlsx` de Kimai, mapeo de proyectos a nombres de cliente, agregación por día, generación del Excel con los estilos de la plantilla, validaciones informativas.

**No entra:** integración con la API de Kimai, envío de mails, facturación, tarifas o montos, interfaz gráfica.

## Formato de entrada — export `.xlsx` de Kimai

Export plano de registros de tiempo ("timesheet"), un archivo por desarrollador por mes. Fila 1 = encabezados. Columnas relevantes:

| Col | Campo | Ejemplo | Uso |
|-----|-------|---------|-----|
| A | Date | `46265.708333333` | Serial Excel, base 1899-12-30 → día del mes |
| D | Duration | `0.14583333` | Fracción de día; × 24 = horas |
| E | User | `Matias Zalazar` | Informativo |
| F | User (username) | `mzalazar` | **Clave** de `personas:` en el mapeo |
| I | Customer | `[608040005] Servicio Nacional de Aduanas` | Informativo |
| J | Project | `[CO2610170] Aduana-Subastas \| Servicio Nacional de Aduanas - Soporte...` | El código entre corchetes es la **clave** de `proyectos:` |
| K | Activity | `Desarrollo` | Tercer nivel de la jerarquía de salida |

Columnas ignoradas: B, C (From/To), G, H, L (Description), M (Billable), N, O, P, Q, R, S, T.

**No se aplica ningún filtro:** todo registro presente en el export entra al Excel.

### Restricción técnica: openpyxl no puede abrir estos archivos

Kimai emite el atributo `showZeroes` en `<sheetView>` donde el esquema OOXML define `showZeros`. `openpyxl.load_workbook()` falla con:

```
TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'
```

Por lo tanto la lectura del input **no usa openpyxl**: se descomprime el `.xlsx` como ZIP y se parsea `xl/worksheets/sheet1.xml` con `xml.etree.ElementTree`, resolviendo `xl/sharedStrings.xml` y también celdas `inlineStr`. Verificado funcionando sobre el archivo real `20260924-kimai-export.xlsx`.

La escritura del output **sí** usa openpyxl (la plantilla es un archivo normal).

## Formato de salida — Excel del partner

Un `.xlsx` por desarrollador, una sola hoja llamada `Worksheet`.

```
     A                        B       C          D          ...  AG
 1   Franco Dodera            Total   10/1/2025  10/2/2025  ...  10/31/2025
 2   Sistemas - C.UNIX        55.5                                          <- CLIENTE
 3   Fan Player               55.5                                    4.0   <- PROYECTO (negrita)
 4   Desarrollo               55.5                                    4.0   <- ACTIVIDAD
 5   Club Atlético Talleres   72.5                                          <- CLIENTE
 6   CRM                      72.5              2.0                   3.0   <- PROYECTO (negrita)
 7   Desarrollo               72.5              2.0                   3.0   <- ACTIVIDAD
 8   Total                   128.0    4.0       7.0        ...       7.0    <- TOTAL
```

Reglas:

- `A1` = nombre completo del desarrollador (de `personas.<username>.nombre`).
- `B1` = literal `"Total"`.
- `C1:…` = un encabezado por día del mes, **como texto** con formato `M/D/YYYY` (`10/1/2025`, sin cero a la izquierda). Se generan tantas columnas como días tenga el mes; para octubre son 31 (C..AG).
- Jerarquía de filas: Cliente → Proyecto → Actividad. Un cliente puede tener varios proyectos; un proyecto, varias actividades.
- Columna `B` de cada fila = total del mes de esa fila.
- Celdas de día vacías cuando no hay horas (no `0`).
- Fila final `Total`: total general en `B`, y total por día en cada columna, **incluyendo `0.0`** en los días sin horas (así está en la plantilla).
- Merge de `C:AG` en cada fila de cliente (en la plantilla: `C2:AG2`, `C5:AG5`).

Nombre de archivo: `<Mes> <archivo>.xlsx` — mes abreviado en inglés con inicial mayúscula (`Oct`) y `personas.<username>.archivo` (`Dodera`) → `Oct Dodera.xlsx`.

### Estilos

`templates/plantilla.xlsx` es la **fuente de estilos**, no un archivo que se muta. Rellenar la plantilla insertando filas rompería estilos de forma silenciosa, y la cantidad de filas (proyectos por dev) y de columnas (días del mes) es variable.

El escritor lee de la plantilla el estilo de una celda representativa por tipo de fila y lo aplica a las filas que genera:

| Tipo de fila | Celda modelo | Estilo observado |
|--------------|--------------|------------------|
| Encabezado | `A1` | default |
| Cliente | `A2` | normal, alineado izquierda |
| Proyecto | `A3` | **negrita** |
| Actividad | `A4` | normal |
| Total | `A8` | normal |

Anchos de columna tomados de la plantilla: `A=34.14`, `B=9.29`, `C=14.0`, `L=15.14`. Sin bordes, sin rellenos, sin colores de fuente especiales.

## Configuración — `config/mapeo.yaml`

```yaml
personas:
  mzalazar:  { nombre: "Matias Zalazar", archivo: "Zalazar" }
  fdodera:   { nombre: "Franco Dodera",  archivo: "Dodera" }

proyectos:
  CO2610170: { cliente: "Servicio Nacional de Aduanas", proyecto: "Subastas" }
  CO2510115: { cliente: "Instituto de Salud Pública de Chile", proyecto: "SIAC-OIRS" }
  PR2510126: { cliente: "MINVU", proyecto: "SELICO" }
```

- La clave de `proyectos` es el código entre corchetes de la columna J (`[CO2610170]` → `CO2610170`). **Match exacto**, nunca por similitud de texto: el código es estable aunque se renombre el proyecto en Kimai, y un match difuso podría imputar horas al cliente equivocado sin que nadie lo note.
- El cliente se declara **por proyecto**, no en una sección aparte. Esto permite agrupar en el Excel proyectos que en Kimai están bajo clientes distintos, o separarlos, según lo que quiera ver el partner.
- La clave de `personas` es la columna F (username de Kimai), no el nombre.

## Arquitectura

```
input/2025-10/*.xlsx
   |
   +--> lector_kimai    -> list[Registro]
   +--> mapeo           -> resuelve cliente/proyecto; ERROR si falta un código
   +--> agregador       -> Reporte: {cliente -> proyecto -> actividad -> {día: horas}}
   +--> validador       -> list[Aviso]
   +--> escritor_excel  -> output/2025-10/<Mes> <Apellido>.xlsx
```

Módulos en `cunix_horas/` (paquete en la raíz del proyecto, no bajo `src/`: así `python -m cunix_horas` funciona sin `pip install -e .`, requisito para el `.bat` de doble clic):

| Módulo | Responsabilidad | Depende de |
|--------|-----------------|------------|
| `lector_kimai.py` | `.xlsx` de Kimai → `list[Registro]`. Parseo XML propio, serial de fecha, duración × 24. | — |
| `mapeo.py` | Carga y valida el YAML. Resuelve código → (cliente, proyecto). Resuelve username → (nombre, archivo). | — |
| `agregador.py` | `list[Registro]` + mapeo → `Reporte` con jerarquía y totales. | `mapeo` |
| `validador.py` | `Reporte` + registros → `list[Aviso]`. | — |
| `escritor_excel.py` | `Reporte` + plantilla → `.xlsx`. | openpyxl |
| `cli.py` | Orquesta: recorre `input/<mes>/`, procesa cada archivo, escribe output y `_validacion.txt`. | todos |

Tipos centrales:

```python
@dataclass(frozen=True)
class Registro:
    fecha: date
    horas: float          # ya convertidas (Duration * 24)
    username: str         # col F
    cod_proyecto: str     # extraído de col J
    actividad: str        # col K

@dataclass(frozen=True)
class Reporte:
    nombre_dev: str
    nombre_archivo: str
    anio: int
    mes: int
    filas: tuple[Fila, ...]   # Fila: (cliente, proyecto, actividad, {día: horas})
```

Todas las estructuras son inmutables; cada etapa devuelve un valor nuevo.

## Manejo de errores

**Frena (no se genera Excel para ese desarrollador):**

- Código de proyecto presente en el export pero ausente de `proyectos:`. El mensaje incluye la línea YAML lista para pegar:

  ```
  Proyecto sin mapear en kimai-mzalazar.xlsx: [PR2610199] MINVU-Portal2
    Agregá a config/mapeo.yaml, bajo proyectos:
      PR2610199:
        cliente: "Subsecretaría del Ministerio de Vivienda y Urbanismo de Chile"
        proyecto: "MINVU-Portal2"   # ajustá el nombre para el partner
  ```

- `username` presente en el export pero ausente de `personas:` (mismo tratamiento).
- Archivo de input ilegible o sin la estructura de columnas esperada.

El fallo de un archivo **no impide** procesar los demás: cada input es independiente. Al final se reporta qué se generó y qué no.

**Avisa (el Excel se genera igual), en `output/<mes>/_validacion.txt`:**

- Registros con fecha fuera del mes de la carpeta → se **excluyen** del Excel y se listan.
- Más de 12 horas cargadas en un mismo día.
- Horas cargadas en sábado o domingo.
- Días hábiles del mes sin ninguna carga.
- Descuadre entre el total general y la suma de los totales por día (red de seguridad contra errores del propio agregador).

## Precisión numérica

Las horas se acumulan como `float` sin redondear y se redondean a 2 decimales **sólo al escribir la celda**. Redondear antes haría que los totales por fila y por día no cierren con el total general.

## Uso

```
python -m cunix_horas 2025-10
```

Y `generar.bat` en la raíz para ejecutarlo con doble clic.

Salida en consola:

```
Procesando input/2025-10/ ...
  kimai-fdodera.xlsx   -> Oct Dodera.xlsx    (128.0 h, 2 clientes)
  kimai-mzalazar.xlsx  -> Oct Zalazar.xlsx   (76.5 h, 3 clientes)
2 archivos generados, 4 avisos en output/2025-10/_validacion.txt
```

## Testing

TDD. Fixtures reales en `tests/fixtures/`:

- `20260924-kimai-export.xlsx` — export real de Kimai (24 registros, 3 proyectos, agosto 2026, 76.5 h) para el lector y el agregador.
- `Oct Dodera.xlsx` — salida real del partner, como contrato **de estructura y estilos** del escritor.

**Advertencia sobre este fixture:** el archivo está recortado. Su fila `Total` (fila 8) declara horas en días (`C8=4.0`, `H8=6.0`, `I8=5.0`) que no aparecen en ninguna fila de cliente/proyecto de arriba: le faltan clientes que sí existían en el original. Por lo tanto **no sirve como contrato de valores** — un test que verifique "los totales cierran" contra este archivo falla por culpa del fixture, no del código. Se usa para validar layout, estilos, merges y posiciones; los valores se prueban con datos sintéticos construidos en el test.

Cobertura por módulo:

- **lector_kimai:** parsea el fixture real; 24 registros; suma 76.5 h; rango 2026-08-03 a 2026-08-31; tolera el atributo `showZeroes`; resuelve tanto `sharedStrings` como `inlineStr`; convierte el serial de fecha correctamente.
- **mapeo:** resuelve códigos conocidos; lanza error con sugerencia YAML ante uno desconocido; ídem para usernames.
- **agregador:** jerarquía correcta; totales por fila, por día y general; días sin horas quedan ausentes; dos registros del mismo día/proyecto se suman.
- **validador:** dispara cada tipo de aviso con un caso mínimo.
- **escritor_excel:** el archivo generado, releído, reproduce celda por celda la estructura de `Oct Dodera.xlsx`; negrita en filas de proyecto; merges de cliente; cantidad de columnas según los días del mes (probar febrero y un mes de 30).
- **cli (integración):** de `input/` a `output/`; un archivo roto no frena a los otros.

## Decisiones tomadas y su razón

| Decisión | Razón |
|----------|-------|
| Mapeo explícito en YAML en vez de derivar los nombres del texto de Kimai | Lo que ve el cliente no debe depender de cómo un dev tipeó el proyecto |
| Match por código exacto, no difuso | Un match difuso imputa horas al cliente equivocado silenciosamente |
| Un Excel por desarrollador | Es el formato que el partner ya recibe; no se cambia lo que funciona |
| La plantilla como fuente de estilos, no como archivo a mutar | Filas y columnas son variables; insertar filas en openpyxl rompe estilos |
| Sin filtro de registros | Decisión del usuario: todo lo cargado en Kimai va al Excel |
| Validaciones que avisan y no frenan | El usuario decide si envía o le reclama al dev; frenar agrega fricción |
| Encabezados de día como texto | Como fechas reales, Excel podría mostrárselas distinto al partner según su configuración regional |
| Parser XML propio para el input | openpyxl no puede abrir los exports de Kimai |

## Fuera de alcance (posible trabajo futuro)

- Integración directa con la API de Kimai (elimina el paso manual de exportar).
- Consolidado mensual de todos los desarrolladores en un solo workbook.
- Comparación mes a mes / detección de desvíos.
