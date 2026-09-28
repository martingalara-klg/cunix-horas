# cunix-horas — Conversión de exports Kimai a Excel mensual del partner

**Fecha:** 2026-09-24
**Estado:** Implementado. Actualizado el 2026-09-28 con el cambio de entregable: del Excel pivoteado por desarrollador al archivo único con el detalle plano.

## Problema

CUNIX (software factory, outsourcing) debe reportar mensualmente a su partner las horas de cada desarrollador en un formato Excel específico. Los desarrolladores cargan sus horas en **Kimai**. Hoy la conversión del export de Kimai al Excel del partner es manual.

## Objetivo

Un proceso repetible: se depositan los exports mensuales de Kimai en una carpeta, se ejecuta un comando, y sale el archivo listo para enviar, con el formato exacto que el partner recibe.

**El entregable cambió** (ver «Formato de salida»): el partner pasó de recibir un Excel pivoteado por desarrollador a recibir **un solo archivo por mes con el detalle plano de todos**, una fila por registro de tiempo. El escritor del formato anterior, su plantilla y sus tests siguen en el repo, pero el CLI ya no los llama.

## Alcance

**Entra:** lectura de los exports de Kimai (`.xlsx` y `.csv`), mapeo opcional de proyectos a nombres de cliente, generación del archivo único con el detalle plano, verificación de integridad y validaciones informativas.

**No entra:** integración con la API de Kimai, envío de mails, facturación, tarifas o montos, interfaz gráfica.

## Formato de entrada — los exports de Kimai

Según con qué reporte de Kimai se exporte, sale un archivo distinto, siempre uno por desarrollador por mes. `lector_kimai.leer()` es un despachador: mira el archivo, elige el lector y devuelve `list[Registro]`, así que agregador, validador y escritor no se enteran del formato. Hoy se aceptan los dos exports del reporte de detalle; el resumen mensual se reconoce y se rechaza (más abajo, y el porqué).

| Se reconoce por | Formato | Lector |
|---|---|---|
| extensión `.csv` | timesheet plano en CSV | `lector_timesheet_csv.py` |
| `.xlsx` con `A1='Date'` | timesheet plano en XLSX | `lector_timesheet_xlsx.py` |
| `.xlsx` con `B1='Total'` | resumen mensual | **rechazado** (ver abajo) |

### El resumen mensual dejó de alcanzar

El partner cambió el entregable: ahora recibe **un solo archivo con el
detalle plano de todos los desarrolladores**, una fila por registro de
tiempo, con la hora de inicio, el nombre para mostrar, el mail, la
descripción y el número de proyecto de cada carga.

El export de resumen mensual no trae nada de eso: es la grilla de días, con
las horas ya sumadas. Aceptarlo generaría filas incompletas en silencio, que
es justo el modo de falla que este proyecto evita en todos lados. Por eso
`lector_kimai.leer()` lo **rechaza con `ErrorLectura`**, y el mensaje le dice
al dueño qué datos faltan y que tiene que volver a exportar esa persona con
el reporte de detalle. Como cualquier fallo de lectura, es de *ese* archivo:
los demás desarrolladores se procesan igual.

`lector_resumen_mensual.py` **no se borró y sigue probado**: el parseo (orden
de la fecha, separador decimal, clasificación de filas, verificación contra
los totales declarados) costó trabajo y el partner podría volver al formato
anterior. Se recupera cambiando una línea en el despachador; sus tests llaman
al lector interno en vez del punto de entrada.

**Ante cualquier otra cosa se falla**, nunca se adivina: el error nombra el archivo, dice qué encontró en A1 y B1, y enumera los formatos que se reconocen. Como todo fallo de lectura, es de *ese* archivo: los demás desarrolladores se procesan igual.

### 1. Timesheet plano `.xlsx`

Export plano de registros de tiempo ("timesheet"), un archivo por desarrollador por mes. Fila 1 = encabezados. Columnas relevantes:

| Col | Campo | Ejemplo | Uso |
|-----|-------|---------|-----|
| A | Date | `46265.708333333` | Serial Excel, base 1899-12-30 → día del mes |
| B | From | `13:00` | Hora de inicio; va en la celda de fecha del entregable |
| D | Duration | `0.14583333` | Fracción de día; × 24 = horas |
| E | Name | `Matias Zalazar` | Nombre para mostrar |
| F | User (username) | `mzalazar` | **Clave** de `personas:` en el mapeo |
| G | E-mail | `matias.zalazar@cunix.net` | Va al entregable |
| I | Customer | `[608040005] Servicio Nacional de Aduanas` | Texto crudo del cliente |
| J | Project | `[CO2610170] Aduana-Subastas \| Servicio Nacional de Aduanas - Soporte...` | El código entre corchetes es la **clave** de `proyectos:` |
| K | Activity | `Desarrollo` | Tercer nivel de la jerarquía de salida |
| L | Description | `Ticket R-012528` | Lo que hizo; **vacía es válido** |
| R | Project number | `210` | Va tal cual al entregable |

Columnas ignoradas: C (To), H, M (Billable), N, O, P, Q, S, T.

**`Project number` no es el código entre corchetes.** Son dos campos
distintos de Kimai y los dos hacen falta. El proyecto de Luciano dice
`[AD2690002]` en el corchete y tiene `Project number` `210` — y en el archivo
de septiembre del partner ese proyecto figura con `210`. El del corchete es
la clave del mapeo, que es estable aunque se renombre el proyecto; el número
va tal cual a la salida. Confundirlos le mandaría al partner un número de
proyecto que no es el suyo, sin ninguna señal.

Este lector toma las columnas **por posición**, así que las del detalle se
verifican igual que las otras: si están y no dicen lo que tienen que decir,
se falla en vez de poner el dato equivocado en cada columna. Si directamente
no están (un export viejo), los campos nuevos quedan vacíos y las horas se
leen igual.

**No se aplica ningún filtro:** todo registro presente en el export entra al Excel.

### Restricción técnica: openpyxl no puede abrir estos archivos

Kimai emite el atributo `showZeroes` en `<sheetView>` donde el esquema OOXML define `showZeros`. `openpyxl.load_workbook()` falla con:

```
TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'
```

Por lo tanto la lectura del input **no usa openpyxl**: se descomprime el `.xlsx` como ZIP y se parsea `xl/worksheets/sheet1.xml` con `xml.etree.ElementTree`, resolviendo `xl/sharedStrings.xml` y también celdas `inlineStr`. Verificado funcionando sobre el archivo real `20260924-kimai-export.xlsx`.

La escritura del output **sí** usa openpyxl (la plantilla es un archivo normal).

El parseo XML vive una sola vez, en `kimai_comun.py`, y lo usan los dos lectores de `.xlsx`.

### 2. Timesheet plano `.csv`

El mismo contenido que el anterior, con los mismos nombres de encabezado, pero con dos diferencias que importan: la fecha viene en ISO (`2026-08-31`) en vez de serial de Excel, y la duración viene en `H:MM` (`2:00`) en vez de fracción de día. El archivo está en UTF-8 y puede traer BOM.

Acá las columnas se leen **por nombre de encabezado, no por posición**: Kimai agrega y reordena columnas de una versión a otra, y leer por posición haría que un export nuevo imputara horas equivocadas sin avisar.

### 3. Resumen mensual `.xlsx`

Tiene la misma forma que el Excel de salida: fila 1 con el nombre del dev, `Total` y un encabezado por día; después una fila de cliente, una de proyecto y una de actividad por cada combinación; y una fila `Total` al final.

```
A1='Lautaro Zalazar'  B1='Total'  C1='1/8/2026'  D1='2/8/2026' ...
A2='[763043541] C.UNIX'                          B2='9,00'      <- CLIENTE (días mergeados)
A3='[GI2680001] C.UNIX - Proyectos Internos ...' B3='9,00'  E3='0,50' ...   <- PROYECTO
A4='Coordinación interna'                        B4='9,00'  E4='0,50' ...   <- ACTIVIDAD
A5='Total'                                       B5='9,00'  ...
```

- **Qué es cada fila** se decide por estructura, no por posición, igual que del lado de la escritura (`escritor_excel._fila_cliente`): la de cliente tiene las celdas de día mergeadas, la de proyecto trae el código entre corchetes y las horas por día, y la de actividad viene abajo. **Sólo las filas de actividad emiten `Registro`**: las de proyecto son subtotales y duplicarían las horas.
- **El orden de la fecha del encabezado se detecta por archivo.** Como están todas las columnas del mes, uno de los dos componentes es constante y ése es el mes: `1/8/2026`…`31/8/2026` es día/mes; `8/1/2026`…`8/31/2026` es mes/día. Si no se puede decidir sin ambigüedad, se falla: adivinar mal movería las horas de día y de mes sin que nadie lo note.
- **El separador decimal también se detecta por archivo**: un export trae `9,00` y otro `150.0`.
- **Red de seguridad:** el archivo declara su propio total en la columna B de la fila `Total`. Después de parsear se compara la suma leída contra ese total, y si no coinciden (con una tolerancia chica por el redondeo del propio archivo) se falla. Es una verificación que los otros dos formatos no permiten, y ataja cualquier error de lectura de la grilla antes de que llegue al Excel del cliente.

Este formato **no trae el username**, sólo el nombre para mostrar (`A1`). El lector pone en `Registro.username` lo que venga en A1 y no sabe nada del mapeo; es `Mapeo.resolver_persona` el que busca primero por username y, si no encuentra, por el `nombre:` que ya está configurado en cada persona. Si dos personas comparten el mismo `nombre:`, falla en vez de elegir una.

## Formato de salida — el detalle plano (vigente)

Un solo `.xlsx` por mes, con **todos** los desarrolladores, una fila por registro de tiempo. El contrato es `Horas KLG-Sept2025.xlsx`, el archivo que mandó el partner (160 filas, septiembre 2025, 5 desarrolladores).

### Por qué cambió

El partner factura sobre el detalle, no sobre el pivot: quiere ver cada carga de horas con su hora de inicio, su descripción y su número de proyecto, y quiere un único adjunto por mes. El pivot agregaba las horas por día y perdía todo eso.

### Las diez columnas, en este orden

```
Date | Duration | Name | User | E-mail | Customer | Project | Activity | Description | Project number
```

- Fila 1: encabezados **en negrita**, con relleno `FFEEEEEE`.
- `Date`: `datetime` que combina la fecha con la **hora de inicio** (`From`), con formato de celda `yyyy-mm-dd`. La hora se guarda pero no se muestra; está ahí para ordenar y para que el dato no se pierda.
- `Duration`: **`timedelta`**, con formato de celda `[hh]:mm`. No es un número de horas: es una duración real, así que el partner totaliza la columna y obtiene `160:30` y no `160,5`.
- Anchos de columna, de A a J: `9.22, 9.56, 11.89, 7.22, 19.67, 22.11, 21.0, 8.67, 40.67, 15.0`.
- Autofiltro sobre el rango de datos.
- `Description` **puede venir vacía y está bien**: 112 de las 160 filas del archivo de referencia lo están. La celda queda vacía, no con un texto vacío.

**Ocho columnas son passthrough directo de Kimai**, sin transformar: `Date`, `Duration`, `Name`, `User`, `E-mail`, `Activity`, `Description` y `Project number`.

Ojo con `Project number`: **no es el código entre corchetes**. Son dos campos distintos de Kimai que se parecen. El del corchete (`[AD2690002]`) es la clave del mapeo; `Project number` (`210`) es un campo propio de Kimai que va tal cual a la salida. `Registro` guarda los dos.

Las otras dos, `Customer` y `Project`, salen del mapeo **si el proyecto está declarado**, y si no se derivan del texto crudo:

- Cliente: se le saca el prefijo `[codigo] ` si lo tiene. La derivación tiene que funcionar igual **cuando no hay prefijo**: el `Customer` de uno de los desarrolladores viene como `CUNIX`, sin corchetes, y asumir el prefijo dejaría esa columna en blanco para toda esa persona.
- Proyecto: se le saca el prefijo `[codigo] ` y se corta en el `|` (lo que sigue es la descripción larga de Kimai).

### Orden de las filas

Agrupadas por desarrollador, los desarrolladores **alfabéticamente por el nombre para mostrar**, y dentro de cada uno **cronológicamente** (fecha, después hora de inicio). A igualdad de momento se conserva el orden del export, para que dos corridas sobre los mismos datos den el mismo archivo y comparar dos envíos siga sirviendo.

### Nombre del archivo

`Horas KLG-<Mes><Año>.xlsx`, siguiendo el archivo del partner. Es configurable en `config/mapeo.yaml` (`archivo_salida:`), con `{mes}` (tabla fija `Jan`…`Dec`) y `{anio}`. Por defecto, agosto de 2026 da `Horas KLG-Aug2026.xlsx`.

### El riesgo del archivo único, y cómo se resuelve

Antes, si el archivo de un desarrollador fallaba, **faltaba un Excel entero** en la carpeta: imposible no notarlo. Ahora todo va junto, así que **un desarrollador que falta es invisible**: el archivo se ve completo y no lo es.

Por eso, si algún export falla:

- el consolidado **se genera igual** con los que sí se pudieron leer (no generarlo dejaría al dueño sin nada que revisar);
- sale con la marca en el **nombre**: `Horas KLG-Aug2026 (INCOMPLETO - FALTAN 2 DESARROLLADORES - NO ENVIAR).xlsx`, que es lo único que el dueño ve al adjuntarlo a un mail;
- y el archivo **limpio** que hubiera quedado de una corrida anterior se aparta a `... (CORRIDA ANTERIOR - NO ENVIAR).xlsx`, para que no se envíe en su lugar.

Cuando no falla nada, el nombre es el limpio.

Arriba de todo, `_validacion.txt` dice qué desarrolladores entraron y cuáles no, con el motivo de cada fallo.

## Formato de salida anterior — Excel pivoteado por desarrollador (conservado, sin ejecutar)

**Este formato ya no se genera.** El escritor (`escritor_excel.py`), su plantilla (`templates/plantilla.xlsx`) y sus tests siguen en el repo y siguen verdes: si el partner vuelve atrás, se recupera sin reescribir nada. Lo que sigue lo describe.

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
- Jerarquía de filas: Cliente → Proyecto → Actividad. Un cliente puede tener varios proyectos; **cada proyecto tiene exactamente una fila de actividad**.
- **Una sola fila de actividad por proyecto.** En Kimai cada desarrollador clasifica sus horas como quiere (`Desarrollo`, `Gestión`, `Testing`…). Esa clasificación es interna: el partner factura sobre el proyecto y viene recibiendo una única fila desde siempre. Todas las actividades de un mismo proyecto se suman en una sola fila, cuyo texto sale de `actividad:` en `config/mapeo.yaml` (`Desarrollo` si no está declarado).
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
archivo_salida: "Horas KLG-{mes}{anio}.xlsx"

actividad: "Desarrollo"

personas:   # opcional
  mzalazar:  { nombre: "Matias Zalazar", archivo: "Zalazar" }
  fdodera:   { nombre: "Franco Dodera",  archivo: "Dodera" }

proyectos:
  CO2610170: { cliente: "Servicio Nacional de Aduanas", proyecto: "Subastas" }
  CO2510115: { cliente: "Instituto de Salud Pública de Chile", proyecto: "SIAC-OIRS" }
  PR2510126: { cliente: "MINVU", proyecto: "SELICO" }
```

- `archivo_salida:` es el nombre del único archivo que recibe el partner. `{mes}` sale de la tabla fija de meses del proyecto (`Jan`…`Dec`) y `{anio}` es el año de cuatro dígitos; agosto de 2026 da `Horas KLG-Aug2026.xlsx`. Es opcional. El archivo de referencia del partner usa `Sept` para septiembre: si se quiere esa forma exacta, o meses en español, se escribe el mes a mano en el patrón, a costa de tener que actualizar la línea cada mes. Se valida al cargar el mapeo —antes de leer ningún export— que el patrón no use reemplazos inventados y que termine en `.xlsx`.
- **`personas:` es opcional.** El archivo del partner trae el nombre, el usuario y el mail de cada desarrollador tal como vienen de Kimai, así que el entregable vigente no necesita declarar a nadie. Si la sección no está, el mapeo carga igual. La sigue usando el escritor por desarrollador, que quedó conservado.
- **Cada entrada de `proyectos:` es opcional también.** Un proyecto sin declarar **ya no frena nada**: sale con el nombre derivado de Kimai y queda listado en `_validacion.txt` con el bloque listo para pegar. El mapeo pasó de ser obligatorio a ser un pulido opcional de nombres, y se puede porque cada fila lleva su `Project number`: la trazabilidad no depende del mapeo.
- La clave de `proyectos` es el código entre corchetes de la columna J (`[CO2610170]` → `CO2610170`). **Match exacto**, nunca por similitud de texto: el código es estable aunque se renombre el proyecto en Kimai, y un match difuso podría imputar horas al cliente equivocado sin que nadie lo note.
- El cliente se declara **por proyecto**, no en una sección aparte. Esto permite agrupar en el Excel proyectos que en Kimai están bajo clientes distintos, o separarlos, según lo que quiera ver el partner.
- La clave de `personas` es la columna F (username de Kimai), no el nombre.
- `actividad:` es el texto de la **única fila de actividad** que lleva cada proyecto en el Excel. Es opcional: si no está declarado vale `Desarrollo`, que es lo que el partner recibió siempre. Es configuración y no una constante en el código para que cambiar lo que ve el partner sea editar una línea. Si está declarado pero vacío o no es texto, el mapeo no carga: un Excel con la fila de actividad en blanco se vería raro del otro lado y nadie sabría de dónde salió.

## Arquitectura

```
input/2025-10/*.xlsx  y  *.csv
   |
   +--> lector_kimai      -> detecta el formato y delega -> list[Registro]
   +--> detalle           -> Detalle: una FilaDetalle por registro, ya ordenada;
   |                         mapeo si el proyecto está declarado, derivación si no
   +--> validador         -> avisos por desarrollador
   +--> escritor_detalle  -> output/2025-10/Horas KLG-Oct2025.xlsx  + integridad
```

El camino anterior (`agregador` → `escritor_excel`, un Excel por desarrollador) sigue en el repo y probado, pero el CLI ya no lo recorre.

Módulos en `cunix_horas/` (paquete en la raíz del proyecto, no bajo `src/`: así `python -m cunix_horas` funciona sin `pip install -e .`, requisito para el `.bat` de doble clic):

| Módulo | Responsabilidad | Depende de |
|--------|-----------------|------------|
| `lector_kimai.py` | Despachador: detecta el formato del export, delega en los dos lectores de timesheet y rechaza el resumen mensual. | los lectores de timesheet |
| `kimai_comun.py` | `Registro`, `ErrorLectura` y el parseo XML del `.xlsx`, compartidos. | — |
| `lector_timesheet_xlsx.py` | Timesheet `.xlsx` → `list[Registro]`. Serial de fecha, duración × 24. | `kimai_comun` |
| `lector_timesheet_csv.py` | Timesheet `.csv` → `list[Registro]`. Fecha ISO, duración `H:MM`, columnas por nombre. | `kimai_comun` |
| `lector_resumen_mensual.py` | Resumen mensual `.xlsx` → `list[Registro]`. Verifica contra el total declarado. **Ya no se usa desde `leer()`**: se conserva probado por si el partner vuelve a ese formato. | `kimai_comun` |
| `mapeo.py` | Carga y valida el YAML. Resuelve código → (cliente, proyecto). Resuelve username *o* nombre para mostrar → (nombre, archivo). | — |
| `detalle.py` | `list[Registro]` + mapeo → `Detalle`: una `FilaDetalle` por registro, ordenadas, más los proyectos sin mapear y los `Project number` ambiguos. | `mapeo` |
| `escritor_detalle.py` | `Detalle` → el `.xlsx` del partner. Incluye la verificación de integridad, que **relee** el archivo escrito. | openpyxl |
| `agregador.py` | **Conservado, sin ejecutar.** `list[Registro]` + mapeo → `Reporte` con jerarquía y totales. | `mapeo` |
| `validador.py` | `avisos_de_desarrollador()` sobre registros crudos (vigente) y `validar()`/`dato_de_desvio()` sobre el `Reporte` (conservados). | — |
| `escritor_excel.py` | **Conservado, sin ejecutar.** `Reporte` + plantilla → `.xlsx` pivoteado. | openpyxl |
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
    # --- lo que pide el detalle plano; al final y con valor por defecto,
    #     para no romper las construcciones posicionales que ya existen ---
    texto_proyecto: str = ""       # col J tal cual
    hora_inicio: time | None = None  # col B (From); None si Kimai no la trae
    nombre: str = ""               # col E (Name), distinto del username
    email: str = ""                # col G
    descripcion: str = ""          # col L; vacía es válido
    numero_proyecto: str = ""      # col R; NO es el código entre corchetes
    texto_cliente: str = ""        # col I tal cual

@dataclass(frozen=True)
class FilaDetalle:          # una fila del archivo del partner
    fecha_hora: datetime    # fecha + hora de inicio; se muestra sólo la fecha
    duracion: timedelta     # duración real, no un número de horas
    nombre: str
    username: str
    email: str
    cliente: str            # del mapeo, o derivado del texto de Kimai
    proyecto: str           # idem
    actividad: str
    descripcion: str        # vacía es válido
    numero_proyecto: str    # el campo propio de Kimai, NO el código del corchete

@dataclass(frozen=True)
class Reporte:              # conservado, para el escritor pivoteado
    nombre_dev: str
    nombre_archivo: str
    anio: int
    mes: int
    filas: tuple[Fila, ...]   # Fila: (cliente, proyecto, actividad, {día: horas})
                              # una sola Fila por (cliente, proyecto);
                              # `actividad` es el texto configurado, no el de Kimai
```

Todas las estructuras son inmutables; cada etapa devuelve un valor nuevo.

## Manejo de errores

**Frena la corrida entera (no se escribe nada):**

- **La verificación de integridad no cierra.** Después de escribir el archivo en el temporal, se lo **relee** y se suman sus duraciones; si el total no coincide *exactamente* con la suma de las horas de los registros leídos, no se reemplaza nada y se falla ruidoso. Con 160 filas de cinco desarrolladores en un solo archivo, una fila perdida no se ve nunca a ojo, y el archivo se vería completo sin serlo. No es un aviso: es lo único que separa un entregable correcto de uno creíble y equivocado.

**Deja a ese desarrollador afuera del archivo (y el consolidado sale marcado como INCOMPLETO):**

- Archivo de input ilegible, de formato desconocido, o sin la estructura de columnas esperada.
- Export de resumen mensual: el formato se reconoce pero no trae los datos que el partner pide por fila.
- Export sin ninguna fila de datos, o con todas sus filas fuera del mes que se está generando: las dos cosas suelen ser el rango de fechas mal puesto en Kimai.
- Hora de inicio (`From`) que no tiene formato de hora: va en la celda de fecha del entregable, y mal leída movería el registro de día.
- Proyecto sin código entre corchetes en la columna `Project`.

El fallo de un archivo **no impide** procesar los demás: cada input es independiente. El informe dice, arriba de todo, quién entró y quién no, con el motivo de cada fallo.

**Ya no frena:** un proyecto ausente de `proyectos:`, y un export con horas de más de un desarrollador. Lo primero porque el nombre se deriva de Kimai y el `Project number` viaja en cada fila. Lo segundo porque el chequeo existía para que el Excel pivoteado no le imputara las horas de todos a una sola persona: en el detalle plano cada fila lleva su propio `Name`, `User` y `E-mail`, así que ese error ya no es posible, y frenar el archivo dejaría afuera a **dos** desarrolladores en vez de incluirlos bien.

**Frenaba en el formato anterior** (conservado, sin ejecutar):

- Código de proyecto presente en el export pero ausente de `proyectos:`. El mensaje incluye la línea YAML lista para pegar. El sangrado del bloque sugerido debe coincidir con el del archivo — 2 espacios para la clave, 4 para los campos: con 4 y 6, el bloque pegado queda anidado dentro de la entrada anterior, el YAML sigue siendo válido, la validación no lo detecta, y el mapeo se pierde en silencio. Hay tests de round-trip que lo cubren.

  ```
  Proyecto sin mapear en kimai-mzalazar.xlsx: [PR2610199] MINVU-Portal2
    Agregá a config/mapeo.yaml, bajo proyectos:
    PR2610199:
      cliente: "Subsecretaría del Ministerio de Vivienda y Urbanismo de Chile"
      proyecto: "MINVU-Portal2"   # ajustá el nombre para el partner
  ```

- Desarrollador presente en el export pero ausente de `personas:` (mismo tratamiento). Lo que llega puede ser el username (timesheet) o el nombre para mostrar (resumen mensual), y el bloque YAML sugerido se arma según cuál sea.
- Dos personas de `personas:` con el mismo `nombre:` y un resumen mensual que trae ese nombre: no hay forma de saber cuál es, y se falla.
- Resumen mensual cuyo total declarado no cierra con lo parseado.
- Resumen mensual con encabezados de día ambiguos (no se puede decidir día/mes vs mes/día).
- Archivo de input ilegible, de formato desconocido, o sin la estructura de columnas esperada.
- Export de resumen mensual: el formato se reconoce pero no trae los datos que el partner pide por fila.
- Hora de inicio (`From`) que no tiene formato de hora: va en la celda de fecha del entregable, y mal leída movería el registro de día.

**Avisa (el archivo se genera igual), en `output/<mes>/_validacion.txt`:**

Agrupados por desarrollador:

- Registros con fecha fuera del mes de la carpeta → se **excluyen** del archivo y se listan.
- Más de 12 horas cargadas en un mismo día.
- Horas cargadas en sábado o domingo.
- Días hábiles del mes sin ninguna carga.

Del mes entero:

- **Proyectos sin mapear**, con el nombre que se usó y el bloque YAML listo para pegar en `config/mapeo.yaml` si el dueño quiere otro.
- **Un mismo `Project number` con más de un nombre de proyecto en el mes**: significa que alguien renombró el proyecto en Kimai a mitad de mes y el partner vería dos nombres para lo mismo.
- Los `.xlsx` que están en `output/<mes>/` y **esta corrida no generó** (los Excel del formato anterior, un consolidado marcado como INCOMPLETO de otra corrida, archivos del dueño). No se borran nunca, pero se nombran uno por uno: el dueño adjunta la carpeta, no la corrida.

**Se eliminó** el aviso de desvío por redondeo: en este formato las duraciones van exactas y no hay nada que redondear. La sección «Precisión numérica» describe el formato anterior.

## Precisión numérica

**Esta sección describe el formato pivoteado, que quedó conservado sin ejecutarse.** En el detalle plano no hay redondeo de presentación: cada duración se escribe como `timedelta` exacto. Lo único que se redondea es el `float` que trae Kimai (1 h llega como `1.000000000000008`), y se redondea **a segundos**, una sola vez, en `detalle.segundos_de()`; tanto la celda escrita como la verificación de integridad parten de ese mismo número, así que la comparación es exacta por construcción y sólo puede fallar si se pierde o se duplica una fila.

Las horas se acumulan como `float` sin redondear y se redondean a 2 decimales **una sola vez: en la celda de día de la fila de actividad**, que es el dato de base. Las horas de un mismo día que en Kimai vienen de actividades distintas se suman **antes** de ese redondeo, no después: por eso la colapsada de actividades vive en el agregador y no en el escritor. Colapsar en el escritor sumaría totales ya redondeados y agregaría un segundo redondeo (`1/3 + 1/3` daría `0.66` en vez de `0.67`), y el Excel podría dejar de cerrar. Todos los demás valores que se muestran —total de la actividad, celdas y total del proyecto, total del cliente, fila `Total` y totales por día— se derivan **sumando valores ya redondeados**.

Este criterio reemplaza al anterior ("acumular sin redondear y redondear sólo al escribir"). Aquel suponía que redondear tarde garantizaba consistencia, y no la garantiza: cada celda se redondeaba por separado sobre una agregación distinta, así que con tercios de hora (20/40/50 minutos, muy comunes en Kimai) una fila mostraba `0.33` cinco veces y declaraba `1.67` de total. El total general era exacto, pero el Excel no cerraba a la vista y el cliente que suma una fila no obtenía el número declarado.

El precio del criterio nuevo es que el total general puede apartarse de las horas reales del export. **El desvío no es "unos centésimos": escala con la cantidad de celdas.** Cada celda de día se redondea por separado y aporta hasta `0.005 h` de error, así que la cota del total es aproximadamente `0.005 h x (celdas de día no vacías)`. Un mes de 154 h repartido en muchas celdas ya muestra `153.72` en el Excel: **0.28 h de desvío**. Con 200 celdas la cota sube a 1 h.

Es igual la elección correcta para un documento que el cliente lee y suma —un Excel cuyas filas no cierran a la vista es peor que un total corrido—, pero el desvío tiene que quedar a la vista del dueño.

De eso se ocupa el aviso **"Desvío por redondeo"** del validador: compara el total del Excel (la suma de los valores redondeados, lo que el cliente ve) contra las horas crudas del export, y avisa si la diferencia supera **0.5 h**. Ese umbral es el punto en que la diferencia empieza a ser discutible en una factura; por debajo es ruido de presentación que no vale la pena poner delante del dueño todos los meses. El mensaje dice cuántas horas de diferencia hay y aclara que es efecto del redondeo y no un error de carga, para que el dueño decida si le importa.

El aviso anterior comparaba valores redondeados contra valores redondeados: la misma cuenta dos veces, una tautología que no podía dispararse nunca. La red de seguridad que esta sección declaraba no existía.

## Uso

```
python -m cunix_horas 2025-10
```

Y `generar.bat` en la raíz para ejecutarlo con doble clic.

Salida en consola:

```
Procesando input/2025-10/ ...
  kimai-fdodera.xlsx  ->  Franco Dodera  (33 registro/s, 128.00 h)
  kimai-mzalazar.xlsx  ->  Matias Zalazar  (24 registro/s, 76.50 h)
  Archivo para el partner: Horas KLG-Oct2025.xlsx  (57 fila/s, 204.50 h)
2 export/s leído/s, 0 no leído/s, 4 aviso/s en output/2025-10/_validacion.txt
```

Un export que no se puede leer no frena a los demás: sale como
`NO ENTRA AL ARCHIVO` con su motivo debajo, entra en la cuenta de la última
línea, y el archivo del mes sale con la marca de INCOMPLETO en el nombre. Y si en
`output/<mes>/` quedó algún `.xlsx` que esta corrida no generó, antes de esa
línea sale el aviso correspondiente:

```
  OJO: en output/2025-10/ hay 1 .xlsx que esta corrida NO generó. No los envíes; están listados en el informe.
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
- **detalle:** passthrough de las ocho columnas; mapeo vs derivación de `Customer` y `Project`; cliente sin corchetes (`CUNIX`); `Project number` del campo propio y no del corchete (`210` contra `AD2690002`); orden agrupado por dev alfabético y cronológico dentro de cada uno; `Project number` con dos nombres.
- **escritor_detalle:** el archivo generado, releído, tiene las diez columnas en orden, los encabezados en negrita con su relleno, los formatos de celda de `Date` (`yyyy-mm-dd`, conservando la hora) y `Duration` (`[hh]:mm`, `timedelta` y no número), los anchos y el autofiltro; la verificación de integridad detecta una fila perdida.
- **cli (integración):** de `input/` a `output/`; un export roto no frena a los otros y el consolidado sale con el nombre que avisa que está incompleto; el informe dice quién falta y por qué; el limpio de una corrida anterior se aparta. **Punta a punta sobre los cinco exports reales de `input/2026-08/`**, verificando el total de horas del archivo generado contra la suma de los exports.

## Decisiones tomadas y su razón

| Decisión | Razón |
|----------|-------|
| Mapeo explícito en YAML en vez de derivar los nombres del texto de Kimai | Lo que ve el cliente no debe depender de cómo un dev tipeó el proyecto |
| Match por código exacto, no difuso | Un match difuso imputa horas al cliente equivocado silenciosamente |
| Un solo archivo por mes, detalle plano | Es lo que el partner pasó a pedir: factura sobre el detalle, no sobre el pivot |
| El formato anterior queda en el repo sin ejecutarse | Si el partner vuelve atrás, se recupera sin reescribir; borrarlo no ahorra nada |
| Con exports fallados el archivo se genera igual, pero marcado en el nombre | No generarlo deja al dueño sin nada; generarlo limpio le deja mandar un archivo incompleto sin saberlo |
| Verificación de integridad que frena, no que avisa | Con todo en un archivo, una fila perdida no se ve a ojo: un aviso se leería tarde |
| El mapeo pasa de obligatorio a opcional | Cada fila lleva su `Project number`: la trazabilidad no depende del mapeo, y frenar por un nombre cuesta un mes de atraso |
| `Duration` como `timedelta` y no como número | El partner totaliza esa columna: con `[hh]:mm` da `160:30`, con número daría `160,5` |
| La plantilla como fuente de estilos, no como archivo a mutar | Filas y columnas son variables; insertar filas en openpyxl rompe estilos |
| Sin filtro de registros | Decisión del usuario: todo lo cargado en Kimai va al Excel |
| Validaciones que avisan y no frenan | El usuario decide si envía o le reclama al dev; frenar agrega fricción |
| Encabezados de día como texto | Como fechas reales, Excel podría mostrárselas distinto al partner según su configuración regional |
| Parser XML propio para el input | openpyxl no puede abrir los exports de Kimai |

## Fuera de alcance (posible trabajo futuro)

- Integración directa con la API de Kimai (elimina el paso manual de exportar).
- Comparación mes a mes / detección de desvíos.
