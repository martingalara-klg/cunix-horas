# cunix-horas

Convierte los exports de Kimai en **lo que KLG le entrega a C.UNIX todos los
meses**: los dos anexos, listos para revisar y enviar.

```
input/AAAA-MM/*.xlsx|*.csv        exports de Kimai, uno por desarrollador
input/AAAA-MM/manual/*.xlsx       horas de quien no está en Kimai, y
                                  descripciones que alguien entrega aparte
        |
        v
python -m cunix_horas AAAA-MM
        |
        v
output/AAAA-MM/  Anexo II (.docx) + Anexo II-A (.xlsx) + _validacion.txt
```

- **Anexo II** — el informe mensual en Word: horas e importe por proyecto,
  horas por persona, principales trabajos y observaciones.
- **Anexo II-A** — el detalle en Excel: una fila por registro de tiempo, con
  las columnas de control y revisión que C.UNIX completa después.

Los dos salen de las plantillas vacías que mandó C.UNIX, que viven en
`templates/` y **nunca se modifican**: se trabaja sobre copias.

---

## Qué hacer cada mes

Esto es todo. Si vas a leer una sola sección, es ésta.

### 1. Exportar de Kimai, un archivo por persona

En Kimai: **Horas → filtrar el mes completo y una sola persona → Exportar a
Excel**, y repetir con cada desarrollador.

Conviene el **reporte de detalle**: trae la hora de inicio y la descripción de
cada registro. El **resumen mensual** también sirve, pero sus filas salen sin
descripción y sin horario, y el informe de validación te lo avisa.

Exportar de a una persona no es un capricho: si un archivo falla, así sabés
cuál es.

### 2. Dejar los archivos en la carpeta del mes

```
input/2026-09/
```

El nombre de cada archivo da igual. La carpeta se crea a mano si no existe.

Si alguien del equipo **no tiene usuario de Kimai**, o **entregó sus
descripciones en una planilla aparte**, esos archivos van en
`input/2026-09/manual/` y además tienen que estar declarados en
`config/mapeo.yaml` (ver «Las planillas que no salen de Kimai», más abajo).

### 3. Correr la herramienta

Doble clic en **`generar.bat`** y escribir el mes cuando lo pida, o desde una
terminal:

```
python -m cunix_horas 2026-09
```

### 4. Leer `output/2026-09/_validacion.txt` ANTES de mandar nada

Describe **lo que quedó en la carpeta** —que es lo que vas a adjuntar— y
alcanza por sí solo para decidir. Lo primero que dice es **quiénes entraron a
los anexos**: contá esa lista contra tu equipo. Después vienen, si corresponde,
los archivos que no entraron y por qué, los archivos de la carpeta que esta
corrida no generó, los proyectos sin valor hora, los registros sin descripción
y los avisos de carga (más de 12 h en un día, horas en fin de semana, días
hábiles sin carga, registros fuera del mes).

### 5. Completar a mano lo que la herramienta no puede saber

Abrí el **Anexo II** (el Word) y completá:

| Qué | Dónde |
|---|---|
| **Principales trabajos** | Tabla 3. Sale un **borrador** agrupado por proyecto y ticket, con las horas sumadas y las descripciones pegadas una atrás de otra. Los totales cierran; la redacción es mecánica. Reescribila, y si fusionás líneas sumá sus horas. |
| **Estado al cierre** | Tabla 3, última columna: `Terminado`, `En curso`, `En revisión de C.UNIX` o `Bloqueado`. Sale con la marca `[●]`. |
| **Observaciones** | Tabla 4, las tres filas. Salen con `[●]`. Si no hay nada que decir, «Sin novedades». |
| **Firmas** | Nombre, cargo y fecha de quien emite. |
| **Contrato de fecha y fecha de emisión** | Salen como `[DD/MM/AAAA]` mientras no estén en `config/mapeo.yaml`. |

Todo lo que quede con `[●]` o con `[DD/MM/AAAA]` es, justamente, lo que falta
completar: está a la vista a propósito.

### 6. Enviar los dos archivos

La entrega del mes son **los dos anexos juntos**. Si alguno salió con
`(INCOMPLETO - ... - NO ENVIAR)` en el nombre, no lo mandes: corregí lo que
dice `_validacion.txt` y volvé a correr.

---

## Lo que sale solo y lo que no

**Automático, de los datos:**

- la hoja `Detalle` completa, una fila por registro;
- la hoja `Datos`: el período y el equipo (persona, perfil, usuario de Kimai);
- la tabla 1 del informe (horas e importe por proyecto): las horas salen de
  los datos, el valor hora de la hoja `Datos` de la plantilla, y el importe se
  calcula;
- la tabla 2 (horas por persona): horas, días con registro y promedio;
- el período en el encabezado del informe.

**Borrador, para que lo edites:** la tabla 3. La agrupación por proyecto y
ticket es mecánica y **los subtotales cierran exactos**, pero decidir qué
trabajos cuentan la misma historia es criterio editorial y no se automatiza.

**Sin tocar, para completar a mano:** la tabla 4 (observaciones), el plazo de
entrega de la tabla 0 si no está configurado, las dos fechas del encabezado y
la tabla de firmas. Las columnas de revisión de C.UNIX del Anexo II-A
(`Revisión C.UNIX`, `Horas aprobadas`, `Observación C.UNIX`) también quedan
vacías: las completa él.

---

## Las planillas que no salen de Kimai

Van en `input/AAAA-MM/manual/` y se declaran en `config/mapeo.yaml`, bajo
`fuentes_manuales:`. Si no hay ninguna, no hace falta tocar nada.

### Alguien trabajó sin usuario de Kimai

Pasa cuando un refuerzo entra al proyecto antes de que C.UNIX le dé el alta:
sus horas quedan cargadas en la cuenta de otra persona. La planilla dice
cuántas horas hizo cada día, y la herramienta **se las resta día por día** a la
otra persona.

```yaml
fuentes_manuales:
  horas_sin_kimai:
    - persona: "Gabriel Denis"
      planilla: "gabriel-denis.xlsx"
      restar_a: "Alexis Carnero"
      proyecto: "SELICO"
```

La planilla se lee de forma tolerante: alcanza con que tenga una fila de
encabezados con **Fecha** y **Horas** (y conviene **Descripción**), en
cualquier hoja y en cualquier orden. La fecha puede ser una fecha de Excel,
`dd/mm/aaaa`, `aaaa-mm-dd` o texto en español sin año (`Dom 02 Ago`). Las filas
de subtotal y de total se saltean solas. Si aun así no se entiende, el error
dice qué se esperaba encontrar.

**Si la resta no cierra —un día en que la otra persona no tiene horas, o no le
alcanzan— no se genera nada.** Un reparto que no cuadra se factura mal y no se
ve en una hoja de ochenta filas.

**Cuando esa persona tenga usuario de Kimai, borrá su entrada de
`config/mapeo.yaml`.** Si no, sus horas se restarían dos veces (y la
herramienta frena avisándolo).

### Alguien entregó sus descripciones aparte

Es la excepción: desde septiembre de 2026 todos cargan la descripción en
Kimai. La planilla tiene las mismas columnas que la hoja `Detalle` y **sólo
aporta la descripción**; el día y las horas se verifican contra lo que ya está,
y si no coinciden tampoco se genera nada.

```yaml
fuentes_manuales:
  descripciones:
    - persona: "Alexis Carnero"
      planilla: "alexis-carnero.xlsx"
```

---

## Qué hay que configurar en `config/mapeo.yaml`

Todo es opcional. Un mes normal no obliga a tocar nada.

```yaml
anexos:
  perfil_por_defecto: "Desarrollador"
  dias_habiles_entrega: 5          # el plazo de la tabla de condiciones
  contrato_de_fecha: "01/01/2026"  # si no está, queda [DD/MM/AAAA]
  fecha_de_emision: "05/10/2026"   # idem

proyectos:
  PR2510126:                       # el código entre corchetes de Kimai
    cliente: "MINVU"
    proyecto: "SELICO"             # IGUAL que en la hoja «Datos»
```

Dos detalles que cuestan plata si se pasan por alto:

- el **nombre del proyecto** tiene que escribirse igual acá y en la hoja
  `Datos` de `templates/Anexo-II-A-Detalle-horas-KLG.xlsx`: es por ese nombre
  que el informe le encuentra el **valor hora**. Si no lo encuentra, el
  proyecto sale con sus horas y sin importe, el total a facturar queda vacío, y
  `_validacion.txt` lo nombra;
- el **valor hora de cada proyecto** se carga una sola vez, a mano, en esa hoja
  `Datos` de la plantilla (columnas Proyecto / Cliente / Valor hora).

Un proyecto que no esté en `mapeo.yaml` **no frena nada**: sale con el nombre
que trae Kimai y queda listado en `_validacion.txt` con el bloque listo para
pegar.

La sección `personas:` sólo hace falta para quien exporta con el **resumen
mensual**, que no trae el usuario: ahí la persona se resuelve por su `nombre:`,
que tiene que coincidir exactamente con lo que muestra Kimai. Cada persona
lleva un solo dato, `nombre:`, y su clave es el usuario de Kimai:

```yaml
personas:
  acarnero:                        # el usuario de Kimai
    nombre: "Alexis Carnero"       # exacto como lo muestra Kimai
```

Si no está declarada, **tampoco frena nada**: esa persona entra con todas sus
horas y la columna «Usuario Kimai» de la hoja `Datos` le queda vacía, igual que
a quien todavía no tiene usuario. `_validacion.txt` la nombra y explica que C.UNIX
pide que cada uno cargue sus horas con su propio usuario.

### Cinco campos que ya no existen

Si tenés un `mapeo.yaml` de antes, puede traer `mail:`, `numero_proyecto:`,
`archivo:`, `actividad:` o `archivo_salida:`. **Los cinco se ignoran al cargar:
podés borrar esas líneas.** Eran de los dos entregables anteriores —el Excel
pivoteado por desarrollador y el archivo plano `Horas KLG-<Mes><Año>.xlsx`—,
que C.UNIX dejó de recibir.

Los anexos de hoy no tienen dónde escribir nada de eso: el Anexo II-A va
`Fecha | Inicio | Fin | Persona | Proyecto | Descripción | Horas | Alertas |
Revisión C.UNIX | Horas aprobadas | Horas a pagar | Observación C.UNIX | Día
nuevo | Nota KLG`. `mail:` y `numero_proyecto:`, además, **frenaban la entrega
entera** si faltaban: trababan un mes por un dato que después no se escribía en
ningún lado.

## Qué frena una entrega y qué sólo se avisa

**Frena** (no se genera nada, o ese export no entra):

- un export que no se puede leer o que no tiene formato de Kimai → **no entra
  ese archivo**, los demás sí, y los anexos salen marcados como INCOMPLETO;
- una planilla manual que falta, o cuyas horas no cierran contra la persona a
  la que se le restan (un día negativo, o un día que la otra persona no tiene)
  → **no se genera ningún anexo**;
- el Anexo II-A escrito cuyas horas no coinciden con las de los exports →
  **no se genera ese archivo**;
- un mes que no entra en la plantilla (más de 600 registros, o más de 15
  personas) → **no se genera ese archivo**, y hay que pedirle a C.UNIX una
  plantilla más grande;
- un anexo que tenés abierto y no se puede escribir → **no se genera ese
  archivo**; cerralo y volvé a correr;
- `config/mapeo.yaml` roto o con un campo declarado en blanco;
- que falte una de las dos plantillas de `templates/`, o la carpeta
  `input/<mes>/`.

**Sólo avisa** en `_validacion.txt`, y la entrega sale igual:

- una persona del resumen mensual que no está en `personas:`, o cuyo `nombre:`
  está declarado en dos entradas → usuario de Kimai vacío;
- un proyecto que no está en `proyectos:` → sale con el nombre que trae Kimai;
- un proyecto sin valor hora en la hoja `Datos` de la plantilla → sale con sus
  horas y sin importe;
- registros sin descripción, días hábiles sin carga, horas en fin de semana,
  más de 12 h en un día, registros fuera del mes.

La regla es una sola: **frena lo que haría salir mal el entregable; avisa todo
lo que el dueño puede mirar y decidir.**

---

## Si falta un desarrollador

Este es el riesgo del entregable y conviene tenerlo claro: con todas las
personas en dos documentos, **una que falta es invisible adentro**.

Por eso, cuando algún export falla:

- los anexos **se generan igual** con los que sí se pudieron leer, para que
  puedas revisar lo que hay;
- pero salen con el nombre
  `Anexo-II-A-Detalle-horas-KLG-2026-09 (INCOMPLETO - FALTAN 2 DESARROLLADORES - NO ENVIAR).xlsx`,
  que es lo único que se ve al adjuntarlos a un mail;
- y si en la carpeta había un anexo limpio de una corrida anterior, se lo
  renombra a `... (CORRIDA ANTERIOR - NO ENVIAR).xlsx` para que no se envíe en
  su lugar. **No se borra**: el dato sigue ahí.

La herramienta nunca borra archivos tuyos.

## La verificación de integridad

Después de escribir el Anexo II-A, la herramienta lo **vuelve a abrir** y suma
la columna `Horas`. Ese total tiene que coincidir exactamente con las horas de
los exports que leyó. Si no coincide, **no genera el archivo** y lo dice
fuerte: no es un aviso. Con ochenta filas de seis personas en una sola hoja,
una fila perdida no se ve nunca a ojo.

Cada anexo se escribe primero en un temporal y recién cuando salió entero se
mueve sobre el nombre final, así que nunca queda un archivo a medio escribir.

## Qué dice `_validacion.txt`

Describe **la carpeta**, no sólo la corrida: no se envía lo que hizo el
programa, se envía lo que hay en `output/<mes>/`.

- **Quiénes entraron**, con cuántos registros y cuántas horas cada uno.
- **Qué archivos NO entraron y por qué**, con el bloque listo para pegar
  cuando el motivo es de configuración.
- **Qué se generó** y qué queda para completar a mano.
- **Qué `.xlsx` y `.docx` hay en la carpeta que esta corrida NO generó.** Esos
  no van en el envío del mes.
- **Los proyectos sin mapear** y **los proyectos sin valor hora**.
- **Las personas que van sin usuario de Kimai** en la hoja `Datos`, con el
  bloque listo para pegar: C.UNIX pide que cada uno cargue con su propio
  usuario, así que esa celda vacía queda explicada. No frena nada.
- **Los registros sin descripción**, por persona y con la cantidad: el
  contrato de C.UNIX pide la descripción para aprobar esas horas.
- **Los avisos de carga**, agrupados por desarrollador: días hábiles sin
  carga, horas en fin de semana, más de 12 h en un día, registros fuera del
  mes.

Si movés o agregás archivos a la carpeta *después* de correr, el informe ya no
corresponde: volvé a correr.

Si `_validacion.txt` no se puede escribir (lo más común: lo tenés abierto), el
informe de esta corrida va a
`_validacion (NO SE PUDO ESCRIBIR _validacion.txt - LEER ESTE).txt`, se avisa
en consola y la corrida termina con error.

## Qué exports de Kimai lee

Según con qué reporte exportes, Kimai da un archivo distinto. La herramienta
lee los tres y los detecta sola.

- **Timesheet en Excel** (`.xlsx` con `Date` en A1): el reporte de detalle, el
  que conviene usar.
- **Timesheet en CSV** (`.csv`): lo mismo, en texto.
- **Resumen mensual** (`.xlsx` con `Total` en B1): la grilla de días, con las
  horas ya sumadas. Sirve, pero sus filas van sin descripción y sin horario.

Si un archivo no es ninguno de esos tres, **ese** desarrollador no entra, el
motivo dice qué encontró en la fila 1, y los demás entran igual.

## Tests

```
python -m pytest -q
```

## Estructura

```
config/mapeo.yaml          configuración editada a mano
templates/                 las plantillas vacías de los dos anexos de C.UNIX
input/AAAA-MM/             exports de Kimai, .xlsx o .csv (se versionan)
input/AAAA-MM/manual/      planillas que NO salen de Kimai
output/AAAA-MM/            lo que se entrega (se versiona con `git add -f`)
cunix_horas/               el código
scripts/                   los scripts de un mes concreto
tests/                     los tests
docs/superpowers/specs/    el diseño: qué hace, y por qué está hecho así
```

### `input/2026-08/manual/`

De agosto de 2026 hay cuatro archivos ahí:

| Archivo | Qué es |
|---|---|
| `gabriel-denis.xlsx` | Las 91 h de Gabriel Denis, que trabajó sin usuario de Kimai y cuyas horas quedaron cargadas en la cuenta de Alexis Carnero. |
| `alexis-carnero.xlsx` | El detalle diario que entregó Alexis Carnero después del reclamo de C.UNIX. Aporta las descripciones que Kimai nunca registró. |
| `anexo-II-informe-cunix.docx` | El Anexo II tal como lo mandó C.UNIX, con agosto cargado como ejemplo. |
| `anexo-II-A-detalle-cunix.xlsx` | Lo mismo, el Anexo II-A. De acá salieron las plantillas vacías de `templates/`. |

Son la **única copia** de esos datos: sin ellas agosto de 2026 no se puede
volver a generar. Por eso están versionadas.

### `scripts/anexos_agosto_2026.py`

Agosto de 2026 se entregó **corrigiendo el ejemplo que mandó C.UNIX**, no
armándolo desde cero. Ese script es el registro auditable de esa corrección y
lo único que reproduce lo entregado tal cual.

```
python scripts/anexos_agosto_2026.py
```

**Escribe el Anexo II-A con el mismo nombre que usa la herramienta**, así que
correrlo pisa lo que haya generado `python -m cunix_horas 2026-08` en esa
carpeta. Es a propósito: lo que KLG entregó en agosto es lo que sale de ahí.
De septiembre en adelante no hace falta para nada.
