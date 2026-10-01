# cunix-horas — Diseño del sistema

**Fecha:** 2026-09-24
**Última revisión:** 2026-10-01 — reescrita para describir el sistema tal como es hoy: los dos anexos de C.UNIX.
**Estado:** Implementado y en uso.

Este documento describe **lo que la herramienta hace hoy** y **por qué está
hecha así**. La fuente de verdad es el código (`cunix_horas/`); esta spec
existe para que las decisiones que costaron caro no haya que volver a
descubrirlas.

Para usarla todos los meses, el documento es el `README.md`. Éste es el de
diseño.

> El plan de implementación original (`docs/superpowers/plans/2026-09-24-cunix-horas.md`)
> describía paso a paso la construcción del **primer** entregable, un Excel
> pivoteado por desarrollador que ya no existe. Quedó obsoleto entero y **se
> eliminó**; el historial de git lo conserva. Las decisiones suyas que siguen
> vigentes están rescatadas más abajo.

---

## 1. El problema

**KLG** es una software factory. **C.UNIX** es su cliente. Todos los meses KLG
le tiene que entregar dos documentos con las horas que su equipo trabajó, en
el formato que C.UNIX definió:

- el **Anexo II** — informe mensual, Word: horas e importe por proyecto, horas
  por persona, principales trabajos y observaciones;
- el **Anexo II-A** — detalle, Excel: una fila por registro de tiempo, con las
  columnas de control y de revisión que C.UNIX completa después.

Los desarrolladores cargan sus horas en **Kimai**. Armar los dos anexos a mano
desde los exports de Kimai es el trabajo que esta herramienta hace.

### Objetivo

Un comando (`python -m cunix_horas AAAA-MM`, o doble clic en `generar.bat`)
que deje en `output/AAAA-MM/` los dos anexos listos para revisar, más un
informe de validación que alcance por sí solo para decidir si se envían.

### Alcance

**Entra:** lectura de los tres formatos de export de Kimai, lectura de las
planillas manuales del mes, mapeo opcional de códigos a nombres, generación de
los dos anexos sobre las plantillas de C.UNIX, verificación de integridad y el
informe de validación.

**No entra:** integración con la API de Kimai, envío de mails, facturación,
tarifas (el valor hora lo carga el dueño a mano en la plantilla), interfaz
gráfica.

---

## 2. Qué entra

```
input/AAAA-MM/*.xlsx  *.csv     exports de Kimai, uno por desarrollador
input/AAAA-MM/manual/*.xlsx     lo que NO sale de Kimai
config/mapeo.yaml               configuración, toda opcional
templates/*.xlsx  *.docx        las plantillas vacías de C.UNIX
```

### 2.1 Los tres formatos de export de Kimai

Según con qué reporte se exporte, Kimai da un archivo distinto. Se aceptan los
tres. `lector_kimai.leer()` es un despachador: mira el archivo, elige el lector
y devuelve `list[Registro]`, así que el resto del pipeline no se entera del
formato.

| Se reconoce por | Formato | Lector |
|---|---|---|
| extensión `.csv` | timesheet plano en CSV | `lector_timesheet_csv.py` |
| `.xlsx` con `A1='Date'` | timesheet plano en XLSX | `lector_timesheet_xlsx.py` |
| `.xlsx` con `B1='Total'` | resumen mensual (la grilla de días) | `lector_resumen_mensual.py` |

**Ante cualquier otra cosa se falla, nunca se adivina:** el error nombra el
archivo, dice qué encontró en la fila 1 y enumera los formatos que se
reconocen. Como todo fallo de lectura, es de *ese* archivo: los demás
desarrolladores entran igual.

#### Timesheet plano `.xlsx`

Fila 1 = encabezados, una fila por registro de tiempo. Columnas relevantes:

| Col | Campo | Ejemplo | Uso |
|-----|-------|---------|-----|
| A | Date | `46265.708333333` | Serial de Excel, época `1899-12-30` |
| B | From | `13:00` | Hora de inicio → columna `Inicio` del anexo |
| D | Duration | `0.14583333` | Fracción de día; × 24 = horas |
| E | Name | `Matias Zalazar` | Nombre con el que se agrupa a la persona |
| F | User | `mzalazar` | Usuario de Kimai; clave de `personas:` |
| G | E-mail | `matias.zalazar@cunix.net` | Se lee, no se escribe en ningún lado |
| I | Customer | `[608040005] Servicio Nacional de Aduanas` | Texto crudo del cliente |
| J | Project | `[CO2610170] Aduana-Subastas \| …` | El código entre corchetes es la clave de `proyectos:` |
| K | Activity | `Desarrollo` | Se lee; hoy ninguna tabla lo escribe |
| L | Description | `Ticket R-012528` | **Vacía es válido** |
| R | Project number | `210` | Se lee, no se escribe en ningún lado |

Este lector toma las columnas **por posición**, porque es lo que este export
viene emitiendo desde siempre. Por eso el encabezado se verifica antes de leer
nada: si una columna está y no dice lo que tiene que decir, se falla en vez de
poner el dato equivocado en cada columna. Si directamente no está (un export
viejo), ese campo queda vacío y las horas se leen igual.

Una fila **sin fecha pero con datos** no se saltea: son horas que
desaparecerían sin que nadie se entere. Se falla nombrando archivo y fila. Una
fila completamente vacía sí se saltea.

**No se aplica ningún filtro:** todo registro del export entra al anexo.

#### Timesheet plano `.csv`

El mismo contenido, con dos diferencias: la fecha viene en ISO (`2026-08-31`) y
la duración en `H:MM` (`2:00`). UTF-8, puede traer BOM.

Acá las columnas se leen **por nombre de encabezado, no por posición**: Kimai
agrega y reordena columnas de una versión a otra, y leer por posición haría que
un export nuevo imputara horas equivocadas sin avisar.

Tres cosas se detectan por archivo en vez de asumirse, porque asumirlas hacía
desaparecer horas sin ruido:

1. el **separador de columnas**, que según el locale de Kimai es `,` o `;`;
2. los **encabezados repetidos**: con dos columnas `Duration` ganaba la última;
3. las **filas sin fecha**: si traen datos, son horas que no se facturan.

Los encabezados se normalizan (se les saca el espacio de los bordes) **una sola
vez**, y tanto la verificación como la lectura usan esas mismas claves. Antes
la verificación normalizaba y la lectura no, así que un archivo con `" Date "`
pasaba la verificación y devolvía cero registros sin error.

#### Resumen mensual `.xlsx`

Es la grilla de días: fila 1 con el nombre del dev, `Total` y un encabezado por
día; después una fila por cliente, una por proyecto y una por actividad; y una
fila `Total` al final.

```
A1='Lautaro Zalazar'  B1='Total'  C1='1/8/2026'  D1='2/8/2026' ...
A2='[763043541] C.UNIX'                          B2='9,00'                 <- CLIENTE (días mergeados)
A3='[GI2680001] C.UNIX - Proyectos Internos ...' B3='9,00'  E3='0,50' ...  <- PROYECTO
A4='Coordinación interna'                        B4='9,00'  E4='0,50' ...  <- ACTIVIDAD
A5='Total'                                       B5='9,00'  ...
```

Tres cosas se resuelven acá, y ninguna se asume:

1. **Qué es cada fila**, por estructura y no por posición: la de cliente tiene
   las celdas de día mergeadas, la de proyecto trae el código entre corchetes y
   las horas por día, y la de actividad viene abajo. **Sólo las filas de
   actividad emiten `Registro`**: las de proyecto son subtotales y duplicarían
   las horas.
2. **El orden de la fecha del encabezado.** Como están todas las columnas del
   mes, uno de los dos componentes es constante y ése es el mes: `1/8/2026`…
   `31/8/2026` es día/mes; `8/1/2026`…`8/31/2026` es mes/día. Si no se puede
   decidir sin ambigüedad, se falla: adivinar mal movería las horas de día y de
   mes sin que nadie lo note.
3. **El separador decimal**, que puede ser coma o punto. Si el archivo mezcla
   los dos, se falla: `"2.00"` leído con coma da 200.

Y hay una **red de seguridad que los otros formatos no permiten**: el archivo
declara sus propios totales en la fila `Total` —el general en la columna B y
uno por día en cada columna—. Se comparan los dos contra lo parseado: el
general ataja las horas que se perdieron, y los de día atajan las que se
leyeron en la columna equivocada.

**El resumen mensual sirve, pero no trae todo.** Le faltan la hora de inicio,
el usuario y la descripción.

| Lo que falta | Qué se hace |
|---|---|
| `Description` | se emite vacía, y el informe cuenta cuántas por persona |
| hora de inicio | `Inicio` y `Fin` quedan vacías, y el informe lo avisa |
| usuario de Kimai | sale de `config/mapeo.yaml`; si no está, **vacío con aviso** |

El usuario es **lo único que el resumen no trae y que alguno de los dos anexos
sí escribe** (la columna «Usuario Kimai» de la hoja `Datos`). Lo completa
`completado.py`, **nunca adivinando**: la persona se resuelve por su `nombre:`,
que es lo único que el resumen trae, y el usuario es la clave de esa entrada en
`personas:`.

**Para los registros del reporte de detalle el mapeo no interviene:** su
usuario es el de Kimai, aunque el mapeo declare otra cosa. El mapeo es un
respaldo para lo que la fuente no trae, no una corrección de lo que sí trae.
`Registro.origen` es lo que distingue los dos casos, y lo pone el lector.

### 2.2 Las fuentes manuales — `input/AAAA-MM/manual/`

Son las planillas que no salen de Kimai. Se declaran en `config/mapeo.yaml`
bajo `fuentes_manuales:` y las dos son opcionales: un mes sin ninguna corre
igual. El CLI no entra solo en esa subcarpeta —recorre `input/<mes>/` con
`glob`, que no es recursivo—: las lee `fuentes_manuales.py`, y sólo las que
estén declaradas.

**a) Horas de alguien que no está en Kimai** (`horas_sin_kimai:`). La persona
trabajó pero no tiene usuario, así que sus horas quedaron cargadas en la cuenta
de otro. La planilla dice cuántas horas hizo cada día; esas horas se le **restan
día por día** a la otra persona y pasan a filas propias, con su explicación en
la columna `Nota KLG`.

**b) Descripciones que alguien entrega aparte** (`descripciones:`). La planilla
trae las mismas columnas que la hoja `Detalle` y **sólo aporta la
descripción**: el día y las horas se verifican contra lo que ya está. Desde
septiembre de 2026 todos cargan la descripción en Kimai, así que esto es la
excepción, no el camino normal.

**Las planillas se leen de forma tolerante.** No se asume ni el nombre de la
hoja, ni el número de fila del encabezado, ni el formato de la fecha: el
encabezado se busca por el nombre de las columnas, y la fecha se acepta como
fecha de Excel, `dd/mm/aaaa`, `aaaa-mm-dd` o texto en español sin año
(`Dom 02 Ago`). Las filas de subtotal y de total se saltean solas. Si aun así
no se entiende, el error dice qué se esperaba encontrar y qué se encontró en su
lugar. El motivo es que estas planillas las arma una persona, no un programa:
exigirles una forma exacta sería garantizar que fallen.

### 2.3 Las plantillas — `templates/`

Las dos plantillas **vacías** que mandó C.UNIX:
`Anexo-II-A-Detalle-horas-KLG.xlsx` y
`Anexo-II-Informe-mensual-horas-KLG.docx`.

Además de dar la estructura, la del Anexo II-A **es la fuente del valor hora de
cada proyecto**: el dueño lo carga una sola vez en su hoja `Datos` (columnas
Proyecto / Cliente / Valor hora, filas 8 a 22) y el informe lo lee de ahí. Los
valores hora **se leen, nunca se escriben**.

### 2.4 La configuración — `config/mapeo.yaml`

**Todo el archivo es opcional.** Un mes normal no obliga a tocar nada.

```yaml
anexos:
  archivo_detalle: "Anexo-II-A-Detalle-horas-KLG-{periodo}.xlsx"
  archivo_informe: "Anexo-II-Informe-mensual-horas-KLG-{periodo}.docx"
  perfil_por_defecto: "Desarrollador"
  perfiles: { "Nombre Apellido": "Lider tecnico" }
  dias_habiles_entrega: 5
  contrato_de_fecha: "01/01/2026"
  fecha_de_emision: "05/10/2026"

fuentes_manuales:
  horas_sin_kimai:
    - persona: "Gabriel Denis"
      planilla: "gabriel-denis.xlsx"
      restar_a: "Alexis Carnero"
      proyecto: "SELICO"
  descripciones: []

personas:
  acarnero: { nombre: "Alexis Carnero" }   # clave = usuario de Kimai

proyectos:
  CO2610170: { cliente: "Servicio Nacional de Aduanas", proyecto: "Subastas" }
```

- **El único campo obligatorio de una persona es `nombre:`**, y es el único que
  existe. La clave de la entrada *es* su usuario de Kimai. `personas:` sólo le
  aporta algo a quien exporta con el resumen mensual; para el reporte de
  detalle no interviene.
- **Un proyecto sin declarar no frena nada:** sale con el nombre derivado del
  texto de Kimai y queda listado en `_validacion.txt` con el bloque YAML listo
  para pegar. El mapeo es un pulido opcional de nombres, no un requisito.
- La clave de `proyectos:` es el **código entre corchetes** de la columna
  `Project` (`[CO2610170]` → `CO2610170`), con **match exacto, nunca por
  similitud**: el código es estable aunque alguien renombre el proyecto en
  Kimai, y un match difuso imputaría horas al cliente equivocado sin que nadie
  lo note.
- El **nombre del proyecto** tiene que escribirse **igual** acá y en la hoja
  `Datos` de la plantilla: es por ese nombre que el informe le encuentra el
  valor hora.
- El cliente se declara **por proyecto**, no en una sección aparte: así se
  pueden agrupar o separar en el anexo proyectos que en Kimai cuelgan de
  clientes distintos.
- `dias_habiles_entrega`, `contrato_de_fecha` y `fecha_de_emision` son textos
  del informe. Si no están, **el documento sale con el marcador a la vista**
  (`[●]`, `[DD/MM/AAAA]`) para que el dueño los complete. Nunca se inventan.
- Un campo **declarado y en blanco sí frena**: lo que quedó escrito a medias
  conviene que falle ahora, con el nombre de la entrada, y no que salga vacío
  en un anexo que ya va camino a C.UNIX.
- Los patrones de nombre de archivo se validan **al cargar el mapeo**, antes de
  leer ningún export: que el único reemplazo sea `{periodo}` y que la extensión
  sea la que corresponde. Un patrón roto tiene que fallar antes de procesar
  todo, no después.

---

## 3. Qué sale

```
output/AAAA-MM/
  Anexo-II-A-Detalle-horas-KLG-AAAA-MM.xlsx
  Anexo-II-Informe-mensual-horas-KLG-AAAA-MM.docx
  _validacion.txt
```

### 3.1 Anexo II-A (detalle, Excel)

Catorce columnas:
`Fecha | Inicio | Fin | Persona | Proyecto | Descripción | Horas | Alertas |
Revisión C.UNIX | Horas aprobadas | Horas a pagar | Observación C.UNIX |
Día nuevo | Nota KLG`.

- **A–G las escribe la herramienta**, una fila por registro, ordenadas
  cronológicamente y, dentro del día, por persona.
- **H–M son de C.UNIX**: tres son fórmulas que ya vienen escritas en las 600
  filas del rango de la plantilla —alcanza con no pisarlas— y tres las completa
  él al revisar.
- **N (`Nota KLG`)** la agrega la herramienta a la derecha de todo, y **sólo
  cuando alguna fila la necesita** (hoy, las horas separadas de una persona sin
  usuario de Kimai). Va fuera de todos los rangos de fórmulas de C.UNIX a
  propósito.
- La hoja `Datos` recibe el período, el equipo (persona, perfil, usuario de
  Kimai) y, si está configurada, la fecha del contrato. Un usuario vacío se
  escribe como celda **vacía, no como `""`**: el anexo se compromete a que esa
  celda vacía signifique «esta persona todavía no tiene usuario de Kimai».
- Las hojas `Instrucciones` y `Resumen` quedan intactas: son de C.UNIX.

`Fin` **no** lo trae Kimai: se calcula como `Inicio + Horas`. Es aritmética
sobre datos registrados, no una reconstrucción —Kimai guarda cada registro como
un bloque continuo—. Cuando el export no trae hora de inicio, `Inicio` y `Fin`
quedan vacíos, que es lo que la plantilla espera.

**Capacidad:** la hoja `Detalle` llega hasta la fila 601 y la hoja `Datos`
tiene 15 lugares de equipo. Un mes que no entre **no se genera**: escribir de
más dejaría registros fuera de todas las fórmulas del anexo y el total no
cerraría. El mensaje dice que hay que pedirle a C.UNIX una plantilla más
grande.

### 3.2 Anexo II (informe, Word)

| Qué | Cómo sale |
|---|---|
| Encabezado, período | automático |
| Fecha de emisión | configurable; si no, queda `[DD/MM/AAAA]` |
| Contrato de fecha | la completa C.UNIX: queda `[DD/MM/AAAA]` a propósito |
| Tabla 0, plazo de entrega | configurable; si no, queda `[●]` |
| Tabla 1, horas e importe por proyecto | automático |
| Tabla 2, horas por persona, días y promedio | automático |
| Tabla 3, principales trabajos | **borrador mecánico**, para editar |
| Tabla 4, observaciones | se deja con `[●]`: la escribe el dueño |
| Tabla 5, firmas | intacta |

Las tres tablas que se llenan se arman **clonando la fila modelo** que dejó la
plantilla, así cada fila nueva hereda sus bordes, su tipografía y su
alineación, y el documento sigue siendo el de C.UNIX.

El párrafo «EJEMPLO: agosto 2026…» con el que C.UNIX mandó el formato **se
quita del documento generado**: es la nota de quien envió la plantilla y no
corresponde en un informe que entrega KLG. En la plantilla queda.

Las tres tablas salen de **las mismas filas** que el Anexo II-A, para que
informe y detalle no puedan discrepar: no hay dos caminos de cálculo.

Si a algún proyecto le falta el valor hora, **el total a facturar se deja
vacío** en vez de escribir una suma parcial: antes que un número creíble y
bajo, un hueco que el informe de validación explica.

### 3.3 `_validacion.txt`

El informe que el dueño lee antes de mandar nada. Tiene que alcanzar por sí
solo para decidir. Arriba de todo va **quién entró a los anexos**, con
registros, horas y de qué reporte salió cada uno: con todas las personas en un
solo par de documentos, esa lista es lo único que distingue un envío completo
de uno que no lo es. Después: qué archivos no entraron y por qué, qué se
generó, qué hay que completar a mano, qué archivos de la carpeta **no** son de
esta corrida, los proyectos sin mapear y sin valor hora, las personas sin
usuario de Kimai, los registros sin descripción y los avisos de carga por
persona.

---

## 4. El pipeline

```
input/AAAA-MM/*.xlsx *.csv
   |
   v
lector_kimai        detecta el formato y delega  -> list[Registro]
   |
   v
completado          le pone al resumen mensual el usuario de Kimai del mapeo
   |
   v
filas_anexo         registros + mapeo -> filas de la hoja Detalle, ordenadas,
   |                 y de ahí las tres tablas del informe
   v
fuentes_manuales    aplica las planillas de input/<mes>/manual/
   |
   +--> escritor_anexo_detalle -> Anexo II-A (.xlsx) + verificación de integridad
   +--> escritor_anexo_informe -> Anexo II (.docx)
   +--> validador              -> avisos de carga por persona
   |
   v
cli                 orquesta todo y escribe _validacion.txt
```

### Módulos

El paquete vive en `cunix_horas/` **en la raíz del proyecto, no bajo `src/`**:
así `python -m cunix_horas` funciona sin `pip install -e .`, que es requisito
para el `.bat` de doble clic.

| Módulo | Responsabilidad |
|---|---|
| `lector_kimai.py` | Despachador: detecta el formato del export y delega. |
| `kimai_comun.py` | `Registro`, `ErrorLectura` y el parseo XML del `.xlsx`, compartidos. |
| `lector_timesheet_xlsx.py` | Timesheet `.xlsx` → `list[Registro]`. Serial de fecha, duración × 24, columnas por posición. |
| `lector_timesheet_csv.py` | Timesheet `.csv` → `list[Registro]`. Fecha ISO, duración `H:MM`, columnas por nombre. |
| `lector_resumen_mensual.py` | Resumen mensual `.xlsx` → `list[Registro]`, verificados contra los totales que el archivo declara. |
| `completado.py` | Le pone a los registros del resumen mensual el usuario de Kimai del mapeo; si no lo resuelve, lo deja vacío y devuelve un aviso. Nunca frena. |
| `mapeo.py` | Carga y valida `config/mapeo.yaml`. Resuelve código → (cliente, proyecto) y persona → usuario. |
| `anexos.py` | `Periodo`, los patrones de nombre, la marca de INCOMPLETO y el formato de números. |
| `filas_anexo.py` | `list[Registro]` + mapeo → las filas de la hoja `Detalle` y las tres tablas del informe. |
| `fuentes_manuales.py` | Las planillas de `input/<mes>/manual/`: la resta día por día y la carga de descripciones. |
| `escritor_anexo_detalle.py` | Escribe el Anexo II-A sobre una copia de la plantilla. Incluye la verificación de integridad. |
| `escritor_anexo_informe.py` | Escribe el Anexo II sobre una copia de la plantilla. |
| `validador.py` | Avisos de carga por persona sobre los registros crudos. Nunca frenan. |
| `cli.py` | Orquesta, decide los nombres de salida y escribe `_validacion.txt`. |

Todos los `dataclass` son `frozen=True` y ninguna etapa muta lo que recibe:
cada una devuelve un valor nuevo. `scripts/anexos_agosto_2026.py` queda fuera
del pipeline (ver §6.11).

---

## 5. Convenciones que atraviesan todo el código

- **Todo en español**: identificadores, mensajes y nombres de archivo. La
  documentación y los errores los lee el dueño, que no es programador.
- **UTF-8 explícito** en toda lectura y escritura de texto: los nombres de
  cliente traen acentos (`Subsecretaría`, `Pública`). El CLI además reconfigura
  `stdout`/`stderr` a UTF-8, para que la consola de Windows no rompa los
  acentos aunque no se haya entrado por `generar.bat`.
- **Nunca `strftime("%B")`** para el nombre del mes: depende del locale de la
  máquina y en Windows suele salir en inglés. Hay una tabla fija en
  `anexos.MESES_ES`.
- **Los números se escriben como los escribe C.UNIX**: coma decimal y punto de
  miles (`306,0`, `3.706,00`), nunca con `str(float)`, que daría `306.0`.
- **Redondeo medio hacia arriba**, no el bancario de Python: con el bancario,
  2,25 h daría 2,2 y el promedio de una persona dejaría de coincidir con el del
  anexo del mes anterior.
- Las horas se comparan con una tolerancia de `1e-6`: las duraciones del `.xlsx`
  de Kimai llegan como fracción de día y 1 h vale `1.000000000000008`.
- Época de los seriales de fecha de Excel: `date(1899, 12, 30)`.

---

## 6. Decisiones de diseño vigentes, y por qué

### 6.1 Los `.xlsx` de Kimai no se leen con openpyxl

Kimai emite el atributo `showZeroes` en `<sheetView>` donde el esquema OOXML
define `showZeros`. `openpyxl.load_workbook()` explota con:

```
TypeError: SheetView.__init__() got an unexpected keyword argument 'showZeroes'
```

No es un bug de openpyxl: el archivo de Kimai está mal. Por eso **la lectura
del input no usa openpyxl**: se descomprime el `.xlsx` como ZIP y se parsea
`xl/worksheets/sheet1.xml` con `xml.etree.ElementTree`, resolviendo
`xl/sharedStrings.xml` y las celdas `inlineStr`. Vale para los dos formatos
`.xlsx` de entrada, así que ese parseo vive una sola vez, en `kimai_comun.py`.

La **escritura** del output sí usa openpyxl: las plantillas son archivos
normales.

### 6.2 Las plantillas son fuente de estructura y nunca se mutan

`templates/` tiene las dos plantillas vacías que mandó C.UNIX. De ellas salen
la estructura, los estilos, las fórmulas, las listas desplegables y los valores
hora. **Nunca se modifican**: se carga la plantilla y se guarda una copia en
`output/<mes>/`.

Dos motivos. Uno, que son el original de C.UNIX: si se corrompen, no hay de
dónde sacar el formato otra vez. Dos, que la alternativa —mutar un archivo e
ir corrigiéndolo— acumula el estado de todas las corridas anteriores, y una
corrida tiene que depender sólo de sus datos.

Dos cosas que openpyxl no hace solo y hay que arreglar después de guardar:

- **el bloque `extLst`** de la hoja `Detalle` —las listas desplegables de
  `Persona` y `Proyecto`—, que openpyxl no soporta y borra al guardar. Se
  reinyecta sobre el `.xlsx` ya escrito, sacándole el `xr:uid`, que sin su
  namespace deja el archivo sin abrir;
- **la verificación de integridad** (§6.4).

### 6.3 La escritura es atómica, y la herramienta nunca borra archivos del dueño

Cada anexo se escribe primero en un temporal `~tmp-<pid>-<nombre>` **en la
misma carpeta** que el destino (mover entre volúmenes no es atómico), y recién
cuando salió entero se lo mueve sobre el nombre final con `os.replace`, que es
atómico y pisa el destino. Así nunca queda un archivo a medio escribir con
nombre de entregable, y una falla durante la generación no toca el que ya
estaba. La verificación corre **sobre el temporal, antes del `os.replace`**:
lo que no cierra no llega a ocupar el nombre bueno.

Y la herramienta **no borra nada**. Si el anexo del mes sale marcado como
INCOMPLETO y en la carpeta había uno limpio de una corrida anterior, se lo
**renombra** a `… (CORRIDA ANTERIOR - NO ENVIAR).xlsx`: el nombre ya no
engaña, pero el dato sigue estando. Lo que el dueño puso en su carpeta es
suyo; lo único que la herramienta se permite es que no se haga pasar por lo de
hoy.

### 6.4 La verificación de integridad frena, no avisa

Después de escribir el Anexo II-A se lo **vuelve a abrir** y se suma la columna
`Horas`. Si el total no coincide con las horas de los exports que se leyeron,
**no se genera el archivo**.

Con ochenta filas de seis personas en una sola hoja, una fila perdida no se ve
nunca a ojo, y el archivo se vería completo sin serlo. Un aviso se leería
tarde. Es lo único que separa un entregable correcto de uno creíble y
equivocado.

### 6.5 El informe de validación describe **la carpeta**, no la corrida

**Esta distinción costó tres rondas de arreglos.** Cada vez que el informe
describía lo que había hecho la corrida, quedaba en `output/<mes>/` un archivo
viejo —un anexo de hace tres corridas, uno de un formato anterior, algo que el
dueño dejó a mano— del que nadie avisaba, y el dueño adjuntaba la carpeta.

El dueño no envía lo que hizo el programa: **envía lo que hay en
`output/<mes>/`**. Por eso el informe recorre la carpeta y nombra uno por uno
todo `.xlsx` y `.docx` que esta corrida no haya generado, separando los que
apartó la herramienta de los que ya estaban. Y por eso `_validacion.txt`
termina diciendo que si se mueven o agregan archivos *después* de correr, el
informe ya no corresponde y hay que volver a correr.

Del mismo principio sale el resto:

- si `_validacion.txt` no se puede escribir (lo más común: está abierto en el
  Bloc de notas), el que quedó en disco es de otra corrida y **describe otra
  cosa**. El de ésta se escribe en
  `_validacion (NO SE PUDO ESCRIBIR _validacion.txt - LEER ESTE).txt`, se avisa
  en consola y la corrida termina con error;
- cuando no hay ningún aviso pero sí archivos ajenos o gente que faltó, el
  informe **nunca dice «Sin avisos.» a secas**: sería el mensaje más
  tranquilizador posible al lado de un envío incompleto;
- los archivos ajenos **no** cambian el código de salida: que la carpeta tenga
  archivos de más no significa que la corrida haya fallado, y un código 1
  recurrente enseñaría a ignorar el código de salida, que es lo que distingue
  una corrida incompleta de una completa.

### 6.6 Qué frena una entrega y qué sólo avisa

El criterio es uno solo: **frena lo que haría salir mal el entregable; avisa
todo lo que el dueño puede mirar y decidir.**

**Frena la corrida entera** (no se genera ningún anexo):

- una fuente manual que falta, o cuyas horas no cierran contra la persona a la
  que se le restan (un día negativo, o un día que la otra persona no tiene).
  Esas planillas **mueven horas de una persona a otra**: un reparto que no
  cuadra se factura mal y no se ve en una hoja de ochenta filas;
- `config/mapeo.yaml` roto, o con un campo declarado en blanco;
- que falte una plantilla, o que no exista `input/<mes>/`.

**Frena ese anexo** (el otro se genera igual, y la entrega queda incompleta):

- la verificación de integridad no cierra (§6.4);
- el mes no entra en la plantilla;
- el archivo está abierto y no se puede escribir;
- la plantilla no tiene la forma que el escritor espera.

**Deja a ese desarrollador afuera** (los demás entran, y los anexos salen
marcados como INCOMPLETO en el nombre):

- export ilegible, de formato desconocido o con las columnas corridas;
- export sin ninguna fila de datos, o con todas sus filas fuera del mes: las
  dos cosas suelen ser el rango de fechas mal puesto en Kimai, y el error lo
  dice;
- hora de inicio que no tiene formato de hora;
- proyecto sin código entre corchetes.

**Sólo avisa** en `_validacion.txt`, y la entrega sale igual:

- una persona del resumen mensual que no está en `personas:`, o cuyo `nombre:`
  está declarado en dos entradas → usuario de Kimai vacío;
- un proyecto que no está en `proyectos:` → sale con el nombre de Kimai;
- un proyecto sin valor hora → sale con sus horas y sin importe;
- registros sin descripción, por persona y con la cantidad;
- filas que van sin descripción y sin horario por venir del resumen mensual;
- días hábiles sin carga, horas en fin de semana, más de 12 h en un día,
  registros fuera del mes.

**Lo que dejó de frenar, y por qué.** Un proyecto sin mapear frenaba: ahora el
nombre se deriva del texto de Kimai, y frenar una entrega por un nombre cuesta
un mes de atraso. Una persona sin declarar frenaba: lo único que el mapeo le
aportaba era el usuario de Kimai, y su ausencia produce **una celda vacía** en
la hoja `Datos` que ya es un valor legítimo del entregable —hay gente que
realmente no tiene usuario—. Un export con horas de dos personas frenaba: el
chequeo existía para que el Excel pivoteado no le imputara todo a una sola
persona, y hoy cada fila lleva su propio nombre, así que frenar dejaría afuera
a **dos** desarrolladores en vez de incluirlos bien.

### 6.7 Si falta un desarrollador, se ve en el nombre del archivo

Éste es el riesgo del entregable: con todas las personas adentro de dos
documentos, **una que falta es invisible**. Nada en el Word ni en el Excel se
ve distinto.

Por eso, cuando algún export falla, los anexos **se generan igual** con lo que
sí se pudo leer —no generarlos dejaría al dueño sin nada que revisar— pero
salen con la marca en el **nombre**:
`Anexo-II-A-Detalle-horas-KLG-2026-09 (INCOMPLETO - FALTAN 2 DESARROLLADORES - NO ENVIAR).xlsx`.
El nombre es lo único que se ve al adjuntar un archivo a un mail.

La marca la decide **sólo lo que no entró de `input/`**, nunca una falla de
escritura posterior: si no, el nombre del segundo anexo dependería del orden en
que se generan.

### 6.8 La tabla 3 sale como borrador mecánico, a propósito

La tabla de principales trabajos se agrupa por proyecto y por el ticket que
cita cada descripción (iTop `R-012532`, ClickUp `CU-8xyz`, y los `BUG-` que el
equipo usó en agosto), con las horas sumadas y las descripciones concatenadas.

**La aritmética está garantizada:** cada fila del mes cae en exactamente un
grupo, así que los subtotales cierran exactos y el total da el total del mes.
El dueño puede reescribir los textos y fusionar líneas sumando sus horas sin
que deje de cuadrar.

**La redacción no se automatiza.** La tabla que C.UNIX armó para agosto de 2026
estaba agrupada **por tema**, leyendo las descripciones una por una: partió las
45,5 h «sin ticket» en tres líneas temáticas, fusionó tres tickets en una sola
línea, y le sumó a un ticket trabajo relacionado que no lo citaba. Saber qué
trabajos cuentan la misma historia es criterio editorial, no cálculo, y
automatizarlo sería inventar.

Por el mismo motivo la columna «Estado al cierre» y las observaciones de la
tabla 4 quedan con la marca `[●]` que trae la plantilla: **un hueco visible es
preferible a un dato inventado**, y todo lo que queda con `[●]` o
`[DD/MM/AAAA]` es exactamente lo que falta completar.

### 6.9 Un export que falla no frena a los demás

Cada archivo de `input/` se procesa aislado. El que falla se registra con su
motivo —en consola y en el informe— y los otros siguen. Frenar todo por un
archivo roto obligaría a arreglarlo antes de poder mirar nada, y el dueño
cierra el mes contra reloj.

### 6.10 Las horas van exactas, sin redondeo de presentación

Los anexos escriben las horas tal como vienen: no hay celdas de día que
redondear. Por eso el aviso de «desvío por redondeo» que existía en el
validador se eliminó junto con el Excel pivoteado por desarrollador, que era el
único formato que redondeaba cada celda por separado. Lo único que se redondea
es la presentación de los totales en el informe (un decimal), y con medio hacia
arriba.

### 6.11 `scripts/anexos_agosto_2026.py` es historia, no camino normal

Agosto de 2026 se entregó **corrigiendo el ejemplo que mandó C.UNIX**, no
armándolo desde cero: la herramienta todavía no generaba los anexos. Ese script
es el registro auditable de esa corrección y lo único que reproduce lo
entregado tal cual. Está versionado para poder auditarlo, no para usarlo: de
septiembre en adelante el camino es `python -m cunix_horas AAAA-MM`.

Las cuatro planillas de `input/2026-08/manual/` son la **única copia** de esos
datos —las horas de Gabriel Denis, el detalle de Alexis Carnero y los dos
anexos tal como los mandó C.UNIX—: sin ellas, agosto de 2026 no se puede volver
a generar.

---

## 7. Qué se eliminó, y por qué

El proyecto pasó por **tres generaciones de entregables**, porque el cliente
cambió dos veces lo que quería recibir. Hoy sólo vive la tercera. Si el
historial de git muestra módulos que ya no están, es por esto:

1. **Un Excel pivoteado por desarrollador** (cliente → proyecto → actividad, un
   día por columna, un archivo por persona). C.UNIX pasó a facturar sobre el
   detalle y no sobre el pivot: quería ver cada carga de horas con su hora de
   inicio y su descripción. Se eliminaron `agregador.py`, `escritor_excel.py`,
   la plantilla de estilos, el fixture de salida del formato viejo y sus tests.
2. **Un archivo plano consolidado** con todos los devs
   (`Horas KLG-<Mes><Año>.xlsx`), con columnas `E-mail` y `Project number`.
   C.UNIX definió sus propios anexos y dejó de recibirlo. Se eliminaron
   `detalle.py`, `escritor_detalle.py` y sus tests.
3. **Los dos anexos de C.UNIX.** Es lo único que se genera hoy.

De las dos primeras generaciones **sobrevive la lectura**: los exports de Kimai
no cambiaron, así que `lector_*.py`, `kimai_comun.py`, `completado.py` y
`validador.py` siguen igual.

De ahí salen también los campos del mapeo que ya no existen: `mail:`,
`numero_proyecto:`, `archivo:`, `actividad:` y `archivo_salida:` servían al
archivo plano o al pivot, y los dos primeros **frenaban la entrega entera** si
faltaban. Era un programa que podía trabar un mes exigiendo un dato que después
no escribía en ningún lado. Se eliminaron con su validación; si quedaron en un
`mapeo.yaml` viejo, se ignoran al cargar.

`Registro` sigue teniendo `email` y `numero_proyecto` porque son datos que el
export de Kimai trae: se leen, no se exigen, y hoy no se escriben en ningún
lado.

---

## 8. Testing

TDD, con `pytest`, sobre fixtures reales. `python -m pytest -q`.

- **Los lectores**, contra exports reales de los tres formatos: cantidad de
  registros, suma de horas, rango de fechas, tolerancia al `showZeroes`,
  `sharedStrings` e `inlineStr`, serial de fecha, duración, separador decimal y
  orden de la fecha.
- **`mapeo`**: resuelve códigos y personas; devuelve `None` —sin lanzar— ante
  lo que no está y ante un `nombre:` repetido; cada sección es opcional; los
  patrones de nombre se validan al cargar.
- **`completado`**: un registro del resumen mensual sale con el usuario del
  mapeo, o vacío con aviso; uno del detalle sale con los valores de Kimai
  aunque el mapeo diga otra cosa.
- **`filas_anexo`**: `Fin = Inicio + Horas`, incluido el cruce de medianoche;
  orden de la hoja; las tres tablas del informe; que los subtotales de la
  tabla 3 cierren contra el total del mes.
- **`fuentes_manuales`**: la resta día por día, sus casos de error, la lectura
  tolerante de encabezados y fechas, y la carga de descripciones con su
  verificación.
- **Los escritores**: el archivo generado, releído, tiene lo que corresponde y
  no tocó lo que es de C.UNIX; la verificación de integridad detecta una fila
  perdida; la plantilla no quedó modificada.
- **`cli` (integración)**: de `input/` a `output/`; un export roto no frena a
  los otros y los anexos salen con el nombre que avisa; el informe dice quién
  falta y por qué; el limpio de una corrida anterior se aparta y no se borra;
  los archivos ajenos se listan. **Punta a punta sobre los cinco exports reales
  de `input/2026-08/`**, verificando el total contra la suma de los exports.

---

## 9. Lo que queda pendiente

1. **La hoja `Instrucciones` de la plantilla del Anexo II-A describe agosto de
   2026 como ejemplo.** Es texto de C.UNIX y no afecta ningún número, pero se
   arrastra a todos los meses. Conviene resolverlo con ellos antes de que
   confunda a alguien. (`Resumen!B5`, la celda «Horas en Kimai según C.UNIX»,
   sí quedó vacía en la plantilla: la carga C.UNIX al conciliar.)

## 10. Fuera de alcance (posible trabajo futuro)

- Integración directa con la API de Kimai, que eliminaría el paso manual de
  exportar uno por uno.
- Comparación mes a mes y detección de desvíos.
