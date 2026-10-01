# cunix-horas

Convierte los exports de Kimai en **lo que KLG le entrega a C.UNIX todos los
meses**.

## El entregable de hoy: los dos anexos

C.UNIX recibe **dos documentos por mes**, con el formato que mandó él mismo:

- **Anexo II** — el informe mensual en Word: horas e importe por proyecto,
  horas por persona, principales trabajos y observaciones.
- **Anexo II-A** — el detalle en Excel: una fila por registro de tiempo, con
  las columnas de control y revisión que C.UNIX completa después.

Las plantillas vacías de los dos viven en `templates/`:

```
templates/Anexo-II-Informe-mensual-horas-KLG.docx
templates/Anexo-II-A-Detalle-horas-KLG.xlsx
```

**Todavía no las usa el CLI.** Agosto de 2026, el primer mes entregado con
este formato, salió de `scripts/anexos_agosto_2026.py`, un script atado a ese
mes. Promover ese script a parte de la herramienta es el trabajo que queda
pendiente; abajo, «Lo que falta».

Lo que el CLI genera hoy sigue siendo el consolidado plano del formato
anterior (`Horas KLG-<Mes><Año>.xlsx`), que es lo que describe el resto de
este README.

## Uso mensual

1. En Kimai, exportar las horas del mes, **un archivo por desarrollador**.
   Conviene el **reporte de detalle**, que trae las diez columnas; el de
   resumen mensual también sirve, con lo que se pierde (ver abajo).
2. Crear la carpeta del mes y poner los archivos adentro:
   `input/2025-10/`  (el nombre de cada archivo da igual)
3. Doble clic en `generar.bat`, o desde una terminal:
   ```
   python -m cunix_horas 2025-10
   ```
4. El archivo del mes queda en `output/2025-10/`, con el nombre
   `Horas KLG-Oct2025.xlsx`.
5. Leer `output/2025-10/_validacion.txt` **antes de mandar nada**: ahí está,
   arriba de todo, la lista de los desarrolladores que entraron al archivo y
   la de los que no, con el motivo de cada uno.

## Qué recibe el partner

**Un solo archivo por mes, con todos los desarrolladores juntos, una fila por
cada carga de horas.** Diez columnas, en este orden:

```
Date | Duration | Name | User | E-mail | Customer | Project | Activity | Description | Project number
```

- `Date` es la fecha del registro; adentro guarda también la hora de inicio,
  aunque la celda muestre sólo la fecha.
- `Duration` es una **duración** (`4:00`), no un número de horas (`4`). Así,
  cuando el partner suma la columna, le da `160:30` y no `160,5`.
- `Description` puede venir vacía y está bien: en el archivo de referencia que
  mandó el partner, 112 de las 160 filas lo están.
- `Project number` es un campo propio de Kimai (`210`) y **no** el código entre
  corchetes del proyecto (`[AD2690002]`). Son dos cosas distintas que se
  parecen: el del corchete es la clave de `config/mapeo.yaml`, el otro es el
  que ve el partner.

Las filas van agrupadas por desarrollador, los desarrolladores en orden
alfabético, y las de cada uno en orden cronológico.

### Antes era un Excel por desarrollador

Hasta septiembre de 2025 el partner recibía un Excel pivoteado por persona,
con las horas sumadas por día. Cambió porque factura sobre el detalle: quiere
ver cada carga con su hora de inicio, su descripción y su número de proyecto,
y un solo adjunto por mes.

Ese formato **se eliminó del repo** (`escritor_excel.py`, sus tests y
`templates/plantilla.xlsx`). Se había conservado «por si el partner vuelve
atrás», pero el partner no volvió atrás: volvió a cambiar, y ahora recibe los
dos anexos. Un escritor que nadie ejecuta no está probado contra nada real,
y mantenerlo costaba más que recuperarlo del historial de git si alguna vez
hiciera falta.

## Si falta un desarrollador

Este es el riesgo del archivo único y conviene tenerlo claro.

Antes, si el archivo de alguien fallaba, **faltaba un Excel entero** en la
carpeta: imposible no notarlo. Ahora todo va junto, así que un desarrollador
que falta **es invisible**: el archivo se ve completo y no lo es.

Por eso, cuando algún export falla:

- el archivo **se genera igual** con los que sí se pudieron leer, para que
  puedas revisar lo que hay;
- pero sale con el nombre
  `Horas KLG-Oct2025 (INCOMPLETO - FALTAN 2 DESARROLLADORES - NO ENVIAR).xlsx`,
  que es lo único que se ve al adjuntarlo a un mail;
- y si en la carpeta había un `Horas KLG-Oct2025.xlsx` limpio de una corrida
  anterior, se lo renombra a `... (CORRIDA ANTERIOR - NO ENVIAR).xlsx` para que
  no se envíe en su lugar. No se borra: el dato sigue ahí.

Cuando no falla nada, el nombre es el limpio.

## Qué exports de Kimai lee

Según con qué reporte exportes, Kimai da un archivo distinto. La herramienta
lee los tres y no hay que decirle cuál es: lo detecta sola.

- **Timesheet en Excel** (`.xlsx` con `Date` en A1): una fila por registro de
  tiempo. Es el **reporte de detalle**, el que conviene usar.
- **Timesheet en CSV** (`.csv`): lo mismo, pero en texto. Kimai lo escribe con
  la fecha al derecho (`2026-08-31`) y la duración en horas y minutos
  (`2:00`).
- **Resumen mensual** (`.xlsx` con `Total` en B1): la grilla de días, con las
  horas ya sumadas. Sirve, pero trae menos (abajo).

Si un archivo no es ninguno de esos tres, **ese** desarrollador no entra al
archivo del mes, el motivo dice qué encontró en la fila 1 y qué formatos se
reconocen, y el archivo sale marcado como INCOMPLETO. Los demás entran igual.

### Qué se pierde con el resumen mensual

Ese export no trae cinco de las diez columnas: la hora de inicio, el usuario,
el mail, la descripción y el número de proyecto.

**Dos se pueden dejar vacías, y el archivo sale igual:**

- la **descripción**, porque el partner la acepta vacía: 112 de las 160 filas
  de su archivo de referencia lo están;
- la **hora de inicio**, que queda en `00:00`: una de sus 160 filas está así.

`_validacion.txt` te dice, por desarrollador, de qué reporte salieron sus
filas, y avisa cuáles van sin descripción y sin hora de inicio. No frena nada:
es para que decidas si lo mandás así o volvés a exportar a esa persona con el
reporte de detalle.

**Las otras tres no pueden ir vacías**, porque las 160 filas del partner las
tienen llenas, sin una sola excepción: el **usuario**, el **mail** y el
**número de proyecto**. Esas tres se completan desde `config/mapeo.yaml`, y
**nunca se adivinan**:

| Columna | De dónde sale |
|---|---|
| `Name` | del propio archivo |
| `User` | la **clave** de la persona en `personas:` |
| `E-mail` | `mail:` de esa persona |
| `Project number` | `numero_proyecto:` de ese proyecto |

Si al mapeo le falta alguno de los tres, **ese archivo no entra**. El motivo
nombra al desarrollador o al proyecto y trae el bloque listo para pegar en
`config/mapeo.yaml`. Los demás desarrolladores entran igual.

El `Project number` es el que menos se puede improvisar: **no** es el código
entre corchetes. El proyecto `[AD2690002]` tiene `Project number` `210`. No se
deduce de nada, así que o está declarado o el archivo frena. Y dejarlo vacío
no es opción: el partner factura sobre esa columna.

Si exportaste con el **reporte de detalle**, nada de esto aplica: el usuario,
el mail y el número vienen de Kimai, y el mapeo no interviene aunque declare
otra cosa.

## Qué dice `_validacion.txt`

Es el informe de lo que quedó en `output/<mes>/` y alcanza por sí solo para
decidir qué se envía:

- **Quiénes entraron al archivo**, con cuántos registros y cuántas horas cada
  uno. Contá esa lista contra tu equipo antes de mandar nada.
- **Qué archivos NO entraron y por qué.** Si aparece esta lista, el archivo
  del mes está incompleto y lleva la marca en el nombre. Hay que corregir el
  motivo y volver a correr antes de enviar.
- **El nombre del archivo generado**, con su cantidad de filas y su total de
  horas.
- **Qué `.xlsx` hay en la carpeta que esta corrida NO generó.** Los Excel del
  formato anterior, un consolidado marcado como INCOMPLETO de otra corrida, o
  archivos que dejaste vos ahí. La herramienta no los borra, pero los nombra
  uno por uno: **esos no van en el envío del mes.**
- **De qué reporte salió cada desarrollador**, y el aviso de quiénes van sin
  descripción y sin hora de inicio por haber exportado el resumen mensual.
- **Los proyectos sin mapear**, con el nombre que se usó y el bloque listo para
  pegar en `config/mapeo.yaml` si querés que el partner vea otro.
- **Los avisos de validación**, agrupados por desarrollador: días hábiles sin
  carga, horas en fin de semana, más de 12 h en un día, registros fuera del
  mes. Y, del mes entero, el aviso de un mismo `Project number` que aparece
  con dos nombres de proyecto distintos.

## La verificación de integridad

Después de escribir el archivo, la herramienta lo **vuelve a abrir** y suma la
columna `Duration`. Ese total tiene que coincidir exactamente con las horas de
los exports que leyó. Si no coincide, **no genera nada** y lo dice fuerte: no
es un aviso.

Con 160 filas de cinco desarrolladores en un solo archivo, una fila perdida no
se ve nunca a ojo.

## Qué hace falta en `config/mapeo.yaml`

Depende de con qué reporte exportó cada persona.

**Si exportó con el reporte de detalle, el mapeo es sólo un pulido de
nombres:**

- Un proyecto que no esté declarado **no frena nada**: sale con el nombre que
  trae Kimai, y queda listado en `_validacion.txt` con el bloque listo para
  pegar por si querés cambiarlo. Cada fila lleva su `Project number`, así que
  la trazabilidad no depende del mapeo.
- La sección `personas:` es opcional: el nombre, el usuario y el mail de cada
  desarrollador vienen de Kimai.

**Si exportó con el resumen mensual, el mapeo es obligatorio para esa persona
y para sus proyectos**, porque ese reporte no trae el usuario, el mail ni el
número de proyecto:

```yaml
personas:
  lzalazar:                                   # esta clave ES el User
    nombre: "Lautaro Zalazar"                 # como lo muestra Kimai
    archivo: "L Zalazar"
    mail: "lautaro.zalazar@cunix.net"         # la columna E-mail

proyectos:
  AD2690002:                                  # el código entre corchetes
    cliente: "Sistemas - C.UNIX"
    proyecto: "VictoriusCP2"
    numero_proyecto: "210"                    # el Project number de Kimai
```

- La **clave** de la persona es el `User` que ve el partner, y el `nombre:`
  tiene que coincidir exactamente con lo que muestra Kimai: es lo único que
  el resumen mensual trae para identificarla.
- `mail:` y `numero_proyecto:` son **opcionales** en el archivo, y sólo los
  usan las filas del resumen mensual. Si los escribís vacíos, el mapeo no
  carga: mejor fallar con el nombre de la entrada que dejar una columna en
  blanco del otro lado.
- `numero_proyecto:` **no** es el código entre corchetes. Miralo en Kimai, en
  la ficha del proyecto.

### El nombre del archivo de salida

Se configura en `config/mapeo.yaml`:

```yaml
archivo_salida: "Horas KLG-{mes}{anio}.xlsx"
```

`{mes}` es el mes en tres letras (`Jan`…`Dec`) y `{anio}` el año. El archivo
de referencia del partner usa `Sept` para septiembre: si querés esa forma
exacta, o los meses en español, se escribe el mes a mano en el patrón
(`"Horas KLG-Sept{anio}.xlsx"`), sabiendo que así hay que actualizar esa línea
todos los meses.

## Qué toca la herramienta en `output/<mes>/`

Sólo el archivo que ella misma genera en esa corrida, y el limpio de una
corrida anterior cuando tiene que apartarlo (lo renombra, nunca lo borra).
Nada más de esa carpeta se toca: si dejás ahí un archivo tuyo, sigue estando
después de correr.

El archivo se escribe primero en un temporal, se verifica ahí la integridad, y
recién cuando salió entero se mueve sobre el nombre final. Así nunca queda un
archivo a medio escribir, y si algo falla, el que ya estaba no se toca.

`_validacion.txt` se reescribe en cada corrida y enumera **todos** los `.xlsx`
que quedan en la carpeta: los que generó y los que no. Es decir, describe la
carpeta, no lo que hizo la corrida. Lo que el informe no puede saber es lo que
pase con la carpeta *después* de correr: si movés o agregás archivos a mano, el
informe ya no corresponde y hay que volver a correr.

Si `_validacion.txt` no se puede escribir (lo más común: lo tenés abierto), el
de la corrida anterior queda en disco describiendo otra cosa. En ese caso la
herramienta escribe el informe de esta corrida en
`_validacion (NO SE PUDO ESCRIBIR _validacion.txt - LEER ESTE).txt`, lo avisa
en consola y termina con error.

## Qué dato de Kimai va a parar a dónde

De cada registro de tiempo se conservan, además de la fecha y las horas:

| Dato de Kimai | Para qué |
|---|---|
| `From` (hora de inicio) | Va junto con la fecha en la columna `Date` |
| `Name` (`Matias Zalazar`) | Columna `Name` |
| `User` (`mzalazar`) | Columna `User` |
| `E-mail` | Columna `E-mail` |
| `Customer` y `Project` (texto crudo) | De ahí salen `Customer` y `Project` si el proyecto no está mapeado |
| `Description` | Columna `Description`; **puede venir vacía y está bien** |
| `Project number` | Columna `Project number`, tal cual |

El `Customer` de algunos proyectos viene con el código entre corchetes
(`[616050001] Instituto de Salud Pública`) y el de otros no (`CUNIX`): los dos
casos se derivan bien.

## Un export por desarrollador

En Kimai conviene exportar **filtrando por un solo desarrollador**, un archivo
por persona: así, si uno falla, se sabe cuál. Pero si un export trae horas de
dos personas, el archivo se genera igual y cada fila lleva su propio `Name`,
`User` y `E-mail`: en este formato es imposible que las horas de uno se le
imputen a otro.

Lo que sí deja a ese archivo afuera es que el export no traiga **ninguna**
fila de datos, o que **todas** sus filas caigan fuera del mes que estás
generando: las dos cosas suelen ser el rango de fechas mal puesto en Kimai. El
motivo, con el rango de fechas que sí trae el archivo, queda en
`_validacion.txt`.

## Tests

```
python -m pytest -v
```

## Estructura

```
config/mapeo.yaml          configuración editada a mano
templates/                 las plantillas vacías de los dos anexos de C.UNIX
input/AAAA-MM/             exports de Kimai, .xlsx o .csv (se versionan)
input/AAAA-MM/manual/      planillas que NO salen de Kimai (ver abajo)
output/AAAA-MM/            lo que se entrega (se versiona con `git add -f`)
cunix_horas/               el código
scripts/                   los scripts de un mes concreto
tests/                     los tests
docs/superpowers/          spec y plan de implementación
```

### `input/AAAA-MM/manual/`

Ahí van las planillas del mes que **no salen de Kimai** y que, sin embargo,
hacen falta para armar los anexos. De agosto de 2026 hay dos:

| Archivo | Qué es |
|---|---|
| `gabriel-denis.xlsx` | Las 91 h de Gabriel Denis, que trabajó sin usuario de Kimai y cuyas horas quedaron cargadas en la cuenta de Alexis Carnero. |
| `alexis-carnero.xlsx` | El detalle diario que entregó Alexis Carnero después del reclamo de C.UNIX. Aporta las descripciones que Kimai nunca registró. |
| `anexo-II-informe-cunix.docx` | El Anexo II tal como lo mandó C.UNIX, con agosto cargado como ejemplo. Es la base que corrige el script. |
| `anexo-II-A-detalle-cunix.xlsx` | Lo mismo, el Anexo II-A. De acá salieron también las plantillas vacías de `templates/`. |

Son la **única copia** de esos datos: sin ellas, agosto de 2026 no se puede
volver a generar. Por eso están versionadas y no en la raíz.

El CLI **no mira** esta carpeta: recorre `input/AAAA-MM/*.xlsx` y `*.csv` sin
entrar en subcarpetas. Hoy la lee sólo `scripts/anexos_agosto_2026.py`.

## Lo que falta

- **Promover los anexos a la herramienta.** Hoy los genera
  `scripts/anexos_agosto_2026.py`, con los textos y los números de agosto de
  2026 escritos adentro. El CLI todavía no sabe nada de las plantillas de
  `templates/`.
- **Decidir qué pasa con el consolidado plano.** `detalle.py` y
  `escritor_detalle.py` siguen siendo lo que corre `python -m cunix_horas`,
  pero C.UNIX ya no recibe ese archivo.
- **La plantilla del Anexo II-A arrastra dos cosas de agosto de 2026**, en
  hojas que C.UNIX arma y KLG no toca: `Resumen!B5` tiene `306` fijo (las
  horas que C.UNIX concilia contra Kimai, que él carga a mano) y la hoja
  `Instrucciones` describe agosto como ejemplo. Hay que resolverlo antes de
  usar la plantilla para otro mes.
- **El párrafo «EJEMPLO: agosto 2026…»** sigue en la plantilla del Anexo II:
  es la nota con la que C.UNIX la mandó, y queda pendiente decidir si se
  borra al generar.
