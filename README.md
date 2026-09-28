# cunix-horas

Convierte los exports de Kimai en **el archivo mensual que recibe el partner**.

## Uso mensual

1. En Kimai, exportar las horas del mes, **un archivo por desarrollador**,
   con el **reporte de detalle** (ver abajo).
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

El escritor de ese formato, su plantilla y sus tests **siguen en el repo** y
siguen probados, pero el programa ya no los usa. Si el partner vuelve atrás,
se recupera sin reescribir nada.

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
lee los dos del **reporte de detalle** y no hay que decirle cuál es: lo
detecta sola.

- **Timesheet en Excel** (`.xlsx` con `Date` en A1): una fila por registro de
  tiempo. Es el que se viene usando.
- **Timesheet en CSV** (`.csv`): lo mismo, pero en texto. Kimai lo escribe con
  la fecha al derecho (`2026-08-31`) y la duración en horas y minutos
  (`2:00`).

Si un archivo no es ninguno de esos dos, **ese** desarrollador no entra al
archivo del mes, el motivo dice qué encontró en la fila 1 y qué formatos se
reconocen, y el archivo sale marcado como INCOMPLETO. Los demás entran igual.

### El resumen mensual ya no sirve

El export de **resumen mensual** (`.xlsx` con `Total` en B1, la grilla de
días) se sigue reconociendo, pero **ya no se acepta**: no trae la hora de
inicio, ni el usuario, ni el mail, ni la descripción, ni el número de
proyecto. Son las horas ya sumadas por día. Aceptarlo dejaría esas columnas en
blanco y el partner recibiría filas incompletas sin que nadie lo note.

Para arreglarlo: volver a exportar a esa persona desde Kimai con el reporte de
detalle, el mismo que usaste para el resto, y dejar ese archivo en lugar del
otro.

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

## El mapeo ya no es obligatorio

`config/mapeo.yaml` sirve para **pulir los nombres** que ve el partner, y nada
más:

- Un proyecto que no esté declarado **no frena nada**: sale con el nombre que
  trae Kimai, y queda listado en `_validacion.txt` con el bloque listo para
  pegar por si querés cambiarlo. Cada fila lleva su `Project number`, así que
  la trazabilidad no depende del mapeo.
- La sección `personas:` **es opcional**: el nombre, el usuario y el mail de
  cada desarrollador vienen de Kimai.

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
config/mapeo.yaml        configuración editada a mano
templates/plantilla.xlsx estilos del formato anterior (ya no se usa)
input/AAAA-MM/           exports de Kimai, .xlsx o .csv (se versionan)
output/AAAA-MM/          el archivo del partner (NO se versiona)
cunix_horas/             el código
tests/                   los tests
docs/superpowers/        spec y plan de implementación
```
