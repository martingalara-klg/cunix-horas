# cunix-horas

Convierte los exports de Kimai en los Excel mensuales que recibe el partner.

## Uso mensual

1. En Kimai, exportar las horas del mes, **un archivo por desarrollador**,
   con el **reporte de detalle** (ver abajo).
2. Crear la carpeta del mes y poner los archivos adentro:
   `input/2025-10/`  (el nombre de cada archivo da igual)
3. Doble clic en `generar.bat`, o desde una terminal:
   ```
   python -m cunix_horas 2025-10
   ```
4. Los Excel del mes quedan en `output/2025-10/`.
5. Leer `output/2025-10/_validacion.txt` **antes de mandar nada**: ahí está
   la lista de los que van en el envío, y la de los que están en esa
   carpeta pero **no** hay que enviar.

## Qué exports de Kimai lee

Según con qué reporte exportes, Kimai da un archivo distinto. La herramienta
lee los dos del **reporte de detalle** y no hay que decirle cuál es: lo
detecta sola.

- **Timesheet en Excel** (`.xlsx` con `Date` en A1): una fila por registro de
  tiempo. Es el que se viene usando.
- **Timesheet en CSV** (`.csv`): lo mismo, pero en texto. Kimai lo escribe con
  la fecha al derecho (`2026-08-31`) y la duración en horas y minutos
  (`2:00`).

Si un archivo no es ninguno de esos dos, **ese** archivo no se genera y el
motivo dice qué encontró en la fila 1 y qué formatos se reconocen. Los demás
desarrolladores se generan igual.

### El resumen mensual ya no sirve

El export de **resumen mensual** (`.xlsx` con `Total` en B1, la grilla de
días) se sigue reconociendo, pero **ya no se acepta**.

El partner cambió lo que recibe: ahora pide el detalle de cada carga de
horas, una fila por registro, con la **hora de inicio**, el **nombre** y el
**mail** del desarrollador, la **descripción** de lo que hizo y el **número
de proyecto**. El resumen mensual no trae nada de eso: son las horas ya
sumadas por día. Generarlo igual dejaría esas columnas en blanco y el partner
recibiría un archivo incompleto sin que nadie lo note.

Si aparece uno de esos archivos en `input/`, **ese** archivo no se genera, el
motivo queda en `_validacion.txt` y los demás desarrolladores salen igual.
Para arreglarlo: volver a exportar a esa persona desde Kimai con el reporte
de detalle, el mismo que usaste para el resto, y dejar ese archivo en lugar
del otro.

## Qué dice `_validacion.txt`

Es el informe de lo que quedó en `output/<mes>/` y alcanza por sí solo para
decidir qué se envía:

- **Qué Excel se generaron**, por nombre.
- **Qué archivos NO se generaron y por qué.** Si aparece esta lista, la corrida
  está incompleta: esos Excel no están en la carpeta. Hay que corregir el
  motivo y volver a correr antes de enviar nada.
- **Qué `.xlsx` hay en la carpeta que esta corrida NO generó.** Excel viejos de
  un desarrollador que este mes ya no tiene export en `input/`, archivos que la
  herramienta apartó en una corrida anterior, o archivos que dejaste vos ahí.
  La herramienta no los borra, pero los nombra uno por uno: **esos no van en el
  envío del mes.**
- **El desvío por redondeo de cada Excel**, siempre, aunque sea chico.
- **Los avisos de validación** de los Excel que sí se generaron (días hábiles
  sin carga, horas en fin de semana, más de 12 h en un día, registros fuera
  del mes, desvío por redondeo por encima del umbral).

## Qué toca la herramienta en `output/<mes>/`

Sólo los Excel que ella misma genera en esa corrida. Nada más de esa carpeta
se borra: si dejás ahí un archivo tuyo, sigue estando después de correr.

Cada Excel se escribe primero en un temporal y recién cuando salió entero se
mueve sobre el nombre final. Así nunca queda un Excel a medio escribir, y si
un archivo falla, el que ya estaba no se toca.

Cuando un archivo falla y en la carpeta había un Excel del mes pasado con ese
mismo nombre, ese archivo viejo **se renombra** a
`... (CORRIDA ANTERIOR - NO ENVIAR).xlsx`. No se borra —el dato sigue ahí— pero
el nombre ya no se puede confundir con el del mes, y `_validacion.txt` lo dice.

Un Excel abierto es fallo de **ese** archivo, no de la corrida: los demás
desarrolladores se generan igual.

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

## Una sola fila de actividad por proyecto

El Excel tiene tres niveles de fila: cliente, proyecto y actividad. En Kimai
cada desarrollador clasifica sus horas como quiere (`Desarrollo`, `Gestión`,
`Testing`…), pero **el partner no ve esa clasificación interna**: factura sobre
el proyecto, y viene recibiendo una sola fila desde siempre.

Por eso todas las actividades de un mismo proyecto se suman en **una única
fila**, incluso día por día:

```
Sistemas - C.UNIX   8
  VictoriusCP2      8
    Desarrollo      8      <- Desarrollo 5 + Gestión 3
Total               8
```

El texto de esa fila sale de `actividad:` en `config/mapeo.yaml`. Si se borra
esa línea vale `Desarrollo`, que es lo que el partner recibió siempre. Para
cambiarlo alcanza con editar esa línea:

```yaml
actividad: "Desarrollo"
```

Proyectos distintos siguen separados: lo que se junta es la clasificación de
Kimai, no los proyectos ni los clientes.

## El desvío por redondeo

Cada celda de día se redondea a 2 decimales para que las filas del Excel
cierren a la vista del cliente. El precio es que el total puede apartarse de
las horas reales del export, y ese error crece con la cantidad de celdas
(hasta 0.005 h por celda; una carga de 50 minutos, por ejemplo, pierde
0.0033 h). `_validacion.txt` muestra el desvío exacto de cada Excel **siempre**,
como dato, y si pasa de 0.25 h además lo avisa: no es un error de carga, pero
decidís vos si importa para facturar.

## Qué dato de Kimai va a parar a dónde

De cada registro de tiempo se conservan, además de la fecha y las horas:

| Dato de Kimai | Para qué |
|---|---|
| `From` (hora de inicio) | Va junto con la fecha en el entregable |
| `Name` (`Matias Zalazar`) | Nombre para mostrar |
| `User` (`mzalazar`) | Clave de `personas:` en `config/mapeo.yaml` |
| `E-mail` | Va al entregable |
| `Customer` y `Project` (texto crudo) | De ahí sale el nombre si el proyecto no está mapeado |
| `Description` | Lo que hizo; **puede venir vacía y está bien** |
| `Project number` | Va tal cual al entregable |

Ojo con las dos últimas columnas de proyecto, que se parecen y no son lo
mismo: el código entre corchetes de `Project` (`[AD2690002]`) es el que se
busca en `config/mapeo.yaml`, y el `Project number` (`210`) es el que ve el
partner. Se guardan los dos.

## Cuando aparece un proyecto o un dev nuevo

El proceso frena para ese archivo y muestra la línea exacta a pegar en
`config/mapeo.yaml`. Se pega, se ajusta el nombre que va a ver el partner,
y se vuelve a correr.

## Un export por desarrollador

En Kimai hay que exportar **filtrando por un solo desarrollador**. Si un export
trae horas de dos personas, el archivo no se genera y el motivo queda en
`_validacion.txt`: sin esa verificación, las horas de todos se le facturarían
a una sola. Lo mismo si el export no trae ninguna fila de datos, o si trae
filas pero **todas** caen fuera del mes que estás generando: las dos cosas
suelen ser el rango de fechas mal puesto en Kimai, y las dos darían un Excel
en blanco. El motivo, con el rango de fechas que sí trae el archivo, queda en
`_validacion.txt`.

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
input/AAAA-MM/           exports de Kimai, .xlsx o .csv (se versionan)
output/AAAA-MM/          Excel generados (NO se versionan)
cunix_horas/             el código
tests/                   los tests
docs/superpowers/        spec y plan de implementación
```
