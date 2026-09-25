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
5. Leer `output/2025-10/_validacion.txt` **antes de mandar nada**.

## Qué dice `_validacion.txt`

Es el informe de la corrida y alcanza por sí solo para decidir si se envía:

- **Qué Excel se generaron**, por nombre.
- **Qué archivos NO se generaron y por qué.** Si aparece esta lista, la corrida
  está incompleta: esos Excel no están en la carpeta. Hay que corregir el
  motivo y volver a correr antes de enviar nada.
- **Los avisos de validación** de los Excel que sí se generaron (días hábiles
  sin carga, horas en fin de semana, más de 12 h en un día, registros fuera
  del mes, desvío por redondeo).

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
desarrolladores se generan igual. Y `_validacion.txt` se reescribe siempre,
así que el informe nunca describe un estado que ya no es el de la carpeta.

## El desvío por redondeo

Cada celda de día se redondea a 2 decimales para que las filas del Excel
cierren a la vista del cliente. El precio es que el total puede apartarse de
las horas reales del export, y ese error crece con la cantidad de celdas
(hasta 0.005 h por celda). Si la diferencia pasa de 0.5 h, `_validacion.txt`
lo avisa con el número exacto: no es un error de carga, pero decidís vos si
importa para facturar.

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
input/AAAA-MM/           exports de Kimai (se versionan)
output/AAAA-MM/          Excel generados (NO se versionan)
cunix_horas/             el código
tests/                   los tests
docs/superpowers/        spec y plan de implementación
```
