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
  del mes).

Cada corrida vacía primero los `.xlsx` de `output/<mes>/` y los vuelve a
generar, así que ahí nunca queda un Excel de una corrida anterior: lo que está
en la carpeta es siempre lo que dice el informe.

## Cuando aparece un proyecto o un dev nuevo

El proceso frena para ese archivo y muestra la línea exacta a pegar en
`config/mapeo.yaml`. Se pega, se ajusta el nombre que va a ver el partner,
y se vuelve a correr.

## Un export por desarrollador

En Kimai hay que exportar **filtrando por un solo desarrollador**. Si un export
trae horas de dos personas, el archivo no se genera y el motivo queda en
`_validacion.txt`: sin esa verificación, las horas de todos se le facturarían
a una sola. Lo mismo si el export no trae ninguna fila de datos, que suele ser
el rango de fechas mal puesto.

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
