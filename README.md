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
