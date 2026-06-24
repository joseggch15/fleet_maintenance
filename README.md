# fleet_maintenance

Cargador del **Fleet Tag Inventory and Maintenance** de la mina Newmont
Merian. Toma el Excel de _submissions_ que exporta el formulario
**"Merian Ops - Form - Fleet maintenance"** y agrega cada inspección como una
fila nueva en la hoja `Full List 2024-2025` del Excel maestro
`Fleet Tag Inventory and Maintenance.xlsx`, **conservando su estructura**
(estilos, fórmulas, la tabla `Table3` y los pivotes/dashboards).

Hermano de [`diesel_report`](../diesel_report/) y
[`lubes_report`](../lubes_report/): mismo stack (PySide6 + openpyxl), misma
forma de trabajar — un click para volcar datos de un export al Excel maestro
sin romper su formato.

## Instalación

```bash
pip install -r requirements.txt
```

Requiere Python 3.10+.

## Uso

```bash
python run.py
```

Pasos en la interfaz:

1. **Cargar Excel de submissions...** — el archivo
   `merian_ops_-_form_-_fleet_maintenance_submissions.xlsx` descargado del
   software del formulario.
2. **Seleccionar Excel destino...** — el maestro
   `Fleet Tag Inventory and Maintenance.xlsx`.
3. **Revisar / corregir** las filas en la tabla de previsualización. Cada fila
   trae una casilla *Incluir* para elegir cuáles se cargan; cualquier celda
   se puede editar antes de escribir.
4. **CARGAR AL EXCEL DESTINO** — agrega las filas a `Full List 2024-2025`.
   Antes de escribir se crea un **respaldo automático** del destino
   (`...backup_AAAAMMDD_HHMMSS.xlsx`).
5. Abrir el Excel y usar **Datos > Actualizar todo** para refrescar los
   pivotes y las hojas de resumen.

## Estructura del proyecto

```
run.py             Punto de entrada
app_qt.py          Interfaz gráfica PySide6
source_reader.py   Lectura del export del formulario (hoja 'Form Submissions')
mapping.py         Mapeo formulario -> columnas de 'Full List 2024-2025'
excel_writer.py    Escritura preservando estilos, fórmulas y la tabla Table3
tests/             Pruebas (pytest) con fixtures sintéticas
```

## Mapeo de columnas

Cada submission se convierte en una fila de `Full List 2024-2025` (tabla
`Table3`). Las columnas **A (`Date`)** y **D (`Verified`)** del destino son
fórmulas que se recalculan solas; la herramienta las reproduce, no las llena.

| Destino (Full List)                     | Origen (Form Submissions)                         |
|-----------------------------------------|---------------------------------------------------|
| A · Date                                | _fórmula_ `=TEXT([Date (mm/dd/yy)], "mmm-yy")`    |
| B · Date (mm/dd/yy)                      | `Date And Time Of Revision` (solo la fecha)       |
| C · Vehicle ID                          | `Vehicle Id`                                      |
| D · Verified                            | _fórmula de matriz_ (cruza con TAG/SMU History)   |
| E · FMS ID                              | `Fms Id`                                          |
| F · System fitted as per standard (Y/N) | _por defecto_ `Y`                                 |
| G · Equipment Hours/ODO                 | `Machine Hourmeter Reading`, si no `… Kilometer`  |
| H · FMS Hours                           | _(sin equivalente — vacío)_                       |
| I · Status                              | _por defecto_ `VIU OK`                            |
| J · # INLETS                            | `Number Of Inlets?`                               |
| K · Are aditional inlets locked?(Y/N)   | `Additional Inlets Locked?` (Yes/No/Not appl. → Y/N/N/A) |
| L · Drain valves locked?(Y/N)           | `Drain Valves Locked` (igual mapeo)               |
| M · Fast fill receiver leaking?(Y/N)    | _por defecto_ `N`                                 |
| N · # SMU/TAGS                          | `Tag Type` → `SMU` / `TAG`                         |
| O · EQUIPMENT TYPE                      | `Equipment Type?`                                 |
| P · REMARKS                             | `Remarks`                                         |
| Q · Inspectors                          | `Name Of Inspector`                               |
| R · OWNER                               | `Company Name?`                                   |
| S · REMEDIAL ACTIONS / UPDPATES         | _(sin equivalente — vacío)_                       |

Los valores por defecto y las conversiones se centralizan en
[`mapping.py`](mapping.py) para ajustarlos fácilmente. Las columnas sin
equivalente directo quedan editables en la previsualización.

### Decisiones de carga

- **Fecha de la fila**: se usa `Date And Time Of Revision` (la fecha de la
  inspección en campo); si falta, se cae a `Submitted At`.
- **Horas/ODO**: se usa el `Machine Hourmeter Reading`; si está vacío, el
  `Machine Kilometer Reading`.
- **Duplicados**: la herramienta **agrega todas** las filas marcadas. Si se
  re-importa el mismo export se duplican filas, así que importe solo
  submissions nuevas (o desmarque las ya cargadas en la previsualización).

## Cómo se preserva la estructura

`excel_writer.py` abre el maestro con openpyxl y, por cada fila nueva:

- copia estilos (fuente, relleno, bordes, alineación y **formato de número**)
  desde la última fila de datos como plantilla;
- reproduce las fórmulas de A y D, incluyendo la **fórmula de matriz**
  (`ArrayFormula`) de `Verified` con su `ref` propio por fila;
- extiende el rango de la tabla `Table3` para incluir las filas nuevas (sin
  esto, Excel reporta *"We found a problem with some content"*).

Validado contra los archivos reales: las **10 hojas**, las **5 tablas
dinámicas (pivotes)**, sus **4 cachés** y los **5 gráficos** se conservan tras
la escritura.

### Limitación conocida

openpyxl conserva los pivotes y los gráficos, pero **descarta el estilo/color
personalizado de los gráficos** (`colors*.xml` / `style*.xml`): tras importar,
los gráficos de los dashboards pueden volver a la paleta por defecto de Excel.
Los datos y los pivotes quedan intactos, y el **respaldo automático** conserva
el original por si se necesita recuperar el formato de los gráficos.

## Pruebas

```bash
pip install pytest
pytest -q
```

Las pruebas usan workbooks sintéticos creados en tiempo de ejecución (no
dependen de los archivos reales del cliente).
