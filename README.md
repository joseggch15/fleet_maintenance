# fleet_maintenance

Software del **Fleet Tag Inventory and Maintenance** de la mina Newmont
Merian. Nació como un cargador de un solo sentido —tomar el Excel de
_submissions_ del formulario y volcarlo en el maestro— y hoy además **guarda
los datos por su cuenta**, los **grafica** y los **exporta** con las mismas
hojas y gráficas que el cliente arma a mano en Excel.

Hermano de [`diesel_report`](../diesel_report/),
[`intakeVersusDelivery`](../intakeVersusDelivery/) y
[`overfillSflSpy`](../overfillSflSpy/): mismo stack (PySide6 + openpyxl +
matplotlib), misma forma de trabajar — un click para volcar datos de un export
al Excel maestro sin romper su formato, y un tablero para verlos sin abrirlo.

## Qué hace

| Pestaña | Para qué |
|---|---|
| **Importar submissions** | El flujo original: leer `merian_ops_-_form_-_fleet_maintenance_submissions.xlsx`, revisar fila por fila y agregarlas a `Full List 2024-2025` del maestro. Ahora también las guarda en la base local y puede **traer el histórico** que el maestro ya tiene. |
| **Full List 2024-2025** | Las inspecciones almacenadas, con filtros por año, propietario y búsqueda libre, y exportación a un Excel con el formato del maestro. |
| **Tablero de mantenimiento** | Las gráficas del maestro —`Reviewed tags/SMU` por mes y `% Inspection per month`— calculadas sobre la base local, más los cortes por tipo de equipo y por estado. |
| **Tags instalados por semana** | Consolida toda la carpeta `Tag Installed Per Week` (subcarpetas incluidas) en una sola tabla, la grafica y la exporta como el `Tag Installed - Consolidado`. |

Todo funciona en **español e inglés** y con **tema claro u oscuro**; el Excel
que se exporta sale en el idioma que tenga la ventana.

## Instalación

```bash
pip install -r requirements.txt
```

Requiere Python 3.10+.

## Uso

```bash
python run.py
```

### 1. Cargar submissions al maestro (flujo original, sin cambios)

1. **Cargar Excel de submissions...** — el archivo descargado del formulario.
2. **Seleccionar Excel destino...** — el maestro
   `Fleet Tag Inventory and Maintenance.xlsx`.
3. **Revisar / corregir** las filas en la previsualización. Cada fila trae una
   casilla *Incluir* y cualquier celda se puede editar antes de escribir.
4. **CARGAR AL EXCEL DESTINO** — agrega las filas a `Full List 2024-2025`.
   Antes de escribir se crea un **respaldo automático** del destino
   (`...backup_AAAAMMDD_HHMMSS.xlsx`).
5. Abrir el Excel y usar **Datos > Actualizar todo** para refrescar los pivotes
   y las hojas de resumen.

### 2. Base local

El software guarda lo que carga en un SQLite propio (`fleet_data.sqlite3`,
junto al código). Hay tres formas de llenarlo:

- la casilla **Guardar también en la base local** al cargar al maestro;
- el botón **GUARDAR EN LA BASE LOCAL**, que guarda sin tocar el maestro;
- **Importar histórico del maestro**, que lee la hoja `Full List 2024-2025` del
  destino seleccionado y trae los años que ya están cargados ahí.

Ese último paso es el que hace que el tablero coincida con el Excel del
cliente: sin él sólo se verían los meses del último export.

**Las filas repetidas no entran dos veces.** La base reconoce una inspección
por su contenido (fecha, equipo, FMS, tipo de dispositivo, inspector y horas),
así que da igual si la misma llega del formulario y más tarde del histórico del
maestro. Dos inspecciones *realmente* idénticas —el mismo camión revisado dos
veces el mismo día por la misma persona— sí se conservan las dos: la dinámica
cuenta inspecciones y colapsarlas cambiaría el número que se le reporta al
cliente.

### Volver a descargar el export del formulario

El export `merian_ops_-_form_-_fleet_maintenance_submissions.xlsx` es
**acumulativo**: cada descarga trae otra vez todo lo anterior más lo nuevo. Al
cargarlo, el software cruza fila por fila contra dos lugares y lo muestra en la
columna **Estado** de la previsualización:

| Estado | Qué significa | Casilla *Incluir* |
|---|---|---|
| **Nueva** | No está ni en la base local ni en el maestro | marcada |
| **Ya en la base** | La misma inspección ya fue guardada | desmarcada |
| **Ya en el maestro** | Está en `Full List 2024-2025` del destino, aunque no en la base | desmarcada |

Es decir: si vuelve a bajar el Excel con 200 submissions viejas y 58 nuevas,
quedan marcadas **sólo las 58**. Puede volver a marcar lo que quiera —el
software no decide por usted— pero si carga al maestro con filas ya cargadas,
avisa cuántas se duplicarían y pide confirmación antes de escribir.

La lectura del maestro se hace una sola vez por archivo (se guarda con su fecha
de modificación) y se puede repetir con **Volver a revisar duplicados**.

La diferencia con la base local es a propósito: la base **se defiende sola** y
nunca duplica, mientras que el maestro recibe lo que se le marque. Ahí manda el
usuario, que puede tener una razón para cargar una fila dos veces; lo que el
software garantiza es que no pase sin que se dé cuenta.

### 3. Tablero

Reproduce el bloque de indicadores de `PIVOT SUMMARY 2025`:

| Indicador | Cómo se calcula |
|---|---|
| **Revisados tags/SMU** | Equipos con **al menos una** inspección ese mes. No es el número de inspecciones: tres visitas al mismo camión en enero cuentan como un equipo revisado. Equivale al `=COUNTIF(columna_del_mes;">=1")` del maestro. |
| **% Inspección por mes** | Revisados del mes / **flota total de ese mes**. |
| **Flota mantenida** | Equipos distintos con alguna inspección en el período elegido (`=COUNTA` de las etiquetas de fila). |
| **% Mantenimiento de flota** | Flota mantenida / flota total del último mes conocido. |

La **flota total de cada mes** no se puede deducir de las inspecciones: incluye
equipos que ese mes nadie tocó. En el Excel del cliente está escrita a mano
dentro de cada fórmula (`=AA4/400`, `=AB4/405`, ...); aquí se edita desde
**Flota por mes...**, se guarda con las preferencias y arranca con los valores
que trae el maestro. Un mes sin flota configurada hereda el último valor
conocido, y si no hay ninguno el porcentaje queda vacío — no en cero: "no se
inspeccionó nada" y "no sabemos cuántos equipos había" no son lo mismo.

El selector de **período** importa para la flota mantenida. Con todos los años
juntos se suman equipos que ya no están en la flota y el porcentaje pasa de
100%; por eso el maestro tiene una hoja de dinámica por período y aquí hay un
selector.

### 4. Tags instalados por semana

**Cargar carpeta...** lee todos los `.xlsx` de la carpeta y sus subcarpetas.
Los archivos cambiaron de formato con el tiempo y el lector los unifica sin
tocar el origen:

- encabezados por nombre, con alias (`CC` → Cost Center, `Dept` → Department);
- los archivos anteriores a junio 2025 **no traen columna `TYPE`**: esos
  movimientos se toman como `NEW INSTALLATION` y quedan marcados como *tipo
  inferido*, para distinguir un dato leído de uno asumido;
- el **tipo de dispositivo** sale del tag: con dos puntos (formato MAC) es un
  SMU, si no es un TAG;
- las semanas que se solapan entre archivos no duplican movimientos.

### Corrección de fechas mal escritas

El nombre de cada archivo declara la semana que cubre, y eso es lo único que
permite afirmar que una fecha está mal: sin esa referencia, `05/12/2027` es una
fecha perfectamente válida y no hay con qué discutirla.

Midiendo las 457 filas con fecha de la carpeta del cliente contra la fecha de
cierre que declara cada nombre:

| Distancia a la semana del archivo | Filas |
|---|---|
| entre 10 días antes y 3 después | 95,6% |
| hasta un mes antes (cargas tardías legítimas) | 1,8% |
| a más de un año | 2,6% — **las 12 filas mal escritas** |

No hay nada entre "un mes antes" y "un año de distancia". Con esa separación,
la casilla **Corregir fechas con el periodo del archivo** (activada por
defecto) aplica una corrección sólo cuando el resultado es **único**:

- **Año equivocado** — `05/12/2027` en el archivo de la semana del `10/12/2025`.
  El día y el mes ya son correctos; cambiando sólo el año la fila cae dentro de
  la semana. Aparece en las dos direcciones: también hay filas de 2025 en
  archivos de 2026.
- **Día y mes invertidos** — `05/12` escrito como `12/05`. No se ha visto en
  estos archivos, pero es el error clásico entre el formato europeo y el
  americano y el período lo delata igual.

Se considera coherente todo lo que cae entre 45 días antes y 3 después, y sólo
se intenta reparar lo que queda fuera; para **aceptar** una corrección la
ventana es más estrecha (31 días antes). Para cambiar un dato hay que estar más
seguro que para dejarlo.

**Lo que no se corrige.** Las fechas imposibles (`31/06/2026` —junio tiene 30
días—, `19/19/2025` —mes 19—) entran sin fecha y marcadas. El período del
archivo puede probar que un *año* está mal, pero no puede decir qué mes quiso
escribir alguien que tecleó `19`; adivinarlo sería inventar el dato. Lo mismo
con cualquier fecha fuera de la ventana que no tenga una corrección única: se
carga tal cual y se marca como sospechosa.

**Toda corrección queda anotada** en la columna *Observación* junto con el valor
original, en la tabla y en el Excel exportado, y **los archivos de origen no se
tocan nunca**. Al terminar la carga el software informa cuántas fechas corrigió
y cuántas quedaron sin resolver.

Sobre los archivos reales corrige exactamente 12 filas —las 10 de 2027 y las 2
de 2025 en archivos de 2026—, todas cambiando sólo el año y sin tocar ningún
otro campo ni ninguna otra fila.

> Al reprocesar una carpeta ya cargada con la corrección recién activada, las
> filas corregidas entran como **nuevas** (la fecha forma parte de la clave de
> deduplicación) y las versiones viejas siguen en la base. Use **Vaciar tags
> almacenados** y vuelva a cargar la carpeta.

Los conteos de instalación excluyen los retiros, igual que el
`=SUMPRODUCT(... <>"REMOVAL")` del consolidado del cliente.

**Agrupar por** — los archivos llegan por semana, pero la pregunta cambia con
quien pregunta. El selector **Diario / Semanal / Mensual / Anual** reagrupa las
gráficas y el resumen del Excel exportado; la tabla de movimientos no cambia,
porque es el detalle y ahí cada fila es un movimiento.

- Con grano **mensual** el resumen conserva el par `Año | Mes` del consolidado
  del cliente; con los otros usa una sola columna de período, y el día y la
  semana van como **fecha real** para poder ordenarlos y armar una dinámica en
  Excel.
- En pantalla la gráfica muestra las últimas cubetas **con datos** (45 días, 30
  semanas, 36 meses) y no las últimas a secas, para que un dato suelto muy
  lejano no deje la vista vacía. Los huecos intermedios se conservan: un hueco
  es información y, cuando es anómalo, es lo que delata un error en el origen.

### 5. Exportar

- **Exportar reporte Excel...** (pestaña Full List) genera un libro con la hoja
  de inspecciones (formato de `Full List 2024-2025`), la hoja de resumen
  dinámico (equipos × meses + bloque de indicadores + la **torta** y las
  **barras con la línea de %**) y una hoja de notas.
- **Exportar consolidado Excel...** (pestaña Tags) genera la tabla de
  movimientos unificada y el **resumen de instalación** por mes con su gráfica
  de barras.

Las gráficas son **nativas de Excel**, no imágenes: quien reciba el archivo
puede filtrar la tabla y ver la gráfica moverse. El libro sale siempre con la
paleta clara aunque la ventana esté en tema oscuro — se imprime y se comparte.

## Estructura del proyecto

```
run.py             Punto de entrada
app_qt.py          Interfaz grafica PySide6 (cuatro pestañas)
i18n.py            Textos español/inglés, formatos de fecha y numero
theme.py           Paletas clara y oscura + hoja de estilos de Qt
settings.py        Preferencias persistidas (idioma, tema, flota por mes)
store.py           Base local SQLite (inspecciones y movimientos de tag)
source_reader.py   Lectura del export del formulario y del histórico del maestro
tag_reader.py      Lectura de los semanales 'Inventory Tag Installed'
mapping.py         Mapeo formulario -> columnas de 'Full List 2024-2025'
analytics.py       Dinámica, indicadores por mes y resumen de tags
charts.py          Gráficas de matplotlib embebidas
excel_writer.py    Escritura en el maestro preservando estilos, fórmulas y Table3
report_export.py   Excel exportados, con gráficas nativas
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
[`mapping.py`](mapping.py). Las columnas sin equivalente directo quedan
editables en la previsualización.

### Decisiones de carga

- **Fecha de la fila**: se usa `Date And Time Of Revision` (la fecha de la
  inspección en campo); si falta, se cae a `Submitted At`.
- **Horas/ODO**: se usa el `Machine Hourmeter Reading`; si está vacío, el
  `Machine Kilometer Reading`.
- **Duplicados al escribir en el maestro**: la herramienta agrega las filas
  marcadas, y las que ya están cargadas llegan **desmarcadas** desde la
  previsualización (ver *Volver a descargar el export del formulario*). Si aun
  así se marcan, el software avisa cuántas se duplicarían antes de escribir.

## Cómo se preserva la estructura del maestro

`excel_writer.py` abre el maestro con openpyxl y, por cada fila nueva:

- copia estilos (fuente, relleno, bordes, alineación y **formato de número**)
  desde la última fila de datos como plantilla;
- reproduce las fórmulas de A y D, incluyendo la **fórmula de matriz**
  (`ArrayFormula`) de `Verified` con su `ref` propio por fila;
- extiende el rango de la tabla `Table3` para incluir las filas nuevas (sin
  esto, Excel reporta *"We found a problem with some content"*).

Validado contra los archivos reales: las **10 hojas**, las **5 tablas
dinámicas**, sus **4 cachés** y los **5 gráficos** se conservan tras la
escritura.

### Limitación conocida

openpyxl conserva los pivotes y los gráficos, pero **descarta el estilo/color
personalizado de los gráficos** (`colors*.xml` / `style*.xml`): tras importar,
los gráficos de los dashboards pueden volver a la paleta por defecto de Excel.
Los datos y los pivotes quedan intactos, y el **respaldo automático** conserva
el original por si se necesita recuperar el formato.

## Diferencias contra el Excel del cliente

Los indicadores del tablero se validaron contra
`260814_Fleet Tag Inventory and Maintenance.xlsx` mes a mes. Coinciden en 17 de
los 19 meses de 2025-2026; las dos diferencias son del Excel, no del cálculo:

- **feb-26**: el maestro cuenta 121 y el software 135. Las fórmulas
  `COUNTIF` de ese bloque quedaron con el rango viejo (`O5:O758`) mientras las
  de los meses siguientes ya usan `P5:P831`; el mes se contó sobre una parte de
  las filas.
- **abr-25**: el maestro cuenta 212 y el software 211. Hay equipos cuyo ID está
  escrito como número en unas filas y como texto en otras; la tabla dinámica de
  Excel los toma como dos equipos distintos y el software los agrupa como uno.

Por lo mismo, `Total maintained fleet` da 826 contra los 827 del maestro. La
hoja `Verified` del maestro **no se reproduce** en el Excel exportado: se
calcula cruzando cada equipo contra `TAG History` y `SMU History`, que viven en
el maestro y no en la base local.

## Pruebas

```bash
pip install pytest
pytest -q
```

Las pruebas usan workbooks sintéticos creados en tiempo de ejecución y una base
SQLite temporal: no dependen de los archivos reales del cliente ni tocan la
base ni las preferencias del usuario.
