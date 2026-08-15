# -*- coding: utf-8 -*-
"""
Internacionalizacion espanol / ingles.

Tres funciones cubren todo:

  t(clave)        -> texto de la INTERFAZ (pestanas, botones, encabezados).
  tr_value(texto) -> traduce un valor de DOMINIO que el motor produjo con su
                     etiqueta canonica (tipos de movimiento, tipos de equipo).
  month_label(d)  -> 'ene-25' / 'Jan-25', la etiqueta de mes de los tableros.

La separacion importa: los modulos de datos (`store`, `analytics`,
`tag_reader`) no saben en que idioma esta la ventana. Producen siempre
etiquetas canonicas y la capa de presentacion las traduce al vuelo; asi el
mismo calculo sirve para la ventana en ingles, para el Excel en espanol y para
un script sin interfaz.

Lo que NO se traduce:

  - Los identificadores de campo (Vehicle ID, FMS ID, Tag, Cost Center): son la
    jerga con la que se cruza contra el FMS y contra los inventarios de
    Newmont; traducirlos rompe el cruce.
  - Los valores de estado que vienen del maestro ('VIU OK', 'NO VIU', 'Y', 'N',
    'N/A') y los tipos de movimiento tal como se escriben en la hoja
    ('NEW INSTALLATION', 'REMOVAL'): se cargan y se vuelven a escribir tal
    cual. `tr_value` solo los traduce para MOSTRARLOS en tableros y leyendas.
"""
from __future__ import annotations

import datetime

ES = "es"
EN = "en"
LANGUAGES = (ES, EN)
LANGUAGE_NAMES = {ES: "Espanol", EN: "English"}

_current = ES


def set_language(lang: str) -> None:
    global _current
    _current = lang if lang in LANGUAGES else ES


def current() -> str:
    return _current


def toggled() -> str:
    """El otro idioma. Para el boton que alterna ES/EN."""
    return EN if _current == ES else ES


def t(key: str, **kwargs) -> str:
    """Texto de interfaz. Si falta la clave devuelve la clave (visible en QA)."""
    entry = _UI.get(key)
    if entry is None:
        return key
    text = entry.get(_current) or entry.get(ES) or key
    return text.format(**kwargs) if kwargs else text


def tr_value(value) -> str:
    """Traduce un valor canonico de dominio para mostrarlo."""
    if value in (None, ""):
        return ""
    return _VALUES.get(str(value).strip().upper(), {}).get(
        _current, str(value))


# ---------------------------------------------------------------------------
# Fechas y numeros
# ---------------------------------------------------------------------------
#
# La FECHA se escribe dd/mm/aaaa en los dos idiomas. No es un descuido: los
# archivos de origen (el export del formulario y los semanales de tags) traen
# dd/mm/aaaa, y este software se lee al lado de ellos. Con mm/dd/aaaa en la
# ventana en ingles, '01/02/2026' seria ambiguo — nadie podria saber si es el 1
# de febrero o el 2 de enero sin saber con que idioma se genero la pantalla.
#
# El numero SI cambia: 1.234,5 en espanol contra 1,234.5 en ingles.

_MONTHS_SHORT = {
    ES: ["ene", "feb", "mar", "abr", "may", "jun",
         "jul", "ago", "sep", "oct", "nov", "dic"],
    EN: ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
}

_MONTHS_LONG = {
    ES: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
    EN: ["January", "February", "March", "April", "May", "June", "July",
         "August", "September", "October", "November", "December"],
}


def date_format() -> str:
    return "%d/%m/%Y"


def excel_date_format() -> str:
    return "DD/MM/YYYY"


def month_short(month: int) -> str:
    """'ene' / 'Jan'. `month` es 1..12."""
    return _MONTHS_SHORT[_current][int(month) - 1]


def month_long(month: int) -> str:
    return _MONTHS_LONG[_current][int(month) - 1]


def month_label(value) -> str:
    """Etiqueta de mes del tablero: 'ene-25' / 'Jan-25'.

    Acepta un date/datetime o la clave canonica 'AAAA-MM' con la que
    `analytics` indexa los meses.
    """
    if isinstance(value, (datetime.date, datetime.datetime)):
        year, month = value.year, value.month
    else:
        text = str(value)
        try:
            year, month = int(text[:4]), int(text[5:7])
        except (ValueError, IndexError):
            return text
    return "%s-%02d" % (month_short(month), year % 100)


def fmt_date(value) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime(date_format())
    return str(value)


def fmt_number(value, decimals: int = 0) -> str:
    """1.234,5 en espanol; 1,234.5 en ingles."""
    if value in (None, ""):
        return ""
    try:
        text = "{:,.{d}f}".format(float(value), d=decimals)
    except (TypeError, ValueError):
        return str(value)
    if _current == ES:
        # Se pasa por un marcador para no pisar las comas ya convertidas.
        text = text.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return text


def fmt_pct(value, decimals: int = 0) -> str:
    if value in (None, ""):
        return ""
    try:
        return fmt_number(float(value) * 100.0, decimals) + "%"
    except (TypeError, ValueError):
        return str(value)


# ---------------------------------------------------------------------------
# Valores de dominio
# ---------------------------------------------------------------------------
_VALUES = {
    "NEW INSTALLATION": {ES: "Instalacion nueva", EN: "New installation"},
    "REPLACEMENT":      {ES: "Reemplazo", EN: "Replacement"},
    "REMOVAL":          {ES: "Retiro", EN: "Removal"},
    "TAG UPDATED":      {ES: "Tag actualizado", EN: "Tag updated"},
    "SMU":              {ES: "SMU", EN: "SMU"},
    "TAG":              {ES: "TAG", EN: "TAG"},
    "Y":                {ES: "Si", EN: "Yes"},
    "N":                {ES: "No", EN: "No"},
    "N/A":              {ES: "N/A", EN: "N/A"},
    # Observaciones que emite `tag_reader` (codigo canonico + detalle).
    "BAD_DATE":         {ES: "Fecha invalida en el archivo de origen",
                         EN: "Invalid date in the source file"},
    "TYPE_UNKNOWN":     {ES: "Movimiento no reconocido, se asumio instalacion",
                         EN: "Unrecognized movement, assumed installation"},
    "TYPE_ASSUMED":     {ES: "Archivo sin columna TYPE, se asumio instalacion",
                         EN: "File has no TYPE column, assumed installation"},
    "FUTURE_DATE":      {ES: "Fecha posterior a hoy, revisar el origen",
                         EN: "Date later than today, check the source"},
}


def tr_note(note) -> str:
    """Traduce una observacion 'CODIGO:detalle' de `tag_reader`.

    El detalle (el texto original que no se pudo interpretar) se conserva tal
    cual: es un dato del archivo del cliente, no una etiqueta del software.
    """
    if not note:
        return ""
    parts = []
    for chunk in str(note).split():
        code, _, detail = chunk.partition(":")
        text = tr_value(code)
        parts.append("%s (%s)" % (text, detail) if detail else text)
    return "; ".join(parts)


# ---------------------------------------------------------------------------
# Textos de interfaz
# ---------------------------------------------------------------------------
_UI = {
    # -- ventana --------------------------------------------------------
    "app.title": {
        ES: "Cargador de Mantenimiento de Flota  -  Newmont Merian",
        EN: "Fleet Maintenance Loader  -  Newmont Merian"},
    "app.ready": {
        ES: "Cargue el Excel de submissions y seleccione el Excel destino.",
        EN: "Load the submissions Excel and pick the target Excel."},

    # -- barra superior -------------------------------------------------
    "top.language": {ES: "Idioma", EN: "Language"},
    "top.theme": {ES: "Tema", EN: "Theme"},
    "top.theme_light": {ES: "Claro", EN: "Light"},
    "top.theme_dark": {ES: "Oscuro", EN: "Dark"},
    "top.stored": {
        ES: "Base local: {inspections} inspecciones · {tags} movimientos de tag",
        EN: "Local database: {inspections} inspections · {tags} tag movements"},

    # -- pestanas -------------------------------------------------------
    "tab.import": {ES: "Importar submissions", EN: "Import submissions"},
    "tab.fulllist": {ES: "Full List 2024-2025", EN: "Full List 2024-2025"},
    "tab.dashboard": {ES: "Tablero de mantenimiento",
                      EN: "Maintenance dashboard"},
    "tab.tags": {ES: "Tags instalados por semana",
                 EN: "Tags installed per week"},

    # -- pestana importar ------------------------------------------------
    "import.files": {ES: "Archivos", EN: "Files"},
    "import.btn_source": {ES: "Cargar Excel de submissions...",
                          EN: "Load submissions Excel..."},
    "import.btn_target": {ES: "Seleccionar Excel destino...",
                          EN: "Pick target Excel..."},
    "import.state": {ES: "Estado:", EN: "Status:"},
    "import.chip_source": {ES: "Submissions", EN: "Submissions"},
    "import.chip_target": {ES: "Excel destino", EN: "Target Excel"},
    "import.chip_missing": {ES: "sin seleccionar", EN: "not selected"},
    "import.chip_rows": {ES: "{n} filas", EN: "{n} rows"},
    "import.preview": {
        ES: "Previsualizacion  (filas que se agregaran a 'Full List 2024-2025')",
        EN: "Preview  (rows to be added to 'Full List 2024-2025')"},
    "import.preview_hint": {
        ES: "Revise y corrija los valores antes de cargar. Las columnas 'Date' "
            "(A) y 'Verified' (D) del destino son formulas y se calculan "
            "solas; aqui no se muestran.",
        EN: "Review and fix the values before loading. The target's 'Date' (A) "
            "and 'Verified' (D) columns are formulas and recalculate on their "
            "own; they are not shown here."},
    "import.col_include": {ES: "Incluir", EN: "Include"},
    "import.check_all": {ES: "Marcar todas", EN: "Check all"},
    "import.uncheck_all": {ES: "Desmarcar todas", EN: "Uncheck all"},
    "import.count": {ES: "{n} submissions cargadas.",
                     EN: "{n} submissions loaded."},
    "import.backup": {ES: "Crear respaldo del destino antes de escribir",
                      EN: "Back up the target before writing"},
    "import.also_store": {
        ES: "Guardar tambien en la base local del software",
        EN: "Also save to the software's local database"},
    "import.btn_append": {ES: "CARGAR AL EXCEL DESTINO",
                          EN: "LOAD INTO TARGET EXCEL"},
    "import.btn_store": {ES: "GUARDAR EN LA BASE LOCAL",
                         EN: "SAVE TO LOCAL DATABASE"},
    "import.btn_history": {ES: "Importar historico del maestro",
                           EN: "Import history from the master"},
    "import.history_hint": {
        ES: "Trae a la base local las filas que el maestro ya tiene en 'Full "
            "List 2024-2025', para que el tablero muestre todo el historico y "
            "no solo las submissions nuevas. Las repetidas se omiten.",
        EN: "Pulls the rows the master already has in 'Full List 2024-2025' "
            "into the local database, so the dashboard shows the whole "
            "history and not only the new submissions. Duplicates are "
            "skipped."},
    "msg.history_loaded": {
        ES: "Se leyeron {read} fila(s) de '{sheet}'.\n"
            "{added} nuevas guardadas, {skipped} ya estaban en la base.",
        EN: "{read} row(s) read from '{sheet}'.\n"
            "{added} new rows saved, {skipped} were already stored."},
    "msg.pick_target_first": {
        ES: "Seleccione primero el Excel destino (el maestro) para poder leer "
            "su historico.",
        EN: "Pick the target Excel (the master) first so its history can be "
            "read."},
    "prog.history_title": {ES: "Leyendo el historico del maestro",
                           EN: "Reading the master's history"},
    "prog.history": {ES: "Leyendo la hoja '{sheet}'...",
                     EN: "Reading sheet '{sheet}'..."},
    "full.limit": {
        ES: "Se muestran las {shown} filas mas recientes de {total}. Los "
            "calculos y la exportacion usan todas.",
        EN: "Showing the {shown} most recent rows out of {total}. Calculations "
            "and the export use all of them."},

    # -- pestana full list -----------------------------------------------
    "full.title": {ES: "Inspecciones almacenadas en el software",
                   EN: "Inspections stored in the software"},
    "full.hint": {
        ES: "Estas filas son las que el software guarda por su cuenta. Se "
            "exportan con el mismo formato de la hoja 'Full List 2024-2025' "
            "del maestro.",
        EN: "These are the rows the software keeps on its own. They export "
            "with the same layout as the master's 'Full List 2024-2025' "
            "sheet."},
    "full.filter_year": {ES: "Ano:", EN: "Year:"},
    "full.filter_month": {ES: "Mes:", EN: "Month:"},
    "full.filter_owner": {ES: "Propietario:", EN: "Owner:"},
    "full.filter_search": {ES: "Buscar equipo / inspector:",
                           EN: "Search equipment / inspector:"},
    "full.all": {ES: "(todos)", EN: "(all)"},
    "full.rows": {ES: "{shown} de {total} filas", EN: "{shown} of {total} rows"},
    "full.btn_export": {ES: "Exportar reporte Excel...",
                        EN: "Export Excel report..."},
    "full.btn_delete": {ES: "Eliminar filas seleccionadas",
                        EN: "Delete selected rows"},
    "full.btn_refresh": {ES: "Actualizar", EN: "Refresh"},

    # -- pestana tablero --------------------------------------------------
    "dash.card_inspections": {ES: "Inspecciones", EN: "Inspections"},
    "dash.card_inspections_hint": {ES: "filas almacenadas", EN: "stored rows"},
    "dash.card_fleet": {ES: "Flota mantenida", EN: "Maintained fleet"},
    "dash.card_fleet_hint": {ES: "equipos distintos inspeccionados",
                             EN: "distinct equipment inspected"},
    "dash.card_pct": {ES: "% Mantenimiento de flota",
                      EN: "% Fleet maintenance"},
    "dash.card_pct_hint": {ES: "sobre la flota total registrada",
                           EN: "over the total registered fleet"},
    "dash.card_last": {ES: "Ultimo mes", EN: "Last month"},
    "dash.card_last_hint": {ES: "equipos revisados", EN: "equipment reviewed"},
    "dash.chart": {ES: "Grafica:", EN: "Chart:"},
    "dash.chart_bars": {ES: "Revisados y % de inspeccion (barras + linea)",
                        EN: "Reviewed and % inspection (bars + line)"},
    "dash.chart_pie": {ES: "Reparto de revisados por mes (torta)",
                       EN: "Reviewed share by month (pie)"},
    "dash.chart_equipment": {ES: "Inspecciones por tipo de equipo",
                             EN: "Inspections by equipment type"},
    "dash.chart_status": {ES: "Inspecciones por estado (Status)",
                          EN: "Inspections by status"},
    "dash.months": {ES: "Meses:", EN: "Months:"},
    "dash.months_all": {ES: "(todos)", EN: "(all)"},
    "dash.months_last": {ES: "ultimos {n}", EN: "last {n}"},
    "dash.fleet_size": {ES: "Flota total del mes:", EN: "Month fleet size:"},
    "dash.fleet_size_hint": {
        ES: "Denominador de '% Inspection per month'. Si se deja en 0 se usa "
            "el ultimo valor conocido.",
        EN: "Denominator of '% Inspection per month'. Left at 0 it falls back "
            "to the last known value."},
    "dash.btn_fleet_sizes": {ES: "Flota por mes...", EN: "Fleet size by month..."},
    "dash.table_month": {ES: "Mes", EN: "Month"},
    "dash.table_reviewed": {ES: "Revisados (tags/SMU)", EN: "Reviewed tags/SMU"},
    "dash.table_fleet": {ES: "Flota total", EN: "Total fleet"},
    "dash.table_pct": {ES: "% Inspeccion del mes", EN: "% Inspection per month"},
    "dash.table_rows": {ES: "Inspecciones", EN: "Inspections"},

    # -- pestana tags ------------------------------------------------------
    "tags.title": {ES: "Consolidado de 'Tag Installed Per Week'",
                   EN: "'Tag Installed Per Week' consolidation"},
    "tags.hint": {
        ES: "Cargue la carpeta completa (incluye subcarpetas) o archivos "
            "sueltos. Los archivos antiguos sin columna TYPE se toman como "
            "'NEW INSTALLATION' y quedan marcados como tipo inferido.",
        EN: "Load the whole folder (subfolders included) or individual files. "
            "Older files with no TYPE column are taken as 'NEW INSTALLATION' "
            "and flagged as inferred type."},
    "tags.btn_folder": {ES: "Cargar carpeta...", EN: "Load folder..."},
    "tags.btn_files": {ES: "Cargar archivos...", EN: "Load files..."},
    "tags.btn_export": {ES: "Exportar consolidado Excel...",
                        EN: "Export consolidated Excel..."},
    "tags.btn_clear": {ES: "Vaciar tags almacenados",
                       EN: "Clear stored tag movements"},
    "tags.filter_type": {ES: "Movimiento:", EN: "Movement:"},
    "tags.filter_device": {ES: "Dispositivo:", EN: "Device:"},
    "tags.filter_dept": {ES: "Departamento:", EN: "Department:"},
    "tags.rows": {ES: "{shown} de {total} movimientos",
                  EN: "{shown} of {total} movements"},
    "tags.chart": {ES: "Grafica:", EN: "Chart:"},
    "tags.chart_month": {ES: "Instalados por mes (SMU / TAG)",
                         EN: "Installed per month (SMU / TAG)"},
    "tags.chart_type": {ES: "Movimientos por mes y tipo",
                        EN: "Movements per month and type"},
    "tags.chart_dept": {ES: "Instalados por departamento",
                        EN: "Installed by department"},
    "tags.chart_week": {ES: "Instalados por semana", EN: "Installed per week"},
    "tags.card_total": {ES: "Movimientos", EN: "Movements"},
    "tags.card_total_hint": {ES: "filas almacenadas", EN: "stored rows"},
    "tags.card_installed": {ES: "Instalados", EN: "Installed"},
    "tags.card_installed_hint": {ES: "altas y reemplazos",
                                 EN: "installations and replacements"},
    "tags.card_removed": {ES: "Retirados", EN: "Removed"},
    "tags.card_removed_hint": {ES: "movimientos de retiro",
                               EN: "removal movements"},
    "tags.card_files": {ES: "Archivos", EN: "Files"},
    "tags.card_files_hint": {ES: "semanales consolidados",
                             EN: "weekly files consolidated"},

    # -- graficas ----------------------------------------------------------
    "chart.empty": {ES: "Sin datos para graficar",
                    EN: "No data to chart"},
    "chart.legend_reviewed": {ES: "Revisados tags/SMU",
                              EN: "Reviewed tags/SMU"},
    "chart.legend_pct": {ES: "% Inspeccion por mes",
                         EN: "% Inspection per month"},
    "chart.legend_smu": {ES: "SMU", EN: "SMU"},
    "chart.legend_tag": {ES: "TAG", EN: "TAG"},
    "chart.axis_count": {ES: "Cantidad", EN: "Count"},
    "chart.axis_pct": {ES: "% de inspeccion", EN: "% inspection"},
    "chart.caption_total": {ES: "Total del periodo: {n}",
                            EN: "Period total: {n}"},
    "chart.caption_last_months": {ES: "se muestran los ultimos {n} meses",
                                  EN: "showing the last {n} months"},
    "chart.other": {ES: "OTROS", EN: "OTHER"},

    # -- columnas de las tablas (mismas etiquetas que el Excel) -----------
    "col.date": {ES: "Fecha", EN: "Date"},
    "col.date_target": {ES: "Fecha (dd/mm/aa)", EN: "Date (mm/dd/yy)"},
    "col.vehicle": {ES: "Vehicle ID", EN: "Vehicle ID"},
    "col.fms": {ES: "FMS ID", EN: "FMS ID"},
    "col.fitted": {ES: "Sistema segun estandar (Y/N)",
                   EN: "System fitted as per standard (Y/N)"},
    "col.hours": {ES: "Horas / ODO del equipo", EN: "Equipment Hours/ODO"},
    "col.fms_hours": {ES: "Horas FMS", EN: "FMS Hours"},
    "col.status": {ES: "Estado", EN: "Status"},
    "col.inlets": {ES: "# Entradas", EN: "# INLETS"},
    "col.addl_locked": {ES: "Entradas adicionales aseguradas (Y/N)",
                        EN: "Are aditional inlets locked?(Y/N)"},
    "col.drain_locked": {ES: "Valvulas de drenaje aseguradas (Y/N)",
                         EN: "Drain valves locked?(Y/N)"},
    "col.leaking": {ES: "Receptor de llenado con fuga (Y/N)",
                    EN: "Fast fill receiver leaking?(Y/N)"},
    "col.smu_tags": {ES: "# SMU/TAGS", EN: "# SMU/TAGS"},
    "col.equipment": {ES: "Tipo de equipo", EN: "EQUIPMENT TYPE"},
    "col.remarks": {ES: "Observaciones", EN: "REMARKS"},
    "col.inspectors": {ES: "Inspectores", EN: "Inspectors"},
    "col.owner": {ES: "Propietario", EN: "OWNER"},
    "col.remedial": {ES: "Acciones correctivas / actualizaciones",
                     EN: "REMEDIAL ACTIONS / UPDPATES"},
    "col.source": {ES: "Archivo origen", EN: "Source file"},

    "col.num": {ES: "#", EN: "#"},
    "col.move_type": {ES: "Movimiento", EN: "TYPE"},
    "col.equipment_id": {ES: "ID", EN: "ID"},
    "col.tag": {ES: "Tag", EN: "Tag"},
    "col.device": {ES: "Tipo de dispositivo", EN: "Device Type"},
    "col.cost_center": {ES: "Centro de costo", EN: "Cost Center"},
    "col.department": {ES: "Departamento", EN: "Department"},
    "col.product": {ES: "Producto", EN: "Product"},
    "col.changed_by": {ES: "Modificado por", EN: "Changed by"},
    "col.year": {ES: "Ano", EN: "Year"},
    "col.month": {ES: "Mes", EN: "Month"},
    "col.week": {ES: "Semana (lunes)", EN: "Week (Monday)"},
    "col.sheet": {ES: "Hoja", EN: "Sheet"},
    "col.inferred": {ES: "Tipo inferido", EN: "Inferred type"},
    "col.note": {ES: "Observacion", EN: "Notes"},
    "col.total": {ES: "Total", EN: "Total"},
    "col.grand_total": {ES: "Total general", EN: "Grand Total"},

    # -- hojas y titulos del Excel exportado ------------------------------
    "sheet.fulllist": {ES: "Lista Completa 2024-2025",
                       EN: "Full List 2024-2025"},
    "sheet.pivot": {ES: "Resumen Dinamico", EN: "PIVOT SUMMARY"},
    "sheet.charts": {ES: "Graficas", EN: "Charts"},
    "sheet.taginstalled": {ES: "Tags Instalados", EN: "Tag Installed"},
    "sheet.tagsummary": {ES: "Resumen de Instalacion",
                         EN: "Installation Summary"},
    "sheet.notes": {ES: "Notas", EN: "Notes"},

    "xls.title_fulllist": {
        ES: "Mantenimiento de flota - Lista completa de inspecciones",
        EN: "Fleet maintenance - Full list of inspections"},
    "xls.title_pivot": {ES: "Resumen dinamico {year}",
                        EN: "Pivot summary {year}"},
    "xls.title_taginstalled": {ES: "Tags instalados - consolidado semanal",
                               EN: "Tags installed - weekly consolidation"},
    "xls.title_tagsummary": {ES: "Resumen de instalacion",
                             EN: "Installation Summary"},
    "xls.generated": {ES: "Generado: {when}", EN: "Generated: {when}"},
    "xls.rows": {ES: "Filas: {n}", EN: "Rows: {n}"},
    "xls.source": {ES: "Origen: base local del software",
                   EN: "Source: software local database"},
    "xls.kpi_reviewed": {ES: "Revisados tags/SMU", EN: "Reviewed tags/SMU"},
    "xls.kpi_pct": {ES: "% Inspeccion por mes", EN: "% Inspection per month"},
    "xls.kpi_fleet": {ES: "Flota total mantenida", EN: "Total maintained fleet"},
    "xls.kpi_pct_fleet": {ES: "% Mantenimiento de flota",
                          EN: "% Fleet maintenance"},
    "xls.chart_pie": {ES: "Revisados tags/SMU  Newmont",
                      EN: "Reviewed tags/SMU  Newmont"},
    "xls.chart_bars": {ES: "Newmont y sus contratistas",
                       EN: "Newmont & BPs"},
    "xls.chart_taginstalled": {ES: "Resumen de instalacion",
                               EN: "Installation Summary"},
    "xls.axis_count": {ES: "Cantidad", EN: "Count"},
    "xls.axis_pct": {ES: "% de inspeccion", EN: "% inspection"},
    "xls.axis_month": {ES: "Mes", EN: "Month"},
    "xls.notes_title": {ES: "Como se construyo este archivo",
                        EN: "How this file was built"},

    # -- notas del Excel exportado ----------------------------------------
    "note.fulllist": {
        ES: "Hoja de inspecciones: una fila por inspeccion almacenada en el "
            "software, con las mismas columnas de la hoja 'Full List "
            "2024-2025' del maestro.",
        EN: "Inspections sheet: one row per inspection stored in the "
            "software, with the same columns as the master's 'Full List "
            "2024-2025' sheet."},
    "note.pivot": {
        ES: "Resumen dinamico: cuenta de inspecciones por equipo y por mes. "
            "'Revisados tags/SMU' es la cantidad de equipos con al menos una "
            "inspeccion en el mes.",
        EN: "Pivot summary: inspection count per equipment and per month. "
            "'Reviewed tags/SMU' is the number of equipment with at least one "
            "inspection in the month."},
    "note.pct": {
        ES: "'% Inspeccion por mes' = revisados del mes / flota total del mes. "
            "La flota total de cada mes se configura en el tablero del "
            "software.",
        EN: "'% Inspection per month' = month reviewed / month total fleet. "
            "Each month's total fleet is configured in the software "
            "dashboard."},
    "note.undated": {
        ES: "{n} fila(s) quedaron sin fecha porque el archivo de origen traia "
            "una fecha que no se pudo interpretar. No se corrigieron ni se "
            "descartaron: quedan al final de la hoja con su observacion.",
        EN: "{n} row(s) have no date because the source file carried a date "
            "that could not be read. They were neither fixed nor dropped: "
            "they sit at the end of the sheet with their note."},
    "note.no_verified": {
        ES: "La columna 'Verified' del maestro no se reproduce: se calcula "
            "cruzando el equipo contra las hojas 'TAG History' y 'SMU "
            "History', que viven en el Excel maestro y no en este software.",
        EN: "The master's 'Verified' column is not reproduced: it is computed "
            "by matching the equipment against the 'TAG History' and 'SMU "
            "History' sheets, which live in the master Excel and not in this "
            "software."},
    "note.tags": {
        ES: "Los archivos anteriores a junio 2025 no traen columna TYPE; esos "
            "movimientos se tomaron como 'NEW INSTALLATION' y quedan marcados "
            "en la columna de tipo inferido.",
        EN: "Files older than June 2025 have no TYPE column; those movements "
            "were taken as 'NEW INSTALLATION' and are flagged in the inferred "
            "type column."},
    "note.tags_device": {
        ES: "El tipo de dispositivo se deduce del tag: con dos puntos (formato "
            "MAC) es SMU, si no es TAG.",
        EN: "Device type is derived from the tag: with colons (MAC format) it "
            "is SMU, otherwise TAG."},
    "note.tags_dedupe": {
        ES: "Los movimientos repetidos entre archivos con solape se cargan una "
            "sola vez (misma combinacion de movimiento, fecha, ID y tag).",
        EN: "Movements repeated across overlapping files are loaded once "
            "(same combination of movement, date, ID and tag)."},

    # -- dialogos ----------------------------------------------------------
    "dlg.error": {ES: "Error", EN: "Error"},
    "dlg.warning": {ES: "Aviso", EN: "Warning"},
    "dlg.info": {ES: "Informacion", EN: "Information"},
    "dlg.confirm": {ES: "Confirmar", EN: "Confirm"},
    "dlg.open_source": {ES: "Cargar Excel de submissions del formulario",
                        EN: "Load the form's submissions Excel"},
    "dlg.open_target": {
        ES: "Seleccionar Excel destino (Fleet Tag Inventory and Maintenance)",
        EN: "Pick target Excel (Fleet Tag Inventory and Maintenance)"},
    "dlg.open_tag_folder": {ES: "Seleccionar la carpeta 'Tag Installed Per Week'",
                            EN: "Pick the 'Tag Installed Per Week' folder"},
    "dlg.open_tag_files": {ES: "Seleccionar archivos semanales de tags",
                           EN: "Pick weekly tag files"},
    "dlg.save_report": {ES: "Guardar reporte de mantenimiento",
                        EN: "Save maintenance report"},
    "dlg.save_tags": {ES: "Guardar consolidado de tags",
                      EN: "Save tag consolidation"},
    "dlg.excel_filter": {ES: "Excel (*.xlsx)", EN: "Excel (*.xlsx)"},

    "msg.read_error": {ES: "No se pudo leer el Excel de submissions:\n{err}",
                       EN: "The submissions Excel could not be read:\n{err}"},
    "msg.no_data": {
        ES: "El archivo no contiene submissions reconocibles en la hoja "
            "'Form Submissions'.",
        EN: "The file has no recognizable submissions in the 'Form "
            "Submissions' sheet."},
    "msg.loaded": {ES: "{n} submissions cargadas desde {file}",
                   EN: "{n} submissions loaded from {file}"},
    "msg.target_set": {ES: "Destino: {file}", EN: "Target: {file}"},
    "msg.nothing_checked": {ES: "No hay filas marcadas para cargar.",
                            EN: "No rows are checked for loading."},
    "msg.confirm_append": {
        ES: "Se agregaran {n} fila(s) a la hoja '{sheet}' de:\n{file}\n\n"
            "Se hara un respaldo automatico del destino.\n\n Continuar?",
        EN: "{n} row(s) will be added to sheet '{sheet}' of:\n{file}\n\n"
            "The target will be backed up automatically.\n\nContinue?"},
    "msg.append_done": {
        ES: "Se agregaron {n} fila(s) a '{sheet}'.\nFilas {first} a {last}.",
        EN: "{n} row(s) were added to '{sheet}'.\nRows {first} to {last}."},
    "msg.backup_made": {ES: "Respaldo creado:\n{path}",
                        EN: "Backup created:\n{path}"},
    "msg.refresh_hint": {
        ES: "Abra el Excel y use 'Datos > Actualizar todo' para refrescar los "
            "pivotes y las hojas de resumen.",
        EN: "Open the Excel and use 'Data > Refresh All' to refresh the pivots "
            "and summary sheets."},
    "msg.file_in_use": {
        ES: "No se pudo escribir el Excel destino. Es probable que este "
            "abierto en Excel. Cierrelo y vuelva a intentar.",
        EN: "The target Excel could not be written. It is probably open in "
            "Excel. Close it and try again."},
    "msg.stored": {
        ES: "Se guardaron {added} inspeccion(es) en la base local.\n"
            "{skipped} ya estaban almacenadas y se omitieron.",
        EN: "{added} inspection(s) saved to the local database.\n"
            "{skipped} were already stored and were skipped."},
    "msg.tags_stored": {
        ES: "Se leyeron {files} archivo(s): {added} movimiento(s) nuevos, "
            "{skipped} repetidos omitidos.",
        EN: "{files} file(s) read: {added} new movement(s), {skipped} "
            "duplicates skipped."},
    "msg.tags_none": {
        ES: "No se encontraron archivos de tags legibles en la seleccion.",
        EN: "No readable tag files were found in the selection."},
    "msg.export_done": {ES: "Reporte generado:\n{path}",
                        EN: "Report generated:\n{path}"},
    "msg.export_empty": {
        ES: "No hay datos almacenados para exportar. Guarde primero "
            "inspecciones o movimientos de tag en la base local.",
        EN: "There is no stored data to export. Save inspections or tag "
            "movements to the local database first."},
    "msg.confirm_delete": {
        ES: "Se eliminaran {n} fila(s) de la base local del software. El Excel "
            "maestro no se toca.\n\n Continuar?",
        EN: "{n} row(s) will be deleted from the software's local database. "
            "The master Excel is not touched.\n\nContinue?"},
    "msg.confirm_clear_tags": {
        ES: "Se eliminaran los {n} movimientos de tag almacenados. Los "
            "archivos semanales no se tocan.\n\n Continuar?",
        EN: "The {n} stored tag movements will be deleted. The weekly files "
            "are not touched.\n\nContinue?"},
    "msg.deleted": {ES: "{n} fila(s) eliminadas.", EN: "{n} row(s) deleted."},
    "msg.no_selection": {ES: "No hay filas seleccionadas.",
                         EN: "No rows are selected."},

    # -- progreso ----------------------------------------------------------
    "prog.reading": {ES: "Leyendo {file} ...", EN: "Reading {file} ..."},
    "prog.reading_title": {ES: "Cargando submissions", EN: "Loading submissions"},
    "prog.writing_title": {ES: "Cargando al Excel destino",
                           EN: "Loading into target Excel"},
    "prog.writing": {ES: "Escribiendo filas en '{sheet}'...",
                     EN: "Writing rows into '{sheet}'..."},
    "prog.tags_title": {ES: "Consolidando archivos de tags",
                        EN: "Consolidating tag files"},
    "prog.tags": {ES: "Leyendo archivos semanales...",
                  EN: "Reading weekly files..."},
    "prog.export_title": {ES: "Exportando", EN: "Exporting"},
    "prog.export": {ES: "Generando el Excel...", EN: "Building the Excel..."},

    # -- dialogo de flota por mes ------------------------------------------
    "fleet.title": {ES: "Flota total por mes", EN: "Total fleet per month"},
    "fleet.hint": {
        ES: "Es el denominador de '% Inspeccion por mes': cuantos equipos "
            "tenia la flota ese mes. Se guarda con las preferencias del "
            "software.",
        EN: "This is the denominator of '% Inspection per month': how many "
            "units the fleet had that month. It is saved with the software "
            "preferences."},
    "fleet.btn_ok": {ES: "Guardar", EN: "Save"},
    "fleet.btn_cancel": {ES: "Cancelar", EN: "Cancel"},
    "fleet.btn_fill": {ES: "Completar hacia abajo", EN: "Fill down"},
}
