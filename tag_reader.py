# -*- coding: utf-8 -*-
"""
Lectura de los archivos semanales 'Inventory Tag Installed *.xlsx'
(carpeta 'Tag Installed Per Week').

Son archivos hechos a mano semana a semana, asi que el formato cambio con el
tiempo. Este modulo los unifica sin tocar el origen:

  encabezados     se buscan por nombre con alias ('CC' -> Cost Center,
                  'Dept' -> Department), nunca por posicion;
  columna TYPE    los archivos anteriores a junio 2025 no la traen; esos
                  movimientos se toman como 'NEW INSTALLATION' y quedan
                  marcados con `type_inferred`, para que en el tablero se
                  distinga un dato leido de uno asumido;
  dispositivo     se deduce del tag: con dos puntos (formato MAC) es un SMU,
                  si no es un TAG. Es la misma regla del consolidado del
                  cliente;
  fechas          se aceptan las fechas de Excel y el texto dd/mm/aaaa. Las
                  que no se pueden interpretar ('19/19/2025', '31/06/2026')
                  NO se corrigen ni se descartan: la fila entra sin fecha y con
                  una observacion, para que alguien la arregle en el origen.

Las observaciones se guardan como CODIGO[:detalle] en ingles canonico y se
traducen al mostrarlas; asi el lector no depende del idioma de la ventana.
"""
from __future__ import annotations

import datetime
import os
import re

import openpyxl

# Tipos de movimiento canonicos, tal como se escriben en los archivos.
MOVE_NEW = "NEW INSTALLATION"
MOVE_REPLACEMENT = "REPLACEMENT"
MOVE_REMOVAL = "REMOVAL"
MOVE_UPDATED = "TAG UPDATED"
MOVE_TYPES = (MOVE_NEW, MOVE_REPLACEMENT, MOVE_REMOVAL, MOVE_UPDATED)

# Los que suman inventario instalado. 'REMOVAL' resta, por eso queda fuera:
# es la misma condicion (<>"REMOVAL") del resumen del cliente.
INSTALL_TYPES = (MOVE_NEW, MOVE_REPLACEMENT, MOVE_UPDATED)

DEVICE_SMU = "SMU"
DEVICE_TAG = "TAG"

# Observaciones (codigo canonico -> se traduce al mostrar).
NOTE_BAD_DATE = "BAD_DATE"
NOTE_TYPE_UNKNOWN = "TYPE_UNKNOWN"
NOTE_NO_TYPE_COLUMN = "TYPE_ASSUMED"
NOTE_FUTURE_DATE = "FUTURE_DATE"

_ALIASES = {
    "move_type":    ("TYPE", "TIPO", "MOVEMENT", "MOVIMIENTO"),
    "date":         ("DATE", "FECHA"),
    "equipment_id": ("ID", "EQUIPMENT ID", "FIELD ID", "EQUIPO"),
    "tag":          ("TAG", "TAG ID", "CODE S/N"),
    "cost_center":  ("COST CENTER", "CC", "COSTCENTER", "CENTRO DE COSTO"),
    "department":   ("DEPARTMENT", "DEPT", "DEPARTAMENTO"),
    "product":      ("PRODUCT", "PRODUCTO"),
    "changed_by":   ("CHANGED BY", "MODIFICADO POR"),
}

# Encabezados que solo aparecen en un consolidado YA generado (por este
# software o por la version anterior hecha a mano). Si una hoja los trae, no es
# un archivo semanal de origen: se ignora para no reimportar la salida.
_OUTPUT_MARKERS = ("ARCHIVO ORIGEN", "SOURCE FILE", "FILE PERIOD",
                   "TYPE INFERIDO", "INFERRED TYPE", "SEMANA (LUNES)",
                   "WEEK (MONDAY)")

_TYPE_LOOKUP = {
    "NEW INSTALLATION": MOVE_NEW,
    "NEW INSTALATION": MOVE_NEW,
    "INSTALLATION": MOVE_NEW,
    "NEW": MOVE_NEW,
    "REPLACEMENT": MOVE_REPLACEMENT,
    "REPLACE": MOVE_REPLACEMENT,
    "REEMPLAZO": MOVE_REPLACEMENT,
    "REMOVAL": MOVE_REMOVAL,
    "REMOVED": MOVE_REMOVAL,
    "RETIRO": MOVE_REMOVAL,
    "TAG UPDATED": MOVE_UPDATED,
    "UPDATED": MOVE_UPDATED,
    "TAG UPDATE": MOVE_UPDATED,
}


# ---------------------------------------------------------------------------
# Normalizacion de valores
# ---------------------------------------------------------------------------
def _clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_type(value):
    """Texto de la columna TYPE -> tipo canonico, o None si no se reconoce."""
    text = _clean(value).upper()
    if not text:
        return None
    return _TYPE_LOOKUP.get(text)


def device_type(tag) -> str:
    """SMU si el tag viene en formato MAC (con ':'), TAG en cualquier otro caso.

    Un mismo equipo puede llevar los dos: el consolidado del cliente los separa
    exactamente asi.
    """
    return DEVICE_SMU if ":" in _clean(tag) else DEVICE_TAG


def parse_date(value):
    """Fecha de Excel o texto dd/mm/aaaa -> date. None si no se interpreta."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = _clean(value)
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y",
                "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def week_monday(value):
    """Lunes de la semana de esa fecha (el consolidado agrupa por semana)."""
    day = value if isinstance(value, datetime.date) else parse_date(value)
    if day is None:
        return None
    if isinstance(day, datetime.datetime):
        day = day.date()
    return day - datetime.timedelta(days=day.weekday())


_DATE_TOKEN = re.compile(r"(\d{8})")


def file_period(filename: str) -> str:
    """Periodo que cubre el archivo, leido de su nombre.

    'Inventory Tag Installed 05082026-12082026.xlsx' -> '05/08/2026 - 12/08/2026'
    'Inventory Tag Installed 01102025.xlsx'          -> '01/10/2025'

    Es informativo: se escribe en el consolidado para saber de que semana salio
    cada fila. Los nombres que no traen 8 digitos (hay al menos uno con un
    digito de menos) devuelven cadena vacia en vez de inventar una fecha.
    """
    stem = os.path.splitext(os.path.basename(filename or ""))[0]
    days = []
    for token in _DATE_TOKEN.findall(stem):
        try:
            days.append(datetime.datetime.strptime(token, "%d%m%Y").date())
        except ValueError:
            continue
    if not days:
        return ""
    if len(days) == 1:
        return days[0].strftime("%d/%m/%Y")
    return "%s - %s" % (min(days).strftime("%d/%m/%Y"),
                        max(days).strftime("%d/%m/%Y"))


# ---------------------------------------------------------------------------
# Lectura de una hoja
# ---------------------------------------------------------------------------
def _map_headers(row):
    """Fila de encabezados -> {campo: indice}. None si la hoja no es de datos."""
    labels = [_clean(v).upper() for v in row]
    if any(marker in labels for marker in _OUTPUT_MARKERS):
        return None
    found = {}
    for field, aliases in _ALIASES.items():
        for idx, label in enumerate(labels):
            if label in aliases and field not in found:
                found[field] = idx
                break
    # Una hoja de movimientos tiene, como minimo, equipo y tag. Sin eso es una
    # hoja de resumen ('KPI'/'Value') o una hoja suelta y se ignora.
    if "equipment_id" not in found or "tag" not in found:
        return None
    return found


def _read_sheet(ws, source_file: str) -> list:
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        return []
    columns = _map_headers(header)
    if columns is None:
        return []

    has_type_column = "move_type" in columns
    records = []
    for raw in rows:
        if raw is None or all(v in (None, "") for v in raw):
            continue

        def cell(field):
            idx = columns.get(field)
            return raw[idx] if idx is not None and idx < len(raw) else None

        equipment_id = _clean(cell("equipment_id"))
        tag = _clean(cell("tag"))
        if not equipment_id and not tag:
            continue

        notes = []
        raw_type = cell("move_type") if has_type_column else None
        move_type = normalize_type(raw_type)
        inferred = False
        if move_type is None:
            move_type = MOVE_NEW
            inferred = True
            text = _clean(raw_type)
            # Hay al menos una fila con el ID escrito en la columna TYPE. Se
            # deja constancia del valor original en vez de perderlo.
            notes.append("%s:%s" % (NOTE_TYPE_UNKNOWN, text) if text
                         else NOTE_NO_TYPE_COLUMN)

        raw_date = cell("date")
        day = parse_date(raw_date)
        if day is None and _clean(raw_date):
            notes.append("%s:%s" % (NOTE_BAD_DATE, _clean(raw_date)))
        elif day is not None and day > datetime.date.today():
            # Los archivos de diciembre 2025 traen unas filas fechadas en 2027.
            # La fecha se conserva tal cual —corregirla seria inventar un dato—
            # pero queda marcada para que salte a la vista en el consolidado.
            notes.append("%s:%s" % (NOTE_FUTURE_DATE, day.strftime("%d/%m/%Y")))

        records.append({
            "move_type": move_type,
            "date": day.isoformat() if day else "",
            "equipment_id": equipment_id,
            "tag": tag,
            "device_type": device_type(tag),
            "cost_center": _clean(cell("cost_center")),
            # El departamento se normaliza a mayusculas: el mismo se escribe
            # 'MINE_OPS' y 'Mine_Ops' segun quien llene la semana, y en el
            # tablero saldrian como dos departamentos distintos.
            "department": _clean(cell("department")).upper(),
            "product": _clean(cell("product")),
            "changed_by": _clean(cell("changed_by")),
            "type_inferred": inferred,
            "source_file": os.path.basename(source_file),
            "sheet": ws.title,
            "note": " ".join(notes),
        })
    return records


def read_file(path: str) -> list:
    """Movimientos de un archivo semanal (todas sus hojas de datos)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        records = []
        for ws in wb.worksheets:
            records.extend(_read_sheet(ws, path))
        return records
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Lectura de una carpeta
# ---------------------------------------------------------------------------
def find_files(folder: str, recursive: bool = True) -> list:
    """Archivos .xlsx de la carpeta, ordenados por nombre.

    Se saltan los temporales de Excel ('~$...') y los respaldos que deja el
    propio software, para no consolidar dos veces la misma semana.
    """
    found = []
    for dirpath, _dirs, files in os.walk(folder):
        for name in sorted(files):
            if not name.lower().endswith(".xlsx"):
                continue
            if name.startswith("~$") or ".backup_" in name:
                continue
            found.append(os.path.join(dirpath, name))
        if not recursive:
            break
    return found


def read_paths(paths: list, progress_cb=None) -> dict:
    """Lee varios archivos y devuelve {records, files, errors}.

    Un archivo ilegible (corrupto, o abierto por Excel con bloqueo) no aborta
    la carga: se anota en `errors` y se sigue con el resto. Consolidar 84 de 85
    semanas y saber cual falto es mas util que no consolidar nada.
    """
    paths = list(paths or [])
    records, files, errors = [], [], []
    total = len(paths)
    for i, path in enumerate(paths):
        if progress_cb:
            try:
                progress_cb(i + 1, total, os.path.basename(path))
            except Exception:
                pass
        try:
            found = read_file(path)
        except Exception as exc:  # noqa: BLE001 - se reporta, no se propaga
            errors.append((os.path.basename(path), str(exc)))
            continue
        if found:
            records.extend(found)
            files.append(path)
    return {"records": records, "files": files, "errors": errors}


def read_folder(folder: str, recursive: bool = True, progress_cb=None) -> dict:
    return read_paths(find_files(folder, recursive), progress_cb)
