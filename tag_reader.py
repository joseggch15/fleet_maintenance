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
NOTE_FIXED_YEAR = "FIXED_YEAR"
NOTE_FIXED_SWAP = "FIXED_SWAP"
NOTE_SUSPECT_DATE = "SUSPECT_DATE"

# ---------------------------------------------------------------------------
# Ventanas de coherencia entre la fecha de una fila y el periodo del archivo
# ---------------------------------------------------------------------------
#
# El nombre del archivo declara la semana que cubre, y eso es lo unico que
# permite saber si una fecha esta mal: sin esa referencia, '05/12/2027' es una
# fecha perfectamente valida y no hay con que discutirla.
#
# Los numeros salen de medir las 457 filas con fecha de la carpeta del cliente
# contra la fecha de cierre que declara cada nombre de archivo:
#
#   95,6%  caen entre 10 dias antes y 3 despues  (la semana del archivo)
#    1,8%  caen hasta un mes antes               (cargas tardias legitimas)
#    2,6%  caen a mas de un ano de distancia     (las 12 filas mal escritas)
#
# No hay nada entre "un mes antes" y "un ano de distancia", asi que el corte es
# holgado y no toca ninguna carga tardia real: se considera coherente todo lo
# que cae entre 45 dias antes y 3 despues, y solo se intenta reparar lo que
# queda fuera. La ventana de ACEPTACION de una correccion es mas estrecha (31
# dias antes): para cambiar un dato hay que estar mas seguro que para dejarlo.
_CONSISTENT_BEFORE = 45
_CONSISTENT_AFTER = 3
_REPAIR_BEFORE = 31
_REPAIR_AFTER = 3

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

# Unos pocos archivos no traen la fecha en digitos sino el mes en ingles
# ('Tag Installed April_2025.xlsx'). Se aceptan porque dan la misma referencia
# —el cierre del periodo— que los demas.
_MONTH_NAMES = ("january", "february", "march", "april", "may", "june",
                "july", "august", "september", "october", "november",
                "december")
_MONTH_TOKEN = re.compile(r"([A-Za-z]+)[ _-]+(\d{4})")


def file_period_days(filename: str) -> list:
    """Fechas que declara el NOMBRE del archivo, ordenadas.

    Es la referencia contra la que se juzga si la fecha de una fila es
    coherente. Los nombres que no traen 8 digitos ni un mes reconocible (hay
    al menos uno con un digito de menos) devuelven lista vacia en vez de
    inventar una fecha: sin referencia no se corrige nada.
    """
    stem = os.path.splitext(os.path.basename(filename or ""))[0]
    days = []
    for token in _DATE_TOKEN.findall(stem):
        try:
            days.append(datetime.datetime.strptime(token, "%d%m%Y").date())
        except ValueError:
            continue
    if days:
        return sorted(days)

    match = _MONTH_TOKEN.search(stem)
    if match and match.group(1).lower() in _MONTH_NAMES:
        month = _MONTH_NAMES.index(match.group(1).lower()) + 1
        year = int(match.group(2))
        # El cierre del mes: el archivo mensual cubre hasta ese dia.
        last = datetime.date(year + month // 12, month % 12 + 1, 1) - \
            datetime.timedelta(days=1)
        return [datetime.date(year, month, 1), last]
    return []


def file_reference(filename: str):
    """Fecha de cierre del periodo del archivo, o None si el nombre no la trae."""
    days = file_period_days(filename)
    return days[-1] if days else None


def file_period(filename: str) -> str:
    """Periodo que cubre el archivo, en texto para el consolidado.

    'Inventory Tag Installed 05082026-12082026.xlsx' -> '05/08/2026 - 12/08/2026'
    'Inventory Tag Installed 01102025.xlsx'          -> '01/10/2025'
    """
    days = file_period_days(filename)
    if not days:
        return ""
    if len(days) == 1:
        return days[0].strftime("%d/%m/%Y")
    return "%s - %s" % (days[0].strftime("%d/%m/%Y"),
                        days[-1].strftime("%d/%m/%Y"))


# ---------------------------------------------------------------------------
# Reparacion de fechas mal escritas
# ---------------------------------------------------------------------------
def _fits(day, reference, before: int, after: int) -> bool:
    delta = (day - reference).days
    return -before <= delta <= after


def repair_date(day, reference, repair: bool = True) -> tuple:
    """(fecha_final, nota). Corrige la fecha si el periodo del archivo lo prueba.

    Solo se corrige cuando el resultado es UNICO. Se prueban dos errores de
    tecleo, los unicos que aparecen en los archivos del cliente y los unicos
    que el periodo del archivo puede confirmar por si solo:

      ano equivocado    '05/12/2027' en el archivo de la semana del 10/12/2025.
                        El dia y el mes ya son correctos; cambiando solo el ano
                        la fila cae dentro de la semana. Aparece en las dos
                        direcciones: tambien hay filas de 2025 en archivos de
                        2026.
      dia y mes al reves '05/12' escrito como '12/05'. No se ha visto en estos
                        archivos, pero es el error clasico entre el formato
                        europeo y el americano y el periodo lo delata igual.

    Si ninguna candidata cae en la ventana, o si caen DOS distintas, la fecha
    se deja como esta y se marca como sospechosa. Una correccion ambigua es
    peor que ninguna: quien lea el consolidado puede revisar una fila marcada,
    pero no puede adivinar que un dato fue cambiado mal.

    Con la ventana actual (34 dias) dos candidatas distintas no llegan a caber
    —dos anos estan a 365 dias— asi que en la practica la correccion es unica o
    no hay. La guarda queda igual: si alguien ensancha la ventana, el codigo
    tiene que seguir negandose a elegir.
    """
    if day is None or reference is None:
        return day, ""
    if _fits(day, reference, _CONSISTENT_BEFORE, _CONSISTENT_AFTER):
        return day, ""

    original = day.strftime("%d/%m/%Y")
    if not repair:
        return day, "%s:%s" % (NOTE_SUSPECT_DATE, original)

    candidates = []
    for year in (reference.year - 1, reference.year, reference.year + 1):
        try:
            fixed = day.replace(year=year)
        except ValueError:
            continue            # 29 de febrero en un ano no bisiesto
        if fixed != day and _fits(fixed, reference, _REPAIR_BEFORE,
                                  _REPAIR_AFTER):
            candidates.append((fixed, NOTE_FIXED_YEAR))
    try:
        swapped = day.replace(day=day.month, month=day.day)
    except ValueError:
        swapped = None
    if swapped is not None and swapped != day and \
            _fits(swapped, reference, _REPAIR_BEFORE, _REPAIR_AFTER):
        candidates.append((swapped, NOTE_FIXED_SWAP))

    if len({fixed for fixed, _code in candidates}) == 1:
        fixed, code = candidates[0]
        return fixed, "%s:%s" % (code, original)
    return day, "%s:%s" % (NOTE_SUSPECT_DATE, original)


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


def _read_sheet(ws, source_file: str, repair: bool = True) -> list:
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        return []
    columns = _map_headers(header)
    if columns is None:
        return []

    has_type_column = "move_type" in columns
    reference = file_reference(source_file)
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
            # Fecha imposible ('31/06/2026', '19/19/2025'): no se repara. El
            # periodo del archivo puede probar que un ANO esta mal, pero no
            # puede decir que mes quiso escribir alguien que tecleo '19'.
            notes.append("%s:%s" % (NOTE_BAD_DATE, _clean(raw_date)))
        elif day is not None:
            day, note = repair_date(day, reference, repair)
            if note:
                notes.append(note)
            if day > datetime.date.today():
                notes.append("%s:%s" % (NOTE_FUTURE_DATE,
                                        day.strftime("%d/%m/%Y")))

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


def read_file(path: str, repair: bool = True) -> list:
    """Movimientos de un archivo semanal (todas sus hojas de datos)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        records = []
        for ws in wb.worksheets:
            records.extend(_read_sheet(ws, path, repair))
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


def count_notes(records: list, code: str) -> int:
    """Cuantos registros llevan una observacion de ese tipo."""
    return sum(1 for r in records or [] if code in (r.get("note") or ""))


def read_paths(paths: list, progress_cb=None, repair: bool = True) -> dict:
    """Lee varios archivos y devuelve {records, files, errors, repaired, suspect}.

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
            found = read_file(path, repair)
        except Exception as exc:  # noqa: BLE001 - se reporta, no se propaga
            errors.append((os.path.basename(path), str(exc)))
            continue
        if found:
            records.extend(found)
            files.append(path)
    return {"records": records, "files": files, "errors": errors,
            "repaired": count_notes(records, NOTE_FIXED_YEAR) +
            count_notes(records, NOTE_FIXED_SWAP),
            "suspect": count_notes(records, NOTE_SUSPECT_DATE) +
            count_notes(records, NOTE_BAD_DATE)}


def read_folder(folder: str, recursive: bool = True, progress_cb=None,
                repair: bool = True) -> dict:
    return read_paths(find_files(folder, recursive), progress_cb, repair)
