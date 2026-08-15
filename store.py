# -*- coding: utf-8 -*-
"""
Base local del software: SQLite.

Hasta ahora la herramienta era un puente de un solo sentido — leia el export
del formulario y lo volcaba en el Excel maestro sin quedarse con nada. Aqui
empieza a tener memoria: guarda las inspecciones cargadas y los movimientos de
tag de los archivos semanales, para poder mostrarlos, graficarlos y exportarlos
sin depender de que el maestro este a mano.

Dos tablas de datos:

  `inspections`   una fila por inspeccion, con las mismas columnas que
                  'Full List 2024-2025' (B, C, E..S del maestro).
  `movements`     una fila por movimiento de tag de 'Tag Installed Per Week'.

Ambas se protegen de la doble carga con una clave `row_key` UNIQUE: volver a
importar el mismo export o una carpeta con semanas solapadas no duplica filas.
Es la diferencia importante contra la carga al Excel maestro, que si agrega
todo lo que se le marque — ahi el usuario decide, aqui la base decide.

SQLite se eligio sobre un JSON o un CSV porque el volumen real ya es de miles
de filas (el maestro tiene 4.400) y porque el filtrado por mes, equipo y
propietario del tablero se resuelve con indices en vez de recorrer todo en
memoria.
"""
from __future__ import annotations

import datetime
import hashlib
import os
import sqlite3

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "fleet_data.sqlite3")

# Campo de la base -> columna de 'Full List 2024-2025'. El orden es el de la
# hoja, asi que sirve tanto para escribir el Excel como para armar la tabla de
# la ventana.
INSPECTION_FIELDS = (
    ("date", "B"),
    ("vehicle_id", "C"),
    ("fms_id", "E"),
    ("system_fitted", "F"),
    ("equipment_hours", "G"),
    ("fms_hours", "H"),
    ("status", "I"),
    ("inlets", "J"),
    ("addl_inlets_locked", "K"),
    ("drain_valves_locked", "L"),
    ("fast_fill_leaking", "M"),
    ("smu_tags", "N"),
    ("equipment_type", "O"),
    ("remarks", "P"),
    ("inspectors", "Q"),
    ("owner", "R"),
    ("remedial", "S"),
)

MOVEMENT_FIELDS = (
    "move_type", "date", "equipment_id", "tag", "device_type",
    "cost_center", "department", "product", "changed_by",
    "type_inferred", "source_file", "sheet", "note",
)

_SCHEMA = """
CREATE TABLE IF NOT EXISTS inspections (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    row_key             TEXT NOT NULL UNIQUE,
    date                TEXT,
    vehicle_id          TEXT,
    fms_id              TEXT,
    system_fitted       TEXT,
    equipment_hours     TEXT,
    fms_hours           TEXT,
    status              TEXT,
    inlets              TEXT,
    addl_inlets_locked  TEXT,
    drain_valves_locked TEXT,
    fast_fill_leaking   TEXT,
    smu_tags            TEXT,
    equipment_type      TEXT,
    remarks             TEXT,
    inspectors          TEXT,
    owner               TEXT,
    remedial            TEXT,
    submission_code     TEXT,
    source_file         TEXT,
    imported_at         TEXT
);
CREATE INDEX IF NOT EXISTS ix_inspections_date    ON inspections(date);
CREATE INDEX IF NOT EXISTS ix_inspections_vehicle ON inspections(vehicle_id);

CREATE TABLE IF NOT EXISTS movements (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    row_key       TEXT NOT NULL UNIQUE,
    move_type     TEXT,
    date          TEXT,
    equipment_id  TEXT,
    tag           TEXT,
    device_type   TEXT,
    cost_center   TEXT,
    department    TEXT,
    product       TEXT,
    changed_by    TEXT,
    type_inferred INTEGER DEFAULT 0,
    source_file   TEXT,
    sheet         TEXT,
    note          TEXT,
    imported_at   TEXT
);
CREATE INDEX IF NOT EXISTS ix_movements_date   ON movements(date);
CREATE INDEX IF NOT EXISTS ix_movements_equip  ON movements(equipment_id);
"""

_db_path = DB_FILE


def set_path(path: str) -> None:
    """Cambia el archivo de base (lo usan las pruebas con un tmp_path)."""
    global _db_path
    _db_path = path


def path() -> str:
    return _db_path


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(_db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init() -> None:
    conn = connect()
    try:
        conn.executescript(_SCHEMA)
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Normalizacion de valores
# ---------------------------------------------------------------------------
def iso_date(value) -> str:
    """Fecha -> 'AAAA-MM-DD'. Devuelve '' si no se puede interpretar.

    Se guarda en ISO y no en dd/mm/aaaa para que el orden alfabetico de SQLite
    sea el orden cronologico y los filtros por mes sean un LIKE 'AAAA-MM%'.
    """
    if value in (None, ""):
        return ""
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    text = str(value).strip()
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S",
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return datetime.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def parse_date(value):
    """'AAAA-MM-DD' -> date, o None."""
    iso = iso_date(value)
    if not iso:
        return None
    try:
        return datetime.date.fromisoformat(iso)
    except ValueError:
        return None


def _text(value) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return iso_date(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def number(value):
    """Devuelve un numero si el texto lo es, si no el texto tal cual.

    Las horas y los contadores llegan como texto de la base; al Excel deben ir
    como numero para que las formulas y los formatos funcionen.
    """
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().replace(",", ".")
    try:
        num = float(text)
    except ValueError:
        return value
    return int(num) if num.is_integer() else num


def _hash(*parts) -> str:
    raw = "|".join(_text(p).upper() for p in parts)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# Inspecciones
# ---------------------------------------------------------------------------
def inspection_from_row(row: dict, submission_code: str = "",
                        source_file: str = "") -> dict:
    """Convierte una fila {columna: valor} de `mapping` en un registro."""
    record = {name: _text(row.get(col)) for name, col in INSPECTION_FIELDS}
    record["date"] = iso_date(row.get("B"))
    record["submission_code"] = _text(submission_code)
    record["source_file"] = os.path.basename(source_file) if source_file else ""
    return record


def inspection_base_key(record: dict) -> str:
    """Huella de CONTENIDO de una inspeccion.

    Deliberadamente no se usa el codigo de submission como clave, aunque sea
    unico: la misma inspeccion puede llegar dos veces por caminos distintos —
    del export del formulario y, mas tarde, de la hoja 'Full List 2024-2025'
    del maestro donde ya se habia volcado. Con el codigo como clave entrarian
    como dos filas; con la huella del contenido, la segunda se reconoce y se
    omite. El codigo se guarda igual, como dato de la fila.
    """
    return "H:" + _hash(record.get("date"), record.get("vehicle_id"),
                        record.get("fms_id"), record.get("smu_tags"),
                        record.get("inspectors"),
                        record.get("equipment_hours"))


def inspection_key(record: dict, occurrence: int = 0) -> str:
    """Clave unica de la fila: huella + numero de repeticion.

    Dos inspecciones REALES pueden ser identicas campo por campo (el mismo
    camion revisado dos veces el mismo dia por el mismo inspector, con el
    contador sin moverse). Por eso la repeticion se numera en vez de
    descartarse: la dinamica cuenta inspecciones, y colapsarlas cambiaria el
    numero que se le reporta al cliente.
    """
    base = inspection_base_key(record)
    return base if occurrence <= 0 else "%s#%d" % (base, occurrence)


def _existing_occurrences(conn) -> dict:
    """{huella: cuantas filas ya hay} para no repetir al reimportar."""
    counts = {}
    for (key,) in conn.execute("SELECT row_key FROM inspections"):
        base = key.split("#", 1)[0]
        counts[base] = counts.get(base, 0) + 1
    return counts


def add_inspections(records: list, progress_cb=None) -> dict:
    """Inserta registros nuevos. Los repetidos se omiten en silencio."""
    records = [r for r in (records or []) if r]
    if not records:
        return {"added": 0, "skipped": 0}

    names = [name for name, _ in INSPECTION_FIELDS] + [
        "submission_code", "source_file"]
    sql = ("INSERT OR IGNORE INTO inspections (row_key, %s, imported_at) "
           "VALUES (%s)" % (", ".join(names),
                            ", ".join("?" * (len(names) + 2))))
    now = datetime.datetime.now().isoformat(timespec="seconds")

    conn = connect()
    try:
        existing = _existing_occurrences(conn)
        seen = {}
        added = 0
        total = len(records)
        for i, record in enumerate(records):
            base = inspection_base_key(record)
            # La n-esima copia de esta huella dentro del lote se compara contra
            # la n-esima ya almacenada: si el archivo trae tres filas iguales y
            # la base ya tiene tres, no se agrega ninguna.
            seen[base] = seen.get(base, 0) + 1
            occurrence = seen[base] - 1
            if occurrence < existing.get(base, 0):
                continue
            values = [inspection_key(record, occurrence)]
            values += [_text(record.get(n)) for n in names]
            values.append(now)
            added += conn.execute(sql, values).rowcount
            if progress_cb and (i % 50 == 0 or i == total - 1):
                progress_cb(i + 1, total, "")
        conn.commit()
    finally:
        conn.close()
    return {"added": added, "skipped": len(records) - added}


def inspections(year=None, month=None, owner=None, search=None,
                limit=None) -> list:
    """Inspecciones almacenadas, de la mas reciente a la mas antigua.

    Las filas sin fecha van al final: son cargas con la fecha ilegible y
    conviene que salten a la vista al revisar, no que se pierdan arriba.
    """
    where, params = [], []
    if year:
        prefix = "%04d-" % int(year)
        if month:
            prefix = "%04d-%02d" % (int(year), int(month))
        where.append("date LIKE ?")
        params.append(prefix + "%")
    elif month:
        where.append("substr(date, 6, 2) = ?")
        params.append("%02d" % int(month))
    if owner:
        where.append("owner = ?")
        params.append(owner)
    if search:
        like = "%" + str(search).strip() + "%"
        where.append("(vehicle_id LIKE ? OR fms_id LIKE ? OR inspectors LIKE ? "
                     "OR equipment_type LIKE ? OR remarks LIKE ?)")
        params += [like] * 5

    sql = "SELECT * FROM inspections"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY (date = '') ASC, date DESC, id DESC"
    if limit:
        sql += " LIMIT %d" % int(limit)

    conn = connect()
    try:
        return [dict(r) for r in conn.execute(sql, params)]
    finally:
        conn.close()


def count_inspections() -> int:
    conn = connect()
    try:
        return conn.execute("SELECT COUNT(*) FROM inspections").fetchone()[0]
    finally:
        conn.close()


def delete_inspections(ids: list) -> int:
    ids = [int(i) for i in (ids or [])]
    if not ids:
        return 0
    conn = connect()
    try:
        cur = conn.execute(
            "DELETE FROM inspections WHERE id IN (%s)" %
            ",".join("?" * len(ids)), ids)
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def inspection_years() -> list:
    conn = connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT substr(date, 1, 4) AS y FROM inspections "
            "WHERE date <> '' ORDER BY y DESC")
        return [r["y"] for r in rows]
    finally:
        conn.close()


def inspection_owners() -> list:
    conn = connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT owner FROM inspections WHERE owner <> '' "
            "ORDER BY owner")
        return [r["owner"] for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Movimientos de tag
# ---------------------------------------------------------------------------
def movement_key(record: dict) -> str:
    """Misma combinacion que usa el consolidado del cliente: tipo, fecha, ID y
    tag. Es lo que hace que dos archivos semanales con dias solapados no
    dupliquen el movimiento compartido."""
    return _hash(record.get("move_type"), record.get("date"),
                 record.get("equipment_id"), record.get("tag"))


def add_movements(records: list, progress_cb=None) -> dict:
    records = [r for r in (records or []) if r]
    if not records:
        return {"added": 0, "skipped": 0}

    sql = ("INSERT OR IGNORE INTO movements (row_key, %s, imported_at) "
           "VALUES (%s)" % (", ".join(MOVEMENT_FIELDS),
                            ", ".join("?" * (len(MOVEMENT_FIELDS) + 2))))
    now = datetime.datetime.now().isoformat(timespec="seconds")

    conn = connect()
    try:
        added = 0
        total = len(records)
        for i, record in enumerate(records):
            values = [movement_key(record)]
            for name in MOVEMENT_FIELDS:
                value = record.get(name)
                if name == "type_inferred":
                    values.append(1 if value else 0)
                elif name == "date":
                    values.append(iso_date(value))
                else:
                    values.append(_text(value))
            values.append(now)
            added += conn.execute(sql, values).rowcount
            if progress_cb and (i % 100 == 0 or i == total - 1):
                progress_cb(i + 1, total, "")
        conn.commit()
    finally:
        conn.close()
    return {"added": added, "skipped": len(records) - added}


def movements(move_type=None, device=None, department=None, year=None,
              search=None) -> list:
    where, params = [], []
    if move_type:
        where.append("move_type = ?")
        params.append(move_type)
    if device:
        where.append("device_type = ?")
        params.append(device)
    if department:
        where.append("department = ?")
        params.append(department)
    if year:
        where.append("date LIKE ?")
        params.append("%04d-%%" % int(year))
    if search:
        like = "%" + str(search).strip() + "%"
        where.append("(equipment_id LIKE ? OR tag LIKE ? OR cost_center LIKE ? "
                     "OR source_file LIKE ?)")
        params += [like] * 4

    sql = "SELECT * FROM movements"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY (date = '') ASC, date DESC, id DESC"

    conn = connect()
    try:
        return [dict(r) for r in conn.execute(sql, params)]
    finally:
        conn.close()


def count_movements() -> int:
    conn = connect()
    try:
        return conn.execute("SELECT COUNT(*) FROM movements").fetchone()[0]
    finally:
        conn.close()


def clear_movements() -> int:
    conn = connect()
    try:
        cur = conn.execute("DELETE FROM movements")
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def movement_values(column: str) -> list:
    """Valores distintos de una columna, para llenar los filtros."""
    if column not in MOVEMENT_FIELDS:
        return []
    conn = connect()
    try:
        rows = conn.execute(
            "SELECT DISTINCT %s AS v FROM movements WHERE %s <> '' "
            "ORDER BY v" % (column, column))
        return [r["v"] for r in rows]
    finally:
        conn.close()


def movement_source_files() -> int:
    conn = connect()
    try:
        return conn.execute(
            "SELECT COUNT(DISTINCT source_file) FROM movements "
            "WHERE source_file <> ''").fetchone()[0]
    finally:
        conn.close()
