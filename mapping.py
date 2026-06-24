# -*- coding: utf-8 -*-
"""
Mapeo entre el export del formulario (merian_ops - Form - Fleet maintenance -
Submissions) y la hoja destino 'Full List 2024-2025' del Excel maestro
'Fleet Tag Inventory and Maintenance.xlsx'.

Cada submission del formulario se convierte en UNA fila nueva de la tabla
`Table3` (columnas A..S) de esa hoja.

Las columnas A ('Date') y D ('Verified') de la hoja destino son FORMULAS
(referencias estructuradas a `Table3`) que se recalculan solas; la herramienta
no las llena con datos, las reproduce tal cual (ver `excel_writer.py`).

El resto de columnas (B, C, E..S) se llenan desde el formulario o con valores
por defecto cuando el formulario no tiene un campo equivalente.
"""
from __future__ import annotations

import datetime


# ---------------------------------------------------------------------------
# Encabezados EXACTOS del formulario (fila 4 de la hoja 'Form Submissions').
# El lector busca las columnas por estos nombres, no por posicion, asi que
# si el software de origen reordena columnas el mapeo sigue funcionando.
# ---------------------------------------------------------------------------
H_REVISION = "Date And Time Of Revision"
H_SUBMITTED = "Submitted At"
H_VEHICLE = "Vehicle Id"
H_FMS = "Fms Id"
H_TAGTYPE = "Tag Type"
H_KM = "Machine Kilometer Reading"
H_HOUR = "Machine Hourmeter Reading"
H_EQUIP = "Equipment Type?"
H_INLETS = "Number Of Inlets?"
H_TAGS = "Number Of Tags?"
H_ADDLOCK = "Additional Inlets Locked?"
H_DRAIN = "Drain Valves Locked"
H_COMPANY = "Company Name?"
H_INSPECTOR = "Name Of Inspector"
H_REMARKS = "Remarks"
H_SIGNATURE = "Signature"
H_CODE = "Submission Code"
H_ID = "Id"


# ---------------------------------------------------------------------------
# Columnas de la hoja destino 'Full List 2024-2025' (tabla Table3).
# ---------------------------------------------------------------------------
TARGET_HEADERS = {
    "A": "Date",
    "B": "Date (mm/dd/yy)",
    "C": "Vehicle ID",
    "D": "Verified",
    "E": "FMS ID",
    "F": "System fitted as per standard (Y/N)",
    "G": "Equipment Hours/ODO",
    "H": "FMS Hours",
    "I": "Status",
    "J": "# INLETS ",
    "K": "Are aditional inlets locked?(Y/N)",
    "L": "Drain valves locked?(Y/N)",
    "M": "Fast fill receiver leaking?(Y/N)",
    "N": "# SMU/TAGS",
    "O": "EQUIPMENT TYPE",
    "P": "REMARKS",
    "Q": "Inspectors",
    "R": "OWNER",
    "S": "REMEDIAL ACTIONS / UPDPATES",
}

# Columnas que la herramienta llena con datos (en orden de la hoja).
# A y D quedan fuera: son formulas reproducidas por el writer.
DATA_COLUMNS = ["B", "C", "E", "F", "G", "H", "I", "J", "K", "L",
                "M", "N", "O", "P", "Q", "R", "S"]
FORMULA_COLUMNS = ["A", "D"]

# Formulas canonicas de respaldo (usadas solo si el writer no puede leerlas
# de la fila plantilla). Usan referencias estructuradas, identicas por fila.
A_FORMULA = '=TEXT(Table3[[#This Row],[Date (mm/dd/yy)]],"[$-en-US]mmm-yy;@")'
D_FORMULA_TEXT = (
    '=IF(OR(ISNUMBER(MATCH(TRUE, EXACT(RIGHT(Table33[Field ID], 3), '
    'RIGHT(Table3[[#This Row],[Vehicle ID]], 3)), 0)), '
    'ISNUMBER(MATCH(TRUE, EXACT(RIGHT(Table4[Field ID], 3), '
    'RIGHT(Table3[[#This Row],[Vehicle ID]], 3)), 0))), "Yes", "No")'
)

# Valores por defecto para columnas del destino sin campo equivalente en el
# formulario. Replican el patron historico de la hoja.
DEFAULT_SYSTEM_FITTED = "Y"     # F  - System fitted as per standard
DEFAULT_STATUS = "VIU OK"       # I  - Status
DEFAULT_FAST_FILL = "N"         # M  - Fast fill receiver leaking?


# ---------------------------------------------------------------------------
# Transformaciones de valores
# ---------------------------------------------------------------------------
def normalize_locked(value):
    """'Yes'/'No'/'Not applicable' -> 'Y'/'N'/'N/A' (formato historico)."""
    if value is None:
        return None
    s = str(value).strip().lower()
    if s in ("yes", "y", "si", "sí"):
        return "Y"
    if s in ("no", "n"):
        return "N"
    if s in ("not applicable", "n/a", "na", "no aplica"):
        return "N/A"
    return value


def tag_kind(tag_type):
    """Tag Type del formulario -> '# SMU/TAGS' del destino.

    'SMU' -> 'SMU';  'Truck Tag' / 'Half-Moon Tag' -> 'TAG'.
    """
    if tag_type is None:
        return None
    s = str(tag_type).strip().lower()
    if not s:
        return None
    if "smu" in s:
        return "SMU"
    return "TAG"


def pick_hours(sub):
    """'Equipment Hours/ODO' (col G): Hourmeter si existe, si no Kilometer."""
    hour = sub.get(H_HOUR)
    if hour not in (None, ""):
        return hour
    return sub.get(H_KM)


def coerce_datetime(raw):
    """Convierte la fecha del formulario (texto o datetime) a datetime."""
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime.datetime):
        return raw
    if isinstance(raw, datetime.date):
        return datetime.datetime(raw.year, raw.month, raw.day)
    s = str(raw).strip()
    formats = (
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y",
    )
    for fmt in formats:
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def revision_datetime(sub):
    """Fecha de inspeccion: 'Date And Time Of Revision'; si falta, 'Submitted At'."""
    dt = coerce_datetime(sub.get(H_REVISION))
    if dt is None:
        dt = coerce_datetime(sub.get(H_SUBMITTED))
    return dt


# ---------------------------------------------------------------------------
# Mapeo principal
# ---------------------------------------------------------------------------
def submission_to_row(sub: dict) -> dict:
    """Convierte una submission (dict por encabezado) en un dict {col: valor}
    para las columnas de datos de 'Full List 2024-2025'.

    Solo incluye columnas de datos (B, C, E..S). A y D las pone el writer.
    """
    dt = revision_datetime(sub)
    date_val = (datetime.datetime(dt.year, dt.month, dt.day)
                if dt is not None else None)
    return {
        "B": date_val,                                  # Date (mm/dd/yy)
        "C": sub.get(H_VEHICLE),                         # Vehicle ID
        "E": sub.get(H_FMS),                             # FMS ID
        "F": DEFAULT_SYSTEM_FITTED,                      # System fitted (Y/N)
        "G": pick_hours(sub),                            # Equipment Hours/ODO
        "H": None,                                       # FMS Hours (sin origen)
        "I": DEFAULT_STATUS,                             # Status
        "J": sub.get(H_INLETS),                          # # INLETS
        "K": normalize_locked(sub.get(H_ADDLOCK)),       # Add. inlets locked?
        "L": normalize_locked(sub.get(H_DRAIN)),         # Drain valves locked?
        "M": DEFAULT_FAST_FILL,                          # Fast fill leaking?
        "N": tag_kind(sub.get(H_TAGTYPE)),               # # SMU/TAGS
        "O": sub.get(H_EQUIP),                           # EQUIPMENT TYPE
        "P": sub.get(H_REMARKS),                         # REMARKS
        "Q": sub.get(H_INSPECTOR),                       # Inspectors
        "R": sub.get(H_COMPANY),                         # OWNER
        "S": None,                                       # REMEDIAL ACTIONS
    }


def submissions_to_rows(subs: list) -> list:
    """Mapea una lista de submissions a filas destino, en el mismo orden."""
    return [submission_to_row(s) for s in subs]
