# -*- coding: utf-8 -*-
"""
Calculos del tablero y del Excel exportado.

Reproduce con Python lo que el maestro resuelve con una tabla dinamica y un
bloque de formulas en 'PIVOT SUMMARY 2025':

  la dinamica       cuenta de inspecciones por equipo (filas) y por mes
                    (columnas);
  Reviewed tags/SMU =COUNTIF(columna_del_mes;">=1") — cuantos EQUIPOS tuvieron
                    al menos una inspeccion ese mes. No es el numero de
                    inspecciones: tres visitas al mismo camion en enero cuentan
                    como un equipo revisado, y esa es justamente la metrica que
                    se le reporta al cliente;
  % Inspection      revisados del mes / flota total de ese mes. La flota total
                    no sale de estos datos (incluye equipos que ese mes no se
                    tocaron), se configura aparte — ver `settings`;
  Total maintained  =COUNTA(etiquetas de fila) — equipos distintos con alguna
                    inspeccion en todo el periodo.

Y el resumen de los archivos semanales de tags, con la misma condicion que usa
el consolidado del cliente: los retiros ('REMOVAL') no cuentan como instalados.

Este modulo no sabe de Qt ni de openpyxl ni del idioma de la ventana: devuelve
estructuras simples que la interfaz y el exportador presentan como quieran.
"""
from __future__ import annotations

import collections
import datetime
from dataclasses import dataclass, field

import tag_reader


# ---------------------------------------------------------------------------
# Utilidades de mes
# ---------------------------------------------------------------------------
def month_key(value) -> str:
    """Fecha o 'AAAA-MM-DD' -> 'AAAA-MM'. Cadena vacia si no hay fecha."""
    if value in (None, ""):
        return ""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return "%04d-%02d" % (value.year, value.month)
    text = str(value).strip()
    return text[:7] if len(text) >= 7 and text[4] == "-" else ""


def month_date(key: str):
    """'AAAA-MM' -> date del dia 1 (para ejes de tiempo y para Excel)."""
    try:
        return datetime.date(int(key[:4]), int(key[5:7]), 1)
    except (ValueError, IndexError):
        return None


def month_range(first: str, last: str) -> list:
    """Todos los meses entre dos claves, incluidos los vacios.

    El tablero necesita el hueco: un mes sin una sola inspeccion es informacion
    (nadie salio a campo), y si se omitiera la columna la grafica mentiria
    haciendo parecer que los meses son consecutivos.
    """
    start, end = month_date(first), month_date(last)
    if start is None or end is None or start > end:
        return []
    months, year, month = [], start.year, start.month
    while (year, month) <= (end.year, end.month):
        months.append("%04d-%02d" % (year, month))
        month += 1
        if month > 12:
            year, month = year + 1, 1
    return months


def available_months(inspections: list, field_name: str = "date") -> list:
    """Meses continuos que cubren los datos, del mas viejo al mas nuevo."""
    keys = {month_key(row.get(field_name)) for row in inspections or []}
    keys.discard("")
    return month_range(min(keys), max(keys)) if keys else []


def last_months(months: list, count: int) -> list:
    """Los ultimos `count` meses de la lista. Con 0 o None devuelve todos."""
    months = list(months or [])
    return months[-count:] if count and count < len(months) else months


def year_months(months: list, year) -> list:
    prefix = "%04d-" % int(year)
    return [m for m in months or [] if m.startswith(prefix)]


def _natural_key(label: str):
    """Orden 'humano' de los IDs de equipo: 3, 10, 127, C-155, LVE0198.

    El Excel ordena los numeros como texto y deja '10' antes que '3'. Aqui los
    numeros van primero y en orden numerico, que es como los lee un operador.
    """
    text = str(label).strip()
    return (0, int(text), "") if text.isdigit() else (1, 0, text.upper())


# ---------------------------------------------------------------------------
# Dinamica de inspecciones
# ---------------------------------------------------------------------------
@dataclass
class Pivot:
    """Cuenta de inspecciones por equipo y por mes."""
    vehicles: list = field(default_factory=list)     # etiquetas de fila
    months: list = field(default_factory=list)       # claves 'AAAA-MM'
    counts: dict = field(default_factory=dict)       # (vehiculo, mes) -> n
    undated: int = 0                                 # filas sin fecha legible

    def count(self, vehicle: str, month: str) -> int:
        return self.counts.get((vehicle, month), 0)

    def row_total(self, vehicle: str) -> int:
        return sum(self.counts.get((vehicle, m), 0) for m in self.months)

    def col_total(self, month: str) -> int:
        return sum(self.counts.get((v, month), 0) for v in self.vehicles)

    def reviewed(self, month: str) -> int:
        """Equipos con al menos una inspeccion en el mes (COUNTIF ">=1")."""
        return sum(1 for v in self.vehicles
                   if self.counts.get((v, month), 0) >= 1)

    @property
    def total_inspections(self) -> int:
        return sum(self.counts.values())

    @property
    def maintained_fleet(self) -> int:
        """Equipos distintos con alguna inspeccion (COUNTA de las filas)."""
        return len(self.vehicles)


def build_pivot(inspections: list, window: list = None) -> Pivot:
    """Arma la dinamica desde las filas almacenadas.

    Los equipos se agrupan sin distinguir mayusculas ('to22' y 'TO22' son el
    mismo camion) pero se muestran con la escritura de la primera fila que los
    trajo, que es la que el inspector uso.

    Con `window` (lista de claves 'AAAA-MM') la dinamica se limita a esos
    meses. Importa para el indicador de flota mantenida: el maestro tiene una
    hoja de dinamica por periodo justamente porque 'cuantos equipos distintos
    se mantuvieron' solo significa algo dentro de un rango — sumar 2024 a 2026
    da un numero mayor que la flota actual y un porcentaje de mas de 100%.
    """
    allowed = set(window) if window else None
    counts = collections.Counter()
    display = {}
    months = set()
    undated = 0

    for row in inspections or []:
        vehicle = str(row.get("vehicle_id") or "").strip()
        if not vehicle:
            continue
        month = month_key(row.get("date"))
        if not month:
            if allowed is None:
                undated += 1
            continue
        if allowed is not None and month not in allowed:
            continue
        key = vehicle.upper()
        display.setdefault(key, vehicle)
        counts[(key, month)] += 1
        months.add(month)

    ordered_months = list(window) if window else (
        month_range(min(months), max(months)) if months else [])
    vehicles = sorted(display.values(), key=_natural_key)
    by_display = {display[k]: k for k in display}
    counts_display = {(name, m): counts.get((by_display[name], m), 0)
                      for name in vehicles for m in ordered_months
                      if counts.get((by_display[name], m), 0)}

    return Pivot(vehicles=vehicles, months=ordered_months,
                 counts=counts_display, undated=undated)


# ---------------------------------------------------------------------------
# Indicadores por mes
# ---------------------------------------------------------------------------
@dataclass
class MonthKpi:
    month: str            # 'AAAA-MM'
    reviewed: int         # equipos con >= 1 inspeccion
    inspections: int      # filas del mes
    fleet_size: int = 0   # flota total configurada para ese mes
    pct: float = None     # revisados / flota total


def monthly_kpis(pivot: Pivot, fleet_sizes: dict = None,
                 months: list = None) -> list:
    """Una fila por mes con revisados, inspecciones y % de inspeccion.

    Un mes sin flota configurada se queda con `pct=None` en vez de con 0: la
    diferencia entre 'no se inspecciono nada' y 'no sabemos cuantos equipos
    habia' es justo la que no se puede perder en un indicador de cumplimiento.
    Para no obligar a cargar los 24 meses a mano, se arrastra hacia adelante el
    ultimo tamano conocido — la flota cambia de a pocos equipos por mes.
    """
    sizes = {str(k): int(v) for k, v in (fleet_sizes or {}).items() if v}
    keys = list(months if months is not None else pivot.months)

    rows, carried = [], 0
    for month in keys:
        size = sizes.get(month, 0)
        if size:
            carried = size
        elif carried:
            size = carried
        reviewed = pivot.reviewed(month)
        rows.append(MonthKpi(
            month=month,
            reviewed=reviewed,
            inspections=pivot.col_total(month),
            fleet_size=size,
            pct=(reviewed / size) if size else None))
    return rows


def fleet_maintenance_pct(pivot: Pivot, kpis: list):
    """Flota mantenida / flota total del ultimo mes con tamano conocido.

    Es el '% Fleet maintenance' del maestro (=827/829): que porcentaje de la
    flota ACTUAL llego a tener alguna inspeccion en todo el periodo.
    """
    size = 0
    for kpi in kpis or []:
        if kpi.fleet_size:
            size = kpi.fleet_size
    if not size:
        return None
    return pivot.maintained_fleet / size


def last_month_kpi(kpis: list):
    """El mes mas reciente CON actividad; si no hay ninguno, el ultimo."""
    for kpi in reversed(kpis or []):
        if kpi.reviewed:
            return kpi
    return kpis[-1] if kpis else None


# ---------------------------------------------------------------------------
# Cortes simples de las inspecciones
# ---------------------------------------------------------------------------
def count_by(inspections: list, field_name: str, top: int = 0,
             other_label: str = "OTHER") -> list:
    """[(etiqueta, cuenta)] ordenado de mayor a menor.

    Con `top` se agrupa la cola en una sola barra: una grafica con 40
    departamentos no se lee, y las 8 primeras ya explican el 90% del volumen.
    """
    counter = collections.Counter()
    for row in inspections or []:
        label = str(row.get(field_name) or "").strip().upper()
        if label:
            counter[label] += 1
    items = counter.most_common()
    if top and len(items) > top:
        rest = sum(count for _label, count in items[top:])
        items = items[:top] + [(other_label, rest)]
    return items


# ---------------------------------------------------------------------------
# Resumen de los movimientos de tag
# ---------------------------------------------------------------------------
@dataclass
class TagMonth:
    month: str
    smu: int = 0
    tag: int = 0
    removals: int = 0

    @property
    def total(self) -> int:
        return self.smu + self.tag


def tag_monthly(movements: list, months: list = None) -> list:
    """Instalados por mes, separados en SMU y TAG.

    'Instalado' excluye los retiros, igual que el =SUMPRODUCT(... <>"REMOVAL")
    del consolidado. Los retiros se llevan aparte en `removals` porque tambien
    hay que verlos: un mes con muchas altas y muchas bajas no es lo mismo que
    un mes de crecimiento.
    """
    smu = collections.Counter()
    tag = collections.Counter()
    removed = collections.Counter()
    seen = set()

    for row in movements or []:
        month = month_key(row.get("date"))
        if not month:
            continue
        seen.add(month)
        move = str(row.get("move_type") or "").strip().upper()
        if move == tag_reader.MOVE_REMOVAL:
            removed[month] += 1
            continue
        if str(row.get("device_type") or "").strip().upper() == \
                tag_reader.DEVICE_SMU:
            smu[month] += 1
        else:
            tag[month] += 1

    keys = list(months) if months is not None else (
        month_range(min(seen), max(seen)) if seen else [])
    return [TagMonth(month=m, smu=smu[m], tag=tag[m], removals=removed[m])
            for m in keys]


def tag_by_move_type(movements: list) -> tuple:
    """(meses, {tipo_de_movimiento: [cuentas por mes]}) para la grafica apilada."""
    counts = collections.Counter()
    seen = set()
    for row in movements or []:
        month = month_key(row.get("date"))
        if not month:
            continue
        seen.add(month)
        move = str(row.get("move_type") or "").strip().upper()
        counts[(move if move in tag_reader.MOVE_TYPES
                else tag_reader.MOVE_NEW, month)] += 1
    months = month_range(min(seen), max(seen)) if seen else []
    series = {move: [counts.get((move, m), 0) for m in months]
              for move in tag_reader.MOVE_TYPES
              if any(counts.get((move, m), 0) for m in months)}
    return months, series


def tag_weekly(movements: list, weeks: int = 26) -> list:
    """[(lunes, instalados)] de las ultimas `weeks` semanas con movimiento."""
    counter = collections.Counter()
    for row in movements or []:
        move = str(row.get("move_type") or "").strip().upper()
        if move == tag_reader.MOVE_REMOVAL:
            continue
        monday = tag_reader.week_monday(row.get("date"))
        if monday:
            counter[monday] += 1
    items = sorted(counter.items())
    return items[-weeks:] if weeks and len(items) > weeks else items


def tag_totals(movements: list) -> dict:
    """Totales para las tarjetas: movimientos, instalados, retirados, archivos."""
    installed = removed = 0
    files = set()
    for row in movements or []:
        move = str(row.get("move_type") or "").strip().upper()
        if move == tag_reader.MOVE_REMOVAL:
            removed += 1
        else:
            installed += 1
        if row.get("source_file"):
            files.add(row["source_file"])
    return {"total": len(movements or []), "installed": installed,
            "removed": removed, "files": len(files)}
