# -*- coding: utf-8 -*-
"""
Preferencias del usuario, persistidas junto al codigo.

Guarda idioma, tema, geometria de la ventana, las ultimas rutas usadas y la
flota total de cada mes en un JSON simple. Se lee UNA vez al arrancar y se
escribe cuando cambia una preferencia.

La escritura es tolerante a fallos a proposito: si el disco esta lleno o la
carpeta es de solo lectura, la aplicacion debe seguir funcionando con los
valores por defecto en vez de caerse al arrancar.
"""
from __future__ import annotations

import json
import os

import i18n
import theme

_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "user_settings.json")


# ---------------------------------------------------------------------------
# Flota total por mes
# ---------------------------------------------------------------------------
#
# Es el denominador de '% Inspection per month'. NO se puede calcular con los
# datos que tiene el software: la flota total incluye equipos que ese mes no se
# inspeccionaron y que por lo tanto no aparecen en ninguna fila. En el Excel
# del cliente esta escrito a mano mes por mes (=AA4/400, =AB4/405, ...).
#
# Estos valores son los que trae 'PIVOT SUMMARY 2025' del maestro
# '260814_Fleet Tag Inventory and Maintenance.xlsx', para que el software
# reproduzca el mismo porcentaje que el Excel desde el primer arranque. El
# usuario los edita desde el tablero y quedan guardados aqui.
DEFAULT_FLEET_SIZES = {
    "2025-01": 400, "2025-02": 405, "2025-03": 625, "2025-04": 620,
    "2025-05": 620, "2025-06": 631, "2025-07": 642, "2025-08": 674,
    "2025-09": 691, "2025-10": 680, "2025-11": 682, "2025-12": 687,
    "2026-01": 687, "2026-02": 708, "2026-03": 747, "2026-04": 751,
    "2026-05": 770, "2026-06": 795, "2026-07": 829,
}

DEFAULTS = {
    "language": i18n.ES,
    "theme": theme.LIGHT,
    "window": {"w": 1420, "h": 900, "maximized": False},
    "fleet_sizes": dict(DEFAULT_FLEET_SIZES),
    # Cuantos meses muestra el tablero. 0 = todos.
    "dashboard_months": 24,
    # Granularidad de las graficas y del resumen de tags: day/week/month/year.
    "tag_grain": "month",
    "backup_target": True,
    "store_on_append": True,
    "last_source_file": "",
    "last_target_file": "",
    "last_tag_folder": "",
    "last_export_dir": "",
}

_data: dict = dict(DEFAULTS)


def load() -> dict:
    """Lee el archivo si existe y aplica idioma y tema."""
    global _data
    _data = dict(DEFAULTS)
    _data["fleet_sizes"] = dict(DEFAULT_FLEET_SIZES)
    try:
        if os.path.exists(_FILE):
            with open(_FILE, encoding="utf-8") as fh:
                stored = json.load(fh)
            if isinstance(stored, dict):
                _data.update({k: v for k, v in stored.items() if k in DEFAULTS})
    except (OSError, ValueError):
        # Archivo corrupto o ilegible: se sigue con los valores por defecto.
        pass
    if not isinstance(_data.get("fleet_sizes"), dict):
        _data["fleet_sizes"] = dict(DEFAULT_FLEET_SIZES)
    i18n.set_language(_data.get("language", i18n.ES))
    theme.set_theme(_data.get("theme", theme.LIGHT))
    return _data


def save() -> None:
    try:
        with open(_FILE, "w", encoding="utf-8") as fh:
            json.dump(_data, fh, indent=2, ensure_ascii=False)
    except OSError:
        pass


def get(key: str, default=None):
    return _data.get(key, DEFAULTS.get(key, default))


def set_(key: str, value) -> None:
    _data[key] = value
    save()


def set_language(lang: str) -> None:
    i18n.set_language(lang)
    set_("language", i18n.current())


def set_theme(name: str) -> None:
    theme.set_theme(name)
    set_("theme", theme.current())


# ---------------------------------------------------------------------------
# Flota por mes
# ---------------------------------------------------------------------------
def fleet_sizes() -> dict:
    sizes = _data.get("fleet_sizes") or {}
    return {k: int(v) for k, v in sizes.items()
            if str(v).strip() not in ("", "0", "None")}


def set_fleet_sizes(sizes: dict) -> None:
    clean = {}
    for month, total in (sizes or {}).items():
        try:
            value = int(total)
        except (TypeError, ValueError):
            continue
        if value > 0:
            clean[str(month)] = value
    set_("fleet_sizes", clean)
