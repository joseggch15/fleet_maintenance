# -*- coding: utf-8 -*-
"""Base local: alta, deduplicacion, filtros y borrado."""
import datetime

import tag_reader


def _row(date, vehicle, hours=100, inspector="RANCHO", smu="SMU"):
    return {"B": date, "C": vehicle, "E": vehicle, "F": "Y", "G": hours,
            "H": None, "I": "VIU OK", "J": 1, "K": "N", "L": "N", "M": "N",
            "N": smu, "O": "HAUL TRUCK", "P": None, "Q": inspector,
            "R": "NEWMONT", "S": None}


def test_inspection_from_row_maps_every_column(temp_store):
    record = temp_store.inspection_from_row(
        _row(datetime.datetime(2026, 5, 15), 836), "MOFFM-1", "export.xlsx")
    assert record["date"] == "2026-05-15"
    assert record["vehicle_id"] == "836"
    assert record["status"] == "VIU OK"
    assert record["submission_code"] == "MOFFM-1"
    assert record["source_file"] == "export.xlsx"
    assert set(name for name, _col in temp_store.INSPECTION_FIELDS) <= \
        set(record)


def test_add_inspections_is_idempotent(temp_store):
    records = [temp_store.inspection_from_row(
        _row(datetime.datetime(2026, 5, 15), 836), "MOFFM-1")]
    assert temp_store.add_inspections(records) == {"added": 1, "skipped": 0}
    assert temp_store.add_inspections(records) == {"added": 0, "skipped": 1}
    assert temp_store.count_inspections() == 1


def test_same_inspection_from_two_sources_stored_once(temp_store):
    """Del formulario y, mas tarde, del historico del maestro."""
    row = _row(datetime.datetime(2026, 5, 15), 836)
    from_form = temp_store.inspection_from_row(row, "MOFFM-1", "export.xlsx")
    from_master = temp_store.inspection_from_row(row, "", "maestro.xlsx")
    temp_store.add_inspections([from_form])
    temp_store.add_inspections([from_master])
    assert temp_store.count_inspections() == 1


def test_genuine_repeats_are_kept(temp_store):
    """Dos revisiones identicas el mismo dia son dos inspecciones, no una."""
    row = temp_store.inspection_from_row(
        _row(datetime.datetime(2026, 5, 15), 836))
    assert temp_store.add_inspections([row, dict(row)])["added"] == 2
    assert temp_store.count_inspections() == 2
    # Reimportar el mismo archivo no agrega una tercera.
    assert temp_store.add_inspections([row, dict(row)])["added"] == 0


def test_filters_by_year_owner_and_search(temp_store):
    temp_store.add_inspections([
        temp_store.inspection_from_row(_row(datetime.datetime(2025, 4, 2), 829)),
        temp_store.inspection_from_row(
            _row(datetime.datetime(2026, 7, 9), "LPL0986",
                 inspector="REGILLIO")),
    ])
    assert len(temp_store.inspections(year=2025)) == 1
    assert len(temp_store.inspections(search="LPL")) == 1
    assert len(temp_store.inspections(search="REGILLIO")) == 1
    assert len(temp_store.inspections(owner="NEWMONT")) == 2
    assert len(temp_store.inspections(owner="SEMC")) == 0


def test_rows_without_date_sort_last(temp_store):
    temp_store.add_inspections([
        temp_store.inspection_from_row(_row(None, "SIN-FECHA")),
        temp_store.inspection_from_row(_row(datetime.datetime(2026, 7, 9), 829)),
    ])
    rows = temp_store.inspections()
    assert rows[-1]["vehicle_id"] == "SIN-FECHA"


def test_delete_inspections(temp_store):
    temp_store.add_inspections([
        temp_store.inspection_from_row(_row(datetime.datetime(2026, 7, 9), 829))])
    row_id = temp_store.inspections()[0]["id"]
    assert temp_store.delete_inspections([row_id]) == 1
    assert temp_store.count_inspections() == 0


def test_movements_dedupe_on_overlapping_weeks(temp_store, weekly_tag_files):
    records = tag_reader.read_folder(weekly_tag_files)["records"]
    result = temp_store.add_movements(records)
    # 7 leidos, uno repetido entre el archivo y su semana solapada.
    assert result == {"added": 6, "skipped": 1}
    assert temp_store.add_movements(records)["added"] == 0
    assert temp_store.count_movements() == 6


def test_movement_filters(temp_store, weekly_tag_files):
    temp_store.add_movements(tag_reader.read_folder(weekly_tag_files)["records"])
    assert len(temp_store.movements(move_type=tag_reader.MOVE_REMOVAL)) == 1
    assert len(temp_store.movements(device=tag_reader.DEVICE_SMU)) == 1
    # Tres: la fila vieja con 'mine_ops' en minusculas cuenta igual.
    assert len(temp_store.movements(department="MINE_OPS")) == 3
    assert len(temp_store.movements(year=2026)) == 4


def test_clear_movements(temp_store, weekly_tag_files):
    temp_store.add_movements(tag_reader.read_folder(weekly_tag_files)["records"])
    assert temp_store.clear_movements() == 6
    assert temp_store.count_movements() == 0


def test_iso_date_accepts_the_formats_of_both_sources(temp_store):
    assert temp_store.iso_date("15/05/2026") == "2026-05-15"
    assert temp_store.iso_date(datetime.date(2026, 5, 15)) == "2026-05-15"
    assert temp_store.iso_date("no es fecha") == ""


def test_number_keeps_text_that_is_not_numeric(temp_store):
    assert temp_store.number("31093.3") == 31093.3
    assert temp_store.number("1") == 1
    assert temp_store.number("N/A") == "N/A"
    assert temp_store.number("") is None
