# -*- coding: utf-8 -*-
import datetime

import mapping


def test_normalize_locked():
    assert mapping.normalize_locked("Yes") == "Y"
    assert mapping.normalize_locked("No") == "N"
    assert mapping.normalize_locked("Not applicable") == "N/A"
    assert mapping.normalize_locked(None) is None


def test_tag_kind():
    assert mapping.tag_kind("SMU") == "SMU"
    assert mapping.tag_kind("Truck Tag") == "TAG"
    assert mapping.tag_kind("Half-Moon Tag") == "TAG"
    assert mapping.tag_kind(None) is None
    assert mapping.tag_kind("") is None


def test_pick_hours_prefers_hourmeter():
    assert mapping.pick_hours({mapping.H_HOUR: 9176, mapping.H_KM: 5000}) == 9176


def test_pick_hours_falls_back_to_kilometer():
    assert mapping.pick_hours({mapping.H_HOUR: None, mapping.H_KM: 14599}) == 14599
    assert mapping.pick_hours({mapping.H_KM: 14599}) == 14599


def test_coerce_datetime_text_ddmmyyyy():
    dt = mapping.coerce_datetime("15/05/2026 13:39:00")
    assert dt == datetime.datetime(2026, 5, 15, 13, 39, 0)


def test_revision_falls_back_to_submitted():
    sub = {mapping.H_SUBMITTED: datetime.datetime(2026, 5, 15, 16, 40)}
    assert mapping.revision_datetime(sub) == datetime.datetime(2026, 5, 15, 16, 40)


def test_submission_to_row_full_mapping():
    sub = {
        mapping.H_REVISION: "15/05/2026 13:39:00",
        mapping.H_VEHICLE: 836, mapping.H_FMS: 836, mapping.H_TAGTYPE: "SMU",
        mapping.H_KM: 14599, mapping.H_EQUIP: "Water Truck",
        mapping.H_INLETS: 1, mapping.H_ADDLOCK: "Not applicable",
        mapping.H_DRAIN: "Not applicable", mapping.H_COMPANY: "Newmont",
        mapping.H_INSPECTOR: "Rouchon Radi", mapping.H_REMARKS: "via ok",
    }
    row = mapping.submission_to_row(sub)
    assert row["B"] == datetime.datetime(2026, 5, 15)   # solo fecha
    assert row["C"] == 836
    assert row["E"] == 836
    assert row["F"] == "Y"            # default System fitted
    assert row["G"] == 14599         # hours/odo (km)
    assert row["I"] == "VIU OK"      # default Status
    assert row["J"] == 1
    assert row["K"] == "N/A"
    assert row["M"] == "N"           # default Fast fill
    assert row["N"] == "SMU"
    assert row["O"] == "Water Truck"
    assert row["P"] == "via ok"
    assert row["Q"] == "Rouchon Radi"
    assert row["R"] == "Newmont"
    # A y D no las pone el mapeo (son formulas del writer)
    assert "A" not in row and "D" not in row
