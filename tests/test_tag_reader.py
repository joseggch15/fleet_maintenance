# -*- coding: utf-8 -*-
"""Lectura de los archivos semanales 'Inventory Tag Installed'."""
import datetime

import tag_reader


def test_device_type_from_tag_format():
    assert tag_reader.device_type("CA:CC:A3:FA:4C:2B") == tag_reader.DEVICE_SMU
    assert tag_reader.device_type("56B2B853") == tag_reader.DEVICE_TAG
    assert tag_reader.device_type("") == tag_reader.DEVICE_TAG


def test_normalize_type_accepts_case_and_spacing():
    assert tag_reader.normalize_type("Tag updated ") == tag_reader.MOVE_UPDATED
    assert tag_reader.normalize_type("new installation") == tag_reader.MOVE_NEW
    assert tag_reader.normalize_type("REMOVAL") == tag_reader.MOVE_REMOVAL
    # Un ID escrito en la columna TYPE no es un movimiento.
    assert tag_reader.normalize_type("CPR0988") is None


def test_parse_date_text_and_excel():
    assert tag_reader.parse_date("10/01/2025") == datetime.date(2025, 1, 10)
    assert tag_reader.parse_date(datetime.datetime(2026, 6, 1, 8, 59)) == \
        datetime.date(2026, 6, 1)
    assert tag_reader.parse_date("19/19/2025") is None


def test_week_monday():
    # 2026-06-03 es miercoles; su lunes es el 1.
    assert tag_reader.week_monday(datetime.date(2026, 6, 3)) == \
        datetime.date(2026, 6, 1)


def test_file_period_from_name():
    assert tag_reader.file_period("Inventory Tag Installed 01102025.xlsx") == \
        "01/10/2025"
    assert tag_reader.file_period(
        "Inventory Tag Installed 05082026-12082026.xlsx") == \
        "05/08/2026 - 12/08/2026"
    # Nombre con un digito de menos: no se inventa una fecha.
    assert tag_reader.file_period("Inventory Tag Installed 2201205.xlsx") == ""


def test_find_files_walks_subfolders(weekly_tag_files):
    files = tag_reader.find_files(weekly_tag_files)
    assert len(files) == 3


def test_reads_all_formats(weekly_tag_files):
    result = tag_reader.read_folder(weekly_tag_files)
    assert not result["errors"]
    records = result["records"]
    # 2 del viejo + 3 del nuevo + 2 del solapado; la hoja 'Summary' se ignora.
    assert len(records) == 7
    assert all(r["source_file"].endswith(".xlsx") for r in records)


def test_old_file_without_type_is_flagged(weekly_tag_files):
    records = tag_reader.read_folder(weekly_tag_files)["records"]
    old = [r for r in records if r["equipment_id"] == "C-155"][0]
    assert old["move_type"] == tag_reader.MOVE_NEW
    assert old["type_inferred"] is True
    assert tag_reader.NOTE_NO_TYPE_COLUMN in old["note"]


def test_unreadable_date_keeps_the_row_with_a_note(weekly_tag_files):
    records = tag_reader.read_folder(weekly_tag_files)["records"]
    broken = [r for r in records if r["equipment_id"] == "CPR0988"][0]
    assert broken["date"] == ""
    assert tag_reader.NOTE_BAD_DATE in broken["note"]
    assert "19/19/2025" in broken["note"]


def test_department_normalized_to_upper(weekly_tag_files):
    records = tag_reader.read_folder(weekly_tag_files)["records"]
    old = [r for r in records if r["equipment_id"] == "C-155"][0]
    assert old["department"] == "MINE_OPS"


def test_alias_headers_are_mapped(weekly_tag_files):
    records = tag_reader.read_folder(weekly_tag_files)["records"]
    old = [r for r in records if r["equipment_id"] == "C-155"][0]
    assert old["cost_center"] == "14386"


def test_output_sheet_is_ignored(tmp_path):
    """Un consolidado ya generado no se vuelve a importar como si fuera origen."""
    import openpyxl
    path = tmp_path / "Tag Installed - Consolidated.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["TYPE", "DATE", "ID", "Tag", "Cost Center", "Department",
               "Source File", "Inferred type"])
    ws.append(["NEW INSTALLATION", datetime.datetime(2026, 6, 1), "LVE2136",
               "56B2E3EE", 14387, "MINE_OPS", "semana.xlsx", "No"])
    wb.save(str(path))
    assert tag_reader.read_file(str(path)) == []
