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


# ---------------------------------------------------------------------------
# Reparacion de fechas contra el periodo del archivo
# ---------------------------------------------------------------------------
REF = datetime.date(2025, 12, 10)      # cierre de la semana del archivo


def test_date_inside_the_period_is_untouched():
    day = datetime.date(2025, 12, 5)
    assert tag_reader.repair_date(day, REF) == (day, "")


def test_late_entry_within_a_month_is_untouched():
    """Cargar hoy un movimiento de hace tres semanas es normal, no un error."""
    day = datetime.date(2025, 11, 20)
    assert tag_reader.repair_date(day, REF) == (day, "")


def test_year_typed_ahead_is_fixed():
    fixed, note = tag_reader.repair_date(datetime.date(2027, 12, 5), REF)
    assert fixed == datetime.date(2025, 12, 5)
    assert note.startswith(tag_reader.NOTE_FIXED_YEAR)
    assert "05/12/2027" in note          # el original queda anotado


def test_year_typed_behind_is_fixed():
    reference = datetime.date(2026, 2, 11)
    fixed, note = tag_reader.repair_date(datetime.date(2025, 2, 10), reference)
    assert fixed == datetime.date(2026, 2, 10)
    assert note.startswith(tag_reader.NOTE_FIXED_YEAR)


def test_day_and_month_swapped_is_fixed():
    # 12/05 en el archivo de la semana del 05/12: se invirtieron dia y mes.
    fixed, note = tag_reader.repair_date(datetime.date(2025, 5, 12), REF)
    assert fixed == datetime.date(2025, 12, 5)
    assert note.startswith(tag_reader.NOTE_FIXED_SWAP)


def test_only_one_year_can_ever_fit_the_window():
    """La ventana de correccion dura 34 dias: dos anos no caben a la vez.

    Es lo que hace que la correccion de ano sea segura y no una apuesta entre
    varias opciones. La guarda contra correcciones ambiguas sigue en el codigo
    por si alguien ensancha la ventana.
    """
    reference = datetime.date(2025, 6, 5)
    fixed, note = tag_reader.repair_date(datetime.date(2027, 5, 6), reference)
    assert fixed == datetime.date(2025, 5, 6)
    assert note.startswith(tag_reader.NOTE_FIXED_YEAR)


def test_unfixable_date_is_flagged_but_not_changed():
    day = datetime.date(2024, 3, 15)     # ningun ano la acerca a la semana
    fixed, note = tag_reader.repair_date(day, REF)
    assert fixed == day
    assert note.startswith(tag_reader.NOTE_SUSPECT_DATE)


def test_repair_can_be_turned_off():
    day = datetime.date(2027, 12, 5)
    fixed, note = tag_reader.repair_date(day, REF, repair=False)
    assert fixed == day
    assert note.startswith(tag_reader.NOTE_SUSPECT_DATE)


def test_without_a_reference_nothing_is_corrected():
    """Sin periodo en el nombre no hay con que probar que la fecha esta mal."""
    day = datetime.date(2027, 12, 5)
    assert tag_reader.repair_date(day, None) == (day, "")


def test_file_reference_reads_digits_and_month_names():
    assert tag_reader.file_reference(
        "Inventory Tag Installed 10122025.xlsx") == datetime.date(2025, 12, 10)
    assert tag_reader.file_reference(
        "Inventory Tag Installed 05082026-12082026.xlsx") == \
        datetime.date(2026, 8, 12)
    assert tag_reader.file_reference(
        "Tag Installed April_2025.xlsx") == datetime.date(2025, 4, 30)
    assert tag_reader.file_reference(
        "Inventory Tag Installed 2201205.xlsx") is None


def test_repair_applies_when_reading_a_file(tmp_path):
    import openpyxl
    path = tmp_path / "Inventory Tag Installed 10122025.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["TYPE", "DATE", "ID", "Tag", "Cost Center", "Department"])
    ws.append(["NEW INSTALLATION", datetime.datetime(2027, 12, 5), "C-MD-FS-21",
               "56B56E1D", 14386, "MINE_GEO"])
    wb.save(str(path))

    fixed = tag_reader.read_file(str(path))[0]
    assert fixed["date"] == "2025-12-05"
    assert tag_reader.NOTE_FIXED_YEAR in fixed["note"]

    raw = tag_reader.read_file(str(path), repair=False)[0]
    assert raw["date"] == "2027-12-05"
    assert tag_reader.NOTE_SUSPECT_DATE in raw["note"]


def test_read_paths_reports_how_many_were_repaired(weekly_tag_files):
    result = tag_reader.read_folder(weekly_tag_files)
    assert result["repaired"] == 0        # la fixture no trae anos mal escritos
    assert result["suspect"] == 1         # la fecha imposible '19/19/2025'


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
