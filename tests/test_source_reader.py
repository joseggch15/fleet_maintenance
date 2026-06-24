# -*- coding: utf-8 -*-
import mapping
import source_reader


def test_reads_submissions(source_workbook):
    subs = source_reader.read_submissions(source_workbook)
    assert len(subs) == 2
    first = subs[0]
    assert first[mapping.H_CODE] == "MOFFM-1"
    assert first[mapping.H_VEHICLE] == 836
    assert first[mapping.H_TAGTYPE] == "SMU"


def test_header_found_below_title_rows(source_workbook):
    # El encabezado real esta en la fila 4 (filas 1-2 son titulos).
    subs = source_reader.read_submissions(source_workbook)
    # Ninguna submission debe ser una fila de titulo.
    for s in subs:
        assert s.get(mapping.H_ID) is not None


def test_submission_label(source_workbook):
    subs = source_reader.read_submissions(source_workbook)
    assert source_reader.submission_label(subs[0]) == "MOFFM-1 (vehiculo 836)"
