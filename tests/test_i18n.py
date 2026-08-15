# -*- coding: utf-8 -*-
"""Idioma, formatos y temas."""
import datetime

import pytest

import i18n
import tag_reader
import theme


@pytest.fixture(autouse=True)
def spanish():
    i18n.set_language(i18n.ES)
    theme.set_theme(theme.LIGHT)
    yield
    i18n.set_language(i18n.ES)
    theme.set_theme(theme.LIGHT)


def test_every_key_exists_in_both_languages():
    """Una clave a medio traducir sale en espanol dentro de la ventana inglesa."""
    missing = [key for key, entry in i18n._UI.items()
               if not entry.get(i18n.ES) or not entry.get(i18n.EN)]
    assert missing == []


def test_unknown_key_returns_the_key():
    assert i18n.t("no.existe") == "no.existe"


def test_month_label_matches_the_excel_of_the_client():
    assert i18n.month_label("2025-01") == "ene-25"
    i18n.set_language(i18n.EN)
    assert i18n.month_label("2025-01") == "Jan-25"


def test_month_label_accepts_dates():
    assert i18n.month_label(datetime.date(2026, 7, 9)) == "jul-26"


def test_date_is_ddmmyyyy_in_both_languages():
    """Los archivos de origen traen dd/mm/aaaa; cambiarlo haria ambiguo 01/02."""
    day = datetime.date(2026, 2, 1)
    assert i18n.fmt_date(day) == "01/02/2026"
    i18n.set_language(i18n.EN)
    assert i18n.fmt_date(day) == "01/02/2026"


def test_number_separators_change_with_the_language():
    assert i18n.fmt_number(1234.5, 1) == "1.234,5"
    i18n.set_language(i18n.EN)
    assert i18n.fmt_number(1234.5, 1) == "1,234.5"


def test_percentage_format():
    assert i18n.fmt_pct(0.3875, 1) == "38,8%"
    assert i18n.fmt_pct(None) == ""


def test_domain_values_are_translated_for_display():
    assert i18n.tr_value(tag_reader.MOVE_REMOVAL) == "Retiro"
    i18n.set_language(i18n.EN)
    assert i18n.tr_value(tag_reader.MOVE_REMOVAL) == "Removal"


def test_unknown_domain_value_passes_through():
    assert i18n.tr_value("VIU OK") == "VIU OK"


def test_note_keeps_the_original_detail():
    note = i18n.tr_note("%s:19/19/2025" % tag_reader.NOTE_BAD_DATE)
    assert "19/19/2025" in note
    assert note.startswith("Fecha invalida")


def test_theme_palettes_define_the_same_keys():
    light = set(theme.palette(theme.LIGHT))
    dark = set(theme.palette(theme.DARK))
    assert light == dark


def test_stylesheet_uses_the_active_palette():
    light = theme.stylesheet(theme.LIGHT)
    dark = theme.stylesheet(theme.DARK)
    assert theme.palette(theme.LIGHT)["bg"] in light
    assert theme.palette(theme.DARK)["bg"] in dark
    assert light != dark


def test_excel_color_is_always_the_light_palette():
    """El Excel se imprime y se comparte: no debe salir con fondo oscuro."""
    theme.set_theme(theme.DARK)
    assert theme.excel_color("primary") == \
        theme.palette(theme.LIGHT)["primary"].lstrip("#")


def test_device_and_move_colors_come_from_the_palette():
    assert theme.device_color("SMU") == theme.palette()["smu"]
    assert theme.device_color("TAG") == theme.palette()["tag"]
    assert theme.move_color(tag_reader.MOVE_REMOVAL) == \
        theme.palette()["removal"]
