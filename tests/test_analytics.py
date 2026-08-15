# -*- coding: utf-8 -*-
"""Dinamica, indicadores por mes y resumen de tags."""
import analytics
import tag_reader


def _insp(date, vehicle, equipment="HAUL TRUCK", status="VIU OK"):
    return {"date": date, "vehicle_id": vehicle, "equipment_type": equipment,
            "status": status}


def _move(date, tag, move_type=tag_reader.MOVE_NEW, dept="MINE_OPS"):
    return {"date": date, "tag": tag, "move_type": move_type,
            "device_type": tag_reader.device_type(tag), "department": dept}


# ---------------------------------------------------------------------------
# Meses
# ---------------------------------------------------------------------------
def test_month_key():
    assert analytics.month_key("2026-07-09") == "2026-07"
    assert analytics.month_key("") == ""
    assert analytics.month_key(None) == ""


def test_month_range_keeps_empty_months():
    """Un mes sin inspecciones es informacion: nadie salio a campo."""
    assert analytics.month_range("2025-11", "2026-02") == [
        "2025-11", "2025-12", "2026-01", "2026-02"]


def test_month_range_rejects_inverted_range():
    assert analytics.month_range("2026-02", "2025-11") == []


def test_last_months():
    months = ["2025-11", "2025-12", "2026-01"]
    assert analytics.last_months(months, 2) == ["2025-12", "2026-01"]
    assert analytics.last_months(months, 0) == months


# ---------------------------------------------------------------------------
# Dinamica
# ---------------------------------------------------------------------------
def test_pivot_counts_inspections_per_vehicle_and_month():
    pivot = analytics.build_pivot([
        _insp("2025-01-10", "829"), _insp("2025-01-20", "829"),
        _insp("2025-02-03", "802"),
    ])
    assert pivot.months == ["2025-01", "2025-02"]
    assert pivot.count("829", "2025-01") == 2
    assert pivot.col_total("2025-01") == 2
    assert pivot.row_total("829") == 2
    assert pivot.total_inspections == 3


def test_reviewed_counts_equipment_not_inspections():
    """Tres visitas al mismo camion son UN equipo revisado."""
    pivot = analytics.build_pivot([
        _insp("2025-01-10", "829"), _insp("2025-01-20", "829"),
        _insp("2025-01-25", "829"), _insp("2025-01-30", "802"),
    ])
    assert pivot.col_total("2025-01") == 4
    assert pivot.reviewed("2025-01") == 2


def test_vehicles_grouped_case_insensitively():
    pivot = analytics.build_pivot([_insp("2025-01-10", "to22"),
                                   _insp("2025-01-11", "TO22")])
    assert pivot.maintained_fleet == 1
    assert pivot.vehicles == ["to22"]      # se muestra como llego la primera


def test_vehicles_sorted_naturally():
    pivot = analytics.build_pivot([
        _insp("2025-01-10", "127"), _insp("2025-01-10", "3"),
        _insp("2025-01-10", "10"), _insp("2025-01-10", "LVE0198"),
    ])
    assert pivot.vehicles == ["3", "10", "127", "LVE0198"]


def test_rows_without_date_are_counted_apart():
    pivot = analytics.build_pivot([_insp("", "829"),
                                   _insp("2025-01-10", "802")])
    assert pivot.undated == 1
    assert pivot.total_inspections == 1


def test_window_limits_the_maintained_fleet():
    """Sin ventana se suman equipos de anos viejos y el % pasa de 100."""
    rows = [_insp("2024-05-10", "999"), _insp("2025-01-10", "829")]
    assert analytics.build_pivot(rows).maintained_fleet == 2
    assert analytics.build_pivot(rows, ["2025-01"]).maintained_fleet == 1


# ---------------------------------------------------------------------------
# Indicadores
# ---------------------------------------------------------------------------
def test_monthly_kpis_percentage():
    pivot = analytics.build_pivot([_insp("2025-01-10", "829"),
                                   _insp("2025-01-11", "802")])
    kpis = analytics.monthly_kpis(pivot, {"2025-01": 400})
    assert kpis[0].reviewed == 2
    assert kpis[0].fleet_size == 400
    assert abs(kpis[0].pct - 0.005) < 1e-9


def test_month_without_fleet_size_has_no_percentage():
    pivot = analytics.build_pivot([_insp("2025-01-10", "829")])
    assert analytics.monthly_kpis(pivot, {})[0].pct is None


def test_fleet_size_is_carried_forward():
    """La flota cambia de a pocos equipos; no hay que cargar los 24 meses."""
    pivot = analytics.build_pivot([_insp("2025-01-10", "829"),
                                   _insp("2025-02-10", "802")])
    kpis = analytics.monthly_kpis(pivot, {"2025-01": 400})
    assert kpis[1].fleet_size == 400


def test_fleet_maintenance_pct_uses_the_last_known_size():
    pivot = analytics.build_pivot([_insp("2025-01-10", "829"),
                                   _insp("2025-02-10", "802")])
    kpis = analytics.monthly_kpis(pivot, {"2025-01": 400, "2025-02": 500})
    assert abs(analytics.fleet_maintenance_pct(pivot, kpis) - 2 / 500) < 1e-9


def test_last_month_kpi_skips_trailing_empty_months():
    pivot = analytics.build_pivot([_insp("2025-01-10", "829"),
                                   _insp("2025-03-10", "802")])
    kpis = analytics.monthly_kpis(pivot, {})
    assert analytics.last_month_kpi(kpis).month == "2025-03"


def test_count_by_groups_the_tail():
    rows = [_insp("2025-01-10", str(i), equipment="TIPO%d" % i)
            for i in range(5)]
    items = analytics.count_by(rows, "equipment_type", top=2,
                               other_label="OTROS")
    assert len(items) == 3
    assert items[-1] == ("OTROS", 3)


# ---------------------------------------------------------------------------
# Tags
# ---------------------------------------------------------------------------
def test_tag_by_period_excludes_removals_from_installed():
    rows = [
        _move("2026-06-01", "CA:CC:A3:FA:4C:2B"),
        _move("2026-06-02", "56B2D9A6"),
        _move("2026-06-03", "56B36ACF", tag_reader.MOVE_REMOVAL),
    ]
    month = analytics.tag_by_period(rows)[0]
    assert (month.smu, month.tag, month.total, month.removals) == (1, 1, 2, 1)


def test_tag_by_period_groups_by_each_grain():
    rows = [_move("2026-06-01", "A"), _move("2026-06-03", "B"),
            _move("2026-07-15", "C")]
    counts = {g: len(analytics.tag_by_period(rows, g))
              for g in analytics.GRAINS}
    # Del 1 de junio al 15 de julio: 45 dias, 7 semanas, 2 meses, 1 ano.
    assert counts[analytics.GRAIN_DAY] == 45
    assert counts[analytics.GRAIN_WEEK] == 7
    assert counts[analytics.GRAIN_MONTH] == 2
    assert counts[analytics.GRAIN_YEAR] == 1
    for grain in analytics.GRAINS:
        assert sum(p.total for p in analytics.tag_by_period(rows, grain)) == 3


def test_week_grain_falls_on_monday():
    # El 3 de junio de 2026 es miercoles: su cubeta es la del lunes 1.
    assert analytics.period_key(
        "2026-06-03", analytics.GRAIN_WEEK) == "2026-06-01"


def test_period_key_per_grain():
    day = "2026-06-03"
    assert analytics.period_key(day, analytics.GRAIN_DAY) == day
    assert analytics.period_key(day, analytics.GRAIN_MONTH) == "2026-06"
    assert analytics.period_key(day, analytics.GRAIN_YEAR) == "2026"
    assert analytics.period_key("", analytics.GRAIN_DAY) == ""


def test_period_range_keeps_empty_buckets():
    assert analytics.period_range(
        "2026", "2028", analytics.GRAIN_YEAR) == ["2026", "2027", "2028"]
    assert len(analytics.period_range("2026-06-01", "2026-06-22",
                                      analytics.GRAIN_WEEK)) == 4


def test_last_periods():
    rows = analytics.tag_by_period([_move("2026-06-01", "A"),
                                    _move("2026-08-01", "B")])
    assert len(rows) == 3
    assert len(analytics.last_periods(rows, 2)) == 2
    assert len(analytics.last_periods(rows, 0)) == 3


def test_tag_by_move_type_only_returns_used_types():
    months, series = analytics.tag_by_move_type([
        _move("2026-06-01", "56B2D9A6"),
        _move("2026-06-02", "56B36ACF", tag_reader.MOVE_REMOVAL)])
    assert months == ["2026-06"]
    assert set(series) == {tag_reader.MOVE_NEW, tag_reader.MOVE_REMOVAL}


def test_tag_by_move_type_honours_the_grain():
    periods, series = analytics.tag_by_move_type(
        [_move("2026-06-01", "A"), _move("2026-06-09", "B")],
        analytics.GRAIN_WEEK)
    assert periods == ["2026-06-01", "2026-06-08"]
    assert series[tag_reader.MOVE_NEW] == [1, 1]


def test_tag_totals():
    rows = [_move("2026-06-01", "A"),
            _move("2026-06-02", "B", tag_reader.MOVE_REMOVAL)]
    rows[0]["source_file"] = "semana1.xlsx"
    rows[1]["source_file"] = "semana1.xlsx"
    totals = analytics.tag_totals(rows)
    assert totals == {"total": 2, "installed": 1, "removed": 1, "files": 1}
