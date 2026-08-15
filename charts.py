# -*- coding: utf-8 -*-
"""
Graficas de matplotlib embebidas en Qt.

Son las mismas del Excel del cliente, hechas aqui para poder verlas sin abrir
el maestro:

  'Newmont & BPs'            barras de equipos revisados por mes y, sobre el
                             eje derecho, la linea del % de inspeccion. La
                             linea es la mitad del mensaje: 121 equipos
                             revisados no significa lo mismo con una flota de
                             400 que con una de 830, y dos barras de altura
                             parecida no dejan ver esa diferencia.
  'Reviewed tags/SMU'        torta con el reparto de revisados entre los meses
                             del periodo elegido.
  Resumen de tags            barras de SMU y TAG instalados por mes, mas los
                             cortes por tipo de movimiento, departamento y
                             semana.

Todo el color sale de `theme` y todo el texto de `i18n`: cambiar de tema o de
idioma es volver a dibujar, no tocar este archivo.
"""
from __future__ import annotations

import matplotlib
matplotlib.use("QtAgg")

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg  # noqa: E402
from matplotlib.figure import Figure                             # noqa: E402
from matplotlib.ticker import FuncFormatter, MaxNLocator         # noqa: E402

import i18n     # noqa: E402
import theme    # noqa: E402
import analytics  # noqa: E402

# Claves de las graficas del tablero de mantenimiento.
CHART_BARS = "bars"
CHART_PIE = "pie"
CHART_EQUIPMENT = "equipment"
CHART_STATUS = "status"
MAINTENANCE_CHARTS = (CHART_BARS, CHART_PIE, CHART_EQUIPMENT, CHART_STATUS)

# Claves de las graficas de tags instalados. Las dos primeras son series de
# tiempo y respetan la granularidad elegida (dia, semana, mes o ano).
CHART_TAG_INSTALLED = "tag_installed"
CHART_TAG_TYPE = "tag_type"
CHART_TAG_DEPT = "tag_dept"
TAG_CHARTS = (CHART_TAG_INSTALLED, CHART_TAG_TYPE, CHART_TAG_DEPT)


def _pct_formatter(value, _pos=None) -> str:
    return i18n.fmt_pct(value, 0)


def _int_formatter(value, _pos=None) -> str:
    return i18n.fmt_number(value, 0)


class ChartCanvas(FigureCanvasQTAgg):
    """Lienzo reutilizable: se le pide una grafica y se redibuja."""

    def __init__(self, parent=None, width=9.0, height=4.4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi,
                          facecolor=theme.color("chart_bg"))
        super().__init__(self.fig)
        if parent is not None:
            self.setParent(parent)

    # -- utilidades ---------------------------------------------------------
    def _fresh(self):
        self.fig.set_facecolor(theme.color("chart_bg"))
        self.fig.clear()
        ax = self.fig.add_subplot(111)
        ax.set_facecolor(theme.color("chart_bg"))
        return ax

    @staticmethod
    def _style(ax):
        p = theme.palette()
        ax.tick_params(colors=p["chart_text"], labelsize=8)
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
        for spine in ("left", "bottom"):
            ax.spines[spine].set_color(p["chart_grid"])
        ax.grid(True, axis="y", color=p["chart_grid"], linewidth=0.7,
                alpha=0.7)
        ax.set_axisbelow(True)
        ax.yaxis.label.set_color(p["chart_text"])
        ax.xaxis.label.set_color(p["chart_text"])

    def _title(self, ax, text, caption=None):
        p = theme.palette()
        ax.set_title(text, fontsize=12, fontweight="bold",
                     color=p["primary"], pad=14 if caption else 8)
        if caption:
            ax.text(0.5, 1.015, caption, transform=ax.transAxes, ha="center",
                    va="bottom", fontsize=8, color=p["chart_muted"])

    def _legend(self, ax, **kw):
        p = theme.palette()
        legend = ax.legend(frameon=False, fontsize=8, **kw)
        for text in legend.get_texts():
            text.set_color(p["chart_text"])
        return legend

    def _empty(self, ax):
        p = theme.palette()
        ax.text(0.5, 0.5, i18n.t("chart.empty"), ha="center", va="center",
                fontsize=11, color=p["chart_muted"], transform=ax.transAxes)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        self.fig.tight_layout()
        self.draw()

    def _done(self, ax):
        self._style(ax)
        self.fig.tight_layout()
        self.draw()

    @staticmethod
    def _period_ticks(ax, periods, grain=analytics.GRAIN_MONTH):
        """Etiquetas de tiempo rotadas y, si son muchas, salteadas.

        Con 24 cubetas las etiquetas se pisan; mostrar una de cada dos es peor
        que rotarlas, porque el ojo pierde la referencia de que barra es cual.
        Por eso se rotan siempre y solo se saltea a partir de 30.
        """
        step = max(1, (len(periods) + 29) // 30)
        ax.set_xticks(range(0, len(periods), step))
        ax.set_xticklabels([i18n.period_label(periods[i], grain)
                            for i in range(0, len(periods), step)],
                           rotation=60, ha="right", fontsize=7.5)

    def _month_ticks(self, ax, months):
        self._period_ticks(ax, months, analytics.GRAIN_MONTH)

    # =======================================================================
    # Tablero de mantenimiento
    # =======================================================================
    def reviewed_bars(self, kpis):
        """'Newmont & BPs': barras de revisados + linea de % de inspeccion."""
        ax = self._fresh()
        kpis = list(kpis or [])
        if not kpis:
            self._empty(ax)
            return

        p = theme.palette()
        x = range(len(kpis))
        values = [k.reviewed for k in kpis]
        ax.bar(list(x), values, 0.62, color=p["reviewed"],
               label=i18n.t("chart.legend_reviewed"))

        top = max(values) if values else 0
        for i, value in enumerate(values):
            if value:
                ax.annotate(i18n.fmt_number(value), (i, value),
                            textcoords="offset points", xytext=(0, 3),
                            ha="center", fontsize=7, color=p["chart_text"])
        ax.set_ylim(0, top * 1.18 if top else 1)
        ax.set_ylabel(i18n.t("chart.axis_count"), fontsize=9)

        # El % va en un eje propio: comparte la grafica con un conteo que va de
        # 0 a 250, y sobre ese eje una serie de 0 a 0,4 seria una linea plana
        # pegada al piso.
        ax2 = ax.twinx()
        ax2.set_facecolor("none")
        pcts = [(k.pct if k.pct is not None else float("nan")) for k in kpis]
        ax2.plot(list(x), pcts, color=p["pct"], linewidth=1.8, marker="o",
                 markersize=3.5, label=i18n.t("chart.legend_pct"))
        ax2.set_ylim(0, max([v for v in pcts if v == v] or [0.5]) * 1.35)
        ax2.yaxis.set_major_formatter(FuncFormatter(_pct_formatter))
        ax2.tick_params(colors=p["chart_text"], labelsize=8)
        ax2.set_ylabel(i18n.t("chart.axis_pct"), fontsize=9,
                       color=p["chart_text"])
        for spine in ("top", "left"):
            ax2.spines[spine].set_visible(False)
        ax2.spines["right"].set_color(p["chart_grid"])
        ax2.grid(False)

        self._month_ticks(ax, [k.month for k in kpis])
        self._title(ax, i18n.t("xls.chart_bars"))

        handles = ax.get_legend_handles_labels()
        extra = ax2.get_legend_handles_labels()
        legend = ax.legend(handles[0] + extra[0], handles[1] + extra[1],
                           frameon=False, fontsize=8, loc="upper left",
                           ncol=2)
        for text in legend.get_texts():
            text.set_color(p["chart_text"])
        self._done(ax)

    # Mas alla de una docena de porciones la torta deja de leerse: las
    # etiquetas se pisan y ningun porcentaje se distingue del vecino. Con un
    # periodo largo se muestran los ultimos meses y se avisa en el subtitulo.
    PIE_MONTHS = 12

    def reviewed_pie(self, kpis):
        """Reparto de equipos revisados entre los meses del periodo."""
        ax = self._fresh()
        data = [(k.month, k.reviewed) for k in (kpis or []) if k.reviewed]
        if not data:
            self._empty(ax)
            return

        trimmed = len(data) > self.PIE_MONTHS
        if trimmed:
            data = data[-self.PIE_MONTHS:]

        p = theme.palette()
        labels = [i18n.month_label(m) for m, _n in data]
        values = [n for _m, n in data]
        colors = theme.series_colors(len(values))
        _wedges, _texts, autotexts = ax.pie(
            values, labels=labels, colors=colors,
            # Las porciones chicas se dejan sin porcentaje adentro: el numero
            # no cabe y termina encima del de al lado.
            autopct=lambda pct: ("%.0f%%" % pct) if pct >= 4 else "",
            startangle=90, counterclock=False, pctdistance=0.72,
            textprops={"fontsize": 8, "color": p["chart_text"]},
            wedgeprops={"linewidth": 0.8, "edgecolor": p["chart_bg"]})
        for text in autotexts:
            text.set_fontsize(7.5)
            text.set_color("#FFFFFF")
            text.set_fontweight("bold")
        ax.axis("equal")
        caption = i18n.t("chart.caption_total", n=i18n.fmt_number(sum(values)))
        if trimmed:
            caption += "   ·   " + i18n.t("chart.caption_last_months",
                                          n=self.PIE_MONTHS)
        self._title(ax, i18n.t("xls.chart_pie"), caption)
        self.fig.tight_layout()
        self.draw()

    def category_bars(self, items, title):
        """Barras horizontales de un corte simple (tipo de equipo, estado)."""
        ax = self._fresh()
        items = list(items or [])
        if not items:
            self._empty(ax)
            return

        p = theme.palette()
        items = items[::-1]          # el mayor arriba
        labels = [str(label) for label, _n in items]
        values = [n for _label, n in items]
        ax.barh(range(len(items)), values, 0.68,
                color=theme.series_colors(len(items)))
        ax.set_yticks(range(len(items)))
        ax.set_yticklabels(labels, fontsize=8)
        ax.xaxis.set_major_formatter(FuncFormatter(_int_formatter))
        ax.xaxis.set_major_locator(MaxNLocator(integer=True))
        span = max(values) if values else 0
        for i, value in enumerate(values):
            ax.annotate(i18n.fmt_number(value), (value, i),
                        textcoords="offset points", xytext=(4, 0),
                        va="center", fontsize=7.5, color=p["chart_text"])
        ax.set_xlim(0, span * 1.12 if span else 1)
        self._title(ax, title)
        self._style(ax)
        ax.grid(True, axis="x", color=p["chart_grid"], linewidth=0.7, alpha=0.7)
        ax.grid(False, axis="y")
        self.fig.tight_layout()
        self.draw()

    # =======================================================================
    # Tags instalados por semana
    # =======================================================================
    @staticmethod
    def _period_caption(total, trimmed):
        caption = i18n.t("chart.caption_total", n=i18n.fmt_number(total))
        if trimmed:
            caption += "   ·   " + i18n.t("chart.caption_last_periods",
                                          n=trimmed)
        return caption

    def tag_installed_bars(self, rows, grain=analytics.GRAIN_MONTH,
                           trimmed=0):
        """SMU y TAG instalados por cubeta de tiempo, uno al lado del otro."""
        ax = self._fresh()
        rows = list(rows or [])
        if not rows or not any(r.total for r in rows):
            self._empty(ax)
            return

        p = theme.palette()
        x = range(len(rows))
        width = 0.4
        ax.bar([i - width / 2 for i in x], [r.smu for r in rows], width,
               color=p["smu"], label=i18n.t("chart.legend_smu"))
        ax.bar([i + width / 2 for i in x], [r.tag for r in rows], width,
               color=p["tag"], label=i18n.t("chart.legend_tag"))
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax.set_ylabel(i18n.t("chart.axis_count"), fontsize=9)
        self._period_ticks(ax, [r.period for r in rows], grain)
        self._title(ax, "%s — %s" % (i18n.t("tags.chart_installed"),
                                     i18n.t("grain." + grain)),
                    self._period_caption(sum(r.total for r in rows), trimmed))
        self._legend(ax, loc="upper left", ncol=2)
        self._done(ax)

    def tag_move_bars(self, periods, series, grain=analytics.GRAIN_MONTH,
                      trimmed=0):
        """Movimientos apilados por tipo (alta, reemplazo, retiro)."""
        ax = self._fresh()
        if not periods or not series:
            self._empty(ax)
            return

        bottom = [0] * len(periods)
        total = 0
        for move, values in series.items():
            ax.bar(range(len(periods)), values, 0.62, bottom=bottom,
                   color=theme.move_color(move), label=i18n.tr_value(move))
            bottom = [b + v for b, v in zip(bottom, values)]
            total += sum(values)
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax.set_ylabel(i18n.t("chart.axis_count"), fontsize=9)
        self._period_ticks(ax, periods, grain)
        self._title(ax, "%s — %s" % (i18n.t("tags.chart_type"),
                                     i18n.t("grain." + grain)),
                    self._period_caption(total, trimmed))
        self._legend(ax, loc="upper left", ncol=2)
        self._done(ax)


# ---------------------------------------------------------------------------
# Despachador
# ---------------------------------------------------------------------------
def render_maintenance(canvas: ChartCanvas, key: str, pivot, kpis,
                       inspections) -> None:
    if key == CHART_PIE:
        canvas.reviewed_pie(kpis)
    elif key == CHART_EQUIPMENT:
        canvas.category_bars(
            analytics.count_by(inspections, "equipment_type", top=12,
                               other_label=i18n.t("chart.other")),
            i18n.t("dash.chart_equipment"))
    elif key == CHART_STATUS:
        canvas.category_bars(
            analytics.count_by(inspections, "status", top=10,
                               other_label=i18n.t("chart.other")),
            i18n.t("dash.chart_status"))
    else:
        canvas.reviewed_bars(kpis)


def render_tags(canvas: ChartCanvas, key: str, movements,
                grain: str = analytics.GRAIN_MONTH, limit: int = 0) -> None:
    """Dibuja la grafica pedida de la pestana de tags.

    `limit` recorta a las ultimas N cubetas. Es necesario en el grano diario:
    dos anos de datos son mas de setecientas barras de un pixel.
    """
    if key == CHART_TAG_DEPT:
        canvas.category_bars(
            analytics.count_by(movements, "department", top=12,
                               other_label=i18n.t("chart.other")),
            i18n.t("tags.chart_dept"))
        return

    if key == CHART_TAG_TYPE:
        periods, series = analytics.tag_by_move_type(movements, grain)
        totals = [sum(values[i] for values in series.values())
                  for i in range(len(periods))]
        cut = analytics.focus_periods(totals, limit, lambda n: not n)
        canvas.tag_move_bars(periods[cut:],
                             {m: v[cut:] for m, v in series.items()},
                             grain, len(periods) - cut if cut else 0)
        return

    rows = analytics.tag_by_period(movements, grain)
    cut = analytics.focus_periods(rows, limit, lambda r: not r.total)
    canvas.tag_installed_bars(rows[cut:], grain,
                              len(rows) - cut if cut else 0)
