# -*- coding: utf-8 -*-
"""
Interfaz grafica (PySide6) del cargador de mantenimiento de flota.

Cuatro pestanas:

  Importar submissions  el flujo original — leer el export del formulario,
                        revisar fila por fila y volcarlo a la hoja
                        'Full List 2024-2025' del Excel maestro. Ademas puede
                        guardar esas filas en la base local y traer de vuelta
                        el historico que el maestro ya tiene.
  Full List             las inspecciones almacenadas, con filtros y exportacion
                        a un Excel con el mismo formato del maestro.
  Tablero               las graficas del maestro (revisados por mes y % de
                        inspeccion) calculadas sobre la base local.
  Tags por semana       consolidacion de la carpeta 'Tag Installed Per Week' y
                        su resumen de instalacion.

Idioma y tema se cambian desde la barra superior. Cambiarlos RECONSTRUYE la
ventana entera (`_rebuild`) en vez de recorrer widget por widget: es la unica
forma de que no queden textos viejos colgados en un dialogo o en una leyenda de
grafica. Los datos ya leidos se conservan en memoria, asi que cambiar de idioma
no cuesta otra lectura de 4.000 filas.

Ejecutar:  python run.py
"""
from __future__ import annotations

import datetime
import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog,
    QDialogButtonBox, QFileDialog, QFrame, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox,
    QProgressDialog, QPushButton, QSpinBox, QSplitter, QTableWidget,
    QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)

from matplotlib.backends.backend_qtagg import NavigationToolbar2QT

import analytics
import charts
import excel_writer
import i18n
import mapping
import report_export
import settings
import source_reader
import store
import tag_reader
import theme

# Columnas de datos que se muestran/editan en la previsualizacion (B, C, E..S).
PREVIEW_COLUMNS = mapping.DATA_COLUMNS

_PREVIEW_LABELS = {
    "B": "col.date_target", "C": "col.vehicle", "E": "col.fms",
    "F": "col.fitted", "G": "col.hours", "H": "col.fms_hours",
    "I": "col.status", "J": "col.inlets", "K": "col.addl_locked",
    "L": "col.drain_locked", "M": "col.leaking", "N": "col.smu_tags",
    "O": "col.equipment", "P": "col.remarks", "Q": "col.inspectors",
    "R": "col.owner", "S": "col.remedial",
}

# La tabla de 'Full List' se llena con widgets, no con un modelo: pasadas unas
# miles de filas eso se nota al filtrar. Se muestran las mas recientes y se
# avisa; los calculos y la exportacion siguen usando todas.
_TABLE_LIMIT = 2000

# Estado de una fila de la previsualizacion frente a lo ya cargado.
_STATE_NEW = "new"
_STATE_STORED = "stored"
_STATE_MASTER = "master"

# Cuantas cubetas muestra la grafica de tags segun la granularidad. Con dos
# anos de datos, 'diario' son mas de 700 barras: ni se leen ni sirven.
_GRAIN_CHART_LIMIT = {
    analytics.GRAIN_DAY: 45,
    analytics.GRAIN_WEEK: 30,
    analytics.GRAIN_MONTH: 36,
    analytics.GRAIN_YEAR: 0,
}


class _BackgroundWorker(QThread):
    """Ejecuta una funcion en un hilo aparte para no congelar la GUI."""

    def __init__(self, func, args=(), kwargs=None, holder=None):
        super().__init__()
        self._func = func
        self._args = args
        self._kwargs = kwargs or {}
        self._holder = holder if holder is not None else {}

    def run(self):
        try:
            self._holder["result"] = self._func(*self._args, **self._kwargs)
        except BaseException as exc:  # noqa: BLE001
            self._holder["error"] = exc


def _run_with_progress(parent, title, message, func, *args, **kwargs):
    """Corre func con una barra de progreso modal (indeterminada o real)."""
    holder = {"result": None, "error": None}
    dlg = QProgressDialog(message, None, 0, 0, parent)
    dlg.setWindowTitle(title)
    dlg.setWindowModality(Qt.ApplicationModal)
    dlg.setMinimumDuration(0)
    dlg.setAutoClose(False)
    dlg.setAutoReset(False)
    dlg.setCancelButton(None)
    dlg.setMinimumWidth(440)

    from PySide6.QtCore import QMetaObject, Qt as _Qt, Q_ARG
    import inspect
    sig_kwargs = dict(kwargs)
    try:
        if "progress_cb" in inspect.signature(func).parameters:
            def _cb(i, total, label=""):
                QMetaObject.invokeMethod(dlg, "setMaximum",
                                         _Qt.QueuedConnection, Q_ARG(int, int(total)))
                QMetaObject.invokeMethod(dlg, "setValue",
                                         _Qt.QueuedConnection, Q_ARG(int, int(i)))
                if label:
                    QMetaObject.invokeMethod(
                        dlg, "setLabelText", _Qt.QueuedConnection,
                        Q_ARG(str, "%s   (%d/%d)" % (label, i, total)))
            sig_kwargs["progress_cb"] = _cb
    except (TypeError, ValueError):
        pass

    worker = _BackgroundWorker(func, args, sig_kwargs, holder)
    worker.finished.connect(dlg.close)
    worker.start()
    dlg.exec()
    worker.wait()
    if holder["error"] is not None:
        raise holder["error"]
    return holder["result"]


def _fmt_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime(i18n.date_format())
    return str(value)


def _card(parent_layout, key_value, key_label, key_hint):
    """Tarjeta de indicador: valor grande, etiqueta y aclaracion."""
    frame = QFrame()
    frame.setObjectName("card")
    lay = QVBoxLayout(frame)
    lay.setContentsMargins(12, 8, 12, 8)
    lay.setSpacing(1)
    value = QLabel(key_value)
    value.setObjectName("cardValue")
    label = QLabel(key_label)
    label.setObjectName("cardLabel")
    hint = QLabel(key_hint)
    hint.setObjectName("cardHint")
    hint.setWordWrap(True)
    lay.addWidget(value)
    lay.addWidget(label)
    lay.addWidget(hint)
    parent_layout.addWidget(frame)
    frame.value_label = value
    return frame


class FleetSizeDialog(QDialog):
    """Editor de la flota total mes a mes.

    Es el denominador de '% Inspection per month' y no se puede deducir de las
    inspecciones: la flota incluye equipos que ese mes nadie toco. En el Excel
    del cliente esta escrito a mano dentro de cada formula; aqui al menos queda
    en un solo lugar y se guarda con las preferencias.
    """

    def __init__(self, parent, months: list, sizes: dict):
        super().__init__(parent)
        self.setWindowTitle(i18n.t("fleet.title"))
        self.setStyleSheet(theme.stylesheet())
        self.resize(420, 560)

        lay = QVBoxLayout(self)
        hint = QLabel(i18n.t("fleet.hint"))
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        self.months = list(months)
        self.table = QTableWidget(len(self.months), 2)
        self.table.setHorizontalHeaderLabels(
            [i18n.t("dash.table_month"), i18n.t("dash.table_fleet")])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.Stretch)
        for row, month in enumerate(self.months):
            item = QTableWidgetItem(i18n.month_label(month))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, item)
            spin = QSpinBox()
            spin.setRange(0, 100000)
            spin.setValue(int(sizes.get(month, 0) or 0))
            self.table.setCellWidget(row, 1, spin)
        lay.addWidget(self.table, 1)

        fill = QPushButton(i18n.t("fleet.btn_fill"))
        fill.setObjectName("ghost")
        fill.clicked.connect(self._fill_down)
        lay.addWidget(fill)

        buttons = QDialogButtonBox(QDialogButtonBox.Save |
                                   QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText(i18n.t("fleet.btn_ok"))
        buttons.button(QDialogButtonBox.Cancel).setText(
            i18n.t("fleet.btn_cancel"))
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)

    def _fill_down(self):
        """Copia hacia abajo el ultimo valor distinto de cero.

        La flota cambia de a pocos equipos por mes; cargar 24 casillas a mano
        cuando 20 son el mismo numero es trabajo inutil.
        """
        last = 0
        for row in range(self.table.rowCount()):
            spin = self.table.cellWidget(row, 1)
            if spin.value():
                last = spin.value()
            elif last:
                spin.setValue(last)

    def sizes(self) -> dict:
        out = {}
        for row, month in enumerate(self.months):
            value = self.table.cellWidget(row, 1).value()
            if value:
                out[month] = value
        return out


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        store.init()

        self.source_path = settings.get("last_source_file") or ""
        self.target_path = settings.get("last_target_file") or ""
        if not os.path.exists(self.target_path):
            self.target_path = ""

        # Estado del formulario en curso (no se guarda en la base hasta que el
        # usuario lo pide).
        self.submissions = []
        self.preview_rows = []      # list[dict] {columna: valor}
        self.preview_codes = []     # codigo de submission por fila
        self.preview_checked = []   # list[bool]
        self.preview_state = []     # 'new' | 'stored' | 'master'

        # Huellas de la hoja 'Full List' del destino, para avisar antes de
        # duplicar. Se guarda con la ruta y la fecha del archivo: si el maestro
        # no cambio, no hace falta releer sus miles de filas.
        self._master_keys = {}
        self._master_stamp = None

        # Cache de la base: se recarga al importar o al borrar, no en cada
        # cambio de filtro ni al cambiar de idioma.
        self.inspections = []
        self.movements = []
        self.pivot = analytics.Pivot()
        self.kpis = []

        # Mientras se construye la ventana, los combos disparan sus senales
        # con datos a medio armar. Este candado hace que los refrescos se
        # ignoren hasta que la ventana este completa.
        self._loading = True
        self._build()
        self.inspections = store.inspections()
        self.movements = store.movements()
        self._refresh_views()

    # ------------------------------------------------------------------
    # Ciclo de vida
    # ------------------------------------------------------------------
    def closeEvent(self, event):
        settings.set_("window", {"w": self.width(), "h": self.height(),
                                 "maximized": self.isMaximized()})
        super().closeEvent(event)

    def _build(self):
        self._loading = True
        try:
            self.setWindowTitle(i18n.t("app.title"))
            self.setStyleSheet(theme.stylesheet())

            central = QWidget()
            outer = QVBoxLayout(central)
            outer.setContentsMargins(10, 8, 10, 8)
            outer.setSpacing(8)
            outer.addWidget(self._build_topbar())

            self.tabs = QTabWidget()
            self.tabs.addTab(self._tab_import(), i18n.t("tab.import"))
            self.tabs.addTab(self._tab_full_list(), i18n.t("tab.fulllist"))
            self.tabs.addTab(self._tab_dashboard(), i18n.t("tab.dashboard"))
            self.tabs.addTab(self._tab_tags(), i18n.t("tab.tags"))
            outer.addWidget(self.tabs, 1)

            self.setCentralWidget(central)
            self.statusBar().showMessage(i18n.t("app.ready"))
        finally:
            self._loading = False

    def _rebuild(self):
        """Reconstruye la ventana tras cambiar idioma o tema.

        Los datos leidos se conservan en memoria; lo unico que se rehace son
        los widgets. Se guarda antes que estaba eligiendo el usuario (pestana,
        grafica, periodo) para devolverlo a donde estaba: cambiar el idioma no
        deberia hacerle perder el sitio.
        """
        combo_names = ("chart_combo", "tag_chart_combo", "period_combo",
                       "grain_combo", "year_combo", "owner_combo",
                       "tag_year_combo", "tag_type_combo", "tag_device_combo",
                       "tag_dept_combo")
        state = {
            "tab": self.tabs.currentIndex(),
            "search": self.search_edit.text(),
            "combos": {name: getattr(self, name).currentData()
                       for name in combo_names},
        }
        self._build()
        self._populate_preview()
        self._populate_filters()
        self._populate_period_combo()
        self._populate_tag_filters()

        self.tabs.setCurrentIndex(state["tab"])
        self.search_edit.blockSignals(True)
        self.search_edit.setText(state["search"])
        self.search_edit.blockSignals(False)
        for name in combo_names:
            combo = getattr(self, name)
            index = combo.findData(state["combos"][name])
            if index >= 0:
                combo.blockSignals(True)
                combo.setCurrentIndex(index)
                combo.blockSignals(False)

        self._refresh_full_list()
        self._refresh_dashboard()
        self._refresh_tag_view()
        self._refresh_status()

    # ------------------------------------------------------------------
    # Barra superior
    # ------------------------------------------------------------------
    def _build_topbar(self) -> QWidget:
        bar = QWidget()
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(2, 0, 2, 0)

        title = QLabel(i18n.t("app.title"))
        title.setObjectName("title")
        lay.addWidget(title)
        lay.addSpacing(16)

        self.lbl_stored = QLabel()
        self.lbl_stored.setObjectName("subtitle")
        lay.addWidget(self.lbl_stored)
        lay.addStretch(1)

        lay.addWidget(QLabel(i18n.t("top.language")))
        self.lang_combo = QComboBox()
        for lang in i18n.LANGUAGES:
            self.lang_combo.addItem(i18n.LANGUAGE_NAMES[lang], lang)
        self.lang_combo.setCurrentIndex(
            self.lang_combo.findData(i18n.current()))
        self.lang_combo.currentIndexChanged.connect(self._on_language)
        lay.addWidget(self.lang_combo)

        lay.addSpacing(10)
        lay.addWidget(QLabel(i18n.t("top.theme")))
        self.theme_combo = QComboBox()
        self.theme_combo.addItem(i18n.t("top.theme_light"), theme.LIGHT)
        self.theme_combo.addItem(i18n.t("top.theme_dark"), theme.DARK)
        self.theme_combo.setCurrentIndex(
            self.theme_combo.findData(theme.current()))
        self.theme_combo.currentIndexChanged.connect(self._on_theme)
        lay.addWidget(self.theme_combo)
        return bar

    def _on_language(self, index):
        settings.set_language(self.lang_combo.itemData(index))
        self._rebuild()

    def _on_theme(self, index):
        settings.set_theme(self.theme_combo.itemData(index))
        self._rebuild()

    # ==================================================================
    # Pestana 1: importar submissions
    # ==================================================================
    def _tab_import(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.addWidget(self._build_controls())
        lay.addWidget(self._build_preview(), 1)
        lay.addWidget(self._build_footer())
        return page

    def _build_controls(self) -> QWidget:
        box = QGroupBox(i18n.t("import.files"))
        outer = QVBoxLayout(box)

        row = QHBoxLayout()
        btn_src = QPushButton(i18n.t("import.btn_source"))
        btn_src.clicked.connect(self._on_load_source)
        btn_tgt = QPushButton(i18n.t("import.btn_target"))
        btn_tgt.clicked.connect(self._on_pick_target)
        self.btn_history = QPushButton(i18n.t("import.btn_history"))
        self.btn_history.setObjectName("ghost")
        self.btn_history.clicked.connect(self._on_import_history)
        row.addWidget(btn_src)
        row.addWidget(btn_tgt)
        row.addWidget(self.btn_history)
        row.addStretch(1)
        outer.addLayout(row)

        hint = QLabel(i18n.t("import.history_hint"))
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        outer.addWidget(hint)

        status_row = QHBoxLayout()
        status_row.setSpacing(8)
        label = QLabel(i18n.t("import.state"))
        label.setObjectName("sectionTitle")
        status_row.addWidget(label)
        self.chip_source = QLabel()
        self.chip_target = QLabel()
        for chip in (self.chip_source, self.chip_target):
            chip.setTextFormat(Qt.RichText)
            chip.setMinimumHeight(26)
            status_row.addWidget(chip)
        status_row.addStretch(1)
        outer.addLayout(status_row)
        return box

    def _build_preview(self) -> QWidget:
        box = QGroupBox(i18n.t("import.preview"))
        lay = QVBoxLayout(box)
        hint = QLabel(i18n.t("import.preview_hint"))
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        headers = [i18n.t("import.col_include"),
                   i18n.t("import.col_state")] + [
            "%s · %s" % (col, i18n.t(_PREVIEW_LABELS[col]))
            for col in PREVIEW_COLUMNS]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.itemChanged.connect(self._on_preview_edited)
        lay.addWidget(self.table)

        sel = QHBoxLayout()
        btn_all = QPushButton(i18n.t("import.check_all"))
        btn_all.setObjectName("ghost")
        btn_all.clicked.connect(lambda: self._set_all_checks(True))
        btn_none = QPushButton(i18n.t("import.uncheck_all"))
        btn_none.setObjectName("ghost")
        btn_none.clicked.connect(lambda: self._set_all_checks(False))
        btn_recheck = QPushButton(i18n.t("import.btn_recheck"))
        btn_recheck.setObjectName("ghost")
        btn_recheck.clicked.connect(lambda: self._mark_duplicates(True))
        sel.addWidget(btn_all)
        sel.addWidget(btn_none)
        sel.addWidget(btn_recheck)
        sel.addStretch(1)
        self.lbl_count = QLabel(i18n.t("import.count", n=0))
        sel.addWidget(self.lbl_count)
        lay.addLayout(sel)
        return box

    def _build_footer(self) -> QWidget:
        widget = QWidget()
        lay = QHBoxLayout(widget)
        lay.setContentsMargins(0, 0, 0, 0)
        self.chk_backup = QCheckBox(i18n.t("import.backup"))
        self.chk_backup.setChecked(bool(settings.get("backup_target")))
        self.chk_backup.toggled.connect(
            lambda v: settings.set_("backup_target", bool(v)))
        self.chk_store = QCheckBox(i18n.t("import.also_store"))
        self.chk_store.setChecked(bool(settings.get("store_on_append")))
        self.chk_store.toggled.connect(
            lambda v: settings.set_("store_on_append", bool(v)))
        lay.addWidget(self.chk_backup)
        lay.addWidget(self.chk_store)
        lay.addStretch(1)

        self.btn_store = QPushButton(i18n.t("import.btn_store"))
        self.btn_store.setMinimumWidth(220)
        self.btn_store.clicked.connect(self._on_store_only)
        lay.addWidget(self.btn_store)

        self.btn_load = QPushButton(i18n.t("import.btn_append"))
        self.btn_load.setObjectName("accent")
        self.btn_load.setMinimumWidth(250)
        font = QFont()
        font.setBold(True)
        self.btn_load.setFont(font)
        self.btn_load.clicked.connect(self._on_append)
        lay.addWidget(self.btn_load)
        return widget

    # -- estado / chips -------------------------------------------------
    def _set_chip(self, chip, loaded, title, detail=""):
        chip.setObjectName("chipOk" if loaded else "chipWarn")
        chip.style().unpolish(chip)
        chip.style().polish(chip)
        if loaded:
            html = "<b>&#10003; %s</b>" % title
            if detail:
                html += " <span>(%s)</span>" % detail
            chip.setText(html)
        else:
            chip.setText("<b>&#9888; %s</b> %s" % (
                title, i18n.t("import.chip_missing")))

    def _refresh_status(self):
        if self.source_path and self.submissions:
            self._set_chip(
                self.chip_source, True, i18n.t("import.chip_source"),
                "%s &mdash; %s" % (os.path.basename(self.source_path),
                                   i18n.t("import.chip_rows",
                                          n=len(self.submissions))))
        else:
            self._set_chip(self.chip_source, False,
                           i18n.t("import.chip_source"))
        if self.target_path:
            self._set_chip(self.chip_target, True,
                           i18n.t("import.chip_target"),
                           os.path.basename(self.target_path))
        else:
            self._set_chip(self.chip_target, False,
                           i18n.t("import.chip_target"))

        has_rows = bool(self.preview_rows)
        self.btn_load.setEnabled(has_rows and bool(self.target_path))
        self.btn_store.setEnabled(has_rows)
        self.btn_history.setEnabled(bool(self.target_path))
        self.lbl_stored.setText(i18n.t(
            "top.stored",
            inspections=i18n.fmt_number(len(self.inspections)),
            tags=i18n.fmt_number(len(self.movements))))

    # -- tabla de previsualizacion --------------------------------------
    def _populate_preview(self):
        self._loading = True
        try:
            self.table.setRowCount(0)
            for index, rowdata in enumerate(self.preview_rows):
                row = self.table.rowCount()
                self.table.insertRow(row)
                check = QCheckBox()
                check.setChecked(self.preview_checked[index])
                check.toggled.connect(
                    lambda value, i=index: self._set_check(i, value))
                holder = QWidget()
                box = QHBoxLayout(holder)
                box.setContentsMargins(0, 0, 0, 0)
                box.setAlignment(Qt.AlignCenter)
                box.addWidget(check)
                self.table.setCellWidget(row, 0, holder)

                state = (self.preview_state[index]
                         if index < len(self.preview_state) else _STATE_NEW)
                label = QTableWidgetItem(i18n.t("import.state_" + state))
                label.setFlags(label.flags() & ~Qt.ItemIsEditable)
                label.setForeground(QColor(theme.color(
                    "accent" if state == _STATE_NEW else "warning")))
                self.table.setItem(row, 1, label)

                for col_index, col in enumerate(PREVIEW_COLUMNS, start=2):
                    self.table.setItem(
                        row, col_index,
                        QTableWidgetItem(_fmt_cell(rowdata.get(col))))
            self.table.resizeColumnsToContents()
        finally:
            self._loading = False
        self.lbl_count.setText(i18n.t("import.count",
                                      n=len(self.preview_rows)))

    def _set_check(self, index, value):
        if 0 <= index < len(self.preview_checked):
            self.preview_checked[index] = bool(value)

    def _set_all_checks(self, checked: bool):
        self.preview_checked = [checked] * len(self.preview_checked)
        self._populate_preview()

    def _on_preview_edited(self, item):
        """Lleva la edicion de la celda al modelo, conservando el tipo.

        Si el texto no cambio se deja el valor original: reconstruirlo desde el
        texto convertiria una fecha en cadena y el Excel la escribiria como
        texto, rompiendo la formula 'Date' del maestro.
        """
        if self._loading or item.column() < 2:
            return
        row, col = item.row(), PREVIEW_COLUMNS[item.column() - 2]
        if row >= len(self.preview_rows):
            return
        original = self.preview_rows[row].get(col)
        text = item.text().strip()
        if _fmt_cell(original) == text:
            return
        if text == "":
            self.preview_rows[row][col] = None
            return
        if col == "B":
            parsed = mapping.coerce_datetime(text)
            if parsed is not None:
                self.preview_rows[row][col] = datetime.datetime(
                    parsed.year, parsed.month, parsed.day)
                return
        self.preview_rows[row][col] = text

    # -- deteccion de duplicados -----------------------------------------
    def _master_key_counts(self) -> dict:
        """Huellas de la hoja 'Full List 2024-2025' del Excel destino.

        La base local se defiende sola de la doble carga, pero el maestro no:
        `append_rows` escribe lo que se le marque. Si el usuario vuelve a bajar
        el export del formulario —que trae las submissions viejas mas las
        nuevas— y lo carga entero, el maestro termina con las viejas dos veces.
        Por eso se leen sus huellas antes de mostrar la previsualizacion.
        """
        if not self.target_path or not os.path.exists(self.target_path):
            return {}
        try:
            stamp = (self.target_path, os.path.getmtime(self.target_path))
        except OSError:
            return {}
        if stamp == self._master_stamp:
            return self._master_keys

        try:
            rows = _run_with_progress(
                self, i18n.t("prog.history_title"),
                i18n.t("prog.history", sheet=source_reader.FULL_LIST_SHEET),
                source_reader.read_full_list, self.target_path)
        except Exception:
            # Un maestro ilegible no puede bloquear la previsualizacion: se
            # sigue con la comparacion contra la base local solamente.
            return {}

        counts = {}
        for row in rows:
            key = store.inspection_base_key(store.inspection_from_row(row))
            counts[key] = counts.get(key, 0) + 1
        self._master_keys = counts
        self._master_stamp = stamp
        return counts

    def _mark_duplicates(self, announce: bool = True):
        """Marca cada fila de la previsualizacion y desmarca las repetidas."""
        if not self.preview_rows:
            self.preview_state = []
            return

        records = [store.inspection_from_row(row, code, self.source_path)
                   for row, code in zip(self.preview_rows,
                                        self.preview_codes)]
        stored = store.count_new(records, store.inspection_key_counts())
        master = store.count_new(records, self._master_key_counts())

        self.preview_state = []
        for is_new_here, is_new_there in zip(stored, master):
            if not is_new_here:
                self.preview_state.append(_STATE_STORED)
            elif not is_new_there:
                self.preview_state.append(_STATE_MASTER)
            else:
                self.preview_state.append(_STATE_NEW)
        self.preview_checked = [state == _STATE_NEW
                                for state in self.preview_state]
        self._populate_preview()

        if not announce:
            return
        dupes = sum(1 for s in self.preview_state if s != _STATE_NEW)
        new = len(self.preview_state) - dupes
        QMessageBox.information(
            self, i18n.t("dlg.info"),
            i18n.t("import.dupes", n=dupes, new=new) if dupes
            else i18n.t("import.dupes_none", new=new))

    def _selected_rows(self) -> list:
        return [row for row, keep in zip(self.preview_rows,
                                         self.preview_checked) if keep]

    def _selected_records(self) -> list:
        """Filas marcadas como registros de la base local."""
        out = []
        for index, keep in enumerate(self.preview_checked):
            if not keep:
                continue
            out.append(store.inspection_from_row(
                self.preview_rows[index],
                self.preview_codes[index] if index < len(self.preview_codes)
                else "",
                self.source_path))
        return out

    # ==================================================================
    # Pestana 2: Full List almacenado
    # ==================================================================
    def _tab_full_list(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)

        title = QLabel(i18n.t("full.title"))
        title.setObjectName("sectionTitle")
        lay.addWidget(title)
        hint = QLabel(i18n.t("full.hint"))
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        filters = QHBoxLayout()
        filters.addWidget(QLabel(i18n.t("full.filter_year")))
        self.year_combo = QComboBox()
        self.year_combo.currentIndexChanged.connect(self._refresh_full_list)
        filters.addWidget(self.year_combo)

        filters.addWidget(QLabel(i18n.t("full.filter_owner")))
        self.owner_combo = QComboBox()
        self.owner_combo.currentIndexChanged.connect(self._refresh_full_list)
        filters.addWidget(self.owner_combo)

        filters.addWidget(QLabel(i18n.t("full.filter_search")))
        self.search_edit = QLineEdit()
        self.search_edit.setMinimumWidth(180)
        self.search_edit.textChanged.connect(self._refresh_full_list)
        filters.addWidget(self.search_edit)
        filters.addStretch(1)

        btn_refresh = QPushButton(i18n.t("full.btn_refresh"))
        btn_refresh.setObjectName("ghost")
        btn_refresh.clicked.connect(self._on_refresh_store)
        filters.addWidget(btn_refresh)
        lay.addLayout(filters)

        columns = [("col.date_target", "date"), ("col.vehicle", "vehicle_id"),
                   ("col.fms", "fms_id"), ("col.fitted", "system_fitted"),
                   ("col.hours", "equipment_hours"),
                   ("col.fms_hours", "fms_hours"), ("col.status", "status"),
                   ("col.inlets", "inlets"),
                   ("col.addl_locked", "addl_inlets_locked"),
                   ("col.drain_locked", "drain_valves_locked"),
                   ("col.leaking", "fast_fill_leaking"),
                   ("col.smu_tags", "smu_tags"),
                   ("col.equipment", "equipment_type"),
                   ("col.remarks", "remarks"),
                   ("col.inspectors", "inspectors"), ("col.owner", "owner"),
                   ("col.remedial", "remedial"), ("col.source", "source_file")]
        self.full_columns = columns
        self.full_table = QTableWidget(0, len(columns))
        self.full_table.setHorizontalHeaderLabels(
            [i18n.t(key) for key, _f in columns])
        self.full_table.setAlternatingRowColors(True)
        self.full_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.full_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.full_table.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.full_table, 1)

        bottom = QHBoxLayout()
        self.lbl_full_rows = QLabel()
        bottom.addWidget(self.lbl_full_rows)
        self.lbl_full_limit = QLabel()
        self.lbl_full_limit.setObjectName("hint")
        bottom.addWidget(self.lbl_full_limit)
        bottom.addStretch(1)
        btn_delete = QPushButton(i18n.t("full.btn_delete"))
        btn_delete.setObjectName("danger")
        btn_delete.clicked.connect(self._on_delete_rows)
        bottom.addWidget(btn_delete)
        btn_export = QPushButton(i18n.t("full.btn_export"))
        btn_export.setObjectName("accent")
        btn_export.clicked.connect(self._on_export_report)
        bottom.addWidget(btn_export)
        lay.addLayout(bottom)
        return page

    def _populate_filters(self):
        # Se bloquean las senales en vez de usar el candado global: llenar un
        # combo emite currentIndexChanged por cada item, y cada uno redibujaria
        # la tabla entera.
        for combo in (self.year_combo, self.owner_combo):
            combo.blockSignals(True)
        try:
            previous_year = self.year_combo.currentData()
            previous_owner = self.owner_combo.currentData()
            self.year_combo.clear()
            self.year_combo.addItem(i18n.t("full.all"), None)
            for year in sorted({r["date"][:4] for r in self.inspections
                                if r.get("date")}, reverse=True):
                self.year_combo.addItem(year, year)
            self.owner_combo.clear()
            self.owner_combo.addItem(i18n.t("full.all"), None)
            for owner in sorted({r["owner"] for r in self.inspections
                                 if r.get("owner")}):
                self.owner_combo.addItem(owner, owner)
            for combo, previous in ((self.year_combo, previous_year),
                                    (self.owner_combo, previous_owner)):
                index = combo.findData(previous)
                combo.setCurrentIndex(max(index, 0))
        finally:
            for combo in (self.year_combo, self.owner_combo):
                combo.blockSignals(False)

    def _filtered_inspections(self) -> list:
        year = self.year_combo.currentData()
        owner = self.owner_combo.currentData()
        text = (self.search_edit.text() or "").strip().upper()
        rows = self.inspections
        if year:
            rows = [r for r in rows if (r.get("date") or "").startswith(year)]
        if owner:
            rows = [r for r in rows if r.get("owner") == owner]
        if text:
            fields = ("vehicle_id", "fms_id", "inspectors", "equipment_type",
                      "remarks", "status")
            rows = [r for r in rows
                    if any(text in str(r.get(f) or "").upper()
                           for f in fields)]
        return rows

    def _refresh_full_list(self, *_args):
        if self._loading:
            return
        rows = self._filtered_inspections()
        shown = rows[:_TABLE_LIMIT]
        self.full_table.setRowCount(0)
        self.full_table.setRowCount(len(shown))
        for r, record in enumerate(shown):
            for c, (_key, field) in enumerate(self.full_columns):
                if field == "date":
                    text = i18n.fmt_date(store.parse_date(record.get("date")))
                elif field in ("system_fitted", "addl_inlets_locked",
                               "drain_valves_locked", "fast_fill_leaking"):
                    text = i18n.tr_value(record.get(field))
                else:
                    text = str(record.get(field) or "")
                item = QTableWidgetItem(text)
                item.setData(Qt.UserRole, record.get("id"))
                self.full_table.setItem(r, c, item)
        self.full_table.resizeColumnsToContents()
        self.lbl_full_rows.setText(i18n.t("full.rows", shown=len(rows),
                                          total=len(self.inspections)))
        self.lbl_full_limit.setText(
            i18n.t("full.limit", shown=len(shown), total=len(rows))
            if len(rows) > len(shown) else "")

    # ==================================================================
    # Pestana 3: tablero
    # ==================================================================
    def _tab_dashboard(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)

        cards = QHBoxLayout()
        cards.setSpacing(8)
        self.card_inspections = _card(cards, "0",
                                      i18n.t("dash.card_inspections"),
                                      i18n.t("dash.card_inspections_hint"))
        self.card_fleet = _card(cards, "0", i18n.t("dash.card_fleet"),
                                i18n.t("dash.card_fleet_hint"))
        self.card_pct = _card(cards, "-", i18n.t("dash.card_pct"),
                              i18n.t("dash.card_pct_hint"))
        self.card_last = _card(cards, "0", i18n.t("dash.card_last"),
                               i18n.t("dash.card_last_hint"))
        lay.addLayout(cards)

        controls = QHBoxLayout()
        controls.addWidget(QLabel(i18n.t("dash.months")))
        self.period_combo = QComboBox()
        self.period_combo.currentIndexChanged.connect(self._refresh_dashboard)
        controls.addWidget(self.period_combo)

        controls.addSpacing(14)
        controls.addWidget(QLabel(i18n.t("dash.chart")))
        self.chart_combo = QComboBox()
        for key, label in ((charts.CHART_BARS, "dash.chart_bars"),
                           (charts.CHART_PIE, "dash.chart_pie"),
                           (charts.CHART_EQUIPMENT, "dash.chart_equipment"),
                           (charts.CHART_STATUS, "dash.chart_status")):
            self.chart_combo.addItem(i18n.t(label), key)
        self.chart_combo.currentIndexChanged.connect(self._render_chart)
        controls.addWidget(self.chart_combo)
        controls.addStretch(1)

        btn_fleet = QPushButton(i18n.t("dash.btn_fleet_sizes"))
        btn_fleet.setObjectName("ghost")
        btn_fleet.clicked.connect(self._on_fleet_sizes)
        controls.addWidget(btn_fleet)
        lay.addLayout(controls)

        splitter = QSplitter(Qt.Vertical)
        chart_box = QWidget()
        chart_layout = QVBoxLayout(chart_box)
        chart_layout.setContentsMargins(0, 0, 0, 0)
        self.canvas = charts.ChartCanvas()
        chart_layout.addWidget(NavigationToolbar2QT(self.canvas, chart_box))
        chart_layout.addWidget(self.canvas, 1)
        splitter.addWidget(chart_box)

        self.kpi_table = QTableWidget(0, 5)
        self.kpi_table.setHorizontalHeaderLabels([
            i18n.t("dash.table_month"), i18n.t("dash.table_reviewed"),
            i18n.t("dash.table_rows"), i18n.t("dash.table_fleet"),
            i18n.t("dash.table_pct")])
        self.kpi_table.setAlternatingRowColors(True)
        self.kpi_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.kpi_table.verticalHeader().setVisible(False)
        self.kpi_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        splitter.addWidget(self.kpi_table)
        splitter.setSizes([520, 260])
        lay.addWidget(splitter, 1)
        return page

    def _populate_period_combo(self):
        self.period_combo.blockSignals(True)
        try:
            current = self.period_combo.currentData()
            self.period_combo.clear()
            self.period_combo.addItem(i18n.t("dash.months_all"), None)
            for count in (12, 24):
                self.period_combo.addItem(
                    i18n.t("dash.months_last", n=count), ("last", count))
            years = sorted({r["date"][:4] for r in self.inspections
                            if r.get("date")}, reverse=True)
            for year in years:
                self.period_combo.addItem(year, ("year", year))
            index = self.period_combo.findData(current)
            self.period_combo.setCurrentIndex(max(index, 0))
        finally:
            self.period_combo.blockSignals(False)

    def _dashboard_window(self):
        months = analytics.available_months(self.inspections)
        choice = self.period_combo.currentData()
        if not choice:
            return months
        kind, value = choice
        if kind == "last":
            return analytics.last_months(months, int(value))
        return analytics.year_months(months, value)

    def _export_window(self, rows: list) -> list:
        """Meses del resumen dinamico que se exporta.

        Se cruzan las dos cosas que el usuario eligio: el periodo del tablero y
        el filtro de la pestana Full List. Si el resumen cubriera un periodo
        mas ancho que las filas exportadas quedarian columnas de meses sin una
        sola fila detras en el mismo archivo.
        """
        available = analytics.available_months(rows)
        window = [m for m in self._dashboard_window() if m in set(available)]
        return window or available

    def _refresh_dashboard(self, *_args):
        if self._loading:
            return
        window = self._dashboard_window()
        self.pivot = analytics.build_pivot(self.inspections, window)
        self.kpis = analytics.monthly_kpis(self.pivot, settings.fleet_sizes())
        fleet_pct = analytics.fleet_maintenance_pct(self.pivot, self.kpis)
        last = analytics.last_month_kpi(self.kpis)

        self.card_inspections.value_label.setText(
            i18n.fmt_number(self.pivot.total_inspections))
        self.card_fleet.value_label.setText(
            i18n.fmt_number(self.pivot.maintained_fleet))
        self.card_pct.value_label.setText(
            i18n.fmt_pct(fleet_pct, 1) if fleet_pct is not None else "-")
        self.card_last.value_label.setText(
            "%s  ·  %s" % (i18n.month_label(last.month),
                           i18n.fmt_number(last.reviewed)) if last else "-")

        self.kpi_table.setRowCount(len(self.kpis))
        for row, kpi in enumerate(self.kpis):
            values = [i18n.month_label(kpi.month),
                      i18n.fmt_number(kpi.reviewed),
                      i18n.fmt_number(kpi.inspections),
                      i18n.fmt_number(kpi.fleet_size) if kpi.fleet_size else "",
                      i18n.fmt_pct(kpi.pct, 1) if kpi.pct is not None else ""]
            for col, text in enumerate(values):
                item = QTableWidgetItem(text)
                if col:
                    item.setTextAlignment(Qt.AlignCenter)
                self.kpi_table.setItem(row, col, item)
        self._render_chart()

    def _render_chart(self, *_args):
        if self._loading:
            return
        rows = [r for r in self.inspections
                if analytics.month_key(r.get("date")) in set(self.pivot.months)]
        charts.render_maintenance(self.canvas, self.chart_combo.currentData(),
                                  self.pivot, self.kpis, rows)

    # ==================================================================
    # Pestana 4: tags instalados por semana
    # ==================================================================
    def _tab_tags(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)

        title = QLabel(i18n.t("tags.title"))
        title.setObjectName("sectionTitle")
        lay.addWidget(title)
        hint = QLabel(i18n.t("tags.hint"))
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        buttons = QHBoxLayout()
        btn_folder = QPushButton(i18n.t("tags.btn_folder"))
        btn_folder.clicked.connect(self._on_load_tag_folder)
        btn_files = QPushButton(i18n.t("tags.btn_files"))
        btn_files.clicked.connect(self._on_load_tag_files)
        buttons.addWidget(btn_folder)
        buttons.addWidget(btn_files)
        self.chk_repair = QCheckBox(i18n.t("tags.repair"))
        self.chk_repair.setChecked(bool(settings.get("repair_dates")))
        self.chk_repair.setToolTip(i18n.t("tags.repair_tip"))
        self.chk_repair.toggled.connect(
            lambda v: settings.set_("repair_dates", bool(v)))
        buttons.addSpacing(12)
        buttons.addWidget(self.chk_repair)
        buttons.addStretch(1)
        btn_clear = QPushButton(i18n.t("tags.btn_clear"))
        btn_clear.setObjectName("danger")
        btn_clear.clicked.connect(self._on_clear_tags)
        buttons.addWidget(btn_clear)
        btn_export = QPushButton(i18n.t("tags.btn_export"))
        btn_export.setObjectName("accent")
        btn_export.clicked.connect(self._on_export_tags)
        buttons.addWidget(btn_export)
        lay.addLayout(buttons)

        cards = QHBoxLayout()
        cards.setSpacing(8)
        self.card_tags_total = _card(cards, "0", i18n.t("tags.card_total"),
                                     i18n.t("tags.card_total_hint"))
        self.card_tags_installed = _card(
            cards, "0", i18n.t("tags.card_installed"),
            i18n.t("tags.card_installed_hint"))
        self.card_tags_removed = _card(cards, "0", i18n.t("tags.card_removed"),
                                       i18n.t("tags.card_removed_hint"))
        self.card_tags_files = _card(cards, "0", i18n.t("tags.card_files"),
                                     i18n.t("tags.card_files_hint"))
        lay.addLayout(cards)

        filters = QHBoxLayout()
        filters.addWidget(QLabel(i18n.t("full.filter_year")))
        self.tag_year_combo = QComboBox()
        self.tag_year_combo.currentIndexChanged.connect(self._refresh_tag_view)
        filters.addWidget(self.tag_year_combo)
        filters.addWidget(QLabel(i18n.t("tags.filter_type")))
        self.tag_type_combo = QComboBox()
        self.tag_type_combo.currentIndexChanged.connect(self._refresh_tag_view)
        filters.addWidget(self.tag_type_combo)
        filters.addWidget(QLabel(i18n.t("tags.filter_device")))
        self.tag_device_combo = QComboBox()
        self.tag_device_combo.currentIndexChanged.connect(
            self._refresh_tag_view)
        filters.addWidget(self.tag_device_combo)
        filters.addWidget(QLabel(i18n.t("tags.filter_dept")))
        self.tag_dept_combo = QComboBox()
        self.tag_dept_combo.currentIndexChanged.connect(self._refresh_tag_view)
        filters.addWidget(self.tag_dept_combo)
        filters.addSpacing(14)
        filters.addWidget(QLabel(i18n.t("tags.chart")))
        self.tag_chart_combo = QComboBox()
        for key, label in ((charts.CHART_TAG_INSTALLED, "tags.chart_installed"),
                           (charts.CHART_TAG_TYPE, "tags.chart_type"),
                           (charts.CHART_TAG_DEPT, "tags.chart_dept")):
            self.tag_chart_combo.addItem(i18n.t(label), key)
        self.tag_chart_combo.currentIndexChanged.connect(self._render_tag_chart)
        filters.addWidget(self.tag_chart_combo)

        filters.addWidget(QLabel(i18n.t("grain.label")))
        self.grain_combo = QComboBox()
        for grain in analytics.GRAINS:
            self.grain_combo.addItem(i18n.t("grain." + grain), grain)
        index = self.grain_combo.findData(
            settings.get("tag_grain") or analytics.GRAIN_MONTH)
        self.grain_combo.setCurrentIndex(max(index, 0))
        self.grain_combo.currentIndexChanged.connect(self._on_grain_changed)
        filters.addWidget(self.grain_combo)
        filters.addStretch(1)
        self.lbl_tag_rows = QLabel()
        filters.addWidget(self.lbl_tag_rows)
        lay.addLayout(filters)

        splitter = QSplitter(Qt.Vertical)
        chart_box = QWidget()
        chart_layout = QVBoxLayout(chart_box)
        chart_layout.setContentsMargins(0, 0, 0, 0)
        self.tag_canvas = charts.ChartCanvas(height=3.8)
        chart_layout.addWidget(NavigationToolbar2QT(self.tag_canvas, chart_box))
        chart_layout.addWidget(self.tag_canvas, 1)
        splitter.addWidget(chart_box)

        self.tag_columns = [
            ("col.move_type", "move_type"), ("col.date", "date"),
            ("col.equipment_id", "equipment_id"), ("col.tag", "tag"),
            ("col.device", "device_type"),
            ("col.cost_center", "cost_center"),
            ("col.department", "department"), ("col.product", "product"),
            ("col.changed_by", "changed_by"), ("col.source", "source_file"),
            ("col.inferred", "type_inferred"), ("col.note", "note")]
        self.tag_table = QTableWidget(0, len(self.tag_columns))
        self.tag_table.setHorizontalHeaderLabels(
            [i18n.t(key) for key, _f in self.tag_columns])
        self.tag_table.setAlternatingRowColors(True)
        self.tag_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tag_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tag_table.horizontalHeader().setStretchLastSection(True)
        splitter.addWidget(self.tag_table)
        splitter.setSizes([420, 380])
        lay.addWidget(splitter, 1)
        return page

    def _populate_tag_filters(self):
        combos = (self.tag_year_combo, self.tag_type_combo,
                  self.tag_device_combo, self.tag_dept_combo)
        for combo in combos:
            combo.blockSignals(True)
        try:
            previous = [combo.currentData() for combo in combos]
            self.tag_year_combo.clear()
            self.tag_year_combo.addItem(i18n.t("full.all"), None)
            for year in sorted({r["date"][:4] for r in self.movements
                                if r.get("date")}, reverse=True):
                self.tag_year_combo.addItem(year, year)
            self.tag_type_combo.clear()
            self.tag_type_combo.addItem(i18n.t("full.all"), None)
            for move in tag_reader.MOVE_TYPES:
                if any(r.get("move_type") == move for r in self.movements):
                    self.tag_type_combo.addItem(i18n.tr_value(move), move)
            self.tag_device_combo.clear()
            self.tag_device_combo.addItem(i18n.t("full.all"), None)
            for device in (tag_reader.DEVICE_SMU, tag_reader.DEVICE_TAG):
                self.tag_device_combo.addItem(device, device)
            self.tag_dept_combo.clear()
            self.tag_dept_combo.addItem(i18n.t("full.all"), None)
            for dept in sorted({r["department"] for r in self.movements
                                if r.get("department")}):
                self.tag_dept_combo.addItem(dept, dept)
            for combo, value in zip(combos, previous):
                index = combo.findData(value)
                combo.setCurrentIndex(max(index, 0))
        finally:
            for combo in combos:
                combo.blockSignals(False)

    def _filtered_movements(self) -> list:
        year = self.tag_year_combo.currentData()
        move = self.tag_type_combo.currentData()
        device = self.tag_device_combo.currentData()
        dept = self.tag_dept_combo.currentData()
        rows = self.movements
        if year:
            rows = [r for r in rows if (r.get("date") or "").startswith(year)]
        if move:
            rows = [r for r in rows if r.get("move_type") == move]
        if device:
            rows = [r for r in rows if r.get("device_type") == device]
        if dept:
            rows = [r for r in rows if r.get("department") == dept]
        return rows

    def _refresh_tag_view(self, *_args):
        if self._loading:
            return
        rows = self._filtered_movements()
        totals = analytics.tag_totals(rows)
        self.card_tags_total.value_label.setText(
            i18n.fmt_number(totals["total"]))
        self.card_tags_installed.value_label.setText(
            i18n.fmt_number(totals["installed"]))
        self.card_tags_removed.value_label.setText(
            i18n.fmt_number(totals["removed"]))
        self.card_tags_files.value_label.setText(
            i18n.fmt_number(totals["files"]))

        shown = rows[:_TABLE_LIMIT]
        self.tag_table.setRowCount(0)
        self.tag_table.setRowCount(len(shown))
        for r, record in enumerate(shown):
            for c, (_key, field) in enumerate(self.tag_columns):
                if field == "date":
                    text = i18n.fmt_date(store.parse_date(record.get("date")))
                elif field == "move_type":
                    text = i18n.tr_value(record.get("move_type"))
                elif field == "type_inferred":
                    text = i18n.tr_value(
                        "Y" if record.get("type_inferred") else "N")
                elif field == "note":
                    text = i18n.tr_note(record.get("note"))
                else:
                    text = str(record.get(field) or "")
                self.tag_table.setItem(r, c, QTableWidgetItem(text))
        self.tag_table.resizeColumnsToContents()
        self.lbl_tag_rows.setText(i18n.t("tags.rows", shown=len(rows),
                                         total=len(self.movements)))
        self._render_tag_chart()

    def _on_grain_changed(self, *_args):
        if self._loading:
            return
        settings.set_("tag_grain", self.grain_combo.currentData())
        self._render_tag_chart()

    def _grain(self) -> str:
        return self.grain_combo.currentData() or analytics.GRAIN_MONTH

    def _render_tag_chart(self, *_args):
        if self._loading:
            return
        grain = self._grain()
        charts.render_tags(self.tag_canvas,
                           self.tag_chart_combo.currentData(),
                           self._filtered_movements(), grain,
                           _GRAIN_CHART_LIMIT.get(grain, 0))

    # ==================================================================
    # Recarga desde la base
    # ==================================================================
    def _refresh_views(self):
        """Vuelve a llenar filtros, tablas, tarjetas y graficas desde la cache."""
        self._populate_filters()
        self._populate_period_combo()
        self._populate_tag_filters()
        self._refresh_full_list()
        self._refresh_dashboard()
        self._refresh_tag_view()
        self._refresh_status()

    def _reload_inspections(self):
        self.inspections = store.inspections()
        self._refresh_views()

    def _reload_movements(self):
        self.movements = store.movements()
        self._refresh_views()

    def _on_refresh_store(self, *_args):
        self.inspections = store.inspections()
        self.movements = store.movements()
        self._refresh_views()

    # ==================================================================
    # Acciones: importar submissions
    # ==================================================================
    def _on_load_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self, i18n.t("dlg.open_source"),
            os.path.dirname(self.source_path or ""),
            i18n.t("dlg.excel_filter"))
        if not path:
            return
        try:
            subs = _run_with_progress(
                self, i18n.t("prog.reading_title"),
                i18n.t("prog.reading", file=os.path.basename(path)),
                source_reader.read_submissions, path)
        except Exception as exc:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 i18n.t("msg.read_error", err=exc))
            return
        if not subs:
            QMessageBox.warning(self, i18n.t("dlg.warning"),
                                i18n.t("msg.no_data"))
            return

        self.source_path = path
        settings.set_("last_source_file", path)
        self.submissions = subs
        self.preview_rows = mapping.submissions_to_rows(subs)
        self.preview_codes = [s.get(mapping.H_CODE) or "" for s in subs]
        self.preview_checked = [True] * len(self.preview_rows)
        self.preview_state = [_STATE_NEW] * len(self.preview_rows)
        self._populate_preview()
        self._refresh_status()
        self.statusBar().showMessage(
            i18n.t("msg.loaded", n=len(subs), file=os.path.basename(path)))
        # El export del formulario es acumulativo: cada descarga trae otra vez
        # todo lo anterior. Marcar lo repetido aqui es lo que evita que el
        # maestro reciba dos veces la misma inspeccion.
        self._mark_duplicates()

    def _on_pick_target(self):
        path, _ = QFileDialog.getOpenFileName(
            self, i18n.t("dlg.open_target"),
            os.path.dirname(self.target_path or ""),
            i18n.t("dlg.excel_filter"))
        if not path:
            return
        self.target_path = path
        settings.set_("last_target_file", path)
        self._refresh_status()
        self.statusBar().showMessage(
            i18n.t("msg.target_set", file=os.path.basename(path)))
        if self.preview_rows:
            # Cambiar de maestro cambia que esta duplicado y que no.
            self._mark_duplicates()

    def _on_import_history(self):
        if not self.target_path:
            QMessageBox.information(self, i18n.t("dlg.info"),
                                    i18n.t("msg.pick_target_first"))
            return
        try:
            rows = _run_with_progress(
                self, i18n.t("prog.history_title"),
                i18n.t("prog.history", sheet=source_reader.FULL_LIST_SHEET),
                source_reader.read_full_list, self.target_path)
        except Exception as exc:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 i18n.t("msg.read_error", err=exc))
            return

        records = [store.inspection_from_row(row, "", self.target_path)
                   for row in rows]
        result = _run_with_progress(
            self, i18n.t("prog.history_title"),
            i18n.t("prog.history", sheet=source_reader.FULL_LIST_SHEET),
            store.add_inspections, records)
        self._reload_inspections()
        if self.preview_rows:
            self._mark_duplicates(announce=False)
        QMessageBox.information(
            self, i18n.t("dlg.info"),
            i18n.t("msg.history_loaded", read=len(rows),
                   sheet=source_reader.FULL_LIST_SHEET,
                   added=result["added"], skipped=result["skipped"]))

    def _on_store_only(self):
        records = self._selected_records()
        if not records:
            QMessageBox.information(self, i18n.t("dlg.info"),
                                    i18n.t("msg.nothing_checked"))
            return
        result = store.add_inspections(records)
        self._reload_inspections()
        self._mark_duplicates(announce=False)
        QMessageBox.information(
            self, i18n.t("dlg.info"),
            i18n.t("msg.stored", added=result["added"],
                   skipped=result["skipped"]))

    def _on_append(self):
        rows = self._selected_rows()
        if not rows or not self.target_path:
            QMessageBox.information(self, i18n.t("dlg.info"),
                                    i18n.t("msg.nothing_checked"))
            return

        # El maestro no deduplica: si el usuario vuelve a marcar filas que ya
        # estan cargadas, se le dice antes de escribir y no despues.
        dupes = sum(1 for keep, state in zip(self.preview_checked,
                                             self.preview_state)
                    if keep and state != _STATE_NEW)
        if dupes:
            answer = QMessageBox.question(
                self, i18n.t("dlg.confirm"),
                i18n.t("import.confirm_dupes", n=len(rows), dupes=dupes),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                return

        confirm = QMessageBox.question(
            self, i18n.t("dlg.confirm"),
            i18n.t("msg.confirm_append", n=len(rows),
                   sheet=excel_writer.SHEET_NAME,
                   file=os.path.basename(self.target_path)),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if confirm != QMessageBox.Yes:
            return
        try:
            result = _run_with_progress(
                self, i18n.t("prog.writing_title"),
                i18n.t("prog.writing", sheet=excel_writer.SHEET_NAME),
                excel_writer.append_rows, self.target_path, rows,
                self.chk_backup.isChecked())
        except PermissionError:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 i18n.t("msg.file_in_use"))
            return
        except Exception as exc:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 "%s\n\n%s" % (exc, traceback.format_exc()))
            return

        message = i18n.t("msg.append_done", n=result["written"],
                         sheet=result["sheet"], first=result["first_row"],
                         last=result["last_row"])
        if result["backup"]:
            message += "\n\n" + i18n.t("msg.backup_made",
                                       path=result["backup"])
        if self.chk_store.isChecked():
            stored = store.add_inspections(self._selected_records())
            self._reload_inspections()
            message += "\n\n" + i18n.t("msg.stored", added=stored["added"],
                                       skipped=stored["skipped"])
        # Lo recien escrito ya es 'viejo' para la proxima pasada.
        self._mark_duplicates(announce=False)
        message += "\n\n" + i18n.t("msg.refresh_hint")
        QMessageBox.information(self, i18n.t("dlg.info"), message)
        self.statusBar().showMessage(
            i18n.t("msg.append_done", n=result["written"],
                   sheet=result["sheet"], first=result["first_row"],
                   last=result["last_row"]).replace("\n", " "))

    # ==================================================================
    # Acciones: Full List almacenado
    # ==================================================================
    def _on_delete_rows(self):
        ids = {self.full_table.item(index.row(), 0).data(Qt.UserRole)
               for index in self.full_table.selectionModel().selectedRows()}
        ids = {i for i in ids if i is not None}
        if not ids:
            QMessageBox.information(self, i18n.t("dlg.info"),
                                    i18n.t("msg.no_selection"))
            return
        confirm = QMessageBox.question(
            self, i18n.t("dlg.confirm"),
            i18n.t("msg.confirm_delete", n=len(ids)),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm != QMessageBox.Yes:
            return
        deleted = store.delete_inspections(list(ids))
        self._reload_inspections()
        QMessageBox.information(self, i18n.t("dlg.info"),
                                i18n.t("msg.deleted", n=deleted))

    def _on_export_report(self):
        if not self.inspections:
            QMessageBox.information(self, i18n.t("dlg.info"),
                                    i18n.t("msg.export_empty"))
            return
        path = self._ask_save_path(i18n.t("dlg.save_report"), "report")
        if not path:
            return
        rows = self._filtered_inspections()
        try:
            _run_with_progress(
                self, i18n.t("prog.export_title"), i18n.t("prog.export"),
                report_export.export_maintenance, path, rows,
                settings.fleet_sizes(), self._export_window(rows))
        except PermissionError:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 i18n.t("msg.file_in_use"))
            return
        except Exception as exc:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 "%s\n\n%s" % (exc, traceback.format_exc()))
            return
        QMessageBox.information(self, i18n.t("dlg.info"),
                                i18n.t("msg.export_done", path=path))

    def _on_fleet_sizes(self):
        months = analytics.available_months(self.inspections)
        if not months:
            months = sorted(settings.fleet_sizes())
        dialog = FleetSizeDialog(self, months, settings.fleet_sizes())
        if dialog.exec() == QDialog.Accepted:
            # Los meses fuera de la ventana editada se conservan: el dialogo
            # solo muestra los que tienen datos, y no puede borrar el resto.
            sizes = dict(settings.fleet_sizes())
            sizes.update(dialog.sizes())
            for month in months:
                if month not in dialog.sizes():
                    sizes.pop(month, None)
            settings.set_fleet_sizes(sizes)
            self._refresh_dashboard()

    # ==================================================================
    # Acciones: tags
    # ==================================================================
    def _on_load_tag_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, i18n.t("dlg.open_tag_folder"),
            settings.get("last_tag_folder") or "")
        if not folder:
            return
        settings.set_("last_tag_folder", folder)
        self._consolidate(tag_reader.find_files(folder))

    def _on_load_tag_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, i18n.t("dlg.open_tag_files"),
            settings.get("last_tag_folder") or "", i18n.t("dlg.excel_filter"))
        if paths:
            self._consolidate(paths)

    def _consolidate(self, paths):
        if not paths:
            QMessageBox.warning(self, i18n.t("dlg.warning"),
                                i18n.t("msg.tags_none"))
            return
        result = _run_with_progress(
            self, i18n.t("prog.tags_title"), i18n.t("prog.tags"),
            tag_reader.read_paths, paths,
            repair=self.chk_repair.isChecked())
        if not result["records"]:
            QMessageBox.warning(self, i18n.t("dlg.warning"),
                                i18n.t("msg.tags_none"))
            return
        stored = store.add_movements(result["records"])
        self._reload_movements()

        message = i18n.t("msg.tags_stored", files=len(result["files"]),
                         added=stored["added"], skipped=stored["skipped"])
        if result["repaired"]:
            message += "\n\n" + i18n.t("msg.tags_repaired",
                                       n=result["repaired"])
        if result["suspect"]:
            message += "\n\n" + i18n.t("msg.tags_suspect", n=result["suspect"])
        if result["errors"]:
            message += "\n\n" + "\n".join(
                "%s: %s" % (name, err) for name, err in result["errors"][:8])
        QMessageBox.information(self, i18n.t("dlg.info"), message)

    def _on_clear_tags(self):
        if not self.movements:
            return
        confirm = QMessageBox.question(
            self, i18n.t("dlg.confirm"),
            i18n.t("msg.confirm_clear_tags", n=len(self.movements)),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm != QMessageBox.Yes:
            return
        deleted = store.clear_movements()
        self._reload_movements()
        QMessageBox.information(self, i18n.t("dlg.info"),
                                i18n.t("msg.deleted", n=deleted))

    def _on_export_tags(self):
        if not self.movements:
            QMessageBox.information(self, i18n.t("dlg.info"),
                                    i18n.t("msg.export_empty"))
            return
        path = self._ask_save_path(i18n.t("dlg.save_tags"), "tags")
        if not path:
            return
        try:
            _run_with_progress(
                self, i18n.t("prog.export_title"), i18n.t("prog.export"),
                report_export.export_tags, path, self._filtered_movements(),
                self._grain())
        except PermissionError:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 i18n.t("msg.file_in_use"))
            return
        except Exception as exc:
            QMessageBox.critical(self, i18n.t("dlg.error"),
                                 "%s\n\n%s" % (exc, traceback.format_exc()))
            return
        QMessageBox.information(self, i18n.t("dlg.info"),
                                i18n.t("msg.export_done", path=path))

    # ------------------------------------------------------------------
    def _ask_save_path(self, title: str, kind: str) -> str:
        suggested = report_export.default_path(
            settings.get("last_export_dir") or "", kind)
        path, _ = QFileDialog.getSaveFileName(self, title, suggested,
                                              i18n.t("dlg.excel_filter"))
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            settings.set_("last_export_dir", os.path.dirname(path))
        return path


def launch() -> int:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication.instance() or QApplication(sys.argv)
    settings.load()
    window = MainWindow()
    geometry = settings.get("window") or {}
    window.resize(int(geometry.get("w", 1420)), int(geometry.get("h", 900)))
    if geometry.get("maximized"):
        window.showMaximized()
    else:
        window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(launch())
