# -*- coding: utf-8 -*-
"""
Interfaz grafica (PySide6) del cargador de mantenimiento de flota.

Flujo:
  1. 'Cargar Excel de submissions...'  -> lee el export del formulario
     (merian_ops - Form - Fleet maintenance - Submissions).
  2. 'Seleccionar Excel destino...'    -> el maestro
     'Fleet Tag Inventory and Maintenance.xlsx'.
  3. Revisar / corregir las filas en la tabla de previsualizacion. Cada fila
     trae una casilla 'Incluir' para decidir cuales se cargan.
  4. 'CARGAR AL EXCEL DESTINO'         -> agrega las filas a la hoja
     'Full List 2024-2025' conservando estilos, formulas y la tabla Table3.
     Se hace un respaldo automatico del destino antes de escribir.

Ejecutar:  python run.py
"""
from __future__ import annotations

import os
import sys
import traceback
import datetime

from PySide6.QtCore import Qt, QThread
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QFrame, QGroupBox, QHBoxLayout,
    QHeaderView, QLabel, QMainWindow, QMessageBox, QProgressDialog,
    QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

import mapping
import source_reader
import excel_writer

PRIMARY = "#1F4E78"
ACCENT = "#2E7D32"
DANGER = "#C62828"
BG = "#F4F6F9"

STYLESHEET = f"""
QMainWindow, QWidget {{ background: {BG}; color: #1A1A1A; }}
QGroupBox {{
    font-weight: bold; color: {PRIMARY};
    border: 1px solid #C9D3DF; border-radius: 8px;
    margin-top: 14px; padding: 10px;
}}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px; }}
QLabel {{ color: #1A1A1A; }}
QPushButton {{
    background: {PRIMARY}; color: white; border: none;
    border-radius: 6px; padding: 7px 14px; font-weight: bold;
}}
QPushButton:hover {{ background: #2A5F92; }}
QPushButton:disabled {{ background: #9AA8B8; color: white; }}
QPushButton#accent {{ background: {ACCENT}; }}
QPushButton#accent:hover {{ background: #388E3C; }}
QTableWidget {{ background: white; gridline-color: #DCE3EB; color: #1A1A1A; }}
QTableWidget QTableCornerButton::section {{ background: {PRIMARY}; }}
QHeaderView::section {{
    background: {PRIMARY}; color: white; padding: 6px; border: none;
    font-weight: bold;
}}
QLabel#title {{ font-size: 16px; font-weight: bold; color: {PRIMARY}; }}
"""

# Columnas de datos que se muestran/editan en la previsualizacion (B, C, E..S).
PREVIEW_COLUMNS = mapping.DATA_COLUMNS


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
    dlg.setMinimumWidth(420)

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
    if isinstance(value, datetime.datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, datetime.date):
        return value.strftime("%d/%m/%Y")
    return str(value)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(
            "Cargador de Mantenimiento de Flota  -  Newmont Merian")
        self.resize(1280, 760)
        self.setStyleSheet(STYLESHEET)

        self.source_path = ""
        self.target_path = ""
        self.submissions = []   # list[dict] crudas del formulario

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.addWidget(self._build_controls())
        layout.addWidget(self._build_preview(), stretch=1)
        layout.addWidget(self._build_footer())

        self.statusBar().showMessage(
            "Cargue el Excel de submissions y seleccione el Excel destino.")
        self._refresh_status()

    # ------------------------------------------------------------------
    # Construccion de la interfaz
    # ------------------------------------------------------------------
    def _build_controls(self) -> QWidget:
        box = QGroupBox("Archivos")
        outer = QVBoxLayout(box)

        row = QHBoxLayout()
        btn_src = QPushButton("Cargar Excel de submissions...")
        btn_src.clicked.connect(self._on_load_source)
        btn_tgt = QPushButton("Seleccionar Excel destino...")
        btn_tgt.clicked.connect(self._on_pick_target)
        row.addWidget(btn_src)
        row.addWidget(btn_tgt)
        row.addStretch(1)
        outer.addLayout(row)

        status_row = QHBoxLayout()
        status_row.setSpacing(8)
        lbl = QLabel("Estado:")
        lbl.setStyleSheet("QLabel { font-weight: bold; color: #1F4E78; }")
        status_row.addWidget(lbl)
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
        box = QGroupBox("Previsualizacion  (filas que se agregaran a "
                        "'Full List 2024-2025')")
        lay = QVBoxLayout(box)
        lay.addWidget(QLabel(
            "Revise y corrija los valores antes de cargar. Las columnas "
            "'Date' (A) y 'Verified' (D) del destino son formulas y se "
            "calculan solas; aqui no se muestran."))

        headers = ["Incluir"] + [
            "%s · %s" % (c, mapping.TARGET_HEADERS[c]) for c in PREVIEW_COLUMNS]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.verticalHeader().setVisible(True)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.table)

        sel = QHBoxLayout()
        btn_all = QPushButton("Marcar todas")
        btn_all.clicked.connect(lambda: self._set_all_checks(True))
        btn_none = QPushButton("Desmarcar todas")
        btn_none.clicked.connect(lambda: self._set_all_checks(False))
        sel.addWidget(btn_all)
        sel.addWidget(btn_none)
        sel.addStretch(1)
        self.lbl_count = QLabel("0 submissions cargadas.")
        sel.addWidget(self.lbl_count)
        lay.addLayout(sel)
        return box

    def _build_footer(self) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        self.chk_backup = QCheckBox("Crear respaldo del destino antes de "
                                    "escribir")
        self.chk_backup.setChecked(True)
        lay.addWidget(self.chk_backup)
        lay.addStretch(1)
        self.btn_load = QPushButton("CARGAR AL EXCEL DESTINO")
        self.btn_load.setObjectName("accent")
        self.btn_load.setMinimumWidth(260)
        f = QFont()
        f.setBold(True)
        self.btn_load.setFont(f)
        self.btn_load.clicked.connect(self._on_append)
        lay.addWidget(self.btn_load)
        return w

    # ------------------------------------------------------------------
    # Estado / chips
    # ------------------------------------------------------------------
    def _set_chip(self, chip, loaded, title, detail=""):
        if loaded:
            chip.setStyleSheet(
                "QLabel { background: #E6F4EA; color: #1B5E20; "
                "border: 1px solid #2E7D32; border-radius: 12px; "
                "padding: 3px 12px; }")
            html = "<b>&#10003; %s</b>" % title
            if detail:
                html += " <span style='color:#33691E;'>(%s)</span>" % detail
            chip.setText(html)
        else:
            chip.setStyleSheet(
                "QLabel { background: #FFF4E5; color: #8B4500; "
                "border: 1px solid #FB8C00; border-radius: 12px; "
                "padding: 3px 12px; }")
            chip.setText("<b>&#9888; %s</b> sin seleccionar" % title)

    def _refresh_status(self):
        if self.source_path and self.submissions:
            self._set_chip(self.chip_source, True, "Submissions",
                           "%s &mdash; %d filas" % (
                               os.path.basename(self.source_path),
                               len(self.submissions)))
        else:
            self._set_chip(self.chip_source, False, "Submissions")
        if self.target_path:
            self._set_chip(self.chip_target, True, "Excel destino",
                           os.path.basename(self.target_path))
        else:
            self._set_chip(self.chip_target, False, "Excel destino")
        self.btn_load.setEnabled(bool(self.submissions and self.target_path))

    # ------------------------------------------------------------------
    # Tabla de previsualizacion
    # ------------------------------------------------------------------
    def _populate_table(self):
        rows = mapping.submissions_to_rows(self.submissions)
        self.table.setRowCount(0)
        for rowdata in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            chk = QCheckBox()
            chk.setChecked(True)
            holder = QWidget()
            hl = QHBoxLayout(holder)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.setAlignment(Qt.AlignCenter)
            hl.addWidget(chk)
            self.table.setCellWidget(r, 0, holder)
            for ci, col in enumerate(PREVIEW_COLUMNS, start=1):
                self.table.setItem(
                    r, ci, QTableWidgetItem(_fmt_cell(rowdata.get(col))))
        self.table.resizeColumnsToContents()
        self.lbl_count.setText("%d submissions cargadas." % len(rows))

    def _set_all_checks(self, checked: bool):
        for r in range(self.table.rowCount()):
            chk = self._row_checkbox(r)
            if chk:
                chk.setChecked(checked)

    def _row_checkbox(self, r):
        holder = self.table.cellWidget(r, 0)
        return holder.findChild(QCheckBox) if holder else None

    def _collect_selected_rows(self) -> list:
        """Lee la tabla y devuelve las filas marcadas como dicts {col: valor}.

        Re-mapea cada submission (para conservar tipos como fecha) y luego
        aplica las ediciones de texto que el usuario haya hecho en la tabla.
        """
        base_rows = mapping.submissions_to_rows(self.submissions)
        selected = []
        for r in range(self.table.rowCount()):
            chk = self._row_checkbox(r)
            if not chk or not chk.isChecked():
                continue
            rowdata = dict(base_rows[r])
            for ci, col in enumerate(PREVIEW_COLUMNS, start=1):
                item = self.table.item(r, ci)
                text = item.text().strip() if item else ""
                rowdata[col] = self._reconcile_value(col, rowdata.get(col), text)
            selected.append(rowdata)
        return selected

    @staticmethod
    def _reconcile_value(col, original, text):
        """Si el usuario no edito la celda (mismo texto que el original), se
        conserva el valor original (con su tipo). Si la edito, se usa el texto;
        para la fecha (col B) se intenta re-parsear."""
        if _fmt_cell(original) == text:
            return original
        if text == "":
            return None
        if col == "B":
            dt = mapping.coerce_datetime(text)
            if dt is not None:
                return datetime.datetime(dt.year, dt.month, dt.day)
        return text

    # ------------------------------------------------------------------
    # Acciones
    # ------------------------------------------------------------------
    def _on_load_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Excel de submissions del formulario", "",
            "Excel (*.xlsx)")
        if not path:
            return
        try:
            subs = _run_with_progress(
                self, "Cargando submissions",
                "Leyendo %s ..." % os.path.basename(path),
                source_reader.read_submissions, path)
        except Exception as exc:
            QMessageBox.critical(
                self, "Error",
                "No se pudo leer el Excel de submissions:\n%s" % exc)
            return
        if not subs:
            QMessageBox.warning(
                self, "Sin datos",
                "El archivo no contiene submissions reconocibles en la hoja "
                "'Form Submissions'.")
            return
        self.source_path = path
        self.submissions = subs
        self._populate_table()
        self._refresh_status()
        self.statusBar().showMessage(
            "%d submissions cargadas desde %s" % (
                len(subs), os.path.basename(path)))

    def _on_pick_target(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Excel destino "
            "(Fleet Tag Inventory and Maintenance)", "", "Excel (*.xlsx)")
        if not path:
            return
        self.target_path = path
        self._refresh_status()
        self.statusBar().showMessage(
            "Destino: %s" % os.path.basename(path))

    def _on_append(self):
        if not self.submissions or not self.target_path:
            return
        rows = self._collect_selected_rows()
        if not rows:
            QMessageBox.information(
                self, "Nada que cargar",
                "No hay filas marcadas para cargar.")
            return
        confirm = QMessageBox.question(
            self, "Confirmar carga",
            "Se agregaran %d fila(s) a la hoja '%s' de:\n%s\n\n"
            "Se hara un respaldo automatico del destino.\n\n¿Continuar?" % (
                len(rows), excel_writer.SHEET_NAME,
                os.path.basename(self.target_path)),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if confirm != QMessageBox.Yes:
            return
        try:
            result = _run_with_progress(
                self, "Cargando al Excel destino",
                "Escribiendo filas en '%s'..." % excel_writer.SHEET_NAME,
                excel_writer.append_rows, self.target_path, rows,
                self.chk_backup.isChecked())
        except PermissionError:
            QMessageBox.critical(
                self, "Archivo en uso",
                "No se pudo escribir el Excel destino. Es probable que este "
                "abierto en Excel. Cierrelo y vuelva a intentar.")
            return
        except Exception as exc:
            QMessageBox.critical(
                self, "Error al cargar",
                "%s\n\n%s" % (exc, traceback.format_exc()))
            return

        msg = ("Se agregaron %d fila(s) a '%s'.\n"
               "Filas %s a %s.\n\n" % (
                   result["written"], result["sheet"],
                   result["first_row"], result["last_row"]))
        if result["backup"]:
            msg += "Respaldo creado:\n%s\n\n" % result["backup"]
        msg += ("Abra el Excel y use 'Datos > Actualizar todo' para refrescar "
                "los pivotes y las hojas de resumen.")
        QMessageBox.information(self, "Carga completada", msg)
        self.statusBar().showMessage(
            "Carga completada: %d fila(s) agregadas." % result["written"])


def launch() -> int:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication.instance() or QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(launch())
