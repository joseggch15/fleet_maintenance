# -*- coding: utf-8 -*-
"""
Temas claro y oscuro: una sola paleta para Qt y para matplotlib.

Toda la apariencia sale de aqui. Los modulos de interfaz, de graficas y de
exportacion leen la paleta ACTIVA (`palette()`) en vez de tener colores
propios, asi que cambiar de tema es cambiar un diccionario y volver a pedir la
hoja de estilos — no hay colores sueltos repartidos por el codigo.

El tema oscuro no es el claro con los grises invertidos: los acentos se aclaran
(el azul corporativo #1F4E78 sobre fondo oscuro es ilegible) y los colores
semanticos se ajustan para mantener contraste sin perder significado.

Tres pares de colores mandan en este software y conviene que no cambien de
sentido nunca:

  `smu` / `tag`        los dos tipos de dispositivo que se instalan en flota;
  `reviewed` / `pct`   las barras de equipos revisados y la linea de % de
                       inspeccion del dashboard (barras y linea comparten
                       grafica, por eso son colores contrastantes, no vecinos);
  `install` / `removal` alta y baja de un tag en el consolidado semanal.

Se eligio azul/naranja para el par principal porque son distinguibles en las
dos formas comunes de daltonismo, a diferencia del par verde/rojo.
"""
from __future__ import annotations

LIGHT = "light"
DARK = "dark"
THEMES = (LIGHT, DARK)

_PALETTES: dict[str, dict[str, str]] = {
    LIGHT: {
        "name": LIGHT,
        "bg":            "#F4F6F9",
        "surface":       "#FFFFFF",
        "surface_alt":   "#EDF1F6",
        "border":        "#C9D3DF",
        "grid":          "#DCE3EB",
        "text":          "#1A1A1A",
        "text_muted":    "#5A6B7D",
        "primary":       "#1F4E78",
        "primary_hover": "#2A5F92",
        "on_primary":    "#FFFFFF",
        "accent":        "#2E7D32",
        "accent_hover":  "#388E3C",
        "danger":        "#C62828",
        "warning":       "#ED7D31",
        "info":          "#00A0A0",
        "selection":     "#D0E4F7",
        "disabled":      "#9AA8B8",

        # -- dispositivos ---------------------------------------------------
        "smu":           "#1F4E78",
        "tag":           "#ED7D31",

        # -- dashboard de mantenimiento -------------------------------------
        "reviewed":      "#1F6FA8",   # barras: equipos revisados en el mes
        "pct":           "#ED7D31",   # linea: % de inspeccion sobre la flota

        # -- movimientos del consolidado semanal ----------------------------
        "install":       "#2E7D32",
        "replacement":   "#4472C4",
        "removal":       "#C62828",
        "updated":       "#B08A00",

        # -- graficas -------------------------------------------------------
        "chart_bg":      "#FFFFFF",
        "chart_grid":    "#D9D9D9",
        "chart_text":    "#404040",
        "chart_muted":   "#7F7F7F",
        "zero_line":     "#404040",
    },
    DARK: {
        "name": DARK,
        "bg":            "#161A20",
        "surface":       "#1E242C",
        "surface_alt":   "#252C36",
        "border":        "#39424F",
        "grid":          "#333B46",
        "text":          "#E8ECF1",
        "text_muted":    "#98A5B5",
        "primary":       "#4A90D9",
        "primary_hover": "#5FA3E8",
        "on_primary":    "#0F1319",
        "accent":        "#4CAF50",
        "accent_hover":  "#66BB6A",
        "danger":        "#EF5350",
        "warning":       "#FFA726",
        "info":          "#26C6DA",
        "selection":     "#2C3E52",
        "disabled":      "#4A5563",

        "smu":           "#64B5F6",
        "tag":           "#FFA726",

        "reviewed":      "#5FA3E8",
        "pct":           "#FFA726",

        "install":       "#66BB6A",
        "replacement":   "#64B5F6",
        "removal":       "#EF5350",
        "updated":       "#D4B106",

        "chart_bg":      "#1E242C",
        "chart_grid":    "#333B46",
        "chart_text":    "#C7D0DB",
        "chart_muted":   "#98A5B5",
        "zero_line":     "#C7D0DB",
    },
}

_current = LIGHT


def set_theme(name: str) -> None:
    global _current
    _current = name if name in _PALETTES else LIGHT


def current() -> str:
    return _current


def toggled() -> str:
    """El otro tema. Util para el boton que alterna claro/oscuro."""
    return DARK if _current == LIGHT else LIGHT


def palette(name: str | None = None) -> dict[str, str]:
    return _PALETTES.get(name or _current, _PALETTES[LIGHT])


def color(key: str, default: str = "#808080") -> str:
    return palette().get(key, default)


def device_color(device: str) -> str:
    """Color de 'SMU' o 'TAG'."""
    return palette().get("smu" if str(device).upper() == "SMU" else "tag")


# Los tipos de movimiento son etiquetas canonicas en ingles (asi vienen en los
# archivos semanales), por eso el mapa se indexa en ingles aunque la ventana
# este en espanol.
_MOVE_COLORS = {
    "NEW INSTALLATION": "install",
    "REPLACEMENT": "replacement",
    "REMOVAL": "removal",
    "TAG UPDATED": "updated",
}


def move_color(move_type: str) -> str:
    return palette().get(_MOVE_COLORS.get(str(move_type).upper(), "text_muted"),
                         "#9E9E9E")


def series_colors(n: int) -> list[str]:
    """Paleta ciclica para series categoricas (departamentos, tipos de equipo).

    Se pide por indice para que dos graficas del mismo dato usen el mismo
    color.
    """
    p = palette()
    base = [p["smu"], p["tag"], p["accent"], p["info"], p["primary"],
            p["updated"], p["danger"], p["replacement"]]
    return [base[i % len(base)] for i in range(max(n, 0))]


# Colores de las graficas nativas de Excel: openpyxl los quiere en hex sin '#'.
# El Excel exportado se imprime y se comparte, asi que siempre sale con la
# paleta CLARA aunque la ventana este en oscuro.
def excel_color(key: str) -> str:
    return _PALETTES[LIGHT].get(key, "808080").lstrip("#")


# ---------------------------------------------------------------------------
# Hoja de estilos de Qt
# ---------------------------------------------------------------------------

def stylesheet(name: str | None = None) -> str:
    """Genera la hoja de estilos completa para el tema pedido.

    Notas de accesibilidad y de defectos ya corregidos:
      - `QHeaderView::section` lleva borde derecho: sin el, en tema oscuro las
        columnas del encabezado se funden en una sola barra.
      - Los `QScrollBar` se estilizan a mano; los nativos de Windows se ven
        blancos sobre el tema oscuro.
      - `QTableWidget::item` tiene padding: sin el, los numeros alineados a la
        derecha quedan pegados a la linea de la grilla y se cortan.
    """
    p = palette(name)
    return f"""
QMainWindow, QWidget {{ background: {p['bg']}; color: {p['text']}; }}
QGroupBox {{
    font-weight: bold; color: {p['primary']};
    border: 1px solid {p['border']}; border-radius: 8px;
    margin-top: 15px; padding: 10px 8px 8px 8px;
    background: {p['surface']};
}}
QGroupBox::title {{
    subcontrol-origin: margin; subcontrol-position: top left;
    left: 10px; padding: 0 5px; background: transparent;
}}
QLabel {{ color: {p['text']}; background: transparent; }}
QLabel#title {{ font-size: 17px; font-weight: bold; color: {p['primary']}; }}
QLabel#subtitle {{ font-size: 11px; color: {p['text_muted']}; }}
QLabel#hint {{ font-size: 10px; color: {p['text_muted']}; font-weight: normal; }}
QLabel#sectionTitle {{
    font-size: 13px; font-weight: bold; color: {p['primary']};
}}

QPushButton {{
    background: {p['primary']}; color: {p['on_primary']}; border: none;
    border-radius: 6px; padding: 7px 13px; font-weight: bold;
}}
QPushButton:hover {{ background: {p['primary_hover']}; }}
QPushButton:disabled {{ background: {p['disabled']}; color: {p['bg']}; }}
QPushButton#accent {{ background: {p['accent']}; color: #FFFFFF; }}
QPushButton#accent:hover {{ background: {p['accent_hover']}; }}
QPushButton#danger {{ background: {p['danger']}; color: #FFFFFF; }}
QPushButton#ghost {{
    background: {p['surface']}; color: {p['primary']};
    border: 1px solid {p['border']}; padding: 5px 9px; font-weight: normal;
}}
QPushButton#ghost:hover {{ background: {p['selection']}; }}
QPushButton#ghost:checked {{
    background: {p['primary']}; color: {p['on_primary']}; font-weight: bold;
}}

QTabWidget::pane {{
    border: 1px solid {p['border']}; border-radius: 6px;
    background: {p['surface']}; top: -1px;
}}
QTabBar::tab {{
    background: {p['surface_alt']}; color: {p['text_muted']};
    padding: 8px 14px; margin-right: 2px;
    border: 1px solid {p['border']}; border-bottom: none;
    border-top-left-radius: 6px; border-top-right-radius: 6px;
}}
QTabBar::tab:selected {{
    background: {p['surface']}; color: {p['primary']}; font-weight: bold;
}}
QTabBar::tab:hover:!selected {{ color: {p['text']}; }}

QTableWidget {{
    background: {p['surface']}; alternate-background-color: {p['surface_alt']};
    gridline-color: {p['grid']}; color: {p['text']};
    border: 1px solid {p['border']}; border-radius: 4px;
    selection-background-color: {p['selection']}; selection-color: {p['text']};
}}
QTableWidget::item {{ padding: 3px 6px; }}
QTableWidget::item:selected {{
    background: {p['selection']}; color: {p['text']};
}}
QTableWidget QTableCornerButton::section {{
    background: {p['primary']}; border: none;
}}
QHeaderView {{ background: {p['primary']}; }}
QHeaderView::section {{
    background: {p['primary']}; color: {p['on_primary']};
    padding: 6px 8px; border: none;
    border-right: 1px solid {p['bg']};
    font-weight: bold;
}}
QHeaderView::section:last {{ border-right: none; }}

QComboBox, QLineEdit, QDateEdit, QDoubleSpinBox, QSpinBox {{
    background: {p['surface']}; color: {p['text']};
    border: 1px solid {p['border']}; border-radius: 5px; padding: 4px 6px;
    selection-background-color: {p['selection']}; selection-color: {p['text']};
}}
QComboBox:hover, QLineEdit:hover, QSpinBox:hover {{ border-color: {p['primary']}; }}
QComboBox::drop-down {{ border: none; width: 18px; }}
QComboBox QAbstractItemView {{
    background: {p['surface']}; color: {p['text']};
    border: 1px solid {p['border']};
    selection-background-color: {p['selection']}; selection-color: {p['text']};
    outline: none;
}}

QCheckBox {{ color: {p['text']}; spacing: 6px; }}
QCheckBox::indicator {{
    width: 15px; height: 15px; border-radius: 3px;
    border: 1px solid {p['border']}; background: {p['surface']};
}}
QCheckBox::indicator:checked {{
    background: {p['primary']}; border-color: {p['primary']};
}}

QFrame#card {{
    background: {p['surface']}; border: 1px solid {p['border']};
    border-radius: 8px;
}}
QLabel#cardValue {{ font-size: 19px; font-weight: bold; color: {p['primary']}; }}
QLabel#cardLabel {{ font-size: 10px; color: {p['text_muted']}; }}
QLabel#cardHint {{ font-size: 9px; color: {p['text_muted']}; }}

QLabel#chipOk {{
    background: {p['selection']}; color: {p['accent']};
    border: 1px solid {p['accent']}; border-radius: 12px; padding: 3px 12px;
}}
QLabel#chipWarn {{
    background: {p['surface_alt']}; color: {p['warning']};
    border: 1px solid {p['warning']}; border-radius: 12px; padding: 3px 12px;
}}

QScrollArea {{ background: transparent; border: none; }}
QScrollBar:vertical {{
    background: {p['bg']}; width: 11px; margin: 0; border-radius: 5px;
}}
QScrollBar::handle:vertical {{
    background: {p['disabled']}; min-height: 28px; border-radius: 5px;
}}
QScrollBar::handle:vertical:hover {{ background: {p['text_muted']}; }}
QScrollBar:horizontal {{
    background: {p['bg']}; height: 11px; margin: 0; border-radius: 5px;
}}
QScrollBar::handle:horizontal {{
    background: {p['disabled']}; min-width: 28px; border-radius: 5px;
}}
QScrollBar::handle:horizontal:hover {{ background: {p['text_muted']}; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

QSplitter::handle {{ background: {p['border']}; }}
QSplitter::handle:horizontal {{ width: 3px; }}
QSplitter::handle:vertical {{ height: 3px; }}

QStatusBar {{ background: {p['surface']}; color: {p['text_muted']}; }}
QToolTip {{
    background: {p['surface']}; color: {p['text']};
    border: 1px solid {p['border']}; padding: 4px;
}}
QProgressDialog {{ background: {p['surface']}; }}
QProgressBar {{
    background: {p['surface_alt']}; border: 1px solid {p['border']};
    border-radius: 5px; text-align: center; color: {p['text']};
}}
QProgressBar::chunk {{ background: {p['primary']}; border-radius: 4px; }}

/* Barra de herramientas de matplotlib embebida. */
QToolBar {{ background: {p['surface']}; border: none; spacing: 2px; }}
QToolButton {{ background: transparent; border: none; padding: 3px; }}
QToolButton:hover {{ background: {p['selection']}; border-radius: 4px; }}
"""
