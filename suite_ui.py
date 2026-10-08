#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BarcodeSuite shell UI: light/dark theming and the collapsible sidebar.

The utility panels in _utilities/ are shared verbatim with ONTbarcoder3 and
hard-code a light palette in their style sheets. Instead of editing them, every
style sheet is remapped on its way into Qt (QWidget.setStyleSheet is wrapped):
each colour is translated according to the property it sets (background,
border or text), so one source colour can become a dark fill and a light text
colour at the same time. Switching theme re-applies the stored source sheets.
"""
from __future__ import annotations
import re
import sys
import colorsys
from typing import Dict, List, Optional, Tuple
from PyQt5 import QtCore, QtGui, QtSvg, QtWidgets


# ═══════════════════════════════════════════════════════════════════════════
# PALETTES
# ═══════════════════════════════════════════════════════════════════════════

# Shell tokens (sidebar, window ground, tooltips) per theme.
TOKENS = {
    "light": {
        "ground": "#F4F6F5", "surface": "#FFFFFF", "line": "#DDE3E0",
        "text": "#1B2420", "text_sec": "#5C6B65", "text_hint": "#8A9A93",
        "accent": "#0F7A70", "accent_fg": "#0F6E66", "accent_tint": "#E3F1EE",
        "hover": "#EEF2F0", "tooltip_bg": "#1B2420", "tooltip_fg": "#FFFFFF",
        "tooltip_border": "#1B2420",
    },
    "dark": {
        "ground": "#11161B", "surface": "#182028", "line": "#2A3540",
        "text": "#E3E8EC", "text_sec": "#9AA6B1", "text_hint": "#6D7A86",
        "accent": "#0F7A70", "accent_fg": "#5CC9BC", "accent_tint": "#15302D",
        "hover": "#1D2731", "tooltip_bg": "#2A3540", "tooltip_fg": "#E3E8EC",
        "tooltip_border": "#3A4754",
    },
}

# Colour roles a CSS property assigns.
_BG, _BORDER, _FG = "bg", "border", "fg"

# Exact translations of the panels' palette (_utilities/shared.py); anything
# not listed goes through the generic rules in _translate().
_EXACT = {
    "light": {
        "#185FA5": "#0F7A70", "#0C4A82": "#0A5E56", "#083460": "#074540",
        "#378ADD": "#2A9D8F", "#E6F1FB": "#E3F1EE", "#203864": "#0F4F49",
        "#1A1A18": "#1B2420", "#6B6960": "#5C6B65", "#A09D96": "#8A9A93",
        "#E0DED8": "#DDE3E0", "#EDEDED": "#EEF2F0",
        ("#F5F5F3", _BG): "#FFFFFF",
        # Success green: soft emerald (same as ONTbarcoder3)
        "#3B6D11": "#0E7A55", "#EAF3DE": "#EAF7F1",
        ("#639922", _BORDER): "#9AD3BC", ("#639922", _BG): "#10875F",
    },
    "dark": {
        ("#185FA5", _BG): "#0F7A70", ("#185FA5", _BORDER): "#2A9D8F",
        ("#185FA5", _FG): "#5CC9BC",
        ("#0C4A82", _BG): "#0A5E56", ("#083460", _BG): "#074540",
        ("#378ADD", _BG): "#2A9D8F", ("#378ADD", _BORDER): "#2A9D8F",
        ("#378ADD", _FG): "#5CC9BC",
        ("#E6F1FB", _BG): "#15302D", ("#203864", _FG): "#8FD8CE",
        ("#1A1A18", _FG): "#E3E8EC", ("#6B6960", _FG): "#9AA6B1",
        ("#A09D96", _FG): "#6D7A86",
        ("#E0DED8", _BG): "#2A3540", ("#E0DED8", _BORDER): "#2A3540",
        ("#F5F5F3", _BG): "#182028", ("#FFFFFF", _BG): "#182028",
        ("#EDEDED", _BG): "#1D2731",
        # Status colours: readable tints on the dark ground.
        ("#3B6D11", _FG): "#5FD3A6", ("#A32D2D", _FG): "#F08A8A",
        ("#639922", _BORDER): "#2F6B58", ("#639922", _BG): "#10875F",
        ("#854F0B", _FG): "#F2B36B",
        ("#EAF3DE", _BG): "#18302B", ("#FCEBEB", _BG): "#3A1E1E",
        ("#FAEEDA", _BG): "#3A2C17",
    },
}

_THEME = ["light"]


def current_theme() -> str:
    return _THEME[0]


def tok(name: str) -> str:
    return TOKENS[_THEME[0]][name]


def _hex(c: str) -> str:
    c = c.upper()
    if len(c) == 4:                                   # #RGB → #RRGGBB
        c = "#" + "".join(ch * 2 for ch in c[1:])
    return c


def _hls(c: str) -> Tuple[float, float, float]:
    r, g, b = (int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return colorsys.rgb_to_hls(r, g, b)


def _from_hls(h: float, l: float, s: float) -> str:
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), max(0.0, min(1.0, s)))
    return "#{:02X}{:02X}{:02X}".format(round(r * 255), round(g * 255), round(b * 255))


def _translate(color: str, role: str, theme: str) -> str:
    c = _hex(color)
    table = _EXACT[theme]
    hit = table.get((c, role), table.get(c))
    if hit:
        return hit
    h, l, s = _hls(c)
    # The panels' blue accent family becomes the suite's teal.
    if s > 0.25 and 195 / 360 <= h <= 235 / 360:
        h = 174 / 360
    if theme == "light":
        return _from_hls(h, l, s)
    # Dark: invert lightness of pale fills/lines and of dark text; solid
    # mid-tone fills (buttons, progress chunks) keep their colour.
    if role == _FG:
        l = max(0.62, 0.10 + (1 - l) * 0.82) if l < 0.6 else l
    elif role == _BORDER:
        l = 0.20 + (1 - l) * 0.9 if l > 0.55 else l
    else:
        if l > 0.55:
            l, s = 0.08 + (1 - l) * 0.9, s * 0.6
    return _from_hls(h, l, s)


_DECL_RE = re.compile(r"([\w-]+)(\s*:\s*)([^;{}]+)")
_COLOR_RE = re.compile(r"#[0-9A-Fa-f]{6}\b|#[0-9A-Fa-f]{3}\b|\bwhite\b", re.I)


def _role(prop: str) -> Optional[str]:
    prop = prop.lower()
    if prop == "color" or prop.endswith("selection-color"):
        return _FG
    if "background" in prop or prop == "gridline-color":
        return _BG
    if prop.startswith("border") or prop.startswith("outline"):
        return _BORDER
    return None


def remap_css(css: str, theme: Optional[str] = None) -> str:
    """Translate every colour in a Qt style sheet to `theme`."""
    theme = theme or _THEME[0]
    if not css:
        return css

    def decl(m):
        role = _role(m.group(1))
        if role is None:
            return m.group(0)

        def col(cm):
            tok_ = cm.group(0)
            if tok_.lower() == "white":
                # 'white' text stays white (it sits on accent fills).
                return tok_ if role == _FG else _translate("#FFFFFF", role, theme)
            return _translate(tok_, role, theme)
        return m.group(1) + m.group(2) + _COLOR_RE.sub(col, m.group(3))
    return _DECL_RE.sub(decl, css)


def remap_qcolor(color: QtGui.QColor, role: str) -> QtGui.QColor:
    return QtGui.QColor(_translate(color.name(), role, _THEME[0]))


# ── Hooks into Qt ───────────────────────────────────────────────────────────

_SRC_PROP = "_suite_css_src"
_orig_set_ss = QtWidgets.QWidget.setStyleSheet
_orig_set_bg = QtWidgets.QTableWidgetItem.setBackground
_orig_set_fg = QtWidgets.QTableWidgetItem.setForeground


def _themed_set_ss(self, css):
    self.setProperty(_SRC_PROP, css)
    _orig_set_ss(self, remap_css(css))


def _themed_set_bg(self, brush):
    b = QtGui.QBrush(brush)
    if _THEME[0] == "dark" and b.style() != QtCore.Qt.NoBrush:
        b.setColor(remap_qcolor(b.color(), _BG))
        if self.foreground().style() == QtCore.Qt.NoBrush:
            _orig_set_fg(self, QtGui.QBrush(QtGui.QColor(tok("text"))))
    _orig_set_bg(self, b)


def _themed_set_fg(self, brush):
    b = QtGui.QBrush(brush)
    if _THEME[0] == "dark" and b.style() != QtCore.Qt.NoBrush:
        b.setColor(remap_qcolor(b.color(), _FG))
    _orig_set_fg(self, b)


class _TipMirror(QtWidgets.QLabel):
    """Rounded, anti-aliased stand-in for Qt's tooltip window.

    Qt's own tip label creates its native window before it can be made
    translucent, so on Windows a border-radius leaves black corners. The
    native tip keeps driving when tooltips show, move and hide (for every
    widget, item view and rich text alike) but is parked off-screen, since
    even fully transparent it keeps Windows' rectangular drop shadow; this
    window mirrors its text at the position Qt chose for it."""

    RADIUS = 16

    def __init__(self):
        super().__init__(None, QtCore.Qt.ToolTip | QtCore.Qt.FramelessWindowHint
                         | QtCore.Qt.NoDropShadowWindowHint)
        self.setObjectName("suite_tooltip")
        for attr in (QtCore.Qt.WA_TranslucentBackground, QtCore.Qt.WA_ShowWithoutActivating,
                     QtCore.Qt.WA_TransparentForMouseEvents):
            self.setAttribute(attr, True)
        self.setContentsMargins(14, 7, 14, 7)

    def sync(self, tip: QtWidgets.QLabel, pos: QtCore.QPoint):
        self.setTextFormat(tip.textFormat())
        self.setWordWrap(tip.wordWrap())
        self.setText(tip.text())
        self.adjustSize()
        geo = QtCore.QRect(pos, self.size())
        screen = QtWidgets.QApplication.screenAt(pos) or QtWidgets.QApplication.primaryScreen()
        avail = screen.availableGeometry()
        geo.moveRight(min(geo.right(), avail.right()))
        geo.moveBottom(min(geo.bottom(), avail.bottom()))
        self.move(geo.topLeft())
        self.show()
        self.raise_()

    def paintEvent(self, event):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        p.setPen(QtGui.QPen(QtGui.QColor(tok("tooltip_border")), 1))
        p.setBrush(QtGui.QColor(tok("tooltip_bg")))
        p.drawRoundedRect(QtCore.QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5),
                          self.RADIUS, self.RADIUS)
        p.end()
        super().paintEvent(event)


class _RoundTooltips(QtCore.QObject):
    """Hides Qt's native tip label and mirrors it with _TipMirror."""

    _SYNC = (QtCore.QEvent.Show, QtCore.QEvent.Move, QtCore.QEvent.Resize)
    _PARKED = QtCore.QPoint(-32000, -32000)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._mirror = None
        self._anchor = QtCore.QPoint()

    def eventFilter(self, obj, event):
        et = event.type()
        if et in self._SYNC or et in (QtCore.QEvent.Hide, QtCore.QEvent.Polish):
            if obj.metaObject().className() == "QTipLabel":
                if et == QtCore.QEvent.Hide:
                    if self._mirror is not None:
                        self._mirror.hide()
                else:
                    obj.setWindowOpacity(0.0)
                    if et != QtCore.QEvent.Polish and obj.isVisible():
                        # Qt (re)placed the tip: remember where, then park it.
                        if obj.pos() != self._PARKED:
                            self._anchor = obj.pos()
                        if self._mirror is None:
                            self._mirror = _TipMirror()
                        self._mirror.sync(obj, self._anchor)
                        if obj.pos() != self._PARKED:
                            obj.move(self._PARKED)
        return False


_TOOLTIP_FILTER = []


class _NoWheelEdits(QtCore.QObject):
    """Stops the mouse wheel from changing spin boxes, combo boxes and
    sliders, so scrolling a page never edits the field under the cursor. The
    event is ignored rather than eaten, so Qt passes it on to the parent
    scroll area and the page still scrolls."""

    _TYPES = (QtWidgets.QAbstractSpinBox, QtWidgets.QComboBox, QtWidgets.QAbstractSlider)

    def eventFilter(self, obj, event):
        if (event.type() == QtCore.QEvent.Wheel
                and isinstance(obj, self._TYPES)
                and not isinstance(obj, QtWidgets.QScrollBar)):
            event.ignore()
            return True
        return False


_WHEEL_FILTER = []


def install_hooks():
    """Must run after QApplication exists and before any panel is built."""
    app = QtWidgets.QApplication.instance()
    if app is not None and not _TOOLTIP_FILTER:
        _TOOLTIP_FILTER.append(_RoundTooltips(app))
        app.installEventFilter(_TOOLTIP_FILTER[0])
    if app is not None and not _WHEEL_FILTER:
        _WHEEL_FILTER.append(_NoWheelEdits(app))
        app.installEventFilter(_WHEEL_FILTER[0])
    QtWidgets.QWidget.setStyleSheet = _themed_set_ss
    QtWidgets.QTableWidgetItem.setBackground = _themed_set_bg
    QtWidgets.QTableWidgetItem.setForeground = _themed_set_fg


# ── Shell style sheet ───────────────────────────────────────────────────────

def _style_assets(t: dict) -> Dict[str, str]:
    """Write the theme's indicator / arrow SVGs to the cache folder and return
    their paths for url() (Qt style sheets cannot take inline images)."""
    base = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.CacheLocation)
    folder = QtCore.QDir(base or QtCore.QDir.tempPath()).filePath(f"suite_ui/{_THEME[0]}")
    QtCore.QDir().mkpath(folder)
    svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" '           'stroke="{c}" stroke-width="{w}" stroke-linecap="round" stroke-linejoin="round">{d}</svg>'
    shapes = {
        "check": svg.format(c="#FFFFFF", w=3.2, d='<path d="M5 12.5l4.5 4.5L19 7.5"/>'),
        "dot":   '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">'
                 '<circle cx="12" cy="12" r="5.5" fill="#FFFFFF"/></svg>',
        "down":  svg.format(c=t["text_sec"], w=2.4, d='<path d="M6 9l6 6 6-6"/>'),
        "up":    svg.format(c=t["text_sec"], w=2.4, d='<path d="M6 15l6-6 6 6"/>'),
    }
    paths = {}
    for name, data in shapes.items():
        path = QtCore.QDir(folder).filePath(f"{name}.svg")
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(data)
        except OSError:
            pass
        paths[name] = path.replace("\\", "/")
    return paths


def shell_css() -> str:
    t = TOKENS[_THEME[0]]
    a = _style_assets(t)
    return f"""
QCheckBox::indicator, QRadioButton::indicator {{
    width: 16px; height: 16px; border: 1.5px solid {t['text_hint']};
    background-color: {t['surface']};
}}
QCheckBox::indicator {{ border-radius: 4px; }}
QRadioButton::indicator {{ border-radius: 9px; }}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{ border-color: {t['accent_fg']}; }}
QCheckBox::indicator:checked {{
    background-color: {t['accent']}; border-color: {t['accent']}; image: url({a['check']});
}}
QRadioButton::indicator:checked {{
    background-color: {t['accent']}; border-color: {t['accent']}; image: url({a['dot']});
}}
QCheckBox::indicator:disabled, QRadioButton::indicator:disabled {{
    background-color: {t['hover']}; border-color: {t['line']};
}}
QCheckBox::indicator:checked:disabled, QRadioButton::indicator:checked:disabled {{
    background-color: {t['text_hint']}; border-color: {t['text_hint']};
}}
QComboBox::drop-down {{ border: none; width: 24px; }}
QComboBox::down-arrow {{ image: url({a['down']}); width: 14px; height: 14px; }}
QComboBox QAbstractItemView {{
    background-color: {t['surface']}; color: {t['text']};
    selection-background-color: {t['accent_tint']}; selection-color: {t['text']};
    border: 1px solid {t['line']}; outline: none;
}}
QAbstractSpinBox::up-button, QAbstractSpinBox::down-button {{
    border: none; width: 18px; background: transparent;
}}
QAbstractSpinBox::up-arrow {{ image: url({a['up']}); width: 12px; height: 12px; }}
QAbstractSpinBox::down-arrow {{ image: url({a['down']}); width: 12px; height: 12px; }}
QMenu {{
    background-color: {t['surface']}; color: {t['text']}; border: 1px solid {t['line']};
}}
QMenu::item:selected {{ background-color: {t['accent_tint']}; }}
QMainWindow, #suite_content {{ background-color: {t['ground']}; }}
QToolTip {{
    background-color: {t['tooltip_bg']}; color: {t['tooltip_fg']};
    border: 1px solid {t['tooltip_border']}; padding: 6px 10px; font-size: 14px;
}}
#suite_tooltip {{
    background: transparent; border: none; padding: 0;
    color: {t['tooltip_fg']}; font-size: 14px;
}}
QScrollBar:vertical {{ width: 10px; background: transparent; margin: 2px; }}
QScrollBar:horizontal {{ height: 10px; background: transparent; margin: 2px; }}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
    background: {t['line']}; border-radius: 3px; min-height: 24px; min-width: 24px;
}}
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {{ background: {t['text_hint']}; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
#suite_sidebar {{
    background-color: {t['surface']};
    border-right: 1px solid {t['line']};
}}
#suite_brand {{ font-size: 18px; font-weight: 700; color: {t['text']}; }}
#suite_version {{ font-size: 13px; color: {t['text_sec']}; }}
#suite_section {{
    font-size: 12px; font-weight: 600; letter-spacing: 1px;
    color: {t['text_hint']}; padding: 14px 12px 4px 12px;
}}
#suite_divider {{ background-color: {t['line']}; }}
#suite_nav, #suite_util {{
    text-align: left; padding: 0 12px; border: none; border-radius: 8px;
    font-size: 16px; color: {t['text']}; background: transparent;
}}
#suite_util {{ color: {t['text_sec']}; font-size: 15px; }}
#suite_nav[collapsed="true"], #suite_util[collapsed="true"] {{
    text-align: center; padding: 0;
}}
#suite_nav:hover, #suite_util:hover {{ background-color: {t['hover']}; }}
#suite_nav:checked {{
    background-color: {t['accent_tint']}; color: {t['accent_fg']}; font-weight: 600;
}}
#suite_nav:focus, #suite_util:focus {{ outline: none; }}
"""


def app_palette() -> QtGui.QPalette:
    t = TOKENS[_THEME[0]]
    pal = QtGui.QPalette()
    for role, key in ((QtGui.QPalette.Window, "ground"), (QtGui.QPalette.Base, "surface"),
                      (QtGui.QPalette.AlternateBase, "hover"), (QtGui.QPalette.Button, "surface"),
                      (QtGui.QPalette.WindowText, "text"), (QtGui.QPalette.Text, "text"),
                      (QtGui.QPalette.ButtonText, "text"), (QtGui.QPalette.ToolTipBase, "tooltip_bg"),
                      (QtGui.QPalette.ToolTipText, "tooltip_fg"), (QtGui.QPalette.Highlight, "accent"),
                      (QtGui.QPalette.PlaceholderText, "text_hint")):
        pal.setColor(role, QtGui.QColor(t[key]))
    pal.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor("#FFFFFF"))
    pal.setColor(QtGui.QPalette.Link, QtGui.QColor(t["accent_fg"]))
    # Fusion draws frames, group boxes and indicators from these roles.
    for role in (QtGui.QPalette.Light, QtGui.QPalette.Midlight, QtGui.QPalette.Mid,
                 QtGui.QPalette.Dark, QtGui.QPalette.Shadow):
        pal.setColor(role, QtGui.QColor(t["line"]))
    for group in (QtGui.QPalette.Disabled,):
        for role in (QtGui.QPalette.Text, QtGui.QPalette.WindowText, QtGui.QPalette.ButtonText):
            pal.setColor(group, role, QtGui.QColor(t["text_hint"]))
    return pal


def apply_theme(app: QtWidgets.QApplication, theme: str, base_css: str):
    """Switch theme live: app sheet, palette and every widget's own sheet."""
    _THEME[0] = theme if theme in TOKENS else "light"
    # Fusion honours the palette everywhere (check boxes, group boxes,
    # combo arrows), which the native Windows style ignores in dark mode.
    if app.style().objectName().lower() != "fusion":
        app.setStyle("Fusion")
    app.setPalette(app_palette())
    app.setStyleSheet(remap_css(base_css) + shell_css())
    for w in app.allWidgets():
        src = w.property(_SRC_PROP)
        if src:
            _orig_set_ss(w, remap_css(src))
    for w in app.topLevelWidgets():
        set_dark_titlebar(w, _THEME[0] == "dark")


def set_dark_titlebar(widget: QtWidgets.QWidget, dark: bool):
    """Windows 10/11: match the native title bar to the theme."""
    if sys.platform != "win32" or not widget.isWindow():
        return
    try:
        import ctypes
        val = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):    # DWMWA_USE_IMMERSIVE_DARK_MODE (new, old builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    int(widget.winId()), attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                break
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════════════
# COMPARE RESULTS WINDOW
# ═══════════════════════════════════════════════════════════════════════════
# The window in compare_panel.py hard-codes a navy title bar and table header,
# plus white rich-text title colours that remap_css cannot reach. Its sheets
# are rewritten here in the panels' own palette so remap_css themes them (and
# re-themes them on a live switch). Styles are matched on their source text.

_CMP_TOPBAR = "#cmp_topbar { background-color: #FFFFFF; border-bottom: 1px solid #E0DED8; }"

# Old saturated pills / beige bands → the palette of the table's state chips.
_CMP_SWAP = {
    "#C5EAAB": "#EAF3DE", "#2D6A0A": "#3B6D11",     # identical
    "#A8CCF0": "#E6F1FB", "#0D4D8A": "#185FA5",     # compatible
    "#F5AAAA": "#FCEBEB", "#8B1A1A": "#A32D2D",     # different
    "#F5D9A8": "#FAEEDA", "#6B3D0A": "#854F0B",     # only in reference
    "#CCCCCC": "#EBEBEB", "#DDDDDD": "#EBEBEB",     # unique / no reference
    "#F0EEE8": "#EDEDED", "#F7F6F2": "#F5F5F3",     # output path / legend bands
}
_CMP_TABLE = """
QTableWidget { background-color: #FFFFFF; border: none; outline: none;
               font-size: 13px; gridline-color: transparent; }
QTableWidget::item { padding: 5px 10px; border-bottom: 1px solid #E0DED8; }
QTableWidget::item:selected { background-color: #E6F1FB; color: #1A1A18; }
QHeaderView::section { background-color: #EDEDED; color: #6B6960; font-weight: 600;
                       font-size: 12px; padding: 8px 10px; border: none;
                       border-bottom: 1px solid #E0DED8; }
"""


def restyle_compare_results(dlg: QtWidgets.QWidget):
    for w in dlg.findChildren(QtWidgets.QWidget):
        src = w.property(_SRC_PROP) or ""
        if "#1A1A2E" in src and "QHeaderView" in src:
            w.setStyleSheet(_CMP_TABLE)
        elif src.strip() == "background-color: #1A1A2E;":
            w.setObjectName("cmp_topbar")
            w.setStyleSheet(_CMP_TOPBAR)
        elif any(k in src for k in _CMP_SWAP):
            for old, new in _CMP_SWAP.items():
                src = src.replace(old, new)
            w.setStyleSheet(src)
        if isinstance(w, QtWidgets.QLabel) and "color:white;font-size:18px;" in w.text():
            w.setText(w.text().replace("color:white;", "")
                      .replace("color:#9AAFCC;", "font-weight:400;"))
            w.setStyleSheet("color: #1A1A18; background: transparent;")
    _fit_state_column(dlg)
    set_dark_titlebar(dlg, current_theme() == "dark")


def _fit_state_column(dlg: QtWidgets.QWidget):
    """The results table sizes its columns from the cell text, but the State
    column holds pill widgets with extra padding, so a long state such as
    "Unique in A" came out clipped. Widen the column to the widest pill."""
    for table in dlg.findChildren(QtWidgets.QTableWidget):
        for col in range(table.columnCount()):
            widths = [table.cellWidget(r, col).sizeHint().width() + 8
                      for r in range(table.rowCount()) if table.cellWidget(r, col)]
            if widths and max(widths) > table.columnWidth(col):
                table.setColumnWidth(col, max(widths))


def install_compare_results_hook(cls):
    """Restyle every instance of compare_panel._CompareResultsWindow."""
    orig_init = cls.__init__

    def __init__(self, *args, **kwargs):
        orig_init(self, *args, **kwargs)
        restyle_compare_results(self)
    cls.__init__ = __init__


# ═══════════════════════════════════════════════════════════════════════════
# ICONS (inline stroke SVG, recoloured per theme / state)
# ═══════════════════════════════════════════════════════════════════════════

_ICON_PATHS = {
    "fastq":   '<path d="M3 20V10M9 20V4M15 20v-7M21 20v-11"/>',
    "fasta":   '<path d="M14 3H6v18h12V7z"/><path d="M14 3v4h4"/><path d="M9 13h6M9 17h4"/>',
    "compare": '<path d="M8 3v18M16 3v18M3 8h5M16 16h5"/>',
    "blast":   '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>',
    "best":    '<path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.5 6.7 19.4l1.2-6L3.4 9.3l6-.7z"/>',
    "genbank": '<ellipse cx="12" cy="6" rx="8" ry="3"/><path d="M4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
    "bold":    '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 10v10"/>',
    "help":    '<circle cx="12" cy="12" r="9.5"/><path d="M9 9.5a3 3 0 1 1 4.5 2.6c-1 .6-1.5 1.1-1.5 2.2"/><circle cx="12" cy="17.4" r=".6"/>',
    "moon":    '<path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/>',
    "sun":     '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
    "collapse": '<path d="M11 17l-5-5 5-5M18 17l-5-5 5-5"/>',
    "expand":  '<path d="M13 17l5-5-5-5M6 17l5-5-5-5"/>',
}

_LOGO_SVG = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 36 36">'
             '<rect width="36" height="36" rx="9" fill="{bg}"/>'
             '<g fill="#FFFFFF"><rect x="8" y="9" width="2" height="18"/>'
             '<rect x="12" y="9" width="4" height="18"/><rect x="18" y="9" width="1.5" height="18"/>'
             '<rect x="21.5" y="9" width="3" height="18"/><rect x="26.5" y="9" width="1.5" height="18"/>'
             '</g></svg>')


def _svg_pixmap(svg: str, size: int, dpr: float) -> QtGui.QPixmap:
    px = QtGui.QPixmap(int(size * dpr), int(size * dpr))
    px.fill(QtCore.Qt.transparent)
    p = QtGui.QPainter(px)
    QtSvg.QSvgRenderer(QtCore.QByteArray(svg.encode())).render(p)
    p.end()
    px.setDevicePixelRatio(dpr)
    return px


def make_icon(name: str, color: str, checked_color: Optional[str] = None,
              size: int = 20) -> QtGui.QIcon:
    dpr = QtWidgets.QApplication.instance().devicePixelRatio()

    def svg(c):
        return ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" '
                f'stroke="{c}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
                f'{_ICON_PATHS[name]}</svg>')
    icon = QtGui.QIcon()
    icon.addPixmap(_svg_pixmap(svg(color), size, dpr), QtGui.QIcon.Normal, QtGui.QIcon.Off)
    if checked_color:
        icon.addPixmap(_svg_pixmap(svg(checked_color), size, dpr), QtGui.QIcon.Normal, QtGui.QIcon.On)
    return icon


# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════

class Sidebar(QtWidgets.QFrame):
    """Grouped tool navigation; collapses to an icon rail with tooltips."""

    toolSelected = QtCore.pyqtSignal(str)        # tool key
    themeToggled = QtCore.pyqtSignal()
    collapsedChanged = QtCore.pyqtSignal(bool)

    EXPANDED_W = 252
    COLLAPSED_W = 72

    def __init__(self, sections: List[Tuple[str, List[Tuple[str, str, str]]]],
                 version: str, parent=None):
        """sections: [(SECTION TITLE, [(key, label, icon name), ...]), ...]"""
        super().__init__(parent)
        self.setObjectName("suite_sidebar")
        self._collapsed = False
        self._buttons: Dict[str, QtWidgets.QPushButton] = {}
        self._labels: Dict[str, str] = {}
        self._icons: Dict[str, str] = {}
        self._section_lbls: List[QtWidgets.QLabel] = []
        self._dividers: List[QtWidgets.QFrame] = []
        self._group = QtWidgets.QButtonGroup(self)
        self._group.setExclusive(True)

        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(12, 18, 12, 14)
        lay.setSpacing(2)

        # Brand
        brand = QtWidgets.QHBoxLayout()
        brand.setContentsMargins(4, 0, 0, 10)
        brand.setSpacing(10)
        self._logo = QtWidgets.QLabel()
        self._logo.setFixedSize(36, 36)
        brand.addWidget(self._logo)
        self._brand_box = QtWidgets.QWidget()
        bl = QtWidgets.QVBoxLayout(self._brand_box)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(0)
        name = QtWidgets.QLabel()
        name.setObjectName("suite_brand")
        name.setTextFormat(QtCore.Qt.RichText)
        self._brand_name = name
        ver = QtWidgets.QLabel(f"v{version}")
        ver.setObjectName("suite_version")
        bl.addWidget(name)
        bl.addWidget(ver)
        brand.addWidget(self._brand_box, 1)
        lay.addLayout(brand)

        for title, items in sections:
            lbl = QtWidgets.QLabel(title.upper())
            lbl.setObjectName("suite_section")
            self._section_lbls.append(lbl)
            lay.addWidget(lbl)
            div = QtWidgets.QFrame()
            div.setObjectName("suite_divider")
            div.setFixedHeight(1)
            div.hide()
            self._dividers.append(div)
            lay.addSpacing(0)
            lay.addWidget(div)
            for key, label, icon in items:
                btn = self._make_button("suite_nav", label)
                btn.setCheckable(True)
                btn.clicked.connect(lambda _=False, k=key: self.toolSelected.emit(k))
                self._group.addButton(btn)
                self._buttons[key] = btn
                self._labels[key] = label
                self._icons[key] = icon
                lay.addWidget(btn)

        lay.addStretch(1)
        self._theme_btn = self._make_button("suite_util", "")
        self._theme_btn.clicked.connect(self.themeToggled.emit)
        lay.addWidget(self._theme_btn)
        self._collapse_btn = self._make_button("suite_util", "")
        self._collapse_btn.clicked.connect(lambda: self.set_collapsed(not self._collapsed, True))
        lay.addWidget(self._collapse_btn)

        self._anim = QtCore.QVariantAnimation(self)
        self._anim.setDuration(180)
        self._anim.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        self._anim.valueChanged.connect(lambda v: self.setFixedWidth(int(v)))
        self.setFixedWidth(self.EXPANDED_W)
        self.refresh()

    @staticmethod
    def _make_button(obj_name: str, text: str) -> QtWidgets.QPushButton:
        b = QtWidgets.QPushButton(text)
        b.setObjectName(obj_name)
        b.setFixedHeight(44)
        b.setIconSize(QtCore.QSize(20, 20))
        b.setCursor(QtCore.Qt.PointingHandCursor)
        return b

    # ── State ───────────────────────────────────────────────────────────────

    def is_collapsed(self) -> bool:
        return self._collapsed

    def select(self, key: str):
        if key in self._buttons:
            self._buttons[key].setChecked(True)

    def set_collapsed(self, collapsed: bool, animate: bool = False):
        self._collapsed = collapsed
        target = self.COLLAPSED_W if collapsed else self.EXPANDED_W
        if animate:
            self._anim.stop()
            self._anim.setStartValue(self.width())
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self.setFixedWidth(target)
        self.refresh()
        self.collapsedChanged.emit(collapsed)

    def refresh(self):
        """Re-apply texts, tooltips and icon colours (collapse / theme change)."""
        c = self._collapsed
        dark = current_theme() == "dark"
        dpr = QtWidgets.QApplication.instance().devicePixelRatio()
        self._logo.setPixmap(_svg_pixmap(_LOGO_SVG.format(bg=tok("accent")), 36, dpr))
        self._brand_box.setVisible(not c)
        self._brand_name.setText(f"Barcode<span style='color:{tok('accent_fg')};'>Suite</span>")
        for lbl in self._section_lbls:
            lbl.setVisible(not c)
        for div in self._dividers:
            div.setVisible(c)

        for key, btn in self._buttons.items():
            btn.setIcon(make_icon(self._icons[key], tok("text_sec"), tok("accent_fg")))
            btn.setText("" if c else f"  {self._labels[key]}")
            btn.setToolTip(self._labels[key] if c else "")

        theme_label = "Light mode" if dark else "Dark mode"
        self._theme_btn.setIcon(make_icon("sun" if dark else "moon", tok("text_sec")))
        self._theme_btn.setText("" if c else f"  {theme_label}")
        self._theme_btn.setToolTip(theme_label if c else "")

        self._collapse_btn.setIcon(make_icon("expand" if c else "collapse", tok("text_sec")))
        self._collapse_btn.setText("" if c else "  Collapse sidebar")
        self._collapse_btn.setToolTip("Expand sidebar (Ctrl+B)" if c else "Ctrl+B")

        for btn in list(self._buttons.values()) + [self._theme_btn, self._collapse_btn]:
            btn.setProperty("collapsed", "true" if c else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)
